// Vercel Edge Function. Keeps the xAI key on the server.
// Set XAI_API_KEY in Vercel Project Settings (Secret, no VITE_ prefix), then redeploy.
//
// The browser still POSTs the old Anthropic-shaped body (PDF base64 + prompt).
// This function unpacks that, sends the PDF to Grok, and streams the reply back
// as Anthropic-style SSE so App.jsx does not need to change.

export const config = { runtime: "edge" };

const XAI_FILES = "https://api.x.ai/v1/files";
const XAI_RESPONSES = "https://api.x.ai/v1/responses";
const MODEL = "grok-4.6";

function jsonError(status, message) {
  return new Response(JSON.stringify({ error: message }), {
    status,
    headers: { "Content-Type": "application/json" },
  });
}

function extractPdfAndPrompt(body) {
  const parts = (((body || {}).messages || [])[0] || {}).content;
  if (!Array.isArray(parts)) return { error: "Request is missing messages[0].content" };
  let base64 = "";
  let prompt = "";
  for (const part of parts) {
    if (part && part.type === "document" && part.source && part.source.data) {
      base64 = String(part.source.data);
    } else if (part && part.type === "text" && part.text) {
      prompt = String(part.text);
    }
  }
  if (!base64) return { error: "Request is missing the PDF (document.source.data)" };
  if (!prompt) return { error: "Request is missing the scan prompt (text)" };
  return { base64, prompt };
}

function base64ToBytes(b64) {
  const bin = atob(b64);
  const out = new Uint8Array(bin.length);
  for (let i = 0; i < bin.length; i += 1) out[i] = bin.charCodeAt(i);
  return out;
}

async function uploadPdf(apiKey, base64) {
  const bytes = base64ToBytes(base64);
  const form = new FormData();
  form.append("purpose", "assistants");
  form.append("file", new Blob([bytes], { type: "application/pdf" }), "packing-list.pdf");
  const res = await fetch(XAI_FILES, {
    method: "POST",
    headers: { Authorization: `Bearer ${apiKey}` },
    body: form,
  });
  const text = await res.text();
  let data;
  try {
    data = JSON.parse(text);
  } catch (err) {
    throw new Error(`xAI files upload returned ${res.status}: ${text.slice(0, 200)}`);
  }
  if (!res.ok || !data.id) {
    throw new Error((data.error && (data.error.message || data.error)) || `xAI files upload returned ${res.status}`);
  }
  return data.id;
}

function anthropicTextEvent(text) {
  return `data: ${JSON.stringify({ type: "content_block_delta", delta: { text } })}\n\n`;
}

function anthropicStop(reason) {
  return `data: ${JSON.stringify({ type: "message_delta", delta: { stop_reason: reason || "end_turn" } })}\n\n`;
}

function pullTextFromXaiEvent(evt) {
  if (!evt || typeof evt !== "object") return "";
  if (typeof evt.delta === "string") return evt.delta;
  if (evt.delta && typeof evt.delta.text === "string") return evt.delta.text;
  if (evt.delta && typeof evt.delta.content === "string") return evt.delta.content;
  if (typeof evt.text === "string" && (evt.type || "").includes("output_text")) return evt.text;
  const choice = (evt.choices && evt.choices[0]) || null;
  if (choice && choice.delta && typeof choice.delta.content === "string") return choice.delta.content;
  return "";
}

export default async function handler(req) {
  if (req.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }

  const apiKey = process.env.XAI_API_KEY;
  if (!apiKey) {
    return jsonError(
      500,
      "Server is missing XAI_API_KEY. Add it in Vercel Project Settings -> Environment Variables (Secret, not VITE_), then redeploy."
    );
  }

  let body;
  try {
    body = await req.json();
  } catch (err) {
    return jsonError(400, "Request body was not valid JSON.");
  }

  const extracted = extractPdfAndPrompt(body);
  if (extracted.error) return jsonError(400, extracted.error);

  let fileId;
  try {
    fileId = await uploadPdf(apiKey, extracted.base64);
  } catch (err) {
    return jsonError(502, `Could not upload PDF to xAI: ${err && err.message ? err.message : err}`);
  }

  let xaiRes;
  try {
    xaiRes = await fetch(XAI_RESPONSES, {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        Authorization: `Bearer ${apiKey}`,
      },
      body: JSON.stringify({
        model: MODEL,
        stream: true,
        input: [
          {
            role: "user",
            content: [
              { type: "input_text", text: extracted.prompt },
              { type: "input_file", file_id: fileId },
            ],
          },
        ],
      }),
    });
  } catch (err) {
    return jsonError(502, `Could not reach the xAI API: ${err && err.message ? err.message : err}`);
  }

  if (!xaiRes.ok || !xaiRes.body) {
    const text = await xaiRes.text();
    let message = `xAI API returned ${xaiRes.status}`;
    try {
      const parsed = JSON.parse(text);
      message = (parsed.error && (parsed.error.message || parsed.error)) || parsed.message || message;
    } catch (err) {
      if (text) message = text.slice(0, 300);
    }
    return jsonError(xaiRes.status === 402 ? 402 : xaiRes.status, message);
  }

  const encoder = new TextEncoder();
  const decoder = new TextDecoder();
  const stream = new ReadableStream({
    async start(controller) {
      const reader = xaiRes.body.getReader();
      let buffer = "";
      const send = (chunk) => controller.enqueue(encoder.encode(chunk));
      try {
        for (;;) {
          const { done, value } = await reader.read();
          if (done) break;
          buffer += decoder.decode(value, { stream: true });
          let sep;
          while ((sep = buffer.indexOf("\n\n")) !== -1) {
            const raw = buffer.slice(0, sep);
            buffer = buffer.slice(sep + 2);
            for (const line of raw.split("\n")) {
              const trimmed = line.trim();
              if (!trimmed.startsWith("data:")) continue;
              const payload = trimmed.slice(5).trim();
              if (!payload || payload === "[DONE]") continue;
              let evt;
              try {
                evt = JSON.parse(payload);
              } catch (err) {
                continue;
              }
              if (evt.type === "error" || evt.error) {
                const msg = (evt.error && (evt.error.message || evt.error)) || "xAI stream error";
                send(`data: ${JSON.stringify({ type: "error", error: { message: msg } })}\n\n`);
                continue;
              }
              const text = pullTextFromXaiEvent(evt);
              if (text) send(anthropicTextEvent(text));
            }
          }
        }
        send(anthropicStop("end_turn"));
        send("data: [DONE]\n\n");
      } catch (err) {
        send(`data: ${JSON.stringify({ type: "error", error: { message: String(err && err.message ? err.message : err) } })}\n\n`);
      } finally {
        controller.close();
      }
    },
  });

  return new Response(stream, {
    status: 200,
    headers: {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
      "X-Accel-Buffering": "no",
    },
  });
}
