// Vercel Edge Function. Keeps the Anthropic API key on the server —
// nobody using the app ever sees or enters it. Set ANTHROPIC_API_KEY as
// a plain (non-VITE_) environment variable in your Vercel project
// settings, then redeploy.
//
// This proxy STREAMS the reply back rather than waiting for the whole of it.
//
// An edge function is limited on how long it may take to start responding, not on how long
// it may go on for. The previous version awaited anthropicRes.json(), so it sat holding the
// connection open for the entire generation and sent nothing at all until the end — a long
// packing list (the 19-page Fujitec one, seventy cases) took longer than that limit and
// Vercel killed the invocation with FUNCTION_INVOCATION_TIMEOUT.
//
// Asking Anthropic to stream means the first bytes arrive in about a second, which satisfies
// the limit, and the rest flows through for as long as it needs. The client reassembles the
// text from the stream.

export const config = { runtime: "edge" };

export default async function handler(req) {
  if (req.method !== "POST") {
    return new Response("Method not allowed", { status: 405 });
  }

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return new Response(
      JSON.stringify({ error: "Server is missing ANTHROPIC_API_KEY. Add it in Vercel Project Settings -> Environment Variables, then redeploy." }),
      { status: 500, headers: { "Content-Type": "application/json" } }
    );
  }

  let body;
  try {
    body = await req.json();
  } catch (err) {
    return new Response(JSON.stringify({ error: "Request body was not valid JSON." }), {
      status: 400,
      headers: { "Content-Type": "application/json" },
    });
  }

  let anthropicRes;
  try {
    anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
      method: "POST",
      headers: {
        "Content-Type": "application/json",
        "x-api-key": apiKey,
        "anthropic-version": "2023-06-01",
      },
      // Streaming is forced here rather than left to the caller, so that no request can
      // accidentally go back to the buffered behaviour that caused the timeout.
      body: JSON.stringify({ ...body, stream: true }),
    });
  } catch (err) {
    return new Response(JSON.stringify({ error: `Could not reach the Anthropic API: ${err && err.message ? err.message : err}` }), {
      status: 502,
      headers: { "Content-Type": "application/json" },
    });
  }

  // An error is returned as ordinary JSON, not as a stream, so pass it through unchanged and
  // let the client report what it says.
  if (!anthropicRes.ok || !anthropicRes.body) {
    const text = await anthropicRes.text();
    return new Response(text || JSON.stringify({ error: `Anthropic API returned ${anthropicRes.status}` }), {
      status: anthropicRes.status,
      headers: { "Content-Type": "application/json" },
    });
  }

  return new Response(anthropicRes.body, {
    status: 200,
    headers: {
      "Content-Type": "text/event-stream; charset=utf-8",
      "Cache-Control": "no-cache, no-transform",
      Connection: "keep-alive",
      // Stops any proxy in front of this from buffering the stream and reintroducing the
      // very delay the streaming is here to avoid.
      "X-Accel-Buffering": "no",
    },
  });
}
