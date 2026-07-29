// Vercel Edge Function. Keeps the Anthropic API key on the server —
// nobody using the app ever sees or enters it. Set ANTHROPIC_API_KEY as
// a plain (non-VITE_) environment variable in your Vercel project
// settings, then redeploy.

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

  const body = await req.json();

  const anthropicRes = await fetch("https://api.anthropic.com/v1/messages", {
    method: "POST",
    headers: {
      "Content-Type": "application/json",
      "x-api-key": apiKey,
      "anthropic-version": "2023-06-01",
    },
    body: JSON.stringify(body),
  });

  const data = await anthropicRes.json();
  return new Response(JSON.stringify(data), {
    status: anthropicRes.status,
    headers: { "Content-Type": "application/json" },
  });
}
