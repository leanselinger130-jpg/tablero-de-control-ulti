const ANTHROPIC_API_URL = "https://api.anthropic.com/v1/messages";
const MAX_CONTEXT_CHARS = 600000; // safety net only — the client already trims to fit, keeping the most recent round intact

const json = (data, status = 200) =>
  new Response(JSON.stringify(data), { status, headers: { "Content-Type": "application/json" } });

export default async (req) => {
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);

  const apiKey = process.env.ANTHROPIC_API_KEY;
  if (!apiKey) {
    return json({ error: "Falta configurar la variable de entorno ANTHROPIC_API_KEY en Netlify (Project configuration → Environment variables)." }, 500);
  }

  let body;
  try { body = await req.json(); } catch { return json({ error: "Invalid JSON body" }, 400); }

  const { question, context } = body || {};
  if (!question || typeof question !== "string") return json({ error: "Falta la pregunta." }, 400);

  const contextJson = JSON.stringify(context || {}).slice(0, MAX_CONTEXT_CHARS);

  const system = `Sos un analista estratégico ayudando a un equipo universitario que juega la simulación Cesim Global Challenge, en el marco de una materia de Dirección General.

Respondé siempre en español, de forma clara, directa y concisa. Basate ÚNICAMENTE en los datos JSON provistos a continuación (resultados de las rondas cargadas del reporte "Results" de Cesim, el plan de estrategia del equipo y las premisas cargadas, si existen). Si la pregunta no se puede responder con esos datos, decilo explícitamente en vez de inventar cifras. Cuando compares equipos o rondas, sé específico con los números. El equipo del usuario es "ourTeam" en los datos; los demás son competidores.

DATOS (JSON):
${contextJson}`;

  try {
    const res = await fetch(ANTHROPIC_API_URL, {
      method: "POST",
      headers: { "Content-Type": "application/json", "x-api-key": apiKey, "anthropic-version": "2023-06-01" },
      body: JSON.stringify({
        model: process.env.ANTHROPIC_MODEL || "claude-sonnet-5",
        max_tokens: 1024,
        system,
        messages: [{ role: "user", content: question }],
      }),
    });
    const data = await res.json();
    if (!res.ok) return json({ error: data?.error?.message || `Error de la API de Anthropic (${res.status}).` }, res.status);
    const answer = (data.content || []).filter((b) => b.type === "text").map((b) => b.text).join("\n").trim();
    return json({ answer: answer || "No obtuve una respuesta de texto." });
  } catch (e) {
    return json({ error: String(e?.message || e) }, 500);
  }
};
