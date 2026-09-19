import { getStore } from "@netlify/blobs";

const json = (data, status = 200) =>
  new Response(JSON.stringify(data), { status, headers: { "Content-Type": "application/json" } });

export default async (req) => {
  if (req.method !== "POST") return json({ error: "Method not allowed" }, 405);
  let body;
  try { body = await req.json(); } catch { return json({ error: "Invalid JSON body" }, 400); }

  const { action, key, value, prefix } = body || {};
  const store = getStore("cesim-dashboard");

  try {
    if (action === "get") {
      if (!key) return json({ error: "Missing key" }, 400);
      const v = await store.get(key);
      return json(v === null ? null : { key, value: v });
    }
    if (action === "set") {
      if (!key) return json({ error: "Missing key" }, 400);
      await store.set(key, value ?? "");
      return json({ key, value });
    }
    if (action === "delete") {
      if (!key) return json({ error: "Missing key" }, 400);
      await store.delete(key);
      return json({ key, deleted: true });
    }
    if (action === "list") {
      const { blobs } = await store.list({ prefix: prefix || "" });
      return json({ keys: blobs.map((b) => b.key) });
    }
    return json({ error: `Unknown action: ${action}` }, 400);
  } catch (e) {
    return json({ error: String(e?.message || e) }, 500);
  }
};
