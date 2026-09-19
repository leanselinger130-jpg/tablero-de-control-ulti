const ENDPOINT = "/.netlify/functions/storage";

async function call(action, payload = {}) {
  const res = await fetch(ENDPOINT, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ action, ...payload }),
  });
  if (!res.ok) throw new Error(`Storage error (${action}): ${res.status}`);
  return res.json();
}

export const storage = {
  async get(key) { return call("get", { key }); },
  async set(key, value) { return call("set", { key, value }); },
  async delete(key) { return call("delete", { key }); },
  async list(prefix) { return call("list", { prefix }); },
};
