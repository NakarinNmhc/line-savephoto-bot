const CLIENT_ID = String(process.env.MS_CLIENT_ID || "").trim();
const TENANT = "consumers"; // personal
const SCOPE = "offline_access User.Read Files.ReadWrite";

async function postForm(url, data) {
  const body = new URLSearchParams(data);
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });
  const text = await res.text();
  let json = {};
  try { json = text ? JSON.parse(text) : {}; } catch { json = { raw: text }; }
  if (!res.ok) throw new Error(JSON.stringify(json, null, 2));
  return json;
}

(async () => {
  if (!CLIENT_ID) throw new Error("Missing MS_CLIENT_ID");

  const deviceCodeUrl = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/devicecode`;
  const tokenUrl = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;

  const dc = await postForm(deviceCodeUrl, { client_id: CLIENT_ID, scope: SCOPE });
  console.log(dc.message);

  const intervalMs = (dc.interval || 5) * 1000;
  const expiresAt = Date.now() + (dc.expires_in || 900) * 1000;

  while (Date.now() < expiresAt) {
    await new Promise(r => setTimeout(r, intervalMs));
    try {
      const tok = await postForm(tokenUrl, {
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: CLIENT_ID,
        device_code: dc.device_code,
      });
      console.log("\nMS_REFRESH_TOKEN:\n");
      console.log(tok.refresh_token);
      process.exit(0);
    } catch (e) {
      const msg = String(e.message || e);
      if (msg.includes("authorization_pending") || msg.includes("slow_down")) continue;
      console.error(msg);
      process.exit(1);
    }
  }
  console.error("Device code expired");
  process.exit(1);
})();