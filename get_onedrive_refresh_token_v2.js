// get_onedrive_refresh_token_v2.js
// Device Code Flow -> get refresh_token
// IMPORTANT: set TENANT below. Start with "organizations" (recommended for work/school accounts)

const TENANT = String(process.env.MS_TENANT || "organizations").trim(); 
const CLIENT_ID = String(process.env.MS_CLIENT_ID || "1e214704-8aa3-4cf8-9024-1da284185fe1").trim();

const SCOPE = "offline_access User.Read Files.ReadWrite.All";

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
  if (!res.ok) {
    console.error("\n--- REQUEST DEBUG ---");
    console.error("URL:", url);
    console.error("BODY:", Object.fromEntries(body.entries()));
    console.error("--- RESPONSE ---");
    throw new Error(JSON.stringify(json, null, 2));
  }
  return json;
}

(async () => {
  if (!CLIENT_ID) {
    console.error("Missing CLIENT_ID");
    process.exit(1);
  }

  const deviceCodeUrl = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/devicecode`;
  const tokenUrl = `https://login.microsoftonline.com/${TENANT}/oauth2/v2.0/token`;

  console.log("TENANT =", TENANT);
  console.log("CLIENT_ID =", CLIENT_ID);
  console.log("deviceCodeUrl =", deviceCodeUrl);

  const dc = await postForm(deviceCodeUrl, {
    client_id: CLIENT_ID,
    scope: SCOPE,
  });

  console.log("\n=== LOGIN ===");
  console.log(dc.message);

  const intervalMs = (dc.interval || 5) * 1000;
  const expiresAt = Date.now() + (dc.expires_in || 900) * 1000;

  while (Date.now() < expiresAt) {
    await new Promise((r) => setTimeout(r, intervalMs));
    try {
      const tok = await postForm(tokenUrl, {
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        client_id: CLIENT_ID,
        device_code: dc.device_code,
      });

      console.log("\n✅ SUCCESS");
      console.log("\nMS_REFRESH_TOKEN (copy this to Render ENV):\n");
      console.log(tok.refresh_token);
      process.exit(0);
    } catch (e) {
      const msg = String(e.message || e);
      if (msg.includes("authorization_pending")) continue;
      if (msg.includes("slow_down")) continue;
      console.error("\n❌ ERROR\n", msg);
      process.exit(1);
    }
  }

  console.error("\n❌ Device code expired. Run again.");
  process.exit(1);
})();
