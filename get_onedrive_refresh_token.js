/**
 * get_onedrive_refresh_token.js (ORG / Tenant)
 *
 * Usage (Windows CMD):
 *   set MS_TENANT=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
 *   set MS_CLIENT_ID=xxxxxxxx-xxxx-xxxx-xxxx-xxxxxxxxxxxx
 *   set MS_SCOPES=offline_access User.Read Files.ReadWrite.All
 *   node get_onedrive_refresh_token.js
 *
 * Notes:
 * - ต้องล็อกอินด้วย "เมลบริษัท" เท่านั้น
 * - ถ้า admin consent ถูกให้แล้ว จะไม่เด้งขอสิทธิ์ซ้ำมาก
 */

const MS_TENANT = process.env.MS_TENANT;
const MS_CLIENT_ID = process.env.MS_CLIENT_ID;
const MS_SCOPES = process.env.MS_SCOPES || "offline_access User.Read Files.ReadWrite.All";

if (!MS_TENANT || !MS_CLIENT_ID) {
  console.error("❌ Missing env: MS_TENANT or MS_CLIENT_ID");
  process.exit(1);
}

const DEVICE_CODE_URL = `https://login.microsoftonline.com/${MS_TENANT}/oauth2/v2.0/devicecode`;
const TOKEN_URL = `https://login.microsoftonline.com/${MS_TENANT}/oauth2/v2.0/token`;

async function postForm(url, data) {
  const body = new URLSearchParams(data);
  const res = await fetch(url, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const text = await res.text();
  let json = {};
  try {
    json = text ? JSON.parse(text) : {};
  } catch {
    json = { raw: text };
  }

  if (!res.ok) {
    const err = new Error(`HTTP ${res.status}: ${JSON.stringify(json)}`);
    err.payload = json;
    throw err;
  }
  return json;
}

async function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}

(async () => {
  console.log("Tenant:", MS_TENANT);
  console.log("Client ID:", MS_CLIENT_ID);
  console.log("Scopes:", MS_SCOPES);
  console.log("Device Code endpoint:", DEVICE_CODE_URL);
  console.log("Token endpoint:", TOKEN_URL);

  // Step 1: get device code
  const dc = await postForm(DEVICE_CODE_URL, {
    client_id: MS_CLIENT_ID,
    scope: MS_SCOPES,
  });

  console.log("\n=== Step 1: User action required ===");
  console.log("Open:", dc.verification_uri || "https://microsoft.com/devicelogin");
  console.log("Enter code:", dc.user_code);
  console.log("Message:", dc.message);
  console.log("Expires in (sec):", dc.expires_in);
  console.log("Polling interval (sec):", dc.interval);

  // Step 2: poll token
  console.log("\n=== Step 2: Waiting for login/consent... ===");
  const started = Date.now();
  const expiresMs = (dc.expires_in || 900) * 1000;
  let intervalMs = (dc.interval || 5) * 1000;

  while (Date.now() - started < expiresMs) {
    try {
      const tok = await postForm(TOKEN_URL, {
        client_id: MS_CLIENT_ID,
        grant_type: "urn:ietf:params:oauth:grant-type:device_code",
        device_code: dc.device_code,
      });

      console.log("\n✅ SUCCESS!");
      console.log("\nMS_ACCESS_TOKEN (short):", String(tok.access_token || "").slice(0, 40) + "...");
      console.log("\nMS_REFRESH_TOKEN (COPY THIS):\n");
      console.log(tok.refresh_token);
      console.log("\n⚠️ Put MS_REFRESH_TOKEN into Render ENV (do NOT share it).");
      process.exit(0);
    } catch (e) {
      const p = e.payload || {};
      const code = p.error;

      // pending / slow_down = ปกติ
      if (code === "authorization_pending") {
        await sleep(intervalMs);
        continue;
      }
      if (code === "slow_down") {
        intervalMs += 2000;
        await sleep(intervalMs);
        continue;
      }

      // อื่น ๆ คือ error จริง
      console.error("\n❌ ERROR:", e.message);
      console.error("Payload:", p);
      process.exit(1);
    }
  }

  console.error("\n❌ TIMEOUT: device code expired. Run again.");
  process.exit(1);
})();