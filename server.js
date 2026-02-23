/**
 * SavePhotoBot - server.js (Render-ready + OneDrive ORG + Date folders + Logs)
 * Requirements:
 * - Silent in group/room (NO replyMessage, NO pushMessage to group/room)
 * - Save images locally into /images/<source-folder>/<YYYY-MM-DD> (optional debug)
 * - Upload images to OneDrive (Microsoft Graph) using ORG tenant refresh token
 * - Notify ADMIN_USER_ID (DM) always when image arrives from group/room (with OneDrive link)
 * - Optional static route /images for viewing with IMAGE_VIEW_TOKEN
 *
 * ENV required:
 *   LINE_ACCESS_TOKEN
 *   LINE_CHANNEL_SECRET
 *   ADMIN_USER_ID
 *
 * OneDrive ORG:
 *   MS_TENANT                 (Directory/Tenant ID)
 *   MS_CLIENT_ID              (Application/Client ID)
 *   MS_REFRESH_TOKEN          (Refresh token from device code flow AFTER admin consent)
 *   MS_SCOPES                 (recommended: "offline_access User.Read Files.ReadWrite.All")
 *   ONEDRIVE_BASE_PATH        (default: "SavePhotoBot")
 *
 * Optional:
 *   IMAGE_VIEW_TOKEN
 *   DELETE_LOCAL_AFTER_UPLOAD (default "1")
 */

require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");

const app = express();

/* -------------------- Basic health routes -------------------- */
app.get("/", (req, res) => res.status(200).send("OK"));
app.get("/health", (req, res) => res.status(200).send("OK"));

/* -------------------- ENV -------------------- */
const LINE_ACCESS_TOKEN = process.env.LINE_ACCESS_TOKEN;
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const ADMIN_USER_ID = process.env.ADMIN_USER_ID;

// Optional: protect image viewing route
const IMAGE_VIEW_TOKEN = process.env.IMAGE_VIEW_TOKEN || "";

// OneDrive (Microsoft Graph) - ORG tenant
const MS_TENANT = process.env.MS_TENANT; // Directory (tenant) ID
const MS_CLIENT_ID = process.env.MS_CLIENT_ID; // Application (client) ID
const MS_REFRESH_TOKEN = process.env.MS_REFRESH_TOKEN; // refresh token (org user)
const MS_SCOPES =
  process.env.MS_SCOPES || "offline_access User.Read Files.ReadWrite.All";

const ONEDRIVE_BASE_PATH = process.env.ONEDRIVE_BASE_PATH || "SavePhotoBot";

// Optional: delete local file after upload
const DELETE_LOCAL_AFTER_UPLOAD = (process.env.DELETE_LOCAL_AFTER_UPLOAD || "1") === "1";

/* -------------------- Validate env -------------------- */
if (!LINE_ACCESS_TOKEN || !LINE_CHANNEL_SECRET) {
  console.error("❌ Missing env: LINE_ACCESS_TOKEN or LINE_CHANNEL_SECRET");
  process.exit(1);
}
if (!ADMIN_USER_ID) {
  console.error("❌ Missing env: ADMIN_USER_ID");
  process.exit(1);
}
if (!MS_TENANT || !MS_CLIENT_ID || !MS_REFRESH_TOKEN) {
  console.error("❌ Missing env: MS_TENANT or MS_CLIENT_ID or MS_REFRESH_TOKEN");
  process.exit(1);
}

/* -------------------- LINE client -------------------- */
const config = {
  channelAccessToken: LINE_ACCESS_TOKEN,
  channelSecret: LINE_CHANNEL_SECRET,
};
const client = new line.Client(config);

/* -------------------- Storage (local temp) -------------------- */
const baseImagesDir = path.join(__dirname, "images");
if (!fs.existsSync(baseImagesDir)) fs.mkdirSync(baseImagesDir, { recursive: true });

/* -------------------- Static route (Express 5 safe) -------------------- */
app.get(/^\/images\/.*/, (req, res, next) => {
  if (!IMAGE_VIEW_TOKEN) return next(); // open if token not set (dev)
  if (req.query.token !== IMAGE_VIEW_TOKEN) return res.sendStatus(403);
  return next();
});
app.use("/images", express.static(baseImagesDir));

/* -------------------- Helpers -------------------- */
function pad(n) {
  return String(n).padStart(2, "0");
}

// Date folder: YYYY-MM-DD (server timezone)
function dateFolder() {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

// Filename: date + messageId
function makeFileName(messageId, ext = "jpg") {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}_${messageId}.${ext}`;
}

function sanitizeFolderName(name) {
  return String(name || "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
}

function saveStreamToFile(stream, filePath) {
  return new Promise((resolve, reject) => {
    const w = fs.createWriteStream(filePath);
    stream.pipe(w);
    w.on("finish", resolve);
    w.on("error", reject);
    stream.on("error", reject);
  });
}

function buildPublicBaseUrl(req) {
  const proto = req.headers["x-forwarded-proto"] || "https";
  const host = req.headers["x-forwarded-host"] || req.headers.host;
  return `${proto}://${host}`;
}

function sourceLabel(event) {
  const s = event.source || {};
  if (s.type === "group") return `GROUP (${(s.groupId || "").slice(-6)})`;
  if (s.type === "room") return `ROOM (${(s.roomId || "").slice(-6)})`;
  if (s.type === "user") return `PRIVATE (${(s.userId || "").slice(-6)})`;
  return "UNKNOWN";
}

/**
 * Graph path must not encode "/" as %2F
 * So encode each segment only.
 */
function encodeGraphPath(p) {
  return String(p || "")
    .split("/")
    .filter(Boolean)
    .map(encodeURIComponent)
    .join("/");
}

/* -------------------- Cache: group name -------------------- */
const nameCache = new Map(); // groupId -> { name, ts }
const CACHE_TTL_MS = 24 * 60 * 60 * 1000;

async function getGroupName(groupId) {
  const cached = nameCache.get(groupId);
  if (cached && Date.now() - cached.ts < CACHE_TTL_MS) return cached.name;

  const summary = await client.getGroupSummary(groupId);
  const name = sanitizeFolderName(summary.groupName || "UnknownGroup");
  nameCache.set(groupId, { name, ts: Date.now() });
  return name;
}

/* -------------------- Source folder -------------------- */
async function getSourceFolder(event) {
  const s = event.source || {};

  if (s.type === "user") return "private";

  if (s.type === "group" && s.groupId) {
    const tail = s.groupId.slice(-6);
    try {
      const name = await getGroupName(s.groupId);
      return `group_${name}_${tail}`;
    } catch (_) {
      return `group_${tail}`;
    }
  }

  if (s.type === "room" && s.roomId) {
    const tail = s.roomId.slice(-6);
    return `room_${tail}`;
  }

  return "unknown";
}

/* -------------------- Dedupe (webhook retry) -------------------- */
const seenMessageIds = new Set();
function rememberMessageId(id) {
  seenMessageIds.add(id);
  setTimeout(() => seenMessageIds.delete(id), 10 * 60 * 1000).unref?.();
}

/* -------------------- Notify admin (DM) -------------------- */
async function notifyAdmin(text) {
  return client.pushMessage(ADMIN_USER_ID, [{ type: "text", text }]);
}

/* -------------------- OneDrive / Microsoft Graph (ORG) -------------------- */
let currentRefreshToken = MS_REFRESH_TOKEN;

async function msPostForm(url, data) {
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
  if (!res.ok) throw new Error(`MS OAuth error: ${JSON.stringify(json)}`);
  return json;
}

async function getGraphAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${MS_TENANT}/oauth2/v2.0/token`;

  const tok = await msPostForm(tokenUrl, {
    client_id: MS_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: currentRefreshToken,
    scope: MS_SCOPES, // ✅ must match the scopes used to obtain refresh token
  });

  // MS may rotate refresh token
  if (tok.refresh_token && tok.refresh_token !== currentRefreshToken) {
    currentRefreshToken = tok.refresh_token;
    console.warn("⚠️ Microsoft rotated refresh token. Update MS_REFRESH_TOKEN on Render ASAP.");
  }

  return tok.access_token;
}

async function graphFetch(url, { accessToken, method = "GET", headers = {}, body } = {}) {
  const res = await fetch(url, {
    method,
    headers: {
      Authorization: `Bearer ${accessToken}`,
      ...headers,
    },
    body,
  });

  const text = await res.text();
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = null;
  }

  return { res, text, json };
}

// Ensure folder exists by creating each segment
async function ensureOneDriveFolder(accessToken, folderPath) {
  const parts = String(folderPath).split("/").filter(Boolean);
  let current = "";

  for (const p of parts) {
    const next = current ? `${current}/${p}` : p;

    // Check exists
    const checkUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeGraphPath(next)}`;
    const check = await graphFetch(checkUrl, { accessToken });

    if (check.res.ok) {
      current = next;
      continue;
    }

    // Create under parent
    const parent = current; // "" or "a/b"
    const createUrl = parent
      ? `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeGraphPath(parent)}:/children`
      : `https://graph.microsoft.com/v1.0/me/drive/root/children`;

    const create = await graphFetch(createUrl, {
      accessToken,
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        name: p,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      }),
    });

    // If already exists due to race, ignore
    if (!create.res.ok) {
      const msg = create.text || "";
      if (!msg.includes("nameAlreadyExists")) {
        throw new Error(`Create folder failed (${next}): ${create.text}`);
      }
    }

    current = next;
  }
}

// Small upload (<=4MB): PUT content
async function uploadSmall(accessToken, oneDrivePath, buffer, contentType) {
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeGraphPath(oneDrivePath)}:/content`;
  const { res, text, json } = await graphFetch(url, {
    accessToken,
    method: "PUT",
    headers: { "Content-Type": contentType || "application/octet-stream" },
    body: buffer,
  });

  if (!res.ok) throw new Error(`Upload small failed: ${text}`);
  return json || {};
}

// Large upload (Create upload session + chunk)
async function uploadLarge(accessToken, oneDrivePath, buffer) {
  const createUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeGraphPath(oneDrivePath)}:/createUploadSession`;
  const created = await graphFetch(createUrl, {
    accessToken,
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({ item: { "@microsoft.graph.conflictBehavior": "rename" } }),
  });

  if (!created.res.ok) throw new Error(`CreateUploadSession failed: ${created.text}`);
  const uploadUrl = created.json?.uploadUrl;
  if (!uploadUrl) throw new Error("CreateUploadSession: missing uploadUrl");

  // chunk size must be multiple of 320 KiB. Use 5 MiB.
  const chunkSize = 5 * 1024 * 1024;
  const total = buffer.length;

  let start = 0;
  while (start < total) {
    const end = Math.min(start + chunkSize, total);
    const chunk = buffer.slice(start, end);

    const res = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${start}-${end - 1}/${total}`,
      },
      body: chunk,
    });

    // 202/201 ok; 200 ok
    if (!(res.status === 200 || res.status === 201 || res.status === 202)) {
      const text = await res.text();
      throw new Error(`Chunk upload failed (${start}-${end - 1}): ${text}`);
    }

    start = end;
  }

  // Query final item by path to get webUrl
  const itemUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeGraphPath(oneDrivePath)}`;
  const item = await graphFetch(itemUrl, { accessToken });

  if (!item.res.ok) throw new Error(`Fetch uploaded item failed: ${item.text}`);
  return item.json || {};
}

async function uploadToOneDrive({ folderName, fileName, localFilePath, contentType }) {
  const accessToken = await getGraphAccessToken();

  // folderName can include subpaths (e.g. "group_xxx/2026-02-23")
  const oneDriveFolder = `${ONEDRIVE_BASE_PATH}/${folderName}`;
  await ensureOneDriveFolder(accessToken, oneDriveFolder);

  const oneDrivePath = `${oneDriveFolder}/${fileName}`;
  const buf = fs.readFileSync(localFilePath);

  const FOUR_MB = 4 * 1024 * 1024;
  const item =
    buf.length <= FOUR_MB
      ? await uploadSmall(accessToken, oneDrivePath, buf, contentType)
      : await uploadLarge(accessToken, oneDrivePath, buf);

  return {
    webUrl: item.webUrl || null,
    id: item.id || null,
    name: item.name || fileName,
    size: item.size || buf.length,
    oneDrivePath,
  };
}

/* -------------------- Webhook -------------------- */
app.post("/webhook", line.middleware(config), async (req, res) => {
  // Reply fast to prevent LINE retry
  res.sendStatus(200);

  const events = req.body?.events || [];
  const baseUrl = buildPublicBaseUrl(req);

  console.log("📩 Webhook triggered. Events:", events.length);

  for (const event of events) {
    try {
      const srcType = event.source?.type;
      console.log("➡️ Event:", {
        type: event.type,
        srcType,
        msgType: event.message?.type,
        messageId: event.message?.id,
      });

      // ---- Silent policy ----
      if (event.type === "join") continue;

      // follow happens in private (user add friend)
      if (event.type === "follow") {
        if (srcType === "user" && event.replyToken) {
          await client.replyMessage(event.replyToken, [
            { type: "text", text: "สวัสดีครับ 🙂 SavePhotoBot พร้อมรับรูปแล้ว" },
          ]);
        }
        continue;
      }

      // Only handle image messages
      if (event.type !== "message" || event.message?.type !== "image") continue;

      const messageId = event.message.id;

      // dedupe
      if (seenMessageIds.has(messageId)) {
        console.log("⚠️ Duplicate ignored:", messageId);
        continue;
      }
      rememberMessageId(messageId);

      const folderName = await getSourceFolder(event);
      const day = dateFolder();

      // local dir: /images/<folderName>/<YYYY-MM-DD>/
      const targetDir = path.join(baseImagesDir, folderName, day);
      if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });

      // get content stream
      const stream = await client.getMessageContent(messageId);

      // best-effort ext from content-type
      const ct = (stream?.headers?.["content-type"] || "").toLowerCase();
      const ext =
        ct.includes("png") ? "png" :
        ct.includes("jpeg") ? "jpg" :
        ct.includes("jpg") ? "jpg" :
        ct.includes("webp") ? "webp" :
        "jpg";

      const fileName = makeFileName(messageId, ext);
      const filePath = path.join(targetDir, fileName);

      await saveStreamToFile(stream, filePath);
      console.log("✅ Saved local:", filePath);

      // local view url (optional)
      const viewPath = `/images/${encodeURIComponent(folderName)}/${encodeURIComponent(day)}/${encodeURIComponent(fileName)}`;
      const localViewUrl = IMAGE_VIEW_TOKEN
        ? `${baseUrl}${viewPath}?token=${encodeURIComponent(IMAGE_VIEW_TOKEN)}`
        : `${baseUrl}${viewPath}`;

      // ---- Upload to OneDrive ----
      const contentType =
        ext === "png" ? "image/png" :
        ext === "webp" ? "image/webp" :
        "image/jpeg";

      // OneDrive path: SavePhotoBot/<folderName>/<YYYY-MM-DD>/<fileName>
      const up = await uploadToOneDrive({
        folderName: `${folderName}/${day}`,
        fileName,
        localFilePath: filePath,
        contentType,
      });

      console.log("☁️ Uploaded OneDrive:", up.oneDrivePath, up.webUrl || "(no webUrl)");

      // Notify admin always when image from group/room
      if (srcType === "group" || srcType === "room") {
        const msg =
          `📸 มีรูปถูกส่งเข้ามา\n` +
          `ที่: ${sourceLabel(event)}\n` +
          `ผู้ส่ง: ${event.source?.userId || "-"}\n` +
          `โฟลเดอร์: ${folderName}\n` +
          `วันที่: ${day}\n` +
          `ไฟล์: ${fileName}\n` +
          `OneDrive: ${up.webUrl || "(สร้างลิงก์ไม่ได้)"}\n` +
          `Local: ${localViewUrl}`;

        await notifyAdmin(msg);

        // IMPORTANT: Silent in group/room -> no reply, no push to group/room
      } else {
        // Private chat (optional)
        if (srcType === "user" && event.replyToken) {
          await client.replyMessage(event.replyToken, [
            { type: "text", text: "✅ บันทึกรูปแล้วครับ (อัปโหลดขึ้น OneDrive แล้ว)" },
          ]);
        }
      }

      // delete local file (optional)
      if (DELETE_LOCAL_AFTER_UPLOAD) {
        try {
          fs.unlinkSync(filePath);

          // remove empty day folder
          try {
            if (fs.existsSync(targetDir) && fs.readdirSync(targetDir).length === 0) fs.rmdirSync(targetDir);
          } catch {}

          // remove empty parent folder (/images/<folderName>)
          const parentDir = path.join(baseImagesDir, folderName);
          try {
            if (fs.existsSync(parentDir) && fs.readdirSync(parentDir).length === 0) fs.rmdirSync(parentDir);
          } catch {}
        } catch {}
      }
    } catch (err) {
      console.error("❌ Error:", err?.message || err);
      console.error("LINE API error body:", err?.originalError?.response?.data);

      // Best-effort admin notify
      try {
        await notifyAdmin(`❌ SavePhotoBot Error: ${String(err?.message || err).slice(0, 900)}`);
      } catch (_) {}
    }
  }
});

/* -------------------- Start -------------------- */
const PORT = Number(process.env.PORT || 3001);
app.listen(PORT, () => console.log(`🚀 SavePhotoBot running on port ${PORT}`));