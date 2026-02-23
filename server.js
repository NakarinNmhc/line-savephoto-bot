/**
 * SavePhotoBot - server.js (Render-ready + SharePoint/OneDrive + Date folders + Logs)
 *
 * Goals:
 * - Silent in group/room (NO replyMessage, NO pushMessage to group/room)
 * - Save images locally into /images/<source-folder>/<YYYY-MM-DD> (optional debug)
 * - Upload images to SharePoint (Site Documents) OR OneDrive via Microsoft Graph (ORG tenant refresh token)
 * - Notify ADMIN_USER_ID (DM) always when image arrives from group/room (with link)
 *
 * ENV required:
 *   LINE_ACCESS_TOKEN
 *   LINE_CHANNEL_SECRET
 *   ADMIN_USER_ID
 *
 * Microsoft Graph (ORG delegated):
 *   MS_TENANT
 *   MS_CLIENT_ID
 *   MS_REFRESH_TOKEN
 *   MS_SCOPES (recommended: "offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All")
 *
 * Storage mode:
 *   STORAGE_MODE=sharepoint | onedrive   (default: sharepoint)
 *
 * SharePoint target (for sharepoint mode):
 *   SP_HOSTNAME=milestonesth.sharepoint.com
 *   SP_SITE_PATH=/sites/SavePhotoBot
 *   SP_DRIVE_NAME=Shared Documents   (optional; try "Documents" too)
 *   SP_DRIVE_ID=xxxxx (optional; if set, skip discovery and use directly - most stable)
 *
 * Folder structure inside target drive:
 *   ONEDRIVE_BASE_PATH/<source>/<YYYY-MM-DD>/<file>
 *   (keep ONEDRIVE_BASE_PATH = "SavePhotoBot" by default)
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

// Microsoft Graph
const MS_TENANT = process.env.MS_TENANT;
const MS_CLIENT_ID = process.env.MS_CLIENT_ID;
const MS_REFRESH_TOKEN = process.env.MS_REFRESH_TOKEN;
const MS_SCOPES =
  process.env.MS_SCOPES ||
  "offline_access User.Read Files.ReadWrite.All Sites.ReadWrite.All";

// Storage selection
const STORAGE_MODE = (process.env.STORAGE_MODE || "sharepoint").toLowerCase(); // sharepoint | onedrive

// Base folder name inside target drive
const ONEDRIVE_BASE_PATH = process.env.ONEDRIVE_BASE_PATH || "SavePhotoBot";

// SharePoint target
const SP_HOSTNAME = process.env.SP_HOSTNAME || "";
const SP_SITE_PATH = process.env.SP_SITE_PATH || ""; // e.g. /sites/SavePhotoBot
const SP_DRIVE_NAME = process.env.SP_DRIVE_NAME || "Shared Documents";
const SP_DRIVE_ID = process.env.SP_DRIVE_ID || "";

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
if (STORAGE_MODE === "sharepoint") {
  // SP_DRIVE_ID is optional, but if not set we need hostname + site_path
  if (!SP_DRIVE_ID && (!SP_HOSTNAME || !SP_SITE_PATH)) {
    console.error("❌ Missing env for SharePoint: SP_HOSTNAME/SP_SITE_PATH (or set SP_DRIVE_ID)");
    process.exit(1);
  }
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
function dateFolder() {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
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

/* -------------------- Microsoft Graph OAuth (ORG) -------------------- */
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
    scope: MS_SCOPES,
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

/* -------------------- Target Drive (SharePoint or OneDrive) -------------------- */
let cachedDriveBase = null;
let cachedDriveBaseTs = 0;
const DRIVE_CACHE_TTL_MS = 6 * 60 * 60 * 1000; // 6 hours

async function getDriveBase(accessToken) {
  // cache
  if (cachedDriveBase && Date.now() - cachedDriveBaseTs < DRIVE_CACHE_TTL_MS) return cachedDriveBase;

  // OneDrive mode (My files)
  if (STORAGE_MODE === "onedrive") {
    cachedDriveBase = "https://graph.microsoft.com/v1.0/me/drive";
    cachedDriveBaseTs = Date.now();
    return cachedDriveBase;
  }

  // SharePoint mode
  if (SP_DRIVE_ID) {
    cachedDriveBase = `https://graph.microsoft.com/v1.0/drives/${SP_DRIVE_ID}`;
    cachedDriveBaseTs = Date.now();
    return cachedDriveBase;
  }

  // Discover siteId
  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${SP_HOSTNAME}:${SP_SITE_PATH}`;
  const site = await graphFetch(siteUrl, { accessToken });
  if (!site.res.ok) throw new Error(`Get site failed: ${site.text}`);
  const siteId = site.json?.id;
  if (!siteId) throw new Error("Get site: missing siteId");

  // List drives (document libraries)
  const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
  const drives = await graphFetch(drivesUrl, { accessToken });
  if (!drives.res.ok) throw new Error(`List drives failed: ${drives.text}`);

  const list = drives.json?.value || [];
  const wantNames = [
    (SP_DRIVE_NAME || "").toLowerCase(),
    "shared documents",
    "documents",
  ].filter(Boolean);

  const found =
    list.find(d => wantNames.includes(String(d.name || "").toLowerCase())) ||
    list.find(d => String(d.name || "").toLowerCase() === "documents") ||
    list[0];

  if (!found?.id) throw new Error("No drive found in this site");

  cachedDriveBase = `https://graph.microsoft.com/v1.0/drives/${found.id}`;
  cachedDriveBaseTs = Date.now();
  return cachedDriveBase;
}

/* -------------------- Drive folder + upload (works for OneDrive & SharePoint) -------------------- */
// Ensure folder exists by creating each segment
async function ensureDriveFolder(accessToken, driveBase, folderPath) {
  const parts = String(folderPath).split("/").filter(Boolean);
  let current = "";

  for (const p of parts) {
    const next = current ? `${current}/${p}` : p;

    // Check exists
    const checkUrl = `${driveBase}/root:/${encodeGraphPath(next)}`;
    const check = await graphFetch(checkUrl, { accessToken });

    if (check.res.ok) {
      current = next;
      continue;
    }

    // Create under parent
    const parent = current; // "" or "a/b"
    const createUrl = parent
      ? `${driveBase}/root:/${encodeGraphPath(parent)}:/children`
      : `${driveBase}/root/children`;

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
async function uploadSmall(accessToken, driveBase, drivePath, buffer, contentType) {
  const url = `${driveBase}/root:/${encodeGraphPath(drivePath)}:/content`;
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
async function uploadLarge(accessToken, driveBase, drivePath, buffer) {
  const createUrl = `${driveBase}/root:/${encodeGraphPath(drivePath)}:/createUploadSession`;
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

    if (!(res.status === 200 || res.status === 201 || res.status === 202)) {
      const text = await res.text();
      throw new Error(`Chunk upload failed (${start}-${end - 1}): ${text}`);
    }

    start = end;
  }

  // Query final item by path to get webUrl
  const itemUrl = `${driveBase}/root:/${encodeGraphPath(drivePath)}`;
  const item = await graphFetch(itemUrl, { accessToken });

  if (!item.res.ok) throw new Error(`Fetch uploaded item failed: ${item.text}`);
  return item.json || {};
}

async function uploadToDrive({ folderName, fileName, localFilePath, contentType }) {
  const accessToken = await getGraphAccessToken();
  const driveBase = await getDriveBase(accessToken);

  // Folder structure: ONEDRIVE_BASE_PATH/<folderName>
  const rootFolder = `${ONEDRIVE_BASE_PATH}/${folderName}`;
  await ensureDriveFolder(accessToken, driveBase, rootFolder);

  const drivePath = `${rootFolder}/${fileName}`;
  const buf = fs.readFileSync(localFilePath);

  const FOUR_MB = 4 * 1024 * 1024;
  const item =
    buf.length <= FOUR_MB
      ? await uploadSmall(accessToken, driveBase, drivePath, buf, contentType)
      : await uploadLarge(accessToken, driveBase, drivePath, buf);

  return {
    webUrl: item.webUrl || null,
    id: item.id || null,
    name: item.name || fileName,
    size: item.size || buf.length,
    drivePath,
    storage: STORAGE_MODE,
  };
}

/* -------------------- Debug route (optional) -------------------- */
app.get("/debug/sharepoint", async (req, res) => {
  try {
    const accessToken = await getGraphAccessToken();

    // show current target base & discovery info
    const driveBase = await getDriveBase(accessToken);

    // If sharepoint discovery, also list drives for your site
    let drivesList = null;
    if (STORAGE_MODE === "sharepoint" && !SP_DRIVE_ID) {
      const siteUrl = `https://graph.microsoft.com/v1.0/sites/${SP_HOSTNAME}:${SP_SITE_PATH}`;
      const site = await graphFetch(siteUrl, { accessToken });
      const siteId = site.json?.id;
      if (siteId) {
        const drives = await graphFetch(`https://graph.microsoft.com/v1.0/sites/${siteId}/drives`, { accessToken });
        drivesList = drives.json?.value || null;
      }
    }

    res.json({
      ok: true,
      STORAGE_MODE,
      ONEDRIVE_BASE_PATH,
      SP_HOSTNAME,
      SP_SITE_PATH,
      SP_DRIVE_NAME,
      SP_DRIVE_ID: SP_DRIVE_ID ? "(set)" : "(not set)",
      driveBase,
      drivesList,
    });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e?.message || e) });
  }
});

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

      const contentType =
        ext === "png" ? "image/png" :
        ext === "webp" ? "image/webp" :
        "image/jpeg";

      // Target path: ONEDRIVE_BASE_PATH/<folderName>/<YYYY-MM-DD>/<fileName>
      const up = await uploadToDrive({
        folderName: `${folderName}/${day}`,
        fileName,
        localFilePath: filePath,
        contentType,
      });

      console.log(`☁️ Uploaded (${up.storage}):`, up.drivePath, up.webUrl || "(no webUrl)");

      // Notify admin always when image from group/room
      if (srcType === "group" || srcType === "room") {
        const msg =
          `📸 มีรูปถูกส่งเข้ามา\n` +
          `ที่: ${sourceLabel(event)}\n` +
          `ผู้ส่ง: ${event.source?.userId || "-"}\n` +
          `โฟลเดอร์: ${folderName}\n` +
          `วันที่: ${day}\n` +
          `ไฟล์: ${fileName}\n` +
          `Link: ${up.webUrl || "(สร้างลิงก์ไม่ได้)"}\n` +
          `Local: ${localViewUrl}`;

        await notifyAdmin(msg);

        // IMPORTANT: Silent in group/room -> no reply, no push to group/room
      } else {
        // Private chat (optional)
        if (srcType === "user" && event.replyToken) {
          await client.replyMessage(event.replyToken, [
            { type: "text", text: `✅ บันทึกรูปแล้วครับ (อัปโหลดขึ้น ${STORAGE_MODE} แล้ว)` },
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