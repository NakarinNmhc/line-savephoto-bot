/**
 * SavePhotoBot - server.js (Production-ready) ✅ SavePhotoBotUser (User-centric folders, NO day folder)
 * Goals:
 * ✅ Production logging template (requestId, event summary, timing, structured error)
 * ✅ Stable webhook (no dropped events, per-event isolation, dedupe, safe reply/push)
 * ✅ Render sleep mitigation (keepalive ping + external ping suggestion)
 * ✅ Support image + video + file
 * ✅ Split folders by type: .../<source>/<sender>/<images|videos|files>/<file>
 * ✅ Video size limit (skip if too large)
 * ✅ File size limit (skip if too large)
 *
 * ✅ STRUCTURE (NO day):
 *   Root: ONEDRIVE_BASE_PATH=SavePhotoBotUser (default)
 *   Path: <root>/<sourceFolder>/<senderFolder>/<images|videos|files>/<file>
 *
 * Notes:
 * - LINE middleware requires raw body verification; keep `line.middleware(config)` as-is.
 * - Reply HTTP 200 ASAP to avoid LINE retry storms.
 */

require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");
const crypto = require("crypto");

const app = express();

/* -------------------- ENV -------------------- */
const LINE_ACCESS_TOKEN = process.env.LINE_ACCESS_TOKEN;
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const ADMIN_USER_ID = process.env.ADMIN_USER_ID; // can be Uxxx or Cxxx or Rxxx

// Optional: protect /images viewing route
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

// ✅ Default root changed to SavePhotoBotUser
const ONEDRIVE_BASE_PATH = process.env.ONEDRIVE_BASE_PATH || "SavePhotoBotUser";

// SharePoint target
const SP_HOSTNAME = process.env.SP_HOSTNAME || "";
const SP_SITE_PATH = process.env.SP_SITE_PATH || ""; // e.g. /sites/SavePhotoBot
const SP_DRIVE_NAME = process.env.SP_DRIVE_NAME || "Documents";
const SP_DRIVE_ID = process.env.SP_DRIVE_ID || "";

// Optional: delete local file after upload
const DELETE_LOCAL_AFTER_UPLOAD =
  (process.env.DELETE_LOCAL_AFTER_UPLOAD || "1") === "1";

// Reliability knobs
const UPLOAD_CONCURRENCY = Math.max(
  1,
  Number(process.env.UPLOAD_CONCURRENCY || 2)
);
const GRAPH_TIMEOUT_MS = Math.max(
  10_000,
  Number(process.env.GRAPH_TIMEOUT_MS || 90_000)
);
const GRAPH_RETRY_MAX = Math.max(0, Number(process.env.GRAPH_RETRY_MAX || 4));
const GRAPH_RETRY_BASE_MS = Math.max(
  200,
  Number(process.env.GRAPH_RETRY_BASE_MS || 600)
);

// Keep-alive for Render
const PUBLIC_BASE_URL = (process.env.PUBLIC_BASE_URL || "").trim(); // e.g. https://xxx.onrender.com
const KEEPALIVE_ENABLED = (process.env.KEEPALIVE_ENABLED || "1") === "1";
const KEEPALIVE_INTERVAL_MS = Math.max(
  60_000,
  Number(process.env.KEEPALIVE_INTERVAL_MS || 300_000)
);

// Media controls
const ALLOW_VIDEO = (process.env.ALLOW_VIDEO || "1") === "1";
const ALLOW_FILE = (process.env.ALLOW_FILE || "1") === "1";

const MAX_VIDEO_MB = Math.max(1, Number(process.env.MAX_VIDEO_MB || 30));
const MAX_VIDEO_BYTES = MAX_VIDEO_MB * 1024 * 1024;

const MAX_FILE_MB = Math.max(1, Number(process.env.MAX_FILE_MB || 20));
const MAX_FILE_BYTES = MAX_FILE_MB * 1024 * 1024;

/* -------------------- Validate env -------------------- */
function looksLikeLineId(id) {
  // LINE IDs typically start with U (user), C (group), R (room)
  return typeof id === "string" && /^[UCR]/.test(id) && id.length >= 10;
}

if (!LINE_ACCESS_TOKEN || !LINE_CHANNEL_SECRET) {
  console.error("❌ Missing env: LINE_ACCESS_TOKEN or LINE_CHANNEL_SECRET");
  process.exit(1);
}

if (!ADMIN_USER_ID || !looksLikeLineId(ADMIN_USER_ID)) {
  console.error(
    "❌ Missing/Invalid env: ADMIN_USER_ID (ต้องเป็น userId(U...) หรือ groupId(C...) หรือ roomId(R...))"
  );
  process.exit(1);
}

if (!MS_TENANT || !MS_CLIENT_ID || !MS_REFRESH_TOKEN) {
  console.error("❌ Missing env: MS_TENANT or MS_CLIENT_ID or MS_REFRESH_TOKEN");
  process.exit(1);
}

if (STORAGE_MODE === "sharepoint") {
  if (!SP_DRIVE_ID && (!SP_HOSTNAME || !SP_SITE_PATH)) {
    console.error(
      "❌ Missing env for SharePoint: SP_HOSTNAME/SP_SITE_PATH (or set SP_DRIVE_ID)"
    );
    process.exit(1);
  }
}

/* -------------------- LINE client -------------------- */
const config = {
  channelAccessToken: LINE_ACCESS_TOKEN,
  channelSecret: LINE_CHANNEL_SECRET,
};
const client = new line.Client(config);

/* -------------------- Basic health routes -------------------- */
app.get("/", (req, res) => res.status(200).send("OK"));
app.get("/health", (req, res) => res.status(200).send("OK"));

/* -------------------- Storage (local temp) -------------------- */
const baseImagesDir = path.join(__dirname, "images");
if (!fs.existsSync(baseImagesDir))
  fs.mkdirSync(baseImagesDir, { recursive: true });

/* -------------------- Static route (Express 5 safe) -------------------- */
app.get(/^\/images\/.*/, (req, res, next) => {
  if (!IMAGE_VIEW_TOKEN) return next();
  if (req.query.token !== IMAGE_VIEW_TOKEN) return res.sendStatus(403);
  return next();
});
app.use("/images", express.static(baseImagesDir));

/* -------------------- Production Logger -------------------- */
function nowISO() {
  return new Date().toISOString();
}
function rid() {
  return crypto.randomBytes(6).toString("hex"); // short request id
}
function safeJson(v) {
  try {
    return JSON.stringify(v);
  } catch {
    return String(v);
  }
}
function log(level, msg, meta = {}) {
  const line = {
    ts: nowISO(),
    level,
    msg,
    ...meta,
  };
  console.log(safeJson(line)); // single-line JSON
}
function msSince(t0) {
  return Date.now() - t0;
}

/* -------------------- Helpers -------------------- */
function pad(n) {
  return String(n).padStart(2, "0");
}
function makeFileNamePrefix(messageId) {
  const d = new Date();
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(
    d.getDate()
  )}_${messageId}`;
}
function makeFileName(messageId, ext = "jpg") {
  return `${makeFileNamePrefix(messageId)}.${ext}`;
}
function sanitizeFolderName(name) {
  return String(name || "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
}
function sanitizeFileName(name) {
  const base = String(name || "")
    .replace(/[/\\]/g, "_")
    .replace(/[<>:"|?*\x00-\x1F]/g, "_")
    .replace(/\s+/g, " ")
    .trim();
  return base ? base.slice(0, 120) : "file";
}
function getExtFromFileName(name) {
  const n = String(name || "").trim();
  const idx = n.lastIndexOf(".");
  if (idx <= 0 || idx === n.length - 1) return "";
  return n
    .slice(idx + 1)
    .toLowerCase()
    .replace(/[^a-z0-9]/g, "")
    .slice(0, 10);
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
function saveStreamToFileWithLimit(stream, filePath, maxBytes) {
  return new Promise((resolve, reject) => {
    let written = 0;
    const w = fs.createWriteStream(filePath);

    const cleanup = () => {
      try {
        w.destroy();
      } catch {}
      try {
        stream.destroy();
      } catch {}
      try {
        if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
      } catch {}
    };

    stream.on("data", (chunk) => {
      written += chunk.length;
      if (written > maxBytes) {
        cleanup();
        reject(new Error("TOO_LARGE"));
      }
    });

    w.on("finish", resolve);
    w.on("error", (e) => {
      cleanup();
      reject(e);
    });
    stream.on("error", (e) => {
      cleanup();
      reject(e);
    });

    stream.pipe(w);
  });
}
function buildPublicBaseUrl(req) {
  if (PUBLIC_BASE_URL) return PUBLIC_BASE_URL;
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
function encodeGraphPath(p) {
  return String(p || "")
    .split("/")
    .filter(Boolean)
    .map(encodeURIComponent)
    .join("/");
}
function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}
function isTransientStatus(status) {
  return status === 408 || status === 429 || (status >= 500 && status <= 599);
}
function extFromContentType(ct) {
  const c = String(ct || "").toLowerCase();
  if (c.includes("image/png")) return "png";
  if (c.includes("image/webp")) return "webp";
  if (c.includes("image/jpeg") || c.includes("image/jpg")) return "jpg";

  if (c.includes("video/mp4")) return "mp4";
  if (c.includes("video/quicktime")) return "mov";
  if (c.includes("video/3gpp")) return "3gp";

  if (c.includes("application/pdf")) return "pdf";
  if (c.includes("application/zip")) return "zip";
  if (c.includes("text/plain")) return "txt";

  return "bin";
}
function mimeFromExt(ext) {
  const e = String(ext || "").toLowerCase();
  if (e === "png") return "image/png";
  if (e === "webp") return "image/webp";
  if (e === "jpg" || e === "jpeg") return "image/jpeg";

  if (e === "mp4") return "video/mp4";
  if (e === "mov") return "video/quicktime";
  if (e === "3gp") return "video/3gpp";

  if (e === "pdf") return "application/pdf";
  if (e === "zip") return "application/zip";
  if (e === "txt") return "text/plain; charset=utf-8";

  return "application/octet-stream";
}
function typeSubFolder(messageType) {
  if (messageType === "image") return "images";
  if (messageType === "video") return "videos";
  return "files";
}

/* -------------------- Simple concurrency limiter -------------------- */
function createLimiter(max) {
  let active = 0;
  const queue = [];

  const next = () => {
    if (active >= max) return;
    const job = queue.shift();
    if (!job) return;

    active++;
    Promise.resolve()
      .then(job.fn)
      .then(job.resolve, job.reject)
      .finally(() => {
        active--;
        next();
      });
  };

  return (fn) =>
    new Promise((resolve, reject) => {
      queue.push({ fn, resolve, reject });
      next();
    });
}
const uploadLimiter = createLimiter(UPLOAD_CONCURRENCY);

/* -------------------- Notify admin (push) with safety -------------------- */
async function notifyAdmin(text, meta = {}) {
  const msg = String(text || "").slice(0, 4900);
  try {
    await client.pushMessage(ADMIN_USER_ID, [{ type: "text", text: msg }]);
    log("INFO", "ADMIN_NOTIFY_OK", meta);
  } catch (e) {
    log("ERROR", "ADMIN_NOTIFY_FAIL", {
      ...meta,
      err: String(e?.message || e),
      hint:
        "ถ้าเจอ to invalid ให้เช็คว่า ADMIN_USER_ID เป็น U.../C.../R... ถูกต้อง และไม่มีช่องว่าง",
    });
  }
}

/* -------------------- Cache: group name -------------------- */
const nameCache = new Map();
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
    } catch {
      return `group_${tail}`;
    }
  }

  if (s.type === "room" && s.roomId) {
    const tail = s.roomId.slice(-6);
    return `room_${tail}`;
  }

  return "unknown";
}

/* -------------------- Cache: sender folder (User-centric) -------------------- */
const senderCache = new Map(); // key: `${srcType}:${scopeId}:${userId}`
const SENDER_CACHE_TTL_MS = 24 * 60 * 60 * 1000;

function buildSenderFolder(userId, displayName) {
  const tail = String(userId || "unknown").slice(-6);
  const name = sanitizeFolderName(displayName || "");
  return name ? `user_${name}_${tail}` : `user_${tail}`;
}

async function getSenderFolder(event) {
  const s = event.source || {};
  const userId = s.userId;
  if (!userId) return "user_unknown";

  const scopeId = s.groupId || s.roomId || "private";
  const cacheKey = `${s.type}:${scopeId}:${userId}`;

  const cached = senderCache.get(cacheKey);
  if (cached && Date.now() - cached.ts < SENDER_CACHE_TTL_MS) return cached.folder;

  try {
    let profile = null;

    if (s.type === "group" && s.groupId) {
      profile = await client.getGroupMemberProfile(s.groupId, userId);
    } else if (s.type === "room" && s.roomId) {
      profile = await client.getRoomMemberProfile(s.roomId, userId);
    } else {
      profile = await client.getProfile(userId);
    }

    const folder = buildSenderFolder(userId, profile?.displayName || "");
    senderCache.set(cacheKey, { folder, ts: Date.now() });
    return folder;
  } catch {
    const folder = buildSenderFolder(userId, "");
    senderCache.set(cacheKey, { folder, ts: Date.now() });
    return folder;
  }
}

/* -------------------- Dedupe (webhook retry) -------------------- */
const seenMessageIds = new Set();
function rememberMessageId(id) {
  seenMessageIds.add(id);
  setTimeout(() => seenMessageIds.delete(id), 10 * 60 * 1000).unref?.();
}

/* -------------------- Microsoft Graph OAuth -------------------- */
let currentRefreshToken = MS_REFRESH_TOKEN;

async function msPostForm(url, data) {
  const body = new URLSearchParams(data);

  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), GRAPH_TIMEOUT_MS);

  try {
    const res = await fetch(url, {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body,
      signal: controller.signal,
    });

    const text = await res.text();
    let json = {};
    try {
      json = text ? JSON.parse(text) : {};
    } catch {
      json = { raw: text };
    }
    if (!res.ok) throw new Error(`MS OAuth error: ${safeJson(json)}`);
    return json;
  } finally {
    clearTimeout(t);
  }
}

async function getGraphAccessToken() {
  const tokenUrl = `https://login.microsoftonline.com/${MS_TENANT}/oauth2/v2.0/token`;

  const tok = await msPostForm(tokenUrl, {
    client_id: MS_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: currentRefreshToken,
    scope: MS_SCOPES,
  });

  if (tok.refresh_token && tok.refresh_token !== currentRefreshToken) {
    currentRefreshToken = tok.refresh_token;
    log("WARN", "MS_REFRESH_TOKEN_ROTATED", {
      hint: "Update MS_REFRESH_TOKEN on Render ASAP (token มีการหมุน)",
    });
  }

  return tok.access_token;
}

/* -------------------- fetch with timeout + retry -------------------- */
async function fetchWithTimeout(url, options = {}, timeoutMs = GRAPH_TIMEOUT_MS) {
  const controller = new AbortController();
  const t = setTimeout(() => controller.abort(), timeoutMs);

  try {
    const res = await fetch(url, { ...options, signal: controller.signal });
    return res;
  } finally {
    clearTimeout(t);
  }
}

async function graphFetch(
  url,
  { accessToken, method = "GET", headers = {}, body, timeoutMs } = {}
) {
  const res = await fetchWithTimeout(
    url,
    {
      method,
      headers: { Authorization: `Bearer ${accessToken}`, ...headers },
      body,
    },
    timeoutMs || GRAPH_TIMEOUT_MS
  );

  const text = await res.text();
  let json = null;
  try {
    json = text ? JSON.parse(text) : null;
  } catch {
    json = null;
  }
  return { res, text, json };
}

async function graphFetchRetry(url, opts, { max = GRAPH_RETRY_MAX } = {}) {
  let attempt = 0;
  let lastErr = null;

  while (attempt <= max) {
    try {
      const out = await graphFetch(url, opts);
      if (out.res.ok) return out;

      if (isTransientStatus(out.res.status)) {
        const wait = GRAPH_RETRY_BASE_MS * Math.pow(2, attempt);
        log("WARN", "GRAPH_TRANSIENT_RETRY", {
          status: out.res.status,
          waitMs: wait,
          url,
          attempt,
        });
        await sleep(wait);
        attempt++;
        continue;
      }

      return out; // non-transient
    } catch (e) {
      lastErr = e;
      const wait = GRAPH_RETRY_BASE_MS * Math.pow(2, attempt);
      log("WARN", "GRAPH_EXCEPTION_RETRY", {
        waitMs: wait,
        attempt,
        err: String(e?.message || e),
      });
      await sleep(wait);
      attempt++;
    }
  }

  throw lastErr || new Error("Graph retry failed");
}

/* -------------------- Target Drive (SharePoint or OneDrive) -------------------- */
let cachedDriveBase = null;
let cachedDriveBaseTs = 0;
const DRIVE_CACHE_TTL_MS = 6 * 60 * 60 * 1000;

async function getDriveBase(accessToken) {
  if (cachedDriveBase && Date.now() - cachedDriveBaseTs < DRIVE_CACHE_TTL_MS) {
    return cachedDriveBase;
  }

  if (STORAGE_MODE === "onedrive") {
    cachedDriveBase = "https://graph.microsoft.com/v1.0/me/drive";
    cachedDriveBaseTs = Date.now();
    return cachedDriveBase;
  }

  if (SP_DRIVE_ID) {
    cachedDriveBase = `https://graph.microsoft.com/v1.0/drives/${SP_DRIVE_ID}`;
    cachedDriveBaseTs = Date.now();
    return cachedDriveBase;
  }

  const siteUrl = `https://graph.microsoft.com/v1.0/sites/${SP_HOSTNAME}:${SP_SITE_PATH}`;
  const site = await graphFetchRetry(siteUrl, { accessToken });
  if (!site.res.ok) throw new Error(`Get site failed: ${site.text}`);
  const siteId = site.json?.id;
  if (!siteId) throw new Error("Get site: missing siteId");

  const drivesUrl = `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`;
  const drives = await graphFetchRetry(drivesUrl, { accessToken });
  if (!drives.res.ok) throw new Error(`List drives failed: ${drives.text}`);

  const list = drives.json?.value || [];
  const wantNames = [
    (SP_DRIVE_NAME || "").toLowerCase(),
    "shared documents",
    "documents",
  ].filter(Boolean);

  const found =
    list.find((d) => wantNames.includes(String(d.name || "").toLowerCase())) ||
    list.find((d) => String(d.name || "").toLowerCase() === "documents") ||
    list[0];

  if (!found?.id) throw new Error("No drive found in this site");

  cachedDriveBase = `https://graph.microsoft.com/v1.0/drives/${found.id}`;
  cachedDriveBaseTs = Date.now();
  return cachedDriveBase;
}

/* -------------------- Drive folder + upload -------------------- */
async function ensureDriveFolder(accessToken, driveBase, folderPath) {
  const parts = String(folderPath).split("/").filter(Boolean);
  let current = "";

  for (const p of parts) {
    const next = current ? `${current}/${p}` : p;

    const checkUrl = `${driveBase}/root:/${encodeGraphPath(next)}`;
    const check = await graphFetchRetry(checkUrl, { accessToken });
    if (check.res.ok) {
      current = next;
      continue;
    }

    const parent = current;
    const createUrl = parent
      ? `${driveBase}/root:/${encodeGraphPath(parent)}:/children`
      : `${driveBase}/root/children`;

    const create = await graphFetchRetry(
      createUrl,
      {
        accessToken,
        method: "POST",
        headers: { "Content-Type": "application/json" },
        body: JSON.stringify({
          name: p,
          folder: {},
          "@microsoft.graph.conflictBehavior": "fail",
        }),
      },
      { max: 2 }
    );

    if (!create.res.ok) {
      const msg = create.text || "";
      if (!msg.includes("nameAlreadyExists")) {
        throw new Error(`Create folder failed (${next}): ${create.text}`);
      }
    }

    current = next;
  }
}

async function uploadSmall(
  accessToken,
  driveBase,
  drivePath,
  buffer,
  contentType
) {
  const url = `${driveBase}/root:/${encodeGraphPath(drivePath)}:/content`;

  const out = await graphFetchRetry(url, {
    accessToken,
    method: "PUT",
    headers: { "Content-Type": contentType || "application/octet-stream" },
    body: buffer,
    timeoutMs: Math.max(GRAPH_TIMEOUT_MS, 120_000),
  });

  if (!out.res.ok) throw new Error(`Upload small failed: ${out.text}`);
  return out.json || {};
}

async function uploadLarge(accessToken, driveBase, drivePath, buffer) {
  const createUrl = `${driveBase}/root:/${encodeGraphPath(
    drivePath
  )}:/createUploadSession`;
  const created = await graphFetchRetry(createUrl, {
    accessToken,
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      item: { "@microsoft.graph.conflictBehavior": "replace" },
    }),
  });

  if (!created.res.ok)
    throw new Error(`CreateUploadSession failed: ${created.text}`);
  const uploadUrl = created.json?.uploadUrl;
  if (!uploadUrl) throw new Error("CreateUploadSession: missing uploadUrl");

  const chunkSize = 5 * 1024 * 1024;
  const total = buffer.length;

  let start = 0;
  while (start < total) {
    const end = Math.min(start + chunkSize, total);
    const chunk = buffer.slice(start, end);

    let attempt = 0;
    while (true) {
      try {
        const res = await fetchWithTimeout(
          uploadUrl,
          {
            method: "PUT",
            headers: {
              "Content-Length": String(chunk.length),
              "Content-Range": `bytes ${start}-${end - 1}/${total}`,
            },
            body: chunk,
          },
          Math.max(GRAPH_TIMEOUT_MS, 120_000)
        );

        if (res.status === 200 || res.status === 201 || res.status === 202)
          break;

        const txt = await res.text();
        if (isTransientStatus(res.status) && attempt < GRAPH_RETRY_MAX) {
          const wait = GRAPH_RETRY_BASE_MS * Math.pow(2, attempt);
          log("WARN", "GRAPH_CHUNK_RETRY", {
            status: res.status,
            waitMs: wait,
            range: `${start}-${end - 1}/${total}`,
          });
          await sleep(wait);
          attempt++;
          continue;
        }

        throw new Error(
          `Chunk upload failed (${start}-${end - 1}) status=${res.status}: ${txt}`
        );
      } catch (e) {
        if (attempt < GRAPH_RETRY_MAX) {
          const wait = GRAPH_RETRY_BASE_MS * Math.pow(2, attempt);
          log("WARN", "GRAPH_CHUNK_EXCEPTION_RETRY", {
            waitMs: wait,
            range: `${start}-${end - 1}/${total}`,
            err: String(e?.message || e),
          });
          await sleep(wait);
          attempt++;
          continue;
        }
        throw e;
      }
    }

    start = end;
  }

  const itemUrl = `${driveBase}/root:/${encodeGraphPath(drivePath)}`;
  const item = await graphFetchRetry(itemUrl, { accessToken }, { max: 2 });

  if (!item.res.ok) {
    log("WARN", "GRAPH_ITEM_FETCH_AFTER_UPLOAD_FAIL", {
      hint: "ถือว่าอัปโหลดสำเร็จได้ (บางที webUrl fetch ล้มเหลว)",
      text: item.text,
    });
    return { id: null, name: path.basename(drivePath), webUrl: null };
  }
  return item.json || {};
}

async function uploadToDrive({
  folderName,
  fileName,
  localFilePath,
  contentType,
}) {
  const accessToken = await getGraphAccessToken();
  const driveBase = await getDriveBase(accessToken);

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

/* -------------------- Debug route -------------------- */
app.get("/debug/sharepoint", async (req, res) => {
  try {
    const accessToken = await getGraphAccessToken();
    const driveBase = await getDriveBase(accessToken);

    let drivesList = null;
    if (STORAGE_MODE === "sharepoint" && !SP_DRIVE_ID) {
      const siteUrl = `https://graph.microsoft.com/v1.0/sites/${SP_HOSTNAME}:${SP_SITE_PATH}`;
      const site = await graphFetchRetry(siteUrl, { accessToken });
      const siteId = site.json?.id;
      if (siteId) {
        const drives = await graphFetchRetry(
          `https://graph.microsoft.com/v1.0/sites/${siteId}/drives`,
          { accessToken }
        );
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
      UPLOAD_CONCURRENCY,
      GRAPH_TIMEOUT_MS,
      GRAPH_RETRY_MAX,
      GRAPH_RETRY_BASE_MS,
      PUBLIC_BASE_URL,
      KEEPALIVE_ENABLED,
      KEEPALIVE_INTERVAL_MS,
      ALLOW_VIDEO,
      ALLOW_FILE,
      MAX_VIDEO_MB,
      MAX_FILE_MB,
    });
  } catch (e) {
    res.status(500).json({ ok: false, error: String(e?.message || e) });
  }
});

/* -------------------- Safe LINE send helpers -------------------- */
async function safeReply(replyToken, messages, meta = {}) {
  if (!replyToken) return;
  try {
    await client.replyMessage(replyToken, messages);
    log("INFO", "LINE_REPLY_OK", meta);
  } catch (e) {
    log("ERROR", "LINE_REPLY_FAIL", { ...meta, err: String(e?.message || e) });
  }
}

/* -------------------- Webhook (stable) -------------------- */
app.post("/webhook", line.middleware(config), async (req, res) => {
  const requestId = rid();
  const t0 = Date.now();

  // Reply fast to prevent LINE retry storms
  res.sendStatus(200);

  const events = req.body?.events || [];
  const baseUrl = buildPublicBaseUrl(req);

  log("INFO", "WEBHOOK_RECEIVED", {
    requestId,
    events: events.length,
    baseUrl,
  });

  for (const event of events) {
    const evT0 = Date.now();

    const srcType = event?.source?.type;
    const groupId = event?.source?.groupId;
    const roomId = event?.source?.roomId;
    const userId = event?.source?.userId;

    const evMeta = {
      requestId,
      eventType: event?.type,
      messageType: event?.message?.type,
      srcType,
      groupTail: groupId ? groupId.slice(-6) : null,
      roomTail: roomId ? roomId.slice(-6) : null,
      userTail: userId ? userId.slice(-6) : null,
    };

    try {
      log("DEBUG", "EVENT_IN", {
        ...evMeta,
        source: event.source,
      });

      // Silent policy for "join"
      if (event.type === "join") {
        log("INFO", "EVENT_JOIN_IGNORED", evMeta);
        continue;
      }

      // follow in private chat
      if (event.type === "follow") {
        if (srcType === "user" && event.replyToken) {
          await safeReply(
            event.replyToken,
            [{ type: "text", text: "สวัสดีครับ 🙂 SavePhotoBot พร้อมรับไฟล์แล้ว" }],
            evMeta
          );
        }
        log("INFO", "EVENT_FOLLOW_HANDLED", { ...evMeta, ms: msSince(evT0) });
        continue;
      }

      if (event.type !== "message") {
        log("INFO", "EVENT_SKIPPED_NOT_MESSAGE", {
          ...evMeta,
          ms: msSince(evT0),
        });
        continue;
      }

      const mtype = event.message?.type; // image | video | file | text | ...
      if (mtype === "video" && !ALLOW_VIDEO) {
        log("INFO", "EVENT_SKIPPED_VIDEO_DISABLED", {
          ...evMeta,
          ms: msSince(evT0),
        });
        continue;
      }
      if (mtype === "file" && !ALLOW_FILE) {
        log("INFO", "EVENT_SKIPPED_FILE_DISABLED", {
          ...evMeta,
          ms: msSince(evT0),
        });
        continue;
      }
      if (!["image", "video", "file"].includes(mtype)) {
        log("INFO", "EVENT_SKIPPED_UNSUPPORTED_TYPE", {
          ...evMeta,
          ms: msSince(evT0),
        });
        continue;
      }

      const messageId = event.message.id;

      // dedupe by messageId
      if (seenMessageIds.has(messageId)) {
        log("WARN", "DEDUPLICATE_IGNORED", { ...evMeta, messageId });
        continue;
      }
      rememberMessageId(messageId);

      // ✅ NO day folder
      const folderName = await getSourceFolder(event);     // group_xxx / room_xxx / private
      const senderFolder = await getSenderFolder(event);   // user_<name>_<tail> or user_<tail>
      const sub = typeSubFolder(mtype);

      const targetDir = path.join(baseImagesDir, folderName, senderFolder, sub);
      if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });

      // Fetch content stream
      const stream = await client.getMessageContent(messageId);
      const ct = (stream?.headers?.["content-type"] || "").toLowerCase();

      // decide filename/ext
      let ext = extFromContentType(ct);
      let fileName = "";

      if (mtype === "file") {
        const original = sanitizeFileName(
          event.message.fileName || `file_${messageId}`
        );
        const fromNameExt = getExtFromFileName(original);
        if (fromNameExt) ext = fromNameExt;

        fileName = `${makeFileNamePrefix(messageId)}_${original}`;
      } else {
        if (mtype === "video" && (ext === "bin" || !ext)) ext = "mp4";
        if (!ext) ext = "bin";
        fileName = makeFileName(messageId, ext);
      }

      // avoid too-long file names for SharePoint
      if (fileName.length > 160) {
        const keepExt = getExtFromFileName(fileName) || ext;
        fileName = `${makeFileNamePrefix(messageId)}.${keepExt}`;
      }

      const filePath = path.join(targetDir, fileName);

      // Save local with limits
      if (mtype === "video") {
        try {
          await saveStreamToFileWithLimit(stream, filePath, MAX_VIDEO_BYTES);
        } catch (e) {
          if (String(e?.message || e).includes("TOO_LARGE")) {
            log("WARN", "VIDEO_TOO_LARGE_SKIPPED", {
              ...evMeta,
              messageId,
              maxMB: MAX_VIDEO_MB,
            });

            await notifyAdmin(
              `🚫 ข้ามวิดีโอ (ใหญ่เกิน ${MAX_VIDEO_MB}MB)\n` +
                `ที่: ${sourceLabel(event)}\n` +
                `โฟลเดอร์: ${folderName}\n` +
                `ผู้ส่ง: ${senderFolder}\n` +
                `ชนิด: ${sub}\n` +
                `messageId: ${messageId}`,
              { ...evMeta, messageId }
            );
            continue;
          }
          throw e;
        }
      } else if (mtype === "file") {
        try {
          await saveStreamToFileWithLimit(stream, filePath, MAX_FILE_BYTES);
        } catch (e) {
          if (String(e?.message || e).includes("TOO_LARGE")) {
            log("WARN", "FILE_TOO_LARGE_SKIPPED", {
              ...evMeta,
              messageId,
              maxMB: MAX_FILE_MB,
              fileName,
            });

            await notifyAdmin(
              `🚫 ข้ามไฟล์ (ใหญ่เกิน ${MAX_FILE_MB}MB)\n` +
                `ที่: ${sourceLabel(event)}\n` +
                `โฟลเดอร์: ${folderName}\n` +
                `ผู้ส่ง: ${senderFolder}\n` +
                `ชนิด: ${sub}\n` +
                `ไฟล์: ${fileName}\n` +
                `messageId: ${messageId}`,
              { ...evMeta, messageId }
            );
            continue;
          }
          throw e;
        }
      } else {
        await saveStreamToFile(stream, filePath);
      }

      log("INFO", "SAVED_LOCAL", {
        ...evMeta,
        messageId,
        filePath,
        contentType: ct,
        senderFolder,
        ms: msSince(evT0),
      });

      // Local view URL (only works if you keep local files; you delete after upload by default)
      const viewPath =
        `/images/${encodeURIComponent(folderName)}` +
        `/${encodeURIComponent(senderFolder)}` +
        `/${encodeURIComponent(sub)}` +
        `/${encodeURIComponent(fileName)}`;

      const localViewUrl = IMAGE_VIEW_TOKEN
        ? `${baseUrl}${viewPath}?token=${encodeURIComponent(IMAGE_VIEW_TOKEN)}`
        : `${baseUrl}${viewPath}`;

      const contentType = mimeFromExt(ext);

      // Upload with concurrency limit
      const up = await uploadLimiter(() =>
        uploadToDrive({
          folderName: `${folderName}/${senderFolder}/${sub}`,
          fileName,
          localFilePath: filePath,
          contentType,
        })
      );

      log("INFO", "UPLOADED_DRIVE", {
        ...evMeta,
        drivePath: up.drivePath,
        webUrl: up.webUrl,
        size: up.size,
        senderFolder,
        ms: msSince(evT0),
      });

      const kindLabel =
        mtype === "image" ? "📸 รูป" : mtype === "video" ? "🎬 วิดีโอ" : "📎 ไฟล์";

      // Notify admin (silent in group/room)
      if (srcType === "group" || srcType === "room") {
        const msg =
          `${kindLabel} ใหม่ถูกส่งเข้ามา\n` +
          `ที่: ${sourceLabel(event)}\n` +
          `โฟลเดอร์: ${folderName}\n` +
          `ผู้ส่ง: ${senderFolder}\n` +
          `ชนิด: ${sub}\n` +
          `ไฟล์: ${fileName}\n` +
          `SharePoint: ${up.webUrl || "(ลิงก์อาจยังไม่พร้อม แต่ไฟล์อัปโหลดแล้ว)"}\n` +
          `Local: ${localViewUrl}`;

        await notifyAdmin(msg, { ...evMeta, messageId });
      } else if (srcType === "user" && event.replyToken) {
        await safeReply(
          event.replyToken,
          [
            {
              type: "text",
              text:
                `✅ บันทึก${
                  mtype === "image" ? "รูป" : mtype === "video" ? "วิดีโอ" : "ไฟล์"
                }แล้วครับ ` + `(อัปโหลดขึ้น ${STORAGE_MODE} แล้ว)`,
            },
          ],
          { ...evMeta, messageId }
        );
      }

      // Cleanup local
      if (DELETE_LOCAL_AFTER_UPLOAD) {
        try {
          fs.unlinkSync(filePath);

          // remove empty sub/sender/source dirs (best-effort)
          try {
            if (fs.existsSync(targetDir) && fs.readdirSync(targetDir).length === 0)
              fs.rmdirSync(targetDir);
          } catch {}

          const senderDir = path.join(baseImagesDir, folderName, senderFolder);
          try {
            if (fs.existsSync(senderDir) && fs.readdirSync(senderDir).length === 0)
              fs.rmdirSync(senderDir);
          } catch {}

          const parentDir = path.join(baseImagesDir, folderName);
          try {
            if (fs.existsSync(parentDir) && fs.readdirSync(parentDir).length === 0)
              fs.rmdirSync(parentDir);
          } catch {}

          log("INFO", "LOCAL_CLEANUP_OK", { ...evMeta, messageId, senderFolder });
        } catch (e) {
          log("WARN", "LOCAL_CLEANUP_FAIL", {
            ...evMeta,
            messageId,
            senderFolder,
            err: String(e?.message || e),
          });
        }
      }

      log("INFO", "EVENT_DONE", {
        ...evMeta,
        messageId,
        senderFolder,
        ms: msSince(evT0),
      });
    } catch (err) {
      const msg = String(err?.message || err);

      log("ERROR", "EVENT_FAIL", {
        ...evMeta,
        err: msg,
        ms: msSince(evT0),
      });

      await notifyAdmin(
        `❌ SavePhotoBot Error\n` +
          `req=${requestId}\n` +
          `event=${event?.type}/${event?.message?.type || "-"}\n` +
          `src=${srcType}\n` +
          `err=${msg.slice(0, 1200)}`,
        evMeta
      );

      continue;
    }
  }

  log("INFO", "WEBHOOK_DONE", { requestId, ms: msSince(t0) });
});

/* -------------------- Keepalive (Render sleep mitigation) -------------------- */
async function keepalivePing() {
  if (!KEEPALIVE_ENABLED) return;
  if (!PUBLIC_BASE_URL) {
    log("WARN", "KEEPALIVE_SKIPPED_NO_PUBLIC_BASE_URL", {
      hint: "ตั้ง PUBLIC_BASE_URL=https://...onrender.com เพื่อให้ keepalive ทำงาน",
    });
    return;
  }

  const url = `${PUBLIC_BASE_URL.replace(/\/$/, "")}/health`;

  try {
    const res = await fetchWithTimeout(url, { method: "GET" }, 15_000);
    log("INFO", "KEEPALIVE_PING", { url, status: res.status });
  } catch (e) {
    log("WARN", "KEEPALIVE_FAIL", { url, err: String(e?.message || e) });
  }
}

if (KEEPALIVE_ENABLED) {
  setInterval(() => keepalivePing(), KEEPALIVE_INTERVAL_MS).unref?.();
  setTimeout(() => keepalivePing(), 10_000).unref?.();

  log("INFO", "KEEPALIVE_ENABLED", {
    PUBLIC_BASE_URL,
    KEEPALIVE_INTERVAL_MS,
    note:
      "บน Render Free: แนะนำใช้ UptimeRobot/cron ภายนอกยิง /health ทุก 5 นาที จะกัน sleep ได้ชัวร์กว่า",
  });
}

/* -------------------- Global crash guards -------------------- */
process.on("unhandledRejection", (reason) => {
  log("ERROR", "UNHANDLED_REJECTION", { reason: String(reason) });
});
process.on("uncaughtException", (err) => {
  log("ERROR", "UNCAUGHT_EXCEPTION", { err: String(err?.message || err) });
});

/* -------------------- Start -------------------- */
const PORT = Number(process.env.PORT || 3001);
app.listen(PORT, () => {
  log("INFO", "SERVER_STARTED", {
    port: PORT,
    STORAGE_MODE,
    ONEDRIVE_BASE_PATH,
    UPLOAD_CONCURRENCY,
    DELETE_LOCAL_AFTER_UPLOAD,
    ALLOW_VIDEO,
    ALLOW_FILE,
    MAX_VIDEO_MB,
    MAX_FILE_MB,
    structure: "<root>/<source>/<sender>/<type>/file",
  });
});