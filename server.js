/**
 * SavePhotoBot - Production-ready (Render)
 * - Silent in group/room
 * - Save image -> OneDrive Personal (Microsoft Graph)
 * - Notify ADMIN (DM) always when image from group/room
 * - Queue + Retry + Backoff + Token cache
 * - DM commands: help, status, last, folders
 */

require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");

// Node 22 has global fetch
const app = express();
app.set("trust proxy", true);

/* -------------------- Basic routes (health for Render) -------------------- */
const STARTED_AT = Date.now();
app.get("/", (req, res) => res.status(200).send("OK"));
app.get("/health", (req, res) => res.status(200).json({ ok: true, uptimeSec: Math.floor((Date.now() - STARTED_AT) / 1000) }));
app.get("/healthz", (req, res) => res.status(200).send("OK"));

/* -------------------- ENV -------------------- */
const LINE_ACCESS_TOKEN = process.env.LINE_ACCESS_TOKEN;
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const ADMIN_USER_ID = process.env.ADMIN_USER_ID;

const MS_TENANT = process.env.MS_TENANT || "consumers"; // OneDrive Personal
const MS_CLIENT_ID = process.env.MS_CLIENT_ID;
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET || ""; // optional
const MS_REFRESH_TOKEN = process.env.MS_REFRESH_TOKEN;

const BOT_ROOT_FOLDER = process.env.BOT_ROOT_FOLDER || "SavePhotoBot";
const UPLOAD_CONCURRENCY = Math.max(1, Number(process.env.UPLOAD_CONCURRENCY || 2));
const UPLOAD_MAX_RETRY = Math.max(1, Number(process.env.UPLOAD_MAX_RETRY || 5));

// Optional local image viewer (dev only)
const IMAGE_VIEW_TOKEN = process.env.IMAGE_VIEW_TOKEN || "";

/* -------------------- Validate env -------------------- */
if (!LINE_ACCESS_TOKEN || !LINE_CHANNEL_SECRET) {
  console.error("❌ Missing env: LINE_ACCESS_TOKEN or LINE_CHANNEL_SECRET");
  process.exit(1);
}
if (!ADMIN_USER_ID) {
  console.error("❌ Missing env: ADMIN_USER_ID");
  process.exit(1);
}
if (!MS_CLIENT_ID || !MS_REFRESH_TOKEN) {
  console.error("❌ Missing env: MS_CLIENT_ID or MS_REFRESH_TOKEN (OneDrive Personal)");
  process.exit(1);
}

/* -------------------- LINE client -------------------- */
const config = {
  channelAccessToken: LINE_ACCESS_TOKEN,
  channelSecret: LINE_CHANNEL_SECRET,
};
const client = new line.Client(config);

/* -------------------- Local storage (optional) -------------------- */
const baseImagesDir = path.join(__dirname, "images");
if (!fs.existsSync(baseImagesDir)) fs.mkdirSync(baseImagesDir, { recursive: true });

// Express 5 safe static route (avoid "/images/*")
app.get(/^\/images\/.*/, (req, res, next) => {
  if (!IMAGE_VIEW_TOKEN) return next(); // dev open
  if (req.query.token !== IMAGE_VIEW_TOKEN) return res.sendStatus(403);
  return next();
});
app.use("/images", express.static(baseImagesDir));

/* -------------------- Helpers -------------------- */
function pad(n) {
  return String(n).padStart(2, "0");
}
function todayStr(d = new Date()) {
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}
function makeFileName(messageId, ext = "jpg") {
  // Requirement: date + messageId
  return `${todayStr()}_${messageId}.${ext}`;
}
function sanitizeFolderName(name) {
  return String(name || "")
    .replace(/[\\/:*?"<>|]/g, "_")
    .replace(/\s+/g, " ")
    .trim()
    .slice(0, 80);
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
function sleep(ms) {
  return new Promise((r) => setTimeout(r, ms));
}
function isRetryableStatus(status) {
  return status === 429 || (status >= 500 && status <= 599);
}

/* -------------------- Dedupe (LINE webhook retries) -------------------- */
const seenMessageIds = new Set();
function rememberMessageId(id) {
  seenMessageIds.add(id);
  setTimeout(() => seenMessageIds.delete(id), 10 * 60 * 1000).unref?.();
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

/* -------------------- Notify admin (DM) -------------------- */
async function notifyAdmin(text) {
  return client.pushMessage(ADMIN_USER_ID, [{ type: "text", text }]);
}

/* ============================================================
 *  OneDrive (Microsoft Graph) - Token + Folder + Upload
 * ============================================================ */

/** In-memory access token cache */
const msTokenCache = {
  accessToken: "",
  expiresAt: 0, // epoch ms
};

async function getMsAccessToken() {
  // reuse if still valid (buffer 60s)
  if (msTokenCache.accessToken && Date.now() < msTokenCache.expiresAt - 60_000) {
    return msTokenCache.accessToken;
  }

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(MS_TENANT)}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.set("client_id", MS_CLIENT_ID);
  if (MS_CLIENT_SECRET) body.set("client_secret", MS_CLIENT_SECRET);
  body.set("grant_type", "refresh_token");
  body.set("refresh_token", MS_REFRESH_TOKEN);

  // IMPORTANT: For v2, scope is recommended. Use Graph default + offline_access to keep refresh stable.
  // (Some setups work without scope; but this is safer)
  body.set("scope", "offline_access https://graph.microsoft.com/.default");

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "content-type": "application/x-www-form-urlencoded" },
    body,
  });

  const json = await res.json();
  if (!res.ok) {
    // don't log tokens
    throw new Error(`MS token error: ${json?.error || res.status} - ${json?.error_description || "unknown"}`);
  }

  msTokenCache.accessToken = json.access_token;
  msTokenCache.expiresAt = Date.now() + Number(json.expires_in || 3600) * 1000;
  return msTokenCache.accessToken;
}

async function graphFetch(url, options = {}, retryLeft = UPLOAD_MAX_RETRY) {
  const token = await getMsAccessToken();
  const headers = {
    ...(options.headers || {}),
    Authorization: `Bearer ${token}`,
  };

  const res = await fetch(url, { ...options, headers });

  if (res.status === 401 && retryLeft > 0) {
    // token expired or revoked -> clear cache and retry once
    msTokenCache.accessToken = "";
    msTokenCache.expiresAt = 0;
    return graphFetch(url, options, retryLeft - 1);
  }

  if (isRetryableStatus(res.status) && retryLeft > 0) {
    const wait = Math.min(30_000, 500 * Math.pow(2, (UPLOAD_MAX_RETRY - retryLeft))); // backoff
    await sleep(wait);
    return graphFetch(url, options, retryLeft - 1);
  }

  return res;
}

/** Ensure folder path exists under root: SavePhotoBot/<folderName>/<date> */
async function ensureOneDriveFolderPath(pathParts) {
  // pathParts is array like ["SavePhotoBot", folderName, "YYYY-MM-DD"]
  // We'll create step by step using children + conflictBehavior=fail
  // Use special root: /me/drive/root

  // Start from root item
  let parentId = null; // null = drive root

  for (const part of pathParts) {
    const name = part;

    // 1) Check if folder exists
    const encodedName = encodeURIComponent(name.replace(/'/g, "''"));
    const listUrl = parentId
      ? `https://graph.microsoft.com/v1.0/me/drive/items/${parentId}/children?$filter=name eq '${encodedName}'`
      : `https://graph.microsoft.com/v1.0/me/drive/root/children?$filter=name eq '${encodedName}'`;

    // NOTE: filter with encodedName is tricky; simplest: list children and find by name would be heavy.
    // Better approach: use path-based addressing:
    // /me/drive/root:/SavePhotoBot:/children and then try create; if exists -> 409 and we then resolve by GET item
    // We'll use path-based method for reliability:

    const folderPath = pathPartsToGraphPath(pathParts.slice(0, pathParts.indexOf(part) + 1));
    // Try GET item by path
    const getUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${folderPath}`;
    let getRes = await graphFetch(getUrl, { method: "GET" });

    if (getRes.ok) {
      const item = await getRes.json();
      parentId = item.id;
      continue;
    }

    // 2) Create folder under current parent
    const createUrl = parentId
      ? `https://graph.microsoft.com/v1.0/me/drive/items/${parentId}/children`
      : `https://graph.microsoft.com/v1.0/me/drive/root/children`;

    const createRes = await graphFetch(createUrl, {
      method: "POST",
      headers: { "content-type": "application/json" },
      body: JSON.stringify({
        name,
        folder: {},
        "@microsoft.graph.conflictBehavior": "fail",
      }),
    });

    if (createRes.ok) {
      const created = await createRes.json();
      parentId = created.id;
      continue;
    }

    // If exists (409), GET again
    if (createRes.status === 409) {
      getRes = await graphFetch(getUrl, { method: "GET" });
      if (!getRes.ok) {
        const t = await safeReadText(getRes);
        throw new Error(`Folder exists but cannot GET: ${getRes.status} ${t}`);
      }
      const item = await getRes.json();
      parentId = item.id;
      continue;
    }

    const txt = await safeReadText(createRes);
    throw new Error(`Create folder failed: ${createRes.status} ${txt}`);
  }

  return parentId;
}

function pathPartsToGraphPath(parts) {
  // join with / and encode each segment safely
  return parts.map((p) => encodeURIComponent(p)).join("/");
}

async function safeReadText(res) {
  try {
    return await res.text();
  } catch {
    return "";
  }
}

/** Upload small file (<4MB) using simple upload */
async function uploadSmallFileToOneDrive(folderParts, fileName, buffer) {
  const fullPath = pathPartsToGraphPath([...folderParts, fileName]);
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${fullPath}:/content`;

  const res = await graphFetch(url, {
    method: "PUT",
    headers: { "content-type": "application/octet-stream" },
    body: buffer,
  });

  if (!res.ok) {
    const txt = await safeReadText(res);
    throw new Error(`Simple upload failed: ${res.status} ${txt}`);
  }
  return res.json();
}

/** Upload large file using upload session (recommended for LINE images that might be >4MB) */
async function uploadLargeFileToOneDrive(folderParts, fileName, buffer) {
  const fullPath = pathPartsToGraphPath([...folderParts, fileName]);
  const createSessionUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${fullPath}:/createUploadSession`;

  const sessionRes = await graphFetch(createSessionUrl, {
    method: "POST",
    headers: { "content-type": "application/json" },
    body: JSON.stringify({
      item: {
        "@microsoft.graph.conflictBehavior": "replace",
        name: fileName,
      },
    }),
  });

  if (!sessionRes.ok) {
    const txt = await safeReadText(sessionRes);
    throw new Error(`Create upload session failed: ${sessionRes.status} ${txt}`);
  }

  const session = await sessionRes.json();
  const uploadUrl = session.uploadUrl;

  // Chunk upload
  const chunkSize = 5 * 1024 * 1024; // 5MB
  const total = buffer.length;

  let start = 0;
  while (start < total) {
    const end = Math.min(total - 1, start + chunkSize - 1);
    const chunk = buffer.subarray(start, end + 1);

    const putRes = await fetch(uploadUrl, {
      method: "PUT",
      headers: {
        "Content-Length": String(chunk.length),
        "Content-Range": `bytes ${start}-${end}/${total}`,
      },
      body: chunk,
    });

    // 202 accepted = continue
    if (putRes.status === 202) {
      start = end + 1;
      continue;
    }

    // 201/200 = finished
    if (putRes.ok) {
      return putRes.json();
    }

    // Retryable?
    if (isRetryableStatus(putRes.status)) {
      await sleep(1000);
      continue;
    }

    const txt = await safeReadText(putRes);
    throw new Error(`Chunk upload failed: ${putRes.status} ${txt}`);
  }

  throw new Error("Chunk upload ended unexpectedly");
}

/** Main upload: ensure folders then upload with correct method */
async function uploadToOneDrive({ folderName, fileName, buffer }) {
  const date = todayStr();
  const folderParts = [BOT_ROOT_FOLDER, folderName, date];

  // Ensure folder path exists
  await ensureOneDriveFolderPath(folderParts);

  // Choose upload strategy
  const FOUR_MB = 4 * 1024 * 1024;
  if (buffer.length <= FOUR_MB) {
    return uploadSmallFileToOneDrive(folderParts, fileName, buffer);
  }
  return uploadLargeFileToOneDrive(folderParts, fileName, buffer);
}

/** Build a OneDrive "open" link (webUrl) from upload response */
function extractWebUrl(uploadResp) {
  return uploadResp?.webUrl || uploadResp?.["@microsoft.graph.downloadUrl"] || "";
}

/* ============================================================
 *  Queue / Concurrency (ข้อ 2)
 * ============================================================ */
function createQueue(concurrency) {
  let running = 0;
  const queue = [];

  const runNext = () => {
    if (running >= concurrency) return;
    const job = queue.shift();
    if (!job) return;

    running++;
    job()
      .catch(() => {})
      .finally(() => {
        running--;
        runNext();
      });
  };

  return {
    add(fn) {
      return new Promise((resolve, reject) => {
        queue.push(async () => {
          try {
            const out = await fn();
            resolve(out);
          } catch (e) {
            reject(e);
          }
        });
        runNext();
      });
    },
    status() {
      return { running, queued: queue.length, concurrency };
    },
  };
}

const uploadQueue = createQueue(UPLOAD_CONCURRENCY);

/* ============================================================
 *  DM commands memory (ข้อ 5)
 * ============================================================ */
const lastUploads = []; // newest last
const LAST_LIMIT = 20;

function pushLastUpload(item) {
  lastUploads.push(item);
  while (lastUploads.length > LAST_LIMIT) lastUploads.shift();
}

/* -------------------- Webhook -------------------- */
app.post("/webhook", line.middleware(config), async (req, res) => {
  // Reply fast to prevent LINE retry storm
  res.sendStatus(200);

  const events = req.body?.events || [];
  const baseUrl = buildPublicBaseUrl(req);

  for (const event of events) {
    const srcType = event.source?.type;

    try {
      // ---- Silent policy ----
      // Never reply in group/room for any event
      if (event.type === "join") continue;

      // follow happens in private (user add friend)
      if (event.type === "follow") {
        if (srcType === "user" && event.replyToken) {
          await client.replyMessage(event.replyToken, [
            { type: "text", text: "สวัสดีครับ 🙂 SavePhotoBot พร้อมรับรูปแล้ว\nพิมพ์ help เพื่อดูคำสั่ง" },
          ]);
        }
        continue;
      }

      // DM commands (only in private user)
      if (event.type === "message" && srcType === "user" && event.message?.type === "text") {
        const text = String(event.message.text || "").trim().toLowerCase();

        if (!event.replyToken) continue;

        if (text === "help") {
          await client.replyMessage(event.replyToken, [{
            type: "text",
            text:
              "คำสั่ง SavePhotoBot:\n" +
              "- help : ดูคำสั่ง\n" +
              "- status : สถานะระบบ/คิวอัปโหลด\n" +
              "- last : ลิงก์รูปล่าสุด\n" +
              "- folders : รายชื่อโฟลเดอร์บน OneDrive (ระดับแรก)\n"
          }]);
          continue;
        }

        if (text === "status") {
          const q = uploadQueue.status();
          await client.replyMessage(event.replyToken, [{
            type: "text",
            text:
              `✅ Status\n` +
              `Uptime: ${Math.floor((Date.now() - STARTED_AT) / 1000)}s\n` +
              `Queue: running=${q.running}, queued=${q.queued}, concurrency=${q.concurrency}\n` +
              `OneDrive root: ${BOT_ROOT_FOLDER}\n`
          }]);
          continue;
        }

        if (text === "last") {
          if (lastUploads.length === 0) {
            await client.replyMessage(event.replyToken, [{ type: "text", text: "ยังไม่มีรายการอัปโหลดล่าสุดครับ" }]);
            continue;
          }
          const items = lastUploads.slice(-5).reverse();
          const msg = items.map((it, i) =>
            `${i + 1}) ${it.when}\nที่: ${it.source}\nโฟลเดอร์: ${it.folder}\nไฟล์: ${it.file}\nลิงก์: ${it.webUrl || "-"}`
          ).join("\n\n");
          await client.replyMessage(event.replyToken, [{ type: "text", text: msg }]);
          continue;
        }

        if (text === "folders") {
          // list folders under root folder on OneDrive
          // GET /me/drive/root:/SavePhotoBot:/children
          const rootPath = pathPartsToGraphPath([BOT_ROOT_FOLDER]);
          const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${rootPath}:/children?$top=50`;
          const r = await graphFetch(url, { method: "GET" });
          if (!r.ok) {
            const t = await safeReadText(r);
            await client.replyMessage(event.replyToken, [{ type: "text", text: `❌ list folders failed: ${r.status} ${t}` }]);
            continue;
          }
          const json = await r.json();
          const names = (json.value || [])
            .filter((x) => x.folder)
            .map((x) => x.name)
            .slice(0, 30);

          await client.replyMessage(event.replyToken, [{
            type: "text",
            text: names.length
              ? `📁 โฟลเดอร์ใน ${BOT_ROOT_FOLDER}:\n- ` + names.join("\n- ")
              : `ยังไม่มีโฟลเดอร์ใน ${BOT_ROOT_FOLDER} ครับ`
          }]);
          continue;
        }

        // unknown command
        await client.replyMessage(event.replyToken, [{ type: "text", text: "พิมพ์ help เพื่อดูคำสั่งที่ใช้ได้ครับ" }]);
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

      // Get content stream from LINE
      const stream = await client.getMessageContent(messageId);

      // Detect ext by content-type (best-effort)
      const ct = (stream?.headers?.["content-type"] || "").toLowerCase();
      const ext =
        ct.includes("png") ? "png" :
        ct.includes("jpeg") ? "jpg" :
        ct.includes("jpg") ? "jpg" :
        ct.includes("webp") ? "webp" :
        "jpg";

      const fileName = makeFileName(messageId, ext);

      // Read stream into Buffer (needed for Graph upload)
      const chunks = [];
      await new Promise((resolve, reject) => {
        stream.on("data", (c) => chunks.push(c));
        stream.on("end", resolve);
        stream.on("error", reject);
      });
      const buffer = Buffer.concat(chunks);

      // Optional local save (dev / debug)
      try {
        const targetDir = path.join(baseImagesDir, folderName);
        if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });
        fs.writeFileSync(path.join(targetDir, fileName), buffer);
      } catch (_) {}

      // Queue upload to OneDrive (ข้อ 2 + 3)
      const uploadResult = await uploadQueue.add(async () => {
        // retry wrapper (ข้อ 3) - network/Graph errors
        let lastErr = null;

        for (let attempt = 1; attempt <= UPLOAD_MAX_RETRY; attempt++) {
          try {
            const resp = await uploadToOneDrive({ folderName, fileName, buffer });
            return resp;
          } catch (e) {
            lastErr = e;
            const wait = Math.min(30_000, 800 * Math.pow(2, attempt - 1));
            console.warn(`⚠️ Upload retry ${attempt}/${UPLOAD_MAX_RETRY}: ${e?.message || e}`);
            await sleep(wait);
          }
        }

        throw lastErr || new Error("Upload failed");
      });

      const webUrl = extractWebUrl(uploadResult);

      pushLastUpload({
        when: new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" }),
        source: sourceLabel(event),
        folder: `${BOT_ROOT_FOLDER}/${folderName}/${todayStr()}`,
        file: fileName,
        webUrl,
      });

      // Notify admin always when image from group/room (ข้อ 5)
      if (srcType === "group" || srcType === "room") {
        const msg =
          `📸 มีรูปถูกส่งเข้ามา\n` +
          `ที่: ${sourceLabel(event)}\n` +
          `โฟลเดอร์: ${BOT_ROOT_FOLDER}/${folderName}/${todayStr()}\n` +
          `ไฟล์: ${fileName}\n` +
          `OneDrive: ${webUrl || "(กำลังได้ลิงก์ไม่ได้)"}\n` +
          `ขนาด: ${(buffer.length / 1024).toFixed(1)} KB`;

        await notifyAdmin(msg);

        // IMPORTANT: Silent in group/room -> no reply, no push to group/room
        continue;
      }

      // Private chat: optional confirm (คุณจะปิดก็ได้)
      if (srcType === "user" && event.replyToken) {
        await client.replyMessage(event.replyToken, [
          { type: "text", text: `✅ อัปโหลดเข้า OneDrive แล้ว\n${webUrl ? `ลิงก์: ${webUrl}` : ""}`.trim() },
        ]);
      }
    } catch (err) {
      console.error("❌ Error:", err?.message || err);
      console.error("LINE API error body:", err?.originalError?.response?.data);

      // notify admin on errors (best effort)
      try {
        await notifyAdmin(`❌ SavePhotoBot Error: ${String(err?.message || err)}`);
      } catch (_) {}
    }
  }
});

/* -------------------- Start -------------------- */
const PORT = Number(process.env.PORT || 3001);
app.listen(PORT, () => console.log(`🚀 SavePhotoBot running on port ${PORT}`));