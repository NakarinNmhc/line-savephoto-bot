/**
 * SavePhotoBot - server.js (Render-ready) + OneDrive Personal upload
 * - Silent in group/room: NO replyMessage, NO pushMessage to group/room
 * - Upload images to OneDrive: /SavePhotoBot/<folderName>/<fileName>
 * - Always notify ADMIN_USER_ID (DM) when image arrives from group/room
 * - Logs every incoming event so you know webhook is working
 */

require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");

// Node 18+ has fetch built-in (Render uses Node 22 OK)
const app = express();

/* -------------------- Health routes -------------------- */
app.get("/", (req, res) => res.status(200).send("OK"));
app.get("/health", (req, res) => res.status(200).send("OK"));

/* -------------------- ENV -------------------- */
const LINE_ACCESS_TOKEN = process.env.LINE_ACCESS_TOKEN;
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const ADMIN_USER_ID = process.env.ADMIN_USER_ID;

const MS_CLIENT_ID = process.env.MS_CLIENT_ID;              // from Azure App Registration
const MS_CLIENT_SECRET = process.env.MS_CLIENT_SECRET || ""; // optional (public client may not have)
const MS_REFRESH_TOKEN = process.env.MS_REFRESH_TOKEN;      // your refresh token

// For OneDrive Personal use "consumers"
const MS_TENANT = process.env.MS_TENANT || "consumers"; // keep "consumers" for personal

/* -------------------- Validate env -------------------- */
function must(name, v) {
  if (!v) {
    console.error(`❌ Missing env: ${name}`);
    process.exit(1);
  }
}

must("LINE_ACCESS_TOKEN", LINE_ACCESS_TOKEN);
must("LINE_CHANNEL_SECRET", LINE_CHANNEL_SECRET);
must("ADMIN_USER_ID", ADMIN_USER_ID);

must("MS_CLIENT_ID", MS_CLIENT_ID);
must("MS_REFRESH_TOKEN", MS_REFRESH_TOKEN);

/* -------------------- LINE client -------------------- */
const config = {
  channelAccessToken: LINE_ACCESS_TOKEN,
  channelSecret: LINE_CHANNEL_SECRET,
};
const client = new line.Client(config);

/* -------------------- Local temp storage (optional) -------------------- */
const tmpDir = path.join(__dirname, "tmp");
if (!fs.existsSync(tmpDir)) fs.mkdirSync(tmpDir, { recursive: true });

/* -------------------- Helpers -------------------- */
function pad(n) {
  return String(n).padStart(2, "0");
}

// date + messageId (requirement)
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

function sourceLabel(event) {
  const s = event.source || {};
  if (s.type === "group") return `GROUP (${(s.groupId || "").slice(-6)})`;
  if (s.type === "room") return `ROOM (${(s.roomId || "").slice(-6)})`;
  if (s.type === "user") return `PRIVATE (${(s.userId || "").slice(-6)})`;
  return "UNKNOWN";
}

/* -------------------- Cache: group name -------------------- */
const nameCache = new Map(); // groupId -> { name, ts }
const CACHE_TTL_MS = 24 * 60 * 60 * 1000;

async function getGroupName(groupId) {
  const cached = nameCache.get(groupId);
  if (cached && Date.now() - cached.ts < CACHE_TTL_MS) return cached.name;

  const summary = await client.getGroupSummary(groupId); // {groupId, groupName, pictureUrl}
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

/* -------------------- Dedupe (LINE webhook retry) -------------------- */
const seenMessageIds = new Set();
function rememberMessageId(id) {
  seenMessageIds.add(id);
  setTimeout(() => seenMessageIds.delete(id), 10 * 60 * 1000).unref?.();
}

/* -------------------- Notify admin (DM) -------------------- */
async function notifyAdmin(text) {
  return client.pushMessage(ADMIN_USER_ID, [{ type: "text", text }]);
}

/* =========================================================
   OneDrive (Microsoft Graph) helpers
   ========================================================= */
let cachedAccessToken = "";
let cachedExpMs = 0;

async function getAccessToken() {
  // reuse until close to expire
  if (cachedAccessToken && Date.now() < cachedExpMs - 60_000) return cachedAccessToken;

  const tokenUrl = `https://login.microsoftonline.com/${encodeURIComponent(MS_TENANT)}/oauth2/v2.0/token`;

  const body = new URLSearchParams();
  body.set("client_id", MS_CLIENT_ID);
  if (MS_CLIENT_SECRET) body.set("client_secret", MS_CLIENT_SECRET);
  body.set("grant_type", "refresh_token");
  body.set("refresh_token", MS_REFRESH_TOKEN);
  body.set("scope", "offline_access Files.ReadWrite");

  const res = await fetch(tokenUrl, {
    method: "POST",
    headers: { "Content-Type": "application/x-www-form-urlencoded" },
    body,
  });

  const json = await res.json();
  if (!res.ok) {
    throw new Error(`Token error: ${JSON.stringify(json, null, 2)}`);
  }

  cachedAccessToken = json.access_token;
  cachedExpMs = Date.now() + (Number(json.expires_in || 3600) * 1000);
  return cachedAccessToken;
}

async function graphFetch(url, options = {}) {
  const token = await getAccessToken();
  const headers = {
    ...(options.headers || {}),
    Authorization: `Bearer ${token}`,
  };

  const res = await fetch(url, { ...options, headers });
  return res;
}

// Create folder if not exists under parentId
async function ensureChildFolder(parentId, folderName) {
  // 1) try list children & find folder
  const listUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${parentId}/children?$select=id,name,folder`;
  const listRes = await graphFetch(listUrl);
  const listJson = await listRes.json();
  if (!listRes.ok) throw new Error(`List children error: ${JSON.stringify(listJson, null, 2)}`);

  const found = (listJson.value || []).find(
    (it) => it?.folder && it?.name?.toLowerCase() === folderName.toLowerCase()
  );
  if (found?.id) return found.id;

  // 2) create folder
  const createUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${parentId}/children`;
  const createRes = await graphFetch(createUrl, {
    method: "POST",
    headers: { "Content-Type": "application/json" },
    body: JSON.stringify({
      name: folderName,
      folder: {},
      "@microsoft.graph.conflictBehavior": "rename",
    }),
  });
  const createJson = await createRes.json();
  if (!createRes.ok) throw new Error(`Create folder error: ${JSON.stringify(createJson, null, 2)}`);

  return createJson.id;
}

// Ensure full path /SavePhotoBot/<folderName> exists, return parent folder id
async function ensureUploadFolder(folderName) {
  // Root
  const rootRes = await graphFetch("https://graph.microsoft.com/v1.0/me/drive/root?$select=id");
  const rootJson = await rootRes.json();
  if (!rootRes.ok) throw new Error(`Root error: ${JSON.stringify(rootJson, null, 2)}`);
  const rootId = rootJson.id;

  // /SavePhotoBot
  const savePhotoBotId = await ensureChildFolder(rootId, "SavePhotoBot");

  // /SavePhotoBot/<folderName>
  const groupFolderId = await ensureChildFolder(savePhotoBotId, folderName);

  return groupFolderId;
}

// Upload file to OneDrive in given folderId
async function uploadFileToOneDrive(folderId, fileName, buffer, contentType) {
  // PUT /items/{folderId}:/{fileName}:/content
  const putUrl = `https://graph.microsoft.com/v1.0/me/drive/items/${folderId}:/${encodeURIComponent(fileName)}:/content`;

  const putRes = await graphFetch(putUrl, {
    method: "PUT",
    headers: {
      "Content-Type": contentType || "application/octet-stream",
    },
    body: buffer,
  });

  const putJson = await putRes.json();
  if (!putRes.ok) throw new Error(`Upload error: ${JSON.stringify(putJson, null, 2)}`);

  return {
    id: putJson.id,
    name: putJson.name,
    webUrl: putJson.webUrl,
  };
}

/* -------------------- Keep "last upload" in memory -------------------- */
let lastUpload = null; // { at, folderName, fileName, webUrl, source }

/* -------------------- Text commands (private only) -------------------- */
async function handlePrivateTextCommand(event) {
  const text = (event.message?.text || "").trim().toLowerCase();
  if (!text) return;

  if (text === "help") {
    await client.replyMessage(event.replyToken, [{
      type: "text",
      text:
        "คำสั่ง SavePhotoBot:\n" +
        "- help : ดูคำสั่ง\n" +
        "- status : สถานะระบบ\n" +
        "- last : ลิงก์รูปล่าสุด\n"
    }]);
    return;
  }

  if (text === "status") {
    const msg =
      "✅ SavePhotoBot ทำงานอยู่\n" +
      "โหมด: silent ใน group/room\n" +
      "Storage: OneDrive/SavePhotoBot/...";
    await client.replyMessage(event.replyToken, [{ type: "text", text: msg }]);
    return;
  }

  if (text === "last") {
    if (!lastUpload) {
      await client.replyMessage(event.replyToken, [{ type: "text", text: "ยังไม่มีรายการอัปโหลดล่าสุดครับ" }]);
      return;
    }
    await client.replyMessage(event.replyToken, [{
      type: "text",
      text:
        "🕒 ล่าสุด\n" +
        `เวลา: ${lastUpload.at}\n` +
        `ที่มา: ${lastUpload.source}\n` +
        `โฟลเดอร์: ${lastUpload.folderName}\n` +
        `ไฟล์: ${lastUpload.fileName}\n` +
        `ลิงก์: ${lastUpload.webUrl || "-"}`
    }]);
    return;
  }
}

/* =========================================================
   Webhook
   ========================================================= */
app.post("/webhook", line.middleware(config), async (req, res) => {
  // respond fast to avoid LINE retry storms
  res.sendStatus(200);

  const events = req.body?.events || [];

  // ✅ LOG: show what came in (so you know webhook is working)
  console.log("📩 Webhook IN | events =", events.length);
  for (const e of events) {
    console.log(
      "   ->",
      JSON.stringify({
        type: e.type,
        messageType: e.message?.type,
        messageId: e.message?.id,
        sourceType: e.source?.type,
        groupIdTail: e.source?.groupId ? e.source.groupId.slice(-6) : undefined,
        roomIdTail: e.source?.roomId ? e.source.roomId.slice(-6) : undefined,
        userIdTail: e.source?.userId ? e.source.userId.slice(-6) : undefined,
      })
    );
  }

  for (const event of events) {
    try {
      const srcType = event.source?.type;

      // Silent policy: never talk in group/room on join
      if (event.type === "join") continue;

      // follow happens in private
      if (event.type === "follow") {
        if (srcType === "user" && event.replyToken) {
          await client.replyMessage(event.replyToken, [
            { type: "text", text: "สวัสดีครับ 🙂 SavePhotoBot พร้อมรับรูปแล้ว ส่งรูปมาได้เลย" },
          ]);
        }
        continue;
      }

      // Text command in private
      if (event.type === "message" && event.message?.type === "text" && srcType === "user") {
        await handlePrivateTextCommand(event);
        continue;
      }

      // Only image messages
      if (event.type !== "message" || event.message?.type !== "image") continue;

      const messageId = event.message.id;

      // dedupe
      if (seenMessageIds.has(messageId)) {
        console.log("⚠️ Duplicate ignored:", messageId);
        continue;
      }
      rememberMessageId(messageId);

      // Determine folder
      const folderName = await getSourceFolder(event);

      // download from LINE -> temp file
      const stream = await client.getMessageContent(messageId);
      const ct = (stream?.headers?.["content-type"] || "").toLowerCase();
      const ext =
        ct.includes("png") ? "png" :
        ct.includes("jpeg") ? "jpg" :
        ct.includes("jpg") ? "jpg" :
        ct.includes("webp") ? "webp" :
        "jpg";

      const fileName = makeFileName(messageId, ext);
      const tempPath = path.join(tmpDir, fileName);

      console.log("📷 Image received:", messageId, "source:", sourceLabel(event), "-> folder:", folderName);

      await saveStreamToFile(stream, tempPath);

      // read file buffer
      const buffer = await fs.promises.readFile(tempPath);

      // ensure OneDrive folder path
      console.log("☁️ OneDrive ensure folder:", `SavePhotoBot/${folderName}`);
      const folderId = await ensureUploadFolder(folderName);

      // upload to OneDrive
      console.log("☁️ OneDrive uploading:", fileName);
      const uploaded = await uploadFileToOneDrive(folderId, fileName, buffer, ct || "image/jpeg");
      console.log("✅ OneDrive uploaded:", uploaded.webUrl);

      // cleanup temp
      fs.promises.unlink(tempPath).catch(() => {});

      // save last upload
      lastUpload = {
        at: new Date().toISOString(),
        folderName,
        fileName,
        webUrl: uploaded.webUrl,
        source: sourceLabel(event),
      };

      // Group/Room -> notify admin only (silent in group/room)
      if (srcType === "group" || srcType === "room") {
        const msg =
          `📸 มีรูปถูกส่งเข้ามา\n` +
          `ที่: ${sourceLabel(event)}\n` +
          `โฟลเดอร์: ${folderName}\n` +
          `ไฟล์: ${fileName}\n` +
          `OneDrive: ${uploaded.webUrl}`;
        await notifyAdmin(msg);
        continue;
      }

      // Private -> reply confirm (so it won't be silent/quiet)
      if (srcType === "user" && event.replyToken) {
        await client.replyMessage(event.replyToken, [
          {
            type: "text",
            text:
              `✅ อัปโหลดไป OneDrive แล้ว\n` +
              `โฟลเดอร์: ${folderName}\n` +
              `ไฟล์: ${fileName}\n` +
              `ลิงก์: ${uploaded.webUrl}`,
          },
        ]);
      }
    } catch (err) {
      console.error("❌ Error:", err?.message || err);
      console.error("DETAIL:", err);

      // try notify admin
      try {
        await notifyAdmin(`❌ SavePhotoBot Error: ${String(err?.message || err)}`);
      } catch (_) {}
    }
  }
});

/* -------------------- Start -------------------- */
const PORT = Number(process.env.PORT || 3001);
app.listen(PORT, () => console.log(`🚀 SavePhotoBot running on port ${PORT}`));