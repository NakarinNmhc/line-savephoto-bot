/**
 * SavePhotoBot - production-ready for Render
 * - Silent in group/room (never reply to group/room)
 * - Save images to /images/<source-folder>
 * - Notify ADMIN_USER_ID via DM whenever image is sent in group/room
 * - Optional: static route to view images with IMAGE_VIEW_TOKEN
 */

require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");

const app = express();

/* -------------------- Health check -------------------- */
app.get("/", (req, res) => res.status(200).send("OK"));
app.get("/health", (req, res) => res.status(200).send("OK"));

/* -------------------- ENV -------------------- */
const LINE_ACCESS_TOKEN = process.env.LINE_ACCESS_TOKEN;
const LINE_CHANNEL_SECRET = process.env.LINE_CHANNEL_SECRET;
const ADMIN_USER_ID = process.env.ADMIN_USER_ID;

// Optional: protect image viewing route
const IMAGE_VIEW_TOKEN = process.env.IMAGE_VIEW_TOKEN || "";

/* -------------------- Validate ENV -------------------- */
if (!LINE_ACCESS_TOKEN || !LINE_CHANNEL_SECRET) {
  console.error("❌ Missing LINE_ACCESS_TOKEN or LINE_CHANNEL_SECRET");
  process.exit(1);
}
if (!ADMIN_USER_ID) {
  console.error("❌ Missing ADMIN_USER_ID (must be set in Render > Environment)");
  process.exit(1);
}

/* -------------------- LINE client -------------------- */
const config = {
  channelAccessToken: LINE_ACCESS_TOKEN,
  channelSecret: LINE_CHANNEL_SECRET,
};
const client = new line.Client(config);

/* -------------------- Storage -------------------- */
const baseImagesDir = path.join(__dirname, "images");
if (!fs.existsSync(baseImagesDir)) fs.mkdirSync(baseImagesDir, { recursive: true });

/* -------------------- Static route (Express 5 safe) -------------------- */
// IMPORTANT: Do NOT use "/images/*" (it can crash with path-to-regexp)
// Use RegExp route instead.
app.get(/^\/images\/.*/, (req, res, next) => {
  // If token not set => open (dev mode)
  if (!IMAGE_VIEW_TOKEN) return next();
  if (req.query.token !== IMAGE_VIEW_TOKEN) return res.sendStatus(403);
  return next();
});
app.use("/images", express.static(baseImagesDir));

/* -------------------- Helpers -------------------- */
function pad(n) {
  return String(n).padStart(2, "0");
}

function makeFileName(messageId, ext = "jpg") {
  const d = new Date();
  // requirement: date + messageId
  // (add time is optional; we keep it clean: YYYY-MM-DD_messageId.ext)
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

/* -------------------- Cache group name (optional) -------------------- */
const nameCache = new Map(); // groupId -> {name, ts}
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
    } catch (_) {
      return `group_${tail}`;
    }
  }

  if (s.type === "room" && s.roomId) {
    // LINE API ไม่มี getRoomSummary => ใช้ roomId แทน
    const tail = s.roomId.slice(-6);
    return `room_${tail}`;
  }

  return "unknown";
}

/* -------------------- Dedupe (webhook retry) -------------------- */
const seenMessageIds = new Set();
function rememberMessageId(id) {
  seenMessageIds.add(id);
  // auto clear after 10 minutes
  setTimeout(() => seenMessageIds.delete(id), 10 * 60 * 1000).unref?.();
}

/* -------------------- Notify ADMIN (DM) -------------------- */
async function notifyAdmin(text) {
  return client.pushMessage(ADMIN_USER_ID, [{ type: "text", text }]);
}

/* -------------------- Webhook -------------------- */
app.post("/webhook", line.middleware(config), async (req, res) => {
  // Respond fast to avoid LINE retry
  res.sendStatus(200);

  const events = req.body?.events || [];
  const baseUrl = buildPublicBaseUrl(req);

  for (const event of events) {
    try {
      // --------- Only handle image messages ----------
      if (event.type !== "message" || event.message?.type !== "image") continue;

      const srcType = event.source?.type; // group/room/user
      const messageId = event.message.id;

      // Dedupe
      if (seenMessageIds.has(messageId)) {
        console.log("⚠️ Duplicate message ignored:", messageId);
        continue;
      }
      rememberMessageId(messageId);

      // Prepare folder
      const folderName = await getSourceFolder(event);
      const targetDir = path.join(baseImagesDir, folderName);
      if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });

      // Get image content stream
      const stream = await client.getMessageContent(messageId);

      // Detect extension (best-effort)
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
      console.log("✅ Saved:", filePath);

      // Build view URL (optional)
      const viewPath = `/images/${encodeURIComponent(folderName)}/${encodeURIComponent(fileName)}`;
      const viewUrl = IMAGE_VIEW_TOKEN
        ? `${baseUrl}${viewPath}?token=${encodeURIComponent(IMAGE_VIEW_TOKEN)}`
        : `${baseUrl}${viewPath}`;

      // --------- Silent in group/room ----------
      // Always notify admin when image comes from group/room.
      if (srcType === "group" || srcType === "room") {
        const msg =
          `📸 มีรูปถูกส่งเข้ามา\n` +
          `ที่: ${sourceLabel(event)}\n` +
          `ผู้ส่ง: ${event.source?.userId || "-"}\n` +
          `โฟลเดอร์: ${folderName}\n` +
          `ไฟล์: ${fileName}\n` +
          `ดูรูป: ${viewUrl}`;

        await notifyAdmin(msg);
        // DO NOT reply to group/room
        continue;
      }

      // --------- Private chat ----------
      // You can choose to reply or not. (Doesn't violate "silent in group")
      // If you want "never auto reply anywhere", comment out the replyMessage below.
      if (srcType === "user" && event.replyToken) {
        await client.replyMessage(event.replyToken, [
          { type: "text", text: "✅ บันทึกรูปแล้วครับ" },
        ]);
      }

      // (Optional) also notify admin even for private images:
      // await notifyAdmin(`📸 Private image saved: ${fileName}\n${viewUrl}`);
    } catch (err) {
      console.error("❌ Error:", err?.message || err);
      try {
        await notifyAdmin(`❌ SavePhotoBot Error: ${String(err?.message || err)}`);
      } catch (_) {}
    }
  }
});

/* -------------------- Start -------------------- */
const PORT = Number(process.env.PORT || 3001);
app.listen(PORT, () => console.log(`🚀 SavePhotoBot running on port ${PORT}`));
