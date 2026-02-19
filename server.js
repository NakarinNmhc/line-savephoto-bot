/**
 * SavePhotoBot - Render/Local
 * - Save images into /images/<sourceFolder>/<YYYY-MM-DD_HH-mm-ss_messageId>.jpg
 * - Never reply in group/room (silent)
 * - Always notify admin privately when any image is received (from any group/room/user)
 *
 * Required env:
 *  - LINE_ACCESS_TOKEN
 *  - LINE_CHANNEL_SECRET
 *  - ADMIN_USER_ID        (userId ของ admin ที่จะรับแจ้งเตือน)
 *  - PORT                 (Render จะ set ให้เอง)
 */

require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");

const app = express();

// ---------- Health check ----------
app.get("/", (req, res) => res.status(200).send("OK"));
app.get("/health", (req, res) => res.status(200).send("OK"));

// ---------- LINE Config ----------
const config = {
  channelAccessToken: process.env.LINE_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const ADMIN_USER_ID = process.env.ADMIN_USER_ID;

if (!config.channelAccessToken || !config.channelSecret) {
  console.error("❌ Missing env: LINE_ACCESS_TOKEN or LINE_CHANNEL_SECRET");
  process.exit(1);
}
if (!ADMIN_USER_ID) {
  console.error("❌ Missing env: ADMIN_USER_ID");
  process.exit(1);
}

const client = new line.Client(config);

// ---------- Storage base folder ----------
const baseImagesDir = path.join(__dirname, "images");
if (!fs.existsSync(baseImagesDir)) fs.mkdirSync(baseImagesDir, { recursive: true });

// ---------- Helpers ----------
function pad(n) {
  return String(n).padStart(2, "0");
}

function makeFileName(messageId) {
  const d = new Date();
  return (
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` +
    `_${pad(d.getHours())}-${pad(d.getMinutes())}-${pad(d.getSeconds())}` +
    `_${messageId}.jpg`
  );
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

// ---------- Cache group/room name ----------
const nameCache = new Map(); // key -> { name, ts }
const CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24h

async function getGroupOrRoomName(source) {
  if (!source?.type) return null;

  // GROUP
  if (source.type === "group" && source.groupId) {
    const key = `group:${source.groupId}`;
    const cached = nameCache.get(key);
    if (cached && Date.now() - cached.ts < CACHE_TTL_MS) return cached.name;

    // ต้องเปิด “Allow bot to join group chats” และบอทต้องอยู่ในกลุ่มนั้น
    const summary = await client.getGroupSummary(source.groupId);
    const name = sanitizeFolderName(summary.groupName || "UnknownGroup");
    nameCache.set(key, { name, ts: Date.now() });
    return name;
  }

  // ROOM
  if (source.type === "room" && source.roomId) {
    const key = `room:${source.roomId}`;
    const cached = nameCache.get(key);
    if (cached && Date.now() - cached.ts < CACHE_TTL_MS) return cached.name;

    const summary = await client.getRoomSummary(source.roomId);
    const name = sanitizeFolderName(summary.roomName || "UnknownRoom");
    nameCache.set(key, { name, ts: Date.now() });
    return name;
  }

  return null;
}

async function getSourceFolder(event) {
  const src = event.source || {};

  if (src.type === "user") return "private";

  const name = await getGroupOrRoomName(src);

  if (src.type === "group" && src.groupId) {
    const tail = src.groupId.slice(-6);
    return name ? `group_${name}_${tail}` : `group_${src.groupId}`;
  }

  if (src.type === "room" && src.roomId) {
    const tail = src.roomId.slice(-6);
    return name ? `room_${name}_${tail}` : `room_${src.roomId}`;
  }

  return "unknown";
}

function sourceLabel(event) {
  const src = event.source || {};
  if (src.type === "group") return `group:${src.groupId}`;
  if (src.type === "room") return `room:${src.roomId}`;
  if (src.type === "user") return `user:${src.userId}`;
  return "unknown";
}

// ---------- Webhook ----------
app.post("/webhook", line.middleware(config), async (req, res) => {
  // ตอบ 200 ให้เร็ว ป้องกัน timeout จาก LINE
  res.sendStatus(200);

  const events = req.body?.events || [];
  console.log("📩 Webhook triggered. Events:", events.length);

  for (const event of events) {
    try {
      // *** เราจะ “เงียบ” ในกลุ่ม/room ทุกกรณี ***
      // และ “ไม่ต้องตอบข้อความ text” เพื่อไม่ให้รก

      // รับเฉพาะรูป
      if (event.type === "message" && event.message?.type === "image") {
        const messageId = event.message.id;

        const folderName = await getSourceFolder(event);
        const targetDir = path.join(baseImagesDir, folderName);
        if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });

        const fileName = makeFileName(messageId);
        const filePath = path.join(targetDir, fileName);

        console.log("📷 Image received:", messageId, "from", sourceLabel(event), "->", folderName);

        // download content แล้ว save
        const stream = await client.getMessageContent(messageId);
        await saveStreamToFile(stream, filePath);

        console.log("✅ Image saved:", filePath);

        // ✅ แจ้ง admin เสมอ (ทางแชทส่วนตัว)
        await client.pushMessage(ADMIN_USER_ID, [
          {
            type: "text",
            text:
              `✅ บันทึกรูปเรียบร้อย\n` +
              `จาก: ${sourceLabel(event)}\n` +
              `โฟลเดอร์: ${folderName}\n` +
              `ไฟล์: ${fileName}`,
          },
        ]);

        continue;
      }

      // ถ้าอยาก log อย่างเดียว (ไม่ตอบกลับ)
      // console.log("ℹ️ Ignored event:", event.type);
    } catch (err) {
      console.error("❌ Error:", err);
      console.error("LINE API error body:", err?.originalError?.response?.data);

      // แจ้ง admin เมื่อ error (optional แต่มีประโยชน์)
      try {
        await client.pushMessage(ADMIN_USER_ID, [
          {
            type: "text",
            text: `❌ SavePhotoBot Error\nจาก: ${sourceLabel(event)}\n${String(err?.message || err)}`,
          },
        ]);
      } catch (_) {}
    }
  }
});

// ---------- Start ----------
const PORT = Number(process.env.PORT || 3001);
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));
