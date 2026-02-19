require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");

const app = express();

const config = {
  channelAccessToken: process.env.LINE_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

if (!config.channelAccessToken || !config.channelSecret) {
  console.error("❌ Missing env: LINE_ACCESS_TOKEN or LINE_CHANNEL_SECRET");
  process.exit(1);
}

const client = new line.Client(config);

// Health check (Render/Browser)
app.get("/", (req, res) => res.status(200).send("OK"));
app.get("/webhook", (req, res) => res.status(200).send("OK")); // กัน verify แบบ GET บางที่

// =====================
// Storage base folder
// =====================
const baseImagesDir = path.join(__dirname, "images");
if (!fs.existsSync(baseImagesDir)) fs.mkdirSync(baseImagesDir, { recursive: true });

// =====================
// Helpers
// =====================
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

// =====================
// Cache: group/room name
// =====================
const nameCache = new Map(); // key -> { name, ts }
const CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24 ชั่วโมง

async function getGroupOrRoomName(source) {
  if (!source?.type) return null;

  if (source.type === "group" && source.groupId) {
    const key = `group:${source.groupId}`;
    const cached = nameCache.get(key);
    if (cached && Date.now() - cached.ts < CACHE_TTL_MS) return cached.name;

    const summary = await client.getGroupSummary(source.groupId);
    const name = sanitizeFolderName(summary.groupName || "UnknownGroup");
    nameCache.set(key, { name, ts: Date.now() });
    return name;
  }

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

  // แชทส่วนตัว
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

function isPrivateChat(event) {
  return event?.source?.type === "user";
}

// =====================
// Webhook
// =====================
app.post("/webhook", line.middleware(config), async (req, res) => {
  // ตอบ 200 เร็ว ๆ กัน timeout
  res.sendStatus(200);

  const events = req.body?.events || [];
  console.log("📩 Webhook triggered. Events:", events.length);

  for (const event of events) {
    try {
      // 0) ไม่ส่งข้อความในกลุ่ม/รูม (กันวุ่นวาย)
      //    แต่ในแชทส่วนตัว เราจะตอบเฉพาะตอนบันทึกเสร็จ
      const privateChat = isPrivateChat(event);

      // 1) รับรูปเท่านั้น (ไม่ตอบข้อความ text, ไม่ทัก join/follow)
      if (event.type === "message" && event.message?.type === "image") {
        const messageId = event.message.id;
        const folderName = await getSourceFolder(event);

        const targetDir = path.join(baseImagesDir, folderName);
        if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });

        const fileName = makeFileName(messageId);
        const filePath = path.join(targetDir, fileName);

        console.log("📷 Image received:", messageId, "->", folderName);

        const stream = await client.getMessageContent(messageId);
        await saveStreamToFile(stream, filePath);

        console.log("✅ Image saved:", filePath);

        // ✅ ตอบกลับเฉพาะแชทส่วนตัวเท่านั้น (ไม่ส่งอะไรในกลุ่ม)
        if (privateChat && event.replyToken) {
          await client.replyMessage(event.replyToken, [
            { type: "text", text: `✅ บันทึกรูปเรียบร้อย\nไฟล์: ${fileName}` },
          ]);
        }

        continue;
      }

      // event อื่น ๆ ไม่ต้องทำอะไร (เงียบ)
    } catch (err) {
      console.error("❌ Error:", err);
      console.error("LINE API error body:", err?.originalError?.response?.data);
    }
  }
});

// =====================
// Start
// =====================
const PORT = Number(process.env.PORT || 3001);
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));
