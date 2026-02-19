//te2Vfvwg4Qe7IbDYgQRPZrn9k5rCTVRP7EaEPgudeGVAVsJwwJquX5mh6+dZMGc4nCftCN7RVbBW9OmH++bZQ4Lye7nldVedlmja3O58c4suHUP/aDnswixvrgbGqZyeHH6+MLPLM0OCjKyOWV35kAdB04t89/1O/w1cDnyilFU=
//f82f6612b4ca51cee0cefafdd641f225

require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");

const app = express();

app.get("/", (req, res) => res.status(200).send("OK"));
app.get("/webhook", (req, res) => res.status(200).send("OK"));


const config = {
  channelAccessToken: process.env.LINE_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

if (!config.channelAccessToken || !config.channelSecret) {
  console.error("❌ Missing env: LINE_ACCESS_TOKEN or LINE_CHANNEL_SECRET");
  process.exit(1);
}

const client = new line.Client(config);

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
  // กันอักขระต้องห้ามใน Windows + ย่อความยาว
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

  // GROUP
  if (source.type === "group" && source.groupId) {
    const key = `group:${source.groupId}`;
    const cached = nameCache.get(key);
    if (cached && Date.now() - cached.ts < CACHE_TTL_MS) return cached.name;

    const summary = await client.getGroupSummary(source.groupId); // { groupId, groupName, pictureUrl }
    const name = sanitizeFolderName(summary.groupName || "UnknownGroup");
    nameCache.set(key, { name, ts: Date.now() });
    return name;
  }

  // ROOM
  if (source.type === "room" && source.roomId) {
    const key = `room:${source.roomId}`;
    const cached = nameCache.get(key);
    if (cached && Date.now() - cached.ts < CACHE_TTL_MS) return cached.name;

    const summary = await client.getRoomSummary(source.roomId); // { roomId, roomName, pictureUrl }
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

// =====================
// Routes
// =====================
app.get("/", (req, res) => res.status(200).send("OK"));

app.post("/webhook", line.middleware(config), async (req, res) => {
  // ตอบ 200 ให้เร็ว (กัน LINE timeout)
  res.sendStatus(200);

  const events = req.body?.events || [];
  console.log("📩 Webhook triggered. Events:", events.length);

  for (const event of events) {
    try {
      // 1) join/follow: ทักทาย
      if (event.type === "join" || event.type === "follow") {
        await client.replyMessage(event.replyToken, [
          { type: "text", text: "สวัสดีครับ 🙂 SavePhotoBot พร้อมรับรูปแล้ว ส่งรูปมาได้เลย" },
        ]);
        console.log("✅ Replied welcome for:", event.type, event.source);
        continue;
      }

      // 2) ข้อความเทส
      if (event.type === "message" && event.message?.type === "text") {
        await client.replyMessage(event.replyToken, [
          { type: "text", text: "✅ เห็นข้อความแล้วครับ ส่งรูปมาได้เลย" },
        ]);
        continue;
      }

      // 3) รับรูป
      if (event.type === "message" && event.message?.type === "image") {
        const messageId = event.message.id;
        const folderName = await getSourceFolder(event);

        const targetDir = path.join(baseImagesDir, folderName);
        if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });

        const fileName = makeFileName(messageId);
        const filePath = path.join(targetDir, fileName);

        console.log("📷 Image received:", messageId, "->", folderName);

        // reply ทันที (กัน replyToken หมดอายุ) — ใช้ได้ครั้งเดียว
        if (event.replyToken) {
          await client.replyMessage(event.replyToken, [
            { type: "text", text: "📥 รับรูปแล้วครับ กำลังบันทึก..." },
          ]);
        }

        // โหลดรูปจาก LINE และบันทึกไฟล์
        const stream = await client.getMessageContent(messageId);
        await saveStreamToFile(stream, filePath);

        console.log("✅ Image saved:", filePath);

        // แจ้ง “บันทึกเสร็จ” (ต้องใช้ push เพราะ replyToken ใช้ไปแล้ว)
        const to = event.source?.userId || event.source?.groupId || event.source?.roomId;
        if (to) {
          await client.pushMessage(to, [
            {
              type: "text",
              text: `✅ บันทึกรูปเรียบร้อย\nโฟลเดอร์: ${folderName}\nไฟล์: ${fileName}`,
            },
          ]);
        }

        continue;
      }

      // event อื่นๆ
      // console.log("ℹ️ Event type:", event.type);
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
