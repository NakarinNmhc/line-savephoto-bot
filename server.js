require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");

const app = express();

// -------------------- Basic routes --------------------
app.get("/", (req, res) => res.status(200).send("OK"));
app.get("/health", (req, res) => res.status(200).send("OK"));

// -------------------- LINE config --------------------
const config = {
  channelAccessToken: process.env.LINE_ACCESS_TOKEN,
  channelSecret: process.env.LINE_CHANNEL_SECRET,
};

const ADMIN_USER_ID = process.env.ADMIN_USER_ID;

// (แนะนำ) ใส่ token เพื่อกันคนอื่นเดา URL แล้วดูรูป
const IMAGE_VIEW_TOKEN = process.env.IMAGE_VIEW_TOKEN || "";

if (!config.channelAccessToken || !config.channelSecret) {
  console.error("❌ Missing env: LINE_ACCESS_TOKEN or LINE_CHANNEL_SECRET");
  process.exit(1);
}
if (!ADMIN_USER_ID) {
  console.error("❌ Missing env: ADMIN_USER_ID");
  process.exit(1);
}

const client = new line.Client(config);

// -------------------- Storage base folder --------------------
const baseImagesDir = path.join(__dirname, "images");
if (!fs.existsSync(baseImagesDir)) fs.mkdirSync(baseImagesDir, { recursive: true });

// -------------------- Static route to view images (optional) --------------------
app.get("/images/*", (req, res, next) => {
  // ถ้าไม่ตั้ง token = เปิดโล่ง (ไม่แนะนำบน production)
  if (!IMAGE_VIEW_TOKEN) return next();
  if (req.query.token !== IMAGE_VIEW_TOKEN) return res.sendStatus(403);
  return next();
});
app.use("/images", express.static(baseImagesDir));

// -------------------- Helpers --------------------
function pad(n) {
  return String(n).padStart(2, "0");
}

function makeFileName(messageId, ext = "jpg") {
  const d = new Date();
  return (
    `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}` +
    `_${pad(d.getHours())}-${pad(d.getMinutes())}-${pad(d.getSeconds())}` +
    `_${messageId}.${ext}`
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

function buildPublicBaseUrl(req) {
  const proto = req.headers["x-forwarded-proto"] || "https";
  const host = req.headers["x-forwarded-host"] || req.headers.host;
  return `${proto}://${host}`;
}

// -------------------- Cache: group name --------------------
const nameCache = new Map(); // key -> { name, ts }
const CACHE_TTL_MS = 24 * 60 * 60 * 1000; // 24h

async function getGroupName(groupId) {
  const key = `group:${groupId}`;
  const cached = nameCache.get(key);
  if (cached && Date.now() - cached.ts < CACHE_TTL_MS) return cached.name;

  const summary = await client.getGroupSummary(groupId); // ✅ มีจริง
  const name = sanitizeFolderName(summary.groupName || "UnknownGroup");
  nameCache.set(key, { name, ts: Date.now() });
  return name;
}

// -------------------- Source folder (group/room/private) --------------------
async function getSourceFolder(event) {
  const src = event.source || {};

  if (src.type === "user") return "private";

  if (src.type === "group" && src.groupId) {
    const tail = src.groupId.slice(-6);
    try {
      const name = await getGroupName(src.groupId);
      return `group_${name}_${tail}`;
    } catch (_) {
      return `group_${tail}`;
    }
  }

  if (src.type === "room" && src.roomId) {
    // ❌ ไม่มี getRoomSummary ใน Messaging API → ใช้ roomId แทน
    const tail = src.roomId.slice(-6);
    return `room_${tail}`;
  }

  return "unknown";
}

function sourceText(event) {
  const src = event.source || {};
  if (src.type === "group") return `GROUP (${src.groupId?.slice(-6) || ""})`;
  if (src.type === "room") return `ROOM (${src.roomId?.slice(-6) || ""})`;
  if (src.type === "user") return `PRIVATE (${src.userId?.slice(-6) || ""})`;
  return "UNKNOWN";
}

// -------------------- Dedupe กัน webhook retry --------------------
const seenMessageIds = new Set();
function rememberMessageId(id) {
  seenMessageIds.add(id);
  setTimeout(() => seenMessageIds.delete(id), 10 * 60 * 1000).unref?.();
}

// -------------------- Main webhook --------------------
app.post("/webhook", line.middleware(config), async (req, res) => {
  // ตอบ 200 ให้เร็ว กัน LINE timeout/retry
  res.sendStatus(200);

  const events = req.body?.events || [];
  const baseUrl = buildPublicBaseUrl(req);

  for (const event of events) {
    try {
      const srcType = event.source?.type;

      // -------------------------------------------------
      // 1) follow/join: ตอบกลับเฉพาะ PRIVATE เท่านั้น
      //    - group/room: เงียบ 100%
      // -------------------------------------------------
      if (event.type === "follow") {
        if (srcType === "user" && event.replyToken) {
          await client.replyMessage(event.replyToken, [
            { type: "text", text: "สวัสดีครับ 🙂 SavePhotoBot พร้อมรับรูปแล้ว" },
          ]);
        }
        continue;
      }

      if (event.type === "join") {
        // join เกิดตอนเข้ากลุ่ม/ห้อง → ต้อง silent
        continue;
      }

      // -------------------------------------------------
      // 2) ข้อความ text: ตอบกลับเฉพาะ PRIVATE เท่านั้น (optional)
      // -------------------------------------------------
      if (event.type === "message" && event.message?.type === "text") {
        if (srcType === "user" && event.replyToken) {
          await client.replyMessage(event.replyToken, [
            { type: "text", text: "✅ รับทราบครับ ส่งรูปมาได้เลย" },
          ]);
        }
        continue;
      }

      // -------------------------------------------------
      // 3) รูปภาพ: save เสมอ + notify ADMIN เสมอ
      //    - group/room: ห้าม reply/push กลับไปที่ group/room
      //    - private: จะ reply สั้นๆ ก็ได้ (optional)
      // -------------------------------------------------
      if (event.type === "message" && event.message?.type === "image") {
        const messageId = event.message.id;

        if (seenMessageIds.has(messageId)) {
          console.log("⚠️ Duplicate messageId ignored:", messageId);
          continue;
        }
        rememberMessageId(messageId);

        const folderName = await getSourceFolder(event);
        const targetDir = path.join(baseImagesDir, folderName);
        if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });

        const stream = await client.getMessageContent(messageId);

        // พยายามเดานามสกุลจาก content-type
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
        console.log("✅ Image saved:", filePath);

        // สร้างลิงก์ดูรูป (ถ้าตั้ง static route)
        const viewPath = `/images/${encodeURIComponent(folderName)}/${encodeURIComponent(fileName)}`;
        const viewUrl = IMAGE_VIEW_TOKEN
          ? `${baseUrl}${viewPath}?token=${encodeURIComponent(IMAGE_VIEW_TOKEN)}`
          : `${baseUrl}${viewPath}`;

        // ✅ แจ้ง ADMIN เสมอ (DM)
        const msg =
          `📸 มีรูปถูกส่งเข้ามา\n` +
          `ที่: ${sourceText(event)}\n` +
          `ผู้ส่ง: ${event.source?.userId || "-"}\n` +
          `โฟลเดอร์: ${folderName}\n` +
          `ไฟล์: ${fileName}\n` +
          `ดูรูป: ${viewUrl}`;

        await client.pushMessage(ADMIN_USER_ID, [{ type: "text", text: msg }]);

        // ❗ silent ใน group/room: ห้าม reply/push กลับกลุ่ม
        if (srcType === "user" && event.replyToken) {
          // optional: private ค่อยตอบกลับสั้นๆ
          await client.replyMessage(event.replyToken, [
            { type: "text", text: "✅ บันทึกรูปแล้วครับ" },
          ]);
        }

        continue;
      }

      // event อื่นๆ: เงียบไปเลย
    } catch (err) {
      console.error("❌ Error:", err?.message || err);
      try {
        await client.pushMessage(ADMIN_USER_ID, [
          { type: "text", text: `❌ SavePhotoBot Error: ${String(err?.message || err)}` },
        ]);
      } catch (_) {}
    }
  }
});

// -------------------- Start --------------------
const PORT = Number(process.env.PORT || 3001);
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));
