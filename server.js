require("dotenv").config();

const express = require("express");
const line = require("@line/bot-sdk");
const fs = require("fs");
const path = require("path");

const app = express();

// -------------------- ENV --------------------
const {
  LINE_ACCESS_TOKEN,
  LINE_CHANNEL_SECRET,
  ADMIN_USER_ID,

  MS_CLIENT_ID,
  MS_REFRESH_TOKEN,
  ONEDRIVE_BASE_PATH = "SavePhotoBot",
} = process.env;

if (!LINE_ACCESS_TOKEN || !LINE_CHANNEL_SECRET) {
  console.error("❌ Missing LINE env: LINE_ACCESS_TOKEN / LINE_CHANNEL_SECRET");
  process.exit(1);
}
if (!ADMIN_USER_ID) {
  console.error("❌ Missing env: ADMIN_USER_ID");
  process.exit(1);
}
if (!MS_CLIENT_ID || !MS_REFRESH_TOKEN) {
  console.error("❌ Missing OneDrive env: MS_CLIENT_ID / MS_REFRESH_TOKEN");
  process.exit(1);
}

const config = {
  channelAccessToken: LINE_ACCESS_TOKEN,
  channelSecret: LINE_CHANNEL_SECRET,
};
const client = new line.Client(config);

// -------------------- Local temp folder (ephemeral on Render) --------------------
const baseImagesDir = path.join(__dirname, "images");
if (!fs.existsSync(baseImagesDir)) fs.mkdirSync(baseImagesDir, { recursive: true });

// -------------------- Helpers --------------------
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

// -------------------- Cache: group/room name --------------------
const nameCache = new Map(); // key -> { name, ts }
const CACHE_TTL_MS = 24 * 60 * 60 * 1000;

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

// -------------------- OneDrive (Microsoft Graph) --------------------
// We'll keep refresh token in memory too (in case MS rotates it)
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
  try { json = text ? JSON.parse(text) : {}; } catch { json = { raw: text }; }
  if (!res.ok) throw new Error(`MS OAuth error: ${JSON.stringify(json)}`);
  return json;
}

async function getAccessToken() {
  // Token endpoint for personal account
  const tokenUrl = "https://login.microsoftonline.com/consumers/oauth2/v2.0/token";

  const tok = await msPostForm(tokenUrl, {
    client_id: MS_CLIENT_ID,
    grant_type: "refresh_token",
    refresh_token: currentRefreshToken,
    scope: "offline_access User.Read Files.ReadWrite",
  });

  // If MS returns a new refresh token, keep it in memory (Render env won't auto-update)
  if (tok.refresh_token && tok.refresh_token !== currentRefreshToken) {
    currentRefreshToken = tok.refresh_token;
    console.warn("⚠️ Refresh token rotated by Microsoft. Update MS_REFRESH_TOKEN on Render ASAP.");
  }

  return tok.access_token;
}

// Ensure folder exists (create by path)
async function ensureOneDriveFolder(accessToken, folderPath) {
  // Create nested folders via /root:/path: trick by creating last segment
  // We'll just create the full path by creating each segment.
  const parts = folderPath.split("/").filter(Boolean);

  let currentPath = "";
  for (const p of parts) {
    const nextPath = currentPath ? `${currentPath}/${p}` : p;

    // check exists
    const checkUrl = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(nextPath)}`;
    const checkRes = await fetch(checkUrl, {
      headers: { Authorization: `Bearer ${accessToken}` },
    });

    if (checkRes.ok) {
      currentPath = nextPath;
      continue;
    }

    // create folder under parent
    const parentPath = currentPath; // may be ""
    const createUrl = parentPath
      ? `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(parentPath)}:/children`
      : `https://graph.microsoft.com/v1.0/me/drive/root/children`;

    const createRes = await fetch(createUrl, {
      method: "POST",
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "Content-Type": "application/json",
      },
      body: JSON.stringify({
        name: p,
        folder: {},
        "@microsoft.graph.conflictBehavior": "rename",
      }),
    });

    if (!createRes.ok) {
      const err = await createRes.text();
      throw new Error(`Create folder failed (${nextPath}): ${err}`);
    }

    currentPath = nextPath;
  }
}

async function uploadSmallFileToOneDrive(accessToken, oneDrivePath, fileBuffer) {
  // PUT /me/drive/root:/path:/content
  const url = `https://graph.microsoft.com/v1.0/me/drive/root:/${encodeURIComponent(oneDrivePath)}:/content`;

  const res = await fetch(url, {
    method: "PUT",
    headers: {
      Authorization: `Bearer ${accessToken}`,
      "Content-Type": "image/jpeg",
    },
    body: fileBuffer,
  });

  if (!res.ok) {
    const err = await res.text();
    throw new Error(`Upload failed: ${err}`);
  }

  const json = await res.json();
  return json?.webUrl || null;
}

// -------------------- Routes --------------------
app.get("/", (req, res) => res.status(200).send("OK"));

app.post("/webhook", line.middleware(config), async (req, res) => {
  // respond fast
  res.sendStatus(200);

  const events = req.body?.events || [];

  for (const event of events) {
    try {
      // join/follow -> welcome (only in 1:1 or when added, ok)
      if (event.type === "follow") {
        await client.replyMessage(event.replyToken, [
          { type: "text", text: "สวัสดีครับ 🙂 SavePhotoBot พร้อมรับรูปแล้ว" },
        ]);
        continue;
      }

      if (event.type === "join") {
        // Silent policy: you can comment this out if you want 100% silent even on join
        // await client.replyMessage(event.replyToken, [{ type: "text", text: "SavePhotoBot พร้อมรับรูปแล้ว ✅" }]);
        continue;
      }

      // only handle image
      if (event.type === "message" && event.message?.type === "image") {
        const messageId = event.message.id;
        const folderName = await getSourceFolder(event);

        // 1) download image from LINE -> save temp
        const targetDir = path.join(baseImagesDir, folderName);
        if (!fs.existsSync(targetDir)) fs.mkdirSync(targetDir, { recursive: true });

        const fileName = makeFileName(messageId);
        const filePath = path.join(targetDir, fileName);

        const stream = await client.getMessageContent(messageId);
        await saveStreamToFile(stream, filePath);

        // 2) upload to OneDrive
        const accessToken = await getAccessToken();

        const oneDriveFolder = `${ONEDRIVE_BASE_PATH}/${folderName}`;
        await ensureOneDriveFolder(accessToken, oneDriveFolder);

        const buf = fs.readFileSync(filePath);
        const oneDrivePath = `${oneDriveFolder}/${fileName}`;
        const webUrl = await uploadSmallFileToOneDrive(accessToken, oneDrivePath, buf);

        // 3) notify admin only (silent in group/room)
        const src = event.source || {};
        const srcType = src.type || "unknown";
        const srcId = src.userId || src.groupId || src.roomId || "-";

        await client.pushMessage(ADMIN_USER_ID, [
          {
            type: "text",
            text:
              `📷 SavePhotoBot: มีรูปเข้าใหม่\n` +
              `Source: ${srcType}\n` +
              `Folder: ${folderName}\n` +
              `File: ${fileName}\n` +
              (webUrl ? `Link: ${webUrl}\n` : "") +
              `ID: ${srcId}`,
          },
        ]);

        // 4) optional: delete temp file to save disk
        try { fs.unlinkSync(filePath); } catch {}

        // IMPORTANT: do not reply/push to group/room/user automatically
        continue;
      }
    } catch (err) {
      console.error("❌ Error:", err?.message || err);
      // also notify admin if something critical
      try {
        await client.pushMessage(ADMIN_USER_ID, [
          { type: "text", text: `❌ SavePhotoBot error: ${String(err?.message || err).slice(0, 900)}` },
        ]);
      } catch {}
    }
  }
});

// -------------------- Start --------------------
const PORT = Number(process.env.PORT || 3001);
app.listen(PORT, () => console.log(`🚀 Server running on port ${PORT}`));