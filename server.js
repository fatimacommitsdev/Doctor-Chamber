// server.js — WhatsApp Bulk Sender Backend (Node.js + Express)
// ─────────────────────────────────────────────────────────────
// npm install express multer@2 xlsx axios cors form-data uuid

const express   = require('express');
const multer    = require('multer');
const XLSX      = require('xlsx');
const axios     = require('axios');
const cors      = require('cors');
const path      = require('path');
const fs        = require('fs');
const { v4: uuidv4 } = require('uuid'); // npm install uuid

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Temp media directory ──────────────────────────────────────
const MEDIA_DIR = path.join(__dirname, 'tmp_media');
if (!fs.existsSync(MEDIA_DIR)) fs.mkdirSync(MEDIA_DIR);

// ── Middleware ────────────────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// ✅ Serve temp media files publicly so Maytapi can fetch them by URL
app.use('/media', express.static(MEDIA_DIR));

// multer v2: accept both the Excel file AND an optional media file
const { memoryStorage } = require('multer');
const upload = multer({
  storage: memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 }, // 50 MB max
});

// ── Utility: Parse Excel / CSV buffer ─────────────────────────
function parseExcel(buffer, originalName) {
  const wb   = XLSX.read(buffer, { type: 'buffer' });
  const ws   = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

// ── Utility: Replace {{placeholders}} with row values ─────────
function fillTemplate(template, row) {
  return template.replace(/\{\{(\w+)\}\}/g, (_, key) => row[key] ?? '');
}

// ── Utility: Normalize phone number ───────────────────────────
function normalizePhone(raw) {
  let phone = String(raw).replace(/[\s\-().]/g, '');
  if (!phone.startsWith('+')) phone = '+' + phone;
  return phone;
}

// ── Utility: Detect Maytapi media type from mime ──────────────
function maytapiType(mimetype) {
  if (!mimetype) return 'text';
  if (mimetype.startsWith('image/'))  return 'image';
  if (mimetype.startsWith('video/'))  return 'video';
  if (mimetype.startsWith('audio/'))  return 'audio';
  return 'media'; // PDF, docs, etc.
}

// ── Utility: Build Maytapi payload ────────────────────────────
// ✅ FIX: Use a public URL for media instead of base64.
// Base64 encoding bloats payload 3-4x and causes HTTP 413 errors.
function buildPayload({ to, message, mediaUrl, mediaMime }) {
  if (!mediaUrl) {
    // Text only
    return { to_number: to, type: 'text', message };
  }

  const type = maytapiType(mediaMime);

  // Maytapi media payload with URL reference
  return {
    to_number: to,
    type,
    message,      // caption shown below the media
    media: mediaUrl,
  };
}

// ── Utility: Send one WhatsApp message via Maytapi ────────────
async function sendWhatsApp({ productId, phoneId, apiToken, to, message, mediaUrl, mediaMime }) {
  const url     = `https://api.maytapi.com/api/${productId}/${phoneId}/sendMessage`;
  const payload = buildPayload({ to, message, mediaUrl, mediaMime });

  // ✅ Add this — see exactly what's being sent
  console.log('📤 Sending to:', to);
  console.log('📦 Payload:', JSON.stringify(payload, null, 2));

  const response = await axios.post(url, payload, {
    headers: {
      'Content-Type': 'application/json',
      'x-maytapi-key': apiToken,
    },
    timeout: 30000,
  });

  // ✅ Add this — see exactly what Maytapi responds
  console.log('📩 Maytapi response:', JSON.stringify(response.data, null, 2));

  if (response.data?.success === false) {
    throw new Error(response.data?.message || 'Maytapi returned success: false');
  }

  return response.data;
}

// ── Utility: Sleep ────────────────────────────────────────────
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// ── Utility: Clean up temp media file ────────────────────────
function cleanupFile(filePath) {
  setTimeout(() => {
    fs.unlink(filePath, () => {}); // delete after 10 min
  }, 10 * 60 * 1000);
}

// ──────────────────────────────────────────────────────────────
// Route: GET /  → serve frontend
// ──────────────────────────────────────────────────────────────
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ──────────────────────────────────────────────────────────────
// Route: POST /api/columns
// ──────────────────────────────────────────────────────────────
app.post('/api/columns', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const rows    = parseExcel(req.file.buffer, req.file.originalname);
    const columns = rows.length > 0 ? Object.keys(rows[0]) : [];
    res.json({ columns, rowCount: rows.length });
  } catch (err) {
    res.status(422).json({ error: 'Failed to parse file: ' + err.message });
  }
});

// ──────────────────────────────────────────────────────────────
// Route: POST /api/send
// Fields (multipart/form-data):
//   excel      — Excel/CSV file
//   media      — (optional) PDF / image / video / audio
//   productId, phoneId, apiToken, message, delay, serverUrl
// ──────────────────────────────────────────────────────────────
app.post('/api/send',
  upload.fields([
    { name: 'excel', maxCount: 1 },
    { name: 'media', maxCount: 1 },
  ]),
  async (req, res) => {
    const { productId, phoneId, apiToken, message, delay: rawDelay, serverUrl } = req.body;
    const delay = Math.max(500, parseInt(rawDelay) || 1500);

    if (!productId || !phoneId || !apiToken || !message) {
      return res.status(400).json({ error: 'Missing required fields' });
    }

    const excelFile = req.files?.excel?.[0];
    const mediaFile = req.files?.media?.[0];

    if (!excelFile) return res.status(400).json({ error: 'No Excel file uploaded' });

    // Parse Excel
    let rows;
    try {
      rows = parseExcel(excelFile.buffer, excelFile.originalname);
    } catch (err) {
      return res.status(422).json({ error: 'Failed to parse Excel: ' + err.message });
    }

    if (rows.length === 0) return res.status(422).json({ error: 'Excel file is empty' });

    // Detect phone column
    const keys     = Object.keys(rows[0]);
    const phoneCol = keys.find(k => /^phone$/i.test(k.trim()))
                  || keys.find(k => /phone|mobile|number|whatsapp/i.test(k))
                  || keys[0];

    // ✅ Save media to disk and create a public URL for Maytapi to fetch
    let mediaUrl  = null;
    let tempFile  = null;

    if (mediaFile) {
      const ext      = path.extname(mediaFile.originalname) || '';
      const filename = `${uuidv4()}${ext}`;
      tempFile       = path.join(MEDIA_DIR, filename);
      fs.writeFileSync(tempFile, mediaFile.buffer);

      // Build the public URL — use serverUrl from frontend or auto-detect
      const base = (serverUrl || `http://localhost:${PORT}`).replace(/\/$/, '');
      mediaUrl   = `${base}/media/${filename}`;

      // Schedule cleanup after 10 minutes
      cleanupFile(tempFile);
    }

    // SSE setup
    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();

    const emit = (obj) => res.write(`data:${JSON.stringify(obj)}\n\n`);

    const total = rows.length;
    let sent = 0, failed = 0;

    for (let i = 0; i < rows.length; i++) {
      const row             = rows[i];
      const phone           = normalizePhone(row[phoneCol]);
      const personalizedMsg = fillTemplate(message, row);

      try {
        await sendWhatsApp({
          productId, phoneId, apiToken,
          to: phone,
          message: personalizedMsg,
          mediaUrl,
          mediaMime: mediaFile?.mimetype || null,
        });
        sent++;
        emit({ type: 'progress', sent, total, failed, phone, index: i + 1 });
      } catch (err) {
        failed++;
        const reason = err.response?.data?.message || err.message || 'Unknown error';
        emit({ type: 'error', phone, reason, index: i + 1 });
      }

      if (i < rows.length - 1) await sleep(delay);
    }

    emit({ type: 'done', sent, failed, total });
    res.end();
  }
);

// ──────────────────────────────────────────────────────────────
// Route: POST /api/preview  (dry run)
// ──────────────────────────────────────────────────────────────
app.post('/api/preview', upload.single('file'), (req, res) => {
  const { message } = req.body;
  if (!req.file || !message) return res.status(400).json({ error: 'Missing file or message' });
  try {
    const rows    = parseExcel(req.file.buffer, req.file.originalname);
    const preview = rows.slice(0, 5).map(row => {
      const ks       = Object.keys(row);
      const phoneCol = ks.find(k => /^phone$/i.test(k.trim())) || ks[0];
      return { phone: normalizePhone(row[phoneCol]), message: fillTemplate(message, row) };
    });
    res.json({ preview, total: rows.length });
  } catch (err) {
    res.status(422).json({ error: err.message });
  }
});

// ── Start server ──────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`✅  Server running at http://localhost:${PORT}`);
});

module.exports = app;