// server.js — WhatsApp Bulk Sender
// ─────────────────────────────────────────────────────────────────────────────
// npm install express multer@2 xlsx axios cors cloudinary sharp
//
// Required Environment Variables:
//   CLOUDINARY_CLOUD_NAME
//   CLOUDINARY_API_KEY
//   CLOUDINARY_API_SECRET
//   PORT (optional, defaults to 3000)

const express    = require('express');
const multer     = require('multer');
const XLSX       = require('xlsx');
const axios      = require('axios');
const cors       = require('cors');
const path       = require('path');
const sharp      = require('sharp');
const cloudinary = require('cloudinary').v2;

const app  = express();
const PORT = process.env.PORT || 3000;

// ── Cloudinary Config (from environment variables) ───────────────────────────
cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key:    process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

// ── Middleware ────────────────────────────────────────────────────────────────
app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

// multer v2 — memory storage, 50 MB limit
const { memoryStorage } = require('multer');
const upload = multer({
  storage: memoryStorage(),
  limits: { fileSize: 50 * 1024 * 1024 },
});

// ── Utility: Upload buffer to Cloudinary ─────────────────────────────────────
function uploadToCloudinary(buffer, mimetype) {
  return new Promise((resolve, reject) => {
    let resourceType = 'raw';
    if (mimetype.startsWith('image/')) resourceType = 'image';
    if (mimetype.startsWith('video/')) resourceType = 'video';
    if (mimetype.startsWith('audio/')) resourceType = 'video'; // Cloudinary uses 'video' for audio

    const stream = cloudinary.uploader.upload_stream(
      { resource_type: resourceType, folder: 'wa-bulk-sender' },
      (error, result) => {
        if (error) return reject(error);
        resolve(result);
      }
    );
    stream.end(buffer);
  });
}

// ── Utility: Parse Excel / CSV buffer ────────────────────────────────────────
function parseExcel(buffer, originalName) {
  const wb  = XLSX.read(buffer, { type: 'buffer' });
  const ws  = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

// ── Utility: Replace {{placeholders}} with row values ────────────────────────
function fillTemplate(template, row) {
  return template.replace(/\{\{(\w+)\}\}/g, (_, key) => row[key] ?? '');
}

// ── Utility: Normalize phone number ──────────────────────────────────────────
function normalizePhone(raw) {
  let phone = String(raw).replace(/[\s\-().]/g, '');
  if (!phone.startsWith('+')) phone = '+' + phone;
  return phone;
}

// ── Utility: Get correct Maytapi type string ──────────────────────────────────
// Maytapi valid types: 'text' | 'image' | 'video' | 'audio' | 'pdf' | 'doc'
function getMaytapiType(mimetype) {
  if (!mimetype)                       return 'text';
  if (mimetype.startsWith('image/'))   return 'image';
  if (mimetype.startsWith('video/'))   return 'video';
  if (mimetype.startsWith('audio/'))   return 'audio';
  if (mimetype === 'application/pdf')  return 'pdf';
  return 'doc'; // Word, Excel, PowerPoint, etc.
}

// ── Utility: Send one WhatsApp message via Maytapi ───────────────────────────
// Correct Maytapi payload:
//   Text:  { to_number, type: 'text', message }
//   Media: { to_number, type: 'image'|'video'|'audio'|'pdf'|'doc', url, message (caption) }
async function sendWhatsApp({ productId, phoneId, apiToken, to, message, mediaUrl, mediaMime }) {
  const endpoint = `https://api.maytapi.com/api/${productId}/${phoneId}/sendMessage`;

  let payload;
  if (mediaUrl) {
    const type = getMaytapiType(mediaMime);
    payload = {
      to_number: to,
      type,
      url:     mediaUrl,    // Cloudinary public URL
      text: message || '', // 'text' is Maytapi's caption field for media
    };
  } else {
    payload = {
      to_number: to,
      type:      'text',
      message,
    };
  }

  console.log(`Sending to: ${to} | type: ${payload.type} | url: ${mediaUrl || '—'}`);

  const response = await axios.post(endpoint, payload, {
    headers: {
      'Content-Type':  'application/json',
      'x-maytapi-key': apiToken,
    },
    timeout: 30000,
  });

  console.log(`Maytapi response [${to}]:`, JSON.stringify(response.data));
  return response.data;
}

// ── Utility: Sleep ────────────────────────────────────────────────────────────
const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

// ─────────────────────────────────────────────────────────────────────────────
// Route: GET /
// ─────────────────────────────────────────────────────────────────────────────
app.get('/', (req, res) => {
  res.sendFile(path.join(__dirname, 'public', 'index.html'));
});

// ─────────────────────────────────────────────────────────────────────────────
// Route: POST /api/columns
// ─────────────────────────────────────────────────────────────────────────────
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

// ─────────────────────────────────────────────────────────────────────────────
// Route: POST /api/send  (SSE streaming)
// Form fields:
//   excel     — Excel/CSV file
//   media     — (optional) image / video / audio / PDF / doc
//   productId, phoneId, apiToken, message, delay
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/send',
  upload.fields([
    { name: 'excel', maxCount: 1 },
    { name: 'media', maxCount: 1 },
  ]),
  async (req, res) => {
    const { productId, phoneId, apiToken, message, delay: rawDelay } = req.body;
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

    // ── Upload media to Cloudinary once, reuse URL for all rows ──────────────
    let mediaUrl  = null;
    let mediaMime = null;

    if (mediaFile) {
      try {
        let fileBuffer = mediaFile.buffer;
        mediaMime      = mediaFile.mimetype;

        // Compress images before upload
        if (mediaFile.mimetype.startsWith('image/')) {
          console.log('Compressing image, original size:', fileBuffer.length, 'bytes');
          fileBuffer = await sharp(fileBuffer)
            .resize(1280, 1280, { fit: 'inside', withoutEnlargement: true })
            .jpeg({ quality: 70 })
            .toBuffer();
          mediaMime = 'image/jpeg';
          console.log('Compressed size:', fileBuffer.length, 'bytes');
        }

        console.log('Uploading to Cloudinary...');
        const result = await uploadToCloudinary(fileBuffer, mediaMime);
        mediaUrl     = result.secure_url;
        console.log('Cloudinary URL:', mediaUrl);
      } catch (err) {
        return res.status(500).json({ error: 'Media upload failed: ' + err.message });
      }
    }

    // ── SSE setup ─────────────────────────────────────────────────────────────
    res.setHeader('Content-Type',  'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection',    'keep-alive');
    res.flushHeaders();

    const emit  = (obj) => res.write(`data:${JSON.stringify(obj)}\n\n`);
    const total = rows.length;
    let sent = 0, failed = 0;

    // ── Send loop ─────────────────────────────────────────────────────────────
    for (let i = 0; i < rows.length; i++) {
      const row             = rows[i];
      const phone           = normalizePhone(row[phoneCol]);
      const personalizedMsg = fillTemplate(message, row);

      try {
        await sendWhatsApp({
          productId, phoneId, apiToken,
          to:       phone,
          message:  personalizedMsg,
          mediaUrl,
          mediaMime,
        });
        sent++;
        emit({ type: 'progress', sent, total, failed, phone, index: i + 1 });
      } catch (err) {
        failed++;
        const reason = err.response?.data?.message || err.message || 'Unknown error';
        console.error(`Failed for ${phone}:`, reason);
        emit({ type: 'error', phone, reason, index: i + 1 });
      }

      if (i < rows.length - 1) await sleep(delay);
    }

    emit({ type: 'done', sent, failed, total });
    res.end();
  }
);

// ─────────────────────────────────────────────────────────────────────────────
// Route: POST /api/preview  — dry run, no messages sent
// ─────────────────────────────────────────────────────────────────────────────
app.post('/api/preview', upload.single('file'), (req, res) => {
  const { message } = req.body;
  if (!req.file || !message) return res.status(400).json({ error: 'Missing file or message' });
  try {
    const rows    = parseExcel(req.file.buffer, req.file.originalname);
    const preview = rows.slice(0, 5).map(row => {
      const ks       = Object.keys(row);
      const phoneCol = ks.find(k => /^phone$/i.test(k.trim())) || ks[0];
      return {
        phone:   normalizePhone(row[phoneCol]),
        message: fillTemplate(message, row),
      };
    });
    res.json({ preview, total: rows.length });
  } catch (err) {
    res.status(422).json({ error: err.message });
  }
});

// ── Start server ──────────────────────────────────────────────────────────────
app.listen(PORT, () => {
  console.log(`✅  Server running at http://localhost:${PORT}`);
});

module.exports = app;
