// server.js — WhatsApp Bulk Sender Backend
const express    = require('express');
const multer     = require('multer');
const XLSX       = require('xlsx');
const axios      = require('axios');
const cors       = require('cors');
const path       = require('path');
const cloudinary = require('cloudinary').v2;

const app  = express();
const PORT = process.env.PORT || 3000;


cloudinary.config({
  cloud_name: process.env.CLOUDINARY_CLOUD_NAME,
  api_key:    process.env.CLOUDINARY_API_KEY,
  api_secret: process.env.CLOUDINARY_API_SECRET,
});

app.use(cors());
app.use(express.json());
app.use(express.static(path.join(__dirname, 'public')));

const { memoryStorage } = require('multer');
const upload = multer({ storage: memoryStorage(), limits: { fileSize: 50 * 1024 * 1024 } });

function parseExcel(buffer) {
  const wb = XLSX.read(buffer, { type: 'buffer' });
  const ws = wb.Sheets[wb.SheetNames[0]];
  return XLSX.utils.sheet_to_json(ws, { defval: '' });
}

function fillTemplate(template, row) {
  return template.replace(/\{\{(\w+)\}\}/g, (_, key) => row[key] ?? '');
}

function normalizePhone(raw) {
  let phone = String(raw).replace(/[\s\-().]/g, '');
  if (!phone.startsWith('+')) phone = '+' + phone;
  return phone;
}

function maytapiType(mimetype) {
  if (!mimetype) return 'text';
  if (mimetype.startsWith('image/')) return 'image';
  if (mimetype.startsWith('video/')) return 'video';
  if (mimetype.startsWith('audio/')) return 'audio';
  return 'media';
}

async function uploadToCloudinary(buffer, mimetype) {
  return new Promise((resolve, reject) => {
    const resourceType = mimetype.startsWith('video/') ? 'video'
                       : mimetype.startsWith('image/') ? 'image'
                       : 'raw';
    const stream = cloudinary.uploader.upload_stream(
      { resource_type: resourceType, folder: 'whatsapp-bulk' },
      (error, result) => error ? reject(error) : resolve(result)
    );
    stream.end(buffer);
  });
}

function buildPayload({ to, message, mediaUrl, mediaMime }) {
  if (!mediaUrl) return { to_number: to, type: 'text', message };
  return { to_number: to, type: maytapiType(mediaMime), message, media: mediaUrl };
}

async function sendWhatsApp({ productId, phoneId, apiToken, to, message, mediaUrl, mediaMime }) {
  const url     = `https://api.maytapi.com/api/${productId}/${phoneId}/sendMessage`;
  const payload = buildPayload({ to, message, mediaUrl, mediaMime });
  console.log('Sending to:', to, '| Payload:', JSON.stringify(payload));
  const response = await axios.post(url, payload, {
    headers: { 'Content-Type': 'application/json', 'x-maytapi-key': apiToken },
    timeout: 30000,
  });
  console.log('Maytapi response:', JSON.stringify(response.data));
  if (response.data?.success === false) throw new Error(response.data?.message || 'Maytapi error');
  return response.data;
}

const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

app.get('/', (req, res) => res.sendFile(path.join(__dirname, 'public', 'index.html')));

app.post('/api/columns', upload.single('file'), (req, res) => {
  if (!req.file) return res.status(400).json({ error: 'No file uploaded' });
  try {
    const rows    = parseExcel(req.file.buffer);
    const columns = rows.length > 0 ? Object.keys(rows[0]) : [];
    res.json({ columns, rowCount: rows.length });
  } catch (err) {
    res.status(422).json({ error: 'Failed to parse file: ' + err.message });
  }
});

app.post('/api/send',
  upload.fields([{ name: 'excel', maxCount: 1 }, { name: 'media', maxCount: 1 }]),
  async (req, res) => {
    const { productId, phoneId, apiToken, message, delay: rawDelay } = req.body;
    const delay = Math.max(500, parseInt(rawDelay) || 1500);

    if (!productId || !phoneId || !apiToken || !message)
      return res.status(400).json({ error: 'Missing required fields' });

    const excelFile = req.files?.excel?.[0];
    const mediaFile = req.files?.media?.[0];
    if (!excelFile) return res.status(400).json({ error: 'No Excel file uploaded' });

    let rows;
    try { rows = parseExcel(excelFile.buffer); }
    catch (err) { return res.status(422).json({ error: 'Failed to parse Excel: ' + err.message }); }
    if (rows.length === 0) return res.status(422).json({ error: 'Excel file is empty' });

    const keys     = Object.keys(rows[0]);
    const phoneCol = keys.find(k => /^phone$/i.test(k.trim()))
                  || keys.find(k => /phone|mobile|number|whatsapp/i.test(k))
                  || keys[0];

    // ✅ Upload to Cloudinary — permanent public URL
    let mediaUrl = null;
    if (mediaFile) {
      try {
        console.log('Uploading to Cloudinary...');
        const result = await uploadToCloudinary(mediaFile.buffer, mediaFile.mimetype);
        mediaUrl = result.secure_url;
        console.log('Cloudinary URL:', mediaUrl);
      } catch (err) {
        return res.status(500).json({ error: 'Media upload failed: ' + err.message });
      }
    }

    res.setHeader('Content-Type', 'text/event-stream');
    res.setHeader('Cache-Control', 'no-cache');
    res.setHeader('Connection', 'keep-alive');
    res.flushHeaders();

    const emit  = (obj) => res.write(`data:${JSON.stringify(obj)}\n\n`);
    const total = rows.length;
    let sent = 0, failed = 0;

    for (let i = 0; i < rows.length; i++) {
      const row             = rows[i];
      const phone           = normalizePhone(row[phoneCol]);
      const personalizedMsg = fillTemplate(message, row);
      try {
        await sendWhatsApp({ productId, phoneId, apiToken, to: phone, message: personalizedMsg, mediaUrl, mediaMime: mediaFile?.mimetype || null });
        sent++;
        emit({ type: 'progress', sent, total, failed, phone, index: i + 1 });
      } catch (err) {
        failed++;
        emit({ type: 'error', phone, reason: err.response?.data?.message || err.message, index: i + 1 });
      }
      if (i < rows.length - 1) await sleep(delay);
    }

    emit({ type: 'done', sent, failed, total });
    res.end();
  }
);

app.post('/api/preview', upload.single('file'), (req, res) => {
  const { message } = req.body;
  if (!req.file || !message) return res.status(400).json({ error: 'Missing file or message' });
  try {
    const rows    = parseExcel(req.file.buffer);
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

app.listen(PORT, () => console.log(`Server running at http://localhost:${PORT}`));
module.exports = app;
