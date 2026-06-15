const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const Papa = require('papaparse');
const crypto = require('crypto');
const { execFile } = require('child_process');
const app = express();
const PORT = process.env.PORT || 3000;
const CORS_ORIGIN = process.env.CORS_ORIGIN || 'https://cos-app.vercel.app';
const UPLOAD_PASSWORD = process.env.UPLOAD_PASSWORD || 'Upload2025';

// Enable CORS for your frontend
app.use(cors({
  origin: CORS_ORIGIN
}));

app.use(express.json({ limit: '50mb' }));

// Directory to store uploaded data. On hosted environments, set DATA_DIR or
// SCHEDULE_DATA_DIR to a mounted persistent disk path so imports survive restarts.
const DEFAULT_DATA_DIR = fs.existsSync('/var/data')
  ? path.join('/var/data', 'cos-app')
  : path.join(__dirname, 'schedules');
const DATA_DIR = path.resolve(process.env.DATA_DIR || process.env.SCHEDULE_DATA_DIR || DEFAULT_DATA_DIR);
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}
const ROOM_CATALOG_PATH = path.join(DATA_DIR, 'rooms.json');
const MODALITY_DEFINITIONS_PATH = path.join(DATA_DIR, 'modalities.json');
const CONVERT_DIR = path.join(DATA_DIR, 'conversions');
if (!fs.existsSync(CONVERT_DIR)) {
  fs.mkdirSync(CONVERT_DIR, { recursive: true });
}

const DEFAULT_MODALITY_DEFINITIONS = [
  { code: 'IP', modality: 'In Person', omitted: false },
  { code: 'ONL', modality: 'Online', omitted: false },
  { code: 'HYB', modality: 'Hybrid', omitted: false },
  { code: 'DE', modality: 'Dual Enrollment', omitted: false },
  { code: 'FLX', modality: 'Flex', omitted: false },
  { code: '02S', modality: 'In Person', omitted: false },
  { code: '022', modality: 'In Person', omitted: false },
  { code: 'OL', modality: 'Online', omitted: false },
  { code: 'ONN', modality: 'Online', omitted: false },
  { code: 'O1', modality: 'Online', omitted: false },
  { code: 'ONS', modality: 'Online', omitted: false },
  { code: '02N', modality: 'In Person', omitted: false },
  { code: 'CPL', modality: 'Omitted from modality analysis', omitted: true },
  { code: '20', modality: 'Omitted from modality analysis', omitted: true }
];

function getSchedulePath(term) {
  if (!/^[a-z0-9 _-]+$/i.test(term)) return null;
  const filePath = path.resolve(DATA_DIR, `${term}.csv`);
  const dataRoot = path.resolve(DATA_DIR) + path.sep;
  return filePath.startsWith(dataRoot) ? filePath : null;
}

function isAuthorized(password) {
  if (!UPLOAD_PASSWORD) return false;
  if (typeof password !== 'string') return false;
  const expected = Buffer.from(UPLOAD_PASSWORD);
  const supplied = Buffer.from(password);
  return expected.length === supplied.length && crypto.timingSafeEqual(expected, supplied);
}

function normalizeRoomCatalog(rooms) {
  if (!Array.isArray(rooms)) return null;
  const normalized = [];
  for (const item of rooms) {
    if (!item || typeof item !== 'object') continue;
    const campus = String(item.campus || item.Campus || '').trim();
    const building = String(item.building || item.Building || '').trim();
    const room = String(item.room || item.Room || '').trim();
    const type = String(item.type || item.Type || item.roomType || item['Room Type'] || '').trim();
    const rawCapacity = item.capacity ?? item.Capacity ?? item.cap ?? item.Cap;
    const capacity = rawCapacity === '' || rawCapacity == null ? null : Number(rawCapacity);
    if (!building || !room) continue;
    normalized.push({
      campus,
      building,
      room,
      buildingRoom: `${building}-${room}`,
      type,
      capacity: Number.isFinite(capacity) ? capacity : null
    });
  }
  return normalized;
}

function normalizeModalityDefinitions(definitions) {
  if (!Array.isArray(definitions)) return null;
  const normalized = [];
  for (const item of definitions) {
    if (!item || typeof item !== 'object') continue;
    const code = String(item.code || item.Code || item.instructionalMethod || item['Instructional Method'] || '').trim().toUpperCase();
    const modality = String(item.modality || item.Modality || item.category || item.Category || '').trim();
    const rawOmitted = item.omitted ?? item.Omitted ?? item.omit ?? item.Omit ?? item.exclude ?? item.Exclude;
    const omitted = rawOmitted === true || String(rawOmitted || '').trim().toLowerCase() === 'true' || String(rawOmitted || '').trim().toLowerCase() === 'yes' || String(rawOmitted || '').trim() === '1';
    if (!code || (!omitted && !modality)) continue;
    normalized.push({
      code,
      modality: omitted ? (modality || 'Omitted from modality analysis') : modality,
      omitted
    });
  }
  return normalized;
}

function readRoomCatalog() {
  if (!fs.existsSync(ROOM_CATALOG_PATH)) {
    return { lastUpdated: null, data: [] };
  }
  const json = fs.readFileSync(ROOM_CATALOG_PATH, 'utf8');
  const stats = fs.statSync(ROOM_CATALOG_PATH);
  return {
    lastUpdated: stats.mtime.toISOString(),
    data: JSON.parse(json)
  };
}

function readModalityDefinitions() {
  if (!fs.existsSync(MODALITY_DEFINITIONS_PATH)) {
    return { lastUpdated: null, data: DEFAULT_MODALITY_DEFINITIONS };
  }
  const json = fs.readFileSync(MODALITY_DEFINITIONS_PATH, 'utf8');
  const stats = fs.statSync(MODALITY_DEFINITIONS_PATH);
  return {
    lastUpdated: stats.mtime.toISOString(),
    data: JSON.parse(json)
  };
}

function safeFilename(name, fallback) {
  const clean = String(name || '').replace(/[^a-z0-9_.-]/gi, '_').slice(0, 120);
  return clean || fallback;
}

function runCommand(command, args, options = {}) {
  return new Promise((resolve, reject) => {
    execFile(command, args, options, (error, stdout, stderr) => {
      if (error) {
        error.stdout = stdout;
        error.stderr = stderr;
        reject(error);
        return;
      }
      resolve({ stdout, stderr });
    });
  });
}

async function convertDocxToPdf(inputPath, outputDir) {
  const commands = [
    process.env.LIBREOFFICE_PATH,
    'soffice',
    'libreoffice'
  ].filter(Boolean);
  let lastError = null;
  for (const command of commands) {
    try {
      await runCommand(command, [
        '--headless',
        '--convert-to',
        'pdf',
        '--outdir',
        outputDir,
        inputPath
      ], { timeout: 30000 });
      const outputPath = path.join(outputDir, `${path.basename(inputPath, path.extname(inputPath))}.pdf`);
      if (fs.existsSync(outputPath)) return outputPath;
      lastError = new Error('LibreOffice finished without producing a PDF.');
    } catch (err) {
      lastError = err;
    }
  }
  const detail = lastError?.stderr || lastError?.message || 'LibreOffice/soffice was not found.';
  throw new Error(`DOCX-to-PDF converter unavailable or failed. ${detail}`);
}

// POST endpoint to upload schedule CSV
app.post('/api/schedule/:term', (req, res) => {
  const term = req.params.term;
  const { csv, password } = req.body;
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  if (typeof csv !== 'string') {
    return res.status(400).json({ error: 'CSV payload is required' });
  }

  const filePath = getSchedulePath(term);
  if (!filePath) {
    return res.status(400).json({ error: 'Invalid term' });
  }
  try {
    fs.writeFileSync(filePath, csv);
    const now = new Date().toISOString();
    return res.json({ success: true, lastUpdated: now });
  } catch (err) {
    console.error('Write error:', err);
    return res.status(500).json({ error: 'File write failed' });
  }
});

// GET endpoint to fetch and parse schedule CSV
app.get('/api/schedule/:term', (req, res) => {
  const term = req.params.term;
  const filePath = getSchedulePath(term);
  if (!filePath) {
    return res.status(400).json({ error: 'Invalid term' });
  }
  if (!fs.existsSync(filePath)) {
    return res.json({ lastUpdated: null, data: [] });
  }

  try {
    const csv = fs.readFileSync(filePath, 'utf8');
    const stats = fs.statSync(filePath);
    const lastUpdated = stats.mtime.toISOString();
    const parsed = Papa.parse(csv, { header: true, skipEmptyLines: true });
    return res.json({ lastUpdated, data: parsed.data });
  } catch (err) {
    console.error('Read error:', err);
    return res.status(500).json({ error: 'File read failed' });
  }
});

app.get('/api/rooms', (req, res) => {
  try {
    return res.json(readRoomCatalog());
  } catch (err) {
    console.error('Room catalog read error:', err);
    return res.status(500).json({ error: 'Room catalog read failed' });
  }
});

app.post('/api/rooms/export', (req, res) => {
  const { password } = req.body || {};
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  try {
    return res.json(readRoomCatalog());
  } catch (err) {
    console.error('Room catalog export error:', err);
    return res.status(500).json({ error: 'Room catalog export failed' });
  }
});

app.post('/api/rooms/import', (req, res) => {
  const { password, rooms } = req.body || {};
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const normalized = normalizeRoomCatalog(rooms);
  if (!normalized || !normalized.length) {
    return res.status(400).json({ error: 'Room catalog payload is required' });
  }
  try {
    fs.writeFileSync(ROOM_CATALOG_PATH, JSON.stringify(normalized, null, 2));
    const now = new Date().toISOString();
    return res.json({ success: true, lastUpdated: now, count: normalized.length, data: normalized });
  } catch (err) {
    console.error('Room catalog write error:', err);
    return res.status(500).json({ error: 'Room catalog write failed' });
  }
});

app.get('/api/modalities', (req, res) => {
  try {
    return res.json(readModalityDefinitions());
  } catch (err) {
    console.error('Modality definitions read error:', err);
    return res.status(500).json({ error: 'Modality definitions read failed' });
  }
});

app.post('/api/modalities/import', (req, res) => {
  const { password, definitions } = req.body || {};
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const normalized = normalizeModalityDefinitions(definitions);
  if (!normalized || !normalized.length) {
    return res.status(400).json({ error: 'Modality definitions payload is required' });
  }
  try {
    fs.writeFileSync(MODALITY_DEFINITIONS_PATH, JSON.stringify(normalized, null, 2));
    const now = new Date().toISOString();
    return res.json({ success: true, lastUpdated: now, count: normalized.length, data: normalized });
  } catch (err) {
    console.error('Modality definitions write error:', err);
    return res.status(500).json({ error: 'Modality definitions write failed' });
  }
});

app.post('/api/convert/docx-to-pdf', async (req, res) => {
  const { filename, docxBase64 } = req.body || {};
  if (typeof docxBase64 !== 'string' || !docxBase64.trim()) {
    return res.status(400).send('DOCX payload is required');
  }

  const requestId = crypto.randomBytes(8).toString('hex');
  const requestDir = path.join(CONVERT_DIR, requestId);
  fs.mkdirSync(requestDir, { recursive: true });
  const inputName = safeFilename(filename, 'schedule-change-form.docx').replace(/\.pdf$/i, '.docx');
  const inputPath = path.join(requestDir, inputName.toLowerCase().endsWith('.docx') ? inputName : `${inputName}.docx`);

  try {
    fs.writeFileSync(inputPath, Buffer.from(docxBase64, 'base64'));
    const pdfPath = await convertDocxToPdf(inputPath, requestDir);
    const downloadName = path.basename(inputPath, path.extname(inputPath)) + '.pdf';
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${downloadName}"`);
    return res.sendFile(pdfPath, err => {
      try {
        fs.rmSync(requestDir, { recursive: true, force: true });
      } catch (cleanupErr) {
        console.error('Conversion cleanup error:', cleanupErr);
      }
      if (err) console.error('PDF send error:', err);
    });
  } catch (err) {
    try {
      fs.rmSync(requestDir, { recursive: true, force: true });
    } catch (cleanupErr) {
      console.error('Conversion cleanup error:', cleanupErr);
    }
    console.error('DOCX conversion error:', err);
    return res.status(500).send(err.message || 'DOCX-to-PDF conversion failed');
  }
});

app.listen(PORT, () => console.log(`Listening on port ${PORT}`));
