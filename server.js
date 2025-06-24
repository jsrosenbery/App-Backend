// server.js
const express   = require('express');
const cors      = require('cors');
const multer    = require('multer');
const Papa      = require('papaparse');
const XLSX      = require('xlsx');

const app = express();
app.use(cors());

// Increase body‐parser limits so large payloads don’t 413
app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ limit: '20mb', extended: true }));

const upload = multer();
let scheduleData = [];
let roomMetadata = [];

/**
 * Upload a term’s schedule CSV.
 * Matches POST /api/schedule/Spring2026 (or any :term).
 */
app.post('/api/schedule/:term', upload.single('file'), (req, res) => {
  const pw = req.body.password;
  if (pw !== 'Upload2025') {
    return res.status(401).json({ error: 'Unauthorized' });
  }
  // Parse CSV buffer into JSON rows
  const parsed = Papa.parse(req.file.buffer.toString(), {
    header: true,
    skipEmptyLines: true
  });
  scheduleData = parsed.data;
  res.json({ success: true });
});

/**
 * Retrieve the last‐uploaded schedule.
 * Supports both /api/schedule and /api/schedule/:term
 */
app.get(['/api/schedule', '/api/schedule/:term'], (req, res) => {
  res.json(scheduleData);
});

/**
 * Upload room metadata Excel.
 */
app.post('/api/rooms/metadata', upload.single('file'), (req, res) => {
  const wb  = XLSX.read(req.file.buffer, { type: 'buffer' });
  const raw = XLSX.utils
    .sheet_to_json(wb.Sheets[wb.SheetNames[0]], { range: 3 });
  roomMetadata = raw.map(r => ({
    campus:   r.Campus,
    building: r.Building,
    room:     String(r['Room Number']),
    type:     r.Type,
    capacity: Number(r['# of Desks in Room'])
  }));
  res.json({ success: true });
});

/**
 * Retrieve room metadata.
 */
app.get('/api/rooms/metadata', (req, res) => {
  res.json(roomMetadata);
});

// Start the server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Backend listening on port ${PORT}`);
});
