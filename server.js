// server.js
const express   = require('express');
const cors      = require('cors');
const multer    = require('multer');
const Papa      = require('papaparse');
const XLSX      = require('xlsx');

const app = express();
app.use(cors());

// allow large payloads
app.use(express.json({ limit: '20mb' }));
app.use(express.urlencoded({ limit: '20mb', extended: true }));

const upload = multer();
let scheduleData = [];
let roomMetadata = [];

/**
 * Upload a term’s schedule CSV.
 * Accepts either multipart/form-data (file upload)
 * or application/json with { password, csv: string }
 */
app.post('/api/schedule/:term', upload.single('file'), (req, res) => {
  const pw = req.body.password;
  if (pw !== 'Upload2025') {
    return res.status(401).json({ error: 'Unauthorized' });
  }

  let csvText;
  if (req.file && req.file.buffer) {
    // file upload path
    csvText = req.file.buffer.toString();
  } else if (req.body.csv) {
    // JSON path
    csvText = req.body.csv;
  } else {
    return res.status(400).json({ error: 'No file or csv payload provided' });
  }

  // parse into JSON rows
  const parsed = Papa.parse(csvText, {
    header: true,
    skipEmptyLines: true
  });

  if (parsed.errors.length) {
    return res.status(400).json({ error: 'CSV parse error', details: parsed.errors });
  }

  scheduleData = parsed.data;
  res.json({ success: true });
});

/**
 * Get the last‐uploaded schedule (ignores :term right now)
 */
app.get(['/api/schedule', '/api/schedule/:term'], (req, res) => {
  res.json(scheduleData);
});

/**
 * Upload room metadata Excel
 */
app.post('/api/rooms/metadata', upload.single('file'), (req, res) => {
  if (!req.file) {
    return res.status(400).json({ error: 'No file provided' });
  }
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
 * Get room metadata
 */
app.get('/api/rooms/metadata', (req, res) => {
  res.json(roomMetadata);
});

// start server
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => {
  console.log(`Backend listening on port ${PORT}`);
});
