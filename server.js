// server.js

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const Papa = require('papaparse');
const XLSX = require('xlsx');

const app = express();
app.use(cors());
app.use(express.json());

const upload = multer();

let scheduleData = [];
let roomMetadata = [];

// POST /api/schedule
app.post('/api/schedule', upload.single('file'), (req, res) => {
  const password = req.body.password;
  if (password !== 'Upload2025') {
    return res.status(401).json({ error: 'Unauthorized: incorrect password' });
  }
  const csvString = req.file.buffer.toString('utf-8');
  const parsed = Papa.parse(csvString, { header: true, skipEmptyLines: true });
  scheduleData = parsed.data;
  return res.json({ success: true, count: scheduleData.length });
});

// GET /api/schedule and /api/schedule/:term
app.get(['/api/schedule', '/api/schedule/:term'], (req, res) => {
  return res.json(scheduleData);
});

// POST /api/rooms/metadata
app.post('/api/rooms/metadata', upload.single('file'), (req, res) => {
  try {
    const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
    const sheet = workbook.Sheets[workbook.SheetNames[0]];
    const raw = XLSX.utils.sheet_to_json(sheet, { range: 3 });
    roomMetadata = raw.map(r => ({
      campus: r.Campus,
      building: r.Building,
      room: r['Room #'].toString(),
      type: r.Type,
      capacity: Number(r['# of Desks in Room'])
    }));
    return res.json({ success: true, count: roomMetadata.length });
  } catch (err) {
    return res.status(500).json({ error: err.message });
  }
});

// GET /api/rooms/metadata
app.get('/api/rooms/metadata', (req, res) => {
  return res.json(roomMetadata);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
