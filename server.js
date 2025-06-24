// server.js

const express = require('express');
const multer = require('multer');
const XLSX = require('xlsx');
const cors = require('cors');

const app = express();
app.use(cors());
app.use(express.json());

let roomMetadata = [];
let scheduleData = [];

// Multer setup
const upload = multer();

// Upload room metadata
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
    res.json({ success: true, count: roomMetadata.length });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

// Get room metadata
app.get('/api/rooms/metadata', (req, res) => {
  res.json(roomMetadata);
});

// Placeholder: schedule upload endpoint
app.post('/api/schedule', upload.single('file'), (req, res) => {
  // TODO: parse CSV and load into scheduleData
  res.json({ success: true });
});

// Get schedule data
app.get('/api/schedule', (req, res) => {
  res.json(scheduleData);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server listening on port ${PORT}`));
