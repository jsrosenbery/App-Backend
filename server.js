// server.js

const express = require('express');
const cors = require('cors');
const multer = require('multer');
const XLSX = require('xlsx');

const app = express();
app.use(cors()); // enable CORS for all origins
app.use(express.json());

const upload = multer();

let scheduleData = []; // in-memory; replace with DB if needed
let roomMetadata = [];

// Schedule upload endpoint
app.post('/api/schedule', upload.single('file'), (req, res) => {
  try {
    // parse CSV into scheduleData (implement as needed)
    // scheduleData = parsedData
    res.json(scheduleData);
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/schedule', (req, res) => {
  res.json(scheduleData);
});

// Room metadata endpoints
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
    res.json({ success: true });
  } catch (err) {
    res.status(500).json({ error: err.message });
  }
});

app.get('/api/rooms/metadata', (req, res) => {
  res.json(roomMetadata);
});

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
