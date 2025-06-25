const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const Papa = require('papaparse');
const app = express();
const PORT = process.env.PORT || 3000;

// Enable CORS for your frontend
app.use(cors({
  origin: 'https://cos-app.vercel.app'
}));

app.use(express.json({ limit: '50mb' }));

// Directory to store CSV files
const DATA_DIR = path.join(__dirname, 'schedules');
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}

// POST endpoint to upload schedule CSV
app.post('/api/schedule/:term', (req, res) => {
  const term = req.params.term;
  const { csv, password } = req.body;
  if (password !== 'Upload2025') {
    return res.status(403).json({ error: 'Unauthorized' });
  }

  const filePath = path.join(DATA_DIR, `${term}.csv`);
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
  const filePath = path.join(DATA_DIR, `${term}.csv`);
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

app.listen(PORT, () => console.log(`Listening on port ${PORT}`));
