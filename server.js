const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const Papa = require('papaparse');
const app = express();
const PORT = process.env.PORT || 3000;

// Build your list of allowed origins
const allowedOrigins = [
  process.env.CORS_ORIGIN,            // staging or prod from env
  'https://cos-app.vercel.app'        // always allow production front-end
].filter(Boolean);

console.log('Allowed CORS origins:', allowedOrigins);

app.use(cors({
  origin: (incomingOrigin, callback) => {
    // allow non-browser/postman/etc requests
    if (!incomingOrigin) return callback(null, true);
    if (allowedOrigins.includes(incomingOrigin)) {
      return callback(null, true);
    }
    callback(new Error(`Origin ${incomingOrigin} not allowed by CORS`));
  }
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
  if (password !== process.env.UPLOAD_PASSWORD) {
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

app.listen(PORT, () => {
  console.log(`Listening on port ${PORT} in ${process.env.NODE_ENV || 'production'} mode`);
});
