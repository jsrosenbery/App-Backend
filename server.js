const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const Papa = require("papaparse");
const fs = require("fs");
const path = require("path");

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json({ limit: "10mb" }));

// Directory for schedules
const scheduleDir = path.join(__dirname, "schedules");
if (!fs.existsSync(scheduleDir)) {
  fs.mkdirSync(scheduleDir);
}

// In-memory schedule store: { [term]: { csv, lastUpdated } }
const scheduleByTerm = {};

// Helper: Save schedule to disk
function saveSchedule(term, csv, lastUpdated) {
  const filePath = path.join(scheduleDir, `${term}.json`);
  fs.writeFileSync(filePath, JSON.stringify({ csv, lastUpdated }), "utf8");
}

// Helper: Load all schedules from disk at startup
function loadAllSchedules() {
  const files = fs.readdirSync(scheduleDir);
  files.forEach(file => {
    if (file.endsWith(".json")) {
      const term = file.replace(".json", "");
      const data = JSON.parse(fs.readFileSync(path.join(scheduleDir, file), "utf8"));
      scheduleByTerm[term] = data;
    }
  });
}

// Load schedules at server start
loadAllSchedules();

// List available terms
app.get("/api/terms", (req, res) => {
  res.json(Object.keys(scheduleByTerm));
});

// Upload schedule for a term (protected)
app.post("/api/schedule/:term", (req, res) => {
  const { csv, password } = req.body;
  const term = req.params.term;

  if (password !== "Upload2025") {
    return res.status(403).json({ error: "Unauthorized" });
  }

  const lastUpdated = new Date().toISOString();
  scheduleByTerm[term] = { csv, lastUpdated };
  saveSchedule(term, csv, lastUpdated);

  res.json({ success: true });
});

// Get parsed schedule for a term with lastUpdated
app.get("/api/schedule/:term", (req, res) => {
  const term = req.params.term;
  const data = scheduleByTerm[term];

  if (!data) {
    return res.json({ lastUpdated: null, data: [] });
  }

  const parsed = Papa.parse(data.csv, { header: true, skipEmptyLines: true });

  res.json({
    lastUpdated: data.lastUpdated,
    data: parsed.data
  });
});

app.listen(port, () => {
  console.log(`Server running on port ${port}`);
});
