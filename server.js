const express = require("express");
const bodyParser = require("body-parser");
const cors = require("cors");
const Papa = require("papaparse");

const app = express();
const port = process.env.PORT || 3000;

app.use(cors());
app.use(bodyParser.json({ limit: "10mb" }));

// In-memory schedule store: { [term]: { csv, lastUpdated } }
const scheduleByTerm = {};

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

  scheduleByTerm[term] = {
    csv,
    lastUpdated: new Date().toISOString()
  };

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
