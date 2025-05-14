// server.js
const express = require("express");
const cors = require("cors");
const fs = require("fs");
const path = require("path");
const Papa = require("papaparse");

const app = express();
const PORT = process.env.PORT || 3001;
const UPLOAD_PASSWORD = "Upload2025";

app.use(cors());
app.use(express.json({ limit: "10mb" }));
app.use(express.urlencoded({ extended: true, limit: "10mb" }));

const dataDir = path.join(__dirname, "data");
if (!fs.existsSync(dataDir)) fs.mkdirSync(dataDir);

// Get schedule for a specific term
app.get("/api/schedule/:term", (req, res) => {
  const filePath = path.join(dataDir, `${req.params.term}.json`);
  if (!fs.existsSync(filePath)) return res.json([]);
  const data = fs.readFileSync(filePath);
  res.json(JSON.parse(data));
});

// Upload and replace schedule for a specific term
app.post("/api/schedule/:term", (req, res) => {
  const { password, csv } = req.body;
  if (password !== UPLOAD_PASSWORD) {
    return res.status(403).json({ error: "Invalid password" });
  }

  try {
    const parsed = Papa.parse(csv, { header: true, skipEmptyLines: true });
    const cleaned = parsed.data
      .filter((r) => {
        const room = (r.ROOM || "").trim().toUpperCase();
        const building = (r.BUILDING || "").trim().toUpperCase();
        return (
          room && !["N/A", "ONLINE", "LIVE"].includes(room) &&
          building && !["ONLINE", "LIVE"].includes(building) &&
          r.Time && r.DAYS
        );
      })
      .map((r) => {
        const [start, end] = r.Time.split(" - ");
        return {
          Term: req.params.term,
          Subject_Course: r.Subject_Course,
          CRN: r.CRN,
          DAYS: r.DAYS,
          Start_Time: start.trim(),
          End_Time: end.trim(),
          Room_ID: `${r.BUILDING.trim()} - ${r.ROOM.trim()}`,
        };
      });

    const outPath = path.join(dataDir, `${req.params.term}.json`);
    fs.writeFileSync(outPath, JSON.stringify(cleaned, null, 2));
    res.json({ success: true });
  } catch (e) {
    res.status(500).json({ error: "Failed to parse or save data" });
  }
});

app.listen(PORT, () => console.log(`Server running on port ${PORT}`));
