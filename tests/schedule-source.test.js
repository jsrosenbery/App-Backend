const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const dataRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'cos-schedule-source-'));
process.env.DATA_DIR = path.join(dataRoot, 'cos-app');
process.env.UPLOAD_PASSWORD = 'Upload2025';
process.env.DEV_PASSWORD = 'DevSecret';
const { app } = require('../server');

test.after(() => fs.rmSync(dataRoot, { recursive: true, force: true }));

function listen() {
  return new Promise(resolve => {
    const server = app.listen(0, () => resolve(server));
  });
}

async function request(baseUrl, pathname, options = {}) {
  const response = await fetch(`${baseUrl}${pathname}`, { ...options, headers: { 'Content-Type': 'application/json', ...(options.headers || {}) } });
  return { response, payload: await response.json() };
}

test('current schedule returns the newest upload and authoritative source metadata', async () => {
  const server = await listen();
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  try {
    const older = await request(baseUrl, '/api/section-seating/FALL%202026/current', {
      method: 'POST',
      body: JSON.stringify({ password: 'Upload2025', sourceName: 'misleading-2099.csv', csv: 'CRN,BUILDING,ROOM\n10001,VIS,101\n' })
    });
    await new Promise(resolve => setTimeout(resolve, 10));
    const newer = await request(baseUrl, '/api/section-seating/FALL%202026/current', {
      method: 'POST',
      body: JSON.stringify({ password: 'Upload2025', sourceName: 'Fall 2026 All Columns.csv', csv: 'CRN,BUILDING,ROOM\n20002,VIS,205\n' })
    });
    const loaded = await request(baseUrl, '/api/section-seating/FALL%202026/current');

    assert.equal(older.response.status, 200);
    assert.equal(newer.response.status, 200);
    assert.ok(Date.parse(newer.payload.source.updatedAt) > Date.parse(older.payload.source.updatedAt));
    assert.equal(loaded.payload.data[0].CRN, '20002');
    assert.equal(loaded.payload.source.name, 'Fall 2026 All Columns.csv');
    assert.equal(loaded.payload.source.updatedAt, newer.payload.source.updatedAt);
    const legacyRoute = await request(baseUrl, '/api/schedule/FALL%202026');
    assert.deepEqual(legacyRoute.payload.data, loaded.payload.data);
    assert.deepEqual(legacyRoute.payload.source, loaded.payload.source);
  } finally {
    server.close();
  }
});

test('analytics archive data cannot override the current schedule source', async () => {
  const server = await listen();
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  try {
    await request(baseUrl, '/api/section-seating/SPRING%202027/current', {
      method: 'POST',
      body: JSON.stringify({ password: 'Upload2025', sourceName: 'Spring 2027 All Columns.csv', csv: 'CRN,BUILDING,ROOM\n30003,VIS,301\n' })
    });
    await request(baseUrl, '/api/analytics-archive/SPRING%202027', {
      method: 'POST',
      body: JSON.stringify({ password: 'Upload2025', csv: 'CRN,BUILDING,ROOM\n99999,VIS,999\n' })
    });
    const loaded = await request(baseUrl, '/api/section-seating/SPRING%202027/current');
    assert.equal(loaded.payload.data[0].CRN, '30003');
    assert.equal(loaded.payload.source.kind, 'section-seating');
  } finally {
    server.close();
  }
});
