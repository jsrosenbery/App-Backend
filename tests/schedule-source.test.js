const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const dataRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'cos-schedule-source-'));
process.env.DATA_DIR = path.join(dataRoot, 'cos-app');
process.env.UPLOAD_PASSWORD = 'Upload2025';
process.env.DEV_PASSWORD = 'DevSecret';
const { app, selectNewestValidAllColumnsArchive } = require('../server');

const legacyCsv = 'CRN,BUILDING,ROOM,DAYS,Time\n10001,VIS,101,MW,09:00-09:50\n';
const allColumnsCsv = 'TERM,CRN,SUBJECT,COURSE,BUILDING,ROOM,DAYS,STARTTIME,ENDTIME,CAMPUS,INSTRUCTIONAL_METHOD_CODE,SCHD_CODE_SSRMEET,ACTUAL_ENROLL,MAX_ENROLL\nFALL 2026,20002,CHEM,001,VIS,205,MW,0900,0950,COS,IP,LEC,24,30\n';

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

test('legacy simple current is migrated from a newer valid All Columns archive without changing the archive', async () => {
  const server = await listen();
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  try {
    await request(baseUrl, '/api/section-seating/FALL%202028/current', {
      method: 'POST', body: JSON.stringify({ password: 'Upload2025', sourceName: 'legacy-simple.csv', csv: legacyCsv })
    });
    await new Promise(resolve => setTimeout(resolve, 10));
    const archived = await request(baseUrl, '/api/analytics-archive/FALL%202028', {
      method: 'POST', body: JSON.stringify({ password: 'Upload2025', csv: allColumnsCsv })
    });
    const archivePath = path.join(process.env.DATA_DIR, 'analytics-archive', 'FALL 2028.csv');
    const archiveBefore = fs.readFileSync(archivePath, 'utf8');
    const archiveMtime = fs.statSync(archivePath).mtimeMs;
    const loaded = await request(baseUrl, '/api/section-seating/FALL%202028/current');

    assert.equal(loaded.payload.data[0].CRN, '20002');
    assert.equal(loaded.payload.source.reportType, 'All Columns Section Seating');
    assert.equal(loaded.payload.source.uploadedAt, archived.payload.lastUpdated);
    assert.equal(fs.readFileSync(archivePath, 'utf8'), archiveBefore);
    assert.equal(fs.statSync(archivePath).mtimeMs, archiveMtime);
  } finally {
    server.close();
  }
});

test('missing current is initialized from All Columns archive and migration is idempotent', async () => {
  const server = await listen();
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  try {
    await request(baseUrl, '/api/analytics-archive/SPRING%202029', {
      method: 'POST', body: JSON.stringify({ password: 'Upload2025', csv: allColumnsCsv.replace('FALL 2026', 'SPRING 2029') })
    });
    const first = await request(baseUrl, '/api/section-seating/SPRING%202029/current');
    const currentPath = path.join(process.env.DATA_DIR, 'SPRING 2029.csv');
    const metadataPath = path.join(process.env.DATA_DIR, 'SPRING 2029.schedule-source.json');
    const currentBefore = fs.readFileSync(currentPath, 'utf8');
    const metadataBefore = fs.readFileSync(metadataPath, 'utf8');
    const second = await request(baseUrl, '/api/section-seating/SPRING%202029/current');

    assert.equal(first.payload.data[0].CRN, '20002');
    assert.deepEqual(second.payload.data, first.payload.data);
    assert.equal(fs.readFileSync(currentPath, 'utf8'), currentBefore);
    assert.equal(fs.readFileSync(metadataPath, 'utf8'), metadataBefore);
  } finally {
    server.close();
  }
});

test('valid newer All Columns current is untouched when an older archive exists', async () => {
  const server = await listen();
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  try {
    await request(baseUrl, '/api/analytics-archive/SUMMER%202029', {
      method: 'POST', body: JSON.stringify({ password: 'Upload2025', csv: allColumnsCsv.replace('20002', '30003') })
    });
    await new Promise(resolve => setTimeout(resolve, 10));
    const currentCsv = allColumnsCsv.replace('20002', '40004');
    const current = await request(baseUrl, '/api/section-seating/SUMMER%202029/current', {
      method: 'POST', body: JSON.stringify({ password: 'Upload2025', sourceName: 'new-current.csv', csv: currentCsv })
    });
    const loaded = await request(baseUrl, '/api/section-seating/SUMMER%202029/current');
    assert.equal(loaded.payload.data[0].CRN, '40004');
    assert.equal(loaded.payload.source.name, 'new-current.csv');
    assert.equal(loaded.payload.source.updatedAt, current.payload.source.updatedAt);
    assert.equal(loaded.payload.source.migration, undefined);
  } finally {
    server.close();
  }
});

test('newest valid All Columns archive candidate is selected by authoritative timestamp', () => {
  const selected = selectNewestValidAllColumnsArchive([
    { csv: allColumnsCsv.replace('20002', '50005'), updatedAt: '2029-01-01T00:00:00Z' },
    { csv: legacyCsv, updatedAt: '2030-01-01T00:00:00Z' },
    { csv: allColumnsCsv.replace('20002', '60006'), uploadedAt: '2029-02-01T00:00:00Z' }
  ]);
  assert.match(selected.csv, /60006/);
});
