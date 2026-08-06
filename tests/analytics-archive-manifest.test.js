const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const dataRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'cos-archive-manifest-'));
const dataDir = path.join(dataRoot, 'cos-app');
fs.mkdirSync(dataDir, { recursive: true });
process.env.DATA_DIR = dataDir;
process.env.UPLOAD_PASSWORD = 'Upload2025';
process.env.DEV_PASSWORD = 'DevSecret';

const { app, displayTermFromArchiveTerm, sortArchiveTermsNewestFirst } = require('../server');

test.after(() => {
  fs.rmSync(dataRoot, { recursive: true, force: true });
});

function listen() {
  return new Promise(resolve => {
    const server = app.listen(0, () => resolve(server));
  });
}

async function jsonRequest(baseUrl, pathname, options = {}) {
  const response = await fetch(`${baseUrl}${pathname}`, {
    ...options,
    headers: {
      'Content-Type': 'application/json',
      ...(options.headers || {})
    }
  });
  const payload = await response.json();
  return { response, payload };
}

const csvA = 'CRN,Course,Enrollment\n10001,COMM C1000,20\n10002,ENGL C1000,18\n';
const csvB = 'CRN,Course,Enrollment\n20001,HIST 017,22\n';

test('Analytics archive manifest returns metadata without full row arrays and preserves old endpoints', async () => {
  const server = await listen();
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  try {
    const savedFall = await jsonRequest(baseUrl, '/api/analytics-archive/FALL 2026', {
      method: 'POST',
      body: JSON.stringify({ password: 'Upload2025', csv: csvA })
    });
    assert.equal(savedFall.response.status, 200);
    assert.equal(savedFall.payload.metadata.rowCount, 2);
    assert.ok(savedFall.payload.metadata.sizeBytes >= csvA.length);

    const savedSpring = await jsonRequest(baseUrl, '/api/analytics-archive/SPRING 2026', {
      method: 'POST',
      body: JSON.stringify({ password: 'Upload2025', csv: csvB })
    });
    assert.equal(savedSpring.response.status, 200);

    const manifest = await jsonRequest(baseUrl, '/api/analytics-archive/manifest');
    assert.equal(manifest.response.status, 200);
    assert.equal(manifest.payload.data.schemaVersion, 1);
    assert.equal(Array.isArray(manifest.payload.data.terms), true);
    assert.deepEqual(manifest.payload.data.terms.map(term => term.termCode), ['FALL 2026', 'SPRING 2026']);
    assert.equal(manifest.payload.data.terms[0].rowCount, 2);
    assert.equal(manifest.payload.data.terms[0].hasArchive, true);
    assert.equal('data' in manifest.payload.data.terms[0], false);
    assert.equal('filePath' in manifest.payload.data.terms[0], false);

    const legacyList = await jsonRequest(baseUrl, '/api/analytics-archive');
    assert.equal(legacyList.response.status, 200);
    assert.deepEqual(legacyList.payload.data.map(term => term.term), ['FALL 2026', 'SPRING 2026']);

    const fullArchive = await jsonRequest(baseUrl, '/api/analytics-archive/FALL%202026');
    assert.equal(fullArchive.response.status, 200);
    assert.equal(fullArchive.payload.data.length, 2);
  } finally {
    server.close();
  }
});

test('Analytics archive replacement updates manifest metadata and corrupt manifest rebuilds safely', async () => {
  const server = await listen();
  const baseUrl = `http://127.0.0.1:${server.address().port}`;
  try {
    const replacementCsv = `${csvA}10003,BIOL 031,24\n`;
    const saved = await jsonRequest(baseUrl, '/api/analytics-archive/FALL 2026', {
      method: 'POST',
      body: JSON.stringify({ password: 'Upload2025', csv: replacementCsv })
    });
    assert.equal(saved.response.status, 200);
    assert.equal(saved.payload.metadata.rowCount, 3);

    const manifestPath = path.join(dataDir, 'analytics-archive', 'manifest.json');
    fs.writeFileSync(manifestPath, '{not valid json');
    const rebuilt = await jsonRequest(baseUrl, '/api/analytics-archive/manifest');
    assert.equal(rebuilt.response.status, 200);
    const fall = rebuilt.payload.data.terms.find(term => term.termCode === 'FALL 2026');
    assert.equal(fall.rowCount, 3);
    assert.equal(fall.hasArchive, true);
  } finally {
    server.close();
  }
});

test('Analytics archive manifest sorting uses newest academic term order', () => {
  const sorted = sortArchiveTermsNewestFirst([
    { termCode: '202610' },
    { termCode: '202620' },
    { termCode: '202630' },
    { termCode: '202710' },
    { termCode: '202720' },
    { termCode: '202730' }
  ]);
  assert.deepEqual(sorted.map(term => term.termCode), ['202730', '202720', '202710', '202630', '202620', '202610']);
});

test('Analytics archive manifest displays Banner term codes using TIMBER conventions', () => {
  assert.equal(displayTermFromArchiveTerm('202710'), 'Fall 2026');
  assert.equal(displayTermFromArchiveTerm('202720'), 'Spring 2027');
  assert.equal(displayTermFromArchiveTerm('202730'), 'Summer 2027');
  assert.equal(displayTermFromArchiveTerm('202610'), 'Fall 2025');
  assert.equal(displayTermFromArchiveTerm('202620'), 'Spring 2026');
  assert.equal(displayTermFromArchiveTerm('202630'), 'Summer 2026');
});

test('Analytics archive manifest implementation uses atomic writes and avoids filesystem paths', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');
  assert.match(source, /app\.get\('\/api\/analytics-archive\/manifest'/);
  assert.match(source, /atomicWriteJson\(ANALYTICS_ARCHIVE_MANIFEST_PATH/);
  assert.doesNotMatch(source, /filePath:\s*filePath/);
});
