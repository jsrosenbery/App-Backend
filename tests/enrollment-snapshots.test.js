const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const dataDir = fs.mkdtempSync(path.join(os.tmpdir(), 'cos-snapshots-'));
process.env.DATA_DIR = dataDir;
process.env.UPLOAD_PASSWORD = 'Upload2025';

const { app } = require('../server');

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

test('backend enrollment snapshot endpoints append, update, list, and delete selected batches', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const auth = await jsonRequest(baseUrl, '/api/auth/enrollment-management', {
      method: 'POST',
      body: JSON.stringify({ password: 'Upload2025' })
    });
    assert.equal(auth.response.status, 200);
    assert.ok(auth.payload.token);
    const headers = { Authorization: `Bearer ${auth.payload.token}` };

    const initial = await jsonRequest(baseUrl, '/api/enrollment-snapshots');
    assert.equal(initial.response.status, 200);
    assert.deepEqual(initial.payload.data, []);

    const first = await jsonRequest(baseUrl, '/api/enrollment-snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({
        records: [
          { term: 'Fall 2027', crn: '10001', snapshotType: 'First Day', snapshotDate: '2027-08-16', enrollment: 22 },
          { term: 'Fall 2027', crn: '10002', snapshotType: 'First Day', snapshotDate: '2027-08-16', enrollment: 18 }
        ]
      })
    });
    assert.equal(first.response.status, 200);
    assert.equal(first.payload.appended, 2);
    assert.equal(first.payload.updated, 0);
    assert.equal(first.payload.data.length, 2);
    assert.ok(first.payload.lastUpdated);

    const second = await jsonRequest(baseUrl, '/api/enrollment-snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({
        records: [
          { term: 'Fall 2027', crn: '10001', snapshotType: 'First Day', snapshotDate: '2027-08-18', enrollment: 24 },
          { term: 'Fall 2027', crn: '10003', snapshotType: 'Census 2', snapshotDate: '2027-10-01', enrollment: 20 }
        ]
      })
    });
    assert.equal(second.response.status, 200);
    assert.equal(second.payload.appended, 1);
    assert.equal(second.payload.updated, 1);
    assert.equal(second.payload.data.length, 3);
    assert.equal(second.payload.data.find(record => record.crn === '10001').enrollment, 24);
    assert.equal(second.payload.data.find(record => record.crn === '10002').enrollment, 18);

    const listed = await jsonRequest(baseUrl, '/api/enrollment-snapshots');
    assert.equal(listed.response.status, 200);
    assert.equal(listed.payload.data.length, 3);
    assert.ok(fs.existsSync(path.join(dataDir, 'enrollment-snapshots.json')));

    const deleted = await jsonRequest(baseUrl, '/api/enrollment-snapshots', {
      method: 'DELETE',
      headers,
      body: JSON.stringify({ term: 'Fall 2027', snapshotType: 'First Day', snapshotDate: '2027-08-16' })
    });
    assert.equal(deleted.response.status, 200);
    assert.equal(deleted.payload.deleted, 1);
    assert.equal(deleted.payload.removed, 1);
    assert.equal(deleted.payload.data.length, 2);
    assert.equal(deleted.payload.data.some(record => record.crn === '10002'), false);
    assert.equal(deleted.payload.data.some(record => record.crn === '10001'), true);
    assert.equal(deleted.payload.data.some(record => record.crn === '10003'), true);
  } finally {
    await new Promise(resolve => server.close(resolve));
    fs.rmSync(dataDir, { recursive: true, force: true });
  }
});
