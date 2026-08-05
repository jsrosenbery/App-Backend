const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const dataRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'cos-low-enrollment-'));
const dataDir = path.join(dataRoot, 'cos-app');
fs.mkdirSync(dataDir, { recursive: true });
process.env.DATA_DIR = dataDir;
process.env.UPLOAD_PASSWORD = 'Upload2025';
process.env.GENERAL_PASSWORD = 'GeneralSecret';
process.env.EM_PASSWORD = 'EmSecret';

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

test('Low Enrollment workspaces save, list, reject unconfirmed overwrite, snapshot, and patch manual fields', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const auth = await jsonRequest(baseUrl, '/api/auth/enrollment-management', {
      method: 'POST',
      body: JSON.stringify({ password: 'EmSecret' })
    });
    assert.equal(auth.response.status, 200);
    const headers = { Authorization: `Bearer ${auth.payload.token}` };
    const workspace = {
      termCode: '202710',
      displayTerm: 'FALL 2026',
      rows: [{ id: 'row-1', course: 'COMM C1000', crns: ['10001'], justification: '', vpComments: '' }],
      snapshots: [{ type: 'initial', snapshotDate: '2026-08-04', values: {} }],
      uploadHistory: [{ type: 'initial', snapshotDate: '2026-08-04' }]
    };

    const denied = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      body: JSON.stringify({ workspace, replaceExisting: false })
    });
    assert.equal(denied.response.status, 403);

    const saved = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace, replaceExisting: false })
    });
    assert.equal(saved.response.status, 200);
    assert.equal(saved.payload.data.rows.length, 1);
    assert.ok(fs.existsSync(path.join(dataDir, 'low-enrollment-tracking', '202710.json')));

    const blockedOverwrite = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: { ...workspace, rows: [] }, replaceExisting: false })
    });
    assert.equal(blockedOverwrite.response.status, 409);

    const listed = await jsonRequest(baseUrl, '/api/low-enrollment-tracking');
    assert.equal(listed.response.status, 200);
    assert.equal(listed.payload.data[0].termCode, '202710');

    const snapshot = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({
        snapshot: { type: 'enrollment-update', snapshotDate: '2026-08-11', values: { 'row-1': { enrollment: 22 } } },
        uploadHistory: { type: 'snapshot', snapshotDate: '2026-08-11' },
        rows: [{ ...workspace.rows[0], latestEnrollment: 22 }],
        replaceExisting: false
      })
    });
    assert.equal(snapshot.response.status, 200);
    assert.equal(snapshot.payload.data.snapshots.length, 2);

    const duplicateSnapshot = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({
        snapshot: { type: 'enrollment-update', snapshotDate: '2026-08-11', values: {} },
        rows: workspace.rows,
        replaceExisting: false
      })
    });
    assert.equal(duplicateSnapshot.response.status, 409);

    const patched = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710/rows/row-1', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ justification: 'Other', vpComments: 'Keep watching', latestEnrollment: 999 })
    });
    assert.equal(patched.response.status, 200);
    assert.equal(patched.payload.data.justification, 'Other');
    assert.equal(patched.payload.data.vpComments, 'Keep watching');
    assert.equal(patched.payload.data.latestEnrollment, 22);
  } finally {
    await new Promise(resolve => server.close(resolve));
    fs.rmSync(dataRoot, { recursive: true, force: true });
  }
});
