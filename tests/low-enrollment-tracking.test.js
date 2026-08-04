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
process.env.DIV_CHAIR_PASSWORD = 'DivChairSecret';
process.env.DEAN_PASSWORD = 'DeanSecret';
process.env.EM_PASSWORD = 'EmSecret';
process.env.DEV_PASSWORD = 'DevSecret';
process.env.ADMIN_PASSWORD = 'AdminSecret';

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

function workspaceFixture(overrides = {}) {
  return {
    termCode: '202710',
    displayTerm: 'FALL 2026',
    sourceFilename: '202710 FA26 Low Enrolled Watchlist_8-4-26.xlsx',
    initialSnapshotDate: '2026-08-04',
    reasons: ['Dual Enrollment', 'Single section offering in Hanford', 'Other'],
    rows: [
      {
        id: 'row-10003',
        termCode: '202710',
        course: 'COMM C1000',
        crnDisplay: '10003',
        crns: ['10003'],
        initialEnrollment: 12,
        latestEnrollment: 12,
        highestEnrollment: 12,
        threshold: 21,
        justification: '',
        vpComments: ''
      }
    ],
    snapshots: [{ snapshotDate: '2026-08-04', type: 'initial', values: { 'row-10003': { enrollment: 12 } } }],
    uploadHistory: [{ type: 'initial', snapshotDate: '2026-08-04', rowsImported: 1 }],
    ...overrides
  };
}

test('Low Enrollment Tracking endpoints persist data, protect writes, and validate row edits', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const denied = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      body: JSON.stringify({ workspace: workspaceFixture() })
    });
    assert.equal(denied.response.status, 403);

    const auth = await jsonRequest(baseUrl, '/api/auth/role', {
      method: 'POST',
      body: JSON.stringify({ password: 'DevSecret', requestedRole: 'development' })
    });
    assert.equal(auth.response.status, 200);
    const headers = { Authorization: `Bearer ${auth.payload.token}` };

    const imported = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: workspaceFixture() })
    });
    assert.equal(imported.response.status, 200);
    assert.equal(imported.payload.summary.rowCount, 1);
    assert.ok(fs.existsSync(path.join(dataDir, 'low-enrollment-tracking', '202710.json')));

    const duplicate = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: workspaceFixture() })
    });
    assert.equal(duplicate.response.status, 409);
    assert.equal(duplicate.payload.requiresConfirmation, true);

    const replaced = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: workspaceFixture({ sourceFilename: 'replacement.xlsx' }), replaceExisting: true })
    });
    assert.equal(replaced.response.status, 200);
    assert.equal(replaced.payload.data.sourceFilename, 'replacement.xlsx');

    const updatedJustification = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710/rows/row-10003', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ justification: 'Other' })
    });
    assert.equal(updatedJustification.response.status, 200);
    assert.equal(updatedJustification.payload.row.justification, 'Other');

    const badJustification = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710/rows/row-10003', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ justification: 'Not in this workbook' })
    });
    assert.equal(badJustification.response.status, 400);

    const badField = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710/rows/row-10003', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ threshold: 99 })
    });
    assert.equal(badField.response.status, 400);

    const savedSnapshot = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({
        snapshot: { snapshotDate: '2026-08-11', type: 'enrollment-update', values: { 'row-10003': { enrollment: 25 } } },
        uploadHistory: { type: 'snapshot', snapshotDate: '2026-08-11', fullyMatchedRows: 1 },
        rows: [{ ...workspaceFixture().rows[0], latestEnrollment: 25, justification: '' }]
      })
    });
    assert.equal(savedSnapshot.response.status, 200);
    assert.equal(savedSnapshot.payload.data.rows[0].justification, 'Other');

    const duplicateSnapshot = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({ snapshot: { snapshotDate: '2026-08-11', values: {} } })
    });
    assert.equal(duplicateSnapshot.response.status, 409);

    const replacedSnapshot = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({
        snapshot: { snapshotDate: '2026-08-11', type: 'enrollment-update', values: { 'row-10003': { enrollment: 26 } } },
        rows: [{ ...workspaceFixture().rows[0], latestEnrollment: 26 }],
        replaceExisting: true
      })
    });
    assert.equal(replacedSnapshot.response.status, 200);
    assert.equal(replacedSnapshot.payload.data.rows[0].latestEnrollment, 26);
    assert.equal(replacedSnapshot.payload.data.rows[0].justification, 'Other');

    const loaded = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710');
    assert.equal(loaded.response.status, 200);
    assert.equal(loaded.payload.data.snapshots.length, 2);
    assert.equal(loaded.payload.data.rows[0].latestEnrollment, 26);
  } finally {
    await new Promise(resolve => server.close(resolve));
    fs.rmSync(dataRoot, { recursive: true, force: true });
  }
});

test('Low Enrollment Tracking rejects invalid workspaces', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const auth = await jsonRequest(baseUrl, '/api/auth/role', {
      method: 'POST',
      body: JSON.stringify({ password: 'DevSecret', requestedRole: 'development' })
    });
    const headers = { Authorization: `Bearer ${auth.payload.token}` };
    const invalid = workspaceFixture({
      rows: [{ id: 'bad', crns: ['10004'], threshold: '' }]
    });
    const response = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: invalid })
    });
    assert.equal(response.response.status, 400);
    assert.match(response.payload.error, /invalid threshold/i);
  } finally {
    await new Promise(resolve => server.close(resolve));
    fs.rmSync(dataRoot, { recursive: true, force: true });
  }
});
