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

async function authHeaders(baseUrl, password = 'DevSecret', requestedRole = 'development') {
  const auth = await jsonRequest(baseUrl, '/api/auth/role', {
    method: 'POST',
    body: JSON.stringify({ password, requestedRole })
  });
  assert.equal(auth.response.status, 200);
  return { Authorization: `Bearer ${auth.payload.token}` };
}

function rowFixture(id, crns, overrides = {}) {
  return {
    id,
    termCode: '202710',
    course: 'COMM C1000',
    crnDisplay: crns.join(' / '),
    crns,
    initialEnrollment: 12,
    currentEnrollment: 12,
    latestEnrollment: 12,
    highestEnrollment: 12,
    threshold: 21,
    snapshotValues: { '2026-08-04': 12 },
    justification: '',
    vpComments: '',
    ...overrides
  };
}

function workspaceFixture(termCode = '202710', overrides = {}) {
  return {
    termCode,
    displayTerm: 'FALL 2026',
    sourceFilename: '202710 FA26 Low Enrolled Watchlist_8-4-26.xlsx',
    initialSnapshotDate: '2026-08-04',
    reasons: ['Dual Enrollment', 'Single section offering in Hanford', 'Other', 'Other'],
    rows: [
      rowFixture('row-10003', ['10003']),
      rowFixture('row-cross', ['10004', '10005'], {
        course: 'ENGL C1000',
        initialEnrollment: null,
        currentEnrollment: null,
        latestEnrollment: null,
        highestEnrollment: null,
        snapshotValues: { '2026-08-04': null }
      })
    ],
    snapshots: [{
      snapshotDate: '2026-08-04',
      type: 'initial',
      values: {
        'row-10003': { enrollment: 12 },
        'row-cross': { enrollment: null }
      }
    }],
    uploadHistory: [{ type: 'initial', snapshotDate: '2026-08-04', rowsImported: 2 }],
    ...overrides
  };
}

function trackerPath(termCode) {
  return path.join(dataDir, 'low-enrollment-tracking', `${termCode}.json`);
}

test('Low Enrollment Tracking rejects invalid term codes and missing workspaces', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const invalid = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/FALL2026');
    assert.equal(invalid.response.status, 400);
    assert.equal(invalid.payload.code, 'INVALID_TERM');

    const missing = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202799');
    assert.equal(missing.response.status, 404);
    assert.equal(missing.payload.code, 'WORKSPACE_NOT_FOUND');
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('Low Enrollment Tracking requires Development/Admin access for writes', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const unauthorized = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      body: JSON.stringify({ workspace: workspaceFixture() })
    });
    assert.equal(unauthorized.response.status, 401);
    assert.equal(unauthorized.payload.code, 'UNAUTHORIZED');

    const generalHeaders = await authHeaders(baseUrl, 'GeneralSecret', 'general');
    const forbidden = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      headers: generalHeaders,
      body: JSON.stringify({ workspace: workspaceFixture() })
    });
    assert.equal(forbidden.response.status, 403);
    assert.equal(forbidden.payload.code, 'FORBIDDEN');
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('Low Enrollment Tracking persists workspaces, replacement rules, warnings, and reloads', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const headers = await authHeaders(baseUrl);

    const created = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: workspaceFixture('202710', {
        rows: [
          rowFixture('row-a', ['10003']),
          rowFixture('row-b', ['10003'], { course: 'COMM C1000B' }),
          rowFixture('row-cross', ['10004', '10005'], { latestEnrollment: null, currentEnrollment: null, highestEnrollment: null })
        ]
      }) })
    });
    assert.equal(created.response.status, 200);
    assert.equal(created.payload.data.reasons.length, 3);
    assert.equal(created.payload.data.validationWarnings[0].code, 'DUPLICATE_CRN');
    assert.equal(created.payload.summary.rowCount, 3);
    assert.ok(fs.existsSync(trackerPath('202710')));
    JSON.parse(fs.readFileSync(trackerPath('202710'), 'utf8'));
    const createdAt = created.payload.data.createdAt;

    const listed = await jsonRequest(baseUrl, '/api/low-enrollment-tracking');
    assert.equal(listed.response.status, 200);
    assert.equal(listed.payload.data.some(item => item.termCode === '202710'), true);

    const loaded = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710');
    assert.equal(loaded.response.status, 200);
    assert.deepEqual(loaded.payload.data.rows.find(row => row.id === 'row-cross').crns, ['10004', '10005']);
    assert.equal(loaded.payload.data.rows.find(row => row.id === 'row-cross').latestEnrollment, null);

    const duplicate = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: workspaceFixture('202710') })
    });
    assert.equal(duplicate.response.status, 409);
    assert.equal(duplicate.payload.code, 'WORKSPACE_EXISTS');

    const replaced = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202710', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: workspaceFixture('202710', { sourceFilename: 'replacement.xlsx' }), replaceExisting: true })
    });
    assert.equal(replaced.response.status, 200);
    assert.equal(replaced.payload.data.sourceFilename, 'replacement.xlsx');
    assert.equal(replaced.payload.data.createdAt, createdAt);
    JSON.parse(fs.readFileSync(trackerPath('202710'), 'utf8'));
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('Low Enrollment Tracking validates invalid workspace shapes', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const headers = await authHeaders(baseUrl);
    const cases = [
      ['202711', workspaceFixture('202710'), 'INVALID_WORKSPACE'],
      ['202712', workspaceFixture('202712', { rows: [{ id: '', crns: ['10001'], threshold: 1 }] }), 'INVALID_WORKSPACE'],
      ['202713', workspaceFixture('202713', { rows: [rowFixture('bad', ['ABC'], { threshold: 1 })] }), 'INVALID_WORKSPACE'],
      ['202714', workspaceFixture('202714', { rows: [rowFixture('bad', ['10001'], { threshold: -1 })] }), 'INVALID_WORKSPACE'],
      ['202715', workspaceFixture('202715', { reasons: [] }), 'INVALID_WORKSPACE'],
      ['202716', workspaceFixture('202716', { snapshots: {} }), 'INVALID_WORKSPACE']
    ];
    for (const [termCode, workspace, code] of cases) {
      const response = await jsonRequest(baseUrl, `/api/low-enrollment-tracking/${termCode}`, {
        method: 'POST',
        headers,
        body: JSON.stringify({ workspace })
      });
      assert.equal(response.response.status, 400, `${termCode} should fail`);
      assert.equal(response.payload.code, code);
    }
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('Low Enrollment Tracking saves snapshots while preserving editable row fields and metadata', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const headers = await authHeaders(baseUrl);
    await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: workspaceFixture('202720') })
    });

    const patchedReason = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720/rows/row-10003', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ justification: 'Other' })
    });
    assert.equal(patchedReason.response.status, 200);

    const patchedComment = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720/rows/row-10003', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ vpComments: 'Reviewed with dean.' })
    });
    assert.equal(patchedComment.response.status, 200);

    const snapshotRows = workspaceFixture('202720').rows.map(row => ({
      ...row,
      latestEnrollment: row.id === 'row-10003' ? 25 : null,
      highestEnrollment: row.id === 'row-10003' ? 25 : null,
      snapshotMatchStatus: { '2026-08-11': row.id === 'row-cross' ? 'partial' : 'matched' },
      snapshotMissingCrns: { '2026-08-11': row.id === 'row-cross' ? ['10005'] : [] },
      justification: '',
      vpComments: ''
    }));
    const savedSnapshot = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({
        snapshot: {
          snapshotDate: '2026-08-11',
          type: 'enrollment-update',
          values: {
            'row-10003': { enrollment: 25, matchedCrns: ['10003'], missingCrns: [] },
            'row-cross': { enrollment: null, matchedCrns: ['10004'], missingCrns: ['10005'] }
          }
        },
        uploadHistory: { type: 'snapshot', snapshotDate: '2026-08-11', partiallyMatchedRows: 1 },
        rows: snapshotRows
      })
    });
    assert.equal(savedSnapshot.response.status, 200);
    const savedRow = savedSnapshot.payload.data.rows.find(row => row.id === 'row-10003');
    assert.equal(savedRow.latestEnrollment, 25);
    assert.equal(savedRow.currentEnrollment, 25);
    assert.equal(savedRow.status, 'Threshold Met');
    assert.equal(savedRow.justification, 'Other');
    assert.equal(savedRow.vpComments, 'Reviewed with dean.');
    const partialRow = savedSnapshot.payload.data.rows.find(row => row.id === 'row-cross');
    assert.deepEqual(partialRow.snapshotMissingCrns['2026-08-11'], ['10005']);
    assert.equal(partialRow.latestEnrollment, null);

    const duplicateSnapshot = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({ snapshot: { snapshotDate: '2026-08-11', values: {} }, rows: snapshotRows })
    });
    assert.equal(duplicateSnapshot.response.status, 409);
    assert.equal(duplicateSnapshot.payload.code, 'SNAPSHOT_EXISTS');

    const replacedSnapshot = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({
        snapshot: { snapshotDate: '2026-08-11', type: 'enrollment-update', values: { 'row-10003': { enrollment: 26 } } },
        rows: snapshotRows.map(row => row.id === 'row-10003' ? { ...row, latestEnrollment: 26, highestEnrollment: 26 } : row),
        replaceExisting: true
      })
    });
    assert.equal(replacedSnapshot.response.status, 200);
    assert.equal(replacedSnapshot.payload.data.snapshots.filter(item => item.snapshotDate === '2026-08-11').length, 1);
    assert.equal(replacedSnapshot.payload.data.rows.find(row => row.id === 'row-10003').justification, 'Other');

    const invalidRowSet = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({ snapshot: { snapshotDate: '2026-08-18', values: {} }, rows: [snapshotRows[0]] })
    });
    assert.equal(invalidRowSet.response.status, 400);
    assert.equal(invalidRowSet.payload.code, 'INVALID_SNAPSHOT');

    const invalidSnapshotValue = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720/snapshots', {
      method: 'POST',
      headers,
      body: JSON.stringify({ snapshot: { snapshotDate: '2026-08-18', values: { unknown: { enrollment: 1 } } }, rows: snapshotRows })
    });
    assert.equal(invalidSnapshotValue.response.status, 400);
    assert.equal(invalidSnapshotValue.payload.code, 'INVALID_SNAPSHOT');

    const loaded = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202720');
    assert.equal(loaded.response.status, 200);
    assert.equal(loaded.payload.data.rows.find(row => row.id === 'row-cross').latestEnrollment, null);
    JSON.parse(fs.readFileSync(trackerPath('202720'), 'utf8'));
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('Low Enrollment Tracking row patch validation covers reason, comments, fields, and missing rows', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const headers = await authHeaders(baseUrl);
    await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202730', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: workspaceFixture('202730') })
    });

    const validReason = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202730/rows/row-10003', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ justification: 'Dual Enrollment' })
    });
    assert.equal(validReason.response.status, 200);
    assert.equal(validReason.payload.data.row.justification, 'Dual Enrollment');

    const invalidReason = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202730/rows/row-10003', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ justification: 'Not a saved reason' })
    });
    assert.equal(invalidReason.response.status, 400);
    assert.equal(invalidReason.payload.code, 'INVALID_JUSTIFICATION');

    const validComment = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202730/rows/row-10003', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ vpComments: 'Reviewed with VP.' })
    });
    assert.equal(validComment.response.status, 200);
    assert.equal(validComment.payload.data.row.vpComments, 'Reviewed with VP.');

    const unsupported = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202730/rows/row-10003', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ threshold: 99 })
    });
    assert.equal(unsupported.response.status, 400);
    assert.equal(unsupported.payload.code, 'INVALID_ROW_UPDATE');

    const missingRow = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202730/rows/missing-row', {
      method: 'PATCH',
      headers,
      body: JSON.stringify({ vpComments: 'No row' })
    });
    assert.equal(missingRow.response.status, 404);
    assert.equal(missingRow.payload.code, 'ROW_NOT_FOUND');
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('Low Enrollment Tracking concurrent writes leave valid JSON and preserve both operations', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const headers = await authHeaders(baseUrl, 'AdminSecret', 'admin');
    await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202740', {
      method: 'POST',
      headers,
      body: JSON.stringify({ workspace: workspaceFixture('202740') })
    });
    const rows = workspaceFixture('202740').rows;
    const [comment, snapshot] = await Promise.all([
      jsonRequest(baseUrl, '/api/low-enrollment-tracking/202740/rows/row-10003', {
        method: 'PATCH',
        headers,
        body: JSON.stringify({ vpComments: 'Concurrent comment save.' })
      }),
      jsonRequest(baseUrl, '/api/low-enrollment-tracking/202740/snapshots', {
        method: 'POST',
        headers,
        body: JSON.stringify({
          snapshot: { snapshotDate: '2026-08-25', type: 'enrollment-update', values: { 'row-10003': { enrollment: 24 } } },
          rows: rows.map(row => row.id === 'row-10003' ? { ...row, latestEnrollment: 24, highestEnrollment: 24 } : row)
        })
      })
    ]);
    assert.equal(comment.response.status, 200);
    assert.equal(snapshot.response.status, 200);
    const saved = JSON.parse(fs.readFileSync(trackerPath('202740'), 'utf8'));
    assert.equal(saved.rows.find(row => row.id === 'row-10003').vpComments, 'Concurrent comment save.');
    assert.equal(saved.snapshots.some(item => item.snapshotDate === '2026-08-25'), true);
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('Low Enrollment Tracking imports manual fields atomically without changing enrollment data', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const headers = await authHeaders(baseUrl);
    const workspace = workspaceFixture('202750');
    workspace.rows[0].justification = 'Other';
    workspace.rows[0].vpComments = 'Clear both';
    await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202750', {
      method: 'POST', headers, body: JSON.stringify({ workspace })
    });

    const imported = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202750/manual-import', {
      method: 'POST', headers, body: JSON.stringify({
        sourceFilename: 'Fall-2026-Tracking.xlsx',
        updates: [
          { rowId: 'row-10003', justification: '', vpComments: '', expectedJustification: 'Other', expectedVpComments: 'Clear both' },
          { rowId: 'row-cross', justification: 'Dual Enrollment', vpComments: 'Reviewed in Excel.', expectedJustification: '', expectedVpComments: '' }
        ]
      })
    });
    assert.equal(imported.response.status, 200);
    assert.equal(imported.payload.data.rows[0].justification, '');
    assert.equal(imported.payload.data.rows[0].vpComments, '');
    assert.equal(imported.payload.data.rows[0].latestEnrollment, 12);
    assert.equal(imported.payload.data.snapshots.length, 1);
    assert.equal(imported.payload.summary.clearedJustifications, 1);
    assert.equal(imported.payload.summary.clearedVpComments, 1);

    const rejected = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202750/manual-import', {
      method: 'POST', headers, body: JSON.stringify({ updates: [
        { rowId: 'row-10003', justification: 'Other', vpComments: 'Would change', expectedJustification: '', expectedVpComments: '' },
        { rowId: 'row-cross', justification: 'Typed custom reason', vpComments: '', expectedJustification: 'Dual Enrollment', expectedVpComments: 'Reviewed in Excel.' }
      ] })
    });
    assert.equal(rejected.response.status, 400);
    assert.equal(rejected.payload.code, 'INVALID_JUSTIFICATION');
    const loaded = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202750');
    assert.equal(loaded.payload.data.rows[0].justification, '');
    assert.equal(loaded.payload.data.rows[0].vpComments, '');

    const conflict = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202750/manual-import', {
      method: 'POST', headers, body: JSON.stringify({ updates: [
        { rowId: 'row-10003', justification: 'Other', vpComments: 'Old export', expectedJustification: 'Other', expectedVpComments: 'Stale value' }
      ] })
    });
    assert.equal(conflict.response.status, 409);
    assert.equal(conflict.payload.code, 'MANUAL_IMPORT_CONFLICT');
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('Low Enrollment Tracking exclusions are reversible and survive enrollment snapshots', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const headers = await authHeaders(baseUrl);
    const workspace = workspaceFixture('202760');
    await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202760', {
      method: 'POST', headers, body: JSON.stringify({ workspace })
    });
    const excluded = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202760/rows/row-10003/exclusion', {
      method: 'POST', headers, body: JSON.stringify({ excluded: true, reason: 'Open lab', note: 'LA 425' })
    });
    assert.equal(excluded.response.status, 200);
    assert.equal(excluded.payload.data.row.exclusion.excluded, true);

    const snapshotRows = workspace.rows.map(row => row.id === 'row-10003'
      ? { ...row, latestEnrollment: 15, highestEnrollment: 15 }
      : row);
    const snapshot = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202760/snapshots', {
      method: 'POST', headers, body: JSON.stringify({
        snapshot: { snapshotDate: '2026-09-01', type: 'enrollment-update', values: { 'row-10003': { enrollment: 15 } } },
        rows: snapshotRows
      })
    });
    assert.equal(snapshot.response.status, 200);
    assert.equal(snapshot.payload.data.rows[0].exclusion.excluded, true);
    assert.equal(snapshot.payload.data.rows[0].latestEnrollment, 15);

    const restored = await jsonRequest(baseUrl, '/api/low-enrollment-tracking/202760/rows/row-10003/exclusion', {
      method: 'POST', headers, body: JSON.stringify({ excluded: false, reason: '', note: '' })
    });
    assert.equal(restored.response.status, 200);
    assert.equal(restored.payload.data.row.exclusion.excluded, false);
    const saved = JSON.parse(fs.readFileSync(trackerPath('202760'), 'utf8'));
    assert.equal(saved.exclusionHistory.length, 2);
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});
