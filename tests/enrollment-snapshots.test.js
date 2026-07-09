const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const dataRoot = fs.mkdtempSync(path.join(os.tmpdir(), 'cos-snapshots-'));
const dataDir = path.join(dataRoot, 'cos-app');
fs.mkdirSync(dataDir, { recursive: true });
process.env.DATA_DIR = dataDir;
process.env.UPLOAD_PASSWORD = 'Upload2025';
process.env.GENERAL_PASSWORD = 'GeneralSecret';
process.env.DEAN_PASSWORD = 'DeanSecret';
process.env.EM_PASSWORD = 'EmSecret';
process.env.DEV_PASSWORD = 'DevSecret';
process.env.ADMIN_PASSWORD = 'AdminSecret';
process.env.MIGRATION_TOKEN = 'MigrationSecret';

const { app, validateMigrationArchiveEntries, MIGRATION_IMPORT_TMP_DIR } = require('../server');

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
      body: JSON.stringify({ password: 'EmSecret' })
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

test('role authentication honors hierarchical password access', async () => {
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const cases = [
      ['GeneralSecret', 'general', 'general'],
      ['DeanSecret', 'general', 'dean'],
      ['DeanSecret', 'dean', 'dean'],
      ['EmSecret', 'dean', 'em'],
      ['EmSecret', 'em', 'em'],
      ['DevSecret', 'em', 'development'],
      ['DevSecret', 'development', 'development'],
      ['AdminSecret', 'development', 'admin'],
      ['AdminSecret', 'admin', 'admin']
    ];

    for (const [password, requestedRole, expectedRole] of cases) {
      const auth = await jsonRequest(baseUrl, '/api/auth/role', {
        method: 'POST',
        body: JSON.stringify({ password, requestedRole })
      });
      assert.equal(auth.response.status, 200, `${expectedRole} should unlock ${requestedRole}`);
      assert.equal(auth.payload.role, expectedRole);
      assert.ok(auth.payload.token);
      assert.ok(auth.payload.roleLevel >= 1);
    }

    const deanDeniedForEm = await jsonRequest(baseUrl, '/api/auth/role', {
      method: 'POST',
      body: JSON.stringify({ password: 'DeanSecret', requestedRole: 'em' })
    });
    assert.equal(deanDeniedForEm.response.status, 403);

    const emCompatibility = await jsonRequest(baseUrl, '/api/auth/enrollment-management', {
      method: 'POST',
      body: JSON.stringify({ password: 'EmSecret' })
    });
    assert.equal(emCompatibility.response.status, 200);
    assert.equal(emCompatibility.payload.role, 'em');
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('temporary migration endpoints require token and report data status', async () => {
  fs.mkdirSync(dataDir, { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'rooms.json'), JSON.stringify([{ building: 'A', room: '1' }]));
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const denied = await fetch(`${baseUrl}/admin/migration/status`);
    assert.equal(denied.status, 403);

    const allowed = await fetch(`${baseUrl}/admin/migration/status`, {
      headers: { 'x-migration-token': 'MigrationSecret' }
    });
    assert.equal(allowed.status, 200);
    const payload = await allowed.json();
    assert.equal(payload.temporaryMigrationOnly, true);
    assert.equal(payload.exists, true);
    assert.ok(payload.diskUsage.files >= 1);
    assert.ok(payload.topLevelEntries.some(entry => entry.name === 'rooms.json'));
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});

test('migration archive validation rejects unsafe archive paths', () => {
  assert.equal(validateMigrationArchiveEntries(['cos-app/', 'cos-app/rooms.json']).valid, true);
  assert.equal(validateMigrationArchiveEntries(['/cos-app/rooms.json']).valid, false);
  assert.equal(validateMigrationArchiveEntries(['cos-app/../evil.txt']).valid, false);
  assert.equal(validateMigrationArchiveEntries(['other/rooms.json']).valid, false);
  assert.equal(validateMigrationArchiveEntries([]).valid, false);
});

test('migration import temporary upload path is outside persistent cos-app data', () => {
  assert.equal(MIGRATION_IMPORT_TMP_DIR, '/tmp/migration-imports');
  assert.equal(MIGRATION_IMPORT_TMP_DIR.includes('/var/data/cos-app'), false);
});

test('room catalog export is public but import remains password protected', async () => {
  fs.mkdirSync(dataDir, { recursive: true });
  fs.writeFileSync(path.join(dataDir, 'rooms.json'), JSON.stringify([
    { campus: 'COS', building: 'A', room: '101', capacity: 40, type: 'Classroom' }
  ]));
  const server = await listen();
  try {
    const baseUrl = `http://127.0.0.1:${server.address().port}`;
    const exported = await jsonRequest(baseUrl, '/api/rooms/export');
    assert.equal(exported.response.status, 200);
    assert.equal(exported.payload.data.length, 1);
    assert.equal(exported.payload.data[0].building, 'A');

    const deniedImport = await jsonRequest(baseUrl, '/api/rooms/import', {
      method: 'POST',
      body: JSON.stringify({ rooms: [{ building: 'B', room: '201' }] })
    });
    assert.equal(deniedImport.response.status, 403);
  } finally {
    await new Promise(resolve => server.close(resolve));
  }
});
