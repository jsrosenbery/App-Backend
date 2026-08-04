const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

test('Low Enrollment Tracking endpoints use dedicated persistent storage and auth', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');

  assert.match(source, /LOW_ENROLLMENT_TRACKING_DIR = path\.join\(DATA_DIR, 'low-enrollment-tracking'\)/);
  assert.match(source, /function getLowEnrollmentTrackingPath\(term\)/);
  assert.match(source, /app\.get\('\/api\/low-enrollment-tracking'/);
  assert.match(source, /app\.get\('\/api\/low-enrollment-tracking\/:term'/);
  assert.match(source, /app\.post\('\/api\/low-enrollment-tracking\/:term'/);
  assert.match(source, /app\.post\('\/api\/low-enrollment-tracking\/:term\/snapshots'/);
  assert.match(source, /app\.patch\('\/api\/low-enrollment-tracking\/:term\/rows\/:rowId'/);
  assert.match(source, /isEnrollmentSessionAuthorized\(req\) && !isAuthorized\(password\)/);
  assert.match(source, /workspace\.snapshots = nextSnapshots/);
  assert.match(source, /workspace\.uploadHistory/);
});

test('Low Enrollment Tracking validates workspace rows before save', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');

  assert.match(source, /Tracker rows are required/);
  assert.match(source, /tracker row\(s\) are missing CRNs/);
  assert.match(source, /saveLowEnrollmentWorkspace\(workspace\)/);
  assert.match(source, /JSON\.stringify\(workspace, null, 2\)/);
});
