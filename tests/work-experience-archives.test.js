const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const test = require('node:test');

const backend = require('../server.js');

test('Work Experience archive metadata counts rows CRNs enrollment and FTES', () => {
  const metadata = backend.workExperienceMetadata('FALL 2026', [
    { CRN: '10001', ACTUAL_ENROLL: '10', FTES: '1.2' },
    { CRN: '10001', ACTUAL_ENROLL: '5', FTES: '0.8' },
    { CRN: '10002', Enrollment: '7', FTES: '0.7' }
  ], { sourceFileName: 'work experience.csv' });

  assert.equal(metadata.term, 'FALL 2026');
  assert.equal(metadata.rawRowCount, 3);
  assert.equal(metadata.distinctCrnCount, 2);
  assert.equal(metadata.enrollmentTotal, 22);
  assert.equal(metadata.ftesTotal, 2.7);
  assert.equal(metadata.sourceFileName, 'work experience.csv');
});

test('Work Experience archive validation requires CRNs', () => {
  assert.equal(backend.validateWorkExperienceRows([{ CRN: '10001' }]).valid, true);
  assert.equal(backend.validateWorkExperienceRows([{ Course: 'WKEX' }]).valid, false);
});

test('Work Experience archives use separate API routes and storage from section schedules', () => {
  const source = fs.readFileSync(path.join(__dirname, '..', 'server.js'), 'utf8');

  assert.match(source, /WORK_EXPERIENCE_DIR = path\.join\(DATA_DIR, 'work-experience'\)/);
  assert.match(source, /app\.get\('\/api\/work-experience'/);
  assert.match(source, /app\.get\('\/api\/work-experience\/:term'/);
  assert.match(source, /app\.post\('\/api\/work-experience\/:term'/);
  assert.match(source, /app\.delete\('\/api\/work-experience\/:term'/);
  assert.match(source, /getSchedulePath\(term\)/);
  assert.match(source, /getAnalyticsArchivePath\(term\)/);
  assert.match(source, /getWorkExperiencePath\(term\)/);
});
