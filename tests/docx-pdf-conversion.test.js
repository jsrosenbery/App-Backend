const assert = require('node:assert/strict');
const fs = require('node:fs');
const os = require('node:os');
const path = require('node:path');
const test = require('node:test');

const backend = require('../server.js');

function tempDir(name) {
  return fs.mkdtempSync(path.join(os.tmpdir(), `${name}-`));
}

function writeFakeLibreOffice(mode = 'success') {
  const dir = tempDir('fake-lo');
  const script = path.join(dir, 'fake-libreoffice.js');
  fs.writeFileSync(script, `
const fs = require('fs');
const path = require('path');
if (process.argv.includes('--version')) {
  console.log('LibreOffice 7.6.0.0 Fake');
  process.exit(0);
}
if (${JSON.stringify(mode)} === 'fail') {
  console.error('fake conversion failed');
  process.exit(7);
}
const outIndex = process.argv.indexOf('--outdir');
const outDir = outIndex >= 0 ? process.argv[outIndex + 1] : process.cwd();
const input = process.argv[process.argv.length - 1];
const pdf = path.join(outDir, path.basename(input, path.extname(input)) + '.pdf');
fs.writeFileSync(pdf, Buffer.from('%PDF-1.4\\n% fake pdf\\n'));
console.log('convert ' + input + ' -> ' + pdf);
`, 'utf8');
  return { dir, script };
}

test('LibreOffice detection reports path and version when command runs', () => {
  const detected = backend.detectDocxPdfConverter({ libreOfficePath: process.execPath });

  assert.equal(detected.available, true);
  assert.equal(detected.installed, true);
  assert.equal(detected.command, process.execPath);
  assert.match(detected.version, /v?\d+\.\d+/);
});

test('DOCX to PDF conversion succeeds with filenames containing spaces', async () => {
  const fake = writeFakeLibreOffice('success');
  const requestDir = tempDir('docx-success');
  const inputPath = path.join(requestDir, 'Schedule Change 10003.docx');
  fs.writeFileSync(inputPath, Buffer.from('PK fake docx'));

  const result = await backend.convertDocxToPdf(inputPath, requestDir, {
    commands: [{ command: process.execPath, argsPrefix: [fake.script] }]
  });

  assert.equal(path.basename(result.outputPath), 'Schedule Change 10003.pdf');
  assert.equal(fs.existsSync(result.outputPath), true);
  assert.equal(result.attempts[0].exitCode, 0);
  assert.match(result.attempts[0].stdout, /convert/);
});

test('DOCX to PDF conversion failure includes exit code stdout stderr and attempts', async () => {
  const fake = writeFakeLibreOffice('fail');
  const requestDir = tempDir('docx-fail');
  const inputPath = path.join(requestDir, 'Schedule Change 10004.docx');
  fs.writeFileSync(inputPath, Buffer.from('PK fake docx'));

  await assert.rejects(
    () => backend.convertDocxToPdf(inputPath, requestDir, {
      commands: [{ command: process.execPath, argsPrefix: [fake.script] }]
    }),
    err => {
      assert.match(err.message, /fake conversion failed|converter unavailable or failed/);
      assert.equal(err.attempts[0].exitCode, 7);
      assert.match(err.attempts[0].stderr, /fake conversion failed/);
      return true;
    }
  );
});

test('DOCX to PDF conversion handles concurrent request directories independently', async () => {
  const fake = writeFakeLibreOffice('success');
  const dirs = [tempDir('docx-concurrent-a'), tempDir('docx-concurrent-b')];
  const inputs = dirs.map((dir, index) => {
    const file = path.join(dir, `Schedule Change ${index + 1}.docx`);
    fs.writeFileSync(file, Buffer.from('PK fake docx'));
    return file;
  });

  const results = await Promise.all(inputs.map(input => backend.convertDocxToPdf(input, path.dirname(input), {
    commands: [{ command: process.execPath, argsPrefix: [fake.script] }]
  })));

  assert.equal(new Set(results.map(result => result.outputPath)).size, 2);
  results.forEach(result => assert.equal(fs.existsSync(result.outputPath), true));
});

test('cleanup removes temporary conversion directory', () => {
  const requestDir = tempDir('docx-cleanup');
  fs.writeFileSync(path.join(requestDir, 'temp.docx'), 'PK fake docx');

  const cleanup = backend.cleanupConversionDir(requestDir);

  assert.equal(cleanup.ok, true);
  assert.equal(fs.existsSync(requestDir), false);
});

test('diagnostics payload includes Node platform and PDF capability fields', () => {
  const diagnostics = backend.diagnosticsPayload();

  assert.equal(diagnostics.nodeVersion, process.version);
  assert.equal(diagnostics.platform, process.platform);
  assert.equal(typeof diagnostics.libreOfficeInstalled, 'boolean');
  assert.equal(typeof diagnostics.pdfConversionAvailable, 'boolean');
  assert.ok('pdfConversionUnavailableReason' in diagnostics);
});
