const express = require('express');
const fs = require('fs');
const path = require('path');
const cors = require('cors');
const Papa = require('papaparse');
const crypto = require('crypto');
const { execFile, spawnSync } = require('child_process');
const { pathToFileURL } = require('url');
const app = express();
const PORT = process.env.PORT || 3000;
const CORS_ORIGIN = process.env.CORS_ORIGIN || 'https://cos-app.vercel.app';
const CORS_ORIGINS = (process.env.CORS_ORIGINS || CORS_ORIGIN)
  .split(',')
  .map(origin => origin.trim())
  .filter(Boolean);
const UPLOAD_PASSWORD = process.env.UPLOAD_PASSWORD || '';
const GENERAL_PASSWORD = process.env.GENERAL_PASSWORD || UPLOAD_PASSWORD;
const DEAN_PASSWORD = process.env.DEAN_PASSWORD || '';
const EM_PASSWORD = process.env.EM_PASSWORD || UPLOAD_PASSWORD;
const DEV_PASSWORD = process.env.DEV_PASSWORD || '';
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';
const ENROLLMENT_SESSION_TTL_MS = Number(process.env.ENROLLMENT_SESSION_TTL_MS || 30 * 60 * 1000);
const enrollmentSessions = new Map();
const ROLE_LEVEL = {
  general: 1,
  dean: 2,
  em: 3,
  development: 4,
  admin: 5
};
const ROLE_LABEL = {
  general: 'General',
  dean: 'Dean / Division Chair',
  em: 'Enrollment Management',
  development: 'Development',
  admin: 'Administrator'
};
const ROLE_PASSWORDS = [
  ['admin', ADMIN_PASSWORD],
  ['development', DEV_PASSWORD],
  ['em', EM_PASSWORD],
  ['dean', DEAN_PASSWORD],
  ['general', GENERAL_PASSWORD]
];

// Enable CORS for your frontend
app.use(cors({
  origin(origin, callback) {
    if (!origin) return callback(null, true);
    if (CORS_ORIGINS.includes(origin)) return callback(null, true);
    if (/^https:\/\/cos-app(?:-[a-z0-9-]+)?\.vercel\.app$/i.test(origin)) return callback(null, true);
    return callback(new Error(`CORS origin not allowed: ${origin}`));
  }
}));

app.use(express.json({ limit: '50mb' }));

// Directory to store all uploaded data: schedules, rooms, modality definitions,
// CAL-GETC mappings, and temporary conversion files. Hosted deployments must
// point this at a persistent disk so redeploys do not reset edited imports.
const configuredDataDir = process.env.DATA_DIR || process.env.SCHEDULE_DATA_DIR;
const renderDiskDataDir = fs.existsSync('/var/data') ? path.join('/var/data', 'cos-app') : '';
const hostedRuntime = process.env.NODE_ENV === 'production' || Boolean(process.env.RENDER);
if (hostedRuntime && !configuredDataDir && !renderDiskDataDir) {
  console.error('Persistent upload storage is not configured. Set DATA_DIR or SCHEDULE_DATA_DIR to a mounted persistent disk path.');
  process.exit(1);
}
const DEFAULT_DATA_DIR = renderDiskDataDir || path.join(__dirname, 'schedules');
const DATA_DIR = path.resolve(configuredDataDir || DEFAULT_DATA_DIR);
if (!fs.existsSync(DATA_DIR)) {
  fs.mkdirSync(DATA_DIR, { recursive: true });
}
const ROOM_CATALOG_PATH = path.join(DATA_DIR, 'rooms.json');
const MODALITY_DEFINITIONS_PATH = path.join(DATA_DIR, 'modalities.json');
const CAL_GETC_MAPPING_PATH = path.join(DATA_DIR, 'cal-getc-mapping.json');
const CURRICULUM_CROSSWALK_PATH = path.join(DATA_DIR, 'curriculum-crosswalk.json');
const ENROLLMENT_SNAPSHOTS_PATH = path.join(DATA_DIR, 'enrollment-snapshots.json');
const CONVERT_DIR = path.join(DATA_DIR, 'conversions');
const MAX_DOCX_CONVERSION_BYTES = Number(process.env.MAX_DOCX_CONVERSION_BYTES || 15 * 1024 * 1024);
const PDF_CONVERSION_UNAVAILABLE_MESSAGE = 'PDF conversion is unavailable on the server. Please export DOCX and save as PDF from Word.';
const DOCX_MIME_TYPE = 'application/vnd.openxmlformats-officedocument.wordprocessingml.document';
const EMAIL_PROVIDER = String(process.env.SCHEDULE_CHANGE_EMAIL_PROVIDER || process.env.EMAIL_PROVIDER || '').trim().toLowerCase();
const MICROSOFT_GRAPH_DRAFT_SUPPORTED = /^true$/i.test(String(process.env.MICROSOFT_GRAPH_DRAFT_ENABLED || process.env.SCHEDULE_CHANGE_GRAPH_DRAFT_ENABLED || ''));
const DIRECT_BACKEND_SEND_SUPPORTED = /^true$/i.test(String(process.env.SCHEDULE_CHANGE_DIRECT_SEND_ENABLED || '')) && Boolean(EMAIL_PROVIDER);
const EMAIL_AUDIT_LOG_PATH = path.join(DATA_DIR, 'schedule-change-email-audit.jsonl');
const MAX_EMAIL_PAYLOAD_BYTES = Number(process.env.MAX_SCHEDULE_CHANGE_EMAIL_BYTES || 20 * 1024 * 1024);
const EMAIL_RATE_LIMIT_WINDOW_MS = Number(process.env.SCHEDULE_CHANGE_EMAIL_RATE_WINDOW_MS || 15 * 60 * 1000);
const EMAIL_RATE_LIMIT_MAX = Number(process.env.SCHEDULE_CHANGE_EMAIL_RATE_MAX || 20);
const emailRateLimit = new Map();
const ANALYTICS_ARCHIVE_DIR = path.join(DATA_DIR, 'analytics-archive');
const FACULTY_SCHEDULES_DIR = path.join(DATA_DIR, 'faculty-schedules');
if (!fs.existsSync(CONVERT_DIR)) {
  fs.mkdirSync(CONVERT_DIR, { recursive: true });
}
if (!fs.existsSync(ANALYTICS_ARCHIVE_DIR)) {
  fs.mkdirSync(ANALYTICS_ARCHIVE_DIR, { recursive: true });
}
if (!fs.existsSync(FACULTY_SCHEDULES_DIR)) {
  fs.mkdirSync(FACULTY_SCHEDULES_DIR, { recursive: true });
}

const DEFAULT_MODALITY_DEFINITIONS = [
  { code: 'IP', modality: 'In Person', omitted: false },
  { code: 'ONL', modality: 'Online', omitted: false },
  { code: 'HYB', modality: 'Hybrid', omitted: false },
  { code: 'DE', modality: 'Dual Enrollment', omitted: false },
  { code: 'FLX', modality: 'Flex', omitted: false },
  { code: '02S', modality: 'In Person', omitted: false },
  { code: '022', modality: 'In Person', omitted: false },
  { code: 'OL', modality: 'Online', omitted: false },
  { code: 'ONN', modality: 'Online', omitted: false },
  { code: 'O1', modality: 'Online', omitted: false },
  { code: 'ONS', modality: 'Online', omitted: false },
  { code: '02N', modality: 'In Person', omitted: false },
  { code: 'CPL', modality: 'Omitted from modality analysis', omitted: true },
  { code: '20', modality: 'Omitted from modality analysis', omitted: true }
];

function getSchedulePath(term) {
  if (!/^[a-z0-9 _-]+$/i.test(term)) return null;
  const filePath = path.resolve(DATA_DIR, `${term}.csv`);
  const dataRoot = path.resolve(DATA_DIR) + path.sep;
  return filePath.startsWith(dataRoot) ? filePath : null;
}

function getAnalyticsArchivePath(term) {
  if (!/^[a-z0-9 _-]+$/i.test(term)) return null;
  const filePath = path.resolve(ANALYTICS_ARCHIVE_DIR, `${term}.csv`);
  const dataRoot = path.resolve(ANALYTICS_ARCHIVE_DIR) + path.sep;
  return filePath.startsWith(dataRoot) ? filePath : null;
}

function getFacultySchedulePath(term) {
  if (!/^[a-z0-9 _-]+$/i.test(term)) return null;
  const filePath = path.resolve(FACULTY_SCHEDULES_DIR, `${term}.json`);
  const dataRoot = path.resolve(FACULTY_SCHEDULES_DIR) + path.sep;
  return filePath.startsWith(dataRoot) ? filePath : null;
}

function passwordMatches(password, expectedPassword) {
  if (!expectedPassword) return false;
  if (typeof password !== 'string') return false;
  const expected = Buffer.from(expectedPassword);
  const supplied = Buffer.from(password);
  return expected.length === supplied.length && crypto.timingSafeEqual(expected, supplied);
}

function isAuthorized(password) {
  return passwordMatches(password, GENERAL_PASSWORD);
}

function authenticateRolePassword(password, minimumRole = 'general') {
  const requiredLevel = ROLE_LEVEL[minimumRole] || ROLE_LEVEL.general;
  for (const [role, expectedPassword] of ROLE_PASSWORDS) {
    if (passwordMatches(password, expectedPassword) && ROLE_LEVEL[role] >= requiredLevel) {
      return role;
    }
  }
  return '';
}

function issueEnrollmentSession(role = 'em') {
  const token = crypto.randomBytes(32).toString('base64url');
  const expiresAtMs = Date.now() + ENROLLMENT_SESSION_TTL_MS;
  enrollmentSessions.set(token, { expiresAtMs, role });
  return {
    token,
    role,
    roleLabel: ROLE_LABEL[role] || role,
    roleLevel: ROLE_LEVEL[role] || 0,
    expiresAt: new Date(expiresAtMs).toISOString(),
    expiresInSeconds: Math.floor(ENROLLMENT_SESSION_TTL_MS / 1000)
  };
}

function cleanupEnrollmentSessions() {
  const now = Date.now();
  for (const [token, session] of enrollmentSessions.entries()) {
    const expiresAtMs = typeof session === 'number' ? session : session?.expiresAtMs;
    if (!expiresAtMs || expiresAtMs <= now) enrollmentSessions.delete(token);
  }
}

function getBearerToken(req) {
  const header = req.get('authorization') || '';
  const match = header.match(/^Bearer\s+(.+)$/i);
  return match ? match[1].trim() : '';
}

function isEnrollmentSessionAuthorized(req) {
  cleanupEnrollmentSessions();
  const token = getBearerToken(req);
  if (!token) return false;
  const session = enrollmentSessions.get(token);
  const expiresAtMs = typeof session === 'number' ? session : session?.expiresAtMs;
  if (!expiresAtMs || expiresAtMs <= Date.now()) {
    enrollmentSessions.delete(token);
    return false;
  }
  return true;
}

function enrollmentSessionRole(req) {
  cleanupEnrollmentSessions();
  const token = getBearerToken(req);
  if (!token) return '';
  const session = enrollmentSessions.get(token);
  const expiresAtMs = typeof session === 'number' ? session : session?.expiresAtMs;
  if (!expiresAtMs || expiresAtMs <= Date.now()) {
    enrollmentSessions.delete(token);
    return '';
  }
  return typeof session === 'number' ? 'em' : session?.role || 'em';
}

function normalizeRoomCatalog(rooms) {
  if (!Array.isArray(rooms)) return null;
  const normalized = [];
  for (const item of rooms) {
    if (!item || typeof item !== 'object') continue;
    const campus = String(item.campus || item.Campus || '').trim();
    const building = String(item.building || item.Building || '').trim();
    const room = String(item.room || item.Room || '').trim();
    const type = String(item.type || item.Type || item.roomType || item['Room Type'] || '').trim();
    const rawCapacity = item.capacity ?? item.Capacity ?? item.cap ?? item.Cap;
    const capacity = rawCapacity === '' || rawCapacity == null ? null : Number(rawCapacity);
    const rawPriorityDivision1 = roomCatalogField(item, [
      'rawPriorityDivision1',
      'priorityDivision1',
      'Priority Division 1',
      'Priority Division',
      'Room Priority',
      'Primary Division',
      'Assigned Division',
      'Preferred Division',
      'Dean Area',
      'Priority Area',
      'priority',
      'roomPriority'
    ]);
    const rawPriorityDivision2 = roomCatalogField(item, [
      'rawPriorityDivision2',
      'priorityDivision2',
      'Priority Division 2',
      'Secondary Division',
      'Secondary Priority',
      'Priority 2',
      'Room Priority 2',
      'Room Priority_2'
    ]);
    const rawRoomFeatures = roomCatalogField(item, [
      'rawRoomFeatures',
      'roomFeaturesText',
      'roomFeatures',
      'Room Features',
      'Features',
      'Preferred Room Features',
      'Technology Features',
      'Instructional Features',
      'Equipment',
      'Notes'
    ]);
    const priorityDivision1 = normalizeRoomPriorityDivision(rawPriorityDivision1, 'Unassigned');
    const priorityDivision2 = normalizeRoomPriorityDivision(rawPriorityDivision2, 'None');
    const roomFeatures = normalizeRoomFeatures(rawRoomFeatures);
    if (!building || !room) continue;
    normalized.push({
      campus,
      building,
      room,
      buildingRoom: `${building}-${room}`,
      type,
      capacity: Number.isFinite(capacity) ? capacity : null,
      rawPriorityDivision1: String(rawPriorityDivision1 || '').trim(),
      rawPriorityDivision2: String(rawPriorityDivision2 || '').trim(),
      priorityDivision1,
      priorityDivision2,
      priority: priorityDivision1,
      rawRoomFeatures: String(rawRoomFeatures || '').trim(),
      roomFeatures,
      roomFeaturesText: roomFeatures.join('; ')
    });
  }
  return normalized;
}

function roomCatalogField(item, names) {
  for (const name of names) {
    const value = item?.[name];
    if (value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  const normalizeKey = value => String(value || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
  const aliases = new Set(names.map(normalizeKey));
  for (const [key, value] of Object.entries(item || {})) {
    if (aliases.has(normalizeKey(key)) && value !== undefined && value !== null && String(value).trim() !== '') return value;
  }
  return '';
}

function normalizeRoomPriorityDivision(value, blankValue) {
  const text = String(value || '').replace(/\s+/g, ' ').trim();
  if (!text) return blankValue;
  return text.toUpperCase() === 'ADMINISTRATION' ? 'Administration' : text;
}

function normalizeRoomFeatures(value) {
  const text = Array.isArray(value) ? value.join('; ') : String(value || '');
  return text
    .split(/[;,]/)
    .map(item => item.replace(/\s+/g, ' ').trim())
    .filter(Boolean);
}

function normalizeModalityDefinitions(definitions) {
  if (!Array.isArray(definitions)) return null;
  const normalized = [];
  for (const item of definitions) {
    if (!item || typeof item !== 'object') continue;
    const code = String(item.code || item.Code || item.instructionalMethod || item['Instructional Method'] || '').trim().toUpperCase();
    const modality = String(item.modality || item.Modality || item.category || item.Category || '').trim();
    const rawOmitted = item.omitted ?? item.Omitted ?? item.omit ?? item.Omit ?? item.exclude ?? item.Exclude;
    const omitted = rawOmitted === true || String(rawOmitted || '').trim().toLowerCase() === 'true' || String(rawOmitted || '').trim().toLowerCase() === 'yes' || String(rawOmitted || '').trim() === '1';
    if (!code || (!omitted && !modality)) continue;
    normalized.push({
      code,
      modality: omitted ? (modality || 'Omitted from modality analysis') : modality,
      omitted
    });
  }
  return normalized;
}

function splitList(value) {
  if (Array.isArray(value)) {
    return value.map(item => String(item || '').trim()).filter(Boolean);
  }
  return String(value || '')
    .split(/[;,|]/)
    .map(item => item.trim())
    .filter(Boolean);
}

function normalizeCalGetcCode(value) {
  return String(value || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function normalizeCalGetcMapping(mapping) {
  if (!Array.isArray(mapping)) return null;
  const normalized = [];
  for (const item of mapping) {
    if (!item || typeof item !== 'object') continue;
    const code = normalizeCalGetcCode(item.code || item.Code || item.course || item.Course || item['Course Code']);
    const areas = splitList(item.areas || item.Areas || item.area || item.Area || item['CAL-GETC Area']);
    const divisions = splitList(item.divisions || item.Divisions || item.division || item.Division || item['CAL-GETC Division']);
    if (!code || (!areas.length && !divisions.length)) continue;
    normalized.push({ code, areas, divisions });
  }
  return normalized;
}

function normalizeCourseCode(value) {
  return String(value || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toUpperCase();
}

function normalizeCurriculumCrosswalk(crosswalk) {
  if (!Array.isArray(crosswalk)) return null;
  const normalized = [];
  for (const item of crosswalk) {
    if (!item || typeof item !== 'object') continue;
    const sourceCourse = normalizeCourseCode(item.sourceCourse || item.SourceCourse || item['Source Course'] || item.oldCourse || item['Old Course'] || item.cosCourse || item['COS Course']);
    const synonymCourse = normalizeCourseCode(item.synonymCourse || item.SynonymCourse || item['Synonym Course'] || item.newCourse || item['New Course'] || item.commonCourse || item['Common Course']);
    if (!sourceCourse || !synonymCourse) continue;
    normalized.push({
      sourceCourse,
      synonymCourse,
      sourceTitle: String(item.sourceTitle || item.SourceTitle || item['Source Title'] || item.cosTitle || item['COS Title'] || '').trim(),
      synonymTitle: String(item.synonymTitle || item.SynonymTitle || item['Synonym Title'] || item.commonTitle || item['Common Title'] || '').trim(),
      changeType: String(item.changeType || item.ChangeType || item['Change Type'] || item.type || item.Type || 'Curriculum Crosswalk').trim(),
      phase: String(item.phase || item.Phase || '').trim(),
      cid: String(item.cid || item.CID || item['C-ID'] || '').trim(),
      template: String(item.template || item.Template || '').trim(),
      effectiveTerm: String(item.effectiveTerm || item.EffectiveTerm || item['Effective Term'] || '').trim(),
      notes: String(item.notes || item.Notes || '').trim()
    });
  }
  return normalized;
}

function readRoomCatalog() {
  if (!fs.existsSync(ROOM_CATALOG_PATH)) {
    return { lastUpdated: null, data: [] };
  }
  const json = fs.readFileSync(ROOM_CATALOG_PATH, 'utf8');
  const stats = fs.statSync(ROOM_CATALOG_PATH);
  return {
    lastUpdated: stats.mtime.toISOString(),
    data: JSON.parse(json)
  };
}

function readModalityDefinitions() {
  if (!fs.existsSync(MODALITY_DEFINITIONS_PATH)) {
    return { lastUpdated: null, data: DEFAULT_MODALITY_DEFINITIONS };
  }
  const json = fs.readFileSync(MODALITY_DEFINITIONS_PATH, 'utf8');
  const stats = fs.statSync(MODALITY_DEFINITIONS_PATH);
  return {
    lastUpdated: stats.mtime.toISOString(),
    data: JSON.parse(json)
  };
}

function readCalGetcMapping() {
  if (!fs.existsSync(CAL_GETC_MAPPING_PATH)) {
    return { lastUpdated: null, data: [] };
  }
  const json = fs.readFileSync(CAL_GETC_MAPPING_PATH, 'utf8');
  const stats = fs.statSync(CAL_GETC_MAPPING_PATH);
  return {
    lastUpdated: stats.mtime.toISOString(),
    data: JSON.parse(json)
  };
}

function readCurriculumCrosswalk() {
  if (!fs.existsSync(CURRICULUM_CROSSWALK_PATH)) {
    return { lastUpdated: null, data: [] };
  }
  const json = fs.readFileSync(CURRICULUM_CROSSWALK_PATH, 'utf8');
  const stats = fs.statSync(CURRICULUM_CROSSWALK_PATH);
  return {
    lastUpdated: stats.mtime.toISOString(),
    data: JSON.parse(json)
  };
}

function snapshotKey(record) {
  return [
    String(record.term || record.Term || '').trim().toUpperCase(),
    String(record.crn || record.CRN || '').trim().toUpperCase(),
    String(record.snapshotType || record['Snapshot Type'] || '').trim().toUpperCase()
  ].join('|');
}

function normalizeEnrollmentSnapshotRecords(records) {
  if (!Array.isArray(records)) return null;
  const normalized = [];
  for (const item of records) {
    if (!item || typeof item !== 'object') continue;
    const record = {
      term: String(item.term || item.Term || '').trim().toUpperCase(),
      crn: String(item.crn || item.CRN || '').trim().toUpperCase(),
      snapshotType: String(item.snapshotType || item['Snapshot Type'] || '').trim().toUpperCase(),
      snapshotDate: String(item.snapshotDate || item['Snapshot Date'] || '').trim(),
      enrollment: Number(item.enrollment ?? item.Enrollment ?? 0),
      sourceFieldUsed: String(item.sourceFieldUsed || item['Source Field Used'] || '').trim(),
      subject: String(item.subject || item.Subject || '').trim().toUpperCase(),
      course: String(item.course || item.Course || '').trim().toUpperCase(),
      section: String(item.section || item.Section || '').trim().toUpperCase(),
      courseTitle: String(item.courseTitle || item['Course Title'] || item.title || '').trim(),
      division: String(item.division || item.Division || '').trim().toUpperCase(),
      department: String(item.department || item.Department || '').trim().toUpperCase(),
      campus: String(item.campus || item.Campus || '').trim().toUpperCase(),
      building: String(item.building || item.Building || '').trim().toUpperCase(),
      room: String(item.room || item.Room || '').trim().toUpperCase(),
      startDate: String(item.startDate || item['Start Date'] || '').trim(),
      endDate: String(item.endDate || item['End Date'] || '').trim(),
      capacity: Number(item.capacity ?? item.Capacity ?? 0),
      waitlist: Number(item.waitlist ?? item.Waitlist ?? 0),
      uploadedAt: String(item.uploadedAt || item['Uploaded At'] || new Date().toISOString()).trim(),
      batchId: String(item.batchId || item['Batch ID'] || '').trim(),
      action: String(item.action || item.Action || '').trim()
    };
    if (!record.term || !record.crn || !record.snapshotType || !record.snapshotDate || !Number.isFinite(record.enrollment)) continue;
    normalized.push(record);
  }
  return normalized;
}

function readEnrollmentSnapshots() {
  if (!fs.existsSync(ENROLLMENT_SNAPSHOTS_PATH)) {
    return { lastUpdated: null, data: [] };
  }
  const json = fs.readFileSync(ENROLLMENT_SNAPSHOTS_PATH, 'utf8');
  const stats = fs.statSync(ENROLLMENT_SNAPSHOTS_PATH);
  return {
    lastUpdated: stats.mtime.toISOString(),
    data: JSON.parse(json)
  };
}

function writeEnrollmentSnapshots(records) {
  fs.writeFileSync(ENROLLMENT_SNAPSHOTS_PATH, JSON.stringify(records, null, 2));
}

function safeFilename(name, fallback) {
  const clean = String(name || '').replace(/[^a-z0-9_. -]/gi, '_').replace(/\s+/g, ' ').trim().slice(0, 120);
  return clean || fallback;
}

function contentDispositionFilename(filename) {
  return String(filename || 'download.pdf').replace(/["\r\n]/g, '_');
}

function runCommand(command, args, options = {}) {
  const startedAt = Date.now();
  return new Promise((resolve, reject) => {
    execFile(command, args, options, (error, stdout, stderr) => {
      const result = {
        command,
        args,
        stdout: String(stdout || ''),
        stderr: String(stderr || ''),
        exitCode: error?.code ?? 0,
        signal: error?.signal || null,
        durationMs: Date.now() - startedAt
      };
      if (error) {
        error.stdout = result.stdout;
        error.stderr = result.stderr;
        error.exitCode = result.exitCode;
        error.signal = result.signal;
        error.durationMs = result.durationMs;
        error.commandResult = result;
        reject(error);
        return;
      }
      resolve(result);
    });
  });
}

function detectDocxPdfConverter(options = {}) {
  const commands = [
    options.libreOfficePath || process.env.LIBREOFFICE_PATH,
    'soffice',
    'libreoffice'
  ].filter(Boolean);
  const attempts = [];
  for (const command of commands) {
    const result = spawnSync(command, ['--version'], {
      encoding: 'utf8',
      windowsHide: true,
      timeout: 5000
    });
    const versionOutput = String(`${result.stdout || ''} ${result.stderr || ''}`).trim();
    attempts.push({
      command,
      exitCode: result.status,
      error: result.error?.message || '',
      stdout: String(result.stdout || '').trim(),
      stderr: String(result.stderr || '').trim()
    });
    if (!result.error && result.status === 0) {
      const version = versionOutput || 'Version unavailable';
      return {
        available: true,
        command,
        installed: true,
        converter: /libreoffice/i.test(`${result.stdout} ${result.stderr}`) ? 'libreoffice' : command,
        version,
        reason: '',
        attempts,
        notes: [`DOCX-to-PDF conversion available through ${command}.`, `LibreOffice version: ${version}`]
      };
    }
  }
  const reason = attempts.length
    ? `LibreOffice/soffice was not found or could not run. Attempts: ${attempts.map(item => `${item.command}${item.error ? ` (${item.error})` : item.exitCode == null ? '' : ` (exit ${item.exitCode})`}`).join('; ')}`
    : 'LibreOffice/soffice was not configured.';
  console.error(`DOCX-to-PDF conversion unavailable: ${reason}`);
  return {
    available: false,
    command: '',
    installed: false,
    converter: 'unavailable',
    version: '',
    reason,
    attempts,
    notes: [PDF_CONVERSION_UNAVAILABLE_MESSAGE, reason]
  };
}

const DOCX_PDF_CAPABILITY = detectDocxPdfConverter();
console.log('[DOCX-PDF] Startup converter diagnostics:', JSON.stringify({
  available: DOCX_PDF_CAPABILITY.available,
  installed: DOCX_PDF_CAPABILITY.installed,
  command: DOCX_PDF_CAPABILITY.command,
  version: DOCX_PDF_CAPABILITY.version,
  reason: DOCX_PDF_CAPABILITY.reason || ''
}));

async function convertDocxToPdf(inputPath, outputDir, options = {}) {
  const commands = options.commands || (DOCX_PDF_CAPABILITY.available
    ? [DOCX_PDF_CAPABILITY.command]
    : [process.env.LIBREOFFICE_PATH, 'soffice', 'libreoffice'].filter(Boolean));
  let lastError = null;
  const attempts = [];
  for (const commandConfig of commands) {
    const command = typeof commandConfig === 'string' ? commandConfig : commandConfig.command;
    const argsPrefix = typeof commandConfig === 'string' ? [] : (commandConfig.argsPrefix || []);
    const profileDir = path.join(outputDir, 'lo-profile');
    fs.mkdirSync(profileDir, { recursive: true });
    const args = [
      ...argsPrefix,
      '--headless',
      '--nologo',
      '--nofirststartwizard',
      '--nodefault',
      '--nolockcheck',
      `-env:UserInstallation=${pathToFileURL(profileDir).href}`,
      '--convert-to',
      'pdf',
      '--outdir',
      outputDir,
      inputPath
    ];
    console.log('[DOCX-PDF] Conversion command:', JSON.stringify({ command, args }));
    try {
      const result = await runCommand(command, args, { timeout: 60000, windowsHide: true });
      attempts.push(result);
      console.log('[DOCX-PDF] Conversion result:', JSON.stringify({
        command,
        exitCode: result.exitCode,
        durationMs: result.durationMs,
        stdout: result.stdout,
        stderr: result.stderr
      }));
      const outputPath = path.join(outputDir, `${path.basename(inputPath, path.extname(inputPath))}.pdf`);
      if (fs.existsSync(outputPath)) {
        return { outputPath, attempts };
      }
      lastError = new Error('LibreOffice finished without producing a PDF.');
      lastError.commandResult = result;
    } catch (err) {
      lastError = err;
      attempts.push(err.commandResult || {
        command,
        args,
        stdout: String(err.stdout || ''),
        stderr: String(err.stderr || ''),
        exitCode: err.exitCode ?? err.code ?? null,
        signal: err.signal || null,
        durationMs: err.durationMs || 0
      });
      console.error('[DOCX-PDF] Conversion failed:', JSON.stringify({
        command,
        exitCode: err.exitCode ?? err.code ?? null,
        durationMs: err.durationMs || 0,
        stdout: err.stdout || '',
        stderr: err.stderr || '',
        message: err.message || ''
      }));
    }
  }
  const detail = lastError?.stderr || lastError?.message || 'LibreOffice/soffice was not found.';
  const err = new Error(`DOCX-to-PDF converter unavailable or failed. ${detail}`);
  err.attempts = attempts;
  throw err;
}

function exportCapabilities() {
  return {
    docxExport: true,
    pdfFromDocx: Boolean(DOCX_PDF_CAPABILITY.available),
    libreOfficeInstalled: Boolean(DOCX_PDF_CAPABILITY.installed),
    libreOfficePath: DOCX_PDF_CAPABILITY.command || '',
    libreOfficeVersion: DOCX_PDF_CAPABILITY.version || '',
    pdfConversionAvailable: Boolean(DOCX_PDF_CAPABILITY.available),
    pdfConversionUnavailableReason: DOCX_PDF_CAPABILITY.available ? '' : DOCX_PDF_CAPABILITY.reason,
    converter: DOCX_PDF_CAPABILITY.converter,
    emailDraftSupported: true,
    microsoftGraphDraftSupported: MICROSOFT_GRAPH_DRAFT_SUPPORTED,
    mailtoFallbackSupported: true,
    directBackendSendSupported: DIRECT_BACKEND_SEND_SUPPORTED,
    emailDelivery: {
      draftSupported: true,
      graphDraft: MICROSOFT_GRAPH_DRAFT_SUPPORTED,
      mailto: true,
      backendSend: DIRECT_BACKEND_SEND_SUPPORTED,
      directBackendSend: DIRECT_BACKEND_SEND_SUPPORTED,
      provider: EMAIL_PROVIDER || 'unavailable',
      attachments: MICROSOFT_GRAPH_DRAFT_SUPPORTED || DIRECT_BACKEND_SEND_SUPPORTED,
      notes: [
        MICROSOFT_GRAPH_DRAFT_SUPPORTED
          ? 'Microsoft 365 draft creation is configured.'
          : 'Microsoft 365 draft creation is not configured. Use the mailto fallback.',
        DIRECT_BACKEND_SEND_SUPPORTED
          ? `Legacy backend direct-send scaffold enabled for ${EMAIL_PROVIDER}.`
          : 'Direct backend sending is disabled by default.'
      ]
    },
    notes: DOCX_PDF_CAPABILITY.notes
  };
}

function diagnosticsPayload() {
  return {
    nodeVersion: process.version,
    platform: process.platform,
    arch: process.arch,
    render: Boolean(process.env.RENDER),
    dataDir: DATA_DIR,
    tempConversionDir: CONVERT_DIR,
    libreOfficeInstalled: Boolean(DOCX_PDF_CAPABILITY.installed),
    libreOfficePath: DOCX_PDF_CAPABILITY.command || '',
    libreOfficeVersion: DOCX_PDF_CAPABILITY.version || '',
    pdfConversionAvailable: Boolean(DOCX_PDF_CAPABILITY.available),
    pdfConversionUnavailableReason: DOCX_PDF_CAPABILITY.available ? '' : DOCX_PDF_CAPABILITY.reason,
    converterAttempts: DOCX_PDF_CAPABILITY.attempts || []
  };
}

function normalizeEmailList(value) {
  const list = Array.isArray(value) ? value : String(value || '').split(/[;,]/);
  return list.map(item => String(item || '').trim()).filter(Boolean);
}

function isValidEmail(value) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(value || '').trim());
}

function validateScheduleChangeEmailPayload(payload) {
  const recipients = normalizeEmailList(payload.recipients || payload.to);
  const cc = normalizeEmailList(payload.cc);
  const bcc = normalizeEmailList(payload.bcc);
  const all = [...recipients, ...cc, ...bcc];
  if (!recipients.length) {
    const err = new Error('At least one recipient is required.');
    err.status = 400;
    throw err;
  }
  const invalid = all.filter(email => !isValidEmail(email));
  if (invalid.length) {
    const err = new Error(`Invalid email address: ${invalid[0]}`);
    err.status = 400;
    throw err;
  }
  const attachments = Array.isArray(payload.attachments) ? payload.attachments : [];
  const totalAttachmentBytes = attachments.reduce((total, item) => {
    const base64 = String(item?.contentBase64 || '');
    return total + Math.ceil(base64.length * 3 / 4);
  }, 0);
  if (totalAttachmentBytes > MAX_EMAIL_PAYLOAD_BYTES) {
    const err = new Error(`Email attachments exceed the ${MAX_EMAIL_PAYLOAD_BYTES} byte limit.`);
    err.status = 413;
    throw err;
  }
  attachments.forEach(item => {
    const filename = safeFilename(item?.filename, '');
    if (!/\.(docx|pdf)$/i.test(filename)) {
      const err = new Error('Only DOCX and PDF Schedule Change Form attachments are allowed.');
      err.status = 400;
      throw err;
    }
    if (!String(item?.contentBase64 || '').trim()) {
      const err = new Error(`Attachment ${filename || '(unnamed)'} is missing content.`);
      err.status = 400;
      throw err;
    }
  });
  return {
    recipients,
    cc,
    bcc,
    subject: String(payload.subject || '').trim() || 'Schedule Change Request',
    body: String(payload.body || '').trim(),
    attachments,
    metadata: payload.metadata && typeof payload.metadata === 'object' ? payload.metadata : {}
  };
}

function checkEmailRateLimit(req) {
  const key = req.ip || req.headers['x-forwarded-for'] || 'unknown';
  const now = Date.now();
  const bucket = emailRateLimit.get(key) || [];
  const recent = bucket.filter(timestamp => now - timestamp < EMAIL_RATE_LIMIT_WINDOW_MS);
  if (recent.length >= EMAIL_RATE_LIMIT_MAX) {
    const err = new Error('Too many email send attempts. Please wait and try again.');
    err.status = 429;
    throw err;
  }
  recent.push(now);
  emailRateLimit.set(key, recent);
}

function appendEmailAudit(entry) {
  fs.appendFileSync(EMAIL_AUDIT_LOG_PATH, `${JSON.stringify(entry)}\n`);
}

async function sendScheduleChangeEmail(_email) {
  const err = new Error(`Schedule Change backend email provider "${EMAIL_PROVIDER}" is not implemented in this deployment.`);
  err.status = 501;
  throw err;
}

async function createScheduleChangeEmailDraft(_email) {
  const err = new Error('Microsoft 365 draft creation is not configured. Opening local email draft instead.');
  err.status = 503;
  throw err;
}

app.post('/api/schedule-change/create-email-draft', (req, res) => {
  const timestamp = new Date().toISOString();
  let audit = {
    timestamp,
    provider: MICROSOFT_GRAPH_DRAFT_SUPPORTED ? 'microsoft-graph' : 'unavailable',
    mode: 'draft',
    status: 'received'
  };
  try {
    checkEmailRateLimit(req);
    if (!MICROSOFT_GRAPH_DRAFT_SUPPORTED) {
      audit.status = 'disabled';
      audit.error = 'Microsoft 365 draft creation is not configured.';
      appendEmailAudit(audit);
      return res.status(503).json({
        success: false,
        error: 'Microsoft 365 draft creation is not configured. Opening local email draft instead.'
      });
    }
    const email = validateScheduleChangeEmailPayload(req.body || {});
    audit = {
      ...audit,
      recipients: email.recipients,
      ccCount: email.cc.length,
      bccCount: email.bcc.length,
      subject: email.subject,
      term: email.metadata.term || '',
      crn: email.metadata.crn || '',
      course: email.metadata.course || '',
      attachmentFilenames: email.attachments.map(item => safeFilename(item.filename, 'attachment'))
    };
    return createScheduleChangeEmailDraft(email)
      .then(result => {
        audit.status = 'draft-created';
        audit.providerMessageId = result?.messageId || '';
        appendEmailAudit(audit);
        return res.json({ success: true, providerMessageId: audit.providerMessageId, webLink: result?.webLink || '' });
      })
      .catch(err => {
        audit.status = 'failed';
        audit.error = err.message || 'Email draft creation failed.';
        appendEmailAudit(audit);
        return res.status(err.status || 500).json({ success: false, error: audit.error });
      });
  } catch (err) {
    audit.status = 'rejected';
    audit.error = err.message || 'Email draft request rejected.';
    appendEmailAudit(audit);
    return res.status(err.status || 400).json({ success: false, error: audit.error });
  }
});

app.post('/api/schedule-change/send-email', (req, res) => {
  const timestamp = new Date().toISOString();
  let audit = {
    timestamp,
    provider: EMAIL_PROVIDER || 'unavailable',
    status: 'received'
  };
  try {
    checkEmailRateLimit(req);
    if (!DIRECT_BACKEND_SEND_SUPPORTED) {
      audit.status = 'disabled';
      audit.error = 'Direct backend sending is disabled.';
      appendEmailAudit(audit);
      return res.status(503).json({
        success: false,
        error: 'Direct backend sending is disabled. Use Open Email Draft or download the DOCX/PDF and send manually.'
      });
    }
    const email = validateScheduleChangeEmailPayload(req.body || {});
    audit = {
      ...audit,
      recipients: email.recipients,
      ccCount: email.cc.length,
      bccCount: email.bcc.length,
      subject: email.subject,
      term: email.metadata.term || '',
      crn: email.metadata.crn || '',
      course: email.metadata.course || '',
      user: email.metadata.user || '',
      attachmentFilenames: email.attachments.map(item => safeFilename(item.filename, 'attachment'))
    };
    return sendScheduleChangeEmail(email)
      .then(result => {
        audit.status = 'sent';
        audit.providerMessageId = result?.messageId || '';
        appendEmailAudit(audit);
        return res.json({ success: true, providerMessageId: audit.providerMessageId });
      })
      .catch(err => {
        audit.status = 'failed';
        audit.error = err.message || 'Email send failed.';
        appendEmailAudit(audit);
        return res.status(err.status || 500).json({ success: false, error: audit.error });
      });
  } catch (err) {
    audit.status = 'rejected';
    audit.error = err.message || 'Email request rejected.';
    appendEmailAudit(audit);
    return res.status(err.status || 400).json({ success: false, error: audit.error });
  }
});

function cleanupConversionDir(requestDir) {
  try {
    fs.rmSync(requestDir, { recursive: true, force: true });
    const removed = !fs.existsSync(requestDir);
    console.log('[DOCX-PDF] Cleanup status:', JSON.stringify({ requestDir, removed }));
    return { ok: removed, message: removed ? 'Temporary conversion directory removed.' : 'Temporary conversion directory still exists after cleanup.' };
  } catch (cleanupErr) {
    console.error('Conversion cleanup error:', cleanupErr);
    return { ok: false, message: cleanupErr.message || 'Cleanup failed.' };
  }
}

function decodeDocxPayload(req) {
  const isBinaryDocx = Buffer.isBuffer(req.body);
  const filename = isBinaryDocx
    ? req.get('x-filename') || req.query?.filename || 'schedule-change-form.docx'
    : req.body?.filename;
  const inputName = safeFilename(filename, 'schedule-change-form.docx').replace(/\.pdf$/i, '.docx');
  if (!/\.docx$/i.test(inputName)) {
    const err = new Error('Invalid file type. Schedule Change PDF conversion requires a .docx file.');
    err.status = 400;
    throw err;
  }
  if (!isBinaryDocx && (typeof req.body?.docxBase64 !== 'string' || !req.body.docxBase64.trim())) {
    const err = new Error('DOCX payload is required.');
    err.status = 400;
    throw err;
  }
  const buffer = isBinaryDocx ? req.body : Buffer.from(req.body.docxBase64, 'base64');
  if (!buffer.length || buffer.length > MAX_DOCX_CONVERSION_BYTES) {
    const err = new Error(`DOCX payload must be between 1 byte and ${MAX_DOCX_CONVERSION_BYTES} bytes.`);
    err.status = 413;
    throw err;
  }
  if (buffer[0] !== 0x50 || buffer[1] !== 0x4b) {
    const err = new Error('Invalid DOCX payload. Expected a DOCX/ZIP file.');
    err.status = 400;
    throw err;
  }
  return { inputName, buffer };
}

async function handleScheduleChangeDocxToPdf(req, res) {
  if (!DOCX_PDF_CAPABILITY.available) {
    return res.status(503).json({
      error: PDF_CONVERSION_UNAVAILABLE_MESSAGE,
      reason: DOCX_PDF_CAPABILITY.reason,
      capabilities: exportCapabilities()
    });
  }

  const requestId = crypto.randomBytes(12).toString('hex');
  const requestDir = path.join(CONVERT_DIR, requestId);
  fs.mkdirSync(requestDir, { recursive: true });
  const startedAt = Date.now();
  console.log('[DOCX-PDF] Conversion request started:', JSON.stringify({ requestId, requestDir }));

  try {
    const { inputName, buffer } = decodeDocxPayload(req);
    const inputPath = path.join(requestDir, inputName);
    fs.writeFileSync(inputPath, buffer, { flag: 'wx' });
    const { outputPath: pdfPath, attempts } = await convertDocxToPdf(inputPath, requestDir);
    const downloadName = path.basename(inputPath, path.extname(inputPath)) + '.pdf';
    console.log('[DOCX-PDF] Conversion request completed:', JSON.stringify({
      requestId,
      inputName,
      outputName: downloadName,
      durationMs: Date.now() - startedAt,
      attempts: attempts.map(item => ({
        command: item.command,
        exitCode: item.exitCode,
        durationMs: item.durationMs,
        stdout: item.stdout,
        stderr: item.stderr
      }))
    }));
    res.setHeader('Content-Type', 'application/pdf');
    res.setHeader('Content-Disposition', `attachment; filename="${contentDispositionFilename(downloadName)}"`);
    return res.sendFile(pdfPath, err => {
      const cleanup = cleanupConversionDir(requestDir);
      console.log('[DOCX-PDF] Response cleanup:', JSON.stringify({ requestId, cleanup }));
      if (err) console.error('PDF send error:', err);
    });
  } catch (err) {
    const cleanup = cleanupConversionDir(requestDir);
    const status = err.status || 500;
    if (status >= 500) {
      console.error('DOCX conversion error:', JSON.stringify({
        requestId,
        status,
        message: err.message || '',
        durationMs: Date.now() - startedAt,
        cleanup,
        attempts: (err.attempts || []).map(item => ({
          command: item.command,
          exitCode: item.exitCode,
          durationMs: item.durationMs,
          stdout: item.stdout,
          stderr: item.stderr
        }))
      }));
    }
    return res.status(status).json({
      error: status === 503 ? PDF_CONVERSION_UNAVAILABLE_MESSAGE : (err.message || 'DOCX-to-PDF conversion failed'),
      reason: err.message || DOCX_PDF_CAPABILITY.reason || '',
      cleanup,
      attempts: err.attempts || []
    });
  }
}

app.get('/api/export-capabilities', (_req, res) => {
  return res.json(exportCapabilities());
});

app.get('/api/admin/diagnostics', (_req, res) => {
  return res.json(diagnosticsPayload());
});

app.post('/api/auth/enrollment-management', (req, res) => {
  const { password } = req.body || {};
  const role = authenticateRolePassword(password, 'em');
  if (!role) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  return res.json(issueEnrollmentSession(role));
});

app.post('/api/auth/role', (req, res) => {
  const { password, requestedRole = 'general' } = req.body || {};
  const minimumRole = ROLE_LEVEL[requestedRole] ? requestedRole : 'general';
  const role = authenticateRolePassword(password, minimumRole);
  if (!role) {
    return res.status(403).json({ error: 'Unauthorized', requiredRole: minimumRole });
  }
  return res.json(issueEnrollmentSession(role));
});

// POST endpoint to upload schedule CSV
app.post('/api/schedule/:term', (req, res) => {
  const term = req.params.term;
  const { csv, password } = req.body;
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  if (typeof csv !== 'string') {
    return res.status(400).json({ error: 'CSV payload is required' });
  }

  const filePath = getSchedulePath(term);
  if (!filePath) {
    return res.status(400).json({ error: 'Invalid term' });
  }
  try {
    fs.writeFileSync(filePath, csv);
    const now = new Date().toISOString();
    return res.json({ success: true, lastUpdated: now });
  } catch (err) {
    console.error('Write error:', err);
    return res.status(500).json({ error: 'File write failed' });
  }
});

// GET endpoint to fetch and parse schedule CSV
app.get('/api/schedule/:term', (req, res) => {
  const term = req.params.term;
  const filePath = getSchedulePath(term);
  if (!filePath) {
    return res.status(400).json({ error: 'Invalid term' });
  }
  if (!fs.existsSync(filePath)) {
    return res.json({ lastUpdated: null, data: [] });
  }

  try {
    const csv = fs.readFileSync(filePath, 'utf8');
    const stats = fs.statSync(filePath);
    const lastUpdated = stats.mtime.toISOString();
    const parsed = Papa.parse(csv, { header: true, skipEmptyLines: true });
    return res.json({ lastUpdated, data: parsed.data });
  } catch (err) {
    console.error('Read error:', err);
    return res.status(500).json({ error: 'File read failed' });
  }
});

app.get('/api/analytics-archive', (req, res) => {
  try {
    const terms = fs.readdirSync(ANALYTICS_ARCHIVE_DIR)
      .filter(file => file.toLowerCase().endsWith('.csv'))
      .map(file => {
        const filePath = path.join(ANALYTICS_ARCHIVE_DIR, file);
        const stats = fs.statSync(filePath);
        return {
          term: path.basename(file, '.csv'),
          lastUpdated: stats.mtime.toISOString()
        };
      })
      .sort((a, b) => a.term.localeCompare(b.term, undefined, { numeric: true }));
    return res.json({ data: terms });
  } catch (err) {
    console.error('Analytics archive list error:', err);
    return res.status(500).json({ error: 'Analytics archive list failed' });
  }
});

app.get('/api/analytics-archive/:term', (req, res) => {
  const term = req.params.term;
  const filePath = getAnalyticsArchivePath(term);
  if (!filePath) {
    return res.status(400).json({ error: 'Invalid term' });
  }
  if (!fs.existsSync(filePath)) {
    return res.json({ term, lastUpdated: null, data: [] });
  }
  try {
    const csv = fs.readFileSync(filePath, 'utf8');
    const stats = fs.statSync(filePath);
    const parsed = Papa.parse(csv, { header: true, skipEmptyLines: true });
    return res.json({ term, lastUpdated: stats.mtime.toISOString(), data: parsed.data });
  } catch (err) {
    console.error('Analytics archive read error:', err);
    return res.status(500).json({ error: 'Analytics archive read failed' });
  }
});

function facultyField(row, names) {
  if (!row || typeof row !== 'object') return '';
  const entries = Object.entries(row);
  for (const name of names) {
    if (row[name] != null && String(row[name]).trim()) return String(row[name]).trim();
    const normalizedName = String(name).replace(/[^a-z0-9]/gi, '').toLowerCase();
    const found = entries.find(([key, value]) =>
      String(key).replace(/[^a-z0-9]/gi, '').toLowerCase() === normalizedName &&
      value != null &&
      String(value).trim()
    );
    if (found) return String(found[1]).trim();
  }
  return '';
}

function validateFacultyScheduleRows(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return { valid: false, error: 'Faculty schedule rows are required.' };
  }
  const has = predicate => rows.some(predicate);
  const missing = [];
  if (!has(row => facultyField(row, ['FCNT_CODE', 'fcntCode']))) missing.push('FCNT_CODE');
  if (!has(row => facultyField(row, ['FACULTYID', 'Faculty ID', 'facultyId']) || facultyField(row, ['FacultyName', 'Faculty Name', 'facultyName']))) missing.push('faculty identity');
  if (!has(row => facultyField(row, ['CRN', 'crn']))) missing.push('CRN');
  if (!has(row => facultyField(row, ['DAYS', 'Days', 'days']))) missing.push('DAYS');
  if (!has(row => facultyField(row, ['STARTTIME', 'Start Time', 'startTime']))) missing.push('STARTTIME');
  if (!has(row => facultyField(row, ['ENDTIME', 'End Time', 'endTime']))) missing.push('ENDTIME');
  if (!has(row => facultyField(row, ['SCHD_CODE_SSRMEET', 'SCHD CODE SSRMEET', 'schdCode']))) missing.push('SCHD_CODE_SSRMEET');
  if (missing.length) {
    return {
      valid: false,
      error: 'This does not appear to be a Faculty Schedule file. Faculty schedules must include FCNT_CODE, faculty identity, CRN, days, times, and SCHD_CODE_SSRMEET.'
    };
  }
  return { valid: true };
}

function facultyTypeFromCode(code) {
  const normalized = String(code || '').trim().toUpperCase();
  if (normalized === 'AE' || normalized === 'X') return 'OMIT';
  if (normalized === 'FT' || normalized === 'TE') return 'FULL_TIME';
  if (normalized === 'JP') return 'PART_TIME';
  return normalized || 'UNKNOWN';
}

function meetingTypeFromCode(code) {
  const normalized = String(code || '').replace(/\D/g, '') || String(code || '').trim().toUpperCase();
  if (normalized === '2') return 'Lecture';
  if (normalized === '4') return 'Lab';
  if (String(code || '').trim().toUpperCase() === 'XX') return 'Activity';
  return 'Other';
}

function facultyScheduleMetadata(term, rows, base = {}) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const facultyTypeCounts = {};
  const meetingTypeCounts = {};
  const faculty = new Set();
  const crns = new Set();
  const meetings = new Set();
  let omittedRowCount = 0;
  safeRows.forEach(row => {
    const fcnt = facultyField(row, ['FCNT_CODE', 'fcntCode']);
    const facultyType = facultyTypeFromCode(fcnt);
    facultyTypeCounts[facultyType] = (facultyTypeCounts[facultyType] || 0) + 1;
    if (facultyType === 'OMIT') omittedRowCount += 1;
    const meetingType = meetingTypeFromCode(facultyField(row, ['SCHD_CODE_SSRMEET', 'SCHD CODE SSRMEET', 'schdCode']));
    meetingTypeCounts[meetingType] = (meetingTypeCounts[meetingType] || 0) + 1;
    const facultyId = facultyField(row, ['FACULTYID', 'Faculty ID', 'facultyId']) || facultyField(row, ['FacultyName', 'Faculty Name', 'facultyName']);
    const crn = facultyField(row, ['CRN', 'crn']);
    if (facultyId) faculty.add(facultyId.toUpperCase());
    if (crn) crns.add(crn.toUpperCase());
    meetings.add([
      crn,
      facultyId,
      facultyField(row, ['DAYS', 'Days', 'days']),
      facultyField(row, ['STARTTIME', 'Start Time', 'startTime']),
      facultyField(row, ['ENDTIME', 'End Time', 'endTime']),
      facultyField(row, ['SCHD_CODE_SSRMEET', 'SCHD CODE SSRMEET', 'schdCode'])
    ].map(value => String(value || '').trim().toUpperCase()).join('|'));
  });
  return {
    term,
    uploadedAt: base.uploadedAt || new Date().toISOString(),
    uploadedByRole: base.uploadedByRole || '',
    sourceFileName: base.sourceFileName || '',
    rawRowCount: safeRows.length,
    normalizedMeetingCount: meetings.size,
    omittedRowCount,
    distinctFacultyCount: faculty.size,
    distinctCrnCount: crns.size,
    facultyTypeCounts,
    meetingTypeCounts
  };
}

function readFacultyScheduleArchive(term) {
  const filePath = getFacultySchedulePath(term);
  if (!filePath || !fs.existsSync(filePath)) return null;
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

app.get('/api/faculty-schedules', (req, res) => {
  try {
    const data = fs.readdirSync(FACULTY_SCHEDULES_DIR)
      .filter(file => file.toLowerCase().endsWith('.json'))
      .map(file => {
        const term = path.basename(file, '.json');
        const payload = readFacultyScheduleArchive(term);
        const stats = fs.statSync(path.join(FACULTY_SCHEDULES_DIR, file));
        return payload?.metadata || { term, uploadedAt: stats.mtime.toISOString() };
      })
      .sort((a, b) => String(a.term || '').localeCompare(String(b.term || ''), undefined, { numeric: true }));
    return res.json({ data });
  } catch (err) {
    console.error('Faculty schedule archive list error:', err);
    return res.status(500).json({ error: 'Faculty schedule archive list failed' });
  }
});

app.get('/api/faculty-schedules/:term', (req, res) => {
  const term = req.params.term;
  const filePath = getFacultySchedulePath(term);
  if (!filePath) return res.status(400).json({ error: 'Invalid term' });
  if (!fs.existsSync(filePath)) return res.json({ term, metadata: null, data: [] });
  try {
    const payload = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    return res.json({ term, metadata: payload.metadata || null, data: Array.isArray(payload.rows) ? payload.rows : [] });
  } catch (err) {
    console.error('Faculty schedule archive read error:', err);
    return res.status(500).json({ error: 'Faculty schedule archive read failed' });
  }
});

app.post('/api/faculty-schedules/:term', (req, res) => {
  const term = req.params.term;
  const { rows, password, sourceFileName = '' } = req.body || {};
  if (!isEnrollmentSessionAuthorized(req) && !isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const filePath = getFacultySchedulePath(term);
  if (!filePath) return res.status(400).json({ error: 'Invalid term' });
  const validation = validateFacultyScheduleRows(rows);
  if (!validation.valid) return res.status(400).json({ error: validation.error });
  try {
    const metadata = facultyScheduleMetadata(term, rows, {
      uploadedByRole: enrollmentSessionRole(req) || (isAuthorized(password) ? 'general' : ''),
      sourceFileName
    });
    fs.writeFileSync(filePath, JSON.stringify({ metadata, rows }, null, 2));
    return res.json({ success: true, term, metadata, data: rows });
  } catch (err) {
    console.error('Faculty schedule archive write error:', err);
    return res.status(500).json({ error: 'Faculty schedule archive write failed' });
  }
});

app.delete('/api/faculty-schedules/:term', (req, res) => {
  const term = req.params.term;
  const { password } = req.body || {};
  if (!isEnrollmentSessionAuthorized(req) && !isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const filePath = getFacultySchedulePath(term);
  if (!filePath) return res.status(400).json({ error: 'Invalid term' });
  try {
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    return res.json({ success: true, term });
  } catch (err) {
    console.error('Faculty schedule archive delete error:', err);
    return res.status(500).json({ error: 'Faculty schedule archive delete failed' });
  }
});

app.post('/api/analytics-archive/:term', (req, res) => {
  const term = req.params.term;
  const { csv, password } = req.body || {};
  if (!isEnrollmentSessionAuthorized(req) && !isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  if (typeof csv !== 'string' || !csv.trim()) {
    return res.status(400).json({ error: 'CSV payload is required' });
  }
  const filePath = getAnalyticsArchivePath(term);
  if (!filePath) {
    return res.status(400).json({ error: 'Invalid term' });
  }
  try {
    fs.writeFileSync(filePath, csv);
    const now = new Date().toISOString();
    return res.json({ success: true, term, lastUpdated: now });
  } catch (err) {
    console.error('Analytics archive write error:', err);
    return res.status(500).json({ error: 'Analytics archive write failed' });
  }
});

app.get('/api/enrollment-snapshots', (req, res) => {
  try {
    return res.json(readEnrollmentSnapshots());
  } catch (err) {
    console.error('Enrollment snapshot read error:', err);
    return res.status(500).json({ error: 'Enrollment snapshot read failed' });
  }
});

app.post('/api/enrollment-snapshots', (req, res) => {
  const { records, password } = req.body || {};
  if (!isEnrollmentSessionAuthorized(req) && !isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const incoming = normalizeEnrollmentSnapshotRecords(records);
  if (!incoming || !incoming.length) {
    return res.status(400).json({ error: 'Enrollment snapshot records are required' });
  }
  try {
    const existing = readEnrollmentSnapshots().data || [];
    const merged = new Map();
    existing.forEach(record => {
      const key = snapshotKey(record);
      if (key !== '||') merged.set(key, { ...record, action: record.action || 'Existing' });
    });
    let appended = 0;
    let updated = 0;
    incoming.forEach(record => {
      const key = snapshotKey(record);
      if (merged.has(key)) {
        updated += 1;
        merged.set(key, { ...merged.get(key), ...record, action: 'Updated' });
      } else {
        appended += 1;
        merged.set(key, { ...record, action: 'Appended' });
      }
    });
    const data = [...merged.values()].sort((a, b) =>
      String(a.term).localeCompare(String(b.term), undefined, { numeric: true }) ||
      String(a.snapshotType).localeCompare(String(b.snapshotType)) ||
      String(a.crn).localeCompare(String(b.crn), undefined, { numeric: true })
    );
    writeEnrollmentSnapshots(data);
    return res.json({ success: true, appended, updated, count: data.length, lastUpdated: new Date().toISOString(), data });
  } catch (err) {
    console.error('Enrollment snapshot write error:', err);
    return res.status(500).json({ error: 'Enrollment snapshot write failed' });
  }
});

app.delete('/api/enrollment-snapshots', (req, res) => {
  const { term, snapshotType, snapshotDate, password } = req.body || {};
  if (!isEnrollmentSessionAuthorized(req) && !isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const termFilter = String(term || '').trim().toUpperCase();
  const typeFilter = String(snapshotType || '').trim().toUpperCase();
  const dateFilter = String(snapshotDate || '').trim();
  if (!termFilter || !typeFilter || !dateFilter) {
    return res.status(400).json({ error: 'Term, snapshot type, and snapshot date are required to clear a batch.' });
  }
  try {
    const existing = readEnrollmentSnapshots().data || [];
    const data = existing.filter(record =>
      String(record.term || '').toUpperCase() !== termFilter ||
      String(record.snapshotType || '').toUpperCase() !== typeFilter ||
      String(record.snapshotDate || '') !== dateFilter
    );
    const deleted = existing.length - data.length;
    writeEnrollmentSnapshots(data);
    return res.json({ success: true, deleted, removed: deleted, count: data.length, lastUpdated: new Date().toISOString(), data });
  } catch (err) {
    console.error('Enrollment snapshot delete error:', err);
    return res.status(500).json({ error: 'Enrollment snapshot delete failed' });
  }
});

app.get('/api/rooms', (req, res) => {
  try {
    return res.json(readRoomCatalog());
  } catch (err) {
    console.error('Room catalog read error:', err);
    return res.status(500).json({ error: 'Room catalog read failed' });
  }
});

app.post('/api/rooms/export', (req, res) => {
  const { password } = req.body || {};
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  try {
    return res.json(readRoomCatalog());
  } catch (err) {
    console.error('Room catalog export error:', err);
    return res.status(500).json({ error: 'Room catalog export failed' });
  }
});

app.post('/api/rooms/import', (req, res) => {
  const { password, rooms } = req.body || {};
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const normalized = normalizeRoomCatalog(rooms);
  if (!normalized || !normalized.length) {
    return res.status(400).json({ error: 'Room catalog payload is required' });
  }
  try {
    fs.writeFileSync(ROOM_CATALOG_PATH, JSON.stringify(normalized, null, 2));
    const now = new Date().toISOString();
    return res.json({ success: true, lastUpdated: now, count: normalized.length, data: normalized });
  } catch (err) {
    console.error('Room catalog write error:', err);
    return res.status(500).json({ error: 'Room catalog write failed' });
  }
});

app.get('/api/modalities', (req, res) => {
  try {
    return res.json(readModalityDefinitions());
  } catch (err) {
    console.error('Modality definitions read error:', err);
    return res.status(500).json({ error: 'Modality definitions read failed' });
  }
});

app.post('/api/modalities/import', (req, res) => {
  const { password, definitions } = req.body || {};
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const normalized = normalizeModalityDefinitions(definitions);
  if (!normalized || !normalized.length) {
    return res.status(400).json({ error: 'Modality definitions payload is required' });
  }
  try {
    fs.writeFileSync(MODALITY_DEFINITIONS_PATH, JSON.stringify(normalized, null, 2));
    const now = new Date().toISOString();
    return res.json({ success: true, lastUpdated: now, count: normalized.length, data: normalized });
  } catch (err) {
    console.error('Modality definitions write error:', err);
    return res.status(500).json({ error: 'Modality definitions write failed' });
  }
});

app.get('/api/cal-getc', (req, res) => {
  try {
    return res.json(readCalGetcMapping());
  } catch (err) {
    console.error('CAL-GETC mapping read error:', err);
    return res.status(500).json({ error: 'CAL-GETC mapping read failed' });
  }
});

app.post('/api/cal-getc/import', (req, res) => {
  const { password, mapping } = req.body || {};
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const normalized = normalizeCalGetcMapping(mapping);
  if (!normalized || !normalized.length) {
    return res.status(400).json({ error: 'CAL-GETC mapping payload is required' });
  }
  try {
    fs.writeFileSync(CAL_GETC_MAPPING_PATH, JSON.stringify(normalized, null, 2));
    const now = new Date().toISOString();
    return res.json({ success: true, lastUpdated: now, count: normalized.length, data: normalized });
  } catch (err) {
    console.error('CAL-GETC mapping write error:', err);
    return res.status(500).json({ error: 'CAL-GETC mapping write failed' });
  }
});

app.get('/api/curriculum-crosswalk', (req, res) => {
  try {
    return res.json(readCurriculumCrosswalk());
  } catch (err) {
    console.error('Curriculum crosswalk read error:', err);
    return res.status(500).json({ error: 'Curriculum crosswalk read failed' });
  }
});

app.post('/api/curriculum-crosswalk/import', (req, res) => {
  const { password, crosswalk } = req.body || {};
  if (!isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const normalized = normalizeCurriculumCrosswalk(crosswalk);
  if (!normalized || !normalized.length) {
    return res.status(400).json({ error: 'Curriculum crosswalk payload is required' });
  }
  try {
    fs.writeFileSync(CURRICULUM_CROSSWALK_PATH, JSON.stringify(normalized, null, 2));
    const now = new Date().toISOString();
    return res.json({ success: true, lastUpdated: now, count: normalized.length, data: normalized });
  } catch (err) {
    console.error('Curriculum crosswalk write error:', err);
    return res.status(500).json({ error: 'Curriculum crosswalk write failed' });
  }
});

const docxRawParser = express.raw({ type: DOCX_MIME_TYPE, limit: MAX_DOCX_CONVERSION_BYTES });
app.post('/api/schedule-change/convert-docx-to-pdf', docxRawParser, handleScheduleChangeDocxToPdf);
app.post('/api/convert/docx-to-pdf', docxRawParser, handleScheduleChangeDocxToPdf);

if (require.main === module) {
  app.listen(PORT, () => console.log(`Listening on port ${PORT}`));
}

module.exports = {
  app,
  detectDocxPdfConverter,
  exportCapabilities,
  diagnosticsPayload,
  safeFilename,
  contentDispositionFilename,
  cleanupConversionDir,
  convertDocxToPdf,
  runCommand
};
