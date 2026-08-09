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
const DIV_CHAIR_PASSWORD = process.env.DIV_CHAIR_PASSWORD || '';
const DEAN_PASSWORD = process.env.DEAN_PASSWORD || '';
const EM_PASSWORD = process.env.EM_PASSWORD || UPLOAD_PASSWORD;
const DEV_PASSWORD = process.env.DEV_PASSWORD || '';
const ADMIN_PASSWORD = process.env.ADMIN_PASSWORD || '';
const ENROLLMENT_SESSION_TTL_MS = Number(process.env.ENROLLMENT_SESSION_TTL_MS || 30 * 60 * 1000);
const enrollmentSessions = new Map();
const ROLE_LEVEL = {
  general: 1,
  divchair: 2,
  dean: 3,
  em: 3,
  development: 4,
  admin: 5
};
const ROLE_LABEL = {
  general: 'General',
  divchair: 'Division Chair / Administrative Assistant',
  dean: 'Dean / Enrollment Management',
  em: 'Dean / Enrollment Management',
  development: 'Developer',
  admin: 'System Administrator'
};
const ROLE_PASSWORDS = [
  ['admin', ADMIN_PASSWORD],
  ['development', DEV_PASSWORD],
  ['dean', DEAN_PASSWORD],
  ['dean', EM_PASSWORD],
  ['divchair', DIV_CHAIR_PASSWORD],
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
const AUTH_FAILURE_LIMIT = Number(process.env.AUTH_FAILURE_LIMIT || 5);
const AUTH_FAILURE_WINDOW_MS = Number(process.env.AUTH_FAILURE_WINDOW_MS || 15 * 60 * 1000);
const AUTH_LOCKOUT_MS = Number(process.env.AUTH_LOCKOUT_MS || 15 * 60 * 1000);
const authFailureState = new Map();
const CONVERSION_RATE_LIMIT_WINDOW_MS = Number(process.env.CONVERSION_RATE_LIMIT_WINDOW_MS || 15 * 60 * 1000);
const CONVERSION_RATE_LIMIT_MAX = Number(process.env.CONVERSION_RATE_LIMIT_MAX || 10);
const CONVERSION_MAX_CONCURRENT = Math.max(1, Number(process.env.CONVERSION_MAX_CONCURRENT || 3));
const conversionRateLimit = new Map();
let activeConversions = 0;
const ANALYTICS_ARCHIVE_DIR = path.join(DATA_DIR, 'analytics-archive');
const ANALYTICS_ARCHIVE_MANIFEST_PATH = path.join(ANALYTICS_ARCHIVE_DIR, 'manifest.json');
const ANALYTICS_ARCHIVE_MANIFEST_SCHEMA_VERSION = 1;
const FACULTY_SCHEDULES_DIR = path.join(DATA_DIR, 'faculty-schedules');
const WORK_EXPERIENCE_DIR = path.join(DATA_DIR, 'work-experience');
const LOW_ENROLLMENT_TRACKING_DIR = path.join(DATA_DIR, 'low-enrollment-tracking');
if (!fs.existsSync(CONVERT_DIR)) {
  fs.mkdirSync(CONVERT_DIR, { recursive: true });
}
if (!fs.existsSync(ANALYTICS_ARCHIVE_DIR)) {
  fs.mkdirSync(ANALYTICS_ARCHIVE_DIR, { recursive: true });
}
if (!fs.existsSync(FACULTY_SCHEDULES_DIR)) {
  fs.mkdirSync(FACULTY_SCHEDULES_DIR, { recursive: true });
}
if (!fs.existsSync(WORK_EXPERIENCE_DIR)) {
  fs.mkdirSync(WORK_EXPERIENCE_DIR, { recursive: true });
}
if (!fs.existsSync(LOW_ENROLLMENT_TRACKING_DIR)) {
  fs.mkdirSync(LOW_ENROLLMENT_TRACKING_DIR, { recursive: true });
}

const DEFAULT_MODALITY_DEFINITIONS = [
  ...['ONL', '71', '72', 'O1', 'OL', 'ONN', 'ONS', 'OO', 'OS', 'OSS', 'OT', 'OTS', 'ON', 'OSL']
    .map(code => ({ code, modality: 'Online', omitted: false })),
  ...['IP', '02', '22', '022', '02H', '02O', '02S', '02T', '02N', '04', '06', '07', '08', '09', '12', 'XX', 'YY']
    .map(code => ({ code, modality: 'In-Person', omitted: false })),
  ...['HYB', 'OH', 'OHF', 'FLX', 'OHS']
    .map(code => ({ code, modality: 'Hybrid', omitted: false })),
  { code: 'DE', modality: 'Dual Enrollment', omitted: false },
  { code: '20', modality: 'Work Experience', omitted: false },
  ...['CPL', 'CBE', '98']
    .map(code => ({ code, modality: 'Omitted from modality analysis', omitted: true }))
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

function displayTermFromArchiveTerm(term) {
  const value = String(term || '').trim();
  const codeMatch = value.match(/^(\d{4})(\d{2})$/);
  if (codeMatch) {
    const year = Number(codeMatch[1]);
    const seasonCode = codeMatch[2];
    if (seasonCode === '10') return `Fall ${year - 1}`;
    if (seasonCode === '20') return `Spring ${year}`;
    if (seasonCode === '30') return `Summer ${year}`;
  }
  return value.replace(/[_-]+/g, ' ').replace(/\s+/g, ' ').trim();
}

function archiveTermSortValue(term) {
  const value = String(term || '').trim().toUpperCase();
  const codeMatch = value.match(/^(\d{4})(\d{2})$/);
  if (codeMatch) {
    const year = Number(codeMatch[1]);
    const seasonCode = codeMatch[2];
    const seasonOrder = seasonCode === '10' ? 1 : seasonCode === '20' ? 2 : seasonCode === '30' ? 3 : 0;
    return year * 10 + seasonOrder;
  }
  const labelMatch = value.match(/\b(SPRING|SUMMER|FALL)\b\D*(\d{4})/i) || value.match(/\b(\d{4})\D*(SPRING|SUMMER|FALL)\b/i);
  if (labelMatch) {
    const season = (Number(labelMatch[2]) ? labelMatch[1] : labelMatch[2]).toUpperCase();
    const year = Number(Number(labelMatch[2]) ? labelMatch[2] : labelMatch[1]);
    const seasonOrder = season === 'SPRING' ? 1 : season === 'SUMMER' ? 2 : season === 'FALL' ? 3 : 0;
    return year * 10 + seasonOrder;
  }
  return Number.NEGATIVE_INFINITY;
}

function sortArchiveTermsNewestFirst(terms = []) {
  return terms.slice().sort((a, b) => {
    const av = archiveTermSortValue(a.termCode || a.term || '');
    const bv = archiveTermSortValue(b.termCode || b.term || '');
    return bv - av || String(b.termCode || b.term || '').localeCompare(String(a.termCode || a.term || ''), undefined, { numeric: true });
  });
}

function atomicWriteJson(filePath, payload) {
  const tempPath = `${filePath}.${process.pid}.${Date.now()}.tmp`;
  fs.writeFileSync(tempPath, JSON.stringify(payload, null, 2));
  fs.renameSync(tempPath, filePath);
}

function analyticsArchiveRowCountFromCsv(csv) {
  if (typeof csv !== 'string' || !csv.trim()) return 0;
  const parsed = Papa.parse(csv, { header: true, skipEmptyLines: true, preview: 0 });
  return Array.isArray(parsed.data) ? parsed.data.length : 0;
}

function normalizeArchiveManifestEntry(entry = {}, stats = null) {
  const termCode = String(entry.termCode || entry.term || '').trim();
  if (!termCode) return null;
  return {
    termCode,
    displayTerm: String(entry.displayTerm || displayTermFromArchiveTerm(termCode)).trim(),
    rowCount: Number.isFinite(Number(entry.rowCount)) ? Number(entry.rowCount) : null,
    updatedAt: String(entry.updatedAt || entry.lastUpdated || stats?.mtime?.toISOString?.() || '').trim(),
    sizeBytes: Number.isFinite(Number(entry.sizeBytes)) ? Number(entry.sizeBytes) : (stats ? stats.size : null),
    hasArchive: entry.hasArchive !== false,
    schemaVersion: entry.schemaVersion || 'csv-v1'
  };
}

function readAnalyticsArchiveManifestFile() {
  if (!fs.existsSync(ANALYTICS_ARCHIVE_MANIFEST_PATH)) return null;
  try {
    const payload = JSON.parse(fs.readFileSync(ANALYTICS_ARCHIVE_MANIFEST_PATH, 'utf8'));
    if (!payload || payload.schemaVersion !== ANALYTICS_ARCHIVE_MANIFEST_SCHEMA_VERSION || !Array.isArray(payload.terms)) return null;
    return {
      schemaVersion: ANALYTICS_ARCHIVE_MANIFEST_SCHEMA_VERSION,
      generatedAt: payload.generatedAt || new Date().toISOString(),
      terms: sortArchiveTermsNewestFirst(payload.terms.map(term => normalizeArchiveManifestEntry(term)).filter(Boolean))
    };
  } catch (err) {
    console.warn('Analytics archive manifest invalid; rebuilding:', err.message || err);
    return null;
  }
}

function writeAnalyticsArchiveManifest(terms = []) {
  const existingFiles = new Set(
    fs.readdirSync(ANALYTICS_ARCHIVE_DIR)
      .filter(file => file.toLowerCase().endsWith('.csv'))
      .map(file => path.basename(file, '.csv'))
  );
  const payload = {
    schemaVersion: ANALYTICS_ARCHIVE_MANIFEST_SCHEMA_VERSION,
    generatedAt: new Date().toISOString(),
    terms: sortArchiveTermsNewestFirst(terms
      .map(term => normalizeArchiveManifestEntry(term))
      .filter(term => term && existingFiles.has(term.termCode))
      .map(term => ({ ...term, hasArchive: true })))
  };
  atomicWriteJson(ANALYTICS_ARCHIVE_MANIFEST_PATH, payload);
  return payload;
}

function rebuildAnalyticsArchiveManifest() {
  const terms = fs.readdirSync(ANALYTICS_ARCHIVE_DIR)
    .filter(file => file.toLowerCase().endsWith('.csv'))
    .map(file => {
      const termCode = path.basename(file, '.csv');
      const filePath = path.join(ANALYTICS_ARCHIVE_DIR, file);
      const stats = fs.statSync(filePath);
      let rowCount = null;
      try {
        rowCount = analyticsArchiveRowCountFromCsv(fs.readFileSync(filePath, 'utf8'));
      } catch (err) {
        console.warn(`Analytics archive manifest row count skipped for ${termCode}:`, err.message || err);
      }
      return normalizeArchiveManifestEntry({ termCode, rowCount, updatedAt: stats.mtime.toISOString(), sizeBytes: stats.size }, stats);
    })
    .filter(Boolean);
  return writeAnalyticsArchiveManifest(terms);
}

function readAnalyticsArchiveManifest() {
  const manifest = readAnalyticsArchiveManifestFile();
  if (manifest) {
    const files = new Set(fs.readdirSync(ANALYTICS_ARCHIVE_DIR).filter(file => file.toLowerCase().endsWith('.csv')).map(file => path.basename(file, '.csv')));
    const complete = manifest.terms.every(term => files.has(term.termCode));
    const includesAll = Array.from(files).every(termCode => manifest.terms.some(term => term.termCode === termCode));
    if (complete && includesAll) return manifest;
  }
  return rebuildAnalyticsArchiveManifest();
}

function updateAnalyticsArchiveManifestEntry(term, csv) {
  const termCode = String(term || '').trim();
  const filePath = getAnalyticsArchivePath(termCode);
  if (!filePath || !fs.existsSync(filePath)) return readAnalyticsArchiveManifest();
  const stats = fs.statSync(filePath);
  const existing = readAnalyticsArchiveManifestFile() || { terms: [] };
  const next = existing.terms.filter(item => item.termCode !== termCode);
  next.push(normalizeArchiveManifestEntry({
    termCode,
    displayTerm: displayTermFromArchiveTerm(termCode),
    rowCount: analyticsArchiveRowCountFromCsv(csv),
    updatedAt: stats.mtime.toISOString(),
    sizeBytes: stats.size,
    hasArchive: true
  }, stats));
  return writeAnalyticsArchiveManifest(next);
}

function getFacultySchedulePath(term) {
  if (!/^[a-z0-9 _-]+$/i.test(term)) return null;
  const filePath = path.resolve(FACULTY_SCHEDULES_DIR, `${term}.json`);
  const dataRoot = path.resolve(FACULTY_SCHEDULES_DIR) + path.sep;
  return filePath.startsWith(dataRoot) ? filePath : null;
}

function getWorkExperiencePath(term) {
  if (!/^[a-z0-9 _-]+$/i.test(term)) return null;
  const filePath = path.resolve(WORK_EXPERIENCE_DIR, `${term}.json`);
  const dataRoot = path.resolve(WORK_EXPERIENCE_DIR) + path.sep;
  return filePath.startsWith(dataRoot) ? filePath : null;
}

function isValidLowEnrollmentTermCode(term) {
  return /^\d{6}$/.test(String(term || '').trim());
}

function getLowEnrollmentTrackingPath(term) {
  if (!isValidLowEnrollmentTermCode(term)) return null;
  const filePath = path.resolve(LOW_ENROLLMENT_TRACKING_DIR, `${term}.json`);
  const dataRoot = path.resolve(LOW_ENROLLMENT_TRACKING_DIR) + path.sep;
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

function requestClientKey(req, scope) {
  return `${scope}:${req.ip || req.socket?.remoteAddress || 'unknown'}`;
}

function checkAuthenticationLock(req) {
  const key = requestClientKey(req, 'auth');
  const now = Date.now();
  const state = authFailureState.get(key);
  if (!state) return key;
  if (state.lockedUntil > now) {
    const err = new Error('Too many failed sign-in attempts. Please wait before trying again.');
    err.status = 429;
    err.retryAfterSeconds = Math.max(1, Math.ceil((state.lockedUntil - now) / 1000));
    throw err;
  }
  if (now - state.firstFailureAt >= AUTH_FAILURE_WINDOW_MS) authFailureState.delete(key);
  return key;
}

function recordAuthenticationFailure(key) {
  const now = Date.now();
  const current = authFailureState.get(key);
  const state = !current || now - current.firstFailureAt >= AUTH_FAILURE_WINDOW_MS
    ? { failures: 0, firstFailureAt: now, lockedUntil: 0 }
    : current;
  state.failures += 1;
  if (state.failures >= AUTH_FAILURE_LIMIT) state.lockedUntil = now + AUTH_LOCKOUT_MS;
  authFailureState.set(key, state);
}

function clearAuthenticationFailures(key) {
  authFailureState.delete(key);
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

function requireEnrollmentRole(req, res, minimumRole = 'development') {
  if (!isEnrollmentSessionAuthorized(req)) {
    res.status(401).json({ error: 'Enrollment Management session is required.', code: 'UNAUTHORIZED' });
    return null;
  }
  const role = enrollmentSessionRole(req);
  const requiredLevel = ROLE_LEVEL[minimumRole] || ROLE_LEVEL.development;
  if (!role || (ROLE_LEVEL[role] || 0) < requiredLevel) {
    res.status(403).json({ error: 'Insufficient role for this action.', code: 'FORBIDDEN' });
    return null;
  }
  return role;
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
  const snapshotType = String(record.snapshotType || record['Snapshot Type'] || '').trim().toUpperCase();
  const parts = [
    String(record.term || record.Term || '').trim().toUpperCase(),
    String(record.crn || record.CRN || '').trim().toUpperCase(),
    snapshotType
  ];
  if (snapshotType === 'CUSTOM') parts.push(String(record.snapshotDate || record['Snapshot Date'] || '').trim());
  return parts.join('|');
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
    '/usr/bin/soffice',
    '/usr/bin/libreoffice',
    '/usr/local/bin/soffice',
    '/usr/local/bin/libreoffice',
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
    : [
        process.env.LIBREOFFICE_PATH,
        '/usr/bin/soffice',
        '/usr/bin/libreoffice',
        '/usr/local/bin/soffice',
        '/usr/local/bin/libreoffice',
        'soffice',
        'libreoffice'
      ].filter(Boolean));
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
    emailDraftSupported: MICROSOFT_GRAPH_DRAFT_SUPPORTED,
    localEmailDraftSupported: true,
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
  if (!requireEnrollmentRole(req, res, 'general')) return;
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
  if (!requireEnrollmentRole(req, res, 'general')) return;
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

function checkConversionRateLimit(req) {
  const key = requestClientKey(req, 'conversion');
  const now = Date.now();
  const recent = (conversionRateLimit.get(key) || []).filter(timestamp => now - timestamp < CONVERSION_RATE_LIMIT_WINDOW_MS);
  if (recent.length >= CONVERSION_RATE_LIMIT_MAX) {
    const err = new Error('Too many document conversion requests. Please wait and try again.');
    err.status = 429;
    err.retryAfterSeconds = Math.max(1, Math.ceil((CONVERSION_RATE_LIMIT_WINDOW_MS - (now - recent[0])) / 1000));
    throw err;
  }
  recent.push(now);
  conversionRateLimit.set(key, recent);
}

async function handleScheduleChangeDocxToPdf(req, res) {
  if (!DOCX_PDF_CAPABILITY.available) {
    return res.status(503).json({
      error: PDF_CONVERSION_UNAVAILABLE_MESSAGE,
      reason: DOCX_PDF_CAPABILITY.reason,
      capabilities: exportCapabilities()
    });
  }
  try {
    checkConversionRateLimit(req);
  } catch (err) {
    res.setHeader('Retry-After', String(err.retryAfterSeconds || Math.ceil(CONVERSION_RATE_LIMIT_WINDOW_MS / 1000)));
    return res.status(429).json({ error: err.message, code: 'CONVERSION_RATE_LIMITED' });
  }
  if (activeConversions >= CONVERSION_MAX_CONCURRENT) {
    res.setHeader('Retry-After', '10');
    return res.status(503).json({ error: 'Document conversion is busy. Please try again shortly.', code: 'CONVERSION_BUSY' });
  }
  activeConversions += 1;
  let conversionSlotHeld = true;
  const releaseConversionSlot = () => {
    if (!conversionSlotHeld) return;
    conversionSlotHeld = false;
    activeConversions = Math.max(0, activeConversions - 1);
  };

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
      releaseConversionSlot();
      const cleanup = cleanupConversionDir(requestDir);
      console.log('[DOCX-PDF] Response cleanup:', JSON.stringify({ requestId, cleanup }));
      if (err) console.error('PDF send error:', err);
    });
  } catch (err) {
    releaseConversionSlot();
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

app.get('/api/admin/diagnostics', (req, res) => {
  if (!requireEnrollmentRole(req, res, 'admin')) return;
  return res.json(diagnosticsPayload());
});

app.post('/api/auth/enrollment-management', (req, res) => {
  let authKey;
  try {
    authKey = checkAuthenticationLock(req);
  } catch (err) {
    res.setHeader('Retry-After', String(err.retryAfterSeconds || Math.ceil(AUTH_LOCKOUT_MS / 1000)));
    return res.status(429).json({ error: err.message, code: 'AUTH_LOCKED' });
  }
  const { password } = req.body || {};
  const role = authenticateRolePassword(password, 'em');
  if (!role) {
    recordAuthenticationFailure(authKey);
    return res.status(403).json({ error: 'Unauthorized' });
  }
  clearAuthenticationFailures(authKey);
  return res.json(issueEnrollmentSession(role));
});

app.post('/api/auth/role', (req, res) => {
  let authKey;
  try {
    authKey = checkAuthenticationLock(req);
  } catch (err) {
    res.setHeader('Retry-After', String(err.retryAfterSeconds || Math.ceil(AUTH_LOCKOUT_MS / 1000)));
    return res.status(429).json({ error: err.message, code: 'AUTH_LOCKED' });
  }
  const { password, requestedRole = 'general' } = req.body || {};
  const minimumRole = ROLE_LEVEL[requestedRole] ? requestedRole : 'general';
  const role = authenticateRolePassword(password, minimumRole);
  if (!role) {
    recordAuthenticationFailure(authKey);
    return res.status(403).json({ error: 'Unauthorized', requiredRole: minimumRole });
  }
  clearAuthenticationFailures(authKey);
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
    const manifest = readAnalyticsArchiveManifest();
    const terms = manifest.terms.map(term => ({
      term: term.termCode,
      lastUpdated: term.updatedAt
    }));
    return res.json({ data: terms });
  } catch (err) {
    console.error('Analytics archive list error:', err);
    return res.status(500).json({ error: 'Analytics archive list failed' });
  }
});

app.get('/api/analytics-archive/manifest', (req, res) => {
  try {
    return res.json({ data: readAnalyticsArchiveManifest() });
  } catch (err) {
    console.error('Analytics archive manifest error:', err);
    return res.status(500).json({ error: 'Analytics archive manifest failed' });
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

function dataField(row, names) {
  if (!row || typeof row !== 'object') return '';
  const normalizeKey = value => String(value || '').replace(/[^a-z0-9]/gi, '').toLowerCase();
  const aliases = new Set(names.map(normalizeKey));
  for (const [key, value] of Object.entries(row)) {
    if (aliases.has(normalizeKey(key)) && value != null && String(value).trim()) return String(value).trim();
  }
  return '';
}

function numericField(row, names) {
  const raw = dataField(row, names);
  const value = Number(String(raw).replace(/,/g, ''));
  return Number.isFinite(value) ? value : 0;
}

function workExperienceMetadata(term, rows, base = {}) {
  const safeRows = Array.isArray(rows) ? rows : [];
  const crns = new Set();
  let enrollmentTotal = 0;
  let ftesTotal = 0;
  safeRows.forEach(row => {
    const crn = dataField(row, ['CRN', 'crn']);
    if (crn) crns.add(crn.toUpperCase());
    enrollmentTotal += numericField(row, ['ACTUAL_ENROLL', 'Actual Enroll', 'Current Enrollment', 'Enrollment', 'CENSUS_ENROLL']);
    ftesTotal += numericField(row, ['FTES', 'ftes']);
  });
  return {
    term,
    uploadedAt: base.uploadedAt || new Date().toISOString(),
    uploadedByRole: base.uploadedByRole || '',
    sourceFileName: base.sourceFileName || '',
    rawRowCount: safeRows.length,
    distinctCrnCount: crns.size,
    enrollmentTotal,
    ftesTotal
  };
}

function validateWorkExperienceRows(rows) {
  if (!Array.isArray(rows) || !rows.length) {
    return { valid: false, error: 'Work Experience rows are required.' };
  }
  const hasCrn = rows.some(row => dataField(row, ['CRN', 'crn']));
  if (!hasCrn) {
    return { valid: false, error: 'This does not appear to be a Work Experience file. Work Experience rows must include CRN values.' };
  }
  return { valid: true };
}

function readWorkExperienceArchive(term) {
  const filePath = getWorkExperiencePath(term);
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

app.get('/api/work-experience', (req, res) => {
  try {
    const data = fs.readdirSync(WORK_EXPERIENCE_DIR)
      .filter(file => file.toLowerCase().endsWith('.json'))
      .map(file => {
        const term = path.basename(file, '.json');
        const payload = readWorkExperienceArchive(term);
        const stats = fs.statSync(path.join(WORK_EXPERIENCE_DIR, file));
        return payload?.metadata || { term, uploadedAt: stats.mtime.toISOString() };
      })
      .sort((a, b) => String(a.term || '').localeCompare(String(b.term || ''), undefined, { numeric: true }));
    return res.json({ data });
  } catch (err) {
    console.error('Work Experience archive list error:', err);
    return res.status(500).json({ error: 'Work Experience archive list failed' });
  }
});

app.get('/api/work-experience/:term', (req, res) => {
  const term = req.params.term;
  const filePath = getWorkExperiencePath(term);
  if (!filePath) return res.status(400).json({ error: 'Invalid term' });
  if (!fs.existsSync(filePath)) return res.json({ term, metadata: null, data: [] });
  try {
    const payload = JSON.parse(fs.readFileSync(filePath, 'utf8'));
    return res.json({ term, metadata: payload.metadata || null, data: Array.isArray(payload.rows) ? payload.rows : [] });
  } catch (err) {
    console.error('Work Experience archive read error:', err);
    return res.status(500).json({ error: 'Work Experience archive read failed' });
  }
});

app.post('/api/work-experience/:term', (req, res) => {
  const term = req.params.term;
  const { rows, password, sourceFileName = '' } = req.body || {};
  if (!isEnrollmentSessionAuthorized(req) && !isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const filePath = getWorkExperiencePath(term);
  if (!filePath) return res.status(400).json({ error: 'Invalid term' });
  const validation = validateWorkExperienceRows(rows);
  if (!validation.valid) return res.status(400).json({ error: validation.error });
  try {
    const metadata = workExperienceMetadata(term, rows, {
      uploadedByRole: enrollmentSessionRole(req) || (isAuthorized(password) ? 'general' : ''),
      sourceFileName
    });
    fs.writeFileSync(filePath, JSON.stringify({ metadata, rows }, null, 2));
    return res.json({ success: true, term, metadata, data: rows });
  } catch (err) {
    console.error('Work Experience archive write error:', err);
    return res.status(500).json({ error: 'Work Experience archive write failed' });
  }
});

app.delete('/api/work-experience/:term', (req, res) => {
  const term = req.params.term;
  const { password } = req.body || {};
  if (!isEnrollmentSessionAuthorized(req) && !isAuthorized(password)) {
    return res.status(403).json({ error: 'Unauthorized' });
  }
  const filePath = getWorkExperiencePath(term);
  if (!filePath) return res.status(400).json({ error: 'Invalid term' });
  try {
    if (fs.existsSync(filePath)) fs.unlinkSync(filePath);
    return res.json({ success: true, term });
  } catch (err) {
    console.error('Work Experience archive delete error:', err);
    return res.status(500).json({ error: 'Work Experience archive delete failed' });
  }
});

const LOW_ENROLLMENT_EDITABLE_FIELDS = new Set(['justification', 'vpComments']);
const lowEnrollmentTermLocks = new Map();

function lowEnrollmentError(message, code = 'LOW_ENROLLMENT_ERROR', statusCode = 400, extra = {}) {
  const err = new Error(message);
  err.code = code;
  err.statusCode = statusCode;
  Object.assign(err, extra);
  return err;
}

function sendLowEnrollmentError(res, err, fallback = 'Low Enrollment Tracking request failed') {
  const status = err?.statusCode || 500;
  const code = err?.code || (status >= 500 ? 'STORAGE_FAILURE' : 'LOW_ENROLLMENT_ERROR');
  const payload = {
    error: err?.message || fallback,
    code
  };
  if (err?.requiresConfirmation) payload.requiresConfirmation = true;
  if (err?.summary) payload.summary = err.summary;
  if (err?.snapshotDate) payload.snapshotDate = err.snapshotDate;
  if (err?.warnings) payload.warnings = err.warnings;
  return res.status(status).json(payload);
}

function lowEnrollmentAllowedReasons(workspace) {
  return Array.from(new Set((Array.isArray(workspace?.reasons) ? workspace.reasons : [])
    .map(value => String(value || '').trim())
    .filter(Boolean)));
}

function normalizeLowEnrollmentDate(value) {
  const text = String(value || '').trim();
  if (!text) return '';
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const parsed = new Date(text);
  return Number.isNaN(parsed.getTime()) ? text : parsed.toISOString().slice(0, 10);
}

function normalizeLowEnrollmentNumber(value, preserveNull = true) {
  if (value === null || value === undefined || String(value).trim?.() === '') return preserveNull ? null : 0;
  const number = Number(String(value).replace(/[$,%]/g, '').trim());
  return Number.isFinite(number) ? number : null;
}

function normalizeLowEnrollmentCrns(crns) {
  return Array.from(new Set((Array.isArray(crns) ? crns : [])
    .map(value => String(value || '').trim())
    .filter(Boolean)));
}

function lowEnrollmentStatusForRow(row) {
  const threshold = normalizeLowEnrollmentNumber(row?.threshold);
  const latest = normalizeLowEnrollmentNumber(row?.latestEnrollment);
  const current = normalizeLowEnrollmentNumber(row?.currentEnrollment);
  const initial = normalizeLowEnrollmentNumber(row?.initialEnrollment);
  const basis = latest !== null ? latest : current !== null ? current : initial;
  if (basis === null || threshold === null) return 'Manual Review';
  return basis >= threshold ? 'Threshold Met' : 'Below Threshold';
}

function lowEnrollmentValidationWarnings(rows = []) {
  const crnRows = new Map();
  rows.forEach(row => {
    (row.crns || []).forEach(crn => {
      const key = String(crn);
      if (!crnRows.has(key)) crnRows.set(key, []);
      crnRows.get(key).push(String(row.id));
    });
  });
  return Array.from(crnRows.entries())
    .filter(([, rowIds]) => rowIds.length > 1)
    .map(([crn, rowIds]) => ({ code: 'DUPLICATE_CRN', crn, rowIds }));
}

function validateLowEnrollmentWorkspace(workspace, expectedTermCode = '') {
  if (!workspace || typeof workspace !== 'object') {
    throw lowEnrollmentError('Workspace payload is required.', 'INVALID_WORKSPACE', 400);
  }
  const termCode = String(workspace.termCode || '').trim();
  if (!isValidLowEnrollmentTermCode(termCode)) {
    throw lowEnrollmentError('A valid six-digit term code is required.', 'INVALID_TERM', 400);
  }
  if (expectedTermCode && termCode !== expectedTermCode) {
    throw lowEnrollmentError('Workspace term code must match the request term.', 'INVALID_WORKSPACE', 400);
  }
  if (!Array.isArray(workspace.reasons) || !lowEnrollmentAllowedReasons(workspace).length) {
    throw lowEnrollmentError('At least one saved Justification reason is required.', 'INVALID_WORKSPACE', 400);
  }
  if (!Array.isArray(workspace.rows) || !workspace.rows.length) {
    throw lowEnrollmentError('Tracker rows are required.', 'INVALID_WORKSPACE', 400);
  }
  if (!Array.isArray(workspace.snapshots)) {
    throw lowEnrollmentError('Snapshots must be an array.', 'INVALID_WORKSPACE', 400);
  }
  if (!Array.isArray(workspace.uploadHistory)) {
    throw lowEnrollmentError('Upload history must be an array.', 'INVALID_WORKSPACE', 400);
  }
  const ids = new Set();
  workspace.rows.forEach((row, index) => {
    if (!row || typeof row !== 'object' || Array.isArray(row)) {
      throw lowEnrollmentError(`Tracker row ${index + 1} must be an object.`, 'INVALID_WORKSPACE', 400);
    }
    const rowId = String(row.id || '').trim();
    if (!rowId) throw lowEnrollmentError(`Tracker row ${index + 1} is missing a stable id.`, 'INVALID_WORKSPACE', 400);
    if (ids.has(rowId)) throw lowEnrollmentError(`Duplicate tracker row id: ${rowId}.`, 'INVALID_WORKSPACE', 400);
    ids.add(rowId);
    const crns = normalizeLowEnrollmentCrns(row.crns);
    if (!crns.length) throw lowEnrollmentError(`Tracker row ${rowId} is missing CRNs.`, 'INVALID_WORKSPACE', 400);
    if (crns.some(crn => !/^\d+$/.test(crn))) throw lowEnrollmentError(`Tracker row ${rowId} contains an invalid CRN.`, 'INVALID_WORKSPACE', 400);
    const threshold = normalizeLowEnrollmentNumber(row.threshold);
    if (threshold === null || threshold < 0) throw lowEnrollmentError(`Tracker row ${rowId} has an invalid threshold.`, 'INVALID_WORKSPACE', 400);
  });
  return lowEnrollmentValidationWarnings(workspace.rows);
}

function normalizeLowEnrollmentWorkspace(workspace, base = {}) {
  const now = new Date().toISOString();
  const prior = base.prior || {};
  if (!Array.isArray(workspace?.snapshots)) {
    throw lowEnrollmentError('Snapshots must be an array.', 'INVALID_WORKSPACE', 400);
  }
  if (!Array.isArray(workspace?.uploadHistory)) {
    throw lowEnrollmentError('Upload history must be an array.', 'INVALID_WORKSPACE', 400);
  }
  const normalized = {
    ...workspace,
    termCode: String(workspace.termCode || '').trim(),
    displayTerm: String(workspace.displayTerm || workspace.termName || workspace.termCode || '').trim(),
    status: workspace.status || 'active',
    sourceFilename: String(workspace.sourceFilename || '').trim(),
    initialSnapshotDate: normalizeLowEnrollmentDate(workspace.initialSnapshotDate),
    reasons: lowEnrollmentAllowedReasons(workspace),
    rows: (Array.isArray(workspace.rows) ? workspace.rows : []).map(row => ({
      ...row,
      id: String(row.id || '').trim(),
      crns: normalizeLowEnrollmentCrns(row.crns),
      threshold: normalizeLowEnrollmentNumber(row.threshold),
      initialEnrollment: normalizeLowEnrollmentNumber(row.initialEnrollment),
      currentEnrollment: normalizeLowEnrollmentNumber(row.currentEnrollment),
      latestEnrollment: normalizeLowEnrollmentNumber(row.latestEnrollment),
      highestEnrollment: normalizeLowEnrollmentNumber(row.highestEnrollment),
      justification: String(row.justification || '').trim(),
      vpComments: String(row.vpComments || '')
    })),
    snapshots: Array.isArray(workspace.snapshots) ? workspace.snapshots.map(snapshot => ({
      ...snapshot,
      snapshotDate: normalizeLowEnrollmentDate(snapshot.snapshotDate)
    })) : [],
    uploadHistory: Array.isArray(workspace.uploadHistory) ? workspace.uploadHistory.map(item => ({
      ...item,
      snapshotDate: normalizeLowEnrollmentDate(item.snapshotDate)
    })) : [],
    createdAt: prior.createdAt || workspace.createdAt || now,
    updatedAt: now,
    importedAt: workspace.importedAt || now
  };
  normalized.validationWarnings = validateLowEnrollmentWorkspace(normalized, normalized.termCode);
  normalized.rows = normalized.rows.map(row => ({ ...row, status: row.status || lowEnrollmentStatusForRow(row) }));
  return normalized;
}

function summarizeLowEnrollmentWorkspace(term, workspace, stats = null) {
  const rows = Array.isArray(workspace?.rows) ? workspace.rows : [];
  const snapshots = Array.isArray(workspace?.snapshots) ? workspace.snapshots : [];
  const history = Array.isArray(workspace?.uploadHistory) ? workspace.uploadHistory : [];
  const crns = new Set();
  rows.forEach(row => (Array.isArray(row.crns) ? row.crns : []).forEach(crn => crn && crns.add(String(crn))));
  return {
    termCode: workspace?.termCode || term,
    displayTerm: workspace?.displayTerm || workspace?.termName || term,
    sourceFilename: workspace?.sourceFilename || '',
    initialSnapshotDate: workspace?.initialSnapshotDate || '',
    createdAt: workspace?.createdAt || stats?.birthtime?.toISOString?.() || '',
    updatedAt: workspace?.updatedAt || stats?.mtime?.toISOString?.() || '',
    rowCount: rows.length,
    crnCount: crns.size,
    snapshotCount: snapshots.length,
    uploadHistoryCount: history.length,
    validationWarnings: workspace?.validationWarnings || []
  };
}

function readLowEnrollmentWorkspace(termCode) {
  const filePath = getLowEnrollmentTrackingPath(termCode);
  if (!filePath || !fs.existsSync(filePath)) return null;
  return JSON.parse(fs.readFileSync(filePath, 'utf8'));
}

function writeLowEnrollmentWorkspaceAtomic(workspace) {
  validateLowEnrollmentWorkspace(workspace, workspace.termCode);
  const filePath = getLowEnrollmentTrackingPath(workspace.termCode);
  if (!filePath) throw lowEnrollmentError('Invalid term code.', 'INVALID_TERM', 400);
  fs.mkdirSync(LOW_ENROLLMENT_TRACKING_DIR, { recursive: true });
  const tempPath = path.join(LOW_ENROLLMENT_TRACKING_DIR, `.${workspace.termCode}.${process.pid}.${Date.now()}.tmp`);
  const json = JSON.stringify(workspace, null, 2);
  let fd = null;
  try {
    fd = fs.openSync(tempPath, 'w');
    fs.writeFileSync(fd, json, 'utf8');
    fs.fsyncSync(fd);
    fs.closeSync(fd);
    fd = null;
    fs.renameSync(tempPath, filePath);
    return workspace;
  } catch (err) {
    if (fd !== null) {
      try { fs.closeSync(fd); } catch (_closeErr) {}
    }
    try { if (fs.existsSync(tempPath)) fs.unlinkSync(tempPath); } catch (_unlinkErr) {}
    throw lowEnrollmentError('Low Enrollment Tracking storage write failed.', 'STORAGE_FAILURE', 500);
  }
}

async function withLowEnrollmentTermLock(termCode, operation) {
  const prior = lowEnrollmentTermLocks.get(termCode) || Promise.resolve();
  let release = () => {};
  const gate = new Promise(resolve => { release = resolve; });
  const queued = prior.catch(() => undefined).then(() => gate);
  lowEnrollmentTermLocks.set(termCode, queued);
  await prior.catch(() => undefined);
  try {
    return await operation();
  } finally {
    release();
    if (lowEnrollmentTermLocks.get(termCode) === queued) lowEnrollmentTermLocks.delete(termCode);
  }
}

function validateLowEnrollmentSnapshotRequest(workspace, snapshot, rows, replaceExisting = false) {
  if (!workspace) throw lowEnrollmentError('Low Enrollment Tracking workspace not found.', 'WORKSPACE_NOT_FOUND', 404);
  const snapshotDate = normalizeLowEnrollmentDate(snapshot?.snapshotDate);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(snapshotDate)) {
    throw lowEnrollmentError('Snapshot date must be ISO YYYY-MM-DD.', 'INVALID_SNAPSHOT', 400);
  }
  if (!Array.isArray(rows) || !rows.length) {
    throw lowEnrollmentError('Snapshot rows are required.', 'INVALID_SNAPSHOT', 400);
  }
  const savedIds = (workspace.rows || []).map(row => String(row.id));
  const savedIdSet = new Set(savedIds);
  const incomingIds = rows.map(row => String(row?.id || '').trim());
  if (incomingIds.length !== savedIds.length || incomingIds.some(id => !savedIdSet.has(id)) || new Set(incomingIds).size !== savedIdSet.size) {
    throw lowEnrollmentError('Snapshot rows must exactly match the saved tracker rows.', 'INVALID_SNAPSHOT', 400);
  }
  const invalidSnapshotRowIds = Object.keys(snapshot?.values || {}).filter(rowId => !savedIdSet.has(String(rowId)));
  if (invalidSnapshotRowIds.length) {
    throw lowEnrollmentError('Snapshot values reference unknown tracker rows.', 'INVALID_SNAPSHOT', 400);
  }
  const existingSnapshot = (workspace.snapshots || []).find(item => String(item.snapshotDate) === snapshotDate);
  if (existingSnapshot && !replaceExisting) {
    throw lowEnrollmentError('A snapshot already exists for this date.', 'SNAPSHOT_EXISTS', 409, { requiresConfirmation: true, snapshotDate });
  }
  return snapshotDate;
}

function mergeLowEnrollmentSnapshotRows(savedRows = [], incomingRows = []) {
  const incomingById = new Map(incomingRows.map(row => [String(row.id), row]));
  return savedRows.map(savedRow => {
    const incoming = incomingById.get(String(savedRow.id)) || {};
    const merged = {
      ...savedRow,
      ...incoming,
      id: savedRow.id,
      crns: savedRow.crns,
      justification: savedRow.justification || '',
      vpComments: savedRow.vpComments || '',
      createdAt: savedRow.createdAt,
      updatedAt: savedRow.updatedAt
    };
    merged.threshold = normalizeLowEnrollmentNumber(merged.threshold);
    merged.initialEnrollment = normalizeLowEnrollmentNumber(merged.initialEnrollment);
    merged.currentEnrollment = normalizeLowEnrollmentNumber(merged.currentEnrollment);
    merged.latestEnrollment = normalizeLowEnrollmentNumber(merged.latestEnrollment);
    merged.highestEnrollment = normalizeLowEnrollmentNumber(merged.highestEnrollment);
    if (merged.latestEnrollment !== null) merged.currentEnrollment = merged.latestEnrollment;
    merged.status = lowEnrollmentStatusForRow(merged);
    return merged;
  });
}

app.get('/api/low-enrollment-tracking', (req, res) => {
  try {
    fs.mkdirSync(LOW_ENROLLMENT_TRACKING_DIR, { recursive: true });
    const data = fs.readdirSync(LOW_ENROLLMENT_TRACKING_DIR)
      .filter(file => /^\d{6}\.json$/i.test(file))
      .map(file => {
        const termCode = path.basename(file, '.json');
        const filePath = path.join(LOW_ENROLLMENT_TRACKING_DIR, file);
        const stats = fs.statSync(filePath);
        const workspace = readLowEnrollmentWorkspace(termCode);
        return summarizeLowEnrollmentWorkspace(termCode, workspace || {}, stats);
      })
      .sort((a, b) => String(b.updatedAt || '').localeCompare(String(a.updatedAt || '')));
    return res.json({ data });
  } catch (err) {
    console.error('Low Enrollment Tracking list error:', err);
    return sendLowEnrollmentError(res, lowEnrollmentError('Low Enrollment Tracking list failed.', 'STORAGE_FAILURE', 500));
  }
});

app.get('/api/low-enrollment-tracking/:termCode', (req, res) => {
  const termCode = String(req.params.termCode || '').trim();
  if (!isValidLowEnrollmentTermCode(termCode)) return sendLowEnrollmentError(res, lowEnrollmentError('Invalid term code.', 'INVALID_TERM', 400));
  try {
    const workspace = readLowEnrollmentWorkspace(termCode);
    if (!workspace) return sendLowEnrollmentError(res, lowEnrollmentError('Low Enrollment Tracking workspace not found.', 'WORKSPACE_NOT_FOUND', 404));
    return res.json({ termCode, data: workspace });
  } catch (err) {
    console.error('Low Enrollment Tracking read error:', err);
    return sendLowEnrollmentError(res, lowEnrollmentError('Low Enrollment Tracking read failed.', 'STORAGE_FAILURE', 500));
  }
});

app.post('/api/low-enrollment-tracking/:termCode', async (req, res) => {
  const termCode = String(req.params.termCode || '').trim();
  if (!isValidLowEnrollmentTermCode(termCode)) return sendLowEnrollmentError(res, lowEnrollmentError('Invalid term code.', 'INVALID_TERM', 400));
  if (!requireEnrollmentRole(req, res, 'development')) return;
  try {
    const result = await withLowEnrollmentTermLock(termCode, async () => {
      const { workspace, replaceExisting = false } = req.body || {};
      const prior = readLowEnrollmentWorkspace(termCode);
      if (prior && replaceExisting !== true) {
        throw lowEnrollmentError('A Low Enrollment Tracking workspace already exists for this term.', 'WORKSPACE_EXISTS', 409, {
          requiresConfirmation: true,
          summary: summarizeLowEnrollmentWorkspace(termCode, prior)
        });
      }
      const normalized = normalizeLowEnrollmentWorkspace(workspace || {}, { prior });
      const warnings = validateLowEnrollmentWorkspace(normalized, termCode);
      normalized.validationWarnings = warnings;
      writeLowEnrollmentWorkspaceAtomic(normalized);
      return normalized;
    });
    return res.json({ success: true, data: result, summary: summarizeLowEnrollmentWorkspace(termCode, result), warnings: result.validationWarnings || [] });
  } catch (err) {
    console.error('Low Enrollment Tracking import error:', err?.code || err?.message || err);
    return sendLowEnrollmentError(res, err, 'Low Enrollment Tracking import failed');
  }
});

app.post('/api/low-enrollment-tracking/:termCode/snapshots', async (req, res) => {
  const termCode = String(req.params.termCode || '').trim();
  if (!isValidLowEnrollmentTermCode(termCode)) return sendLowEnrollmentError(res, lowEnrollmentError('Invalid term code.', 'INVALID_TERM', 400));
  if (!requireEnrollmentRole(req, res, 'development')) return;
  try {
    const result = await withLowEnrollmentTermLock(termCode, async () => {
      const { snapshot, uploadHistory, rows, replaceExisting = false } = req.body || {};
      const workspace = readLowEnrollmentWorkspace(termCode);
      const snapshotDate = validateLowEnrollmentSnapshotRequest(workspace, snapshot, rows, replaceExisting);
      const normalizedSnapshot = { ...snapshot, snapshotDate };
      workspace.snapshots = [...(workspace.snapshots || []).filter(item => String(item.snapshotDate) !== snapshotDate), normalizedSnapshot]
        .sort((a, b) => String(a.snapshotDate).localeCompare(String(b.snapshotDate)));
      workspace.rows = mergeLowEnrollmentSnapshotRows(workspace.rows || [], rows);
      if (uploadHistory && typeof uploadHistory === 'object') {
        const normalizedHistory = { ...uploadHistory, type: uploadHistory.type || 'snapshot', snapshotDate };
        workspace.uploadHistory = [...(workspace.uploadHistory || []).filter(item => !(item.type === normalizedHistory.type && item.snapshotDate === snapshotDate)), normalizedHistory];
      }
      workspace.updatedAt = new Date().toISOString();
      workspace.validationWarnings = lowEnrollmentValidationWarnings(workspace.rows);
      writeLowEnrollmentWorkspaceAtomic(workspace);
      return workspace;
    });
    return res.json({ success: true, data: result, summary: summarizeLowEnrollmentWorkspace(termCode, result), warnings: result.validationWarnings || [] });
  } catch (err) {
    console.error('Low Enrollment Tracking snapshot error:', err?.code || err?.message || err);
    return sendLowEnrollmentError(res, err, 'Low Enrollment Tracking snapshot save failed');
  }
});

app.patch('/api/low-enrollment-tracking/:termCode/rows/:rowId', async (req, res) => {
  const termCode = String(req.params.termCode || '').trim();
  const rowId = String(req.params.rowId || '').trim();
  if (!isValidLowEnrollmentTermCode(termCode)) return sendLowEnrollmentError(res, lowEnrollmentError('Invalid term code.', 'INVALID_TERM', 400));
  if (!requireEnrollmentRole(req, res, 'development')) return;
  try {
    const result = await withLowEnrollmentTermLock(termCode, async () => {
      const workspace = readLowEnrollmentWorkspace(termCode);
      if (!workspace) throw lowEnrollmentError('Low Enrollment Tracking workspace not found.', 'WORKSPACE_NOT_FOUND', 404);
      const attemptedFields = Object.keys(req.body || {});
      const disallowedFields = attemptedFields.filter(field => !LOW_ENROLLMENT_EDITABLE_FIELDS.has(field));
      if (disallowedFields.length || !attemptedFields.length) {
        throw lowEnrollmentError(`Unsupported update field(s): ${disallowedFields.join(', ') || 'none'}.`, 'INVALID_ROW_UPDATE', 400);
      }
      const row = (workspace.rows || []).find(item => String(item.id) === rowId);
      if (!row) throw lowEnrollmentError('Tracker row not found.', 'ROW_NOT_FOUND', 404);
      if (Object.prototype.hasOwnProperty.call(req.body || {}, 'justification')) {
        const nextJustification = String(req.body.justification || '').trim();
        const allowedReasons = lowEnrollmentAllowedReasons(workspace);
        if (nextJustification.length > 500) throw lowEnrollmentError('Justification is too long.', 'INVALID_JUSTIFICATION', 400);
        if (nextJustification && !allowedReasons.includes(nextJustification)) {
          throw lowEnrollmentError('Invalid justification for this tracker workspace.', 'INVALID_JUSTIFICATION', 400);
        }
        row.justification = nextJustification;
      }
      if (Object.prototype.hasOwnProperty.call(req.body || {}, 'vpComments')) {
        if (typeof req.body.vpComments !== 'string') throw lowEnrollmentError('VP comments must be a string.', 'INVALID_ROW_UPDATE', 400);
        if (req.body.vpComments.length > 10000) throw lowEnrollmentError('VP comments are too long.', 'INVALID_ROW_UPDATE', 400);
        row.vpComments = req.body.vpComments;
      }
      row.updatedAt = new Date().toISOString();
      workspace.updatedAt = new Date().toISOString();
      writeLowEnrollmentWorkspaceAtomic(workspace);
      return { workspace, row };
    });
    return res.json({
      success: true,
      row: result.row,
      updatedAt: result.workspace.updatedAt,
      data: { row: result.row, workspaceUpdatedAt: result.workspace.updatedAt }
    });
  } catch (err) {
    console.error('Low Enrollment Tracking row update error:', err?.code || err?.message || err);
    return sendLowEnrollmentError(res, err, 'Low Enrollment Tracking row update failed');
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
    const manifest = updateAnalyticsArchiveManifestEntry(term, csv);
    const entry = manifest.terms.find(item => item.termCode === String(term || '').trim());
    return res.json({ success: true, term, lastUpdated: entry?.updatedAt || new Date().toISOString(), metadata: entry || null });
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

function handleRoomCatalogExport(_req, res) {
  try {
    return res.json(readRoomCatalog());
  } catch (err) {
    console.error('Room catalog export error:', err);
    return res.status(500).json({ error: 'Room catalog export failed' });
  }
}

app.get('/api/rooms/export', handleRoomCatalogExport);
app.post('/api/rooms/export', handleRoomCatalogExport);

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
  runCommand,
  validateWorkExperienceRows,
  workExperienceMetadata,
  displayTermFromArchiveTerm,
  archiveTermSortValue,
  sortArchiveTermsNewestFirst,
  readAnalyticsArchiveManifest,
  rebuildAnalyticsArchiveManifest,
  updateAnalyticsArchiveManifestEntry
};
