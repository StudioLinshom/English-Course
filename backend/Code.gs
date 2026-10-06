// Progress sync for the vocab app, one row per child in the "users" sheet.
//
//   GET  ?u=CODE                              -> {ok, name, data}
//   POST {action:"create", invite, name}      -> {ok, u, name}
//   POST {action:"create", sibling, name}     -> {ok, u, name}  (sibling = an existing child's code)
//   POST {action:"save", u, data}             -> {ok}
//   POST {action:"rename", u, name}           -> {ok}
//   POST {action:"checkInvite", invite}       -> {ok}
//   POST {action:"claim", key}                -> {ok, u, name}  (legacy single-user data -> Yoav's user)
//
// A user code is the only credential for that child's data, so codes are long and random.
// New users need the invite code from the "settings" sheet, which the owner can change.

// Key of the original single-user version. Accepted only until settings.legacy_until,
// so devices that still run the old page keep syncing while they update.
const SECRET = 'b936d1c344f0f173a3fc6df1eb1fca73';
const LEGACY_NAME = 'יואב';

const CODE_ABC = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789'; // no 0/O, 1/I
const CODE_LEN = 10;
const MAX_JSON = 400000;
const CHUNK = 45000;  // a Sheets cell holds at most 50,000 characters
const DATA_COL = 5;   // users: A code, B name, C created, D updated, E.. data chunks

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function newCode_(n) {
  let s = '';
  while (s.length < n) {
    const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid());
    for (let i = 0; i < bytes.length && s.length < n; i++) s += CODE_ABC[(bytes[i] + 256) % 32];
  }
  return s;
}

const normCode_ = c => String(c || '').toUpperCase().replace(/[^A-Z0-9]/g, '');

function settings_() {
  const sh = sheet_('settings');
  if (sh.getLastRow() === 0) {
    sh.getRange('B:B').setNumberFormat('@');
    const until = new Date(); until.setDate(until.getDate() + 21);
    sh.getRange(1, 1, 3, 3).setValues([
      ['invite', newCode_(6), 'קוד ההזמנה להורים של חברים. אפשר לשנות אותו בכל רגע.'],
      ['max_users', '100', 'מספר המשתמשים המרבי'],
      ['legacy_until', until.toISOString().slice(0, 10), 'עד התאריך הזה הגרסה הישנה של האפליקציה עוד מסתנכרנת'],
    ]);
  }
  const out = {};
  sh.getRange(1, 1, sh.getLastRow(), 2).getDisplayValues().forEach(r => { out[r[0]] = r[1]; });
  return out;
}

function legacyOpen_() {
  return new Date().toISOString().slice(0, 10) <= String(settings_().legacy_until || '');
}

function users_() {
  const sh = sheet_('users');
  if (sh.getLastRow() === 0) {
    sh.getRange('A:D').setNumberFormat('@');
    sh.appendRow(['code', 'name', 'created', 'updated', 'data']);
    sh.setFrozenRows(1);
  }
  return sh;
}

function findRow_(code) {
  if (!code) return 0;
  const cell = users_().getRange('A:A').createTextFinder(code).matchEntireCell(true).findNext();
  return cell ? cell.getRow() : 0;
}

function readUser_(row) {
  const sh = users_();
  const width = Math.max(1, sh.getLastColumn() - DATA_COL + 1);
  const head = sh.getRange(row, 1, 1, 2).getValues()[0];
  const raw = sh.getRange(row, DATA_COL, 1, width).getValues()[0].join('');
  return { name: head[1], data: raw ? JSON.parse(raw) : {} };
}

function writeData_(row, data) {
  const json = JSON.stringify(data || {});
  if (json.length > MAX_JSON) throw new Error('data too large');
  const sh = users_();
  const chunks = [];
  for (let i = 0; i < json.length; i += CHUNK) chunks.push(json.slice(i, i + CHUNK));
  const width = Math.max(chunks.length, sh.getLastColumn() - DATA_COL + 1);
  while (chunks.length < width) chunks.push('');
  const range = sh.getRange(row, DATA_COL, 1, width);
  range.setNumberFormat('@');
  range.setValues([chunks]);
  sh.getRange(row, 4).setValue(new Date().toISOString());
}

function createUser_(name, data) {
  const sh = users_();
  let code;
  do { code = newCode_(CODE_LEN); } while (findRow_(code));
  const now = new Date().toISOString();
  sh.appendRow([code, name, now, now]);
  writeData_(sh.getLastRow(), data || {});
  return code;
}

// One summary row per child per day, updated on every save.
function logDay_(code, name, data) {
  const sh = sheet_('daily');
  if (sh.getLastRow() === 0) {
    sh.appendRow(['date', 'name', 'words known', 'sentences known', 'practice days', 'last update', 'key']);
    sh.setFrozenRows(1);
    sh.getRange('A:A').setNumberFormat('@');
    sh.getRange('G:G').setNumberFormat('@');
  }
  const day = Utilities.formatDate(new Date(), 'Asia/Jerusalem', 'yyyy-MM-dd');
  const key = day + '|' + code;
  const row = [day, name, countAtLeast_(data.score, 3), countAtLeast_(data.gscore, 3),
               Array.isArray(data.days) ? data.days.length : 0, new Date(), key];
  const hit = sh.getRange('G:G').createTextFinder(key).matchEntireCell(true).findNext();
  if (hit) sh.getRange(hit.getRow(), 1, 1, row.length).setValues([row]);
  else sh.appendRow(row);
}

const countAtLeast_ = (map, n) => Object.keys(map || {}).filter(k => map[k] >= n).length;

function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.u) {
    const row = findRow_(normCode_(p.u));
    if (!row) return json_({ error: 'unknown user' });
    const user = readUser_(row);
    return json_({ ok: true, name: user.name, data: user.data });
  }
  if (p.key === SECRET && legacyOpen_()) return json_(legacyRead_());
  return json_({ error: 'unauthorized' });
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad json' }); }
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const action = body.action;
    if (action === 'create') {
      const s = settings_();
      // A parent adding a sibling proves membership with an existing child's code instead of the invite.
      const sibling = body.sibling && findRow_(normCode_(body.sibling));
      if (!sibling && normCode_(body.invite) !== normCode_(s.invite)) return json_({ error: 'bad invite' });
      const name = String(body.name || '').trim().slice(0, 30);
      if (!name) return json_({ error: 'missing name' });
      if (users_().getLastRow() - 1 >= Number(s.max_users || 100)) return json_({ error: 'too many users' });
      return json_({ ok: true, u: createUser_(name, body.data), name: name });
    }
    if (action === 'checkInvite') {
      // Lets a parent who forgot the parent PIN set a new one.
      return json_(normCode_(body.invite) === normCode_(settings_().invite) ? { ok: true } : { error: 'bad invite' });
    }
    if (action === 'save' || action === 'rename') {
      const code = normCode_(body.u);
      const row = findRow_(code);
      if (!row) return json_({ error: 'unknown user' });
      if (action === 'rename') {
        const name = String(body.name || '').trim().slice(0, 30);
        if (!name) return json_({ error: 'missing name' });
        users_().getRange(row, 2).setValue(name);
        return json_({ ok: true });
      }
      writeData_(row, body.data);
      logDay_(code, readUser_(row).name, body.data || {});
      return json_({ ok: true });
    }
    if (action === 'claim') {
      if (body.key !== SECRET || !legacyOpen_()) return json_({ error: 'unauthorized' });
      const code = legacyCode_();
      return json_({ ok: true, u: code, name: LEGACY_NAME });
    }
    if (!action && body.key === SECRET && legacyOpen_()) {
      // An old page saving its single-user state.
      const code = legacyCode_();
      writeData_(findRow_(code), body.data);
      logDay_(code, LEGACY_NAME, body.data || {});
      return json_({ ok: true });
    }
    return json_({ error: 'unauthorized' });
  } finally {
    lock.releaseLock();
  }
}

// The original single-user data lives in data!A1; it becomes Yoav's user the first time it is needed.
function legacyCode_() {
  const props = PropertiesService.getScriptProperties();
  let code = props.getProperty('LEGACY_CODE');
  if (code && findRow_(code)) return code;
  const raw = sheet_('data').getRange('A1').getValue();
  code = createUser_(LEGACY_NAME, raw ? JSON.parse(raw) : {});
  props.setProperty('LEGACY_CODE', code);
  return code;
}

function legacyRead_() {
  const props = PropertiesService.getScriptProperties();
  const code = props.getProperty('LEGACY_CODE');
  const row = code ? findRow_(code) : 0;
  if (row) return readUser_(row).data;
  const raw = sheet_('data').getRange('A1').getValue();
  return raw ? JSON.parse(raw) : {};
}
