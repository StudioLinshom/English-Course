const SECRET = 'b936d1c344f0f173a3fc6df1eb1fca73';

function sheet_(name) {
  const ss = SpreadsheetApp.getActive();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  if (!e || !e.parameter || e.parameter.key !== SECRET) return json_({ error: 'unauthorized' });
  const raw = sheet_('data').getRange('A1').getValue();
  return json_(raw ? JSON.parse(raw) : {});
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ error: 'bad json' }); }
  if (body.key !== SECRET) return json_({ error: 'unauthorized' });
  const data = body.data || {};
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    sheet_('data').getRange('A1').setValue(JSON.stringify(data));
    const log = sheet_('log');
    if (log.getLastRow() === 0) log.appendRow(['timestamp', 'words mastered', 'practice days']);
    log.appendRow([new Date(), countMastered_(data), countDays_(data)]);
  } finally {
    lock.releaseLock();
  }
  return json_({ ok: true });
}

// Words with score >= 3. Accepts {words:{w:{score}}}, {scores:{w:n}}, or arrays of {score}.
function countMastered_(data) {
  const src = data.words || data.scores || data.progress || {};
  const vals = Array.isArray(src) ? src : Object.keys(src).map(function (k) { return src[k]; });
  return vals.filter(function (v) {
    const s = typeof v === 'number' ? v : (v && v.score);
    return s >= 3;
  }).length;
}

// Distinct practice days. Accepts data.days / data.practiceDays as array, object, or number.
function countDays_(data) {
  const d = data.days || data.practiceDays || data.history;
  if (typeof d === 'number') return d;
  if (Array.isArray(d)) return d.length;
  if (d && typeof d === 'object') return Object.keys(d).length;
  return 0;
}
