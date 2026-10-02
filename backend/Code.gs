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

// Words with score >= 3 (data.score is {word: 0..3}).
function countMastered_(data) {
  const sc = data.score || {};
  return Object.keys(sc).filter(function (k) { return sc[k] >= 3; }).length;
}

// Distinct practice days (data.days is an array of YYYY-MM-DD).
function countDays_(data) {
  return Array.isArray(data.days) ? data.days.length : 0;
}
