/**
 * Vert Academy sign-up backend (Google Apps Script, bound to the sign-up Google Sheet).
 *
 * - GET  ?action=status&sessions=2026-11-14,2026-11-21  -> { ok, capacity, counts: { id: n } }
 * - POST { session, name, phone, email }                -> { ok, spot, count } or { ok:false, code, message }
 *
 * The server is the source of truth: it checks the registration window
 * (Monday 6:00 AM to Saturday 10:00 AM Pacific) and the 8-spot cap under a lock,
 * so two people can never take the last spot at the same time.
 *
 * Optional Zenzap alert: set Script Properties ZENZAP_API_KEY, ZENZAP_API_SECRET
 * and ZENZAP_TOPIC_ID (Project Settings > Script Properties). Leave them unset to skip.
 */

var VA = {
  TIMEZONE: 'America/Los_Angeles',
  CAPACITY: 8,
  SESSION_TIME: '10:00',   // Saturday
  OPEN_DAYS_BEFORE: 5,     // Monday before
  OPEN_TIME: '06:00',
  SHEET_NAME: 'Signups',
  HEADERS: ['Submitted At', 'Session Date', 'Session', 'Spot #', 'Full Name', 'Phone', 'Email']
};

/* ------------------------------------------------------------------ */
/* Web app entry points                                                */
/* ------------------------------------------------------------------ */

function doGet(e) {
  var p = (e && e.parameter) || {};
  if (p.action === 'status') {
    var ids = String(p.sessions || '').split(',').filter(isSessionId_).slice(0, 6);
    var all = countBySession_();
    var counts = {};
    ids.forEach(function (id) { counts[id] = all[id] || 0; });
    return json_({ ok: true, capacity: VA.CAPACITY, counts: counts, serverTime: Date.now() });
  }
  return json_({ ok: true, service: 'vert-academy' });
}

function doPost(e) {
  var body;
  try {
    body = JSON.parse(e.postData.contents);
  } catch (err) {
    return json_({ ok: false, code: 'INVALID', message: 'Bad request.' });
  }

  // Honeypot: bots fill the hidden "website" field. Pretend success, store nothing.
  if (body.website) return json_({ ok: true, spot: 0, count: 0, capacity: VA.CAPACITY });

  var name = clean_(body.name, 80).replace(/\s+/g, ' ');
  var digits = String(body.phone || '').replace(/\D/g, '');
  if (digits.length === 11 && digits.charAt(0) === '1') digits = digits.slice(1);
  var email = clean_(body.email, 120).toLowerCase();
  var id = String(body.session || '');

  if (name.length < 2) return json_({ ok: false, code: 'INVALID', message: 'Please enter your full name.' });
  if (digits.length !== 10) return json_({ ok: false, code: 'INVALID', message: 'Please enter a 10-digit phone number.' });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(email)) return json_({ ok: false, code: 'INVALID', message: 'Please enter a valid email address.' });

  var win = sessionWindow_(id);
  if (!win) return json_({ ok: false, code: 'INVALID', message: 'That session is not available.' });
  var now = new Date();
  if (now < win.opens) return json_({ ok: false, code: 'NOT_OPEN', message: 'Registration for this session is not open yet.' });
  if (now >= win.start) return json_({ ok: false, code: 'CLOSED', message: 'Registration for this session has closed.' });

  var phone = '(' + digits.slice(0, 3) + ') ' + digits.slice(3, 6) + '-' + digits.slice(6);
  var label = Utilities.formatDate(win.start, VA.TIMEZONE, 'EEE MMM d, yyyy') + ' · 10:00 AM';
  var spot;

  var lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) {
    return json_({ ok: false, code: 'BUSY', message: 'Lots of sign-ups right now. Please try again in a few seconds.' });
  }
  try {
    var sheet = sheet_();
    var rows = sessionRows_(sheet, id);
    if (rows.length >= VA.CAPACITY) {
      return json_({ ok: false, code: 'FULL', count: rows.length, message: 'Sorry, this session just filled up.' });
    }
    var dup = rows.some(function (r) {
      return String(r[6]).toLowerCase() === email || String(r[5]).replace(/\D/g, '') === digits;
    });
    if (dup) {
      return json_({ ok: false, code: 'DUPLICATE', count: rows.length, message: 'You are already registered for this session.' });
    }
    spot = rows.length + 1;
    sheet.appendRow([
      Utilities.formatDate(now, VA.TIMEZONE, 'yyyy-MM-dd HH:mm:ss'),
      id, label, spot, name, phone, email
    ]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }

  try {
    var text = 'New Vert Academy sign-up (' + spot + '/' + VA.CAPACITY + ')\n' +
      label + '\n' + name + '\n' + phone + '\n' + email +
      (spot >= VA.CAPACITY ? '\nSESSION IS NOW FULL. Registration closed.' : '');
    notifyZenzap_(text, 'va-' + id + '-' + spot);
  } catch (err) {
    console.error('Zenzap notify failed: ' + err);
  }

  return json_({ ok: true, spot: spot, count: spot, capacity: VA.CAPACITY });
}

/* ------------------------------------------------------------------ */
/* Helpers                                                             */
/* ------------------------------------------------------------------ */

function isSessionId_(id) { return /^\d{4}-\d{2}-\d{2}$/.test(id); }

/** Returns {start, opens} for a Saturday session id (yyyy-MM-dd), or null. */
function sessionWindow_(id) {
  if (!isSessionId_(id)) return null;
  var start = Utilities.parseDate(id + ' ' + VA.SESSION_TIME, VA.TIMEZONE, 'yyyy-MM-dd HH:mm');
  if (Utilities.formatDate(start, VA.TIMEZONE, 'yyyy-MM-dd') !== id) return null;
  if (Utilities.formatDate(start, VA.TIMEZONE, 'u') !== '6') return null; // 6 = Saturday
  var openDay = Utilities.formatDate(new Date(start.getTime() - VA.OPEN_DAYS_BEFORE * 86400000), VA.TIMEZONE, 'yyyy-MM-dd');
  var opens = Utilities.parseDate(openDay + ' ' + VA.OPEN_TIME, VA.TIMEZONE, 'yyyy-MM-dd HH:mm');
  return { start: start, opens: opens };
}

function sheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    var sid = PropertiesService.getScriptProperties().getProperty('SHEET_ID');
    ss = SpreadsheetApp.openById(sid);
  }
  var sh = ss.getSheetByName(VA.SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(VA.SHEET_NAME);
    formatSheet_(sh);
  }
  return sh;
}

function formatSheet_(sh) {
  sh.getRange(1, 1, 1, VA.HEADERS.length).setValues([VA.HEADERS])
    .setFontWeight('bold').setBackground('#060606').setFontColor('#ffffff');
  sh.setFrozenRows(1);
  // Keep session ids and phone numbers as plain text
  sh.getRange('B:B').setNumberFormat('@');
  sh.getRange('F:F').setNumberFormat('@');
}

function sessionKey_(v) {
  if (Object.prototype.toString.call(v) === '[object Date]') return Utilities.formatDate(v, VA.TIMEZONE, 'yyyy-MM-dd');
  return String(v);
}

function sessionRows_(sheet, id) {
  var last = sheet.getLastRow();
  if (last < 2) return [];
  return sheet.getRange(2, 1, last - 1, VA.HEADERS.length).getValues()
    .filter(function (r) { return sessionKey_(r[1]) === id; });
}

function countBySession_() {
  var sheet = sheet_();
  var last = sheet.getLastRow();
  var counts = {};
  if (last < 2) return counts;
  sheet.getRange(2, 2, last - 1, 1).getValues().forEach(function (r) {
    var k = sessionKey_(r[0]);
    if (k) counts[k] = (counts[k] || 0) + 1;
  });
  return counts;
}

/** Trim, cap length, and block spreadsheet formula injection. */
function clean_(v, max) {
  var s = String(v == null ? '' : v).trim().slice(0, max);
  return s.replace(/^[=+\-@\t\r]+/, '');
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ------------------------------------------------------------------ */
/* Zenzap notification (optional)                                      */
/* Docs: https://docs.zenzap.co/api-reference/messages/send-a-message  */
/* ------------------------------------------------------------------ */

function notifyZenzap_(text, externalId) {
  var props = PropertiesService.getScriptProperties();
  var key = props.getProperty('ZENZAP_API_KEY');
  var secret = props.getProperty('ZENZAP_API_SECRET');
  var topic = props.getProperty('ZENZAP_TOPIC_ID');
  if (!key || !secret || !topic) return null;

  var payload = { topicId: topic, text: text };
  if (externalId) payload.externalId = String(externalId).slice(0, 60);
  var body = JSON.stringify(payload);
  var ts = String(Date.now());
  var sigBytes = Utilities.computeHmacSignature(
    Utilities.MacAlgorithm.HMAC_SHA_256, ts + '.' + body, secret, Utilities.Charset.UTF_8);
  var sig = sigBytes.map(function (b) { return ('0' + (b & 0xff).toString(16)).slice(-2); }).join('');

  var res = UrlFetchApp.fetch('https://api.zenzap.co/v2/messages', {
    method: 'post',
    contentType: 'application/json',
    payload: body,
    headers: { 'Authorization': 'Bearer ' + key, 'X-Signature': sig, 'X-Timestamp': ts },
    muteHttpExceptions: true
  });
  if (res.getResponseCode() >= 300) {
    console.error('Zenzap ' + res.getResponseCode() + ': ' + res.getContentText());
  }
  return res.getResponseCode();
}

/* ------------------------------------------------------------------ */
/* Run these from the Apps Script editor                               */
/* ------------------------------------------------------------------ */

/** Run once: creates/formats the Signups tab and asks for permissions. */
function setup() {
  var sh = sheet_();
  formatSheet_(sh);
  UrlFetchApp.fetch('https://www.google.com', { muteHttpExceptions: true }); // pre-authorize outbound requests for Zenzap
  Logger.log('Ready. Signups tab: ' + sh.getName());
}

/** Sends a test message to your Zenzap chat (after setting the Script Properties). */
function testZenzap() {
  var code = notifyZenzap_('Vert Academy test: sign-up alerts are connected.', 'va-test-' + Date.now());
  Logger.log(code === null ? 'Zenzap properties not set.' : 'Zenzap responded with HTTP ' + code);
}
