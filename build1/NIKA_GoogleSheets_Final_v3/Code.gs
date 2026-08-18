/* ============================================================================
   НІКА — Google Apps Script backend (фінальна версія)
   Сумісність зі старим фронтендом і структурою Google Sheets збережена.

   Головні принципи:
   - таблиця відкривається за стабільним ID, а не через "active spreadsheet" у web-app;
   - усі читання/записи виконуються пакетно;
   - одночасні записи захищені ScriptLock;
   - бізнес-формули та зв'язки 1:1 з попередньою версією;
   - структуровані логи з requestId, етапом і тривалістю;
   - помилки повертаються у зрозумілому вигляді.
   ============================================================================ */

var APP = Object.freeze({
  NAME: 'НІКА',
  VERSION: '2026.08.09-final-v2',
  SPREADSHEET_ID_KEY: 'SF_SPREADSHEET_ID',
  LOCK_TIMEOUT_MS: 15000,
  TELEGRAM_TIMEOUT_SEC: 12,
  TELEGRAM_DELAY_MS: 30 * 60 * 1000,
  TELEGRAM_RETRY_MS: 10 * 60 * 1000,
  TELEGRAM_PENDING_KEY: 'SF_TG_PENDING',
  TELEGRAM_LAST_CHANGE_KEY: 'SF_TG_LAST_CHANGE_MS',
  TELEGRAM_TRIGGER_HANDLER: 'flushTelegramDigest',
  TELEGRAM_MAX_PENDING: 80,
  MAX_CLIENT_LOG_LEN: 3500
});

var TABLES = Object.freeze({
  projects:  { name: 'Стройки',      headers: ['id','name','sort_order','created_by','created_at','updated_by','updated_at'] },
  entries:   { name: 'Витрати',      headers: ['id','project_id','section','category','description','amount','date','note','created_by','created_at','updated_by','updated_at'] },
  overheads: { name: 'Накладні',     headers: ['id','group_type','category','description','amount','date','created_by','created_at','updated_by','updated_at'] },
  contracts: { name: 'Договори',     headers: ['id','project_id','name','date','amount','created_by','created_at','updated_by','updated_at'] },
  acts:      { name: 'Акти',         headers: ['id','contract_id','number','date','amount','created_by','created_at','updated_by','updated_at'] },
  users:     { name: 'Користувачі',  headers: ['code','name','password'] },
  settings:  { name: 'Налаштування', headers: ['key','value'] },
  journal:   { name: 'Журнал',       headers: ['created_at','user','action','details'] }
});

var SECTION_KEYS = Object.freeze(['revenue','cashless','asset','cash','payroll']);
var SECTION_LABELS = Object.freeze({ revenue: 'Виручка', cashless: 'Безготівка', asset: 'Активи', cash: 'Готівка', payroll: 'ФОП' });
var GROUP_LABELS = Object.freeze({ cashless: 'Безготівкові', cash: 'Готівкові' });

var _spreadsheetCache = null;
var _requestContext = null;

/* ================================ LOGGING ================================= */

function makeRequestId_() {
  return Utilities.getUuid().replace(/-/g, '').slice(0, 10);
}

function nowMs_() { return new Date().getTime(); }

function log_(level, message, data) {
  var ctx = _requestContext || {};
  var prefix = '[' + APP.NAME + '][' + APP.VERSION + ']'
    + (ctx.requestId ? '[' + ctx.requestId + ']' : '')
    + '[' + String(level || 'INFO').toUpperCase() + '] ';
  var tail = '';
  if (data !== undefined) {
    try { tail = ' ' + JSON.stringify(data); }
    catch (e) { tail = ' ' + String(data); }
  }
  console.log(prefix + String(message || '') + tail);
}

function withRequestContext_(label, fn) {
  var prev = _requestContext;
  var started = nowMs_();
  _requestContext = { requestId: makeRequestId_(), label: label || 'request', started: started };
  log_('INFO', 'START ' + _requestContext.label);
  try {
    var result = fn();
    log_('INFO', 'DONE ' + _requestContext.label, { ms: nowMs_() - started });
    return result;
  } catch (err) {
    var safe = normalizeError_(err);
    log_('ERROR', 'FAIL ' + _requestContext.label, { ms: nowMs_() - started, message: safe.message, stack: safe.stack });
    throw new Error('[' + _requestContext.requestId + '] ' + safe.message);
  } finally {
    _requestContext = prev;
  }
}

function normalizeError_(err) {
  var message = err && err.message ? String(err.message) : String(err || 'Невідома помилка');
  var stack = err && err.stack ? String(err.stack) : '';
  if (stack.length > 1800) stack = stack.slice(0, 1800) + '…';
  return { message: message, stack: stack };
}

function reportClientError(payload, userCode) {
  return withRequestContext_('clientError', function () {
    try {
      var p = payload && typeof payload === 'object' ? payload : { message: String(payload || '') };
      var user = resolveUser_(userCode);
      log_('CLIENT', 'Browser error', {
        user: user ? user.name : String(userCode || ''),
        message: String(p.message || '').slice(0, APP.MAX_CLIENT_LOG_LEN),
        source: String(p.source || ''),
        line: p.line || '',
        column: p.column || '',
        stack: String(p.stack || '').slice(0, APP.MAX_CLIENT_LOG_LEN),
        href: String(p.href || '').slice(0, 500)
      });
    } catch (e) {
      log_('ERROR', 'Client-log handler failed', { error: String(e) });
    }
    return true;
  });
}

/* ========================== SPREADSHEET CONNECTION ========================= */

function getScriptProperties_() {
  return PropertiesService.getScriptProperties();
}

function setSpreadsheetId(spreadsheetId) {
  var id = String(spreadsheetId || '').trim();
  if (!id) throw new Error('Передайте ID Google Таблиці');
  var ss = SpreadsheetApp.openById(id);
  getScriptProperties_().setProperty(APP.SPREADSHEET_ID_KEY, ss.getId());
  _spreadsheetCache = ss;
  log_('INFO', 'Spreadsheet ID saved', { id: ss.getId(), name: ss.getName() });
  return { id: ss.getId(), name: ss.getName() };
}

function getSpreadsheet_() {
  if (_spreadsheetCache) return _spreadsheetCache;

  var props = getScriptProperties_();
  var savedId = String(props.getProperty(APP.SPREADSHEET_ID_KEY) || '').trim();
  if (savedId) {
    try {
      _spreadsheetCache = SpreadsheetApp.openById(savedId);
      return _spreadsheetCache;
    } catch (e) {
      log_('WARN', 'Saved spreadsheet ID cannot be opened', { id: savedId, error: String(e) });
    }
  }

  var active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) {
    props.setProperty(APP.SPREADSHEET_ID_KEY, active.getId());
    _spreadsheetCache = active;
    log_('INFO', 'Spreadsheet ID auto-detected and saved', { id: active.getId(), name: active.getName() });
    return _spreadsheetCache;
  }

  throw new Error(
    'Google Таблицю не прив’язано. Відкрийте Apps Script із потрібної таблиці та один раз запустіть setupSheets(), ' +
    'або виконайте setSpreadsheetId("ID_ТАБЛИЦІ").'
  );
}

function ensureTable_(key) {
  var spec = TABLES[key];
  if (!spec) throw new Error('Невідома таблиця: ' + key);
  var ss = getSpreadsheet_();
  var sheet = ss.getSheetByName(spec.name);
  if (!sheet) {
    sheet = ss.insertSheet(spec.name);
    log_('INFO', 'Created sheet', { key: key, name: spec.name });
  }
  var headerRange = sheet.getRange(1, 1, 1, spec.headers.length);
  var current = headerRange.getValues()[0];
  var differs = false;
  for (var i = 0; i < spec.headers.length; i++) {
    if (String(current[i] || '') !== spec.headers[i]) { differs = true; break; }
  }
  if (differs) headerRange.setValues([spec.headers.slice()]);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
  return sheet;
}

function getTab_(key) {
  var spec = TABLES[key];
  if (!spec) throw new Error('Невідома таблиця: ' + key);
  var sheet = getSpreadsheet_().getSheetByName(spec.name);
  return sheet || ensureTable_(key);
}

function readRows_(key) {
  var spec = TABLES[key];
  if (!spec) throw new Error('Невідома таблиця: ' + key);
  var sheet = getSpreadsheet_().getSheetByName(spec.name);
  if (!sheet || sheet.getLastRow() < 2) return [];

  var count = sheet.getLastRow() - 1;
  var values = sheet.getRange(2, 1, count, spec.headers.length).getValues();
  var out = [];

  for (var i = 0; i < values.length; i++) {
    var first = values[i][0];
    if (first === '' || first === null || first === undefined) continue;
    var row = {};
    for (var j = 0; j < spec.headers.length; j++) row[spec.headers[j]] = values[i][j];
    out.push(row);
  }
  return out;
}

function writeRows_(key, rows) {
  var spec = TABLES[key];
  if (!spec) throw new Error('Невідома таблиця: ' + key);
  rows = Array.isArray(rows) ? rows : [];

  var sheet = getTab_(key);
  var existing = Math.max(0, sheet.getLastRow() - 1);
  var target = rows.length;
  var clearCount = Math.max(existing, target);

  if (clearCount > 0) sheet.getRange(2, 1, clearCount, spec.headers.length).clearContent();
  if (target > 0) {
    var matrix = rows.map(function (row) {
      return spec.headers.map(function (h) {
        var v = row[h];
        return (v === undefined || v === null) ? '' : v;
      });
    });
    sheet.getRange(2, 1, matrix.length, spec.headers.length).setValues(matrix);
  }
}

function appendRecord_(key, row) {
  var spec = TABLES[key];
  var sheet = getTab_(key);
  sheet.appendRow(spec.headers.map(function (h) {
    var v = row[h];
    return (v === undefined || v === null) ? '' : v;
  }));
}

function withWriteLock_(label, fn) {
  var lock = LockService.getScriptLock();
  if (!lock.tryLock(APP.LOCK_TIMEOUT_MS)) {
    throw new Error('Дані зараз змінює інший користувач. Повторіть операцію через кілька секунд.');
  }
  try {
    log_('INFO', 'WRITE LOCK acquired', { action: label });
    return fn();
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

/* =============================== WEB APP ================================== */

function getIndexHtml_() {
  var names = ['Index', 'Index.html', 'index', 'index.html'];
  for (var i = 0; i < names.length; i++) {
    try { return HtmlService.createHtmlOutputFromFile(names[i]).getContent(); }
    catch (e) {}
  }
  throw new Error('Файл HTML із назвою Index не знайдено.');
}

function doGet(e) {
  return withRequestContext_('doGet', function () {
    var code = (e && e.parameter && e.parameter.u) ? String(e.parameter.u).trim() : '';
    var content = getIndexHtml_();
    var safeCode = code.replace(/\\/g, '\\\\').replace(/'/g, "\\'").replace(/\r/g, '').replace(/\n/g, '');
    return HtmlService.createHtmlOutput(content.replace('__USER_CODE__', safeCode))
      .setTitle(APP.NAME)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  });
}

function dispatch(method, url, body, userCode) {
  return withRequestContext_('dispatch', function () {
    var methodNorm = String(method || 'GET').toUpperCase();
    var urlNorm = String(url || '');
    var codeNorm = String(userCode || '').trim();
    log_('INFO', 'ROUTE', { method: methodNorm, url: urlNorm, userCode: codeNorm || '(empty)' });

    var user = resolveUser_(codeNorm);
    if (!user) {
      throw new Error('Невідомий код користувача (отримано: "' + String(userCode) + '"). Відкрийте застосунок за своїм персональним посиланням.');
    }

    var m;
    if (methodNorm === 'GET' && urlNorm === '/api/state') {
      return getState_();
    }

    if (methodNorm === 'POST' && urlNorm === '/api/projects') {
      return withWriteLock_('addProject', function () {
        var id = addProject_(user, body || {});
        return { id: id, state: getState_() };
      });
    }
    if (methodNorm === 'PUT' && (m = urlNorm.match(/^\/api\/projects\/(\d+)$/))) {
      return withWriteLock_('updateProject', function () { updateProject_(user, body || {}, +m[1]); return { state: getState_() }; });
    }
    if (methodNorm === 'DELETE' && (m = urlNorm.match(/^\/api\/projects\/(\d+)$/))) {
      return withWriteLock_('deleteProject', function () { deleteProject_(user, +m[1]); return { state: getState_() }; });
    }

    if (methodNorm === 'POST' && (m = urlNorm.match(/^\/api\/projects\/(\d+)\/entries$/))) {
      return withWriteLock_('addEntry', function () { addEntry_(user, body || {}, +m[1]); return { state: getState_() }; });
    }
    if (methodNorm === 'PUT' && (m = urlNorm.match(/^\/api\/entries\/(\d+)$/))) {
      return withWriteLock_('updateEntry', function () { updateEntry_(user, body || {}, +m[1]); return { state: getState_() }; });
    }
    if (methodNorm === 'DELETE' && (m = urlNorm.match(/^\/api\/entries\/(\d+)$/))) {
      return withWriteLock_('deleteEntry', function () { deleteEntry_(user, +m[1]); return { state: getState_() }; });
    }

    if (methodNorm === 'POST' && urlNorm === '/api/overheads') {
      return withWriteLock_('addOverhead', function () { addOverhead_(user, body || {}); return { state: getState_() }; });
    }
    if (methodNorm === 'PUT' && (m = urlNorm.match(/^\/api\/overheads\/(\d+)$/))) {
      return withWriteLock_('updateOverhead', function () { updateOverhead_(user, body || {}, +m[1]); return { state: getState_() }; });
    }
    if (methodNorm === 'DELETE' && (m = urlNorm.match(/^\/api\/overheads\/(\d+)$/))) {
      return withWriteLock_('deleteOverhead', function () { deleteOverhead_(user, +m[1]); return { state: getState_() }; });
    }

    if (methodNorm === 'POST' && urlNorm === '/api/contracts') {
      return withWriteLock_('addContract', function () { addContract_(user, body || {}); return { state: getState_() }; });
    }
    if (methodNorm === 'PUT' && (m = urlNorm.match(/^\/api\/contracts\/(\d+)$/))) {
      return withWriteLock_('updateContract', function () { updateContract_(user, body || {}, +m[1]); return { state: getState_() }; });
    }
    if (methodNorm === 'DELETE' && (m = urlNorm.match(/^\/api\/contracts\/(\d+)$/))) {
      return withWriteLock_('deleteContract', function () { deleteContract_(user, +m[1]); return { state: getState_() }; });
    }

    if (methodNorm === 'POST' && (m = urlNorm.match(/^\/api\/contracts\/(\d+)\/acts$/))) {
      return withWriteLock_('addAct', function () { addAct_(user, body || {}, +m[1]); return { state: getState_() }; });
    }
    if (methodNorm === 'PUT' && (m = urlNorm.match(/^\/api\/acts\/(\d+)$/))) {
      return withWriteLock_('updateAct', function () { updateAct_(user, body || {}, +m[1]); return { state: getState_() }; });
    }
    if (methodNorm === 'DELETE' && (m = urlNorm.match(/^\/api\/acts\/(\d+)$/))) {
      return withWriteLock_('deleteAct', function () { deleteAct_(user, +m[1]); return { state: getState_() }; });
    }

    if (methodNorm === 'POST' && urlNorm === '/api/settings') {
      return withWriteLock_('settings', function () {
        setSetting_('conversion_rate', num_(body && body.conversion_rate, 0.14));
        logAndNotify_(user, 'updated', 'ставку конвертації: ' + pctStr_(body && body.conversion_rate));
        return { state: getState_() };
      });
    }

    if (methodNorm === 'GET' && urlNorm === '/api/health') {
      return healthCheck_(user);
    }

    throw new Error('Невідомий запит: ' + methodNorm + ' ' + urlNorm);
  });
}

/* ========================= USERS / AUTH / SETTINGS ======================== */

function normCode_(s) {
  return String(s || '').trim().toLowerCase()
    .replace(/\u0456/g, 'i')
    .replace(/\u0406/g, 'i');
}

function num_(v, dflt) {
  var n = Number(v);
  if (isNaN(n)) return dflt !== undefined ? dflt : 0;
  return n;
}

function findById_(arr, id) {
  for (var i = 0; i < arr.length; i++) if (Number(arr[i].id) === Number(id)) return arr[i];
  return null;
}

function nextIdFromRows_(rows) {
  var max = 0;
  for (var i = 0; i < rows.length; i++) {
    var n = Number(rows[i].id) || 0;
    if (n > max) max = n;
  }
  return max + 1;
}

function readSettings_() {
  var out = {};
  readRows_('settings').forEach(function (r) { out[String(r.key)] = r.value; });
  return out;
}

function setSetting_(key, value) {
  var rows = readRows_('settings');
  var found = false;
  for (var i = 0; i < rows.length; i++) {
    if (String(rows[i].key) === String(key)) {
      rows[i].value = value;
      found = true;
      break;
    }
  }
  if (!found) rows.push({ key: key, value: value });
  writeRows_('settings', rows);
}

function readUsers_() {
  var map = {};
  readRows_('users').forEach(function (r) {
    if (!r.code) return;
    var code = String(r.code).trim();
    map[code] = { code: code, name: r.name, password: r.password };
  });
  return map;
}

function resolveUser_(code) {
  if (!code) return null;
  var needle = normCode_(code);
  var map = readUsers_();
  for (var key in map) {
    if (Object.prototype.hasOwnProperty.call(map, key) && normCode_(key) === needle) return map[key];
  }
  return null;
}

function resolveByName_(name) {
  if (!name) return null;
  var needle = normCode_(name);
  var map = readUsers_();
  for (var key in map) {
    if (Object.prototype.hasOwnProperty.call(map, key) && normCode_(map[key].name) === needle) return map[key];
  }
  return null;
}

function hashPw_(pw) {
  var raw = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, String(pw), Utilities.Charset.UTF_8);
  return raw.map(function (b) {
    var v = b < 0 ? b + 256 : b;
    var s = v.toString(16);
    return s.length < 2 ? '0' + s : s;
  }).join('');
}

function hashPw(pw) { return hashPw_(pw); }

function login(loginValue, password) {
  return withRequestContext_('login', function () {
    ensureTable_('users');
    var user = resolveUser_(loginValue) || resolveByName_(loginValue);
    if (!user) throw new Error('Користувача з таким логіном не знайдено');

    var stored = String(user.password || '').trim();
    var typed = String(password || '');
    if (!stored || hashPw_(typed) !== stored) throw new Error('Невірний пароль');

    log_('INFO', 'Login success', { code: user.code, name: user.name });
    return { code: user.code, name: user.name };
  });
}

function getUserInfo(userCode) {
  var user = resolveUser_(userCode);
  return user ? { name: user.name, code: String(user.code).trim() } : null;
}

function setPasswords() {
  return withRequestContext_('setPasswords', function () {
    var passwordMap = { oleksandr: '1952', vitaliy: '2341', andriy: '2242', anya: '3242' };
    ensureTable_('users');
    var rows = readRows_('users');
    rows.forEach(function (r) {
      var key = normCode_(r.code);
      if (Object.prototype.hasOwnProperty.call(passwordMap, key)) r.password = hashPw_(passwordMap[key]);
    });
    writeRows_('users', rows);
    log_('INFO', 'Passwords updated', { users: rows.length });
    return true;
  });
}

/* ======================== AUDIT + TELEGRAM NOTIFY ========================= */

function fmtDateTime_(d) {
  return Utilities.formatDate(new Date(d), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm');
}

function dateOnly_(value) {
  if (value === '' || value === null || value === undefined) return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  var s = String(value).trim();
  if (!s) return '';
  var iso = s.match(/^(\d{4}-\d{2}-\d{2})/);
  if (iso) return iso[1];
  var d = new Date(value);
  if (!isNaN(d.getTime())) return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  return s;
}

function pctStr_(v) { return (Math.round(num_(v, 0) * 1000) / 10) + '%'; }
function moneyStr_(v) { return Math.round(num_(v, 0) * 100) / 100 + ' грн'; }

function escTg_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
}

function logAndNotify_(user, action, details) {
  var when = new Date();
  appendRecord_('journal', {
    created_at: fmtDateTime_(when),
    user: user.name,
    action: action,
    details: details
  });
  log_('AUDIT', action, { user: user.name, details: details });
  queueTelegramChange_(user.name, action, details, when);
}

/*
  Telegram працює як debounce: повідомлення відправляється один раз,
  коли після ОСТАННЬОЇ зміни минуло не менше 30 хвилин.
  Нові зміни накопичуються у зведення. Якщо одноразовий тригер спрацював
  раніше ніж через 30 хвилин від останньої зміни, він просто переноситься.
*/
function queueTelegramChange_(userName, action, details, when) {
  var props = getScriptProperties_();
  var pending = [];
  try { pending = JSON.parse(props.getProperty(APP.TELEGRAM_PENDING_KEY) || '[]'); }
  catch (e) { pending = []; }
  if (!Array.isArray(pending)) pending = [];

  pending.push({
    ts: (when || new Date()).getTime(),
    user: String(userName || 'Користувач'),
    action: String(action || 'updated'),
    details: String(details || '')
  });
  if (pending.length > APP.TELEGRAM_MAX_PENDING) {
    pending = pending.slice(pending.length - APP.TELEGRAM_MAX_PENDING);
  }

  var now = (when || new Date()).getTime();
  props.setProperty(APP.TELEGRAM_PENDING_KEY, JSON.stringify(pending));
  props.setProperty(APP.TELEGRAM_LAST_CHANGE_KEY, String(now));
  ensureTelegramDigestTrigger_(APP.TELEGRAM_DELAY_MS);
  log_('INFO', 'Telegram digest queued', { pending: pending.length, delayMin: 30 });
}

function telegramDigestTriggers_() {
  return ScriptApp.getProjectTriggers().filter(function (tr) {
    return tr.getHandlerFunction && tr.getHandlerFunction() === APP.TELEGRAM_TRIGGER_HANDLER;
  });
}

function ensureTelegramDigestTrigger_(delayMs) {
  var existing = telegramDigestTriggers_();
  if (existing.length) return;
  ScriptApp.newTrigger(APP.TELEGRAM_TRIGGER_HANDLER)
    .timeBased()
    .after(Math.max(60 * 1000, Number(delayMs) || APP.TELEGRAM_DELAY_MS))
    .create();
  log_('INFO', 'Telegram digest trigger created', { delayMs: delayMs });
}

function clearTelegramDigestTriggers_() {
  telegramDigestTriggers_().forEach(function (tr) {
    try { ScriptApp.deleteTrigger(tr); } catch (e) {}
  });
}

function flushTelegramDigest() {
  return withRequestContext_('flushTelegramDigest', function () {
    return withWriteLock_('telegramDigest', function () {
      var props = getScriptProperties_();
      var lastChange = Number(props.getProperty(APP.TELEGRAM_LAST_CHANGE_KEY) || 0);
      var now = nowMs_();
      var elapsed = lastChange ? now - lastChange : APP.TELEGRAM_DELAY_MS;

      clearTelegramDigestTriggers_();
      if (lastChange && elapsed < APP.TELEGRAM_DELAY_MS) {
        var remaining = APP.TELEGRAM_DELAY_MS - elapsed;
        ensureTelegramDigestTrigger_(remaining);
        log_('INFO', 'Telegram digest postponed', { remainingMs: remaining });
        return { ok: true, postponed: true, remainingMs: remaining };
      }

      var pending = [];
      try { pending = JSON.parse(props.getProperty(APP.TELEGRAM_PENDING_KEY) || '[]'); }
      catch (e) { pending = []; }
      if (!Array.isArray(pending) || !pending.length) {
        props.deleteProperty(APP.TELEGRAM_PENDING_KEY);
        props.deleteProperty(APP.TELEGRAM_LAST_CHANGE_KEY);
        return { ok: true, sent: false, reason: 'empty' };
      }

      var sendResult = sendTelegramDigest_(pending);
      if (sendResult.ok) {
        props.deleteProperty(APP.TELEGRAM_PENDING_KEY);
        props.deleteProperty(APP.TELEGRAM_LAST_CHANGE_KEY);
        log_('INFO', 'Telegram digest sent and cleared', { items: pending.length });
        return { ok: true, sent: true, items: pending.length };
      }

      if (!sendResult.configured) {
        props.deleteProperty(APP.TELEGRAM_PENDING_KEY);
        props.deleteProperty(APP.TELEGRAM_LAST_CHANGE_KEY);
        log_('WARN', 'Telegram digest dropped because Telegram is not configured', { items: pending.length });
        return { ok: false, sent: false, configured: false };
      }

      // Do not retry permanent Telegram API errors every 10 minutes. A bad
      // token or chat_id would otherwise resend the same digest forever to
      // every chat that accepted it.
      if (sendResult.permanentError) {
        props.deleteProperty(APP.TELEGRAM_PENDING_KEY);
        props.deleteProperty(APP.TELEGRAM_LAST_CHANGE_KEY);
        log_('WARN', 'Telegram digest dropped after permanent API error', {
          items: pending.length,
          error: sendResult.error || ''
        });
        return { ok: false, sent: false, configured: true, permanentError: true, error: sendResult.error || '' };
      }

      ensureTelegramDigestTrigger_(APP.TELEGRAM_RETRY_MS);
      log_('WARN', 'Telegram digest retained for retry', { items: pending.length });
      return { ok: false, sent: false, retry: true, error: sendResult.error || '' };
    });
  });
}

function sendTelegramDigest_(pending) {
  var settings = readSettings_();
  var token = String(settings.telegram_bot_token || '').trim();
  var chatIds = String(settings.chat_ids || '')
    .split(',')
    .map(function (x) { return String(x).trim(); })
    .filter(Boolean);

  if (!token || !chatIds.length) {
    log_('WARN', 'Telegram is not configured', { token: !!token, chats: chatIds.length });
    return { ok: false, configured: false };
  }

  var actionIcon = { added: '🟢', updated: '🟡', deleted: '🔴' };
  var lines = pending.map(function (item) {
    var t = Utilities.formatDate(new Date(item.ts), Session.getScriptTimeZone(), 'dd.MM HH:mm');
    return (actionIcon[item.action] || '🔹') + ' <b>' + escTg_(item.user) + '</b> — ' + escTg_(item.details) + ' <i>(' + t + ')</i>';
  });
  var header = '📋 <b>' + escTg_(APP.NAME) + ' — зміни за останній період</b>\n'
    + 'Остання зміна: ' + Utilities.formatDate(new Date(pending[pending.length - 1].ts), Session.getScriptTimeZone(), 'dd.MM.yyyy HH:mm') + '\n\n';
  var chunks = splitTelegramText_(header, lines, 3600);
  var url = 'https://api.telegram.org/bot' + token + '/sendMessage';
  var requests = [];
  chatIds.forEach(function (cid) {
    chunks.forEach(function (text) {
      requests.push({
        url: url,
        method: 'post',
        payload: { chat_id: cid, text: text, parse_mode: 'HTML', disable_web_page_preview: true },
        muteHttpExceptions: true,
        timeoutSeconds: APP.TELEGRAM_TIMEOUT_SEC
      });
    });
  });

  try {
    var responses = UrlFetchApp.fetchAll(requests);
    var codes = responses.map(function (r) { return r.getResponseCode(); });
    var ok = codes.every(function (c) { return c >= 200 && c < 300; });
    var errors = responses.map(function (r, i) {
      if (r.getResponseCode() >= 200 && r.getResponseCode() < 300) return '';
      var body = String(r.getContentText() || '').slice(0, 500);
      return 'request ' + (i + 1) + ', HTTP ' + r.getResponseCode() + ': ' + body;
    }).filter(Boolean);
    var error = errors.join(' | ');
    var permanentError = responses.some(function (r) {
      var code = r.getResponseCode();
      return code >= 400 && code < 500 && code !== 429;
    });
    log_(ok ? 'INFO' : 'WARN', 'Telegram digest response', {
      chats: chatIds.length,
      chunks: chunks.length,
      codes: codes,
      error: error
    });
    return { ok: ok, configured: true, permanentError: permanentError, error: error };
  } catch (e) {
    var fetchError = String(e && e.message ? e.message : e);
    log_('WARN', 'Telegram digest failed', { error: fetchError });
    return { ok: false, configured: true, error: fetchError };
  }
}

function splitTelegramText_(header, lines, maxLen) {
  maxLen = Math.max(1000, Number(maxLen) || 3600);
  var out = [];
  var cur = String(header || '');
  (lines || []).forEach(function (line) {
    var piece = String(line || '') + '\n';
    if (cur.length + piece.length > maxLen && cur.trim()) {
      out.push(cur.trim());
      cur = '';
    }
    if (piece.length > maxLen) piece = piece.slice(0, maxLen - 2) + '…\n';
    cur += piece;
  });
  if (cur.trim()) out.push(cur.trim());
  return out.length ? out : [String(header || '').trim()];
}

/* Встановлюваний onEdit-тригер: ловить ручні зміни у бізнес-аркушах. */
function onSpreadsheetEdit_(e) {
  try {
    if (!e || !e.range) return;
    var sheet = e.range.getSheet();
    var businessSheets = ['Стройки','Витрати','Накладні','Договори','Акти'];
    if (businessSheets.indexOf(sheet.getName()) < 0) return;

    withRequestContext_('onSpreadsheetEdit', function () {
      withWriteLock_('manualSheetEdit', function () {
        var editor = 'Ручна зміна';
        try {
          if (e.user && e.user.getEmail) editor = e.user.getEmail() || editor;
        } catch (ignore) {}
        var details = 'ручна зміна в аркуші «' + sheet.getName() + '», діапазон ' + e.range.getA1Notation();
        appendRecord_('journal', {
          created_at: fmtDateTime_(new Date()),
          user: editor,
          action: 'updated',
          details: details
        });
        queueTelegramChange_(editor, 'updated', details, new Date());
      });
    });
  } catch (err) {
    console.log('[NIKA][onSpreadsheetEdit][ERROR] ' + String(err && err.stack ? err.stack : err));
  }
}

function ensureSpreadsheetEditTrigger_() {
  var ss = getSpreadsheet_();
  var triggers = ScriptApp.getProjectTriggers().filter(function (tr) {
    return tr.getHandlerFunction && tr.getHandlerFunction() === 'onSpreadsheetEdit_';
  });
  if (triggers.length) return;
  ScriptApp.newTrigger('onSpreadsheetEdit_').forSpreadsheet(ss).onEdit().create();
  log_('INFO', 'Installable spreadsheet edit trigger created', { spreadsheetId: ss.getId() });
}

function testTelegramNow() {
  return withRequestContext_('testTelegramNow', function () {
    var result = sendTelegramDigest_([{ ts: nowMs_(), user: 'TEST', action: 'updated', details: 'тестове повідомлення Telegram' }]);
    if (!result.ok) {
      var reason = result.error ? ' Деталі: ' + result.error : '';
      throw new Error('Telegram не налаштований або API повернув помилку. Перевірте telegram_bot_token та chat_ids у «Налаштування».' + reason);
    }
    return { ok: true };
  });
}

/* ============================== BUSINESS DATA ============================== */

function projectName_(id) {
  var p = findById_(readRows_('projects'), id);
  return p ? p.name : String(id);
}

function contractName_(id) {
  var c = findById_(readRows_('contracts'), id);
  return c ? c.name : String(id);
}

function addProject_(user, payload) {
  var name = String(payload.name || '').trim();
  if (!name) throw new Error('Назва стройки порожня');

  var rows = readRows_('projects');
  var id = nextIdFromRows_(rows);
  var now = fmtDateTime_(new Date());
  rows.push({ id: id, name: name, sort_order: id, created_by: user.name, created_at: now, updated_by: user.name, updated_at: now });
  writeRows_('projects', rows);
  logAndNotify_(user, 'added', 'стройку «' + name + '»');
  return id;
}

function updateProject_(user, payload, id) {
  var rows = readRows_('projects');
  var project = findById_(rows, id);
  if (!project) throw new Error('Стройку не знайдено');

  var oldName = project.name;
  var name = String(payload.name || '').trim();
  if (!name) throw new Error('Назва стройки порожня');

  project.name = name;
  project.updated_by = user.name;
  project.updated_at = fmtDateTime_(new Date());
  writeRows_('projects', rows);
  logAndNotify_(user, 'updated', 'стройку: «' + oldName + '» → «' + name + '»');
}

function deleteProject_(user, id) {
  var projectsAll = readRows_('projects');
  var project = findById_(projectsAll, id);
  if (!project) throw new Error('Стройку не знайдено');

  var contractsAll = readRows_('contracts');
  var contractIds = contractsAll
    .filter(function (r) { return Number(r.project_id) === Number(id); })
    .map(function (r) { return Number(r.id); });

  writeRows_('projects', projectsAll.filter(function (r) { return Number(r.id) !== Number(id); }));
  writeRows_('entries', readRows_('entries').filter(function (r) { return Number(r.project_id) !== Number(id); }));
  writeRows_('contracts', contractsAll.filter(function (r) { return Number(r.project_id) !== Number(id); }));
  writeRows_('acts', readRows_('acts').filter(function (r) { return contractIds.indexOf(Number(r.contract_id)) < 0; }));
  logAndNotify_(user, 'deleted', 'стройку «' + project.name + '» і всі її записи');
}

function addEntry_(user, payload, projectId) {
  var project = findById_(readRows_('projects'), projectId);
  if (!project) throw new Error('Стройку не знайдено');
  var section = payload.section;
  if (SECTION_KEYS.indexOf(section) < 0) throw new Error('Невірний розділ');

  var rows = readRows_('entries');
  var id = nextIdFromRows_(rows);
  var now = fmtDateTime_(new Date());
  var row = {
    id: id, project_id: projectId, section: section,
    category: payload.category || '', description: payload.description || '',
    amount: num_(payload.amount), date: payload.date || '', note: payload.note || '',
    created_by: user.name, created_at: now, updated_by: user.name, updated_at: now
  };
  rows.push(row);
  writeRows_('entries', rows);
  logAndNotify_(user, 'added', entryDetails_(row, project.name));
}

function updateEntry_(user, payload, id) {
  var rows = readRows_('entries');
  var entry = findById_(rows, id);
  if (!entry) throw new Error('Запис не знайдено');

  var section = payload.section || entry.section;
  if (SECTION_KEYS.indexOf(section) < 0) throw new Error('Невірний розділ');

  entry.section = section;
  entry.category = payload.category !== undefined ? payload.category : entry.category;
  entry.description = payload.description !== undefined ? payload.description : entry.description;
  entry.amount = payload.amount !== undefined ? num_(payload.amount) : entry.amount;
  entry.date = payload.date !== undefined ? payload.date : entry.date;
  entry.note = payload.note !== undefined ? payload.note : entry.note;
  entry.updated_by = user.name;
  entry.updated_at = fmtDateTime_(new Date());
  writeRows_('entries', rows);
  logAndNotify_(user, 'updated', entryDetails_(entry, projectName_(entry.project_id)) + ' (змінено)');
}

function deleteEntry_(user, id) {
  var rows = readRows_('entries');
  var idx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (Number(rows[i].id) === Number(id)) { idx = i; break; }
  }
  if (idx < 0) throw new Error('Запис не знайдено');

  var entry = rows[idx];
  rows.splice(idx, 1);
  writeRows_('entries', rows);
  logAndNotify_(user, 'deleted', entryDetails_(entry, projectName_(entry.project_id)));
}

function entryDetails_(e, projectName) {
  var sec = SECTION_LABELS[e.section] || e.section;
  var s = 'запис у проєкті «' + projectName + '», розділ «' + sec + '»';
  if (e.category) s += ', «' + e.category + '»';
  if (e.description) s += ', «' + e.description + '»';
  s += ', ' + moneyStr_(e.amount);
  if (e.date) s += ' (' + dateOnly_(e.date) + ')';
  return s;
}

function addOverhead_(user, payload) {
  var group = payload.group_type;
  if (group !== 'cashless' && group !== 'cash') throw new Error('Невірний тип накладних');

  var rows = readRows_('overheads');
  var id = nextIdFromRows_(rows);
  var now = fmtDateTime_(new Date());
  var row = {
    id: id, group_type: group, category: payload.category || '', description: payload.description || '',
    amount: num_(payload.amount), date: payload.date || '',
    created_by: user.name, created_at: now, updated_by: user.name, updated_at: now
  };
  rows.push(row);
  writeRows_('overheads', rows);
  logAndNotify_(user, 'added', overheadDetails_(row));
}

function updateOverhead_(user, payload, id) {
  var rows = readRows_('overheads');
  var entry = findById_(rows, id);
  if (!entry) throw new Error('Накладну витрату не знайдено');

  var group = payload.group_type || entry.group_type;
  if (group !== 'cashless' && group !== 'cash') throw new Error('Невірний тип накладних');

  entry.group_type = group;
  entry.category = payload.category !== undefined ? payload.category : entry.category;
  entry.description = payload.description !== undefined ? payload.description : entry.description;
  entry.amount = payload.amount !== undefined ? num_(payload.amount) : entry.amount;
  entry.date = payload.date !== undefined ? payload.date : entry.date;
  entry.updated_by = user.name;
  entry.updated_at = fmtDateTime_(new Date());
  writeRows_('overheads', rows);
  logAndNotify_(user, 'updated', overheadDetails_(entry) + ' (змінено)');
}

function deleteOverhead_(user, id) {
  var rows = readRows_('overheads');
  var idx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (Number(rows[i].id) === Number(id)) { idx = i; break; }
  }
  if (idx < 0) throw new Error('Накладну витрату не знайдено');

  var entry = rows[idx];
  rows.splice(idx, 1);
  writeRows_('overheads', rows);
  logAndNotify_(user, 'deleted', overheadDetails_(entry));
}

function overheadDetails_(e) {
  var group = GROUP_LABELS[e.group_type] || e.group_type;
  var s = 'накладну витрату «' + (e.description || '') + '» (' + group;
  if (e.category) s += ', ' + e.category;
  s += '), ' + moneyStr_(e.amount);
  if (e.date) s += ' (' + dateOnly_(e.date) + ')';
  return s;
}

function addContract_(user, payload) {
  var project = findById_(readRows_('projects'), payload.project_id);
  if (!project) throw new Error('Стройку не знайдено');

  var rows = readRows_('contracts');
  var id = nextIdFromRows_(rows);
  var now = fmtDateTime_(new Date());
  var row = {
    id: id, project_id: num_(payload.project_id), name: payload.name || '',
    date: payload.date || '', amount: num_(payload.amount),
    created_by: user.name, created_at: now, updated_by: user.name, updated_at: now
  };
  rows.push(row);
  writeRows_('contracts', rows);
  logAndNotify_(user, 'added', 'договір «' + (row.name || '') + '», ' + moneyStr_(row.amount) + ' (проєкт «' + project.name + '»)');
}

function updateContract_(user, payload, id) {
  var rows = readRows_('contracts');
  var contract = findById_(rows, id);
  if (!contract) throw new Error('Договір не знайдено');

  contract.project_id = payload.project_id !== undefined ? num_(payload.project_id) : contract.project_id;
  contract.name = payload.name !== undefined ? payload.name : contract.name;
  contract.date = payload.date !== undefined ? payload.date : contract.date;
  contract.amount = payload.amount !== undefined ? num_(payload.amount) : contract.amount;
  contract.updated_by = user.name;
  contract.updated_at = fmtDateTime_(new Date());
  writeRows_('contracts', rows);
  logAndNotify_(user, 'updated', 'договір «' + (contract.name || '') + '» (змінено)');
}

function deleteContract_(user, id) {
  var rows = readRows_('contracts');
  var contract = findById_(rows, id);
  if (!contract) throw new Error('Договір не знайдено');

  writeRows_('contracts', rows.filter(function (r) { return Number(r.id) !== Number(id); }));
  writeRows_('acts', readRows_('acts').filter(function (r) { return Number(r.contract_id) !== Number(id); }));
  logAndNotify_(user, 'deleted', 'договір «' + (contract.name || '') + '» і всі його акти');
}

function addAct_(user, payload, contractId) {
  var contract = findById_(readRows_('contracts'), contractId);
  if (!contract) throw new Error('Договір не знайдено');

  var rows = readRows_('acts');
  var id = nextIdFromRows_(rows);
  var now = fmtDateTime_(new Date());
  var row = {
    id: id, contract_id: contractId, number: payload.number || '',
    date: payload.date || '', amount: num_(payload.amount),
    created_by: user.name, created_at: now, updated_by: user.name, updated_at: now
  };
  rows.push(row);
  writeRows_('acts', rows);
  logAndNotify_(user, 'added', 'акт №' + (row.number || '') + ' до договору «' + contract.name + '», ' + moneyStr_(row.amount));
}

function updateAct_(user, payload, id) {
  var rows = readRows_('acts');
  var act = findById_(rows, id);
  if (!act) throw new Error('Акт не знайдено');

  act.number = payload.number !== undefined ? payload.number : act.number;
  act.date = payload.date !== undefined ? payload.date : act.date;
  act.amount = payload.amount !== undefined ? num_(payload.amount) : act.amount;
  act.updated_by = user.name;
  act.updated_at = fmtDateTime_(new Date());
  writeRows_('acts', rows);
  logAndNotify_(user, 'updated', 'акт №' + (act.number || '') + ' до договору «' + contractName_(act.contract_id) + '» (змінено)');
}

function deleteAct_(user, id) {
  var rows = readRows_('acts');
  var idx = -1;
  for (var i = 0; i < rows.length; i++) {
    if (Number(rows[i].id) === Number(id)) { idx = i; break; }
  }
  if (idx < 0) throw new Error('Акт не знайдено');

  var act = rows[idx];
  rows.splice(idx, 1);
  writeRows_('acts', rows);
  logAndNotify_(user, 'deleted', 'акт №' + (act.number || '') + ' до договору «' + contractName_(act.contract_id) + '»');
}

/* =========================== STATE + CALCULATIONS ========================== */

function readStateSnapshot_() {
  return {
    settings: readSettings_(),
    projects: readRows_('projects'),
    entries: readRows_('entries'),
    overheads: readRows_('overheads'),
    contracts: readRows_('contracts'),
    acts: readRows_('acts')
  };
}

function getState_() {
  var started = nowMs_();
  var raw = readStateSnapshot_();
  var rate = num_(raw.settings.conversion_rate, 0.14);

  var projects = raw.projects
    .map(function (r) {
      return { id: num_(r.id), name: r.name, sort_order: num_(r.sort_order), entries: [] };
    })
    .sort(function (a, b) { return (a.sort_order || a.id) - (b.sort_order || b.id); });

  var projectMap = {};
  projects.forEach(function (p) { projectMap[String(p.id)] = p; });

  raw.entries.forEach(function (r) {
    var p = projectMap[String(num_(r.project_id))];
    if (!p) return;
    p.entries.push({
      id: num_(r.id), section: r.section, category: r.category, description: r.description,
      amount: num_(r.amount), date: dateOnly_(r.date), note: r.note
    });
  });
  projects.forEach(function (p) {
    p.entries.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)) || a.id - b.id; });
  });

  var overheads = raw.overheads
    .map(function (r) {
      return {
        id: num_(r.id), group_type: r.group_type, category: r.category,
        description: r.description, amount: num_(r.amount), date: dateOnly_(r.date)
      };
    })
    .sort(function (a, b) { return String(a.date).localeCompare(String(b.date)) || a.id - b.id; });

  var contracts = raw.contracts
    .map(function (r) {
      return {
        id: num_(r.id), project_id: num_(r.project_id), name: r.name,
        date: dateOnly_(r.date), amount: num_(r.amount), acts: []
      };
    })
    .sort(function (a, b) { return String(a.date).localeCompare(String(b.date)) || a.id - b.id; });

  var contractMap = {};
  contracts.forEach(function (c) { contractMap[String(c.id)] = c; });
  raw.acts.forEach(function (r) {
    var c = contractMap[String(num_(r.contract_id))];
    if (!c) return;
    c.acts.push({ id: num_(r.id), number: r.number, date: dateOnly_(r.date), amount: num_(r.amount) });
  });
  contracts.forEach(function (c) {
    c.acts.sort(function (a, b) { return String(a.date).localeCompare(String(b.date)) || a.id - b.id; });
  });

  var state = {
    settings: { conversion_rate: rate },
    projects: projects,
    overheads: overheads,
    contracts: contracts
  };
  state.results = calculateResults_(state);

  log_('INFO', 'State built', {
    projects: projects.length,
    entries: raw.entries.length,
    overheads: overheads.length,
    contracts: contracts.length,
    acts: raw.acts.length,
    ms: nowMs_() - started
  });
  return state;
}

/*
  Формули нижче перенесені з остаточної локальної версії.
  НЕ змінювати без окремого рішення власника даних.
*/
function calculateResults_(state) {
  var rate = num_(state.settings.conversion_rate, 0);
  var overheadCashless = 0;
  var overheadCash = 0;

  (state.overheads || []).forEach(function (e) {
    if (e.group_type === 'cashless') overheadCashless += num_(e.amount);
    else if (e.group_type === 'cash') overheadCash += num_(e.amount);
  });

  var overheadAdjustment = overheadCashless * rate;
  var distributable = overheadCashless + overheadCash - overheadAdjustment;
  var baseRows = [];
  var totalTurnover = 0;
  var payrollTotal = 0;

  (state.projects || []).forEach(function (p) {
    var sums = { revenue: 0, cashless: 0, asset: 0, cash: 0, payroll: 0 };
    (p.entries || []).forEach(function (en) {
      if (Object.prototype.hasOwnProperty.call(sums, en.section)) {
        sums[en.section] += num_(en.amount);
        if (en.section === 'payroll' && !isAdminPayrollCategory_(en.category)) {
          payrollTotal += num_(en.amount);
        }
      }
    });

    var turnover = sums.revenue;
    var cashless = sums.cashless + sums.asset;
    var balance = turnover - cashless;
    var conversion = balance * rate;
    var directCash = sums.cash + sums.payroll;
    var profitBefore = turnover - cashless - conversion - directCash;

    baseRows.push({
      project_id: p.id,
      name: p.name,
      turnover: turnover,
      cashless: cashless,
      balance: balance,
      conversion: conversion,
      cash_expenses: directCash,
      profit_before_overhead: profitBefore
    });
    totalTurnover += turnover;
  });

  var rows = baseRows.map(function (r) {
    var allocated = totalTurnover ? distributable * r.turnover / totalTurnover : 0;
    var net = r.profit_before_overhead - allocated;
    var percent = r.turnover ? net / r.turnover : 0;
    return Object.assign({}, r, {
      allocated_overhead: allocated,
      net_profit: net,
      profit_percent: percent
    });
  });

  var sum = function (key) {
    return rows.reduce(function (acc, r) { return acc + num_(r[key]); }, 0);
  };
  var turnoverTotal = sum('turnover');
  var netProfitTotal = sum('net_profit');

  var totals = {
    turnover: turnoverTotal,
    allocated_overhead: distributable,
    cashless: sum('cashless'),
    balance: sum('balance') - overheadCashless,
    conversion: sum('conversion'),
    cash_expenses: sum('cash_expenses'),
    net_profit: netProfitTotal,
    profit_percent: turnoverTotal ? netProfitTotal / turnoverTotal : 0,
    overhead_cashless: overheadCashless,
    overhead_cash: overheadCash,
    overhead_adjustment: overheadAdjustment,
    overhead_coefficient: totalTurnover ? distributable / totalTurnover : 0,
    payroll_coefficient: totalTurnover ? payrollTotal / totalTurnover : 0
  };

  return { rows: rows, totals: totals };
}

function isAdminPayrollCategory_(category) {
  var c = String(category || '').toLowerCase();
  return c.indexOf('адміністрація') >= 0 || c.indexOf('администрация') >= 0 || c.indexOf('admin') >= 0;
}

/* ================================ EXPORT ================================== */

function exportExcel(userCode, page, projectId) {
  return withRequestContext_('exportExcel', function () {
    var user = resolveUser_(userCode);
    if (!user) throw new Error('Невідомий код користувача. Відкрийте застосунок за своїм персональним посиланням.');

    var specs = buildExportSpecs_(page, projectId);
    var filename = exportFileName_(page, projectId);
    var xlsxBlob = buildXlsxBlob_(specs, filename);
    var bytes = xlsxBlob.getBytes();

    log_('INFO', 'Excel generated in memory for direct download', {
      user: user.name,
      page: page || 'all',
      projectId: projectId || '',
      sheets: specs.map(function (x) { return x.name; }),
      bytes: bytes.length
    });
    return {
      filename: filename,
      mimeType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
      base64: Utilities.base64Encode(bytes)
    };
  });
}

function exportFileName_(page, projectId) {
  var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
  var suffix = 'all';
  var p = String(page || '').trim();
  if (p) suffix = p;
  if ((p === 'projects' || p === 'contracts') && projectId) {
    try { suffix += '_' + projectName_(projectId); } catch (e) {}
  }
  suffix = String(suffix).replace(/[\/:*?"<>|]+/g, '_').replace(/\s+/g, '_').slice(0, 70);
  return APP.NAME + '_' + suffix + '_' + stamp + '.xlsx';
}

/* Справжній XLSX формується в пам'яті без створення файлів у Google Drive. */
function buildXlsxBlob_(specs, filename) {
  var normalized = normalizeXlsxSpecs_(specs);
  var blobs = [];

  blobs.push(Utilities.newBlob(xlsxContentTypesXml_(normalized.length), 'application/xml', '[Content_Types].xml'));
  blobs.push(Utilities.newBlob(xlsxRootRelsXml_(), 'application/xml', '_rels/.rels'));
  blobs.push(Utilities.newBlob(xlsxWorkbookXml_(normalized), 'application/xml', 'xl/workbook.xml'));
  blobs.push(Utilities.newBlob(xlsxWorkbookRelsXml_(normalized.length), 'application/xml', 'xl/_rels/workbook.xml.rels'));

  normalized.forEach(function (spec, i) {
    blobs.push(Utilities.newBlob(xlsxSheetXml_(spec.rows), 'application/xml', 'xl/worksheets/sheet' + (i + 1) + '.xml'));
  });

  return Utilities.zip(blobs, filename || (APP.NAME + '.xlsx'));
}

function normalizeXlsxSpecs_(specs) {
  specs = specs && specs.length ? specs : [{ name: 'Результат', rows: [['Немає даних']] }];
  var used = {};
  return specs.map(function (spec, i) {
    var base = sanitizeXlsxSheetName_(spec.name || ('Sheet' + (i + 1)));
    var name = base;
    var n = 2;
    while (used[name.toLowerCase()]) {
      var suffix = ' (' + n++ + ')';
      name = base.slice(0, Math.max(1, 31 - suffix.length)) + suffix;
    }
    used[name.toLowerCase()] = true;
    return { name: name, rows: Array.isArray(spec.rows) ? spec.rows : [] };
  });
}

function sanitizeXlsxSheetName_(name) {
  var s = String(name || 'Sheet').replace(/[\/\?\*\[\]:]/g, ' ').replace(/\s+/g, ' ').trim();
  if (!s) s = 'Sheet';
  if (s.length > 31) s = s.slice(0, 31).trim();
  return s || 'Sheet';
}

function xlsxContentTypesXml_(sheetCount) {
  var overrides = '';
  for (var i = 1; i <= sheetCount; i++) {
    overrides += '<Override PartName="/xl/worksheets/sheet' + i + '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
    + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
    + '<Default Extension="xml" ContentType="application/xml"/>'
    + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
    + overrides + '</Types>';
}

function xlsxRootRelsXml_() {
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
    + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>'
    + '</Relationships>';
}

function xlsxWorkbookXml_(specs) {
  var sheets = specs.map(function (spec, i) {
    return '<sheet name="' + xmlEscape_(spec.name) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>';
  }).join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">'
    + '<sheets>' + sheets + '</sheets></workbook>';
}

function xlsxWorkbookRelsXml_(sheetCount) {
  var rels = '';
  for (var i = 1; i <= sheetCount; i++) {
    rels += '<Relationship Id="rId' + i + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' + i + '.xml"/>';
  }
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' + rels + '</Relationships>';
}

function xlsxSheetXml_(rows) {
  rows = Array.isArray(rows) ? rows : [];
  var sheetRows = rows.map(function (row, rIdx) {
    row = Array.isArray(row) ? row : [row];
    var cells = row.map(function (value, cIdx) {
      return xlsxCellXml_(value, cIdx + 1, rIdx + 1);
    }).join('');
    return '<row r="' + (rIdx + 1) + '">' + cells + '</row>';
  }).join('');
  return '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
    + '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'
    + '<sheetData>' + sheetRows + '</sheetData></worksheet>';
}

function xlsxCellXml_(value, col, row) {
  var ref = xlsxColName_(col) + row;
  if (value === null || value === undefined || value === '') return '<c r="' + ref + '"/>';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    value = dateOnly_(value);
  }
  if (typeof value === 'number' && isFinite(value)) {
    return '<c r="' + ref + '"><v>' + String(value) + '</v></c>';
  }
  if (typeof value === 'boolean') {
    return '<c r="' + ref + '" t="b"><v>' + (value ? '1' : '0') + '</v></c>';
  }
  var text = cleanXmlText_(String(value));
  return '<c r="' + ref + '" t="inlineStr"><is><t xml:space="preserve">' + xmlEscape_(text) + '</t></is></c>';
}

function xlsxColName_(n) {
  var out = '';
  while (n > 0) {
    n--;
    out = String.fromCharCode(65 + (n % 26)) + out;
    n = Math.floor(n / 26);
  }
  return out;
}

function cleanXmlText_(s) {
  return String(s || '').replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '');
}

function xmlEscape_(s) {
  return String(s || '').replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

function buildExportSpecs_(page, projectId) {
  var state = getState_();
  var pageNorm = String(page || '').trim();
  var pid = num_(projectId, 0);
  var project = null;
  for (var i = 0; i < state.projects.length; i++) {
    if (Number(state.projects[i].id) === Number(pid)) { project = state.projects[i]; break; }
  }

  if (pageNorm === 'dashboard') {
    var t = state.results.totals;
    var overview = [
      ['Валовий прибуток (усього)', t.turnover],
      ['Безготівковий залишок', t.balance],
      ['Накладні витрати', t.allocated_overhead],
      ['Чистий прибуток', t.net_profit],
      ['Кількість будівництв', state.projects.length],
      [''],
      ['№','Найменування','Валовий прибуток','Накладні','Безготівкові','Конвертація','Готівкові','Чистий прибуток','%']
    ];
    state.results.rows.forEach(function (r, idx) {
      overview.push([idx + 1, r.name, r.turnover, r.allocated_overhead, r.cashless, r.conversion, r.cash_expenses, r.net_profit, r.profit_percent]);
    });
    return [{ name: 'Огляд', rows: overview }];
  }

  if (pageNorm === 'projects') {
    var projectRows = [['Розділ','Категорія','Опис','Сума','Дата','Примітка']];
    if (project) {
      (project.entries || []).forEach(function (e) {
        projectRows.push([SECTION_LABELS[e.section] || e.section, e.category, e.description, e.amount, e.date, e.note || '']);
      });
    }
    return [{ name: 'Витрати — ' + (project ? project.name : ''), rows: projectRows }];
  }

  if (pageNorm === 'overheads') {
    var overheadRows = [['Тип','Категорія','Опис','Сума','Дата']];
    (state.overheads || []).forEach(function (o) {
      overheadRows.push([GROUP_LABELS[o.group_type] || o.group_type, o.category, o.description, o.amount, dateOnly_(o.date)]);
    });
    return [{ name: 'Накладні', rows: overheadRows }];
  }

  if (pageNorm === 'results') {
    return [{ name: 'Результат', rows: resultsSheetRows_(state) }];
  }

  if (pageNorm === 'contracts') {
    var contracts = (state.contracts || []).filter(function (c) {
      return !pid || Number(c.project_id) === Number(pid);
    });
    var contractRows = [['Назва','Дата','Сума']];
    var actRows = [['Договір','№','Дата','Сума']];
    contracts.forEach(function (c) {
      contractRows.push([c.name, dateOnly_(c.date), num_(c.amount)]);
      (c.acts || []).forEach(function (a) {
        actRows.push([c.name, a.number, dateOnly_(a.date), num_(a.amount)]);
      });
    });
    return [
      { name: 'Договори', rows: contractRows },
      { name: 'Акти', rows: actRows }
    ];
  }

  var specs = [];
  ['projects','entries','overheads','contracts','acts'].forEach(function (key) {
    var table = TABLES[key];
    var rows = [table.headers.slice()];
    readRows_(key).forEach(function (r) {
      rows.push(table.headers.map(function (h) {
        var v = r[h];
        return (v === undefined || v === null) ? '' : v;
      }));
    });
    specs.push({ name: table.name, rows: rows });
  });
  specs.push({ name: 'Результат', rows: resultsSheetRows_(state) });
  return specs;
}

function resultsSheetRows_(state) {
  var head = ['№','Найменування','Валовий прибуток','Накладні','Безготівкові','Конвертація','Готівкові','Чистий прибуток','%'];
  var rows = [head];
  state.results.rows.forEach(function (r, i) {
    rows.push([i + 1, r.name, r.turnover, r.allocated_overhead, r.cashless, r.conversion, r.cash_expenses, r.net_profit, r.profit_percent]);
  });
  var t = state.results.totals;
  rows.push(['','ВСЬОГО',t.turnover,t.allocated_overhead,t.cashless,t.conversion,t.cash_expenses,t.net_profit,t.profit_percent]);
  return rows;
}


/* ============================== SETUP / TESTS ============================== */

function setupSheets() {
  return withRequestContext_('setupSheets', function () {
    var ss = getSpreadsheet_();
    log_('INFO', 'Setup spreadsheet', { id: ss.getId(), name: ss.getName() });

    Object.keys(TABLES).forEach(function (key) {
      var t0 = nowMs_();
      ensureTable_(key);
      log_('INFO', 'Sheet checked', { key: key, ms: nowMs_() - t0 });
    });

    var settingRows = readRows_('settings');
    var settingMap = {};
    settingRows.forEach(function (r) { settingMap[String(r.key)] = true; });
    if (!settingMap.conversion_rate) settingRows.push({ key: 'conversion_rate', value: 0.14 });
    if (!settingMap.chat_ids) settingRows.push({ key: 'chat_ids', value: '' });
    if (!settingMap.telegram_bot_token) settingRows.push({ key: 'telegram_bot_token', value: '' });
    writeRows_('settings', settingRows);
    ensureSpreadsheetEditTrigger_();

    return {
      ok: true,
      appVersion: APP.VERSION,
      spreadsheetId: ss.getId(),
      spreadsheetName: ss.getName(),
      sheets: Object.keys(TABLES).map(function (k) { return TABLES[k].name; })
    };
  });
}

function healthCheck_(user) {
  var ss = getSpreadsheet_();
  var state = getState_();
  return {
    ok: true,
    version: APP.VERSION,
    user: user ? user.name : '',
    spreadsheetId: ss.getId(),
    spreadsheetName: ss.getName(),
    counts: {
      projects: state.projects.length,
      overheads: state.overheads.length,
      contracts: state.contracts.length
    }
  };
}

function testSpreadsheetConnection() {
  return withRequestContext_('testSpreadsheetConnection', function () {
    var ss = getSpreadsheet_();
    var result = { ok: true, id: ss.getId(), name: ss.getName(), sheets: ss.getSheets().map(function (s) { return s.getName(); }) };
    Logger.log(JSON.stringify(result));
    return result;
  });
}

function testState() {
  return withRequestContext_('testState', function () {
    var state = getState_();
    Logger.log(JSON.stringify(state).slice(0, 1500));
    return state;
  });
}

function testDispatchState() {
  return withRequestContext_('testDispatchState', function () {
    var result = dispatch('GET', '/api/state', null, 'oleksandr');
    Logger.log('dispatch GET /api/state => OK, projects=' + result.projects.length + ', json_len=' + JSON.stringify(result).length);
    return result;
  });
}

function testLogin() {
  return withRequestContext_('testLogin', function () {
    var u = login('oleksandr', '1952');
    Logger.log('login => OK code="' + u.code + '", name="' + u.name + '"');
    return u;
  });
}

function testDispatch() {
  return withRequestContext_('testDispatch', function () {
    var created = dispatch('POST', '/api/projects', { name: 'Тест-перевірка' }, 'Oleksandr');
    Logger.log('POST projects => OK id=' + created.id);
    dispatch('DELETE', '/api/projects/' + created.id, {}, 'Oleksandr');
    Logger.log('DELETE => OK');
    return true;
  });
}
