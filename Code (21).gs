/**
 * نَسَق — خادم استقبال الطلبات (Google Apps Script)
 * يُربط بجدول Google Sheets، وينشئ ورقتَي «الطلبات» و«الإعدادات» تلقائياً.
 */

// ===== غيّري رمز الدخول إلى لوحة الإدارة =====
const ADMIN_PIN = '1234';

const SHEET_REQ = 'الطلبات';
const SHEET_SET = 'الإعدادات';
const NEW = 'جديد';
const DONE = 'تمّ التواصل';
const TZ = 'Asia/Riyadh';
const HEADERS = ['المعرّف', 'تاريخ الطلب', 'الاسم', 'الجوال', 'الفكرة', 'الحالة'];

// ===== نقطة الدخول =====
function doPost(e) {
  let out;
  try {
    const body = JSON.parse((e && e.postData && e.postData.contents) || '{}');
    out = handle_(body);
  } catch (err) {
    out = { ok: false, error: 'حدث خطأ في الخادم، أعد المحاولة.' };
  }
  return json_(out);
}

function doGet() {
  return json_(status_());
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function handle_(b) {
  switch (b.action) {
    case 'status': return status_();
    case 'submit': return submit_(b);
    case 'admin_list': return auth_(b) ? adminList_() : deny_();
    case 'admin_update': return auth_(b) ? adminUpdate_(b) : deny_();
    case 'admin_delete': return auth_(b) ? adminDelete_(b) : deny_();
    case 'admin_settings': return auth_(b) ? adminSettings_(b) : deny_();
    default: return { ok: false, error: 'طلب غير معروف.' };
  }
}

function auth_(b) { return String(b.pin || '') === String(ADMIN_PIN); }
function deny_() { return { ok: false, error: 'رمز الدخول غير صحيح.' }; }

// ===== الأوراق =====
function reqSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_REQ);
  if (!sh) {
    sh = ss.insertSheet(SHEET_REQ);
    sh.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold');
    sh.setFrozenRows(1);
    sh.setRightToLeft(true);
  }
  return sh;
}

function setSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_SET);
  if (!sh) {
    sh = ss.insertSheet(SHEET_SET);
    sh.getRange(1, 1, 2, 2).setValues([['الحد الأقصى للطلبات الجديدة', 5], ['إيقاف الاستقبال يدوياً', false]]);
    sh.setRightToLeft(true);
  }
  return sh;
}

function settings_() {
  const v = setSheet_().getRange(1, 2, 2, 1).getValues();
  const limit = parseInt(v[0][0], 10);
  const closed = v[1][0] === true || String(v[1][0]).toUpperCase() === 'TRUE';
  return { limit: limit >= 1 ? limit : 5, closed: closed };
}

function rows_() {
  const sh = reqSheet_();
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, HEADERS.length).getDisplayValues().map(function (r, i) {
    return { row: i + 2, id: r[0], date: r[1], name: r[2], phone: r[3], idea: r[4], status: r[5] };
  });
}

function status_() {
  const s = settings_();
  const newCount = rows_().filter(function (r) { return r.status === NEW; }).length;
  return {
    ok: true,
    open: !s.closed && newCount < s.limit,
    closed: s.closed,
    limit: s.limit,
    newCount: newCount,
    remaining: Math.max(0, s.limit - newCount)
  };
}

// ===== استقبال طلب جديد =====
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '').trim().slice(0, max);
}

function normPhone_(s) {
  var p = String(s || '')
    .replace(/[٠-٩]/g, function (d) { return '٠١٢٣٤٥٦٧٨٩'.indexOf(d); })
    .replace(/[^\d+]/g, '');
  if (p.indexOf('+966') === 0) p = '0' + p.slice(4);
  else if (p.indexOf('00966') === 0) p = '0' + p.slice(5);
  else if (p.indexOf('966') === 0) p = '0' + p.slice(3);
  else if (/^5\d{8}$/.test(p)) p = '0' + p;
  return p;
}

function submit_(b) {
  const name = clean_(b.name, 60);
  const phone = normPhone_(b.phone);
  const idea = clean_(b.idea, 500);
  if (name.length < 2) return { ok: false, error: 'اكتب اسمك.' };
  if (!/^05\d{8}$/.test(phone)) return { ok: false, error: 'رقم الجوال غير صحيح.' };
  if (idea.length < 10) return { ok: false, error: 'اكتب فكرة الموقع باختصار.' };

  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    if (!status_().open) return { ok: false, full: true };
    const sh = reqSheet_();
    const row = sh.getLastRow() + 1;
    const id = Utilities.getUuid().slice(0, 8);
    const date = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm');
    sh.getRange(row, 1, 1, HEADERS.length)
      .setNumberFormat('@')
      .setValues([[id, date, name, phone, idea, NEW]]);
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

// ===== لوحة الإدارة =====
function adminList_() {
  const st = status_();
  st.items = rows_().reverse().map(function (r) {
    return { id: r.id, date: r.date, name: r.name, phone: r.phone, idea: r.idea, status: r.status };
  });
  return st;
}

function findRow_(id) {
  const r = rows_().filter(function (x) { return x.id === String(id); })[0];
  return r ? r.row : 0;
}

function adminUpdate_(b) {
  const status = b.status === DONE ? DONE : NEW;
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const row = findRow_(b.id);
    if (!row) return { ok: false, error: 'لم يُعثر على الطلب.' };
    reqSheet_().getRange(row, 6).setValue(status);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function adminDelete_(b) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const row = findRow_(b.id);
    if (!row) return { ok: false, error: 'لم يُعثر على الطلب.' };
    reqSheet_().deleteRow(row);
    return { ok: true };
  } finally { lock.releaseLock(); }
}

function adminSettings_(b) {
  const sh = setSheet_();
  if (b.limit !== undefined) {
    const n = parseInt(b.limit, 10);
    if (!(n >= 1 && n <= 100)) return { ok: false, error: 'الحد الأقصى يجب أن يكون من 1 إلى 100.' };
    sh.getRange(1, 2).setValue(n);
  }
  if (b.closed !== undefined) sh.getRange(2, 2).setValue(b.closed === true);
  return { ok: true };
}

// ===== شغّليها مرة واحدة من المحرّر لإنشاء الأوراق =====
function setup() {
  reqSheet_();
  setSheet_();
}
