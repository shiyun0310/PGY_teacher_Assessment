/**
 * 一般醫學長期導師評核平台 — Google Sheets 後端（Google Apps Script）
 * 部署步驟請見專案 README.md 的「Google Sheets 資料庫設定」章節。
 *
 * 工作表（第一次執行或開啟 API 網址時會自動建立）：
 *   teachers    — 導師名單：name（姓名）/ unit（單位）/ active（填 N 表示停用，其他皆視為啟用）
 *   assessments — 評核紀錄（由系統自動寫入，請勿變動欄位順序）
 */

/* ── 設定：登入密碼與自動簽名人。修改後需「部署 → 管理部署作業 → 編輯 → 新版本」才會生效 ── */
var PASSWORDS = { chief: '940963', dean: '010101' };
var SIGNERS   = { chief: '郭妍伶', dean: '蘇慧真' };

var TEACHER_SHEET = 'teachers';
var REC_SHEET = 'assessments';
var REC_HEADERS = [
  'id', 'period', 'name', 'unit',
  'self1', 'self2', 'self3', 'self4', 'self_total',
  'chief1', 'chief2', 'chief3', 'chief4', 'chief_total', 'chief_feedback', 'satisfaction',
  'dean1', 'dean2', 'dean3', 'dean4', 'dean_total', 'dean_feedback',
  'status', 'teacher_signed_at', 'chief_signed_at', 'dean_signed_at',
];

function doGet() {
  ensureSheets();
  return json({ ok: true, message: 'PGY 導師評核平台 API 運作中' });
}

function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    return json(handle(req));
  } catch (err) {
    return json({ error: String(err) });
  }
}

function handle(req) {
  switch (req.action) {
    case 'login':
      return { ok: PASSWORDS[req.role] === String(req.password) };
    case 'teacherLogin':
      var t = findTeacher(req.name);
      return t ? { ok: true, unit: t.unit } : { ok: false };
    case 'getRecords':
      return { records: getRecords() };
    case 'saveRecord':
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try { saveRecord(req.record); } finally { lock.releaseLock(); }
      return { ok: true, records: getRecords() };
    default:
      return { error: 'unknown action: ' + req.action };
  }
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function getSheet(name, headers) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, headers.length).setValues([headers]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}

function ensureSheets() {
  getSheet(TEACHER_SHEET, ['name', 'unit', 'active']);
  getSheet(REC_SHEET, REC_HEADERS);
}

function findTeacher(name) {
  var sh = getSheet(TEACHER_SHEET, ['name', 'unit', 'active']);
  var rows = sh.getDataRange().getValues().slice(1);
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[0]).trim() === String(name).trim() &&
        String(r[2]).trim().toUpperCase() !== 'N') {
      return { name: String(r[0]).trim(), unit: String(r[1]).trim() };
    }
  }
  return null;
}

/* ── 讀取所有評核紀錄，轉為前端使用的資料格式 ── */
function getRecords() {
  var sh = getSheet(REC_SHEET, REC_HEADERS);
  var rows = sh.getDataRange().getValues().slice(1);
  var records = [];
  for (var i = 0; i < rows.length; i++) {
    var o = {};
    for (var c = 0; c < REC_HEADERS.length; c++) o[REC_HEADERS[c]] = rows[i][c];
    if (!o.id) continue;
    var chiefDate = fmtDate(o.chief_signed_at);
    var deanDate = fmtDate(o.dean_signed_at);
    records.push({
      id: String(o.id),
      period: String(o.period),
      name: String(o.name),
      unit: String(o.unit),
      selfScores: [Number(o.self1), Number(o.self2), Number(o.self3), Number(o.self4)],
      selfTotal: Number(o.self_total),
      chiefScores: o.chief_total === '' ? null
        : [Number(o.chief1), Number(o.chief2), Number(o.chief3), Number(o.chief4)],
      chiefTotal: numOrNull(o.chief_total),
      chiefFeedback: String(o.chief_feedback || ''),
      satisfaction: numOrNull(o.satisfaction),
      deanScores: o.dean_total === '' ? null
        : [Number(o.dean1), Number(o.dean2), Number(o.dean3), Number(o.dean4)],
      deanTotal: numOrNull(o.dean_total),
      deanFeedback: String(o.dean_feedback || ''),
      status: String(o.status),
      submittedAt: fmtDate(o.teacher_signed_at),
      sigs: {
        teacher: { name: String(o.name), date: fmtDate(o.teacher_signed_at) },
        chief: chiefDate ? { name: SIGNERS.chief, date: chiefDate } : null,
        dean: deanDate ? { name: SIGNERS.dean, date: deanDate } : null,
      },
    });
  }
  return records;
}

/* ── 新增或更新一筆評核紀錄（依 id 比對）── */
function saveRecord(rec) {
  var sh = getSheet(REC_SHEET, REC_HEADERS);
  var row = [
    rec.id, rec.period, rec.name, rec.unit,
    rec.selfScores[0], rec.selfScores[1], rec.selfScores[2], rec.selfScores[3], rec.selfTotal,
  ]
    .concat(rec.chiefScores || ['', '', '', ''])
    .concat([nv(rec.chiefTotal), nv(rec.chiefFeedback), nv(rec.satisfaction)])
    .concat(rec.deanScores || ['', '', '', ''])
    .concat([nv(rec.deanTotal), nv(rec.deanFeedback)])
    .concat([
      rec.status,
      sigDate(rec, 'teacher'), sigDate(rec, 'chief'), sigDate(rec, 'dean'),
    ]);
  var last = sh.getLastRow();
  if (last > 1) {
    var ids = sh.getRange(2, 1, last - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(rec.id)) {
        sh.getRange(i + 2, 1, 1, REC_HEADERS.length).setValues([row]);
        return;
      }
    }
  }
  sh.appendRow(row);
}

function sigDate(rec, who) {
  return (rec.sigs && rec.sigs[who] && rec.sigs[who].date) ? rec.sigs[who].date : '';
}

function nv(v) { return (v === null || v === undefined) ? '' : v; }

function numOrNull(v) { return (v === '' || v === null || v === undefined) ? null : Number(v); }

function fmtDate(v) {
  if (v instanceof Date) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return (v === '' || v === null || v === undefined) ? null : String(v);
}
