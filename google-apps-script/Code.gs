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

/* ── 導師座談記錄單所在的 Google 雲端硬碟資料夾 ID ──
   從資料夾網址取得：https://drive.google.com/drive/folders/【這一段就是 ID】
   設定後在編輯器選擇 fillMeetingLinks 執行一次，系統會依檔名中的姓名
   自動把連結填入 teachers 工作表第 4 欄。 */
var MEETING_FOLDER_ID = '';

var TEACHER_SHEET = 'teachers';
var REC_SHEET = 'assessments';
/* 程式內部使用的欄位鍵值（依「欄位位置」對應，工作表第 1 列標題可自由改名） */
var REC_HEADERS = [
  'id', 'period', 'name', 'unit',
  'self1', 'self2', 'self3', 'self4', 'self_total',
  'chief1', 'chief2', 'chief3', 'chief4', 'chief_total', 'chief_feedback', 'satisfaction',
  'dean1', 'dean2', 'dean3', 'dean4', 'dean_total', 'dean_feedback',
  'status', 'teacher_signed_at', 'chief_signed_at', 'dean_signed_at',
];
/* 自動建立工作表時顯示的中文標題（僅供閱讀，順序需與 REC_HEADERS 一致） */
var TEACHER_LABELS = ['姓名', '單位', '停用請填N', '人事號', '座談記錄單連結'];
/* teachers 工作表欄位位置（0 起算） */
var T_NAME = 0, T_UNIT = 1, T_ACTIVE = 2, T_STAFF = 3, T_LINK = 4;
var REC_LABELS = [
  '編號', '期別', '姓名', '單位',
  '自評1', '自評2', '自評3', '自評4', '自評總分',
  '主任評分1', '主任評分2', '主任評分3', '主任評分4', '主任總分', '主任回饋', '滿意度',
  '部長評分1', '部長評分2', '部長評分3', '部長評分4', '部長總分', '部長回饋',
  '狀態', '導師簽核日', '主任簽核日', '部長簽核日',
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
      /* 座談記錄單僅供主任與部長調閱，不回傳給導師 */
      var t = findTeacher(req.staffId || req.name);
      return t
        ? { ok: true, name: t.name, unit: t.unit, staffId: t.staffId }
        : { ok: false };
    case 'getRecords':
      return { records: scopedRecords(req) };
    case 'saveRecord':
      if (!canSave(req)) return { error: '未授權的存取' };
      var lock = LockService.getScriptLock();
      lock.waitLock(10000);
      try { saveRecord(req.record); } finally { lock.releaseLock(); }
      return { ok: true, records: scopedRecords(req) };
    default:
      return { error: 'unknown action: ' + req.action };
  }
}

/* 依身分限定可讀取的資料：導師只拿得到自己那筆（且不含記錄單連結），
   主管須通過密碼驗證，才能取得全部資料與座談記錄單連結 */
function scopedRecords(req) {
  if (req.role === 'chief' || req.role === 'dean') {
    if (PASSWORDS[req.role] !== String(req.password)) throw new Error('密碼驗證失敗');
    return getRecords(null, true);
  }
  if (req.role === 'teacher' && req.name) return getRecords(String(req.name).trim(), false);
  return [];
}

function canSave(req) {
  if (req.role === 'chief' || req.role === 'dean') {
    return PASSWORDS[req.role] === String(req.password);
  }
  if (req.role === 'teacher') {
    return !!(req.record && String(req.record.name).trim() === String(req.name).trim());
  }
  return false;
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
  getSheet(TEACHER_SHEET, TEACHER_LABELS);
  getSheet(REC_SHEET, REC_LABELS);
}

/* 以人事號查詢導師（找不到時亦接受輸入姓名，方便人事號尚未建檔時使用） */
function findTeacher(key) {
  var sh = getSheet(TEACHER_SHEET, TEACHER_LABELS);
  var rows = sh.getDataRange().getValues().slice(1);
  var k = String(key).trim();
  if (!k) return null;
  var byName = null;
  for (var i = 0; i < rows.length; i++) {
    var r = rows[i];
    if (String(r[T_ACTIVE]).trim().toUpperCase() === 'N') continue;
    var hit = {
      name: String(r[T_NAME]).trim(),
      unit: String(r[T_UNIT]).trim(),
      staffId: String(r[T_STAFF] || '').trim(),
      meetingLink: String(r[T_LINK] || '').trim(),
    };
    if (hit.staffId && hit.staffId.toUpperCase() === k.toUpperCase()) return hit;
    if (!byName && hit.name === k) byName = hit;
  }
  return byName;
}

/* 姓名 → 座談記錄單連結 對照表 */
function meetingLinkMap() {
  var sh = getSheet(TEACHER_SHEET, TEACHER_LABELS);
  var rows = sh.getDataRange().getValues().slice(1);
  var map = {};
  for (var i = 0; i < rows.length; i++) {
    var n = String(rows[i][T_NAME]).trim();
    if (n) map[n] = String(rows[i][T_LINK] || '').trim();
  }
  return map;
}

/* ── 一次性工具：掃描 Drive 資料夾，依檔名中的姓名自動填入連結 ──
   在 Apps Script 編輯器上方選擇本函式後按「執行」，完成後查看執行紀錄。 */
function fillMeetingLinks() {
  if (!MEETING_FOLDER_ID) throw new Error('請先在程式最上方設定 MEETING_FOLDER_ID');
  var sh = getSheet(TEACHER_SHEET, TEACHER_LABELS);
  var last = sh.getLastRow();
  if (last < 2) throw new Error('teachers 工作表尚無導師名單');

  var files = [];
  var it = DriveApp.getFolderById(MEETING_FOLDER_ID).getFiles();
  while (it.hasNext()) {
    var f = it.next();
    files.push({ name: f.getName(), url: f.getUrl() });
  }

  var names = sh.getRange(2, T_NAME + 1, last - 1, 1).getValues();
  var linkRange = sh.getRange(2, T_LINK + 1, last - 1, 1);
  var existing = linkRange.getValues();
  var out = [], matched = 0, missing = [];
  for (var i = 0; i < names.length; i++) {
    var n = String(names[i][0]).trim();
    var prev = String(existing[i][0] || '').trim();
    if (!n) { out.push([prev]); continue; }
    var hit = '';
    for (var j = 0; j < files.length; j++) {
      if (files[j].name.indexOf(n) >= 0) { hit = files[j].url; break; }
    }
    /* 比對不到時保留原本欄位內容，避免覆蓋手動填入的連結 */
    out.push([hit || prev]);
    if (hit) matched++; else missing.push(n);
  }
  linkRange.setValues(out);
  Logger.log('資料夾檔案數：' + files.length + '，成功對應：' + matched +
    '，查無檔案：' + (missing.length ? missing.join('、') : '無'));
}

/* ── 讀取評核紀錄，轉為前端使用的資料格式
   onlyName 不為空時僅回傳該導師；withLinks 為 false 時不附座談記錄單連結 ── */
function getRecords(onlyName, withLinks) {
  var sh = getSheet(REC_SHEET, REC_LABELS);
  var rows = sh.getDataRange().getValues().slice(1);
  var links = withLinks ? meetingLinkMap() : {};
  var records = [];
  for (var i = 0; i < rows.length; i++) {
    var o = {};
    for (var c = 0; c < REC_HEADERS.length; c++) o[REC_HEADERS[c]] = rows[i][c];
    if (!o.id) continue;
    if (onlyName && String(o.name).trim() !== onlyName) continue;
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
      meetingLink: links[String(o.name).trim()] || '',
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
  var sh = getSheet(REC_SHEET, REC_LABELS);
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
  if (v && typeof v.getTime === 'function') {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return (v === '' || v === null || v === undefined) ? null : String(v);
}
