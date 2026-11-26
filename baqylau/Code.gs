/**
 * Бақылау тесті — нәтижелерді Google Sheets-ке жинау (Білім Арена · Шың 140)
 * Орнату: Google Кесте → Кеңейтімдер → Apps Script → осы кодты қойыңыз →
 * TEACHER_KEY-ді өзгертіңіз → Іске қосу (Deploy) → Жаңа іске қосу → Веб-қосымша →
 * «Орындаушы: Мен», «Кімге қолжетімді: Барлығы» → сілтемені (…/exec) көшіріңіз.
 */
const TEACHER_KEY = 'OZGERTINIZ-2026';   // ← өз құпия кілтіңізді жазыңыз (мұғалім панеліне енгізесіз)
const SHEET_NAME = 'Бақылау тесті';
const HEAD = ['Уақыт', 'ID', 'Аты-жөні', 'Сынып', 'Деңгей', 'Пән', 'Балл', 'Сұрақ саны', 'Баға', 'Күйі', 'Себебі', 'Секунд', 'Жауаптар'];

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) { sh = ss.insertSheet(SHEET_NAME); sh.appendRow(HEAD); sh.setFrozenRows(1); sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold'); }
  return sh;
}
function clean_(x) {
  let s = String(x == null ? '' : x).slice(0, 200);
  if (/^[=+\-@]/.test(s)) s = "'" + s;   // формула енгізуден қорғау
  return s;
}
function doPost(e) {
  const d0 = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  if (d0.a === 'live' || d0.a === 'room' || d0.a === 'consent' || d0.a === 'mailreg') return tjPost_(d0);   // TelefonJarys.gs
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    const d = JSON.parse(e.postData.contents);
    sheet_().appendRow([
      new Date(), clean_(d.id), clean_(d.name), clean_(d.cls), clean_(d.lv), clean_(d.subj),
      d.score === undefined || d.score === '' ? '' : Number(d.score), Number(d.total) || '',
      clean_(d.grade), clean_(d.status), clean_(d.reason), Number(d.sec) || '', clean_(d.detail)
    ]);
  } finally { lock.releaseLock(); }
  return ContentService.createTextOutput('ok');
}
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (p.a) return tjGet_(p);                                                     // TelefonJarys.gs
  if (p.key !== TEACHER_KEY) return out_({ ok: false, error: 'Кілт қате' });
  const since = Number(p.since || 0);
  const v = sheet_().getDataRange().getValues(); v.shift();
  const rows = [];
  v.forEach(r => { const t = new Date(r[0]).getTime(); if (t && t >= since) rows.push([t].concat(r.slice(1).map(x => x instanceof Date ? x.getTime() : x))); });
  return out_({ ok: true, n: v.length, rows: rows });
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
