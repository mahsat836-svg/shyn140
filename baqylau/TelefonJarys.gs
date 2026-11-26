/**
 * Білім Арена · ТЕЛЕФОН ЖАРЫСЫ және АТА-АНА КЕЛІСІМІ (Шың 140)
 *
 * Code.gs-пен бір жобада тұрады (бөлек файл). Code.gs осы файлдағы
 * tjPost_ / tjGet_ функцияларын шақырады.
 *
 * Орнату (бір рет):
 *   1) Google Кесте → Кеңейтімдер → Apps Script → «Файлдар» жанындағы «+» → «Сценарий» → атауы: TelefonJarys.
 *      Осы кодты түгел қойып, сақтаңыз. Code.gs-ті де репозиторийдегі жаңа нұсқамен ауыстырыңыз.
 *   2) ⚙️ Жоба баптаулары → «Сценарий сипаттары» → сипат қосу: T_PIN = өз PIN-іңіз (тек сіз білесіз).
 *   3) Іске қосу (Deploy) → «Іске қосуларды басқару» → ✏️ → Нұсқа: «Жаңа нұсқа» → Іске қосу.
 *      Сілтеме (…/exec) өзгермейді.
 *
 * Парақтар (өздері жасалады):
 *   «Телефон жарысы»       — жарыс кезіндегі тірі нәтиже (бір оқушыға бір жол)
 *   «Ата-ана келісімдері»  — ата-аналардың онлайн келісімі (қолтаңбасымен)
 *
 * Құпиялылық: ата-ана телефоны мен келісімдерді тек T_PIN білетін мұғалім көреді.
 */

function tjPin_() { return String(PropertiesService.getScriptProperties().getProperty('T_PIN') || ''); }
function tjOut_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function tjNorm_(s) { return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ').trim(); }
function tjCls_(s) { return tjNorm_(s).replace(/[\s\-«»\x22\x27]/g, ''); }          // «5 А», «5А», «5-а» — бірдей
function tjClean_(x, n) { let s = String(x == null ? '' : x).slice(0, n || 200); if (/^[=+\-@]/.test(s)) s = "'" + s; return s; }
function tjSheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); sh.getRange(1, 1, 1, head.length).setFontWeight('bold'); }
  return sh;
}
const TJ_LIVE_HEAD = ['Жаңартылды', 'Жарыс', 'Сынып', 'Аты-жөні', 'Дұрыс', 'Жауап берді', 'Барлығы', 'Күйі', 'Секунд', 'Әлсіз тақырып', 'Шығу саны'];
const TJ_CONS_HEAD = ['Уақыты', 'Оқушы', 'Сынып', 'Ата-анасы', 'Телефон', 'Тест/нәтиже', 'WhatsApp', 'Фото/бейне', 'Жариялау', 'Тілі', 'Қолтаңба', 'Күйі'];
const tjLive_ = () => tjSheet_('Телефон жарысы', TJ_LIVE_HEAD);
const tjCons_ = () => tjSheet_('Ата-ана келісімдері', TJ_CONS_HEAD);

/* ---------------- POST ---------------- */
function tjPost_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try {
    if (d.a === 'room') {                       // мұғалім жарыс ашты → оқушыларға сынып тізімін 6 сағат көрсетеміз
      if (!tjPin_() || String(d.pin) !== tjPin_()) return tjOut_({ ok: false });
      if (/^\d{4}$/.test(String(d.r))) CacheService.getScriptCache().put('room_' + d.r, String(d.c || ''), 21600);
      return tjOut_({ ok: true });
    }
    if (d.a === 'consent') {
      if (!d.child || !d.cls || !d.parent) return tjOut_({ ok: false });
      const sig = String(d.sig || '');
      tjCons_().appendRow([new Date(), tjClean_(d.child, 80), tjClean_(d.cls, 20), tjClean_(d.parent, 80), tjClean_(d.phone, 20),
        d.c1 ? 'иә' : 'жоқ', d.c2 ? 'иә' : 'жоқ', d.c3 ? 'иә' : 'жоқ', d.c4 ? 'иә' : 'жоқ', tjClean_(d.lang, 5),
        sig.indexOf('data:image/png;base64,') === 0 && sig.length < 45000 ? sig : '', d.revoke ? 'қайтарып алды' : 'берді']);
      return tjOut_({ ok: true });
    }
    if (d.a === 'live') {
      if (!/^\d{4}$/.test(String(d.r)) || !d.name) return tjOut_({ ok: false });
      const sh = tjLive_(), v = sh.getDataRange().getValues();
      const key = d.r + '|' + tjCls_(d.c) + '|' + tjNorm_(d.name);
      let row = -1;
      for (let i = v.length - 1; i >= 1; i--) {
        if (String(v[i][1]) + '|' + tjCls_(v[i][2]) + '|' + tjNorm_(v[i][3]) === key) {
          if (v[i][7] === 'done' || v[i][7] === 'out') return tjOut_({ ok: true, locked: true });   // аяқталған нәтиже өзгермейді
          row = i + 1; break;
        }
      }
      const st = d.status === 'done' || d.status === 'out' ? d.status : 'test';
      const vals = [new Date(), String(d.r), tjClean_(d.c, 20), tjClean_(d.name, 60), +d.correct || 0, +d.answered || 0, +d.total || 0, st, +d.secs || 0, tjClean_(d.weak, 80), +d.leaves || 0];
      if (row > 0) sh.getRange(row, 1, 1, vals.length).setValues([vals]); else sh.appendRow(vals);
      return tjOut_({ ok: true });
    }
    return tjOut_({ ok: false });
  } finally { lock.releaseLock(); }
}

/* ---------------- GET ---------------- */
function tjConsents_() {                         // әр оқушы бойынша соңғы келісім
  const v = tjCons_().getDataRange().getValues(); v.shift();
  const m = {};
  v.forEach(r => { m[tjCls_(r[2]) + '|' + tjNorm_(r[1])] = r; });
  return Object.keys(m).map(k => m[k]);
}
function tjGet_(p) {
  if (p.a === 'ping') return tjOut_({ ok: true, app: 'Телефон жарысы' });
  if (p.a === 'roster') {                        // тек ашық жарыс коды болса, сол сыныптың аттары
    const c = CacheService.getScriptCache().get('room_' + p.r);
    if (c === null) return tjOut_({ ok: false, names: [] });
    const names = tjConsents_().filter(r => tjCls_(r[2]) === tjCls_(c) && r[11] !== 'қайтарып алды' && r[5] === 'иә').map(r => String(r[1]));
    return tjOut_({ ok: true, cls: c, names: names.sort() });
  }
  if (!tjPin_() || String(p.pin) !== tjPin_()) return tjOut_({ ok: false, error: 'pin' });
  if (p.a === 'live') {
    const v = tjLive_().getDataRange().getValues(); v.shift();
    const rows = v.filter(r => String(r[1]) === String(p.r)).map(r => ({ c: r[2], name: r[3], correct: +r[4], answered: +r[5], total: +r[6], status: r[7], secs: +r[8], weak: r[9], leaves: +r[10] }));
    return tjOut_({ ok: true, rows: rows });
  }
  if (p.a === 'consents') {
    const rows = tjConsents_().map(r => ({ t: new Date(r[0]).getTime(), child: r[1], cls: r[2], parent: r[3], phone: String(r[4]), c1: r[5] === 'иә', c2: r[6] === 'иә', c3: r[7] === 'иә', c4: r[8] === 'иә', lang: r[9], sig: p.sig ? r[10] : (r[10] ? 1 : 0), revoked: r[11] === 'қайтарып алды' }));
    return tjOut_({ ok: true, rows: rows });
  }
  return tjOut_({ ok: false });
}
