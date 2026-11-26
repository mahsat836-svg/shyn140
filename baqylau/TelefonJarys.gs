/**
 * Білім Арена · ТЕЛЕФОН ЖАРЫСЫ және АТА-АНА КЕЛІСІМІ (Шың 140)
 *
 * Code.gs-пен бір жобада тұрады (бөлек файл). Code.gs осы файлдағы
 * tjPost_ / tjGet_ функцияларын шақырады.
 *
 * Не істейді:
 *   • Ата-ана онлайн келісім берсе — «Ата-ана келісімдері» парағына жазады және
 *     қолтаңбасы бар ресми PDF құжатты TJ_EMAIL поштасына жібереді.
 *   • Мұғалім «📧 PDF тізілім» басса — сыныптың барлық келісімі бір PDF болып поштаға келеді.
 *   • Телефон жарысының тірі нәтижесін «Телефон жарысы» парағына жазады.
 *
 * Орнату (бір рет):
 *   1) Apps Script → «+» → «Сценарий» → TelefonJarys → осы кодты қойыңыз. Code.gs-ті репозиторийдегі нұсқамен ауыстырыңыз.
 *   2) ⚙️ Жоба баптаулары → «Сценарий сипаттары» → T_PIN = өз PIN-іңіз.
 *   3) Жоғарыдан tjAuthorize функциясын таңдап «▶ Орындау» → «Рұқсат беру» (пошта жіберу үшін бір рет).
 *   4) Іске қосу → «Іске қосуларды басқару» → ✏️ → «Жаңа нұсқа» → Іске қосу.
 */

const TJ_EMAIL = 'jmahsat@inbox.ru';
const TJ_SCHOOL = 'Ө. Жәнібеков атындағы жалпы білім беретін мектеп';
const TJ_PLACE = 'Сауран ауданы, Түркістан облысы';
const TJ_TEACHER = 'Юлдашев Кадыржан Фазылович';

function tjAuthorize() { Logger.log('Пошта квотасы: ' + MailApp.getRemainingDailyQuota()); Logger.log('PIN бар ма: ' + (tjPin_() ? 'иә' : 'жоқ')); }
function tjPin_() { return String(PropertiesService.getScriptProperties().getProperty('T_PIN') || ''); }
function tjOut_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function tjNorm_(s) { return String(s || '').toLowerCase().replace(/ё/g, 'е').replace(/\s+/g, ' ').trim(); }
function tjCls_(s) { return tjNorm_(s).replace(/[\s\-«»\x22\x27]/g, ''); }
function tjClean_(x, n) { let s = String(x == null ? '' : x).slice(0, n || 200); if (/^[=+\-@]/.test(s)) s = "'" + s; return s; }
function tjEsc_(s) { return String(s == null ? '' : s).replace(/[&<>\x22\x27]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]); }
function tjDate_(d) { return Utilities.formatDate(d, 'Asia/Tashkent', 'dd.MM.yyyy, HH:mm'); }
function tjSheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) { sh = ss.insertSheet(name); sh.appendRow(head); sh.setFrozenRows(1); sh.getRange(1, 1, 1, head.length).setFontWeight('bold'); }
  return sh;
}
const TJ_LIVE_HEAD = ['Жаңартылды', 'Жарыс', 'Сынып', 'Аты-жөні', 'Дұрыс', 'Жауап берді', 'Барлығы', 'Күйі', 'Секунд', 'Әлсіз тақырып', 'Шығу саны'];
const TJ_CONS_HEAD = ['Уақыты', 'Оқушы', 'Сынып', 'Ата-анасы', 'Телефон', 'Тест/нәтиже', 'WhatsApp', 'Фото/бейне', 'Жариялау', 'Тілі', 'Қолтаңба', 'Күйі', 'Құжат №', 'Пошта'];
const tjLive_ = () => tjSheet_('Телефон жарысы', TJ_LIVE_HEAD);
const tjCons_ = () => tjSheet_('Ата-ана келісімдері', TJ_CONS_HEAD);

/* ---------------- PDF мәтіндері ---------------- */
const TJ_TX = {
  kk: { h: 'АТА-АНАНЫҢ (ЗАҢДЫ ӨКІЛДІҢ) КЕЛІСІМІ', sub: 'электронды нысан · «Білім Арена» / «Шың 140»', child: 'Оқушы', cls: 'Сыныбы', parent: 'Ата-анасы (заңды өкілі)', phone: 'Телефоны', yes: 'КЕЛІСЕМІН', no: 'КЕЛІСПЕЙМІН', iam: 'Мен осы баланың ата-анасы (заңды өкілі) екенімді растаймын.', sign: 'Қолы', date: 'Толтырылған уақыты', rev: 'КЕЛІСІМ ТОЛЫҚ ҚАЙТАРЫП АЛЫНДЫ', what: 'Келісім мазмұны', dec: 'Ата-ананың шешімі',
    c: ['Баламның «Білім Арена» мен «Шың 140» платформаларында тест пен жарыстарға қатысуына және нәтижесі мұғалімнің кестесінде сақталуына', 'Баламның нәтижесі туралы мұғалімнен WhatsApp арқылы хабарлама алуға', 'Сабақ, тест, жарыс және олимпиада кезінде баламды фото/бейнеге түсіруге (мектептің оқу мақсатында)', 'Сол фото/бейнені мектептің ресми парақшаларында (Instagram, YouTube) жариялауға'] },
  uz: { h: 'ОТА-ОНАНИНГ (ҚОНУНИЙ ВАКИЛНИНГ) РОЗИЛИГИ', sub: 'электрон шакл · «Билим Арена» / «Шың 140»', child: 'Ўқувчи', cls: 'Синфи', parent: 'Ота-онаси (қонуний вакили)', phone: 'Телефони', yes: 'РОЗИМАН', no: 'РОЗИ ЭМАСМАН', iam: 'Мен шу боланинг ота-онаси (қонуний вакили) эканимни тасдиқлайман.', sign: 'Имзо', date: 'Тўлдирилган вақти', rev: 'РОЗИЛИК ТЎЛИҚ ҚАЙТАРИБ ОЛИНДИ', what: 'Розилик мазмуни', dec: 'Ота-онанинг қарори',
    c: ['Фарзандимнинг «Билим Арена» ва «Шың 140» платформаларида тест ва мусобақаларда қатнашишига ҳамда натижаси ўқитувчининг жадвалида сақланишига', 'Фарзандимнинг натижаси ҳақида ўқитувчидан WhatsApp орқали хабар олишга', 'Дарс, тест, мусобақа ва олимпиада пайтида фарзандимни фото/видеога олишга (мактабнинг ўқув мақсадида)', 'Ўша фото/видеоларни мактабнинг расмий саҳифаларида (Instagram, YouTube) эълон қилишга'] }
};
const TJ_CSS = 'body{font-family:Arial,sans-serif;color:#111;font-size:12pt;margin:0}' +
  '.top{border-bottom:3px solid #0b2a5b;padding-bottom:8px;margin-bottom:14px}.top b{font-size:13pt;color:#0b2a5b}.top div{font-size:10pt;color:#444}' +
  'h1{font-size:15pt;text-align:center;margin:10px 0 2px;color:#0b2a5b}.sub{text-align:center;font-size:9.5pt;color:#555;margin-bottom:14px}' +
  'table{width:100%;border-collapse:collapse}td,th{border:1px solid #999;padding:6px 8px;vertical-align:top;font-size:11pt}th{background:#e8eef8;text-align:left}' +
  '.y{color:#0a7a3b;font-weight:bold;white-space:nowrap}.n{color:#b00020;font-weight:bold;white-space:nowrap}.no{font-size:9pt;color:#555}' +
  '.sig{height:70px}.foot{margin-top:18px;font-size:8.5pt;color:#555;border-top:1px solid #ccc;padding-top:6px}.rev{color:#b00020;font-weight:bold;text-align:center;font-size:13pt;margin:10px 0}';

function tjConsentHtml_(r) {
  const t = TJ_TX[r.lang === 'uz' ? 'uz' : 'kk'];
  const yn = v => v ? '<span class="y">✔ ' + t.yes + '</span>' : '<span class="n">✖ ' + t.no + '</span>';
  let h = '<html><head><meta charset="utf-8"><style>' + TJ_CSS + '</style></head><body>';
  h += '<div class="top"><b>' + tjEsc_(TJ_SCHOOL) + '</b><div>' + tjEsc_(TJ_PLACE) + ' · Құжат № ' + tjEsc_(r.no) + '</div></div>';
  h += '<h1>' + t.h + '</h1><div class="sub">' + t.sub + '</div>';
  if (r.revoke) h += '<div class="rev">' + t.rev + '</div>';
  h += '<table><tr><th style="width:38%">' + t.child + '</th><td><b>' + tjEsc_(r.child) + '</b></td></tr><tr><th>' + t.cls + '</th><td>' + tjEsc_(r.cls) + '</td></tr><tr><th>' + t.parent + '</th><td>' + tjEsc_(r.parent) + '</td></tr><tr><th>' + t.phone + '</th><td>' + tjEsc_(r.phone) + '</td></tr></table><br>';
  h += '<table><tr><th style="width:6%">№</th><th>' + t.what + '</th><th style="width:22%">' + t.dec + '</th></tr>';
  [r.c1, r.c2, r.c3, r.c4].forEach((v, i) => { h += '<tr><td>' + (i + 1) + '</td><td>' + t.c[i] + '</td><td>' + yn(v) + '</td></tr>'; });
  h += '</table><p>' + (r.revoke ? '' : '☑ ' + t.iam) + '</p>';
  h += '<table><tr><th style="width:38%">' + t.date + '</th><td>' + tjEsc_(tjDate_(r.when)) + ' (UTC+5)</td></tr><tr><th>' + t.sign + '</th><td>' + (r.sig ? '<img class="sig" src="' + r.sig + '">' : '—') + '</td></tr></table>';
  h += '<div class="foot">Құжат ата-ананың телефонынан «Білім Арена · Шың 140» онлайн нысаны арқылы толтырылып, саусақпен қол қойылды. Деректер тек мұғалімнің Google кестесінде сақталады, үшінші тұлғаға берілмейді. Келісімді кез келген уақытта өзгертуге немесе қайтарып алуға болады. Жауапты мұғалім: ' + tjEsc_(TJ_TEACHER) + '.</div>';
  return h + '</body></html>';
}
function tjPdf_(html, name) { return Utilities.newBlob(html, 'text/html', name + '.html').getAs('application/pdf').setName(name + '.pdf'); }
function tjSafe_(s) { return String(s || '').replace(/[\\\/:*?\x22<>|]+/g, ' ').trim().slice(0, 60); }

function tjMailConsent_(r) {
  const pdf = tjPdf_(tjConsentHtml_(r), 'Келісім ' + r.no + ' — ' + tjSafe_(r.child) + ' (' + tjSafe_(r.cls) + ')');
  const yn = v => v ? '✔' : '✖';
  const subj = (r.revoke ? '⛔ Келісім қайтарылды: ' : '✅ Ата-ана келісімі: ') + r.child + ' · ' + r.cls;
  const body = (r.revoke ? 'Ата-ана келісімін ТОЛЫҚ ҚАЙТАРЫП АЛДЫ.\n\n' : 'Жаңа онлайн келісім келді.\n\n') +
    'Оқушы: ' + r.child + ' (' + r.cls + ')\nАта-анасы: ' + r.parent + ', ' + r.phone + '\n' +
    'Тест ' + yn(r.c1) + ' · WhatsApp ' + yn(r.c2) + ' · Фото/бейне ' + yn(r.c3) + ' · Жариялау ' + yn(r.c4) + '\n' +
    'Уақыты: ' + tjDate_(r.when) + '\nҚұжат № ' + r.no + ' — PDF тіркелді.\n\nБілім Арена · Шың 140';
  MailApp.sendEmail({ to: TJ_EMAIL, subject: subj, body: body, attachments: [pdf], name: 'Білім Арена · Келісімдер' });
}

/* ---------------- POST ---------------- */
function tjPost_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  let rec = null, row = 0;
  try {
    if (d.a === 'room') {
      if (!tjPin_() || String(d.pin) !== tjPin_()) return tjOut_({ ok: false });
      if (/^\d{4}$/.test(String(d.r))) CacheService.getScriptCache().put('room_' + d.r, String(d.c || ''), 21600);
      return tjOut_({ ok: true });
    }
    if (d.a === 'consent') {
      if (!d.child || !d.cls || !d.parent) return tjOut_({ ok: false, error: 'fields' });
      const sig = String(d.sig || '');
      const sh = tjCons_();
      const no = Utilities.formatDate(new Date(), 'Asia/Tashkent', 'yyMMdd') + '-' + String(sh.getLastRow()).padStart(3, '0');
      rec = { when: new Date(), child: String(d.child).slice(0, 80), cls: String(d.cls).slice(0, 20), parent: String(d.parent).slice(0, 80), phone: String(d.phone || '').slice(0, 20),
        c1: !!d.c1, c2: !!d.c2, c3: !!d.c3, c4: !!d.c4, lang: d.lang === 'uz' ? 'uz' : 'kk', sig: sig.indexOf('data:image/png;base64,') === 0 && sig.length < 45000 ? sig : '', revoke: !!d.revoke, no: no };
      sh.appendRow([rec.when, tjClean_(rec.child), tjClean_(rec.cls), tjClean_(rec.parent), tjClean_(rec.phone), rec.c1 ? 'иә' : 'жоқ', rec.c2 ? 'иә' : 'жоқ', rec.c3 ? 'иә' : 'жоқ', rec.c4 ? 'иә' : 'жоқ', rec.lang, rec.sig, rec.revoke ? 'қайтарып алды' : 'берді', no, '…']);
      row = sh.getLastRow();
    } else if (d.a === 'live') {
      if (!/^\d{4}$/.test(String(d.r)) || !d.name) return tjOut_({ ok: false });
      const sh = tjLive_(), v = sh.getDataRange().getValues();
      const key = d.r + '|' + tjCls_(d.c) + '|' + tjNorm_(d.name);
      let rw = -1;
      for (let i = v.length - 1; i >= 1; i--) {
        if (String(v[i][1]) + '|' + tjCls_(v[i][2]) + '|' + tjNorm_(v[i][3]) === key) {
          if (v[i][7] === 'done' || v[i][7] === 'out') return tjOut_({ ok: true, locked: true });
          rw = i + 1; break;
        }
      }
      const st = d.status === 'done' || d.status === 'out' ? d.status : 'test';
      const vals = [new Date(), String(d.r), tjClean_(d.c, 20), tjClean_(d.name, 60), +d.correct || 0, +d.answered || 0, +d.total || 0, st, +d.secs || 0, tjClean_(d.weak, 80), +d.leaves || 0];
      if (rw > 0) sh.getRange(rw, 1, 1, vals.length).setValues([vals]); else sh.appendRow(vals);
      return tjOut_({ ok: true });
    } else return tjOut_({ ok: false });
  } finally { lock.releaseLock(); }
  let mailed = false, err = '';
  try { tjMailConsent_(rec); mailed = true; } catch (e) { err = String(e.message || e).slice(0, 120); }
  tjCons_().getRange(row, 14).setValue(mailed ? 'жіберілді ' + tjDate_(new Date()) : 'қате: ' + err);
  return tjOut_({ ok: true, no: rec.no, mailed: mailed });
}

/* ---------------- GET ---------------- */
function tjConsents_() {
  const v = tjCons_().getDataRange().getValues(); v.shift();
  const m = {};
  v.forEach(r => { m[tjCls_(r[2]) + '|' + tjNorm_(r[1])] = r; });
  return Object.keys(m).map(k => m[k]);
}
function tjRegisterHtml_(rows, cls) {
  const yn = v => v ? '<span class="y">✔</span>' : '<span class="n">✖</span>';
  let h = '<html><head><meta charset="utf-8"><style>' + TJ_CSS + 'td,th{font-size:9.5pt;padding:4px 5px}.sig{height:38px}</style></head><body>';
  h += '<div class="top"><b>' + tjEsc_(TJ_SCHOOL) + '</b><div>' + tjEsc_(TJ_PLACE) + '</div></div>';
  h += '<h1>АТА-АНА КЕЛІСІМДЕРІНІҢ ТІЗІЛІМІ</h1><div class="sub">' + tjEsc_(cls || 'Барлық сыныптар') + ' · ' + tjDate_(new Date()) + ' · «Білім Арена / Шың 140» электронды нысаны</div>';
  const act = rows.filter(r => r[11] !== 'қайтарып алды');
  h += '<p>Барлығы: <b>' + act.length + '</b> келісім · фото/бейнеге рұқсат: <b>' + act.filter(r => r[7] === 'иә').length + '</b> · WhatsApp: <b>' + act.filter(r => r[6] === 'иә').length + '</b>' + (rows.length > act.length ? ' · қайтарып алғандар: <b>' + (rows.length - act.length) + '</b>' : '') + '</p>';
  h += '<table><tr><th>№</th><th>Оқушы</th><th>Сынып</th><th>Ата-анасы, телефоны</th><th>Тест</th><th>WA</th><th>Фото</th><th>Жария</th><th>Күні</th><th>Қолы</th></tr>';
  rows.forEach((r, i) => {
    const rv = r[11] === 'қайтарып алды';
    h += '<tr' + (rv ? ' style="color:#999"' : '') + '><td>' + (i + 1) + '</td><td><b>' + tjEsc_(r[1]) + '</b>' + (rv ? '<br><span class="n">қайтарып алды</span>' : '') + '</td><td>' + tjEsc_(r[2]) + '</td><td>' + tjEsc_(r[3]) + '<br><span class="no">' + tjEsc_(r[4]) + '</span></td>' +
      [5, 6, 7, 8].map(k => '<td>' + yn(!rv && r[k] === 'иә') + '</td>').join('') + '<td>' + tjEsc_(tjDate_(new Date(r[0]))) + '<br><span class="no">№ ' + tjEsc_(r[12] || '') + '</span></td><td>' + (r[10] ? '<img class="sig" src="' + r[10] + '">' : '—') + '</td></tr>';
  });
  h += '</table><div class="foot">Әр жолдың толық құжаты (қолтаңбасымен) мұғалімнің поштасына жеке PDF болып келген. Жауапты мұғалім: ' + tjEsc_(TJ_TEACHER) + '. Сынып жетекшінің қолы: ____________________</div>';
  return h + '</body></html>';
}
function tjGet_(p) {
  if (p.a === 'ping') return tjOut_({ ok: true, app: 'Телефон жарысы' });
  if (p.a === 'roster') {
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
    const rows = tjConsents_().map(r => ({ t: new Date(r[0]).getTime(), child: r[1], cls: r[2], parent: r[3], phone: String(r[4]), c1: r[5] === 'иә', c2: r[6] === 'иә', c3: r[7] === 'иә', c4: r[8] === 'иә', lang: r[9], sig: r[10] ? 1 : 0, revoked: r[11] === 'қайтарып алды', no: r[12] || '', mail: String(r[13] || '') }));
    return tjOut_({ ok: true, rows: rows });
  }
  if (p.a === 'mailreg') {
    const c = tjCls_(p.c || '');
    const rows = tjConsents_().filter(r => !c || tjCls_(r[2]) === c).sort((a, b) => String(a[1]).localeCompare(String(b[1])));
    const name = 'Келісімдер тізілімі — ' + tjSafe_(p.c || 'барлық сынып') + ' — ' + Utilities.formatDate(new Date(), 'Asia/Tashkent', 'dd.MM.yyyy');
    try {
      MailApp.sendEmail({ to: TJ_EMAIL, subject: '📋 ' + name, body: 'Ата-ана келісімдерінің тізілімі PDF түрінде тіркелді (' + rows.length + ' жол).\n\nБілім Арена · Шың 140', attachments: [tjPdf_(tjRegisterHtml_(rows, p.c), name)], name: 'Білім Арена · Келісімдер' });
      return tjOut_({ ok: true, n: rows.length, to: TJ_EMAIL });
    } catch (e) { return tjOut_({ ok: false, error: String(e.message || e).slice(0, 160) }); }
  }
  return tjOut_({ ok: false });
}
