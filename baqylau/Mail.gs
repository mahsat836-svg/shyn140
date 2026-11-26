/**
 * Білім Арена · Бақылау тесті — ҮЗДІКТЕРДІ ПОШТАҒА ЖІБЕРУ (Шың 140)
 *
 * Не істейді:
 *   • Күн сайын 18:00-де — бүгінгі нәтижелер: әр сынып пен пән бойынша 1–2–3 орын.
 *   • Әр жұма 17:00-де — «Апта үздіктері» (жалпы ТОП-10) және «Сынып үздіктері».
 *   • Хат осы мекенжайға барады: EMAIL (төменде).
 *
 * Орнату (бір рет, 3 минут) — Code.gs-ке ТИМЕЙСІЗ, бөлек файл қосасыз:
 *   1) Google Кесте → Кеңейтімдер → Apps Script.
 *   2) Сол жақтағы «Файлдар» жанындағы «+» → «Сценарий» → атауы: Mail.
 *   3) Осы кодты түгел қойып, 💾 сақтаңыз.
 *   4) Жоғарыдағы тізімнен setupMailTriggers функциясын таңдап → «Орындау» (▶).
 *      Google рұқсат сұрайды (пошта жіберу) → рұқсат беріңіз.
 *   5) Тексеру үшін testMail функциясын «Орындау» → поштаңызды ашыңыз.
 * Қайта іске қосу (Deploy) қажет емес.
 */
const EMAIL = 'jmahsat@inbox.ru';      // ← хат баратын пошта
const DAILY_HOUR = 18;                  // күнделікті есеп сағаты
const WEEKLY_DAY = ScriptApp.WeekDay.FRIDAY, WEEKLY_HOUR = 17;
const TZ = 'Asia/Tashkent';             // UTC+5 (Қазақстан уақытымен бірдей)
const TOP_PER_GROUP = 3;                // әр сынып+пән бойынша неше орын
const TOP_WEEK = 10;                    // аптаның жалпы рейтингі

/* ---------- деректер ---------- */
// Кесте жолы: [Уақыт, ID, Аты-жөні, Сынып, Деңгей, Пән, Балл, Сұрақ саны, Баға, Күйі, Себебі, Секунд, Жауаптар]
function mailRows_(from, to) {
  const v = sheet_().getDataRange().getValues(); v.shift();
  const done = [], bad = [];
  v.forEach(r => {
    const t = new Date(r[0]).getTime();
    if (!t || t < from || t >= to) return;
    const a = { t, name: String(r[2]).trim(), cls: String(r[3]), lv: String(r[4]), subj: String(r[5]),
      score: Number(r[6]) || 0, total: Number(r[7]) || 0, grade: String(r[8] || ''), status: String(r[9]), sec: Number(r[11]) || 0 };
    if (a.status === 'Аяқталды' && a.total) { a.pct = a.score / a.total; done.push(a); }
    else if (a.status === 'Өшірілді') bad.push(a);
  });
  return { done, bad };
}
// Әр оқушының ең жақсы нәтижесі; балл тең болса — тезірек орындаған жоғары
function best_(list, keyFn) {
  const m = new Map();
  list.forEach(a => { const k = keyFn(a), b = m.get(k);
    if (!b || a.pct > b.pct || (a.pct === b.pct && a.sec < b.sec)) m.set(k, a); });
  return [...m.values()].sort((a, b) => b.pct - a.pct || a.sec - b.sec);
}
const who_ = a => a.name.toLowerCase() + '|' + a.cls;

/* ---------- хат безендіру ---------- */
function esc_(s) { return String(s).replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c])); }
function dur_(s) { s = Math.round(s); return Math.floor(s / 60) + ' мин ' + (s % 60) + ' с'; }
const MEDAL = ['🥇', '🥈', '🥉'];
function rowsHtml_(L) {
  return L.map((a, i) => `<tr><td style="padding:4px 8px">${MEDAL[i] || (i + 1) + '.'}</td><td style="padding:4px 8px"><b>${esc_(a.name)}</b></td>` +
    `<td style="padding:4px 8px">${esc_(a.cls)}-сынып</td><td style="padding:4px 8px">${esc_(a.subj)}</td>` +
    `<td style="padding:4px 8px"><b>${a.score}/${a.total}</b>${a.grade ? ' · баға ' + esc_(a.grade) : ''}</td>` +
    `<td style="padding:4px 8px;color:#777">${dur_(a.sec)}</td></tr>`).join('');
}
function table_(title, L) {
  return `<h3 style="margin:18px 0 6px;color:#0b1f4d">${esc_(title)}</h3>` +
    `<table style="border-collapse:collapse;font-size:14px">${rowsHtml_(L)}</table>`;
}
function wrap_(title, sub, body) {
  return `<div style="font-family:Arial,sans-serif;max-width:720px">` +
    `<div style="background:#0b1f4d;color:#fff;padding:14px 18px;border-radius:10px">` +
    `<div style="font-size:12px;color:#f2c14e;letter-spacing:1px">БІЛІМ АРЕНА · ШЫҢ 140</div>` +
    `<div style="font-size:20px;font-weight:bold">${esc_(title)}</div><div style="font-size:13px;opacity:.85">${esc_(sub)}</div></div>` +
    body +
    `<p style="color:#888;font-size:12px;margin-top:22px">Хат автоматты түрде жіберілді. Толық нәтиже: ` +
    `<a href="https://mahsat836-svg.github.io/shyn140/baqylau/mugalim.html">мұғалім панелі</a> · Google Кесте «${esc_(SHEET_NAME)}».</p></div>`;
}
function day0_(d) { const s = Utilities.formatDate(d, TZ, 'yyyy-MM-dd'); return new Date(s + 'T00:00:00+05:00').getTime(); }
function fmt_(t) { return Utilities.formatDate(new Date(t), TZ, 'dd.MM.yyyy'); }

/* ---------- күнделікті есеп ---------- */
function sendDailyTop() {
  const to = Date.now(), from = day0_(new Date());
  const { done, bad } = mailRows_(from, to);
  if (!done.length) return;                         // бүгін тест болмаса — хат жоқ
  const groups = new Map();
  done.forEach(a => { const k = a.cls + '|' + a.subj; (groups.get(k) || groups.set(k, []).get(k)).push(a); });
  const keys = [...groups.keys()].sort((x, y) => parseInt(x) - parseInt(y) || x.localeCompare(y));
  let body = `<p style="font-size:14px">Бүгін аяқталған тест: <b>${done.length}</b> · сайттан шығып өшірілген: <b>${bad.length}</b></p>`;
  keys.forEach(k => { const [c, s] = k.split('|'); body += table_(`${c}-сынып · ${s}`, best_(groups.get(k), who_).slice(0, TOP_PER_GROUP)); });
  MailApp.sendEmail({ to: EMAIL, subject: `Білім Арена: бүгінгі үздіктер (${fmt_(to)})`,
    htmlBody: wrap_('Бүгінгі жүлделі орындар', fmt_(to) + ' · әр сынып пен пән бойынша 1–2–3 орын', body) });
}

/* ---------- апталық есеп ---------- */
function sendWeeklyTop() {
  const now = new Date(), to = now.getTime();
  const wd = Number(Utilities.formatDate(now, TZ, 'u'));       // 1 = дүйсенбі
  const from = day0_(now) - (wd - 1) * 864e5;
  const { done, bad } = mailRows_(from, to);
  if (!done.length) return;
  const all = best_(done, who_);
  const byCls = new Map();
  all.forEach(a => { if (!byCls.has(a.cls)) byCls.set(a.cls, a); });   // all сұрыпталған → бірінші = үздік
  const clsTop = [...byCls.values()].sort((a, b) => parseInt(a.cls) - parseInt(b.cls));
  const body = `<p style="font-size:14px">Аптада аяқталған тест: <b>${done.length}</b> · қатысқан оқушы: <b>${all.length}</b> · өшірілген: <b>${bad.length}</b></p>` +
    table_(`Апта үздіктері — ТОП-${TOP_WEEK}`, all.slice(0, TOP_WEEK)) +
    table_('Сынып үздіктері', clsTop).replace(/🥇|🥈|🥉|\d+\./g, '⭐');
  MailApp.sendEmail({ to: EMAIL, subject: `Білім Арена: апта үздіктері (${fmt_(from)} – ${fmt_(to)})`,
    htmlBody: wrap_('Апта үздіктері', fmt_(from) + ' – ' + fmt_(to) + ' · Instagram-ға дайын тізім', body) });
}

/* ---------- баптау және тексеру ---------- */
function setupMailTriggers() {
  ScriptApp.getProjectTriggers().forEach(t => {
    if (['sendDailyTop', 'sendWeeklyTop'].includes(t.getHandlerFunction())) ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyTop').timeBased().everyDays(1).atHour(DAILY_HOUR).inTimezone(TZ).create();
  ScriptApp.newTrigger('sendWeeklyTop').timeBased().onWeekDay(WEEKLY_DAY).atHour(WEEKLY_HOUR).inTimezone(TZ).create();
  Logger.log('Дайын: күн сайын %s:00 және жұма %s:00 → %s', DAILY_HOUR, WEEKLY_HOUR, EMAIL);
}
function testMail() {
  const n = mailRows_(0, Date.now()).done.length;
  MailApp.sendEmail({ to: EMAIL, subject: 'Білім Арена: тексеру хаты',
    htmlBody: wrap_('Пошта жұмыс істеп тұр ✓', 'Тексеру хаты', `<p>Кестеде барлығы <b>${n}</b> аяқталған тест бар. ` +
      `Күн сайын ${DAILY_HOUR}:00-де бүгінгі үздіктер, жұма ${WEEKLY_HOUR}:00-де апта үздіктері келеді.</p>`) });
  sendWeeklyTop();   // осы аптада нәтиже болса — апталық хатты да бірден жібереді
}
