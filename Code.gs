/* =====================================================================
   I.EEA306 «Цахилгаан системийн тогтсон горим» — оноо хадгалах сервер
   Google Apps Script (Google Sheets-тэй холбоотой). Үнэгүй.
   Тохиргоо: Project Settings → Script Properties → PW = багшийн нууц үг
   Хуудсууд: Log (бүх үйлдэл), Students (оюутны жагсаалт: код, нэр, анги)
   ===================================================================== */
const SHEET = 'Log';
const HEAD = ['Огноо','SID','Код','Нэр','Анги','Төрөл','Түлхүүр','Оноо','Нийт','Секунд','Өдөр','Лекц','Асуулт'];
const ROSTER_SHEET = 'Students';

function ss_() { return SpreadsheetApp.getActiveSpreadsheet(); }
function sh_() {
  let s = ss_().getSheetByName(SHEET);
  if (!s) { s = ss_().insertSheet(SHEET); s.appendRow(HEAD); s.setFrozenRows(1); }
  return s;
}
function rosterSheet_() {
  let s = ss_().getSheetByName(ROSTER_SHEET);
  if (!s) { s = ss_().insertSheet(ROSTER_SHEET); s.appendRow(['Код', 'Овог нэр', 'Сургууль/анги']); s.setFrozenRows(1); }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function clip_(v, n) { return String(v == null ? '' : v).slice(0, n); }
function norm_(s) { return String(s || '').toUpperCase().replace(/\s+/g, '').replace(/Ё/g, 'Е'); }
function props_() { return PropertiesService.getScriptProperties(); }
function pwOk_(pw) { const p = props_().getProperty('PW'); return !!p && pw === p; }

function roster_() {
  const v = rosterSheet_().getDataRange().getValues();
  const out = [];
  for (let i = 1; i < v.length; i++) if (v[i][0]) out.push({ code: String(v[i][0]).trim(), name: String(v[i][1]).trim(), group: String(v[i][2] || '').trim() });
  return out;
}
function findStudent_(code) {
  const c = norm_(code);
  return roster_().find(r => norm_(r.code) === c) || null;
}
function theme_() {
  try { return JSON.parse(props_().getProperty('THEME') || 'null'); } catch (e) { return null; }
}

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = JSON.parse(e.postData.contents);
    const act = String(p.action || '');
    /* --- багшийн үйлдэл --- */
    if (act === 'setRoster' || act === 'setTheme') {
      if (!pwOk_(p.pw)) return out_({ ok: false, error: 'bad_pw' });
      if (act === 'setTheme') {
        const t = p.theme || {};
        const clean = { id: clip_(t.id, 20) };
        if (/^#[0-9a-fA-F]{6}$/.test(t.accent || '')) clean.accent = t.accent;
        props_().setProperty('THEME', JSON.stringify(clean));
        return out_({ ok: true });
      }
      const list = (p.roster || []).slice(0, 2000).map(r => [clip_(r.code, 20).trim(), clip_(r.name, 80).trim(), clip_(r.group, 60).trim()]).filter(r => r[0] && r[1]);
      const s = rosterSheet_();
      if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 3).clearContent();
      if (list.length) s.getRange(2, 1, list.length, 3).setValues(list);
      return out_({ ok: true, count: list.length });
    }
    /* --- оюутны үйлдэл: зөвхөн жагсаалтад байгаа код --- */
    if (['register', 'study', 'result'].indexOf(act) < 0) return out_({ ok: false, error: 'bad_action' });
    const st = findStudent_(p.code);
    if (!st) return out_({ ok: false, error: 'not_enrolled' });
    const s = sh_();
    const who = st.code;
    if (act === 'result' && p.key === 'final') {
      const v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if (norm_(v[i][2]) === norm_(who) && v[i][5] === 'result' && v[i][6] === 'final')
          return out_({ ok: false, error: 'final_exists' });
      }
    }
    s.appendRow([new Date(), clip_(p.sid, 40), who, st.name, st.group, act,
      clip_(p.key, 10), p.score == null ? '' : Number(p.score), p.total == null ? '' : Number(p.total),
      p.sec == null ? '' : Number(p.sec), clip_(p.date, 10),
      (p.lec || []).map(Number).join(','), (p.items || []).join(',')]);
    return out_({ ok: true });
  } catch (err) {
    return out_({ ok: false, error: String(err) });
  } finally {
    try { lock.releaseLock(); } catch (x) {}
  }
}

function doGet(e) {
  const prm = (e && e.parameter) || {};
  if (prm.action === 'settings') return out_({ ok: true, theme: theme_() });
  if (prm.action === 'login') {
    const st = findStudent_(prm.code);
    if (!st || norm_(st.name) !== norm_(prm.name)) return out_({ ok: false, error: 'not_found' });
    return out_({ ok: true, code: st.code, name: st.name, group: st.group });
  }
  if (prm.action === 'teacher') {
    if (!pwOk_(prm.pw)) return out_({ ok: false, error: 'bad_pw' });
    return out_({ ok: true, rows: aggregate_(), roster: roster_(), theme: theme_() });
  }
  return out_({ ok: true, msg: 'I.EEA306 server ажиллаж байна' });
}

function aggregate_() {
  const m = {};
  roster_().forEach(r => { m[norm_(r.code)] = { _id: r.code, name: r.name, group: r.group, code: r.code, study: {} }; });
  const v = sh_().getDataRange().getValues();
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    const key = norm_(r[2]);
    if (!key || !m[key]) continue;           // жагсаалтаас хасагдсан оюутныг тооцохгүй
    const o = m[key];
    const at = r[0] instanceof Date ? r[0].toISOString() : String(r[0]);
    if (r[5] === 'study' && r[10]) {
      const d = o.study[r[10]] || (o.study[r[10]] = { sec: 0, lec: [] });
      d.sec += Number(r[9]) || 0;
      String(r[11] || '').split(',').filter(String).forEach(x => { x = Number(x); if (d.lec.indexOf(x) < 0) d.lec.push(x); });
      d.lec.sort((a, b) => a - b);
    } else if (r[5] === 'result') {
      const k = String(r[6]), sc = Number(r[7]) || 0;
      if (k === 'final') {
        if (!o.final) o.final = { score: sc, at: at, sec: Number(r[9]) || 0 };
      } else if (/^t\d\d$/.test(k)) {
        const t = o[k] || (o[k] = { best: 0, attempts: 0, at: '' });
        t.best = Math.max(t.best, sc); t.attempts++; t.at = at; t.last = sc;
      }
    }
  }
  return Object.keys(m).map(k => m[k]);
}

/* Эхний жагсаалт: Apps Script засварлагч дээр энэ функцийг нэг удаа ▶ Run хийнэ. */
function seedRoster() {
  const list = [
    ['B220000166', 'Б.Энхбаатар', 'ӨмТДС'],
    ['B252183024', 'Т.Эрдэнэсүх', 'ӨмТДС'],
    ['B252183025', 'Б.Сувданзаяа', 'ӨмТДС'],
    ['B252183028', 'Ч.Цогнаран', 'ӨмТДС'],
    ['B252183032', 'Г.Ганганживаа', 'ӨмТДС'],
    ['B252183035', 'О.Одгариг', 'ӨмТДС'],
    ['B252183038', 'Б.Одбаяр', 'ӨмТДС']
  ];
  const s = rosterSheet_();
  if (s.getLastRow() > 1) s.getRange(2, 1, s.getLastRow() - 1, 3).clearContent();
  s.getRange(2, 1, list.length, 3).setValues(list);
}
