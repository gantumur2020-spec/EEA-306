/* =====================================================================
   I.EEA306 «Цахилгаан системийн тогтсон горим» — оноо хадгалах сервер
   Google Apps Script (Google Sheets-тэй холбоотой). Үнэгүй.
   Тохиргоо: Төслийн тохиргоо → Script Properties → PW = багшийн нууц үг
   ===================================================================== */
const SHEET = 'Log';
const HEAD = ['Огноо','SID','Код','Нэр','Анги','Төрөл','Түлхүүр','Оноо','Нийт','Секунд','Өдөр','Лекц','Асуулт'];

function sh_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let s = ss.getSheetByName(SHEET);
  if (!s) { s = ss.insertSheet(SHEET); s.appendRow(HEAD); s.setFrozenRows(1); }
  return s;
}
function out_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
function clip_(v, n) { return String(v == null ? '' : v).slice(0, n); }

function doPost(e) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(15000);
    const p = JSON.parse(e.postData.contents);
    const act = String(p.action || '');
    if (['register', 'study', 'result'].indexOf(act) < 0) return out_({ ok: false, error: 'bad_action' });
    const s = sh_();
    const sid = clip_(p.sid, 40), code = clip_(p.code, 30), who = code || sid;
    if (act === 'result' && p.key === 'final') {
      const v = s.getDataRange().getValues();
      for (let i = 1; i < v.length; i++) {
        if ((v[i][2] || v[i][1]) === who && v[i][5] === 'result' && v[i][6] === 'final')
          return out_({ ok: false, error: 'final_exists' });
      }
    }
    s.appendRow([new Date(), sid, code, clip_(p.name, 80), clip_(p.group, 40), act,
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
  if (prm.action === 'teacher') {
    const pw = PropertiesService.getScriptProperties().getProperty('PW');
    if (!pw || prm.pw !== pw) return out_({ ok: false, error: 'bad_pw' });
    return out_({ ok: true, rows: aggregate_() });
  }
  return out_({ ok: true, msg: 'I.EEA306 server ажиллаж байна' });
}

function aggregate_() {
  const v = sh_().getDataRange().getValues();
  const m = {};
  for (let i = 1; i < v.length; i++) {
    const r = v[i];
    const id = String(r[2] || r[1]);
    if (!id) continue;
    const at = r[0] instanceof Date ? r[0].toISOString() : String(r[0]);
    const o = m[id] || (m[id] = { _id: id, name: '', group: '', code: '', study: {} });
    if (r[3]) o.name = r[3];
    if (r[4]) o.group = r[4];
    if (r[2]) o.code = r[2];
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
