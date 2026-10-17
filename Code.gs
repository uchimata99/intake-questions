/** @OnlyCurrentDoc */
// ==== Intake questions - Google Sheet backend ====
// Admin code: needed on the big screen to drag, edit topics and start sessions.
const ADMIN_CODE = '1234';

const Q_SHEET = 'שאלות';
const H_SHEET = 'כותרות';
const S_SHEET = 'סשנים';
const Q_HEAD = ['id', 'text', 'author', 'created', 'heading', 'order', 'deleted', 'topic'];
const H_HEAD = ['id', 'title', 'order'];
const S_HEAD = ['id', 'name', 'created', 'qsheet', 'hsheet', 'count'];

function doGet(e) {
  return handle_((e && e.parameter) || {});
}

function doPost(e) {
  let p = {};
  try { p = JSON.parse(e.postData.contents); } catch (err) {}
  return handle_(p);
}

function handle_(p) {
  let out;
  try {
    const action = p.action || 'list';
    if (action === 'list') {
      out = state_();
    } else if (action === 'add') {
      out = withLock_(function () { return add_(p); });
    } else if (String(p.code || '') !== String(ADMIN_CODE)) {
      out = { error: 'code' };
    } else if (action === 'check') {
      out = { ok: true };
    } else if (action === 'layout') {
      out = withLock_(function () { return layout_(p); });
    } else if (action === 'edit') {
      out = withLock_(function () { return edit_(p); });
    } else if (action === 'delete') {
      out = withLock_(function () { return del_(p); });
    } else if (action === 'newSession') {
      out = withLock_(function () { return newSession_(p); });
    } else if (action === 'sessions') {
      out = sessions_();
    } else if (action === 'session') {
      out = session_(p);
    } else {
      out = { error: 'action' };
    }
  } catch (err) {
    out = { error: String(err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out))
    .setMimeType(ContentService.MimeType.JSON);
}

function withLock_(fn) {
  const lock = LockService.getScriptLock();
  lock.waitLock(15000);
  try { return fn(); } finally { lock.releaseLock(); }
}

function sheet_(name, head) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, head.length).setValues([head]);
    sh.setFrozenRows(1);
  } else if (sh.getLastColumn() < head.length) {
    sh.getRange(1, 1, 1, head.length).setValues([head]);
  }
  return sh;
}

function rows_(sh, head) {
  const n = sh.getLastRow() - 1;
  if (n < 1) return [];
  return sh.getRange(2, 1, n, head.length).getValues().map(function (r, i) {
    const o = { _row: i + 2 };
    head.forEach(function (k, j) { o[k] = r[j]; });
    return o;
  });
}

function read_(qSh, hSh) {
  const qs = rows_(qSh, Q_HEAD)
    .filter(function (q) { return q.id && !q.deleted; })
    .map(function (q) {
      return {
        id: String(q.id), text: String(q.text), author: String(q.author || ''),
        created: Number(q.created) || 0, heading: String(q.heading || ''),
        order: Number(q.order) || 0, topic: String(q.topic || '')
      };
    });
  const hs = rows_(hSh, H_HEAD)
    .filter(function (h) { return h.id; })
    .map(function (h) { return { id: String(h.id), title: String(h.title), order: Number(h.order) || 0 }; })
    .sort(function (a, b) { return a.order - b.order; });
  return { ok: true, questions: qs, headings: hs, now: Date.now() };
}

function state_() {
  return read_(sheet_(Q_SHEET, Q_HEAD), sheet_(H_SHEET, H_HEAD));
}

function clean_(s, max) {
  return String(s || '').replace(/[\u0000-\u0008\u000B-\u001F]/g, '').trim().slice(0, max);
}

function safe_(s) {
  return /^[=+\-@]/.test(s) ? "'" + s : s;
}

function add_(p) {
  const text = clean_(p.text, 500);
  if (!text) return { error: 'empty' };
  const author = clean_(p.author, 40);
  const topic = clean_(p.topic, 80);
  let heading = '';
  if (topic) {
    const key = topic.replace(/\s+/g, ' ').toLowerCase();
    rows_(sheet_(H_SHEET, H_HEAD), H_HEAD).forEach(function (h) {
      if (!heading && h.id && String(h.title).replace(/\s+/g, ' ').trim().toLowerCase() === key) heading = String(h.id);
    });
  }
  const id = 'q' + Utilities.getUuid().slice(0, 8);
  const now = Date.now();
  sheet_(Q_SHEET, Q_HEAD).appendRow([id, safe_(text), safe_(author), now, heading, now, '', safe_(topic)]);
  return { ok: true, id: id };
}

function layout_(p) {
  const hSh = sheet_(H_SHEET, H_HEAD);
  if (Array.isArray(p.headings)) {
    const last = hSh.getLastRow();
    if (last > 1) hSh.getRange(2, 1, last - 1, H_HEAD.length).clearContent();
    const vals = p.headings.slice(0, 60).map(function (h, i) {
      return [clean_(h.id, 20), safe_(clean_(h.title, 80)), i];
    }).filter(function (r) { return r[0] && r[1]; });
    if (vals.length) hSh.getRange(2, 1, vals.length, H_HEAD.length).setValues(vals);
  }
  if (Array.isArray(p.placement)) {
    const qSh = sheet_(Q_SHEET, Q_HEAD);
    const all = rows_(qSh, Q_HEAD);
    const byId = {};
    p.placement.forEach(function (x) { byId[String(x.id)] = x; });
    const hCol = Q_HEAD.indexOf('heading') + 1;
    if (all.length) {
      const vals = qSh.getRange(2, hCol, all.length, 2).getValues();
      all.forEach(function (q, i) {
        const x = byId[String(q.id)];
        if (x) vals[i] = [clean_(x.heading, 20), Number(x.order) || 0];
      });
      qSh.getRange(2, hCol, all.length, 2).setValues(vals);
    }
  }
  return state_();
}

function find_(id) {
  const sh = sheet_(Q_SHEET, Q_HEAD);
  const q = rows_(sh, Q_HEAD).filter(function (r) { return String(r.id) === String(id); })[0];
  return { sh: sh, q: q };
}

function edit_(p) {
  const f = find_(p.id);
  if (!f.q) return { error: 'missing' };
  const text = clean_(p.text, 500);
  if (!text) return { error: 'empty' };
  f.sh.getRange(f.q._row, Q_HEAD.indexOf('text') + 1).setValue(safe_(text));
  return state_();
}

function del_(p) {
  const f = find_(p.id);
  if (!f.q) return { error: 'missing' };
  f.sh.getRange(f.q._row, Q_HEAD.indexOf('deleted') + 1).setValue(1);
  return state_();
}

// ---- sessions: archive the current board and start a new one ----
function newSession_(p) {
  const qSh = sheet_(Q_SHEET, Q_HEAD);
  const hSh = sheet_(H_SHEET, H_HEAD);
  const now = new Date();
  const name = clean_(p.name, 60) ||
    Utilities.formatDate(now, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
  const id = 's' + Utilities.getUuid().slice(0, 8);
  const count = rows_(qSh, Q_HEAD).filter(function (q) { return q.id && !q.deleted; }).length;
  const keep = p.keepTopics !== false;
  const hVals = keep ? rows_(hSh, H_HEAD).filter(function (h) { return h.id; })
    .map(function (h) { return [h.id, h.title, h.order]; }) : [];
  qSh.setName('ארכיון ' + id + ' שאלות');
  hSh.setName('ארכיון ' + id + ' נושאים');
  sheet_(S_SHEET, S_HEAD).appendRow([id, safe_(name), now.getTime(), qSh.getName(), hSh.getName(), count]);
  sheet_(Q_SHEET, Q_HEAD);
  const nh = sheet_(H_SHEET, H_HEAD);
  if (hVals.length) nh.getRange(2, 1, hVals.length, H_HEAD.length).setValues(hVals);
  return state_();
}

function sessions_() {
  const list = rows_(sheet_(S_SHEET, S_HEAD), S_HEAD)
    .filter(function (s) { return s.id; })
    .map(function (s) {
      return { id: String(s.id), name: String(s.name), created: Number(s.created) || 0, count: Number(s.count) || 0 };
    })
    .sort(function (a, b) { return b.created - a.created; });
  return { ok: true, sessions: list };
}

function session_(p) {
  const s = rows_(sheet_(S_SHEET, S_HEAD), S_HEAD).filter(function (r) { return String(r.id) === String(p.id); })[0];
  if (!s) return { error: 'missing' };
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const qSh = ss.getSheetByName(String(s.qsheet));
  const hSh = ss.getSheetByName(String(s.hsheet));
  if (!qSh || !hSh) return { error: 'missing' };
  const out = read_(qSh, hSh);
  out.session = { id: String(s.id), name: String(s.name), created: Number(s.created) || 0 };
  return out;
}
