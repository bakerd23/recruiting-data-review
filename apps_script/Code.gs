/**
 * Recruiting Data Review: the back end for the intern review site (Google Apps Script attached to a private Sheet).
 *
 * The site on GitHub Pages holds NO player data. Every request carries the reviewer's Google sign-in
 * token; this script checks it with Google, looks the email up in the Reviewers tab, and only then
 * hands out items or records decisions. Contact details (phone, email, birth date) never reach the
 * Sheet: items only say whether they match.
 *
 * The API (POST, JSON body sent as text/plain so the browser needs no CORS preflight):
 *   {action, token, ...}
 *   me            -> who you are, your role, counts per tool
 *   next          {tool}             -> your reserved item for that tool, or reserves the next open one
 *   release       {item_id}          -> give a reserved item back (Skip)
 *   decide        {item_id, decision, details, note}   decision: see DECISIONS
 *   undo          {}                 -> cancels YOUR latest decision; the item is reserved for you again
 *   mine          {tool}             -> your own decisions only (nobody sees anyone else's)
 *   ceeb          {}                 -> College Board school list (high school tool search)
 *   hs_index      {}                 -> id/name/city/state of every high school item ("same as" search)
 *   escalated     {tool}             -> items marked "needs review" (senior and admin only)
 *   resolve       {item_id, decision, details, note}  (senior and admin only)
 *   reviewers     {}                 -> the reviewer list (admin only)
 *   set_reviewer  {email, role, active, name, note}   (admin only)
 * The future Cloud Run service answers exactly the same actions; the site only changes its endpoint.
 *
 * Tabs (each mirrors a planned BigQuery table):
 *   Reviewers  email | role (intern, senior, admin) | active | name | added_by | added_at | note
 *   Queue      item_id | tool | priority | added_at | payload (JSON, built by export_queues.py)
 *   Claims     item_id | tool | email | claimed_at | expires_at
 *   Decisions  append-only: decision_id | item_id | tool | email | role | kind | decision | details |
 *              note | decided_at | undoes
 *   AdminLog   every change to Reviewers
 *
 * Script Properties (never in the repo):
 *   CLIENT_ID      the Google sign-in client ID (same as in config.js)
 *   CEEB_FILE_ID   Drive file id of ceeb.json (College Board list for the high school tool)
 */

const TOOLS = ['hs', 'dup', 'pair'];
const ROLES = ['intern', 'senior', 'admin'];
const DECISIONS = {
  hs: ['match', 'same_as', 'newhs', 'nohs', 'needs_review'],
  dup: ['same', 'different', 'needs_review'],
  pair: ['same', 'different', 'needs_review'],
};
const CLAIM_MINUTES = 30;

// ---------------------------------------------------------------- two reviewers (OFF)
// How many reviewers must give the SAME answer before an item counts as decided. 1 = one look
// (decided Nov 9). To require two interns to agree:
//   1. set REQUIRED_AGREEMENTS = 2
//   2. in nextOpen_(), skip items the caller has already decided (the line marked AGREEMENT)
//   3. itemState_() then keeps an item open until two matching answers exist, and sends it to the
//      senior reviewers ("escalated") when two reviewers disagree.
const REQUIRED_AGREEMENTS = 1;

const H = {
  Reviewers: ['email', 'role', 'active', 'name', 'added_by', 'added_at', 'note'],
  Queue: ['item_id', 'tool', 'priority', 'added_at', 'payload'],
  Claims: ['item_id', 'tool', 'email', 'claimed_at', 'expires_at'],
  Decisions: ['decision_id', 'item_id', 'tool', 'email', 'role', 'kind', 'decision', 'details', 'note', 'decided_at', 'undoes'],
  AdminLog: ['at', 'by', 'email', 'role', 'active', 'name', 'note'],
};

// ---------------------------------------------------------------- helpers
function json_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function text_(v) { const s = String(v == null ? '' : v); return /^[=+\-@]/.test(s) ? "'" + s : s; }   // never a formula
function prop_(k) { return PropertiesService.getScriptProperties().getProperty(k); }

function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(name);
  if (!sh) {
    sh = ss.insertSheet(name);
    sh.getRange(1, 1, 1, H[name].length).setValues([H[name]]).setFontWeight('bold');
    sh.setFrozenRows(1);
  }
  return sh;
}
function rows_(name, ncols) {
  const sh = sheet_(name), n = sh.getLastRow() - 1;
  return n > 0 ? sh.getRange(2, 1, n, ncols || H[name].length).getValues() : [];
}
function append_(name, rows) {
  if (!rows.length) return;
  const sh = sheet_(name);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}
function iso_(d) { return (d instanceof Date ? d : new Date(d || Date.now())).toISOString(); }

// ---------------------------------------------------------------- sign-in
function verify_(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const key = 'tok_' + Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)).slice(0, 40);
  const hit = cache.get(key);
  if (hit) return hit;
  const r = UrlFetchApp.fetch('https://oauth2.googleapis.com/tokeninfo?id_token=' + encodeURIComponent(token), { muteHttpExceptions: true });
  if (r.getResponseCode() !== 200) return null;
  const t = JSON.parse(r.getContentText());
  if (t.aud !== prop_('CLIENT_ID') || String(t.email_verified) !== 'true' || Number(t.exp) * 1000 < Date.now()) return null;
  const email = String(t.email).toLowerCase();
  cache.put(key, email, 300);
  return email;
}
function reviewer_(email) {
  for (const r of rows_('Reviewers')) {
    if (String(r[0]).trim().toLowerCase() === email && (r[2] === true || String(r[2]).toUpperCase() === 'TRUE'))
      return { email: email, role: ROLES.indexOf(String(r[1])) >= 0 ? String(r[1]) : 'intern', name: String(r[3] || '') };
  }
  return null;
}
const isSenior_ = (u) => u.role === 'senior' || u.role === 'admin';

// ---------------------------------------------------------------- item state
function decisionsByItem_() {
  const all = rows_('Decisions');
  const undone = new Set(all.filter(r => r[5] === 'undo').map(r => String(r[10])));
  const by = {};
  for (const r of all) {
    if (r[5] === 'undo' || undone.has(String(r[0]))) continue;
    (by[String(r[1])] = by[String(r[1])] || []).push(
      { id: String(r[0]), email: String(r[3]), role: String(r[4]), kind: String(r[5]), decision: String(r[6]), at: r[9] });
  }
  return by;
}
// open | escalated | decided
function itemState_(ds) {
  if (!ds || !ds.length) return 'open';
  if (ds.some(d => d.kind === 'resolve')) return 'decided';
  if (ds.some(d => d.decision === 'needs_review')) return 'escalated';
  if (REQUIRED_AGREEMENTS <= 1) return 'decided';
  // Two-reviewer mode (OFF while REQUIRED_AGREEMENTS = 1):
  // const answers = {}; ds.forEach(d => { answers[d.decision] = (answers[d.decision] || 0) + 1; });
  // if (Object.keys(answers).length > 1) return 'escalated';                  // reviewers disagree
  // return Math.max.apply(null, Object.values(answers)) >= REQUIRED_AGREEMENTS ? 'decided' : 'open';
  return 'decided';
}

function queueIndex_() {          // item_id, tool, priority, row number (payloads are read one at a time)
  return rows_('Queue', 3).map((r, i) => ({ id: String(r[0]), tool: String(r[1]), priority: Number(r[2]) || 0, row: i + 2 }));
}
function payload_(row) { return JSON.parse(String(sheet_('Queue').getRange(row, 5).getValue() || '{}')); }

function activeClaims_() {
  const now = Date.now(), out = {};
  rows_('Claims').forEach((r, i) => {
    if (new Date(r[4]).getTime() > now) out[String(r[0])] = { email: String(r[2]), tool: String(r[1]), row: i + 2 };
  });
  return out;
}
function claim_(item, email) {
  const now = new Date();
  append_('Claims', [[item.id, item.tool, email, iso_(now), iso_(now.getTime() + CLAIM_MINUTES * 60000)]]);
}
function release_(itemId, email) {
  const sh = sheet_('Claims'), now = iso_();
  rows_('Claims').forEach((r, i) => {
    if (String(r[0]) === itemId && String(r[2]) === email && new Date(r[4]).getTime() > Date.now())
      sh.getRange(i + 2, 5).setValue(now);                           // expire it now
  });
}

function nextOpen_(tool, email, byItem, claims) {
  const items = queueIndex_().filter(q => q.tool === tool).sort((a, b) => b.priority - a.priority || a.row - b.row);
  const mine = items.find(q => claims[q.id] && claims[q.id].email === email && itemState_(byItem[q.id]) === 'open');
  if (mine) return mine;
  return items.find(q => !claims[q.id] && itemState_(byItem[q.id]) === 'open'
    // AGREEMENT: in two-reviewer mode, also require: && !(byItem[q.id] || []).some(d => d.email === email)
  );
}

function counts_(email) {
  const byItem = decisionsByItem_(), out = {};
  TOOLS.forEach(t => { out[t] = { total: 0, open: 0, decided: 0, escalated: 0, mine: 0 }; });
  for (const q of queueIndex_()) {
    if (!out[q.tool]) continue;
    const c = out[q.tool], st = itemState_(byItem[q.id]);
    c.total++; c[st]++;
    if ((byItem[q.id] || []).some(d => d.email === email && d.kind === 'decision')) c.mine++;
  }
  return out;
}

// ---------------------------------------------------------------- actions
function act_(body, u) {
  const a = body.action;
  if (a === 'me') return { ok: true, email: u.email, role: u.role, name: u.name, counts: counts_(u.email),
                           claim_minutes: CLAIM_MINUTES, required_agreements: REQUIRED_AGREEMENTS };

  if (a === 'next') {
    if (TOOLS.indexOf(body.tool) < 0) return { ok: false, error: 'bad_tool' };
    const byItem = decisionsByItem_(), claims = activeClaims_();
    const q = nextOpen_(body.tool, u.email, byItem, claims);
    if (!q) return { ok: true, item: null };
    if (!claims[q.id]) claim_(q, u.email);
    return { ok: true, item: payload_(q.row), item_id: q.id };
  }

  if (a === 'release') { release_(String(body.item_id), u.email); return { ok: true }; }

  if (a === 'decide' || a === 'resolve') {
    if (a === 'resolve' && !isSenior_(u)) return { ok: false, error: 'not_allowed' };
    const q = queueIndex_().find(x => x.id === String(body.item_id));
    if (!q) return { ok: false, error: 'no_item' };
    if (DECISIONS[q.tool].indexOf(body.decision) < 0 || (a === 'resolve' && body.decision === 'needs_review'))
      return { ok: false, error: 'bad_decision' };
    const byItem = decisionsByItem_(), st = itemState_(byItem[q.id]);
    if (a === 'decide') {
      const claims = activeClaims_();
      if (st !== 'open') return { ok: false, error: 'already_decided' };
      if (claims[q.id] && claims[q.id].email !== u.email) return { ok: false, error: 'reserved_by_someone_else' };
    } else if (st !== 'escalated') return { ok: false, error: 'not_escalated' };
    const id = Utilities.getUuid();
    append_('Decisions', [[id, q.id, q.tool, u.email, u.role, a === 'resolve' ? 'resolve' : 'decision', body.decision,
      text_(JSON.stringify(body.details || {}).slice(0, 5000)), text_(String(body.note || '').slice(0, 2000)), iso_(), '']]);
    release_(q.id, u.email);
    return { ok: true, decision_id: id };
  }

  if (a === 'undo') {             // the caller's latest decision that is still in effect (and not yet resolved by a senior)
    const all = rows_('Decisions');
    const undone = new Set(all.filter(r => r[5] === 'undo').map(r => String(r[10])));
    const byItem = decisionsByItem_();
    for (let i = all.length - 1; i >= 0; i--) {
      const r = all[i];
      if (String(r[3]) !== u.email || r[5] === 'undo' || undone.has(String(r[0]))) continue;
      if (r[5] === 'decision' && (byItem[String(r[1])] || []).some(d => d.kind === 'resolve'))
        return { ok: false, error: 'already_resolved_by_senior' };
      append_('Decisions', [[Utilities.getUuid(), r[1], r[2], u.email, u.role, 'undo', '', '', '', iso_(), r[0]]]);
      const q = queueIndex_().find(x => x.id === String(r[1]));
      if (q && r[5] === 'decision') claim_(q, u.email);
      return { ok: true, item_id: String(r[1]), tool: String(r[2]) };
    }
    return { ok: false, error: 'nothing_to_undo' };
  }

  if (a === 'mine') {              // ONLY the caller's own decisions
    const all = rows_('Decisions');
    const undone = new Set(all.filter(r => r[5] === 'undo').map(r => String(r[10])));
    const out = all.filter(r => String(r[3]) === u.email && r[5] !== 'undo' && (!body.tool || String(r[2]) === body.tool))
      .map(r => ({ decision_id: String(r[0]), item_id: String(r[1]), tool: String(r[2]), kind: String(r[5]),
                   decision: String(r[6]), note: String(r[8]), at: iso_(r[9]), undone: undone.has(String(r[0])) }));
    return { ok: true, rows: out.slice(-200).reverse() };
  }

  if (a === 'ceeb') {
    const id = prop_('CEEB_FILE_ID');
    if (!id) return { ok: false, error: 'no_ceeb_file' };
    return { ok: true, ceeb: JSON.parse(DriveApp.getFileById(id).getBlob().getDataAsString('UTF-8')) };
  }

  if (a === 'hs_index') {
    const sh = sheet_('Queue'), idx = queueIndex_().filter(q => q.tool === 'hs');
    const vals = idx.length ? sh.getRange(2, 5, sh.getLastRow() - 1, 1).getValues() : [];
    return { ok: true, rows: idx.map(q => { const p = JSON.parse(String(vals[q.row - 2][0] || '{}'));
      return { id: q.id, name: p.name, city: p.city, state: p.state, source: p.source }; }) };
  }

  if (a === 'escalated') {
    if (!isSenior_(u)) return { ok: false, error: 'not_allowed' };
    const byItem = decisionsByItem_();
    const out = queueIndex_().filter(q => (!body.tool || q.tool === body.tool) && itemState_(byItem[q.id]) === 'escalated')
      .slice(0, 50).map(q => ({ item: payload_(q.row), item_id: q.id, tool: q.tool,
        flagged: (byItem[q.id] || []).filter(d => d.decision === 'needs_review').map(d => ({ by: d.email, at: iso_(d.at) })) }));
    // seniors also see the note left with "needs review"
    const notes = {};
    rows_('Decisions').forEach(r => { if (String(r[6]) === 'needs_review') notes[String(r[1])] = String(r[8]); });
    out.forEach(o => { o.note = notes[o.item_id] || ''; });
    return { ok: true, rows: out };
  }

  if (a === 'reviewers' || a === 'set_reviewer') {
    if (u.role !== 'admin') return { ok: false, error: 'not_allowed' };
    if (a === 'set_reviewer') {
      const email = String(body.email || '').trim().toLowerCase();
      if (!/^[^@\s]+@[^@\s]+\.[a-z]{2,}$/.test(email) || ROLES.indexOf(body.role) < 0) return { ok: false, error: 'bad_reviewer' };
      const sh = sheet_('Reviewers'), rs = rows_('Reviewers');
      const i = rs.findIndex(r => String(r[0]).trim().toLowerCase() === email);
      const row = [email, body.role, body.active !== false, text_(String(body.name || '').slice(0, 80)),
                   i >= 0 ? rs[i][4] : u.email, i >= 0 ? rs[i][5] : iso_(), text_(String(body.note || '').slice(0, 300))];
      if (i >= 0) sh.getRange(i + 2, 1, 1, row.length).setValues([row]); else append_('Reviewers', [row]);
      append_('AdminLog', [[iso_(), u.email, email, body.role, body.active !== false, row[3], row[6]]]);
    }
    return { ok: true, rows: rows_('Reviewers').map(r => ({ email: String(r[0]), role: String(r[1]),
      active: r[2] === true || String(r[2]).toUpperCase() === 'TRUE', name: String(r[3]), added_by: String(r[4]),
      added_at: iso_(r[5]), note: String(r[6]) })) };
  }
  return { ok: false, error: 'bad_action' };
}

function doPost(e) {
  let body;
  try { body = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'bad_json' }); }
  const email = verify_(body.token);
  if (!email) return json_({ ok: false, error: 'not_signed_in' });
  const u = reviewer_(email);
  if (!u) return json_({ ok: false, error: 'not_on_reviewer_list', email: email });
  const lock = LockService.getScriptLock();          // one change at a time (reservations, decisions)
  lock.waitLock(25000);
  try { return json_(act_(body, u)); }
  catch (err) { return json_({ ok: false, error: 'server_error', detail: String(err).slice(0, 300) }); }
  finally { lock.releaseLock(); }
}
function doGet() { return json_({ ok: true, service: 'recruiting-data-review' }); }

// One-time setup: run this once from the Apps Script editor (select setupTabs, click Run).
// It creates the five tabs with their header rows. Then add yourself to Reviewers.
function setupTabs() { Object.keys(H).forEach(name => sheet_(name)); }
