/**
 * UED London Project 2026/27: group allocation and pairing form backend.
 *
 * This script lives inside your private Google Sheet (Extensions > Apps Script).
 * It is never published on GitHub, so the access code below stays private.
 *
 * Sheets used:
 *   Roster     Name | UCL email          (you paste the class list here, kept private)
 *   Responses  one row per student, updated if they submit again
 *   Pairs      rebuilt automatically after every submission
 */

const CONFIG = {
  ACCESS_CODE: 'UED26LP27',      // share this code with students in class or on Moodle only
  EMAIL_DOMAIN: '@ucl.ac.uk',
  FORM_TITLE: 'Group Allocation London Project 2026/27',
  SEND_CONFIRMATION: true,           // emails each student a copy of their answers
  CLOSED: false,                     // set to true to stop accepting responses
  REPLY_TO: ''                       // optional, e.g. the UED team mailbox
};

const SKILLS = [
  'Project Management', 'Oral Presentations', 'Graphic Design',
  'Qualitative Research', 'Quantitative Research', 'Critical Literature Review', 'Report Writing'
];
const EMPLOYMENT = ['Part-time', 'Full time', 'Not currently working'];
const NO_PARTNER = 'NONE';

const RESP_HEADERS = ['Last updated', 'UCL email', 'Roster name', 'Name given']
  .concat(SKILLS.map(function (s, i) { return 'Rank ' + (i + 1); }))
  .concat(SKILLS.map(function (s) { return s + ' (rank)'; }))
  .concat(['Other skill', 'GIS familiarity', 'Most relevant work experience', 'Current employment',
           'Degrees and subject areas', 'Native English speaker', 'Chosen partner', 'Pair status', 'Times submitted']);

/* ---------- one-off setup: run this once from the editor ---------- */
function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const roster = sheet_('Roster');
  if (roster.getLastRow() === 0) {
    roster.appendRow(['Name', 'UCL email']);
    roster.setFrozenRows(1);
  }
  const resp = sheet_('Responses');
  if (resp.getLastRow() === 0) {
    resp.appendRow(RESP_HEADERS);
    resp.setFrozenRows(1);
    resp.getRange(1, 1, 1, RESP_HEADERS.length).setFontWeight('bold').setBackground('#eedeff');
  }
  sheet_('Pairs');
  rebuildPairs_();
  ss.toast('Setup complete. Paste the class list into the Roster tab.');
}

/* ---------- web endpoints ---------- */
function doGet() {
  return json_({ ok: true, message: 'London Project form backend is running.' });
}

function doPost(e) {
  let req;
  try { req = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request.' }); }

  if (req.website) return json_({ ok: true });                 // honeypot: silently ignore bots
  if (tooManyFailures_()) return json_({ ok: false, error: 'Too many attempts. Please try again in a few minutes.' });
  if (!codeOk_(req.code)) {
    recordFailure_();
    Utilities.sleep(800);
    return json_({ ok: false, error: 'That access code is not correct.' });
  }
  if (CONFIG.CLOSED) return json_({ ok: false, error: 'This form is now closed. Please contact the UED Team.' });

  if (req.action === 'roster') return json_({ ok: true, title: CONFIG.FORM_TITLE, students: publicRoster_() });
  if (req.action === 'submit') return submit_(req);
  return json_({ ok: false, error: 'Unknown action.' });
}

/* ---------- submission ---------- */
function submit_(req) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return json_({ ok: false, error: 'The form is busy. Please try again in a moment.' });
  try {
    const roster = roster_();
    const email = clean_(req.email, 120).toLowerCase();
    if (!email.endsWith(CONFIG.EMAIL_DOMAIN)) return json_({ ok: false, error: 'Please use your UCL email address.' });
    const me = roster.find(function (r) { return r.email === email; });
    if (!me) return json_({ ok: false, error: 'That email is not on the London Project class list. Please check it or contact the UED Team.' });

    const name = clean_(req.name, 100);
    const ranking = Array.isArray(req.ranking) ? req.ranking.map(function (x) { return String(x); }) : [];
    const other = clean_(req.otherSkill, 150);
    const gis = clean_(req.gis, 200);
    const work = clean_(req.work, 400);
    const employment = String(req.employment || '');
    const degrees = clean_(req.degrees, 1000);
    const native = String(req.native || '');
    const partner = String(req.partner || '');

    const validRanking = ranking.length === SKILLS.length && SKILLS.every(function (s) { return ranking.indexOf(s) !== -1; });
    if (!name || !validRanking || !other || !gis || !work || !degrees ||
        EMPLOYMENT.indexOf(employment) === -1 || ['Yes', 'No'].indexOf(native) === -1) {
      return json_({ ok: false, error: 'Some answers are missing. Please complete every question.' });
    }
    if (partner !== NO_PARTNER && !roster.some(function (r) { return r.name === partner; })) {
      return json_({ ok: false, error: 'Please choose your partner from the list.' });
    }
    if (partner === me.name) return json_({ ok: false, error: 'You cannot choose yourself as your partner.' });

    const sh = sheet_('Responses');
    const data = sh.getDataRange().getValues();
    let rowIndex = -1, times = 0;
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][1]).toLowerCase() === email) { rowIndex = i + 1; times = Number(data[i][RESP_HEADERS.length - 1]) || 0; break; }
    }
    const rankOf = SKILLS.map(function (s) { return ranking.indexOf(s) + 1; });
    const row = [new Date(), email, me.name, safe_(name)]
      .concat(ranking).concat(rankOf)
      .concat([safe_(other), safe_(gis), safe_(work), employment, safe_(degrees), native,
               partner === NO_PARTNER ? 'Please pair me' : partner, '', times + 1]);
    if (rowIndex > 0) sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);
    else sh.appendRow(row);

    const status = rebuildPairs_()[me.name] || '';
    if (CONFIG.SEND_CONFIRMATION) sendConfirmation_(email, name, ranking, other, gis, work, employment, degrees, native, partner, status);
    return json_({ ok: true, updated: rowIndex > 0, status: status });
  } finally {
    lock.releaseLock();
  }
}

/* ---------- pairs ---------- */
function rebuildPairs_() {
  const roster = roster_();
  const resp = sheet_('Responses');
  const data = resp.getLastRow() > 1 ? resp.getRange(2, 1, resp.getLastRow() - 1, RESP_HEADERS.length).getValues() : [];
  const pCol = RESP_HEADERS.indexOf('Chosen partner'), sCol = RESP_HEADERS.indexOf('Pair status');
  const choice = {};
  data.forEach(function (r) { choice[r[2]] = r[pCol]; });

  const status = {};
  roster.forEach(function (r) {
    const c = choice[r.name];
    if (c === undefined) status[r.name] = 'Not yet submitted';
    else if (c === 'Please pair me') status[r.name] = 'Asked the team to find a partner';
    else if (choice[c] === r.name) status[r.name] = 'Confirmed pair';
    else if (choice[c] === undefined) status[r.name] = 'Waiting for ' + c + ' to submit';
    else status[r.name] = c + ' chose someone else';
  });

  data.forEach(function (r, i) { resp.getRange(i + 2, sCol + 1).setValue(status[r[2]] || ''); });

  const ps = sheet_('Pairs');
  ps.clear();
  const out = [['Student', 'UCL email', 'Chose', 'Their partner chose', 'Status']];
  roster.forEach(function (r) {
    const c = choice[r.name];
    out.push([r.name, r.email, c === undefined ? '' : c, (c && choice[c] !== undefined) ? choice[c] : '', status[r.name]]);
  });
  ps.getRange(1, 1, out.length, 5).setValues(out);
  ps.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#eedeff');
  ps.setFrozenRows(1);
  return status;
}

function publicRoster_() {
  const status = rebuildPairsReadOnly_();
  return roster_().map(function (r) { return { name: r.name, paired: status[r.name] === 'Confirmed pair' }; });
}

function rebuildPairsReadOnly_() {
  const resp = sheet_('Responses');
  const data = resp.getLastRow() > 1 ? resp.getRange(2, 1, resp.getLastRow() - 1, RESP_HEADERS.length).getValues() : [];
  const pCol = RESP_HEADERS.indexOf('Chosen partner'), choice = {}, status = {};
  data.forEach(function (r) { choice[r[2]] = r[pCol]; });
  Object.keys(choice).forEach(function (n) { if (choice[choice[n]] === n) status[n] = 'Confirmed pair'; });
  return status;
}

/* ---------- email ---------- */
function sendConfirmation_(email, name, ranking, other, gis, work, employment, degrees, native, partner, status) {
  const lines = [
    'Dear ' + name + ',', '',
    'Thank you for completing the ' + CONFIG.FORM_TITLE + ' form. Here is a copy of your answers.', '',
    'Skills ranking: ' + ranking.map(function (s, i) { return (i + 1) + '. ' + s; }).join(', '),
    'Other skill: ' + other,
    'GIS familiarity: ' + gis,
    'Most relevant work experience: ' + work,
    'Current employment: ' + employment,
    'Degrees and subject areas: ' + degrees,
    'Native English speaker: ' + native,
    'Chosen partner: ' + (partner === NO_PARTNER ? 'You asked the team to find you a partner' : partner),
    'Pair status: ' + status, '',
    'If anything is wrong, simply submit the form again and your answers will be updated.', '',
    'With best wishes,', 'UED Team'
  ];
  const opts = { name: 'UED Team' };
  if (CONFIG.REPLY_TO) opts.replyTo = CONFIG.REPLY_TO;
  try { MailApp.sendEmail(email, CONFIG.FORM_TITLE + ': your answers', lines.join('\n'), opts); } catch (err) { console.log(err); }
}

/* ---------- helpers ---------- */
function roster_() {
  const sh = sheet_('Roster');
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, 2).getValues()
    .filter(function (r) { return String(r[0]).trim(); })
    .map(function (r) { return { name: String(r[0]).trim(), email: String(r[1]).trim().toLowerCase() }; });
}
function sheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}
function codeOk_(code) {
  return String(code || '').trim().toUpperCase() === CONFIG.ACCESS_CODE.trim().toUpperCase();
}
function tooManyFailures_() {
  return Number(CacheService.getScriptCache().get('fails') || 0) >= 40;
}
function recordFailure_() {
  const c = CacheService.getScriptCache();
  c.put('fails', String(Number(c.get('fails') || 0) + 1), 600);
}
function clean_(v, max) {
  return String(v == null ? '' : v).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F\u007F]/g, '').trim().slice(0, max);
}
function safe_(v) {           // stops spreadsheet formula injection
  return /^[=+\-@\t\r]/.test(v) ? "'" + v : v;
}
function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
