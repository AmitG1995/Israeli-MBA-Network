/**
 * Israeli MBA Network: sign-up backend (Google Apps Script).
 *
 * - The app POSTs each sign-up here. It's saved as a row in the "Requests" sheet
 *   and the admin gets an email with a link to approve or decline it.
 * - The approve link opens a confirmation page. Approving marks the row approved
 *   and sends the member their official welcome email from this Google account.
 * - The app can ask for a request's status (pending / approved / declined).
 * - Members log in on any device with LinkedIn: the app sends the LinkedIn sign-in code here,
 *   we look up their verified LinkedIn email, and return their profile if they're a member.
 *   As a fallback they can log in with a 6-digit code sent to their email.
 *
 * Setup: see backend/README.md.
 */

const SHEET_NAME = 'Requests';
const COLUMNS = ['id', 'status', 'submittedAt', 'decidedAt', 'first', 'last', 'email', 'headline', 'school', 'classYear',
  'location', 'industry', 'phone', 'linkedin', 'source', 'answer', 'about', 'career', 'token'];
const NETWORK_NAME = 'Israeli MBA Network';

/* Admin email: the ADMIN_EMAIL script property, or the account that deployed the script. */
function adminEmail_() {
  return PropertiesService.getScriptProperties().getProperty('ADMIN_EMAIL') || Session.getEffectiveUser().getEmail();
}

function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) {
    sh = ss.insertSheet(SHEET_NAME);
    sh.appendRow(COLUMNS);
    sh.setFrozenRows(1);
  }
  return sh;
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function findRow_(id) {
  const sh = sheet_();
  const values = sh.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (values[i][0] === id) {
      const rec = {};
      COLUMNS.forEach((c, j) => rec[c] = values[i][j]);
      return { sh, row: i + 1, rec };
    }
  }
  return null;
}

function esc_(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
}

/* ---------- sign-up ---------- */
function doPost(e) {
  let d;
  try { d = JSON.parse(e.postData.contents); } catch (err) { return json_({ ok: false, error: 'Bad request' }); }
  if (d.action === 'login-code') return sendLoginCode_(d.email);
  if (d.action === 'login') return login_(d.email, d.code);
  if (d.action === 'linkedin-login') return linkedinLogin_(d.code, d.redirectUri);
  const need = ['first', 'last', 'email', 'headline', 'school', 'classYear'];
  const missing = need.filter(k => !String(d[k] || '').trim());
  if (missing.length) return json_({ ok: false, error: 'Missing: ' + missing.join(', ') });
  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(d.email)) return json_({ ok: false, error: 'Invalid email' });

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const id = Utilities.getUuid();
    const token = Utilities.getUuid().replace(/-/g, '');
    const rec = {
      id, status: 'pending', submittedAt: new Date(), decidedAt: '',
      first: d.first, last: d.last, email: d.email, headline: d.headline, school: d.school, classYear: d.classYear,
      location: d.location || '', industry: d.industry || '', phone: d.phone || '', linkedin: d.linkedin || '',
      source: d.source || 'manual', answer: d.answer || '', about: d.about || '',
      career: JSON.stringify(d.career || []), token,
    };
    sheet_().appendRow(COLUMNS.map(c => rec[c]));
    sendAdminEmail_(rec);
    return json_({ ok: true, id });
  } finally {
    lock.releaseLock();
  }
}

function sendAdminEmail_(r) {
  const link = ScriptApp.getService().getUrl() + '?action=review&id=' + encodeURIComponent(r.id) + '&token=' + r.token;
  const career = JSON.parse(r.career || '[]');
  const li = r.linkedin ? (/^https?:/.test(r.linkedin) ? r.linkedin : 'https://' + r.linkedin) : '';
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:14px;color:#161A26;max-width:560px">
      <h2 style="margin:0 0 4px">New request to join ${NETWORK_NAME}</h2>
      <p style="color:#5C6272;margin:0 0 16px">Signed up ${r.source === 'linkedin' ? 'with their LinkedIn PDF' : 'manually'}.</p>
      <table style="border-collapse:collapse;width:100%">
        ${[['Name', r.first + ' ' + r.last], ['Email', r.email], ['Headline', r.headline], ['MBA', r.school + ', class of ' + r.classYear],
           ['Location', r.location], ['Industry', r.industry], ['WhatsApp', r.phone]]
          .filter(([, v]) => v).map(([k, v]) => `<tr><td style="padding:6px 12px 6px 0;color:#5C6272;vertical-align:top">${k}</td><td style="padding:6px 0">${esc_(v)}</td></tr>`).join('')}
        ${li ? `<tr><td style="padding:6px 12px 6px 0;color:#5C6272">LinkedIn</td><td style="padding:6px 0"><a href="${esc_(li)}">${esc_(r.linkedin)}</a></td></tr>` : ''}
      </table>
      ${r.answer ? `<p style="margin:16px 0 4px;color:#5C6272">Where in Israel are you from, and what did you do before your MBA?</p><p style="margin:0;background:#F7F8F4;padding:10px 12px;border-radius:8px">${esc_(r.answer)}</p>` : ''}
      ${career.length ? `<p style="margin:16px 0 4px;color:#5C6272">Career</p><ul style="margin:0;padding-left:18px">${career.map(c => `<li>${esc_(c.role)} · ${esc_(c.company)} (${esc_(c.years)})</li>`).join('')}</ul>` : ''}
      <p style="margin:24px 0"><a href="${link}" style="background:#2E49C2;color:#fff;padding:12px 20px;border-radius:10px;text-decoration:none;font-weight:bold">Review and approve</a></p>
      <p style="color:#8C91A0;font-size:12px">The link opens a page where you confirm Approve or Decline. All requests are in the "${SHEET_NAME}" sheet.</p>
    </div>`;
  MailApp.sendEmail({
    to: adminEmail_(), subject: `Join request: ${r.first} ${r.last} (${r.school} '${String(r.classYear).slice(-2)})`,
    htmlBody: html, name: NETWORK_NAME, replyTo: r.email,
  });
}

/* ---------- log in with an emailed code ---------- */
const CODE_TTL_SECONDS = 600;
const MAX_CODE_TRIES = 5;

/* The member's most recent request, matched by email. */
function findByEmail_(email) {
  const want = String(email || '').trim().toLowerCase();
  if (!want) return null;
  const values = sheet_().getDataRange().getValues();
  const col = COLUMNS.indexOf('email');
  for (let i = values.length - 1; i >= 1; i--) {
    if (String(values[i][col]).trim().toLowerCase() === want) {
      const rec = {};
      COLUMNS.forEach((c, j) => rec[c] = values[i][j]);
      return rec;
    }
  }
  return null;
}

function sendLoginCode_(email) {
  const r = findByEmail_(email);
  if (!r || r.status === 'declined') return json_({ ok: false, error: 'not-member' });
  const code = String(Math.floor(100000 + Math.random() * 900000));
  const key = 'login:' + r.email.toLowerCase();
  CacheService.getScriptCache().put(key, JSON.stringify({ code, tries: 0 }), CODE_TTL_SECONDS);
  MailApp.sendEmail({
    to: r.email, subject: `${code} is your ${NETWORK_NAME} login code`, name: NETWORK_NAME,
    htmlBody: `<div style="font-family:Arial,sans-serif;font-size:15px;color:#161A26;max-width:480px">
      <p>Hi ${esc_(r.first)},</p><p>Your login code is:</p>
      <p style="font-size:28px;font-weight:bold;letter-spacing:4px;margin:8px 0 16px">${code}</p>
      <p style="color:#5C6272">It expires in 10 minutes. If you didn't try to log in, you can ignore this email.</p></div>`,
  });
  return json_({ ok: true });
}

function login_(email, code) {
  const r = findByEmail_(email);
  if (!r) return json_({ ok: false, error: 'not-member' });
  const cache = CacheService.getScriptCache(), key = 'login:' + r.email.toLowerCase();
  const saved = JSON.parse(cache.get(key) || 'null');
  if (!saved) return json_({ ok: false, error: 'expired' });
  if (saved.code !== String(code || '').trim()) {
    saved.tries++;
    if (saved.tries >= MAX_CODE_TRIES) cache.remove(key); else cache.put(key, JSON.stringify(saved), CODE_TTL_SECONDS);
    return json_({ ok: false, error: saved.tries >= MAX_CODE_TRIES ? 'expired' : 'wrong-code' });
  }
  cache.remove(key);
  return json_({ ok: true, profile: profile_(r) });
}

/* What the app gets back after a successful login. Never includes the review token. */
function profile_(r) {
  const profile = {};
  ['id', 'status', 'first', 'last', 'email', 'headline', 'school', 'classYear', 'location', 'industry', 'phone', 'linkedin', 'about']
    .forEach(k => profile[k] = r[k]);
  try { profile.career = JSON.parse(r.career || '[]'); } catch (err) { profile.career = []; }
  return profile;
}

/* ---------- log in with LinkedIn (OpenID Connect) ----------
 * Needs the script properties LINKEDIN_CLIENT_ID and LINKEDIN_CLIENT_SECRET (see README).
 * LinkedIn only accepts redirectUri if it's one of the app's registered redirect URLs. */
function linkedinLogin_(code, redirectUri) {
  const props = PropertiesService.getScriptProperties();
  const id = props.getProperty('LINKEDIN_CLIENT_ID'), secret = props.getProperty('LINKEDIN_CLIENT_SECRET');
  if (!id || !secret) return json_({ ok: false, error: 'not-configured' });
  const tok = UrlFetchApp.fetch('https://www.linkedin.com/oauth/v2/accessToken', {
    method: 'post', muteHttpExceptions: true,
    payload: { grant_type: 'authorization_code', code: String(code || ''), redirect_uri: String(redirectUri || ''), client_id: id, client_secret: secret },
  });
  if (tok.getResponseCode() !== 200) return json_({ ok: false, error: 'linkedin-failed' });
  const info = UrlFetchApp.fetch('https://api.linkedin.com/v2/userinfo', {
    headers: { Authorization: 'Bearer ' + JSON.parse(tok.getContentText()).access_token }, muteHttpExceptions: true,
  });
  if (info.getResponseCode() !== 200) return json_({ ok: false, error: 'linkedin-failed' });
  const u = JSON.parse(info.getContentText());
  if (!u.email || u.email_verified === false) return json_({ ok: false, error: 'no-email' });
  const r = findByEmail_(u.email);
  if (!r || r.status === 'declined') return json_({ ok: false, error: 'not-member', email: u.email });
  const profile = profile_(r);
  profile.photo = u.picture || '';
  return json_({ ok: true, profile });
}

/* ---------- admin review page + status ---------- */
function doGet(e) {
  const p = e.parameter || {};
  if (p.action === 'status') {
    const f = p.id && findRow_(p.id);
    return json_(f ? { ok: true, status: f.rec.status } : { ok: false, error: 'Not found' });
  }
  if (p.action === 'review') {
    const f = p.id && findRow_(p.id);
    if (!f || f.rec.token !== p.token) return HtmlService.createHtmlOutput('<p style="font-family:Arial">This link is not valid.</p>');
    const t = HtmlService.createTemplate(REVIEW_PAGE);
    t.r = f.rec; t.token = p.token; t.network = NETWORK_NAME;
    return t.evaluate().setTitle('Review join request');
  }
  return HtmlService.createHtmlOutput('<p style="font-family:Arial">' + NETWORK_NAME + ' sign-up backend is running.</p>');
}

/* Called from the review page's buttons. */
function decide(id, token, approve) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const f = findRow_(id);
    if (!f || f.rec.token !== token) return 'This link is not valid.';
    if (f.rec.status !== 'pending') return `Already ${f.rec.status}.`;
    const status = approve ? 'approved' : 'declined';
    f.sh.getRange(f.row, COLUMNS.indexOf('status') + 1).setValue(status);
    f.sh.getRange(f.row, COLUMNS.indexOf('decidedAt') + 1).setValue(new Date());
    if (approve) sendWelcomeEmail_(f.rec);
    return approve ? `Approved. ${f.rec.first} is now an official member and their welcome email is on its way to ${f.rec.email}.`
                   : `Declined. ${f.rec.first} was not notified.`;
  } finally {
    lock.releaseLock();
  }
}

function sendWelcomeEmail_(r) {
  const html = `
    <div style="font-family:Arial,sans-serif;font-size:15px;line-height:1.5;color:#161A26;max-width:560px">
      <h2 style="margin:0 0 12px">Welcome to ${NETWORK_NAME}, ${esc_(r.first)}!</h2>
      <p>Your request to join has been approved. You're now an official member of ${NETWORK_NAME}, the network of Israeli MBA students and alumni in the US.</p>
      <p style="margin-bottom:4px"><b>What you can do now:</b></p>
      <ul style="margin-top:0">
        <li>Search for Israeli MBAs by school, company or industry</li>
        <li>Connect with members and ask for warm intros</li>
        <li>Keep your profile current, so others can find you</li>
      </ul>
      <p>To log in on any device, choose <b>Log in</b> in the app and enter ${esc_(r.email)}. We'll email you a code.</p>
      <p>Welcome aboard,<br>The ${NETWORK_NAME} team</p>
    </div>`;
  MailApp.sendEmail({
    to: r.email, subject: `You're in: welcome to ${NETWORK_NAME}`,
    htmlBody: html, name: NETWORK_NAME, replyTo: adminEmail_(),
  });
}

const REVIEW_PAGE = `<!doctype html><html><head><base target="_top"><meta name="viewport" content="width=device-width,initial-scale=1">
<style>body{font-family:Arial,sans-serif;background:#F1F2EE;color:#161A26;margin:0;padding:24px 16px}
.card{max-width:520px;margin:0 auto;background:#fff;border:1px solid #E1E3DC;border-radius:16px;padding:24px}
h2{margin:0 0 4px}.muted{color:#5C6272}td{padding:5px 12px 5px 0;vertical-align:top}td:first-child{color:#5C6272}
.row{display:flex;gap:10px;margin-top:20px}button{font:inherit;font-weight:bold;border:0;border-radius:10px;padding:12px 18px;cursor:pointer}
.ok{background:#2E49C2;color:#fff}.no{background:#F7F8F4;border:1px solid #E1E3DC}#msg{margin-top:16px;font-weight:bold}</style></head>
<body><div class="card" id="card" data-id="<?= r.id ?>" data-token="<?= token ?>">
<h2>Join request</h2><p class="muted">Status: <b><?= r.status ?></b></p>
<table><tr><td>Name</td><td><?= r.first ?> <?= r.last ?></td></tr><tr><td>Email</td><td><?= r.email ?></td></tr>
<tr><td>Headline</td><td><?= r.headline ?></td></tr><tr><td>MBA</td><td><?= r.school ?>, class of <?= r.classYear ?></td></tr>
<tr><td>LinkedIn</td><td><?= r.linkedin ?></td></tr><tr><td>Answer</td><td><?= r.answer ?></td></tr></table>
<? if (r.status === 'pending') { ?>
<div class="row" id="btns"><button class="ok" onclick="go(true)">Approve and send welcome email</button><button class="no" onclick="go(false)">Decline</button></div>
<? } ?>
<div id="msg"></div></div>
<script>
function go(ok){ document.getElementById('btns').style.display='none'; document.getElementById('msg').textContent='Saving…';
  google.script.run.withSuccessHandler(function(m){ document.getElementById('msg').textContent=m; })
    .withFailureHandler(function(e){ document.getElementById('msg').textContent='Something went wrong: '+e.message; })
    .decide(document.getElementById('card').dataset.id, document.getElementById('card').dataset.token, ok); }
</script></body></html>`;
