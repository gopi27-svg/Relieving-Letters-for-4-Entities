/****************************************************************
 *  Relieving Letter Generator – Render server (Node.js)
 *  - Serves public/index.html
 *  - Asks Google Apps Script to generate the PDF (Drive) and log (Sheet)
 *  - Sends the mail from the Microsoft 365 mailbox
 *      MAIL_MODE=smtp   -> Outlook SMTP with password (default; needs paid Render plan)
 *      MAIL_MODE=graph  -> Microsoft Graph API app (optional; works on free plan)
 *  All settings come from Render → Environment.
 ****************************************************************/
const express = require('express');
const path = require('path');
const crypto = require('crypto');
const nodemailer = require('nodemailer');

const env = k => String(process.env[k] || '').trim();
const splitList = s => String(s || '').split(/[,;\s]+/).map(x => x.trim().toLowerCase()).filter(Boolean);

const CFG = {
  port:         process.env.PORT || 3000,
  appKey:       env('APP_ACCESS_KEY'),
  gasUrl:       env('GAS_URL'),
  gasKey:       env('GAS_KEY'),
  mode:         (env('MAIL_MODE') || 'smtp').toLowerCase(),
  from:         env('MAIL_FROM').toLowerCase(),
  fromName:     env('MAIL_FROM_NAME') || 'HR - USDC Global',
  defaultCc:    splitList(env('DEFAULT_CC')),
  // Graph
  tenantId:     env('MS_TENANT_ID'),
  clientId:     env('MS_CLIENT_ID'),
  clientSecret: env('MS_CLIENT_SECRET'),
  // SMTP
  smtpHost:     env('SMTP_HOST') || 'smtp.office365.com',
  smtpPort:     Number(env('SMTP_PORT') || 587),
  smtpUser:     env('SMTP_USER') || env('MAIL_FROM'),
  smtpPass:     process.env.SMTP_PASS || ''
};

function missingConfig() {
  const need = ['APP_ACCESS_KEY', 'GAS_URL', 'GAS_KEY', 'MAIL_FROM'];
  if (CFG.mode === 'smtp') need.push('SMTP_PASS');
  else need.push('MS_TENANT_ID', 'MS_CLIENT_ID', 'MS_CLIENT_SECRET');
  return need.filter(k => !process.env[k] || !String(process.env[k]).trim());
}

const sleep = ms => new Promise(r => setTimeout(r, ms));
const isEmail = s => /^[^\s@]+@[^\s@]+\.[^\s@]{2,}$/.test(String(s || '').trim());
const esc = s => String(s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');

/* ================= APP ================= */
const app = express();
app.use(express.json({ limit: '1mb' }));
app.use(express.static(path.join(__dirname, 'public')));
app.get('/healthz', (req, res) => res.send('ok'));

// Serve the page whether index.html is in /public, /Public or the repo root
const fs = require('fs');
const PAGE = ['public/index.html', 'Public/index.html', 'index.html']
  .map(p => path.join(__dirname, p)).find(p => fs.existsSync(p));
app.get('/', (req, res) => PAGE
  ? res.sendFile(PAGE)
  : res.status(404).send('index.html not found – upload it to the repo (root or a "public" folder)'));

function checkKey(req, res, next) {
  if (!CFG.appKey) return res.status(500).json({ ok: false, error: 'Server not configured: APP_ACCESS_KEY is missing' });
  const given = String(req.get('X-Access-Key') || '');
  const a = crypto.createHash('sha256').update(given).digest();
  const b = crypto.createHash('sha256').update(CFG.appKey).digest();
  if (!crypto.timingSafeEqual(a, b)) return res.status(401).json({ ok: false, error: 'Invalid access key' });
  next();
}

/* ================= GOOGLE APPS SCRIPT ================= */
async function callGas(action, data) {
  const ctrl = new AbortController();
  const timer = setTimeout(() => ctrl.abort(), 90000);
  try {
    const res = await fetch(CFG.gasUrl, {
      method: 'POST',
      headers: { 'Content-Type': 'text/plain;charset=utf-8' },
      body: JSON.stringify(Object.assign({}, data, { action, key: CFG.gasKey })),
      redirect: 'follow',
      signal: ctrl.signal
    });
    const txt = await res.text();
    try { return JSON.parse(txt); }
    catch (e) { throw new Error('Google Apps Script gave an unexpected response – check GAS_URL and that the deployment access is "Anyone"'); }
  } catch (e) {
    if (e.name === 'AbortError') throw new Error('Google Apps Script took too long to respond');
    throw e;
  } finally { clearTimeout(timer); }
}

/* ================= MAIL CONTENT ================= */
function buildMail(entity, r, cc, pdf) {
  const n = esc(r.name), en = esc(entity);
  const html =
    'Dear ' + n + ',<br><br>' +
    'Please find the attached ' + en + ' Experience Letter for your reference.<br><br>' +
    'We appreciate your contributions during your time with ' + en + '. ' +
    'We wish you all the best in your future endeavours.<br><br>' +
    'For any challenges feel free to reach out to us.<br><br>' +
    'Regards,<br>HR Team<br>' + en;
  const text = html.replace(/<br>/g, '\n').replace(/&amp;/g, '&').replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&quot;/g, '"');
  return {
    to: r.email, cc,
    subject: 'Relieving Letter/Experience Letter - ' + entity,
    html, text,
    fileName: pdf.fileName, pdfBase64: pdf.base64
  };
}

/* ================= MICROSOFT GRAPH ================= */
let tokenCache = { token: null, exp: 0 };

async function graphToken() {
  if (tokenCache.token && Date.now() < tokenCache.exp - 60000) return tokenCache.token;
  const body = new URLSearchParams({
    client_id: CFG.clientId, client_secret: CFG.clientSecret,
    scope: 'https://graph.microsoft.com/.default', grant_type: 'client_credentials'
  });
  const res = await fetch('https://login.microsoftonline.com/' + encodeURIComponent(CFG.tenantId) + '/oauth2/v2.0/token', { method: 'POST', body });
  const j = await res.json().catch(() => ({}));
  if (!res.ok || !j.access_token) {
    const d = String(j.error_description || j.error || ('HTTP ' + res.status)).split(/\r?\n/)[0];
    throw new Error('Microsoft 365 sign-in failed (check MS_TENANT_ID / MS_CLIENT_ID / MS_CLIENT_SECRET): ' + d);
  }
  tokenCache = { token: j.access_token, exp: Date.now() + Number(j.expires_in || 3600) * 1000 };
  return tokenCache.token;
}

function graphReason(status, j) {
  const code = (j.error && j.error.code) || '';
  const msg = (j.error && j.error.message) || ('HTTP ' + status);
  if (status === 401) { tokenCache = { token: null, exp: 0 }; return 'Microsoft 365 rejected the app login: ' + msg; }
  if (status === 403 || code === 'ErrorAccessDenied')
    return 'Microsoft 365 denied sending as ' + CFG.from + ' (Mail.Send permission / admin consent missing, or app not allowed for this mailbox): ' + msg;
  if (status === 404 || code === 'ErrorInvalidUser' || code === 'MailboxNotEnabledForRESTAPI')
    return 'Mailbox ' + CFG.from + ' not found or not enabled in Microsoft 365: ' + msg;
  if (code === 'ErrorInvalidRecipients') return 'Invalid recipient address: ' + msg;
  if (status === 429) return 'Microsoft 365 is throttling sends – wait a minute and retry';
  return msg;
}

async function sendGraph(m) {
  const message = {
    subject: m.subject,
    body: { contentType: 'HTML', content: m.html },
    toRecipients: [{ emailAddress: { address: m.to } }],
    ccRecipients: m.cc.map(a => ({ emailAddress: { address: a } })),
    attachments: [{
      '@odata.type': '#microsoft.graph.fileAttachment',
      name: m.fileName, contentType: 'application/pdf', contentBytes: m.pdfBase64
    }]
  };
  for (let attempt = 1; attempt <= 3; attempt++) {
    const token = await graphToken();
    const res = await fetch('https://graph.microsoft.com/v1.0/users/' + encodeURIComponent(CFG.from) + '/sendMail', {
      method: 'POST',
      headers: { Authorization: 'Bearer ' + token, 'Content-Type': 'application/json' },
      body: JSON.stringify({ message, saveToSentItems: true })
    });
    if (res.status === 202) return;
    const j = await res.json().catch(() => ({}));
    const retryable = res.status === 429 || res.status >= 500 || (res.status === 401 && attempt === 1);
    if (retryable && attempt < 3) {
      if (res.status === 401) tokenCache = { token: null, exp: 0 };
      const wait = Number(res.headers.get('retry-after')) * 1000 || attempt * 3000;
      await sleep(Math.min(wait, 20000));
      continue;
    }
    throw new Error(graphReason(res.status, j));
  }
}

/* ================= OUTLOOK SMTP ================= */
let transporter = null;
function getTransport() {
  if (!transporter) {
    transporter = nodemailer.createTransport({
      host: CFG.smtpHost, port: CFG.smtpPort,
      secure: CFG.smtpPort === 465, requireTLS: CFG.smtpPort !== 465,
      auth: { user: CFG.smtpUser, pass: CFG.smtpPass },
      tls: { minVersion: 'TLSv1.2' },
      connectionTimeout: 20000, greetingTimeout: 20000, socketTimeout: 45000
    });
  }
  return transporter;
}

function smtpReason(e) {
  const s = String(e.response || e.message || e);
  if (/5\.7\.139|SmtpClientAuthentication is disabled|basic authentication is disabled/i.test(s))
    return 'SMTP sign-in with password is disabled for this mailbox/tenant – ask the Microsoft 365 admin to enable Authenticated SMTP for ' + CFG.from + ', or switch to MAIL_MODE=graph';
  if (/535|5\.7\.3|Authentication unsuccessful|Invalid login/i.test(s))
    return 'Login failed – wrong SMTP_PASS, MFA/security defaults blocking password login, or SMTP AUTH disabled (' + s.slice(0, 150) + ')';
  if (['ETIMEDOUT', 'ECONNECTION', 'ESOCKET', 'ECONNREFUSED', 'ENETUNREACH'].includes(e.code) || /Greeting never received|Connection timeout/i.test(s))
    return 'Cannot reach ' + CFG.smtpHost + ':' + CFG.smtpPort + ' – Render free plan blocks SMTP ports. Use MAIL_MODE=graph or a paid Render instance';
  if (/5\.7\.60|SendAsDenied/i.test(s)) return CFG.smtpUser + ' is not allowed to send as ' + CFG.from;
  if (/^5\.1\.|Recipient/i.test(s)) return 'Recipient rejected: ' + s.slice(0, 160);
  return s.slice(0, 250);
}

async function sendSmtp(m) {
  try {
    await getTransport().sendMail({
      from: { name: CFG.fromName, address: CFG.from },
      to: m.to, cc: m.cc.length ? m.cc : undefined,
      subject: m.subject, html: m.html, text: m.text,
      attachments: [{ filename: m.fileName, content: Buffer.from(m.pdfBase64, 'base64'), contentType: 'application/pdf' }]
    });
  } catch (e) {
    transporter = null;
    throw new Error(smtpReason(e));
  }
}

/* ================= HELPERS ================= */
function cleanRow(row) {
  row = row || {};
  const t = v => String(v == null ? '' : v).replace(/\s+/g, ' ').trim();
  return { name: t(row.name), designation: t(row.designation), doj: t(row.doj), doe: t(row.doe), email: t(row.email).toLowerCase() };
}
function validateRow(r, cc) {
  const p = [];
  if (!r.name) p.push('Employee Full Name is missing');
  if (!r.designation) p.push('Designation is missing');
  if (!r.doj) p.push('Date of Joining is missing');
  if (!r.doe) p.push('Date of Exit is missing');
  if (!r.email) p.push('Personal Email ID is missing');
  else if (!isEmail(r.email)) p.push('Invalid Personal Email ID "' + r.email + '"');
  cc.forEach(c => { if (!isEmail(c)) p.push('Invalid CC email "' + c + '"'); });
  return p;
}

/* ================= API ================= */
app.post('/api/ping', checkKey, async (req, res) => {
  const miss = missingConfig();
  if (miss.length) return res.json({ ok: false, error: 'Server not configured – missing Render environment variable(s): ' + miss.join(', ') });
  try {
    const g = await callGas('ping', {});
    if (!g.ok) return res.json({ ok: false, error: 'Google Apps Script: ' + g.error });
    // check the mail login right away, so Connect shows any problem before sending
    if (CFG.mode === 'smtp') {
      try { await getTransport().verify(); }
      catch (e) { transporter = null; throw new Error(smtpReason(e)); }
    } else {
      await graphToken();
    }
    res.json({ ok: true, sender: CFG.from, senderName: CFG.fromName, mode: CFG.mode, defaultCc: CFG.defaultCc });
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.post('/api/preview', checkKey, async (req, res) => {
  try {
    const r = cleanRow(req.body.row);
    const p = validateRow(r, []);
    if (p.length) return res.json({ ok: false, error: p.join('; ') });
    res.json(await callGas('preview', { entity: req.body.entity, row: r }));
  } catch (e) {
    res.json({ ok: false, error: e.message });
  }
});

app.post('/api/send', checkKey, async (req, res) => {
  const entity = String(req.body.entity || '').trim();
  const mode = req.body.mode === 'Bulk' ? 'Bulk' : 'Single';
  const cc = [...new Set(splitList(Array.isArray(req.body.cc) ? req.body.cc.join(',') : req.body.cc))];
  const r = cleanRow(req.body.row);
  let pdf = {}, status = 'Failed', reason = '';

  try {
    const miss = missingConfig();
    if (miss.length) throw new Error('Server not configured – missing: ' + miss.join(', '));
    const p = validateRow(r, cc);
    if (p.length) throw new Error(p.join('; '));

    const g = await callGas('generate', { entity, row: r });
    if (!g.ok) throw new Error(g.error || 'PDF generation failed');
    pdf = g;

    const m = buildMail(entity, r, cc, g);
    if (CFG.mode === 'smtp') await sendSmtp(m); else await sendGraph(m);
    status = 'Sent';
  } catch (e) {
    reason = e.message || String(e);
  }

  try {
    await callGas('log', {
      entity, mode, row: r, cc: cc.join(','), status, reason,
      pdfName: pdf.fileName || '', pdfUrl: pdf.pdfUrl || ''
    });
  } catch (e) { console.error('Log failed:', e.message); }

  res.json(status === 'Sent'
    ? { ok: true, status, pdfUrl: pdf.pdfUrl, pdfName: pdf.fileName }
    : { ok: false, status, error: reason, pdfUrl: pdf.pdfUrl || '' });
});

app.listen(CFG.port, () => {
  const miss = missingConfig();
  console.log('Relieving Letter server on port ' + CFG.port + ' | mail mode: ' + CFG.mode);
  if (miss.length) console.warn('Missing environment variables: ' + miss.join(', '));
});
