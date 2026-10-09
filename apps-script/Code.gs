/**
 * SCU Partnership Form -> Email
 * Google Apps Script web app. Receives the "Partner with SCU" form from the
 * website and emails a formatted summary to the SCU board inbox.
 *
 * Uploaded flyers (PDF/JPG/PNG, max 5 MB) are re-checked here and saved to a
 * private Google Drive folder; the email links to them instead of attaching.
 *
 * Setup (one time):
 *  1. Go to https://script.google.com -> New project -> paste this file.
 *  2. Deploy -> New deployment -> type "Web app"
 *       Execute as: Me      Who has access: Anyone
 *  3. Authorize when asked, then copy the Web app URL (ends in /exec).
 *  4. Paste that URL into PARTNER_FORM_ENDPOINT in index.html.
 */

var TO_EMAIL = 'scuboard@mcw.edu';
var LOGO_URL = 'https://jamessrn.github.io/scu-website/img/scu-logo.png';
var SITE_URL = 'https://jamessrn.github.io/scu-website/';
var FLYER_FOLDER_NAME = 'SCU Partnership Flyers';   // private Drive folder, created automatically
var MAX_FLYER_BYTES = 5 * 1024 * 1024;              // 5 MB
var ALLOWED_TYPES = { 'application/pdf': 'pdf', 'image/png': 'png', 'image/jpeg': 'jpg' };

function doPost(e) {
  var d = {};
  try { d = JSON.parse(e.postData.contents); } catch (err) { d = e.parameter || {}; }

  // Basic spam protection: hidden "website" field must stay empty.
  if (d.website) return json_({ ok: true });
  if (!d.name || !d.org || !d.email || !d.idea) return json_({ ok: false, error: 'missing fields' });

  // Light rate limit: max 5 submissions per email address per hour.
  var cache = CacheService.getScriptCache();
  var key = 'rl_' + String(d.email).toLowerCase();
  var n = Number(cache.get(key) || 0);
  if (n >= 5) return json_({ ok: false, error: 'rate limited' });
  cache.put(key, String(n + 1), 3600);

  var flyer = null, inline = {};
  if (d.fileData) {
    flyer = saveFlyer_(d);                 // returns null if the file fails checks
    if (flyer && flyer.isImage) inline.flyerthumb = flyer.blob;
    if (!flyer) d.flyerRejected = true;
  }
  d.flyer = flyer;
  d.flyerLink = safeUrl_(d.flyerLink);

  MailApp.sendEmail({
    to: TO_EMAIL,
    replyTo: d.email,
    name: 'SCU Website',
    subject: buildSubject(d),
    htmlBody: buildEmailHtml(d),
    body: buildPlainText(d),
    inlineImages: inline
  });
  return json_({ ok: true });
}

function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}


// ---- Flyer handling -------------------------------------------------------
function saveFlyer_(d) {
  var type = String(d.fileType || '').toLowerCase();
  if (!ALLOWED_TYPES[type]) return null;
  var bytes;
  try { bytes = Utilities.base64Decode(String(d.fileData).replace(/^data:[^,]*,/, '')); } catch (e) { return null; }
  if (!bytes.length || bytes.length > MAX_FLYER_BYTES) return null;
  if (!magicMatches_(bytes, type)) return null;  // file contents must really be PDF/PNG/JPG

  var safeName = String(d.fileName || 'flyer').replace(/[^\w.\- ]+/g, '').slice(0, 80) || 'flyer';
  safeName = safeName.replace(/\.[^.]*$/, '') + '.' + ALLOWED_TYPES[type];
  var stamp = Utilities.formatDate(new Date(), 'America/Chicago', 'yyyy-MM-dd');
  var blob = Utilities.newBlob(bytes, type, stamp + ' - ' + clean(d.org).slice(0, 60) + ' - ' + safeName);
  var file = getFlyerFolder_().createFile(blob);   // stays private to this Google account
  return { url: file.getUrl(), name: safeName, size: bytes.length, isImage: type !== 'application/pdf', blob: blob };
}

function magicMatches_(b, type) {
  function at(i) { return (b[i] + 256) % 256; }
  if (type === 'application/pdf') return at(0) === 0x25 && at(1) === 0x50 && at(2) === 0x44 && at(3) === 0x46; // %PDF
  if (type === 'image/png') return at(0) === 0x89 && at(1) === 0x50 && at(2) === 0x4E && at(3) === 0x47;
  if (type === 'image/jpeg') return at(0) === 0xFF && at(1) === 0xD8 && at(2) === 0xFF;
  return false;
}

function getFlyerFolder_() {
  var it = DriveApp.getFoldersByName(FLYER_FOLDER_NAME);
  return it.hasNext() ? it.next() : DriveApp.createFolder(FLYER_FOLDER_NAME);
}

function safeUrl_(u) {
  u = String(u || '').trim();
  return /^https:\/\/[^\s<>"]+$/i.test(u) ? u.slice(0, 500) : '';
}

function fmtSize_(n) { return n > 1048576 ? (n / 1048576).toFixed(1) + ' MB' : Math.round(n / 1024) + ' KB'; }

function buildSubject(d) {
  return '\uD83E\uDD1D New Partnership Interest: ' + clean(d.org) + ' (' + clean(d.name) + ')' + ((d.flyer || d.flyerLink) ? ' \u2014 flyer included' : '');
}

function esc(s) {
  return String(s == null ? '' : s).replace(/[&<>"']/g, function (c) {
    return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c];
  });
}
function clean(s) { return String(s == null ? '' : s).replace(/[\r\n]+/g, ' ').slice(0, 120); }

function row(label, value) {
  return '<tr><td style="padding:10px 0;border-bottom:1px solid #eef1f0;width:150px;vertical-align:top;font-size:13px;color:#6b7775;font-weight:600;">' + label +
    '</td><td style="padding:10px 0;border-bottom:1px solid #eef1f0;vertical-align:top;font-size:15px;color:#1d2b2a;">' + value + '</td></tr>';
}

function buildEmailHtml(d) {
  var when = d.submittedAt ? new Date(d.submittedAt) : new Date();
  var whenStr = Utilities && Utilities.formatDate
    ? Utilities.formatDate(when, 'America/Chicago', "EEEE, MMMM d, yyyy 'at' h:mm a z")
    : when.toString();
  var dash = '<span style="color:#9aa5a3;">&mdash;</span>';
  var emailLink = '<a href="mailto:' + esc(d.email) + '" style="color:#2f6e66;font-weight:600;text-decoration:none;">' + esc(d.email) + '</a>';
  var phoneLink = d.phone ? '<a href="tel:' + esc(String(d.phone).replace(/[^\d+]/g, '')) + '" style="color:#2f6e66;text-decoration:none;">' + esc(d.phone) + '</a>' : dash;
  var replyHref = 'mailto:' + encodeURIComponent(d.email) + '?subject=' +
    encodeURIComponent('Re: Partnering with the Saturday Clinic for the Uninsured');
  var first = esc(String(d.name).split(' ')[0]);

  return '' +
  '<!doctype html><html><body style="margin:0;padding:0;background:#f3f4f3;">' +
  '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#f3f4f3;padding:28px 12px;font-family:Helvetica,Arial,sans-serif;">' +
  '<tr><td align="center">' +
  '<table role="presentation" width="600" cellpadding="0" cellspacing="0" style="max-width:600px;width:100%;background:#ffffff;border-radius:10px;overflow:hidden;box-shadow:0 2px 10px rgba(0,0,0,.06);">' +

  // Header
  '<tr><td style="background:#2f6e66;padding:26px 32px 22px;border-bottom:4px solid #c79a3b;">' +
    '<table role="presentation" width="100%"><tr>' +
    '<td style="vertical-align:middle;"><div style="display:inline-block;background:rgba(255,255,255,.18);color:#ffffff;font-size:11px;font-weight:700;letter-spacing:1px;text-transform:uppercase;padding:4px 10px;border-radius:12px;">New partnership inquiry</div>' +
    '<div style="color:#ffffff;font-size:24px;font-weight:700;margin-top:10px;line-height:1.25;">' + esc(d.org) + '</div>' +
    '<div style="color:#d8ebe6;font-size:14px;margin-top:4px;">wants to partner with SCU</div></td>' +
    '<td width="84" align="right" style="vertical-align:middle;"><img src="' + LOGO_URL + '" width="76" alt="SCU" style="display:block;background:#ffffff;border-radius:10px;padding:6px;"></td>' +
    '</tr></table></td></tr>' +

  // Intro
  '<tr><td style="padding:24px 32px 6px;font-size:15px;color:#33403e;line-height:1.55;">' +
    '<strong>' + esc(d.name) + '</strong>' + (d.role ? ' (' + esc(d.role) + ')' : '') + ' from <strong>' + esc(d.org) + '</strong> submitted the <em>Partner with SCU</em> form on the clinic website.' +
  '</td></tr>' +

  // Contact info
  '<tr><td style="padding:12px 32px 4px;"><div style="font-size:12px;font-weight:700;color:#c79a3b;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px;">Contact</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
    row('Name', esc(d.name)) +
    row('Organization', esc(d.org)) +
    row('Role / Title', d.role ? esc(d.role) : dash) +
    row('Email', emailLink) +
    row('Phone', phoneLink) +
    '</table></td></tr>' +

  // Idea
  '<tr><td style="padding:18px 32px 6px;"><div style="font-size:12px;font-weight:700;color:#c79a3b;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">Partnership idea</div>' +
    '<div style="background:#f6f8f7;border-left:4px solid #2f6e66;border-radius:4px;padding:14px 16px;font-size:15px;color:#1d2b2a;line-height:1.6;white-space:pre-wrap;">' + esc(d.idea) + '</div>' +
  '</td></tr>' +

  // Flyer
  ((d.flyer || d.flyerLink || d.flyerRejected) ? '<tr><td style="padding:18px 32px 6px;"><div style="font-size:12px;font-weight:700;color:#c79a3b;letter-spacing:1px;text-transform:uppercase;margin-bottom:8px;">Flyer</div>' +
    (d.flyer ? '<table role="presentation" cellpadding="0" cellspacing="0" style="border:1px solid #e1e4e3;border-radius:8px;width:100%;"><tr>' +
        (d.flyer.isImage ? '<td width="96" style="padding:10px;"><img src="cid:flyerthumb" width="84" style="display:block;border-radius:4px;border:1px solid #e1e4e3;"></td>'
                         : '<td width="96" align="center" style="padding:10px;"><div style="width:64px;height:80px;border-radius:4px;background:#712b6b;color:#fff;font-weight:700;font-size:14px;line-height:80px;text-align:center;">PDF</div></td>') +
        '<td style="padding:10px 14px 10px 4px;vertical-align:middle;"><div style="font-size:14px;color:#1d2b2a;font-weight:600;">' + esc(d.flyer.name) + '</div>' +
        '<div style="font-size:12px;color:#6b7775;margin:2px 0 8px;">' + fmtSize_(d.flyer.size) + ' &middot; saved to the &ldquo;' + FLYER_FOLDER_NAME + '&rdquo; Drive folder</div>' +
        '<a href="' + esc(d.flyer.url) + '" style="display:inline-block;background:#ffffff;border:1.5px solid #2f6e66;color:#2f6e66;font-size:13px;font-weight:700;text-decoration:none;padding:7px 14px;border-radius:18px;">&#128206; View flyer</a></td></tr></table>' : '') +
    (d.flyerLink ? '<div style="font-size:14px;margin-top:10px;color:#33403e;">Link they shared: <a href="' + esc(d.flyerLink) + '" style="color:#2f6e66;font-weight:600;">' + esc(d.flyerLink) + '</a></div>' : '') +
    (d.flyerRejected ? '<div style="font-size:13px;margin-top:8px;color:#b3261e;">An uploaded file was blocked because it wasn&rsquo;t a valid PDF, JPG or PNG under 5 MB.</div>' : '') +
  '</td></tr>' : '') +

  // Reply button
  '<tr><td align="center" style="padding:22px 32px 8px;">' +
    '<a href="' + replyHref + '" style="display:inline-block;background:#2f6e66;color:#ffffff;font-size:15px;font-weight:700;text-decoration:none;padding:12px 28px;border-radius:24px;">Reply to ' + first + ' &rarr;</a>' +
  '</td></tr>' +

  // Submission details
  '<tr><td style="padding:18px 32px 22px;"><div style="font-size:12px;font-weight:700;color:#c79a3b;letter-spacing:1px;text-transform:uppercase;margin-bottom:4px;">Submission details</div>' +
    '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="font-size:13px;">' +
    row('Submitted', esc(whenStr)) +
    row('Their time zone', d.timezone ? esc(d.timezone) : dash) +
    row('Device', d.device ? esc(d.device) : dash) +
    row('Browser language', d.language ? esc(d.language) : dash) +
    row('Came from', d.referrer ? esc(d.referrer) : 'Direct visit') +
    row('Form page', d.page ? '<a href="' + esc(d.page) + '" style="color:#2f6e66;">' + esc(d.page) + '</a>' : dash) +
    '</table></td></tr>' +

  // Footer
  '<tr><td style="background:#f6f8f7;padding:16px 32px;font-size:12px;color:#6b7775;line-height:1.5;text-align:center;">' +
    'Sent automatically from the <a href="' + SITE_URL + '" style="color:#2f6e66;">Saturday Clinic for the Uninsured website</a>.<br>Replying to this email goes directly to ' + esc(d.email) + '.' +
  '</td></tr>' +

  '</table></td></tr></table></body></html>';
}

function buildPlainText(d) {
  return 'New partnership inquiry from the SCU website\n\n' +
    'Name: ' + d.name + '\nOrganization: ' + d.org + '\nRole: ' + (d.role || '-') +
    '\nEmail: ' + d.email + '\nPhone: ' + (d.phone || '-') + '\n\nPartnership idea:\n' + d.idea +
    (d.flyer ? '\n\nFlyer: ' + d.flyer.url : '') + (d.flyerLink ? '\nFlyer link: ' + d.flyerLink : '') + '\n\nSubmitted: ' + (d.submittedAt || '') + '\nTime zone: ' + (d.timezone || '-') +
    '\nDevice: ' + (d.device || '-') + '\nCame from: ' + (d.referrer || 'Direct visit') + '\n';
}

// Run this once from the Apps Script editor to send yourself a test email.
function sendTestEmail() {
  doPost({ postData: { contents: JSON.stringify({
    name: 'Maria Lopez', org: 'Neighborhood House of Milwaukee', role: 'Community Programs Director',
    email: 'mlopez@example.org', phone: '(414) 555-0142',
    idea: 'We host family programming on the north side and would love to bring SCU volunteers for a monthly blood-pressure screening and health-education table.',
    submittedAt: new Date().toISOString(), timezone: 'America/Chicago', device: 'iPhone · Safari',
    language: 'en-US', referrer: 'instagram.com', page: SITE_URL
  }) } });
}
