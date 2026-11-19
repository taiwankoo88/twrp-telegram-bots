// Google Apps Script web app behind the static order page (site/index.html on GitHub Pages).
// The page builds the Excel file in the browser and posts it here; this script emails it to the administrator
// and lists the items in the Google Sheet the script is attached to.
//
// Set-up (once, by the administrator, in their own Google account):
//   1. Create a Google Sheet (for example "TWRP Web Orders"), then Extensions > Apps Script.
//   2. Replace the sample code with this file and put the administrator's email address in TO below.
//   3. Deploy > New deployment > type "Web app"; Execute as: Me; Who has access: Anyone. Authorize when asked.
//   4. Copy the web app URL (ends in /exec) into site/config.js as MAIL_URL.
// After changing this code: Deploy > Manage deployments > edit > Version: New version (the URL stays the same).

const TO = 'PUT-YOUR-EMAIL-ADDRESS-HERE';
const MAX_PER_HOUR = 30; // anyone with the link can order, and every order sends an email

function doPost(e) {
  try {
    const d = JSON.parse(e.postData.contents);
    const cache = CacheService.getScriptCache();
    // A confirm sent twice (double tap, retry after a timeout) must not email twice.
    const seen = 'order-' + String(d.nonce).slice(0, 64);
    if (cache.get(seen)) return reply({ok: true, repeated: true});
    if (!/^C3D[12] (?:(?:HDG|SS316) Fastener|EPDM Gasket) Order\.xlsx$/.test(d.filename)) return reply({error: 'Unexpected file name.'});
    if (typeof d.file !== 'string' || d.file.length > 700000) return reply({error: 'The order file is too large.'});
    const hour = 'count-' + Math.floor(Date.now() / 3600000), count = Number(cache.get(hour) || 0);
    if (count >= MAX_PER_HOUR) return reply({error: 'Too many orders in the last hour. Please try again later or tell the administrator.'});
    cache.put(hour, String(count + 1), 3600);

    const file = Utilities.newBlob(Utilities.base64Decode(d.file), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', d.filename);
    MailApp.sendEmail({to: TO, subject: String(d.subject).slice(0, 250), body: String(d.body).slice(0, 20000), attachments: [file], name: 'TWRP Order Form'});
    cache.put(seen, '1', 21600);

    // Record of web orders: one row per item. A failure here must not hide that the email was sent.
    try {
      const book = SpreadsheetApp.getActiveSpreadsheet();
      if (book && Array.isArray(d.rows)) {
        const sheet = book.getSheets()[0];
        if (sheet.getLastRow() === 0) sheet.appendRow(['Received', 'Requester', 'Project', 'Material', 'Title', 'Designation', 'Standard', 'Strength / PN', 'Qty', 'Remarks']);
        const now = new Date();
        // A leading apostrophe keeps text that starts with = + - @ from being read as a formula.
        const rows = d.rows.slice(0, 200).map(function (row) {
          return [now].concat(row.slice(0, 9).map(function (v) { return typeof v === 'string' && /^[=+@-]/.test(v) ? "'" + v : v; }));
        });
        if (rows.length) sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 10).setValues(rows);
      }
    } catch (err) {
      console.error(err);
    }
    return reply({ok: true});
  } catch (err) {
    console.error(err);
    return reply({error: 'The email could not be sent. Please download the file and send it to the administrator yourself.'});
  }
}

function reply(data) {
  return ContentService.createTextOutput(JSON.stringify(data)).setMimeType(ContentService.MimeType.JSON);
}

// ---- Weekly summary (added 2026-11-16) ----
// Every Monday at 09:00 Singapore time: emails TO the web orders of the previous Monday-to-Sunday (who ordered what,
// grouped by project, material, requester and title), with a link to this Google Sheet and the week's rows as a CSV.
// Set-up, once: in the Apps Script editor pick `installWeeklyTrigger` in the function list, click Run and authorize.
// To check it now without waiting for Monday: run `sendWeeklySummary` (it summarizes the last full Monday-to-Sunday week).
// The trigger runs the latest saved code, so no new deployment is needed for this part.

const TZ = 'Asia/Singapore';

function installWeeklyTrigger() {
  ScriptApp.getProjectTriggers().filter(function (t) { return t.getHandlerFunction() === 'sendWeeklySummary'; }).forEach(function (t) { ScriptApp.deleteTrigger(t); });
  ScriptApp.newTrigger('sendWeeklySummary').timeBased().onWeekDay(ScriptApp.WeekDay.MONDAY).atHour(9).inTimezone(TZ).create();
}

function sendWeeklySummary() {
  const book = SpreadsheetApp.getActiveSpreadsheet(), sheet = book.getSheets()[0], day = 86400000;
  // Start of today (Singapore) = the Monday this runs on; the week is the 7 days before it.
  const todayStart = new Date(Utilities.formatDate(new Date(), TZ, "yyyy-MM-dd'T'00:00:00+08:00"));
  const fromDate = new Date(todayStart.getTime() - 7 * day), toDate = todayStart;
  const label = Utilities.formatDate(fromDate, TZ, 'd MMM yyyy') + ' to ' + Utilities.formatDate(new Date(toDate.getTime() - day), TZ, 'd MMM yyyy');
  const data = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 10).getValues() : [];
  const rows = data.filter(function (r) { return r[0] instanceof Date && r[0] >= fromDate && r[0] < toDate; });
  const link = book.getUrl();
  if (!rows.length) {
    MailApp.sendEmail({to: TO, subject: 'Web orders ' + label + ': none', body: 'No web orders were received from ' + label + '.\n\nRecord: ' + link, name: 'TWRP Order Form'});
    return;
  }
  const groups = {}, order = [];
  rows.forEach(function (r) {
    const key = [r[2], r[3], r[1], r[4]].join(' · ');
    if (!groups[key]) { groups[key] = []; order.push(key); }
    groups[key].push('• ' + r[5] + (r[7] ? ' (' + r[7] + ')' : '') + ' × ' + Number(r[8]).toLocaleString('en-US'));
  });
  const who = {};
  rows.forEach(function (r) { who[r[1]] = true; });
  const body = ['Web orders ' + label + ': ' + Object.keys(who).length + ' requester(s), ' + order.length + ' title(s), ' + rows.length + ' item line(s).']
    .concat(order.map(function (key) { return key + '\n' + groups[key].join('\n'); }))
    .concat(['Record: ' + link]).join('\n\n');
  const csv = [['Received', 'Requester', 'Project', 'Material', 'Title', 'Designation', 'Standard', 'Strength / PN', 'Qty', 'Remarks']].concat(rows.map(function (r) {
    return [Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd HH:mm')].concat(r.slice(1));
  })).map(function (r) { return r.map(function (v) { return '"' + String(v).replace(/"/g, '""') + '"'; }).join(','); }).join('\r\n');
  MailApp.sendEmail({to: TO, subject: 'Web orders ' + label + ': ' + order.length + ' title(s)', body: body.slice(0, 20000), attachments: [Utilities.newBlob('\ufeff' + csv, 'text/csv', 'web-orders.csv')], name: 'TWRP Order Form'});
}
