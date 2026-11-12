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
