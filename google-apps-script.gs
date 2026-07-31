/**
 * Saum Studio — Lead form → Google Sheet + email notification
 * ---------------------------------------------------------------
 * Paste this into the Apps Script editor of your Google Sheet
 * (Extensions → Apps Script), then Deploy as a Web App.
 * Full step-by-step setup is in README.md ("Google Sheet setup").
 *
 * This version writes each value into the column whose HEADER matches,
 * so your own column order / extra columns (Budget, Remark, etc.) are
 * respected. Just make sure row 1 has headers named like:
 *   Timestamp | Name | Phone | Home type | Locality | Source |
 *   Campaign Name | Ad Set | Ad Name | Placement | Page
 * (Spelling is flexible — see ALIASES below. Extra columns are left blank.)
 */

// Leads are also emailed here (set to '' to turn email off).
var NOTIFY_EMAIL = 'design@saumstudio.com';

// Default header row, only used if the sheet is completely empty.
var DEFAULT_HEADERS = ['Timestamp', 'Name', 'Phone', 'Home type', 'Locality', 'Source',
                       'Campaign Name', 'Ad Set', 'Ad Name', 'Placement', 'Page'];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheets()[0];
    var now = new Date();

    // Phone kept as text (leading ' stops Sheets rounding / dropping the +).
    var phoneText = data.phone ? "'" + data.phone : '';

    // value for each normalized header name (lowercased, trimmed)
    var ALIASES = {
      'timestamp': now, 'time': now, 'date': now,
      'name': data.name || '', 'full name': data.name || '',
      'phone': phoneText, 'phone number': phoneText, 'mobile': phoneText, 'whatsapp': phoneText,
      'home type': data.bhk || '', 'home': data.bhk || '', 'bhk': data.bhk || '',
      'locality': data.area || '', 'area': data.area || '',
      'source': data.source || '',
      'campaign name': data.utm_campaign || '', 'campaign': data.utm_campaign || '', 'utm_campaign': data.utm_campaign || '',
      'ad set': data.utm_medium || '', 'ad set name': data.utm_medium || '', 'adset': data.utm_medium || '', 'utm_medium': data.utm_medium || '',
      'ad name': data.utm_content || '', 'ad': data.utm_content || '', 'utm_content': data.utm_content || '',
      'placement': data.utm_source || '', 'utm_source': data.utm_source || '',
      'page': data.page || '', 'url': data.page || ''
    };

    // If the sheet is empty, lay down a sensible default header row first.
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(DEFAULT_HEADERS);
    }

    // Read the actual header row and place each value under its matching header.
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    var row = new Array(lastCol).fill('');
    for (var c = 0; c < headers.length; c++) {
      var key = String(headers[c]).trim().toLowerCase();
      if (ALIASES.hasOwnProperty(key)) {
        row[c] = ALIASES[key];
      }
    }
    sheet.appendRow(row);

    // Email notification so leads are seen instantly.
    if (NOTIFY_EMAIL) {
      MailApp.sendEmail({
        to: NOTIFY_EMAIL,
        subject: 'New lead — ' + (data.name || 'Website') + ' (' + (data.bhk || '') + ')',
        body:
          'New consultation request from the website:\n\n' +
          'Name: ' + (data.name || '-') + '\n' +
          'Phone: ' + (data.phone || '-') + '\n' +
          'Home type: ' + (data.bhk || '-') + '\n' +
          'Locality: ' + (data.area || '-') + '\n' +
          'Source: ' + (data.source || '-') + '\n\n' +
          'Campaign: ' + (data.utm_campaign || '-') + '\n' +
          'Ad set: ' + (data.utm_medium || '-') + '\n' +
          'Ad: ' + (data.utm_content || '-') + '\n' +
          'Placement: ' + (data.utm_source || '-') + '\n\n' +
          'Page: ' + (data.page || '-') + '\n' +
          'Time: ' + now
      });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ result: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ result: 'error', message: String(err) }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Lets you open the /exec URL in a browser to confirm it's live.
function doGet() {
  return ContentService.createTextOutput('Saum Studio lead endpoint is running.');
}
