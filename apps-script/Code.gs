/**
 * Your Retirement Insights — website lead handler (US)
 *
 * Receives inquiries from the modal on yourretirementinsights.com, appends each
 * one to this spreadsheet, and creates the contact in GoHighLevel via the API.
 *
 * SETUP
 *  1. Create a Google Sheet named exactly:
 *       Retirement Insights - Website Leads (US)
 *  2. Extensions > Apps Script. Rename the project to:
 *       Retirement Insights Lead Handler (US)
 *     Delete anything in the editor, paste this file in.
 *  3. Fill in GHL_LOCATION_ID and GHL_API_TOKEN below.
 *     In the US GoHighLevel sub-account (name it "Your Retirement Insights (US)"):
 *       Settings > Private Integrations > Create new integration
 *       Name:   Retirement Insights Website (US)
 *       Scopes: contacts.write and contacts.readonly
 *     The token starts "pit-". The location id is in the sub-account URL.
 *     Leave both empty to skip GHL; the sheet and the email alert still work,
 *     and the "CRM result" column will read "not configured".
 *  4. Deploy > New deployment > type "Web app".
 *       Execute as:      Me
 *       Who has access:  Anyone
 *     Copy the /exec URL and paste it into FORM_ENDPOINT in index.html.
 *
 * NOTE: keep the token out of the public GitHub repo. It lives here in Apps
 * Script only. Anyone holding it can read and write your CRM contacts.
 *
 * Re-deploying after an edit: Deploy > Manage deployments > pencil icon >
 * Version "New version" > Deploy. The URL stays the same.
 */

var GHL_API_TOKEN = '';
var GHL_LOCATION_ID = '';          // US sub-account, NOT the UK egvzjbE2j8uCngMIIDh7
var GHL_API_VERSION = '2021-07-28';

// info@yourretirementinsights.com has no mailbox yet (the domain has no MX
// records), so alerts go somewhere that actually receives mail. Switch this to
// info@yourretirementinsights.com once that mailbox is live.
var NOTIFY_EMAIL = 'jamie@adsorbed.co.uk';

var SHEET_NAME = 'Leads';

var HEADERS = [
  'Timestamp',
  'First name',
  'Last name',
  'Email',
  'Phone',
  'Investable assets',
  'Requested',
  'Page',
  'CRM result'
];

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);

    // Reject anything missing the fields the form marks as required, so junk
    // does not silently fill the sheet.
    if (!data.email || !data.firstName) {
      return json({ ok: false, error: 'Missing required fields' });
    }

    var crmResult = sendToGhl(data);
    appendRow(data, crmResult);
    notify(data, crmResult);

    return json({ ok: true });
  } catch (err) {
    console.error('Lead handler failed: ' + err);
    return json({ ok: false, error: String(err) });
  }
}

function doGet() {
  return json({ ok: true, message: 'Your Retirement Insights lead handler is running.' });
}

function appendRow(data, crmResult) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    sheet.getRange(1, 1, 1, HEADERS.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  sheet.appendRow([
    new Date(),
    data.firstName || '',
    data.lastName || '',
    data.email || '',
    data.phone || '',
    data.assets || '',
    data.resource || '',
    data.page || '',
    crmResult
  ]);
}

/**
 * Upsert the contact in GoHighLevel. Upsert rather than create, so a repeat
 * inquiry from the same email updates the existing contact instead of erroring.
 *
 * Returns a short status string that gets written to the sheet, so a broken
 * CRM connection is visible as a column of failures rather than as leads that
 * quietly never arrived.
 */
function sendToGhl(data) {
  if (!GHL_API_TOKEN || !GHL_LOCATION_ID) return 'not configured';

  try {
    // Upsert the core fields only. Tags are deliberately NOT sent here:
    // upsert REPLACES the whole tag array, so a second inquiry from the same
    // person would wipe the assets tag from their first one.
    var res = ghlFetch('https://services.leadconnectorhq.com/contacts/upsert', {
      locationId: GHL_LOCATION_ID,
      firstName: data.firstName || '',
      lastName: data.lastName || '',
      email: data.email || '',
      phone: data.phone || '',
      source: 'Retirement Insights website (US)'
    });

    var code = res.getResponseCode();
    var body = res.getContentText();

    if (code < 200 || code >= 300) {
      console.error('GHL upsert returned ' + code + ': ' + body);
      return 'failed (' + code + ')';
    }

    var contact = (JSON.parse(body).contact) || {};
    if (!contact.id) return 'ok';

    addTags(contact.id, data);
    return 'ok · ' + contact.id;
  } catch (err) {
    // Never let a CRM outage lose the lead. The sheet row still gets written.
    console.error('GHL upsert threw: ' + err);
    return 'failed';
  }
}

/**
 * Tags are applied through the dedicated endpoint, which appends rather than
 * overwrites, so a returning enquirer accumulates their history instead of
 * losing it. Qualifying detail rides on tags because they need no custom-field
 * setup and they are what the callers actually filter on.
 */
function addTags(contactId, data) {
  var tags = ['Website Lead', 'US'];
  if (data.assets) tags.push('Assets: ' + data.assets);
  if (data.resource) tags.push('Requested: ' + data.resource);

  try {
    var res = ghlFetch(
      'https://services.leadconnectorhq.com/contacts/' + contactId + '/tags',
      { tags: tags }
    );
    if (res.getResponseCode() >= 300) {
      console.error('GHL tagging returned ' + res.getResponseCode() + ': ' + res.getContentText());
    }
  } catch (err) {
    // A tagging failure must not fail the lead; the contact already exists.
    console.error('GHL tagging threw: ' + err);
  }
}

function ghlFetch(url, payload) {
  return UrlFetchApp.fetch(url, {
    method: 'post',
    contentType: 'application/json',
    headers: {
      Authorization: 'Bearer ' + GHL_API_TOKEN,
      Version: GHL_API_VERSION,
      Accept: 'application/json'
    },
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
}

function notify(data, crmResult) {
  if (!NOTIFY_EMAIL) return;

  var subject = 'New US website lead: ' + (data.firstName || '') + ' ' + (data.lastName || '');
  var body = [
    'A new inquiry came in from yourretirementinsights.com.',
    '',
    'Name:               ' + (data.firstName || '') + ' ' + (data.lastName || ''),
    'Email:              ' + (data.email || ''),
    'Phone:              ' + (data.phone || ''),
    'Investable assets:  ' + (data.assets || ''),
    'Requested:          ' + (data.resource || ''),
    'Page:               ' + (data.page || ''),
    '',
    'CRM:                ' + crmResult
  ].join('\n');

  try {
    MailApp.sendEmail(NOTIFY_EMAIL, subject, body);
  } catch (err) {
    console.error('Notification email failed: ' + err);
  }
}

function json(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
