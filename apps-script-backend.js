/**
 * Vegas Fire Dancers — Registration Backend
 * Google Apps Script (deploy as Web App)
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and create a new project
 * 2. Paste this entire file into Code.gs
 * 3. Update the CONFIG section below with your details
 * 4. Create a Google Sheet with two tabs: "Registrations" and "Settings"
 *    - "Registrations" headers (Row 1):
 *      Timestamp | Week Date | Week Label | Name | Email | Phone | Activities |
 *      Donations/Sharing | Referral | Consent | Status | Approval Sent | Confirmation Sent | Confirmed | Location Sent
 *    - "Settings" tab, Cell A1: "Location", Cell B1: your actual jam location address
 * 5. Deploy > New Deployment > Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the Web App URL and paste it into index.html (replace YOUR_APPS_SCRIPT_URL)
 * 7. Set up Twilio webhook (for auto-reply with location):
 *    - In Twilio Console > Phone Numbers > your number > Messaging
 *    - Set "A MESSAGE COMES IN" webhook URL to your Apps Script Web App URL
 *    - Method: HTTP POST
 * 8. Set up time-based triggers in Apps Script (Triggers > + Add Trigger):
 *    - sendDayBeforeConfirmations > Time-driven > Week timer > Every Saturday > 6pm-7pm
 *    - onEdit > From spreadsheet > On edit (for auto-sending approval notifications)
 *
 * ADMIN WORKFLOW:
 * 1. Someone registers on the site for a specific week → you get an email
 * 2. Open Google Sheet → change Status to "APPROVED"
 * 3. They receive an approval email (no location yet)
 * 4. Saturday evening (24hrs before), they get a confirmation text automatically
 * 5. They reply YES → location is sent automatically via Twilio webhook
 * 6. They show up Sunday!
 */

// ===================== CONFIG =====================
const CONFIG = {
  SPREADSHEET_ID: '1VSNr-SUocsMek2fmjjwMerSP5fFJ4F74D_M3Zhy5Kbo',
  TWILIO_SID: 'AC3710e1af3d457327b4d622247a30a079',
  TWILIO_AUTH: 'YOUR_AUTH_TOKEN_HERE',              // Add your Auth Token from Twilio Console — DO NOT commit this to GitHub
  TWILIO_PHONE: '+18446702635',
  TWILIO_MESSAGING_SID: 'MGc141149365714dcf2101d99cb36fe788',
  ADMIN_EMAIL: 'caroline@hwtalentllc.com',
  FROM_NAME: 'Vegas Fire Dancers',
};

// Column indexes (0-based) matching the spreadsheet
const COL = {
  TIMESTAMP: 0,
  WEEK_DATE: 1,
  WEEK_LABEL: 2,
  NAME: 3,
  EMAIL: 4,
  PHONE: 5,
  ACTIVITIES: 6,
  DONATE: 7,
  REFERRAL: 8,
  CONSENT: 9,
  STATUS: 10,       // PENDING / APPROVED / DENIED
  APPROVAL_SENT: 11,
  CONFIRM_SENT: 12,
  CONFIRMED: 13,    // YES after they reply
  LOCATION_SENT: 14,
};
// ==================================================

/**
 * Handle POST requests — either registration form or Twilio incoming SMS
 */
function doPost(e) {
  try {
    const contentType = e.postData.type || '';

    // Twilio sends form-urlencoded; our form sends JSON
    if (contentType.includes('urlencoded')) {
      return handleIncomingSMS(e);
    }

    return handleRegistration(e);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle registration form submission
 */
function handleRegistration(e) {
  const data = JSON.parse(e.postData.contents);
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Registrations');

  sheet.appendRow([
    data.timestamp,
    data.weekDate,
    data.weekLabel,
    data.name,
    data.email,
    data.phone,
    data.activities.join(', '),
    data.donate || '',
    data.referral || '',
    'YES',
    'PENDING',
    'NO',
    'NO',
    'NO',
    'NO'
  ]);

  // Notify admin
  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: 'New Registration (' + data.weekLabel + '): ' + data.name,
    htmlBody: buildAdminEmail(data)
  });

  return ContentService
    .createTextOutput(JSON.stringify({ success: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Handle incoming SMS from Twilio (someone replying YES to confirm)
 * This is called when someone texts your Twilio number.
 * Twilio webhook must be pointed at this Apps Script Web App URL.
 */
function handleIncomingSMS(e) {
  const params = {};
  const pairs = e.postData.contents.split('&');
  pairs.forEach(p => {
    const [k, v] = p.split('=');
    params[decodeURIComponent(k)] = decodeURIComponent(v.replace(/\+/g, ' '));
  });

  const fromPhone = (params.From || '').replace(/[^0-9+]/g, '');
  const body = (params.Body || '').trim().toUpperCase();

  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Registrations');
  const data = sheet.getDataRange().getValues();

  // Find the most recent APPROVED registration for this phone number
  // that has had a confirmation sent but not yet confirmed
  let matchRow = -1;
  for (let i = data.length - 1; i >= 1; i--) {
    let phone = (data[i][COL.PHONE] || '').toString().replace(/[^0-9]/g, '');
    let fromClean = fromPhone.replace(/[^0-9]/g, '');
    // Match last 10 digits
    if (phone.slice(-10) === fromClean.slice(-10)) {
      const status = (data[i][COL.STATUS] || '').toString().toUpperCase();
      const confirmSent = (data[i][COL.CONFIRM_SENT] || '').toString().toUpperCase();
      const confirmed = (data[i][COL.CONFIRMED] || '').toString().toUpperCase();
      if (status === 'APPROVED' && confirmSent === 'YES' && confirmed !== 'YES') {
        matchRow = i;
        break;
      }
    }
  }

  let replyMsg;

  if (matchRow === -1) {
    replyMsg = 'Vegas Fire Dancers: We could not find a pending confirmation for your number. ' +
      'Make sure you have registered and been approved at our website.';
  } else if (body === 'YES' || body === 'Y' || body === 'CONFIRM') {
    // Mark as confirmed
    sheet.getRange(matchRow + 1, COL.CONFIRMED + 1).setValue('YES');

    // Get location and send it
    const settings = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Settings');
    const location = settings.getRange('B1').getValue();
    const name = data[matchRow][COL.NAME];
    const weekLabel = data[matchRow][COL.WEEK_LABEL];

    replyMsg = 'Confirmed! See you at the jam, ' + name + '! ' +
      weekLabel + ' — Setup 6:30 PM, Safety 7 PM, Flow 7-10 PM. ' +
      'Location: ' + location + '. ' +
      'Do not share this location. See you there!';

    // Mark location sent
    sheet.getRange(matchRow + 1, COL.LOCATION_SENT + 1).setValue('YES');

    // Also send location via email
    const email = data[matchRow][COL.EMAIL];
    const activities = data[matchRow][COL.ACTIVITIES];
    MailApp.sendEmail({
      to: email,
      subject: 'Confirmed! Here\'s Your Location for ' + weekLabel,
      htmlBody: buildLocationEmail(name, location, activities, weekLabel)
    });

    // Notify admin
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: 'Attendance Confirmed: ' + name + ' (' + weekLabel + ')',
      htmlBody: '<p><strong>' + name + '</strong> confirmed attendance for ' + weekLabel + '.</p>'
    });
  } else if (body === 'NO' || body === 'N' || body === 'CANCEL') {
    sheet.getRange(matchRow + 1, COL.STATUS + 1).setValue('CANCELLED');
    const name = data[matchRow][COL.NAME];
    replyMsg = 'Got it, ' + name + '. You\'re cancelled for this week. ' +
      'Register again anytime at our website. Hope to see you soon!';

    // Notify admin
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: 'Cancellation: ' + name + ' (' + data[matchRow][COL.WEEK_LABEL] + ')',
      htmlBody: '<p><strong>' + name + '</strong> cancelled for ' + data[matchRow][COL.WEEK_LABEL] + '.</p>'
    });
  } else {
    replyMsg = 'Vegas Fire Dancers: Reply YES to confirm your attendance and receive the location, ' +
      'or NO to cancel. Other messages are not monitored.';
  }

  // Return TwiML response
  const twiml = '<?xml version="1.0" encoding="UTF-8"?>' +
    '<Response><Message>' + escapeXml(replyMsg) + '</Message></Response>';

  return ContentService
    .createTextOutput(twiml)
    .setMimeType(ContentService.MimeType.XML);
}

function escapeXml(str) {
  return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
}

/**
 * Handle CORS preflight / health check
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Triggered on edit — when Status changes to APPROVED, send approval email
 * (No location sent at this stage — location is only sent after they confirm)
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'Registrations') return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  // Column 11 (1-indexed) = Status
  if (col !== COL.STATUS + 1 || row <= 1) return;

  const status = e.range.getValue().toString().toUpperCase();
  if (status !== 'APPROVED') return;

  const rowData = sheet.getRange(row, 1, 1, 15).getValues()[0];
  const approvalSent = (rowData[COL.APPROVAL_SENT] || '').toString().toUpperCase();
  if (approvalSent === 'YES') return;

  const name = rowData[COL.NAME];
  const email = rowData[COL.EMAIL];
  const weekLabel = rowData[COL.WEEK_LABEL];
  const activities = rowData[COL.ACTIVITIES];

  // Send approval email (NO location — they get that after confirming)
  MailApp.sendEmail({
    to: email,
    subject: 'You\'re Approved for ' + weekLabel + '!',
    htmlBody: buildApprovalEmail(name, activities, weekLabel)
  });

  // Mark approval sent
  sheet.getRange(row, COL.APPROVAL_SENT + 1).setValue('YES');
}

/**
 * Send confirmation texts 24 hours before jam (Saturday evening)
 * Set as trigger: sendDayBeforeConfirmations > Time-driven > Week timer > Saturday > 6pm-7pm
 */
function sendDayBeforeConfirmations() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Registrations');
  const data = sheet.getDataRange().getValues();

  // Calculate tomorrow's date (Sunday)
  const tomorrow = new Date();
  tomorrow.setDate(tomorrow.getDate() + 1);
  const tomorrowStr = tomorrow.toISOString().split('T')[0];

  let sent = 0;

  for (let i = 1; i < data.length; i++) {
    const weekDate = (data[i][COL.WEEK_DATE] || '').toString().split('T')[0];
    const status = (data[i][COL.STATUS] || '').toString().toUpperCase();
    const confirmSent = (data[i][COL.CONFIRM_SENT] || '').toString().toUpperCase();
    const phone = data[i][COL.PHONE];
    const name = data[i][COL.NAME];
    const weekLabel = data[i][COL.WEEK_LABEL];

    if (weekDate === tomorrowStr && status === 'APPROVED' && confirmSent !== 'YES' && phone) {
      sendSMS(phone,
        'Hey ' + name + '! Tomorrow is the Vegas Fire Dancers jam (' + weekLabel + '). ' +
        'Setup at 6:30 PM, safety at 7 PM, flow until 10 PM. ' +
        'Reply YES to confirm your attendance and receive the location, or NO to cancel.'
      );
      sheet.getRange(i + 1, COL.CONFIRM_SENT + 1).setValue('YES');
      sent++;
    }
  }

  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: 'Confirmation Texts Sent: ' + sent + ' for tomorrow\'s jam',
    htmlBody: '<p>' + sent + ' confirmation texts sent for tomorrow\'s jam (' + tomorrowStr + ').</p>' +
      '<p>Attendees who reply YES will automatically receive the location.</p>'
  });
}

/**
 * Send SMS via Twilio
 */
function sendSMS(to, body) {
  if (CONFIG.TWILIO_AUTH === 'YOUR_AUTH_TOKEN_HERE') return;

  const url = 'https://api.twilio.com/2010-04-01/Accounts/' + CONFIG.TWILIO_SID + '/Messages.json';

  let phone = to.toString().replace(/[^0-9+]/g, '');
  if (!phone.startsWith('+')) phone = '+1' + phone;

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(CONFIG.TWILIO_SID + ':' + CONFIG.TWILIO_AUTH)
    },
    payload: {
      'To': phone,
      'MessagingServiceSid': CONFIG.TWILIO_MESSAGING_SID,
      'Body': body
    },
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  const result = JSON.parse(response.getContentText());

  if (result.error_code) {
    Logger.log('Twilio error for ' + phone + ': ' + result.error_message);
  }

  return result;
}

// ===================== EMAIL TEMPLATES =====================

function buildAdminEmail(data) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto">
      <h2 style="color:#FF4500">New Registration — ${data.weekLabel}</h2>
      <table style="width:100%;border-collapse:collapse">
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Week</td><td style="padding:8px;border-bottom:1px solid #eee">${data.weekLabel} (${data.weekDate})</td></tr>
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Name</td><td style="padding:8px;border-bottom:1px solid #eee">${data.name}</td></tr>
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Email</td><td style="padding:8px;border-bottom:1px solid #eee">${data.email}</td></tr>
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Phone</td><td style="padding:8px;border-bottom:1px solid #eee">${data.phone}</td></tr>
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Activities</td><td style="padding:8px;border-bottom:1px solid #eee">${data.activities.join(', ')}</td></tr>
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Donating/Sharing</td><td style="padding:8px;border-bottom:1px solid #eee">${data.donate || 'None'}</td></tr>
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Referring</td><td style="padding:8px;border-bottom:1px solid #eee">${data.referral || 'None'}</td></tr>
        <tr><td style="padding:8px;font-weight:bold">Waiver</td><td style="padding:8px">Accepted</td></tr>
      </table>
      <p style="margin-top:16px;color:#666">Open your <a href="https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}">registration sheet</a> and change their Status to <strong>APPROVED</strong>.</p>
    </div>
  `;
}

function buildApprovalEmail(name, activities, weekLabel) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;background:#1a1a1a;color:#fff;padding:30px;border-radius:8px">
      <h2 style="color:#FF4500;margin-top:0">You're Approved!</h2>
      <p>Hey ${name},</p>
      <p>You've been approved for the <strong>${weekLabel}</strong> Vegas Fire Dancers jam!</p>
      <div style="background:#222;padding:16px;border-radius:4px;margin:16px 0;border-left:3px solid #FF4500">
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Date:</strong> ${weekLabel}</p>
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Setup:</strong> 6:30 PM</p>
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Safety Briefing:</strong> 7:00 PM</p>
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Flow Time:</strong> 7:00 – 10:00 PM</p>
        <p style="margin:0"><strong style="color:#FF6B35">Activities:</strong> ${activities}</p>
      </div>
      <div style="background:#331a00;padding:16px;border-radius:4px;margin:16px 0;border:1px solid #FF4500">
        <p style="margin:0;color:#FF6B35;font-weight:bold">HOW TO GET THE LOCATION:</p>
        <p style="margin:8px 0 0;color:#ccc">24 hours before the jam, you will receive a confirmation text. <strong style="color:#fff">Reply YES</strong> to confirm your attendance — the location will be sent to you automatically.</p>
      </div>
      <p style="color:#999;font-size:13px">By attending, you acknowledge the waiver of liability you agreed to during registration. Participation is at your own risk.</p>
      <p>See you soon!<br><strong style="color:#FF4500">Vegas Fire Dancers</strong><br><span style="color:#999;font-size:12px">Organized by HW Talent LLC</span></p>
    </div>
  `;
}

function buildLocationEmail(name, location, activities, weekLabel) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;background:#1a1a1a;color:#fff;padding:30px;border-radius:8px">
      <h2 style="color:#FF4500;margin-top:0">You're Confirmed! Here's the Location</h2>
      <p>Hey ${name},</p>
      <p>You're confirmed for <strong>${weekLabel}</strong>. Here's everything you need:</p>
      <div style="background:#222;padding:16px;border-radius:4px;margin:16px 0;border-left:3px solid #FF4500">
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Location:</strong> ${location}</p>
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Setup:</strong> 6:30 PM</p>
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Safety Briefing:</strong> 7:00 PM</p>
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Flow Time:</strong> 7:00 – 10:00 PM</p>
        <p style="margin:0"><strong style="color:#FF6B35">Activities:</strong> ${activities}</p>
      </div>
      <p style="color:#FF4500;font-weight:bold;font-size:14px">Please do not share this location publicly.</p>
      <p style="color:#999;font-size:13px">Participation is at your own risk per the waiver you agreed to during registration.</p>
      <p>See you tonight!<br><strong style="color:#FF4500">Vegas Fire Dancers</strong><br><span style="color:#999;font-size:12px">Organized by HW Talent LLC</span></p>
    </div>
  `;
}
