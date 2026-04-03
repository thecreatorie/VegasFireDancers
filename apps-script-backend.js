/**
 * Vegas Fire Dancers — Registration Backend
 * Google Apps Script (deploy as Web App)
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to https://script.google.com and create a new project
 * 2. Paste this entire file into Code.gs
 * 3. Update the CONFIG section below with your details
 * 4. Create a Google Sheet with two tabs: "Registrations" and "Settings"
 *    - "Registrations" headers (Row 1): Timestamp | Name | Email | Phone | Activities | Consent | Status | Location Sent | Confirmation Sent
 *    - "Settings" tab, Cell A1: "Location", Cell B1: your actual location address
 * 5. Deploy > New Deployment > Web App
 *    - Execute as: Me
 *    - Who has access: Anyone
 * 6. Copy the Web App URL and paste it into index.html (replace the TODO comment in submitRegistration)
 *
 * ADMIN WORKFLOW:
 * - New registrations appear in the Google Sheet with Status = "PENDING"
 * - Change Status to "APPROVED" — the script auto-sends location email + welcome text
 * - Day-of confirmations run via a time-based trigger (set up in step 7)
 *
 * 7. Set up triggers:
 *    - In Apps Script, go to Triggers (clock icon)
 *    - Add trigger: sendDayOfConfirmations > Time-driven > Day timer > 9am-10am (Sunday)
 *    - Add trigger: onEdit > From spreadsheet > On edit (for auto-sending approval emails)
 */

// ===================== CONFIG =====================
const CONFIG = {
  SPREADSHEET_ID: 'YOUR_GOOGLE_SHEET_ID_HERE',     // The ID from your Google Sheet URL
  TWILIO_SID: 'YOUR_TWILIO_ACCOUNT_SID',            // From twilio.com/console
  TWILIO_AUTH: 'YOUR_TWILIO_AUTH_TOKEN',             // From twilio.com/console
  TWILIO_PHONE: '+1XXXXXXXXXX',                      // Your Twilio phone number
  ADMIN_EMAIL: 'caroline@hwtalentllc.com',           // Your email for notifications
  FROM_NAME: 'Vegas Fire Dancers',
  JAM_DAY: 0,                                        // 0 = Sunday
};
// ==================================================

/**
 * Handle incoming POST requests (registration form submissions)
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Registrations');

    // Append registration row
    sheet.appendRow([
      data.timestamp,
      data.name,
      data.email,
      data.phone,
      data.activities.join(', '),
      'YES',
      'PENDING',
      'NO',
      'NO'
    ]);

    // Notify admin
    MailApp.sendEmail({
      to: CONFIG.ADMIN_EMAIL,
      subject: 'New Registration: ' + data.name,
      htmlBody: buildAdminEmail(data)
    });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

/**
 * Handle CORS preflight
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}

/**
 * Triggered on edit — auto-sends location when status changes to APPROVED
 */
function onEdit(e) {
  const sheet = e.source.getActiveSheet();
  if (sheet.getName() !== 'Registrations') return;

  const col = e.range.getColumn();
  const row = e.range.getRow();

  // Column 7 = Status
  if (col !== 7 || row <= 1) return;

  const status = e.range.getValue().toString().toUpperCase();
  if (status !== 'APPROVED') return;

  const rowData = sheet.getRange(row, 1, 1, 9).getValues()[0];
  const name = rowData[1];
  const email = rowData[2];
  const phone = rowData[3];
  const activities = rowData[4];
  const locationSent = rowData[7];

  if (locationSent === 'YES') return; // Already sent

  // Get location from Settings tab
  const settings = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Settings');
  const location = settings.getRange('B1').getValue();

  // Send approval email with location
  MailApp.sendEmail({
    to: email,
    subject: 'You\'re Approved! Location Details for Sunday\'s Fire Jam',
    htmlBody: buildApprovalEmail(name, location, activities)
  });

  // Send SMS with location if Twilio is configured
  if (CONFIG.TWILIO_SID !== 'YOUR_TWILIO_ACCOUNT_SID' && phone) {
    sendSMS(phone,
      'Hey ' + name + '! You\'re approved for this Sunday\'s Vegas Fire Dancers jam. ' +
      'Location: ' + location + '. Setup starts at 6:30 PM, safety briefing at 7 PM. See you there!'
    );
  }

  // Mark location as sent
  sheet.getRange(row, 8).setValue('YES');
}

/**
 * Send day-of confirmation texts to all approved registrants
 * Set this up as a time-based trigger to run Sunday morning
 */
function sendDayOfConfirmations() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Registrations');
  const data = sheet.getDataRange().getValues();
  const settings = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Settings');
  const location = settings.getRange('B1').getValue();

  let confirmed = 0;

  for (let i = 1; i < data.length; i++) {
    const status = (data[i][6] || '').toString().toUpperCase();
    const confirmSent = (data[i][8] || '').toString().toUpperCase();
    const phone = data[i][3];
    const name = data[i][1];

    if (status === 'APPROVED' && confirmSent !== 'YES' && phone) {
      sendSMS(phone,
        'Reminder: Tonight\'s Vegas Fire Dancers jam! ' +
        'Setup at 6:30 PM, safety at 7 PM, flow until 10 PM. ' +
        'Location: ' + location + '. ' +
        'Reply YES to confirm attendance or NO to cancel. See you tonight!'
      );
      sheet.getRange(i + 1, 9).setValue('YES');
      confirmed++;
    }
  }

  // Notify admin of confirmation count
  MailApp.sendEmail({
    to: CONFIG.ADMIN_EMAIL,
    subject: 'Day-Of Confirmations Sent: ' + confirmed + ' attendees',
    htmlBody: '<p>' + confirmed + ' confirmation texts sent for tonight\'s jam.</p>'
  });
}

/**
 * Send SMS via Twilio
 */
function sendSMS(to, body) {
  const url = 'https://api.twilio.com/2010-04-01/Accounts/' + CONFIG.TWILIO_SID + '/Messages.json';

  // Normalize phone number
  let phone = to.toString().replace(/[^0-9+]/g, '');
  if (!phone.startsWith('+')) phone = '+1' + phone;

  const options = {
    method: 'post',
    headers: {
      'Authorization': 'Basic ' + Utilities.base64Encode(CONFIG.TWILIO_SID + ':' + CONFIG.TWILIO_AUTH)
    },
    payload: {
      'To': phone,
      'From': CONFIG.TWILIO_PHONE,
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

/**
 * Reset all confirmation statuses (run weekly before Sunday to allow re-registration tracking)
 * Set as a time-based trigger for Saturday night or early Sunday
 */
function resetWeeklyConfirmations() {
  const sheet = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID).getSheetByName('Registrations');
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    sheet.getRange(i + 1, 9).setValue('NO'); // Reset confirmation sent
  }
}

// ===================== EMAIL TEMPLATES =====================

function buildAdminEmail(data) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto">
      <h2 style="color:#FF4500">New Registration</h2>
      <table style="width:100%;border-collapse:collapse">
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Name</td><td style="padding:8px;border-bottom:1px solid #eee">${data.name}</td></tr>
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Email</td><td style="padding:8px;border-bottom:1px solid #eee">${data.email}</td></tr>
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Phone</td><td style="padding:8px;border-bottom:1px solid #eee">${data.phone}</td></tr>
        <tr><td style="padding:8px;border-bottom:1px solid #eee;font-weight:bold">Activities</td><td style="padding:8px;border-bottom:1px solid #eee">${data.activities.join(', ')}</td></tr>
        <tr><td style="padding:8px;font-weight:bold">Waiver</td><td style="padding:8px">Accepted</td></tr>
      </table>
      <p style="margin-top:16px;color:#666">Open your <a href="https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}">registration sheet</a> and change their status to APPROVED to send them the location.</p>
    </div>
  `;
}

function buildApprovalEmail(name, location, activities) {
  return `
    <div style="font-family:Arial,sans-serif;max-width:500px;margin:0 auto;background:#1a1a1a;color:#fff;padding:30px;border-radius:8px">
      <h2 style="color:#FF4500;margin-top:0">You're Approved!</h2>
      <p>Hey ${name},</p>
      <p>You've been approved for this Sunday's Vegas Fire Dancers jam. Here are your details:</p>
      <div style="background:#222;padding:16px;border-radius:4px;margin:16px 0;border-left:3px solid #FF4500">
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Location:</strong> ${location}</p>
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Setup:</strong> 6:30 PM</p>
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Safety Briefing:</strong> 7:00 PM</p>
        <p style="margin:0 0 8px"><strong style="color:#FF6B35">Flow Time:</strong> 7:00 – 10:00 PM</p>
        <p style="margin:0"><strong style="color:#FF6B35">Activities:</strong> ${activities}</p>
      </div>
      <p style="color:#999;font-size:13px">Please do not share this location publicly. You'll receive a confirmation text on Sunday to confirm your attendance.</p>
      <p style="color:#999;font-size:13px">By attending, you acknowledge the waiver of liability you agreed to during registration. Participation is at your own risk.</p>
      <p>See you Sunday!<br><strong style="color:#FF4500">Vegas Fire Dancers</strong><br><span style="color:#999;font-size:12px">Organized by HW Talent LLC</span></p>
    </div>
  `;
}
