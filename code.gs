// ═══════════════════════════════════════════════════════
//  FireHire – Ignition Program | Google Apps Script
//  Sheet ID : 1Lid0PGZ2pmjNTcgSvpy8J_iQ3WtfhDHumAOiovnsuWc
//  Sheet Name: The Course
// ═══════════════════════════════════════════════════════

const SHEET_ID   = '1Lid0PGZ2pmjNTcgSvpy8J_iQ3WtfhDHumAOiovnsuWc';
const SHEET_NAME = 'The Course';

// ── Entry point for POST requests from the form ──────────
function doPost(e) {
  try {
    const data   = JSON.parse(e.postData.contents);
    const sheet  = SpreadsheetApp
                    .openById(SHEET_ID)
                    .getSheetByName(SHEET_NAME);

    // Add headers if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        'Aplicant Name',
        'Whatsapp Number',
        'Email',
        'English Level',
        'How did it reach us?',
        'If you came via Recruiter, the name of the Recruiter',
        'Submission Date'
      ]);

      // Style the header row
      const headerRange = sheet.getRange(1, 1, 1, 7);
      headerRange.setBackground('#FF4D00')
                 .setFontColor('#FFFFFF')
                 .setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    // Append the new row
    sheet.appendRow([
      data.applicantName  || '',
      data.whatsapp       || '',
      data.email          || '',
      data.englishLevel   || '',
      data.howReach       || '',
      data.recruiterName  || '',
      new Date().toLocaleString('en-EG', { timeZone: 'Africa/Cairo' })
    ]);

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success' }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── Allow GET (health check) ─────────────────────────────
function doGet() {
  // تأكد إن اسم الملف بين القوسين هو نفس اسم ملف الـ HTML عندك (غالباً index أو login)
  return HtmlService.createHtmlOutputFromFile('index') 
      .setTitle('FireHire Academy - Form')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}
