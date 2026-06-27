// ============================================================
// Harsh Bharti Portfolio — Career Guidance Booking Receiver
// Paste this entire file into:
//   Google Sheet → Extensions → Apps Script → Code.gs
// Then deploy as a Web App (see bottom of file)
// ============================================================

const SHEET_NAME = 'Sheet1'; // your existing tab name

// Called automatically for every POST from the portfolio form
function doPost(e) {
  try {
    const data  = JSON.parse(e.postData.contents);
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(SHEET_NAME);

    if (!sheet) {
      throw new Error('Sheet "' + SHEET_NAME + '" not found. Check the tab name.');
    }

    const lastRow  = sheet.getLastRow();
    const serialNo = lastRow; // row 1 = header, so data rows start at 2 → S.No. = 1, 2, 3…

    const row = [
      serialNo,
      data.submittedAt   || '',
      data.fullName      || '',
      data.email         || '',
      data.phone         || '',
      data.currentClass  || '',
      data.college       || '',
      data.questions     || '',
      data.preferredDate || '',
      data.preferredTime || '',
      'Pending'
    ];

    sheet.appendRow(row);

    // Alternate row shading for readability
    const newRow = sheet.getLastRow();
    if (newRow % 2 === 0) {
      sheet.getRange(newRow, 1, 1, 11).setBackground('#F0FBF4');
    }

    return ContentService
      .createTextOutput(JSON.stringify({ status: 'success', row: newRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ status: 'error', message: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// GET handler — open the deployed URL in a browser to confirm it's live
function doGet() {
  return ContentService
    .createTextOutput('✅ Booking receiver is live. Waiting for POST requests from the portfolio form.')
    .setMimeType(ContentService.MimeType.TEXT);
}

// ============================================================
// HOW TO DEPLOY (do this once):
//
// 1. In the Apps Script editor, click Deploy → New deployment
// 2. Click the gear ⚙ next to "Select type" → Web app
// 3. Set:
//      Execute as  → Me
//      Who has access → Anyone
// 4. Click Deploy → Authorise when prompted
// 5. Copy the Web app URL  (looks like https://script.google.com/macros/s/ABC.../exec)
// 6. Paste that URL into harsh-bharti-portfolio.html at:
//      const APPS_SCRIPT_URL = 'PASTE_URL_HERE';
// ============================================================
