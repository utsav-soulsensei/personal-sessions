/**
 * SoulSensei 1:1 Personal Guidance Sessions - Google Apps Script
 * Destination Spreadsheet ID: 1VsgRLyWoD2KigCwJDDUg4bIEskOA7hjR8qVzRxNrA_M
 * Sheet: https://docs.google.com/spreadsheets/d/1VsgRLyWoD2KigCwJDDUg4bIEskOA7hjR8qVzRxNrA_M/edit?usp=sharing
 *
 * HOW TO DEPLOY:
 * 1. Open the Google Sheet: https://docs.google.com/spreadsheets/d/1VsgRLyWoD2KigCwJDDUg4bIEskOA7hjR8qVzRxNrA_M/edit
 * 2. In the top menu, click Extensions > Apps Script.
 * 3. Delete any default code in Code.gs and paste the code below.
 * 4. Click Save (disk icon).
 * 5. Click "Deploy" > "New deployment".
 * 6. Under "Select type" (gear icon), select "Web app".
 * 7. Set:
 *    - Description: "SoulSensei Kavyal Event 1:1 Form Webhook"
 *    - Execute as: "Me"
 *    - Who has access: "Anyone" (essential so the event web form can submit without login).
 * 8. Click "Deploy", review/grant permissions, and copy the Web App URL (ends in /exec).
 * 9. Paste this URL into the HTML form (or the Admin Sync panel) as GOOGLE_SHEET_WEBHOOK_URL.
 */

var SPREADSHEET_ID = "1VsgRLyWoD2KigCwJDDUg4bIEskOA7hjR8qVzRxNrA_M";
var SHEET_NAME = "Form Responses";

function getOrCreateSheet() {
  var ss;
  try {
    ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  } catch (err) {
    ss = SpreadsheetApp.getActiveSpreadsheet();
  }
  
  var sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    // If "Form Responses" doesn't exist, check active sheet or create one
    sheet = ss.getActiveSheet();
    if (sheet && sheet.getLastRow() === 0) {
      sheet.setName(SHEET_NAME);
    } else {
      sheet = ss.insertSheet(SHEET_NAME);
    }
  }

  // Ensure header row exists
  if (sheet.getLastRow() === 0) {
    var headers = [
      "Timestamp",
      "Booking ID",
      "Full Name",
      "Mobile Number",
      "Area of Clarity",
      "Specific Topic / Notes",
      "Selected Session",
      "Status",
      "Source"
    ];
    sheet.appendRow(headers);
    var headerRange = sheet.getRange(1, 1, 1, headers.length);
    headerRange.setFontWeight("bold");
    headerRange.setBackground("#471C35");
    headerRange.setFontColor("#FFFFFF");
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, headers.length);
  }

  return sheet;
}

function doPost(e) {
  var lock = LockService.getScriptLock();
  try {
    // Wait up to 10 seconds for concurrent submissions
    lock.waitLock(10000);
    
    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (jsonErr) {
        data = e.parameter || {};
      }
    } else if (e.parameter) {
      data = e.parameter;
    }

    var sheet = getOrCreateSheet();
    var now = new Date();
    var formattedDate = Utilities.formatDate(now, "Asia/Kolkata", "yyyy-MM-dd HH:mm:ss");

    var row = [
      formattedDate,
      data.bookingId || ("SS-KAV-" + Math.floor(1000 + Math.random() * 9000)),
      data.fullName || "",
      "'" + (data.mobileNumber || ""), // Prefix with single quote to preserve leading zeros in phone numbers
      data.clarityArea || "",
      data.otherDetails || "",
      data.sessionType || "",
      "Slot Held",
      data.source || "Kavyal Event 1:1 Page"
    ];

    sheet.appendRow(row);

    return ContentService.createTextOutput(JSON.stringify({
      status: "success",
      message: "Slot held successfully",
      bookingId: data.bookingId
    })).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: error.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  try {
    var sheet = getOrCreateSheet();
    var data = sheet.getDataRange().getValues();
    return ContentService.createTextOutput(JSON.stringify({
      status: "online",
      spreadsheet: sheet.getParent().getName(),
      totalEntries: Math.max(0, data.length - 1)
    })).setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService.createTextOutput(JSON.stringify({
      status: "error",
      message: err.toString()
    })).setMimeType(ContentService.MimeType.JSON);
  }
}
