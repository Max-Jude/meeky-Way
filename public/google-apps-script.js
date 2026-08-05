/**
 * De Wisdom Comprehensive Academy - Admission Form Apps Script
 * =============================================================
 * How to setup:
 * 1. Open Google Sheets (https://sheets.google.com) and create a new sheet.
 * 2. Click "Extensions" -> "Apps Script" in the top menu.
 * 3. Delete any code in the editor and paste this entire code below.
 * 4. Click "Deploy" (top right) -> "New deployment".
 * 5. Select type: "Web app" (click the gear icon next to Select type).
 * 6. Set Description: "Admission Form Handler"
 * 7. Set Execute as: "Me"
 * 8. Set Who has access: "Anyone"
 * 9. Click "Deploy", grant permissions if prompted.
 * 10. Copy the Web App URL (starts with https://script.google.com/macros/s/...)
 * 11. Paste the URL into your .env file as VITE_GOOGLE_APPS_SCRIPT_URL="..."
 *     OR directly in index.html as const MANUAL_APPS_SCRIPT_URL = "YOUR_WEB_APP_URL";
 */

function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.tryLock(10000); // Prevent concurrent write race conditions

  try {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
    
    // Automatically set up headers if the sheet is empty
    if (sheet.getLastRow() === 0) {
      sheet.appendRow([
        "Timestamp",
        "Student Full Name",
        "Age",
        "Gender",
        "Target Class",
        "Date of Birth",
        "Guardian / Parent Name",
        "Parent WhatsApp Number",
        "Previous School Attended",
        "Medical Condition / Special Needs"
      ]);
      
      // Format Header Row (Bold + Navy Fill)
      var headerRange = sheet.getRange(1, 1, 1, 10);
      headerRange.setFontWeight("bold");
      headerRange.setBackground("#000080");
      headerRange.setFontColor("#FFFFFF");
    }

    // Parse incoming form data (handles both JSON body and URL parameters)
    var data = {};
    if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        data = e.parameter;
      }
    } else {
      data = e.parameter;
    }

    var timestamp = new Date();

    // Extract values with safe fallbacks
    var studentName = data.studentName || "";
    var studentAge = data.studentAge || "";
    var studentGender = data.studentGender || "";
    var targetClass = data.targetClass || "";
    var dob = data.dob || "";
    var guardianName = data.guardianName || "";
    var parentWhatsapp = data.parentWhatsapp || "";
    var previousSchool = data.previousSchool || "";
    var medicalCondition = data.medicalCondition || "None";

    // Append new admission record to Google Sheet
    sheet.appendRow([
      timestamp,
      studentName,
      studentAge,
      studentGender,
      targetClass,
      dob,
      guardianName,
      parentWhatsapp,
      previousSchool,
      medicalCondition
    ]);

    // Format phone column as text to preserve leading zeros / country code +
    var lastRow = sheet.getLastRow();
    sheet.getRange(lastRow, 8).setNumberFormat('@');

    return ContentService.createTextOutput(
      JSON.stringify({ result: "success", row: lastRow })
    ).setMimeType(ContentService.MimeType.JSON);

  } catch (error) {
    return ContentService.createTextOutput(
      JSON.stringify({ result: "error", error: error.toString() })
    ).setMimeType(ContentService.MimeType.JSON);

  } finally {
    lock.releaseLock();
  }
}

function doGet(e) {
  return ContentService.createTextOutput("De Wisdom Comprehensive Academy Admission Web App is Active!");
}
