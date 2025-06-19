// ========== NANBA BUILD v2.3 — Centralized Configurable Logging Version ==========

// === CONFIG ===
const SHEET_ID = '1criS0D-fjnZpHS12hXBQquvOxqqJv7k_ViyCS7HkRPU';

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action;

    logMessage("Incoming request: " + JSON.stringify(payload));

    if (action === "checkAndNotifyRenewals") {
      checkAndNotifyRenewals();
    } else if (action === "notifyMonthlyWorkingLeads") {
      notifyMonthlyWorkingLeads();
    } else if (action === "updateRenewalReminderDates") {
      updateRenewalReminderDates();
    } else if (action === "importLeadsFromStaging") {
      importLeadsFromStaging();
    } else {
      logMessage("Unknown action received: " + action, true);
    }

    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
  } catch (error) {
    logMessage("Exception in doPost: " + error.message, true);
    return ContentService.createTextOutput("Error: " + error.message).setMimeType(ContentService.MimeType.TEXT);
  }
}

// === RENEWAL REMINDER DAILY ===
function checkAndNotifyRenewals() {
  try {
    logMessage("Started checkAndNotifyRenewals", true);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("CancellationTrack");
    const data = sheet.getDataRange().getValues();
    const today = new Date();
    let remindersSent = 0;
    let eligibleLeads = [];

    logMessage(`Fetched ${data.length - 1} rows from CancellationTrack`, true);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const firstName = row[0];

      if (!firstName || firstName.toString().trim() === '') {
        logMessage(`Row ${i + 1}: Blank row detected — stopping further processing.`);
        break;
      }

      const followUpDate = parseDateSafe(row[5], i + 1);
      const renewalDate = parseDateSafe(row[4], i + 1);
      const notifiedFlag = row[6];

      if (!followUpDate) continue;

      if (isSameDate(followUpDate, today) && notifiedFlag !== "Yes") {
        eligibleLeads.push({
          firstName: row[0],
          lastName: row[1],
          contactNumber: row[2],
          renewalDate,
          rowIndex: i + 1
        });
        remindersSent++;
      } else {
        logMessage(`Row ${i + 1}: Not eligible for notification.`);
      }
    }

    if (eligibleLeads.length > 0) {
      sendConsolidatedEmail(eligibleLeads);
      eligibleLeads.forEach(lead => {
        sheet.getRange(lead.rowIndex, 7).setValue("Yes");
        sheet.getRange(lead.rowIndex, 7).setBackground("#00FF00");
      });
    }

    logMessage(`Total reminders sent: ${remindersSent}`, true);
  } catch (error) {
    logMessage("Exception in checkAndNotifyRenewals: " + error.message, true);
  }
}

// === UPDATE REMINDER DATE ===
function updateRenewalReminderDates() {
  try {
    logMessage("Started updateRenewalReminderDates", true);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("CancellationTrack");
    const data = sheet.getDataRange().getValues();
    const today = new Date();
    let updatedRows = 0;

    logMessage(`Fetched ${data.length - 1} rows from CancellationTrack`, true);

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const firstName = row[0];
      if (!firstName || firstName.toString().trim() === '') {
        logMessage(`Row ${i + 1}: Blank row detected — stopping further processing.`);
        break;
      }

      const renewalDate = parseDateSafe(row[4], i + 1);
      if (!renewalDate) continue;

      const followUpDate = new Date(renewalDate);
      followUpDate.setDate(followUpDate.getDate() - 42);
      sheet.getRange(i + 1, 6).setValue(followUpDate);
      updatedRows++;
    }

    logMessage(`Total renewal reminder dates updated: ${updatedRows}`, true);
  } catch (error) {
    logMessage("Exception in updateRenewalReminderDates: " + error.message, true);
  }
}

// === MONTHLY WORKING LEADS ===
function notifyMonthlyWorkingLeads() {
  try {
    logMessage("Started notifyMonthlyWorkingLeads", true);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName("WorkingLeads");
    const data = sheet.getDataRange().getValues();

    logMessage(`Fetched ${data.length - 1} rows from WorkingLeads`, true);

    const recipientEmail = getNotificationEmail("MonthlyLeadsEmails");
    const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy");
    const subject = `Monthly Working Leads Summary - ${todayFormatted}`;

    let body = `Dear Admin,\n\nBelow are the current active working leads:\n\n`;

    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      if (!row[0]) break;
      body += `${i}. ${row.join(" | ")}\n`;
    }

    body += `\nRegards,\nAMAIA Team\n(This is an automated message)`;

    GmailApp.sendEmail(recipientEmail, subject, body);
    logMessage("Monthly working leads email sent successfully.");

  } catch (error) {
    logMessage("Exception in notifyMonthlyWorkingLeads: " + error.message, true);
  }
}

// === CONSOLIDATED EMAIL ===
function sendConsolidatedEmail(leads) {
  try {
    logMessage("Inside sendConsolidatedEmail(): Started");

    const recipientEmail = getNotificationEmail("RenewalEmails");
    logMessage("Email fetched from SetUp: " + recipientEmail);

    const todayFormatted = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "MMMM dd, yyyy");
    const subject = `Daily Follow Up - ${todayFormatted}`;

    let body = `Dear Admin,\n\nThe following leads are due for follow-up today:\n\n`;
    leads.forEach((lead, index) => {
      const renewalDateFormatted = lead.renewalDate ? Utilities.formatDate(lead.renewalDate, Session.getScriptTimeZone(), "MMMM dd, yyyy") : "N/A";
      body += `${index + 1}. ${lead.firstName} ${lead.lastName} – Contact: ${lead.contactNumber} – Renewal Date: ${renewalDateFormatted}\n`;
    });

    body += `\nRegards,\nAMAIA Team\n\n(This is an automated email, please do not reply.)`;

    GmailApp.sendEmail(recipientEmail, subject, body);
    logMessage("Consolidated email sent successfully.");
  } catch (error) {
    logMessage("Exception inside sendConsolidatedEmail(): " + error.message, true);
  }
}

// === UTILITIES ===
function logMessage(message, forceLog = false) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const setupSheet = ss.getSheetByName("SetUp");
  const setupData = setupSheet.getDataRange().getValues();

  let logLevel = 'Minimal';
  for (let i = 0; i < setupData.length; i++) {
    if (setupData[i][0] === 'LogLevel') {
      logLevel = setupData[i][1];
      break;
    }
  }

  if (logLevel === 'Detailed' || forceLog) {
    const logSheet = ss.getSheetByName("Logs");
    logSheet.appendRow([new Date(), message]);
  }
}

function parseDateSafe(rawDate, rowNum) {
  try {
    if (rawDate instanceof Date) {
      return rawDate;
    } else if (typeof rawDate === 'string' && rawDate.trim() !== '') {
      const normalizedDateStr = rawDate.replace('T', ' ').replace('Z', '');
      const parsedDate = new Date(normalizedDateStr);
      if (isNaN(parsedDate.getTime())) {
        logMessage(`Row ${rowNum}: Invalid date format after parsing`);
        return null;
      }
      return parsedDate;
    } else {
      logMessage(`Row ${rowNum}: Invalid or missing date`);
      return null;
    }
  } catch (e) {
    logMessage(`Row ${rowNum}: Exception while parsing date - ${e.message}`, true);
    return null;
  }
}

function isSameDate(date1, date2) {
  return date1.getFullYear() === date2.getFullYear() &&
         date1.getMonth() === date2.getMonth() &&
         date1.getDate() === date2.getDate();
}

function getNotificationEmail(attribute) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const setupSheet = ss.getSheetByName("SetUp");
  const data = setupSheet.getDataRange().getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === attribute) {
      return data[i][1];
    }
  }
  throw new Error(`No email found for attribute: ${attribute}`);
}

// === LEAD IMPORT FUNCTION ===
function importLeadsFromStaging() {
  try {
    logMessage("Started importLeadsFromStaging", true);

    const ss = SpreadsheetApp.openById(SHEET_ID);
    const setupSheet = ss.getSheetByName("SetUp");
    const leadStageSheet = ss.getSheetByName("LeadStage");
    const workingLeadsSheet = ss.getSheetByName("WorkingLeads");

    const setupData = setupSheet.getRange(2, 3, setupSheet.getLastRow(), 2).getValues(); // Columns C & D
    const statusCol = getSetupValue("LeadImportStatusColumn");
    const primaryKeyDestCol = setupData[1][0]; // C3 - e.g., "C"
    const primaryKeySourceCol = setupData[1][1]; // D3 - e.g., "A"
    const mappings = setupData.slice(2); // From C4 onwards

    logMessage(`Primary Key Destination Column: ${primaryKeyDestCol}`);
    logMessage(`Primary Key Source Column: ${primaryKeySourceCol}`);
    logMessage(`Lead Import Status Column: ${statusCol}`);
    logMessage(`Column Mappings: ${JSON.stringify(mappings)}`);

    const leadStageData = leadStageSheet.getDataRange().getValues();
    const workingLeadsData = workingLeadsSheet.getDataRange().getValues();

    const leadStageHeaders = leadStageData[0];
    let importCount = 0, duplicateCount = 0;

    for (let i = 1; i < leadStageData.length; i++) {
      const row = leadStageData[i];
      const primaryValue = row[columnToIndex(primaryKeySourceCol)];

      if (!primaryValue || primaryValue.toString().trim() === '') continue;

      const isDuplicate = workingLeadsData.some(wr => wr[columnToIndex(primaryKeyDestCol)] == primaryValue);
      const statusCell = leadStageSheet.getRange(i + 1, columnToIndex(statusCol) + 1);

      if (isDuplicate) {
        statusCell.setValue("Duplicate").setBackground("#FFC7CE");
        duplicateCount++;
        continue;
      }

      const newRow = [];
      mappings.forEach(([destCol, srcCol]) => {
        newRow[columnToIndex(destCol)] = row[columnToIndex(srcCol)];
      });

      workingLeadsSheet.appendRow(newRow);
      statusCell.setValue("Imported").setBackground("#C6EFCE");
      importCount++;
    }

    logMessage(`Lead Import Complete. Imported: ${importCount}, Duplicates: ${duplicateCount}`, true);

  } catch (error) {
    logMessage("Exception in importLeadsFromStaging: " + error.message, true);
  }
}

function getSetupValue(attribute) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const setupSheet = ss.getSheetByName("SetUp");
  const data = setupSheet.getDataRange().getValues();

  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === attribute) {
      return data[i][1];
    }
  }
  throw new Error(`No setup value found for attribute: ${attribute}`);
}

function columnToIndex(col) {
  return col.toUpperCase().charCodeAt(0) - 65;
}