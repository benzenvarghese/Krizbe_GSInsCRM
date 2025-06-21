// === CONFIG: Web App URL ===
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzZfI005n68Yc96yKkFPCRbxmNzTNMUU4qhAd_BbvHY_Wj-qKkceGRaEoI7PZII_5hM/exec';

// ===   BUILD MENU ON OPEN  ===
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu("📣 CRM Actions")
    .addItem("🔔 Send Renewal Reminders", "runCheckAndNotify")
    .addItem("📅 Send Monthly Working Leads", "runMonthlyWorkingLeads")
    .addSeparator() // Added for better visual separation
    .addItem("🔄 Update Renewal Reminder Dates", "runUpdateRenewalDates") // Added this menu item
	.addSeparator()
	.addItem("🧹 Clear Logs", "clearLogs")
    .addToUi();
}

// === MENU ACTION HANDLERS ===

function runCheckAndNotify() {
  triggerWebAppAction("checkAndNotifyRenewals");
}

function runMonthlyWorkingLeads() {
  triggerWebAppAction("notifyMonthlyWorkingLeads");
}

function runUpdateRenewalDates() {
  triggerWebAppAction("updateRenewalReminderDates");
}

//Clear logs sheet
function clearLogs() {
  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Logs");
    if (!sheet) throw new Error("Logs sheet not found.");

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).clearContent();
    }

    SpreadsheetApp.getUi().alert("All logs cleared (except header).");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Error clearing logs: " + e.message);
  }
}


// === CORE CALLER ===

function triggerWebAppAction(actionName) {
  const ui = SpreadsheetApp.getUi(); // Get the UI object to display alerts

  try {
    // Show an initial message indicating the action is starting
    ui.alert('Initiating Action', `Sending request for "${actionName}"... Please wait.`, ui.ButtonSet.OK);

    const payload = { action: actionName };
    Logger.log("Sending payload: " + JSON.stringify(payload));

    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true // Crucial for catching web app errors gracefully
    };

    const response = UrlFetchApp.fetch(WEB_APP_URL, options);
    const responseText = response.getContentText();
    const responseCode = response.getResponseCode();

    Logger.log(`Action ${actionName} triggered. Web App Response Code: ${responseCode}, Content: ${responseText}`);

    // Attempt to parse JSON response for structured feedback
    let feedbackMessage = `Action "${actionName}" completed.`;
    if (responseCode >= 200 && responseCode < 300) { // Success range
      try {
        const responseData = JSON.parse(responseText);
        if (responseData && responseData.status === 'success' && responseData.message) {
          feedbackMessage = `Success: ${responseData.message}`;
        } else if (responseData && responseData.status === 'error' && responseData.message) {
          feedbackMessage = `Failed: ${responseData.message}`;
        } else {
          // Fallback if JSON structure is unexpected
          feedbackMessage = `Action completed with web app response: ${responseText}`;
        }
      } catch (e) {
        // Response was not valid JSON, treat as raw text
        feedbackMessage = `Action completed. Web App Response (text): ${responseText}`;
      }
      ui.alert('Action Completed', feedbackMessage, ui.ButtonSet.OK);
    } else { // HTTP Error from Web App
      feedbackMessage = `Web App returned an error (Code: ${responseCode}): ${responseText}`;
      ui.alert('Web App Error', feedbackMessage, ui.ButtonSet.OK);
    }

  } catch (err) {
    const errorMessage = "An unexpected error occurred while calling the Web App: " + err.message;
    Logger.log(errorMessage);
    ui.alert('Script Error', errorMessage, ui.ButtonSet.OK); // Notify user of script-side errors
  }
}
//    ====  LIST IMPORT CALL
function triggerLeadImport() {
  const url = WEB_APP_URL; // Replace with your new Web App URL

  const payload = {
    action: "importLeadsFromStaging"
  };

  const options = {
    method: 'post',
    contentType: 'application/json',
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  };

  const response = UrlFetchApp.fetch(url, options);
  Logger.log("Lead Import Triggered: " + response.getContentText());
}

// === TIMESTAMP AND FUTURE DATE LOGIC ===

function onEdit(e) {
  try {
    logMessage("➡️ onEdit triggered");

    const sheet = e.range.getSheet();
    const editedCell = e.range;
    const row = editedCell.getRow();
    const col = editedCell.getColumn();

    const sheetName1 = "WorkingLeads";
    const colF = 6;
    const colH = 8;
    const timestampCol = 9;
    const now = new Date();

    logMessage(`📝 Edited Sheet: ${sheet.getName()}, Row: ${row}, Column: ${col}`);

    if (sheet.getName() === sheetName1 && row !== 1) {
      // Timestamp Logic
      const dateOptions = { year: 'numeric', month: '2-digit', day: '2-digit' };
      const timeOptions = { hour: 'numeric', minute: '2-digit', hour12: true };
      const dayOptions = { weekday: 'short' };
      const shortDate = now.toLocaleDateString('en-US', dateOptions);
      const time = now.toLocaleTimeString('en-US', timeOptions);
      const day = now.toLocaleDateString('en-US', dayOptions);
      const finalStamp = `${shortDate}, ${time} - ${day}`;
      sheet.getRange(row, timestampCol).setValue(finalStamp);

      logMessage(`✅ Timestamp written at I${row}: ${finalStamp}`);

      // Column F change logic
      if (col === colF) {
        let value = editedCell.getValue();
        logMessage(`🔍 F${row} value: "${value}"`);

        if (!value || typeof value !== "string") {
          logMessage(`⚠️ Invalid or empty value at F${row}. Skipping...`);
          return;
        }

        const match = value.toString().trim().match(/^(\d+)\s*(?:[-]?\s*(mo|month|months)?)?/i);
        const numMonths = match ? parseInt(match[1]) : NaN;

        if (!isNaN(numMonths) && numMonths >= 0) {
          const futureDate = new Date(now.getFullYear(), now.getMonth() + numMonths, 1);
          const futureOptions = { year: 'numeric', month: 'long' };
          const formattedDate = futureDate.toLocaleDateString('en-US', futureOptions);
          sheet.getRange(row, colH).setValue(formattedDate);
          logMessage(`✅ H${row} updated with: ${formattedDate} for ${numMonths} month(s)`);
        } else {
          SpreadsheetApp.getActive().toast(`❌ Could not extract month from: "${value}"`);
          logMessage(`❌ Extraction failed for F${row}: "${value}"`);
        }
      } else {
        logMessage(`ℹ️ Edit was not in Column F. No future date update attempted.`);
      }
    } else {
      logMessage(`ℹ️ Edit not on '${sheetName1}' or first row. Ignored.`);
    }
  } catch (error) {
    logMessage("🔥 Exception in onEdit: " + error.message, true);
  }
}

