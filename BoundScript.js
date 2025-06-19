// === CONFIG: Web App URL ===
const WEB_APP_URL = 'https://script.google.com/macros/s/AKfycbzZfI005n68Yc96yKkFPCRbxmNzTNMUU4qhAd_BbvHY_Wj-qKkceGRaEoI7PZII_5hM/exec';

// === BUILD MENU ON OPEN ===
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
//LIST IMPORT CALL
function triggerLeadImport() {
  const url = 'https://script.google.com/macros/s/AKfycbyFjgpRv_Iq48W0kp9ekzVWFPaZYaCNkYHdAcFU9v_IMWZ8vSyZ20e7fszBVvMCflNj/exec'; // Replace with your new Web App URL

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
