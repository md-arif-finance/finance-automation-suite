function onOpen() {
  var ui = SpreadsheetApp.getUi();
    ui.createMenu('Finance Tool')
      .addItem('🚀 Initialize System', 'initializeSystem') 
      .addSeparator()
      .addItem('🔄 Run Follow-ups Now', 'runAutoFollowups')
      .addItem('🧹 Repair Layout (Safe Mode)', 'setupAllSheets')
      .addToUi();
}

/**
 * SETUP: Installs Triggers safely (Prevents Duplicates)
 */
function createInstallableOnEditTrigger() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // 1. DELETE EXISTING TRIGGERS FIRST
  var triggers = ScriptApp.getProjectTriggers();
  
  for (var i = 0; i < triggers.length; i++) {
    var handlerFunction = triggers[i].getHandlerFunction();
    if (handlerFunction === "onEditTrigger" || handlerFunction === "runAutoFollowups") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  
  // 2. CREATE FRESH TRIGGERS
  ScriptApp.newTrigger("onEditTrigger").forSpreadsheet(ss).onEdit().create();
  ScriptApp.newTrigger("runAutoFollowups").timeBased().everyHours(1).create();
}

function onEditTrigger(e) {
  // Guard clause: If the trigger runs without an event object (e.g. manually), stop.
  if (!e || !e.range) return;

  var sheet = e.range.getSheet();
  var row = e.range.getRow();
  var col = e.range.getColumn();
  var val = e.value;

  // --- A. ACTION PANEL TRIGGERS (Col 15 / O) ---
  if (sheet.getName() === SHEET_NAME_UI && col === 15) {
    
    if (e.range.isChecked()) {
      
      // ROW 5: SAVE & SEND
      if (row === 5) {
        e.source.toast("🚀 Sending Invoice...");
        try { processInvoice("SEND"); } 
        catch (err) { SpreadsheetApp.getUi().alert("Error: " + err.message); }
      }
      
      // ROW 6: SAVE AS DRAFT
      else if (row === 6) {
        e.source.toast("💾 Saving Draft...");
        try { processInvoice("DRAFT"); } 
        catch (err) { SpreadsheetApp.getUi().alert("Error: " + err.message); }
      }
      
      // ROW 7: CLEAR FORM
      else if (row === 7) {
        if (Browser.msgBox("Clear Form?", "Are you sure you want to clear all data?", Browser.Buttons.YES_NO) == "yes") {
          clearInvoiceForm();
          e.source.toast("🧹 Form Cleared");
        }
      }
      
      // Always uncheck the button after running
      e.range.uncheck();
    }
    return; 
  }

  // --- C. TRACKER STATUS UPDATES (The Fix) ---
  if (sheet.getName() === SHEET_NAME_TRACKER) {
    // Column 7 is "Status"
    // We check if value is "Ready" AND row > 1
    if (col === 7 && val === "Ready" && row > 1) {
      
      e.source.toast("🚀 Generating PDF & Sending Email...");
      
      try {
        // Call the reconstruction function
        processInvoiceFromHistory(row, false); 
      } catch (err) {
        SpreadsheetApp.getUi().alert("Error sending invoice: " + err.message);
        // Optional: Revert status to Draft if failed
        sheet.getRange(row, 7).setValue("Draft");
      }
    }
  }
}

function runAutoFollowups() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_TRACKER);
  if (!sheet) return;

  var data = sheet.getDataRange().getValues();
  var now = new Date();
  var count = 0;

  ss.toast("🔄 Checking for overdue invoices...");

  for (var i = 1; i < data.length; i++) {
    var rowNum = i + 1;
    var status = data[i][6];       // Column G
    var nextFollowUp = data[i][10]; // Column K

    if (status === "Sent" && nextFollowUp instanceof Date && nextFollowUp <= now) {
      console.log("Found overdue invoice at Row " + rowNum);
      try {
        // We pass "true" to indicate this is an Auto-Followup (to update next date)
        processInvoiceFromHistory(rowNum, true); 
        count++;
      } catch (e) {
        console.error("Failed row " + rowNum + ": " + e.message);
      }
      sendInvoiceEmail(sheet, rowNum, "Follow-up Reminder");
    }
  }

  // FINAL SUMMARY TOAST
  if (count > 0) {
    ss.toast("✅ Sent " + count + " follow-up emails.");
  } else {
    ss.toast("👍 No overdue invoices found.");
  }
}