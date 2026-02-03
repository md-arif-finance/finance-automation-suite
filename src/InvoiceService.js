/**
 * 7) updateDashboard()
 * Calculates financial totals and updates the Dashboard tab.
 * Run this manually or link it to the onEdit trigger.
 */
function updateDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackerSheet = ss.getSheetByName("Invoice_Tracker");
  var dashSheet = ss.getSheetByName("Dashboard");

  // Create Dashboard tab if it doesn't exist
  if (!dashSheet) {
    dashSheet = ss.insertSheet("Dashboard");
    // Move it to the first position
    ss.setActiveSheet(dashSheet);
    ss.moveActiveSheet(1);
  }

  // Get all data from Tracker
  // (We assume data starts at row 2, columns A to G are relevant)
  var lastRow = trackerSheet.getLastRow();
  if (lastRow < 2) return; // No data yet

  var data = trackerSheet.getRange(2, 1, lastRow - 1, 12).getValues();

  // Initialize totals
  var totalInvoiced = 0;
  var totalCollected = 0;
  var totalOutstanding = 0;
  var overdueCount = 0;

  var now = new Date();

  // Loop through rows to calculate
  data.forEach(function(row) {
    var amount = row[3]; // Col D is Amount
    var status = row[6]; // Col G is Status
    var dueDate = row[5]; // Col F is Due Date

    // Skip empty rows
    if (!amount) return;

    totalInvoiced += amount;

    if (status === "Paid") {
      totalCollected += amount;
    } else {
      totalOutstanding += amount;
      // Check if overdue
      if (dueDate instanceof Date && dueDate < now) {
        overdueCount++;
      }
    }
  });

  // --- DRAW THE DASHBOARD ---
  dashSheet.clear(); // Wipe clean to redraw
  
  // Title
  dashSheet.getRange("B2").setValue("FINANCIAL DASHBOARD")
    .setFontSize(24).setFontWeight("bold").setFontColor("#202124");

  // KPI Cards (Key Performance Indicators)
  // 1. Total Invoiced
  createKpiCard(dashSheet, "B4", "Total Invoiced", totalInvoiced, "#4285F4"); // Blue
  // 2. Total Collected
  createKpiCard(dashSheet, "E4", "Collected", totalCollected, "#0F9D58"); // Green
  // 3. Outstanding
  createKpiCard(dashSheet, "H4", "Outstanding", totalOutstanding, "#DB4437"); // Red

  // Stats Text
  dashSheet.getRange("B10").setValue("Overdue Invoices: " + overdueCount)
    .setFontSize(14).setFontColor("#DB4437").setFontWeight("bold");

  dashSheet.getRange("B11").setValue("Last Updated: " + now.toLocaleString())
    .setFontSize(10).setFontColor("#666");

  // Remove gridlines for a cleaner look
  dashSheet.setHiddenGridlines(true);
}

/**
 * HELPER: Draws a pretty box for the KPI
 */
function createKpiCard(sheet, cellAddress, title, value, color) {
  var range = sheet.getRange(cellAddress);
  var row = range.getRow();
  var col = range.getColumn();
  
  // Format the box area (3 rows x 2 columns)
  var boxRange = sheet.getRange(row, col, 3, 2);
  boxRange.setBackground("#f8f9fa").setBorder(true, true, true, true, null, null, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // Set Title
  sheet.getRange(row, col).setValue(title)
    .setFontSize(11).setFontColor("#5f6368").setFontWeight("bold");

  // Set Value
  sheet.getRange(row + 1, col).setValue("$" + value.toFixed(2))
    .setFontSize(20).setFontColor(color).setFontWeight("bold");
}

/**
 * MAIN CONTROLLER
 * mode = "SEND" (Default) or "DRAFT"
 */
function processInvoice(mode) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var uiSheet = ss.getSheetByName(SHEET_NAME_UI);
  var trackerSheet = ss.getSheetByName(SHEET_NAME_TRACKER);
  var itemSheet = ss.getSheetByName(SHEET_NAME_ITEMS);

  // --- 1. SCRAPE DATA ---
  var clientName = uiSheet.getRange("D4").getValue();
  var clientEmail = uiSheet.getRange("D5").getValue();
  var invDate = uiSheet.getRange("K4").getValue();
  var invNo = uiSheet.getRange("K5").getValue();
  var dueDate = uiSheet.getRange("K6").getValue();

  if (clientName === "") throw new Error("Client Name is required!");

  // --- 2. SCRAPE ITEMS ---
  var startRow = 12;
  var tableData = uiSheet.getRange(startRow, 2, 10, 12).getValues(); 
  var cleanItems = [];
  var subTotal = 0, totalCGST = 0, totalSGST = 0, totalIGST = 0;
  var productData = ss.getSheetByName(SHEET_NAME_PRODUCTS).getDataRange().getValues();

  for (var i = 0; i < tableData.length; i++) {
    var row = tableData[i];
    if (row[1] !== "") { // If Item Name exists

      var itemName = row[1];
      var description = fetchProductDescription(itemName, productData);

      var itemObj = {
        srNo: row[0], name: row[1], description: description, hsn: row[2], qty: row[3], 
        rate: row[4], disc: row[5], taxable: row[6], gstRate: row[7],
        cgst: row[8], sgst: row[9], igst: row[10], total: row[11]
      };
      cleanItems.push(itemObj);
      subTotal += (itemObj.taxable || 0);
      totalCGST += (itemObj.cgst || 0);
      totalSGST += (itemObj.sgst || 0);
      totalIGST += (itemObj.igst || 0);
      
      // Save to Items Archive
      itemSheet.appendRow([
        invNo, itemObj.srNo, itemObj.name, itemObj.hsn, itemObj.qty, 
        itemObj.rate, itemObj.disc, itemObj.taxable, itemObj.gstRate, 
        itemObj.cgst, itemObj.sgst, itemObj.igst, itemObj.total
      ]);
    }
  }

  if (cleanItems.length === 0) throw new Error("No items found!");

  var grandTotal = subTotal + totalCGST + totalSGST + totalIGST;
  var totalTax = totalCGST + totalSGST + totalIGST;

  // --- 3. DETERMINE STATUS ---
  var status = (mode === "SEND") ? "Sent" : "Draft";
  var pdfLink = "Not Generated";

  // --- 4. HANDLE EMAILING (Only if SEND mode) ---
  if (mode === "SEND") {
    var clientGST = uiSheet.getRange("D6").getValue();
    var clientAddress = uiSheet.getRange("D7").getValue();
    var placeOfSupply = uiSheet.getRange("D9").getValue();

    // FIX: FORCE DATE FORMAT (DD/MM/YYYY)
    var tz = Session.getScriptTimeZone();
    var formattedDate = Utilities.formatDate(new Date(invDate), tz, "dd/MM/yyyy");
    var formattedDueDate = Utilities.formatDate(new Date(dueDate), tz, "dd/MM/yyyy");

    var invoiceData = {
      invNo: invNo,
      date: formattedDate,    // <--- Usage
      dueDate: formattedDueDate, // <--- Usage
      client: { name: clientName, email: clientEmail, gstin: clientGST, address: clientAddress, state: placeOfSupply },
      items: cleanItems,
      totals: { taxable: subTotal, cgst: totalCGST, sgst: totalSGST, igst: totalIGST, grandTotal: grandTotal, amountInWords: NUMBER_TO_WORDS(grandTotal) }
    };
    
    // Send Email & Get Link
    pdfLink = sendInvoiceEmail(invoiceData); // Ensure EmailService returns the file URL
  }

  // --- 5. SAVE TO TRACKER ---
  trackerSheet.appendRow([
    invNo,           // 1. Invoice No
    clientName,      // 2. Client Name
    clientEmail,     // 3. Email
    grandTotal,      // 4. Amount (We use Grand Total here)
    invDate,         // 5. Invoice Date
    dueDate,         // 6. Due Date
    status,          // 7. Status ("Sent" or "Draft")
    3,               // 8. Follow-up Value (Default: 3)
    "Days",          // 9. Follow-up Unit (Default: Days)
    "",              // 10. Last Follow-up Sent (Empty initially)
    "",              // 11. Next Follow-up At (Empty initially)
    mode === "DRAFT" ? "Draft saved manually" : "Initial Invoice Sent" // 12. Notes
  ]);

  // --- 6. CLEANUP & INCREMENT ---
  clearInvoiceForm(); // Call helper to clear form
  incrementInvoiceNumber(uiSheet, invNo); // Helper to increment
  
  var msg = (mode === "SEND") ? "✅ Invoice Sent & Saved!" : "💾 Invoice Saved as Draft!";
  ss.toast(msg);
}

/**
 * HELPER: Clears Inputs & Restores Logic
 */
function clearInvoiceForm() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_UI);
  
  // 1. Clear Header Inputs
  sheet.getRange("D4").clearContent(); // Name
  
  // 2. Clear Table Inputs (Item, Qty, Disc)
  // We only clear the columns the USER types in.
  var startRow = 12; 
  var numRows = 10;
  
  sheet.getRange(startRow, 3, numRows, 1).clearContent(); // Col C: Item
  sheet.getRange(startRow, 5, numRows, 1).clearContent(); // Col E: Qty
  sheet.getRange(startRow, 7, numRows, 1).clearContent(); // Col G: Disc
  
  // 3. RESTORE FORMULAS (The Magic Step)
  // This overwrites Rate, Tax, Total, etc. with fresh formulas
  restoreInvoiceFormulas();

  // 4. Reset Meta Data
  sheet.getRange("K4").setValue(new Date()); 
  sheet.getRange("O5:O7").uncheck();
}

/**
 * HELPER: Increments Invoice Number
 */
function incrementInvoiceNumber(sheet, currentInvNo) {
  var parts = currentInvNo.split("-");
  if (parts.length > 1) {
    var num = parseInt(parts[1]) + 1;
    var newNum = num < 10 ? "00" + num : (num < 100 ? "0" + num : num);
    var newInvNo = parts[0] + "-" + newNum;
    sheet.getRange("K5").setValue(newInvNo);
  }
}

/**
 * RECONSTRUCT & SEND: Sends email when Status changes in History Sheet
 */
function processInvoiceFromHistory(row, isAutoFollowup) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var trackerSheet = ss.getSheetByName(SHEET_NAME_TRACKER);
  var itemSheet = ss.getSheetByName(SHEET_NAME_ITEMS);
  var clientSheet = ss.getSheetByName(SHEET_NAME_CLIENTS);

  // 1. FETCH TRACKER DATA
  // Cols: 1:InvNo, 2:Name, 3:Email, 4:Amt, 5:Date, 6:Due, 7:Status
  var trackerData = trackerSheet.getRange(row, 1, 1, 7).getValues()[0];

  var invNo = trackerData[0];
  var clientName = trackerData[1];
  var clientEmail = trackerData[2];
  var grandTotal = trackerData[3];
  var invDate = trackerData[4];
  var dueDate = trackerData[5];
  var freqVal = trackerData[7];
  var freqUnit = trackerData[8];

  if (!invNo || !clientName) {
    throw new Error("Missing Invoice Number or Client Name");
  }

  // 2. FETCH CLIENT DETAILS (Address, GST, State)
  // We need to look up the client in the DB to get the missing details for the PDF
  var clientDb = clientSheet.getDataRange().getValues();
  var clientObj = { name: clientName, email: clientEmail, gstin: "", address: "", state: "" };
  
  for (var i = 1; i < clientDb.length; i++) {
    // Col 0 is Name
    if (clientDb[i][0] === clientName) {
      clientObj.gstin = clientDb[i][2];   // Col C
      clientObj.address = clientDb[i][3]; // Col D
      clientObj.state = clientDb[i][4];   // Col E
      break;
    }
  }

  // 3. FETCH & RECONSTRUCT ITEMS
  // We look through the Archive to find items matching this Invoice No
  var archiveData = itemSheet.getDataRange().getValues();
  var productData = ss.getSheetByName(SHEET_NAME_PRODUCTS).getDataRange().getValues();
  var cleanItems = [];
  var subTotal = 0, totalCGST = 0, totalSGST = 0, totalIGST = 0;

  for (var j = 1; j < archiveData.length; j++) {
    // Archive Col 0 is Invoice No
    if (archiveData[j][0] === invNo) {
      var itemRow = archiveData[j];
      var itemName = itemRow[2];
      
      var itemObj = {
        srNo: itemRow[1], 
        name: itemName, 
        description: fetchProductDescription(itemName, productData), 
        hsn: itemRow[3], qty: itemRow[4], rate: itemRow[5],
        disc: itemRow[6], taxable: itemRow[7], gstRate: itemRow[8],
        cgst: itemRow[9], sgst: itemRow[10], igst: itemRow[11], total: itemRow[12]
      };
      
      cleanItems.push(itemObj);
      subTotal += itemObj.taxable;
      totalCGST += itemObj.cgst;
      totalSGST += itemObj.sgst;
      totalIGST += itemObj.igst;
    }
  }

  if (cleanItems.length === 0) {
    throw new Error("No items found for Invoice " + invNo + " in Archive.");
  }

  // 4. PREPARE PDF DATA
  var tz = Session.getScriptTimeZone();
  var invoiceData = {
    invNo: invNo,
    date: Utilities.formatDate(new Date(invDate), tz, "dd/MM/yyyy"),
    dueDate: Utilities.formatDate(new Date(dueDate), tz, "dd/MM/yyyy"),
    client: clientObj,
    items: cleanItems,
    totals: {
      taxable: subTotal, cgst: totalCGST, sgst: totalSGST, igst: totalIGST,
      grandTotal: grandTotal, amountInWords: NUMBER_TO_WORDS(grandTotal)
    }
  };

  // 5. SEND EMAIL
  sendInvoiceEmail(invoiceData);

  // 6. UPDATE STATUS TO "Sent"
  trackerSheet.getRange(row, 7).setValue("Sent"); // Update Status
  trackerSheet.getRange(row, 10).setValue(new Date()); // Last Sent Time

  // CRITICAL: Calculate Next Follow-up (Prevent Infinite Loop)
  if (isAutoFollowup) {
    var now = new Date();
    var nextDate = new Date(now.getTime());
    
    if (freqUnit === "Minutes") nextDate.setMinutes(nextDate.getMinutes() + freqVal);
    else if (freqUnit === "Hours") nextDate.setHours(nextDate.getHours() + freqVal);
    else if (freqUnit === "Days") nextDate.setDate(nextDate.getDate() + freqVal);
    
    trackerSheet.getRange(row, 11).setValue(nextDate); // Update "Next Follow-up"
    trackerSheet.getRange(row, 12).setValue("Auto-Reminder Sent"); 
  } else {
    var now = new Date();
    var nextDate = new Date(now.getTime());
    nextDate.setDate(nextDate.getDate() + 3); // Default 3 days if manual start
    trackerSheet.getRange(row, 11).setValue(nextDate);
    trackerSheet.getRange(row, 12).setValue("Manually Sent"); 
  }
  
  SpreadsheetApp.getActiveSpreadsheet().toast("✅ Invoice " + invNo + " sent successfully!");
}

/**
 * UTILITY: Fetches Description from Product Master
 * Accepts productData array to avoid repeated API calls
 */
function fetchProductDescription(itemName, productData) {
  // Loop through product rows (Skip header row 0)
  for (var i = 1; i < productData.length; i++) {
    // Col 0 = Item Name, Col 1 = Description
    if (productData[i][0] === itemName) {
      return productData[i][1] || ""; // Return Desc or empty string
    }
  }
  return ""; // Not found
}