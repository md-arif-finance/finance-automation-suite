/**
 * MASTER INITIALIZATION FUNCTION
 * This is the "One Button" the user clicks to start.
 */
function initializeSystem() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  ss.toast("⚙️ System initializing... Please wait.", "Setup Started", 10);

  // 1. Build/Repair all Sheet Structures
  setupAllSheets();

  // 2. Install Automation Triggers
  createInstallableOnEditTrigger(); // This function is in Triggers.js

  ss.toast("✅ System Ready! Triggers installed & Sheets verified.", "Success", 5);
}

/**
 * Master Setup Function - Run this first!
 */
function setupAllSheets() {
  setupInvoiceTracker();      
  setupDashboard();
  setupCustomerDatabase();
  setupProductMaster();
  setupAdvancedStructure();   
  setupLogSheet();            
  
  organizeSheets();
}

/**
 * BUILDS THE "COMMAND CENTER" DASHBOARD (Fixed Layout & Borders)
 */
function setupDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_DASHBOARD, 0);
  }

  // 1. BACKUP EXISTING SETTINGS (Removed Subject)
  var savedSettings = {
    name: sheet.getRange("D11").getValue(),
    addr: sheet.getRange("D12").getValue(),
    state: sheet.getRange("D14").getValue()
  };
  
  // 1. CLEAN SLATE
  sheet.getRange("A1:Z100").breakApart(); 
  sheet.clear(); 
  sheet.setHiddenGridlines(true);

  // --- A. THE HEADER (Rows 1-3) ---
  sheet.getRange("B2:M3").merge().setValue("INVOICE DASHBOARD")
    .setFontWeight("bold").setFontSize(24).setHorizontalAlignment("center").setVerticalAlignment("middle")
    .setBackground("#2c3e50").setFontColor("white");

  // --- C. LEFT COLUMN: SETTINGS (Rows 9+) ---
  var setRow = 9;
  sheet.getRange(setRow, 3).setValue("🏢 MY COMPANY PROFILE").setFontWeight("bold").setFontSize(14).setFontColor("#2c3e50");
  
  // Labels (Removed Email Subject)
  sheet.getRange(setRow + 2, 3).setValue("Company Name:").setFontWeight("bold");       
  sheet.getRange(setRow + 3, 3).setValue("Address:").setFontWeight("bold");            
  sheet.getRange(setRow + 5, 3).setValue("My State:").setFontWeight("bold");           
  
  // 2. INPUTS STYLING & MERGING
  var inputStyle = { bg: "#fff2cc", border: true };

    // Re-merge Cells
  var nameCell = sheet.getRange(setRow + 2, 4, 1, 3).mergeAcross();
  var addrCell = sheet.getRange(setRow + 3, 4, 2, 3).merge();
  var stateCell = sheet.getRange(setRow + 5, 4, 1, 3).mergeAcross();

  // Style Cells
  [nameCell, addrCell, stateCell].forEach(c => {
    c.setBackground(inputStyle.bg).setBorder(true,true,true,true,null,null).setFontWeight("bold");
  });
  addrCell.setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top");

  // RESTORE VALUES (Use backup, or default)
  nameCell.setValue(savedSettings.name || "My AI Finance Corp");
  addrCell.setValue(savedSettings.addr || "123, Tech Park, New Delhi");
  stateCell.setValue(savedSettings.state || "Delhi (07)");

  // 4. DROPDOWN
  var stateList = getIndianStateDropdownList();
  var rule = SpreadsheetApp.newDataValidation().requireValueInList(stateList, true).build();
  stateCell.setDataValidation(rule);

  // --- D. RIGHT COLUMN: USER GUIDE ---
  var guideCol = 8; 
  sheet.getRange(setRow, guideCol).setValue("📘 QUICK START GUIDE").setFontWeight("bold").setFontSize(14).setFontColor("#2c3e50");

    var steps = [
    ["1. Initial Setup", "Go to Finance Tool > 🚀 Initialize System. Authorize the script."],
    ["2. Configure", "Update your Company Name and State on the left."],
    ["3. Databases", "Add clients in '👥 Clients' and items in '📦 Products'."],
    ["4. Create", "Go to '🧾 Create Invoice', select a client, and add items."],
    ["5. Send", "Click the Checkbox in 'ACTIONS' panel to Email PDF."],
    ["6. History", "Track status in '📊 Invoice History'. Change status to 'Paid' to close"]
  ];

  for (var i = 0; i < steps.length; i++) {
    var r = setRow + 2 + (i * 2);
    sheet.getRange(r, guideCol + 1, 1, 4).breakApart(); 
    sheet.getRange(r, guideCol).setValue(steps[i][0]).setFontWeight("bold");
    sheet.getRange(r, guideCol + 1, 1, 4).merge().setValue(steps[i][1]).setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
    sheet.getRange(r, guideCol, 1, 5).setBackground(i % 2 == 0 ? "#f9f9f9" : "white");
  }

  // --- E. LAYOUT & WIDTHS ---
  sheet.setColumnWidth(1, 20);
  sheet.setColumnWidth(2, 20);  
  sheet.setColumnWidth(3, 120); 
  sheet.setColumnWidth(13, 20); 

  // --- F. OUTER BORDER (The Final Touch) ---
  var lastRow = setRow + 2 + (steps.length * 2); 
  sheet.getRange(2, 2, lastRow - 1, 12).setBorder(true, true, true, true, false, false, "#2c3e50", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  setupCard(sheet, 5, 3, "TOTAL INVOICED", `=SUM('${SHEET_NAME_TRACKER}'!D:D)`, "#2ecc71", "💰"); 
  setupCard(sheet, 5, 7, "INVOICES SENT", `=COUNTA('${SHEET_NAME_TRACKER}'!A2:A)`, "#3498db", "📄"); 
  setupCard(sheet, 5, 11, "PENDING PAYMENTS", `=SUMIFS('${SHEET_NAME_TRACKER}'!D:D, '${SHEET_NAME_TRACKER}'!G:G, "<>Paid")`, "#e74c3c", "⏳");

  // Protect Structure
  var protection = sheet.protect().setDescription("Dashboard Structure");
  var unprotected = [nameCell, addrCell, stateCell]; 
  protection.setUnprotectedRanges(unprotected);
  protection.setWarningOnly(true);
}

/**
 * HELPER: Creates a "Pro-Style" Stat Card
 * Features: Icons, Colored Values, and crisp borders.
 */
function setupCard(sheet, row, col, title, formula, color, icon) {
  // 1. SAFETY: Clear the area
  var cardArea = sheet.getRange(row, col, 2, 2);
  cardArea.breakApart(); 

  // 2. BORDER: Use the theme color for the outline
  cardArea.setBorder(true, true, true, true, true, true, color, SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 3. HEADER (Top Row)
  sheet.getRange(row, col, 1, 2).merge()
       .setValue(icon + " " + title) // Add Icon
       .setBackground(color)
       .setFontColor("white")
       .setFontWeight("bold")
       .setFontSize(10)
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");
       
  // 4. VALUE (Bottom Row)
  var valueCell = sheet.getRange(row + 1, col, 1, 2).merge();
  valueCell.setFormula(formula)
       .setBackground("white")
       .setFontColor(color) // <--- Number matches the header color
       .setFontSize(20)     // <--- Much bigger font
       .setFontWeight("bold")
       .setHorizontalAlignment("center")
       .setVerticalAlignment("middle");

  // 5. FORMATTING
  if (title.includes("REVENUE") || title.includes("PAYMENTS") || title.includes("INVOICED")) {
    // Force Indian Formatting: "₹ 1,50,000.00"
    valueCell.setNumberFormat('[$₹] #,##,##0.00');
  } else {
    valueCell.setNumberFormat("0");
  }
}

/**
 * SETUP: PRODUCT MASTER (Clean Table Style + Orange Theme)
 */
function setupProductMaster() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_PRODUCTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_PRODUCTS);
  }
  
  // 1. HEADERS
  // Indices for VLOOKUP later: 1=Name, 2=Desc, 3=HSN, 4=Rate, 5=GST
  var headers = [["Item Name", "Description", "HSN/SAC", "Rate", "GST Rate"]];
  sheet.getRange("A1:E1").setValues(headers)
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground("#e67e22") // Orange Theme (Matches Tab Color)
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 2. STYLING
  var numRows = sheet.getMaxRows();
  
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange("A1:E" + numRows).createFilter();

  // Zebra Stripes
  var bandings = sheet.getBandings();
  for (var i = 0; i < bandings.length; i++) { bandings[i].remove(); }
  sheet.getRange("A2:E" + numRows).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  // 4. FREEZE & WIDTHS
  sheet.setFrozenRows(1); 
  sheet.setRowHeight(1, 35); 
  
  sheet.setColumnWidth(1, 200); // Item Name (Bold)
  sheet.setColumnWidth(2, 250); // Description
  sheet.setColumnWidth(3, 120); // HSN
  sheet.setColumnWidth(4, 120); // Rate
  sheet.setColumnWidth(5, 100); // GST

  // 3. FORMATTING (Safe to re-apply)
  sheet.getRange("A2:A" + numRows).setFontWeight("bold");
  sheet.getRange("D2:D" + numRows).setNumberFormat("##,##,##0.00");
  sheet.getRange("E2:E" + numRows).setNumberFormat("0%");

  // Add Dropdown for common GST rates
  var gstRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["0%", "5%", "12%", "18%", "28%"], true)
    .setAllowInvalid(true).build();
  sheet.getRange("E2:E" + numRows).setDataValidation(gstRule);

  // 6. SAMPLE DATA (If empty)
if (sheet.getLastRow() === 1) {
      var products = [
      ["Web Development Service", "Full Stack Website Dev (5 Pages)", "998314", 25000, 0.18],
      ["Hosting Charges (Yearly)", "Cloud Server AWS t2.micro", "998315", 8500, 0.18],
      ["Digital Marketing Retainer", "SEO & Social Media Management", "9983", 35000, 0.18],
      ["Business Consultation", "Financial Advisory (Per Hour)", "9982", 5000, 0.18],
      ["Accounting Software", "1 Year Subscription License", "997331", 12000, 0.18],
      ["Laptop Stand (Metal)", "Aluminum Adjustable Stand", "9403", 1200, 0.18],
      ["Freight Charges", "Transport & Logistics", "9965", 500, 0.05]
    ];

    sheet.getRange(2, 1, products.length, 5).setValues(products);
  }

  // 5. PROTECTION
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) { protections[i].remove(); }
  
  var protection = sheet.protect().setDescription("Product Header Protection");
  protection.setUnprotectedRanges([sheet.getRange("A2:E")]); 
  protection.setWarningOnly(true);
}

function setupInvoiceTracker() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(SHEET_NAME_TRACKER);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_NAME_TRACKER);
    }

    var headers = [
        "Invoice No", "Client Name", "Email", "Amount", "Invoice Date",
        "Due Date", "Status", "Follow-up Value", "Follow-up Unit",
        "Last Follow-up Sent", "Next Follow-up At", "Notes"
    ];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight("bold").setBackground("#d9ead3").setFontSize(10)
       .setHorizontalAlignment("center").setVerticalAlignment("middle")
       .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 3. RE-APPLY VALIDATION (Safe for existing data)
  var numRows = sheet.getMaxRows();
  
  // Status Dropdown (Col G / 7)
  var statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Draft", "Ready", "Sent", "Paid", "Stop Follow-up"], true).build();
  sheet.getRange(2, 7, numRows, 1).setDataValidation(statusRule);

  // Unit Dropdown (Col I / 9)
  var unitRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(["Minutes", "Hours", "Days"], true).build();
  sheet.getRange(2, 9, numRows, 1).setDataValidation(unitRule);
    sheet.setFrozenRows(1);

  // 4. COLUMN WIDTHS
  sheet.setColumnWidth(1, 100); // Inv No
  sheet.setColumnWidth(2, 200); // Name
  sheet.setColumnWidth(3, 200); // Email
  sheet.setColumnWidth(4, 100); // Amount
  sheet.setColumnWidth(5, 100); // Date
  sheet.setColumnWidth(6, 100); // Due Date
  sheet.setColumnWidth(7, 120); // Status
  sheet.setColumnWidth(8, 50);  // Value
  sheet.setColumnWidth(9, 80);  // Unit
  sheet.setColumnWidth(10, 150); // Last Sent
  sheet.setColumnWidth(11, 150); // Next Followup
  sheet.setColumnWidth(12, 300); // Notes
  
  // 5. DATA FORMATTING (The Fix)
  // Col D (Amount): Indian Currency
  sheet.getRange("D2:D").setNumberFormat("##,##,##0.00");
  
  // Col E, F, J, K (Dates): Date Format
  var dateFormat = "dd-MMM-yyyy HH:mm";
  sheet.getRange("E2:F").setNumberFormat("dd-MMM-yyyy"); // Inv Date, Due Date (No time needed)
  sheet.getRange("J2:K").setNumberFormat(dateFormat);    // Tracking timestamps (Time needed)

  // Col H (Follow-up Value): PLAIN NUMBER (Fixes the 1900 issue)
  sheet.getRange("H2:H").setNumberFormat("0"); 

  // 5. PROTECTION (Header Only)
  var protection = sheet.protect().setDescription("History Headers");
  protection.setUnprotectedRanges([sheet.getRange("A2:L")]);
  protection.setWarningOnly(true);  
}

/**
 * MAIN SETUP: Coordinator Function
 */
function setupAdvancedStructure() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  // 1. Setup Backend DB
  setupItemsDatabase(ss);

  // 2. Prepare UI Sheet (Clear & Reset)
  var uiSheet = ss.getSheetByName(SHEET_NAME_UI);
  if (!uiSheet) {
    uiSheet = ss.insertSheet(SHEET_NAME_UI);
    ss.setActiveSheet(uiSheet);
    ss.moveActiveSheet(1);
  }
  var protection = uiSheet.getProtections(SpreadsheetApp.ProtectionType.SHEET)[0];
  if (protection) { protection.remove(); }
  uiSheet.clear(); 
  uiSheet.setHiddenGridlines(true);

  // 3. Build Components
  var startRow = 4;
  setupHeaderAndInputs(uiSheet, startRow);
  var tableInfo = setupTableStructure(ss, uiSheet, startRow); // Returns dataStart & numRows
  setupFooterAndBorders(uiSheet, tableInfo.dataStart, tableInfo.numRows);

  // 4. Inject Logic & Actions
  setupActionPanel(uiSheet);
  
  // THE MASTER SWITCH: This applies all VLOOKUPs, Maths, and Totals
  restoreInvoiceFormulas();

  protectSheetLayout(uiSheet, tableInfo.dataStart, tableInfo.numRows);
}

/**
 * SETUP: ITEMS ARCHIVE (Basic styling & Widths)
 */
function setupItemsDatabase(ss) {
  var itemSheet = ss.getSheetByName(SHEET_NAME_ITEMS);
  if (!itemSheet) {
    itemSheet = ss.insertSheet(SHEET_NAME_ITEMS);
  }
  
  // 1. CLEAR STYLING
  itemSheet.clearFormats();
  
  // 2. HEADERS
  // Headers: [InvNo, S.No, Item, HSN, Qty, Rate, Disc, Taxable, GST%, CGST, SGST, IGST, Total]
  var dbHeaders = [["Invoice No", "S. No.", "Item Name", "HSN/SAC", "Qty", "Rate", "Disc", "Taxable Val", "GST Rate", "CGST", "SGST", "IGST", "Total"]];
  
  var headerRange = itemSheet.getRange("A1:M1");
  headerRange.setValues(dbHeaders)
    .setFontWeight("bold")
    .setFontSize(10)
    .setBackground("#34495e") // Dark Blue Theme
    .setFontColor("white")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle");

  // 3. FREEZE & FILTER
  itemSheet.setFrozenRows(1);
  if (itemSheet.getFilter()) itemSheet.getFilter().remove();
  itemSheet.getRange("A1:M1000").createFilter(); // Apply filter to header

  // 4. COLUMN WIDTHS (Context-Aware)
  itemSheet.setColumnWidth(1, 100);  // Invoice No
  itemSheet.setColumnWidth(2, 50);   // S.No (Narrow)
  itemSheet.setColumnWidth(3, 250);  // Item Name (Wide)
  itemSheet.setColumnWidth(4, 80);   // HSN
  itemSheet.setColumnWidth(5, 50);   // Qty
  itemSheet.setColumnWidth(6, 80);   // Rate
  itemSheet.setColumnWidth(7, 60);   // Disc
  itemSheet.setColumnWidth(8, 90);   // Taxable
  itemSheet.setColumnWidth(9, 60);   // GST %
  itemSheet.setColumnWidth(10, 70);  // CGST
  itemSheet.setColumnWidth(11, 70);  // SGST
  itemSheet.setColumnWidth(12, 70);  // IGST
  itemSheet.setColumnWidth(13, 100); // Total

  // 5. ZEBRA STRIPING (For Readability)
  // Remove old bands first
  var bandings = itemSheet.getBandings();
  for (var i = 0; i < bandings.length; i++) {
    bandings[i].remove();
  }
  // Apply new grey banding to the data area
  itemSheet.getRange("A2:M1000").applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);
}

function setupHeaderAndInputs(sheet, startRow) {
  // --- MAIN HEADER ---
  sheet.getRange("B2:M2").merge().setValue("TAX INVOICE")
    .setFontSize(22).setFontWeight("bold").setFontColor("#2c3e50")
    .setHorizontalAlignment("center").setVerticalAlignment("middle");

  // --- STYLING ---
  var labelStyle = { weight: "bold", size: 10, color: "#444" };
  var inputStyle = { bg: "#fff2cc", border: true, size: 10 };

  // --- LABELS ---
  sheet.getRange(startRow, 3).setValue("Customer Name:").setFontWeight(labelStyle.weight);
  sheet.getRange(startRow + 1, 3).setValue("Email:").setFontWeight(labelStyle.weight);
  sheet.getRange(startRow + 2, 3).setValue("GSTIN:").setFontWeight(labelStyle.weight);
  sheet.getRange(startRow + 3, 3).setValue("Billing Address:").setFontWeight(labelStyle.weight);
  sheet.getRange(startRow + 5, 3).setValue("Place of Supply:").setFontWeight(labelStyle.weight);

  // --- INPUT FIELDS ---
  var inputs = [
    sheet.getRange(startRow, 4, 1, 4).merge(),    // Name
    sheet.getRange(startRow + 1, 4, 1, 4).merge(),// Email
    sheet.getRange(startRow + 2, 4, 1, 4).merge(),// GSTIN
    sheet.getRange(startRow + 3, 4, 2, 4).merge(),// Address
    sheet.getRange(startRow + 5, 4, 1, 4).merge() // Place
  ];

  inputs.forEach(cell => {
    cell.setBackground(inputStyle.bg).setBorder(true, true, true, true, null, null)
        .setFontSize(inputStyle.size).setVerticalAlignment("middle");
  });
  inputs[3].setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP).setVerticalAlignment("top"); // Address

  // Client Dropdown
  var rule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_CLIENTS).getRange("A2:A"), true)
    .setAllowInvalid(true).build();
  inputs[0].setDataValidation(rule).setFontWeight("bold");

  // --- RIGHT SIDE META DATA ---
  // Values
  sheet.getRange(startRow, 11, 1, 3).merge().setValue(new Date()).setNumberFormat("dd-MMM-yyyy").setHorizontalAlignment("left");
  sheet.getRange(startRow + 1, 11, 1, 3).merge().setValue("INV-001").setFontWeight("bold").setFontColor("#c0392b").setHorizontalAlignment("left");
  
  // Due Date (Formula Removed - Just Formatting)
  sheet.getRange(startRow + 2, 11, 1, 3).merge()
       .setNumberFormat("dd-MMM-yyyy").setBackground("#fff2cc")
       .setBorder(true, true, true, true, null, null).setHorizontalAlignment("left");
}

function setupTableStructure(ss, sheet, startRow) {
  var tableHeadRow = startRow + 7;
  var headers = [["S. No.", "Item Name", "HSN/SAC", "Qty", "Rate", "Disc", "Taxable Val", "GST Rate", "CGST", "SGST", "IGST", "Total"]];
  
  // Table Header
  sheet.getRange(tableHeadRow, 2, 1, 12).setValues(headers)
    .setFontWeight("bold").setBackground("#2c3e50").setFontColor("white").setHorizontalAlignment("center")
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  var numRows = 10;
  var dataStart = tableHeadRow + 1;
  var indianFormat = "##,##,##0.00";

  // Product Dropdown Rule
  var productRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(SHEET_NAME_PRODUCTS).getRange("A2:A"), true)
    .setAllowInvalid(true).build();

  // FORMATTING LOOP (Visuals only)
  for (var i = 0; i < numRows; i++) {
    var r = dataStart + i;
    
    // Backgrounds
    sheet.getRange(r, 2).setBackground("#f3f3f3").setHorizontalAlignment("center"); // S.No
    sheet.getRange(r, 3).setBackground("white").setDataValidation(productRule);     // Item (Input)
    sheet.getRange(r, 4).setBackground("#f9f9f9").setFontColor("#555");             // HSN (ReadOnly)
    sheet.getRange(r, 5).setBackground("white");                                    // Qty (Input)
    sheet.getRange(r, 6).setBackground("white");                                    // Rate (Input/Auto)
    sheet.getRange(r, 7).setBackground("white");                                    // Disc (Input)
    sheet.getRange(r, 8).setBackground("#f3f3f3");                                  // Taxable
    sheet.getRange(r, 9).setBackground("white").setNumberFormat("0%");              // GST Rate
    sheet.getRange(r, 10, 1, 4).setBackground("#f3f3f3");                           // Taxes + Total
    sheet.getRange(r, 13).setFontWeight("bold");
  }

  // Number Formats
  sheet.getRange(dataStart, 6, numRows, 3).setNumberFormat(indianFormat); // Rate, Disc, Taxable
  sheet.getRange(dataStart, 10, numRows, 4).setNumberFormat(indianFormat); // Taxes, Total

  // Column Widths
  sheet.setColumnWidth(1, 20); sheet.setColumnWidth(2, 40); sheet.setColumnWidth(3, 200); 
  sheet.setColumnWidth(4, 70); sheet.setColumnWidth(5, 50); sheet.setColumnWidth(6, 70); 
  sheet.setColumnWidth(7, 60); sheet.setColumnWidth(8, 90); sheet.setColumnWidth(9, 60); 
  sheet.setColumnWidth(10, 70); sheet.setColumnWidth(11, 70); sheet.setColumnWidth(12, 70); 
  sheet.setColumnWidth(13, 100);

  return { dataStart: dataStart, numRows: numRows };
}

function setupFooterAndBorders(sheet, dataStart, numRows) {
  var footerRow = dataStart + numRows;
  var indianFormat = "##,##,##0.00";

  // 1. Footer Visuals
  sheet.getRange(footerRow, 2, 1, 6).merge().setValue("Total").setFontWeight("bold").setHorizontalAlignment("right").setBackground("#fff2cc");
  
  // Styling the Sum Cells (But NOT setting formulas)
  var sumCols = [8, 10, 11, 12, 13];
  sumCols.forEach(col => {
    sheet.getRange(footerRow, col).setFontWeight("bold").setBackground("#fff2cc").setNumberFormat(indianFormat);
  });
  sheet.getRange(footerRow, 9).setBackground("#fff2cc");

  // 2. Summary Block Visuals
  var summaryStart = footerRow + 2;
  
  sheet.getRange(summaryStart, 12).setValue("Taxable Amount").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange(summaryStart + 1, 12).setValue("Total Tax").setFontWeight("bold").setHorizontalAlignment("right");
  sheet.getRange(summaryStart + 3, 12).setValue("Invoice Total").setFontWeight("bold").setHorizontalAlignment("right");

  // Format the Value Cells (Formulas removed)
  sheet.getRange(summaryStart, 13).setNumberFormat(indianFormat).setFontWeight("bold");
  sheet.getRange(summaryStart + 1, 13).setNumberFormat(indianFormat).setFontWeight("bold");
  sheet.getRange(summaryStart + 3, 13).setBackground("#fff2cc").setFontWeight("bold").setBorder(true,true,true,true,null,null).setNumberFormat(indianFormat);

  // Amount in Words Visuals
  sheet.getRange(summaryStart + 3, 3).setValue("Total amount (in words):").setFontWeight("bold");
  sheet.getRange(summaryStart + 3, 4, 1, 6).merge().setHorizontalAlignment("left").setFontStyle("italic");

  // 3. Borders
  var borderEndRow = summaryStart + 4; 
  sheet.getRange(dataStart, 2, numRows + 1, 12).setBorder(true, true, true, true, true, true, "#999999", SpreadsheetApp.BorderStyle.SOLID);
  sheet.getRange(2, 2, borderEndRow - 1, 12).setBorder(true, true, true, true, false, false, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);
}

function protectSheetLayout(sheet, dataStart, numRows) {
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) protections[i].remove();

  var protection = sheet.protect().setDescription('Invoice Layout');
  var unprotected = [];
  
  // Client Name Input (C4:F4) <--- Only Name is "Input", others are Formulas now!
  unprotected.push(sheet.getRange(4, 3, 1, 4));
  
  // OPTIONAL: Unlock other client fields if you want to override formulas manually
  // C5:F9 (Email, GST, Address, Place)
  unprotected.push(sheet.getRange(5, 3, 5, 4));

  // Due Date
  unprotected.push(sheet.getRange(6, 11, 1, 3));
  
  // Table Inputs
  unprotected.push(sheet.getRange(dataStart, 3, numRows, 5)); // Item to Disc
  unprotected.push(sheet.getRange(dataStart, 9, numRows, 1)); // GST Rate

  protection.setUnprotectedRanges(unprotected);
  protection.setWarningOnly(true);

  unprotected.push(sheet.getRange("O5:07"));

  protection.setUnprotectedRanges(unprotected);
}

function setupSettingsSheet() {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName("Settings");
    var stateList = getIndianStateDropdownList();

    if (!sheet) {
        sheet = ss.insertSheet("Settings");
        var defaults = [
            ["Company Name", "My AI Finance Corp"],
            ["My State", "Delhi (07)"],
            ["Currency Symbol", "₹"],
            ["Email Subject", "Invoice Notification"],
            ["Email Body Intro", "This is a notification regarding your invoice."],
            ["Email Footer", "Thank you for your business."]
        ];

        sheet.getRange("A1:B1").setValues([["Setting Name", "Value"]]).setFontWeight("bold");
        sheet.getRange(2, 1, defaults.length, 2).setValues(defaults);

        // Formatting help
        sheet.setColumnWidth(1, 150);
        sheet.setColumnWidth(2, 400);
        sheet.getRange("A:B").setVerticalAlignment("top");

        // APPLY DROPDOWN to "My State" (Row 3, Col 2)
        // We assume "My State" is the 2nd item in defaults, so it's at Row 3
        var rule = SpreadsheetApp.newDataValidation().requireValueInList(stateList, true).build();
        sheet.getRange("B3").setDataValidation(rule);
    }
}

function setupLogSheet() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_LOGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_LOGS);
  }

  // 1. HEADERS
  var headers = [["Timestamp", "Action", "Details"]];
  sheet.getRange("A1:C1").setValues(headers)
       .setFontWeight("bold").setBackground("#e6b8af") // Light Red Theme
       .setHorizontalAlignment("center").setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID);

  // 2. STYLING
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 200); // Time
  sheet.setColumnWidth(2, 150); // Action
  sheet.setColumnWidth(3, 500); // Details (Very Wide)

  // 4. PROTECTION (Strict)
  // Logs should ideally NOT be edited by users manually.
  var protection = sheet.protect().setDescription("Audit Logs");
  // We leave the whole sheet protected (Warning Only for owner)
  protection.setWarningOnly(true);
}

function logAction(action, details) {
    var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_LOGS);
    if (!sheet) {
        setupLogSheet();
        sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME_LOGS);
    }
    sheet.appendRow([new Date(), action, details]);
}

/**
 * SETUP: CLIENTS DB (Clean Table Style)
 */
function setupCustomerDatabase() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_CLIENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME_CLIENTS);
  }
  
  // 1. HEADERS (The 5 Required Columns)
  var headers = [["Customer Name", "Email", "GSTIN", "Billing Address", "Place of Supply"]];
  sheet.getRange("A1:E1").setValues(headers)
    .setFontWeight("bold")
    .setFontSize(11)
    .setBackground("#f1c40f") 
    .setFontColor("#2c3e50")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setBorder(true, true, true, true, true, true, "black", SpreadsheetApp.BorderStyle.SOLID_MEDIUM);

  // 2. REPAIR STYLING
  var numRows = sheet.getMaxRows(); // Use actual sheet size
  
  if (sheet.getFilter()) sheet.getFilter().remove();
  sheet.getRange("A1:E" + numRows).createFilter();

  var bandings = sheet.getBandings();
  for (var i = 0; i < bandings.length; i++) { bandings[i].remove(); }
  sheet.getRange("A2:E" + numRows).applyRowBanding(SpreadsheetApp.BandingTheme.LIGHT_GREY);

  // 3. FREEZE & WIDTHS
  sheet.setFrozenRows(1); // Keep headers visible
  sheet.setRowHeight(1, 35); 
  
  // Custom Widths (As requested)
  sheet.setColumnWidth(1, 250); // A: Name (Big)
  sheet.setColumnWidth(2, 200); // B: Email
  sheet.setColumnWidth(3, 150); // C: GSTIN
  sheet.setColumnWidth(4, 300); // D: Address (Very Big)
  sheet.setColumnWidth(5, 150); // E: State

  sheet.getRange("A2:A" + numRows).setFontWeight("bold");

  // 4. DATA VALIDATION & SAMPLE
  // Email (Col B)
  var emailRule = SpreadsheetApp.newDataValidation().requireTextIsEmail().setAllowInvalid(true).build();
  sheet.getRange(2, 2, numRows, 1).setDataValidation(emailRule);

  // State Dropdown (Col E)
  var stateList = getIndianStateDropdownList();
  var stateRule = SpreadsheetApp.newDataValidation().requireValueInList(stateList, true).build();
  sheet.getRange(2, 5, numRows, 1).setDataValidation(stateRule);
  
  // Sample Data (Only if empty)
  if (sheet.getLastRow() === 1) { // Only if just header exists
    var customers = [
      ["Tech Solutions Pvt Ltd", "client.tech@example.com", "07AAACT0000A1Z5", "123 Tech Park, Okhla Phase III, New Delhi", "Delhi (07)"],
      ["Kabir Jewellers", "accounts@kabirjewels.com", "37AABCU9603R1Z2", "45 Airport Road, Begumpet, Hyderabad", "Andhra Pradesh (37)"],
      ["Blue Chip Marketing", "billing@bluechip.co.in", "27AABCB1234P1Z1", "99 Nariman Point, Mumbai", "Maharashtra (27)"],
      ["Green Leaf Organics", "purchase@greenleaf.org", "29AADCG5678L1Z9", "88 Indiranagar, Bangalore", "Karnataka (29)"],
      ["Rahul Traders", "rahul.traders@testmail.com", "09AABCR9012K1Z3", "12 Civil Lines, Kanpur", "Uttar Pradesh (09)"]
    ];

    sheet.getRange(2, 1, customers.length, 5).setValues(customers);
  }

  // 5. PROTECTION
  var protections = sheet.getProtections(SpreadsheetApp.ProtectionType.SHEET);
  for (var i = 0; i < protections.length; i++) { protections[i].remove(); }
  
  var protection = sheet.protect().setDescription("Header Protection");
  protection.setUnprotectedRanges([sheet.getRange("A2:E")]); 
  protection.setWarningOnly(true);
}

/**
 * HELPER: Creates the "Multi-Action Panel" in Column O
 */
function setupActionPanel(sheet) {
  // 2. Setup the "Panel" visual background (Rows 4 to 8)
  var panelRange = sheet.getRange("O4:P8");
  panelRange.setBackground("white")
    .setBorder(true, true, true, true, true, true, "#eeeeee", SpreadsheetApp.BorderStyle.SOLID);

  // 3. The Header
  sheet.getRange("O4").setValue("ACTIONS")
    .setFontWeight("bold").setFontSize(10).setHorizontalAlignment("center").setBackground("#f3f3f3");

  // --- BUTTON 1: SAVE & SEND (Row 5) ---
  sheet.getRange("O5").insertCheckboxes().uncheck().setHorizontalAlignment("center");
  sheet.getRange("P5").setValue("Save & Send Email")
    .setFontSize(9).setFontColor("#2c3e50").setFontWeight("bold").setVerticalAlignment("middle");

  // --- BUTTON 2: SAVE AS DRAFT (Row 6) ---
  sheet.getRange("O6").insertCheckboxes().uncheck().setHorizontalAlignment("center");
  sheet.getRange("P6").setValue("Save as Draft (No Email)")
    .setFontSize(9).setFontColor("#2c3e50").setVerticalAlignment("middle");

  // --- BUTTON 3: CLEAR FORM (Row 7) ---
  sheet.getRange("O7").insertCheckboxes().uncheck().setHorizontalAlignment("center");
  sheet.getRange("P7").setValue("Clear / Reset Form")
    .setFontSize(9).setFontColor("#c0392b").setVerticalAlignment("middle"); // Red text for caution

  // 4. Set Column Widths
  sheet.setColumnWidth(14, 20);  // N (Padding)
  sheet.setColumnWidth(15, 40);  // O (Checkbox)
  sheet.setColumnWidth(16, 140); // P (Labels)
}

function organizeSheets() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  
  var layout = [
    { name: SHEET_NAME_DASHBOARD, color: "#2c3e50" }, // Dark Grey/Blue
    { name: SHEET_NAME_UI, color: "#3498db" },        // Blue
    { name: SHEET_NAME_TRACKER, color: "#9b59b6" },   // Purple
    { name: SHEET_NAME_CLIENTS, color: "#f1c40f" },   // Yellow
    { name: SHEET_NAME_PRODUCTS, color: "#e67e22" },  // Orange
    { name: SHEET_NAME_ITEMS, color: "#34495e" },     // Dark Blue
    { name: SHEET_NAME_LOGS, color: "#c0392b" }       // Red
  ];

  for (var i = 0; i < layout.length; i++) {
    var sheet = ss.getSheetByName(layout[i].name);
    if (sheet) {
      sheet.activate();
      ss.moveActiveSheet(i + 1);
      sheet.setTabColor(layout[i].color);
    }
  }
}

/**
 * MASTER HELPER: Re-applies ALL formulas (Client, Table, Footer)
 * This makes the sheet "Self-Healing."
 */
function restoreInvoiceFormulas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_UI);
  
  // --- A. CONSTANTS (Must match Setup Layout) ---
  var startRow = 4;           // Where Inputs Start
  var tableHeadRow = startRow + 7; 
  var dataStart = tableHeadRow + 1; // Row 12
  var numRows = 10;
  var footerRow = dataStart + numRows; // Row 22
  
  // --- B. CLIENT DETAILS (VLOOKUPs) ---
  // Name is at D4. We look up details based on D4.
  // Email (Col 2), GSTIN (Col 3), Address (Col 4), State (Col 5)
  sheet.getRange(startRow + 1, 4).setFormula(`=IFERROR(VLOOKUP(D${startRow}, '${SHEET_NAME_CLIENTS}'!A:E, 2, FALSE), "")`); // Email
  sheet.getRange(startRow + 2, 4).setFormula(`=IFERROR(VLOOKUP(D${startRow}, '${SHEET_NAME_CLIENTS}'!A:E, 3, FALSE), "")`); // GSTIN
  sheet.getRange(startRow + 3, 4).setFormula(`=IFERROR(VLOOKUP(D${startRow}, '${SHEET_NAME_CLIENTS}'!A:E, 4, FALSE), "")`); // Address
  sheet.getRange(startRow + 5, 4).setFormula(`=IFERROR(VLOOKUP(D${startRow}, '${SHEET_NAME_CLIENTS}'!A:E, 5, FALSE), "")`); // Place

  // --- C. META DATA ---
  // Due Date (K6) = Invoice Date (K4) + 15 Days
  sheet.getRange(startRow + 2, 11).setFormula(`=K${startRow} + 15`);

    // --- D. TABLE CALCULATIONS (The Smart Logic) ---
  // We define the logic strings once to keep code clean
  // Condition: IF(PlaceOfSupply = MyState, Intra, Inter)
  // My State is at '🏠 Dashboard'!C14
  // Place of Supply is at $C$9 (Merged C9:F9, accessed via C9)
  
  var myStateRef = `'${SHEET_NAME_DASHBOARD}'!$D$14`;
  var custStateRef = `$D$9`;

  for (var i = 0; i < numRows; i++) {
    var r = dataStart + i;
    
    // 1. S.No
    sheet.getRange(r, 2).setFormula(`=IF(C${r}="", "", ${i + 1})`);

    // 2. Default Quantity Formula
    sheet.getRange(r, 5).setFormula(`=IF(C${r}="", "", 1)`);
    
    // 2. HSN (Col D)
    sheet.getRange(r, 4).setFormula(`=IFERROR(VLOOKUP(C${r}, '${SHEET_NAME_PRODUCTS}'!A:E, 3, FALSE), "")`);
    
    // 3. Rate (Col F)
    sheet.getRange(r, 6).setFormula(`=IFERROR(VLOOKUP(C${r}, '${SHEET_NAME_PRODUCTS}'!A:E, 4, FALSE), "")`);
    
    // 4. Taxable Val (Col H) -> (Qty * Rate) - Disc
    sheet.getRange(r, 8).setFormula(`=IF(C${r}="", "", (E${r}*F${r})-G${r})`);
    
    // 5. GST Rate (Col I)
    sheet.getRange(r, 9).setFormula(`=IFERROR(VLOOKUP(C${r}, '${SHEET_NAME_PRODUCTS}'!A:E, 5, FALSE), "")`);

    // 6. Taxes (CGST, SGST, IGST)
    // CGST: If Row Empty -> "". Else If States Diff -> 0. Else Calc.
    sheet.getRange(r, 10).setFormula(`=IF(C${r}="", "", IF(${custStateRef} <> ${myStateRef}, 0, H${r} * (I${r}/2)))`);
    
    // SGST: Same as CGST
    sheet.getRange(r, 11).setFormula(`=IF(C${r}="", "", IF(${custStateRef} <> ${myStateRef}, 0, H${r} * (I${r}/2)))`);
    
    // IGST: If Row Empty -> "". Else If States Same -> 0. Else Calc.
    sheet.getRange(r, 12).setFormula(`=IF(C${r}="", "", IF(${custStateRef} = ${myStateRef}, 0, H${r} * I${r}))`);
    
    // 7. Line Total
    sheet.getRange(r, 13).setFormula(`=IF(C${r}="", "", SUM(H${r}, J${r}, K${r}, L${r}))`);
  }

  // --- E. FOOTER TOTALS ---
  var colMap = {8:"H", 10:"J", 11:"K", 12:"L", 13:"M"};
  // Sum Columns H, J, K, L, M
  for (var colIndex in colMap) {
    var letter = colMap[colIndex];
    sheet.getRange(footerRow, parseInt(colIndex)).setFormula(`=SUM(${letter}${dataStart}:${letter}${footerRow-1})`);
  }

  // --- F. SUMMARY BLOCK ---
  // Taxable Amount (Matches H Total)
  sheet.getRange(footerRow + 2, 13).setFormula(`=H${footerRow}`);
  
  // Total Tax (J+K+L)
  sheet.getRange(footerRow + 3, 13).setFormula(`=J${footerRow} + K${footerRow} + L${footerRow}`);
  
  // Invoice Total (Matches M Total)
  sheet.getRange(footerRow + 5, 13).setFormula(`=M${footerRow}`);

  // --- G. AMOUNT IN WORDS ---
  // Uses M Total
  sheet.getRange(footerRow + 5, 4).setFormula(`=NUMBER_TO_WORDS(M${footerRow})`);
}