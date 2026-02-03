function sendInvoiceEmail(data) {
  // 1. Fetch Configuration
  var companyName = fetchSetting("Company Name") || "My Company Name";

  // 2. Prepare Metadata
  var recipient = data.client.email;
  var subject = "Invoice #" + data.invNo + " from " + companyName; 
  var fileName = "Invoice_" + data.invNo + "_" + data.client.name + ".pdf";
  
  // 3. GENERATE PDF BLOB
  var pdfBlob = createPdfBlob(data, fileName);

  // 4. SAVE TO GOOGLE DRIVE
  var folderName = FOLDER_NAME; // From Config.js
  var folder;
  var folders = DriveApp.getFoldersByName(folderName);

  // Check if folder exists, otherwise create it
  if (folders.hasNext()) {
    folder = folders.next();
  } else {
    folder = DriveApp.createFolder(folderName);
  }

  var file = folder.createFile(pdfBlob);
  var fileUrl = file.getUrl();

  // 5. SEND EMAIL
  var htmlMessage = getBeautifulEmailBody(data);
  var plainBody = "Dear " + data.client.name + ",\n\n" +
             "Please find attached Invoice #" + data.invNo + " for " + data.totals.grandTotal + ".\n\n" +
             "Thank you for your business.\n\n" +
             "Regards,\n" + companyName;
  
  try {
    GmailApp.sendEmail(recipient, subject, plainBody, { // plainBody is fallback
      name: companyName,
      attachments: [pdfBlob],
      htmlBody: htmlMessage
    });
    
    // Log the success
    logAction("Email Sent", "To: " + recipient + " | Drive Link: " + fileUrl);

    return fileUrl;
    
  } catch (e) {
    logAction("Email Failed", "Error: " + e.toString());
    throw new Error("Failed to send email: " + e.toString());
  }
}

function getHtmlTemplate(data, type) {
  var body = fetchSetting("Email Body Intro");
  var footer = fetchSetting("Email Footer");
  var currency = "₹";
  return `<div style="font-family:Arial; padding:20px;"><h2>${data.company}</h2><p>${body}</p><p><strong>${type}</strong>: ${currency}${data.amount}</p><p>${footer}</p></div>`;
}

/**
 * GENERATE PROFESSIONAL PDF
 * Features: Dynamic Columns (Hides/Shows Tax columns based on State)
 */
function createPdfBlob(data, fileName) {
  
  // 1. FETCH SETTINGS & DETERMINE TAX TYPE
  var companyName = fetchSetting("Company Name") || "Finance AI Corp";
  var companyAddress = fetchSetting("Company Address") || "Delhi, India";
  var myState = fetchSetting("My State") || "";
  var currency = "₹";
  
  // Logic: Comparison logic (Normalize strings to avoid case sensitivity)
  // Assuming format "State (Code)" or just "State". Direct string match is safest for now.
  var isIGST = (data.client.state !== myState); 

  // 2. CONSTRUCT DYNAMIC TABLE HEADERS
  var taxHeaders = "";
  if (isIGST) {
    taxHeaders = `<th width="5%">GST%</th> <th width="15%">IGST</th>`;
  } else {
    taxHeaders = `<th width="5%">GST%</th> <th width="10%">CGST</th> <th width="10%">SGST</th>`;
  }

  // 3. BUILD ITEM ROWS HTML
  var itemsHtml = "";
  data.items.forEach(function(item) {
    var descHtml = item.description 
      ? `<div style="font-size: 8pt; font-style: italic; color: #666; margin-top: 2px;">${item.description}</div>` 
      : "";

    // Dynamic Tax Cells
    var taxCells = "";
    if (isIGST) {
      taxCells = `<td class="text-right">${formatMoney(item.gstRate * 100)}%</td>
                  <td class="text-right">${formatMoney(item.igst)}</td>`;
    } else {
      taxCells = `<td class="text-right">${formatMoney(item.gstRate * 100)}%</td>
                  <td class="text-right">${formatMoney(item.cgst)}</td>
                  <td class="text-right">${formatMoney(item.sgst)}</td>`;
    }

    itemsHtml += `
      <tr>
        <td class="text-center">${item.srNo}</td>
        <td>
          <div style="font-weight: bold;">${item.name}</div>
          ${descHtml}
        </td>
        <td class="text-center">${item.hsn}</td>
        <td class="text-center">${item.qty}</td>
        <td class="text-right">${formatMoney(item.rate)}</td>
        <td class="text-right">${formatMoney(item.taxable)}</td>
        ${taxCells}
        <td class="text-right"><strong>${formatMoney(item.total)}</strong></td>
      </tr>
    `;
  });

  // 4. BUILD DYNAMIC FOOTER ROW
  var footerCells = "";
  if (isIGST) {
    footerCells = `
      <td class="text-right"></td>
      <td class="text-right">${formatMoney(data.totals.igst)}</td>`;
  } else {
    footerCells = `
      <td class="text-right"></td>
      <td class="text-right">${formatMoney(data.totals.cgst)}</td>
      <td class="text-right">${formatMoney(data.totals.sgst)}</td>`;
  }

  // 5. BUILD DYNAMIC SUMMARY BLOCK
  var summaryRows = "";
  if (isIGST) {
    summaryRows = `<tr><td>IGST:</td><td>${formatMoney(data.totals.igst)}</td></tr>`;
  } else {
    summaryRows = `
      <tr><td>CGST:</td><td>${formatMoney(data.totals.cgst)}</td></tr>
      <tr><td>SGST:</td><td>${formatMoney(data.totals.sgst)}</td></tr>`;
  }


  // 6. CONSTRUCT FULL HTML
  var htmlTemplate = `
    <html>
      <head>
        <style>
          @page { size: A4; margin: 1.0in 0.5in 0.5in 0.5in; }
          body { font-family: 'Helvetica', 'Arial', sans-serif; font-size: 9pt; color: #333; margin: 0; padding: 0; -webkit-print-color-adjust: exact; print-color-adjust: exact; }
          
          .header-box { width: 100%; border-bottom: 2px solid #c0392b; margin-bottom: 15px; padding-bottom: 5px; }
          .company-name { font-size: 18pt; font-weight: bold; color: #c0392b; margin-bottom: 3px; }
          .title { font-size: 24pt; font-weight: bold; color: #2c3e50; text-align: right; vertical-align: top; }
          
          .info-table { width: 100%; margin-bottom: 20px; border-collapse: collapse; }
          .info-table td { vertical-align: top; padding: 2px 0; } 
          .label { font-weight: bold; color: #555; white-space: nowrap; padding-right: 10px; }
          .section-title { font-size: 10pt; font-weight: bold; color: #2c3e50; border-bottom: 1px solid #ddd; margin-bottom: 5px; padding-bottom: 2px; display: block; }
          .client-name { font-size: 12pt; font-weight: bold; color: #333; margin-bottom: 3px; }

          .items-table { width: 100%; border-collapse: collapse; margin-top: 10px; }
          .items-table th { background-color: #2c3e50 !important; color: white; padding: 6px; font-size: 8pt; border: 1px solid #333; text-align: center; -webkit-print-color-adjust: exact; }
          .items-table td { padding: 6px; border: 1px solid #ccc; font-size: 9pt; }
          .text-center { text-align: center; }
          .text-right { text-align: right; }
          
          .yellow-row { background-color: #fff2cc !important; font-weight: bold; -webkit-print-color-adjust: exact; }
          .yellow-row td { border-top: 2px solid #333; border-bottom: 2px solid #333; }
          
          .summary-section { width: 100%; margin-top: 20px; }
          .summary-table { float: right; width: 45%; border-collapse: collapse; } 
          .summary-table td { padding: 4px; text-align: right; }
          
          .total-row { background-color: #f1c40f !important; font-weight: bold; border-top: 2px solid black; border-bottom: 2px solid black; font-size: 11pt; -webkit-print-color-adjust: exact; }
          .words-section { float: left; width: 50%; margin-top: 10px; font-style: italic; border-top: 1px solid #ccc; padding-top: 5px; }
          .disclaimer { margin-top: 40px; text-align: center; font-size: 8pt; color: #666; clear: both; border-top: 1px solid #eee; padding-top: 10px; }
        </style>
      </head>
      <body>
      
        <div class="header-box">
          <table width="100%">
            <tr>
              <td>
                <div class="company-name">${companyName}</div>
                <div>GSTIN: 07AABCU9603R1Z2</div>
                <div>${companyAddress}</div>
              </td>
              <td class="title">TAX INVOICE</td>
            </tr>
          </table>
        </div>

        <table class="info-table">
          <tr>
            <td width="55%">
              <span class="section-title">Bill To:</span>
              <div style="margin-top: 5px;">
                <div class="client-name">${data.client.name}</div>
                <div>GSTIN: ${data.client.gstin}</div>
                <div>Email: ${data.client.email}</div>
                <div style="margin-top: 5px; line-height: 1.3;">${data.client.address}</div>
              </div>
            </td>
            <td width="45%" style="text-align: right;">
              <table align="right">
                <tr><td class="label">Invoice No:</td><td><strong>${data.invNo}</strong></td></tr>
                <tr><td class="label">Invoice Date:</td><td>${data.date}</td></tr>
                <tr><td class="label">Due Date:</td><td>${data.dueDate}</td></tr>
                <tr><td class="label">Place of Supply:</td><td>${data.client.state}</td></tr>
              </table>
            </td>
          </tr>
        </table>

        <table class="items-table">
          <thead>
            <tr>
              <th width="5%">Sr.</th>
              <th width="30%">Item Description</th>
              <th width="10%">HSN</th>
              <th width="5%">Qty</th>
              <th width="10%">Rate</th>
              <th width="10%">Taxable</th>
              ${taxHeaders} <!-- Dynamic Header -->
              <th width="10%">Total</th>
            </tr>
          </thead>
          <tbody>
            ${itemsHtml} <!-- Dynamic Rows -->
            
            <tr class="yellow-row">
              <td colspan="5" class="text-right">Total:</td>
              <td class="text-right">${formatMoney(data.totals.taxable)}</td>
              ${footerCells} <!-- Dynamic Footer -->
              <td class="text-right">${formatMoney(data.totals.grandTotal)}</td>
            </tr>
          </tbody>
        </table>

        <div class="summary-section">
          <table class="summary-table">
            <tr><td>Taxable Amount:</td><td>${formatMoney(data.totals.taxable)}</td></tr>
            ${summaryRows} <!-- Dynamic Summary Breakdown -->
            <tr class="total-row"><td>Invoice Total:</td><td>${currency} ${formatMoney(data.totals.grandTotal)}</td></tr>
          </table>

          <div class="words-section">
            <strong>Total amount (in words):</strong><br>
            ${data.totals.amountInWords}
          </div>
          <div style="clear: both;"></div>
        </div>

        <div class="disclaimer">
          This is an electronically generated document, no signature is required.
        </div>

      </body>
    </html>
  `;

  var blob = Utilities.newBlob(htmlTemplate, "text/html", fileName);
  return blob.getAs("application/pdf");
}

/**
 * HELPER: Generates the "Modern SaaS" Email Body
 */
function getBeautifulEmailBody(data) {
  var companyName = fetchSetting("Company Name") || "My Finance AI";
  var address = fetchSetting("Company Address") || "";
  var currency = "₹";

  return `
    <div style="background-color: #f6f6f6; padding: 40px 0; font-family: 'Helvetica', Arial, sans-serif;">
      <div style="max-width: 600px; margin: 0 auto; background: white; border-radius: 8px; overflow: hidden; box-shadow: 0 2px 5px rgba(0,0,0,0.1);">
        
        <!-- HEADER -->
        <div style="background-color: #2c3e50; padding: 30px; text-align: center;">
          <h1 style="color: white; margin: 0; font-size: 24px; font-weight: 300;">${companyName}</h1>
        </div>
        
        <!-- CONTENT -->
        <div style="padding: 40px; color: #333; line-height: 1.6;">
          <div style="font-size: 18px; font-weight: bold; margin-bottom: 20px;">Hi ${data.client.name},</div>
          
          <p>Please find attached the invoice for our recent services. We appreciate your continued business.</p>
          
          <!-- HIGHLIGHT BOX -->
          <div style="background-color: #f8f9fa; border-left: 5px solid #3498db; padding: 20px; margin: 25px 0; border-radius: 4px;">
            <table width="100%" cellpadding="0" cellspacing="0">
              <tr>
                <td>
                  <div style="font-size: 14px; margin-bottom: 5px;"><span style="color: #777; font-weight: bold;">Invoice No:</span> <strong>${data.invNo}</strong></div>
                  <div style="font-size: 14px; margin-bottom: 5px;"><span style="color: #777; font-weight: bold;">Date:</span> ${data.date}</div>
                  <div style="font-size: 14px;"><span style="color: #777; font-weight: bold;">Due Date:</span> <strong style="color: #e67e22;">${data.dueDate}</strong></div>
                </td>
                <td align="right" style="vertical-align: top;">
                  <div style="font-size: 11px; color: #999; text-transform: uppercase; letter-spacing: 1px;">Amount Due</div>
                  <div style="font-size: 24px; font-weight: bold; color: #2c3e50; margin-top: 5px;">${currency} ${formatMoney(data.totals.grandTotal)}</div>
                </td>
              </tr>
            </table>
          </div>
          
          <p>The PDF invoice is attached to this email.</p>
          
          <p style="margin-top: 30px;">
            Best regards,<br>
            <strong>Accounts Team</strong><br>
            ${companyName}
          </p>
        </div>
        
        <!-- FOOTER -->
        <div style="background-color: #f4f4f4; padding: 20px; text-align: center; font-size: 12px; color: #999; border-top: 1px solid #eee;">
          ${address}
        </div>
        
      </div>
    </div>
  `;
}

/**
 * HELPER: Formats numbers to Indian Currency (Comma Separated)
 * 123456 -> 1,23,456.00
 */
function formatMoney(amount) {
  if (!amount) return "0.00";
  return amount.toLocaleString('en-IN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
}

function calculateNextFollowUp(sheet, row, freqVal, freqUnit) {
  var now = new Date();
  var nextDate = new Date(now.getTime());
  if (freqUnit === "Minutes") nextDate.setMinutes(nextDate.getMinutes() + freqVal);
  else if (freqUnit === "Days") nextDate.setDate(nextDate.getDate() + freqVal);
  sheet.getRange(row, 11).setValue(nextDate);
}