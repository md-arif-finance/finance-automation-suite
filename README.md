# Automated AR Control System

A powerful, Google Sheets-based Invoice Management System designed to streamline your Accounts Receivable (AR) process. This tool allows you to generate professional PDF invoices, manage clients and products, track payments, and automate email notifications directly from a Google Spreadsheet.

![Manual vs Our Way](docs/images/manual-vs-our-way.jpg)

## 🚀 Get Started

Get started and get the copy of the sheet for this project here:
[Get Started](https://md-arif-finance.github.io/finance-automation-suite/)

## 🚀 Features

*   **📊 Financial Dashboard:** Get a real-time overview of your business health with key metrics like Total Invoiced, Collected Amount, Outstanding Payments, and Overdue Invoices.
*   **🧾 Smart Invoice Generation:**
    *   User-friendly interface to create invoices.
    *   Automatic fetching of Client and Product details.
    *   **Auto-Tax Calculation:** Automatically applies CGST/SGST or IGST based on the client's state.
    *   **Amount in Words:** Automatically converts numerical totals to words (Indian Rupee format).
*   **📧 Email Automation:**
    *   Send professional HTML emails with the Invoice PDF attached in one click.
    *   Customizable email body and subject.
*   **👥 Client & Product Management:** Dedicated sheets to manage your customer database (`👥 Clients`) and product inventory (`📦 Products`) with HSN codes and tax rates.
*   **📝 Invoice History & Tracking:** automatically records every generated invoice in the `📊 Invoice History` sheet for easy tracking of payment status and follow-ups.
*   **🛡️ Secure & Validated:** Includes data validation, protected ranges, and input locking to prevent accidental errors.

## 📖 Usage Guide

### 1. Add Master Data
*   **Clients:** Go to the `👥 Clients` sheet and add your customer details (Name, Email, GSTIN, Address, State).
*   **Products:** Go to the `📦 Products` sheet and list your services/items with HSN codes and unit rates.

### 2. Create an Invoice
1.  Navigate to the `🧾 Create Invoice` sheet.
2.  Select a **Customer Name** from the dropdown.
3.  Add items in the table below (Select Item, Enter Qty, Discount).
4.  The system will automatically calculate Taxable Value, GST, and Total Amount.

### 3. Send or Save
Look for the **ACTIONS** panel on the right side of the Invoice sheet:
*   **Save & Send Email:** Generates a PDF, saves it to Google Drive (`Invoices_Sent` folder), sends it to the client, and records it in History.
*   **Save as Draft:** Saves the invoice to History without sending an email.
*   **Clear / Reset Form:** Wipes the form clean for the next invoice.

### 4. Track Payments
*   Go to `📊 Invoice History` to see all your invoices.
*   Update the **Status** column (e.g., change from "Sent" to "Paid") to update your Dashboard numbers.

## 📂 Project Structure

*   `Config.js`: Central configuration for sheet names, state codes, and helper functions (like `NUMBER_TO_WORDS`).
*   `Setup.js`: Logic for building the sheet structure, formatting, and applying protections.
*   `InvoiceService.js`: Core logic for scraping invoice data, calculating totals, and managing the invoice lifecycle.
*   `EmailService.js`: Handles PDF generation (HTML to PDF) and sending emails via Gmail.
*   `Triggers.js`: (If applicable) Manages automation triggers like `onEdit` for dynamic updates.

## ⚠️ Notes
*   **Do not rename sheets:** The script relies on specific sheet names (defined in `Config.js`) to function correctly.
*   **Drive Folder:** A folder named `Invoices_Sent` will be automatically created in your Google Drive to store PDF copies of all invoices.

## 🤝 Contributing
Feel free to fork this project and customize the `src/` scripts to fit your specific business needs!
