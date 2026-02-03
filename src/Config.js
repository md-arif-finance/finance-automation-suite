// --- CONFIGURATION ---
var SHEET_NAME_DASHBOARD = "🏠 Dashboard";
var SHEET_NAME_UI        = "🧾 Create Invoice";
var SHEET_NAME_TRACKER   = "📊 Invoice History";
var SHEET_NAME_CLIENTS   = "👥 Clients";
var SHEET_NAME_PRODUCTS  = "📦 Products";
var SHEET_NAME_ITEMS     = "🗄️ Items Archive";
var SHEET_NAME_LOGS      = "⚠️ Logs";
var FOLDER_NAME          = "Invoices_Sent";

/**
 * HELPER: Reads a setting from the Dashboard Sheet
 * We will define fixed cells for these settings to make it faster.
 * Company Name: C5, State: C6, Email Subject: C7
 */
function fetchSetting(key) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME_DASHBOARD);
  if (!sheet) return null;

  if (key === "Company Name") return sheet.getRange("D11").getValue();
  if (key === "Company Address") return sheet.getRange("D12").getValue();
  if (key === "My State") return sheet.getRange("D14").getValue();
  
  return null;
}

/**
 * MASTER DATA: List of Indian States with GST Codes
 */
var RAW_STATES = [
  { code: 1, name: "Jammu and Kashmir" },
  { code: 2, name: "Himachal Pradesh" },
  { code: 3, name: "Punjab" },
  { code: 4, name: "Chandigarh" },
  { code: 5, name: "Uttarakhand" },
  { code: 6, name: "Haryana" },
  { code: 7, name: "Delhi" },
  { code: 8, name: "Rajasthan" },
  { code: 9, name: "Uttar Pradesh" },
  { code: 10, name: "Bihar" },
  { code: 11, name: "Sikkim" },
  { code: 12, name: "Arunachal Pradesh" },
  { code: 13, name: "Nagaland" },
  { code: 14, name: "Manipur" },
  { code: 15, name: "Mizoram" },
  { code: 16, name: "Tripura" },
  { code: 17, name: "Meghalaya" },
  { code: 18, name: "Assam" },
  { code: 19, name: "West Bengal" },
  { code: 20, name: "Jharkhand" },
  { code: 21, name: "Odisha" },
  { code: 22, name: "Chhattisgarh" },
  { code: 23, name: "Madhya Pradesh" },
  { code: 24, name: "Gujarat" },
  { code: 26, name: "Dadra and Nagar Haveli and Daman and Diu" },
  { code: 27, name: "Maharashtra" },
  { code: 29, name: "Karnataka" },
  { code: 30, name: "Goa" },
  { code: 31, name: "Lakshadweep" },
  { code: 32, name: "Kerala" },
  { code: 33, name: "Tamil Nadu" },
  { code: 34, name: "Puducherry" },
  { code: 35, name: "Andaman and Nicobar Islands" },
  { code: 36, name: "Telangana" },
  { code: 37, name: "Andhra Pradesh" },
  { code: 38, name: "Ladakh" },
  { code: 97, name: "Other Territory" },
  { code: 99, name: "Centre Jurisdiction" }
];

/**
 * HELPER: Returns sorted list of "State Name (Code)" for dropdowns
 */
function getIndianStateDropdownList() {
  var formatted = RAW_STATES.map(function(s) {
    // Add leading zero if code is single digit (e.g., 7 -> "07")
    var paddedCode = s.code < 10 ? "0" + s.code : "" + s.code;
    return s.name + " (" + paddedCode + ")";
  });
  
  // Sort Alphabetically
  return formatted.sort(); 
}

// ... (Previous Config code) ...

/**
 * HELPER: Converts Number to Indian Currency Words
 * Used in the "Amount in Words" cell
 */
function NUMBER_TO_WORDS(amount) {
  if (amount === 0) return "Zero Rupees Only";
  
  var words = new Array();
  words[0] = '';
  words[1] = 'One';
  words[2] = 'Two';
  words[3] = 'Three';
  words[4] = 'Four';
  words[5] = 'Five';
  words[6] = 'Six';
  words[7] = 'Seven';
  words[8] = 'Eight';
  words[9] = 'Nine';
  words[10] = 'Ten';
  words[11] = 'Eleven';
  words[12] = 'Twelve';
  words[13] = 'Thirteen';
  words[14] = 'Fourteen';
  words[15] = 'Fifteen';
  words[16] = 'Sixteen';
  words[17] = 'Seventeen';
  words[18] = 'Eighteen';
  words[19] = 'Nineteen';
  words[20] = 'Twenty';
  words[30] = 'Thirty';
  words[40] = 'Forty';
  words[50] = 'Fifty';
  words[60] = 'Sixty';
  words[70] = 'Seventy';
  words[80] = 'Eighty';
  words[90] = 'Ninety';
  
  amount = amount.toString();
  var atemp = amount.split(".");
  var number = atemp[0].split(",").join("");
  var n_length = number.length;
  var words_string = "";
  
  if (n_length <= 9) {
    var n_array = new Array(0, 0, 0, 0, 0, 0, 0, 0, 0);
    var received_n_array = new Array();
    for (var i = 0; i < n_length; i++) {
      received_n_array[i] = number.substr(i, 1);
    }
    for (var i = 9 - n_length, j = 0; i < 9; i++, j++) {
      n_array[i] = received_n_array[j];
    }
    for (var i = 0, j = 1; i < 9; i++, j++) {
      if (i == 0 || i == 2 || i == 4 || i == 7) {
        if (n_array[i] == 1) {
          n_array[j] = 10 + parseInt(n_array[j]);
          n_array[i] = 0;
        }
      }
    }
    var value = "";
    for (var i = 0; i < 9; i++) {
      if (i == 0 || i == 2 || i == 4 || i == 7) {
        value = n_array[i] * 10;
      } else {
        value = n_array[i];
      }
      if (value != 0) {
        words_string += words[value] + " ";
      }
      if ((i == 1 && value != 0) || (i == 0 && value != 0 && n_array[i + 1] == 0)) {
        words_string += "Crore ";
      }
      if ((i == 3 && value != 0) || (i == 2 && value != 0 && n_array[i + 1] == 0)) {
        words_string += "Lakh ";
      }
      if ((i == 5 && value != 0) || (i == 4 && value != 0 && n_array[i + 1] == 0)) {
        words_string += "Thousand ";
      }
      if (i == 6 && value != 0 && (n_array[i + 1] != 0 && n_array[i + 2] != 0)) {
        words_string += "Hundred and ";
      } else if (i == 6 && value != 0) {
        words_string += "Hundred ";
      }
    }
    words_string = words_string.split("  ").join(" ");
  }
  return words_string + "Rupees Only";
}