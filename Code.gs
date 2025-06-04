// Global variables - USER MUST FILL THESE OUT
var SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE"; // Please replace with your actual Spreadsheet ID
var COUPON_SHEET_NAME = "Coupons";
var USAGE_LOG_SHEET_NAME = "UsageLog";

/**
 * Serves the HTML interface for the web app.
 * @param {Object} e Event parameter.
 * @return {HtmlOutput} HTML output for the web app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index.html')
      .setTitle('Coupon Management System')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Saves a new coupon to the spreadsheet.
 * @param {Object} couponData An object containing coupon details:
 *                            {barcode: string, expiryDate: string, isGiftCertificate: boolean, initialBalance: number | null}
 * @return {String} A success or error message.
 */
function saveCoupon(couponData) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    if (!sheet) {
      // If sheet doesn't exist, create it with headers
      sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(COUPON_SHEET_NAME);
      sheet.appendRow(["Barcode", "Entry Date", "Expiration Date", "Is Gift Certificate", "Balance", "Original Amount"]);
    }

    var entryDate = new Date();
    var balance = couponData.isGiftCertificate ? parseFloat(couponData.initialBalance) : null;
    var originalAmount = balance; // Store the original amount for reference

    // Validate data
    if (!couponData.barcode || couponData.barcode.trim() === "") {
      return "Error: Barcode cannot be empty.";
    }
    if (!couponData.expiryDate) {
      return "Error: Expiration date cannot be empty.";
    }
    if (couponData.isGiftCertificate && (isNaN(balance) || balance === null || balance < 0)) {
        return "Error: Initial balance for gift certificate must be a non-negative number.";
    }


    sheet.appendRow([
      couponData.barcode,
      entryDate,
      new Date(couponData.expiryDate),
      couponData.isGiftCertificate,
      balance,
      originalAmount // Add original amount to the row
    ]);
    return "Coupon saved successfully: " + couponData.barcode;
  } catch (e) {
    Logger.log("Error in saveCoupon: " + e.toString());
    return "Error saving coupon: " + e.toString();
  }
}

/**
 * Retrieves all coupons from the spreadsheet.
 * Sorts by Entry Date (newest first).
 * @return {Array<Object>} An array of coupon objects.
 */
function getCoupons() {
  var returnValue = [];
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
      Logger.log("Error in getCoupons: Spreadsheet not found with ID: " + SPREADSHEET_ID + ". Returning empty array string.");
    } else {
      var sheet = ss.getSheetByName(COUPON_SHEET_NAME);
      if (!sheet) {
        Logger.log("Info in getCoupons: Coupon sheet '" + COUPON_SHEET_NAME + "' not found. Returning empty array string.");
      } else {
        if (sheet.getLastRow() === 0) {
          Logger.log("Info in getCoupons: Coupon sheet is completely empty. Returning empty array string.");
        } else {
          var dataRange = sheet.getDataRange();
          if (!dataRange) {
            Logger.log("Info in getCoupons: No data range found in sheet. Returning empty array string.");
          } else {
            var data = dataRange.getValues();
            if (!data || sheet.getLastRow() <= 1) {
              Logger.log("Info in getCoupons: No actual coupon data found. Last row: " + sheet.getLastRow() + ". Data length: " + (data ? data.length : "null") + ". Returning empty array string.");
            } else {
              var couponsData = [];
              for (var i = 1; i < data.length; i++) {
                if (data[i] && data[i][0] != "") {
                  couponsData.push({
                    barcode: data[i][0],
                    entryDate: data[i][1],
                    expiryDate: data[i][2],
                    isGiftCertificate: data[i][3],
                    balance: data[i][4],
                    originalAmount: data[i][5]
                  });
                }
              }
              couponsData.sort(function(a, b) {
                var dateA = (a.entryDate instanceof Date) ? a.entryDate : new Date(a.entryDate);
                var dateB = (b.entryDate instanceof Date) ? b.entryDate : new Date(b.entryDate);
                return dateB - dateA;
              });
              returnValue = couponsData;
            }
          }
        }
      }
    }
    Logger.log("Returning from getCoupons (try), value before stringify: " + JSON.stringify(returnValue));
    return JSON.stringify(returnValue); // Stringify here
  } catch (e) {
    Logger.log("Exception in getCoupons: " + e.toString() + " Stack: " + e.stack);
    var fallbackValue = [];
    Logger.log("Returning from getCoupons (catch), value before stringify: " + JSON.stringify(fallbackValue));
    return JSON.stringify(fallbackValue); // Stringify here
  }
}

/**
 * Retrieves the 5 most recently added coupons.
 * @return {String} A JSON string representing an array of the 5 latest coupon objects.
 */
function getLatestCoupons() {
  try {
    var couponsJsonString = getCoupons(); // This is now a JSON string
    var allCoupons = JSON.parse(couponsJsonString); // Parse it

    // Ensure allCoupons is an array after parsing, though getCoupons should guarantee it
    if (!Array.isArray(allCoupons)) {
        Logger.log("Error in getLatestCoupons: Parsed data from getCoupons is not an array. Data: " + couponsJsonString);
        return JSON.stringify([]); // Return stringified empty array
    }

    var latest = allCoupons.slice(0, 5);
    Logger.log("Returning from getLatestCoupons, value before stringify: " + JSON.stringify(latest));
    return JSON.stringify(latest); // Re-stringify
  } catch (e) {
    Logger.log("Exception in getLatestCoupons: " + e.toString() + " Stack: " + e.stack);
    return JSON.stringify([]); // Return stringified empty array on error
  }
}

/**
 * Logs usage of a gift certificate and updates its balance.
 * @param {String} barcode The barcode of the gift certificate.
 * @param {Number} amountUsed The amount used.
 * @return {String} A JSON string representing an object with success status, message, and newBalance or error.
 */
function logGiftCertificateUsage(barcode, amountUsed) {
  try {
    if (typeof amountUsed !== 'number' || amountUsed <= 0) {
      return JSON.stringify({ success: false, message: "Error: Amount used must be a positive number." });
    }

    var couponSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    if (!couponSheet) {
      return JSON.stringify({ success: false, message: "Error: Coupons sheet not found." });
    }

    var data = couponSheet.getDataRange().getValues();
    var couponRow = -1;
    var currentBalance = 0;
    var isGift = false;

    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == barcode) {
        couponRow = i;
        isGift = data[i][3];
        currentBalance = parseFloat(data[i][4]);
        break;
      }
    }

    if (couponRow === -1) {
      return JSON.stringify({ success: false, message: "Error: Coupon with barcode '" + barcode + "' not found." });
    }
    if (!isGift) {
      return JSON.stringify({ success: false, message: "Error: Coupon '" + barcode + "' is not a gift certificate." });
    }
    if (isNaN(currentBalance) || currentBalance < amountUsed) {
      return JSON.stringify({ success: false, message: "Error: Insufficient balance. Current balance: " + (isNaN(currentBalance) ? 0 : currentBalance.toFixed(2)) });
    }

    var newBalance = currentBalance - amountUsed;
    couponSheet.getRange(couponRow + 1, 5).setValue(newBalance);

    var logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USAGE_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(USAGE_LOG_SHEET_NAME);
      logSheet.appendRow(["Timestamp", "Barcode", "Amount Used", "New Balance"]);
    }
    logSheet.appendRow([new Date(), barcode, amountUsed, newBalance]);

    return JSON.stringify({ success: true, message: "Usage logged successfully. New balance for " + barcode + ": " + newBalance.toFixed(2), newBalance: newBalance });
  } catch (e) {
    Logger.log("Error in logGiftCertificateUsage: " + e.toString());
    return JSON.stringify({ success: false, message: "Error logging usage: " + e.toString() });
  }
}

// Helper function to test (optional, can be run from Apps Script editor)
function testSaveCoupon() {
  // Note: Test functions will now log JSON strings for return values from saveCoupon etc.
  Logger.log(saveCoupon({barcode: "TEST12345", expiryDate: "2024-12-31", isGiftCertificate: false, initialBalance: null}));
  Logger.log(saveCoupon({barcode: "GIFT001", expiryDate: "2025-06-30", isGiftCertificate: true, initialBalance: 100}));
}

function testGetCoupons() {
  var couponsJson = getCoupons();
  Logger.log("Raw JSON from getCoupons():");
  Logger.log(couponsJson);
  Logger.log("Parsed coupons from getCoupons():");
  Logger.log(JSON.parse(couponsJson));
}

function testGetLatestCoupons() {
  var latestCouponsJson = getLatestCoupons();
  Logger.log("Raw JSON from getLatestCoupons():");
  Logger.log(latestCouponsJson);
  Logger.log("Parsed coupons from getLatestCoupons():");
  Logger.log(JSON.parse(latestCouponsJson));
}

function testLogUsage() {
  // Make sure a coupon with barcode "GIFT001" exists and is a gift certificate with balance
  // Logger.log(saveCoupon({barcode: "GIFT001", expiryDate: "2025-06-30", isGiftCertificate: true, initialBalance: 100})); // if not already present
  Logger.log(logGiftCertificateUsage("GIFT001", 25));
  // Logger.log(logGiftCertificateUsage("GIFT001", 100)); // Test insufficient
  // Logger.log(logGiftCertificateUsage("NONEXISTENT", 10));
}

// --- saveCoupon modification starts here ---
// (The prompt had saveCoupon code mixed with test functions, separating it)
/**
 * Saves a new coupon to the spreadsheet.
 * @param {Object} couponData An object containing coupon details:
 *                            {barcode: string, expiryDate: string, isGiftCertificate: boolean, initialBalance: number | null}
 * @return {String} A JSON string representing a success or error object.
 */
function saveCoupon(couponData) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    if (!sheet) {
      // If sheet doesn't exist, create it with headers
      sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(COUPON_SHEET_NAME);
      sheet.appendRow(["Barcode", "Entry Date", "Expiration Date", "Is Gift Certificate", "Balance", "Original Amount"]);
    }

    var entryDate = new Date();
    var balance = couponData.isGiftCertificate ? parseFloat(couponData.initialBalance) : null;
    var originalAmount = balance; // Store the original amount for reference

    // Validate data
    if (!couponData.barcode || couponData.barcode.trim() === "") {
      return JSON.stringify({ success: false, message: "Error: Barcode cannot be empty." });
    }
    if (!couponData.expiryDate) {
      return JSON.stringify({ success: false, message: "Error: Expiration date cannot be empty." });
    }
    // Ensure expiryDate is treated as a date, even if it comes as a string
    var expiryDateObj = new Date(couponData.expiryDate);
    if (isNaN(expiryDateObj.getTime())) {
        return JSON.stringify({ success: false, message: "Error: Invalid expiration date format."});
    }

    if (couponData.isGiftCertificate && (isNaN(balance) || balance === null || balance < 0)) {
        return JSON.stringify({ success: false, message: "Error: Initial balance for gift certificate must be a non-negative number."});
    }


    sheet.appendRow([
      couponData.barcode,
      entryDate,
      expiryDateObj, // Use the date object
      couponData.isGiftCertificate,
      balance,
      originalAmount // Add original amount to the row
    ]);
    return JSON.stringify({ success: true, message: "Coupon saved successfully: " + couponData.barcode });
  } catch (e) {
    Logger.log("Error in saveCoupon: " + e.toString());
    return JSON.stringify({ success: false, message: "Error saving coupon: " + e.toString() });
  }
}
// --- saveCoupon modification ends here ---
