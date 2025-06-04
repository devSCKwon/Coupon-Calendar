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
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    if (!sheet) {
      return []; // Or throw an error: throw new Error("Coupon sheet not found");
    }
    // Get all data, excluding the header row
    var data = sheet.getDataRange().getValues();
    if (data.length < 2) {
      return []; // No data other than header
    }

    var coupons = [];
    // Start from 1 to skip header row
    for (var i = 1; i < data.length; i++) {
      coupons.push({
        barcode: data[i][0],
        entryDate: data[i][1],
        expiryDate: data[i][2],
        isGiftCertificate: data[i][3],
        balance: data[i][4],
        originalAmount: data[i][5]
      });
    }

    // Sort by Entry Date (column 1, which is index 1 in data[i]), newest first
    coupons.sort(function(a, b) {
      return new Date(b.entryDate) - new Date(a.entryDate);
    });

    return coupons;
  } catch (e) {
    Logger.log("Error in getCoupons: " + e.toString());
    return []; // Return empty array on error
  }
}

/**
 * Retrieves the 5 most recently added coupons.
 * @return {Array<Object>} An array of the 5 latest coupon objects.
 */
function getLatestCoupons() {
  var allCoupons = getCoupons(); // Leverages the sorting in getCoupons
  return allCoupons.slice(0, 5);
}

/**
 * Logs usage of a gift certificate and updates its balance.
 * @param {String} barcode The barcode of the gift certificate.
 * @param {Number} amountUsed The amount used.
 * @return {Object} An object with success status, message, and newBalance or error.
 */
function logGiftCertificateUsage(barcode, amountUsed) {
  try {
    if (typeof amountUsed !== 'number' || amountUsed <= 0) {
      return { success: false, message: "Error: Amount used must be a positive number." };
    }

    var couponSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    if (!couponSheet) {
      return { success: false, message: "Error: Coupons sheet not found." };
    }

    var data = couponSheet.getDataRange().getValues();
    var couponRow = -1;
    var currentBalance = 0;
    var isGift = false;

    // Find the coupon (skip header)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == barcode) { // Barcode in column A
        couponRow = i;
        isGift = data[i][3]; // isGiftCertificate in column D
        currentBalance = parseFloat(data[i][4]); // Balance in column E
        break;
      }
    }

    if (couponRow === -1) {
      return { success: false, message: "Error: Coupon with barcode '" + barcode + "' not found." };
    }

    if (!isGift) {
      return { success: false, message: "Error: Coupon '" + barcode + "' is not a gift certificate." };
    }

    if (isNaN(currentBalance) || currentBalance < amountUsed) {
      return { success: false, message: "Error: Insufficient balance. Current balance: " + (isNaN(currentBalance) ? 0 : currentBalance) };
    }

    var newBalance = currentBalance - amountUsed;
    couponSheet.getRange(couponRow + 1, 5).setValue(newBalance); // Update balance (column E)

    // Log the usage
    var logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USAGE_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(USAGE_LOG_SHEET_NAME);
      logSheet.appendRow(["Timestamp", "Barcode", "Amount Used", "New Balance"]);
    }
    logSheet.appendRow([new Date(), barcode, amountUsed, newBalance]);

    return { success: true, message: "Usage logged successfully. New balance for " + barcode + ": " + newBalance, newBalance: newBalance };
  } catch (e) {
    Logger.log("Error in logGiftCertificateUsage: " + e.toString());
    return { success: false, message: "Error logging usage: " + e.toString() };
  }
}

// Helper function to test (optional, can be run from Apps Script editor)
function testSaveCoupon() {
  Logger.log(saveCoupon({barcode: "TEST12345", expiryDate: "2024-12-31", isGiftCertificate: false, initialBalance: null}));
  Logger.log(saveCoupon({barcode: "GIFT001", expiryDate: "2025-06-30", isGiftCertificate: true, initialBalance: 100}));
}

function testGetCoupons() {
  Logger.log(getCoupons());
}

function testGetLatestCoupons() {
  Logger.log(getLatestCoupons());
}

function testLogUsage() {
  // Make sure a coupon with barcode "GIFT001" exists and is a gift certificate with balance
  // saveCoupon({barcode: "GIFT001", expiryDate: "2025-06-30", isGiftCertificate: true, initialBalance: 100}); // if not already present
  Logger.log(logGiftCertificateUsage("GIFT001", 25));
  Logger.log(logGiftCertificateUsage("GIFT001", 100)); // Test insufficient
  Logger.log(logGiftCertificateUsage("NONEXISTENT", 10));
}
