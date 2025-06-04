// Global Variables
var SPREADSHEET_ID = "YOUR_SPREADSHEET_ID_HERE";
var DRIVE_FOLDER_ID = "YOUR_DRIVE_FOLDER_ID_HERE";

var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
var couponsSheet = ss.getSheetByName("Coupons");
var usageLogSheet = ss.getSheetByName("UsageLog");

/**
 * Serves the HTML for the web app.
 * @param {Object} e event parameter.
 * @return {HtmlOutput} HTML output for the web app.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index.html')
      .setTitle('Coupon Manager Web App')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * Saves coupon data to the "Coupons" sheet.
 *
 * @param {string} barcode The barcode string.
 * @param {string} inputDate The date the coupon was input (YYYY-MM-DD string).
 * @param {string} expiryDate The expiry date of the coupon (YYYY-MM-DD string).
 * @param {boolean} isMonetaryValue True if the coupon has monetary value.
 * @param {number} initialValue The initial monetary value (can be 0 if not monetary).
 * @param {string} imageFileName The filename of the uploaded barcode image.
 * @param {string} notes Additional notes for the coupon.
 * @return {string} A success or error message.
 */
function saveCouponData(barcode, inputDate, expiryDate, isMonetaryValue, initialValue, imageFileName, notes) {
  try {
    if (!couponsSheet) {
      throw new Error("Coupons sheet not found!");
    }
    if (!barcode) {
      throw new Error("Barcode cannot be empty.");
    }
    if (!inputDate) {
      throw new Error("Input date cannot be empty.");
    }
     if (!expiryDate) {
      throw new Error("Expiry date cannot be empty.");
    }

    var currentBalance = isMonetaryValue ? initialValue : 0;

    couponsSheet.appendRow([
      barcode,
      new Date(inputDate),
      new Date(expiryDate),
      isMonetaryValue,
      initialValue || 0,
      currentBalance,
      imageFileName || "",
      notes || ""
    ]);

    Logger.log("Coupon data saved: " + barcode);
    return "Coupon '" + barcode + "' saved successfully.";

  } catch (error) {
    Logger.log("Error in saveCouponData: " + error.message);
    Logger.log("Stack trace: " + error.stack);
    return "Error saving coupon: " + error.message;
  }
}

// Helper function to test (Optional - you can run this from Apps Script Editor)
function testSaveCoupon() {
  Logger.log(saveCouponData(
    "TEST12345",
    "2024-01-15",
    "2024-12-31",
    true,
    50.00,
    "test_image.jpg",
    "This is a test coupon."
  ));
  Logger.log(saveCouponData(
    "PROMOABC",
    "2024-03-01",
    "2024-06-30",
    false,
    0,
    "",
    "Promotional item."
  ));
   Logger.log(saveCouponData(
    "NOEXPIRY",
    "2023-01-01",
    "9999-12-31", // A very distant future date for no expiry
    false,
    0,
    "no_expiry_img.png",
    "Coupon with no practical expiry."
  ));
   Logger.log(saveCouponData(
    "GIFTCARD789",
    "2024-05-20",
    "2025-05-20",
    true,
    100.00,
    "giftcard.png",
    "Birthday gift card."
  ));
}

```
/**
 * Uploads a file (image) to the specified Google Drive folder.
 * The image data is expected to be a base64 encoded string.
 *
 * @param {string} base64Data The base64 encoded data of the image.
 * @param {string} fileName The desired name for the file in Google Drive.
 * @return {object} An object containing the fileId and fileName of the uploaded file, or an error message.
 */
function uploadBarcodeImage(base64Data, fileName) {
  try {
    if (!DRIVE_FOLDER_ID) {
      throw new Error("DRIVE_FOLDER_ID is not set in Code.gs.");
    }
    var folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);

    // Decode the base64 string
    // The data URL format is "data:[<mediatype>][;base64],<data>"
    // We need to extract the mediatype and the base64 part.
    var contentType = base64Data.substring(base64Data.indexOf(':') + 1, base64Data.indexOf(';'));
    var GDriveSafeFileName = makeFileNameGDriveSafe(fileName);
    var data = base64Data.substring(base64Data.indexOf(',') + 1);
    var decodedData = Utilities.base64Decode(data);
    var blob = Utilities.newBlob(decodedData, contentType, GDriveSafeFileName);

    var file = folder.createFile(blob);

    Logger.log("File uploaded: " + file.getName() + ", ID: " + file.getId());
    return {
      fileId: file.getId(),
      fileName: file.getName()
    };

  } catch (error) {
    Logger.log("Error in uploadBarcodeImage: " + error.message);
    Logger.log("Stack trace: " + error.stack);
    // It's often better to throw the error or return a structured error object
    // for the client-side to handle.
    return {
      error: "Error uploading image: " + error.message
    };
  }
}

/**
 * Makes a filename safe for Google Drive by removing or replacing characters
 * that might be problematic. Google Drive is generally quite permissive,
 * but this can help avoid edge cases.
 * This version is basic; more complex sanitization might be needed for specific cases.
 * @param {string} fileName The original filename.
 * @return {string} A "safer" version of the filename.
 */
function makeFileNameGDriveSafe(fileName) {
  // Replace characters that are problematic in some filesystems or URLs.
  // Google Drive is flexible, but it's good practice.
  // This replaces slashes, backslashes, colons, etc., with underscores.
  // It also removes leading/trailing whitespace.
  var newName = fileName.trim().replace(/[\/:*?"<>|]/g, '_');

  // Truncate if too long (Google Drive has a limit, though it's quite large)
  if (newName.length > 200) {
    var extension = "";
    var dotIndex = newName.lastIndexOf('.');
    if (dotIndex > 0 && newName.length - dotIndex <= 5) { // simple extension check
        extension = newName.substring(dotIndex);
        newName = newName.substring(0, 200 - extension.length);
    } else {
        newName = newName.substring(0, 200);
    }
    newName += extension;
  }
  return newName;
}


/**
 * Retrieves all coupons from the "Coupons" sheet.
 *
 * @return {Array<Object>} An array of coupon objects, or an empty array if error/no data.
 */
function getCoupons() {
  try {
    if (!couponsSheet) {
      throw new Error("Coupons sheet not found!");
    }
    // Get all data, excluding the header row
    // getDataRange() gets the entire block of data.
    // getValues() returns a 2D array.
    var data = couponsSheet.getDataRange().getValues();

    if (data.length <= 1) { // Only header or empty
      return [];
    }

    var headers = data[0];
    var coupons = [];

    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      var coupon = {};
      var isEmptyRow = true; // Flag to check if the row is entirely empty
      for (var j = 0; j < headers.length; j++) {
        var header = headers[j];
        var value = row[j];
        // Check if this cell has content, to avoid processing empty rows
        if (value !== "") {
            isEmptyRow = false;
        }
        // Format dates to ISO string (YYYY-MM-DD) for easier client-side handling
        if ((header === 'InputDate' || header === 'ExpiryDate') && value instanceof Date) {
          coupon[header] = Utilities.formatDate(value, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          coupon[header] = value;
        }
      }
       if (!isEmptyRow) { // Only add coupon if it's not an empty row
           coupons.push(coupon);
       }
    }
    // Logger.log("Retrieved " + coupons.length + " coupons.");
    return coupons;

  } catch (error) {
    Logger.log("Error in getCoupons: " + error.message);
    return []; // Return empty array on error
  }
}

/**
 * Retrieves the 5 most recently added coupons.
 * Assumes coupons are added chronologically and InputDate is a good proxy,
 * or simply takes the last 5 rows. For more robust "recent", sorting by an ID or timestamp is better.
 * This implementation takes the last 5 non-empty rows.
 *
 * @return {Array<Object>} An array of the 5 most recent coupon objects.
 */
function getRecentCoupons() {
  try {
    var allCoupons = getCoupons(); // This already handles empty/header checks

    // Sort by InputDate descending to be sure, if not already guaranteed by sheet order
    // Requires InputDate to be consistently a Date object or string that sorts correctly.
    // For simplicity, we'll rely on sheet order (last rows are newest) or getCoupons providing them ordered.
    // If getCoupons doesn't guarantee order, uncomment and adapt sorting:
    /*
    allCoupons.sort(function(a,b){
      // Assuming InputDate is in a format that can be compared, like YYYY-MM-DD
      // or convert to Date objects for comparison: new Date(b.InputDate) - new Date(a.InputDate)
      if (a.InputDate && b.InputDate) {
        return new Date(b.InputDate) - new Date(a.InputDate);
      }
      return 0; // no change in order if dates are missing
    });
    */

    // Get the last 5 items
    var recentCoupons = allCoupons.slice(-5);
    // Logger.log("Retrieved " + recentCoupons.length + " recent coupons.");
    return recentCoupons.reverse(); // To show newest first if taking from end of array

  } catch (error) {
    Logger.log("Error in getRecentCoupons: " + error.message);
    return [];
  }
}

// Helper function to test (Optional)
function testGetCoupons() {
  var coupons = getCoupons();
  if (coupons.length > 0) {
    Logger.log("First coupon: " + JSON.stringify(coupons[0]));
    Logger.log("Number of coupons: " + coupons.length);
  } else {
    Logger.log("No coupons found or error occurred.");
  }

  var recent = getRecentCoupons();
  if (recent.length > 0) {
    Logger.log("Most recent coupon: " + JSON.stringify(recent[0]));
    Logger.log("Number of recent coupons: " + recent.length);
  } else {
    Logger.log("No recent coupons found.");
  }
}

function testUpload() {
  // This test function cannot be run directly without providing actual base64 data and a filename.
  // You would typically call uploadBarcodeImage from the client-side (HTML JavaScript)
  // with data from a file input.
  // For a direct backend test, you'd need a sample base64 string.
  Logger.log("To test uploadBarcodeImage, call it from the client-side HTML or provide a sample base64 string here.");
  // Example (replace with actual base64 data and filename):
  // var sampleBase64 = "data:image/png;base64,iVBORw0KGgoAAAANSUhEUgAAAAUA..."; // very short, not a real image
  // var sampleFileName = "test_image_from_backend.png";
  // var result = uploadBarcodeImage(sampleBase64, sampleFileName);
  // Logger.log(result);
}

/**
 * Retrieves coupons that are expiring within a specified number of days from today.
 *
 * @param {number} daysThreshold The number of days to look ahead for expiring coupons. E.g., 7 for next 7 days.
 * @return {Array<Object>} An array of coupon objects that are nearing expiry.
 */
function getExpiringCoupons(daysThreshold) {
  try {
    var allCoupons = getCoupons(); // Re-use the existing function
    var expiringSoon = [];
    var today = new Date();
    today.setHours(0, 0, 0, 0); // Normalize today to the start of the day for accurate date comparison

    var thresholdDate = new Date(today);
    thresholdDate.setDate(today.getDate() + parseInt(daysThreshold));

    for (var i = 0; i < allCoupons.length; i++) {
      var coupon = allCoupons[i];
      if (coupon.ExpiryDate) {
        var expiryDate = new Date(coupon.ExpiryDate);
        // Normalize expiryDate as well if it includes time, though YYYY-MM-DD from getCoupons should be fine
        expiryDate.setHours(0,0,0,0);

        // Check if expiryDate is after or equal to today AND before or equal to the thresholdDate
        if (expiryDate >= today && expiryDate <= thresholdDate) {
          expiringSoon.push(coupon);
        }
      }
    }

    // Sort by expiry date, soonest first
    expiringSoon.sort(function(a,b){
      return new Date(a.ExpiryDate) - new Date(b.ExpiryDate);
    });

    // Logger.log("Found " + expiringSoon.length + " coupons expiring within " + daysThreshold + " days.");
    return expiringSoon;

  } catch (error) {
    Logger.log("Error in getExpiringCoupons: " + error.message);
    return [];
  }
}

/**
 * Filters coupons based on a date range for either InputDate or ExpiryDate.
 *
 * @param {string} startDateString The start date of the range (YYYY-MM-DD string).
 * @param {string} endDateString The end date of the range (YYYY-MM-DD string).
 * @param {string} dateType The type of date to filter on: "InputDate" or "ExpiryDate".
 * @return {Array<Object>} An array of coupon objects that match the filter.
 */
function filterCouponsByDateRange(startDateString, endDateString, dateType) {
  try {
    var allCoupons = getCoupons();
    var filteredCoupons = [];

    if (!startDateString || !endDateString || !dateType) {
        Logger.log("filterCouponsByDateRange: Missing one or more parameters.");
        return []; // Or throw error
    }

    var startDate = new Date(startDateString);
    startDate.setHours(0,0,0,0); // Normalize
    var endDate = new Date(endDateString);
    endDate.setHours(23,59,59,999); // Normalize to end of day for inclusive range

    for (var i = 0; i < allCoupons.length; i++) {
      var coupon = allCoupons[i];
      var couponDateValue = coupon[dateType]; // e.g., coupon['InputDate']

      if (couponDateValue) {
        var couponDate = new Date(couponDateValue);
         // Normalize couponDate if it's just a date string without time
        couponDate.setHours(0,0,0,0);


        if (couponDate >= startDate && couponDate <= endDate) {
          filteredCoupons.push(coupon);
        }
      }
    }
    // Logger.log("Filtered " + filteredCoupons.length + " coupons by " + dateType + " between " + startDateString + " and " + endDateString);
    return filteredCoupons;

  } catch (error) {
    Logger.log("Error in filterCouponsByDateRange: " + error.message);
    return [];
  }
}

/**
 * Logs the usage of a monetary coupon and updates its balance.
 *
 * @param {string} barcode The barcode of the coupon to use.
 * @param {number} amountUsed The amount to deduct from the coupon's balance.
 * @return {object} An object with a status message and the updated balance, or an error message.
 */
function logCouponUsage(barcode, amountUsed) {
  try {
    if (!couponsSheet || !usageLogSheet) {
      throw new Error("One or more sheets (Coupons, UsageLog) not found.");
    }
    if (!barcode || amountUsed == null || isNaN(parseFloat(amountUsed)) || parseFloat(amountUsed) <= 0) {
      throw new Error("Barcode or valid positive amountUsed not provided.");
    }

    var amount = parseFloat(amountUsed);
    var data = couponsSheet.getDataRange().getValues();
    var headers = data[0];
    var barcodeCol = headers.indexOf("Barcode") + 1; // 1-based index
    var isMonetaryCol = headers.indexOf("IsMonetaryValue") + 1;
    var balanceCol = headers.indexOf("CurrentBalance") + 1;
    var initialValueCol = headers.indexOf("InitialValue") + 1;


    if (barcodeCol === 0 || isMonetaryCol === 0 || balanceCol === 0 || initialValueCol === 0) {
        throw new Error("Could not find required columns (Barcode, IsMonetaryValue, CurrentBalance, InitialValue) in Coupons sheet.");
    }


    var couponFound = false;
    var updatedBalance = 0;

    for (var i = 1; i < data.length; i++) { // Start from 1 to skip header
      if (data[i][barcodeCol - 1] == barcode) {
        couponFound = true;
        if (!data[i][isMonetaryCol - 1]) { // Check IsMonetaryValue
          return { error: "Coupon '" + barcode + "' is not a monetary value coupon." };
        }

        var currentBalance = parseFloat(data[i][balanceCol - 1]);
        if (currentBalance < amount) {
          return { error: "Insufficient balance for coupon '" + barcode + "'. Available: " + currentBalance };
        }

        updatedBalance = currentBalance - amount;
        // Update the balance in the "Coupons" sheet (Row i+1 because sheet rows are 1-based, data array is 0-based for rows after header)
        couponsSheet.getRange(i + 1, balanceCol).setValue(updatedBalance);

        // Log the usage
        var logId = Utilities.getUuid(); // Generate a unique ID for the log
        var usageDate = new Date();
        usageLogSheet.appendRow([logId, barcode, usageDate, amount, updatedBalance]);

        Logger.log("Usage logged for coupon: " + barcode + ". Amount used: " + amount + ". New balance: " + updatedBalance);
        return {
          success: "Usage logged successfully for '" + barcode + "'.",
          newBalance: updatedBalance,
          barcode: barcode
        };
      }
    }

    if (!couponFound) {
      return { error: "Coupon with barcode '" + barcode + "' not found." };
    }

  } catch (error) {
    Logger.log("Error in logCouponUsage: " + error.message);
    Logger.log("Stack trace: " + error.stack);
    return { error: "Error logging usage: " + error.message };
  }
}

// Helper function to test (Optional)
function testExpiringAndFilter() {
  Logger.log("--- Testing getExpiringCoupons ---");
  var expiring = getExpiringCoupons(7); // Expiring in next 7 days
  if (expiring.length > 0) {
    Logger.log("Expiring coupons (next 7 days): " + expiring.length);
    expiring.forEach(function(c) { Logger.log(JSON.stringify(c)); });
  } else {
    Logger.log("No coupons expiring in the next 7 days.");
  }

  Logger.log("--- Testing filterCouponsByDateRange (InputDate) ---");
  // Adjust dates based on your test data
  var filteredByInput = filterCouponsByDateRange("2024-01-01", "2024-01-31", "InputDate");
  if (filteredByInput.length > 0) {
    Logger.log("Filtered by InputDate (Jan 2024): " + filteredByInput.length);
    filteredByInput.forEach(function(c) { Logger.log(JSON.stringify(c)); });
  } else {
    Logger.log("No coupons found for InputDate in Jan 2024.");
  }

  Logger.log("--- Testing filterCouponsByDateRange (ExpiryDate) ---");
  // Adjust dates based on your test data
  var filteredByExpiry = filterCouponsByDateRange("2024-12-01", "2024-12-31", "ExpiryDate");
    if (filteredByExpiry.length > 0) {
    Logger.log("Filtered by ExpiryDate (Dec 2024): " + filteredByExpiry.length);
    filteredByExpiry.forEach(function(c) { Logger.log(JSON.stringify(c)); });
  } else {
    Logger.log("No coupons found for ExpiryDate in Dec 2024.");
  }
}

function testLogUsage() {
  Logger.log("--- Testing logCouponUsage ---");
  // NB! This test MODIFIES data in your sheet.
  // Ensure you have a test coupon, e.g., "TEST12345" which is monetary and has balance.
  // First, check current state or add one if needed via testSaveCoupon or manually.
  // Example: Logger.log(saveCouponData("MONEYTEST01", "2024-01-01", "2025-01-01", true, 100, "mt.png", "Test monetary"));

  var result1 = logCouponUsage("TEST12345", 10.00); // Assuming TEST12345 exists and has >=10
  Logger.log("Usage attempt 1: " + JSON.stringify(result1));

  var result2 = logCouponUsage("TEST12345", 100.00); // Attempt to use more (may fail if balance is low)
  Logger.log("Usage attempt 2: " + JSON.stringify(result2));

  var result3 = logCouponUsage("NONEXISTENT123", 5.00); // Non-existent coupon
  Logger.log("Usage attempt 3 (non-existent): " + JSON.stringify(result3));

  var result4 = logCouponUsage("PROMOABC", 5.00); // Assuming PROMOABC is not monetary
  Logger.log("Usage attempt 4 (not monetary): " + JSON.stringify(result4));
}
