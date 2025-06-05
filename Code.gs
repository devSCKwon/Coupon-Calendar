// 전역 변수 - 사용자가 직접 입력해야 하는 부분도 있었으나, 제공된 ID로 일부 설정됨
var SPREADSHEET_ID = "1jLqURqn9hoHZB1r32RC3q6A_U52YiU5Vjc4_0LfV1Tk";
var DRIVE_FOLDER_ID = "19Gm6k5Jf1qQDT4YfddElLC7QO7jvySfQ"; // 바코드 이미지 저장 폴더 ID
var COUPON_SHEET_NAME = "Coupons"; // 쿠폰 정보 시트 이름
var USAGE_LOG_SHEET_NAME = "UsageLog"; // 사용 기록 시트 이름

/**
 * 웹 앱의 HTML 인터페이스를 제공합니다.
 * @param {Object} e 이벤트 매개변수.
 * @return {HtmlOutput} 웹 앱의 HTML 출력.
 */
function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index.html')
      .setTitle('쿠폰 관리 시스템')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT);
}

/**
 * 이미지를 Base64 데이터로 받아 Google Drive의 지정된 폴더에 업로드하고 파일 정보를 반환합니다.
 * @param {string} base64ImageData Base64로 인코딩된 이미지 데이터.
 * @param {string} imageName 저장할 이미지 파일 이름.
 * @param {string} mimeType 이미지의 MIME 타입 (예: 'image/png', 'image/jpeg').
 * @return {string} JSON 문자열 형태의 결과 객체: { success: true, fileId: string, webViewLink: string } 또는 { success: false, message: string }.
 */
function uploadImageToDrive(base64ImageData, imageName, mimeType) {
  try {
    var folder;
    try {
      folder = DriveApp.getFolderById(DRIVE_FOLDER_ID);
    } catch (e) {
      Logger.log("uploadImageToDrive 오류: Drive 폴더 ID '" + DRIVE_FOLDER_ID + "'를 찾는 중 오류 발생: " + e.toString());
      return JSON.stringify({ success: false, message: "지정된 구글 드라이브 폴더를 찾을 수 없습니다. 폴더 ID (" + DRIVE_FOLDER_ID + ")를 확인하거나 폴더를 생성해주세요." });
    }

    if (!base64ImageData || !imageName || !mimeType) {
        return JSON.stringify({ success: false, message: "이미지 데이터, 파일 이름, MIME 타입은 필수입니다."});
    }

    var pureBase64Data = base64ImageData.includes(',') ? base64ImageData.split(',')[1] : base64ImageData;
    var decodedData = Utilities.base64Decode(pureBase64Data, Utilities.Charset.UTF_8);
    var blob = Utilities.newBlob(decodedData, mimeType, imageName);
    var file = folder.createFile(blob);

    Logger.log("이미지가 성공적으로 업로드되었습니다: " + file.getName() + ", ID: " + file.getId() + ", Link: " + file.getUrl());
    return JSON.stringify({
      success: true,
      fileId: file.getId(),
      webViewLink: file.getUrl()
    });

  } catch (e) {
    Logger.log("uploadImageToDrive 오류: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify({ success: false, message: "이미지 업로드 중 오류 발생: " + e.toString() });
  }
}

/**
 * 새 쿠폰을 스프레드시트에 저장합니다. 이미지 파일 ID, 링크, 사용 상태, 만료 상태를 포함합니다.
 * @param {Object} couponData 쿠폰 상세 정보 객체
 * @return {string} JSON 문자열 형태의 성공 또는 오류 메시지 객체.
 */
function saveCoupon(couponData) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    var headers = ["바코드", "입력일", "만료일", "금액권 여부", "잔액", "최초 금액", "이미지 파일 ID", "이미지 보기 링크", "사용 완료", "만료 상태"];

    if (!sheet) {
      sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(COUPON_SHEET_NAME);
      sheet.appendRow(headers);
      Logger.log("'" + COUPON_SHEET_NAME + "' 시트가 생성되고 헤더가 추가되었습니다.");
    } else {
      if (sheet.getLastRow() > 0) {
        var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var headersMatch = headers.length === currentHeaders.length && headers.every(function(value, index) {
          return value === currentHeaders[index];
        });
        if (!headersMatch) {
          Logger.log("주의: '" + COUPON_SHEET_NAME + "' 시트의 헤더가 예상과 다릅니다. 현재 헤더: [" + currentHeaders.join(", ") + "], 예상 헤더: [" + headers.join(", ") + "]. 기능이 올바르게 작동하지 않을 수 있습니다.");
        }
      } else {
        sheet.appendRow(headers);
        Logger.log("'" + COUPON_SHEET_NAME + "' 시트가 비어있어 헤더가 추가되었습니다.");
      }
    }

    var entryDate = new Date();
    var balance = couponData.isGiftCertificate ? parseFloat(couponData.initialBalance) : null;
    var originalAmount = balance;

    if (!couponData.barcode || couponData.barcode.trim() === "") {
      return JSON.stringify({ success: false, message: "바코드는 필수 항목입니다." });
    }
    if (!couponData.expiryDate) {
      return JSON.stringify({ success: false, message: "만료일은 필수 항목입니다." });
    }
    var expiryDateObj = new Date(couponData.expiryDate);
    if (isNaN(expiryDateObj.getTime())) {
        return JSON.stringify({ success: false, message: "유효하지 않은 만료일 형식입니다."});
    }
    if (couponData.isGiftCertificate && (isNaN(balance) || balance === null || balance < 0)) {
        return JSON.stringify({ success: false, message: "금액권의 초기 금액은 0 이상의 숫자여야 합니다."});
    }

    sheet.appendRow([
      couponData.barcode,
      entryDate,
      expiryDateObj,
      couponData.isGiftCertificate,
      balance,
      originalAmount,
      couponData.imageFileId || null,
      couponData.imageWebViewLink || null,
      false, // 사용 완료 (기본값: false)
      "사용가능" // 만료 상태 (기본값: "사용가능")
    ]);
    return JSON.stringify({ success: true, message: "쿠폰이 성공적으로 저장되었습니다: " + couponData.barcode });
  } catch (e) {
    Logger.log("saveCoupon 함수 오류: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify({ success: false, message: "쿠폰 저장 중 오류 발생: " + e.toString() });
  }
}

/**
 * 스프레드시트에서 모든 쿠폰 정보를 가져옵니다 (사용 상태 및 만료 상태 포함).
 * @return {String} 쿠폰 배열의 JSON 문자열.
 */
function getCoupons() {
  var returnValue = [];
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    if (!ss) {
      Logger.log("getCoupons 오류: 스프레드시트를 찾을 수 없습니다.");
    } else {
      var sheet = ss.getSheetByName(COUPON_SHEET_NAME);
      if (!sheet) {
        Logger.log("getCoupons 정보: 쿠폰 시트 '" + COUPON_SHEET_NAME + "'를 찾을 수 없습니다.");
      } else {
        if (sheet.getLastRow() === 0) {
          Logger.log("getCoupons 정보: 쿠폰 시트가 비어있습니다.");
        } else {
          var dataRange = sheet.getDataRange();
          if (!dataRange) {
            Logger.log("getCoupons 정보: 데이터 범위를 찾을 수 없습니다.");
          } else {
            var data = dataRange.getValues();
            if (!data || sheet.getLastRow() <= 1) {
              Logger.log("getCoupons 정보: 실제 쿠폰 데이터가 없습니다.");
            } else {
              var couponsData = [];
              var headers = data[0]; // 첫 번째 행을 헤더로 간주
              // 필요한 열의 인덱스를 헤더 기반으로 찾기
              var barcodeIdx = headers.indexOf("바코드");
              var entryDateIdx = headers.indexOf("입력일");
              var expiryDateIdx = headers.indexOf("만료일");
              var isGiftCertIdx = headers.indexOf("금액권 여부");
              var balanceIdx = headers.indexOf("잔액");
              var originalAmountIdx = headers.indexOf("최초 금액");
              var imageFileIdIdx = headers.indexOf("이미지 파일 ID");
              var imageWebViewLinkIdx = headers.indexOf("이미지 보기 링크");
              var isUsedIdx = headers.indexOf("사용 완료");
              var expiryStatusIdx = headers.indexOf("만료 상태");

              for (var i = 1; i < data.length; i++) {
                if (data[i] && data[i][barcodeIdx] != "") {
                  couponsData.push({
                    barcode: data[i][barcodeIdx],
                    entryDate: data[i][entryDateIdx],
                    expiryDate: data[i][expiryDateIdx],
                    isGiftCertificate: data[i][isGiftCertIdx],
                    balance: data[i][balanceIdx],
                    originalAmount: data[i][originalAmountIdx],
                    imageFileId: imageFileIdIdx !== -1 ? (data[i][imageFileIdIdx] || null) : null,
                    imageWebViewLink: imageWebViewLinkIdx !== -1 ? (data[i][imageWebViewLinkIdx] || null) : null,
                    isUsed: isUsedIdx !== -1 ? (data[i][isUsedIdx]) : false, // 기본값 false
                    expiryStatus: expiryStatusIdx !== -1 ? (data[i][expiryStatusIdx]) : "정보없음" // 기본값
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
    Logger.log("getCoupons (try) 반환 전 값 (문자열화 전): " + JSON.stringify(returnValue));
    return JSON.stringify(returnValue);
  } catch (e) {
    Logger.log("getCoupons 예외 발생: " + e.toString() + " 스택: " + e.stack);
    var fallbackValue = [];
    Logger.log("getCoupons (catch) 반환 전 값 (문자열화 전): " + JSON.stringify(fallbackValue));
    return JSON.stringify(fallbackValue);
  }
}

/**
 * 일반 쿠폰을 사용 완료로 처리하고 만료 상태를 업데이트합니다.
 * @param {string} barcode 처리할 쿠폰의 바코드.
 * @return {string} JSON 문자열 형태의 결과 객체.
 */
function markCouponAsUsed(barcode) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    if (!sheet) {
      return JSON.stringify({ success: false, message: "쿠폰 시트를 찾을 수 없습니다." });
    }
    if (sheet.getLastRow() === 0) {
        return JSON.stringify({ success: false, message: "쿠폰 시트가 비어있습니다." });
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var barcodeColIdx = headers.indexOf("바코드");
    var isGiftCertColIdx = headers.indexOf("금액권 여부");
    var isUsedColIdx = headers.indexOf("사용 완료");
    var expiryStatusColIdx = headers.indexOf("만료 상태");

    if (barcodeColIdx === -1 || isGiftCertColIdx === -1 || isUsedColIdx === -1 || expiryStatusColIdx === -1) {
        return JSON.stringify({ success: false, message: "시트 헤더 구성이 올바르지 않습니다. ('바코드', '금액권 여부', '사용 완료', '만료 상태' 열 필요)" });
    }

    for (var i = 1; i < data.length; i++) {
      if (data[i][barcodeColIdx] == barcode) {
        if (data[i][isGiftCertColIdx] === true || data[i][isGiftCertColIdx] === 'TRUE' || data[i][isGiftCertColIdx].toString().toLowerCase() === 'true') {
          return JSON.stringify({ success: false, message: "금액권은 이 기능을 사용할 수 없습니다. '사용 기록' 기능을 이용해주세요." });
        }

        if (data[i][isUsedColIdx] === true || data[i][isUsedColIdx] === 'TRUE' || data[i][isUsedColIdx].toString().toLowerCase() === 'true') {
            return JSON.stringify({ success: false, message: "이미 사용 완료 처리된 쿠폰입니다." });
        }

        sheet.getRange(i + 1, isUsedColIdx + 1).setValue(true);
        sheet.getRange(i + 1, expiryStatusColIdx + 1).setValue("사용완료");

        Logger.log("일반 쿠폰 사용 완료: " + barcode);
        return JSON.stringify({ success: true, message: "'" + barcode + "' 쿠폰이 사용 완료 처리되었습니다." });
      }
    }
    return JSON.stringify({ success: false, message: "바코드 '" + barcode + "'에 해당하는 쿠폰을 찾을 수 없습니다." });
  } catch (e) {
    Logger.log("markCouponAsUsed 오류: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify({ success: false, message: "쿠폰 사용 완료 처리 중 오류 발생: " + e.toString() });
  }
}

/**
 * 가장 최근에 추가된 5개의 쿠폰을 가져옵니다.
 * @return {String} 5개 최신 쿠폰 객체 배열의 JSON 문자열.
 */
function getLatestCoupons() {
  try {
    var couponsJsonString = getCoupons();
    var allCoupons = JSON.parse(couponsJsonString);

    if (!Array.isArray(allCoupons)) {
        Logger.log("getLatestCoupons 오류: getCoupons에서 파싱한 데이터가 배열이 아닙니다. 데이터: " + couponsJsonString);
        return JSON.stringify([]);
    }

    var latest = allCoupons.slice(0, 5);
    Logger.log("getLatestCoupons 반환 전 값 (문자열화 전): " + JSON.stringify(latest));
    return JSON.stringify(latest);
  } catch (e) {
    Logger.log("getLatestCoupons 예외 발생: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify([]);
  }
}

/**
 * 금액권 사용 내역을 기록하고 잔액을 업데이트하며, 잔액이 0이 되면 만료 상태를 업데이트합니다.
 * @param {String} barcode 금액권의 바코드.
 * @param {Number} amountUsed 사용한 금액.
 * @return {String} 성공 여부, 메시지, 새 잔액 등을 포함한 객체의 JSON 문자열.
 */
function logGiftCertificateUsage(barcode, amountUsed) {
  try {
    if (typeof amountUsed !== 'number' || amountUsed <= 0) {
      return JSON.stringify({ success: false, message: "오류: 사용 금액은 0보다 큰 숫자여야 합니다." });
    }

    var couponSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    if (!couponSheet) {
      return JSON.stringify({ success: false, message: "오류: 쿠폰 시트를 찾을 수 없습니다." });
    }
    if (couponSheet.getLastRow() === 0) {
        return JSON.stringify({ success: false, message: "오류: 쿠폰 시트가 비어있습니다." });
    }

    var data = couponSheet.getDataRange().getValues();
    var headers = data[0];
    var barcodeColIdx = headers.indexOf("바코드");
    var isGiftCertColIdx = headers.indexOf("금액권 여부");
    var balanceColIdx = headers.indexOf("잔액");
    var expiryStatusColIdx = headers.indexOf("만료 상태");

    if (barcodeColIdx === -1 || isGiftCertColIdx === -1 || balanceColIdx === -1 || expiryStatusColIdx === -1) {
        return JSON.stringify({ success: false, message: "시트 헤더 구성이 올바르지 않습니다. ('바코드', '금액권 여부', '잔액', '만료 상태' 열 필요)" });
    }

    var couponRow = -1;
    var currentBalance = 0;
    var isGift = false;

    for (var i = 1; i < data.length; i++) {
      if (data[i][barcodeColIdx] == barcode) {
        couponRow = i; // 실제 데이터 행의 인덱스 (0부터 시작하는 배열 기준)
        isGift = data[i][isGiftCertColIdx];
        currentBalance = parseFloat(data[i][balanceColIdx]);
        break;
      }
    }

    if (couponRow === -1) {
      return JSON.stringify({ success: false, message: "오류: 바코드 '" + barcode + "'에 해당하는 쿠폰을 찾을 수 없습니다." });
    }
    if (!(isGift === true || isGift === 'TRUE' || isGift.toString().toLowerCase() === 'true')) {
      return JSON.stringify({ success: false, message: "오류: 쿠폰 '" + barcode + "'은 금액권이 아닙니다." });
    }
    if (isNaN(currentBalance) || currentBalance < amountUsed) {
      return JSON.stringify({ success: false, message: "오류: 잔액이 부족합니다. 현재 잔액: " + (isNaN(currentBalance) ? 0 : currentBalance.toFixed(2)) });
    }

    var newBalance = currentBalance - amountUsed;
    couponSheet.getRange(couponRow + 1, balanceColIdx + 1).setValue(newBalance);

    if (newBalance === 0) {
        couponSheet.getRange(couponRow + 1, expiryStatusColIdx + 1).setValue("잔액없음");
        Logger.log("금액권 잔액 없음 처리: " + barcode);
    }

    var logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USAGE_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(USAGE_LOG_SHEET_NAME);
      logSheet.appendRow(["Timestamp", "Barcode", "Amount Used", "New Balance"]);
    }
    logSheet.appendRow([new Date(), barcode, amountUsed, newBalance]);

    var message = "사용 내역이 성공적으로 기록되었습니다. " + barcode + "의 새 잔액: " + newBalance.toFixed(2);
    if (newBalance === 0) {
        message += " (잔액 소진)";
    }
    return JSON.stringify({ success: true, message: message, newBalance: newBalance });
  } catch (e) {
    Logger.log("logGiftCertificateUsage 함수 오류: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify({ success: false, message: "사용 내역 기록 중 오류 발생: " + e.toString() });
  }
}

// 테스트용 헬퍼 함수
function testSaveCoupon() {
  Logger.log(saveCoupon({barcode: "NEW_COUPON_001", expiryDate: "2027-12-31", isGiftCertificate: false, initialBalance: null}));
  Logger.log(saveCoupon({barcode: "GIFT_COUPON_003", expiryDate: "2027-06-30", isGiftCertificate: true, initialBalance: 200}));
}

function testMarkCouponAsUsed() {
    // "NEW_COUPON_001" 쿠폰이 일반 쿠폰으로 존재해야 함
    Logger.log(markCouponAsUsed("NEW_COUPON_001"));
    // 이미 사용된 쿠폰 테스트
    // Logger.log(markCouponAsUsed("NEW_COUPON_001"));
    // 금액권 사용 시도 테스트
    // Logger.log(markCouponAsUsed("GIFT_COUPON_003"));
}

function testGetCoupons() {
  var couponsJson = getCoupons();
  Logger.log("getCoupons() 원시 JSON:");
  Logger.log(couponsJson);
  Logger.log("getCoupons() 파싱된 쿠폰:");
  Logger.log(JSON.parse(couponsJson));
}

function testGetLatestCoupons() {
  var latestCouponsJson = getLatestCoupons();
  Logger.log("getLatestCoupons() 원시 JSON:");
  Logger.log(latestCouponsJson);
  Logger.log("getLatestCoupons() 파싱된 쿠폰:");
  Logger.log(JSON.parse(latestCouponsJson));
}

function testLogUsage() {
  // "GIFT_COUPON_003" 쿠폰이 존재하고 잔액이 있어야 함
  // Logger.log(saveCoupon({barcode: "GIFT_COUPON_003", expiryDate: "2027-06-30", isGiftCertificate: true, initialBalance: 200}));
  Logger.log(logGiftCertificateUsage("GIFT_COUPON_003", 50));
  Logger.log(logGiftCertificateUsage("GIFT_COUPON_003", 150)); // 잔액 0 만드는 테스트
}

function testUploadImage() {
  var testBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
  var testName = "testImage.png";
  var testMime = "image/png";
  Logger.log(uploadImageToDrive(testBase64, testName, testMime));
}
