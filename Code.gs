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
      .setTitle('쿠폰 관리 시스템') // 웹 앱 제목 설정
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

    // 데이터 URL 프리픽스 제거 (예: "data:image/png;base64,")
    var pureBase64Data = base64ImageData.includes(',') ? base64ImageData.split(',')[1] : base64ImageData;

    var decodedData = Utilities.base64Decode(pureBase64Data, Utilities.Charset.UTF_8);
    var blob = Utilities.newBlob(decodedData, mimeType, imageName);

    var file = folder.createFile(blob);
    // file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW); // 선택 사항: 공유 설정 - 개인 정보 보호에 주의

    Logger.log("이미지가 성공적으로 업로드되었습니다: " + file.getName() + ", ID: " + file.getId() + ", Link: " + file.getUrl()); // Log with getUrl()
    return JSON.stringify({
      success: true,
      fileId: file.getId(),
      webViewLink: file.getUrl() // Changed to getUrl()
    });

  } catch (e) {
    Logger.log("uploadImageToDrive 오류: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify({ success: false, message: "이미지 업로드 중 오류 발생: " + e.toString() });
  }
}


/**
 * 새 쿠폰을 스프레드시트에 저장합니다. 이미지 파일 ID와 링크를 포함할 수 있습니다.
 * @param {Object} couponData 쿠폰 상세 정보 객체:
 *                            {barcode: string, expiryDate: string, isGiftCertificate: boolean,
 *                             initialBalance: number | null, imageFileId?: string, imageWebViewLink?: string}
 * @return {string} JSON 문자열 형태의 성공 또는 오류 메시지 객체.
 */
function saveCoupon(couponData) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    // 헤더 정의 (이미지 관련 열 추가)
    var headers = ["바코드", "입력일", "만료일", "금액권 여부", "잔액", "최초 금액", "이미지 파일 ID", "이미지 보기 링크"];

    if (!sheet) {
      // 시트가 없으면 새로 생성하고 헤더를 추가합니다.
      sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(COUPON_SHEET_NAME);
      sheet.appendRow(headers);
      Logger.log("'" + COUPON_SHEET_NAME + "' 시트가 생성되고 헤더가 추가되었습니다.");
    } else {
      // 시트가 이미 존재하는 경우, 헤더를 확인합니다.
      if (sheet.getLastRow() > 0) {
        var currentHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
        var headersMatch = headers.length === currentHeaders.length && headers.every(function(value, index) {
          return value === currentHeaders[index];
        });
        if (!headersMatch) {
          Logger.log("주의: '" + COUPON_SHEET_NAME + "' 시트의 헤더가 예상과 다릅니다. 현재 헤더: [" + currentHeaders.join(", ") + "], 예상 헤더: [" + headers.join(", ") + "]. 이미지 관련 기능이 올바르게 작동하지 않을 수 있습니다. 수동으로 헤더를 업데이트하거나 새 시트를 사용하세요.");
          // 이 단계에서는 기존 시트의 헤더를 강제로 수정하지 않습니다.
          // 필요시 사용자가 수동으로 "이미지 파일 ID", "이미지 보기 링크" 열을 추가해야 합니다.
        }
      } else {
        // 시트는 존재하지만 비어있는 경우 (행이 없음), 헤더를 추가합니다.
        sheet.appendRow(headers);
        Logger.log("'" + COUPON_SHEET_NAME + "' 시트가 비어있어 헤더가 추가되었습니다.");
      }
    }

    var entryDate = new Date(); // 현재 날짜를 입력일로 사용
    var balance = couponData.isGiftCertificate ? parseFloat(couponData.initialBalance) : null;
    var originalAmount = balance; // 참고용으로 원래 금액 저장

    // 데이터 유효성 검사
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

    // 시트에 새 행으로 쿠폰 정보 추가
    sheet.appendRow([
      couponData.barcode,
      entryDate,
      expiryDateObj, // Date 객체로 저장
      couponData.isGiftCertificate,
      balance,
      originalAmount,
      couponData.imageFileId || null, // 이미지 파일 ID (없으면 null)
      couponData.imageWebViewLink || null // 이미지 보기 링크 (없으면 null)
    ]);
    return JSON.stringify({ success: true, message: "쿠폰이 성공적으로 저장되었습니다: " + couponData.barcode });
  } catch (e) {
    Logger.log("saveCoupon 함수 오류: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify({ success: false, message: "쿠폰 저장 중 오류 발생: " + e.toString() });
  }
}

/**
 * 스프레드시트에서 모든 쿠폰을 가져옵니다.
 * 입력일 기준 내림차순 정렬 (최신순).
 * @return {String} 쿠폰 배열의 JSON 문자열.
 */
function getCoupons() {
  var returnValue = []; // 반환될 배열 초기화
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    // 스프레드시트를 찾을 수 없는 경우
    if (!ss) {
      Logger.log("getCoupons 오류: ID가 " + SPREADSHEET_ID + "인 스프레드시트를 찾을 수 없습니다. 빈 배열 문자열을 반환합니다.");
    } else {
      var sheet = ss.getSheetByName(COUPON_SHEET_NAME);
      // 쿠폰 시트를 찾을 수 없는 경우
      if (!sheet) {
        Logger.log("getCoupons 정보: 쿠폰 시트 '" + COUPON_SHEET_NAME + "'를 찾을 수 없습니다. 빈 배열 문자열을 반환합니다.");
      } else {
        // 시트가 완전히 비어있는 경우 (행이 없는 경우)
        if (sheet.getLastRow() === 0) {
          Logger.log("getCoupons 정보: 쿠폰 시트가 완전히 비어있습니다. 빈 배열 문자열을 반환합니다.");
        } else {
          var dataRange = sheet.getDataRange();
          // 데이터 범위가 없는 경우 (예: 시트는 있지만 데이터가 전혀 없는 경우)
          if (!dataRange) {
            Logger.log("getCoupons 정보: 시트에서 데이터 범위를 찾을 수 없습니다. 빈 배열 문자열을 반환합니다.");
          } else {
            var data = dataRange.getValues();
            // 데이터가 없거나 헤더 행만 있는 경우
            // 이미지 관련 열이 추가되었으므로, data[i][6] (이미지 파일 ID), data[i][7] (이미지 보기 링크)을 읽도록 인덱스 조정 필요
            if (!data || sheet.getLastRow() <= 1) {
              Logger.log("getCoupons 정보: 실제 쿠폰 데이터가 없습니다. 마지막 행: " + sheet.getLastRow() + ". 데이터 길이: " + (data ? data.length : "null") + ". 빈 배열 문자열을 반환합니다.");
            } else {
              // 모든 검사를 통과하면 데이터를 처리합니다.
              var couponsData = [];
              // data[0]은 헤더 행이므로 1부터 시작합니다.
              for (var i = 1; i < data.length; i++) {
                // 바코드(0번 열)가 비어있지 않은지 확인하여 유효한 행만 처리합니다.
                if (data[i] && data[i][0] != "") {
                  couponsData.push({
                    barcode: data[i][0],
                    entryDate: data[i][1],
                    expiryDate: data[i][2],
                    isGiftCertificate: data[i][3],
                    balance: data[i][4],
                    originalAmount: data[i][5],
                    imageFileId: data[i][6] || null, // 이미지 파일 ID
                    imageWebViewLink: data[i][7] || null // 이미지 보기 링크
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
 * 금액권 사용 내역을 기록하고 잔액을 업데이트합니다.
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
      return JSON.stringify({ success: false, message: "오류: 바코드 '" + barcode + "'에 해당하는 쿠폰을 찾을 수 없습니다." });
    }
    if (!isGift) {
      return JSON.stringify({ success: false, message: "오류: 쿠폰 '" + barcode + "'은 금액권이 아닙니다." });
    }
    if (isNaN(currentBalance) || currentBalance < amountUsed) {
      return JSON.stringify({ success: false, message: "오류: 잔액이 부족합니다. 현재 잔액: " + (isNaN(currentBalance) ? 0 : currentBalance.toFixed(2)) });
    }

    var newBalance = currentBalance - amountUsed;
    couponSheet.getRange(couponRow + 1, 5).setValue(newBalance);

    var logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USAGE_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(USAGE_LOG_SHEET_NAME);
      logSheet.appendRow(["Timestamp", "Barcode", "Amount Used", "New Balance"]);
    }
    logSheet.appendRow([new Date(), barcode, amountUsed, newBalance]);

    return JSON.stringify({ success: true, message: "사용 내역이 성공적으로 기록되었습니다. " + barcode + "의 새 잔액: " + newBalance.toFixed(2), newBalance: newBalance });
  } catch (e) {
    Logger.log("logGiftCertificateUsage 함수 오류: " + e.toString());
    return JSON.stringify({ success: false, message: "사용 내역 기록 중 오류 발생: " + e.toString() });
  }
}

// 테스트용 헬퍼 함수 (선택 사항, Apps Script 편집기에서 실행 가능)
function testSaveCoupon() {
  // 참고: 이제 saveCoupon 등의 함수 반환값은 JSON 문자열로 로깅됩니다.
  Logger.log(saveCoupon({barcode: "TEST_IMG_001", expiryDate: "2025-12-31", isGiftCertificate: false, initialBalance: null, imageFileId: "someFileId", imageWebViewLink: "someLink"}));
  Logger.log(saveCoupon({barcode: "GIFT_IMG_002", expiryDate: "2026-06-30", isGiftCertificate: true, initialBalance: 150, imageFileId: "anotherFileId", imageWebViewLink: "anotherLink"}));
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
  // "GIFT_IMG_002" 바코드의 쿠폰이 존재하고 금액권이며 잔액이 있는지 확인하세요.
  // Logger.log(saveCoupon({barcode: "GIFT_IMG_002", expiryDate: "2026-06-30", isGiftCertificate: true, initialBalance: 150})); // 미리 저장되어 있지 않다면 주석 해제
  Logger.log(logGiftCertificateUsage("GIFT_IMG_002", 30));
}

function testUploadImage() {
  // 이 함수를 테스트하려면 실제 base64 문자열이 필요합니다.
  // 간단한 1x1 픽셀 PNG 이미지 Base64 예시 (실제로는 더 큰 데이터 사용)
  var testBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
  var testName = "testImage.png";
  var testMime = "image/png";
  Logger.log(uploadImageToDrive(testBase64, testName, testMime));
}
