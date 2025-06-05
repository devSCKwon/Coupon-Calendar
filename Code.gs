// 전역 변수 - 사용자가 직접 입력해야 하는 부분도 있었으나, 제공된 ID로 일부 설정됨
var SPREADSHEET_ID = "1jLqURqn9hoHZB1r32RC3q6A_U52YiU5Vjc4_0LfV1Tk";
var DRIVE_FOLDER_ID = "19Gm6k5Jf1qQDT4YfddElLC7QO7jvySfQ"; // 바코드 이미지 저장 폴더 ID (현재 기능에서는 직접 사용되지 않음)
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
 * 새 쿠폰을 스프레드시트에 저장합니다.
 * @param {Object} couponData 쿠폰 상세 정보 객체:
 *                            {barcode: string, expiryDate: string, isGiftCertificate: boolean, initialBalance: number | null}
 * @return {String} JSON 문자열 형태의 성공 또는 오류 객체.
 */
function saveCoupon(couponData) {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    // 시트가 없으면 헤더와 함께 새로 생성합니다.
    if (!sheet) {
      sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(COUPON_SHEET_NAME);
      sheet.appendRow(["Barcode", "Entry Date", "Expiration Date", "Is Gift Certificate", "Balance", "Original Amount"]);
    }

    var entryDate = new Date(); // 현재 날짜를 입력일로 사용
    var balance = couponData.isGiftCertificate ? parseFloat(couponData.initialBalance) : null;
    var originalAmount = balance; // 참고용으로 원래 금액 저장

    // 데이터 유효성 검사
    if (!couponData.barcode || couponData.barcode.trim() === "") {
      return JSON.stringify({ success: false, message: "오류: 바코드는 비워둘 수 없습니다." });
    }
    if (!couponData.expiryDate) {
      return JSON.stringify({ success: false, message: "오류: 만료일은 비워둘 수 없습니다." });
    }
    // 만료일이 문자열로 오더라도 Date 객체로 처리합니다.
    var expiryDateObj = new Date(couponData.expiryDate);
    if (isNaN(expiryDateObj.getTime())) {
        return JSON.stringify({ success: false, message: "오류: 유효하지 않은 만료일 형식입니다."});
    }

    if (couponData.isGiftCertificate && (isNaN(balance) || balance === null || balance < 0)) {
        return JSON.stringify({ success: false, message: "오류: 금액권의 초기 잔액은 0 이상의 숫자여야 합니다."});
    }

    // 시트에 새 행으로 쿠폰 정보 추가
    sheet.appendRow([
      couponData.barcode,
      entryDate,
      expiryDateObj, // Date 객체로 저장
      couponData.isGiftCertificate,
      balance,
      originalAmount
    ]);
    return JSON.stringify({ success: true, message: "쿠폰이 성공적으로 저장되었습니다: " + couponData.barcode });
  } catch (e) {
    Logger.log("saveCoupon 함수 오류: " + e.toString());
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
                    entryDate: data[i][1], // Date 객체 형태입니다 (스프레드시트에서 직접 가져올 때)
                    expiryDate: data[i][2], // Date 객체 형태입니다
                    isGiftCertificate: data[i][3],
                    balance: data[i][4],
                    originalAmount: data[i][5]
                  });
                }
              }

              // 입력일을 기준으로 내림차순 정렬 (최신 쿠폰이 먼저 오도록)
              couponsData.sort(function(a, b) {
                // getValues()로 가져온 날짜는 이미 Date 객체일 가능성이 높지만, 문자열일 경우를 대비하여 Date 객체로 변환합니다.
                var dateA = (a.entryDate instanceof Date) ? a.entryDate : new Date(a.entryDate);
                var dateB = (b.entryDate instanceof Date) ? b.entryDate : new Date(b.entryDate);
                return dateB - dateA; // 최신순 정렬
              });
              returnValue = couponsData; // 기본 반환 변수에 처리된 데이터를 할당합니다.
            }
          }
        }
      }
    }
    // try 블록의 반환값 로깅 (문자열로 변환하기 전의 객체)
    Logger.log("getCoupons (try) 반환 전 값 (문자열화 전): " + JSON.stringify(returnValue));
    return JSON.stringify(returnValue); // 여기서 최종 반환값을 문자열로 변환합니다.
  } catch (e) {
    Logger.log("getCoupons 예외 발생: " + e.toString() + " 스택: " + e.stack);
    var fallbackValue = []; // 예외 발생 시 반환할 기본값 (빈 배열)
    // catch 블록의 반환값 로깅 (문자열로 변환하기 전의 객체)
    Logger.log("getCoupons (catch) 반환 전 값 (문자열화 전): " + JSON.stringify(fallbackValue));
    return JSON.stringify(fallbackValue); // 여기서 최종 반환값을 문자열로 변환합니다.
  }
}

/**
 * 가장 최근에 추가된 5개의 쿠폰을 가져옵니다.
 * @return {String} 5개 최신 쿠폰 객체 배열의 JSON 문자열.
 */
function getLatestCoupons() {
  try {
    var couponsJsonString = getCoupons(); // getCoupons()는 이제 JSON 문자열을 반환합니다.
    var allCoupons = JSON.parse(couponsJsonString); // JSON 문자열을 파싱합니다.

    // 파싱 후 allCoupons가 배열인지 확인 (getCoupons가 이를 보장해야 함)
    if (!Array.isArray(allCoupons)) {
        Logger.log("getLatestCoupons 오류: getCoupons에서 파싱한 데이터가 배열이 아닙니다. 데이터: " + couponsJsonString);
        return JSON.stringify([]); // 빈 배열의 문자열을 반환합니다.
    }

    var latest = allCoupons.slice(0, 5); // 처음 5개 요소를 가져옵니다.
    Logger.log("getLatestCoupons 반환 전 값 (문자열화 전): " + JSON.stringify(latest));
    return JSON.stringify(latest); // 다시 문자열로 변환하여 반환합니다.
  } catch (e) {
    Logger.log("getLatestCoupons 예외 발생: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify([]); // 오류 발생 시 빈 배열의 문자열을 반환합니다.
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
    // 사용 금액 유효성 검사
    if (typeof amountUsed !== 'number' || amountUsed <= 0) {
      return JSON.stringify({ success: false, message: "오류: 사용 금액은 0보다 큰 숫자여야 합니다." });
    }

    var couponSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    if (!couponSheet) {
      return JSON.stringify({ success: false, message: "오류: 쿠폰 시트를 찾을 수 없습니다." });
    }

    var data = couponSheet.getDataRange().getValues();
    var couponRow = -1; // 해당 쿠폰의 행 번호 (0부터 시작)
    var currentBalance = 0; // 현재 잔액
    var isGift = false; // 금액권 여부

    // 헤더를 제외하고 쿠폰 검색 (일반적으로 1번 행부터 데이터 시작)
    for (var i = 1; i < data.length; i++) {
      if (data[i][0] == barcode) { // 바코드(A열) 일치 확인
        couponRow = i;
        isGift = data[i][3]; // 금액권 여부(D열)
        currentBalance = parseFloat(data[i][4]); // 잔액(E열)
        break;
      }
    }

    // 쿠폰을 찾지 못한 경우
    if (couponRow === -1) {
      return JSON.stringify({ success: false, message: "오류: 바코드 '" + barcode + "'에 해당하는 쿠폰을 찾을 수 없습니다." });
    }
    // 금액권이 아닌 경우
    if (!isGift) {
      return JSON.stringify({ success: false, message: "오류: 쿠폰 '" + barcode + "'은 금액권이 아닙니다." });
    }
    // 잔액 부족 또는 유효하지 않은 잔액
    if (isNaN(currentBalance) || currentBalance < amountUsed) {
      return JSON.stringify({ success: false, message: "오류: 잔액이 부족합니다. 현재 잔액: " + (isNaN(currentBalance) ? 0 : currentBalance.toFixed(2)) });
    }

    var newBalance = currentBalance - amountUsed;
    // 쿠폰 시트의 잔액 업데이트 (getRange는 1부터 시작하는 인덱스 사용, 행은 couponRow + 1)
    couponSheet.getRange(couponRow + 1, 5).setValue(newBalance);

    // 사용 내역 기록
    var logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USAGE_LOG_SHEET_NAME);
    // 사용 내역 시트가 없으면 헤더와 함께 새로 생성
    if (!logSheet) {
      logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(USAGE_LOG_SHEET_NAME);
      logSheet.appendRow(["Timestamp", "Barcode", "Amount Used", "New Balance"]);
    }
    logSheet.appendRow([new Date(), barcode, amountUsed, newBalance]); // 새 사용 내역 추가

    return JSON.stringify({ success: true, message: "사용 내역이 성공적으로 기록되었습니다. " + barcode + "의 새 잔액: " + newBalance.toFixed(2), newBalance: newBalance });
  } catch (e) {
    Logger.log("logGiftCertificateUsage 함수 오류: " + e.toString());
    return JSON.stringify({ success: false, message: "사용 내역 기록 중 오류 발생: " + e.toString() });
  }
}

// 테스트용 헬퍼 함수 (선택 사항, Apps Script 편집기에서 실행 가능)
function testSaveCoupon() {
  // 참고: 이제 saveCoupon 등의 함수 반환값은 JSON 문자열로 로깅됩니다.
  Logger.log(saveCoupon({barcode: "TEST12345", expiryDate: "2024-12-31", isGiftCertificate: false, initialBalance: null}));
  Logger.log(saveCoupon({barcode: "GIFT001", expiryDate: "2025-06-30", isGiftCertificate: true, initialBalance: 100}));
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
  // "GIFT001" 바코드의 쿠폰이 존재하고 금액권이며 잔액이 있는지 확인하세요.
  // Logger.log(saveCoupon({barcode: "GIFT001", expiryDate: "2025-06-30", isGiftCertificate: true, initialBalance: 100})); // 미리 저장되어 있지 않다면 주석 해제
  Logger.log(logGiftCertificateUsage("GIFT001", 25));
  // Logger.log(logGiftCertificateUsage("GIFT001", 100)); // 잔액 부족 테스트
  // Logger.log(logGiftCertificateUsage("NONEXISTENT", 10)); // 존재하지 않는 쿠폰 테스트
}
