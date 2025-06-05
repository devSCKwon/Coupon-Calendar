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
 * 지정된 바코드에 대한 쿠폰 사용 내역을 가져옵니다.
 * @param {string} barcodeForFilter 필터링할 바코드 (선택 사항). 제공되면 해당 바코드의 내역만 반환합니다.
 * @return {string} 사용 내역 객체 배열의 JSON 문자열. 각 객체는 {timestamp: string, barcode: string, amountUsed: number, newBalance: number} 형태를 가집니다.
 *                  오류 발생 시 { error: string } 형태의 객체를 JSON 문자열로 반환할 수 있습니다.
 */
function getUsageHistory(barcodeForFilter) {
  try {
    var ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    var sheet = ss.getSheetByName(USAGE_LOG_SHEET_NAME);

    if (!sheet) {
      Logger.log("getUsageHistory: 사용 기록 시트 ('" + USAGE_LOG_SHEET_NAME + "')를 찾을 수 없습니다.");
      return JSON.stringify([]); // 시트가 없으면 빈 배열 반환
    }

    if (sheet.getLastRow() <= 1) { // 헤더만 있거나 데이터가 없는 경우
      Logger.log("getUsageHistory: 사용 기록 시트에 데이터가 없습니다.");
      return JSON.stringify([]); // 데이터가 없으면 빈 배열 반환
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    var expectedHeaders = ["Timestamp", "Barcode", "Amount Used", "New Balance"];

    // 헤더 인덱스 매핑
    var colIndices = {};
    for (var i = 0; i < expectedHeaders.length; i++) {
      var headerIndex = headers.indexOf(expectedHeaders[i]);
      if (headerIndex === -1) {
        Logger.log("getUsageHistory 오류: 필수 헤더 '" + expectedHeaders[i] + "'를 찾을 수 없습니다.");
        return JSON.stringify({ error: "사용 기록 시트의 헤더 구성이 올바르지 않습니다. 누락된 헤더: " + expectedHeaders[i] });
      }
      colIndices[expectedHeaders[i]] = headerIndex;
    }

    var historyRecords = [];
    var normalizedFilterBarcode = null;
    if (barcodeForFilter && String(barcodeForFilter).trim() !== "") {
      normalizedFilterBarcode = String(barcodeForFilter).replace(/\D/g, "");
    }

    // 데이터는 1번 인덱스부터 시작 (0번 인덱스는 헤더)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      
      // 바코드 필터링 (제공된 경우)
      // 시트의 바코드는 이미 정규화되어 저장된 것으로 가정 (logGiftCertificateUsage에서 처리)
      if (normalizedFilterBarcode) { Logger.log("Filtering comparison: sheetBarcode='" + row[colIndices["Barcode"]] + "' (type: " + typeof row[colIndices["Barcode"]] + "), filterBarcode='" + normalizedFilterBarcode + "' (type: " + typeof normalizedFilterBarcode + ")"); }
      if (normalizedFilterBarcode && String(row[colIndices["Barcode"]]) !== normalizedFilterBarcode) {
        continue;
      }

      var timestamp = row[colIndices["Timestamp"]];
      // 날짜/시간 객체인 경우 ISO 문자열로 변환, 그렇지 않으면 그대로 사용 (이미 문자열일 수 있음)
      var timestampStr = (timestamp instanceof Date) ? timestamp.toISOString() : timestamp;
      
      historyRecords.push({
        timestamp: timestampStr,
        barcode: row[colIndices["Barcode"]],
        amountUsed: parseFloat(row[colIndices["Amount Used"]]), // 숫자형으로 변환
        newBalance: parseFloat(row[colIndices["New Balance"]])  // 숫자형으로 변환
      });
    }

    // Timestamp 기준으로 내림차순 정렬 (최신 기록이 위로)
    historyRecords.sort(function(a, b) {
      var dateA = new Date(a.timestamp);
      var dateB = new Date(b.timestamp);
      return dateB - dateA; // 내림차순
    });
    
    Logger.log("getUsageHistory: " + historyRecords.length + "개의 기록을 반환합니다. 필터 바코드: " + (normalizedFilterBarcode || "없음"));
    return JSON.stringify(historyRecords);

  } catch (e) {
    Logger.log("getUsageHistory 함수 오류: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify({ error: "사용 내역 조회 중 오류 발생: " + e.toString() });
  }
}

/**
 * 새 쿠폰을 스프레드시트에 저장합니다. 바코드 정규화 및 유일성 검사를 포함합니다.
 * @param {Object} couponData 쿠폰 상세 정보 객체
 * @return {string} JSON 문자열 형태의 성공 또는 오류 메시지 객체.
 */
function saveCoupon(couponData) {
  try {
    var originalBarcode = couponData.barcode;

    if (couponData.barcode === null || couponData.barcode === undefined) {
        return JSON.stringify({ success: false, message: "바코드 값이 유효하지 않습니다. 바코드를 입력해주세요." });
    }
    var normalizedBarcode = String(couponData.barcode).replace(/\D/g, "");

    if (!normalizedBarcode || normalizedBarcode.trim() === "") {
      return JSON.stringify({ success: false, message: "유효한 바코드 형식이 아닙니다. 숫자만 포함된 바코드를 입력해주세요. (입력값: " + originalBarcode + ")" });
    }

    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    var headers = ["바코드", "입력일", "만료일", "금액권 여부", "잔액", "최초 금액", "이미지 파일 ID", "이미지 보기 링크", "사용 완료", "만료 상태"];

    if (!sheet) {
      sheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(COUPON_SHEET_NAME);
      sheet.appendRow(headers);
      Logger.log("'" + COUPON_SHEET_NAME + "' 시트가 생성되고 헤더가 추가되었습니다.");
    } else {
      if (sheet.getLastRow() == 0) {
        sheet.appendRow(headers);
        Logger.log("'" + COUPON_SHEET_NAME + "' 시트가 비어있어 헤더가 추가되었습니다.");
      }
    }

    var barcodeColumnIdx = headers.indexOf("바코드");
    if (barcodeColumnIdx === -1) {
        Logger.log("saveCoupon 심각한 오류: '바코드' 열 헤더를 찾을 수 없습니다. headers 배열을 확인하세요.");
        return JSON.stringify({ success: false, message: "내부 서버 오류: 바코드 열 설정을 찾을 수 없습니다." });
    }

    if (sheet.getLastRow() > 1) {
        var existingBarcodesRange = sheet.getRange(2, barcodeColumnIdx + 1, sheet.getLastRow() - 1, 1);
        var existingBarcodes = existingBarcodesRange.getValues();
        for (var i = 0; i < existingBarcodes.length; i++) {
            var existingNormalizedBarcode = String(existingBarcodes[i][0]).replace(/\D/g, "");
            if (existingNormalizedBarcode == normalizedBarcode) {
                Logger.log("중복 바코드 감지: 입력된 바코드(원본): " + originalBarcode + ", 정규화된 바코드: " + normalizedBarcode + ", 시트의 기존 바코드: " + existingBarcodes[i][0]);
                return JSON.stringify({ success: false, message: "이미 등록된 바코드 번호입니다: " + originalBarcode + " (처리된 값: " + normalizedBarcode + ")" });
            }
        }
    }

    var entryDate = new Date();
    var balance = couponData.isGiftCertificate ? parseFloat(couponData.initialBalance) : null;
    var originalAmount = balance;

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
      normalizedBarcode,
      entryDate,
      expiryDateObj,
      couponData.isGiftCertificate,
      balance,
      originalAmount,
      couponData.imageFileId || null,
      couponData.imageWebViewLink || null,
      false,
      "사용가능"
    ]);
    Logger.log("쿠폰 저장 성공: 원본 바코드=" + originalBarcode + ", 정규화된 바코드=" + normalizedBarcode);
    return JSON.stringify({ success: true, message: "쿠폰이 성공적으로 저장되었습니다: " + originalBarcode + " (처리된 바코드: " + normalizedBarcode + ")" });
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
              var headers = data[0];
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
                if (data[i] && barcodeIdx !==-1 && data[i][barcodeIdx] != "") {
                  couponsData.push({
                    barcode: data[i][barcodeIdx],
                    entryDate: entryDateIdx !==-1 ? data[i][entryDateIdx] : null,
                    expiryDate: expiryDateIdx !==-1 ? data[i][expiryDateIdx] : null,
                    isGiftCertificate: isGiftCertIdx !==-1 ? data[i][isGiftCertIdx] : false,
                    balance: balanceIdx !==-1 ? data[i][balanceIdx] : null,
                    originalAmount: originalAmountIdx !==-1 ? data[i][originalAmountIdx] : null,
                    imageFileId: imageFileIdIdx !== -1 ? (data[i][imageFileIdIdx] || null) : null,
                    imageWebViewLink: imageWebViewLinkIdx !== -1 ? (data[i][imageWebViewLinkIdx] || null) : null,
                    isUsed: isUsedIdx !== -1 ? (data[i][isUsedIdx]) : false,
                    expiryStatus: expiryStatusIdx !== -1 ? (data[i][expiryStatusIdx]) : "정보없음"
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
 * @param {string} barcode 처리할 쿠폰의 바코드 (사용자 입력값).
 * @return {string} JSON 문자열 형태의 결과 객체.
 */
function markCouponAsUsed(barcode) {
  try {
    if (barcode === null || barcode === undefined) {
        return JSON.stringify({ success: false, message: "바코드 값이 유효하지 않습니다." });
    }
    var originalInputBarcode = barcode;
    var normalizedBarcode = String(barcode).replace(/\D/g, "");
    if (!normalizedBarcode) {
      return JSON.stringify({ success: false, message: "처리할 유효한 바코드가 없습니다 (정규화 후 빈 값). 원본: " + originalInputBarcode });
    }

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
      // 시트에 저장된 바코드 (data[i][barcodeColIdx])도 정규화된 값이라고 가정하고 비교
      if (data[i][barcodeColIdx] == normalizedBarcode) {
        if (String(data[i][isGiftCertColIdx]).toLowerCase() === 'true') {
          return JSON.stringify({ success: false, message: "금액권은 이 기능을 사용할 수 없습니다. '사용 기록' 기능을 이용해주세요. (입력 바코드: " + originalInputBarcode + ")" });
        }

        if (String(data[i][isUsedColIdx]).toLowerCase() === 'true') {
            return JSON.stringify({ success: false, message: "이미 사용 완료 처리된 쿠폰입니다. (입력 바코드: " + originalInputBarcode + ")" });
        }

        sheet.getRange(i + 1, isUsedColIdx + 1).setValue(true);
        sheet.getRange(i + 1, expiryStatusColIdx + 1).setValue("사용완료");

        Logger.log("일반 쿠폰 사용 완료: " + normalizedBarcode + " (원본 입력: " + originalInputBarcode + ")");
        return JSON.stringify({ success: true, message: "'" + originalInputBarcode + "' 쿠폰이 사용 완료 처리되었습니다." });
      }
    }
    return JSON.stringify({ success: false, message: "바코드 '" + originalInputBarcode + "'에 해당하는 쿠폰을 찾을 수 없습니다." });
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
 * @param {String} barcode 금액권의 바코드 (사용자 입력값).
 * @param {Number} amountUsed 사용한 금액.
 * @return {String} 성공 여부, 메시지, 새 잔액 등을 포함한 객체의 JSON 문자열.
 */
function logGiftCertificateUsage(barcode, amountUsed) {
  try {
    if (barcode === null || barcode === undefined) {
        return JSON.stringify({ success: false, message: "바코드 값이 유효하지 않습니다." });
    }
    var originalInputBarcode = barcode;
    var normalizedBarcode = String(barcode).replace(/\D/g, "");
    if (!normalizedBarcode) {
       return JSON.stringify({ success: false, message: "처리할 유효한 바코드가 없습니다 (정규화 후 빈 값). 원본: " + originalInputBarcode });
    }

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
      // 시트에 저장된 바코드 (data[i][barcodeColIdx])도 정규화된 값이라고 가정하고 비교
      if (data[i][barcodeColIdx] == normalizedBarcode) {
        couponRow = i;
        isGift = data[i][isGiftCertColIdx];
        currentBalance = parseFloat(data[i][balanceColIdx]);
        break;
      }
    }

    if (couponRow === -1) {
      return JSON.stringify({ success: false, message: "오류: 바코드 '" + originalInputBarcode + "'에 해당하는 쿠폰을 찾을 수 없습니다." });
    }
    if (!(String(isGift).toLowerCase() === 'true')) {
      return JSON.stringify({ success: false, message: "오류: 쿠폰 '" + originalInputBarcode + "'은 금액권이 아닙니다." });
    }
    if (isNaN(currentBalance) || currentBalance < amountUsed) {
      return JSON.stringify({ success: false, message: "오류: 잔액이 부족합니다. 현재 잔액 (" + originalInputBarcode + "): " + (isNaN(currentBalance) ? 0 : currentBalance.toFixed(2)) });
    }

    var newBalance = currentBalance - amountUsed;
    couponSheet.getRange(couponRow + 1, balanceColIdx + 1).setValue(newBalance);

    if (newBalance === 0) {
        couponSheet.getRange(couponRow + 1, expiryStatusColIdx + 1).setValue("잔액없음");
        Logger.log("금액권 잔액 없음 처리: " + normalizedBarcode + " (원본 입력: " + originalInputBarcode + ")");
    }

    var logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(USAGE_LOG_SHEET_NAME);
    if (!logSheet) {
      logSheet = SpreadsheetApp.openById(SPREADSHEET_ID).insertSheet(USAGE_LOG_SHEET_NAME);
      logSheet.appendRow(["Timestamp", "Barcode", "Amount Used", "New Balance"]);
    }
    logSheet.appendRow([new Date(), normalizedBarcode, amountUsed, newBalance]); // 로그에는 정규화된 바코드 사용

    var message = "사용 내역이 성공적으로 기록되었습니다. " + originalInputBarcode + "의 새 잔액: " + newBalance.toFixed(2);
    if (newBalance === 0) {
        message += " (잔액 소진)";
    }
    return JSON.stringify({ success: true, message: message, newBalance: newBalance });
  } catch (e) {
    Logger.log("logGiftCertificateUsage 함수 오류: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify({ success: false, message: "사용 내역 기록 중 오류 발생: " + e.toString() });
  }
}

/**
 * 대시보드 통계 정보를 계산하여 반환합니다.
 * @return {String} 통계 정보를 담은 객체의 JSON 문자열.
 *                   { totalCoupons: number, availableCoupons: number, expiringSoonCoupons: number, totalGiftCertificateBalance: number }
 */
function getDashboardStats() {
  try {
    var sheet = SpreadsheetApp.openById(SPREADSHEET_ID).getSheetByName(COUPON_SHEET_NAME);
    if (!sheet || sheet.getLastRow() <= 1) { // 헤더만 있거나 시트가 없는 경우
      Logger.log("getDashboardStats: 쿠폰 시트가 비어있거나 찾을 수 없습니다.");
      return JSON.stringify({ totalCoupons: 0, availableCoupons: 0, expiringSoonCoupons: 0, totalGiftCertificateBalance: 0 });
    }

    var data = sheet.getDataRange().getValues();
    var headers = data[0];
    
    // 헤더 인덱스 찾기 - 실제 열 이름과 일치해야 함
    var expiryDateIdx = headers.indexOf("만료일");
    var isGiftCertIdx = headers.indexOf("금액권 여부");
    var balanceIdx = headers.indexOf("잔액");
    var isUsedIdx = headers.indexOf("사용 완료"); // "사용 완료" 열
    var expiryStatusIdx = headers.indexOf("만료 상태"); // "만료 상태" 열

    if (expiryDateIdx === -1 || isGiftCertIdx === -1 || balanceIdx === -1 || isUsedIdx === -1 || expiryStatusIdx === -1) {
      Logger.log("getDashboardStats 오류: 필요한 열 헤더 중 일부를 찾을 수 없습니다. (" + 
                 "만료일: " + expiryDateIdx + ", 금액권 여부: " + isGiftCertIdx + ", 잔액: " + balanceIdx + 
                 ", 사용 완료: " + isUsedIdx + ", 만료 상태: " + expiryStatusIdx + ")");
      // 오류 발생 시 기본값 반환 또는 오류 객체 반환 고려
      return JSON.stringify({ error: "시트 헤더 구성이 올바르지 않습니다.", totalCoupons: 0, availableCoupons: 0, expiringSoonCoupons: 0, totalGiftCertificateBalance: 0 });
    }

    var totalCoupons = 0;
    var availableCoupons = 0;
    var expiringSoonCoupons = 0;
    var totalGiftCertificateBalance = 0;

    var today = new Date();
    today.setHours(0, 0, 0, 0); // 날짜 비교를 위해 시간 부분 초기화

    var sevenDaysFromNow = new Date();
    sevenDaysFromNow.setDate(today.getDate() + 7);
    sevenDaysFromNow.setHours(0,0,0,0);

    // 데이터는 1번 인덱스부터 시작 (0번 인덱스는 헤더)
    for (var i = 1; i < data.length; i++) {
      var row = data[i];
      if (!row || row[0] === "") continue; // 빈 행이나 바코드 없는 행은 건너뛰기

      totalCoupons++;

      var expiryDateStr = row[expiryDateIdx];
      var isGiftCertificate = row[isGiftCertIdx]; // boolean 또는 문자열 'true'/'false'
      var balance = parseFloat(row[balanceIdx]);
      var isUsed = row[isUsedIdx]; // boolean 또는 문자열 'true'/'false'
      var currentExpiryStatus = row[expiryStatusIdx]; // 문자열, 예: "사용가능", "만료됨", "사용완료"

      // isUsed와 isGiftCertificate가 문자열일 수 있으므로 boolean으로 변환
      var isUsedBool = (String(isUsed).toLowerCase() === 'true');
      var isGiftCertBool = (String(isGiftCertificate).toLowerCase() === 'true');
      
      var expiryDate = new Date(expiryDateStr);
      expiryDate.setHours(0,0,0,0); // 날짜 비교를 위해 시간 부분 초기화

      var isActuallyAvailable = false;

      // 1. 만료되지 않았는지 (만료일 > 오늘)
      if (expiryDate > today) {
        // 2. "사용 완료" 상태가 아닌지
        if (!isUsedBool) {
          // 3. "만료 상태" 필드가 "사용가능" 또는 유사한 긍정적 상태인지 (추가적인 안전장치)
          //    (만료 상태가 "만료됨"으로 이미 표시된 경우는 제외)
          if (currentExpiryStatus !== "만료됨" && currentExpiryStatus !== "사용완료" && currentExpiryStatus !== "잔액없음") {
             // 4. 금액권의 경우 잔액이 0보다 큰지
            if (isGiftCertBool) {
              if (!isNaN(balance) && balance > 0) {
                isActuallyAvailable = true;
              }
            } else { // 일반 쿠폰
              isActuallyAvailable = true;
            }
          }
        }
      }


      if (isActuallyAvailable) {
        availableCoupons++;

        // 만료 예정 쿠폰 (사용 가능하면서 7일 이내 만료)
        if (expiryDate <= sevenDaysFromNow) {
          expiringSoonCoupons++;
        }

        // 사용 가능한 금액권의 총 잔액 합산
        if (isGiftCertBool) {
          if (!isNaN(balance)) { // 이미 balance > 0 조건은 isActuallyAvailable에서 확인됨
             totalGiftCertificateBalance += balance;
          }
        }
      }
    }

    Logger.log("대시보드 통계: 총 쿠폰=" + totalCoupons + ", 사용 가능=" + availableCoupons + 
               ", 만료 예정=" + expiringSoonCoupons + ", 총 잔액=" + totalGiftCertificateBalance);

    return JSON.stringify({
      totalCoupons: totalCoupons,
      availableCoupons: availableCoupons,
      expiringSoonCoupons: expiringSoonCoupons,
      totalGiftCertificateBalance: totalGiftCertificateBalance.toFixed(2) // 소수점 2자리로 포맷
    });

  } catch (e) {
    Logger.log("getDashboardStats 함수 오류: " + e.toString() + " 스택: " + e.stack);
    return JSON.stringify({ 
      error: "통계 계산 중 오류 발생: " + e.toString(),
      totalCoupons: 0, 
      availableCoupons: 0, 
      expiringSoonCoupons: 0, 
      totalGiftCertificateBalance: 0 
    });
  }
}


// 테스트용 헬퍼 함수
function testSaveCoupon() {
  Logger.log(saveCoupon({barcode: "NEW_COUPON_001", expiryDate: "2027-12-31", isGiftCertificate: false, initialBalance: null}));
  Logger.log(saveCoupon({barcode: "GIFT_COUPON_003", expiryDate: "2027-06-30", isGiftCertificate: true, initialBalance: 200}));
}

function testMarkCouponAsUsed() {
    Logger.log(markCouponAsUsed("NEW_COUPON_001"));
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
  Logger.log(logGiftCertificateUsage("GIFT_COUPON_003", 50));
  Logger.log(logGiftCertificateUsage("GIFT_COUPON_003", 150));
}

function testUploadImage() {
  var testBase64 = "iVBORw0KGgoAAAANSUhEUgAAAAEAAAABCAQAAAC1HAwCAAAAC0lEQVR42mNkYAAAAAYAAjCB0C8AAAAASUVORK5CYII=";
  var testName = "testImage.png";
  var testMime = "image/png";
  Logger.log(uploadImageToDrive(testBase64, testName, testMime));
}
