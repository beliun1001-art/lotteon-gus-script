/**
 * LOTTEON v6.50 identifier no-comma format fix
 *
 * 문제:
 * - 주문번호, 상품번호처럼 금액이 아닌 식별자 컬럼에 1,000단위 쉼표가 적용될 수 있었습니다.
 * - 특히 부가세_신고자료의 주문번호/상품번호는 금액이 아니므로 #,##0 적용 대상이 아닙니다.
 *
 * 수정:
 * - 금액 서식 적용 전후로 식별자/번호 컬럼을 일반 텍스트(@)로 강제합니다.
 * - 식별자 컬럼의 기존 쉼표 표시를 제거합니다.
 * - 금액 판별에서 주문번호/상품번호/마켓주문번호/마켓상품번호 등은 확실히 제외합니다.
 */

var LOTTEON_PATCH_V650_IDENTIFIER_NO_COMMA_FORMAT_LOADED = true;

var __baseApplyAmountThousandsFormat_v650 = typeof applyAmountThousandsFormat_v649 === 'function' ? applyAmountThousandsFormat_v649 : null;

applyAmountThousandsFormat_v649 = function() {
  var started = Date.now();
  var before = applyIdentifierNoCommaFormatAllSheets_v650_();
  var result = __baseApplyAmountThousandsFormat_v650 ? __baseApplyAmountThousandsFormat_v650.apply(this, arguments) : applySplitDisplayFormatAllSheets_v649_('amount');
  var after = applyIdentifierNoCommaFormatAllSheets_v650_();
  try {
    SpreadsheetApp.getUi().alert(
      '금액 1,000단위 쉼표 적용 완료\n\n' +
      '금액 컬럼 처리: ' + (result && result.columns ? result.columns : 0) + '\n' +
      '번호/ID 컬럼 쉼표 제거: ' + (before.columns + after.columns) + '\n' +
      '소요초: ' + Math.round((Date.now() - started) / 1000) + '\n\n' +
      '주문번호/상품번호/마켓주문번호/마켓상품번호는 금액이 아니므로 쉼표를 적용하지 않습니다.'
    );
  } catch (e) {}
  return result;
};

function applyIdentifierNoCommaFormat_v650() {
  var started = Date.now();
  var result = applyIdentifierNoCommaFormatAllSheets_v650_();
  try {
    SpreadsheetApp.getUi().alert(
      '번호/ID 컬럼 쉼표 제거 완료\n\n' +
      '처리 시트: ' + result.sheets + '\n' +
      '처리 컬럼: ' + result.columns + '\n' +
      '소요초: ' + Math.round((Date.now() - started) / 1000) + '\n\n' +
      '서식: 일반 텍스트(@)\n대상: 주문번호, 상품번호, 마켓주문번호, 마켓상품번호, URL, ID 등'
    );
  } catch (e) {}
  return result;
}

function applyIdentifierNoCommaFormatAllSheets_v650_() {
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var totalSheets = 0;
  var totalColumns = 0;
  sheets.forEach(function(sheet) {
    var result = applyIdentifierNoCommaFormatToSheet_v650_(sheet);
    if (result.columns > 0) {
      totalSheets++;
      totalColumns += result.columns;
    }
  });
  try { log_('patch_v650_identifier_no_comma', 'sheets=' + totalSheets + ' columns=' + totalColumns); } catch (e) {}
  return { ok: true, sheets: totalSheets, columns: totalColumns };
}

function applyIdentifierNoCommaFormatToSheet_v650_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return { columns: 0 };
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var headerScanRows = Math.min(lastRow, 120);
  var values = sheet.getRange(1, 1, headerScanRows, lastCol).getValues();
  var processed = 0;

  for (var r = 1; r <= headerScanRows; r++) {
    var row = values[r - 1];
    for (var c = 1; c <= lastCol; c++) {
      var header = normalizeIdentifierHeader_v650_(row[c - 1]);
      if (!isIdentifierHeader_v650_(header)) continue;
      var endRow = findIdentifierSectionEndRow_v650_(values, r + 1, headerScanRows, lastRow);
      var rowCount = Math.max(0, endRow - r - 1);
      if (rowCount <= 0) continue;
      var range = sheet.getRange(r + 1, c, rowCount, 1);
      removeCommaAndTextFormat_v650_(range);
      processed++;
    }
  }
  return { columns: processed };
}

function findIdentifierSectionEndRow_v650_(headerArea, startRowNo, maxScanRowNo, lastRow) {
  for (var r = startRowNo; r <= maxScanRowNo; r++) {
    var row = headerArea[r - 1];
    var idHits = 0;
    for (var c = 0; c < row.length; c++) {
      if (isIdentifierHeader_v650_(normalizeIdentifierHeader_v650_(row[c]))) idHits++;
    }
    if (idHits > 0) return r;
  }
  return lastRow + 1;
}

function removeCommaAndTextFormat_v650_(range) {
  var values = range.getValues();
  var changed = false;
  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var v = values[r][c];
      if (v === '' || v == null) continue;
      var s = String(v).trim();
      if (!s) continue;
      // 번호/ID 컬럼은 표시용 식별자이므로 쉼표를 제거하고 텍스트로 둡니다.
      var cleaned = s.replace(/,/g, '');
      if (cleaned !== s || typeof v === 'number') {
        values[r][c] = cleaned;
        changed = true;
      }
    }
  }
  if (changed) range.setValues(values);
  range.setNumberFormat('@').setHorizontalAlignment('left');
}

function normalizeIdentifierHeader_v650_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, '').trim();
}

function isIdentifierHeader_v650_(header) {
  header = normalizeIdentifierHeader_v650_(header);
  if (!header) return false;
  // 금액/율/수량성은 식별자 아님
  if (/금액|매출|정산|매입|이익|수수료|부가세|공급|율|비율|수량|건수|상품수|고객수|주문수/.test(header)) return false;
  return /주문번호|상품번호|마켓주문번호|마켓상품번호|상품코드|품목코드|품번|주문ID|상품ID|vendorItemId|URL|링크|계정ID|쿠팡계정ID|사업자등록번호|운송장|송장|전화번호|휴대폰|바코드|옵션ID/i.test(header);
}

// v6.49의 금액 헤더 판별을 더 보수적으로 덮어씁니다.
isAmountHeader_v649_ = function(header) {
  header = String(header == null ? '' : header).replace(/\n/g, '').replace(/\s+/g, '').trim();
  if (!header) return false;
  if (isIdentifierHeader_v650_(header)) return false;
  if (/율|비율|rate|%/i.test(header)) return false;
  if (/상품수|건수|주문수|고객수|수량|행수|일수|경과|날짜|일자|번호|ID|아이디|URL|링크|코드|품번|바코드|송장|운송장/.test(header)) return false;
  return /금액|매출|정산|매입|이익|수수료|부가세|공급가액|공급대가|구매가격|판매금액|결제금액|순수매출|총매출|예상정산|납부예상/.test(header);
};
