/**
 * LOTTEON v6.25 dashboard hard display format fix
 *
 * 문제:
 * - 대시보드에 통화기호(₩)가 남아 있음
 * - 매출상품수/미정산 주문건수/예상이익률 같은 비금액 컬럼이 금액 서식으로 오인됨
 * - 일부 데이터 행 높이와 가운데 정렬이 헤더 서식처럼 남아 있음
 *
 * 수정:
 * - 대시보드 전체 사용 범위의 ₩ 텍스트를 숫자로 복구
 * - 헤더/데이터 행 서식을 명확히 분리
 * - 헤더 행: 색상, 굵은 글씨, 가운데 정렬, 줄바꿈
 * - 데이터 행: 흰 배경, 보통 글씨, 동일 행높이, 텍스트 좌측/숫자 우측
 * - 금액 서식은 #,##0, 통화기호 없음
 * - 날짜는 MM/dd
 */

var LOTTEON_PATCH_V625_DASHBOARD_FORMAT_HARD_FIX_LOADED = true;

var __baseApplyDisplayStandardsOnlyFast_v625 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseApplyFastOutputSheetFormatting_v625 = typeof applyFastOutputSheetFormatting_v620_ === 'function' ? applyFastOutputSheetFormatting_v620_ : null;

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v625 ? __baseApplyDisplayStandardsOnlyFast_v625.apply(this, arguments) : null;
  hardFixDashboardDisplay_v625_();
  try { SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n대시보드 통화기호 제거, 헤더/데이터 행 서식, 행높이 정리를 다시 적용했습니다.'); } catch (e) {}
  return result;
};

applyFastOutputSheetFormatting_v620_ = function() {
  var result = __baseApplyFastOutputSheetFormatting_v625 ? __baseApplyFastOutputSheetFormatting_v625.apply(this, arguments) : null;
  hardFixDashboardDisplay_v625_();
  return result;
};

// v6.23 money 판정 보정: 매출상품수/주문건수/미정산건수/이익률 등을 금액으로 오인하지 않게 합니다.
isMoneyHeader_v623_ = function(header) {
  var h = normalizeHeaderText_v625_(header);
  if (!h) return false;
  if (/율|rate|%/.test(h)) return false;
  if (/매출상품|상품수|주문건수|주문$|미정산주문건수|미정산건수|30일초과.*건수|수집수|판매수량|수량|고객수|브랜드수|미판매수|경과일|분석브랜드수|건수|count|qty/i.test(h)) return false;
  return /금액|매출액|순수매출액|총매출액|정산|매입|이익|수수료|부가세|공급가액|공급대가|배송비|원가|비용|amount|price|cost|fee|vat/i.test(h);
};

function hardFixDashboardDisplay_v625_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return { skipped: true };

  var lastRow = Math.min(sheet.getLastRow(), 260);
  var lastCol = Math.min(sheet.getLastColumn(), 24);
  var range = sheet.getRange(1, 1, lastRow, lastCol);
  var values = range.getValues();
  var changed = false;

  for (var r = 0; r < values.length; r++) {
    for (var c = 0; c < values[r].length; c++) {
      var converted = convertCurrencyTextToNumber_v625_(values[r][c]);
      if (converted !== values[r][c]) {
        values[r][c] = converted;
        changed = true;
      }
    }
  }
  if (changed) range.setValues(values);

  // 전체 기본 데이터 서식 초기화
  range
    .setFontFamily('Arial')
    .setFontSize(10)
    .setFontWeight('normal')
    .setFontColor('#000000')
    .setBackground(null)
    .setVerticalAlignment('middle')
    .setHorizontalAlignment('left')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  try { sheet.setFrozenRows(1); } catch (e) {}
  try { sheet.setRowHeights(1, lastRow, 22); } catch (e) {}

  var displayValues = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headerRows = [];
  for (var i = 0; i < displayValues.length; i++) {
    var first = String(displayValues[i][0] || '').trim();
    if (i === 0 || first === '구분') headerRows.push(i + 1);
  }

  // 헤더 행 서식
  headerRows.forEach(function(rowNo) {
    try {
      sheet.getRange(rowNo, 1, 1, lastCol)
        .setBackground('#d9eaf7')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      sheet.setRowHeight(rowNo, 32);
    } catch (e) {}
  });

  // 요약 영역 C열 서식
  applyDashboardSummaryFormats_v625_(sheet, displayValues, lastRow);

  // 표 영역 헤더 기준 서식
  for (var h = 0; h < headerRows.length; h++) {
    var headerRowNo = headerRows[h];
    var endRowNo = h + 1 < headerRows.length ? headerRows[h + 1] - 1 : lastRow;
    applyDashboardTableFormats_v625_(sheet, headerRowNo, endRowNo, lastCol);
  }

  // 빈 행/데이터 행 최종 보정
  for (var rr = 1; rr <= lastRow; rr++) {
    if (headerRows.indexOf(rr) >= 0) continue;
    try {
      sheet.getRange(rr, 1, 1, lastCol)
        .setFontWeight('normal')
        .setBackground(null)
        .setVerticalAlignment('middle');
      sheet.setRowHeight(rr, 22);
    } catch (e) {}
  }

  try {
    log_('patch_v625_dashboard_format_fix', 'rows=' + lastRow + ' cols=' + lastCol + ' converted=' + (changed ? 'Y' : 'N'));
  } catch (e) {}

  return { ok: true };
}

function applyDashboardSummaryFormats_v625_(sheet, values, lastRow) {
  var max = Math.min(lastRow, 35);
  for (var r = 2; r <= max; r++) {
    var item = normalizeHeaderText_v625_(values[r - 1][1]);
    if (!item) continue;
    try {
      var cell = sheet.getRange(r, 3);
      if (isRateHeader_v625_(item)) cell.setNumberFormat('0.00%').setHorizontalAlignment('right');
      else if (isCountHeader_v625_(item)) cell.setNumberFormat('#,##0').setHorizontalAlignment('right');
      else if (isMoneyHeader_v623_(item)) cell.setNumberFormat('#,##0').setHorizontalAlignment('right');
      else if (isDateHeader_v625_(item)) cell.setNumberFormat('MM/dd').setHorizontalAlignment('right');
    } catch (e) {}
  }
}

function applyDashboardTableFormats_v625_(sheet, headerRowNo, endRowNo, lastCol) {
  if (endRowNo <= headerRowNo) return;
  var rowCount = endRowNo - headerRowNo;
  var headers = sheet.getRange(headerRowNo, 1, 1, lastCol).getValues()[0];

  for (var c = 1; c <= headers.length; c++) {
    var h = normalizeHeaderText_v625_(headers[c - 1]);
    if (!h) continue;
    try {
      var dataRange = sheet.getRange(headerRowNo + 1, c, rowCount, 1);
      if (isRateHeader_v625_(h)) dataRange.setNumberFormat('0.00%').setHorizontalAlignment('right');
      else if (isMoneyHeader_v623_(h)) dataRange.setNumberFormat('#,##0').setHorizontalAlignment('right');
      else if (isCountHeader_v625_(h)) dataRange.setNumberFormat('#,##0').setHorizontalAlignment('right');
      else if (isDateHeader_v625_(h)) dataRange.setNumberFormat('MM/dd').setHorizontalAlignment('center');
      else if (/주문번호|상품번호|계정ID|ID$/i.test(h)) dataRange.setNumberFormat('@').setHorizontalAlignment('left');
      else dataRange.setHorizontalAlignment('left');
    } catch (e) {}
  }
}

function convertCurrencyTextToNumber_v625_(value) {
  if (typeof value !== 'string') return value;
  var text = value.trim();
  if (text.indexOf('₩') < 0) return value;
  var cleaned = text.replace(/₩/g, '').replace(/,/g, '').trim();
  if (/^-?\d+(\.\d+)?$/.test(cleaned)) return Number(cleaned);
  return text.replace(/₩/g, '').trim();
}

function normalizeHeaderText_v625_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, ' ').trim();
}

function isRateHeader_v625_(h) {
  h = normalizeHeaderText_v625_(h);
  return /율|rate|%/i.test(h);
}

function isDateHeader_v625_(h) {
  h = normalizeHeaderText_v625_(h);
  if (h === '년' || h === '월' || h === '일') return false;
  return /날짜|일자|주문일|결제일|수집일|생성일|등록일|전송일|갱신일|완료일|최근|최초|date/i.test(h);
}

function isCountHeader_v625_(h) {
  h = normalizeHeaderText_v625_(h);
  if (isRateHeader_v625_(h)) return false;
  if (isMoneyHeader_v623_(h)) return false;
  return /매출상품|상품수|주문건수|주문$|미정산주문건수|미정산건수|30일초과.*건수|수집수|판매수량|수량|고객수|브랜드수|미판매수|경과일|분석브랜드수|건수|count|qty/i.test(h);
}
