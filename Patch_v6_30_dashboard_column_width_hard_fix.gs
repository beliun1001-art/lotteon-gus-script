/**
 * LOTTEON v6.30 dashboard column width hard fix
 *
 * 문제:
 * - 대시보드 시트가 생성/갱신된 뒤 열너비 자동조정이 실제 보기 상태에 반영되지 않았습니다.
 * - 숫자 컬럼이 좌측 정렬로 남아 있어 금액/건수 가독성이 떨어졌습니다.
 * - 대시보드 표에서 대표검색필터명, 계정ID 컬럼이 더 이상 필요하지 않습니다.
 *
 * 수정:
 * - 대시보드 생성 자체에서 대표검색필터명/계정ID 컬럼을 제거합니다.
 * - 대시보드 전용 열너비 강제 정리 함수를 추가합니다.
 * - 대시보드 생성/빠른 갱신/표시서식 정리/열너비 자동조정 후 항상 대시보드 전용 폭을 재적용합니다.
 * - 숫자/금액/비율/건수는 우측 정렬로 고정합니다.
 * - 텍스트 컬럼만 좌측 정렬합니다.
 */

var LOTTEON_PATCH_V630_DASHBOARD_COLUMN_WIDTH_HARD_FIX_LOADED = true;

var __baseHardFixDashboardDisplay_v630 = typeof hardFixDashboardDisplay_v625_ === 'function' ? hardFixDashboardDisplay_v625_ : null;
var __baseApplyDisplayStandardsOnlyFast_v630 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseApplyFastOutputSheetFormatting_v630 = typeof applyFastOutputSheetFormatting_v620_ === 'function' ? applyFastOutputSheetFormatting_v620_ : null;
var __baseRunColumnWidthStep_v630 = typeof runColumnWidthAutoAdjustStep_v623 === 'function' ? runColumnWidthAutoAdjustStep_v623 : null;

rebuildDashboardSingleSource_v628_ = function(salesAgg, filterAgg) {
  return rebuildDashboardNoFilterAccount_v630_(salesAgg, filterAgg);
};

hardFixDashboardDisplay_v625_ = function() {
  var result = __baseHardFixDashboardDisplay_v630 ? __baseHardFixDashboardDisplay_v630.apply(this, arguments) : null;
  hardFixDashboardDisplay_v630_();
  return result || { ok: true };
};

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v630 ? __baseApplyDisplayStandardsOnlyFast_v630.apply(this, arguments) : null;
  hardFixDashboardDisplay_v630_();
  try { SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n대시보드 전용 열너비, 숫자 우측정렬, 불필요 컬럼 제거 기준을 다시 적용했습니다.'); } catch (e) {}
  return result;
};

applyFastOutputSheetFormatting_v620_ = function() {
  var result = __baseApplyFastOutputSheetFormatting_v630 ? __baseApplyFastOutputSheetFormatting_v630.apply(this, arguments) : null;
  hardFixDashboardDisplay_v630_();
  return result || { ok: true };
};

runColumnWidthAutoAdjustStep_v623 = function() {
  var result = __baseRunColumnWidthStep_v630 ? __baseRunColumnWidthStep_v630.apply(this, arguments) : null;
  hardFixDashboardDisplay_v630_();
  return result;
};

function rebuildDashboardNoFilterAccount_v630_(salesAgg, filterAgg) {
  salesAgg = salesAgg || (typeof buildSingleSourceSalesAgg_v628_ === 'function' ? buildSingleSourceSalesAgg_v628_() : null);
  filterAgg = filterAgg || (typeof aggregateFiltersByBrand_v611_ === 'function' ? aggregateFiltersByBrand_v611_() : {});
  if (!salesAgg) throw new Error('대시보드 생성용 salesAgg를 만들 수 없습니다.');

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD) || ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  var rows = [];
  var nowText = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  var supply = Math.round((salesAgg.totalNetSales || 0) / 1.1);
  var vat = (salesAgg.totalNetSales || 0) - supply;
  var purchaseVat = (salesAgg.totalPurchase || 0) - Math.round((salesAgg.totalPurchase || 0) / 1.1);
  var payableVat = vat - purchaseVat;

  function summary(item, value, memo) {
    rows.push(['요약', item, value, '', '', '', '', '', '', '', '', '', '', '', '', memo || '']);
  }

  summary('갱신기준', 'v6.30 단일 원천 + 대시보드 간소화', '대표검색필터명/계정ID 제거');
  summary('갱신시각', nowText, 'Asia/Seoul');
  summary('원본 전체 행수', salesAgg.rawRowCount || 0, '매출데이터_붙여넣기');
  summary('분석대상 행수', (salesAgg.detailRows || []).length, '취소/반품 제외 + lotteon + 상품번호 + 브랜드');
  summary('분석브랜드수', Object.keys(salesAgg.byBrand || {}).length, '브랜드 기준');
  summary('수집상품수', (filterAgg && filterAgg.totalCollectCount) || 0, '필터별_상품수 API_totalCount 총합');
  summary('매출상품수', salesAgg.totalProducts || 0, '분석대상 상품수');
  summary('주문건수', salesAgg.totalOrders || 0, '분석대상 주문수');
  summary('미정산 주문건수', salesAgg.unsettledOrderCount || 0, '정산예정금액 공란/0');
  summary('30일초과 미정산 주문건수', salesAgg.overdueUnsettledOrderCount || 0, '주문일 기준 30일 초과');
  summary('총매출액', salesAgg.totalGrossSales || 0, '원본 전체');
  summary('취소/반품매출', salesAgg.totalCancelSales || 0, '상태값 제외');
  summary('분석제외매출', salesAgg.totalAnalysisExcludedSales || 0, '취소 외 분석대상 제외');
  summary('제외매출합계', salesAgg.totalExcludedSales || 0, '취소/반품 + 분석제외');
  summary('순수매출액', salesAgg.totalNetSales || 0, '분석대상 기준');
  summary('실제정산금액', salesAgg.totalActualSettlement || 0, '정산예정금액 입력 완료 주문');
  summary('예상정산금액(미정산)', salesAgg.totalEstimatedSettlement || 0, '미정산 순수매출액 × 0.901');
  summary('정산기준금액', salesAgg.totalSettlementBasis || 0, '실제 + 예상');
  summary('마켓수수료/비용', salesAgg.totalMarketFee || 0, '순수매출액 - 정산기준금액');
  summary('마켓수수료율', salesAgg.totalMarketFeeRate || 0, '마켓수수료 ÷ 순수매출액');
  summary('매입금액', salesAgg.totalPurchase || 0, 'AC열 구매가격 직접 합계');
  summary('예상이익', salesAgg.totalEstimatedProfit || 0, '정산기준금액 - 매입금액');
  summary('예상이익률', salesAgg.totalEstimatedProfitRate || 0, '예상이익 ÷ 순수매출액');
  summary('부가세 신고 공급대가', salesAgg.totalNetSales || 0, '순수매출액 기준');
  summary('부가세 신고 공급가액', supply, '공급대가 ÷ 1.1');
  summary('매출부가세', vat, '공급대가 - 공급가액');
  summary('매입부가세', purchaseVat, '매입금액 부가세 포함 기준');
  summary('납부예상부가세', payableVat, '매출부가세 - 매입부가세');
  rows.push(['','','','','','','','','','','','','','','','']);

  var brandRows = typeof buildDashboardBrandRows_v628_ === 'function' ? buildDashboardBrandRows_v628_(salesAgg, filterAgg) : buildDashboardBrandRowsFallback_v630_(salesAgg);

  rows.push(['구분','브랜드명','수집수','매출상품','주문','순수매출액','실제정산금액','예상정산금액(미정산)','정산기준금액','매입금액','예상이익','예상이익률','미정산건수','30일초과건수','메모','']);
  brandRows.slice().sort(function(a,b){ return (b.netSalesAmount || 0) - (a.netSalesAmount || 0); }).slice(0, 60).forEach(function(b) {
    rows.push(['브랜드TOP', b.brand, b.collectCount || 0, b.salesProductCount || 0, b.orderCount || 0, b.netSalesAmount || 0, b.actualSettlementAmount || 0, b.estimatedSettlementAmount || 0, b.settlementBasisAmount || 0, b.purchaseAmount || 0, b.estimatedProfitAmount || 0, b.estimatedProfitRate || 0, b.unsettledOrderCount || 0, b.overdueUnsettledOrderCount || 0, '브랜드 기준', '']);
  });
  rows.push(['','','','','','','','','','','','','','','','']);
  rows.push(['구분','브랜드명','수집수','매출상품','주문','순수매출액','정산기준금액','예상이익률','미정산건수','30일초과건수','다음작업','최근수집일','미판매수','메모','','']);
  brandRows.slice().sort(function(a,b){ return (b.unsoldCount || 0) - (a.unsoldCount || 0); }).slice(0, 80).forEach(function(b) {
    rows.push(['상품갈이', b.brand, b.collectCount || 0, b.salesProductCount || 0, b.orderCount || 0, b.netSalesAmount || 0, b.settlementBasisAmount || 0, b.estimatedProfitRate || 0, b.unsettledOrderCount || 0, b.overdueUnsettledOrderCount || 0, decideNextRotationSafe_v630_(b), formatShortDateSafe_v630_(b.latestRecentDate), Math.max(0, b.unsoldCount || 0), 'v6.30 브랜드 기준', '', '']);
  });

  replaceSheetValuesSafe_v628_(sheet, ['구분','항목','값1','값2','값3','값4','값5','값6','값7','값8','값9','값10','값11','값12','값13','메모'], rows);
  hardFixDashboardDisplay_v630_();
  return { rows: rows.length };
}

function buildDashboardBrandRowsFallback_v630_(salesAgg) {
  return Object.keys(salesAgg.byBrand || {}).map(function(k) {
    var b = salesAgg.byBrand[k];
    return {
      brand: b.brand,
      collectCount: 0,
      salesProductCount: b.salesProductCount || 0,
      orderCount: b.orderCount || 0,
      netSalesAmount: b.netSalesAmount || 0,
      actualSettlementAmount: b.actualSettlementAmount || 0,
      estimatedSettlementAmount: b.estimatedSettlementAmount || 0,
      settlementBasisAmount: b.settlementBasisAmount || 0,
      purchaseAmount: b.purchaseAmount || 0,
      estimatedProfitAmount: b.estimatedProfitAmount || 0,
      estimatedProfitRate: b.estimatedProfitRate || 0,
      unsettledOrderCount: b.unsettledOrderCount || 0,
      overdueUnsettledOrderCount: b.overdueUnsettledOrderCount || 0,
      unsoldCount: 0,
      latestRecentDate: ''
    };
  });
}

function hardFixDashboardDisplay_v630_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD);
  if (!sheet || sheet.getLastRow() < 1) return { skipped: true };
  var lastRow = Math.min(sheet.getLastRow(), 260);
  var lastCol = Math.min(sheet.getLastColumn(), 16);
  try {
    applyDashboardColumnWidths_v630_(sheet);
    applyDashboardRowsAndAlignment_v630_(sheet, lastRow, lastCol);
    applyDashboardNumberFormats_v630_(sheet, lastRow, lastCol);
    try { sheet.setFrozenRows(1); } catch (e) {}
    try { log_('patch_v630_dashboard_width_fix', 'rows=' + lastRow + ' cols=' + lastCol); } catch (e) {}
  } catch (e) {
    try { log_('patch_v630_dashboard_width_fix_error', String(e && e.message ? e.message : e)); } catch (ignore) {}
  }
  return { ok: true };
}

function applyDashboardColumnWidths_v630_(sheet) {
  var widths = {
    1: 62,
    2: 142,
    3: 86,
    4: 82,
    5: 76,
    6: 116,
    7: 118,
    8: 136,
    9: 118,
    10: 116,
    11: 112,
    12: 88,
    13: 92,
    14: 92,
    15: 132,
    16: 230
  };
  Object.keys(widths).forEach(function(k) {
    try { sheet.setColumnWidth(Number(k), widths[k]); } catch (e) {}
  });
}

function applyDashboardRowsAndAlignment_v630_(sheet, lastRow, lastCol) {
  if (lastRow < 1 || lastCol < 1) return;
  var range = sheet.getRange(1, 1, lastRow, lastCol);
  range
    .setFontFamily('Arial')
    .setFontSize(10)
    .setVerticalAlignment('middle')
    .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);

  try { sheet.setRowHeights(1, lastRow, 22); } catch (e) {}

  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headerRows = [];
  for (var r = 0; r < values.length; r++) {
    var first = String(values[r][0] || '').replace(/\n/g, '').trim();
    if (r === 0 || first === '구분') headerRows.push(r + 1);
  }

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, lastCol)
      .setFontWeight('normal')
      .setBackground(null)
      .setHorizontalAlignment('right')
      .setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
  }

  headerRows.forEach(function(rowNo) {
    try {
      sheet.getRange(rowNo, 1, 1, lastCol)
        .setBackground('#d9eaf7')
        .setFontWeight('bold')
        .setHorizontalAlignment('center')
        .setVerticalAlignment('middle')
        .setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
      sheet.setRowHeight(rowNo, 34);
    } catch (e) {}
  });

  // 텍스트 컬럼만 좌측 정렬. 숫자는 기본 우측 정렬 유지.
  [1,2,15,16].forEach(function(col) {
    if (col <= lastCol && lastRow > 1) {
      try { sheet.getRange(2, col, lastRow - 1, 1).setHorizontalAlignment('left'); } catch (e) {}
    }
  });
}

function applyDashboardNumberFormats_v630_(sheet, lastRow, lastCol) {
  if (lastRow < 2) return;
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();

  var summaryMax = Math.min(30, lastRow);
  for (var r = 2; r <= summaryMax; r++) {
    var item = normalizeDashboardHeader_v630_(values[r - 1][1]);
    if (!item) continue;
    try {
      var cell = sheet.getRange(r, 3).setHorizontalAlignment('right');
      if (/율/.test(item)) cell.setNumberFormat('0.00%');
      else if (/일자|날짜|시각|최근|최초/.test(item)) cell.setNumberFormat('@').setHorizontalAlignment('left');
      else if (/행수|상품수|건수|브랜드수|주문수|수집수/.test(item)) cell.setNumberFormat('#,##0');
      else if (/매출|정산|매입|이익|수수료|부가세|공급/.test(item)) cell.setNumberFormat('#,##0');
    } catch (e) {}
  }

  for (var hr = 1; hr <= lastRow; hr++) {
    var first = String(values[hr - 1][0] || '').replace(/\n/g, '').trim();
    if (!(hr === 1 || first === '구분')) continue;
    var nextHeader = findNextDashboardHeaderRow_v630_(values, hr + 1);
    var endRow = nextHeader ? nextHeader - 1 : lastRow;
    if (endRow <= hr) continue;
    var rowCount = endRow - hr;
    for (var c = 1; c <= lastCol; c++) {
      var h = normalizeDashboardHeader_v630_(values[hr - 1][c - 1]);
      if (!h) continue;
      try {
        var rg = sheet.getRange(hr + 1, c, rowCount, 1);
        if (/율/.test(h)) rg.setNumberFormat('0.00%').setHorizontalAlignment('right');
        else if (/날짜|일자|최근수집일/.test(h)) rg.setNumberFormat('@').setHorizontalAlignment('center');
        else if (/금액|매출|정산|매입|이익|수수료|부가세|공급/.test(h) && !/상품수|건수|수량/.test(h)) rg.setNumberFormat('#,##0').setHorizontalAlignment('right');
        else if (/수집수|상품|주문|건수|수량|미판매수|30일초과/.test(h) && !/주문번호/.test(h)) rg.setNumberFormat('#,##0').setHorizontalAlignment('right');
        else rg.setNumberFormat('@').setHorizontalAlignment('left');
      } catch (e) {}
    }
  }
}

function findNextDashboardHeaderRow_v630_(values, fromRowNo) {
  for (var r = fromRowNo; r <= values.length; r++) {
    var first = String(values[r - 1][0] || '').replace(/\n/g, '').trim();
    if (first === '구분') return r;
  }
  return 0;
}

function normalizeDashboardHeader_v630_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, '').trim();
}

function decideNextRotationSafe_v630_(b) {
  if (typeof decideNextRotationSafe_v628_ === 'function') return decideNextRotationSafe_v628_(b);
  if (typeof decideNextRotation_v611_ === 'function') return decideNextRotation_v611_(b);
  return (b.unsoldCount || 0) > 0 ? '상품갈이 검토' : '유지';
}

function formatShortDateSafe_v630_(d) {
  if (!d) return '';
  if (typeof formatShortDateSafe_v628_ === 'function') return formatShortDateSafe_v628_(d);
  if (typeof formatShortDate_ === 'function') return formatShortDate_(d);
  return d;
}
