/**
 * LOTTEON v6.19 all-sheet autowidth and dashboard split columns patch
 *
 * 수정 내용:
 * 1) v6.18의 일부 지정 시트만 서식 정리하던 방식을 전체 시트 기준으로 확장합니다.
 * 2) 모든 시트는 autoResizeColumns를 먼저 적용합니다.
 * 3) 상품명/메모/비고/주문번호처럼 긴 텍스트 열은 최대 폭을 제한하고 화면 표시를 CLIP 처리합니다.
 * 4) 원본/붙여넣기 시트의 실제 값은 축약하지 않고, 출력용 시트만 긴 텍스트를 축약합니다.
 * 5) 대시보드의 긴 메모 셀을 실제정산금액/예상정산금액/미정산건수/예상이익 컬럼으로 분리합니다.
 * 6) 대시보드 섹션별 숫자/금액/비율 서식을 헤더명 기준으로 다시 적용합니다.
 */

var LOTTEON_PATCH_V619_ALL_SHEET_AUTOWIDTH_DASHBOARD_SPLIT_LOADED = true;

var __baseRefreshDashboardFastOnly_v619 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseRunPendingChangesApproval_v619 = typeof runPendingChangesApproval === 'function' ? runPendingChangesApproval : null;
var __baseRefreshCoreSummaryAndDashboard_v619 = typeof refreshCoreSummaryAndDashboardWithRetransmitLogDates === 'function' ? refreshCoreSummaryAndDashboardWithRetransmitLogDates : null;

rebuildDashboardBrandBased_v611_ = function(salesAgg, filterAgg) {
  return rebuildDashboardSplitColumns_v619_(salesAgg || aggregateSalesByBrand_v611_(), filterAgg || aggregateFiltersByBrand_v611_());
};

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v619 ? __baseRefreshDashboardFastOnly_v619.apply(this, arguments) : null;
  applyAllSheetsAutoWidthAndCompact_v619_();
  return result;
};

runPendingChangesApproval = function() {
  var result = __baseRunPendingChangesApproval_v619 ? __baseRunPendingChangesApproval_v619.apply(this, arguments) : null;
  applyAllSheetsAutoWidthAndCompact_v619_();
  return result;
};

refreshCoreSummaryAndDashboardWithRetransmitLogDates = function() {
  var result = __baseRefreshCoreSummaryAndDashboard_v619 ? __baseRefreshCoreSummaryAndDashboard_v619.apply(this, arguments) : null;
  applyAllSheetsAutoWidthAndCompact_v619_();
  return result;
};

function rebuildDashboardSplitColumns_v619_(salesAgg, filterAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD) || ss.insertSheet(CONFIG.SHEETS.DASHBOARD);
  salesAgg = salesAgg || aggregateSalesByBrand_v611_();
  filterAgg = filterAgg || aggregateFiltersByBrand_v611_();

  var brandRows = buildBrandRowsForDashboard_v611_(salesAgg, filterAgg);
  brandRows.forEach(function(b) {
    var source = salesAgg.byBrand && salesAgg.byBrand[b.key] ? salesAgg.byBrand[b.key] : {};
    b.grossSalesAmount = source.grossSalesAmount || 0;
    b.cancelSalesAmount = source.cancelSalesAmount || 0;
    b.netSalesAmount = source.netSalesAmount || b.salesAmount || 0;
    b.actualSettlementAmount = source.actualSettlementAmount || 0;
    b.estimatedSettlementAmount = source.estimatedSettlementAmount || 0;
    b.settlementBasisAmount = source.settlementBasisAmount || source.settlementAmount || b.settlementAmount || 0;
    b.purchaseAmount = source.purchaseAmount || b.purchaseAmount || 0;
    b.marketFeeAmount = source.marketFeeAmount || (b.netSalesAmount - b.settlementBasisAmount);
    b.estimatedProfitAmount = source.estimatedProfitAmount || source.netProfitAmount || b.netProfitAmount || 0;
    b.estimatedProfitRate = source.estimatedProfitRate || source.netProfitRate || b.marginRate || 0;
    b.unsettledOrderCount = source.unsettledOrderCount || 0;
    b.overdueUnsettledOrderCount = source.overdueUnsettledOrderCount || 0;
    b.unsettledNetSales = source.unsettledNetSales || 0;
  });

  var nowText = Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy-MM-dd HH:mm:ss');
  var overdueCount = Object.keys(salesAgg.overdueUnsettledOrderSet || {}).length;
  var unsettledCount = Object.keys(salesAgg.unsettledOrderSet || {}).length;
  var actualSettlement = salesAgg.totalActualSettlement || 0;
  var estimatedSettlement = salesAgg.totalEstimatedSettlement || 0;
  var settlementBasis = salesAgg.totalSettlementBasis || salesAgg.totalSettlement || 0;
  var totalProfit = salesAgg.totalEstimatedProfit || salesAgg.totalNetProfit || 0;
  var profitRate = salesAgg.totalEstimatedProfitRate || salesAgg.totalNetProfitRate || 0;
  var marketFee = salesAgg.totalMarketFee || ((salesAgg.totalNetSales || salesAgg.totalSales || 0) - settlementBasis);
  var marketFeeRate = salesAgg.totalMarketFeeRate || ((salesAgg.totalNetSales || salesAgg.totalSales || 0) ? marketFee / (salesAgg.totalNetSales || salesAgg.totalSales || 0) : 0);
  var totalNetSales = salesAgg.totalNetSales || salesAgg.totalSales || 0;
  var vatBaseTotal = Math.round(totalNetSales / 1.1);

  var rows = [];
  rows.push(['요약','갱신기준','브랜드 기준 + 미정산 예상정산 90.1% 반영 + 컬럼분리','','','','','','','','','','','','','v6.19']);
  rows.push(['요약','갱신시각',nowText,'','','','','','','','','','','','','Asia/Seoul']);
  rows.push(['요약','분석브랜드수',brandRows.length,'','','','','','','','','','','','','브랜드명 기준']);
  rows.push(['요약','수집상품수',filterAgg.totalCollectCount || 0,'','','','','','','','','','','','','필터별_상품수 API_totalCount 총합']);
  rows.push(['요약','매출상품수',salesAgg.totalProducts || 0,'','','','','','','','','','','','','취소/반품 제외']);
  rows.push(['요약','주문건수',salesAgg.totalOrders || 0,'','','','','','','','','','','','','취소/반품 제외']);
  rows.push(['요약','미정산 주문건수',unsettledCount,'','','','','','','','','','','','','정산예정금액 공란 정상 주문']);
  rows.push(['요약','30일초과 미정산 주문건수',overdueCount,'','','','','','','','','','','','','주문일 기준 30일 초과']);
  rows.push(['요약','총매출액',salesAgg.totalGrossSales || 0,'','','','','','','','','','','','','취소/반품 포함']);
  rows.push(['요약','취소/반품매출',salesAgg.totalCancelSales || 0,'','','','','','','','','','','','','취소/반품 판정 row']);
  rows.push(['요약','순수매출액',totalNetSales,'','','','','','','','','','','','','취소/반품 제외']);
  rows.push(['요약','실제정산금액',actualSettlement,'','','','','','','','','','','','','정산예정금액 입력 완료 주문']);
  rows.push(['요약','예상정산금액(미정산)',estimatedSettlement,'','','','','','','','','','','','','미정산 순수매출액 × 0.901']);
  rows.push(['요약','정산기준금액',settlementBasis,'','','','','','','','','','','','','실제정산금액 + 예상정산금액']);
  rows.push(['요약','마켓수수료/비용',marketFee,'','','','','','','','','','','','','순수매출액 - 정산기준금액']);
  rows.push(['요약','마켓수수료율',marketFeeRate,'','','','','','','','','','','','','마켓수수료 ÷ 순수매출액']);
  rows.push(['요약','매입금액',salesAgg.totalPurchase || 0,'','','','','','','','','','','','',salesAgg.hasPurchaseColumn ? '취소제외 매입/원가 컬럼 기준' : '매입/원가 컬럼 미확인']);
  rows.push(['요약','예상이익',totalProfit,'','','','','','','','','','','','','정산기준금액 - 매입금액']);
  rows.push(['요약','예상이익률',profitRate,'','','','','','','','','','','','','예상이익 ÷ 순수매출액']);
  rows.push(['요약','부가세 신고 공급대가',totalNetSales,'','','','','','','','','','','','','순수매출액 기준']);
  rows.push(['요약','부가세 신고 공급가액',vatBaseTotal,'','','','','','','','','','','','','공급대가 ÷ 1.1']);
  rows.push(['요약','부가세',totalNetSales - vatBaseTotal,'','','','','','','','','','','','','공급대가 - 공급가액']);
  rows.push(['','','','','','','','','','','','','','','','']);

  var tableHeader = ['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','순수매출액','실제정산금액','예상정산금액(미정산)','정산기준금액','매입금액','예상이익','예상이익률','미정산건수','30일초과건수'];
  rows.push(tableHeader);
  brandRows.slice().sort(function(a, b) { return (b.netSalesAmount || 0) - (a.netSalesAmount || 0); }).slice(0, 30).forEach(function(b) {
    rows.push([
      '브랜드TOP', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount,
      b.netSalesAmount, b.actualSettlementAmount, b.estimatedSettlementAmount, b.settlementBasisAmount,
      b.purchaseAmount, b.estimatedProfitAmount, b.estimatedProfitRate, b.unsettledOrderCount, b.overdueUnsettledOrderCount
    ]);
  });
  rows.push(['','','','','','','','','','','','','','','','']);

  rows.push(['구분','브랜드명','대표검색필터명','계정ID','수집수','매출상품','주문','순수매출액','정산기준금액','예상이익률','미정산건수','30일초과건수','다음작업','최근수집일','미판매수','메모']);
  brandRows.slice().sort(function(a, b) { return (b.unsoldCount || 0) - (a.unsoldCount || 0); }).slice(0, 40).forEach(function(b) {
    rows.push([
      '상품갈이', b.brand, b.filterName, b.accountId, b.collectCount, b.salesProductCount, b.orderCount,
      b.netSalesAmount, b.settlementBasisAmount, b.estimatedProfitRate, b.unsettledOrderCount, b.overdueUnsettledOrderCount,
      decideNextRotation_v611_(b), formatShortDate_(b.latestRecentDate), b.unsoldCount, '브랜드 기준'
    ]);
  });

  replaceSheetValues_v611_(sheet, buildHeader16_v619_(), rows);
  formatDashboardByHeader_v619_(sheet);
  if (typeof buildUnsettledSettlementByAccountSheet_v616_ === 'function') buildUnsettledSettlementByAccountSheet_v616_(salesAgg);
  applyAllSheetsAutoWidthAndCompact_v619_();
  return { rows: rows.length };
}

function buildHeader16_v619_() {
  return ['구분','항목','값1','값2','값3','값4','값5','값6','값7','값8','값9','값10','값11','값12','값13','메모'];
}

function formatDashboardByHeader_v619_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var lastRow = sheet.getLastRow();
  var lastCol = Math.max(sheet.getLastColumn(), 16);
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  try {
    sheet.getRange(1, 1, lastRow, lastCol).setFontFamily('Arial').setFontSize(10).setFontWeight('normal').setBackground(null).setVerticalAlignment('middle');
    sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#d9eaf7').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
  } catch (e) {}

  for (var r = 1; r <= lastRow; r++) {
    var row = values[r - 1];
    var first = String(row[0] || '').trim();
    if (r === 1 || first === '구분') {
      try {
        sheet.getRange(r, 1, 1, lastCol).setFontWeight('bold').setBackground('#d9eaf7').setHorizontalAlignment('center');
      } catch (e) {}
      applyTableFormatsFromHeaderRow_v619_(sheet, r, findTableEndRow_v619_(values, r));
    }
  }

  // 요약 영역 값 서식: 항목명 기준으로 값1만 처리
  for (var i = 2; i <= Math.min(lastRow, 23); i++) {
    var item = String(values[i - 1][1] || '').trim();
    try {
      if (/율/.test(item)) sheet.getRange(i, 3).setNumberFormat('0.00%');
      else if (/금액|매출|정산|수수료|매입|이익|공급|부가세/.test(item)) sheet.getRange(i, 3).setNumberFormat('₩#,##0');
      else if (/시각/.test(item)) sheet.getRange(i, 3).setNumberFormat('yyyy-mm-dd hh:mm:ss');
      else sheet.getRange(i, 3).setNumberFormat('#,##0');
    } catch (e) {}
  }
}

function applyTableFormatsFromHeaderRow_v619_(sheet, headerRowNo, endRowNo) {
  if (endRowNo <= headerRowNo) return;
  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(headerRowNo, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').trim(); });
  var rowCount = endRowNo - headerRowNo;
  headers.forEach(function(h, idx) {
    var col = idx + 1;
    if (!h) return;
    try {
      var range = sheet.getRange(headerRowNo + 1, col, rowCount, 1);
      if (/순수매출액|실제정산금액|예상정산금액|정산기준금액|매입금액|예상이익|총매출|취소|수수료|공급|부가세/.test(h)) range.setNumberFormat('₩#,##0');
      else if (/율/.test(h)) range.setNumberFormat('0.00%');
      else if (/수집수|매출상품|주문|미정산건수|30일초과건수|미판매수|수량|건수/.test(h)) range.setNumberFormat('#,##0');
      else if (/주문번호|상품번호|계정ID/.test(h)) range.setNumberFormat('@');
    } catch (e) {}
  });
}

function findTableEndRow_v619_(values, headerRowNo) {
  var startIdx = headerRowNo;
  for (var i = startIdx; i < values.length; i++) {
    var row = values[i];
    var first = String(row[0] || '').trim();
    var nonBlank = row.some(function(v) { return v !== '' && v != null; });
    if (i > startIdx && first === '구분') return i;
    if (!nonBlank && i > startIdx) return i;
  }
  return values.length;
}

function applyAllSheetsAutoWidthAndCompact_v619_() {
  var ss = SpreadsheetApp.getActive();
  ss.getSheets().forEach(function(sheet) {
    try {
      autoWidthAndLimitSheet_v619_(sheet);
    } catch (e) {
      try { log_('patch_v619_sheet_format_error', sheet.getName() + ': ' + e); } catch (ignore) {}
    }
  });
  try { log_('patch_v619_all_sheet_autowidth', 'done'); } catch (e) {}
}

function autoWidthAndLimitSheet_v619_(sheet) {
  if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').trim(); });

  sheet.getRange(1, 1, lastRow, lastCol).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#d9eaf7').setHorizontalAlignment('center');
  sheet.autoResizeColumns(1, lastCol);

  for (var c = 1; c <= lastCol; c++) {
    var h = header[c - 1] || guessHeaderFromColumn_v619_(sheet, c);
    var maxWidth = maxColumnWidth_v619_(h);
    if (sheet.getColumnWidth(c) > maxWidth) sheet.setColumnWidth(c, maxWidth);
    if (/상품명|메모|비고|사유|주문번호|필터명/.test(h)) {
      sheet.getRange(1, c, lastRow, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    }
    if (/년|월|일/.test(h)) sheet.getRange(2, c, Math.max(lastRow - 1, 1), 1).setNumberFormat('0');
    if (/주문번호|상품번호|계정ID/.test(h)) sheet.getRange(2, c, Math.max(lastRow - 1, 1), 1).setNumberFormat('@');
  }

  if (isOutputSheetForTextCompact_v619_(sheet.getName())) compactLongTextCells_v619_(sheet);
}

function isOutputSheetForTextCompact_v619_(sheetName) {
  return [
    CONFIG.SHEETS.DASHBOARD,
    CONFIG.SHEETS.RETRANSMIT_LOG,
    '브랜드별_마진율',
    '부가세_신고자료',
    '부가세_상품별',
    '부가세_고객별',
    '부가세_주문번호별',
    '미정산_쿠팡계정별'
  ].indexOf(sheetName) >= 0;
}

function compactLongTextCells_v619_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').trim(); });
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var changed = false;

  for (var c = 0; c < header.length; c++) {
    var limit = textLimitByHeader_v619_(header[c]);
    if (!limit) continue;
    for (var r = 0; r < values.length; r++) {
      var before = values[r][c];
      var after = compactText_v619_(before, limit);
      if (after !== before) {
        values[r][c] = after;
        changed = true;
      }
    }
  }
  if (changed) sheet.getRange(2, 1, lastRow - 1, lastCol).setValues(values);
}

function textLimitByHeader_v619_(header) {
  var h = String(header || '').trim();
  if (/상품명/.test(h)) return 54;
  if (/메모|비고|사유/.test(h)) return 36;
  if (/대표검색필터명/.test(h)) return 34;
  if (/고객명/.test(h)) return 18;
  if (/주문번호/.test(h)) return 34;
  return 0;
}

function maxColumnWidth_v619_(header) {
  var h = String(header || '').trim();
  if (/상품명/.test(h)) return 360;
  if (/메모|비고|사유/.test(h)) return 260;
  if (/대표검색필터명/.test(h)) return 210;
  if (/주문번호/.test(h)) return 180;
  if (/고객명/.test(h)) return 120;
  if (/브랜드명/.test(h)) return 140;
  if (/상품번호/.test(h)) return 130;
  if (/계정ID/.test(h)) return 120;
  if (/년|월|일/.test(h)) return 70;
  if (/율/.test(h)) return 90;
  if (/금액|매출|정산|매입|이익|수수료|부가세/.test(h)) return 130;
  return 150;
}

function guessHeaderFromColumn_v619_(sheet, col) {
  try {
    var lastRow = Math.min(sheet.getLastRow(), 80);
    var vals = sheet.getRange(1, col, lastRow, 1).getValues();
    for (var i = 0; i < vals.length; i++) {
      var text = String(vals[i][0] || '').trim();
      if (/상품명|메모|비고|주문번호|브랜드명|금액|율|계정ID|필터명/.test(text)) return text;
    }
  } catch (e) {}
  return '';
}

function compactText_v619_(value, limit) {
  var text = String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  if (!limit || text.length <= limit) return text;
  return text.slice(0, Math.max(0, limit - 1)) + '…';
}
