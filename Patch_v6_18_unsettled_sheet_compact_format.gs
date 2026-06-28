/**
 * LOTTEON v6.18 unsettled sheet compact format patch
 *
 * 수정 내용:
 * 1) 미정산_쿠팡계정별 요약 행의 주문번호 셀에 전체 미정산 주문번호를 쉼표로 몰아넣지 않습니다.
 * 2) 주문번호 컬럼에는 30일초과 미정산 주문번호만 표시합니다.
 * 3) 30일초과 주문번호가 계정별 2개 이상이면 한 셀에 쉼표로 넣지 않고 아래 행으로 1건씩 분리합니다.
 * 4) 미정산_쿠팡계정별 헤더를 대시보드처럼 색상/굵은 글씨/가운데 정렬로 꾸밉니다.
 * 5) 주요 출력 시트는 자동 열너비 조정 후 긴 텍스트 열은 최대 폭으로 제한하고, 상품명/비고는 축약 표시합니다.
 */

var LOTTEON_PATCH_V618_UNSETTLED_SHEET_COMPACT_FORMAT_LOADED = true;

var __baseRefreshDashboardFastOnly_v618 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseRunPendingChangesApproval_v618 = typeof runPendingChangesApproval === 'function' ? runPendingChangesApproval : null;
var __baseRefreshCoreSummaryAndDashboard_v618 = typeof refreshCoreSummaryAndDashboardWithRetransmitLogDates === 'function' ? refreshCoreSummaryAndDashboardWithRetransmitLogDates : null;
var __baseRebuildDashboardBrandBased_v618 = typeof rebuildDashboardBrandBased_v611_ === 'function' ? rebuildDashboardBrandBased_v611_ : null;

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v618 ? __baseRefreshDashboardFastOnly_v618.apply(this, arguments) : null;
  applyWorkbookCompactFormatting_v618_();
  return result;
};

runPendingChangesApproval = function() {
  var result = __baseRunPendingChangesApproval_v618 ? __baseRunPendingChangesApproval_v618.apply(this, arguments) : null;
  applyWorkbookCompactFormatting_v618_();
  return result;
};

refreshCoreSummaryAndDashboardWithRetransmitLogDates = function() {
  var result = __baseRefreshCoreSummaryAndDashboard_v618 ? __baseRefreshCoreSummaryAndDashboard_v618.apply(this, arguments) : null;
  applyWorkbookCompactFormatting_v618_();
  return result;
};

if (__baseRebuildDashboardBrandBased_v618) {
  rebuildDashboardBrandBased_v611_ = function() {
    var result = __baseRebuildDashboardBrandBased_v618.apply(this, arguments);
    applyWorkbookCompactFormatting_v618_();
    return result;
  };
}

function buildUnsettledSettlementByAccountSheet_v616_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('미정산_쿠팡계정별') || ss.insertSheet('미정산_쿠팡계정별');
  salesAgg = salesAgg || (typeof aggregateSalesWithUnsettledEstimate_v616_ === 'function' ? aggregateSalesWithUnsettledEstimate_v616_() : null);
  salesAgg = salesAgg || { unsettledRows: [], overdueRows: [] };

  var headers = [
    '쿠팡계정ID','구분','미정산건수','30일초과건수','주문일','경과일','30일초과 주문번호',
    '브랜드명','고객명','상품번호','상품명','순수매출액','예상정산금액','매입금액','예상이익','비고'
  ];

  var accountSummary = {};
  (salesAgg.unsettledRows || []).forEach(function(d) {
    var account = d.accountId || '계정미확인';
    if (!accountSummary[account]) {
      accountSummary[account] = createUnsettledAccountSummary_v618_(account);
    }
    var s = accountSummary[account];
    s.unsettledCount += 1;
    s.netSales += d.netSales || 0;
    s.estimatedSettlement += d.estimatedSettlement || 0;
    s.purchase += d.purchase || 0;
    s.estimatedProfit += d.estimatedProfit || 0;
    if (d.orderNo) s.unsettledOrders[d.orderNo] = true;
    if (d.overdueStatus) {
      s.overdueCount += 1;
      if (d.orderNo) s.overdueOrders[d.orderNo] = true;
      s.overdueRows.push(d);
    }
  });

  var rows = [];
  Object.keys(accountSummary).sort().forEach(function(account) {
    var s = accountSummary[account];
    rows.push([
      account,
      '요약',
      s.unsettledCount,
      s.overdueCount,
      '',
      '',
      '',
      '',
      '',
      '',
      '30일초과 주문번호만 아래 행에 표시',
      s.netSales,
      s.estimatedSettlement,
      s.purchase,
      s.estimatedProfit,
      s.overdueCount ? '30일초과 미정산 확인 필요' : '30일초과 없음'
    ]);

    s.overdueRows.sort(function(a, b) {
      return String(a.orderDate || '').localeCompare(String(b.orderDate || '')) || String(a.orderNo || '').localeCompare(String(b.orderNo || ''));
    }).forEach(function(d) {
      rows.push([
        account,
        '30일초과',
        '',
        1,
        d.orderDate || '',
        d.elapsedDays || '',
        d.orderNo || '',
        d.brand || '',
        compactText_v618_(d.customer || '', 18),
        d.productNo || '',
        compactText_v618_(d.productName || '', 54),
        d.netSales || 0,
        d.estimatedSettlement || 0,
        d.purchase || 0,
        d.estimatedProfit || 0,
        '30일초과 미정산'
      ]);
    });
  });

  replaceSheetValues_v611_(sheet, headers, rows);
  formatUnsettledSheet_v618_(sheet);
  applyWorkbookCompactFormatting_v618_();
  return { rows: rows.length };
}

function createUnsettledAccountSummary_v618_(account) {
  return {
    account: account,
    unsettledCount: 0,
    overdueCount: 0,
    netSales: 0,
    estimatedSettlement: 0,
    purchase: 0,
    estimatedProfit: 0,
    unsettledOrders: {},
    overdueOrders: {},
    overdueRows: []
  };
}

function formatUnsettledSheet_v618_(sheet) {
  if (!sheet || sheet.getLastRow() < 1) return;
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  try {
    sheet.getRange(1, 1, 1, lastCol)
      .setBackground('#d9eaf7')
      .setFontWeight('bold')
      .setHorizontalAlignment('center')
      .setVerticalAlignment('middle');
    if (lastRow > 1) {
      sheet.getRange(2, 1, lastRow - 1, lastCol).setFontWeight('normal').setBackground(null).setVerticalAlignment('middle');
      sheet.getRange(2, 12, lastRow - 1, 4).setNumberFormat('₩#,##0');
      sheet.getRange(2, 3, lastRow - 1, 2).setNumberFormat('#,##0');
      sheet.getRange(2, 6, lastRow - 1, 1).setNumberFormat('#,##0');
      sheet.getRange(2, 7, lastRow - 1, 1).setNumberFormat('@');
    }
    // 요약 행 강조
    var values = sheet.getRange(2, 2, Math.max(lastRow - 1, 1), 1).getValues();
    for (var r = 0; r < values.length; r++) {
      if (String(values[r][0] || '') === '요약') {
        sheet.getRange(r + 2, 1, 1, lastCol).setBackground('#eef5fb').setFontWeight('bold');
      }
    }
    sheet.setFrozenRows(1);
  } catch (e) {}
}

function applyWorkbookCompactFormatting_v618_() {
  var ss = SpreadsheetApp.getActive();
  var sheetNames = [
    CONFIG.SHEETS.DASHBOARD,
    CONFIG.SHEETS.RETRANSMIT_LOG,
    '브랜드별_마진율',
    '부가세_신고자료',
    '부가세_상품별',
    '부가세_고객별',
    '부가세_주문번호별',
    '미정산_쿠팡계정별'
  ];

  sheetNames.forEach(function(name) {
    var sheet = ss.getSheetByName(name);
    if (!sheet || sheet.getLastRow() < 1 || sheet.getLastColumn() < 1) return;
    compactSheetDisplay_v618_(sheet);
  });

  try { log_('patch_compact_format_v618', 'formatted=' + sheetNames.join(',')); } catch (e) {}
}

function compactSheetDisplay_v618_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return;

  try {
    sheet.getRange(1, 1, lastRow, lastCol).setFontFamily('Arial').setFontSize(10).setVerticalAlignment('middle');
    sheet.getRange(1, 1, 1, lastCol).setFontWeight('bold').setBackground('#d9eaf7').setHorizontalAlignment('center');
    sheet.setFrozenRows(1);
    sheet.autoResizeColumns(1, lastCol);
  } catch (e) {}

  // 긴 텍스트가 많은 열은 자동 너비 후 최대 폭 제한 + 잘라 보이기 처리
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').trim(); });
  for (var c = 1; c <= lastCol; c++) {
    var h = header[c - 1];
    var maxWidth = maxColumnWidthByHeader_v618_(h);
    try {
      if (sheet.getColumnWidth(c) > maxWidth) sheet.setColumnWidth(c, maxWidth);
      if (/상품명|메모|비고|사유|주문번호/.test(h)) {
        sheet.getRange(1, c, lastRow, 1).setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
      }
      if (/년|월|일/.test(h)) {
        sheet.getRange(2, c, Math.max(lastRow - 1, 1), 1).setNumberFormat('0');
      }
      if (/주문번호|상품번호|계정ID/.test(h)) {
        sheet.getRange(2, c, Math.max(lastRow - 1, 1), 1).setNumberFormat('@');
      }
    } catch (e) {}
  }

  // 값 자체를 과하게 길게 만들 필요 없는 출력 시트만 텍스트 축약
  if (['미정산_쿠팡계정별', '부가세_신고자료', '부가세_상품별', '부가세_주문번호별'].indexOf(sheet.getName()) >= 0) {
    compactLongTextCellsByHeader_v618_(sheet);
  }
}

function compactLongTextCellsByHeader_v618_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  if (lastRow < 2) return;
  var header = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return String(h || '').trim(); });
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var changed = false;

  for (var c = 0; c < header.length; c++) {
    var h = header[c];
    var limit = 0;
    if (/상품명/.test(h)) limit = 54;
    else if (/비고|메모|사유/.test(h)) limit = 34;
    else if (/고객명/.test(h)) limit = 18;
    else if (/주문번호/.test(h)) limit = 30;
    if (!limit) continue;

    for (var r = 0; r < values.length; r++) {
      var original = values[r][c];
      var compact = compactText_v618_(original, limit);
      if (compact !== original) {
        values[r][c] = compact;
        changed = true;
      }
    }
  }
  if (changed) sheet.getRange(2, 1, lastRow - 1, lastCol).setValues(values);
}

function maxColumnWidthByHeader_v618_(header) {
  var h = String(header || '').trim();
  if (/상품명/.test(h)) return 360;
  if (/주문번호/.test(h)) return 170;
  if (/고객명/.test(h)) return 120;
  if (/브랜드명/.test(h)) return 130;
  if (/비고|메모|사유/.test(h)) return 260;
  if (/대표검색필터명/.test(h)) return 210;
  if (/상품번호/.test(h)) return 130;
  if (/계정ID/.test(h)) return 115;
  return 150;
}

function compactText_v618_(value, limit) {
  var text = String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  if (!limit || text.length <= limit) return text;
  return text.slice(0, Math.max(0, limit - 1)) + '…';
}
