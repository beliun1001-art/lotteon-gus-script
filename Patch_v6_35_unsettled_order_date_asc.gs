/**
 * LOTTEON v6.35 unsettled order date ascending sort
 *
 * 확정 기준:
 * - 미정산_쿠팡계정별 시트 데이터는 주문일 오름차순으로 정렬합니다.
 * - 같은 주문일 안에서는 쿠팡계정ID → 30일초과 여부 → 주문번호 순으로 정렬합니다.
 * - 주문일이 비어 있는 행은 아래로 보냅니다.
 * - 미정산_쿠팡계정별 생성/부가세 생성/대시보드 빠른 갱신/표시서식 정리 후 마지막에 다시 적용합니다.
 */

var LOTTEON_PATCH_V635_UNSETTLED_ORDER_DATE_ASC_LOADED = true;

var __baseBuildUnsettledSheet_v635 = typeof buildUnsettledSettlementByAccountSheet_v629_ === 'function' ? buildUnsettledSettlementByAccountSheet_v629_ : null;
var __baseBuildUnsettledSheetV616_v635 = typeof buildUnsettledSettlementByAccountSheet_v616_ === 'function' ? buildUnsettledSettlementByAccountSheet_v616_ : null;
var __baseGenerateVatReportsFullSeparated_v635 = typeof generateVatReportsFullSeparated_v622 === 'function' ? generateVatReportsFullSeparated_v622 : null;
var __baseRefreshDashboardFastOnly_v635 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseApplyDisplayStandardsOnlyFast_v635 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseApplyFastOutputSheetFormatting_v635 = typeof applyFastOutputSheetFormatting_v620_ === 'function' ? applyFastOutputSheetFormatting_v620_ : null;

buildUnsettledSettlementByAccountSheet_v629_ = function(salesAgg) {
  var result = __baseBuildUnsettledSheet_v635 ? __baseBuildUnsettledSheet_v635.apply(this, arguments) : null;
  sortUnsettledSheetByOrderDateAsc_v635_();
  return result || { ok: true };
};

buildUnsettledSettlementByAccountSheet_v616_ = function(salesAgg) {
  var result = __baseBuildUnsettledSheetV616_v635 ? __baseBuildUnsettledSheetV616_v635.apply(this, arguments) : null;
  sortUnsettledSheetByOrderDateAsc_v635_();
  return result || { ok: true };
};

generateVatReportsFullSeparated_v622 = function() {
  var result = __baseGenerateVatReportsFullSeparated_v635 ? __baseGenerateVatReportsFullSeparated_v635.apply(this, arguments) : null;
  sortUnsettledSheetByOrderDateAsc_v635_();
  return result || { ok: true };
};

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v635 ? __baseRefreshDashboardFastOnly_v635.apply(this, arguments) : null;
  sortUnsettledSheetByOrderDateAsc_v635_();
  return result || { ok: true };
};

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v635 ? __baseApplyDisplayStandardsOnlyFast_v635.apply(this, arguments) : null;
  sortUnsettledSheetByOrderDateAsc_v635_();
  try { SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n미정산_쿠팡계정별 시트를 주문일 오름차순으로 정렬했습니다.'); } catch (e) {}
  return result || { ok: true };
};

applyFastOutputSheetFormatting_v620_ = function() {
  var result = __baseApplyFastOutputSheetFormatting_v635 ? __baseApplyFastOutputSheetFormatting_v635.apply(this, arguments) : null;
  sortUnsettledSheetByOrderDateAsc_v635_();
  return result || { ok: true };
};

function sortUnsettledSheetByOrderDateAsc_v635_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('미정산_쿠팡계정별');
  if (!sheet || sheet.getLastRow() < 3 || sheet.getLastColumn() < 1) return { skipped: true };

  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();
  var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  var headers = values[0].map(function(h) { return normalizeUnsettledHeader_v635_(h); });

  var idx = {
    account: findUnsettledHeader_v635_(headers, ['쿠팡계정ID','계정ID']),
    status: findUnsettledHeader_v635_(headers, ['구분']),
    orderDate: findUnsettledHeader_v635_(headers, ['주문일','주문일자','날짜']),
    elapsed: findUnsettledHeader_v635_(headers, ['경과일']),
    over30OrderNo: findUnsettledHeader_v635_(headers, ['30일초과주문번호']),
    orderNo: findUnsettledHeader_v635_(headers, ['주문번호'])
  };

  if (idx.orderDate < 0) return { skipped: true, reason: '주문일 헤더 없음' };

  var body = values.slice(1).filter(function(row) {
    return row.some(function(v) { return String(v == null ? '' : v).trim() !== ''; });
  });

  body.sort(function(a, b) {
    var da = orderDateSortKey_v635_(a[idx.orderDate]);
    var db = orderDateSortKey_v635_(b[idx.orderDate]);
    if (da !== db) return da - db;

    var aa = idx.account >= 0 ? String(a[idx.account] || '') : '';
    var ab = idx.account >= 0 ? String(b[idx.account] || '') : '';
    var ac = aa.localeCompare(ab);
    if (ac !== 0) return ac;

    var overA = idx.status >= 0 && /30일초과/.test(String(a[idx.status] || '')) ? 0 : 1;
    var overB = idx.status >= 0 && /30일초과/.test(String(b[idx.status] || '')) ? 0 : 1;
    if (overA !== overB) return overA - overB;

    var oa = idx.over30OrderNo >= 0 ? String(a[idx.over30OrderNo] || '') : '';
    var ob = idx.over30OrderNo >= 0 ? String(b[idx.over30OrderNo] || '') : '';
    if (oa || ob) return oa.localeCompare(ob);

    var ra = idx.orderNo >= 0 ? String(a[idx.orderNo] || '') : '';
    var rb = idx.orderNo >= 0 ? String(b[idx.orderNo] || '') : '';
    return ra.localeCompare(rb);
  });

  sheet.getRange(2, 1, lastRow - 1, lastCol).clearContent();
  if (body.length) sheet.getRange(2, 1, body.length, lastCol).setValues(body);

  try {
    sheet.getRange(1, 1, 1, lastCol).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    if (body.length) {
      sheet.getRange(2, 1, body.length, lastCol).setVerticalAlignment('middle');
      if (idx.orderDate >= 0) sheet.getRange(2, idx.orderDate + 1, body.length, 1).setHorizontalAlignment('center');
      if (idx.elapsed >= 0) sheet.getRange(2, idx.elapsed + 1, body.length, 1).setNumberFormat('#,##0').setHorizontalAlignment('right');
    }
    log_('patch_v635_unsettled_order_date_asc', 'rows=' + body.length + ' orderDateCol=' + (idx.orderDate + 1));
  } catch (e) {}

  return { ok: true, rows: body.length };
}

function orderDateSortKey_v635_(v) {
  if (v == null || v === '') return 99999999;
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v)) {
    return Number(Utilities.formatDate(v, CONFIG.TIMEZONE, 'yyyyMMdd'));
  }
  var s = String(v).trim();
  if (!s) return 99999999;

  var ymd = s.match(/(20\d{2})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  if (ymd) return Number(ymd[1] + ('0' + Number(ymd[2])).slice(-2) + ('0' + Number(ymd[3])).slice(-2));

  var md = s.match(/^(\d{1,2})[-.\/](\d{1,2})$/);
  if (md) {
    var year = Number(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy'));
    return Number(String(year) + ('0' + Number(md[1])).slice(-2) + ('0' + Number(md[2])).slice(-2));
  }

  return 99999999;
}

function normalizeUnsettledHeader_v635_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, '').trim();
}

function findUnsettledHeader_v635_(headers, candidates) {
  for (var i = 0; i < headers.length; i++) {
    for (var j = 0; j < candidates.length; j++) {
      if (headers[i] === normalizeUnsettledHeader_v635_(candidates[j])) return i;
    }
  }
  return -1;
}
