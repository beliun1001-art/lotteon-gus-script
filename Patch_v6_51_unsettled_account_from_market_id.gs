/**
 * LOTTEON v6.51 unsettled account fix from D column market id
 *
 * 문제:
 * - 미정산_쿠팡계정별 시트에 계정미확인이 남아 있습니다.
 * - 원본 매출데이터_붙여넣기 D열 마켓아이디가 계정 단일 원천이므로,
 *   미정산 시트도 주문번호 기준으로 D열 마켓아이디를 재매핑해야 합니다.
 *
 * 기준:
 * - 매출데이터_붙여넣기 C열 마켓주문번호 → 주문번호 key
 * - 매출데이터_붙여넣기 D열 마켓아이디 → 쿠팡계정ID
 * - 주문번호/마켓주문번호는 쉼표 제거 후 비교
 */

var LOTTEON_PATCH_V651_UNSETTLED_ACCOUNT_FROM_MARKET_ID_LOADED = true;

var __baseBuildUnsettledSheet_v651 = typeof buildUnsettledSettlementByAccountSheet_v629_ === 'function' ? buildUnsettledSettlementByAccountSheet_v629_ : null;
var __baseRefreshDashboardFastOnly_v651 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseApplyDisplayStandardsOnlyFast_v651 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;
var __baseGenerateVatReportsFullSeparated_v651 = typeof generateVatReportsFullSeparated_v622 === 'function' ? generateVatReportsFullSeparated_v622 : null;

buildUnsettledSettlementByAccountSheet_v629_ = function() {
  var result = __baseBuildUnsettledSheet_v651 ? __baseBuildUnsettledSheet_v651.apply(this, arguments) : null;
  try { fixUnsettledAccountFromSalesD_v651(false); } catch (e) { try { log_('patch_v651_unsettled_after_build_error', String(e && e.message ? e.message : e)); } catch(ignore) {} }
  return result || { ok: true };
};

refreshDashboardFastOnly = function() {
  var result = __baseRefreshDashboardFastOnly_v651 ? __baseRefreshDashboardFastOnly_v651.apply(this, arguments) : null;
  try { fixUnsettledAccountFromSalesD_v651(false); } catch (e) {}
  return result || { ok: true };
};

applyDisplayStandardsOnlyFast_v623 = function() {
  var result = __baseApplyDisplayStandardsOnlyFast_v651 ? __baseApplyDisplayStandardsOnlyFast_v651.apply(this, arguments) : null;
  try { fixUnsettledAccountFromSalesD_v651(false); } catch (e) {}
  return result || { ok: true };
};

generateVatReportsFullSeparated_v622 = function() {
  var result = __baseGenerateVatReportsFullSeparated_v651 ? __baseGenerateVatReportsFullSeparated_v651.apply(this, arguments) : null;
  try { fixUnsettledAccountFromSalesD_v651(false); } catch (e) {}
  return result || { ok: true };
};

function fixUnsettledAccountFromSalesD_v651(showAlert) {
  var ss = SpreadsheetApp.getActive();
  var unsettled = ss.getSheetByName('미정산_쿠팡계정별');
  if (!unsettled || unsettled.getLastRow() < 2) {
    writeUnsettledAccountFixDiagnostic_v651_({ skipped: true, reason: '미정산_쿠팡계정별 시트 또는 데이터 없음' });
    if (showAlert !== false) safeAlert_v651_('미정산_쿠팡계정별 시트 또는 데이터가 없습니다.');
    return { ok: true, skipped: true };
  }

  var orderAccountMap = buildOrderAccountMapFromSalesD_v651_();
  var lastRow = unsettled.getLastRow();
  var lastCol = unsettled.getLastColumn();
  var headers = unsettled.getRange(1, 1, 1, lastCol).getValues()[0].map(function(h) { return normalizeHeader_v651_(h); });
  var accountCol = headers.indexOf('쿠팡계정ID') + 1;
  var orderCol = headers.indexOf('주문번호') + 1;
  if (accountCol <= 0 || orderCol <= 0) throw new Error('미정산_쿠팡계정별에서 쿠팡계정ID/주문번호 헤더를 찾지 못했습니다.');

  var values = unsettled.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var fixed = 0;
  var already = 0;
  var unmatched = 0;
  var blankOrder = 0;
  var sampleUnmatched = [];

  for (var r = 0; r < values.length; r++) {
    var current = cleanText_v651_(values[r][accountCol - 1]);
    var orderKey = normalizeOrderNo_v651_(values[r][orderCol - 1]);
    if (!orderKey) {
      blankOrder++;
      continue;
    }
    var mapped = orderAccountMap[orderKey] || '';
    if (mapped) {
      if (current !== mapped) {
        values[r][accountCol - 1] = mapped;
        fixed++;
      } else {
        already++;
      }
    } else {
      unmatched++;
      if (sampleUnmatched.length < 10) sampleUnmatched.push(values[r][orderCol - 1]);
    }
  }

  unsettled.getRange(2, 1, values.length, lastCol).setValues(values);
  try { unsettled.getRange(2, accountCol, values.length, 1).setNumberFormat('@').setHorizontalAlignment('left'); } catch (e) {}
  writeUnsettledAccountFixDiagnostic_v651_({
    sourceMapCount: Object.keys(orderAccountMap).length,
    targetRows: values.length,
    fixed: fixed,
    already: already,
    unmatched: unmatched,
    blankOrder: blankOrder,
    sampleUnmatched: sampleUnmatched.join(', ')
  });
  try { log_('patch_v651_unsettled_account_fix', 'fixed=' + fixed + ' unmatched=' + unmatched + ' map=' + Object.keys(orderAccountMap).length); } catch (e) {}
  if (showAlert !== false) {
    safeAlert_v651_(
      '미정산 계정ID 보정 완료\n\n' +
      '보정: ' + fixed + '행\n' +
      '이미 정상: ' + already + '행\n' +
      '주문번호 미매칭: ' + unmatched + '행\n' +
      '원본 주문번호 map: ' + Object.keys(orderAccountMap).length + '건\n\n' +
      '검증 시트: 미정산_계정ID보정검증'
    );
  }
  return { ok: true, fixed: fixed, already: already, unmatched: unmatched, mapCount: Object.keys(orderAccountMap).length };
}

function buildOrderAccountMapFromSalesD_v651_() {
  var ss = SpreadsheetApp.getActive();
  var source = ss.getSheetByName('매출데이터_붙여넣기');
  if (!source || source.getLastRow() < 2 || source.getLastColumn() < 4) return {};
  var lastRow = source.getLastRow();
  var values = source.getRange(2, 3, lastRow - 1, 2).getValues(); // C:D = 마켓주문번호, 마켓아이디
  var map = {};
  values.forEach(function(row) {
    var orderKey = normalizeOrderNo_v651_(row[0]);
    var account = cleanText_v651_(row[1]);
    if (orderKey && account) map[orderKey] = account;
  });
  return map;
}

function writeUnsettledAccountFixDiagnostic_v651_(info) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('미정산_계정ID보정검증') || ss.insertSheet('미정산_계정ID보정검증');
  var rows = [
    ['항목','값','메모'],
    ['버전','v6.51','D열 마켓아이디 기준'],
    ['원천','매출데이터_붙여넣기 C:D','C=마켓주문번호, D=마켓아이디'],
    ['원본 주문번호 map', info.sourceMapCount || 0, '쉼표 제거 후 주문번호 key'],
    ['미정산 행수', info.targetRows || 0, '미정산_쿠팡계정별'],
    ['보정 행수', info.fixed || 0, '계정미확인 또는 다른 값 → D열 마켓아이디'],
    ['이미 정상', info.already || 0, '현재값과 D열 마켓아이디 일치'],
    ['주문번호 미매칭', info.unmatched || 0, '원본 C열에서 주문번호 찾지 못함'],
    ['주문번호 공란', info.blankOrder || 0, ''],
    ['미매칭 예시', info.sampleUnmatched || '', '최대 10개'],
    ['메모', info.reason || '', '']
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 3).setValues(rows);
  try {
    sheet.getRange(1, 1, 1, 3).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    sheet.autoResizeColumns(1, 3);
  } catch (e) {}
}

function normalizeOrderNo_v651_(v) {
  return String(v == null ? '' : v).replace(/,/g, '').replace(/\s+/g, '').trim();
}

function cleanText_v651_(v) {
  return String(v == null ? '' : v).replace(/,/g, '').trim();
}

function normalizeHeader_v651_(v) {
  return String(v == null ? '' : v).replace(/\n/g, '').replace(/\s+/g, '').trim();
}

function safeAlert_v651_(message) {
  try { SpreadsheetApp.getUi().alert(message); } catch (e) {}
}
