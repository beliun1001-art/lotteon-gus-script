/**
 * LOTTEON v6.39 fast operation sheet visibility
 *
 * 문제:
 * - ⑥ 시트 정리: 운영 시트만 표시 메뉴가 시간초과됩니다.
 * - 단순 숨김/표시 작업인데 시간초과가 나는 것은 기존 함수가 시트 생성/서식/데이터 재계산/autoResize 등
 *   숨김과 무관한 작업까지 함께 수행했을 가능성이 높습니다.
 *
 * 수정:
 * - showOperationSheetsOnly / showOnlyOperationSheets / hideNonOperationSheets 를 모두 초경량 함수로 재정의합니다.
 * - 데이터 읽기, 시트 생성, 서식 변경, 열너비 조정, 부가세/대시보드 재계산을 하지 않습니다.
 * - 현재 존재하는 시트명만 확인하고, 운영 whitelist에 포함된 시트만 표시합니다.
 * - 이미 표시/숨김 상태가 맞는 시트는 건드리지 않습니다.
 */

var LOTTEON_PATCH_V639_FAST_OPERATION_SHEET_VISIBILITY_LOADED = true;

var LOTTEON_V639_OPERATION_SHEET_NAMES = [
  '기준',
  '대시보드',
  '필터별_상품수',
  '쿠팡재전송_로그',
  '브랜드별_마진율',
  '사업자별_계정별_써머리',
  '부가세_신고자료',
  '부가세_상품별',
  '부가세_고객별',
  '부가세_주문번호별',
  '미정산_쿠팡계정별',
  '시트별_금액검증',
  '원본대상행_검증',
  '분석제외_행_리스트',
  '계정구분_D열검증',
  '미정산_정리검증',
  '고객주소_구분검증',
  '매입금액_AC검증',
  '매입금액_미매칭',
  '매입금액_보강로그',
  '매출데이터_붙여넣기'
];

showOperationSheetsOnly = function() {
  return showOperationSheetsOnlyFast_v639_();
};

showOnlyOperationSheets = function() {
  return showOperationSheetsOnlyFast_v639_();
};

hideNonOperationSheets = function() {
  return showOperationSheetsOnlyFast_v639_();
};

function showOperationSheetsOnlyFast_v639_() {
  var started = Date.now();
  var ss = SpreadsheetApp.getActive();
  var sheets = ss.getSheets();
  var keep = buildOperationSheetKeepMap_v639_();
  var visibleBefore = sheets.filter(function(s) { return !s.isSheetHidden(); });
  var active = ss.getActiveSheet();
  var activeName = active ? active.getName() : '';
  var shown = 0;
  var hidden = 0;
  var skipped = 0;

  // Google Sheets는 모든 시트를 숨길 수 없으므로, 운영 시트 중 첫 번째 존재 시트를 먼저 표시합니다.
  var firstKeepSheet = null;
  for (var i = 0; i < sheets.length; i++) {
    if (keep[sheets[i].getName()]) {
      firstKeepSheet = sheets[i];
      break;
    }
  }
  if (!firstKeepSheet) firstKeepSheet = active || sheets[0];
  try {
    if (firstKeepSheet && firstKeepSheet.isSheetHidden()) {
      firstKeepSheet.showSheet();
      shown++;
    }
    if (firstKeepSheet) ss.setActiveSheet(firstKeepSheet);
  } catch (e) {}

  sheets.forEach(function(sheet) {
    var name = sheet.getName();
    var shouldShow = !!keep[name];
    try {
      if (shouldShow) {
        if (sheet.isSheetHidden()) {
          sheet.showSheet();
          shown++;
        } else {
          skipped++;
        }
      } else {
        // 현재 유일한 표시 시트면 숨김 시도하지 않습니다.
        if (!sheet.isSheetHidden()) {
          if (countVisibleSheetsFast_v639_(sheets) > 1) {
            sheet.hideSheet();
            hidden++;
          } else {
            skipped++;
          }
        } else {
          skipped++;
        }
      }
    } catch (e) {
      skipped++;
    }
  });

  try {
    if (keep[activeName]) ss.setActiveSheet(ss.getSheetByName(activeName));
    else if (ss.getSheetByName('대시보드') && !ss.getSheetByName('대시보드').isSheetHidden()) ss.setActiveSheet(ss.getSheetByName('대시보드'));
  } catch (e) {}

  var elapsed = Math.round((Date.now() - started) / 1000);
  try { log_('patch_v639_operation_sheet_visibility', 'shown=' + shown + ' hidden=' + hidden + ' skipped=' + skipped + ' elapsed=' + elapsed); } catch (e) {}
  try {
    SpreadsheetApp.getUi().alert(
      '운영 시트만 표시 완료\n\n' +
      '표시 처리: ' + shown + '\n' +
      '숨김 처리: ' + hidden + '\n' +
      '건너뜀: ' + skipped + '\n' +
      '소요초: ' + elapsed + '\n\n' +
      'v6.39 기준: 숨김/표시만 수행, 데이터 재계산/서식/열너비 조정 없음'
    );
  } catch (e) {}
  return { ok: true, shown: shown, hidden: hidden, skipped: skipped, elapsedSeconds: elapsed };
}

function buildOperationSheetKeepMap_v639_() {
  var keep = {};
  LOTTEON_V639_OPERATION_SHEET_NAMES.forEach(function(name) { keep[name] = true; });
  return keep;
}

function countVisibleSheetsFast_v639_(sheets) {
  var count = 0;
  for (var i = 0; i < sheets.length; i++) {
    try { if (!sheets[i].isSheetHidden()) count++; } catch (e) {}
  }
  return count;
}
