/**
 * LOTTEON v6.05 cleanup patch
 *
 * 목적:
 * - 쿠팡전송수_수동입력 시트를 전송수 기준에서 완전히 제외한 뒤 삭제 준비
 * - 변경감지 대상은 매출데이터_붙여넣기만 유지
 * - ⑥ 운영 시트 정리에서 쿠팡전송수_수동입력은 숨김 처리
 *
 * 주의:
 * - 이 패치는 실제 시트를 자동 삭제하지 않습니다.
 * - 1일 운영 확인 후 사용자가 직접 삭제하는 것을 기준으로 합니다.
 */

var LOTTEON_DEPRECATED_COUPANG_SENT_MANUAL_SHEET_V605 = '쿠팡전송수_수동입력';

function startChangeDetectionApproval() {
  const ui = SpreadsheetApp.getUi();
  try {
    const deleted = deleteChangeDetectionTriggers_();
    ScriptApp.newTrigger('handleWatchedSheetEdit')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    const props = PropertiesService.getScriptProperties();
    props.setProperty('CHANGE_DETECTION_ENABLED', 'Y');
    props.setProperty('CHANGE_DETECTION_STARTED_AT', now_());
    props.setProperty('CHANGE_DETECTION_LAST_MESSAGE', '변경감지 기능 시작 - 매출데이터_붙여넣기만 감지');
    props.setProperty('CHANGE_DETECTION_LAST_ERROR', '');
    props.setProperty('CHANGE_MANUAL_DIRTY', 'N');
    props.deleteProperty('CHANGE_MANUAL_DIRTY_AT');
    props.deleteProperty('CHANGE_MANUAL_DIRTY_RANGE');

    writeChangeDetectionStatus_('ENABLED', '변경감지 기능 시작', 'v6.05: 감지 대상은 매출데이터_붙여넣기만 사용 / 기존 감지 트리거 정리 수: ' + deleted);

    ui.alert(
      '변경감지 기능을 시작했습니다.\n\n' +
      'v6.05 기준 감지 대상:\n' +
      '1. 매출데이터_붙여넣기\n\n' +
      '쿠팡전송수는 이제 필터별_상품수 API_totalCount를 사용하므로\n' +
      '쿠팡전송수_수동입력 시트는 감지 대상에서 제외했습니다.\n\n' +
      '실제 반영은 LOTTEON 자동화 → ① 변경사항 반영 실행을 눌러주세요.'
    );
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    PropertiesService.getScriptProperties().setProperty('CHANGE_DETECTION_LAST_ERROR', msg);
    writeChangeDetectionStatus_('ERROR', '변경감지 기능 시작 실패', msg);
    ui.alert('변경감지 기능 시작 중 오류가 발생했습니다.\n\n' + msg);
    throw e;
  }
}

function handleWatchedSheetEdit(e) {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('CHANGE_DETECTION_ENABLED') !== 'Y') return;
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();
    const sheetName = sheet ? sheet.getName() : '';
    const a1 = e.range.getA1Notation ? e.range.getA1Notation() : '';
    const salesSheetName = CONFIG.SHEETS.SALES_IN;

    if (sheetName !== salesSheetName) return;

    const editedAt = now_();
    props.setProperty('CHANGE_LAST_EDIT_SHEET', sheetName);
    props.setProperty('CHANGE_LAST_EDIT_RANGE', a1);
    props.setProperty('CHANGE_LAST_EDIT_AT', editedAt);
    props.setProperty('CHANGE_LAST_MESSAGE', sheetName + ' 변경 감지');
    props.setProperty('CHANGE_SALES_DIRTY', 'Y');
    props.setProperty('CHANGE_SALES_DIRTY_AT', editedAt);
    props.setProperty('CHANGE_SALES_DIRTY_RANGE', a1);

    props.setProperty('CHANGE_MANUAL_DIRTY', 'N');
    props.deleteProperty('CHANGE_MANUAL_DIRTY_AT');
    props.deleteProperty('CHANGE_MANUAL_DIRTY_RANGE');

    writeChangeDetectionStatus_('DIRTY', sheetName + ' 변경 감지', 'LOTTEON 자동화 → ① 변경사항 반영 실행 필요 / v6.05 수동전송수 감지 제외');

    try { SpreadsheetApp.getActive().toast('변경 감지됨: ① 변경사항 반영 실행 필요', 'LOTTEON 자동화', 5); } catch (err) {}
  } catch (err) {
    const msg = String(err && err.message ? err.message : err);
    PropertiesService.getScriptProperties().setProperty('CHANGE_DETECTION_LAST_ERROR', msg);
    writeChangeDetectionStatus_('ERROR', '변경감지 오류', msg);
  }
}

function resetChangeDetectionFlags() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  [
    'CHANGE_SALES_DIRTY', 'CHANGE_SALES_DIRTY_AT', 'CHANGE_SALES_DIRTY_RANGE',
    'CHANGE_MANUAL_DIRTY', 'CHANGE_MANUAL_DIRTY_AT', 'CHANGE_MANUAL_DIRTY_RANGE',
    'CHANGE_LAST_EDIT_SHEET', 'CHANGE_LAST_EDIT_RANGE', 'CHANGE_LAST_EDIT_AT',
    'CHANGE_LAST_RESULT', 'CHANGE_LAST_MESSAGE', 'CHANGE_DETECTION_LAST_ERROR'
  ].forEach(function(key) { props.deleteProperty(key); });
  props.setProperty('CHANGE_MANUAL_DIRTY', 'N');
  writeChangeDetectionStatus_('RESET', '변경감지 플래그 초기화', 'v6.05: 수동전송수 플래그 제거 완료');
  ui.alert('변경감지 플래그를 초기화했습니다.\n\n수동전송수 관련 플래그도 함께 제거했습니다.');
}

function refreshManualApiCheckOnly() {
  const ss = SpreadsheetApp.getActive();
  const deprecated = ss.getSheetByName(LOTTEON_DEPRECATED_COUPANG_SENT_MANUAL_SHEET_V605);
  if (deprecated) {
    try { deprecated.hideSheet(); } catch (e) {}
  }
  SpreadsheetApp.getUi().alert(
    'v6.05 기준 이 메뉴는 더 이상 필요하지 않습니다.\n\n' +
    '쿠팡전송수 기준:\n' +
    '필터별_상품수 → API_totalCount\n\n' +
    '쿠팡전송수_수동입력 시트는 숨김 처리했습니다.\n' +
    '하루 정도 운영 확인 후 문제가 없으면 시트 탭에서 직접 삭제해도 됩니다.'
  );
}

function showOnlyMainSheets() {
  const ss = SpreadsheetApp.getActive();
  const visibleNames = [
    CONFIG.SHEETS.DASHBOARD,
    CONFIG.SHEETS.BRAND_SUMMARY,
    CONFIG.SHEETS.RETRANSMIT_LOG,
    CONFIG.SHEETS.FILTERS,
    CONFIG.SHEETS.SALES_IN,
    CONFIG.SHEETS.SYNC_STATUS,
    '자동갱신_상태',
    '필터_대시보드'
  ];
  const visibleSet = {};
  visibleNames.forEach(function(name) { if (name) visibleSet[String(name)] = true; });

  let shown = 0;
  let hidden = 0;
  ss.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    if (visibleSet[name]) {
      try { sheet.showSheet(); shown++; } catch (e) {}
    } else {
      try { sheet.hideSheet(); hidden++; } catch (e) {}
    }
  });

  const firstVisible = visibleNames.map(function(n) { return ss.getSheetByName(n); }).filter(Boolean)[0];
  if (firstVisible) ss.setActiveSheet(firstVisible);

  SpreadsheetApp.getUi().alert(
    '운영 시트만 표시했습니다.\n\n' +
    'v6.05 기준 쿠팡전송수_수동입력은 운영 시트에서 제외했습니다.\n\n' +
    '표시 시트:\n' + visibleNames.filter(function(n) { return !!ss.getSheetByName(n); }).join('\n') + '\n\n' +
    '표시=' + shown + ' / 숨김=' + hidden
  );
}

function checkDeprecatedCoupangSentManualSheetUsage_v605() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(LOTTEON_DEPRECATED_COUPANG_SENT_MANUAL_SHEET_V605);
  const exists = !!sheet;
  const props = PropertiesService.getScriptProperties();
  props.setProperty('CHANGE_MANUAL_DIRTY', 'N');

  SpreadsheetApp.getUi().alert(
    '쿠팡전송수_수동입력 연결 점검 결과\n\n' +
    '시트 존재: ' + (exists ? 'Y' : 'N') + '\n' +
    '현재 전송수 기준: 필터별_상품수 API_totalCount\n' +
    '변경감지 대상: 매출데이터_붙여넣기만 사용\n\n' +
    '결론: 대시보드 전송수 산출 기준에서는 삭제 가능합니다.\n' +
    '단, 삭제 전 하루 정도는 숨김 상태로 운영 확인하는 것을 권장합니다.'
  );
}

// GitHub 코드 연결 테스트 문구 보강용 marker
var LOTTEON_PATCH_V605_CLEANUP_LOADED = true;
