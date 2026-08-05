/**
 * Permanent remote task slot.
 * Current task: one-time non-destructive operation sheet view cleanup.
 */
const LOTTEON_REMOTE_TASK = {
  id: 'SHEET-VIEW-v1.0-20260805',
  title: '운영 시트 보기 정리',
  enabled: true
};

const LOTTEON_OPERATION_VISIBLE_SHEETS = [
  '기준',
  '대시보드',
  '매출데이터_붙여넣기',
  '소싱 List',
  'LOTTEON_상품목록',
  '필터_대시보드',
  '필터별_상품수',
  '핵심_브랜드요약',
  '브랜드별_마진율',
  '월별_브랜드성과',
  '미정산_쿠팡계정별',
  '쿠팡전송_확인',
  '부가세_신고자료',
  '부가세_기간별',
  '부가세_상품별',
  '부가세_카드매칭검증',
  '카드사용내역_붙여넣기',
  '카드_마스터'
];

function runLotteonRemoteTaskStartRemote_() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');

  const visible = {};
  LOTTEON_OPERATION_VISIBLE_SHEETS.forEach(function(name) { visible[name] = true; });

  const dashboard = ss.getSheetByName('대시보드') || ss.getSheets()[0];
  if (dashboard) dashboard.activate();

  const shown = [];
  const hidden = [];
  const missing = [];

  LOTTEON_OPERATION_VISIBLE_SHEETS.forEach(function(name) {
    if (!ss.getSheetByName(name)) missing.push(name);
  });

  ss.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    if (visible[name]) {
      sheet.showSheet();
      shown.push(name);
    } else {
      sheet.hideSheet();
      hidden.push(name);
    }
  });

  SpreadsheetApp.flush();
  try {
    SpreadsheetApp.getUi().alert(
      '운영 시트 보기 정리 완료',
      '표시: ' + shown.length + '개\n숨김: ' + hidden.length + '개\n삭제: 0개' +
      (missing.length ? '\n\n현재 없는 표시 대상: ' + missing.join(', ') : '') +
      '\n\n복구: LOTTEON 자동화 → 고급/복구 → 시트 복구: 전체 시트 표시',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (e) {}

  return {
    ok: true,
    taskId: LOTTEON_REMOTE_TASK.id,
    shown: shown,
    hidden: hidden,
    missing: missing,
    deleted: 0
  };
}

function runLotteonRemoteTaskContinueRemote_() {
  return runLotteonRemoteTaskStartRemote_();
}
