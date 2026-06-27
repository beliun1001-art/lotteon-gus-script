/**
 * LOTTEON combined patch bootstrap v6.05.1
 *
 * Apps Script loader는 이 파일을 Code.gs와 같은 eval 스코프에서 붙여 실행합니다.
 * 따라서 하위 patch도 같은 스코프의 direct eval로 로드해야 now_(), log_(), CONFIG 등을 정상 참조합니다.
 *
 * 로드 순서:
 * 1) v6.04 full patch: 필터별_상품수 자동 갱신 + API_totalCount 전송수 + 숫자 접두어 브랜드 정규화
 * 2) v6.05 cleanup patch: 쿠팡전송수_수동입력 연결 정리/숨김
 *
 * loader 연결 테스트 호환 marker:
 * function startDailyFilterCountsSchedule
 */

var LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.05.1';
var LOTTEON_PATCH_BOOTSTRAP_URLS = [
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/e7ba5932687a74fd1bfe4ec6a2ef154fadf516ca/Patch_v6_01_daily_filter_auto.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_05_cleanup_manual_sheet.gs'
];

for (var LOTTEON_PATCH_BOOTSTRAP_I = 0; LOTTEON_PATCH_BOOTSTRAP_I < LOTTEON_PATCH_BOOTSTRAP_URLS.length; LOTTEON_PATCH_BOOTSTRAP_I++) {
  var LOTTEON_PATCH_BOOTSTRAP_URL = LOTTEON_PATCH_BOOTSTRAP_URLS[LOTTEON_PATCH_BOOTSTRAP_I];
  var LOTTEON_PATCH_BOOTSTRAP_RESPONSE = UrlFetchApp.fetch(LOTTEON_PATCH_BOOTSTRAP_URL + '?ts=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  var LOTTEON_PATCH_BOOTSTRAP_CODE = LOTTEON_PATCH_BOOTSTRAP_RESPONSE.getResponseCode();
  var LOTTEON_PATCH_BOOTSTRAP_TEXT = LOTTEON_PATCH_BOOTSTRAP_RESPONSE.getContentText('UTF-8');
  if (LOTTEON_PATCH_BOOTSTRAP_CODE < 200 || LOTTEON_PATCH_BOOTSTRAP_CODE >= 300) {
    throw new Error('LOTTEON patch 로드 실패 HTTP ' + LOTTEON_PATCH_BOOTSTRAP_CODE + ': ' + LOTTEON_PATCH_BOOTSTRAP_URL + '\n' + LOTTEON_PATCH_BOOTSTRAP_TEXT.slice(0, 500));
  }

  // 중요: indirect eval((0, eval))을 쓰면 Code.gs의 now_(), log_(), CONFIG 스코프가 끊어집니다.
  // 여기서는 loader가 만든 Code.gs와 같은 eval 스코프에 patch 함수를 직접 주입합니다.
  eval(LOTTEON_PATCH_BOOTSTRAP_TEXT);
}
