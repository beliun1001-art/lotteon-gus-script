/**
 * LOTTEON combined patch bootstrap v6.05
 *
 * Apps Script loader는 이 파일만 읽도록 되어 있으므로,
 * 실제 운영 patch를 아래 순서로 동적 로드합니다.
 *
 * 1) v6.04 full patch: 필터별_상품수 자동 갱신 + API_totalCount 전송수 + 숫자 접두어 브랜드 정규화
 * 2) v6.05 cleanup patch: 쿠팡전송수_수동입력 연결 정리/숨김
 *
 * loader 연결 테스트 호환 marker:
 * function startDailyFilterCountsSchedule
 */

var LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.05';

(function loadLotteonCombinedPatches_v605_() {
  const urls = [
    'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/e7ba5932687a74fd1bfe4ec6a2ef154fadf516ca/Patch_v6_01_daily_filter_auto.gs',
    'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_05_cleanup_manual_sheet.gs'
  ];

  urls.forEach(function(url) {
    const response = UrlFetchApp.fetch(url + '?ts=' + new Date().getTime(), {
      method: 'get',
      muteHttpExceptions: true,
      followRedirects: true
    });
    const code = response.getResponseCode();
    const text = response.getContentText('UTF-8');
    if (code < 200 || code >= 300) {
      throw new Error('LOTTEON patch 로드 실패 HTTP ' + code + ': ' + url + '\n' + text.slice(0, 500));
    }
    (0, eval)(text);
  });
})();
