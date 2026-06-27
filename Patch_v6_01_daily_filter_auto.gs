var LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.08';
var LOTTEON_PATCH_BOOTSTRAP_URLS = [
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/e7ba5932687a74fd1bfe4ec6a2ef154fadf516ca/Patch_v6_01_daily_filter_auto.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_05_cleanup_manual_sheet.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_06_dashboard_matchdiag_sync.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_07_dashboard_total_sent_sync.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_08_filter_brand_after_second_underscore.gs'
];

// loader compatibility marker: function startDailyFilterCountsSchedule
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
    throw new Error('LOTTEON patch load failed HTTP ' + LOTTEON_PATCH_BOOTSTRAP_CODE + ': ' + LOTTEON_PATCH_BOOTSTRAP_URL + '\n' + LOTTEON_PATCH_BOOTSTRAP_TEXT.slice(0, 500));
  }
  eval(LOTTEON_PATCH_BOOTSTRAP_TEXT);
}
