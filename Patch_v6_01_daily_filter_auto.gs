var LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.22';
var LOTTEON_PATCH_BOOTSTRAP_URLS = [
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/e7ba5932687a74fd1bfe4ec6a2ef154fadf516ca/Patch_v6_01_daily_filter_auto.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_05_cleanup_manual_sheet.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_06_dashboard_matchdiag_sync.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_07_dashboard_total_sent_sync.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_08_filter_brand_after_second_underscore.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_09_dashboard_brand_filter_sync.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_10_dashboard_force_brand_filter_sync.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_11_brand_dashboard_log_margin_vat.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_12_profit_logic_settlement_basis.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_13_net_sales_vat_breakdown.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_14_purchase_and_year_format_fix.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_15_dashboard_header_style_auto.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_16_unsettled_estimated_profit.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_17_purchase_lookup_split_fix.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_18_unsettled_sheet_compact_format.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_19_all_sheet_autowidth_dashboard_split.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_20_fast_dashboard_timeout_guard.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_21_vat_input_credit_profit.gs',
  'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_22_fast_dashboard_skip_vat_reports.gs'
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
