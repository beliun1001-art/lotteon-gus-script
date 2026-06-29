var LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.31';
var LOTTEON_PATCH_BASE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/';
var LOTTEON_PATCH_BOOTSTRAP_URLS = [
  'Patch_v6_05_cleanup_manual_sheet.gs',
  'Patch_v6_06_dashboard_matchdiag_sync.gs',
  'Patch_v6_07_dashboard_total_sent_sync.gs',
  'Patch_v6_08_filter_brand_after_second_underscore.gs',
  'Patch_v6_09_dashboard_brand_filter_sync.gs',
  'Patch_v6_10_dashboard_force_brand_filter_sync.gs',
  'Patch_v6_11_brand_dashboard_log_margin_vat.gs',
  'Patch_v6_12_profit_logic_settlement_basis.gs',
  'Patch_v6_13_net_sales_vat_breakdown.gs',
  'Patch_v6_14_purchase_and_year_format_fix.gs',
  'Patch_v6_15_dashboard_header_style_auto.gs',
  'Patch_v6_16_unsettled_estimated_profit.gs',
  'Patch_v6_17_purchase_lookup_split_fix.gs',
  'Patch_v6_18_unsettled_sheet_compact_format.gs',
  'Patch_v6_19_all_sheet_autowidth_dashboard_split.gs',
  'Patch_v6_20_fast_dashboard_timeout_guard.gs',
  'Patch_v6_21_vat_input_credit_profit.gs',
  'Patch_v6_22_fast_dashboard_skip_vat_reports.gs',
  'Patch_v6_23_column_width_batch_display_standard.gs',
  'Patch_v6_24_column_width_auto_continue.gs',
  'Patch_v6_25_dashboard_format_hard_fix.gs',
  'Patch_v6_26_purchase_fallback_for_vat_reports.gs',
  'Patch_v6_27_ac_column_purchase_source_of_truth.gs',
  'Patch_v6_28_single_source_financial_validation.gs',
  'Patch_v6_29_customer_address_unsettled_restore.gs',
  'Patch_v6_30_dashboard_column_width_hard_fix.gs',
  'Patch_v6_31_dashboard_final_alignment_width_fix.gs'
];

for (var LOTTEON_PATCH_BOOTSTRAP_I = 0; LOTTEON_PATCH_BOOTSTRAP_I < LOTTEON_PATCH_BOOTSTRAP_URLS.length; LOTTEON_PATCH_BOOTSTRAP_I++) {
  var LOTTEON_PATCH_BOOTSTRAP_URL = LOTTEON_PATCH_BASE_URL + LOTTEON_PATCH_BOOTSTRAP_URLS[LOTTEON_PATCH_BOOTSTRAP_I];
  var LOTTEON_PATCH_BOOTSTRAP_RESPONSE = UrlFetchApp.fetch(LOTTEON_PATCH_BOOTSTRAP_URL + '?ts=' + new Date().getTime(), { method: 'get', muteHttpExceptions: true, followRedirects: true });
  var LOTTEON_PATCH_BOOTSTRAP_CODE = LOTTEON_PATCH_BOOTSTRAP_RESPONSE.getResponseCode();
  var LOTTEON_PATCH_BOOTSTRAP_TEXT = LOTTEON_PATCH_BOOTSTRAP_RESPONSE.getContentText('UTF-8');
  if (LOTTEON_PATCH_BOOTSTRAP_CODE < 200 || LOTTEON_PATCH_BOOTSTRAP_CODE >= 300) {
    throw new Error('LOTTEON patch load failed HTTP ' + LOTTEON_PATCH_BOOTSTRAP_CODE + ': ' + LOTTEON_PATCH_BOOTSTRAP_URL + '\n' + LOTTEON_PATCH_BOOTSTRAP_TEXT.slice(0, 500));
  }
  eval(LOTTEON_PATCH_BOOTSTRAP_TEXT);
}
