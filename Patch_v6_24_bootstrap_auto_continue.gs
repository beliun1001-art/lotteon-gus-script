var LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.44';
var LOTTEON_PATCH_BASE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/';
var LOTTEON_PATCH_BOOTSTRAP_URLS = [
'Patch_v6_05_cleanup_manual_sheet.gs','Patch_v6_06_dashboard_matchdiag_sync.gs','Patch_v6_07_dashboard_total_sent_sync.gs','Patch_v6_08_filter_brand_after_second_underscore.gs','Patch_v6_09_dashboard_brand_filter_sync.gs','Patch_v6_10_dashboard_force_brand_filter_sync.gs','Patch_v6_11_brand_dashboard_log_margin_vat.gs','Patch_v6_12_profit_logic_settlement_basis.gs','Patch_v6_13_net_sales_vat_breakdown.gs','Patch_v6_14_purchase_and_year_format_fix.gs','Patch_v6_15_dashboard_header_style_auto.gs','Patch_v6_16_unsettled_estimated_profit.gs','Patch_v6_17_purchase_lookup_split_fix.gs','Patch_v6_18_unsettled_sheet_compact_format.gs','Patch_v6_19_all_sheet_autowidth_dashboard_split.gs','Patch_v6_20_fast_dashboard_timeout_guard.gs','Patch_v6_21_vat_input_credit_profit.gs','Patch_v6_22_fast_dashboard_skip_vat_reports.gs','Patch_v6_23_column_width_batch_display_standard.gs','Patch_v6_24_column_width_auto_continue.gs','Patch_v6_25_dashboard_format_hard_fix.gs','Patch_v6_26_purchase_fallback_for_vat_reports.gs','Patch_v6_27_ac_column_purchase_source_of_truth.gs','Patch_v6_28_single_source_financial_validation.gs','Patch_v6_29_customer_address_unsettled_restore.gs','Patch_v6_30_dashboard_column_width_hard_fix.gs','Patch_v6_31_dashboard_final_alignment_width_fix.gs','Patch_v6_32_profit_rate_recalculate_fix.gs','Patch_v6_33_percent_one_decimal_all_sheets.gs','Patch_v6_34_brand_margin_rate_hard_fix.gs','Patch_v6_35_unsettled_order_date_asc.gs','Patch_v6_36_unsettled_20day_simple.gs','Patch_v6_37_account_summary_and_vat_account_field.gs','Patch_v6_38_market_id_col_d_account_source.gs','Patch_v6_39_fast_operation_sheet_visibility.gs','Patch_v6_40_business_registration_no_vat_field.gs','Patch_v6_41_fast_vat_single_pass_no_timeout.gs','Patch_v6_42_vat_auto_continue_steps.gs','Patch_v6_43_vat_launcher_status_first.gs','Patch_v6_44_vat_detail_chunked_batch.gs'
];
for (var i = 0; i < LOTTEON_PATCH_BOOTSTRAP_URLS.length; i++) {
  var url = LOTTEON_PATCH_BASE_URL + LOTTEON_PATCH_BOOTSTRAP_URLS[i];
  var res = UrlFetchApp.fetch(url + '?ts=' + new Date().getTime(), { method:'get', muteHttpExceptions:true, followRedirects:true });
  var code = res.getResponseCode();
  var text = res.getContentText('UTF-8');
  if (code < 200 || code >= 300) throw new Error('LOTTEON patch load failed HTTP ' + code + ': ' + url + '\n' + text.slice(0, 500));
  eval(text);
}
