const assert = require('assert');
const fs = require('fs');

const dedicated = fs.readFileSync('FilterCountLightweight_v6_56.gs', 'utf8');
const loader = fs.readFileSync('lotteon_github_remote_loader_v1_8_filter_schedule_menu_restore.gs', 'utf8');
const primaryLoader = fs.readFileSync('loader.gs', 'utf8');
const bootstrap = fs.readFileSync('Patch_v6_24_bootstrap_auto_continue.gs', 'utf8');

for (const source of [loader, primaryLoader]) {
  assert(source.includes('LOTTEON_FILTER_COUNT_LIGHTWEIGHT_URL'));
  assert(source.includes("function runDailyFilterCountsContinue() { return runFilterCountLightweightFunction_('runDailyFilterCountsContinue'); }"));
  assert(source.includes("function runDailyFilterCountsOnceManual() { return runFilterCountLightweightFunction_('runDailyFilterCountsOnceManual'); }"));
  assert(!source.includes("function runDailyFilterCountsOnceManual() { return runRemoteFunction_"));
}
assert(dedicated.includes("FILTER_COUNT_LW_BATCH_SIZE = 10"));
assert(dedicated.includes("FILTER_COUNT_LW_GUARD_MS = 22000"));
assert(dedicated.includes("state.status = 'TRIGGER_PERMISSION_ERROR'"));
assert(dedicated.includes("filterCountLightweightWriteStatus_(state)"));
assert(!dedicated.includes('buildDashboard_'));
assert(!dedicated.includes('generateVatReports'));
assert(!dedicated.includes('loadLotteonRemoteBundle_'));
assert(bootstrap.includes("'Patch_v6_54_vat_validation_control_tower.gs'"));

console.log('v6.56 lightweight route: OK (1 remote file, local continuation handler, full bundle excluded)');

