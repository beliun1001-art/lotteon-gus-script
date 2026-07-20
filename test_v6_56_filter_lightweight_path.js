const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

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
assert(dedicated.includes('filterCountLightweightExistingCounts_()'));
assert(!dedicated.includes('buildDashboard_'));
assert(!dedicated.includes('generateVatReports'));
assert(!dedicated.includes('loadLotteonRemoteBundle_'));
assert(bootstrap.includes("'Patch_v6_54_vat_validation_control_tower.gs'"));

const sandbox = { String, Number, Date, JSON, Math };
vm.createContext(sandbox);
vm.runInContext(dedicated, sandbox);
const expectedAccounts = {
  '01': [1, 'beliun1021'],
  '02': [2, 'beliun1024'],
  '03': [3, 'beliun1023'],
  '04': [4, 'beliun1024']
};
Object.entries(expectedAccounts).forEach(([code, expected]) => {
  const row = sandbox.filterCountLightweightRow_({ filterName: `롯백_${code}_브랜드`, itemCount: 1 });
  assert.deepStrictEqual(Array.from(row.slice(2, 4)), expected, `롯백_${code} 계정 매핑`);
});

const preserved = [10, 3];
const rows = [['롯백_01_브랜드', '브랜드', 1, 'beliun1021', ...preserved, 0, '01', '', '', '', '', 'old']];
const sheet = {
  getLastRow: () => 2,
  getRange: () => ({ getValues: () => rows.map(row => row.slice()), setValues: values => { rows[0] = values[0]; } })
};
sandbox.filterCountLightweightSheet_ = () => sheet;
sandbox.filterCountLightweightApi_ = () => { throw new Error('HTTP_500'); };
const state = { currentIndex: 0, totalFilters: 1, processedCount: 0, successCount: 0, errorCount: 0 };
sandbox.filterCountLightweightCount_(state, Date.now());
assert.deepStrictEqual(rows[0].slice(4, 6), preserved, 'productList 실패 시 기존 count/page 유지');
assert.equal(state.errorCount, 1);

console.log('v6.56 lightweight route: OK (1 remote file, local continuation handler, full bundle excluded)');
console.log('v6.56 account mapping and failed productList retention: OK');

