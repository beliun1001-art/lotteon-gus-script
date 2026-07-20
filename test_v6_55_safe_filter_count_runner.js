const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

const source = fs.readFileSync('Patch_v6_55_safe_filter_count_runner.gs', 'utf8');
const bootstrap = fs.readFileSync('Patch_v6_24_bootstrap_auto_continue.gs', 'utf8');
const operatingLoader = fs.readFileSync('lotteon_github_remote_loader_v1_8_filter_schedule_menu_restore.gs', 'utf8');
assert(source.includes("LOTTEON_FILTER_COUNT_SAFE_JOB_STATE"));
assert(source.includes('LOTTEON_FILTER_COUNT_SAFE_BATCH_SIZE = 10'));
assert(source.includes('LOTTEON_FILTER_COUNT_SAFE_TIMEOUT_MS = 22000'));
assert(source.includes("runDailyFilterCountsOnceManual()"));
assert(source.includes("runDailyFilterCountsStart()"));
assert(source.includes("runDailyFilterCountsContinue()"));
assert(source.includes("fetchProductListTotalCountFast_(filterNames)"));
assert(!source.includes('buildDashboard_('));
assert(!source.includes('generateVatReports'));
assert(!source.includes('autoResize'));
assert(bootstrap.includes("LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.55'"));
assert(bootstrap.includes("'Patch_v6_54_vat_validation_control_tower.gs','Patch_v6_55_safe_filter_count_runner.gs'"));
assert(operatingLoader.includes('필터별_상품수 안전 갱신 초기화'));
assert(operatingLoader.includes("function resetDailyFilterCountsSafeState()"));

const values = [
  ['검색필터명', '브랜드명', '계정번호', '쿠팡계정ID', 'API_totalCount', 'API_totalPage', '이번조회_행수', '필터코드', '최근', '생성', '최근필드', '생성필드', '메모'],
  ['롯백_01_A', 'A', '', '', 3, 1, 0, '', '', '', '', '', 'old'],
  ['롯백_01_B', 'B', '', '', 4, 1, 0, '', '', '', '', '', 'old']
];
const sheet = {
  getLastRow: () => values.length,
  getRange(row, col, numRows, numCols) {
    return {
      getValues: () => values.slice(row - 1, row - 1 + numRows).map(r => r.slice(col - 1, col - 1 + numCols)),
      setValues: rows => rows.forEach((r, i) => { for (let j = 0; j < numCols; j++) values[row - 1 + i][col - 1 + j] = r[j]; })
    };
  }
};
const sandbox = {
  Date,
  JSON,
  String,
  Number,
  Math,
  CONFIG: { SHEETS: { FILTERS: '필터별_상품수' }, HEADERS: { FILTERS: values[0] } },
  getSheet_: () => sheet,
  fetchProductListTotalCountFast_: names => ({
    [names[0]]: { totalCount: 11, totalPage: 2, error: '' },
    [names[1]]: { totalCount: 0, totalPage: 0, error: 'HTTP_500' }
  })
};
vm.createContext(sandbox);
vm.runInContext(source, sandbox);
const state = { currentIndex: 0, totalFilters: 2, processedCount: 0, successCount: 0, errorCount: 0, lastError: '' };
sandbox.runSafeFilterCountTick_v655_(state, Date.now());
assert.equal(state.phase, 'DONE');
assert.equal(state.successCount, 1);
assert.equal(state.errorCount, 1);
assert.equal(values[1][4], 11);
assert.equal(values[1][5], 2);
assert.equal(values[2][4], 4, 'failed filter retains prior API_totalCount');
assert(values[2][12].includes('HTTP_500'));

const propertyStore = {};
const triggerSandbox = {
  Date,
  JSON,
  String,
  Number,
  Math,
  PropertiesService: { getScriptProperties: () => ({
    getProperty: key => propertyStore[key] || null,
    setProperty: (key, value) => { propertyStore[key] = value; }
  }) },
  LockService: { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) },
  now_: () => '2026-07-20T00:00:00Z'
};
vm.createContext(triggerSandbox);
vm.runInContext(source, triggerSandbox);
const triggerState = { version: 'v6.55', phase: 'COUNT', status: 'COUNT_RUNNING', totalFilters: 1, currentIndex: 0, processedCount: 0, successCount: 0, errorCount: 0, lastError: '' };
propertyStore.LOTTEON_FILTER_COUNT_SAFE_JOB_STATE = JSON.stringify(triggerState);
triggerSandbox.runSafeFilterCountTick_v655_ = () => {};
triggerSandbox.scheduleSafeFilterCountContinuation_v655_ = () => ({ scheduled: false, error: '트리거 권한 오류: denied' });
triggerSandbox.deleteSafeFilterCountContinuationTriggers_v655_ = () => 0;
let displayedStatus = null;
triggerSandbox.writeSafeFilterCountStatus_v655_ = state => { displayedStatus = state.status; };
const triggerResult = triggerSandbox.runDailyFilterCountsStep_({ showAlert: false });
assert.equal(triggerResult.state.status, 'TRIGGER_PERMISSION_ERROR');
assert.equal(triggerResult.state.nextExecutionScheduled, 'N');
assert(triggerResult.state.lastError.includes('denied'));
assert.equal(JSON.parse(propertyStore.LOTTEON_FILTER_COUNT_SAFE_JOB_STATE).status, 'TRIGGER_PERMISSION_ERROR');
assert.equal(displayedStatus, 'TRIGGER_PERMISSION_ERROR');
console.log('v6.55 safe filter count runner mock: OK (success=1, error=1, failed count retained)');
console.log('v6.55 continuation permission mock: OK (TRIGGER_PERMISSION_ERROR retained)');

