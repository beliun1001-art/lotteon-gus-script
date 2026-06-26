/**
 * LOTTEON v6.01 + v6.02 + v6.03 + v6.04 operating patch
 *
 * Loaded after Code.gs by loader.gs.
 *
 * v6.01:
 * - 필터별_상품수 매일 자동 갱신
 *
 * v6.02:
 * - 쿠팡전송수_수동입력의 쿠팡전체전송수 대신
 *   필터별_상품수.API_totalCount를 전송수 기준으로 사용
 *
 * v6.03:
 * - 이미 생성된 핵심_브랜드요약/대시보드의 전송수도
 *   필터별_상품수.API_totalCount로 직접 보정
 *
 * v6.04:
 * - 1K2, 4보브, 5숲, 3플라스틱아일랜드_01 같은 숫자 접두어 필터 브랜드를
 *   K2, 보브, 숲, 플라스틱아일랜드로 정규화해 기존 매출 행에 병합
 */

function startDailyFilterCountsSchedule() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const deleted = deleteDailyFilterCountsTriggers_();

  ScriptApp.newTrigger('runDailyFilterCountsStart')
    .timeBased()
    .atHour(6)
    .nearMinute(10)
    .everyDays(1)
    .create();

  props.setProperty('FILTERLIST_AUTO_ENABLED', 'Y');
  props.setProperty('FILTERLIST_AUTO_RUNNING', 'N');
  props.setProperty('FILTERLIST_AUTO_SCHEDULE', '매일 06:10 전후');
  props.setProperty('FILTERLIST_AUTO_LAST_MESSAGE', '필터별_상품수 매일 자동 갱신 예약 완료');
  props.setProperty('FILTERLIST_AUTO_LAST_ERROR', '');
  props.setProperty('FILTERLIST_AUTO_LAST_UPDATED', now_());

  writeDailyFilterCountsStatus_('SCHEDULED', '필터별_상품수 매일 자동 갱신 예약 완료', '정리한 기존 트리거 수: ' + deleted);
  log_('filterlist_daily_schedule_start_v601', '매일 06:10 전후 자동 갱신 예약 / deleted=' + deleted);

  ui.alert(
    '필터별_상품수 매일 자동 갱신을 예약했습니다.\n\n' +
    '예약 시간: 매일 06:10 전후\n' +
    '실행 방식: filterList 1페이지씩 자동 이어실행\n' +
    '최종 반영 시트: 필터별_상품수\n\n' +
    'v6.04 기준으로 숫자 접두어 필터도 원 브랜드로 병합합니다.'
  );
}

function stopDailyFilterCountsSchedule() {
  const ui = SpreadsheetApp.getUi();
  const props = PropertiesService.getScriptProperties();
  const deleted = deleteDailyFilterCountsTriggers_();

  props.setProperty('FILTERLIST_AUTO_ENABLED', 'N');
  props.setProperty('FILTERLIST_AUTO_RUNNING', 'N');
  props.setProperty('FILTERLIST_AUTO_STAGE', 'STOPPED');
  props.setProperty('FILTERLIST_AUTO_LAST_MESSAGE', '필터별_상품수 자동 갱신 중지');
  props.setProperty('FILTERLIST_AUTO_LAST_ERROR', '');
  props.setProperty('FILTERLIST_AUTO_LAST_UPDATED', now_());

  writeDailyFilterCountsStatus_('STOPPED', '필터별_상품수 자동 갱신 중지', '삭제한 트리거 수: ' + deleted);
  log_('filterlist_daily_schedule_stop_v601', 'deleted=' + deleted);
  ui.alert('필터별_상품수 자동 갱신을 중지했습니다.\n\n삭제한 트리거 수: ' + deleted);
}

function runDailyFilterCountsOnceManual() {
  const ui = SpreadsheetApp.getUi();
  const result = runDailyFilterCountsStep_({ startFresh: true, source: 'MANUAL_NOW', showAlert: true });
  if (result && result.done) return;
  ui.alert(
    '필터별_상품수 자동 이어실행을 시작했습니다.\n\n' +
    '첫 1페이지 처리가 끝났고, 남은 페이지는 약 1분 간격으로 자동 이어실행됩니다.\n' +
    '진행상태는 동기화상태 시트 또는 자동 상태 확인 메뉴에서 볼 수 있습니다.'
  );
}

function runDailyFilterCountsStart() {
  return runDailyFilterCountsStep_({ startFresh: true, source: 'DAILY_START', showAlert: false });
}

function runDailyFilterCountsContinue() {
  return runDailyFilterCountsStep_({ startFresh: false, source: 'DAILY_CONTINUE', showAlert: false });
}

function runDailyFilterCountsStep_(options) {
  options = options || {};
  const props = PropertiesService.getScriptProperties();
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) {
    writeDailyFilterCountsStatus_('LOCKED', '다른 필터별_상품수 갱신 작업 실행 중', '잠시 후 continuation trigger가 다시 실행되거나 수동으로 상태 확인하세요.');
    return { done: false, locked: true };
  }

  const started = Date.now();
  try {
    if (options.startFresh) {
      deleteDailyFilterCountsContinuationTriggers_();
      props.setProperty('FILTERLIST_AUTO_STARTED_AT', now_());
      props.setProperty('FILTERLIST_AUTO_PAGE_RUNS', '0');
    }

    const runCount = toNumber_(props.getProperty('FILTERLIST_AUTO_PAGE_RUNS')) + 1;
    props.setProperty('FILTERLIST_AUTO_PAGE_RUNS', String(runCount));
    props.setProperty('FILTERLIST_AUTO_ENABLED', props.getProperty('FILTERLIST_AUTO_ENABLED') || 'Y');
    props.setProperty('FILTERLIST_AUTO_RUNNING', 'Y');
    props.setProperty('FILTERLIST_AUTO_STAGE', 'RUNNING');
    props.setProperty('FILTERLIST_AUTO_LAST_MESSAGE', '필터별_상품수 자동 갱신 페이지 처리 중');
    props.setProperty('FILTERLIST_AUTO_LAST_ERROR', '');
    props.setProperty('FILTERLIST_AUTO_LAST_UPDATED', now_());

    writeDailyFilterCountsStatus_(
      'RUNNING',
      '필터별_상품수 자동 갱신 페이지 처리 중',
      'source=' + (options.source || '') + ' / pageRun=' + runCount
    );

    const result = runFilterListResumeBatch_({ startFresh: !!options.startFresh });
    const elapsedSec = Math.round((Date.now() - started) / 1000);

    if (result.done) {
      deleteDailyFilterCountsContinuationTriggers_();
      props.setProperty('FILTERLIST_AUTO_RUNNING', 'N');
      props.setProperty('FILTERLIST_AUTO_STAGE', 'DONE');
      props.setProperty('FILTERLIST_AUTO_LAST_DONE_AT', now_());
      props.setProperty('FILTERLIST_AUTO_LAST_MESSAGE', '필터별_상품수 자동 갱신 완료');
      props.setProperty('FILTERLIST_AUTO_LAST_RESULT', '필터수=' + result.filterCount + ' / 페이지실행=' + runCount + ' / 마지막소요초=' + elapsedSec);
      props.setProperty('FILTERLIST_AUTO_LAST_ERROR', '');
      props.setProperty('FILTERLIST_AUTO_LAST_UPDATED', now_());

      writeDailyFilterCountsStatus_(
        'DONE',
        '필터별_상품수 자동 갱신 완료',
        '필터수=' + result.filterCount + ' / 페이지실행=' + runCount + ' / 마지막소요초=' + elapsedSec
      );
      log_('filterlist_daily_auto_done_v601', 'filterCount=' + result.filterCount + ', pageRuns=' + runCount + ', elapsedSec=' + elapsedSec);

      if (options.showAlert) {
        SpreadsheetApp.getUi().alert(
          '필터별_상품수 자동 갱신 완료\n\n' +
          '필터 수: ' + result.filterCount + '개\n' +
          '페이지 실행: ' + runCount + '회\n' +
          '마지막 실행 소요초: ' + elapsedSec + '초\n\n' +
          '대시보드까지 반영하려면 ⑤ 대시보드만 빠른 갱신을 실행하세요.'
        );
      }
      return result;
    }

    props.setProperty('FILTERLIST_AUTO_STAGE', 'WAIT_CONTINUE');
    props.setProperty('FILTERLIST_AUTO_LAST_MESSAGE', '필터별_상품수 자동 갱신 이어실행 대기');
    props.setProperty('FILTERLIST_AUTO_LAST_RESULT', '누적필터수=' + result.filterCount + ' / 다음페이지=' + result.nextPage + ' / 페이지실행=' + runCount);
    props.setProperty('FILTERLIST_AUTO_LAST_UPDATED', now_());

    writeDailyFilterCountsStatus_(
      'WAIT_CONTINUE',
      '필터별_상품수 자동 갱신 이어실행 대기',
      '누적필터수=' + result.filterCount + ' / 다음페이지=' + result.nextPage + ' / 페이지실행=' + runCount
    );

    scheduleDailyFilterCountsContinuation_();
    log_('filterlist_daily_auto_continue_scheduled_v601', 'nextPage=' + result.nextPage + ', filterCount=' + result.filterCount + ', pageRuns=' + runCount);
    return result;
  } catch (e) {
    const msg = String(e && e.message ? e.message : e);
    props.setProperty('FILTERLIST_AUTO_RUNNING', 'N');
    props.setProperty('FILTERLIST_AUTO_STAGE', 'ERROR');
    props.setProperty('FILTERLIST_AUTO_LAST_ERROR', msg);
    props.setProperty('FILTERLIST_AUTO_LAST_MESSAGE', '필터별_상품수 자동 갱신 오류');
    props.setProperty('FILTERLIST_AUTO_LAST_UPDATED', now_());
    writeDailyFilterCountsStatus_('ERROR', '필터별_상품수 자동 갱신 오류', msg);
    log_('filterlist_daily_auto_error_v601', msg);
    if (options.showAlert) SpreadsheetApp.getUi().alert('필터별_상품수 자동 갱신 오류\n\n' + msg);
    throw e;
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

function showDailyFilterCountsStatus() {
  const props = PropertiesService.getScriptProperties();
  writeDailyFilterCountsStatus_(
    props.getProperty('FILTERLIST_AUTO_STAGE') || 'UNKNOWN',
    props.getProperty('FILTERLIST_AUTO_LAST_MESSAGE') || '',
    props.getProperty('FILTERLIST_AUTO_LAST_ERROR') || props.getProperty('FILTERLIST_AUTO_LAST_RESULT') || ''
  );

  SpreadsheetApp.getUi().alert(
    '필터별_상품수 자동 갱신 상태\n\n' +
    '사용여부: ' + (props.getProperty('FILTERLIST_AUTO_ENABLED') || 'N') + '\n' +
    '실행중: ' + (props.getProperty('FILTERLIST_AUTO_RUNNING') || 'N') + '\n' +
    '현재단계: ' + (props.getProperty('FILTERLIST_AUTO_STAGE') || '') + '\n' +
    '예약시간: ' + (props.getProperty('FILTERLIST_AUTO_SCHEDULE') || '매일 06:10 전후') + '\n' +
    '최근시작: ' + (props.getProperty('FILTERLIST_AUTO_STARTED_AT') || '') + '\n' +
    '최근완료: ' + (props.getProperty('FILTERLIST_AUTO_LAST_DONE_AT') || '') + '\n' +
    '페이지실행: ' + (props.getProperty('FILTERLIST_AUTO_PAGE_RUNS') || '') + '\n' +
    '최근결과: ' + (props.getProperty('FILTERLIST_AUTO_LAST_RESULT') || '') + '\n' +
    '최근오류: ' + (props.getProperty('FILTERLIST_AUTO_LAST_ERROR') || '')
  );
}

function scheduleDailyFilterCountsContinuation_() {
  deleteDailyFilterCountsContinuationTriggers_();
  ScriptApp.newTrigger('runDailyFilterCountsContinue')
    .timeBased()
    .after(60 * 1000)
    .create();
}

function deleteDailyFilterCountsTriggers_() {
  return deleteTriggersByHandlerNames_(['runDailyFilterCountsStart', 'runDailyFilterCountsContinue']);
}

function deleteDailyFilterCountsContinuationTriggers_() {
  return deleteTriggersByHandlerNames_(['runDailyFilterCountsContinue']);
}

function writeDailyFilterCountsStatus_(stage, message, memo) {
  const props = PropertiesService.getScriptProperties();
  try {
    writeSyncStatus_({
      phase: stage || props.getProperty('FILTERLIST_AUTO_STAGE') || '',
      status: message || props.getProperty('FILTERLIST_AUTO_LAST_MESSAGE') || '',
      currentFilter: props.getProperty('FILTERLIST_AUTO_SCHEDULE') || '매일 06:10 전후',
      lastUrl: buildApiUrl_('filterList'),
      tmpRows: props.getProperty('FILTERLIST_AUTO_LAST_RESULT') || '',
      lastError: props.getProperty('FILTERLIST_AUTO_LAST_ERROR') || '',
      memo: memo || ''
    });
  } catch (e) {}
}

// -----------------------------------------------------------------------------
// v6.02~v6.04: 쿠팡전송수 기준 변경 및 숫자 접두어 브랜드 정규화
// -----------------------------------------------------------------------------

function readCoupangSentManualMap_() {
  return readCoupangSentCountFromFilterTotalCountMap_();
}

function readCoupangSentCountFromFilterTotalCountMap_() {
  const map = buildFilterApiTotalCountMap_v604_();
  const result = { byBrand: {}, byFilter: {}, byBrandAll: {}, byFilterAll: {}, rows: [] };

  Object.keys(map.byFilter).forEach(function(filterName) {
    const item = map.byFilter[filterName];
    const accountCounts = { beliun1021: 0, beliun1023: 0, beliun1024: 0 };
    if (item.accountId === 'beliun1021') accountCounts.beliun1021 = item.total;
    if (item.accountId === 'beliun1023') accountCounts.beliun1023 = item.total;
    if (item.accountId === 'beliun1024') accountCounts.beliun1024 = item.total;

    const out = {
      confirmDate: item.recentDate || item.createDate || '',
      statusRaw: '필터별_상품수 API_totalCount 기준',
      statusType: 'API_TOTALCOUNT',
      isAdded: false,
      isDeleted: false,
      isMoved: false,
      inactive: false,
      filterName: item.filterName,
      brand: item.canonicalBrand || item.brand,
      total: item.total,
      hasTotal: item.hasTotal,
      mangoCount: item.total,
      hasMangoCount: item.hasTotal,
      accountCounts: accountCounts,
      accountId: item.accountId,
      accountNo: item.accountNo,
      sendDate: '',
      memo: 'v6.04: 필터별_상품수 API_totalCount 사용 / 숫자 접두어 브랜드 정규화'
    };

    result.rows.push(out);
    result.byFilter[item.filterName] = out;
    result.byFilterAll[item.filterName] = out;

    const brandKey = normalizePatchKey_v604_(out.brand);
    if (brandKey) {
      if (!result.byBrand[brandKey]) result.byBrand[brandKey] = out;
      if (!result.byBrandAll[brandKey]) result.byBrandAll[brandKey] = out;
    }

    const mergeKey = normalizeCoupangWorkLogBrandMergeKey_(out.brand, item.filterName);
    if (mergeKey) {
      if (!result.byBrand[mergeKey]) result.byBrand[mergeKey] = out;
      if (!result.byBrandAll[mergeKey]) result.byBrandAll[mergeKey] = out;
    }
  });

  return result;
}

function applyCoupangSentManualToBrandMetrics_(brandMap) {
  const apiMap = readCoupangSentCountFromFilterTotalCountMap_();
  (apiMap.rows || []).forEach(function(item) {
    if (!item || item.inactive) return;
    const brand = item.brand || canonicalBrandFromFilterName_v604_(item.filterName) || '브랜드미확인';
    if (!brandMap[brand]) brandMap[brand] = createBrandMetric_(brand, item.filterName, item.accountNo || '', item.accountId || '');
  });

  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand];
    const filterName = String(b.filterName || '').trim();
    const item = apiMap.byFilter[filterName] || apiMap.byBrand[normalizeCoupangWorkLogBrandMergeKey_(brand, filterName)] || apiMap.byBrand[normalizePatchKey_v604_(brand)];
    if (!item || item.inactive) return;

    if (item.hasTotal) {
      b.coupangSentCount = item.total || 0;
      b.manualCoupangSentCount = 0;
    }
    if (item.hasMangoCount && item.mangoCount > 0) b.productCount = item.mangoCount;
    b.manualCoupangAccountCounts = item.accountCounts || {};
    if (item.filterName) b.filterName = item.filterName;
    if (item.accountId) {
      b.accountId = item.accountId;
      b.accountNo = item.accountNo || b.accountNo;
    }
    b.coupangSentStatus = b.coupangSentCount > 0
      ? '필터별_상품수 API_totalCount 기준 전송수 확인'
      : '필터별_상품수 API_totalCount 0 또는 미확인';
  });
}

var __baseBuildDashboard_v604 = typeof buildDashboard_ === 'function' ? buildDashboard_ : null;
var __baseRefreshDashboardFastOnly_v604 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseRunPendingChangesApproval_v604 = typeof runPendingChangesApproval === 'function' ? runPendingChangesApproval : null;
var __baseRefreshCoreSummaryAndDashboard_v604 = typeof refreshCoreSummaryAndDashboardWithRetransmitLogDates === 'function' ? refreshCoreSummaryAndDashboardWithRetransmitLogDates : null;

if (__baseBuildDashboard_v604) {
  buildDashboard_ = function() {
    patchCoreSummarySentCountsFromFilterApiTotalCount_v604_();
    return __baseBuildDashboard_v604.apply(this, arguments);
  };
}

refreshDashboardFastOnly = function() {
  patchCoreSummarySentCountsFromFilterApiTotalCount_v604_();
  return __baseRefreshDashboardFastOnly_v604 ? __baseRefreshDashboardFastOnly_v604.apply(this, arguments) : null;
};

runPendingChangesApproval = function() {
  const result = __baseRunPendingChangesApproval_v604 ? __baseRunPendingChangesApproval_v604.apply(this, arguments) : null;
  patchCoreSummarySentCountsFromFilterApiTotalCount_v604_();
  return result;
};

refreshCoreSummaryAndDashboardWithRetransmitLogDates = function() {
  const result = __baseRefreshCoreSummaryAndDashboard_v604 ? __baseRefreshCoreSummaryAndDashboard_v604.apply(this, arguments) : null;
  patchCoreSummarySentCountsFromFilterApiTotalCount_v604_();
  if (__baseRefreshDashboardFastOnly_v604) __baseRefreshDashboardFastOnly_v604();
  return result;
};

function patchCoreSummarySentCountsFromFilterApiTotalCount_v604_() {
  const ss = SpreadsheetApp.getActive();
  const summarySheet = ss.getSheetByName(CONFIG.SHEETS.BRAND_SUMMARY);
  if (!summarySheet || summarySheet.getLastRow() < 2) return { updated: 0, reason: 'NO_SUMMARY' };

  const filterMap = buildFilterApiTotalCountMap_v604_();
  const summaryRange = summarySheet.getDataRange();
  const values = summaryRange.getValues();
  if (!values || values.length < 2) return { updated: 0, reason: 'EMPTY_SUMMARY' };

  const header = values[0];
  const colBrand = findHeaderIndex_v604_(header, ['브랜드명']);
  const colFilter = findHeaderIndex_v604_(header, ['대표검색필터명', '대표검색필터']);
  const colSent = findHeaderIndex_v604_(header, ['전송수']);
  const colSalesProduct = findHeaderIndex_v604_(header, ['매출상품수']);
  const colRate = findHeaderIndex_v604_(header, ['매출상품률']);
  const colMemo = findHeaderIndex_v604_(header, ['요약메모']);

  if (colSent < 0) return { updated: 0, reason: 'NO_SENT_COLUMN' };

  let updated = 0;
  for (let r = 1; r < values.length; r++) {
    const brand = colBrand >= 0 ? String(values[r][colBrand] || '').trim() : '';
    const filterName = colFilter >= 0 ? String(values[r][colFilter] || '').trim() : '';
    const exact = filterName ? filterMap.byFilter[filterName] : null;
    const brandItem = brand ? filterMap.byBrand[normalizePatchKey_v604_(brand)] : null;
    const filterCanonicalItem = filterName ? filterMap.byBrand[normalizePatchKey_v604_(canonicalBrandFromFilterName_v604_(filterName))] : null;
    const item = exact || brandItem || filterCanonicalItem;
    if (!item || !item.hasTotal) continue;

    values[r][colSent] = item.total;
    if (colRate >= 0 && colSalesProduct >= 0) {
      const salesProduct = Number(String(values[r][colSalesProduct] || '0').replace(/,/g, '')) || 0;
      values[r][colRate] = item.total > 0 ? salesProduct / item.total : 0;
    }
    if (colMemo >= 0) {
      const oldMemo = String(values[r][colMemo] || '');
      const newMemo = '전송수 API_totalCount 기준';
      if (oldMemo.indexOf(newMemo) < 0) values[r][colMemo] = oldMemo ? oldMemo + ' / ' + newMemo : newMemo;
    }
    updated++;
  }

  summaryRange.setValues(values);
  if (colSent >= 0 && values.length > 1) summarySheet.getRange(2, colSent + 1, values.length - 1, 1).setNumberFormat('#,##0');
  if (colRate >= 0 && values.length > 1) summarySheet.getRange(2, colRate + 1, values.length - 1, 1).setNumberFormat('0.00%');
  log_('patch_summary_sent_count_v604', 'updated=' + updated + ' / source=필터별_상품수.API_totalCount / numberedBrandNormalized=Y');
  return { updated: updated };
}

function buildFilterApiTotalCountMap_v604_() {
  const result = { byFilter: {}, byBrand: {} };
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(CONFIG.SHEETS.FILTERS);
  if (!sheet || sheet.getLastRow() < 2) return result;

  const values = sheet.getDataRange().getValues();
  const header = values[0];
  const colFilter = findHeaderIndex_v604_(header, ['검색필터명']);
  const colBrand = findHeaderIndex_v604_(header, ['브랜드명']);
  const colTotal = findHeaderIndex_v604_(header, ['API_totalCount', 'APItotalCount']);
  const colAccountId = findHeaderIndex_v604_(header, ['쿠팡계정ID']);
  const colAccountNo = findHeaderIndex_v604_(header, ['계정번호']);
  const colRecent = findHeaderIndex_v604_(header, ['API_최근수집일자']);
  const colCreate = findHeaderIndex_v604_(header, ['API_필터생성일']);
  if (colFilter < 0 || colTotal < 0) return result;

  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const filterName = String(row[colFilter] || '').trim();
    if (!filterName || filterName.indexOf(CONFIG.FILTER_PREFIX) !== 0) continue;

    const rawTotal = row[colTotal];
    const hasTotal = String(rawTotal || '').trim() !== '';
    const total = Number(String(rawTotal || '0').replace(/,/g, '')) || 0;
    const rawBrand = colBrand >= 0 ? String(row[colBrand] || '').trim() : '';
    const canonicalBrand = canonicalBrandFromRawBrandOrFilter_v604_(rawBrand, filterName);
    const item = {
      filterName: filterName,
      brand: rawBrand || canonicalBrand,
      canonicalBrand: canonicalBrand,
      total: total,
      hasTotal: hasTotal,
      accountId: colAccountId >= 0 ? String(row[colAccountId] || '').trim() : '',
      accountNo: colAccountNo >= 0 ? String(row[colAccountNo] || '').trim() : '',
      recentDate: colRecent >= 0 ? String(row[colRecent] || '').trim() : '',
      createDate: colCreate >= 0 ? String(row[colCreate] || '').trim() : '',
      isNumbered: isNumberedFilterBrand_v604_(rawBrand, filterName)
    };

    result.byFilter[filterName] = item;

    const brandKey = normalizePatchKey_v604_(canonicalBrand);
    if (brandKey && shouldPreferFilterApiItem_v604_(item, result.byBrand[brandKey])) result.byBrand[brandKey] = item;
  }

  return result;
}

function shouldPreferFilterApiItem_v604_(candidate, current) {
  if (!current) return true;
  if (!!candidate.hasTotal !== !!current.hasTotal) return !!candidate.hasTotal;
  if (!!candidate.isNumbered !== !!current.isNumbered) return !!candidate.isNumbered;
  const cDate = String(candidate.recentDate || candidate.createDate || '');
  const pDate = String(current.recentDate || current.createDate || '');
  if (cDate && pDate && cDate !== pDate) return cDate > pDate;
  if (candidate.total !== current.total) return Number(candidate.total || 0) > Number(current.total || 0);
  return String(candidate.filterName || '') > String(current.filterName || '');
}

function canonicalBrandFromRawBrandOrFilter_v604_(brand, filterName) {
  const fromBrand = canonicalizeBrandText_v604_(brand);
  if (fromBrand) return fromBrand;
  return canonicalBrandFromFilterName_v604_(filterName);
}

function canonicalBrandFromFilterName_v604_(filterName) {
  const text = String(filterName || '').trim();
  if (!text) return '';
  const parts = text.split('_');
  if (parts.length >= 3) return canonicalizeBrandText_v604_(parts.slice(2).join('_'));
  return canonicalizeBrandText_v604_(text.replace(/^롯백_?\d*_?/, ''));
}

function canonicalizeBrandText_v604_(value) {
  let text = String(value || '').trim();
  if (!text) return '';
  text = text.replace(/^\d+/, '');
  text = text.replace(/_\d+$/, '');
  text = text.replace(/\s+/g, ' ').trim();
  return text;
}

function isNumberedFilterBrand_v604_(brand, filterName) {
  const raw = String(brand || '').trim();
  if (/^\d+/.test(raw)) return true;
  const fromFilter = String(filterName || '').split('_').slice(2).join('_');
  return /^\d+/.test(fromFilter);
}

function findHeaderIndex_v604_(headerRow, candidates) {
  const normalized = (headerRow || []).map(function(h) { return normalizeHeaderKey_v604_(h); });
  for (let i = 0; i < candidates.length; i++) {
    const key = normalizeHeaderKey_v604_(candidates[i]);
    const idx = normalized.indexOf(key);
    if (idx >= 0) return idx;
  }
  return -1;
}

function normalizeHeaderKey_v604_(v) {
  return String(v || '').replace(/\s+/g, '').replace(/\n/g, '').replace(/_/g, '').trim();
}

function normalizePatchKey_v604_(v) {
  return canonicalizeBrandText_v604_(v).toLowerCase().replace(/\s+/g, '').replace(/[\[\]\(\)\{\}\-_]/g, '').trim();
}
