/**
 * LOTTEON v6.01 + v6.02 operating patch
 *
 * Loaded after Code.gs by loader.gs.
 *
 * v6.01:
 * - 필터별_상품수 매일 자동 갱신
 *
 * v6.02:
 * - 쿠팡전송수_수동입력의 쿠팡전체전송수 대신
 *   필터별_상품수.API_totalCount를 전송수 기준으로 사용
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
    'v6.02 기준으로 대시보드/핵심요약의 전송수는\n' +
    '쿠팡전송수_수동입력이 아니라 필터별_상품수 API_totalCount를 사용합니다.\n\n' +
    '바로 테스트하려면 같은 메뉴의 "필터별_상품수 자동 갱신 지금 시작"을 실행하세요.'
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
          'v6.02 기준 전송수는 필터별_상품수 API_totalCount를 사용합니다.\n' +
          '대시보드까지 반영하려면 ③ 쿠팡재전송_로그 갱신 → ④ 핵심요약+대시보드 갱신을 실행하세요.'
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
// v6.02: 쿠팡전송수 기준 변경
// -----------------------------------------------------------------------------

function readCoupangSentManualMap_() {
  return readCoupangSentCountFromFilterTotalCountMap_();
}

function readCoupangSentCountFromFilterTotalCountMap_() {
  const result = { byBrand: {}, byFilter: {}, byBrandAll: {}, byFilterAll: {}, rows: [] };
  let rows = [];
  try {
    rows = readTable_(getSheet_(CONFIG.SHEETS.FILTERS));
  } catch (e) {
    return result;
  }

  (rows || []).forEach(function(r) {
    const filterName = String(getObjectValueByHeader_(r, '검색필터명') || '').trim();
    if (!isValidLotteonFilterName_(filterName)) return;

    const brand = String(getObjectValueByHeader_(r, '브랜드명') || '').trim() || brandFromFilterName_(filterName);
    if (!filterName && !brand) return;

    const account = accountFromFilterName_(filterName);
    const accountId = String(getObjectValueByHeader_(r, '쿠팡계정ID') || '').trim() || account.accountId || '';
    const accountNo = String(getObjectValueByHeader_(r, '계정번호') || '').trim() || account.accountNo || '';
    const totalRaw = getObjectValueByHeader_(r, 'API_totalCount');
    const hasTotal = String(totalRaw || '').trim() !== '';
    const total = hasTotal ? toNumber_(totalRaw) : 0;
    const recentDate = normalizeDateText_(getObjectValueByHeader_(r, 'API_최근수집일자'));
    const createDate = normalizeDateText_(getObjectValueByHeader_(r, 'API_필터생성일'));

    const accountCounts = { beliun1021: 0, beliun1023: 0, beliun1024: 0 };
    if (accountId === 'beliun1021') accountCounts.beliun1021 = total;
    if (accountId === 'beliun1023') accountCounts.beliun1023 = total;
    if (accountId === 'beliun1024') accountCounts.beliun1024 = total;

    const item = {
      confirmDate: recentDate || createDate || '',
      statusRaw: '필터별_상품수 API_totalCount 기준',
      statusType: 'API_TOTALCOUNT',
      isAdded: false,
      isDeleted: false,
      isMoved: false,
      inactive: false,
      filterName: filterName,
      brand: brand,
      total: total,
      hasTotal: hasTotal,
      mangoCount: total,
      hasMangoCount: hasTotal,
      accountCounts: accountCounts,
      accountId: accountId,
      accountNo: accountNo,
      sendDate: '',
      memo: 'v6.02: 쿠팡전송수_수동입력 대신 필터별_상품수 API_totalCount 사용'
    };

    result.rows.push(item);
    if (filterName) {
      result.byFilter[filterName] = item;
      result.byFilterAll[filterName] = item;
    }

    const brandKey = normalizeCoupangWorkLogBrandMergeKey_(brand, filterName);
    if (brandKey) {
      if (!result.byBrand[brandKey] || preferApiTotalCountItem_(item, result.byBrand[brandKey])) result.byBrand[brandKey] = item;
      if (!result.byBrandAll[brandKey] || preferApiTotalCountItem_(item, result.byBrandAll[brandKey])) result.byBrandAll[brandKey] = item;
    }

    const normalizedBrandKey = normalizeBrandKey_(brand);
    if (normalizedBrandKey) {
      if (!result.byBrand[normalizedBrandKey] || preferApiTotalCountItem_(item, result.byBrand[normalizedBrandKey])) result.byBrand[normalizedBrandKey] = item;
      if (!result.byBrandAll[normalizedBrandKey] || preferApiTotalCountItem_(item, result.byBrandAll[normalizedBrandKey])) result.byBrandAll[normalizedBrandKey] = item;
    }
  });

  return result;
}

function preferApiTotalCountItem_(candidate, current) {
  if (!current) return true;
  if (!!candidate.hasTotal !== !!current.hasTotal) return !!candidate.hasTotal;
  const cDate = normalizeDateText_(candidate.confirmDate);
  const pDate = normalizeDateText_(current.confirmDate);
  if (cDate && pDate && cDate !== pDate) return cDate > pDate;
  if (candidate.total !== current.total) return toNumber_(candidate.total) > toNumber_(current.total);
  return String(candidate.filterName || '') > String(current.filterName || '');
}

function applyCoupangSentManualToBrandMetrics_(brandMap) {
  const apiMap = readCoupangSentCountFromFilterTotalCountMap_();

  (apiMap.rows || []).forEach(function(item) {
    if (!item || item.inactive) return;
    const brand = item.brand || brandFromFilterName_(item.filterName) || '브랜드미확인';
    if (!brandMap[brand]) brandMap[brand] = createBrandMetric_(brand, item.filterName, item.accountNo || '', item.accountId || '');
  });

  Object.keys(brandMap).forEach(function(brand) {
    const b = brandMap[brand];
    const filterName = String(b.filterName || '').trim();
    const item = apiMap.byFilter[filterName] || apiMap.byBrand[normalizeCoupangWorkLogBrandMergeKey_(brand, filterName)] || apiMap.byBrand[normalizeBrandKey_(brand)];
    if (!item || item.inactive) return;

    if (item.hasTotal) {
      b.coupangSentCount = item.total || 0;
      b.manualCoupangSentCount = 0;
    } else {
      b.coupangSentCount = b.coupangSentCount || 0;
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
