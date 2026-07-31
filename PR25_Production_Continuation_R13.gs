/**
 * PR #25 production-path operating smoke R13.
 *
 * Resumes the existing v6.48 VAT continuation state without resetting written rows.
 * The repository branch is downloaded once as a ZIP, the production bundle is
 * assembled locally, gzip-compressed, and cached in chunks for later triggers.
 * This avoids re-fetching dozens of raw GitHub patch files on every continuation.
 */
const PR25_R13_VERSION = 'v1.24-PR25-CACHED-PRODUCTION-RESUME-R13';
const PR25_R13_HANDLER = 'continuePr25ProductionVatR13';
const PR25_R13_STATUS_SHEET = 'PR25_실행상태';
const PR25_R13_VAT_STATE_KEY = 'LOTTEON_V648_LIGHT_VAT_JOB_STATE';
const PR25_R13_CACHE_PREFIX = 'PR25_R13_BUNDLE_V1_';
const PR25_R13_CACHE_MANIFEST = PR25_R13_CACHE_PREFIX + 'MANIFEST';
const PR25_R13_CACHE_CHUNK_CHARS = 80000;
const PR25_R13_CACHE_TTL_SECONDS = 21600;
const PR25_R13_ZIP_URL = 'https://codeload.github.com/beliun1001-art/lotteon-gus-script/zip/refs/heads/codex/issue-24-tracking-payment-primary';

/** Resume the existing R12 state. Do not reset or recreate prior detail rows. */
function resumePr25ProductionVatR13() {
  var state = pr25r13_getVatState_();
  var ss = pr25r13_resolveSpreadsheet_(state);
  pr25r13_assertResumeSafe_(ss, state);
  pr25r13_clearTriggers_();
  pr25r13_writeStatus_(ss, state, '기존 production 상태에서 캐시형 이어실행 시작', 'running', 'preparing');
  return pr25r13_runProduction_('resume');
}

/** Time-trigger continuation entrypoint. */
function continuePr25ProductionVatR13() {
  return pr25r13_runProduction_('continue');
}

/** Optional cache reset. Does not modify VAT sheets or VAT continuation state. */
function resetPr25ProductionBundleCacheR13() {
  pr25r13_clearBundleCache_();
  var state = pr25r13_getVatState_();
  var ss = pr25r13_resolveSpreadsheet_(state);
  pr25r13_writeStatus_(ss, state, 'R13 production bundle 캐시만 초기화', state && state.status || '', 'cache_reset');
  return { ok:true };
}

function pr25r13_runProduction_(mode) {
  var stateBefore = pr25r13_getVatState_();
  var ss = pr25r13_resolveSpreadsheet_(stateBefore);
  try {
    if (!stateBefore) throw new Error('이어실행할 부가세 production 상태가 없습니다. R12 결과를 유지한 상태에서 실행하세요.');
    var bundleInfo = pr25r13_getOrPrimeBundle_();
    var invocation = [
      "LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.67-PR25-R13';",
      "LOTTEON_V648_HANDLER = '" + PR25_R13_HANDLER + "';",
      'generateVatReportsFullSeparated_v622();'
    ].join('\n');
    var result = eval(bundleInfo.bundle + '\n\n;\n\n' + invocation);
    var state = pr25r13_getVatState_();
    ss = pr25r13_resolveSpreadsheet_(state || stateBefore);
    var message = state && state.status === 'done'
      ? 'production VAT 이어실행 완료'
      : (mode === 'resume' ? '저장 지점에서 이어실행 및 다음 배치 예약' : '캐시 bundle로 자동 이어실행 진행');
    pr25r13_writeStatus_(ss, state, message, state && state.status || 'running', bundleInfo.source);
    if (state && state.status === 'done') pr25r13_clearTriggers_();
    return result;
  } catch (e) {
    var stateAfter = pr25r13_getVatState_() || stateBefore || {};
    ss = pr25r13_resolveSpreadsheet_(stateAfter);
    pr25r13_writeStatus_(ss, stateAfter, 'R13 production 이어실행 오류', 'failed', 'error', String(e && e.message ? e.message : e));
    pr25r13_clearTriggers_();
    throw e;
  }
}

function pr25r13_getOrPrimeBundle_() {
  var cached = pr25r13_readBundleCache_();
  if (cached) return { bundle:cached, source:'cache' };
  var bundle = pr25r13_buildBundleFromZip_();
  pr25r13_writeBundleCache_(bundle);
  return { bundle:bundle, source:'branch_zip_1_fetch' };
}

function pr25r13_buildBundleFromZip_() {
  var response = UrlFetchApp.fetch(PR25_R13_ZIP_URL + '?ts=' + new Date().getTime(), {
    method:'get', muteHttpExceptions:true, followRedirects:true
  });
  var code = response.getResponseCode();
  if (code < 200 || code >= 300) {
    throw new Error('PR25 branch ZIP 로드 실패 HTTP ' + code + ': ' + response.getContentText('UTF-8').slice(0, 500));
  }
  var blobs = Utilities.unzip(response.getBlob());
  var files = {};
  for (var i = 0; i < blobs.length; i++) {
    var name = String(blobs[i].getName() || '').replace(/\\/g, '/');
    var base = name.substring(name.lastIndexOf('/') + 1);
    if (!base || files[base] || !/\.gs$/i.test(base)) continue;
    files[base] = blobs[i].getDataAsString('UTF-8');
  }
  if (!files['Code.gs']) throw new Error('branch ZIP에서 Code.gs를 찾지 못했습니다.');
  var bootstrap = files['Patch_v6_24_bootstrap_auto_continue.gs'];
  if (!bootstrap) throw new Error('branch ZIP에서 patch bootstrap을 찾지 못했습니다.');
  var patchNames = pr25r13_parseBootstrapPatchNames_(bootstrap);
  if (!patchNames.length) throw new Error('patch bootstrap에서 로드 파일 목록을 찾지 못했습니다.');
  var parts = [files['Code.gs']];
  for (var p = 0; p < patchNames.length; p++) {
    var patch = files[patchNames[p]];
    if (typeof patch !== 'string') throw new Error('branch ZIP에 patch 파일이 없습니다: ' + patchNames[p]);
    parts.push(patch);
  }
  var bundle = parts.join('\n\n;\n\n');
  if (bundle.indexOf('LOTTEON_PATCH_V666_VAT_TRACKING_PAYMENT_PRIMARY_LOADED') < 0 ||
      bundle.indexOf('LOTTEON_PATCH_V667_VAT_TRACKING_PRODUCTION_PATH_FIX_LOADED') < 0) {
    throw new Error('조립된 production bundle에 v6.66/v6.67이 포함되지 않았습니다.');
  }
  return bundle;
}

function pr25r13_parseBootstrapPatchNames_(bootstrap) {
  var text = String(bootstrap || '');
  var block = text.match(/LOTTEON_PATCH_BOOTSTRAP_URLS\s*=\s*\[([\s\S]*?)\];/);
  if (!block) return [];
  var out = [], seen = {}, re = /['"]([^'"]+\.gs)['"]/g, match;
  while ((match = re.exec(block[1])) !== null) {
    var name = match[1];
    if (!seen[name]) { seen[name] = true; out.push(name); }
  }
  return out;
}

function pr25r13_writeBundleCache_(bundle) {
  pr25r13_clearBundleCache_();
  var gzip = Utilities.gzip(Utilities.newBlob(String(bundle || ''), 'text/plain', 'pr25-r13-bundle.js'));
  var encoded = Utilities.base64Encode(gzip.getBytes());
  var count = Math.ceil(encoded.length / PR25_R13_CACHE_CHUNK_CHARS);
  var cache = CacheService.getScriptCache();
  for (var i = 0; i < count; i++) {
    cache.put(PR25_R13_CACHE_PREFIX + 'CHUNK_' + i,
      encoded.substring(i * PR25_R13_CACHE_CHUNK_CHARS, (i + 1) * PR25_R13_CACHE_CHUNK_CHARS),
      PR25_R13_CACHE_TTL_SECONDS);
  }
  cache.put(PR25_R13_CACHE_MANIFEST, JSON.stringify({
    version:PR25_R13_VERSION,
    chunks:count,
    encodedLength:encoded.length,
    sourceLength:String(bundle || '').length,
    createdAt:new Date().toISOString()
  }), PR25_R13_CACHE_TTL_SECONDS);
}

function pr25r13_readBundleCache_() {
  var cache = CacheService.getScriptCache();
  var raw = cache.get(PR25_R13_CACHE_MANIFEST);
  if (!raw) return '';
  var manifest;
  try { manifest = JSON.parse(raw); } catch (e) { return ''; }
  if (!manifest || !manifest.chunks) return '';
  var encoded = '';
  for (var i = 0; i < Number(manifest.chunks); i++) {
    var chunk = cache.get(PR25_R13_CACHE_PREFIX + 'CHUNK_' + i);
    if (!chunk) return '';
    encoded += chunk;
  }
  try {
    var bytes = Utilities.base64Decode(encoded);
    return Utilities.ungzip(Utilities.newBlob(bytes)).getDataAsString('UTF-8');
  } catch (e) {
    return '';
  }
}

function pr25r13_clearBundleCache_() {
  var cache = CacheService.getScriptCache();
  var raw = cache.get(PR25_R13_CACHE_MANIFEST);
  var count = 30;
  try { var m = JSON.parse(raw || '{}'); if (m.chunks) count = Math.max(count, Number(m.chunks)); } catch (e) {}
  cache.remove(PR25_R13_CACHE_MANIFEST);
  for (var i = 0; i < count; i++) cache.remove(PR25_R13_CACHE_PREFIX + 'CHUNK_' + i);
}

function pr25r13_assertResumeSafe_(ss, state) {
  if (!state) throw new Error('이어실행할 VAT 상태가 없습니다.');
  if (state.status === 'done') throw new Error('VAT 작업이 이미 완료 상태입니다.');
  if (state.phase !== 'detail' && state.phase !== 'summaries') {
    throw new Error('이어실행할 수 없는 VAT 단계입니다: ' + String(state.phase || ''));
  }
  var detail = ss && ss.getSheetByName && ss.getSheetByName('부가세_신고자료');
  if (!detail) throw new Error('부가세_신고자료 시트를 찾지 못했습니다.');
  var actualRows = Math.max(detail.getLastRow() - 1, 0);
  var expectedRows = Number(state.writtenRows || 0);
  if (actualRows !== expectedRows) {
    throw new Error('이어실행 안전검증 실패: 상태 작성행 ' + expectedRows + ' / 실제 상세행 ' + actualRows + '. 아무 작업도 시작하지 않았습니다.');
  }
  if (Number(state.sourceRow || 0) < 2 || Number(state.sourceLastRow || 0) < Number(state.sourceRow || 0) - 1) {
    throw new Error('이어실행 상태의 원본 진행행이 올바르지 않습니다.');
  }
}

function pr25r13_getVatState_() {
  var raw = PropertiesService.getScriptProperties().getProperty(PR25_R13_VAT_STATE_KEY);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (e) { return null; }
}

function pr25r13_resolveSpreadsheet_(state) {
  state = state || pr25r13_getVatState_();
  if (state && state.spreadsheetId) return SpreadsheetApp.openById(state.spreadsheetId);
  return SpreadsheetApp.getActive();
}

function pr25r13_clearTriggers_() {
  var handlers = {
    generateVatReportsFullSeparated_v622:true,
    continuePr25ProductionVatR12:true,
    continuePr25ProductionVatR13:true,
    continuePr25TrackingContinuationR6:true,
    continuePr25TrackingContinuationR7:true,
    continuePr25TrackingContinuationR8:true,
    continuePr25TrackingContinuationR9:true,
    continuePr25TrackingContinuationR11:true
  };
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    try { if (handlers[trigger.getHandlerFunction()]) ScriptApp.deleteTrigger(trigger); } catch (e) {}
  });
}

function pr25r13_writeStatus_(ss, state, message, forcedStatus, bundleSource, forcedError) {
  state = state || {};
  var sheet = ss.getSheetByName(PR25_R13_STATUS_SHEET) || ss.insertSheet(PR25_R13_STATUS_SHEET);
  var rows = [
    ['항목','값'],
    ['버전',PR25_R13_VERSION],
    ['실행 경로','branch ZIP 1회 조립 + gzip ScriptCache + production continuation'],
    ['bundle 출처',bundleSource || ''],
    ['상태',forcedStatus || state.status || ''],
    ['단계',state.phase || ''],
    ['메시지',message || ''],
    ['다음 실행 예약',state.nextRunScheduled ? 'Y' : 'N'],
    ['시작시각',state.startedAt || ''],
    ['갱신시각',state.updatedAt || new Date().toISOString()],
    ['원본 진행행',Number(state.sourceRow || 0)],
    ['원본 마지막행',Number(state.sourceLastRow || 0)],
    ['원본 읽기 열수',Number(state.sourceReadColumns || 0)],
    ['트래킹 열번호',Number(state.trackingPaymentColumn || 0)],
    ['fallback 열번호',Number(state.fallbackPaymentColumn || 0)],
    ['부가세 상세 작성행',Number(state.writtenRows || 0)],
    ['제외행',Number(state.skippedRows || 0)],
    ['계정미확인행',Number(state.accountMissingRows || 0)],
    ['마지막 오류',forcedError || state.lastError || '']
  ];
  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.getRange(1, 1, 1, 2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 210);
  sheet.setColumnWidth(2, 700);
  SpreadsheetApp.flush();
}
