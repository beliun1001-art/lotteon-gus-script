/** PR #25 v6.66 tracking-payment operating smoke runner. Temporary Apps Script file. */
const LOTTEON_PR25_SMOKE_RAW_BASE = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-24-tracking-payment-primary/';
const LOTTEON_PR25_SMOKE_VERSION = 'v1.15-PR25-SMOKE';

function testPr25TrackingPaymentConnection() {
  const bundle = loadPr25TrackingPaymentBundle_();
  const version = detectPr25TrackingPaymentVersion_(bundle);
  SpreadsheetApp.getUi().alert(
    'PR #25 GitHub 코드 연결 성공\n\n' +
    'Branch Base: ' + LOTTEON_PR25_SMOKE_RAW_BASE + '\n' +
    'Smoke Runner: ' + LOTTEON_PR25_SMOKE_VERSION + '\n' +
    '로드 크기: ' + bundle.length.toLocaleString('ko-KR') + '자\n' +
    '버전 추정: ' + version + '\n\n' +
    'v6.66이 표시되면 runPr25TrackingPaymentVatSmoke를 실행하세요.'
  );
}

function runPr25TrackingPaymentVatSmoke() {
  const bundle = loadPr25TrackingPaymentBundle_();
  return eval(
    bundle +
    '\n\n; if (typeof generateVatReportsFullSeparated_v622 !== "function") {' +
    ' throw new Error("Remote function not found: generateVatReportsFullSeparated_v622"); }' +
    '\n generateVatReportsFullSeparated_v622();'
  );
}

function loadPr25TrackingPaymentBundle_() {
  const codeText = fetchPr25TrackingPaymentText_('Code.gs', 'Code.gs');
  const originalPatchText = fetchPr25TrackingPaymentText_('Patch_v6_24_bootstrap_auto_continue.gs', 'Patch bootstrap');
  const patchText = originalPatchText.replace(
    /var\s+LOTTEON_PATCH_BASE_URL\s*=\s*['"][^'"]+['"]\s*;/,
    "var LOTTEON_PATCH_BASE_URL = '" + LOTTEON_PR25_SMOKE_RAW_BASE + "';"
  );
  if (patchText === originalPatchText) throw new Error('PR #25 bootstrap branch base 교체에 실패했습니다.');
  const bundle = codeText + '\n\n;\n\n' + patchText;
  const version = detectPr25TrackingPaymentVersion_(bundle);
  if (version !== 'v6.66') throw new Error('PR #25 smoke expected v6.66, actual ' + version);
  return bundle;
}

function fetchPr25TrackingPaymentText_(path, label) {
  const url = LOTTEON_PR25_SMOKE_RAW_BASE + path;
  const response = UrlFetchApp.fetch(url + '?ts=' + new Date().getTime(), {
    method:'get', muteHttpExceptions:true, followRedirects:true
  });
  const code = response.getResponseCode();
  const text = response.getContentText('UTF-8');
  if (code < 200 || code >= 300) throw new Error(label + ' 로드 실패 HTTP ' + code + '\n' + url + '\n' + text.slice(0, 500));
  return text;
}

function detectPr25TrackingPaymentVersion_(text) {
  const src = String(text || '');
  const m = src.match(/LOTTEON_PATCH_BOOTSTRAP_VERSION\s*=\s*['"]([^'"]+)['"]/);
  return m ? m[1] : '확인불가';
}
