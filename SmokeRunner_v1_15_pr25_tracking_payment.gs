/** PR #25 v6.66 tracking-payment operating smoke runner. Temporary Apps Script file. */
const LOTTEON_PR25_SMOKE_RAW_BASE = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-24-tracking-payment-primary/';
const LOTTEON_PR25_SMOKE_VERSION = 'v1.15-PR25-SMOKE-R2';

function testPr25TrackingPaymentConnection() {
  const bundle = loadPr25TrackingPaymentBundle_();
  const version = detectPr25TrackingPaymentVersion_(bundle);
  SpreadsheetApp.getUi().alert(
    'PR #25 GitHub 코드 연결 성공\n\n' +
    'Branch Base: ' + LOTTEON_PR25_SMOKE_RAW_BASE + '\n' +
    'Smoke Runner: ' + LOTTEON_PR25_SMOKE_VERSION + '\n' +
    '로드 크기: ' + bundle.length.toLocaleString('ko-KR') + '자\n' +
    '버전 추정: ' + version + '\n\n' +
    'v6.66이 표시되면 실행 가능합니다.'
  );
}

/**
 * Full VAT rebuild. This can exceed the Apps Script 6-minute limit on large workbooks.
 * After a timeout, do not rerun this function. Use resumePr25TrackingPaymentCardMatch instead.
 */
function runPr25TrackingPaymentVatSmoke() {
  const bundle = loadPr25TrackingPaymentBundle_();
  return eval(
    bundle +
    '\n\n; if (typeof generateVatReportsFullSeparated_v622 !== "function") {' +
    ' throw new Error("Remote function not found: generateVatReportsFullSeparated_v622"); }' +
    '\n generateVatReportsFullSeparated_v622();'
  );
}

/**
 * Lightweight resume path after the full rebuild timed out.
 * Uses the already-written 부가세_신고자료, rebuilds only 부가세_기간별,
 * then runs v6.66 purchase-card reconciliation.
 */
function resumePr25TrackingPaymentCardMatch() {
  const coverage = inspectPr25TrackingPaymentCoverage_();
  if (coverage.paymentHeaderIndex < 0) {
    throw new Error('부가세_신고자료에서 롯데결제수단 열을 찾지 못했습니다. 전체 실행이 상세 시트 생성 전 종료되었습니다.');
  }
  if (coverage.nonblankPayments < 1) {
    throw new Error('롯데결제수단 값이 아직 0건입니다. 전체 실행이 트래킹 번호 반영 전에 종료되었습니다.');
  }

  const bundle = loadPr25TrackingPaymentBundle_();
  const result = eval(
    bundle +
    '\n\n; (function(){' +
    ' var ss = SpreadsheetApp.getActive();' +
    ' if (typeof buildVatPeriodSummary_v657_ !== "function") {' +
    '   throw new Error("Remote function not found: buildVatPeriodSummary_v657_");' +
    ' }' +
    ' return buildVatPeriodSummary_v657_(ss);' +
    ' })();'
  );

  SpreadsheetApp.getUi().alert(
    'PR #25 경량 이어실행 완료\n\n' +
    '부가세 상세 데이터행: ' + coverage.detailRows.toLocaleString('ko-KR') + '건\n' +
    '롯데결제수단 입력행: ' + coverage.nonblankPayments.toLocaleString('ko-KR') + '건\n' +
    '기간별 집계행: ' + String(result && result.rows != null ? result.rows : '확인 필요') + '\n\n' +
    '이제 사업자별 반기 신고요약과 부가세_카드매칭검증 시트를 확인하세요.'
  );
  return result;
}

function inspectPr25TrackingPaymentState() {
  const coverage = inspectPr25TrackingPaymentCoverage_();
  SpreadsheetApp.getUi().alert(
    'PR #25 현재 생성 상태\n\n' +
    '부가세 상세 데이터행: ' + coverage.detailRows.toLocaleString('ko-KR') + '건\n' +
    '롯데결제수단 열: ' + (coverage.paymentHeaderIndex >= 0 ? '있음' : '없음') + '\n' +
    '롯데결제수단 입력행: ' + coverage.nonblankPayments.toLocaleString('ko-KR') + '건'
  );
  return coverage;
}

function inspectPr25TrackingPaymentCoverage_() {
  const ss = SpreadsheetApp.getActive();
  const detail = ss.getSheetByName('부가세_신고자료');
  if (!detail || detail.getLastRow() < 2) {
    return { detailRows: 0, paymentHeaderIndex: -1, nonblankPayments: 0 };
  }
  const values = detail.getDataRange().getDisplayValues();
  const headers = values[0] || [];
  let paymentHeaderIndex = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').replace(/\s/g, '') === '롯데결제수단') {
      paymentHeaderIndex = i;
      break;
    }
  }
  let nonblankPayments = 0;
  if (paymentHeaderIndex >= 0) {
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][paymentHeaderIndex] || '').trim()) nonblankPayments++;
    }
  }
  return {
    detailRows: Math.max(values.length - 1, 0),
    paymentHeaderIndex: paymentHeaderIndex,
    nonblankPayments: nonblankPayments
  };
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
