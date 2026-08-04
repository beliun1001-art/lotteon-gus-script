/**
 * PR15 diagnostic hotfix v1.2.
 * Broad raw scan that does not depend on v6.48 inclusion or period queue matching.
 * Writes only PR15 diagnostic/status sheets.
 */
var PR15_DIAG_VERSION = 'v1.2-ISSUE15-RAW-SCAN-DIAGNOSTIC';

function runPr15FutureAccountDiagnosticV12() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  pr15WriteRawStatusV12_(ss, 'RUNNING', '원본·신고자료 전체 raw scan 시작', {}, '');

  try {
    var source = ss.getSheetByName(PR15_SOURCE_SHEET);
    var detail = ss.getSheetByName(PR15_DETAIL_SHEET);
    if (!source || source.getLastRow() < 2) throw new Error('원본 시트가 없거나 비어 있습니다: ' + PR15_SOURCE_SHEET);
    if (!detail || detail.getLastRow() < 2) throw new Error('부가세 상세 시트가 없거나 비어 있습니다: ' + PR15_DETAIL_SHEET);

    var sourceScan = pr15RawScanSheetV12_(source, '원본');
    var detailScan = pr15RawScanSheetV12_(detail, '신고자료');
    var rows = sourceScan.rows.concat(detailScan.rows);
    rows.sort(function(a, b) {
      return String(a[0]).localeCompare(String(b[0])) || Number(a[1]) - Number(b[1]);
    });

    var headers = [
      '시트구분','행번호','계정ID','날짜값','날짜표시','날짜타입','복원월',
      '주문번호','사업자번호','상품번호','순수매출액','매입금액','탐지근거',
      'root_cause_status','production 수정 필요 여부','행 전체 요약'
    ];
    var diag = ss.getSheetByName(PR15_DIAG_SHEET) || ss.insertSheet(PR15_DIAG_SHEET);
    diag.clear();
    diag.getRange(1, 1, 1, headers.length).setValues([headers]);
    if (rows.length) diag.getRange(2, 1, rows.length, headers.length).setValues(rows);
    diag.setFrozenRows(1);
    diag.getRange(1, 1, 1, headers.length)
      .setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    if (rows.length) {
      diag.getRange(2, 1, rows.length, 10).setNumberFormat('@');
      diag.getRange(2, 11, rows.length, 2).setNumberFormat('#,##0');
      diag.getRange(2, 13, rows.length, 4).setNumberFormat('@');
    }
    [90,80,190,180,120,100,100,170,150,150,110,110,260,200,250,600]
      .forEach(function(width, index) { diag.setColumnWidth(index + 1, width); });

    var stats = {
      sourceRows: source.getLastRow() - 1,
      detailRows: detail.getLastRow() - 1,
      sourceHits: sourceScan.rows.length,
      detailHits: detailScan.rows.length,
      rows: rows.length,
      currentEvidenceNotFound: rows.length ? 0 : 1
    };
    var message = rows.length
      ? 'Issue #15 raw scan 완료: 대상 증거 ' + rows.length + '행 / 운영 원본 변경 0건'
      : 'Issue #15 raw scan 완료: 현재 원본·신고자료에서 대상 계정·2026-10·11 증거 없음 / 운영 원본 변경 0건';
    pr15WriteRawStatusV12_(ss, 'PASS', message, stats, '');
    return {ok:true, rows:rows.length, stats:stats, productionWrites:0};
  } catch (error) {
    var message = String(error && error.message ? error.message : error);
    pr15WriteRawStatusV12_(ss, 'ERROR', 'Issue #15 raw scan 실패', {}, message);
    throw error;
  }
}

function runPr15FutureAccountDiagnosticContinueV12() {
  return runPr15FutureAccountDiagnosticV12();
}

function pr15RawScanSheetV12_(sheet, sheetKind) {
  var range = sheet.getDataRange();
  var values = range.getValues();
  var displays = range.getDisplayValues();
  var headers = values[0] || [];
  var ix = pr15RawIndexesV12_(headers);
  var out = [];
  var today = pr15Today_();

  for (var r = 1; r < values.length; r++) {
    var account = pr15RawAccountV12_(values[r], displays[r], ix.account);
    var rawDate = ix.date >= 0 ? values[r][ix.date] : '';
    var displayDate = ix.date >= 0 ? displays[r][ix.date] : '';
    var date = pr15ParseDate_(rawDate || displayDate);
    var month = pr15RawMonthV12_(values[r], displays[r], ix, date);
    var rowText = displays[r].map(pr15Text_).join(' | ');
    var targetAccount = pr15IsTargetAccount_(account) || pr15RowContainsTargetAccountV12_(displays[r]);
    var targetMonth = month === '2026-10' || month === '2026-11' || /(^|\D)2026[-\/.](10|11)(\D|$)/.test(rowText);
    if (!targetAccount && !targetMonth) continue;

    var reasons = [];
    if (targetAccount) reasons.push('대상 이상 계정 탐지');
    if (targetMonth) reasons.push('2026-10·11 미래월 탐지');
    if (date && pr15IsFutureDate_(date, today)) reasons.push('원본 날짜가 실행일보다 미래');

    var business = pr15Cell_(values[r], ix.business);
    var root = 'UNRESOLVED';
    var action = '추가 원본 증거 필요';
    if (date && pr15IsFutureDate_(date, today)) {
      root = 'SOURCE_DATA_ERROR';
      action = '원본 주문일 정정 여부 확인 필요';
    } else if (targetAccount && !business) {
      root = 'ACCOUNT_MAPPING_REQUIRED';
      action = '신고 대상 계정·사업자 매핑 사용자 확인 필요';
    } else if (targetMonth) {
      root = 'SOURCE_CONFIRMED_VALID';
      action = '현재 원본 날짜가 실제값인지 사용자 확인 필요';
    }

    out.push([
      sheetKind,
      String(r + 1),
      account,
      pr15RawValueText_(rawDate),
      pr15Text_(displayDate),
      pr15ValueType_(rawDate),
      month,
      pr15Cell_(values[r], ix.orderNo),
      business,
      pr15Cell_(values[r], ix.productNo),
      pr15Number_(pr15Cell_(values[r], ix.sales)),
      pr15Number_(pr15Cell_(values[r], ix.purchase)),
      reasons.join(' / '),
      root,
      action,
      rowText.slice(0, 1000)
    ]);
  }
  return {rows:out};
}

function pr15RawIndexesV12_(headers) {
  return {
    account: pr15FindHeader_(headers, ['쿠팡계정ID','마켓아이디','계정ID','마켓ID','아이디']),
    date: pr15FindHeader_(headers, ['마켓주문일자','주문일자','주문일','날짜']),
    month: pr15FindHeader_(headers, ['신고월','월']),
    orderNo: pr15FindHeader_(headers, ['마켓주문번호','주문번호','주문ID','주문ID(마켓)']),
    business: pr15FindHeader_(headers, ['사업자등록번호','사업자번호']),
    productNo: pr15FindHeader_(headers, ['사이트상품번호','마켓상품번호','상품번호']),
    sales: pr15FindHeader_(headers, ['순수매출액','결제금액합계','결제금액']),
    purchase: pr15FindHeader_(headers, ['매입금액','구매가격','구매금액'])
  };
}

function pr15RawAccountV12_(row, displayRow, index) {
  if (index >= 0) return pr15Text_(row[index] || displayRow[index]);
  for (var i = 0; i < displayRow.length; i++) {
    var text = pr15Text_(displayRow[i]).toLowerCase();
    if (PR15_TARGET_ACCOUNTS[text]) return text;
  }
  return '';
}

function pr15RowContainsTargetAccountV12_(row) {
  for (var i = 0; i < row.length; i++) {
    if (PR15_TARGET_ACCOUNTS[pr15Text_(row[i]).toLowerCase()]) return true;
  }
  return false;
}

function pr15RawMonthV12_(row, displayRow, ix, date) {
  if (ix.month >= 0) {
    var monthText = pr15Text_(row[ix.month] || displayRow[ix.month]);
    var direct = monthText.match(/^(20\d{2})-(\d{2})$/);
    if (direct) return direct[1] + '-' + direct[2];
  }
  if (date) return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM');
  var text = ix.date >= 0 ? pr15Text_(displayRow[ix.date]) : '';
  var match = text.match(/^(20\d{2})[-\/.](\d{1,2})/);
  return match ? match[1] + '-' + (Number(match[2]) < 10 ? '0' : '') + Number(match[2]) : '';
}

function pr15WriteRawStatusV12_(ss, status, message, stats, error) {
  var sheet = ss.getSheetByName(PR15_STATUS_SHEET) || ss.insertSheet(PR15_STATUS_SHEET);
  stats = stats || {};
  var rows = [
    ['항목','값'],
    ['버전',PR15_DIAG_VERSION],
    ['상태',status],
    ['메시지',message || ''],
    ['원본 스캔행',stats.sourceRows || 0],
    ['신고자료 스캔행',stats.detailRows || 0],
    ['원본 탐지행',stats.sourceHits || 0],
    ['신고자료 탐지행',stats.detailHits || 0],
    ['전체 진단행',stats.rows || 0],
    ['CURRENT_EVIDENCE_NOT_FOUND',stats.currentEvidenceNotFound || 0],
    ['운영 원본 변경','0건'],
    ['오류',error || ''],
    ['갱신시각',new Date().toISOString()]
  ];
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setColumnWidth(1,280);
  sheet.setColumnWidth(2,700);
}
