/** PR #25 v6.66 tracking-payment lightweight operating smoke runner R3. */
const LOTTEON_PR25_SMOKE_RAW_BASE = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-24-tracking-payment-primary/';
const LOTTEON_PR25_SMOKE_VERSION = 'v1.15-PR25-SMOKE-R3';
const LOTTEON_PR25_LIGHT_FILES = [
  'Patch_v6_48_lightweight_vat_control_tower.gs',
  'Patch_v6_57_vat_period_summary.gs',
  'Patch_v6_60_vat_business_half_summary.gs',
  'Patch_v6_60_vat_card_input_runtime_fix.gs',
  'Patch_v6_61_vat_card_date_object_fix.gs',
  'Patch_v6_62_vat_card_unmatched_diagnostic.gs',
  'Patch_v6_63_vat_card_forward_allocation.gs',
  'Patch_v6_64_vat_card_canonical_cancel_safe.gs',
  'Patch_v6_65_vat_card_text_safe_output.gs',
  'Patch_v6_66_vat_tracking_payment_primary.gs'
];

function testPr25TrackingPaymentConnection() {
  const bundle = loadPr25TrackingPaymentLightBundle_();
  SpreadsheetApp.getUi().alert(
    'PR #25 경량 코드 연결 성공\n\n' +
    'Smoke Runner: ' + LOTTEON_PR25_SMOKE_VERSION + '\n' +
    'Branch Base: ' + LOTTEON_PR25_SMOKE_RAW_BASE + '\n' +
    '로드 파일: ' + LOTTEON_PR25_LIGHT_FILES.length + '개\n' +
    '로드 크기: ' + bundle.length.toLocaleString('ko-KR') + '자\n' +
    '대상 버전: v6.66'
  );
}

function backfillPr25TrackingPaymentAndMatch() {
  const bundle = loadPr25TrackingPaymentLightBundle_();
  const result = eval(
    bundle +
    '\n\n;(function(){' +
    ' var ss = SpreadsheetApp.getActive();' +
    ' var source = ss.getSheetByName("매출데이터_붙여넣기");' +
    ' var detail = ss.getSheetByName("부가세_신고자료");' +
    ' if (!source || source.getLastRow() < 2) throw new Error("매출데이터_붙여넣기 시트를 찾지 못했거나 데이터가 없습니다.");' +
    ' if (!detail || detail.getLastRow() < 2) throw new Error("부가세_신고자료 시트를 찾지 못했거나 데이터가 없습니다.");' +
    ' function normText(v){ return String(v == null ? "" : v).trim().replace(/\\s+/g," "); }' +
    ' function normNumber(v){ var n=Number(String(v == null ? 0 : v).replace(/[원,\\s]/g,"")); return isNaN(n)?0:Math.round(n); }' +
    ' function findIx(headers,names){ for(var n=0;n<names.length;n++){ var target=String(names[n]).replace(/\\s/g,""); for(var i=0;i<headers.length;i++) if(String(headers[i]||"").replace(/\\s/g,"")===target) return i; } return -1; }' +
    ' function keyFrom(row, ix){ return [normText(row[ix.date]),normText(row[ix.account]),normText(row[ix.business]),normText(row[ix.orderNo]),normText(row[ix.customer]),normText(row[ix.brand]),normText(row[ix.productNo]),normText(row[ix.productName]),normNumber(row[ix.quantity]),normNumber(row[ix.sales]),normNumber(row[ix.purchase])].join("|"); }' +
    ' function indexes(headers){ return {date:findIx(headers,["날짜","주문일","주문일자","마켓주문일자"]),account:findIx(headers,["쿠팡계정ID"]),business:findIx(headers,["사업자등록번호"]),orderNo:findIx(headers,["주문번호","마켓주문번호","주문ID","주문ID(마켓)"]),customer:findIx(headers,["고객명"]),brand:findIx(headers,["브랜드명"]),productNo:findIx(headers,["상품번호","마켓상품번호"]),productName:findIx(headers,["상품명"]),quantity:findIx(headers,["판매수량","수량"]),sales:findIx(headers,["순수매출액"]),purchase:findIx(headers,["매입금액"]),payment:findIx(headers,["롯데결제수단"])}; }' +
    ' function requireIndexes(ix,label){ Object.keys(ix).forEach(function(k){ if(ix[k] < 0) throw new Error(label+" 필수 열 누락: "+k); }); }' +
    ' var sourceMaxCol=Math.min(LOTTEON_V648_MAX_COL || 29, source.getLastColumn());' +
    ' var sourceHeaders=source.getRange(1,1,1,sourceMaxCol).getValues()[0];' +
    ' var sourceRows=source.getRange(2,1,source.getLastRow()-1,sourceMaxCol).getValues();' +
    ' var sourceIx=vatHeaderIndexes_v648_(sourceHeaders);' +
    ' var outputHeaders=vatDetailHeaders_v648_();' +
    ' var outputIx=indexes(outputHeaders);' +
    ' requireIndexes(outputIx,"v6.66 생성행");' +
    ' var queues={}; var generated=0; var generatedNonblank=0;' +
    ' sourceRows.forEach(function(row,offset){ var result=vatDetailRow_v648_(row,sourceIx,offset+2); if(!result || !result.row) return; var k=keyFrom(result.row,outputIx); if(!queues[k]) queues[k]=[]; var p=normText(result.row[outputIx.payment]); queues[k].push(p); generated++; if(p) generatedNonblank++; });' +
    ' var detailValues=detail.getDataRange().getValues();' +
    ' var detailHeaders=detailValues[0]||[];' +
    ' var detailIx=indexes(detailHeaders);' +
    ' requireIndexes(detailIx,"부가세_신고자료");' +
    ' var detailRows=detailValues.slice(1);' +
    ' if(generated !== detailRows.length){ throw new Error("안전검증 실패: 원본 기준 생성행 "+generated+"건 / 기존 상세행 "+detailRows.length+"건. 아무 값도 쓰지 않았습니다."); }' +
    ' var paymentValues=[]; var unmatched=[]; var matchedNonblank=0;' +
    ' detailRows.forEach(function(row,offset){ var k=keyFrom(row,detailIx); var q=queues[k]; if(!q || !q.length){ unmatched.push(offset+2); paymentValues.push([""]); return; } var p=q.shift(); paymentValues.push([p]); if(p) matchedNonblank++; });' +
    ' var leftovers=0; Object.keys(queues).forEach(function(k){ leftovers += queues[k].length; });' +
    ' if(unmatched.length || leftovers){ throw new Error("안전검증 실패: 미매칭 상세행 "+unmatched.length+"건 / 남은 원본행 "+leftovers+"건. 아무 값도 쓰지 않았습니다."); }' +
    ' if(matchedNonblank < 1){ throw new Error("트래킹 번호/결제수단 값이 원본에서 0건입니다. 아무 값도 쓰지 않았습니다."); }' +
    ' detail.getRange(2,detailIx.payment+1,paymentValues.length,1).setNumberFormat("@").setValues(paymentValues);' +
    ' SpreadsheetApp.flush();' +
    ' var periodResult=buildVatPeriodSummary_v657_(ss);' +
    ' return {detailRows:detailRows.length,paymentRows:matchedNonblank,blankPaymentRows:detailRows.length-matchedNonblank,periodRows:periodResult && periodResult.rows != null ? periodResult.rows : ""};' +
    '})();'
  );

  SpreadsheetApp.getUi().alert(
    'PR #25 트래킹 결제수단 보정 및 카드 재매칭 완료\n\n' +
    '부가세 상세행: ' + Number(result.detailRows || 0).toLocaleString('ko-KR') + '건\n' +
    '결제수단 입력행: ' + Number(result.paymentRows || 0).toLocaleString('ko-KR') + '건\n' +
    '결제수단 공란행: ' + Number(result.blankPaymentRows || 0).toLocaleString('ko-KR') + '건\n' +
    '기간별 집계행: ' + String(result.periodRows) + '\n\n' +
    '사업자별 반기 신고요약과 부가세_카드매칭검증 시트를 확인하세요.'
  );
  return result;
}

function inspectPr25TrackingPaymentState() {
  const ss = SpreadsheetApp.getActive();
  const detail = ss.getSheetByName('부가세_신고자료');
  if (!detail || detail.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('부가세_신고자료가 없거나 데이터가 없습니다.');
    return;
  }
  const values = detail.getDataRange().getDisplayValues();
  const headers = values[0] || [];
  let paymentIx = -1;
  for (let i = 0; i < headers.length; i++) {
    if (String(headers[i] || '').replace(/\s/g, '') === '롯데결제수단') {
      paymentIx = i;
      break;
    }
  }
  let nonblank = 0;
  if (paymentIx >= 0) {
    for (let r = 1; r < values.length; r++) {
      if (String(values[r][paymentIx] || '').trim()) nonblank++;
    }
  }
  SpreadsheetApp.getUi().alert(
    'PR #25 현재 상태\n\n' +
    '부가세 상세행: ' + Math.max(values.length - 1, 0).toLocaleString('ko-KR') + '건\n' +
    '롯데결제수단 열: ' + (paymentIx >= 0 ? '있음' : '없음') + '\n' +
    '결제수단 입력행: ' + nonblank.toLocaleString('ko-KR') + '건'
  );
}

function loadPr25TrackingPaymentLightBundle_() {
  const chunks = [];
  for (let i = 0; i < LOTTEON_PR25_LIGHT_FILES.length; i++) {
    chunks.push(fetchPr25TrackingPaymentText_(LOTTEON_PR25_LIGHT_FILES[i]));
  }
  const bundle = chunks.join('\n\n;\n\n');
  if (bundle.indexOf('LOTTEON_PATCH_V666_VAT_TRACKING_PAYMENT_PRIMARY_LOADED') < 0) {
    throw new Error('v6.66 patch를 로드하지 못했습니다.');
  }
  return bundle;
}

function fetchPr25TrackingPaymentText_(path) {
  const url = LOTTEON_PR25_SMOKE_RAW_BASE + path;
  const response = UrlFetchApp.fetch(url + '?ts=' + new Date().getTime(), {
    method: 'get',
    muteHttpExceptions: true,
    followRedirects: true
  });
  const code = response.getResponseCode();
  const text = response.getContentText('UTF-8');
  if (code < 200 || code >= 300) {
    throw new Error(path + ' 로드 실패 HTTP ' + code + '\n' + url + '\n' + text.slice(0, 500));
  }
  return text;
}
