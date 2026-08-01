/** PR #29 operating preview smoke. Writes PR29_* sheets only; production VAT sheets are read-only. */
const PR29_PREVIEW_VERSION = 'v1.0-PR29-V669-PREVIEW';
const PR29_PATCH_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-28-vat-period-card-fallback/Patch_v6_69_vat_tracking_period_card_fallback.gs';

function runPr29PreviewSmoke() {
  if (typeof loadLotteonRemoteBundle_ !== 'function') throw new Error('Code.gs v1.14-MAIN loader를 찾지 못했습니다.');
  const ss = SpreadsheetApp.getActive();
  const mainBundle = loadLotteonRemoteBundle_();
  const patch = pr29FetchText_(PR29_PATCH_URL);
  const spreadsheetId = ss.getId();
  const invocation = [
    ';(function(){',
    'var ss=SpreadsheetApp.openById(' + JSON.stringify(spreadsheetId) + ');',
    "var detail=ss.getSheetByName('부가세_신고자료');",
    "if(!detail||detail.getLastRow()<2)throw new Error('부가세_신고자료가 없습니다.');",
    'var orders=groupVatDetailByOrder_v660_(detail.getDataRange().getValues());',
    'var history=loadVatCardHistory_v660_(ss);',
    'var master=loadVatCardMaster_v660_(ss);',
    'var canonical=canonicalizeVatHistory_v664_(history,master);',
    'allocateVatPurchaseCards_v664_(orders,canonical,master);',
    'var stats={orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,fallback:0,fallbackMatched:0,fallbackNonCard:0,fallbackAmbiguous:0,invalidIdentity:0,invalidFallbackEvidence:0};',
    'var totals={sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};',
    'orders.forEach(function(o){var m=o.cardMatch||noMatch_v660_("미실행");stats.orders++;if(m.status==="MATCHED"||m.status==="MASTER_MATCHED")stats.matched++;else if(m.status==="NON_CARD")stats.nonCard++;else if(m.status==="AMBIGUOUS")stats.ambiguous++;else stats.noMatch++;if(m.v669Fallback){stats.fallback++;if(m.status==="MATCHED")stats.fallbackMatched++;else if(m.status==="NON_CARD")stats.fallbackNonCard++;else if(m.status==="AMBIGUOUS")stats.fallbackAmbiguous++;if(m.approvalDate||m.approvalNo||Number(m.approvalAmount||0)!==0||String(m.reason||"").indexOf("금액비교없음")<0&&m.status!=="AMBIGUOUS")stats.invalidFallbackEvidence++;}var c=normalizeCardCompany_v660_(m.company);var e=normalizeVatCardEnd4_v667_(m.cardEnd4,m.cardNumber);if((c==="KB국민카드"&&e!=="4091")||(c==="우리카드"&&e!=="7680")||(m.cardName==="Trip to 로카"&&e!=="0126")||(m.cardName==="LOCA LIKIT 1.2"&&e!=="0036"))stats.invalidIdentity++;Object.keys(totals).forEach(function(k){totals[k]+=Number(o[k]||0);});});',
    'var summaryHeaders=vatBusinessCardHalfHeaders_v660_();',
    'var summaryRows=aggregateVatBusinessCardHalf_v660_(orders);',
    'var diagHeaders=vatCardDiagnosticHeaders_v660_().concat(["v6.69 2차귀속"]);',
    'var diagRows=orders.map(function(o){var m=o.cardMatch||noMatch_v660_("미실행");return [o.year,o.half,o.orderDate,o.business,o.account,o.orderNo,o.lottePayment,o.detailRows,o.purchase,m.company,m.alias,m.cardName,m.cardNumber,m.cardEnd4,m.approvalDate,m.approvalTime,m.approvalNo,m.approvalAmount,m.status,m.reason,m.candidateCount,m.merchant,m.merchantOrderNo,m.evidenceType,m.cancelMemo,m.sourceFile,m.candidateSummary,m.v669Fallback?"Y":""];});',
    'return {stats:stats,totals:totals,summaryHeaders:summaryHeaders,summaryRows:summaryRows,diagHeaders:diagHeaders,diagRows:diagRows,canonicalRows:canonical.length};',
    '})()'
  ].join('\n');
  const out = eval(mainBundle + '\n\n;\n\n' + patch + '\n\n;\n\n' + invocation);
  pr29Validate_(out);
  pr29WritePreview_(ss, out);
  SpreadsheetApp.flush();
  SpreadsheetApp.getUi().alert(
    'PR #29 미리보기 완료\n\n' +
    '주문: ' + out.stats.orders + '건\n' +
    'MATCHED: ' + out.stats.matched + '건\n' +
    'NON_CARD: ' + out.stats.nonCard + '건\n' +
    'AMBIGUOUS: ' + out.stats.ambiguous + '건\n' +
    'NO_MATCH: ' + out.stats.noMatch + '건\n' +
    'v6.69 2차귀속: ' + out.stats.fallback + '건\n\n' +
    '생성 시트: PR29_사업자별반기요약, PR29_카드매칭검증, PR29_실행상태\n' +
    '기존 부가세 시트는 변경하지 않았습니다.'
  );
  return out;
}

function pr29FetchText_(url) {
  const r = UrlFetchApp.fetch(url + '?ts=' + Date.now(), {muteHttpExceptions:true, followRedirects:true});
  const code = r.getResponseCode(), text = r.getContentText('UTF-8');
  if (code < 200 || code >= 300) throw new Error('PR29 patch 로드 실패 HTTP ' + code + '\n' + text.slice(0, 300));
  return text;
}

function pr29Validate_(out) {
  if (!out || !out.stats || !out.totals) throw new Error('PR29 결과가 비어 있습니다.');
  const expected = {orders:1355,sales:71838700,salesSupply:65307938,salesVat:6530762,settlement:64726771,fee:7111929,purchase:54807644,purchaseSupply:49825146,purchaseVat:4982498,payable:1548264,profit:9919127,vatProfit:8370863};
  Object.keys(expected).forEach(function(k){
    const actual = k === 'orders' ? Number(out.stats.orders||0) : Math.round(Number(out.totals[k]||0));
    if (actual !== expected[k]) throw new Error('PR29 불변합계 검증 실패: ' + k + ' 실제 ' + actual + ' / 기대 ' + expected[k]);
  });
  const classified = out.stats.matched + out.stats.nonCard + out.stats.ambiguous + out.stats.noMatch;
  if (classified !== out.stats.orders) throw new Error('PR29 상태 합계 불일치: ' + classified + ' / ' + out.stats.orders);
  if (out.stats.invalidIdentity) throw new Error('PR29 잘못된 카드 식별자 ' + out.stats.invalidIdentity + '건');
  if (out.stats.invalidFallbackEvidence) throw new Error('PR29 2차귀속 증빙필드 오류 ' + out.stats.invalidFallbackEvidence + '건');
  if (out.stats.fallback < 1) throw new Error('PR29 2차귀속 결과가 0건입니다.');
  if (out.stats.noMatch >= 664) throw new Error('PR29 NO_MATCH가 줄지 않았습니다: ' + out.stats.noMatch);
}

function pr29WritePreview_(ss, out) {
  pr29WriteTable_(ss, 'PR29_사업자별반기요약', out.summaryHeaders, out.summaryRows, 2);
  pr29WriteTable_(ss, 'PR29_카드매칭검증', out.diagHeaders, out.diagRows, 1);
  const s = ss.getSheetByName('PR29_실행상태') || ss.insertSheet('PR29_실행상태');
  const rows = [
    ['항목','값'],['버전',PR29_PREVIEW_VERSION],['상태','PASS'],['운영시트 변경','없음'],
    ['주문',out.stats.orders],['MATCHED',out.stats.matched],['NON_CARD',out.stats.nonCard],
    ['AMBIGUOUS',out.stats.ambiguous],['NO_MATCH',out.stats.noMatch],['v6.69 2차귀속',out.stats.fallback],
    ['2차귀속 MATCHED',out.stats.fallbackMatched],['2차귀속 NON_CARD',out.stats.fallbackNonCard],
    ['2차귀속 AMBIGUOUS',out.stats.fallbackAmbiguous],['canonical 증빙행',out.canonicalRows],
    ['순수매출액',out.totals.sales],['매입금액',out.totals.purchase],['납부예상부가세',out.totals.payable],
    ['잘못된 카드 식별자',out.stats.invalidIdentity],['2차귀속 증빙필드 오류',out.stats.invalidFallbackEvidence]
  ];
  s.clearContents(); s.getRange(1,1,rows.length,2).setValues(rows); s.setFrozenRows(1);
  s.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold'); s.setColumnWidth(1,220); s.setColumnWidth(2,500);
}

function pr29WriteTable_(ss, name, headers, rows, frozenRows) {
  const s = ss.getSheetByName(name) || ss.insertSheet(name);
  s.clearContents(); s.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) s.getRange(2,1,rows.length,headers.length).setValues(rows);
  s.setFrozenRows(frozenRows || 1); s.getRange(1,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold');
}
