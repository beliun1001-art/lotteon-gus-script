/** PR #25 v6.68 production summary-only rerun after KakaoPay Money identity normalization. */
const PR25_R14_VERSION = 'v1.25-PR25-PRODUCTION-SUMMARY-R14';

function runPr25ProductionSummaryR14() {
  if (typeof pr25r13_getVatState_ !== 'function' ||
      typeof pr25r13_getOrPrimeBundle_ !== 'function' ||
      typeof pr25r13_clearBundleCache_ !== 'function' ||
      typeof pr25r13_resolveSpreadsheet_ !== 'function') {
    throw new Error('PR25_R13 helper를 찾지 못했습니다. PR25_R13 파일을 유지하세요.');
  }

  var state = pr25r13_getVatState_();
  var ss = pr25r13_resolveSpreadsheet_(state);
  if (!state || state.status !== 'done' || state.phase !== 'done') {
    throw new Error('R13 production VAT가 done/done 상태가 아닙니다.');
  }

  try {
    pr25r13_clearBundleCache_();
    var bundleInfo = pr25r13_getOrPrimeBundle_();
    var spreadsheetId = ss.getId();
    var invocation = [
      "LOTTEON_PATCH_BOOTSTRAP_VERSION = 'v6.68-PR25-R14';",
      "buildVatPurchaseCardReconciliation_v660_(SpreadsheetApp.openById(" + JSON.stringify(spreadsheetId) + "));"
    ].join('\n');
    var result = eval(bundleInfo.bundle + '\n\n;\n\n' + invocation);
    var validation = pr25r14_validateOutput_(ss);
    pr25r14_assertValidation_(validation);
    pr25r14_writeStatus_(ss, 'done', 'v6.68 production 카드 요약 재생성 완료', bundleInfo.source, result, validation, '');
    return { ok:true, result:result, validation:validation };
  } catch (e) {
    pr25r14_writeStatus_(ss, 'failed', 'v6.68 production 카드 요약 재생성 오류', 'error', null, null, String(e && e.message ? e.message : e));
    throw e;
  }
}

function pr25r14_validateOutput_(ss) {
  var sheet = ss.getSheetByName('부가세_기간별');
  if (!sheet) throw new Error('부가세_기간별 시트를 찾지 못했습니다.');
  var values = sheet.getDataRange().getValues();
  var headerRow = -1;
  for (var r = 0; r < values.length; r++) {
    if (values[r].indexOf('카드매칭상태') >= 0 && values[r].indexOf('구매카드명') >= 0) {
      headerRow = r;
      break;
    }
  }
  if (headerRow < 0) throw new Error('구매카드별 신고요약 헤더를 찾지 못했습니다.');
  var headers = values[headerRow];
  var ix = function(name) { return headers.indexOf(name); };
  var p = {
    half:ix('반기'), company:ix('구매카드사'), alias:ix('구매카드별칭'),
    name:ix('구매카드명'), end4:ix('카드번호끝4'), status:ix('카드매칭상태'),
    orders:ix('주문건수'), sales:ix('순수매출액'), purchase:ix('매입금액'),
    salesSupply:ix('매출공급가액'), salesVat:ix('매출부가세'),
    settlement:ix('정산기준금액'), fee:ix('마켓수수료'),
    purchaseSupply:ix('매입공급가액'), purchaseVat:ix('매입부가세'),
    payable:ix('납부예상부가세'), profit:ix('예상이익'), vatProfit:ix('부가세반영예상이익')
  };
  var out = {
    halfOrders:0, matched:0, nonCard:0, ambiguous:0, noMatch:0,
    sales:0, salesSupply:0, salesVat:0, settlement:0, fee:0,
    purchase:0, purchaseSupply:0, purchaseVat:0, payable:0, profit:0, vatProfit:0,
    kakaoMoneySummaryRows:0, kakaoMoneyOrders:0, blankKakaoMoneyAliasRows:0,
    invalidCardIdentityRows:0
  };
  function n(v) { return Math.round(Number(String(v == null ? '' : v).replace(/,/g,'')) || 0); }
  for (var i = headerRow + 1; i < values.length; i++) {
    var row = values[i];
    if (!String(row[0] || '').trim()) break;
    if (String(row[p.half] || '').trim() !== '상반기') continue;

    var orders = n(row[p.orders]);
    var status = String(row[p.status] || '').trim();
    out.halfOrders += orders;
    if (status === 'MATCHED') out.matched += orders;
    else if (status === 'NON_CARD') out.nonCard += orders;
    else if (status === 'AMBIGUOUS') out.ambiguous += orders;
    else if (status === 'NO_MATCH') out.noMatch += orders;

    out.sales += n(row[p.sales]);
    out.salesSupply += n(row[p.salesSupply]);
    out.salesVat += n(row[p.salesVat]);
    out.settlement += n(row[p.settlement]);
    out.fee += n(row[p.fee]);
    out.purchase += n(row[p.purchase]);
    out.purchaseSupply += n(row[p.purchaseSupply]);
    out.purchaseVat += n(row[p.purchaseVat]);
    out.payable += n(row[p.payable]);
    out.profit += n(row[p.profit]);
    out.vatProfit += n(row[p.vatProfit]);

    var company = String(row[p.company] || '').trim();
    var alias = String(row[p.alias] || '').trim();
    var name = String(row[p.name] || '').trim();
    var end4 = String(row[p.end4] || '').trim();

    if (status === 'NON_CARD' && name === '카카오페이 페이머니') {
      out.kakaoMoneySummaryRows++;
      out.kakaoMoneyOrders += orders;
      if (!alias) out.blankKakaoMoneyAliasRows++;
    }
    if (status === 'MATCHED') {
      if (company === 'KB국민카드' && end4 !== '4091') out.invalidCardIdentityRows++;
      if (company === '우리카드' && end4 !== '7680') out.invalidCardIdentityRows++;
      if (name === 'Trip to 로카' && end4 !== '0126') out.invalidCardIdentityRows++;
      if (name === 'LOCA LIKIT 1.2' && end4 !== '0036') out.invalidCardIdentityRows++;
    }
  }
  return out;
}

function pr25r14_assertValidation_(v) {
  var expected = {
    halfOrders:1355, sales:71838700, salesSupply:65307938, salesVat:6530762,
    settlement:64726771, fee:7111929, purchase:54807644,
    purchaseSupply:49825146, purchaseVat:4982498, payable:1548264,
    profit:9919127, vatProfit:8370863,
    matched:425, nonCard:265, ambiguous:1, noMatch:664,
    kakaoMoneySummaryRows:3, kakaoMoneyOrders:265,
    blankKakaoMoneyAliasRows:0, invalidCardIdentityRows:0
  };
  Object.keys(expected).forEach(function(key) {
    if (Number(v[key] || 0) !== expected[key]) {
      throw new Error('R14 출력검증 실패: ' + key + ' 실제 ' + Number(v[key] || 0) + ' / 기대 ' + expected[key]);
    }
  });
}

function pr25r14_writeStatus_(ss, status, message, bundleSource, result, validation, errorText) {
  var sheet = ss.getSheetByName('PR25_실행상태') || ss.insertSheet('PR25_실행상태');
  result = result || {};
  validation = validation || {};
  var rows = [
    ['항목','값'],
    ['버전',PR25_R14_VERSION],
    ['실행 경로','R13 branch ZIP production bundle + v6.68 summary-only rebuild'],
    ['bundle 출처',bundleSource || ''],
    ['상태',status || ''],
    ['단계','summary'],
    ['메시지',message || ''],
    ['상세 재생성','N'],
    ['상반기 주문건수',Number(validation.halfOrders || 0)],
    ['상반기 MATCHED',Number(validation.matched || 0)],
    ['상반기 NON_CARD',Number(validation.nonCard || 0)],
    ['상반기 AMBIGUOUS',Number(validation.ambiguous || 0)],
    ['상반기 NO_MATCH',Number(validation.noMatch || 0)],
    ['상반기 순수매출액',Number(validation.sales || 0)],
    ['상반기 매입금액',Number(validation.purchase || 0)],
    ['카카오머니 요약행',Number(validation.kakaoMoneySummaryRows || 0)],
    ['카카오머니 주문건수',Number(validation.kakaoMoneyOrders || 0)],
    ['카카오머니 별칭 공란행',Number(validation.blankKakaoMoneyAliasRows || 0)],
    ['잘못된 카드식별자 행',Number(validation.invalidCardIdentityRows || 0)],
    ['마지막 오류',errorText || '']
  ];
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,240);
  sheet.setColumnWidth(2,700);
  SpreadsheetApp.flush();
}
