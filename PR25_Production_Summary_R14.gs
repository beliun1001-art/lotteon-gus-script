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
    if (validation.blankKakaoMoneyAliasRows !== 0) {
      throw new Error('카카오페이 페이머니 별칭 공란 행이 남았습니다: ' + validation.blankKakaoMoneyAliasRows);
    }
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
  var iHalf=ix('반기'), iAlias=ix('구매카드별칭'), iName=ix('구매카드명'),
      iStatus=ix('카드매칭상태'), iOrders=ix('주문건수');
  var blankRows=0, kakaoRows=0, kakaoOrders=0;
  for (var i = headerRow + 1; i < values.length; i++) {
    var row = values[i];
    if (!String(row[0] || '').trim()) break;
    if (String(row[iHalf] || '').trim() !== '상반기') continue;
    if (String(row[iStatus] || '').trim() !== 'NON_CARD') continue;
    if (String(row[iName] || '').trim() !== '카카오페이 페이머니') continue;
    kakaoRows++;
    kakaoOrders += Number(row[iOrders] || 0);
    if (!String(row[iAlias] || '').trim()) blankRows++;
  }
  return {
    kakaoMoneySummaryRows:kakaoRows,
    kakaoMoneyOrders:kakaoOrders,
    blankKakaoMoneyAliasRows:blankRows
  };
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
    ['카드요약 행수',Number(result.summaryRows || 0)],
    ['주문건수',Number(result.orderRows || 0)],
    ['MATCHED',Number(result.matchedOrders || 0)],
    ['NON_CARD',Number(result.nonCardOrders || 0)],
    ['AMBIGUOUS',Number(result.ambiguousOrders || 0)],
    ['NO_MATCH',Number(result.noMatchOrders || 0)],
    ['카카오머니 요약행',Number(validation.kakaoMoneySummaryRows || 0)],
    ['카카오머니 주문건수',Number(validation.kakaoMoneyOrders || 0)],
    ['카카오머니 별칭 공란행',Number(validation.blankKakaoMoneyAliasRows || 0)],
    ['마지막 오류',errorText || '']
  ];
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1,220);
  sheet.setColumnWidth(2,700);
  SpreadsheetApp.flush();
}
