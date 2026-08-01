/**
 * PR #29 operating preview smoke v1.2.
 *
 * - Reads production VAT sheets only.
 * - Processes 2026 H1 orders in resumable batches.
 * - Writes PR29_* sheets only.
 */
const PR29_PREVIEW_VERSION = 'v1.2-PR29-V669-H1-BATCHED';
const PR29_PATCH_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/codex/issue-28-vat-period-card-fallback/Patch_v6_69_vat_tracking_period_card_fallback.gs';
const PR29_BATCH_SIZE = 150;
const PR29_WORK_SHEET = 'PR29_작업데이터';
const PR29_STATUS_SHEET = 'PR29_실행상태';
const PR29_CONTINUE_HANDLER = 'runPr29PreviewSmokeContinue';
const PR29_SPREADSHEET_ID_KEY = 'PR29_PREVIEW_SPREADSHEET_ID';

function runPr29PreviewSmoke() {
  if (typeof loadLotteonRemoteBundle_ !== 'function') {
    throw new Error('Code.gs v1.14-MAIN loader를 찾지 못했습니다.');
  }

  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');

  pr29DeleteContinuationTriggers_();
  PropertiesService.getScriptProperties().setProperty(PR29_SPREADSHEET_ID_KEY, ss.getId());

  ['PR29_사업자별반기요약', 'PR29_카드매칭검증', PR29_WORK_SHEET, PR29_STATUS_SHEET].forEach(function(name) {
    const sheet = ss.getSheetByName(name);
    if (sheet) ss.deleteSheet(sheet);
  });

  const work = ss.insertSheet(PR29_WORK_SHEET);
  const headers = pr29WorkHeaders_();
  work.getRange(1, 1, 1, headers.length).setValues([headers]);
  work.setFrozenRows(1);
  work.getRange(1, 1, 1, headers.length).setBackground('#d9eaf7').setFontWeight('bold');

  pr29WriteStatus_(ss, {
    status: 'RUNNING',
    message: '초기화 완료; 1차 배치 실행',
    processed: 0,
    target: 1355,
    sourceOrders: '',
    excludedOrders: '',
    startedAt: new Date().toISOString()
  });

  return runPr29PreviewSmokeContinue();
}

function runPr29PreviewSmokeContinue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(5000)) return { ok: false, reason: 'LOCK_BUSY' };

  try {
    const spreadsheetId = PropertiesService.getScriptProperties().getProperty(PR29_SPREADSHEET_ID_KEY);
    const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActive();
    if (!ss) throw new Error('PR29 대상 스프레드시트를 찾지 못했습니다.');

    const work = ss.getSheetByName(PR29_WORK_SHEET);
    if (!work) throw new Error(PR29_WORK_SHEET + ' 시트가 없습니다. runPr29PreviewSmoke를 다시 시작하세요.');

    const processed = Math.max(0, work.getLastRow() - 1);
    const used = pr29ReadUsedEvidence_(work);
    const batch = pr29RunRemoteBatch_(ss, processed, PR29_BATCH_SIZE, used);

    if (batch.rows && batch.rows.length) {
      work.getRange(work.getLastRow() + 1, 1, batch.rows.length, pr29WorkHeaders_().length).setValues(batch.rows);
    }

    const next = processed + (batch.rows ? batch.rows.length : 0);
    pr29WriteStatus_(ss, {
      status: next >= batch.targetOrders ? 'FINALIZING' : 'RUNNING',
      message: next >= batch.targetOrders ? '전체 배치 완료; 최종 검증 중' : '다음 배치 자동 예약',
      processed: next,
      target: batch.targetOrders,
      sourceOrders: batch.sourceOrders,
      excludedOrders: batch.excludedOrders,
      startedAt: pr29ReadStatusValue_(ss, '시작시각') || new Date().toISOString()
    });

    SpreadsheetApp.flush();

    if (next >= batch.targetOrders) {
      const result = pr29Finalize_(ss, batch.canonicalRows);
      pr29DeleteContinuationTriggers_();
      PropertiesService.getScriptProperties().deleteProperty(PR29_SPREADSHEET_ID_KEY);
      return result;
    }

    pr29ScheduleContinuation_();
    return {
      ok: true,
      done: false,
      processed: next,
      target: batch.targetOrders,
      nextScheduled: true
    };
  } catch (error) {
    try {
      const spreadsheetId = PropertiesService.getScriptProperties().getProperty(PR29_SPREADSHEET_ID_KEY);
      const ss = spreadsheetId ? SpreadsheetApp.openById(spreadsheetId) : SpreadsheetApp.getActive();
      if (ss) {
        pr29WriteStatus_(ss, {
          status: 'ERROR',
          message: String(error && error.message ? error.message : error),
          processed: Math.max(0, (ss.getSheetByName(PR29_WORK_SHEET) || { getLastRow:function(){return 1;} }).getLastRow() - 1),
          target: 1355,
          sourceOrders: '',
          excludedOrders: '',
          startedAt: pr29ReadStatusValue_(ss, '시작시각') || ''
        });
      }
    } catch (ignore) {}
    pr29DeleteContinuationTriggers_();
    throw error;
  } finally {
    lock.releaseLock();
  }
}

function pr29RunRemoteBatch_(ss, startIndex, batchSize, used) {
  const mainBundle = loadLotteonRemoteBundle_();
  const patch = pr29FetchText_(PR29_PATCH_URL);
  const spreadsheetId = ss.getId();
  const invocation = [
    ';(function(){',
    'var ss=SpreadsheetApp.openById(' + JSON.stringify(spreadsheetId) + ');',
    "var detail=ss.getSheetByName('부가세_신고자료');",
    "if(!detail||detail.getLastRow()<2)throw new Error('부가세_신고자료가 없습니다.');",
    'var allOrders=groupVatDetailByOrder_v660_(detail.getDataRange().getValues());',
    'var orders=allOrders.filter(function(o){return String(o.year)==="2026"&&String(o.half)==="상반기";});',
    'orders.sort(function(a,b){return String(a.orderDate||"").localeCompare(String(b.orderDate||""))||String(a.orderNo||"").localeCompare(String(b.orderNo||""))||Number(a.purchase||0)-Number(b.purchase||0);});',
    'var history=loadVatCardHistory_v660_(ss);',
    'var master=loadVatCardMaster_v660_(ss);',
    'var canonical=canonicalizeVatHistory_v664_(history,master);',
    'var used=' + JSON.stringify(used || {}) + ';',
    'var start=' + Number(startIndex || 0) + ';',
    'var end=Math.min(orders.length,start+' + Number(batchSize || PR29_BATCH_SIZE) + ');',
    'var rows=[];',
    'for(var i=start;i<end;i++){',
      'var o=orders[i];',
      'var m=matchVatOrderCardCanonical_v664_(o,canonical,master,used)||noMatch_v660_("미실행");',
      'rows.push([',
        'o.year,o.half,o.orderDate,o.business,o.account,o.orderNo,o.lottePayment,o.detailRows,',
        'o.sales,o.salesSupply,o.salesVat,o.settlement,o.fee,o.purchase,o.purchaseSupply,o.purchaseVat,o.payable,o.profit,o.vatProfit,',
        'm.company||"",m.alias||"",m.cardName||"",m.cardNumber||"",m.cardEnd4||"",',
        'm.approvalDate||"",m.approvalTime||"",m.approvalNo||"",Number(m.approvalAmount||0),',
        'm.status||"NO_MATCH",m.reason||"",Number(m.candidateCount||0),m.merchant||"",m.merchantOrderNo||"",',
        'm.evidenceType||"",m.cancelMemo||"",m.sourceFile||"",m.candidateSummary||"",m.v669Fallback?"Y":"",m.canonicalEvidenceKey||""',
      ']);',
    '}',
    'return {rows:rows,targetOrders:orders.length,sourceOrders:allOrders.length,excludedOrders:allOrders.length-orders.length,canonicalRows:canonical.length};',
    '})()'
  ].join('\n');

  return eval(mainBundle + '\n\n;\n\n' + patch + '\n\n;\n\n' + invocation);
}

function pr29Finalize_(ss, canonicalRows) {
  const work = ss.getSheetByName(PR29_WORK_SHEET);
  if (!work || work.getLastRow() < 2) throw new Error('PR29 작업 결과가 비어 있습니다.');

  const values = work.getDataRange().getValues();
  const headers = values[0];
  const index = pr29HeaderIndex_(headers);
  const orders = values.slice(1).map(function(row) {
    return pr29OrderFromWorkRow_(row, index);
  });

  const sourceOrders = Number(pr29ReadStatusValue_(ss, '전체 기간 주문') || orders.length);
  const excludedOrders = Number(pr29ReadStatusValue_(ss, '대상 제외 주문') || 0);
  const out = pr29BuildFinalResult_(orders, canonicalRows, sourceOrders, excludedOrders);
  pr29Validate_(out);

  const summaryRows = pr29AggregateSummary_(orders);
  pr29WriteTable_(ss, 'PR29_사업자별반기요약', pr29SummaryHeaders_(), summaryRows, 2);
  pr29WriteTable_(ss, 'PR29_카드매칭검증', pr29DiagnosticHeaders_(), pr29DiagnosticRows_(orders), 1);
  pr29WriteFinalStatus_(ss, out);

  ss.deleteSheet(work);
  SpreadsheetApp.flush();

  return {
    ok: true,
    done: true,
    version: PR29_PREVIEW_VERSION,
    stats: out.stats,
    totals: out.totals
  };
}

function pr29BuildFinalResult_(orders, canonicalRows, sourceOrders, excludedOrders) {
  const stats = {
    sourceOrders: Number(sourceOrders || orders.length),
    excludedOrders: Number(excludedOrders || 0),
    orders: 0,
    matched: 0,
    nonCard: 0,
    ambiguous: 0,
    noMatch: 0,
    fallback: 0,
    fallbackMatched: 0,
    fallbackNonCard: 0,
    fallbackAmbiguous: 0,
    invalidIdentity: 0,
    invalidFallbackEvidence: 0
  };
  const totals = {
    sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,
    purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0
  };

  orders.forEach(function(o) {
    const m = o.cardMatch || {};
    stats.orders++;
    if (m.status === 'MATCHED' || m.status === 'MASTER_MATCHED') stats.matched++;
    else if (m.status === 'NON_CARD') stats.nonCard++;
    else if (m.status === 'AMBIGUOUS') stats.ambiguous++;
    else stats.noMatch++;

    if (m.v669Fallback) {
      stats.fallback++;
      if (m.status === 'MATCHED') stats.fallbackMatched++;
      else if (m.status === 'NON_CARD') stats.fallbackNonCard++;
      else if (m.status === 'AMBIGUOUS') stats.fallbackAmbiguous++;

      if (m.approvalDate || m.approvalNo || Number(m.approvalAmount || 0) !== 0 ||
          (String(m.reason || '').indexOf('금액비교없음') < 0 && m.status !== 'AMBIGUOUS')) {
        stats.invalidFallbackEvidence++;
      }
    }

    const company = pr29NormalizeCompany_(m.company);
    const end4 = pr29NormalizeEnd4_(m.cardEnd4, m.cardNumber);
    if ((company === 'KB국민카드' && end4 !== '4091') ||
        (company === '우리카드' && end4 !== '7680') ||
        (m.cardName === 'Trip to 로카' && end4 !== '0126') ||
        (m.cardName === 'LOCA LIKIT 1.2' && end4 !== '0036')) {
      stats.invalidIdentity++;
    }

    Object.keys(totals).forEach(function(key) {
      totals[key] += Number(o[key] || 0);
    });
  });

  return { stats:stats, totals:totals, canonicalRows:Number(canonicalRows || 0) };
}

function pr29Validate_(out) {
  if (!out || !out.stats || !out.totals) throw new Error('PR29 결과가 비어 있습니다.');
  const expected = {
    orders:1355,sales:71838700,salesSupply:65307938,salesVat:6530762,
    settlement:64726771,fee:7111929,purchase:54807644,purchaseSupply:49825146,
    purchaseVat:4982498,payable:1548264,profit:9919127,vatProfit:8370863
  };

  Object.keys(expected).forEach(function(key) {
    const actual = key === 'orders' ? Number(out.stats.orders || 0) : Math.round(Number(out.totals[key] || 0));
    if (actual !== expected[key]) {
      throw new Error('PR29 상반기 불변합계 검증 실패: ' + key + ' 실제 ' + actual + ' / 기대 ' + expected[key]);
    }
  });

  const classified = out.stats.matched + out.stats.nonCard + out.stats.ambiguous + out.stats.noMatch;
  if (classified !== out.stats.orders) throw new Error('PR29 상태 합계 불일치: ' + classified + ' / ' + out.stats.orders);
  if (out.stats.invalidIdentity) throw new Error('PR29 잘못된 카드 식별자 ' + out.stats.invalidIdentity + '건');
  if (out.stats.invalidFallbackEvidence) throw new Error('PR29 2차귀속 증빙필드 오류 ' + out.stats.invalidFallbackEvidence + '건');
  if (out.stats.fallback < 1) throw new Error('PR29 2차귀속 결과가 0건입니다.');
  if (out.stats.noMatch >= 664) throw new Error('PR29 NO_MATCH가 줄지 않았습니다: ' + out.stats.noMatch);
}

function pr29WriteFinalStatus_(ss, out) {
  const s = ss.getSheetByName(PR29_STATUS_SHEET) || ss.insertSheet(PR29_STATUS_SHEET);
  const rows = [
    ['항목','값'],
    ['버전',PR29_PREVIEW_VERSION],
    ['상태','PASS'],
    ['운영시트 변경','없음'],
    ['검증 대상','2026년 상반기'],
    ['전체 기간 주문',out.stats.sourceOrders],
    ['대상 제외 주문',out.stats.excludedOrders],
    ['상반기 주문',out.stats.orders],
    ['MATCHED',out.stats.matched],
    ['NON_CARD',out.stats.nonCard],
    ['AMBIGUOUS',out.stats.ambiguous],
    ['NO_MATCH',out.stats.noMatch],
    ['v6.69 2차귀속',out.stats.fallback],
    ['2차귀속 MATCHED',out.stats.fallbackMatched],
    ['2차귀속 NON_CARD',out.stats.fallbackNonCard],
    ['2차귀속 AMBIGUOUS',out.stats.fallbackAmbiguous],
    ['canonical 증빙행',out.canonicalRows],
    ['순수매출액',out.totals.sales],
    ['매입금액',out.totals.purchase],
    ['납부예상부가세',out.totals.payable],
    ['잘못된 카드 식별자',out.stats.invalidIdentity],
    ['2차귀속 증빙필드 오류',out.stats.invalidFallbackEvidence],
    ['완료시각',new Date().toISOString()]
  ];
  s.clearContents();
  s.getRange(1,1,rows.length,2).setValues(rows);
  s.setFrozenRows(1);
  s.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  s.setColumnWidth(1,220);
  s.setColumnWidth(2,500);
}

function pr29WriteStatus_(ss, state) {
  const s = ss.getSheetByName(PR29_STATUS_SHEET) || ss.insertSheet(PR29_STATUS_SHEET);
  const rows = [
    ['항목','값'],
    ['버전',PR29_PREVIEW_VERSION],
    ['상태',state.status || 'RUNNING'],
    ['메시지',state.message || ''],
    ['처리주문',Number(state.processed || 0)],
    ['대상주문',Number(state.target || 0)],
    ['전체 기간 주문',state.sourceOrders === '' ? '' : Number(state.sourceOrders || 0)],
    ['대상 제외 주문',state.excludedOrders === '' ? '' : Number(state.excludedOrders || 0)],
    ['시작시각',state.startedAt || ''],
    ['갱신시각',new Date().toISOString()]
  ];
  s.clearContents();
  s.getRange(1,1,rows.length,2).setValues(rows);
  s.setFrozenRows(1);
  s.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  s.setColumnWidth(1,220);
  s.setColumnWidth(2,500);
}

function pr29ReadStatusValue_(ss, wanted) {
  const s = ss.getSheetByName(PR29_STATUS_SHEET);
  if (!s || s.getLastRow() < 2) return '';
  const values = s.getRange(2,1,s.getLastRow()-1,2).getValues();
  for (let i=0;i<values.length;i++) if (String(values[i][0]) === wanted) return values[i][1];
  return '';
}

function pr29ReadUsedEvidence_(work) {
  const used = {};
  if (!work || work.getLastRow() < 2) return used;
  const headers = work.getRange(1,1,1,work.getLastColumn()).getValues()[0];
  const ix = headers.indexOf('canonicalEvidenceKey');
  if (ix < 0) return used;
  const values = work.getRange(2,ix+1,work.getLastRow()-1,1).getValues();
  values.forEach(function(row) {
    const key = String(row[0] || '').trim();
    if (key) used[key] = true;
  });
  return used;
}

function pr29ScheduleContinuation_() {
  pr29DeleteContinuationTriggers_();
  ScriptApp.newTrigger(PR29_CONTINUE_HANDLER).timeBased().after(60 * 1000).create();
}

function pr29DeleteContinuationTriggers_() {
  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === PR29_CONTINUE_HANDLER) {
      try { ScriptApp.deleteTrigger(trigger); } catch (ignore) {}
    }
  });
}

function pr29FetchText_(url) {
  const response = UrlFetchApp.fetch(url + '?ts=' + Date.now(), {
    muteHttpExceptions:true,
    followRedirects:true
  });
  const code = response.getResponseCode();
  const text = response.getContentText('UTF-8');
  if (code < 200 || code >= 300) {
    throw new Error('PR29 patch 로드 실패 HTTP ' + code + '\n' + text.slice(0,300));
  }
  return text;
}

function pr29WorkHeaders_() {
  return [
    '신고연도','반기','주문일','사업자등록번호','쿠팡계정ID','주문번호','롯데결제수단','상세행수',
    '순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익',
    '구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','승인일','승인시각','승인번호','승인금액','카드매칭상태','카드매칭근거','후보수','가맹점명','가맹점주문번호','증빙유형','취소/부분취소메모','원본파일','후보요약','v6.69 2차귀속','canonicalEvidenceKey'
  ];
}

function pr29HeaderIndex_(headers) {
  const out = {};
  (headers || []).forEach(function(header, index) { out[String(header)] = index; });
  return out;
}

function pr29OrderFromWorkRow_(row, ix) {
  function value(name) { return ix[name] == null ? '' : row[ix[name]]; }
  function number(name) { return Number(value(name) || 0); }

  return {
    year:String(value('신고연도') || ''),
    half:String(value('반기') || ''),
    orderDate:String(value('주문일') || ''),
    business:String(value('사업자등록번호') || ''),
    account:String(value('쿠팡계정ID') || ''),
    orderNo:String(value('주문번호') || ''),
    lottePayment:String(value('롯데결제수단') || ''),
    detailRows:number('상세행수'),
    sales:number('순수매출액'),
    salesSupply:number('매출공급가액'),
    salesVat:number('매출부가세'),
    settlement:number('정산기준금액'),
    fee:number('마켓수수료'),
    purchase:number('매입금액'),
    purchaseSupply:number('매입공급가액'),
    purchaseVat:number('매입부가세'),
    payable:number('납부예상부가세'),
    profit:number('예상이익'),
    vatProfit:number('부가세반영예상이익'),
    cardMatch:{
      company:String(value('구매카드사') || ''),
      alias:String(value('구매카드별칭') || ''),
      cardName:String(value('구매카드명') || ''),
      cardNumber:String(value('카드번호') || ''),
      cardEnd4:String(value('카드번호끝4') || ''),
      approvalDate:String(value('승인일') || ''),
      approvalTime:String(value('승인시각') || ''),
      approvalNo:String(value('승인번호') || ''),
      approvalAmount:number('승인금액'),
      status:String(value('카드매칭상태') || 'NO_MATCH'),
      reason:String(value('카드매칭근거') || ''),
      candidateCount:number('후보수'),
      merchant:String(value('가맹점명') || ''),
      merchantOrderNo:String(value('가맹점주문번호') || ''),
      evidenceType:String(value('증빙유형') || ''),
      cancelMemo:String(value('취소/부분취소메모') || ''),
      sourceFile:String(value('원본파일') || ''),
      candidateSummary:String(value('후보요약') || ''),
      v669Fallback:String(value('v6.69 2차귀속') || '') === 'Y',
      canonicalEvidenceKey:String(value('canonicalEvidenceKey') || '')
    }
  };
}

function pr29AggregateSummary_(orders) {
  const map = {};
  (orders || []).forEach(function(o) {
    const m = o.cardMatch || {};
    const business = o.business || '사업자번호 미매핑';
    const identity = m.status === 'AMBIGUOUS' || m.status === 'NO_MATCH'
      ? m.status
      : [m.company,m.alias,m.cardName,m.cardNumber,m.cardEnd4,m.status].join('|');
    const key = [o.year,o.half,business,identity].join('|');

    if (!map[key]) {
      map[key] = {
        year:o.year,half:o.half,business:business,accounts:{},company:m.company||'',alias:m.alias||'',
        cardName:m.cardName||'',cardNumber:m.cardNumber||'',cardEnd4:m.cardEnd4||'',status:m.status||'NO_MATCH',
        reasons:{},orders:{},blankOrders:0,sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,
        purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0
      };
    }

    const x = map[key];
    if (o.account) x.accounts[o.account] = true;
    if (o.orderNo) x.orders[o.orderNo] = true;
    else x.blankOrders++;
    if (m.reason) x.reasons[m.reason] = true;
    x.sales += o.sales;
    x.salesSupply += o.salesSupply;
    x.salesVat += o.salesVat;
    x.settlement += o.settlement;
    x.fee += o.fee;
    x.purchase += o.purchase;
    x.purchaseSupply += o.purchaseSupply;
    x.purchaseVat += o.purchaseVat;
    x.payable += o.payable;
    x.profit += o.profit;
    x.vatProfit += o.vatProfit;
  });

  const halfRank = {'상반기':0,'하반기':1};
  return Object.keys(map).map(function(key) {
    const x = map[key];
    const notes = [];
    if (x.business === '사업자번호 미매핑') notes.push('사업자번호 미매핑');
    if (x.blankOrders) notes.push('주문번호 공란 ' + x.blankOrders + '건');
    if (x.status === 'AMBIGUOUS' || x.status === 'NO_MATCH') notes.push('카드 미확정');
    return [
      x.year,x.half,x.business,Object.keys(x.accounts).sort().join(', '),x.company,x.alias,x.cardName,
      x.cardNumber,x.cardEnd4,x.status,Object.keys(x.reasons).sort().join(' / '),Object.keys(x.orders).length,
      x.sales,x.salesSupply,x.salesVat,x.settlement,x.fee,x.purchase,x.purchaseSupply,x.purchaseVat,
      x.payable,x.profit,x.vatProfit,notes.join(' / ')
    ];
  }).sort(function(a,b) {
    return String(a[0]).localeCompare(String(b[0])) ||
      (halfRank[a[1]] - halfRank[b[1]]) ||
      String(a[2]).localeCompare(String(b[2])) ||
      String(a[4]).localeCompare(String(b[4])) ||
      String(a[7]).localeCompare(String(b[7]));
  });
}

function pr29SummaryHeaders_() {
  return ['신고연도','반기','사업자등록번호','연결 쿠팡계정ID','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','카드매칭상태','카드매칭근거','주문건수','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];
}

function pr29DiagnosticHeaders_() {
  return ['신고연도','반기','주문일','사업자등록번호','쿠팡계정ID','주문번호','롯데결제수단','상세행수','주문매입금액','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','승인일','승인시각','승인번호','승인금액','카드매칭상태','카드매칭근거','후보수','가맹점명','가맹점주문번호','증빙유형','취소/부분취소메모','원본파일','후보요약','v6.69 2차귀속'];
}

function pr29DiagnosticRows_(orders) {
  return (orders || []).map(function(o) {
    const m = o.cardMatch || {};
    return [
      o.year,o.half,o.orderDate,o.business,o.account,o.orderNo,o.lottePayment,o.detailRows,o.purchase,
      m.company||'',m.alias||'',m.cardName||'',m.cardNumber||'',m.cardEnd4||'',m.approvalDate||'',m.approvalTime||'',
      m.approvalNo||'',Number(m.approvalAmount||0),m.status||'NO_MATCH',m.reason||'',Number(m.candidateCount||0),
      m.merchant||'',m.merchantOrderNo||'',m.evidenceType||'',m.cancelMemo||'',m.sourceFile||'',m.candidateSummary||'',
      m.v669Fallback?'Y':''
    ];
  });
}

function pr29WriteTable_(ss, name, headers, rows, frozenRows) {
  const s = ss.getSheetByName(name) || ss.insertSheet(name);
  s.clearContents();
  s.getRange(1,1,1,headers.length).setValues([headers]);
  if (rows.length) s.getRange(2,1,rows.length,headers.length).setValues(rows);
  s.setFrozenRows(frozenRows || 1);
  s.getRange(1,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold');
}

function pr29NormalizeCompany_(value) {
  const s = String(value == null ? '' : value).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');
  if (!s) return '';
  if (s.indexOf('kb') >= 0 || s.indexOf('국민') >= 0) return 'KB국민카드';
  if (s.indexOf('롯데') >= 0) return '롯데카드';
  if (s.indexOf('우리') >= 0) return '우리카드';
  if (s.indexOf('신한') >= 0) return '신한카드';
  if (s.indexOf('농협') >= 0 || s.indexOf('nh') >= 0) return 'NH농협카드';
  if (s.indexOf('삼성') >= 0) return '삼성카드';
  if (s.indexOf('하나') >= 0) return '하나카드';
  if (s.indexOf('현대') >= 0) return '현대카드';
  if (s.indexOf('비카드') >= 0 || s.indexOf('카카오') >= 0 || s.indexOf('머니') >= 0) return '비카드';
  return String(value == null ? '' : value).trim();
}

function pr29NormalizeEnd4_(end4, cardNumber) {
  const explicit = String(end4 == null ? '' : end4).replace(/\D/g,'');
  if (explicit) return ('0000' + explicit).slice(-4);
  const digits = String(cardNumber == null ? '' : cardNumber).replace(/\D/g,'');
  return digits.length >= 4 ? digits.slice(-4) : '';
}
