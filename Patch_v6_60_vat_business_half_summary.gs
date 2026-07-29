/**
 * v6.60 Issue #20 - purchase-card VAT reconciliation and business/half-year filing summary.
 *
 * Core rules:
 * - The card shown in the filing summary is the actual LOTTEON/Lotte purchase instrument.
 * - 2026-06-30 and earlier: NEVER infer a card from business/account alone. Card-statement evidence is required.
 * - 2026-07-01 and later: statement evidence still wins; card master may be used only when
 *   business/account + issuer + applicable period resolves to exactly one active card.
 * - Multi-item VAT detail rows are grouped to order level before statement amount matching.
 * - Ambiguous/no-match orders stay visible and are never force-assigned.
 */
var LOTTEON_PATCH_V660_VAT_BUSINESS_HALF_SUMMARY_LOADED = true;
var LOTTEON_V660_CARD_MASTER_SHEET = '카드_마스터';
var LOTTEON_V660_CARD_HISTORY_SHEET = '카드사용내역_붙여넣기';
var LOTTEON_V660_CARD_DIAG_SHEET = '부가세_카드매칭검증';
var LOTTEON_V660_MASTER_CUTOFF = '2026-07-01';

/* Keep new card input/diagnostic sheets visible and eligible for safe-width handling. */
if (typeof LOTTEON_V653_CORE_VISIBLE_SHEETS !== 'undefined') {
  [LOTTEON_V660_CARD_MASTER_SHEET, LOTTEON_V660_CARD_HISTORY_SHEET, LOTTEON_V660_CARD_DIAG_SHEET].forEach(function(name) {
    if (LOTTEON_V653_CORE_VISIBLE_SHEETS.indexOf(name) < 0) LOTTEON_V653_CORE_VISIBLE_SHEETS.push(name);
  });
}
var __baseGetSafeWidthTargetSheets_v660_ = typeof getSafeWidthTargetSheets_v652_ === 'function' ? getSafeWidthTargetSheets_v652_ : null;
if (__baseGetSafeWidthTargetSheets_v660_) {
  getSafeWidthTargetSheets_v652_ = function(ss) {
    var result = __baseGetSafeWidthTargetSheets_v660_.apply(this, arguments) || [];
    var seen = {};
    result.forEach(function(sheet) { try { seen[sheet.getName()] = true; } catch (e) {} });
    [LOTTEON_V660_CARD_MASTER_SHEET, LOTTEON_V660_CARD_HISTORY_SHEET, LOTTEON_V660_CARD_DIAG_SHEET].forEach(function(name) {
      var sheet = ss && ss.getSheetByName && ss.getSheetByName(name);
      if (!sheet || seen[name]) return;
      try { if (sheet.isSheetHidden()) return; } catch (e) {}
      seen[name] = true; result.push(sheet);
    });
    return result;
  };
}

/* ---------- Optional LOTTEON purchase-payment evidence in VAT detail ---------- */
var __baseVatHeaderIndexes_v660_ = typeof vatHeaderIndexes_v648_ === 'function' ? vatHeaderIndexes_v648_ : null;
var __baseVatDetailHeaders_v660_ = typeof vatDetailHeaders_v648_ === 'function' ? vatDetailHeaders_v648_ : null;
var __baseVatDetailRow_v660_ = typeof vatDetailRow_v648_ === 'function' ? vatDetailRow_v648_ : null;

if (__baseVatHeaderIndexes_v660_) {
  vatHeaderIndexes_v648_ = function(headers) {
    var indexes = __baseVatHeaderIndexes_v660_.apply(this, arguments) || {};
    indexes.lottePayment = findVatPaymentHeader_v660_(headers || []);
    return indexes;
  };
}

if (__baseVatDetailHeaders_v660_) {
  vatDetailHeaders_v648_ = function() {
    var headers = __baseVatDetailHeaders_v660_.apply(this, arguments).slice();
    if (headers.indexOf('롯데결제수단') < 0) headers.push('롯데결제수단');
    return headers;
  };
}

if (__baseVatDetailRow_v660_) {
  vatDetailRow_v648_ = function(row, ix, sourceRow) {
    var result = __baseVatDetailRow_v660_.apply(this, arguments);
    if (!result || !result.row) return result;
    var payment = '';
    if (ix && typeof ix.lottePayment === 'number' && ix.lottePayment >= 0) {
      payment = cleanVatText_v648_(valueAt_v648_(row, ix.lottePayment));
    }
    result.row.push(payment);
    return result;
  };
}

function findVatPaymentHeader_v660_(headers) {
  return findHeaderAlias_v660_(headers, ['결제수단','결제정보','결제방법','카드사','결제수단/카드사','결제수단(카드사)','구매결제수단'], -1);
}

/* ---------- Wrap v6.57 period build ---------- */
var __baseBuildVatPeriodSummary_v660_ = typeof buildVatPeriodSummary_v657_ === 'function' ? buildVatPeriodSummary_v657_ : null;
if (__baseBuildVatPeriodSummary_v660_) {
  buildVatPeriodSummary_v657_ = function(ss) {
    ss = ss || SpreadsheetApp.getActive();
    ensureVatCardInputSheets_v660_(ss);
    var result = __baseBuildVatPeriodSummary_v660_.apply(this, arguments);
    var rendered = buildVatPurchaseCardReconciliation_v660_(ss);
    if (result && typeof result === 'object') {
      result.cardSummaryRows = rendered.summaryRows;
      result.cardMatchedOrders = rendered.matchedOrders;
      result.cardAmbiguousOrders = rendered.ambiguousOrders;
      result.cardNoMatchOrders = rendered.noMatchOrders;
    }
    return result;
  };
}

function ensureVatCardInputSheets_v660_(ss) {
  if (!ss || !ss.getSheetByName) return;
  ensureSheetHeaders_v660_(ss, LOTTEON_V660_CARD_MASTER_SHEET, vatCardMasterHeaders_v660_());
  ensureSheetHeaders_v660_(ss, LOTTEON_V660_CARD_HISTORY_SHEET, vatCardHistoryHeaders_v660_());
}

function ensureSheetHeaders_v660_(ss, name, headers) {
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sheet.getLastRow() < 1) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  }
  return sheet;
}

function vatCardMasterHeaders_v660_() {
  return ['카드사','카드별칭','카드명','카드구분','상태','카드번호','카드번호끝4','사업자코드','사업자등록번호','적용시작일','적용종료일','한도','메모'];
}
function vatCardHistoryHeaders_v660_() {
  return ['카드사','카드명','카드번호','카드번호끝4','승인일','승인시각','가맹점명','가맹점사업자번호','승인금액','승인번호','승인상태','취소일','취소금액','가맹점주문번호','증빙유형','롯데계열여부','원본파일','메모'];
}

/* ---------- Reconciliation orchestration ---------- */
function buildVatPurchaseCardReconciliation_v660_(ss) {
  var detail = ss && ss.getSheetByName && ss.getSheetByName('부가세_신고자료');
  var periodSheet = ss && ss.getSheetByName && ss.getSheetByName('부가세_기간별');
  if (!detail || !periodSheet || detail.getLastRow() < 1 || periodSheet.getLastRow() < 1) {
    return { summaryRows: 0, matchedOrders: 0, ambiguousOrders: 0, noMatchOrders: 0, reason: 'MISSING_VAT_SHEET' };
  }

  var detailValues = detail.getDataRange().getValues();
  var periodValues = periodSheet.getDataRange().getValues();
  var orders = groupVatDetailByOrder_v660_(detailValues);
  var history = loadVatCardHistory_v660_(ss);
  var master = loadVatCardMaster_v660_(ss);

  var stats = { matchedOrders: 0, ambiguousOrders: 0, noMatchOrders: 0, nonCardOrders: 0 };
  orders.forEach(function(order) {
    order.cardMatch = matchVatOrderCard_v660_(order, history, master);
    if (order.cardMatch.status === 'MATCHED' || order.cardMatch.status === 'MASTER_MATCHED') stats.matchedOrders++;
    else if (order.cardMatch.status === 'NON_CARD') stats.nonCardOrders++;
    else if (order.cardMatch.status === 'AMBIGUOUS') stats.ambiguousOrders++;
    else stats.noMatchOrders++;
  });

  writeVatCardMatchDiagnostic_v660_(ss, orders);
  var summary = aggregateVatBusinessCardHalf_v660_(orders);
  prependVatBusinessCardHalfSummary_v660_(periodSheet, periodValues, summary);
  return {
    summaryRows: summary.length,
    orderRows: orders.length,
    matchedOrders: stats.matchedOrders,
    nonCardOrders: stats.nonCardOrders,
    ambiguousOrders: stats.ambiguousOrders,
    noMatchOrders: stats.noMatchOrders
  };
}

/* ---------- VAT detail -> order level ---------- */
function groupVatDetailByOrder_v660_(values) {
  if (!values || values.length < 2) return [];
  var headers = values[0] || [];
  var ix = function(names) { return findHeaderAlias_v660_(headers, names, -1); };
  var indexes = {
    date: ix(['날짜','주문일','주문일자','마켓주문일자']),
    year: ix(['신고연도']), half: ix(['반기']), month: ix(['신고월']),
    account: ix(['쿠팡계정ID']), business: ix(['사업자등록번호']),
    orderNo: ix(['주문번호','마켓주문번호','주문ID','주문ID(마켓)']),
    payment: ix(['롯데결제수단','구매결제수단','결제수단']),
    sales: ix(['순수매출액']), salesSupply: ix(['매출공급가액']), salesVat: ix(['매출부가세']),
    settlement: ix(['정산기준금액']), fee: ix(['마켓수수료/비용','마켓수수료']),
    purchase: ix(['매입금액']), purchaseSupply: ix(['매입공급가액']), purchaseVat: ix(['매입부가세']),
    payable: ix(['납부예상부가세']), profit: ix(['예상이익']), vatProfit: ix(['부가세반영예상이익'])
  };
  var required = [indexes.year,indexes.half,indexes.account,indexes.business,indexes.sales,indexes.salesSupply,indexes.salesVat,indexes.settlement,indexes.purchase,indexes.purchaseSupply,indexes.purchaseVat,indexes.payable,indexes.profit,indexes.vatProfit];
  if (required.some(function(i) { return i < 0; })) throw new Error('v6.60 카드매칭용 부가세 상세 필수 헤더를 찾을 수 없습니다.');

  var map = {};
  function num(row, index) { return index < 0 ? 0 : number_v660_(row[index]); }
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var year = text_v660_(row[indexes.year]);
    var half = text_v660_(row[indexes.half]);
    if (!year || year === '기간미확인' || (half !== '상반기' && half !== '하반기')) continue;
    var account = text_v660_(row[indexes.account]);
    var business = text_v660_(row[indexes.business]);
    var orderNo = indexes.orderNo >= 0 ? text_v660_(row[indexes.orderNo]) : '';
    var orderDate = normalizeOrderDate_v660_(row[indexes.date], year);
    var payment = indexes.payment >= 0 ? text_v660_(row[indexes.payment]) : '';
    var key = orderNo ? [year,half,business,account,orderNo].join('|') : [year,half,business,account,'BLANK',r].join('|');
    if (!map[key]) {
      map[key] = {
        key:key, year:year, half:half, month:indexes.month >= 0 ? text_v660_(row[indexes.month]) : '',
        orderDate:orderDate, business:business, account:account, orderNo:orderNo,
        lottePayments:{}, detailRows:0,
        sales:0, salesSupply:0, salesVat:0, settlement:0, fee:0,
        purchase:0, purchaseSupply:0, purchaseVat:0, payable:0, profit:0, vatProfit:0
      };
    }
    var item = map[key];
    item.detailRows++;
    if (!item.orderDate && orderDate) item.orderDate = orderDate;
    if (payment) item.lottePayments[payment] = true;
    item.sales += num(row,indexes.sales);
    item.salesSupply += num(row,indexes.salesSupply);
    item.salesVat += num(row,indexes.salesVat);
    item.settlement += num(row,indexes.settlement);
    item.fee += indexes.fee >= 0 ? num(row,indexes.fee) : num(row,indexes.sales) - num(row,indexes.settlement);
    item.purchase += num(row,indexes.purchase);
    item.purchaseSupply += num(row,indexes.purchaseSupply);
    item.purchaseVat += num(row,indexes.purchaseVat);
    item.payable += num(row,indexes.payable);
    item.profit += num(row,indexes.profit);
    item.vatProfit += num(row,indexes.vatProfit);
  }
  return Object.keys(map).map(function(key) {
    var item = map[key];
    item.lottePayment = Object.keys(item.lottePayments).sort().join(', ');
    return item;
  }).sort(function(a,b) {
    return String(a.year).localeCompare(String(b.year)) || String(a.orderDate).localeCompare(String(b.orderDate)) || String(a.orderNo).localeCompare(String(b.orderNo));
  });
}

/* ---------- Input loaders ---------- */
function loadVatCardHistory_v660_(ss) {
  var sheet = ss && ss.getSheetByName && ss.getSheetByName(LOTTEON_V660_CARD_HISTORY_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues();
  var h = values[0] || [];
  var ix = function(names) { return findHeaderAlias_v660_(h, names, -1); };
  var p = {
    company:ix(['카드사']), name:ix(['카드명']), number:ix(['카드번호']), end4:ix(['카드번호끝4']),
    date:ix(['승인일','이용일','거래일']), time:ix(['승인시각','이용시각','거래시각']),
    merchant:ix(['가맹점명','이용가맹점']), merchantBusiness:ix(['가맹점사업자번호']), amount:ix(['승인금액','이용금액','거래금액']),
    approval:ix(['승인번호']), status:ix(['승인상태','승인/취소구분','상태']), cancelDate:ix(['취소일']), cancelAmount:ix(['취소금액']),
    orderNo:ix(['가맹점주문번호','주문번호']), evidence:ix(['증빙유형']), lotte:ix(['롯데계열여부']), source:ix(['원본파일']), memo:ix(['메모'])
  };
  var out = [];
  for (var r=1; r<values.length; r++) {
    var row=values[r], amount=number_v660_(p.amount >= 0 ? row[p.amount] : 0);
    var obj = {
      rowNo:r+1,
      company:textAt_v660_(row,p.company), cardName:textAt_v660_(row,p.name), cardNumber:textAt_v660_(row,p.number), cardEnd4:textAt_v660_(row,p.end4),
      date:normalizeDateText_v660_(textAt_v660_(row,p.date)), time:textAt_v660_(row,p.time), merchant:textAt_v660_(row,p.merchant), merchantBusiness:textAt_v660_(row,p.merchantBusiness),
      amount:amount, approvalNo:textAt_v660_(row,p.approval), status:textAt_v660_(row,p.status), cancelDate:normalizeDateText_v660_(textAt_v660_(row,p.cancelDate)),
      cancelAmount:number_v660_(p.cancelAmount >= 0 ? row[p.cancelAmount] : 0), merchantOrderNo:textAt_v660_(row,p.orderNo), evidenceType:textAt_v660_(row,p.evidence),
      lotteFlag:textAt_v660_(row,p.lotte), sourceFile:textAt_v660_(row,p.source), memo:textAt_v660_(row,p.memo)
    };
    obj.nonCard = isNonCardEvidence_v660_(obj);
    obj.cancelRow = isCancellationHistoryRow_v660_(obj);
    obj.lotteEvidence = isLotteEvidence_v660_(obj);
    obj.amountVariants = historyAmountVariants_v660_(obj);
    if (obj.date || obj.merchantOrderNo || obj.amount) out.push(obj);
  }
  return out;
}

function loadVatCardMaster_v660_(ss) {
  var sheet = ss && ss.getSheetByName && ss.getSheetByName(LOTTEON_V660_CARD_MASTER_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values=sheet.getDataRange().getValues(), h=values[0] || [];
  var ix=function(names){return findHeaderAlias_v660_(h,names,-1);};
  var p={company:ix(['카드사']),alias:ix(['카드별칭']),name:ix(['카드명']),type:ix(['카드구분']),status:ix(['상태']),number:ix(['카드번호']),end4:ix(['카드번호끝4']),account:ix(['사업자코드','쿠팡계정ID']),business:ix(['사업자등록번호']),start:ix(['적용시작일']),end:ix(['적용종료일']),limit:ix(['한도']),memo:ix(['메모'])};
  var out=[];
  for (var r=1;r<values.length;r++) {
    var row=values[r];
    var obj={rowNo:r+1,company:textAt_v660_(row,p.company),alias:textAt_v660_(row,p.alias),cardName:textAt_v660_(row,p.name),cardType:textAt_v660_(row,p.type),status:textAt_v660_(row,p.status),cardNumber:textAt_v660_(row,p.number),cardEnd4:textAt_v660_(row,p.end4),account:textAt_v660_(row,p.account),business:textAt_v660_(row,p.business),startDate:normalizeDateText_v660_(textAt_v660_(row,p.start)),endDate:normalizeDateText_v660_(textAt_v660_(row,p.end)),limit:number_v660_(p.limit>=0?row[p.limit]:0),memo:textAt_v660_(row,p.memo)};
    if (obj.company || obj.cardName || obj.cardNumber || obj.business || obj.account) out.push(obj);
  }
  return out;
}

/* ---------- Matching ---------- */
function matchVatOrderCard_v660_(order, history, master) {
  var direct = [];
  if (order.orderNo) {
    direct = (history || []).filter(function(h) {
      return !h.cancelRow && h.merchantOrderNo && text_v660_(h.merchantOrderNo) === text_v660_(order.orderNo);
    });
    direct = dedupeHistoryCandidates_v660_(direct);
    if (direct.length === 1) return matchFromHistory_v660_(order, direct[0], master, direct[0].nonCard ? 'NON_CARD' : 'MATCHED', '거래내역_주문번호');
    if (direct.length > 1) return ambiguousMatch_v660_(direct, '가맹점주문번호 동일 후보 ' + direct.length + '건');
  }

  var base = (history || []).filter(function(h) {
    if (h.cancelRow || !h.lotteEvidence) return false;
    if (!order.orderDate || h.date !== order.orderDate) return false;
    return historyMatchesAmount_v660_(h, order.purchase);
  });
  var paymentFiltered = filterEvidenceByLottePayment_v660_(base, order.lottePayment);
  var candidates = paymentFiltered.length ? paymentFiltered : base;
  candidates = dedupeHistoryCandidates_v660_(candidates);
  if (candidates.length === 1) {
    var reason = paymentFiltered.length ? '거래내역_일자+금액+롯데결제수단' : '거래내역_일자+금액';
    return matchFromHistory_v660_(order, candidates[0], master, candidates[0].nonCard ? 'NON_CARD' : 'MATCHED', reason);
  }
  if (candidates.length > 1) return ambiguousMatch_v660_(candidates, (paymentFiltered.length ? '일자+금액+결제수단' : '일자+금액') + ' 후보 ' + candidates.length + '건');

  if (order.orderDate && order.orderDate >= LOTTEON_V660_MASTER_CUTOFF) {
    var masterCandidates = findPostJulyMasterCandidates_v660_(order, master);
    if (masterCandidates.length === 1) return matchFromMaster_v660_(masterCandidates[0], '카드마스터_7월이후_사업자+카드사+적용기간');
    if (masterCandidates.length > 1) return ambiguousMasterMatch_v660_(masterCandidates);
  }
  return noMatch_v660_(order.orderDate && order.orderDate < LOTTEON_V660_MASTER_CUTOFF ? '상반기 거래내역 증빙 매칭 없음' : '거래내역/카드마스터 매칭 없음');
}

function filterEvidenceByLottePayment_v660_(rows, payment) {
  var p=text_v660_(payment);
  if (!p) return [];
  return (rows || []).filter(function(h) { return evidenceMatchesPayment_v660_(h,p); });
}
function evidenceMatchesPayment_v660_(h,payment) {
  var p=compact_v660_(payment);
  if (!p) return false;
  if (p.indexOf('카카오')>=0) return compact_v660_(h.cardName+h.evidenceType+h.company).indexOf('카카오')>=0;
  if (p.indexOf('네이버')>=0 || p.indexOf('npay')>=0) return compact_v660_(h.cardName+h.evidenceType+h.company).indexOf('npay')>=0 || compact_v660_(h.cardName+h.evidenceType+h.company).indexOf('네이버')>=0;
  var issuer=normalizeCardCompany_v660_(payment);
  return issuer && normalizeCardCompany_v660_(h.company) === issuer;
}

function findPostJulyMasterCandidates_v660_(order, master) {
  var issuer=normalizeCardCompany_v660_(order.lottePayment);
  if (!issuer) return [];
  return (master || []).filter(function(m) {
    if (!isActiveMaster_v660_(m)) return false;
    if (!m.startDate || m.startDate > order.orderDate) return false;
    if (m.endDate && m.endDate < order.orderDate) return false;
    if (normalizeCardCompany_v660_(m.company) !== issuer) return false;
    var businessMatch = order.business && m.business && normalizeBusinessNo_v660_(order.business) === normalizeBusinessNo_v660_(m.business);
    var accountMatch = order.account && m.account && normalizeAccountCode_v660_(order.account) === normalizeAccountCode_v660_(m.account);
    return businessMatch || accountMatch;
  });
}
function isActiveMaster_v660_(m) { return !/휴면|해지|사용불가|정지|폐기/i.test(text_v660_(m.status)); }

function matchFromHistory_v660_(order,h,master,status,reason) {
  var enriched = enrichHistoryFromMaster_v660_(h, master);
  var amountKind = historyAmountKind_v660_(h, order.purchase);
  var cancelMemo = '';
  if (h.cancelAmount) cancelMemo = '취소일 ' + (h.cancelDate || '-') + ' / 취소금액 ' + h.cancelAmount + (amountKind === 'NET_AFTER_CANCEL' ? ' / 취소반영금액 일치' : ' / 원승인금액 일치');
  return {
    status:status, reason:reason + (amountKind ? ' / ' + amountKind : ''), candidateCount:1,
    company:enriched.company || (h.nonCard ? '비카드' : ''), alias:enriched.alias || '', cardName:enriched.cardName || '', cardNumber:enriched.cardNumber || '', cardEnd4:enriched.cardEnd4 || '',
    approvalDate:h.date || '', approvalTime:h.time || '', approvalNo:h.approvalNo || '', approvalAmount:h.amount || 0,
    merchant:h.merchant || '', merchantOrderNo:h.merchantOrderNo || '', evidenceType:h.evidenceType || '', sourceFile:h.sourceFile || '', cancelMemo:cancelMemo,
    candidateSummary:historyCandidateLabel_v660_(h)
  };
}
function enrichHistoryFromMaster_v660_(h, master) {
  var matches=(master || []).filter(function(m){return samePhysicalCard_v660_(h,m);});
  var m=matches.length===1?matches[0]:null;
  return {
    company:h.company || (m && m.company) || '', alias:(m && m.alias)||'', cardName:(m && m.cardName)||h.cardName||'',
    cardNumber:h.cardNumber || (m && m.cardNumber) || '', cardEnd4:h.cardEnd4 || (m && m.cardEnd4) || ''
  };
}
function samePhysicalCard_v660_(h,m) {
  if (normalizeCardCompany_v660_(h.company) !== normalizeCardCompany_v660_(m.company)) return false;
  if (h.cardEnd4 && m.cardEnd4 && digits_v660_(h.cardEnd4) === digits_v660_(m.cardEnd4)) return true;
  var hNum=compact_v660_(h.cardNumber), mNum=compact_v660_(m.cardNumber);
  if (hNum && mNum && (hNum.indexOf(mNum)>=0 || mNum.indexOf(hNum)>=0)) return true;
  var hName=normalizeCardName_v660_(h.cardName), mName=normalizeCardName_v660_(m.cardName);
  return hName.length>=4 && mName.length>=4 && (hName.indexOf(mName)>=0 || mName.indexOf(hName)>=0);
}
function matchFromMaster_v660_(m,reason) {
  return {status:'MASTER_MATCHED',reason:reason,candidateCount:1,company:m.company||'',alias:m.alias||'',cardName:m.cardName||'',cardNumber:m.cardNumber||'',cardEnd4:m.cardEnd4||'',approvalDate:'',approvalTime:'',approvalNo:'',approvalAmount:0,merchant:'',merchantOrderNo:'',evidenceType:'카드_마스터',sourceFile:'',cancelMemo:'',candidateSummary:masterCandidateLabel_v660_(m)};
}
function ambiguousMatch_v660_(candidates,reason) { return {status:'AMBIGUOUS',reason:reason,candidateCount:candidates.length,company:'',alias:'',cardName:'',cardNumber:'',cardEnd4:'',approvalDate:'',approvalTime:'',approvalNo:'',approvalAmount:0,merchant:'',merchantOrderNo:'',evidenceType:'',sourceFile:'',cancelMemo:'',candidateSummary:candidates.map(historyCandidateLabel_v660_).join(' || ')}; }
function ambiguousMasterMatch_v660_(candidates) { return {status:'AMBIGUOUS',reason:'카드마스터_7월이후 후보 '+candidates.length+'건',candidateCount:candidates.length,company:'',alias:'',cardName:'',cardNumber:'',cardEnd4:'',approvalDate:'',approvalTime:'',approvalNo:'',approvalAmount:0,merchant:'',merchantOrderNo:'',evidenceType:'카드_마스터',sourceFile:'',cancelMemo:'',candidateSummary:candidates.map(masterCandidateLabel_v660_).join(' || ')}; }
function noMatch_v660_(reason) { return {status:'NO_MATCH',reason:reason,candidateCount:0,company:'',alias:'',cardName:'',cardNumber:'',cardEnd4:'',approvalDate:'',approvalTime:'',approvalNo:'',approvalAmount:0,merchant:'',merchantOrderNo:'',evidenceType:'',sourceFile:'',cancelMemo:'',candidateSummary:''}; }

/* ---------- Evidence rules ---------- */
function isCancellationHistoryRow_v660_(h) {
  var s=compact_v660_(h.status);
  if (!s) return false;
  if (s.indexOf('취소있음')>=0) return false; // original approval carrying later cancellation metadata
  return s.indexOf('취소')>=0 || s.indexOf('환불')>=0;
}
function isNonCardEvidence_v660_(h) {
  var t=compact_v660_((h.company||'')+' '+(h.cardName||'')+' '+(h.evidenceType||''));
  return t.indexOf('비카드')>=0 || t.indexOf('현금영수증')>=0 || t.indexOf('페이머니')>=0 || t.indexOf('머니')>=0;
}
function isLotteEvidence_v660_(h) {
  if (text_v660_(h.lotteFlag).toUpperCase()==='Y') return true;
  return /롯데|LOTTE/i.test(text_v660_(h.merchant));
}
function historyAmountVariants_v660_(h) {
  var out=[];
  function add(n,label){n=Math.round(Number(n)||0);if(n>0&&!out.some(function(x){return x.amount===n;}))out.push({amount:n,label:label});}
  add(h.amount,'APPROVAL');
  if (h.cancelAmount) add(Number(h.amount||0)+Number(h.cancelAmount||0),'NET_AFTER_CANCEL');
  return out;
}
function historyMatchesAmount_v660_(h,amount) { var target=Math.round(Number(amount)||0); return (h.amountVariants||historyAmountVariants_v660_(h)).some(function(x){return x.amount===target;}); }
function historyAmountKind_v660_(h,amount) { var target=Math.round(Number(amount)||0), v=(h.amountVariants||historyAmountVariants_v660_(h)); for(var i=0;i<v.length;i++)if(v[i].amount===target)return v[i].label; return ''; }
function dedupeHistoryCandidates_v660_(rows) {
  var seen={},out=[];
  (rows||[]).forEach(function(h){var key=[h.company,h.cardName,h.cardNumber,h.date,h.time,h.amount,h.approvalNo,h.merchantOrderNo,h.sourceFile].join('|');if(!seen[key]){seen[key]=true;out.push(h);}});
  return out;
}

/* ---------- Diagnostic ---------- */
function vatCardDiagnosticHeaders_v660_() {
  return ['신고연도','반기','주문일','사업자등록번호','쿠팡계정ID','주문번호','롯데결제수단','상세행수','주문매입금액','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','승인일','승인시각','승인번호','승인금액','카드매칭상태','카드매칭근거','후보수','가맹점명','가맹점주문번호','증빙유형','취소/부분취소메모','원본파일','후보요약'];
}
function writeVatCardMatchDiagnostic_v660_(ss,orders) {
  var headers=vatCardDiagnosticHeaders_v660_(), sheet=ss.getSheetByName(LOTTEON_V660_CARD_DIAG_SHEET)||ss.insertSheet(LOTTEON_V660_CARD_DIAG_SHEET);
  var rows=(orders||[]).map(function(o){var m=o.cardMatch||noMatch_v660_('미실행');return [o.year,o.half,o.orderDate,o.business,o.account,o.orderNo,o.lottePayment,o.detailRows,o.purchase,m.company,m.alias,m.cardName,m.cardNumber,m.cardEnd4,m.approvalDate,m.approvalTime,m.approvalNo,m.approvalAmount,m.status,m.reason,m.candidateCount,m.merchant,m.merchantOrderNo,m.evidenceType,m.cancelMemo,m.sourceFile,m.candidateSummary];});
  sheet.clearContents();sheet.getRange(1,1,1,headers.length).setValues([headers]);if(rows.length)sheet.getRange(2,1,rows.length,headers.length).setValues(rows);sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  if(rows.length){sheet.getRange(2,9,rows.length,1).setNumberFormat('#,##0');sheet.getRange(2,18,rows.length,1).setNumberFormat('#,##0');}
  for(var i=0;i<headers.length;i++){var h=headers[i];sheet.setColumnWidth(i+1,/후보요약|취소|근거|원본파일/.test(h)?220:(/카드|사업자|계정|주문번호/.test(h)?135:(/금액/.test(h)?105:90)));}
  return rows.length;
}

/* ---------- Filing summary ---------- */
function aggregateVatBusinessCardHalf_v660_(orders) {
  var map={};
  (orders||[]).forEach(function(o){
    var m=o.cardMatch||noMatch_v660_('미실행');
    var business=o.business||'사업자번호 미매핑';
    var identity=m.status==='AMBIGUOUS'||m.status==='NO_MATCH' ? m.status : [m.company,m.alias,m.cardName,m.cardNumber,m.cardEnd4,m.status].join('|');
    var key=[o.year,o.half,business,identity].join('|');
    if(!map[key])map[key]={year:o.year,half:o.half,business:business,accounts:{},company:m.company||'',alias:m.alias||'',cardName:m.cardName||'',cardNumber:m.cardNumber||'',cardEnd4:m.cardEnd4||'',status:m.status,reasons:{},orders:{},blankOrders:0,sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};
    var x=map[key];if(o.account)x.accounts[o.account]=true;if(o.orderNo)x.orders[o.orderNo]=true;else x.blankOrders++;if(m.reason)x.reasons[m.reason]=true;
    x.sales+=o.sales;x.salesSupply+=o.salesSupply;x.salesVat+=o.salesVat;x.settlement+=o.settlement;x.fee+=o.fee;x.purchase+=o.purchase;x.purchaseSupply+=o.purchaseSupply;x.purchaseVat+=o.purchaseVat;x.payable+=o.payable;x.profit+=o.profit;x.vatProfit+=o.vatProfit;
  });
  var halfRank={'상반기':0,'하반기':1};
  return Object.keys(map).map(function(k){var x=map[k],notes=[];if(x.business==='사업자번호 미매핑')notes.push('사업자번호 미매핑');if(x.blankOrders)notes.push('주문번호 공란 '+x.blankOrders+'건');if(x.status==='AMBIGUOUS'||x.status==='NO_MATCH')notes.push('카드 미확정');return [x.year,x.half,x.business,Object.keys(x.accounts).sort().join(', '),x.company,x.alias,x.cardName,x.cardNumber,x.cardEnd4,x.status,Object.keys(x.reasons).sort().join(' / '),Object.keys(x.orders).length,x.sales,x.salesSupply,x.salesVat,x.settlement,x.fee,x.purchase,x.purchaseSupply,x.purchaseVat,x.payable,x.profit,x.vatProfit,notes.join(' / ')];}).sort(function(a,b){return String(a[0]).localeCompare(String(b[0]))||(halfRank[a[1]]-halfRank[b[1]])||String(a[2]).localeCompare(String(b[2]))||String(a[4]).localeCompare(String(b[4]))||String(a[7]).localeCompare(String(b[7]));});
}
function vatBusinessCardHalfHeaders_v660_() {
  return ['신고연도','반기','사업자등록번호','연결 쿠팡계정ID','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','카드매칭상태','카드매칭근거','주문건수','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];
}
function prependVatBusinessCardHalfSummary_v660_(sheet,periodValues,summary) {
  var headers=vatBusinessCardHalfHeaders_v660_(), detailHeaders=periodValues[0]||[], detailStart=(summary||[]).length+5, maxCols=Math.max(headers.length,detailHeaders.length,1);
  sheet.clearContents();sheet.getRange(1,1).setValue('사업자별 반기 신고요약 (구매카드별)');sheet.getRange(2,1,1,headers.length).setValues([headers]);if(summary.length)sheet.getRange(3,1,summary.length,headers.length).setValues(summary);if(periodValues.length&&detailHeaders.length)sheet.getRange(detailStart,1,periodValues.length,detailHeaders.length).setValues(periodValues);
  sheet.getRange(1,1,1,maxCols).setBackground('#b4c6e7').setFontWeight('bold');sheet.getRange(2,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');if(detailHeaders.length)sheet.getRange(detailStart,1,1,detailHeaders.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  if(summary.length){for(var c=12;c<=23;c++)sheet.getRange(3,c,summary.length,1).setNumberFormat('#,##0');sheet.getRange(3,1,summary.length,11).setNumberFormat('@');}
  if(typeof vatPeriodFormatPlan_v657_==='function'&&detailHeaders.length){var plan=vatPeriodFormatPlan_v657_(detailHeaders);Object.keys(plan).forEach(function(k){if(periodValues.length>1)sheet.getRange(detailStart+1,Number(k)+1,periodValues.length-1,1).setNumberFormat(plan[k]);});}
  for(var i=0;i<maxCols;i++){var h=headers[i]||detailHeaders[i]||'';sheet.setColumnWidth(i+1,/근거|비고/.test(h)?180:(/카드|사업자|계정/.test(h)?135:(/금액|부가세|이익|수수료/.test(h)?110:90)));}
  sheet.setFrozenRows(2);return {summaryRows:summary.length,detailStartRow:detailStart};
}

/* ---------- Generic helpers ---------- */
function findHeaderAlias_v660_(headers,names,fallback){for(var n=0;n<names.length;n++){var wanted=compact_v660_(names[n]);for(var i=0;i<headers.length;i++)if(compact_v660_(headers[i])===wanted)return i;}return fallback;}
function textAt_v660_(row,index){return index<0?'':text_v660_(row[index]);}
function text_v660_(v){return String(v==null?'':v).trim();}
function number_v660_(v){if(typeof v==='number')return isNaN(v)?0:v;var n=Number(String(v==null?'0':v).replace(/[원,\s]/g,''));return isNaN(n)?0:n;}
function compact_v660_(v){return text_v660_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function digits_v660_(v){return text_v660_(v).replace(/\D/g,'');}
function normalizeDateText_v660_(v){var s=text_v660_(v);if(!s)return '';var m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(!m)return '';return m[1]+'-'+pad2_v660_(m[2])+'-'+pad2_v660_(m[3]);}
function normalizeOrderDate_v660_(raw,year){if(Object.prototype.toString.call(raw)==='[object Date]'&&!isNaN(raw.getTime()))return raw.getFullYear()+'-'+pad2_v660_(raw.getMonth()+1)+'-'+pad2_v660_(raw.getDate());var s=text_v660_(raw),full=normalizeDateText_v660_(s);if(full)return full;var m=s.match(/^(\d{1,2})[.\/-](\d{1,2})$/);return m&&year?String(year)+'-'+pad2_v660_(m[1])+'-'+pad2_v660_(m[2]):'';}
function pad2_v660_(n){n=String(n);return n.length<2?'0'+n:n;}
function normalizeBusinessNo_v660_(v){return digits_v660_(v);}
function normalizeAccountCode_v660_(v){return text_v660_(v).toLowerCase().replace(/^beliun/,'');}
function normalizeCardCompany_v660_(v){var s=compact_v660_(v);if(!s)return '';if(s.indexOf('kb')>=0||s.indexOf('국민')>=0)return 'KB국민카드';if(s.indexOf('신한')>=0)return '신한카드';if(s.indexOf('롯데')>=0)return '롯데카드';if(s.indexOf('우리')>=0)return '우리카드';if(s.indexOf('농협')>=0||s.indexOf('nh')>=0)return 'NH농협카드';if(s.indexOf('삼성')>=0)return '삼성카드';if(s.indexOf('하나')>=0)return '하나카드';if(s.indexOf('현대')>=0)return '현대카드';if(s.indexOf('비카드')>=0||s.indexOf('카카오')>=0||s.indexOf('페이머니')>=0)return '비카드';return text_v660_(v);}
function normalizeCardName_v660_(v){return compact_v660_(v).replace(/^본인/,'').replace(/[0-9*]+$/,'');}
function historyCandidateLabel_v660_(h){return [h.company,h.cardName,h.cardEnd4||h.cardNumber,h.date,h.time,h.amount,h.approvalNo,h.merchant].filter(Boolean).join(' / ');}
function masterCandidateLabel_v660_(m){return [m.company,m.alias,m.cardName,m.cardEnd4||m.cardNumber,m.business||m.account,m.startDate,m.endDate].filter(Boolean).join(' / ');}
