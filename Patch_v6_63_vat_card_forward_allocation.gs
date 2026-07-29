/**
 * v6.63 Issue #20 - purchase-card reconciliation using market-order -> purchase forward allocation.
 *
 * Root cause from operating data:
 * - 부가세_신고자료의 날짜 is the Coupang/market order date, not the LOTTEON card approval date.
 * - LOTTEON purchase happens after the resale order, so exact same-calendar-day matching is too strict.
 *
 * Safety rules:
 * - Pre-July cards are never inferred from business/account.
 * - Card-statement evidence must have the exact purchase amount and LOTTE evidence.
 * - Search only from market order date forward, never before it.
 * - Evidence rows are allocated one-to-one and cannot be reused by multiple resale orders.
 * - Earliest evidence day wins. Multiple candidates are auto-resolved only when every candidate
 *   on that earliest day is the same physical card/payment identity; otherwise AMBIGUOUS.
 * - Search window is 0..7 days to cover delayed sourcing while keeping evidence bounded.
 */
var LOTTEON_PATCH_V663_VAT_CARD_FORWARD_ALLOCATION_LOADED = true;
var LOTTEON_V663_MAX_FORWARD_DAYS = 7;

buildVatPurchaseCardReconciliation_v660_ = function(ss) {
  ss = ss || SpreadsheetApp.getActive();
  var detail = ss && ss.getSheetByName && ss.getSheetByName('부가세_신고자료');
  var periodSheet = ss && ss.getSheetByName && ss.getSheetByName('부가세_기간별');
  if (!detail || !periodSheet || detail.getLastRow() < 1 || periodSheet.getLastRow() < 1) {
    return { summaryRows:0, matchedOrders:0, ambiguousOrders:0, noMatchOrders:0, reason:'MISSING_VAT_SHEET' };
  }

  var detailValues = detail.getDataRange().getValues();
  var periodValues = periodSheet.getDataRange().getValues();
  var orders = groupVatDetailByOrder_v660_(detailValues);
  var history = loadVatCardHistory_v660_(ss);
  var master = loadVatCardMaster_v660_(ss);

  allocateVatPurchaseCards_v663_(orders, history, master);

  var stats = { matchedOrders:0, ambiguousOrders:0, noMatchOrders:0, nonCardOrders:0, masterMatchedOrders:0 };
  orders.forEach(function(order) {
    var s = order.cardMatch && order.cardMatch.status || 'NO_MATCH';
    if (s === 'MATCHED') stats.matchedOrders++;
    else if (s === 'MASTER_MATCHED') { stats.matchedOrders++; stats.masterMatchedOrders++; }
    else if (s === 'NON_CARD') stats.nonCardOrders++;
    else if (s === 'AMBIGUOUS') stats.ambiguousOrders++;
    else stats.noMatchOrders++;
  });

  writeVatCardMatchDiagnostic_v660_(ss, orders);
  var summary = aggregateVatBusinessCardHalf_v660_(orders);
  prependVatBusinessCardHalfSummary_v660_(periodSheet, periodValues, summary);

  var result = {
    summaryRows:summary.length,
    orderRows:orders.length,
    matchedOrders:stats.matchedOrders,
    masterMatchedOrders:stats.masterMatchedOrders,
    nonCardOrders:stats.nonCardOrders,
    ambiguousOrders:stats.ambiguousOrders,
    noMatchOrders:stats.noMatchOrders,
    allocationWindowDays:LOTTEON_V663_MAX_FORWARD_DAYS
  };
  try {
    if (typeof writeVatCardUnmatchedAnalysis_v662_ === 'function') {
      result.unmatchedAnalysisRows = writeVatCardUnmatchedAnalysis_v662_(ss);
    }
  } catch (e) {
    result.unmatchedAnalysisError = String(e && e.message ? e.message : e);
  }
  return result;
};

function allocateVatPurchaseCards_v663_(orders, history, master) {
  var used = {};
  var sortedOrders = (orders || []).slice().sort(function(a,b) {
    return String(a.orderDate || '').localeCompare(String(b.orderDate || '')) ||
      String(a.orderNo || '').localeCompare(String(b.orderNo || '')) ||
      Number(a.purchase || 0) - Number(b.purchase || 0);
  });

  sortedOrders.forEach(function(order) {
    order.cardMatch = matchVatOrderCardForwardAllocated_v663_(order, history || [], master || [], used);
  });
  return { usedHistoryRows:Object.keys(used).length };
}

function matchVatOrderCardForwardAllocated_v663_(order, history, master, used) {
  if (!order || !order.orderDate || !Number(order.purchase || 0)) {
    return noMatch_v660_('주문일/매입금액 없음');
  }

  var issuer = normalizeCardCompany_v660_(order.lottePayment);
  for (var lag = 0; lag <= LOTTEON_V663_MAX_FORWARD_DAYS; lag++) {
    var targetDate = shiftDate_v662_(order.orderDate, lag);
    var base = (history || []).filter(function(h) {
      if (!h || h.cancelRow || !h.lotteEvidence) return false;
      if (used[String(h.rowNo)]) return false;
      if (h.date !== targetDate) return false;
      if (!historyMatchesAmount_v660_(h, order.purchase)) return false;
      return true;
    });

    if (issuer) {
      var issuerRows = base.filter(function(h) { return normalizeCardCompany_v660_(h.company) === issuer || h.nonCard; });
      if (issuerRows.length) base = issuerRows;
    } else if (order.lottePayment) {
      var paymentRows = filterEvidenceByLottePayment_v660_(base, order.lottePayment);
      if (paymentRows.length) base = paymentRows;
    }

    var candidates = dedupeHistoryCandidates_v660_(base);
    if (!candidates.length) continue;

    var identityMap = {};
    var invalidIdentity = false;
    candidates.forEach(function(h) {
      var identity = cardIdentityKey_v662_(h, master);
      if (!identity) invalidIdentity = true;
      else identityMap[identity] = true;
    });
    var identities = Object.keys(identityMap);
    if (invalidIdentity || identities.length !== 1) {
      return ambiguousForwardMatch_v663_(candidates, lag, invalidIdentity ? '카드식별정보 불충분' : '서로 다른 구매카드 후보');
    }

    candidates.sort(function(a,b) {
      return String(a.time || '').localeCompare(String(b.time || '')) || Number(a.rowNo || 0) - Number(b.rowNo || 0);
    });
    var chosen = candidates[0];
    used[String(chosen.rowNo)] = true;
    var allNonCard = candidates.every(function(h) { return !!h.nonCard; });
    var status = allNonCard ? 'NON_CARD' : 'MATCHED';
    var reason = '거래내역_마켓주문일' + (lag ? '+' + lag + '일' : '당일') + '_일자+금액_1:1할당';
    if (candidates.length > 1) reason += '_동일구매수단확정(' + candidates.length + '건)';
    var result = matchFromHistory_v660_(order, chosen, master, status, reason);
    result.historyRowNo = chosen.rowNo;
    result.allocationLagDays = lag;
    result.allocationCandidateCount = candidates.length;
    return result;
  }

  if (order.orderDate >= LOTTEON_V660_MASTER_CUTOFF) {
    var masterCandidates = findPostJulyMasterCandidates_v660_(order, master);
    if (masterCandidates.length === 1) return matchFromMaster_v660_(masterCandidates[0], '카드마스터_7월이후_사업자+카드사+적용기간');
    if (masterCandidates.length > 1) return ambiguousMasterMatch_v660_(masterCandidates);
  }
  return noMatch_v660_(order.orderDate < LOTTEON_V660_MASTER_CUTOFF ? '상반기 거래내역 0~+' + LOTTEON_V663_MAX_FORWARD_DAYS + '일 증빙 매칭 없음' : '거래내역/카드마스터 매칭 없음');
}

function ambiguousForwardMatch_v663_(candidates, lag, detail) {
  return {
    status:'AMBIGUOUS',
    reason:'마켓주문일' + (lag ? '+' + lag + '일' : '당일') + ' 동일금액 후보 ' + candidates.length + '건 / ' + detail,
    candidateCount:candidates.length,
    company:'', alias:'', cardName:'', cardNumber:'', cardEnd4:'',
    approvalDate:'', approvalTime:'', approvalNo:'', approvalAmount:0,
    merchant:'', merchantOrderNo:'', evidenceType:'', sourceFile:'', cancelMemo:'',
    candidateSummary:candidates.map(historyCandidateLabel_v660_).join(' || '),
    allocationLagDays:lag
  };
}

/* Extend diagnostics so remaining NO_MATCH shows forward exact-amount evidence by lag. */
var __baseAnalyzeUnmatchedOrder_v663_ = typeof analyzeUnmatchedOrder_v662_ === 'function' ? analyzeUnmatchedOrder_v662_ : null;
if (__baseAnalyzeUnmatchedOrder_v663_) {
  analyzeUnmatchedOrder_v662_ = function(orderDate, purchase, payment, history, status) {
    var result = __baseAnalyzeUnmatchedOrder_v663_.apply(this, arguments) || {};
    var counts = [];
    for (var lag=0; lag<=LOTTEON_V663_MAX_FORWARD_DAYS; lag++) {
      var d = shiftDate_v662_(orderDate, lag);
      var rows = (history || []).filter(function(h) {
        return h && !h.cancelRow && h.lotteEvidence && h.date === d && historyMatchesAmount_v660_(h, purchase);
      });
      counts.push(rows.length);
    }
    result.forwardExactByLag = counts.join('/');
    if (status === 'NO_MATCH' && counts.some(function(n){return n>0;})) {
      result.suggestion = '0~+7일 동일금액 후보가 있으나 1:1 evidence 소진/식별 충돌 확인';
    }
    return result;
  };
}
