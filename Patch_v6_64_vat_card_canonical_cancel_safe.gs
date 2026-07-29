/** v6.64 Issue #20: canonical purchase evidence + cancellation-safe one-to-one allocation. */
var LOTTEON_PATCH_V664_VAT_CARD_CANONICAL_CANCEL_SAFE_LOADED = true;
var LOTTEON_V664_MAX_FORWARD_DAYS = 7;

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
  var rawHistory = loadVatCardHistory_v660_(ss);
  var master = loadVatCardMaster_v660_(ss);
  var canonical = canonicalizeVatHistory_v664_(rawHistory, master);

  allocateVatPurchaseCards_v664_(orders, canonical, master);

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
    canonicalEvidenceRows:canonical.length,
    fullCancelEvidenceExcluded:canonical.filter(function(h){ return !!h.v664FullyCanceled; }).length,
    allocationWindowDays:LOTTEON_V664_MAX_FORWARD_DAYS
  };
  try {
    if (typeof writeVatCardUnmatchedAnalysis_v662_ === 'function') result.unmatchedAnalysisRows = writeVatCardUnmatchedAnalysis_v662_(ss);
  } catch (e) {
    result.unmatchedAnalysisError = String(e && e.message ? e.message : e);
  }
  return result;
};

function canonicalizeVatHistory_v664_(history, master) {
  var cardGroups = {}, singles = [];
  (history || []).forEach(function(h) {
    if (!h) return;
    var issuer = normalizeCardCompany_v660_(h.company);
    var approval = text_v660_(h.approvalNo);
    if (!h.nonCard && issuer && approval) {
      var key = 'CARD|' + issuer + '|' + approval;
      if (!cardGroups[key]) cardGroups[key] = [];
      cardGroups[key].push(h);
    } else {
      var single = canonicalSingleEvidence_v664_(h, 'ROW|' + String(h.rowNo || singles.length + 1));
      if (single) singles.push(single);
    }
  });

  var out = singles;
  Object.keys(cardGroups).forEach(function(key) {
    var item = canonicalCardGroup_v664_(key, cardGroups[key], master || []);
    if (item) out.push(item);
  });
  return out;
}

function canonicalSingleEvidence_v664_(h, key) {
  var x = shallowCopyEvidence_v664_(h);
  x.v664CanonicalKey = key;
  var original = Math.abs(Number(h.amount || 0));
  var cancel = Math.abs(Number(h.cancelAmount || 0));
  if (h.cancelRow && Number(h.amount || 0) < 0) return null;
  var effective = Math.max(original - cancel, 0);
  x.v664OriginalAmount = original;
  x.v664CancelAmount = cancel;
  x.v664EffectiveAmount = effective;
  x.v664FullyCanceled = original > 0 && effective === 0;
  x.amount = effective;
  x.cancelRow = !!x.v664FullyCanceled || !!(h.cancelRow && original <= 0);
  x.amountVariants = effective > 0 ? [effective] : [];
  return x;
}

function canonicalCardGroup_v664_(key, rows, master) {
  if (!rows || !rows.length) return null;
  var positives = rows.filter(function(h){ return Number(h.amount || 0) > 0; });
  if (!positives.length) return null;
  positives.sort(function(a,b){ return evidenceRichness_v664_(b) - evidenceRichness_v664_(a); });
  var rep = positives[0];
  var original = 0, cancel = 0, cancelDate = '';
  rows.forEach(function(h) {
    var a = Number(h.amount || 0);
    if (a > original) original = a;
    var c = Math.abs(Number(h.cancelAmount || 0));
    if (a < 0) c = Math.max(c, Math.abs(a));
    if (c > cancel) { cancel = c; cancelDate = h.cancelDate || h.date || cancelDate; }
  });
  cancel = Math.min(cancel, original);
  var effective = Math.max(original - cancel, 0);
  var x = shallowCopyEvidence_v664_(rep);
  x.v664CanonicalKey = key;
  x.v664OriginalAmount = original;
  x.v664CancelAmount = cancel;
  x.v664EffectiveAmount = effective;
  x.v664FullyCanceled = original > 0 && effective === 0;
  x.v664MergedRows = rows.map(function(h){ return h.rowNo; });
  x.v664CancelDate = cancelDate;
  x.amount = effective;
  x.cancelAmount = cancel ? -cancel : 0;
  x.cancelDate = cancelDate || x.cancelDate || '';
  x.cancelRow = !!x.v664FullyCanceled;
  x.lotteEvidence = rows.some(function(h){ return !!h.lotteEvidence; });
  x.nonCard = false;
  x.amountVariants = effective > 0 ? [effective] : [];

  var enriched = enrichHistoryFromMaster_v660_(x, master || []);
  if (enriched.company) x.company = enriched.company;
  if (enriched.cardName) x.cardName = enriched.cardName;
  if (enriched.cardNumber) x.cardNumber = enriched.cardNumber;
  if (enriched.cardEnd4) x.cardEnd4 = enriched.cardEnd4;
  return x;
}

function evidenceRichness_v664_(h) {
  var n = 0;
  if (text_v660_(h.evidenceType) === '카드이용내역') n += 100;
  if (text_v660_(h.cardEnd4)) n += 20;
  if (text_v660_(h.cardNumber)) n += 10;
  if (text_v660_(h.cardName)) n += 5;
  if (!h.cancelRow) n += 1;
  return n;
}

function shallowCopyEvidence_v664_(h) {
  var x = {};
  Object.keys(h || {}).forEach(function(k){ x[k] = h[k]; });
  return x;
}

function allocateVatPurchaseCards_v664_(orders, history, master) {
  var used = {};
  var sortedOrders = (orders || []).slice().sort(function(a,b) {
    return String(a.orderDate || '').localeCompare(String(b.orderDate || '')) ||
      String(a.orderNo || '').localeCompare(String(b.orderNo || '')) || Number(a.purchase || 0) - Number(b.purchase || 0);
  });
  sortedOrders.forEach(function(order) {
    order.cardMatch = matchVatOrderCardCanonical_v664_(order, history || [], master || [], used);
  });
  return { usedEvidence:Object.keys(used).length };
}

function matchVatOrderCardCanonical_v664_(order, history, master, used) {
  if (!order || !order.orderDate || !Number(order.purchase || 0)) return noMatch_v660_('주문일/매입금액 없음');
  var issuer = normalizeCardCompany_v660_(order.lottePayment);

  for (var lag=0; lag<=LOTTEON_V664_MAX_FORWARD_DAYS; lag++) {
    var targetDate = shiftDate_v662_(order.orderDate, lag);
    var base = (history || []).filter(function(h) {
      if (!h || h.cancelRow || h.v664FullyCanceled || !h.lotteEvidence) return false;
      if (used[String(h.v664CanonicalKey)]) return false;
      if (h.date !== targetDate) return false;
      return Number(h.v664EffectiveAmount || h.amount || 0) === Number(order.purchase || 0);
    });

    if (issuer) {
      var issuerRows = base.filter(function(h){ return normalizeCardCompany_v660_(h.company) === issuer; });
      if (issuerRows.length) base = issuerRows;
    } else if (order.lottePayment) {
      var paymentRows = filterEvidenceByLottePayment_v660_(base, order.lottePayment);
      if (paymentRows.length) base = paymentRows;
    }

    if (!base.length) continue;
    var identities = {};
    base.forEach(function(h) {
      var id = cardIdentityKey_v662_(h, master || []);
      if (id) identities[id] = true;
      else identities['UNKNOWN|' + String(h.v664CanonicalKey)] = true;
    });
    if (Object.keys(identities).length !== 1) return ambiguousForwardMatch_v663_(base, lag, '서로 다른 구매카드 후보');

    base.sort(function(a,b){ return String(a.time || '').localeCompare(String(b.time || '')) || String(a.v664CanonicalKey).localeCompare(String(b.v664CanonicalKey)); });
    var chosen = base[0];
    used[String(chosen.v664CanonicalKey)] = true;
    return matchCanonicalHistory_v664_(order, chosen, master, lag, base.length);
  }

  if (order.orderDate >= LOTTEON_V660_MASTER_CUTOFF) {
    var masterCandidates = findPostJulyMasterCandidates_v660_(order, master);
    if (masterCandidates.length === 1) return matchFromMaster_v660_(masterCandidates[0], '카드마스터_7월이후_사업자+카드사+적용기간');
    if (masterCandidates.length > 1) return ambiguousMasterMatch_v660_(masterCandidates);
  }

  var canceled = (history || []).filter(function(h) {
    if (!h || !h.v664FullyCanceled || !h.lotteEvidence) return false;
    if (Number(h.v664OriginalAmount || 0) !== Number(order.purchase || 0)) return false;
    var d = daysBetween_v664_(order.orderDate, h.date);
    return d >= 0 && d <= LOTTEON_V664_MAX_FORWARD_DAYS;
  }).length;
  var reason = order.orderDate < LOTTEON_V660_MASTER_CUTOFF ? '상반기 거래내역 0~+7일 유효증빙 매칭 없음' : '거래내역/카드마스터 매칭 없음';
  if (canceled) reason += ' / 완전취소 증빙 ' + canceled + '건 제외';
  return noMatch_v660_(reason);
}

function matchCanonicalHistory_v664_(order, h, master, lag, candidateCount) {
  var clone = shallowCopyEvidence_v664_(h);
  clone.amount = Number(h.v664EffectiveAmount || h.amount || 0);
  clone.cancelAmount = 0;
  var status = h.nonCard ? 'NON_CARD' : 'MATCHED';
  var reason = '거래내역_정합성보정_마켓주문일' + (lag ? '+' + lag + '일' : '당일') + '_실결제금액_1:1할당';
  if (candidateCount > 1) reason += '_동일구매수단확정(' + candidateCount + '건)';
  var result = matchFromHistory_v660_(order, clone, master, status, reason);
  result.reason = reason + (Number(h.v664CancelAmount || 0) ? ' / NET_AFTER_CANCEL' : ' / APPROVAL');
  result.approvalAmount = Number(h.v664OriginalAmount || h.amount || 0);
  result.historyRowNo = h.rowNo;
  result.canonicalEvidenceKey = h.v664CanonicalKey;
  result.allocationLagDays = lag;
  result.allocationCandidateCount = candidateCount;
  if (Number(h.v664CancelAmount || 0)) {
    result.cancelMemo = '취소일 ' + (h.v664CancelDate || h.cancelDate || '-') + ' / 취소금액 -' + Number(h.v664CancelAmount || 0) + ' / 취소반영 실결제금액 ' + Number(h.v664EffectiveAmount || 0);
  }
  return result;
}

function daysBetween_v664_(fromText, toText) {
  var a = String(fromText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  var b = String(toText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!a || !b) return 99999;
  var ad = Date.UTC(Number(a[1]), Number(a[2])-1, Number(a[3]));
  var bd = Date.UTC(Number(b[1]), Number(b[2])-1, Number(b[3]));
  return Math.round((bd-ad)/86400000);
}
