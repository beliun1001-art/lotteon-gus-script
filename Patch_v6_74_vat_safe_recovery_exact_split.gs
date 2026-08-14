/**
 * v6.74 Issue #74: guarded recovery for residual VAT-card NO_MATCH orders.
 *
 * Production rule is intentionally two-pass:
 * 1) run the existing v6.64 -> v6.70 matcher for every order first and reserve all
 *    canonical evidence consumed by existing exact matches;
 * 2) only then recover orders that are still NO_MATCH.
 *
 * This prevents a recovered order from stealing evidence that the pre-v6.74 matcher
 * would have assigned to a later order.
 *
 * Recovery A: one unique valid LOTTE card evidence outside the old 0..+7 window,
 *             within +/-14 days, exact amount, and that unique evidence is exactly -1 day.
 * Recovery B: one unique unused pair of valid LOTTE card evidence inside 0..+7 days,
 *             same physical-card identity, whose effective amounts sum exactly to purchase.
 */
var LOTTEON_PATCH_V674_VAT_SAFE_RECOVERY_EXACT_SPLIT_LOADED = true;
var LOTTEON_V674_OUTSIDE_ABS_DAYS = 14;
var LOTTEON_V674_FORWARD_DAYS = 7;

var __baseMatchVatOrderCardCanonical_v674_ = typeof matchVatOrderCardCanonical_v664_ === 'function'
  ? matchVatOrderCardCanonical_v664_
  : null;

if (__baseMatchVatOrderCardCanonical_v674_) {
  allocateVatPurchaseCards_v664_ = function(orders, history, master) {
    var used = {};
    var sortedOrders = (orders || []).slice().sort(function(a,b) {
      return String(a.orderDate || '').localeCompare(String(b.orderDate || '')) ||
        String(a.orderNo || '').localeCompare(String(b.orderNo || '')) ||
        Number(a.purchase || 0) - Number(b.purchase || 0);
    });

    // Pass 1: reproduce the complete pre-v6.74 result first.
    sortedOrders.forEach(function(order) {
      order.cardMatch = __baseMatchVatOrderCardCanonical_v674_.call(
        this, order, history || [], master || [], used
      );
    });

    var recoveredDate = 0, recoveredSplit = 0;

    // Pass 2: only residual NO_MATCH orders can use the new recovery rules.
    sortedOrders.forEach(function(order) {
      var current = order && order.cardMatch || {};
      if (current.status !== 'NO_MATCH') return;

      var single = recoverMinusOneExact_v674_(order, history || [], master || [], used);
      if (single) {
        order.cardMatch = single.result;
        used[String(single.key)] = true;
        recoveredDate++;
        return;
      }

      var split = recoverSplitExact_v674_(order, history || [], master || [], used);
      if (split) {
        order.cardMatch = split.result;
        used[String(split.keys[0])] = true;
        used[String(split.keys[1])] = true;
        recoveredSplit++;
      }
    });

    return {
      usedEvidence:Object.keys(used).length,
      v674RecoveredDate:recoveredDate,
      v674RecoveredSplit:recoveredSplit,
      v674RecoveredTotal:recoveredDate + recoveredSplit
    };
  };
}

function validRecoveryEvidence_v674_(h) {
  return !!h && !h.nonCard && !h.cancelRow && !h.v664FullyCanceled && !!h.lotteEvidence &&
    Number(h.v664EffectiveAmount || h.amount || 0) > 0;
}

function recoveryEvidenceKey_v674_(h) {
  return String(h && h.v664CanonicalKey || '');
}

function paymentIssuerConflict_v674_(order, h) {
  if (!order || !h) return false;
  if (typeof classifyVatTrackingPayment_v666_ === 'function') {
    var rule = classifyVatTrackingPayment_v666_(order.lottePayment);
    if (rule && rule.kind === 'ISSUER_CARD' && rule.issuer) {
      return normalizeCardCompany_v660_(h.company) !== rule.issuer;
    }
    return false;
  }
  var raw = typeof compact_v660_ === 'function' ? compact_v660_(order.lottePayment) : String(order.lottePayment || '').toLowerCase();
  if (!raw || raw.indexOf('lpay') >= 0 || raw.indexOf('엘페이') >= 0 || raw.indexOf('토스') >= 0 || raw.indexOf('카카오') >= 0) return false;
  var issuer = typeof normalizeCardCompany_v660_ === 'function' ? normalizeCardCompany_v660_(order.lottePayment) : '';
  return !!issuer && normalizeCardCompany_v660_(h.company) !== issuer;
}

function outsideExactCandidates_v674_(order, history) {
  if (!order || !order.orderDate || !Number(order.purchase || 0)) return [];
  return (history || []).filter(function(h) {
    if (!validRecoveryEvidence_v674_(h)) return false;
    if (Number(h.v664EffectiveAmount || h.amount || 0) !== Number(order.purchase || 0)) return false;
    var lag = daysBetween_v664_(order.orderDate, h.date);
    return Math.abs(lag) <= LOTTEON_V674_OUTSIDE_ABS_DAYS && !(lag >= 0 && lag <= LOTTEON_V674_FORWARD_DAYS);
  });
}

function recoverMinusOneExact_v674_(order, history, master, used) {
  var candidates = outsideExactCandidates_v674_(order, history || []);
  if (candidates.length !== 1) return null;
  var h = candidates[0];
  if (daysBetween_v664_(order.orderDate, h.date) !== -1) return null;
  var key = recoveryEvidenceKey_v674_(h);
  if (!key || used[String(key)]) return null;
  if (paymentIssuerConflict_v674_(order, h)) return null;

  var result = matchCanonicalHistory_v664_(order, h, master || [], -1, 1) || {};
  if (result.status !== 'MATCHED') return null;
  result.reason = 'v6.74_-1일_exact_단일미사용증빙_1:1회수';
  result.v674Fallback = true;
  result.v674RecoveryKind = 'MINUS_ONE_EXACT';
  result.allocationLagDays = -1;
  result.allocationCandidateCount = 1;
  result.canonicalEvidenceKey = key;
  return { key:key, result:result };
}

function forwardRecoveryEvidence_v674_(order, history) {
  if (!order || !order.orderDate || !Number(order.purchase || 0)) return [];
  return (history || []).filter(function(h) {
    if (!validRecoveryEvidence_v674_(h)) return false;
    var lag = daysBetween_v664_(order.orderDate, h.date);
    return lag >= 0 && lag <= LOTTEON_V674_FORWARD_DAYS;
  });
}

function physicalIdentity_v674_(h, master) {
  if (!h) return '';
  if (typeof cardIdentityKey_v662_ === 'function') return String(cardIdentityKey_v662_(h, master || []) || '');
  var e = typeof enrichHistoryFromMaster_v660_ === 'function' ? enrichHistoryFromMaster_v660_(h, master || []) : h;
  var company = normalizeCardCompany_v660_(e.company || h.company || '');
  var number = String(e.cardNumber || h.cardNumber || '').replace(/\D/g, '');
  var end4 = String(e.cardEnd4 || h.cardEnd4 || '').replace(/\D/g, '').slice(-4);
  var name = typeof compact_v660_ === 'function' ? compact_v660_(e.cardName || h.cardName || '') : String(e.cardName || h.cardName || '').toLowerCase();
  if (!company) return '';
  if (end4) return company + '|END4:' + end4;
  if (number) return company + '|NUM:' + number;
  if (name) return company + '|NAME:' + name;
  return '';
}

function splitRecoveryPairs_v674_(order, history, master, used) {
  var rows = forwardRecoveryEvidence_v674_(order, history || []);
  var pairs = [];
  for (var i=0; i<rows.length; i++) {
    for (var j=i+1; j<rows.length; j++) {
      var a = rows[i], b = rows[j];
      if (Number(a.v664EffectiveAmount || a.amount || 0) + Number(b.v664EffectiveAmount || b.amount || 0) !== Number(order.purchase || 0)) continue;
      var ka = recoveryEvidenceKey_v674_(a), kb = recoveryEvidenceKey_v674_(b);
      if (!ka || !kb || ka === kb || used[String(ka)] || used[String(kb)]) continue;
      var ia = physicalIdentity_v674_(a, master || []), ib = physicalIdentity_v674_(b, master || []);
      if (!ia || ia !== ib) continue;
      pairs.push({ a:a, b:b, keys:[ka,kb], identity:ia });
    }
  }
  return pairs;
}

function enrichedRecoveryEvidence_v674_(h, master) {
  return typeof enrichHistoryFromMaster_v660_ === 'function'
    ? (enrichHistoryFromMaster_v660_(h, master || []) || h)
    : h;
}

function recoveryCandidateLabel_v674_(h) {
  if (typeof historyCandidateLabel_v660_ === 'function') return historyCandidateLabel_v660_(h);
  return [h.date, Number(h.v664EffectiveAmount || h.amount || 0), h.company || '', h.approvalNo || ''].join('|');
}

function recoverSplitExact_v674_(order, history, master, used) {
  var pairs = splitRecoveryPairs_v674_(order, history || [], master || [], used || {});
  if (pairs.length !== 1) return null;
  var p = pairs[0], a = p.a, b = p.b;
  var ea = enrichedRecoveryEvidence_v674_(a, master || []), eb = enrichedRecoveryEvidence_v674_(b, master || []);
  var company = ea.company || eb.company || a.company || b.company || '';
  var alias = ea.alias || eb.alias || '';
  var cardName = ea.cardName || eb.cardName || a.cardName || b.cardName || '';
  var cardNumber = ea.cardNumber || eb.cardNumber || a.cardNumber || b.cardNumber || '';
  var cardEnd4 = ea.cardEnd4 || eb.cardEnd4 || a.cardEnd4 || b.cardEnd4 || '';
  if (!cardEnd4 && cardNumber) cardEnd4 = String(cardNumber).replace(/\D/g, '').slice(-4);
  var dateA = String(a.date || ''), dateB = String(b.date || '');
  var sourceA = String(a.sourceFile || ''), sourceB = String(b.sourceFile || '');
  var merchantA = String(a.merchant || ''), merchantB = String(b.merchant || '');
  var approvalA = String(a.approvalNo || ''), approvalB = String(b.approvalNo || '');

  var result = {
    status:'MATCHED',
    reason:'v6.74_split_exact_동일카드_2증빙_1:1회수',
    candidateCount:2,
    company:company,
    alias:alias,
    cardName:cardName,
    cardNumber:cardNumber,
    cardEnd4:cardEnd4,
    approvalDate:dateA === dateB ? dateA : dateA + '+' + dateB,
    approvalTime:'',
    approvalNo:'SPLIT:' + approvalA + '+' + approvalB,
    approvalAmount:Number(order.purchase || 0),
    merchant:merchantA === merchantB ? merchantA : merchantA + ' + ' + merchantB,
    merchantOrderNo:'',
    evidenceType:'분할결제_2증빙',
    sourceFile:sourceA === sourceB ? sourceA : sourceA + ' + ' + sourceB,
    cancelMemo:'',
    candidateSummary:recoveryCandidateLabel_v674_(a) + ' || ' + recoveryCandidateLabel_v674_(b),
    allocationCandidateCount:2,
    canonicalEvidenceKeys:p.keys.slice(),
    v674Fallback:true,
    v674RecoveryKind:'SPLIT_EXACT'
  };
  return { keys:p.keys.slice(), result:result };
}
