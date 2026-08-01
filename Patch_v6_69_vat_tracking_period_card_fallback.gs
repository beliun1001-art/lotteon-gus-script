/** v6.69 Issue #28: tracking-payment period-single-card fallback after exact amount matching fails. */
var LOTTEON_PATCH_V669_VAT_TRACKING_PERIOD_CARD_FALLBACK_LOADED = true;

function vatHalfBounds_v669_(orderDate) {
  var m = String(orderDate || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return null;
  return Number(m[2]) <= 6
    ? { start:m[1] + '-01-01', end:m[1] + '-06-30' }
    : { start:m[1] + '-07-01', end:m[1] + '-12-31' };
}

function vatDateInSameHalf_v669_(dateText, orderDate) {
  var b = vatHalfBounds_v669_(orderDate), d = String(dateText || '');
  return !!b && !!d && d >= b.start && d <= b.end;
}

function vatMasterActiveOnDate_v669_(m, orderDate) {
  if (!m || (typeof isActiveMaster_v660_ === 'function' && !isActiveMaster_v660_(m))) return false;
  var d = String(orderDate || '');
  if (!d || (m.startDate && String(m.startDate) > d) || (m.endDate && String(m.endDate) < d)) return false;
  return true;
}

function vatPeriodCandidateIdentity_v669_(candidate, master) {
  if (!candidate) return null;
  var source = candidate.v669Source || 'history';
  var e = source === 'master' ? candidate :
    (typeof enrichHistoryFromMaster_v660_ === 'function' ? enrichHistoryFromMaster_v660_(candidate, master || []) : candidate);
  var company = normalizeCardCompany_v660_(e.company || candidate.company || '');
  var number = String(e.cardNumber || candidate.cardNumber || '').trim();
  var end4 = typeof normalizeVatCardEnd4_v667_ === 'function'
    ? normalizeVatCardEnd4_v667_(e.cardEnd4 || candidate.cardEnd4 || '', number)
    : String(e.cardEnd4 || candidate.cardEnd4 || '').replace(/\D/g, '').slice(-4);
  var name = String(e.cardName || candidate.cardName || '').trim();
  var alias = String(e.alias || candidate.alias || '').trim();
  var digits = number.replace(/\D/g, '');
  var key = company + '|';
  if (end4) key += 'END4:' + end4;
  else if (digits) key += 'NUM:' + digits;
  else if (name) key += 'NAME:' + compact_v660_(name);
  else return null;
  return {
    key:key, company:company, alias:alias, cardName:name, cardNumber:number, cardEnd4:end4,
    sourceFile:String(candidate.sourceFile || ''),
    label:source === 'master' && typeof masterCandidateLabel_v660_ === 'function'
      ? masterCandidateLabel_v660_(candidate)
      : (typeof historyCandidateLabel_v660_ === 'function' ? historyCandidateLabel_v660_(candidate) : key)
  };
}

function vatPeriodTrackingCandidates_v669_(order, history, master, rule) {
  var rows = [], date = String(order && order.orderDate || '');
  if (!date || !rule || rule.kind === 'UNKNOWN' || rule.kind === 'KAKAO_MONEY') return rows;
  (history || []).forEach(function(h) {
    if (!h || h.nonCard || h.cancelRow || h.v664FullyCanceled || !vatDateInSameHalf_v669_(h.date, date)) return;
    if (rule.kind === 'KAKAO_CARD') {
      if (!(h.v666KakaoCard || (typeof isKakaoCardEvidence_v666_ === 'function' && isKakaoCardEvidence_v666_(h)))) return;
    } else if (normalizeCardCompany_v660_(h.company) !== rule.issuer) return;
    var x = {}; Object.keys(h).forEach(function(k) { x[k] = h[k]; }); x.v669Source = 'history'; rows.push(x);
  });
  if (rule.kind === 'ISSUER_CARD') {
    (master || []).forEach(function(m) {
      if (!vatMasterActiveOnDate_v669_(m, date) || normalizeCardCompany_v660_(m.company) !== rule.issuer) return;
      var x = {}; Object.keys(m).forEach(function(k) { x[k] = m[k]; }); x.v669Source = 'master'; rows.push(x);
    });
  }
  if (rule.kind === 'ISSUER_CARD' && rule.issuer === '롯데카드' && typeof filterLotteCardByOrderDate_v666_ === 'function')
    rows = filterLotteCardByOrderDate_v666_(rows, date, master || []);
  else if (rule.kind === 'ISSUER_CARD' && typeof filterKnownSingleCard_v666_ === 'function')
    rows = filterKnownSingleCard_v666_(rows, rule.issuer, master || []);
  return rows;
}

function uniqueVatPeriodCardIdentities_v669_(rows, master) {
  var map = {};
  (rows || []).forEach(function(row) {
    var x = vatPeriodCandidateIdentity_v669_(row, master || []); if (!x) return;
    if (!map[x.key]) map[x.key] = x;
    else {
      var y = map[x.key];
      if (!y.alias && x.alias) y.alias = x.alias;
      if (!y.cardName && x.cardName) y.cardName = x.cardName;
      if (!y.cardNumber && x.cardNumber) y.cardNumber = x.cardNumber;
      if (!y.cardEnd4 && x.cardEnd4) y.cardEnd4 = x.cardEnd4;
      if (!y.sourceFile && x.sourceFile) y.sourceFile = x.sourceFile;
    }
  });
  return Object.keys(map).sort().map(function(k) { return map[k]; });
}

function vatPeriodFallbackReason_v669_(order, rule) {
  var p = typeof trackingPaymentReasonPrefix_v666_ === 'function' ? trackingPaymentReasonPrefix_v666_(order, rule) : '';
  return (p ? p + ' / ' : '') + '트래킹번호_기간단일카드_2차귀속_금액비교없음';
}

function vatPeriodBaseMatch_v669_(order, rule, status, identity) {
  identity = identity || {};
  return {
    status:status, reason:vatPeriodFallbackReason_v669_(order, rule), candidateCount:1,
    company:identity.company || '', alias:identity.alias || '', cardName:identity.cardName || '',
    cardNumber:identity.cardNumber || '', cardEnd4:identity.cardEnd4 || '',
    approvalDate:'', approvalTime:'', approvalNo:'', approvalAmount:0,
    merchant:'', merchantOrderNo:'', evidenceType:status === 'NON_CARD' ? '트래킹번호_기간결제수단귀속' : '기간단일카드_트래킹귀속',
    sourceFile:identity.sourceFile || '', cancelMemo:'', candidateSummary:identity.label || identity.key || '',
    trackingPayment:rule.raw || '', trackingPaymentRule:rule.kind || 'UNKNOWN', v669Fallback:true
  };
}

function ambiguousVatPeriodCards_v669_(order, rule, identities) {
  var p = typeof trackingPaymentReasonPrefix_v666_ === 'function' ? trackingPaymentReasonPrefix_v666_(order, rule) : '';
  return {
    status:'AMBIGUOUS', reason:(p ? p + ' / ' : '') + '기간내_서로다른구매카드후보_' + identities.length + '건',
    candidateCount:identities.length, company:'', alias:'', cardName:'', cardNumber:'', cardEnd4:'',
    approvalDate:'', approvalTime:'', approvalNo:'', approvalAmount:0, merchant:'', merchantOrderNo:'',
    evidenceType:'기간단일카드_트래킹귀속', sourceFile:'', cancelMemo:'',
    candidateSummary:identities.map(function(x) { return x.label || x.key; }).join(' || '),
    trackingPayment:rule.raw || '', trackingPaymentRule:rule.kind || 'UNKNOWN', v669Fallback:true
  };
}

var __baseMatchVatOrderCardCanonical_v669_ = typeof matchVatOrderCardCanonical_v664_ === 'function' ? matchVatOrderCardCanonical_v664_ : null;
if (__baseMatchVatOrderCardCanonical_v669_) {
  matchVatOrderCardCanonical_v664_ = function(order, history, master, used) {
    var result = __baseMatchVatOrderCardCanonical_v669_.apply(this, arguments) || {};
    if (result.status !== 'NO_MATCH') return result;
    var rule = typeof classifyVatTrackingPayment_v666_ === 'function'
      ? classifyVatTrackingPayment_v666_(order && order.lottePayment) : { kind:'UNKNOWN', raw:'', issuer:'' };
    if (rule.kind === 'UNKNOWN') return result;
    if (rule.kind === 'KAKAO_MONEY') return vatPeriodBaseMatch_v669_(order, rule, 'NON_CARD', {
      company:'비카드', alias:'신한은행 계좌결제', cardName:'카카오페이 페이머니',
      label:'카카오페이 페이머니 / 신한은행 계좌결제'
    });
    var identities = uniqueVatPeriodCardIdentities_v669_(
      vatPeriodTrackingCandidates_v669_(order, history || [], master || [], rule), master || []);
    if (identities.length === 1) return vatPeriodBaseMatch_v669_(order, rule, 'MATCHED', identities[0]);
    if (identities.length > 1) return ambiguousVatPeriodCards_v669_(order, rule, identities);
    result.reason = String(result.reason || '') + ' / 트래킹번호_기간단일카드_2차귀속후보없음';
    result.trackingPayment = rule.raw || ''; result.trackingPaymentRule = rule.kind || 'UNKNOWN';
    return result;
  };
}
