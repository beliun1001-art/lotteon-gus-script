/** v6.66 Issue #24: tracking-number payment method as primary purchase-card evidence. */
var LOTTEON_PATCH_V666_VAT_TRACKING_PAYMENT_PRIMARY_LOADED = true;
var LOTTEON_V666_LOTTE_TRIP_LAST_ORDER_DATE = '2026-05-28';
var LOTTEON_V666_LOTTE_LIKIT_FIRST_ORDER_DATE = '2026-05-29';

function findVatTrackingPaymentHeader_v666_(headers) {
  return findHeaderAlias_v660_(headers || [], ['트래킹 번호','트래킹번호','tracking number','trackingnumber'], -1);
}

function findVatFallbackPaymentHeader_v666_(headers) {
  return findHeaderAlias_v660_(headers || [], ['결제수단','결제정보','결제방법','카드사','결제수단/카드사','결제수단(카드사)','구매결제수단'], -1);
}

var __baseFindVatPaymentHeader_v666_ = typeof findVatPaymentHeader_v660_ === 'function' ? findVatPaymentHeader_v660_ : null;
findVatPaymentHeader_v660_ = function(headers) {
  var tracking = findVatTrackingPaymentHeader_v666_(headers || []);
  if (tracking >= 0) return tracking;
  if (__baseFindVatPaymentHeader_v666_) return __baseFindVatPaymentHeader_v666_.apply(this, arguments);
  return findVatFallbackPaymentHeader_v666_(headers || []);
};

var __baseVatHeaderIndexes_v666_ = typeof vatHeaderIndexes_v648_ === 'function' ? vatHeaderIndexes_v648_ : null;
if (__baseVatHeaderIndexes_v666_) {
  vatHeaderIndexes_v648_ = function(headers) {
    var indexes = __baseVatHeaderIndexes_v666_.apply(this, arguments) || {};
    indexes.v666TrackingPayment = findVatTrackingPaymentHeader_v666_(headers || []);
    indexes.v666FallbackPayment = findVatFallbackPaymentHeader_v666_(headers || []);
    return indexes;
  };
}

var __baseVatDetailRow_v666_ = typeof vatDetailRow_v648_ === 'function' ? vatDetailRow_v648_ : null;
if (__baseVatDetailRow_v666_) {
  vatDetailRow_v648_ = function(row, ix, sourceRow) {
    var result = __baseVatDetailRow_v666_.apply(this, arguments);
    if (!result || !result.row || !result.row.length) return result;
    var tracking = ix && ix.v666TrackingPayment >= 0 ? cleanVatText_v648_(valueAt_v648_(row, ix.v666TrackingPayment)) : '';
    var fallback = ix && ix.v666FallbackPayment >= 0 ? cleanVatText_v648_(valueAt_v648_(row, ix.v666FallbackPayment)) : '';
    // v6.60 appends the normalized purchase-payment value as the last VAT detail column.
    result.row[result.row.length - 1] = tracking || fallback || '';
    return result;
  };
}

function normalizeTrackingCardIssuer_v666_(value) {
  var s = compact_v660_(value);
  if (!s) return '';
  if (s.indexOf('kb') >= 0 || s.indexOf('국민') >= 0) return 'KB국민카드';
  if (s.indexOf('롯데') >= 0) return '롯데카드';
  if (s.indexOf('우리') >= 0) return '우리카드';
  if (s.indexOf('신한') >= 0) return '신한카드';
  if (s.indexOf('농협') >= 0 || s.indexOf('nh') >= 0) return 'NH농협카드';
  if (s.indexOf('삼성') >= 0) return '삼성카드';
  if (s.indexOf('하나') >= 0) return '하나카드';
  if (s.indexOf('현대') >= 0) return '현대카드';
  return '';
}

function classifyVatTrackingPayment_v666_(value) {
  var raw = text_v660_(value);
  var s = compact_v660_(raw);
  if (!s) return { kind:'UNKNOWN', raw:raw, issuer:'' };
  if (s.indexOf('카카오') >= 0 && (s.indexOf('머니') >= 0 || s.indexOf('계좌') >= 0 || s.indexOf('현금') >= 0)) {
    return { kind:'KAKAO_MONEY', raw:raw, issuer:'비카드' };
  }
  if (s.indexOf('카카오') >= 0) return { kind:'KAKAO_CARD', raw:raw, issuer:'' };
  var issuer = normalizeTrackingCardIssuer_v666_(raw);
  if (issuer) return { kind:'ISSUER_CARD', raw:raw, issuer:issuer };
  return { kind:'UNKNOWN', raw:raw, issuer:'' };
}

function isKakaoCardEvidence_v666_(h) {
  if (!h || h.nonCard) return false;
  var s = compact_v660_((h.cardName || '') + ' ' + (h.evidenceType || '') + ' ' + (h.sourceFile || '') + ' ' + (h.memo || ''));
  return s.indexOf('카카오') >= 0 && s.indexOf('머니') < 0;
}

function isKakaoMoneyEvidence_v666_(h) {
  if (!h || !h.nonCard) return false;
  var s = compact_v660_((h.company || '') + ' ' + (h.cardName || '') + ' ' + (h.evidenceType || '') + ' ' + (h.sourceFile || '') + ' ' + (h.memo || ''));
  return s.indexOf('카카오') >= 0 && s.indexOf('머니') >= 0;
}

var __baseCanonicalizeVatHistory_v666_ = typeof canonicalizeVatHistory_v664_ === 'function' ? canonicalizeVatHistory_v664_ : null;
if (__baseCanonicalizeVatHistory_v666_) {
  canonicalizeVatHistory_v664_ = function(history, master) {
    var out = __baseCanonicalizeVatHistory_v666_.apply(this, arguments) || [];
    var tags = {};
    (history || []).forEach(function(h) {
      if (!h || h.nonCard) return;
      var issuer = normalizeCardCompany_v660_(h.company);
      var approval = text_v660_(h.approvalNo);
      if (!issuer || !approval) return;
      var key = 'CARD|' + issuer + '|' + approval;
      if (!tags[key]) tags[key] = { kakaoCard:false };
      if (isKakaoCardEvidence_v666_(h)) tags[key].kakaoCard = true;
    });
    out.forEach(function(h) {
      var tag = tags[String(h && h.v664CanonicalKey || '')];
      h.v666KakaoCard = !!(tag && tag.kakaoCard) || isKakaoCardEvidence_v666_(h);
      h.v666KakaoMoney = isKakaoMoneyEvidence_v666_(h);
    });
    return out;
  };
}

function lotteCardKind_v666_(h, master) {
  if (!h) return '';
  var e = typeof enrichHistoryFromMaster_v660_ === 'function' ? enrichHistoryFromMaster_v660_(h, master || []) : h;
  var s = compact_v660_((e.cardName || h.cardName || '') + ' ' + (e.cardNumber || h.cardNumber || '') + ' ' + (e.cardEnd4 || h.cardEnd4 || ''));
  if (s.indexOf('tripto로카') >= 0 || s.indexOf('트립투로카') >= 0) return 'TRIP';
  if (s.indexOf('localikit') >= 0 || s.indexOf('로카likit') >= 0 || s.indexOf('로카리킷') >= 0) return 'LIKIT';
  return '';
}

function filterLotteCardByOrderDate_v666_(rows, orderDate, master) {
  var recognized = (rows || []).filter(function(h) { return !!lotteCardKind_v666_(h, master || []); });
  if (!recognized.length) return rows || [];
  var expected = '';
  if (orderDate && orderDate <= LOTTEON_V666_LOTTE_TRIP_LAST_ORDER_DATE) expected = 'TRIP';
  else if (orderDate && orderDate >= LOTTEON_V666_LOTTE_LIKIT_FIRST_ORDER_DATE) expected = 'LIKIT';
  if (!expected) return rows || [];
  return recognized.filter(function(h) { return lotteCardKind_v666_(h, master || []) === expected; });
}

function filterVatHistoryByTrackingPayment_v666_(order, history, master, rule) {
  var rows = (history || []).slice();
  if (!rule || rule.kind === 'UNKNOWN') return rows;
  if (rule.kind === 'KAKAO_MONEY') {
    return rows.filter(function(h) { return !!h.nonCard && (!!h.v666KakaoMoney || isKakaoMoneyEvidence_v666_(h)); });
  }
  if (rule.kind === 'KAKAO_CARD') {
    return rows.filter(function(h) { return !h.nonCard && !!h.v666KakaoCard; });
  }
  if (rule.kind === 'ISSUER_CARD') {
    rows = rows.filter(function(h) { return normalizeCardCompany_v660_(h.company) === rule.issuer; });
    if (rule.issuer === '롯데카드') rows = filterLotteCardByOrderDate_v666_(rows, order && order.orderDate, master || []);
    return rows;
  }
  return rows;
}

function trackingPaymentReasonPrefix_v666_(order, rule) {
  if (!rule || rule.kind === 'UNKNOWN') return '';
  if (rule.kind === 'KAKAO_MONEY') return '트래킹번호_카카오페이페이머니_신한은행계좌_현금결제_1차필터';
  if (rule.kind === 'KAKAO_CARD') return '트래킹번호_카카오페이카드_원카드승인증빙_1차필터';
  var prefix = '트래킹번호_' + rule.issuer + '_1차필터';
  if (rule.issuer === 'KB국민카드') prefix += '_HERITAGE단일카드';
  if (rule.issuer === '우리카드') prefix += '_EVERY_POINT단일카드';
  if (rule.issuer === '롯데카드') {
    if (order && order.orderDate <= LOTTEON_V666_LOTTE_TRIP_LAST_ORDER_DATE) prefix += '_Trip_to_로카(~2026-05-28)';
    else if (order && order.orderDate >= LOTTEON_V666_LOTTE_LIKIT_FIRST_ORDER_DATE) prefix += '_LOCA_LIKIT(2026-05-29~)';
  }
  return prefix;
}

var __baseMatchVatOrderCardCanonical_v666_ = typeof matchVatOrderCardCanonical_v664_ === 'function' ? matchVatOrderCardCanonical_v664_ : null;
if (__baseMatchVatOrderCardCanonical_v666_) {
  matchVatOrderCardCanonical_v664_ = function(order, history, master, used) {
    var rule = classifyVatTrackingPayment_v666_(order && order.lottePayment);
    var filtered = filterVatHistoryByTrackingPayment_v666_(order, history || [], master || [], rule);
    var result = __baseMatchVatOrderCardCanonical_v666_.call(this, order, filtered, master, used);
    result = result || {};
    result.trackingPayment = rule.raw || '';
    result.trackingPaymentRule = rule.kind || 'UNKNOWN';
    var prefix = trackingPaymentReasonPrefix_v666_(order, rule);
    if (prefix) result.reason = prefix + (result.reason ? ' / ' + result.reason : '');
    if (rule.kind === 'KAKAO_MONEY' && result.status === 'NON_CARD') {
      result.company = '비카드';
      result.alias = '신한은행 계좌결제';
      result.cardName = '카카오페이 페이머니';
    }
    return result;
  };
}
