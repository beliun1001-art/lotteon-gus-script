/** v6.68 Issue #24: normalize every KakaoPay Money non-card match to one summary identity. */
var LOTTEON_PATCH_V668_VAT_KAKAO_MONEY_SUMMARY_NORMALIZE_LOADED = true;

function isVatKakaoMoneyMatch_v668_(result) {
  if (!result || result.status !== 'NON_CARD') return false;
  var s = compact_v660_([
    result.company || '', result.alias || '', result.cardName || '',
    result.evidenceType || '', result.candidateSummary || '', result.reason || ''
  ].join(' '));
  return s.indexOf('카카오') >= 0 && s.indexOf('머니') >= 0;
}

function normalizeVatKakaoMoneyMatch_v668_(result) {
  if (!isVatKakaoMoneyMatch_v668_(result)) return result;
  result.company = '비카드';
  result.alias = '신한은행 계좌결제';
  result.cardName = '카카오페이 페이머니';
  result.cardNumber = '';
  result.cardEnd4 = '';
  return result;
}

var __baseMatchVatOrderCardCanonical_v668_ = typeof matchVatOrderCardCanonical_v664_ === 'function'
  ? matchVatOrderCardCanonical_v664_ : null;
if (__baseMatchVatOrderCardCanonical_v668_) {
  matchVatOrderCardCanonical_v664_ = function(order, history, master, used) {
    var result = __baseMatchVatOrderCardCanonical_v668_.apply(this, arguments);
    return normalizeVatKakaoMoneyMatch_v668_(result || {});
  };
}
