/** v6.70 Issue #32: date-window tracking fallback for generic L.PAY and KakaoPay card. */
var LOTTEON_PATCH_V670_VAT_DATE_WINDOW_FALLBACK_LOADED = true;

function vatDateWindowRule_v670_(order, result) {
  var date = String(order && order.orderDate || '');
  var raw = typeof text_v660_ === 'function' ? text_v660_(order && order.lottePayment) : String(order && order.lottePayment || '').trim();
  var compact = typeof compact_v660_ === 'function'
    ? compact_v660_(raw)
    : raw.toLowerCase().replace(/[\s._()\[\]{}\-\/]/g, '');

  if (result && result.status === 'NO_MATCH' && compact === 'lpay' &&
      date >= '2026-06-22' && date <= '2026-06-23') {
    return {
      end4:'4091', company:'KB국민카드', cardName:'HERITAGE Smart(할인형)',
      code:'LPAY_20260622_20260623_KB4091'
    };
  }

  if (result && (result.status === 'AMBIGUOUS' || result.status === 'NO_MATCH') && compact === '카카오페이') {
    if (date >= '2026-06-11' && date <= '2026-06-22') {
      return {
        end4:'0036', company:'롯데카드', cardName:'LOCA LIKIT 1.2',
        code:'KAKAOPAY_20260611_20260622_LOCA0036'
      };
    }
    if (date >= '2026-06-23' && date <= '2026-06-25') {
      return {
        end4:'7680', company:'우리카드', cardName:'카드의정석 EVERY POINT',
        code:'KAKAOPAY_20260623_20260625_WOORI7680'
      };
    }
    if (date >= '2026-06-29' && date <= '2026-06-30') {
      return {
        end4:'0036', company:'롯데카드', cardName:'LOCA LIKIT 1.2',
        code:'KAKAOPAY_20260629_20260630_LOCA0036'
      };
    }
  }
  return null;
}

function vatExpectedIdentity_v670_(order, history, master, rule) {
  if (!rule || !rule.end4) return null;
  var date = String(order && order.orderDate || '');
  var rows = [];

  (history || []).forEach(function(h) {
    if (!h || h.nonCard || h.cancelRow || h.v664FullyCanceled) return;
    if (typeof vatDateInSameHalf_v669_ === 'function' && !vatDateInSameHalf_v669_(h.date, date)) return;
    rows.push(h);
  });

  (master || []).forEach(function(m) {
    if (typeof vatMasterActiveOnDate_v669_ === 'function' && !vatMasterActiveOnDate_v669_(m, date)) return;
    var x = {};
    Object.keys(m || {}).forEach(function(k) { x[k] = m[k]; });
    x.v669Source = 'master';
    rows.push(x);
  });

  var map = {};
  rows.forEach(function(row) {
    var id = typeof vatPeriodCandidateIdentity_v669_ === 'function'
      ? vatPeriodCandidateIdentity_v669_(row, master || [])
      : null;
    if (!id || id.cardEnd4 !== rule.end4) return;
    if (rule.company && typeof normalizeCardCompany_v660_ === 'function' &&
        normalizeCardCompany_v660_(id.company) !== normalizeCardCompany_v660_(rule.company)) return;
    if (!map[id.key]) map[id.key] = id;
  });
  var ids = Object.keys(map).sort().map(function(k) { return map[k]; });
  return ids.length === 1 ? ids[0] : null;
}

function vatDateWindowMatch_v670_(order, rule, identity) {
  identity = identity || {};
  var raw = typeof text_v660_ === 'function' ? text_v660_(order && order.lottePayment) : String(order && order.lottePayment || '').trim();
  return {
    status:'MATCHED',
    reason:'트래킹번호_일자구간단일카드_3차귀속_' + rule.code + '_금액비교없음',
    candidateCount:1,
    company:identity.company || rule.company || '',
    alias:identity.alias || '',
    cardName:identity.cardName || rule.cardName || '',
    cardNumber:identity.cardNumber || '',
    cardEnd4:identity.cardEnd4 || rule.end4 || '',
    approvalDate:'',
    approvalTime:'',
    approvalNo:'',
    approvalAmount:0,
    merchant:'',
    merchantOrderNo:'',
    evidenceType:'일자구간단일카드_트래킹귀속',
    sourceFile:identity.sourceFile || '',
    cancelMemo:'',
    candidateSummary:identity.label || [rule.company, rule.cardName, rule.end4].filter(Boolean).join(' / '),
    trackingPayment:raw,
    trackingPaymentRule:'V670_DATE_WINDOW',
    v669Fallback:false,
    v670Fallback:true
  };
}

var __baseMatchVatOrderCardCanonical_v670_ = typeof matchVatOrderCardCanonical_v664_ === 'function'
  ? matchVatOrderCardCanonical_v664_
  : null;

if (__baseMatchVatOrderCardCanonical_v670_) {
  matchVatOrderCardCanonical_v664_ = function(order, history, master, used) {
    var result = __baseMatchVatOrderCardCanonical_v670_.apply(this, arguments) || {};
    if (result.status === 'MATCHED' || result.status === 'MASTER_MATCHED' || result.status === 'NON_CARD') return result;

    var rule = vatDateWindowRule_v670_(order, result);
    if (!rule) return result;

    var identity = vatExpectedIdentity_v670_(order, history || [], master || [], rule);
    if (!identity) {
      result.reason = String(result.reason || '') + ' / v6.70_일자구간단일카드_후보확정실패_' + rule.code;
      return result;
    }
    return vatDateWindowMatch_v670_(order, rule, identity);
  };
}
