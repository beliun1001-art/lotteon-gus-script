/** v6.62 Issue #20: safely collapse same-card ambiguity and diagnose unresolved purchase-card matches. */
var LOTTEON_PATCH_V662_VAT_CARD_UNMATCHED_DIAGNOSTIC_LOADED = true;
var LOTTEON_V662_UNMATCHED_SHEET = '부가세_카드미매칭분석';

var __baseMatchVatOrderCard_v662_ = typeof matchVatOrderCard_v660_ === 'function' ? matchVatOrderCard_v660_ : null;
if (__baseMatchVatOrderCard_v662_) {
  matchVatOrderCard_v660_ = function(order, history, master) {
    var result = __baseMatchVatOrderCard_v662_.apply(this, arguments);
    if (!result || result.status !== 'AMBIGUOUS') return result;
    var candidates = ambiguityCandidatesForOrder_v662_(order, history);
    var resolved = resolveSamePhysicalCardCandidates_v662_(order, candidates, master);
    return resolved || result;
  };
}

function ambiguityCandidatesForOrder_v662_(order, history) {
  var direct = [];
  if (order && order.orderNo) {
    direct = (history || []).filter(function(h) {
      return !h.cancelRow && h.merchantOrderNo && text_v660_(h.merchantOrderNo) === text_v660_(order.orderNo);
    });
    direct = dedupeHistoryCandidates_v660_(direct);
    if (direct.length > 1) return direct;
  }
  var base = (history || []).filter(function(h) {
    if (h.cancelRow || !h.lotteEvidence) return false;
    if (!order.orderDate || h.date !== order.orderDate) return false;
    return historyMatchesAmount_v660_(h, order.purchase);
  });
  var paymentFiltered = filterEvidenceByLottePayment_v660_(base, order.lottePayment);
  var candidates = paymentFiltered.length ? paymentFiltered : base;
  return dedupeHistoryCandidates_v660_(candidates);
}

function resolveSamePhysicalCardCandidates_v662_(order, candidates, master) {
  if (!candidates || candidates.length < 2) return null;
  var identities = {}, first = null, allNonCard = true;
  for (var i = 0; i < candidates.length; i++) {
    var h = candidates[i];
    var key = cardIdentityKey_v662_(h, master);
    if (!key) return null;
    identities[key] = true;
    if (!first) first = h;
    if (!h.nonCard) allNonCard = false;
  }
  if (Object.keys(identities).length !== 1) return null;
  var status = allNonCard ? 'NON_CARD' : 'MATCHED';
  return matchFromHistory_v660_(
    order,
    first,
    master,
    status,
    '거래내역_다중후보_동일구매카드확정(' + candidates.length + '건)'
  );
}

function cardIdentityKey_v662_(h, master) {
  if (!h) return '';
  if (h.nonCard) {
    var nonCardName = compact_v660_((h.company || '') + '|' + (h.cardName || '') + '|' + (h.evidenceType || ''));
    return nonCardName ? 'NONCARD|' + nonCardName : '';
  }
  var e = enrichHistoryFromMaster_v660_(h, master || []);
  var company = normalizeCardCompany_v660_(e.company || h.company);
  if (!company) return '';
  var end4 = digits_v660_(e.cardEnd4 || h.cardEnd4);
  if (end4) return company + '|END4|' + end4;
  var number = compact_v660_(e.cardNumber || h.cardNumber);
  if (number) return company + '|NUMBER|' + number;
  var name = normalizeCardName_v660_(e.cardName || h.cardName);
  if (name && name.length >= 4) return company + '|NAME|' + name;
  return '';
}

var __baseBuildVatPurchaseCardReconciliation_v662_ = typeof buildVatPurchaseCardReconciliation_v660_ === 'function' ? buildVatPurchaseCardReconciliation_v660_ : null;
if (__baseBuildVatPurchaseCardReconciliation_v662_) {
  buildVatPurchaseCardReconciliation_v660_ = function(ss) {
    ss = ss || SpreadsheetApp.getActive();
    var result = __baseBuildVatPurchaseCardReconciliation_v662_.apply(this, arguments);
    try {
      result = result || {};
      result.unmatchedAnalysisRows = writeVatCardUnmatchedAnalysis_v662_(ss);
    } catch (e) {
      result = result || {};
      result.unmatchedAnalysisError = String(e && e.message ? e.message : e);
    }
    return result;
  };
}

function writeVatCardUnmatchedAnalysis_v662_(ss) {
  var diag = ss && ss.getSheetByName && ss.getSheetByName(LOTTEON_V660_CARD_DIAG_SHEET);
  if (!diag || diag.getLastRow() < 2) return 0;
  var values = diag.getDataRange().getValues();
  var headers = values[0] || [];
  function ix(name) { return findHeaderAlias_v660_(headers, [name], -1); }
  var p = {
    year: ix('신고연도'), half: ix('반기'), date: ix('주문일'), business: ix('사업자등록번호'), account: ix('쿠팡계정ID'),
    orderNo: ix('주문번호'), payment: ix('롯데결제수단'), purchase: ix('주문매입금액'), status: ix('카드매칭상태'), reason: ix('카드매칭근거'), candidates: ix('후보요약')
  };
  var history = loadVatCardHistory_v660_(ss);
  var rows = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var status = textAt_v660_(row, p.status);
    if (status !== 'NO_MATCH' && status !== 'AMBIGUOUS') continue;
    var orderDate = normalizeDateText_v660_(p.date >= 0 ? row[p.date] : '');
    var purchase = number_v660_(p.purchase >= 0 ? row[p.purchase] : 0);
    var payment = textAt_v660_(row, p.payment);
    var insight = analyzeUnmatchedOrder_v662_(orderDate, purchase, payment, history, status);
    rows.push([
      textAt_v660_(row,p.year), textAt_v660_(row,p.half), orderDate, textAt_v660_(row,p.business), textAt_v660_(row,p.account), textAt_v660_(row,p.orderNo),
      payment, purchase, status, textAt_v660_(row,p.reason), insight.exactAll, insight.exactNonLotte, insight.prevDayExact, insight.nextDayExact,
      insight.closestDiff, insight.closestCandidates, insight.suggestion, textAt_v660_(row,p.candidates)
    ]);
  }
  var outHeaders = ['신고연도','반기','주문일','사업자등록번호','쿠팡계정ID','주문번호','롯데결제수단','주문매입금액','현재상태','현재근거','동일일자+금액전체후보','그중롯데계열아님','전일동일금액롯데후보','익일동일금액롯데후보','동일일자최근접금액차이','최근접후보','다음확인권고','기존후보요약'];
  var sheet = ss.getSheetByName(LOTTEON_V662_UNMATCHED_SHEET) || ss.insertSheet(LOTTEON_V662_UNMATCHED_SHEET);
  sheet.clearContents();
  sheet.getRange(1,1,1,outHeaders.length).setValues([outHeaders]);
  if (rows.length) sheet.getRange(2,1,rows.length,outHeaders.length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,outHeaders.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  if (rows.length) {
    sheet.getRange(2,8,rows.length,1).setNumberFormat('#,##0');
    sheet.getRange(2,11,rows.length,5).setNumberFormat('#,##0');
  }
  for (var c=0;c<outHeaders.length;c++) sheet.setColumnWidth(c+1, /후보|권고|근거/.test(outHeaders[c]) ? 220 : (/사업자|계정|주문번호|결제수단/.test(outHeaders[c]) ? 135 : 95));
  return rows.length;
}

function analyzeUnmatchedOrder_v662_(orderDate, purchase, payment, history, status) {
  var exactAll = [], exactLotte = [], exactNonLotte = [];
  var prev = [], next = [], sameDay = [];
  var prevDate = shiftDate_v662_(orderDate, -1), nextDate = shiftDate_v662_(orderDate, 1);
  var issuer = normalizeCardCompany_v660_(payment);
  (history || []).forEach(function(h) {
    if (h.cancelRow) return;
    var amountMatch = historyMatchesAmount_v660_(h, purchase);
    if (h.date === orderDate) {
      if (amountMatch) {
        exactAll.push(h);
        if (h.lotteEvidence) exactLotte.push(h); else exactNonLotte.push(h);
      }
      if (h.lotteEvidence && (!issuer || normalizeCardCompany_v660_(h.company) === issuer)) sameDay.push(h);
    }
    if (amountMatch && h.lotteEvidence && h.date === prevDate) prev.push(h);
    if (amountMatch && h.lotteEvidence && h.date === nextDate) next.push(h);
  });
  sameDay.sort(function(a,b){ return Math.abs(Number(a.amount||0)-purchase) - Math.abs(Number(b.amount||0)-purchase); });
  var closest = sameDay.slice(0,3);
  var closestDiff = closest.length ? Math.abs(Number(closest[0].amount||0)-purchase) : '';
  var suggestion = '';
  if (status === 'AMBIGUOUS') suggestion = '서로 다른 구매카드 후보면 수동확인; 동일카드 후보는 v6.62 자동확정';
  else if (exactNonLotte.length && !exactLotte.length) suggestion = '가맹점 롯데계열 분류 누락 확인';
  else if (prev.length || next.length) suggestion = '승인일 ±1일 후보 확인';
  else if (closest.length && closestDiff !== '' && closestDiff <= 1000) suggestion = '동일일자 금액차이/쿠폰·부분취소 확인';
  else suggestion = '원본 카드증빙 부재 또는 주문-승인금액 구조 확인';
  return {
    exactAll: exactAll.length,
    exactNonLotte: exactNonLotte.length,
    prevDayExact: prev.length,
    nextDayExact: next.length,
    closestDiff: closestDiff,
    closestCandidates: closest.map(historyCandidateLabel_v660_).join(' || '),
    suggestion: suggestion
  };
}

function shiftDate_v662_(dateText, days) {
  var m = String(dateText || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return '';
  var d = new Date(Date.UTC(Number(m[1]), Number(m[2])-1, Number(m[3])));
  d.setUTCDate(d.getUTCDate() + Number(days || 0));
  return d.getUTCFullYear() + '-' + pad2_v660_(d.getUTCMonth()+1) + '-' + pad2_v660_(d.getUTCDate());
}

if (typeof LOTTEON_V653_CORE_VISIBLE_SHEETS !== 'undefined' && LOTTEON_V653_CORE_VISIBLE_SHEETS.indexOf(LOTTEON_V662_UNMATCHED_SHEET) < 0) {
  LOTTEON_V653_CORE_VISIBLE_SHEETS.push(LOTTEON_V662_UNMATCHED_SHEET);
}
