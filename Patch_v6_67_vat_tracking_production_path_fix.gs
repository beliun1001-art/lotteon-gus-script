/**
 * v6.67 Issue #24: production VAT continuation path for tracking payment evidence.
 *
 * - Discover payment headers across the full source header row.
 * - Read only through the furthest required payment column while keeping AC (index 28)
 *   as the immutable purchase source of truth.
 * - Preserve tracking-first / fallback-second behavior in the normal 500-row continuation.
 * - Restore four-digit card endings and make the Lotte 2026-05-28/29 boundary depend on
 *   the physical card number/end4 before any card-name label.
 */
var LOTTEON_PATCH_V667_VAT_TRACKING_PRODUCTION_PATH_FIX_LOADED = true;

/* ---------- Source payment header discovery ---------- */
findVatTrackingPaymentHeader_v666_ = function(headers) {
  return findHeaderAlias_v660_(headers || [], [
    '현지트래킹번호','현지 트래킹 번호','트래킹 번호','트래킹번호','tracking number','trackingnumber'
  ], -1);
};

findVatFallbackPaymentHeader_v666_ = function(headers) {
  headers = headers || [];
  var direct = findHeaderAlias_v660_(headers, [
    '롯데결제수단','롯데 결제수단','롯데결제정보','롯데 결제정보',
    '결제수단','결제정보','결제방법','카드사','결제카드','결제카드사',
    '결제수단/카드사','결제수단(카드사)','구매결제수단'
  ], -1);
  if (direct >= 0) return direct;

  for (var i = 0; i < headers.length; i++) {
    var h = compact_v660_(headers[i]);
    if (!h || h.indexOf('트래킹') >= 0 || h.indexOf('tracking') >= 0) continue;
    if (h.indexOf('결제수단') >= 0 || h.indexOf('결제방법') >= 0 ||
        h.indexOf('결제카드') >= 0 || h.indexOf('결제정보') >= 0) return i;
  }
  return -1;
};

/* ---------- Normal VAT 500-row continuation: widen source read only as needed ---------- */
var __baseRunVatDetailBatch_v667_ = typeof runVatDetailBatch_v648_ === 'function' ? runVatDetailBatch_v648_ : null;
if (__baseRunVatDetailBatch_v667_) {
  runVatDetailBatch_v648_ = function(ss, state) {
    var source = ss.getSheetByName(LOTTEON_V648_SOURCE_SHEET);
    if (!source) throw new Error(LOTTEON_V648_SOURCE_SHEET + ' 시트를 찾을 수 없습니다.');
    var lastRow = Math.max(Number(state.sourceLastRow || 0), source.getLastRow());
    var startRow = Math.max(2, Number(state.sourceRow || 2));

    if (startRow > lastRow) {
      state.phase = 'summaries';
      state.updatedAt = new Date().toISOString();
      saveVatState_v648_(state);
      scheduleVatTrigger_v648_(state);
      writeVatStatus_v648_(ss, state, '상세 완료; 요약 생성 예약');
      return { ok: true, done: false, state: state };
    }

    var sourceLastColumn = Math.max(1, source.getLastColumn());
    var allHeaders = source.getRange(1, 1, 1, sourceLastColumn).getValues()[0] || [];
    var allIndexes = vatHeaderIndexes_v648_(allHeaders);
    var requiredLastColumn = Math.max(7, Number(LOTTEON_V648_MAX_COL || 29));
    ['lottePayment','v666TrackingPayment','v666FallbackPayment'].forEach(function(key) {
      var index = Number(allIndexes && allIndexes[key]);
      if (isFinite(index) && index >= 0) requiredLastColumn = Math.max(requiredLastColumn, index + 1);
    });
    var maxCol = Math.min(sourceLastColumn, requiredLastColumn);
    var count = Math.min(LOTTEON_V648_CHUNK_SIZE, lastRow - startRow + 1);
    var headers = allHeaders.slice(0, maxCol);
    var values = source.getRange(startRow, 1, count, maxCol).getValues();
    var indexes = vatHeaderIndexes_v648_(headers);
    var output = [];
    var skipped = 0;
    var missing = 0;

    values.forEach(function(row, offset) {
      var result = vatDetailRow_v648_(row, indexes, startRow + offset);
      if (!result.row) {
        skipped += 1;
        if (result.accountMissing) missing += 1;
        return;
      }
      if (result.accountMissing) missing += 1;
      output.push(result.row);
    });

    if (output.length) {
      var detail = ss.getSheetByName(LOTTEON_V648_DETAIL_SHEET);
      detail.getRange(2 + Number(state.writtenRows || 0), 1, output.length, output[0].length).setValues(output);
    }
    state.sourceRow = startRow + count;
    state.writtenRows = Number(state.writtenRows || 0) + output.length;
    state.skippedRows = Number(state.skippedRows || 0) + skipped;
    state.accountMissingRows = Number(state.accountMissingRows || 0) + missing;
    state.sourceReadColumns = maxCol;
    state.trackingPaymentColumn = Number(indexes.v666TrackingPayment) >= 0 ? Number(indexes.v666TrackingPayment) + 1 : 0;
    state.fallbackPaymentColumn = Number(indexes.v666FallbackPayment) >= 0 ? Number(indexes.v666FallbackPayment) + 1 : 0;
    state.updatedAt = new Date().toISOString();

    if (state.sourceRow > lastRow) state.phase = 'summaries';
    saveVatState_v648_(state);
    scheduleVatTrigger_v648_(state);
    writeVatStatus_v648_(ss, state, state.phase === 'summaries' ? '상세 완료; 요약 생성 예약' : '다음 상세 배치 예약');
    return { ok: true, done: false, state: state };
  };
}

/* ---------- Card identifier text safety ---------- */
function normalizeVatCardEnd4_v667_(end4, cardNumber) {
  var explicitDigits = digits_v660_(end4);
  if (explicitDigits) return ('0000' + explicitDigits).slice(-4);
  var numberDigits = digits_v660_(cardNumber);
  return numberDigits.length >= 4 ? numberDigits.slice(-4) : '';
}

function normalizeVatCardEvidenceIdentity_v667_(item) {
  if (!item) return item;
  item.cardNumber = text_v660_(item.cardNumber);
  item.cardEnd4 = normalizeVatCardEnd4_v667_(item.cardEnd4, item.cardNumber);
  return item;
}

var __baseLoadVatCardHistory_v667_ = typeof loadVatCardHistory_v660_ === 'function' ? loadVatCardHistory_v660_ : null;
if (__baseLoadVatCardHistory_v667_) {
  loadVatCardHistory_v660_ = function(ss) {
    var rows = __baseLoadVatCardHistory_v667_.apply(this, arguments) || [];
    rows.forEach(normalizeVatCardEvidenceIdentity_v667_);
    return rows;
  };
}

var __baseLoadVatCardMaster_v667_ = typeof loadVatCardMaster_v660_ === 'function' ? loadVatCardMaster_v660_ : null;
if (__baseLoadVatCardMaster_v667_) {
  loadVatCardMaster_v660_ = function(ss) {
    var rows = __baseLoadVatCardMaster_v667_.apply(this, arguments) || [];
    rows.forEach(normalizeVatCardEvidenceIdentity_v667_);
    return rows;
  };
}

/* ---------- Lotte physical-card boundary ---------- */
lotteCardKind_v666_ = function(h, master) {
  if (!h) return '';
  var enriched = typeof enrichHistoryFromMaster_v660_ === 'function'
    ? enrichHistoryFromMaster_v660_(h, master || []) : h;
  var cardNumber = text_v660_(enriched.cardNumber || h.cardNumber || '');
  var cardEnd4 = normalizeVatCardEnd4_v667_(enriched.cardEnd4 || h.cardEnd4 || '', cardNumber);

  // Physical identifiers take precedence over potentially stale/mislabeled card names.
  if (cardEnd4 === '0126') return 'TRIP';
  if (cardEnd4 === '0036') return 'LIKIT';

  var numberDigits = digits_v660_(cardNumber);
  if (numberDigits.slice(-4) === '0126') return 'TRIP';
  if (numberDigits.slice(-4) === '0036') return 'LIKIT';

  var name = compact_v660_(enriched.cardName || h.cardName || '');
  if (name.indexOf('tripto로카') >= 0 || name.indexOf('트립투로카') >= 0) return 'TRIP';
  if (name.indexOf('localikit') >= 0 || name.indexOf('로카likit') >= 0 || name.indexOf('로카리킷') >= 0) return 'LIKIT';
  return '';
};

function normalizeVatCardMatchIdentity_v667_(result, master) {
  if (!result) return result;
  result.cardNumber = text_v660_(result.cardNumber);
  result.cardEnd4 = normalizeVatCardEnd4_v667_(result.cardEnd4, result.cardNumber);
  if (normalizeCardCompany_v660_(result.company) === '롯데카드') {
    var kind = lotteCardKind_v666_(result, master || []);
    if (kind === 'TRIP') result.cardName = 'Trip to 로카';
    else if (kind === 'LIKIT') result.cardName = 'LOCA LIKIT 1.2';
  }
  return result;
}

var __baseMatchVatOrderCardCanonical_v667_ = typeof matchVatOrderCardCanonical_v664_ === 'function' ? matchVatOrderCardCanonical_v664_ : null;
if (__baseMatchVatOrderCardCanonical_v667_) {
  matchVatOrderCardCanonical_v664_ = function(order, history, master, used) {
    var result = __baseMatchVatOrderCardCanonical_v667_.apply(this, arguments);
    return normalizeVatCardMatchIdentity_v667_(result, master || []);
  };
}
