/**
 * v6.61 Issue #20 - Google Sheets Date-object preservation for purchase-card reconciliation.
 *
 * Root cause confirmed by operating smoke:
 * - 카드사용내역_붙여넣기 dates pasted from Excel are often real Google Sheets Date objects.
 * - v6.60 loadVatCardHistory_v660_ converted the Date to text before normalizeDateText_v660_.
 * - The resulting locale Date string was not YYYY-MM-DD, so date+amount evidence could not match.
 * - LOTTE-card rows that remained text dates matched, while KB/우리 evidence was largely NO_MATCH.
 *
 * This patch changes only input loading. Matching rules, VAT math, AC purchase source and
 * pre-July no-force-match policy remain unchanged.
 */
var LOTTEON_PATCH_V661_VAT_CARD_DATE_OBJECT_FIX_LOADED = true;

loadVatCardHistory_v660_ = function(ss) {
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
  for (var r = 1; r < values.length; r++) {
    var row = values[r], amount = number_v660_(p.amount >= 0 ? row[p.amount] : 0);
    var rawDate = p.date >= 0 ? row[p.date] : '';
    var rawCancelDate = p.cancelDate >= 0 ? row[p.cancelDate] : '';
    var obj = {
      rowNo:r + 1,
      company:textAt_v660_(row,p.company), cardName:textAt_v660_(row,p.name), cardNumber:textAt_v660_(row,p.number), cardEnd4:textAt_v660_(row,p.end4),
      date:normalizeDateText_v660_(rawDate), time:textAt_v660_(row,p.time), merchant:textAt_v660_(row,p.merchant), merchantBusiness:textAt_v660_(row,p.merchantBusiness),
      amount:amount, approvalNo:textAt_v660_(row,p.approval), status:textAt_v660_(row,p.status), cancelDate:normalizeDateText_v660_(rawCancelDate),
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
};

loadVatCardMaster_v660_ = function(ss) {
  var sheet = ss && ss.getSheetByName && ss.getSheetByName(LOTTEON_V660_CARD_MASTER_SHEET);
  if (!sheet || sheet.getLastRow() < 2) return [];
  var values = sheet.getDataRange().getValues(), h = values[0] || [];
  var ix = function(names) { return findHeaderAlias_v660_(h, names, -1); };
  var p = {
    company:ix(['카드사']), alias:ix(['카드별칭']), name:ix(['카드명']), type:ix(['카드구분']), status:ix(['상태']),
    number:ix(['카드번호']), end4:ix(['카드번호끝4']), account:ix(['사업자코드','쿠팡계정ID']), business:ix(['사업자등록번호']),
    start:ix(['적용시작일']), end:ix(['적용종료일']), limit:ix(['한도']), memo:ix(['메모'])
  };
  var out = [];
  for (var r = 1; r < values.length; r++) {
    var row = values[r];
    var rawStart = p.start >= 0 ? row[p.start] : '';
    var rawEnd = p.end >= 0 ? row[p.end] : '';
    var obj = {
      rowNo:r + 1, company:textAt_v660_(row,p.company), alias:textAt_v660_(row,p.alias), cardName:textAt_v660_(row,p.name),
      cardType:textAt_v660_(row,p.type), status:textAt_v660_(row,p.status), cardNumber:textAt_v660_(row,p.number), cardEnd4:textAt_v660_(row,p.end4),
      account:textAt_v660_(row,p.account), business:textAt_v660_(row,p.business), startDate:normalizeDateText_v660_(rawStart), endDate:normalizeDateText_v660_(rawEnd),
      limit:number_v660_(p.limit >= 0 ? row[p.limit] : 0), memo:textAt_v660_(row,p.memo)
    };
    if (obj.company || obj.cardName || obj.cardNumber || obj.business || obj.account) out.push(obj);
  }
  return out;
};
