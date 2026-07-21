/**
 * LOTTEON v6.53 VAT filtering and operation sheet cleanup
 *
 * 확정 기준:
 * - beliun1021-1 사업자등록번호: 176-71-00758
 * - 부가세_신고자료 / 부가세_상품별은 취소·반품·환불·교환 상태 제외 기준으로 생성
 * - 부가세 생성 대상은 매출데이터_붙여넣기 기준:
 *   1) 마켓주문상태가 취소/반품/환불/교환이 아님
 *   2) 구매사이트명에 lotteon 포함
 *   3) 사이트상품번호 또는 마켓상품번호 존재
 *   4) 브랜드 존재
 * - 고객별/주문번호별/수동전송수/TMP류 시트는 삭제
 * - 운영 핵심 11개 시트만 표시하고 진단/상태/보강/검증 시트는 숨김
 */

var LOTTEON_PATCH_V653_VAT_FILTER_AND_SHEET_CLEANUP_LOADED = true;
LOTTEON_V648_VERSION = 'v6.53';

var LOTTEON_V653_CORE_VISIBLE_SHEETS = [
  '대시보드',
  '매출데이터_붙여넣기',
  '필터별_상품수',
  '쿠팡재전송_로그',
  '브랜드별_마진율',
  '사업자별_계정별_써머리',
  '미정산_쿠팡계정별',
  '부가세_신고자료',
  '부가세_상품별',
  '부가세_기간별',
  '시트별_금액검증'
];

var LOTTEON_V653_DELETE_EXACT_SHEETS = [
  '부가세_고객별',
  '부가세_주문번호별',
  '고객주소_구분검증',
  '쿠팡전송수_수동입력',
  '수동입력_API검증',
  '전송일_수동입력',
  '빠른갱신_TMP',
  '필터별_상품수_TMP',
  'LOTTEON_상품목록_TMP',
  'LOTTEON_상품목록_샘플',
  '시트2',
  '시트3'
];

businessNoForMarketId_v648_ = function(marketId) {
  var normalized = cleanVatText_v648_(marketId).toLowerCase();
  if (normalized === 'beliun1021' || normalized === '1021') return '227-27-04928';
  if (normalized === 'beliun1021-1' || normalized === '1021-1') return '176-71-00758';
  if (normalized === 'beliun1023' || normalized === '1023') return '835-58-00765';
  if (normalized === 'beliun1024' || normalized === '1024') return '606-45-93763';
  return '';
};

vatHeaderIndexes_v648_ = function(headers) {
  var normalized = {};
  headers.forEach(function(header, index) { normalized[cleanVatText_v648_(header).replace(/\s/g, '')] = index; });
  function find(names, fallback) {
    for (var i = 0; i < names.length; i++) {
      var key = names[i].replace(/\s/g, '');
      if (Object.prototype.hasOwnProperty.call(normalized, key)) return normalized[key];
    }
    return fallback;
  }
  return {
    date: find(['마켓주문일자','주문일자','결제일자','주문일시'], 0),
    orderNo: find(['마켓주문번호','주문번호','주문ID','주문ID(마켓)'], 2),
    marketAccount: find(['마켓아이디','쿠팡계정ID','계정ID'], 3),
    marketProductNo: find(['마켓상품번호'], 4),
    sales: find(['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'], 6),
    status: find(['마켓주문상태','주문상태','상태','클레임상태','처리상태'], 15),
    customer: find(['고객명','수령인명','수령인','수취인','구매자','주문자'], 8),
    siteName: find(['구매사이트명','구매사이트','사이트명'], 21),
    siteProductNo: find(['사이트상품번호','상품번호','상품코드','판매자상품코드'], 22),
    brand: find(['브랜드','브랜드명'], 23),
    productName: find(['마켓상품명','상품명(원문)','상품명','상품명(옵션포함)','등록상품명'], 16),
    quantity: find(['결제수량','판매수량','수량','구매수량'], 18),
    settlement: find(['정산예정금액(원)','정산예정금액','실제정산금액','정산금액'], -1)
  };
};

vatDetailRow_v648_ = function(row, ix, sourceRow) {
  var accountId = cleanVatText_v648_(valueAt_v648_(row, ix.marketAccount >= 0 ? ix.marketAccount : 3));
  var accountMissing = !accountId;
  var status = cleanVatText_v648_(valueAt_v648_(row, ix.status));
  if (isExcludedOrderStatus_v653_(status)) return { row: null, accountMissing: accountMissing, excluded: 'status' };

  var siteName = cleanVatText_v648_(valueAt_v648_(row, ix.siteName));
  if (!/lotteon/i.test(siteName)) return { row: null, accountMissing: accountMissing, excluded: 'site' };

  var sales = vatNumber_v648_(valueAt_v648_(row, ix.sales));
  if (!sales) return { row: null, accountMissing: accountMissing, excluded: 'sales' };

  var brand = cleanVatText_v648_(valueAt_v648_(row, ix.brand));
  if (!brand) return { row: null, accountMissing: accountMissing, excluded: 'brand' };

  var siteProductNo = cleanVatText_v648_(valueAt_v648_(row, ix.siteProductNo));
  var marketProductNo = cleanVatText_v648_(valueAt_v648_(row, ix.marketProductNo));
  var productNo = siteProductNo || marketProductNo;
  if (!productNo) return { row: null, accountMissing: accountMissing, excluded: 'product' };

  var settlementActual = vatNumber_v648_(valueAt_v648_(row, ix.settlement));
  var settlement = settlementActual || Math.round(sales * 0.901);
  var purchase = vatNumber_v648_(row[28]); // AC 구매가격 단일 원천.
  var salesVat = splitVat_v648_(sales);
  var purchaseVat = splitVat_v648_(purchase);
  var fee = sales - settlement;
  var profit = settlement - purchase;
  var payableVat = salesVat.vat - purchaseVat.vat;
  var businessNo = businessNoForMarketId_v648_(accountId);
  var productName = cleanVatText_v648_(valueAt_v648_(row, ix.productName));

  return { accountMissing: accountMissing, row: [
    vatDateText_v648_(valueAt_v648_(row, ix.date)),
    accountId,
    businessNo,
    cleanVatText_v648_(valueAt_v648_(row, ix.orderNo)),
    cleanVatText_v648_(valueAt_v648_(row, ix.customer)),
    brand,
    productNo,
    productName,
    vatNumber_v648_(valueAt_v648_(row, ix.quantity)) || 1,
    sales,
    salesVat.supply,
    salesVat.vat,
    settlement,
    fee,
    purchase,
    purchaseVat.supply,
    purchaseVat.vat,
    payableVat,
    profit,
    profit - payableVat,
    accountMissing ? 'D열 마켓아이디 공란 (원본 ' + sourceRow + '행)' : (businessNo ? LOTTEON_V648_VERSION + ' 취소/반품 제외 + D열 기준' : '사업자번호 미매핑')
  ] };
};

function isExcludedOrderStatus_v653_(status) {
  return /취소|반품|환불|교환/.test(String(status || ''));
}

// 기존 메뉴 ⑥ 시트 정리: 운영 시트만 표시에서 호출되는 함수명 override.
showOperationSheetsOnly = function() {
  return cleanupOperationSheets_v653(true);
};

function cleanupOperationSheets_v653(showAlert) {
  var ss = SpreadsheetApp.getActive();
  var deleted = [];
  var hidden = [];
  var shown = [];
  var failed = [];
  var core = {};
  LOTTEON_V653_CORE_VISIBLE_SHEETS.forEach(function(name) { core[name] = true; });

  var dashboard = ss.getSheetByName('대시보드') || ss.getActiveSheet();
  if (dashboard) dashboard.activate();

  ss.getSheets().slice().forEach(function(sheet) {
    var name = sheet.getName();
    if (shouldDeleteSheet_v653_(name)) {
      try {
        if (ss.getSheets().length > 1) {
          ss.deleteSheet(sheet);
          deleted.push(name);
        }
      } catch (e) {
        failed.push(name + ': ' + String(e && e.message ? e.message : e));
      }
    }
  });

  ss.getSheets().forEach(function(sheet) {
    var name = sheet.getName();
    try {
      if (core[name]) {
        sheet.showSheet();
        shown.push(name);
      } else {
        sheet.hideSheet();
        hidden.push(name);
      }
    } catch (e) {
      failed.push(name + ': ' + String(e && e.message ? e.message : e));
    }
  });

  try { PropertiesService.getDocumentProperties().setProperty('LOTTEON_LAST_SHEET_CLEANUP_V653', JSON.stringify({ at: new Date().toISOString(), deleted: deleted, hiddenCount: hidden.length, shown: shown, failed: failed })); } catch (e) {}
  try { log_('patch_v653_sheet_cleanup', 'deleted=' + deleted.length + ' hidden=' + hidden.length + ' shown=' + shown.length + ' failed=' + failed.length); } catch (e) {}
  if (showAlert !== false) {
    safeAlert_v653_(
      '운영 시트 정리 완료\n\n' +
      '표시 시트: ' + shown.length + '개\n' +
      '숨김 시트: ' + hidden.length + '개\n' +
      '삭제 시트: ' + deleted.length + '개\n' +
      (failed.length ? '\n실패: ' + failed.slice(0, 5).join('\n') : '') +
      '\n\n표시 기준: 운영 핵심 ' + LOTTEON_V653_CORE_VISIBLE_SHEETS.length + '개 시트'
    );
  }
  return { ok: true, version: 'v6.53', shown: shown, hidden: hidden, deleted: deleted, failed: failed };
}

function shouldDeleteSheet_v653_(name) {
  if (LOTTEON_V653_CORE_VISIBLE_SHEETS.indexOf(name) >= 0) return false;
  if (LOTTEON_V653_DELETE_EXACT_SHEETS.indexOf(name) >= 0) return true;
  if (/TMP|_TMP|샘플|테스트/.test(name)) return true;
  if (/수동입력|수동전송수|전송수_수동|전송일_수동/.test(name)) return true;
  if (/부가세_고객별|부가세_주문번호별|고객주소_구분검증/.test(name)) return true;
  return false;
}

function safeAlert_v653_(message) {
  try { SpreadsheetApp.getUi().alert(message); } catch (e) {}
}

