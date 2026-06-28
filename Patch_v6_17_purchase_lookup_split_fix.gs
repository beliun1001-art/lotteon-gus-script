/**
 * LOTTEON v6.17 purchase lookup split fix
 *
 * 문제:
 * - 동일 주문번호 안에 동일 상품번호/상품명 옵션이 2개 이상 있을 때,
 *   매출데이터_붙여넣기 원본의 매입금액을 주문번호+상품번호 기준으로 SUM한 뒤
 *   각 상세행에 같은 SUM 금액을 반복 적용하는 문제가 있었습니다.
 * - 예: 주문번호 5100191148599, 상품번호 LE1221234392 2행
 *   원본 총 매입 24,240원이 각 행에 24,240원씩 반복되어 2배 집계됨.
 *
 * 수정:
 * - lookup index는 amount와 count를 같이 보관합니다.
 * - 주문번호+상품번호 또는 주문번호+상품명 매칭이 여러 행이면 amount / count로 1행당 매입금액을 배분합니다.
 * - 주문번호 단독 fallback도 여러 원본 행이면 쓰지 않습니다.
 */

var LOTTEON_PATCH_V617_PURCHASE_LOOKUP_SPLIT_FIX_LOADED = true;

function buildRawPurchaseLookup_v614_() {
  var result = {
    byOrderProduct: {},
    byOrderProductName: {},
    byOrderSingle: {},
    orderCounts: {},
    hasPurchaseColumn: false,
    purchaseHeaders: []
  };

  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(CONFIG.SHEETS.SALES_IN);
  if (!sheet || sheet.getLastRow() < 2) return result;

  var values = sheet.getDataRange().getValues();
  var header = values[0].map(function(h) { return String(h || '').trim(); });
  header.forEach(function(h) {
    if (isPurchaseHeader_v614_(h)) result.purchaseHeaders.push(h);
  });
  result.hasPurchaseColumn = result.purchaseHeaders.length > 0;

  for (var r = 1; r < values.length; r++) {
    var obj = {};
    for (var c = 0; c < header.length; c++) if (header[c]) obj[header[c]] = values[r][c];
    var purchase = getPurchaseAmountFromRow_v614_(obj);
    if (!purchase) continue;

    var orderNo = String(getFirstValue_v613_(obj, ['마켓주문번호','주문번호','주문ID','주문번호_원본']) || '').trim();
    var productNo = String(getFirstValue_v613_(obj, ['사이트상품번호','마켓상품번호','상품번호','상품ID','상품코드']) || '').trim();
    var productName = String(getFirstValue_v613_(obj, ['마켓상품명','원문상품명','상품명','상품명_원본','옵션명']) || '').trim();
    if (!orderNo) continue;

    result.orderCounts[orderNo] = (result.orderCounts[orderNo] || 0) + 1;
    if (productNo) addLookupAmountCount_v617_(result.byOrderProduct, orderNo + '||' + productNo, purchase);
    if (productName) addLookupAmountCount_v617_(result.byOrderProductName, orderNo + '||' + normalizeTextKey_v614_(productName), purchase);
    addLookupAmountCount_v617_(result.byOrderSingle, orderNo, purchase);
  }

  return result;
}

function lookupPurchaseAmount_v614_(index, orderNo, productNo, productName, netSales, qty) {
  orderNo = String(orderNo || '').trim();
  productNo = String(productNo || '').trim();
  productName = String(productName || '').trim();
  if (!orderNo || !index) return 0;

  if (productNo) {
    var byProduct = getLookupAmountPerRow_v617_(index.byOrderProduct, orderNo + '||' + productNo);
    if (byProduct) return byProduct;
  }
  if (productName) {
    var byName = getLookupAmountPerRow_v617_(index.byOrderProductName, orderNo + '||' + normalizeTextKey_v614_(productName));
    if (byName) return byName;
  }

  // 주문번호 하나에 원본 행이 하나뿐일 때만 주문번호 단독 fallback 사용.
  // 여러 옵션/상품이 섞인 주문에서 주문번호 단독 SUM을 각 행에 반복 적용하지 않기 위함.
  if ((index.orderCounts[orderNo] || 0) === 1) {
    return getLookupAmountPerRow_v617_(index.byOrderSingle, orderNo);
  }
  return 0;
}

function addLookupAmountCount_v617_(map, key, amount) {
  if (!map[key]) map[key] = { amount: 0, count: 0 };
  map[key].amount += toNumber_(amount);
  map[key].count += 1;
}

function getLookupAmountPerRow_v617_(map, key) {
  var item = map && map[key];
  if (!item) return 0;
  if (typeof item === 'number') return item;
  var amount = toNumber_(item.amount);
  var count = Math.max(1, toNumber_(item.count));
  return amount / count;
}
