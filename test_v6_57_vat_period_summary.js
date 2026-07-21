const assert = require('assert');
const fs = require('fs');
const vm = require('vm');
const source = fs.readFileSync('Patch_v6_57_vat_period_summary.gs', 'utf8');
const ctx = {}; vm.createContext(ctx); vm.runInContext(source, ctx);

assert.deepStrictEqual(JSON.parse(JSON.stringify(ctx.vatPeriodFromDate_v657_('2026-01-01'))), { year:'2026', half:'상반기', month:'2026-01' });
assert.equal(ctx.vatPeriodFromDate_v657_('2026-06-30').half, '상반기');
assert.equal(ctx.vatPeriodFromDate_v657_('2026-07-01').half, '하반기');
assert.equal(ctx.vatPeriodFromDate_v657_('2026-12-31').month, '2026-12');
assert.equal(ctx.vatPeriodFromDate_v657_('').year, '기간미확인');

const values = [
  ['주문일','쿠팡계정ID','사업자등록번호','주문번호','상품번호','순수매출액','정산기준금액','마켓수수료','매입금액','매출부가세','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'],
  ['2026-01-01','beliun1021','x','same-order','A',100,90,10,50,10,5,5,40,35],
  ['2026-06-30','beliun1021','x','same-order','B',200,180,20,100,20,10,10,80,70],
  ['2026-07-01','beliun1023','y','order-2','C',300,270,30,120,30,12,18,150,132],
  ['','beliun1024','z','','D',400,360,40,150,40,15,25,170,145]
];
const enriched = ctx.enrichVatDetailPeriods_v657_(values);
const rows = ctx.aggregateVatPeriods_v657_(enriched.values);
const half1 = rows.find(r => r[0] === '반기' && r[1] === '2026' && r[2] === '상반기');
const half2 = rows.find(r => r[0] === '반기' && r[2] === '하반기');
const unknown = rows.find(r => r[0] === '기간미확인');
assert.equal(half1[6], 1); // same order, two products => one distinct order
assert.equal(half1[7], 300); assert.equal(half2[7], 300); assert.equal(unknown[6], 0); assert.equal(unknown[7], 400);
assert.equal(rows.blankOrderDetailRows, 1); // blank order number is never guessed as one order
const monthlySales = rows.filter(r => r[0] === '월별').reduce((sum, r) => sum + r[7], 0);
const halfSales = rows.filter(r => r[0] === '반기').reduce((sum, r) => sum + r[7], 0);
assert.equal(monthlySales, halfSales); assert.equal(halfSales + unknown[7], 1000);

// Integration-style path: source keeps year, detail shows only MM/dd, duplicate keys consume stable queues.
const sourceRows = [
  ['2025-01-01','acct','','ORD','','SKU',100],
  ['2026-01-01','acct','','ORD','','SKU',100]
];
ctx.LOTTEON_V648_SOURCE_SHEET = '매출데이터_붙여넣기'; ctx.LOTTEON_V648_MAX_COL = 29;
ctx.vatHeaderIndexes_v648_ = () => ({ date: 0 }); ctx.valueAt_v648_ = (row, index) => row[index];
ctx.vatDetailRow_v648_ = row => ({ row: [row[0].slice(5).replace('-', '/'), row[1], 'biz', row[3], '', '', row[5], '', '', row[6]] });
const sourceSheet = { getLastColumn: () => 7, getLastRow: () => 3, getRange: (row) => ({ getValues: () => row === 1 ? [['마켓주문일자']] : sourceRows }) };
const ss = { getSheetByName: name => name === '매출데이터_붙여넣기' ? sourceSheet : null };
const detailMmdd = [['주문일','쿠팡계정ID','사업자등록번호','주문번호','고객명','브랜드명','상품번호','상품명','수량','순수매출액'], ['01/01','acct','biz','ORD','','','SKU','',1,100], ['01/01','acct','biz','ORD','','','SKU','',1,100]];
const restored = ctx.enrichVatDetailPeriods_v657_(detailMmdd, ss).values;
assert.deepStrictEqual(JSON.parse(JSON.stringify(restored.slice(1).map(r => [r[1], r[2], r[3]]))), [['2025','상반기','2025-01'], ['2026','상반기','2026-01']]);

const plan = ctx.vatPeriodFormatPlan_v657_(['신고연도','반기','신고월','주문번호','상품번호','계정ID','사업자등록번호','순수매출액']);
assert.deepStrictEqual(JSON.parse(JSON.stringify(plan)), { 0:'@', 1:'@', 2:'@', 3:'@', 4:'@', 5:'@', 6:'@', 7:'#,##0' });
let saved, written;
ctx.saveVatState_v648_ = state => { saved = state; }; ctx.writeVatStatus_v648_ = (ss, state) => { written = state; }; ctx.toastVat_v648_ = () => {};
const failed = ctx.finalizeVatPeriodFailure_v657_({}, { status: 'done' }, new Error('mock failure'));
assert.equal(failed.status, 'failed'); assert.match(saved.lastError, /mock failure/); assert.equal(written.status, 'failed');
console.log('v6.57 VAT period mock: OK (boundaries, unknown, MM/dd restore, distinct order, reconciliation, format, failed status)');

