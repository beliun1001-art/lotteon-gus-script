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
const values = [['주문일','쿠팡계정ID','사업자등록번호','순수매출액','정산기준금액','마켓수수료','매입금액','매출부가세','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'],['2026-01-01','beliun1021','x',100,90,10,50,10,5,5,40,35],['2026-06-30','beliun1021','x',200,180,20,100,20,10,10,80,70],['2026-07-01','beliun1023','y',300,270,30,120,30,12,18,150,132],['','beliun1024','z',400,360,40,150,40,15,25,170,145]];
const enriched = ctx.enrichVatDetailPeriods_v657_(values); const rows = ctx.aggregateVatPeriods_v657_(enriched.values);
const half1 = rows.find(r => r[0] === '반기' && r[1] === '2026' && r[2] === '상반기'); const half2 = rows.find(r => r[0] === '반기' && r[2] === '하반기'); const unknown = rows.find(r => r[0] === '기간미확인');
assert.equal(half1[7], 300); assert.equal(half2[7], 300); assert.equal(unknown[6], 1); assert.equal(unknown[7], 400);
console.log('v6.57 VAT period reconciliation mock: OK (month/half/year boundary + unknown date)');

