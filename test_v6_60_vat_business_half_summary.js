'use strict';
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const sandbox = {
  console,
  buildVatPeriodSummary_v657_: function(){ return { rows: 1 }; },
  findVatPeriodHeader_v657_: function(headers, names, fallback) {
    for (const name of names) {
      const wanted = String(name).replace(/\s/g, '');
      for (let i = 0; i < headers.length; i++) {
        if (String(headers[i] || '').replace(/\s/g, '') === wanted) return i;
      }
    }
    return fallback;
  },
  vatPeriodNumber_v657_: function(value) {
    const n = Number(String(value == null ? 0 : value).replace(/[원,\s]/g, ''));
    return Number.isNaN(n) ? 0 : n;
  },
  vatPeriodFormatPlan_v657_: function(){ return {}; }
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Patch_v6_60_vat_business_half_summary.gs', 'utf8'), sandbox);

const headers = [
  '날짜','신고연도','반기','신고월','쿠팡계정ID','사업자등록번호','주문번호','순수매출액','매출공급가액','매출부가세',
  '정산기준금액','마켓수수료/비용','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'
];
const rows = [headers,
  ['05/01','2026','상반기','2026-05','acctA','111-11-11111','O1',1100,1000,100,900,200,550,500,50,50,350,300],
  ['05/02','2026','상반기','2026-05','acctB','111-11-11111','O2',2200,2000,200,1800,400,1100,1000,100,100,700,600],
  ['05/03','2026','상반기','2026-05','acctB','111-11-11111','O2',3300,3000,300,2700,600,1650,1500,150,150,1050,900],
  ['07/01','2026','하반기','2026-07','acctC','222-22-22222','O3',4400,4000,400,3600,800,2200,2000,200,200,1400,1200],
  ['10/01','2026','하반기','2026-10','unknown','', '', 550,500,50,450,100,0,0,0,50,450,400],
  ['??','기간미확인','기간미확인','기간미확인','x','333-33-33333','O4',999,908,91,900,99,100,91,9,82,800,718]
];

const out = sandbox.aggregateVatBusinessHalf_v660_(rows);
assert.strictEqual(out.length, 3);

const first = out[0];
assert.deepStrictEqual(Array.from(first.slice(0, 5)), ['2026','상반기','111-11-11111','acctA, acctB',2]);
assert.deepStrictEqual(Array.from(first.slice(5, 16)), [6600,6000,600,5400,1200,3300,3000,300,300,2100,1800]);

const second = out[1];
assert.deepStrictEqual(Array.from(second.slice(0, 5)), ['2026','하반기','222-22-22222','acctC',1]);

const third = out[2];
assert.strictEqual(third[2], '사업자번호 미매핑');
assert.strictEqual(third[3], 'unknown');
assert.strictEqual(third[4], 0);
assert.ok(String(third[16]).includes('사업자번호 미매핑'));
assert.ok(String(third[16]).includes('주문번호 공란 1행'));

assert.strictEqual(first[6], 6000);
assert.strictEqual(first[7], 600);
assert.strictEqual(first[11], 3000);
assert.strictEqual(first[12], 300);

console.log('v6.60 VAT business half summary tests PASS');
