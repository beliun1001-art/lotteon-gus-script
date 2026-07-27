'use strict';
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const baseHeaders = ['날짜','쿠팡계정ID','사업자등록번호','주문번호','순수매출액'];
const sandbox = {
  console,
  buildVatPeriodSummary_v657_: function(){ return { rows: 1 }; },
  vatHeaderIndexes_v648_: function(){ return { date:0, orderNo:2, sales:6 }; },
  vatDetailHeaders_v648_: function() { return baseHeaders.slice(); },
  vatDetailRow_v648_: function() { return { row: ['05/01','acct','111-11-11111','O1',1100], accountMissing:false }; },
  valueAt_v648_: function(row, index) { return index >= 0 && index < row.length ? row[index] : ''; },
  cleanVatText_v648_: function(value) { return String(value == null ? '' : value).trim(); },
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

const sourceHeaders = ['날짜','주문번호','기타','결제수단'];
const ix = sandbox.vatHeaderIndexes_v648_(sourceHeaders);
assert.strictEqual(ix.payment, 3);
assert.strictEqual(sandbox.findVatPaymentHeader_v660_(['x','카드사']), 1);
assert.strictEqual(sandbox.findVatPaymentHeader_v660_(['x','없는필드']), -1);
assert.strictEqual(sandbox.vatDetailHeaders_v648_().slice(-1)[0], '결제수단');
const detailResult = sandbox.vatDetailRow_v648_(['05/01','O1','x','우리카드(일시불)'], ix, 2);
assert.strictEqual(detailResult.row.slice(-1)[0], '우리카드(일시불)');
const unknownPayment = sandbox.vatDetailRow_v648_(['05/01','O1','x',''], ix, 3);
assert.strictEqual(unknownPayment.row.slice(-1)[0], '결제수단 미확인');

const headers = [
  '날짜','신고연도','반기','신고월','쿠팡계정ID','사업자등록번호','주문번호','순수매출액','매출공급가액','매출부가세',
  '정산기준금액','마켓수수료/비용','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','결제수단'
];
const rows = [headers,
  ['05/01','2026','상반기','2026-05','acctA','111-11-11111','O1',1100,1000,100,900,200,550,500,50,50,350,300,'신한카드'],
  ['05/02','2026','상반기','2026-05','acctB','111-11-11111','O2',2200,2000,200,1800,400,1100,1000,100,100,700,600,'우리카드(일시불)'],
  ['05/03','2026','상반기','2026-05','acctB','111-11-11111','O2',3300,3000,300,2700,600,1650,1500,150,150,1050,900,'우리카드(일시불)'],
  ['06/01','2026','상반기','2026-06','acctA','111-11-11111','O5',550,500,50,450,100,0,0,0,50,450,400,'L.PAY신용카드'],
  ['07/01','2026','하반기','2026-07','acctC','222-22-22222','O3',4400,4000,400,3600,800,2200,2000,200,200,1400,1200,'카카오페이 머니'],
  ['10/01','2026','하반기','2026-10','unknown','', '', 550,500,50,450,100,0,0,0,50,450,400,''],
  ['??','기간미확인','기간미확인','기간미확인','x','333-33-33333','O4',999,908,91,900,99,100,91,9,82,800,718,'신한카드']
];

const out = sandbox.aggregateVatBusinessHalf_v660_(rows);
assert.strictEqual(out.length, 5);
const sh = out.find(r => r[2] === '111-11-11111' && r[3] === '신한카드');
const woori = out.find(r => r[2] === '111-11-11111' && r[3] === '우리카드(일시불)');
const lpay = out.find(r => r[2] === '111-11-11111' && r[3] === 'L.PAY신용카드');
assert.ok(sh && woori && lpay);
assert.deepStrictEqual(Array.from(sh.slice(0, 6)), ['2026','상반기','111-11-11111','신한카드','acctA',1]);
assert.deepStrictEqual(Array.from(woori.slice(0, 6)), ['2026','상반기','111-11-11111','우리카드(일시불)','acctB',1]);
assert.deepStrictEqual(Array.from(woori.slice(6, 17)), [5500,5000,500,4500,1000,2750,2500,250,250,1750,1500]);
assert.strictEqual(lpay[6], 550);
const unknown = out.find(r => r[2] === '사업자번호 미매핑');
assert.strictEqual(unknown[3], '결제수단 미확인');
assert.strictEqual(unknown[5], 0);
assert.ok(String(unknown[17]).includes('사업자번호 미매핑'));
assert.ok(String(unknown[17]).includes('결제수단 미확인'));
assert.ok(String(unknown[17]).includes('주문번호 공란 1행'));
const kakao = out.find(r => r[3] === '카카오페이 머니');
assert.ok(kakao);
assert.strictEqual(kakao[6], 4400);
console.log('v6.60 VAT business/payment half summary tests PASS');
