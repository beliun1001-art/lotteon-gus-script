'use strict';
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const ctx = { console };
ctx.LOTTEON_V660_CARD_HISTORY_SHEET = '카드사용내역_붙여넣기';
ctx.LOTTEON_V660_CARD_MASTER_SHEET = '카드_마스터';
ctx.findHeaderAlias_v660_ = function(headers, names, fallback) {
  const compact = v => String(v == null ? '' : v).trim().toLowerCase().replace(/[\s._()\[\]{}\-\/]/g, '');
  for (const name of names) {
    const wanted = compact(name);
    for (let i = 0; i < headers.length; i++) if (compact(headers[i]) === wanted) return i;
  }
  return fallback;
};
ctx.text_v660_ = v => String(v == null ? '' : v).trim();
ctx.textAt_v660_ = (row, index) => index < 0 ? '' : ctx.text_v660_(row[index]);
ctx.number_v660_ = v => { const n = Number(String(v == null ? 0 : v).replace(/[원,\s]/g, '')); return Number.isNaN(n) ? 0 : n; };
ctx.pad2_v660_ = n => String(n).padStart(2, '0');
ctx.normalizeDateText_v660_ = function(v) {
  if (Object.prototype.toString.call(v) === '[object Date]' && !isNaN(v.getTime())) return v.getFullYear() + '-' + ctx.pad2_v660_(v.getMonth()+1) + '-' + ctx.pad2_v660_(v.getDate());
  const s = ctx.text_v660_(v), m = s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);
  return m ? m[1] + '-' + ctx.pad2_v660_(m[2]) + '-' + ctx.pad2_v660_(m[3]) : '';
};
ctx.isNonCardEvidence_v660_ = () => false;
ctx.isCancellationHistoryRow_v660_ = () => false;
ctx.isLotteEvidence_v660_ = h => h.lotteFlag === 'Y';
ctx.historyAmountVariants_v660_ = h => [{ amount: h.amount, label:'APPROVAL' }];
vm.createContext(ctx);
vm.runInContext(fs.readFileSync('Patch_v6_61_vat_card_date_object_fix.gs', 'utf8'), ctx);

function fakeSheet(values) {
  return { getLastRow: () => values.length, getDataRange: () => ({ getValues: () => values }) };
}
const historyHeaders = ['카드사','카드명','카드번호','카드번호끝4','승인일','승인시각','가맹점명','가맹점사업자번호','승인금액','승인번호','승인상태','취소일','취소금액','가맹점주문번호','증빙유형','롯데계열여부','원본파일','메모'];
const historyValues = [historyHeaders, ['KB국민카드','HERITAGE Smart','4091','4091',new Date(2026,5,30),'12:00','롯데쇼핑', '', 7650,'A1','승인','','','','카드이용내역','Y','KB.xls','']];
const historySs = { getSheetByName: name => name === ctx.LOTTEON_V660_CARD_HISTORY_SHEET ? fakeSheet(historyValues) : null };
const loadedHistory = ctx.loadVatCardHistory_v660_(historySs);
assert.strictEqual(loadedHistory[0].date, '2026-06-30');
assert.strictEqual(loadedHistory[0].amount, 7650);

const masterHeaders = ['카드사','카드별칭','카드명','카드구분','상태','카드번호','카드번호끝4','사업자코드','사업자등록번호','적용시작일','적용종료일','한도','메모'];
const masterValues = [masterHeaders, ['KB국민카드','KB4','HERITAGE Smart','신용','사용 가능','4091','4091','1024','606-45-93763',new Date(2026,6,1),new Date(2026,11,31),0,'']];
const masterSs = { getSheetByName: name => name === ctx.LOTTEON_V660_CARD_MASTER_SHEET ? fakeSheet(masterValues) : null };
const loadedMaster = ctx.loadVatCardMaster_v660_(masterSs);
assert.strictEqual(loadedMaster[0].startDate, '2026-07-01');
assert.strictEqual(loadedMaster[0].endDate, '2026-12-31');

const src = fs.readFileSync('Patch_v6_61_vat_card_date_object_fix.gs','utf8');
assert(src.includes('normalizeDateText_v660_(rawDate)'));
assert(src.includes('normalizeDateText_v660_(rawStart)'));
assert(!src.includes('date:normalizeDateText_v660_(textAt_v660_(row,p.date))'));
console.log('v6.61 card Date-object input fix PASS');
