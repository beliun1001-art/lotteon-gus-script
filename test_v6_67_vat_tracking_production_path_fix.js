'use strict';
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const compact = v => String(v == null ? '' : v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g, '');
const text = v => String(v == null ? '' : v).trim();
const digits = v => text(v).replace(/\D/g, '');
function findAlias(headers, names, fallback) {
  for (const name of names) {
    const wanted = compact(name);
    for (let i = 0; i < headers.length; i++) if (compact(headers[i]) === wanted) return i;
  }
  return fallback;
}
function normalizeCompany(v) {
  const s = compact(v);
  if (s.includes('롯데')) return '롯데카드';
  if (s.includes('kb') || s.includes('국민')) return 'KB국민카드';
  if (s.includes('우리')) return '우리카드';
  if (s.includes('비카드') || s.includes('머니')) return '비카드';
  return text(v);
}

let detailWrite = null;
let sourceReadWidths = [];
let savedState = null;
const headers = Array.from({ length: 39 }, (_, i) => '열' + (i + 1));
headers[3] = '마켓아이디';
headers[28] = '매입금액';
headers[29] = '현지트래킹번호';
headers[34] = '롯데결제수단';

const sourceRows = [
  Array.from({ length: 39 }, () => ''),
  Array.from({ length: 39 }, () => '')
];
sourceRows[0][3] = 'beliun1021';
sourceRows[0][28] = 21910;
sourceRows[0][29] = '롯데카드(일시불)';
sourceRows[0][34] = '우리카드';
sourceRows[1][3] = 'beliun1021';
sourceRows[1][28] = 31500;
sourceRows[1][29] = '';
sourceRows[1][34] = '우리카드(일시불)';

const source = {
  getLastRow: () => 3,
  getLastColumn: () => 39,
  getRange(row, col, count, width) {
    if (row === 1) return { getValues: () => [headers.slice(0, width)] };
    sourceReadWidths.push(width);
    const start = row - 2;
    return { getValues: () => sourceRows.slice(start, start + count).map(r => r.slice(0, width)) };
  }
};
const detail = {
  getRange(row, col, count, width) {
    return { setValues(values) { detailWrite = { row, col, count, width, values }; return this; } };
  }
};
const ss = { getSheetByName: name => name === '매출데이터_붙여넣기' ? source : detail };

const sandbox = {
  console,
  isFinite,
  LOTTEON_V648_SOURCE_SHEET: '매출데이터_붙여넣기',
  LOTTEON_V648_DETAIL_SHEET: '부가세_신고자료',
  LOTTEON_V648_MAX_COL: 29,
  LOTTEON_V648_CHUNK_SIZE: 500,
  findHeaderAlias_v660_: findAlias,
  compact_v660_: compact,
  text_v660_: text,
  digits_v660_: digits,
  normalizeCardCompany_v660_: normalizeCompany,
  saveVatState_v648_: state => { savedState = JSON.parse(JSON.stringify(state)); },
  scheduleVatTrigger_v648_: () => {},
  writeVatStatus_v648_: () => {},
  runVatDetailBatch_v648_: () => ({ base: true }),
  enrichHistoryFromMaster_v660_: h => ({
    company: h.company, cardName: h.cardName, cardNumber: h.cardNumber, cardEnd4: h.cardEnd4
  }),
  loadVatCardHistory_v660_: () => [
    { company: '롯데카드', cardName: 'Trip to 로카', cardNumber: '3779-731600-80126', cardEnd4: 126 },
    { company: '롯데카드', cardName: 'LOCA LIKIT 1.2', cardNumber: '3762-776436-56036', cardEnd4: 36 }
  ],
  loadVatCardMaster_v660_: () => [
    { company: '롯데카드', cardName: 'Trip to 로카', cardNumber: '3779-731600-80126', cardEnd4: 126 },
    { company: '롯데카드', cardName: 'LOCA LIKIT 1.2', cardNumber: '3762-776436-56036', cardEnd4: 36 }
  ],
  matchVatOrderCardCanonical_v664_: (order, history) => Object.assign({ status: 'MATCHED' }, history[0] || {})
};

sandbox.findVatTrackingPaymentHeader_v666_ = hs => findAlias(hs, ['트래킹 번호'], -1);
sandbox.findVatFallbackPaymentHeader_v666_ = hs => findAlias(hs, ['결제수단'], -1);
sandbox.vatHeaderIndexes_v648_ = hs => {
  const tracking = sandbox.findVatTrackingPaymentHeader_v666_(hs);
  const fallback = sandbox.findVatFallbackPaymentHeader_v666_(hs);
  return { lottePayment: tracking >= 0 ? tracking : fallback, v666TrackingPayment: tracking, v666FallbackPayment: fallback };
};
sandbox.vatDetailRow_v648_ = (row, ix) => ({
  row: ['base', Number(row[28] || 0), text(row[ix.v666TrackingPayment]) || text(row[ix.v666FallbackPayment])],
  accountMissing: !text(row[3])
});
sandbox.filterLotteCardByOrderDate_v666_ = (rows, orderDate, master) => {
  const expected = orderDate <= '2026-05-28' ? 'TRIP' : 'LIKIT';
  return rows.filter(h => sandbox.lotteCardKind_v666_(h, master) === expected);
};

vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Patch_v6_67_vat_tracking_production_path_fix.gs', 'utf8'), sandbox);

assert.strictEqual(sandbox.findVatTrackingPaymentHeader_v666_(headers), 29);
assert.strictEqual(sandbox.findVatFallbackPaymentHeader_v666_(headers), 34);

const state = { sourceLastRow: 3, sourceRow: 2, writtenRows: 0, skippedRows: 0, accountMissingRows: 0, phase: 'detail' };
const batch = sandbox.runVatDetailBatch_v648_(ss, state);
assert.strictEqual(batch.ok, true);
assert.strictEqual(sourceReadWidths[0], 35);
assert.strictEqual(detailWrite.values.length, 2);
assert.strictEqual(detailWrite.values[0][2], '롯데카드(일시불)');
assert.strictEqual(detailWrite.values[1][2], '우리카드(일시불)');
assert.strictEqual(detailWrite.values[0][1], 21910);
assert.strictEqual(savedState.sourceReadColumns, 35);
assert.strictEqual(savedState.trackingPaymentColumn, 30);
assert.strictEqual(savedState.fallbackPaymentColumn, 35);

assert.strictEqual(sandbox.normalizeVatCardEnd4_v667_(126, '3779-731600-80126'), '0126');
assert.strictEqual(sandbox.normalizeVatCardEnd4_v667_(36, '3762-776436-56036'), '0036');
assert.strictEqual(sandbox.normalizeVatCardEnd4_v667_('4091', '5598-69**-****-4091'), '4091');
assert.strictEqual(sandbox.normalizeVatCardEnd4_v667_('7680', '7680'), '7680');

const history = sandbox.loadVatCardHistory_v660_({});
assert.strictEqual(history[0].cardEnd4, '0126');
assert.strictEqual(history[1].cardEnd4, '0036');
const master = sandbox.loadVatCardMaster_v660_({});
assert.strictEqual(master[0].cardEnd4, '0126');
assert.strictEqual(master[1].cardEnd4, '0036');

const mislabeledLoca = { company: '롯데카드', cardName: 'Trip to 로카', cardNumber: '3762-776436-56036', cardEnd4: '36' };
const physicalTrip = { company: '롯데카드', cardName: 'LOCA LIKIT 1.2', cardNumber: '3779-731600-80126', cardEnd4: '126' };
assert.strictEqual(sandbox.lotteCardKind_v666_(mislabeledLoca, []), 'LIKIT');
assert.strictEqual(sandbox.lotteCardKind_v666_(physicalTrip, []), 'TRIP');
assert.deepStrictEqual(
  sandbox.filterLotteCardByOrderDate_v666_([mislabeledLoca, physicalTrip], '2026-05-27', []),
  [physicalTrip]
);
assert.deepStrictEqual(
  sandbox.filterLotteCardByOrderDate_v666_([mislabeledLoca, physicalTrip], '2026-05-29', []),
  [mislabeledLoca]
);

const normalizedMatch = sandbox.matchVatOrderCardCanonical_v664_(
  { orderDate: '2026-05-29' }, [mislabeledLoca], [], {}
);
assert.strictEqual(normalizedMatch.cardEnd4, '0036');
assert.strictEqual(normalizedMatch.cardName, 'LOCA LIKIT 1.2');

console.log('v6.67 tracking production-path tests PASS');
