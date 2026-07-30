'use strict';
const fs = require('fs'), vm = require('vm'), assert = require('assert');
const ctx = {
  console,
  LOTTEON_V660_CARD_DIAG_SHEET:'부가세_카드매칭검증',
  matchVatOrderCard_v660_: () => ({status:'AMBIGUOUS', reason:'base'}),
  buildVatPurchaseCardReconciliation_v660_: () => ({summaryRows:1}),
  text_v660_: v => String(v == null ? '' : v).trim(),
  compact_v660_: v => String(v == null ? '' : v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,''),
  digits_v660_: v => String(v == null ? '' : v).replace(/\D/g,''),
  normalizeCardCompany_v660_: v => /kb|국민/i.test(String(v)) ? 'KB국민카드' : (/우리/i.test(String(v)) ? '우리카드' : ''),
  normalizeCardName_v660_: v => String(v == null ? '' : v).toLowerCase().replace(/\s/g,''),
  enrichHistoryFromMaster_v660_: h => ({company:h.company, cardName:h.cardName, cardNumber:h.cardNumber, cardEnd4:h.cardEnd4}),
  dedupeHistoryCandidates_v660_: rows => rows,
  filterEvidenceByLottePayment_v660_: rows => rows,
  historyMatchesAmount_v660_: (h,a) => Number(h.amount) === Number(a),
  matchFromHistory_v660_: (o,h,m,status,reason) => ({status,reason,company:h.company,cardEnd4:h.cardEnd4}),
  historyCandidateLabel_v660_: h => [h.company,h.cardEnd4,h.date,h.amount].join('/'),
  pad2_v660_: n => String(n).padStart(2,'0'),
  number_v660_: v => Number(v)||0,
  textAt_v660_: (row,i) => i < 0 ? '' : String(row[i] == null ? '' : row[i]).trim(),
  findHeaderAlias_v660_: () => -1,
  normalizeDateText_v660_: v => String(v || ''),
  loadVatCardHistory_v660_: () => [],
  SpreadsheetApp:{getActive:()=>null}
};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync('Patch_v6_62_vat_card_unmatched_diagnostic.gs','utf8'), ctx);

const order={orderNo:'O1',orderDate:'2026-06-30',purchase:10000,lottePayment:'KB국민카드'};
const same=[
  {company:'KB국민카드',cardEnd4:'4091',cardName:'HERITAGE',date:'2026-06-30',amount:10000,merchantOrderNo:'O1',cancelRow:false,lotteEvidence:true,nonCard:false},
  {company:'KB국민카드',cardEnd4:'4091',cardName:'HERITAGE',date:'2026-06-30',amount:10000,merchantOrderNo:'O1',cancelRow:false,lotteEvidence:true,nonCard:false}
];
let r=ctx.matchVatOrderCard_v660_(order,same,[]);
assert.strictEqual(r.status,'MATCHED');
assert.ok(r.reason.includes('동일구매카드확정'));

const different=[same[0], {...same[1], cardEnd4:'9554'}];
r=ctx.matchVatOrderCard_v660_(order,different,[]);
assert.strictEqual(r.status,'AMBIGUOUS');
assert.strictEqual(ctx.shiftDate_v662_('2026-06-30',1),'2026-07-01');
assert.strictEqual(ctx.shiftDate_v662_('2026-01-01',-1),'2025-12-31');

const insight=ctx.analyzeUnmatchedOrder_v662_('2026-06-30',10000,'KB국민카드',[
  {company:'KB국민카드',cardEnd4:'4091',date:'2026-06-29',amount:10000,cancelRow:false,lotteEvidence:true},
  {company:'KB국민카드',cardEnd4:'4091',date:'2026-06-30',amount:10500,cancelRow:false,lotteEvidence:true},
  {company:'KB국민카드',cardEnd4:'4091',date:'2026-06-30',amount:10000,cancelRow:false,lotteEvidence:false}
],'NO_MATCH');
assert.strictEqual(insight.exactNonLotte,1);
assert.strictEqual(insight.prevDayExact,1);
assert.strictEqual(insight.closestDiff,500);
console.log('v6.62 VAT card unmatched diagnostic tests PASS');
