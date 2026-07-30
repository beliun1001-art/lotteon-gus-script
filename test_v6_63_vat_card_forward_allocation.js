'use strict';
const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const sandbox = {
  console,
  SpreadsheetApp: { getActive(){ return null; } },
  LOTTEON_V660_MASTER_CUTOFF: '2026-07-01',
  text_v660_: v => String(v == null ? '' : v).trim(),
  normalizeCardCompany_v660_: v => {
    const s = String(v || '').toLowerCase();
    if (s.includes('kb') || s.includes('국민')) return 'KB국민카드';
    if (s.includes('우리')) return '우리카드';
    if (s.includes('롯데')) return '롯데카드';
    if (s.includes('신한')) return '신한카드';
    return '';
  },
  shiftDate_v662_: (dateText, days) => {
    const d = new Date(dateText + 'T00:00:00Z');
    d.setUTCDate(d.getUTCDate() + days);
    return d.toISOString().slice(0,10);
  },
  historyMatchesAmount_v660_: (h, amount) => Number(h.amount) === Number(amount),
  filterEvidenceByLottePayment_v660_: (rows, payment) => rows,
  dedupeHistoryCandidates_v660_: rows => rows.slice(),
  cardIdentityKey_v662_: h => h.nonCard ? 'NONCARD|kakao' : `${h.company}|${h.cardEnd4 || h.cardNumber || h.cardName}`,
  matchFromHistory_v660_: (order,h,master,status,reason) => ({
    status, reason, company:h.company || (h.nonCard ? '비카드' : ''), cardName:h.cardName || '',
    cardNumber:h.cardNumber || '', cardEnd4:h.cardEnd4 || '', candidateCount:1
  }),
  noMatch_v660_: reason => ({status:'NO_MATCH',reason,candidateCount:0}),
  findPostJulyMasterCandidates_v660_: () => [],
  matchFromMaster_v660_: () => ({status:'MASTER_MATCHED'}),
  ambiguousMasterMatch_v660_: () => ({status:'AMBIGUOUS'}),
  historyCandidateLabel_v660_: h => `${h.company}|${h.cardEnd4}|${h.date}|${h.amount}`,
  analyzeUnmatchedOrder_v662_: () => ({}),
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Patch_v6_63_vat_card_forward_allocation.gs', 'utf8'), sandbox);

function order(no, date, amount, payment='') {
  return {orderNo:no, orderDate:date, purchase:amount, lottePayment:payment};
}
function h(rowNo, date, amount, company, end4, nonCard=false) {
  return {rowNo,date,amount,company,cardEnd4:end4,cardName:company,nonCard,lotteEvidence:true,cancelRow:false,time:'10:00'};
}

// +2 day evidence should match because source date is market order date, not purchase date.
{
  const o = order('C1','2026-05-01',10000);
  const used = {};
  const r = sandbox.matchVatOrderCardForwardAllocated_v663_(o,[h(1,'2026-05-03',10000,'KB국민카드','4091')],[],used);
  assert.strictEqual(r.status,'MATCHED');
  assert.strictEqual(r.allocationLagDays,2);
  assert.ok(r.reason.includes('+2일'));
  assert.strictEqual(used['1'],true);
}

// Evidence must never be reused for a second resale order.
{
  const orders = [order('C1','2026-05-01',10000), order('C2','2026-05-01',10000)];
  sandbox.allocateVatPurchaseCards_v663_(orders,[h(1,'2026-05-01',10000,'KB국민카드','4091')],[]);
  const statuses = orders.map(x=>x.cardMatch.status).sort();
  assert.deepStrictEqual(statuses,['MATCHED','NO_MATCH']);
}

// Two same-card evidence rows can support two orders one-to-one.
{
  const orders = [order('C1','2026-05-01',10000), order('C2','2026-05-01',10000)];
  const hist = [h(1,'2026-05-02',10000,'우리카드','7680'),h(2,'2026-05-02',10000,'우리카드','7680')];
  sandbox.allocateVatPurchaseCards_v663_(orders,hist,[]);
  assert.deepStrictEqual(orders.map(x=>x.cardMatch.status),['MATCHED','MATCHED']);
  assert.ok(orders.every(x=>x.cardMatch.cardEnd4==='7680'));
}

// Different physical-card candidates on the earliest evidence day remain ambiguous.
{
  const o = order('C1','2026-05-01',10000);
  const hist = [h(1,'2026-05-02',10000,'KB국민카드','4091'),h(2,'2026-05-02',10000,'우리카드','7680')];
  const r = sandbox.matchVatOrderCardForwardAllocated_v663_(o,hist,[],{});
  assert.strictEqual(r.status,'AMBIGUOUS');
}

// Previous-day evidence must not be used: purchase cannot precede the market order in this workflow.
{
  const o = order('C1','2026-05-02',10000);
  const r = sandbox.matchVatOrderCardForwardAllocated_v663_(o,[h(1,'2026-05-01',10000,'KB국민카드','4091')],[],{});
  assert.strictEqual(r.status,'NO_MATCH');
}

// Known issuer must not be polluted by non-card evidence.
{
  const o = order('C1','2026-05-01',10000,'KB국민카드');
  const hist = [h(1,'2026-05-01',10000,'비카드','',true),h(2,'2026-05-01',10000,'KB국민카드','4091')];
  const r = sandbox.matchVatOrderCardForwardAllocated_v663_(o,hist,[],{});
  assert.strictEqual(r.status,'MATCHED');
  assert.strictEqual(r.cardEnd4,'4091');
}

console.log('v6.63 VAT forward allocation tests PASS');
