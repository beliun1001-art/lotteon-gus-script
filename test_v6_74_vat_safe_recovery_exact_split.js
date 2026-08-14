'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');

const compact=v=>String(v==null?'':v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');
const normCompany=v=>{
  const s=compact(v);
  if(s.includes('kb')||s.includes('국민'))return'KB국민카드';
  if(s.includes('롯데'))return'롯데카드';
  if(s.includes('우리'))return'우리카드';
  if(s.includes('신한'))return'신한카드';
  return String(v||'').trim();
};
const days=(a,b)=>Math.round((Date.parse(b+'T00:00:00Z')-Date.parse(a+'T00:00:00Z'))/86400000);
const baseMatcher=(order,history,master,used)=>{
  if(order.baseMatchKey){
    used[order.baseMatchKey]=true;
    return {status:'MATCHED',reason:'BASE',canonicalEvidenceKey:order.baseMatchKey};
  }
  if(order.baseStatus)return {status:order.baseStatus,reason:'BASE'};
  return {status:'NO_MATCH',reason:'BASE_NO_MATCH'};
};
const sandbox={
  console,Object,
  matchVatOrderCardCanonical_v664_:baseMatcher,
  daysBetween_v664_:days,
  compact_v660_:compact,
  normalizeCardCompany_v660_:normCompany,
  classifyVatTrackingPayment_v666_:v=>{
    const c=compact(v);
    if(c.includes('국민')||c.includes('kb'))return{kind:'ISSUER_CARD',issuer:'KB국민카드',raw:v};
    if(c.includes('롯데'))return{kind:'ISSUER_CARD',issuer:'롯데카드',raw:v};
    if(c.includes('우리'))return{kind:'ISSUER_CARD',issuer:'우리카드',raw:v};
    if(c.includes('카카오')&&c.includes('머니'))return{kind:'KAKAO_MONEY',issuer:'비카드',raw:v};
    if(c.includes('카카오'))return{kind:'KAKAO_CARD',issuer:'',raw:v};
    return{kind:'UNKNOWN',issuer:'',raw:v};
  },
  cardIdentityKey_v662_:(h)=>h.identity||'',
  enrichHistoryFromMaster_v660_:(h)=>h,
  historyCandidateLabel_v660_:(h)=>[h.date,h.amount,h.identity,h.v664CanonicalKey].join('|'),
  matchCanonicalHistory_v664_:(order,h,master,lag,count)=>({
    status:h.nonCard?'NON_CARD':'MATCHED',reason:'OLD',candidateCount:count,
    company:h.company||'',alias:h.alias||'',cardName:h.cardName||'',cardNumber:h.cardNumber||'',cardEnd4:h.cardEnd4||'',
    approvalDate:h.date||'',approvalTime:h.time||'',approvalNo:h.approvalNo||'',approvalAmount:h.v664OriginalAmount||h.amount||0,
    merchant:h.merchant||'',merchantOrderNo:h.merchantOrderNo||'',evidenceType:h.evidenceType||'',sourceFile:h.sourceFile||'',cancelMemo:'',
    canonicalEvidenceKey:h.v664CanonicalKey,allocationLagDays:lag,allocationCandidateCount:count
  })
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Patch_v6_74_vat_safe_recovery_exact_split.gs','utf8'),sandbox);

const ev=(key,date,amount,identity='CARD|A',company='KB국민카드',extra={})=>Object.assign({
  v664CanonicalKey:key,date,amount,v664EffectiveAmount:amount,v664OriginalAmount:amount,
  identity,company,cardName:'TEST',cardNumber:'1111-2222-3333-4091',cardEnd4:'4091',
  lotteEvidence:true,nonCard:false,cancelRow:false,v664FullyCanceled:false,
  approvalNo:key,merchant:'롯데',sourceFile:'test.csv'
},extra);
const order=(date,purchase,no,extra={})=>Object.assign({orderDate:date,purchase,orderNo:no,lottePayment:''},extra);

// A. unique -1 exact is recovered.
{
  const orders=[order('2026-06-10',10000,'A')];
  const history=[ev('E1','2026-06-09',10000)];
  const r=sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'MATCHED');
  assert.strictEqual(orders[0].cardMatch.v674RecoveryKind,'MINUS_ONE_EXACT');
  assert.strictEqual(orders[0].cardMatch.allocationLagDays,-1);
  assert.strictEqual(r.v674RecoveredDate,1);
}

// A guard: two outside exact candidates are not auto-recovered.
{
  const orders=[order('2026-06-10',10000,'B')];
  const history=[ev('E1','2026-06-09',10000),ev('E2','2026-06-08',10000)];
  sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'NO_MATCH');
}

// A guard: +8 exact is outside old window but is not the approved -1 rule.
{
  const orders=[order('2026-06-10',10000,'C')];
  const history=[ev('E1','2026-06-18',10000)];
  sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'NO_MATCH');
}

// A guard: explicit issuer conflict blocks recovery.
{
  const orders=[order('2026-06-10',10000,'D',{lottePayment:'국민카드'})];
  const history=[ev('E1','2026-06-09',10000,'CARD|L','롯데카드')];
  sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'NO_MATCH');
}

// A guard: non-card evidence never becomes MATCHED via v6.74.
{
  const orders=[order('2026-06-10',10000,'E')];
  const history=[ev('E1','2026-06-09',10000,'NONCARD|M','비카드',{nonCard:true})];
  sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'NO_MATCH');
}

// B. one exact same-card split pair is recovered and consumes both evidence rows.
{
  const orders=[order('2026-06-10',11720,'F')];
  const history=[ev('S1','2026-06-10',5000,'CARD|S'),ev('S2','2026-06-11',6720,'CARD|S')];
  const r=sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'MATCHED');
  assert.strictEqual(orders[0].cardMatch.v674RecoveryKind,'SPLIT_EXACT');
  assert.strictEqual(orders[0].cardMatch.approvalAmount,11720);
  assert.deepStrictEqual(Array.from(orders[0].cardMatch.canonicalEvidenceKeys),['S1','S2']);
  assert.strictEqual(r.v674RecoveredSplit,1);
}

// B guard: different physical-card identities are not recovered.
{
  const orders=[order('2026-06-10',11720,'G')];
  const history=[ev('S1','2026-06-10',5000,'CARD|A'),ev('S2','2026-06-11',6720,'CARD|B')];
  sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'NO_MATCH');
}

// B guard: more than one safe pair is review, not automatic.
{
  const orders=[order('2026-06-10',10000,'H')];
  const history=[
    ev('S1','2026-06-10',4000,'CARD|A'),ev('S2','2026-06-11',6000,'CARD|A'),
    ev('S3','2026-06-12',3000,'CARD|A'),ev('S4','2026-06-13',7000,'CARD|A')
  ];
  sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'NO_MATCH');
}

// Two-pass invariant: an earlier residual order cannot steal evidence reserved by a later base match.
{
  const orders=[
    order('2026-06-10',10000,'I'),
    order('2026-06-11',9999,'J',{baseMatchKey:'RESERVED'})
  ];
  const history=[ev('RESERVED','2026-06-09',10000)];
  const r=sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'NO_MATCH');
  assert.strictEqual(orders[1].cardMatch.status,'MATCHED');
  assert.strictEqual(r.v674RecoveredDate,0);
}

// Existing non-NO_MATCH result is never overwritten.
{
  const orders=[order('2026-06-10',10000,'K',{baseStatus:'NON_CARD'})];
  const history=[ev('E1','2026-06-09',10000)];
  sandbox.allocateVatPurchaseCards_v664_(orders,history,[]);
  assert.strictEqual(orders[0].cardMatch.status,'NON_CARD');
  assert.strictEqual(orders[0].cardMatch.reason,'BASE');
}

console.log('v6.74 safe recovery exact/split tests PASS');
