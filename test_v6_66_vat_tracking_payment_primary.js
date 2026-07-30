'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');

const compact=v=>String(v==null?'':v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');
const text=v=>String(v==null?'':v).trim();
function findAlias(headers,names,fallback){for(const name of names){const wanted=compact(name);for(let i=0;i<headers.length;i++)if(compact(headers[i])===wanted)return i;}return fallback;}
function normalizeCompany(v){const s=compact(v);if(s.includes('kb')||s.includes('국민'))return'KB국민카드';if(s.includes('롯데'))return'롯데카드';if(s.includes('우리'))return'우리카드';if(s.includes('신한'))return'신한카드';if(s.includes('비카드')||s.includes('머니'))return'비카드';return text(v);}

let lastFiltered=[];
const sandbox={
  console,
  findHeaderAlias_v660_:findAlias,
  cleanVatText_v648_:text,
  valueAt_v648_:(row,i)=>row[i],
  text_v660_:text,
  compact_v660_:compact,
  normalizeCardCompany_v660_:normalizeCompany,
  enrichHistoryFromMaster_v660_:(h)=>({company:h.company,cardName:h.cardName,cardNumber:h.cardNumber,cardEnd4:h.cardEnd4}),
  findVatPaymentHeader_v660_:(headers)=>findAlias(headers,['결제수단'],-1),
  vatHeaderIndexes_v648_:(headers)=>({lottePayment:findAlias(headers,['결제수단'],-1)}),
  vatDetailRow_v648_:(row,ix)=>({row:['base',ix.lottePayment>=0?text(row[ix.lottePayment]):'']}),
  canonicalizeVatHistory_v664_:(history)=>{
    const map={};
    for(const h of history){
      if(h.nonCard){map['ROW|'+h.rowNo]=Object.assign({},h,{v664CanonicalKey:'ROW|'+h.rowNo});continue;}
      const key='CARD|'+normalizeCompany(h.company)+'|'+h.approvalNo;
      if(!map[key]||h.evidenceType==='카드이용내역')map[key]=Object.assign({},h,{v664CanonicalKey:key});
    }
    return Object.values(map);
  },
  matchVatOrderCardCanonical_v664_:(order,history)=>{
    lastFiltered=history.slice();
    if(!history.length)return{status:'NO_MATCH',reason:'base no match'};
    const h=history[0];
    return{status:h.nonCard?'NON_CARD':'MATCHED',reason:'base matched',company:h.company,cardName:h.cardName,alias:''};
  }
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Patch_v6_66_vat_tracking_payment_primary.gs','utf8'),sandbox);

{
  const headers=['주문번호','결제수단','트래킹 번호'];
  const ix=sandbox.vatHeaderIndexes_v648_(headers);
  assert.strictEqual(ix.v666TrackingPayment,2);
  assert.strictEqual(ix.v666FallbackPayment,1);
  let r=sandbox.vatDetailRow_v648_(['1','롯데카드','카카오페이'],ix);
  assert.strictEqual(r.row[r.row.length-1],'카카오페이');
  r=sandbox.vatDetailRow_v648_(['1','우리카드',''],ix);
  assert.strictEqual(r.row[r.row.length-1],'우리카드');
}

assert.strictEqual(sandbox.classifyVatTrackingPayment_v666_('카카오페이').kind,'KAKAO_CARD');
assert.strictEqual(sandbox.classifyVatTrackingPayment_v666_('카카오페이 페이머니').kind,'KAKAO_MONEY');
assert.strictEqual(sandbox.classifyVatTrackingPayment_v666_('KB국민카드').issuer,'KB국민카드');
assert.strictEqual(sandbox.classifyVatTrackingPayment_v666_('롯데카드').issuer,'롯데카드');

{
  const raw=[
    {rowNo:1,company:'우리카드',cardName:'카드의정석 EVERY POINT',approvalNo:'A1',evidenceType:'카드이용내역',nonCard:false},
    {rowNo:2,company:'우리카드',cardName:'카카오페이 카드결제',approvalNo:'A1',evidenceType:'카카오페이 카드',nonCard:false}
  ];
  const c=sandbox.canonicalizeVatHistory_v664_(raw,[]);
  assert.strictEqual(c.length,1);
  assert.strictEqual(c[0].cardName,'카드의정석 EVERY POINT');
  assert.strictEqual(c[0].v666KakaoCard,true);
}

{
  const history=[
    {company:'롯데카드',cardName:'LOCA LIKIT 1.2',nonCard:false,v666KakaoCard:false},
    {company:'우리카드',cardName:'카드의정석 EVERY POINT',nonCard:false,v666KakaoCard:true}
  ];
  const r=sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-06-10',lottePayment:'카카오페이'},history,[],{});
  assert.strictEqual(lastFiltered.length,1);
  assert.strictEqual(lastFiltered[0].company,'우리카드');
  assert.strictEqual(r.status,'MATCHED');
  assert.ok(r.reason.includes('카카오페이카드'));
}

{
  const history=[
    {company:'우리카드',cardName:'카카오페이 카드결제',nonCard:false,v666KakaoCard:true},
    {company:'비카드',cardName:'카카오페이 페이머니',evidenceType:'카카오페이 페이머니',nonCard:true,v666KakaoMoney:true}
  ];
  const r=sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-06-10',lottePayment:'카카오페이 페이머니'},history,[],{});
  assert.strictEqual(lastFiltered.length,1);
  assert.strictEqual(lastFiltered[0].nonCard,true);
  assert.strictEqual(r.status,'NON_CARD');
  assert.strictEqual(r.alias,'신한은행 계좌결제');
  assert.strictEqual(r.cardName,'카카오페이 페이머니');
}

{
  const history=[
    {company:'롯데카드',cardName:'Trip to 로카',nonCard:false},
    {company:'롯데카드',cardName:'LOCA LIKIT 1.2',nonCard:false}
  ];
  sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-05-28',lottePayment:'롯데카드'},history,[],{});
  assert.strictEqual(lastFiltered.length,1);
  assert.strictEqual(lastFiltered[0].cardName,'Trip to 로카');
  sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-05-29',lottePayment:'롯데카드'},history,[],{});
  assert.strictEqual(lastFiltered.length,1);
  assert.strictEqual(lastFiltered[0].cardName,'LOCA LIKIT 1.2');
}

{
  const history=[
    {company:'KB국민카드',cardName:'HERITAGE Smart(할인형)',nonCard:false},
    {company:'우리카드',cardName:'카드의정석 EVERY POINT',nonCard:false}
  ];
  const r=sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-06-10',lottePayment:'KB국민카드'},history,[],{});
  assert.strictEqual(lastFiltered.length,1);
  assert.strictEqual(lastFiltered[0].company,'KB국민카드');
  assert.ok(r.reason.includes('HERITAGE단일카드'));
}

console.log('v6.66 tracking-payment-primary tests PASS');
