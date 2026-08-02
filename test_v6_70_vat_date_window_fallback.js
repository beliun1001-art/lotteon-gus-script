'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const compact=v=>String(v==null?'':v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');
const text=v=>String(v==null?'':v).trim();
const company=v=>{const s=compact(v);if(s.includes('kb')||s.includes('국민'))return'KB국민카드';if(s.includes('롯데'))return'롯데카드';if(s.includes('우리'))return'우리카드';return text(v);};
let baseResult={status:'NO_MATCH',reason:'base'};
const master=[
 {company:'KB국민카드',cardName:'HERITAGE Smart(할인형)',cardNumber:'5598-69**-****-4091',cardEnd4:'4091',status:'사용'},
 {company:'우리카드',cardName:'카드의정석 EVERY POINT',cardNumber:'7680',cardEnd4:'7680',status:'사용'},
 {company:'롯데카드',cardName:'LOCA LIKIT 1.2',cardNumber:'3762-776436-56036',cardEnd4:'0036',status:'사용'}
];
const sandbox={
 console,Object,
 text_v660_:text,
 compact_v660_:compact,
 normalizeCardCompany_v660_:company,
 vatDateInSameHalf_v669_:(d,o)=>String(d||'').slice(0,4)===String(o||'').slice(0,4),
 vatMasterActiveOnDate_v669_:()=>true,
 vatPeriodCandidateIdentity_v669_:(row)=>({
   key:company(row.company)+'|END4:'+row.cardEnd4,
   company:company(row.company),
   alias:row.alias||'',
   cardName:row.cardName||'',
   cardNumber:row.cardNumber||'',
   cardEnd4:row.cardEnd4||'',
   sourceFile:row.sourceFile||'',
   label:[company(row.company),row.cardName,row.cardEnd4].filter(Boolean).join(' / ')
 }),
 matchVatOrderCardCanonical_v664_:()=>Object.assign({},baseResult)
};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Patch_v6_70_vat_date_window_fallback.gs','utf8'),sandbox);
const call=(order,result)=>{baseResult=result||{status:'NO_MATCH',reason:'base'};return sandbox.matchVatOrderCardCanonical_v664_(order,[],master,{});};

{
 const r=call({orderDate:'2026-06-22',lottePayment:'L.PAY'});
 assert.strictEqual(r.status,'MATCHED');assert.strictEqual(r.cardEnd4,'4091');assert.strictEqual(r.v670Fallback,true);
 assert.strictEqual(r.approvalAmount,0);assert.ok(r.reason.includes('LPAY_20260622_20260623_KB4091'));
}
{
 const r=call({orderDate:'2026-06-23',lottePayment:'L.PAY'});
 assert.strictEqual(r.cardEnd4,'4091');
}
{
 const r=call({orderDate:'2026-06-21',lottePayment:'L.PAY'});
 assert.strictEqual(r.status,'NO_MATCH');
}
{
 const r=call({orderDate:'2026-06-15',lottePayment:'토스페이'});
 assert.strictEqual(r.status,'NO_MATCH');
}
{
 const r=call({orderDate:'2026-06-12',lottePayment:'카카오페이'},{status:'AMBIGUOUS',reason:'period candidates'});
 assert.strictEqual(r.status,'MATCHED');assert.strictEqual(r.cardEnd4,'0036');
}
{
 const r=call({orderDate:'2026-06-24',lottePayment:'카카오페이'},{status:'AMBIGUOUS',reason:'period candidates'});
 assert.strictEqual(r.cardEnd4,'7680');
}
{
 const r=call({orderDate:'2026-06-30',lottePayment:'카카오페이'},{status:'AMBIGUOUS',reason:'period candidates'});
 assert.strictEqual(r.cardEnd4,'0036');
}
{
 const r=call({orderDate:'2026-06-26',lottePayment:'카카오페이'},{status:'AMBIGUOUS',reason:'period candidates'});
 assert.strictEqual(r.status,'AMBIGUOUS');
}
{
 const r=call({orderDate:'2026-06-22',lottePayment:''},{status:'NO_MATCH',reason:'blank'});
 assert.strictEqual(r.status,'NO_MATCH');
}
{
 const r=call({orderDate:'2026-06-22',lottePayment:'L.PAY'},{status:'MATCHED',reason:'exact',cardEnd4:'7680'});
 assert.strictEqual(r.reason,'exact');assert.strictEqual(r.cardEnd4,'7680');
}
console.log('v6.70 date-window-fallback tests PASS');
