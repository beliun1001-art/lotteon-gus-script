'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const compact=v=>String(v==null?'':v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');
const text=v=>String(v==null?'':v).trim();
const digits=v=>text(v).replace(/\D/g,'');
function normalizeCompany(v){const s=compact(v);if(s.includes('kb')||s.includes('국민'))return'KB국민카드';if(s.includes('롯데'))return'롯데카드';if(s.includes('우리'))return'우리카드';if(s.includes('신한'))return'신한카드';return text(v);}
function classify(v){const raw=text(v),s=compact(v);if(s.includes('카카오')&&(s.includes('머니')||s.includes('계좌')||s.includes('현금')))return{kind:'KAKAO_MONEY',raw,issuer:'비카드'};if(s.includes('카카오'))return{kind:'KAKAO_CARD',raw,issuer:''};const issuer=normalizeCompany(raw);return issuer?{kind:'ISSUER_CARD',raw,issuer}:{kind:'UNKNOWN',raw,issuer:''};}
function reason(order,rule){if(rule.kind==='KAKAO_MONEY')return'트래킹번호_카카오페이페이머니';if(rule.kind==='KAKAO_CARD')return'트래킹번호_카카오페이카드';return'트래킹번호_'+rule.issuer;}
function end4(v,n){const x=digits(v);if(x)return('0000'+x).slice(-4);const d=digits(n);return d.length>=4?d.slice(-4):'';}
let baseResult={status:'NO_MATCH',reason:'base no match'};
const sandbox={console,Object,compact_v660_:compact,text_v660_:text,digits_v660_:digits,normalizeCardCompany_v660_:normalizeCompany,classifyVatTrackingPayment_v666_:classify,trackingPaymentReasonPrefix_v666_:reason,normalizeVatCardEnd4_v667_:end4,isActiveMaster_v660_:m=>!/해지|정지/.test(text(m.status)),enrichHistoryFromMaster_v660_:(h,master)=>{const m=master.find(x=>normalizeCompany(x.company)===normalizeCompany(h.company)&&end4(x.cardEnd4,x.cardNumber)===end4(h.cardEnd4,h.cardNumber));return{company:h.company||(m&&m.company)||'',alias:(m&&m.alias)||'',cardName:(m&&m.cardName)||h.cardName||'',cardNumber:h.cardNumber||(m&&m.cardNumber)||'',cardEnd4:h.cardEnd4||(m&&m.cardEnd4)||''};},historyCandidateLabel_v660_:h=>[h.company,h.cardName,end4(h.cardEnd4,h.cardNumber),h.date].filter(Boolean).join(' / '),masterCandidateLabel_v660_:m=>[m.company,m.cardName,end4(m.cardEnd4,m.cardNumber)].filter(Boolean).join(' / '),isKakaoCardEvidence_v666_:h=>!!h.v666KakaoCard,matchVatOrderCardCanonical_v664_:()=>Object.assign({},baseResult)};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Patch_v6_69_vat_tracking_period_card_fallback.gs','utf8'),sandbox);
const call=(order,history=[],master=[])=>sandbox.matchVatOrderCardCanonical_v664_(order,history,master,{});
baseResult={status:'MATCHED',reason:'exact',company:'우리카드',cardEnd4:'7680'};
assert.strictEqual(call({orderDate:'2026-06-01',lottePayment:'우리카드'}).reason,'exact');
baseResult={status:'AMBIGUOUS',reason:'exact ambiguous'};
assert.strictEqual(call({orderDate:'2026-06-01',lottePayment:'우리카드'}).reason,'exact ambiguous');
baseResult={status:'NO_MATCH',reason:'amount mismatch'};
{
 const r=call({orderDate:'2026-06-10',lottePayment:'KB국민카드'},[{company:'KB국민카드',cardName:'HERITAGE Smart(할인형)',cardEnd4:'4091',date:'2026-05-02'},{company:'KB국민카드',cardName:'과거카드',cardEnd4:'1111',date:'2025-12-31'},{company:'우리카드',cardName:'EVERY POINT',cardEnd4:'7680',date:'2026-05-02'}]);
 assert.strictEqual(r.status,'MATCHED');assert.strictEqual(r.cardEnd4,'4091');assert.strictEqual(r.approvalAmount,0);assert.ok(r.reason.includes('금액비교없음'));
}
{
 const history=[{company:'롯데카드',cardName:'Trip to 로카',cardEnd4:'0126',date:'2026-05-01'},{company:'롯데카드',cardName:'LOCA LIKIT 1.2',cardEnd4:'0036',date:'2026-06-01'}];
 assert.strictEqual(call({orderDate:'2026-05-28',lottePayment:'롯데카드'},history).cardEnd4,'0126');
 assert.strictEqual(call({orderDate:'2026-05-29',lottePayment:'롯데카드'},history).cardEnd4,'0036');
}
{
 const r=call({orderDate:'2026-06-10',lottePayment:'신한카드'},[{company:'신한카드',cardName:'A',cardEnd4:'1111',date:'2026-05-01'},{company:'신한카드',cardName:'B',cardEnd4:'2222',date:'2026-06-01'}]);
 assert.strictEqual(r.status,'AMBIGUOUS');assert.strictEqual(r.candidateCount,2);
}
{
 const r=call({orderDate:'2026-06-10',lottePayment:'카카오페이 페이머니'});
 assert.strictEqual(r.status,'NON_CARD');assert.strictEqual(r.alias,'신한은행 계좌결제');
}
{
 const r=call({orderDate:'2026-06-10',lottePayment:'카카오페이'},[{company:'우리카드',cardName:'EVERY POINT',cardEnd4:'7680',date:'2026-06-02',v666KakaoCard:true},{company:'롯데카드',cardName:'LOCA',cardEnd4:'0036',date:'2026-06-03',v666KakaoCard:false}]);
 assert.strictEqual(r.status,'MATCHED');assert.strictEqual(r.cardEnd4,'7680');
}
{
 const r=call({orderDate:'2026-06-10',lottePayment:'우리카드'},[{company:'우리카드',cardName:'EVERY POINT',cardEnd4:'7680',date:'2025-12-01'}]);
 assert.strictEqual(r.status,'NO_MATCH');assert.ok(r.reason.includes('후보없음'));
}
{
 const r=call({orderDate:'2026-06-10',lottePayment:'KB국민카드'},[{company:'KB국민카드',cardName:'알수없는 국민카드',cardEnd4:'1111',date:'2026-05-01'}]);
 assert.strictEqual(r.status,'NO_MATCH');assert.ok(r.reason.includes('후보없음'));
}
console.log('v6.69 tracking-period-card-fallback tests PASS');
