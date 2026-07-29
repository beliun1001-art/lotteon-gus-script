'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const sandbox={
  console,
  SpreadsheetApp:{getActive(){return null;}},
  text_v660_:v=>String(v==null?'':v).trim(),
  normalizeCardCompany_v660_:v=>{const s=String(v||'').toLowerCase(); if(s.includes('롯데'))return'롯데카드'; if(s.includes('우리'))return'우리카드'; if(s.includes('kb')||s.includes('국민'))return'KB국민카드'; if(s.includes('비카드'))return'비카드'; return String(v||'');},
  enrichHistoryFromMaster_v660_:(h)=>({company:h.company,cardName:h.cardName,cardNumber:h.cardNumber,cardEnd4:h.cardEnd4}),
  shiftDate_v662_:(dateText,days)=>{const d=new Date(dateText+'T00:00:00Z'); d.setUTCDate(d.getUTCDate()+days); return d.toISOString().slice(0,10);},
  cardIdentityKey_v662_:(h)=>h.nonCard?'NONCARD|'+h.cardName:(h.company+'|'+(h.cardEnd4||h.cardNumber||h.cardName)),
  filterEvidenceByLottePayment_v660_:(rows)=>rows,
  ambiguousForwardMatch_v663_:(candidates,lag,detail)=>({status:'AMBIGUOUS',candidateCount:candidates.length,allocationLagDays:lag,reason:detail}),
  noMatch_v660_:reason=>({status:'NO_MATCH',reason,candidateCount:0}),
  findPostJulyMasterCandidates_v660_:()=>[], matchFromMaster_v660_:()=>({status:'MASTER_MATCHED'}), ambiguousMasterMatch_v660_:()=>({status:'AMBIGUOUS'}),
  LOTTEON_V660_MASTER_CUTOFF:'2026-07-01',
  matchFromHistory_v660_:(order,h,master,status,reason)=>({status,reason,company:h.company,cardName:h.cardName,cardNumber:h.cardNumber,cardEnd4:h.cardEnd4,approvalAmount:h.amount,cancelMemo:''}),
  writeVatCardMatchDiagnostic_v660_:()=>{},aggregateVatBusinessCardHalf_v660_:()=>[],prependVatBusinessCardHalfSummary_v660_:()=>{},groupVatDetailByOrder_v660_:()=>[],loadVatCardHistory_v660_:()=>[],loadVatCardMaster_v660_:()=>[],
};
vm.createContext(sandbox); vm.runInContext(fs.readFileSync('Patch_v6_64_vat_card_canonical_cancel_safe.gs','utf8'),sandbox);
function h(row,company,name,number,end4,date,amount,approval,status,cancel=0,evidence='카드이용내역'){
 return {rowNo:row,company,cardName:name,cardNumber:number,cardEnd4:end4,date,time:'10:00',amount,approvalNo:approval,status,cancelDate:cancel?'2026-06-20':'',cancelAmount:cancel,merchantOrderNo:'',evidenceType:evidence,lotteEvidence:true,nonCard:company==='비카드',cancelRow:/^취소$/.test(status)};
}
{
 const rows=[h(1,'롯데카드','LOCA LIKIT','3762','036','2026-06-11',96140,'46693138','정상'),h(2,'롯데카드','카카오페이 카드결제','3762-77**-****-****','','2026-06-11',96140,'46693138','승인',0,'카카오페이 카드')];
 const c=sandbox.canonicalizeVatHistory_v664_(rows,[]);
 assert.strictEqual(c.length,1); assert.strictEqual(c[0].cardEnd4,'036'); assert.strictEqual(c[0].v664EffectiveAmount,96140);
}
{
 const rows=[h(1,'우리카드','EVERY POINT','7680','7680','2026-06-24',117740,'83482661','취소',-117740),h(2,'우리카드','카카오페이 카드결제','4679-14**','','2026-06-24',117740,'83482661','승인(취소있음)',-117740,'카카오페이 카드')];
 const c=sandbox.canonicalizeVatHistory_v664_(rows,[]); assert.strictEqual(c.length,1); assert.strictEqual(c[0].v664FullyCanceled,true); assert.strictEqual(c[0].v664EffectiveAmount,0);
 const r=sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-06-24',purchase:117740,lottePayment:''},c,[],{}); assert.strictEqual(r.status,'NO_MATCH'); assert.ok(r.reason.includes('완전취소'));
}
{
 const rows=[h(1,'롯데카드','LOCA','3762','036','2026-06-12',39840,'49119192','취소있음',-36840)];
 const c=sandbox.canonicalizeVatHistory_v664_(rows,[]); assert.strictEqual(c[0].v664EffectiveAmount,3000);
 let r=sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-06-12',purchase:39840,lottePayment:''},c,[],{}); assert.strictEqual(r.status,'NO_MATCH');
 r=sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-06-12',purchase:3000,lottePayment:''},c,[],{}); assert.strictEqual(r.status,'MATCHED'); assert.ok(r.reason.includes('NET_AFTER_CANCEL'));
}
{
 const c=sandbox.canonicalizeVatHistory_v664_([h(1,'KB국민카드','HERITAGE','5598','4091','2026-06-10',10000,'A1','정상')],[]); const used={};
 let r=sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-06-10',purchase:10000,lottePayment:''},c,[],used); assert.strictEqual(r.status,'MATCHED');
 r=sandbox.matchVatOrderCardCanonical_v664_({orderDate:'2026-06-10',purchase:10000,lottePayment:''},c,[],used); assert.strictEqual(r.status,'NO_MATCH');
}
console.log('v6.64 canonical cancellation-safe tests PASS');
