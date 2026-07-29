'use strict';
const fs=require('fs'),vm=require('vm'),assert=require('assert');
const sandbox={console};
vm.createContext(sandbox);
vm.runInContext(fs.readFileSync('Patch_v6_60_vat_business_half_summary.gs','utf8'),sandbox);

// 1) Multi-item order must be summed BEFORE card statement amount matching.
const detailHeaders=['날짜','신고연도','반기','신고월','쿠팡계정ID','사업자등록번호','주문번호','고객명','상품번호','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','롯데결제수단'];
const details=[detailHeaders,
 ['06/30','2026','상반기','2026-06','beliun1024','606-45-93763','O-MULTI','x','P1',50000,45455,4545,45000,5000,60000,54545,5455,-910,-15000,-14090,'우리카드'],
 ['06/30','2026','상반기','2026-06','beliun1024','606-45-93763','O-MULTI','x','P2',50000,45455,4545,45000,5000,48660,44236,4424,121,-3660,-3781,'우리카드'],
 ['06/30','2026','상반기','2026-06','beliun1024','606-45-93763','2026063012399247','x','P3',20000,18182,1818,19000,1000,19000,17273,1727,91,0,-91,'카카오페이 머니'],
 ['06/30','2026','상반기','2026-06','beliun1024','606-45-93763','O-AMB','x','P4',10000,9091,909,9000,1000,9900,9000,900,9,-900,-909,'KB국민카드'],
 ['06/30','2026','상반기','2026-06','beliun1024','606-45-93763','O-PREMASTER','x','P5',10000,9091,909,9000,1000,7777,7070,707,202,-1223,-1425,'KB국민카드'],
 ['07/02','2026','하반기','2026-07','beliun1021','227-27-04928','O-JULY','x','P6',20000,18182,1818,18000,2000,15000,13636,1364,454,3000,2546,'KB국민카드']
];
const orders=sandbox.groupVatDetailByOrder_v660_(details);
assert.equal(orders.length,5);
const multi=orders.find(o=>o.orderNo==='O-MULTI');
assert.equal(multi.purchase,108660);
assert.equal(multi.detailRows,2);
assert.equal(multi.orderDate,'2026-06-30');
assert.equal(multi.sales,100000);

function hist(obj){
 const h=Object.assign({company:'',cardName:'',cardNumber:'',cardEnd4:'',date:'',time:'',merchant:'롯데쇼핑(주)',amount:0,approvalNo:'',status:'정상',cancelDate:'',cancelAmount:0,merchantOrderNo:'',evidenceType:'카드이용내역',lotteFlag:'Y',sourceFile:'test',memo:''},obj);
 h.nonCard=sandbox.isNonCardEvidence_v660_(h);h.cancelRow=sandbox.isCancellationHistoryRow_v660_(h);h.lotteEvidence=sandbox.isLotteEvidence_v660_(h);h.amountVariants=sandbox.historyAmountVariants_v660_(h);return h;
}
const history=[
 hist({company:'우리카드',cardName:'카드의정석 EVERY POINT',cardNumber:'7680',cardEnd4:'7680',date:'2026-06-30',time:'18:29:46',amount:108660,approvalNo:'85452686'}),
 hist({company:'비카드',cardName:'카카오페이 페이머니',date:'2026-06-30',time:'21:23:26',amount:19000,merchantOrderNo:'2026063012399247',evidenceType:'카카오페이 거래확인증'}),
 hist({company:'KB국민카드',cardName:'HERITAGE Smart(할인형)',cardNumber:'5598-69**-****-4091',cardEnd4:'4091',date:'2026-06-30',time:'17:01',amount:9900,approvalNo:'A1'}),
 hist({company:'KB국민카드',cardName:'HERITAGE Smart(할인형)',cardNumber:'5598-69**-****-4091',cardEnd4:'4091',date:'2026-06-30',time:'17:02',amount:9900,approvalNo:'A2'}),
 hist({company:'롯데카드',cardName:'LOCA LIKIT 1.2',date:'2026-06-29',time:'21:02',amount:119500,cancelDate:'2026-07-13',cancelAmount:-92424,status:'취소있음'}),
 hist({company:'우리카드',cardName:'카드의정석 EVERY POINT',date:'2026-06-24',time:'07:58',amount:117740,status:'취소',cancelAmount:-117740})
];
const master=[
 {company:'우리카드',alias:'우리4',cardName:'카드의정석 EVERY POINT',cardNumber:'7680',cardEnd4:'7680',account:'1024',business:'606-45-93763',startDate:'2026-07-01',endDate:'',status:'사용 가능'},
 {company:'KB국민카드',alias:'KB4',cardName:'HERITAGE Smart(할인형)',cardNumber:'4091',cardEnd4:'4091',account:'1024',business:'606-45-93763',startDate:'2026-07-01',endDate:'',status:'사용 가능'},
 {company:'KB국민카드',alias:'KB1',cardName:'KB 신규카드',cardNumber:'2121',cardEnd4:'2121',account:'1021',business:'227-27-04928',startDate:'2026-07-01',endDate:'',status:'사용 가능'},
 {company:'KB국민카드',alias:'KB1-old',cardName:'기간미입력',cardNumber:'9999',cardEnd4:'9999',account:'1021',business:'227-27-04928',startDate:'',endDate:'',status:'사용 가능'}
];

// 2) Exact date+order-level amount+issuer matches Woori and enriches from master by physical card.
const m1=sandbox.matchVatOrderCard_v660_(multi,history,master);
assert.equal(m1.status,'MATCHED');
assert.equal(m1.company,'우리카드');
assert.equal(m1.cardEnd4,'7680');
assert.equal(m1.alias,'우리4');
assert(m1.reason.includes('일자+금액+롯데결제수단'));

// 3) Direct merchant order number is strongest and returns NON_CARD for Kakao Pay Money.
const kakaoOrder=orders.find(o=>o.orderNo==='2026063012399247');
const mk=sandbox.matchVatOrderCard_v660_(kakaoOrder,history,master);
assert.equal(mk.status,'NON_CARD');
assert.equal(mk.cardName,'카카오페이 페이머니');
assert(mk.reason.includes('거래내역_주문번호'));

// 4) Same date+amount+issuer multiple candidates must remain AMBIGUOUS.
const amb=orders.find(o=>o.orderNo==='O-AMB');
const ma=sandbox.matchVatOrderCard_v660_(amb,history,master);
assert.equal(ma.status,'AMBIGUOUS');
assert.equal(ma.candidateCount,2);

// 5) Pre-July MUST NOT fall back to business master.
const pre=orders.find(o=>o.orderNo==='O-PREMASTER');
const mp=sandbox.matchVatOrderCard_v660_(pre,history,master);
assert.equal(mp.status,'NO_MATCH');
assert(mp.reason.includes('상반기'));

// 6) Post-July can use exactly one active/applicable master row; blank start date is ignored.
const july=orders.find(o=>o.orderNo==='O-JULY');
const mj=sandbox.matchVatOrderCard_v660_(july,history,master);
assert.equal(mj.status,'MASTER_MATCHED');
assert.equal(mj.alias,'KB1');
assert.equal(mj.cardEnd4,'2121');

// 7) Master fallback becomes ambiguous if two applicable active cards resolve for same business+issuer.
const masterAmb=master.concat([{company:'KB국민카드',alias:'KB1-B',cardName:'second',cardNumber:'3131',cardEnd4:'3131',account:'1021',business:'227-27-04928',startDate:'2026-07-01',endDate:'',status:'사용 가능'}]);
const mj2=sandbox.matchVatOrderCard_v660_(july,history,masterAmb);
assert.equal(mj2.status,'AMBIGUOUS');

// 8) Cancellation rows are not normal candidates; '취소있음' original approval remains candidate and exposes net variant.
assert.equal(history[5].cancelRow,true);
assert.equal(history[4].cancelRow,false);
assert(sandbox.historyMatchesAmount_v660_(history[4],27076));
assert.equal(sandbox.historyAmountKind_v660_(history[4],27076),'NET_AFTER_CANCEL');

// 9) Summary totals must equal order-level totals even when card matching is unresolved.
orders.forEach(o=>{o.cardMatch=sandbox.matchVatOrderCard_v660_(o,history,master);});
const summary=sandbox.aggregateVatBusinessCardHalf_v660_(orders);
const sum=(arr,idx)=>arr.reduce((n,r)=>n+Number(r[idx]||0),0);
assert.equal(sum(summary,12),orders.reduce((n,o)=>n+o.sales,0));
assert.equal(sum(summary,17),orders.reduce((n,o)=>n+o.purchase,0));
assert.equal(sum(summary,19),orders.reduce((n,o)=>n+o.purchaseVat,0));
assert(summary.some(r=>r[9]==='AMBIGUOUS'));
assert(summary.some(r=>r[9]==='NON_CARD'));

// 10) Historical statement-specific identifiers enrich without using business as the identifier.
const shinHist=hist({company:'신한카드',cardName:'본인038*',cardNumber:'본인038*',date:'2026-06-28',amount:3000});
const shinMaster=[{company:'신한카드',alias:'신한4',cardName:'Deep Dream[딥 드림]',cardNumber:'038*',cardEnd4:'',account:'1024',business:'606-45-93763',startDate:'2026-07-01',status:'사용 가능'}];
const enriched=sandbox.enrichHistoryFromMaster_v660_(shinHist,shinMaster);
assert.equal(enriched.alias,'신한4');
assert.equal(enriched.cardName,'Deep Dream[딥 드림]');

console.log('v6.60 purchase-card reconciliation tests PASS');
