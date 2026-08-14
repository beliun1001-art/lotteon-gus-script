var LOTTEON_REMOTE_TASK={id:'ISSUE78-v1.0-20260814',title:'사업자별 경정청구 입력 준비표 + 추가증빙 체크리스트',enabled:true,statusSheet:'ISSUE78_실행상태'};
var I78V='v1.0-ISSUE78-FILING-PREP-HOMETAX-EVIDENCE-GAP';
var I78OUT='ISSUE78_경정청구입력준비',I78CHK='ISSUE78_추가증빙체크리스트',I78S='ISSUE78_실행상태';
var I78CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
function runLotteonRemoteTaskStartRemote_(){return i78run_();}
function runLotteonRemoteTaskContinueRemote_(){return i78run_();}

function i78run_(){
  var ss=SpreadsheetApp.getActive(),before={};
  try{
    i78status_(ss,[['version',I78V],['상태','RUNNING'],['단계','GUARD'],['메시지','사업자별 경정청구 입력 준비표 및 홈택스 추가증빙 체크리스트 생성 중']]);
    I78CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락: '+n);before[n]=i78sig_(sh);});
    i78guardStatus_(ss);
    var vat=i78vatBiz_(ss.getSheetByName('부가세_신고자료'));
    var b77=i78issue77Biz_(ss.getSheetByName('ISSUE77_사업자별증빙요약'));
    var joined=i78join_(vat,b77);
    i78guardJoined_(joined);
    var rows=i78prepRows_(joined);
    i78writePrep_(ss,rows);
    var checklist=i78checklist_(ss,joined);
    i78writeChecklist_(ss,checklist);
    var changed=[];I78CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh||i78sig_(sh)!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반: '+changed.join(', '));
    var agg=i78agg_(rows);
    var st=[
      ['version',I78V],['상태','PASS'],['단계','DONE'],['메시지','사업자 4개 경정청구 입력 준비표 및 홈택스 추가증빙 체크리스트 생성 완료'],
      ['사업자수',rows.length],['주문수',agg.orders],['순수매출액',agg.sales],['매출부가세',agg.salesVat],['총매입',agg.purchase],['총스크립트매입부가세',agg.purchaseVat],
      ['카드증빙연결건수',agg.matchedCount],['카드증빙연결_매입',agg.matchedPurchase],['카드증빙연결_스크립트매입부가세',agg.matchedVat],
      ['현금영수증명시확정건수',agg.cashCount],['현금영수증명시확정_매입',agg.cashPurchase],['현금영수증명시확정_스크립트매입부가세',agg.cashVat],
      ['카카오머니_현금영수증미확인건수',agg.kakaoCount],['카카오머니_미확인_매입',agg.kakaoPurchase],['카카오머니_미확인_스크립트매입부가세',agg.kakaoVat],
      ['NO_MATCH건수',agg.noMatchCount],['NO_MATCH_매입',agg.noMatchPurchase],['NO_MATCH_스크립트매입부가세',agg.noMatchVat],
      ['현재증빙연결분_스크립트매입부가세',agg.connectedVat],['추가증빙필요_스크립트매입부가세',agg.pendingVat],
      ['보수적_임시납부세액',agg.conservativePayable],['전체스크립트_계산납부세액',agg.scriptPayable],['증빙확인차이',agg.gap],
      ['추가증빙체크행',checklist.length],['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
    ];
    i78status_(ss,st);try{ss.toast('Issue78 PASS: 사업자별 입력 준비표 생성 / 추가증빙 VAT '+agg.pendingVat.toLocaleString()+'원','LOTTEON',10);}catch(_e){}
    return{ok:true,done:true,version:I78V,businesses:rows.length,pendingVat:agg.pendingVat};
  }catch(e){var m=String(e&&e.message?e.message:e);try{i78status_(ss,[['version',I78V],['상태','ERROR'],['단계','FAILED'],['메시지','Issue78 입력 준비표 생성 실패'],['오류',m],['완료시각',new Date().toISOString()]]);}catch(_e2){}throw e;}
}

function i78guardStatus_(ss){
  var s77=ss.getSheetByName('ISSUE77_실행상태'),s76=ss.getSheetByName('ISSUE76_최종검증상태');if(!s77||!s76)throw new Error('Issue76/77 상태시트 누락');
  var a=i78kv_(s77),b=i78kv_(s76);
  if(i78t_(a['상태'])!=='PASS'||i78t_(a['단계'])!=='DONE'||Number(a['운영주문'])!==1355||Number(a['MATCHED'])!==835||Number(a['NON_CARD'])!==498||Number(a['NO_MATCH'])!==22||Number(a['NONCARD_현금영수증명시확정'])!==0||Number(a['NONCARD_카카오머니_현금영수증미확인'])!==498||Number(a['핵심시트변경수'])!==0)throw new Error('Issue77 PASS exact guard 실패');
  if(i78t_(b['상태'])!=='PASS'||i78t_(b['단계'])!=='DONE'||Number(b['VAT_주문수'])!==1355||Number(b['MATCHED'])!==835||Number(b['NON_CARD'])!==498||Number(b['NO_MATCH'])!==22||Number(b['핵심시트변경수'])!==0)throw new Error('Issue76 PASS exact guard 실패');
}
function i78vatBiz_(sh){
  var v=sh.getDataRange().getValues(),h=v[0]||[],x={year:i78ix_(h,['신고연도']),half:i78ix_(h,['반기']),business:i78ix_(h,['사업자등록번호']),account:i78ix_(h,['쿠팡계정ID']),order:i78ix_(h,['주문번호']),sales:i78ix_(h,['순수매출액']),sv:i78ix_(h,['매출부가세']),purchase:i78ix_(h,['매입금액']),pv:i78ix_(h,['매입부가세'])};Object.keys(x).forEach(function(k){if(x[k]<0)throw new Error('VAT header 누락 '+k);});
  var m={},seen={};for(var r=1;r<v.length;r++){var z=v[r];if(i78t_(z[x.year])!=='2026'||i78t_(z[x.half])!=='상반기')continue;var no=i78t_(z[x.order]);if(!no)continue;var biz=i78t_(z[x.business]);if(!m[biz])m[biz]={business:biz,accounts:{},orders:{},sales:0,salesVat:0,purchase:0,purchaseVat:0};var q=m[biz];q.accounts[i78t_(z[x.account])]=1;q.orders[no]=1;q.sales+=i78n_(z[x.sales]);q.salesVat+=i78n_(z[x.sv]);q.purchase+=i78n_(z[x.purchase]);q.purchaseVat+=i78n_(z[x.pv]);seen[biz+'|'+i78t_(z[x.account])+'|'+no]=1;}
  return m;
}
function i78issue77Biz_(sh){
  if(!sh)throw new Error('ISSUE77_사업자별증빙요약 누락');var v=sh.getDataRange().getValues(),h=v[0]||[],x={business:i78ix_(h,['사업자등록번호']),total:i78ix_(h,['총주문수']),mc:i78ix_(h,['카드매칭확정건수']),mp:i78ix_(h,['카드매칭확정_매입']),mv:i78ix_(h,['카드매칭확정_스크립트매입부가세']),cc:i78ix_(h,['현금영수증명시확정건수']),cp:i78ix_(h,['현금영수증명시확정_매입']),cv:i78ix_(h,['현금영수증명시확정_스크립트매입부가세']),kc:i78ix_(h,['카카오머니_현금영수증미확인건수']),kp:i78ix_(h,['카카오머니_미확인_매입']),kv:i78ix_(h,['카카오머니_미확인_스크립트매입부가세']),nc:i78ix_(h,['NO_MATCH건수']),np:i78ix_(h,['NO_MATCH_매입']),nv:i78ix_(h,['NO_MATCH_스크립트매입부가세']),tp:i78ix_(h,['총매입']),tv:i78ix_(h,['총스크립트매입부가세'])};Object.keys(x).forEach(function(k){if(x[k]<0)throw new Error('Issue77 사업자요약 header 누락 '+k);});
  var m={};for(var r=1;r<v.length;r++){var z=v[r],biz=i78t_(z[x.business]);if(!biz)continue;if(m[biz])throw new Error('Issue77 사업자 중복 '+biz);m[biz]={business:biz,total:i78n_(z[x.total]),matchedCount:i78n_(z[x.mc]),matchedPurchase:i78n_(z[x.mp]),matchedVat:i78n_(z[x.mv]),cashCount:i78n_(z[x.cc]),cashPurchase:i78n_(z[x.cp]),cashVat:i78n_(z[x.cv]),kakaoCount:i78n_(z[x.kc]),kakaoPurchase:i78n_(z[x.kp]),kakaoVat:i78n_(z[x.kv]),noMatchCount:i78n_(z[x.nc]),noMatchPurchase:i78n_(z[x.np]),noMatchVat:i78n_(z[x.nv]),totalPurchase:i78n_(z[x.tp]),totalVat:i78n_(z[x.tv])};}return m;
}
function i78join_(vat,b77){var out={};Object.keys(vat).forEach(function(b){if(!b77[b])throw new Error('Issue77 사업자요약 누락 '+b);out[b]={vat:vat[b],e:b77[b]};});Object.keys(b77).forEach(function(b){if(!vat[b])throw new Error('VAT 사업자 누락 '+b);});return out;}
function i78guardJoined_(m){var ks=Object.keys(m).sort();if(ks.length!==4)throw new Error('사업자수 '+ks.length+' (기대4)');var t={orders:0,sales:0,sv:0,p:0,pv:0,mc:0,kc:0,nc:0};ks.forEach(function(k){var q=m[k],v=q.vat,e=q.e,orders=Object.keys(v.orders).length;if(orders!==e.total||v.purchase!==e.totalPurchase||v.purchaseVat!==e.totalVat)throw new Error('사업자 요약 불일치 '+k);if(e.cashCount!==0||e.cashPurchase!==0||e.cashVat!==0)throw new Error('현금영수증 명시확정이 0이 아님 '+k);t.orders+=orders;t.sales+=v.sales;t.sv+=v.salesVat;t.p+=v.purchase;t.pv+=v.purchaseVat;t.mc+=e.matchedCount;t.kc+=e.kakaoCount;t.nc+=e.noMatchCount;});if(t.orders!==1355||t.sales!==138432300||t.sv!==12584695||t.p!==105762969||t.pv!==9614786||t.mc!==835||t.kc!==498||t.nc!==22)throw new Error('전체 guard 실패 '+JSON.stringify(t));}
function i78prepRows_(m){return Object.keys(m).sort().map(function(k){var q=m[k],v=q.vat,e=q.e,orders=Object.keys(v.orders).length,connected=e.matchedVat+e.cashVat,pending=e.kakaoVat+e.noMatchVat,conservative=v.salesVat-connected,script=v.salesVat-v.purchaseVat,gap=conservative-script;return['2026','상반기',k,Object.keys(v.accounts).sort().join(', '),orders,v.sales,v.salesVat,v.purchase,v.purchaseVat,e.matchedCount,e.matchedPurchase,e.matchedVat,e.cashCount,e.cashPurchase,e.cashVat,e.kakaoCount,e.kakaoPurchase,e.kakaoVat,e.noMatchCount,e.noMatchPurchase,e.noMatchVat,connected,pending,conservative,script,gap,'추가증빙확인필요: 홈택스 카드 공제여부 + 지출증빙용 현금영수증 + NO_MATCH'];});}
function i78writePrep_(ss,rows){var sh=ss.getSheetByName(I78OUT)||ss.insertSheet(I78OUT);sh.clear();var h=['신고연도','반기','사업자등록번호','연결쿠팡계정ID','주문수','순수매출액','매출부가세','총매입','총스크립트매입부가세','카드증빙연결건수','카드증빙연결_매입','카드증빙연결_스크립트매입부가세','현금영수증명시확정건수','현금영수증명시확정_매입','현금영수증명시확정_스크립트매입부가세','카카오머니_현금영수증미확인건수','카카오머니_미확인_매입','카카오머니_미확인_스크립트매입부가세','NO_MATCH건수','NO_MATCH_매입','NO_MATCH_스크립트매입부가세','현재증빙연결분_스크립트매입부가세','추가증빙필요_스크립트매입부가세','보수적_임시납부세액','전체스크립트_계산납부세액','증빙확인차이','신고준비상태'];sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length)sh.getRange(2,1,rows.length,h.length).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');if(rows.length){sh.getRange(2,1,rows.length,4).setNumberFormat('@');sh.getRange(2,5,rows.length,h.length-5).setNumberFormat('#,##0');sh.getRange(2,27,rows.length,1).setNumberFormat('@');}try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}}
function i78checklist_(ss,m){var rows=[],priority=1;Object.keys(m).sort().forEach(function(k){var e=m[k].e;rows.push([priority++,k,'사업용신용카드 공제여부 확인','','',e.matchedCount,e.matchedPurchase,e.matchedVat,'카드증빙 내부연결 완료','홈택스 사업용신용카드 사용내역에서 매입세액 공제 확인/변경 상태를 확인하고 공제대상 원본을 확보','MATCHED','공제/불공제 최종확정 전']);rows.push([priority++,k,'현금영수증 지출증빙 확인','','',e.kakaoCount,e.kakaoPurchase,e.kakaoVat,'카카오페이머니/계좌 결제 확인, 현금영수증 발급 미확인','홈택스 현금영수증 매입내역에서 사업자등록번호로 발급된 지출증빙용 현금영수증 원본을 확보','NON_CARD','카카오머니 자체는 현금영수증 확정 아님']);});var sh=ss.getSheetByName('ISSUE77_경정청구증빙진단');if(!sh)throw new Error('ISSUE77_경정청구증빙진단 누락');var v=sh.getDataRange().getValues(),h=v[0]||[],x={kind:i78ix_(h,['구분']),biz:i78ix_(h,['사업자등록번호']),date:i78ix_(h,['주문일']),order:i78ix_(h,['주문번호']),purchase:i78ix_(h,['주문매입금액']),vat:i78ix_(h,['스크립트매입부가세']),cause:i78ix_(h,['Issue72원인']),verdict:i78ix_(h,['Issue73판정'])};Object.keys(x).forEach(function(k){if(x[k]<0)throw new Error('Issue77 진단 header 누락 '+k);});var n=0;for(var r=1;r<v.length;r++){var z=v[r];if(i78t_(z[x.kind])!=='NO_MATCH')continue;n++;var ver=i78t_(z[x.verdict]),act=ver==='MULTI_CANDIDATE'?'복수 후보 중 실제 승인증빙 수동확정':ver==='USED_BY_OTHER'?'동일 증빙이 다른 주문에 사용된 충돌 해소':ver==='AMOUNT_REVIEW'?'금액차이 원인(쿠폰/부분취소/분할결제) 원본확인':ver==='BLOCKED_ZERO'?'매입금액 0 원천 확인; 공제대상 금액 없음 여부 확인':ver==='BLOCKED_CANCELED'?'완전취소 여부 및 대체 승인증빙 존재 확인':'추가 원본증빙 확인';rows.push([priority++,i78t_(z[x.biz]),'NO_MATCH 추가확인',i78t_(z[x.date]),i78t_(z[x.order]),1,i78n_(z[x.purchase]),i78n_(z[x.vat]),'미확정',act,i78t_(z[x.cause])+' / '+ver,'확정 전 경정청구 공제액에서 격리']);}if(n!==22)throw new Error('NO_MATCH 체크리스트 '+n+' (기대22)');return rows;}
function i78writeChecklist_(ss,rows){var sh=ss.getSheetByName(I78CHK)||ss.insertSheet(I78CHK);sh.clear();var h=['우선순위','사업자등록번호','대상구분','주문일','주문번호','건수','매입금액','스크립트매입부가세','현재상태','필요한확인','원인/판정','비고'];sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length)sh.getRange(2,1,rows.length,h.length).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');if(rows.length){sh.getRange(2,1,rows.length,1).setNumberFormat('0');sh.getRange(2,2,rows.length,4).setNumberFormat('@');sh.getRange(2,6,rows.length,3).setNumberFormat('#,##0');}try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}}
function i78agg_(rows){var t={orders:0,sales:0,salesVat:0,purchase:0,purchaseVat:0,matchedCount:0,matchedPurchase:0,matchedVat:0,cashCount:0,cashPurchase:0,cashVat:0,kakaoCount:0,kakaoPurchase:0,kakaoVat:0,noMatchCount:0,noMatchPurchase:0,noMatchVat:0,connectedVat:0,pendingVat:0,conservativePayable:0,scriptPayable:0,gap:0};rows.forEach(function(r){t.orders+=i78n_(r[4]);t.sales+=i78n_(r[5]);t.salesVat+=i78n_(r[6]);t.purchase+=i78n_(r[7]);t.purchaseVat+=i78n_(r[8]);t.matchedCount+=i78n_(r[9]);t.matchedPurchase+=i78n_(r[10]);t.matchedVat+=i78n_(r[11]);t.cashCount+=i78n_(r[12]);t.cashPurchase+=i78n_(r[13]);t.cashVat+=i78n_(r[14]);t.kakaoCount+=i78n_(r[15]);t.kakaoPurchase+=i78n_(r[16]);t.kakaoVat+=i78n_(r[17]);t.noMatchCount+=i78n_(r[18]);t.noMatchPurchase+=i78n_(r[19]);t.noMatchVat+=i78n_(r[20]);t.connectedVat+=i78n_(r[21]);t.pendingVat+=i78n_(r[22]);t.conservativePayable+=i78n_(r[23]);t.scriptPayable+=i78n_(r[24]);t.gap+=i78n_(r[25]);});if(t.orders!==1355||t.sales!==138432300||t.salesVat!==12584695||t.purchase!==105762969||t.purchaseVat!==9614786||t.matchedCount!==835||t.cashCount!==0||t.kakaoCount!==498||t.noMatchCount!==22||t.scriptPayable!==2969909||t.gap!==t.pendingVat)throw new Error('Issue78 aggregate guard '+JSON.stringify(t));return t;}
function i78status_(ss,p){var sh=ss.getSheetByName(I78S)||ss.insertSheet(I78S);sh.clearContents();sh.getRange(1,1,1,2).setValues([['항목','값']]);if(p&&p.length)sh.getRange(2,1,p.length,2).setValues(p);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}
function i78kv_(sh){var v=sh.getDataRange().getValues(),m={};for(var r=1;r<v.length;r++)m[i78t_(v[r][0])]=v[r][1];return m;}
function i78ix_(h,n){var m={};(h||[]).forEach(function(v,i){m[i78c_(v)]=i;});for(var j=0;j<n.length;j++){var k=i78c_(n[j]);if(Object.prototype.hasOwnProperty.call(m,k))return m[k];}return-1;}
function i78t_(v){if(v===null||v===undefined)return'';if(Object.prototype.toString.call(v)==='[object Date]')return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');return String(v).trim();}
function i78c_(v){return i78t_(v).toLowerCase().replace(/[\s_\-\/.()\[\]:]+/g,'');}
function i78n_(v){if(typeof v==='number')return isNaN(v)?0:Math.round(v);var n=Number(String(v==null?'0':v).replace(/[원,\s]/g,''));return isNaN(n)?0:Math.round(n);}
function i78sig_(sh){var v=sh.getDataRange().getDisplayValues(),h=2166136261;function add(s){s=String(s);for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619);}}for(var r=0;r<v.length;r++)for(var c=0;c<v[r].length;c++){add(v[r][c]);add('|');}return String(h>>>0);}
