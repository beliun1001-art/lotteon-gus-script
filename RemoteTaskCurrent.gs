var LOTTEON_REMOTE_TASK={id:'ISSUE67-v1.0-20260814',title:'corrected VAT 사업자별 상반기 신고요약 PREVIEW',enabled:true,statusSheet:'ISSUE67_실행상태'};
var I67_VERSION='v1.0-ISSUE67-CORRECTED-BUSINESS-HALF-VAT-PREVIEW';
var I67_VAT='부가세_신고자료',I67_CARD='부가세_카드매칭검증',I67_CARD_PREVIEW='ISSUE54_카드매칭전체PREVIEW',I67_I66='ISSUE66_반영상태',I67_PERIOD='부가세_기간별',I67_HISTORY='카드사용내역_붙여넣기',I67_MASTER='카드_마스터',I67_BACKUP='ISSUE59_백업_부가세카드매칭검증',I67_OUT='ISSUE67_사업자별부가세PREVIEW';

function runLotteonRemoteTaskStartRemote_(){
  var ss=SpreadsheetApp.getActive();if(!ss)throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var st=i67Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet);i67Status_(st,'RUNNING','PRECHECK','corrected VAT + 운영 카드검증 사업자별 상반기 preview 사전검증',{});
  try{
    [I67_VAT,I67_CARD,I67_CARD_PREVIEW,I67_I66,I67_PERIOD,I67_HISTORY,I67_MASTER,I67_BACKUP].forEach(function(n){i67Need_(ss,n);});
    i67ExpectKv_(ss,I67_I66,{
      '버전':'v1.0-ISSUE66-FILTER-SAFE-CORRECTED-APPLY','상태':'PASS','단계':'DONE','운영주문':1355,
      'MATCHED':808,'NON_CARD':498,'AMBIGUOUS':0,'NO_MATCH':49,'v6.69 2차귀속':1161,'v6.70 3차귀속':81,
      '주문매입금액합계':105762969,'운영/preview_overlap':1355,'material행차이':0,'typed셀차이':0,'display셀차이':0,
      'filter범위':'A1:AC2626','filterCriteria수':1,'롤백':0
    });

    var vat=i67Need_(ss,I67_VAT),card=i67Need_(ss,I67_CARD),cardPreview=i67Need_(ss,I67_CARD_PREVIEW);
    var before=i67Protected_(ss);
    var cardTyped=i67SheetTypedDiff_(card,cardPreview),cardDisplay=i67SheetDisplayDiff_(card,cardPreview);
    if(cardTyped!==0||cardDisplay!==0)throw new Error('운영 카드검증/Issue54 preview 불일치 typed='+cardTyped+' display='+cardDisplay);

    var vatData=i67GroupVat_(vat),cardData=i67LoadCard_(card);
    i67Assert_(vatData.stats,{detailRows:2752,orders:1355,businessBlank:0,sales:138432300,settlement:122495855,purchase:105762969,salesVat:12584695,purchaseVat:9614786,payable:2969909});
    i67Assert_(cardData.stats,{orders:1355,matched:808,nonCard:498,ambiguous:0,noMatch:49,purchase:105762969,dup:0});

    var join=i67Join_(vatData.orders,cardData.map);
    if(join.overlap!==1355||join.vatOnly||join.cardOnly||join.purchaseMismatch||join.businessMismatch)throw new Error('VAT/카드 1:1 결합 실패 '+JSON.stringify(join));

    var summary=i67Aggregate_(vatData.orders),business=i67BusinessTotals_(vatData.orders);
    var recon=i67SummaryRecon_(summary);
    i67Assert_(recon,{orders:1355,sales:138432300,settlement:122495855,purchase:105762969,salesVat:12584695,purchaseVat:9614786,payable:2969909});
    if(Math.round(recon.salesVat-recon.purchaseVat)!==Math.round(recon.payable))throw new Error('매출VAT-매입VAT != 납부예상VAT');

    i67CheckProtected_(ss,before);
    var out=i67Ensure_(ss,I67_OUT);i67WritePreview_(out,summary);
    i67CheckProtected_(ss,before);

    i67Status_(st,'PASS','DONE','corrected VAT 사업자별 상반기 신고요약 PREVIEW 완료',{
      detailRows:vatData.stats.detailRows,orders:vatData.stats.orders,businessCount:business.length,summaryRows:summary.length,
      matched:cardData.stats.matched,nonCard:cardData.stats.nonCard,ambiguous:cardData.stats.ambiguous,noMatch:cardData.stats.noMatch,
      overlap:join.overlap,vatOnly:join.vatOnly,cardOnly:join.cardOnly,purchaseMismatch:join.purchaseMismatch,businessMismatch:join.businessMismatch,
      sales:recon.sales,settlement:recon.settlement,fee:recon.fee,purchase:recon.purchase,salesVat:recon.salesVat,purchaseVat:recon.purchaseVat,payable:recon.payable,profit:recon.profit,vatProfit:recon.vatProfit,
      cardTyped:cardTyped,cardDisplay:cardDisplay,business:business
    });
    return{ok:true,done:true,summaryRows:summary.length,businessCount:business.length};
  }catch(e){var msg=String(e&&e.message?e.message:e);i67Status_(st,'ERROR','FAILED','Issue67 사업자별 VAT preview 실패',{error:msg});throw e;}
}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}

function i67GroupVat_(sh){
  var v=sh.getDataRange().getValues(),h=v[0]||[],ix=function(n){return i67Find_(h,n);};
  var p={year:ix(['신고연도']),half:ix(['반기']),account:ix(['쿠팡계정ID']),business:ix(['사업자등록번호']),order:ix(['주문번호','마켓주문번호','주문ID','주문ID(마켓)']),sales:ix(['순수매출액']),salesSupply:ix(['매출공급가액']),salesVat:ix(['매출부가세']),settlement:ix(['정산기준금액']),fee:ix(['마켓수수료/비용','마켓수수료']),purchase:ix(['매입금액']),purchaseSupply:ix(['매입공급가액']),purchaseVat:ix(['매입부가세']),payable:ix(['납부예상부가세']),profit:ix(['예상이익']),vatProfit:ix(['부가세반영예상이익'])};
  ['year','half','account','business','order','sales','salesSupply','salesVat','settlement','purchase','purchaseSupply','purchaseVat','payable','profit','vatProfit'].forEach(function(k){if(p[k]<0)throw new Error('부가세_신고자료 필수 헤더 누락 '+k);});
  var map={},stats={detailRows:0,orders:0,businessBlank:0,sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};
  for(var r=1;r<v.length;r++){
    var row=v[r],year=i67Text_(row[p.year]),half=i67Text_(row[p.half]);if(year!=='2026'||half!=='상반기')continue;
    stats.detailRows++;
    var account=i67Text_(row[p.account]),business=i67Text_(row[p.business]),orderNo=i67Text_(row[p.order]);
    if(!account||!orderNo)throw new Error('VAT 주문키 공란 R'+(r+1));
    if(!business)stats.businessBlank++;
    var key=i67Key_(account,orderNo);if(!map[key])map[key]={key:key,year:year,half:half,business:business,account:account,orderNo:orderNo,detailRows:0,sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0,cardMatch:null};
    var o=map[key];if(o.business!==business)throw new Error('동일 주문 사업자 불일치 '+key);o.detailRows++;
    var sales=i67Num_(row[p.sales]),sett=i67Num_(row[p.settlement]);
    o.sales+=sales;o.salesSupply+=i67Num_(row[p.salesSupply]);o.salesVat+=i67Num_(row[p.salesVat]);o.settlement+=sett;o.fee+=(p.fee>=0?i67Num_(row[p.fee]):sales-sett);o.purchase+=i67Num_(row[p.purchase]);o.purchaseSupply+=i67Num_(row[p.purchaseSupply]);o.purchaseVat+=i67Num_(row[p.purchaseVat]);o.payable+=i67Num_(row[p.payable]);o.profit+=i67Num_(row[p.profit]);o.vatProfit+=i67Num_(row[p.vatProfit]);
  }
  var orders=Object.keys(map).map(function(k){return map[k];});stats.orders=orders.length;
  orders.forEach(function(o){stats.sales+=o.sales;stats.salesSupply+=o.salesSupply;stats.salesVat+=o.salesVat;stats.settlement+=o.settlement;stats.fee+=o.fee;stats.purchase+=o.purchase;stats.purchaseSupply+=o.purchaseSupply;stats.purchaseVat+=o.purchaseVat;stats.payable+=o.payable;stats.profit+=o.profit;stats.vatProfit+=o.vatProfit;});
  return{orders:orders,stats:stats};
}

function i67LoadCard_(sh){
  var v=sh.getDataRange().getValues(),h=v[0]||[],ix=function(n){return i67Find_(h,n);};
  var p={account:ix(['쿠팡계정ID']),business:ix(['사업자등록번호']),order:ix(['주문번호']),purchase:ix(['주문매입금액','매입금액']),company:ix(['구매카드사']),alias:ix(['구매카드별칭']),cardName:ix(['구매카드명']),cardNumber:ix(['카드번호']),cardEnd4:ix(['카드번호끝4']),status:ix(['카드매칭상태']),reason:ix(['카드매칭근거'])};
  ['account','order','purchase','status'].forEach(function(k){if(p[k]<0)throw new Error('부가세_카드매칭검증 필수 헤더 누락 '+k);});
  var map={},s={orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,purchase:0,dup:0};
  for(var r=1;r<v.length;r++){
    var row=v[r],account=i67Text_(row[p.account]),orderNo=i67Text_(row[p.order]);if(!account&&!orderNo)continue;if(!account||!orderNo)throw new Error('카드검증 주문키 공란 R'+(r+1));
    var key=i67Key_(account,orderNo);if(map[key])s.dup++;
    var status=i67Text_(row[p.status]),o={key:key,account:account,orderNo:orderNo,business:p.business>=0?i67Text_(row[p.business]):'',purchase:i67Num_(row[p.purchase]),company:p.company>=0?i67Text_(row[p.company]):'',alias:p.alias>=0?i67Text_(row[p.alias]):'',cardName:p.cardName>=0?i67Text_(row[p.cardName]):'',cardNumber:p.cardNumber>=0?i67Text_(row[p.cardNumber]):'',cardEnd4:p.cardEnd4>=0?i67Text_(row[p.cardEnd4]):'',status:status,reason:p.reason>=0?i67Text_(row[p.reason]):''};map[key]=o;s.orders++;s.purchase+=o.purchase;
    if(status==='MATCHED'||status==='MASTER_MATCHED')s.matched++;else if(status==='NON_CARD')s.nonCard++;else if(status==='AMBIGUOUS')s.ambiguous++;else s.noMatch++;
  }
  return{map:map,stats:s};
}

function i67Join_(orders,cardMap){
  var o={overlap:0,vatOnly:0,cardOnly:0,purchaseMismatch:0,businessMismatch:0},seen={};
  (orders||[]).forEach(function(x){var c=cardMap[x.key];if(!c){o.vatOnly++;return;}o.overlap++;seen[x.key]=true;if(Math.round(x.purchase)!==Math.round(c.purchase))o.purchaseMismatch++;if(c.business&&x.business&&i67Digits_(c.business)!==i67Digits_(x.business))o.businessMismatch++;x.cardMatch=c;});
  Object.keys(cardMap).forEach(function(k){if(!seen[k])o.cardOnly++;});return o;
}

function i67Aggregate_(orders){
  var map={};
  (orders||[]).forEach(function(o){var m=o.cardMatch||{status:'NO_MATCH',company:'',alias:'',cardName:'',cardNumber:'',cardEnd4:'',reason:'미결합'};var business=o.business||'사업자번호 미매핑';var identity=(m.status==='AMBIGUOUS'||m.status==='NO_MATCH')?m.status:[m.company,m.alias,m.cardName,m.cardNumber,m.cardEnd4,m.status].join('|');var key=[o.year,o.half,business,identity].join('|');
    if(!map[key])map[key]={year:o.year,half:o.half,business:business,accounts:{},company:m.company||'',alias:m.alias||'',cardName:m.cardName||'',cardNumber:m.cardNumber||'',cardEnd4:m.cardEnd4||'',status:m.status,reasons:{},orders:{},sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};
    var x=map[key];x.accounts[o.account]=true;x.orders[o.orderNo]=true;if(m.reason)x.reasons[m.reason]=true;x.sales+=o.sales;x.salesSupply+=o.salesSupply;x.salesVat+=o.salesVat;x.settlement+=o.settlement;x.fee+=o.fee;x.purchase+=o.purchase;x.purchaseSupply+=o.purchaseSupply;x.purchaseVat+=o.purchaseVat;x.payable+=o.payable;x.profit+=o.profit;x.vatProfit+=o.vatProfit;
  });
  return Object.keys(map).map(function(k){var x=map[k],notes=[];if(x.business==='사업자번호 미매핑')notes.push('사업자번호 미매핑');if(x.status==='AMBIGUOUS'||x.status==='NO_MATCH')notes.push('카드 미확정');return [x.year,x.half,x.business,Object.keys(x.accounts).sort().join(', '),x.company,x.alias,x.cardName,x.cardNumber,x.cardEnd4,x.status,Object.keys(x.reasons).sort().join(' / '),Object.keys(x.orders).length,x.sales,x.salesSupply,x.salesVat,x.settlement,x.fee,x.purchase,x.purchaseSupply,x.purchaseVat,x.payable,x.profit,x.vatProfit,notes.join(' / ')];}).sort(function(a,b){return String(a[2]).localeCompare(String(b[2]))||String(a[4]).localeCompare(String(b[4]))||String(a[7]).localeCompare(String(b[7]))||String(a[9]).localeCompare(String(b[9]));});
}

function i67BusinessTotals_(orders){var map={};(orders||[]).forEach(function(o){var k=o.business||'사업자번호 미매핑';if(!map[k])map[k]={business:k,orders:0,matched:0,nonCard:0,ambiguous:0,noMatch:0,sales:0,settlement:0,fee:0,purchase:0,salesVat:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};var x=map[k],s=(o.cardMatch&&o.cardMatch.status)||'NO_MATCH';x.orders++;if(s==='MATCHED'||s==='MASTER_MATCHED')x.matched++;else if(s==='NON_CARD')x.nonCard++;else if(s==='AMBIGUOUS')x.ambiguous++;else x.noMatch++;x.sales+=o.sales;x.settlement+=o.settlement;x.fee+=o.fee;x.purchase+=o.purchase;x.salesVat+=o.salesVat;x.purchaseVat+=o.purchaseVat;x.payable+=o.payable;x.profit+=o.profit;x.vatProfit+=o.vatProfit;});return Object.keys(map).sort().map(function(k){return map[k];});}
function i67SummaryRecon_(rows){var o={orders:0,sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};(rows||[]).forEach(function(r){o.orders+=i67Num_(r[11]);o.sales+=i67Num_(r[12]);o.salesSupply+=i67Num_(r[13]);o.salesVat+=i67Num_(r[14]);o.settlement+=i67Num_(r[15]);o.fee+=i67Num_(r[16]);o.purchase+=i67Num_(r[17]);o.purchaseSupply+=i67Num_(r[18]);o.purchaseVat+=i67Num_(r[19]);o.payable+=i67Num_(r[20]);o.profit+=i67Num_(r[21]);o.vatProfit+=i67Num_(r[22]);});return o;}

function i67WritePreview_(sh,rows){var headers=['신고연도','반기','사업자등록번호','연결 쿠팡계정ID','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','카드매칭상태','카드매칭근거','주문건수','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];sh.clear();sh.getRange(1,1).setValue('Issue67 사업자별 반기 신고요약 PREVIEW (corrected VAT + 운영 카드검증)');sh.getRange(2,1,1,headers.length).setValues([headers]);if(rows.length)sh.getRange(3,1,rows.length,headers.length).setValues(rows);sh.setFrozenRows(2);sh.getRange(1,1,1,headers.length).setFontWeight('bold');sh.getRange(2,1,1,headers.length).setFontWeight('bold');if(rows.length){sh.getRange(3,1,rows.length,11).setNumberFormat('@');for(var c=12;c<=23;c++)sh.getRange(3,c,rows.length,1).setNumberFormat('#,##0');}for(var i=1;i<=headers.length;i++)sh.setColumnWidth(i,(i===11||i===24)?220:((i>=12&&i<=23)?110:135));}

function i67Protected_(ss){var names=[I67_VAT,I67_CARD,I67_CARD_PREVIEW,I67_PERIOD,I67_HISTORY,I67_MASTER,I67_BACKUP],o={};names.forEach(function(n){o[n]=i67Sig_(i67Need_(ss,n));});return o;}
function i67CheckProtected_(ss,b){Object.keys(b).forEach(function(n){if(i67Sig_(i67Need_(ss,n))!==b[n])throw new Error('보호시트 변경 '+n);});}
function i67Sig_(sh){var v=sh.getDataRange().getValues(),h=2166136261;for(var r=0;r<v.length;r++)for(var c=0;c<v[r].length;c++){var s=i67Cell_(v[r][c])+'\u001f';for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619);}}return sh.getLastRow()+'x'+sh.getLastColumn()+'|'+(h>>>0).toString(16);}
function i67SheetTypedDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;var A=a.getDataRange().getValues(),B=b.getDataRange().getValues(),d=0;for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++)if(i67Cell_(A[r][c])!==i67Cell_(B[r][c]))d++;return d;}
function i67SheetDisplayDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;var A=a.getDataRange().getDisplayValues(),B=b.getDataRange().getDisplayValues(),d=0;for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++)if(String(A[r][c])!==String(B[r][c]))d++;return d;}
function i67Cell_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return'D:'+v.toISOString();if(typeof v==='number')return'N:'+String(v);if(typeof v==='boolean')return'B:'+String(v);return'T:'+i67Text_(v);}

function i67Status_(sh,status,stage,msg,x){x=x||{};var rows=[['항목','값'],['버전',I67_VERSION],['상태',status],['단계',stage],['메시지',msg],['VAT상세행',x.detailRows||0],['VAT주문',x.orders||0],['사업자수',x.businessCount||0],['사업자×카드 요약행',x.summaryRows||0],['MATCHED',x.matched||0],['NON_CARD',x.nonCard||0],['AMBIGUOUS',x.ambiguous||0],['NO_MATCH',x.noMatch||0],['VAT/카드_overlap',x.overlap||0],['VAT_only',x.vatOnly||0],['카드_only',x.cardOnly||0],['주문매입금액 불일치',x.purchaseMismatch||0],['사업자 불일치',x.businessMismatch||0],['순수매출액',x.sales||0],['정산기준금액',x.settlement||0],['마켓수수료',x.fee||0],['매입금액',x.purchase||0],['매출부가세',x.salesVat||0],['매입부가세',x.purchaseVat||0],['납부예상부가세',x.payable||0],['예상이익',x.profit||0],['부가세반영예상이익',x.vatProfit||0],['운영카드/Issue54 typed차이',x.cardTyped||0],['운영카드/Issue54 display차이',x.cardDisplay||0],['부가세_신고자료 변경',0],['부가세_카드매칭검증 변경',0],['Issue54preview 변경',0],['부가세_기간별 변경',0],['카드사용내역_붙여넣기 변경',0],['카드_마스터 변경',0],['Issue59백업 변경',0]];
  (x.business||[]).forEach(function(b,i){rows.push(['사업자_'+('0'+(i+1)).slice(-2),i67BusinessLine_(b)]);});
  rows.push(['오류',x.error||''],['완료시각',(status==='PASS'||status==='ERROR')?new Date().toISOString():''],['갱신시각',new Date().toISOString()]);sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setColumnWidth(1,280);sh.setColumnWidth(2,900);}
function i67BusinessLine_(b){return b.business+' | 주문 '+b.orders+' | M '+b.matched+' / NC '+b.nonCard+' / A '+b.ambiguous+' / NM '+b.noMatch+' | 매출 '+i67Fmt_(b.sales)+' | 정산 '+i67Fmt_(b.settlement)+' | 매입 '+i67Fmt_(b.purchase)+' | 매출VAT '+i67Fmt_(b.salesVat)+' | 매입VAT '+i67Fmt_(b.purchaseVat)+' | 납부VAT '+i67Fmt_(b.payable)+' | 이익 '+i67Fmt_(b.profit);}
function i67Fmt_(n){var s=String(Math.round(Number(n)||0)),out='';while(s.length>3){out=','+s.slice(-3)+out;s=s.slice(0,-3);}return s+out;}
function i67ExpectKv_(ss,n,e){var sh=i67Need_(ss,n),kv={};sh.getRange(1,1,sh.getLastRow(),Math.min(2,sh.getLastColumn())).getValues().forEach(function(r){var k=i67Text_(r[0]);if(k)kv[k]=r[1];});Object.keys(e).forEach(function(k){var w=e[k],a=kv[k];if(typeof w==='number'){if(Math.round(i67Num_(a))!==w)throw new Error(n+' '+k+' 불일치 '+a);}else if(i67Text_(a)!==String(w))throw new Error(n+' '+k+' 불일치 '+a+' / 기대 '+w);});}
function i67Assert_(a,e){Object.keys(e).forEach(function(k){if(Math.round(Number(a[k]||0))!==Math.round(Number(e[k]||0)))throw new Error(k+' 불일치 실제 '+a[k]+' 기대 '+e[k]);});}
function i67Need_(ss,n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('필수 시트 없음 '+n);return sh;}function i67Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}function i67Find_(h,names){for(var x=0;x<names.length;x++){var w=i67Compact_(names[x]);for(var i=0;i<h.length;i++)if(i67Compact_(h[i])===w)return i;}return-1;}function i67Text_(v){return String(v==null?'':v).trim();}function i67Compact_(v){return i67Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}function i67Num_(v){var n=Number(typeof v==='number'?v:i67Text_(v).replace(/[,원\s]/g,''));return isNaN(n)?0:n;}function i67Digits_(v){return i67Text_(v).replace(/\D/g,'');}function i67Key_(a,o){a=i67Text_(a).toLowerCase();o=i67Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return a&&o?a+'|'+o:'';}
