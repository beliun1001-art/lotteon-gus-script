var LOTTEON_REMOTE_TASK={id:'ISSUE79-V52-CARD-HISTORY-COVERAGE',title:'Issue79 카드원본 coverage/메타데이터 진단 READ-ONLY',enabled:true,statusSheet:'ISSUE79_NOMATCH4_원본Coverage상태'};
var V='v5.2-ISSUE79-CARD-HISTORY-COVERAGE-METADATA-READONLY';
var CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
var TARGETS=[
 {id:'12100196902987',date:'2026-06-14',amt:179640},
 {id:'12100197310918',date:'2026-06-16',amt:100620},
 {id:'5100198099688',date:'2026-06-19',amt:179640}
];
var ZERO='30100189414967';
function runLotteonRemoteTaskStartRemote_(){return run_();}
function runLotteonRemoteTaskContinueRemote_(){return run_();}
function run_(){
 var ss=SpreadsheetApp.getActive(),before={},started=new Date().toISOString();
 try{
  status_(ss,[['version',V],['상태','RUNNING'],['단계','COVERAGE'],['메시지','카드원본 전체 coverage/헤더/끝4/금액/과거진단 원문 점검 중'],['실행시작',started]]);
  CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락 '+n);before[n]=sig_(sh);});
  var card=card_(ss.getSheetByName('부가세_카드매칭검증'));
  if(card.rows.length!==1355||(card.s.MATCHED||0)!==842||(card.s.NON_CARD||0)!==509||(card.s.NO_MATCH||0)!==4||card.sum!==105314779)throw new Error('CARD baseline '+JSON.stringify({n:card.rows.length,s:card.s,p:card.sum}));
  var nm=card.rows.filter(function(x){return x.st==='NO_MATCH';}).map(function(x){return x.id;}).sort();
  var exp=TARGETS.map(function(x){return x.id;}).concat([ZERO]).sort();
  if(nm.join('|')!==exp.join('|'))throw new Error('NO_MATCH4 set '+JSON.stringify(nm));
  var hist=history_(ss.getSheetByName('카드사용내역_붙여넣기'));
  var master=scanMaster_(ss.getSheetByName('카드_마스터'));
  var prior=scanPrior_(ss,TARGETS.map(function(x){return x.id;}).concat([ZERO]));
  var detail=[],summary=[];
  TARGETS.forEach(function(t){
   var w30=hist.rows.filter(function(r){return Math.abs(days_(t.date,r.date))<=30;});
   var exactAll=hist.rows.filter(function(r){return r.amount===t.amt;});
   var exact30=w30.filter(function(r){return r.amount===t.amt;});
   var lotte30=w30.filter(function(r){return r.isLotte||r.isLoca;});
   var e430=w30.filter(function(r){return r.e4==='0036';});
   var loca036_30=w30.filter(function(r){return (r.isLotte||r.isLoca)&&r.e4==='0036';});
   var direct=hist.rows.filter(function(r){return r.rowText.indexOf(t.id)>=0;});
   exact30.forEach(function(r){detail.push(rowOut_(t,'EXACT30_ANY',r));});
   lotte30.forEach(function(r){detail.push(rowOut_(t,'LOTTE_OR_LOCA_30',r));});
   e430.forEach(function(r){detail.push(rowOut_(t,'END4_0036_30',r));});
   direct.forEach(function(r){detail.push(rowOut_(t,'DIRECT_ORDER_ID',r));});
   summary.push([t.id,t.date,t.amt,exactAll.length,exact30.length,lotte30.length,e430.length,loca036_30.length,direct.length,prior.byId[t.id]||0]);
  });
  var allDates=hist.rows.map(function(r){return r.date;}).filter(Boolean).sort();
  var lotte=hist.rows.filter(function(r){return r.isLotte;});
  var loca=hist.rows.filter(function(r){return r.isLoca;});
  var e4=hist.rows.filter(function(r){return r.e4==='0036';});
  var loca036=hist.rows.filter(function(r){return (r.isLotte||r.isLoca)&&r.e4==='0036';});
  var a179=hist.rows.filter(function(r){return r.amount===179640;});
  var a100=hist.rows.filter(function(r){return r.amount===100620;});
  var cls='UNRESOLVED';
  if(!loca036.length&&!a179.length&&!a100.length)cls='TARGET_CARD_OR_TRANSACTIONS_MISSING_FROM_HISTORY';
  else if(!loca036.length&&(lotte.length||loca.length))cls='CARD_END4_OR_CARDNAME_METADATA_GAP';
  else if(loca036.length)cls='LOCA036_PRESENT_BUT_TARGET_MATCH_MISSING';
  write_(ss,summary,detail,prior.rows,master.rows,hist.headers);
  var changed=[];CORE.forEach(function(n){if(sig_(ss.getSheetByName(n))!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반 '+changed.join(','));
  status_(ss,[['version',V],['상태','PASS'],['단계','DONE'],['메시지','카드원본 coverage/메타데이터/금액/과거진단 원문 점검 완료'],
   ['현재_CARD주문',card.rows.length],['현재_MATCHED',card.s.MATCHED||0],['현재_NON_CARD',card.s.NON_CARD||0],['현재_NO_MATCH',card.s.NO_MATCH||0],['현재_CARD매입합',card.sum],
   ['카드원본_헤더행',hist.headerRow],['카드원본_데이터행',hist.rows.length],['카드원본_최초일',allDates[0]||''],['카드원본_최종일',allDates.length?allDates[allDates.length-1]:'' ],
   ['카드원본_롯데계열행',lotte.length],['카드원본_LOCA_LIKIT텍스트행',loca.length],['카드원본_끝4_0036행',e4.length],['카드원본_롯데LOCA+0036행',loca036.length],
   ['카드원본_179640행',a179.length],['카드원본_100620행',a100.length],
   ['카드마스터_롯데LOCA관련행',master.rows.length],['과거진단_TARGET원문hit',prior.rows.length],['진단분류',cls],['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]]);
  return{ok:true,done:true,version:V,classification:cls,loca036:loca036.length,priorHits:prior.rows.length};
 }catch(e){var m=String(e&&e.stack?e.stack:e);try{status_(ss,[['version',V],['상태','ERROR'],['단계','FAILED'],['메시지','카드원본 coverage 진단 실패'],['오류',m],['완료시각',new Date().toISOString()]])}catch(_e){}throw e;}
}
function card_(sh){var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=hdr_(v,['주문번호','카드매칭상태']),h=v[hr],x={y:ix_(h,['신고연도']),hf:ix_(h,['반기']),id:ix_(h,['주문번호']),pur:ix_(h,['주문매입금액','매입금액']),st:ix_(h,['카드매칭상태'])},a=[],s={},sum=0;for(var r=hr+1;r<v.length;r++){if(t_(v[r][x.y])!=='2026'||t_(v[r][x.hf])!=='상반기')continue;var id=id_(d[r][x.id]);if(!id)continue;var st=t_(v[r][x.st]).toUpperCase(),p=n_(v[r][x.pur]);a.push({id:id,st:st,pur:p});s[st]=(s[st]||0)+1;sum+=p}return{rows:a,s:s,sum:Math.round(sum)}}
function history_(sh){
 var range=sh.getDataRange(),v=range.getValues(),d=range.getDisplayValues(),hr=detectHistoryHeader_(v),h=v[hr]||[];
 var x={co:ix_(h,['카드사','발급사','카드회사']),name:ix_(h,['카드명','카드상품명','상품명']),num:ix_(h,['카드번호','이용카드번호']),e4:ix_(h,['카드번호끝4','끝4','카드끝4','카드번호뒤4자리']),date:ix_(h,['승인일','이용일','거래일','승인일자','이용일자']),mer:ix_(h,['가맹점명','이용가맹점','가맹점']),amt:ix_(h,['승인금액','이용금액','거래금액','금액']),app:ix_(h,['승인번호']),state:ix_(h,['승인상태','승인/취소구분','상태']),src:ix_(h,['원본파일','파일명','출처']),mo:ix_(h,['가맹점주문번호','주문번호'])};
 var rows=[];for(var r=hr+1;r<v.length;r++){var rowDisp=d[r]||[],txt=rowDisp.join(' | '),date=x.date>=0?date_(v[r][x.date]):findDate_(txt),amt=x.amt>=0?n_(v[r][x.amt]):findMoney_(rowDisp),co=x.co>=0?t_(v[r][x.co]):'',name=x.name>=0?t_(v[r][x.name]):'',num=x.num>=0?t_(rowDisp[x.num]):'',e4v=x.e4>=0?t_(rowDisp[x.e4]):'',e4=normE4_(e4v||num),mer=x.mer>=0?t_(v[r][x.mer]):'',app=x.app>=0?id_(rowDisp[x.app]):'',state=x.state>=0?t_(v[r][x.state]):'',src=x.src>=0?t_(v[r][x.src]):'',mo=x.mo>=0?id_(rowDisp[x.mo]):'';if(!date&&!amt&&!co&&!name&&!mer&&!app&&!mo)continue;var text=c_([co,name,num,e4v,mer,src,txt].join(' '));rows.push({row:r+1,date:date,amount:amt,company:co,cardName:name,cardNumber:num,e4:e4,merchant:mer,approval:app,state:state,source:src,merchantOrder:mo,rowText:txt,isLotte:/롯데|lotte/.test(text),isLoca:/loca|likit/.test(text)});}return{rows:rows,headerRow:hr+1,headers:h.map(t_).join(' | ')};
}
function detectHistoryHeader_(v){for(var r=0;r<Math.min(30,v.length);r++){var h=v[r]||[],hasDate=ix_(h,['승인일','이용일','거래일','승인일자','이용일자'])>=0,hasAmt=ix_(h,['승인금액','이용금액','거래금액','금액'])>=0;if(hasDate&&hasAmt)return r;}return 0;}
function scanMaster_(sh){var d=sh.getDataRange().getDisplayValues(),out=[];for(var r=0;r<d.length;r++){var txt=(d[r]||[]).join(' | '),c=c_(txt);if(/롯데|lotte|loca|likit/.test(c)||normE4_(txt)==='0036')out.push([r+1,txt.slice(0,4000)]);}return{rows:out};}
function scanPrior_(ss,ids){var names=['ISSUE54_카드매칭전체PREVIEW','ISSUE55_카드매칭차이진단','ISSUE56_상태변경12건판정','ISSUE57_차단9건심층진단','ISSUE58_4건원인분리','ISSUE74_카드회수PREVIEW'],rows=[],byId={};names.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)return;var d=sh.getDataRange().getDisplayValues();for(var r=0;r<d.length;r++){var txt=(d[r]||[]).join(' | ');ids.forEach(function(id){if(txt.indexOf(id)>=0){rows.push([n,r+1,id,txt.slice(0,6000)]);byId[id]=(byId[id]||0)+1;}});}});return{rows:rows,byId:byId};}
function rowOut_(t,kind,r){return[t.id,t.date,t.amt,kind,r.row,r.date,r.amount,r.company,r.cardName,r.e4,r.approval,r.merchant,r.state,r.merchantOrder,r.source,r.rowText.slice(0,3000)];}
function write_(ss,sum,detail,prior,master,headers){
 var s=ss.getSheetByName('ISSUE79_NOMATCH4_원본Coverage상세')||ss.insertSheet('ISSUE79_NOMATCH4_원본Coverage상세');s.clear();var h=['주문번호','주문일','목표금액','전기간exact','±30일exact','±30일롯데/LOCA','±30일끝4_0036','±30일LOCA036','카드원본직접IDhit','과거진단hit'];s.getRange(1,1,1,h.length).setValues([h]);if(sum.length)s.getRange(2,1,sum.length,h.length).setValues(sum);s.getRange(7,1).setValue('카드원본 감지 헤더');s.getRange(7,2).setValue(headers);s.setFrozenRows(1);s.getRange(1,1,1,h.length).setFontWeight('bold');
 var d=ss.getSheetByName('ISSUE79_NOMATCH4_원본후보')||ss.insertSheet('ISSUE79_NOMATCH4_원본후보');d.clear();var h2=['주문번호','주문일','목표금액','후보종류','원본행','승인/이용일','금액','카드사','카드명','끝4','승인번호','가맹점','상태','가맹점주문번호','원본파일','원본행텍스트'];d.getRange(1,1,1,h2.length).setValues([h2]);if(detail.length)d.getRange(2,1,detail.length,h2.length).setValues(detail);d.setFrozenRows(1);d.getRange(1,1,1,h2.length).setFontWeight('bold');
 var p=ss.getSheetByName('ISSUE79_NOMATCH4_과거증빙원문')||ss.insertSheet('ISSUE79_NOMATCH4_과거증빙원문');p.clear();p.getRange(1,1,1,4).setValues([['시트','행','주문번호','원문']]);if(prior.length)p.getRange(2,1,prior.length,4).setValues(prior);p.setFrozenRows(1);p.getRange(1,1,1,4).setFontWeight('bold');
 var m=ss.getSheetByName('ISSUE79_NOMATCH4_카드마스터원문')||ss.insertSheet('ISSUE79_NOMATCH4_카드마스터원문');m.clear();m.getRange(1,1,1,2).setValues([['행','롯데/LOCA관련 원문']]);if(master.length)m.getRange(2,1,master.length,2).setValues(master);m.setFrozenRows(1);m.getRange(1,1,1,2).setFontWeight('bold');
}
function status_(ss,p){var sh=ss.getSheetByName('ISSUE79_NOMATCH4_원본Coverage상태')||ss.insertSheet('ISSUE79_NOMATCH4_원본Coverage상태');sh.clearContents();sh.getRange(1,1,1,2).setValues([['항목','값']]);if(p.length)sh.getRange(2,1,p.length,2).setValues(p);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}
function hdr_(v,req){for(var r=0;r<Math.min(v.length,30);r++){var h=v[r]||[],ok=req.every(function(x){return ix_(h,[x])>=0;});if(ok)return r;}return 0;}
function ix_(h,names){var nn=names.map(c_);for(var i=0;i<(h||[]).length;i++){var c=c_(h[i]);if(nn.indexOf(c)>=0)return i;}return-1;}
function t_(x){return x===null||x===undefined?'':String(x).trim();}
function c_(x){return t_(x).replace(/\s/g,'').replace(/[()\[\]{}_:：\-/.]/g,'').toLowerCase();}
function id_(x){return t_(x).replace(/,/g,'').replace(/\s/g,'').replace(/\.0+$/,'');}
function n_(x){if(typeof x==='number'&&isFinite(x))return Math.round(x);var s=t_(x).replace(/,/g,'').replace(/원/g,'').replace(/\s/g,'');if(!s)return 0;var n=Number(s);return isFinite(n)?Math.round(n):0;}
function normE4_(x){var s=t_(x);var m=s.match(/(\d{1,4})\D*$/);if(!m)return'';var d=m[1];if(d.length>4)d=d.slice(-4);return('0000'+d).slice(-4);}
function date_(x){if(x instanceof Date&&!isNaN(x.getTime()))return Utilities.formatDate(x,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=t_(x),m=s.match(/(20\d{2})[-/.년\s]*(\d{1,2})[-/.월\s]*(\d{1,2})/);return m?m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2):'';}
function findDate_(s){return date_(s);}
function findMoney_(row){for(var i=0;i<(row||[]).length;i++){var s=t_(row[i]);if(/^[-+]?\d{1,3}(,\d{3})+$/.test(s)){var n=n_(s);if(Math.abs(n)>=1000)return n;}}return 0;}
function days_(a,b){var x=new Date(a+'T00:00:00+09:00'),y=new Date(b+'T00:00:00+09:00');if(isNaN(x)||isNaN(y))return 9999;return Math.round((y-x)/86400000);}
function sig_(sh){if(!sh)return'MISSING';var v=sh.getDataRange().getDisplayValues(),h=2166136261;for(var i=0;i<v.length;i++)for(var j=0;j<v[i].length;j++){var s=String(v[i][j]);for(var k=0;k<s.length;k++){h^=s.charCodeAt(k);h=Math.imul(h,16777619);}h^=31;h=Math.imul(h,16777619);}return[sh.getLastRow(),sh.getLastColumn(),h>>>0].join('|');}
