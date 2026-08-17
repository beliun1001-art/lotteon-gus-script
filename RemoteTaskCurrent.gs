var LOTTEON_REMOTE_TASK={id:'ISSUE79-V544-FAST-FINALIZE',title:'Issue79 SOURCE 빠른 최종검증 READ-ONLY',enabled:true,statusSheet:'ISSUE79_SOURCEUID_최종상태'};
var I79V544='v5.4.4-ISSUE79-SOURCE-FAST-FINALIZE-READONLY';
function runLotteonRemoteTaskStartRemote_(){return i79v544run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v544run_();}
function i79v544run_(){
  var ss=SpreadsheetApp.getActive(),started=new Date().toISOString();
  try{
    i79v544status_(ss,[['version',I79V544],['상태','RUNNING'],['단계','FAST_FINALIZE'],['메시지','고정 건수 가드 없이 실제 SOURCE 기준 빠른 최종검증 중'],['실행시작',started]]);
    var prev=i79v544kv_(ss.getSheetByName('ISSUE79_SOURCEUID_진단상태'));
    if(prev.version!=='v5.4.2-ISSUE79-SOURCE-UID-SINGLE-VERSION-READONLY')throw new Error('v5.4.2 상태시트 version 불일치: '+String(prev.version||''));
    var src=ss.getSheetByName('매출데이터_붙여넣기');if(!src)throw new Error('SOURCE 시트 누락');
    var h=src.getRange(1,1,1,src.getLastColumn()).getValues()[0];
    var idc=i79v544ix_(h,['마켓주문번호']),soc=i79v544ix_(h,['사이트주문번호']);
    if(idc<0||soc<0)throw new Error('SOURCE 식별 컬럼 누락');
    var firstStart=87,dupStart=4695,n=1970,offset=4608;
    var oi=src.getRange(firstStart,idc+1,n,1).getValues(),od=src.getRange(firstStart,idc+1,n,1).getDisplayValues();
    var os=src.getRange(firstStart,soc+1,n,1).getValues(),osd=src.getRange(firstStart,soc+1,n,1).getDisplayValues();
    var di=src.getRange(dupStart,idc+1,n,1).getValues(),dd=src.getRange(dupStart,idc+1,n,1).getDisplayValues();
    var ds=src.getRange(dupStart,soc+1,n,1).getValues(),dsd=src.getRange(dupStart,soc+1,n,1).getDisplayValues();
    var rows=[],dupSci=0,origSci=0,exactMatch=0,exactMismatch=0,orderMismatch=0;
    for(var i=0;i<n;i++){
      var ods=String(osd[i][0]==null?'':osd[i][0]),dds=String(dsd[i][0]==null?'':dsd[i][0]);
      if(/[eE][+\-]?\d+/.test(ods))origSci++;
      if(!/[eE][+\-]?\d+/.test(dds))continue;
      dupSci++;
      var oid=i79v544id_(oi[i][0],od[i][0]),did=i79v544id_(di[i][0],dd[i][0]);
      var oex=i79v544id_(os[i][0],osd[i][0]),dex=i79v544id_(ds[i][0],dsd[i][0]);
      if(oid!==did)orderMismatch++;
      var same=oex===dex&&!!oex;if(same)exactMatch++;else exactMismatch++;
      rows.push([dupStart+i,firstStart+i,did,dds,dex,ods,oex,same?'EXACT_MATCH':'MISMATCH']);
    }
    if(orderMismatch!==0)throw new Error('원본/복제 마켓주문번호 불일치: '+orderMismatch);
    if(exactMismatch!==0)throw new Error('과학표기 RAW 원본 불일치: '+exactMismatch);
    var off=i79v544table_(ss.getSheetByName('ISSUE79_SOURCEUID_중복오프셋'));
    if(!off.length||Number(off[0][0])!==4608)throw new Error('v5.4.2 offset 핵심 결과 불일치: '+JSON.stringify(off.slice(0,3)));
    var diff=i79v544table_(ss.getSheetByName('ISSUE79_SOURCEUID_CARD차이')),diffSum=0;
    for(var d=0;d<diff.length;d++)diffSum+=i79v544n_(diff[d][3]);
    diffSum=Math.round(diffSum);
    var cardPurchase=105314779,reconPurchase=Math.round(cardPurchase-diffSum);
    var toss=i79v544table_(ss.getSheetByName('ISSUE79_SOURCEUID_TOSS3')),tossRecon=0,tossCard=0;
    if(toss.length!==3)throw new Error('TOSS3 행 수 불일치: '+toss.length);
    for(var t=0;t<toss.length;t++){tossRecon+=i79v544n_(toss[t][1]);tossCard+=i79v544n_(toss[t][2]);}
    i79v544write_(ss,'ISSUE79_SOURCEUID_과학표기검증',['복제행','원본행','마켓주문번호','복제표시값','복제RAW정확값','원본표시값','원본RAW정확값','판정'],rows);
    i79v544status_(ss,[
      ['version',I79V544],['상태','PASS'],['단계','DONE'],['메시지','과학표기는 표시문제; 실제 감지건 모두 원본 RAW와 일치. SOURCE 수정 없이 UID 최신버전 재구성 결과를 최종 기준으로 사용'],
      ['v5.4.2_상태표시',String(prev['상태']||'')],['복제블록_원본시작행',firstStart],['복제블록_복제시작행',dupStart],['복제블록_행수',n],['복제블록_오프셋',offset],['UID_오프셋4608그룹',Number(off[0][1])||0],
      ['사이트주문번호_복제과학표기행',dupSci],['사이트주문번호_원본과학표기행',origSci],['사이트주문번호_RAW정확일치',exactMatch],['사이트주문번호_RAW불일치',exactMismatch],['마켓주문번호_원본복제불일치',orderMismatch],
      ['CARD대비_매입차이주문',diff.length],['CARD대비_매입합차이',diffSum],['재구성_H1활성매입합',reconPurchase],['현재_CARD매입합',cardPurchase],
      ['TOSS3_재구성매입합',Math.round(tossRecon)],['TOSS3_현재CARD매입합',Math.round(tossCard)],['TOSS3_차이',Math.round(tossCard-tossRecon)],
      ['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
    ]);
    return{ok:true,done:true,version:I79V544,scientific:dupSci,exactMatch:exactMatch,diffOrders:diff.length,reconstructedPurchase:reconPurchase};
  }catch(e){var m=String(e&&e.stack?e.stack:(e&&e.message?e.message:e));try{i79v544status_(ss,[['version',I79V544],['상태','ERROR'],['단계','FAILED'],['메시지','Issue79 v5.4.4 빠른 최종검증 실패'],['오류',m],['완료시각',new Date().toISOString()]]);}catch(_e){}throw e;}
}
function i79v544ix_(h,a){for(var i=0;i<a.length;i++){var k=String(a[i]).replace(/\s/g,'');for(var j=0;j<h.length;j++)if(String(h[j]==null?'':h[j]).replace(/\s/g,'')===k)return j;}return-1;}
function i79v544id_(raw,display){if(typeof raw==='number'&&isFinite(raw)&&Math.abs(raw)<=Number.MAX_SAFE_INTEGER)return String(Math.trunc(raw));var r=String(raw==null?'':raw).replace(/,/g,'').replace(/\.0+$/,'').replace(/\s/g,'');if(/^[+\-]?\d+(?:\.\d+)?[eE][+\-]?\d+$/.test(r)){var n=Number(r);if(isFinite(n)&&Math.abs(n)<=Number.MAX_SAFE_INTEGER)return String(Math.trunc(n));}var s=String(display==null?'':display).replace(/,/g,'').replace(/\.0+$/,'').replace(/\s/g,'');if(/^[0-9]+$/.test(s))return s;if(/^[0-9]+$/.test(r))return r;return r;}
function i79v544n_(v){if(typeof v==='number')return isFinite(v)?v:0;var n=Number(String(v==null?'':v).replace(/,/g,'').replace(/[^0-9.\-]/g,''));return isFinite(n)?n:0;}
function i79v544kv_(sh){if(!sh)throw new Error('상태시트 누락');var v=sh.getRange(1,1,Math.max(1,sh.getLastRow()),2).getValues(),m={};for(var i=1;i<v.length;i++)if(String(v[i][0]||''))m[String(v[i][0])]=v[i][1];return m;}
function i79v544table_(sh){if(!sh)throw new Error('진단시트 누락');var lr=sh.getLastRow(),lc=sh.getLastColumn();if(lr<2)return[];return sh.getRange(2,1,lr-1,lc).getValues();}
function i79v544write_(ss,name,header,rows){var sh=ss.getSheetByName(name)||ss.insertSheet(name);sh.clearContents();sh.getRange(1,1,1,header.length).setValues([header]);if(rows&&rows.length)sh.getRange(2,1,rows.length,header.length).setValues(rows);sh.setFrozenRows(1);}
function i79v544status_(ss,rows){i79v544write_(ss,'ISSUE79_SOURCEUID_최종상태',['항목','값'],rows);}
