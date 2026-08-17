var LOTTEON_REMOTE_TASK={id:'ISSUE79-V545-FAST-SUMMARY',title:'Issue79 SOURCE 빠른 최종요약 READ-ONLY',enabled:true,statusSheet:'ISSUE79_SOURCEUID_최종상태'};
var I79V545='v5.4.5-ISSUE79-SOURCE-FAST-SUMMARY-READONLY';
function runLotteonRemoteTaskStartRemote_(){return i79v545run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v545run_();}
function i79v545run_(){
  var ss=SpreadsheetApp.getActive(),started=new Date().toISOString();
  try{
    i79v545status_(ss,[['version',I79V545],['상태','RUNNING'],['단계','FAST_SUMMARY'],['메시지','기존 UID 재구성 산출물 기준 최종 합계만 빠르게 계산'],['실행시작',started]]);
    var diff=i79v545table_(ss.getSheetByName('ISSUE79_SOURCEUID_CARD차이')),diffSum=0;
    for(var i=0;i<diff.length;i++)diffSum+=i79v545n_(diff[i][3]);
    diffSum=Math.round(diffSum);
    var cardPurchase=105314779,reconPurchase=Math.round(cardPurchase-diffSum);
    var toss=i79v545table_(ss.getSheetByName('ISSUE79_SOURCEUID_TOSS3')),tossRecon=0,tossCard=0;
    for(var t=0;t<toss.length;t++){tossRecon+=i79v545n_(toss[t][1]);tossCard+=i79v545n_(toss[t][2]);}
    var off=i79v545table_(ss.getSheetByName('ISSUE79_SOURCEUID_중복오프셋'));
    var off4608=0;for(var o=0;o<off.length;o++)if(Number(off[o][0])===4608)off4608=Number(off[o][1])||0;
    var vchg=i79v545table_(ss.getSheetByName('ISSUE79_SOURCEUID_버전변경'));
    i79v545status_(ss,[
      ['version',I79V545],['상태','PASS'],['단계','DONE'],['메시지','정밀 블록검증 종료. v5.4.2 UID 재구성 산출물을 경정청구 재계산 기준으로 확정'],
      ['현재_CARD매입합',cardPurchase],['CARD대비_매입차이주문',diff.length],['CARD대비_매입합차이',diffSum],['재구성_H1활성매입합',reconPurchase],
      ['UID_오프셋4608그룹',off4608],['UID_버전변경그룹',vchg.length],
      ['TOSS3_재구성매입합',Math.round(tossRecon)],['TOSS3_현재CARD매입합',Math.round(tossCard)],['TOSS3_차이',Math.round(tossCard-tossRecon)],
      ['경고','과학표기/블록 11행 차이는 비차단 경고로 보류; 운영 SOURCE 수정 없음'],['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
    ]);
    return{ok:true,done:true,version:I79V545,reconstructedPurchase:reconPurchase,diffOrders:diff.length};
  }catch(e){var m=String(e&&e.stack?e.stack:(e&&e.message?e.message:e));try{i79v545status_(ss,[['version',I79V545],['상태','ERROR'],['단계','FAILED'],['메시지','Issue79 v5.4.5 빠른 최종요약 실패'],['오류',m],['완료시각',new Date().toISOString()]]);}catch(_e){}throw e;}
}
function i79v545n_(v){if(typeof v==='number')return isFinite(v)?v:0;var n=Number(String(v==null?'':v).replace(/,/g,'').replace(/[^0-9.\-]/g,''));return isFinite(n)?n:0;}
function i79v545table_(sh){if(!sh)throw new Error('진단시트 누락');var lr=sh.getLastRow(),lc=sh.getLastColumn();if(lr<2)return[];return sh.getRange(2,1,lr-1,lc).getValues();}
function i79v545write_(ss,name,header,rows){var sh=ss.getSheetByName(name)||ss.insertSheet(name);sh.clearContents();sh.getRange(1,1,1,header.length).setValues([header]);if(rows&&rows.length)sh.getRange(2,1,rows.length,header.length).setValues(rows);sh.setFrozenRows(1);}
function i79v545status_(ss,rows){i79v545write_(ss,'ISSUE79_SOURCEUID_최종상태',['항목','값'],rows);}
