var LOTTEON_REMOTE_TASK={id:'ISSUE79-V42-SCOPED-H1',title:'Issue79 2026상반기 범위고정 심층재검수',enabled:true,statusSheet:'ISSUE79_V4_실행확인'};
var ISSUE79_V4_PINNED_SOURCE='https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/327e007b7bfb4e751cf8141417421a0366c4ce27/RemoteTaskCurrent.gs';
var ISSUE79_V42='v4.2-ISSUE79-2026H1-SCOPED-DEEP-READONLY';
function runLotteonRemoteTaskStartRemote_(){return issue79V42Exec_();}
function runLotteonRemoteTaskContinueRemote_(){return issue79V42Exec_();}
function issue79V42Probe_(ss,rows){var sh=ss.getSheetByName('ISSUE79_V4_실행확인')||ss.insertSheet('ISSUE79_V4_실행확인');sh.clearContents();sh.getRange(1,1,1,2).setValues([['항목','값']]);if(rows&&rows.length)sh.getRange(2,1,rows.length,2).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}
function issue79V42Exec_(){
 var ss=SpreadsheetApp.getActive(),started=new Date().toISOString();
 issue79V42Probe_(ss,[['version',ISSUE79_V42],['상태','RUNNING'],['실행시작',started],['범위','2026 상반기'],['고정sourceCommit','327e007b7bfb4e751cf8141417421a0366c4ce27']]);
 try{
  var res=UrlFetchApp.fetch(ISSUE79_V4_PINNED_SOURCE+'?ts='+new Date().getTime(),{method:'get',muteHttpExceptions:true,followRedirects:true});
  var src=res.getContentText('UTF-8');
  if(res.getResponseCode()!==200||!src)throw new Error('v4 pinned source fetch 실패 HTTP '+res.getResponseCode());
  if(src.indexOf("ISSUE79-NOMATCH16-DEEP-v4")<0)throw new Error('v4 pinned source marker 불일치');
  var oldLine="  return eval(code+'\\n'+extra+'\\n;i79runV2_();\\ni79v4deep_();');";
  if(src.indexOf(oldLine)<0)throw new Error('v4 final eval anchor 누락');
  var scoped=String.raw`
var I79V42_ALL_USED={};
function i79v42scopeHalf_(v){var s=i79c_(v);return s==='상반기'||s==='1h'||s==='h1'||s==='1'||s==='1반기';}
function i79v42usedKey_(o){return o.approvalNo?'APP|'+i79c_(o.company)+'|'+i79c_(o.approvalNo):(o.approvalDate||o.approvalAmount||o.merchant||o.source?'ROW|'+o.approvalDate+'|'+o.approvalAmount+'|'+i79c_(o.merchant)+'|'+i79c_(o.source):'');}
function i79v42markUsed_(o){if(o.status!=='MATCHED'&&o.status!=='NON_CARD')return;var k=i79v42usedKey_(o);if(k)I79V42_ALL_USED[k]=(I79V42_ALL_USED[k]||0)+1;}
function i79readCardScoped_(sh){
 var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=i79header_(v,['주문번호','카드매칭상태']);if(hr<0)throw new Error('카드검증 header 탐지 실패');
 var h=v[hr],x={year:i79ix_(h,['신고연도']),half:i79ix_(h,['반기']),date:i79ix_(h,['주문일']),business:i79ix_(h,['사업자등록번호']),account:i79ix_(h,['쿠팡계정ID']),order:i79ix_(h,['주문번호']),purchase:i79ix_(h,['주문매입금액']),payment:i79ix_(h,['롯데결제수단']),status:i79ix_(h,['카드매칭상태']),company:i79ix_(h,['구매카드사']),name:i79ix_(h,['구매카드명']),end4:i79ix_(h,['카드번호끝4']),adate:i79ix_(h,['승인일']),approval:i79ix_(h,['승인번호']),aamount:i79ix_(h,['승인금액']),merchant:i79ix_(h,['가맹점명']),source:i79ix_(h,['원본파일'])};
 if(x.year<0||x.half<0)throw new Error('카드검증 신고연도/반기 header 누락');
 var rows=[],stats={},p=0;I79V42_ALL_USED={};
 for(var r=hr+1;r<v.length;r++){
  var no=x.order>=0?i79t_(d[r][x.order]):'';if(!no)continue;
  var st=i79t_(v[r][x.status]).toUpperCase(),amt=i79n_(v[r][x.purchase]);
  var o={business:i79t_(v[r][x.business]),account:i79t_(v[r][x.account]),orderDate:i79date_(v[r][x.date]),orderNo:no,purchase:amt,payment:x.payment>=0?i79t_(v[r][x.payment]):'',status:st,company:x.company>=0?i79t_(v[r][x.company]):'',cardName:x.name>=0?i79t_(v[r][x.name]):'',end4:x.end4>=0?i79t_(d[r][x.end4]):'',approvalDate:x.adate>=0?i79date_(v[r][x.adate]):'',approvalNo:x.approval>=0?i79t_(d[r][x.approval]):'',approvalAmount:x.aamount>=0?i79n_(v[r][x.aamount]):0,merchant:x.merchant>=0?i79t_(v[r][x.merchant]):'',source:x.source>=0?i79t_(v[r][x.source]):''};
  i79v42markUsed_(o);
  var yr=i79c_(v[r][x.year]);if(yr!=='2026'||!i79v42scopeHalf_(v[r][x.half]))continue;
  stats[st]=(stats[st]||0)+1;p+=amt;rows.push(o);
 }
 return{rows:rows,stats:stats,purchase:Math.round(p),scope:'2026 상반기'};
}
function i79usedScoped_(rows){var out={},k;for(k in I79V42_ALL_USED)if(Object.prototype.hasOwnProperty.call(I79V42_ALL_USED,k))out[k]=I79V42_ALL_USED[k];return out;}
`;
  var replacement="  return eval(code+'\\n'+extra+'\\n'+"+JSON.stringify(scoped)+"+'\\n;i79readCard_=i79readCardScoped_;i79used_=i79usedScoped_;i79runV2_();\\ni79v4deep_();');";
  src=src.replace(oldLine,replacement);
  var result=eval(src+'\n;issue79V4Exec_();');
  issue79V42Probe_(ss,[['version',ISSUE79_V42],['상태','PASS'],['실행시작',started],['실행완료',new Date().toISOString()],['범위','2026 상반기'],['scopeMarker','2026H1_OK'],['고정sourceCommit','327e007b7bfb4e751cf8141417421a0366c4ce27'],['심층version',result&&result.version?result.version:''],['최종SAFE',result&&result.totalSafe!==undefined?result.totalSafe:''],['최종잔여',result&&result.remain!==undefined?result.remain:'']]);
  return result;
 }catch(e){var msg=String(e&&e.message?e.message:e);issue79V42Probe_(ss,[['version',ISSUE79_V42],['상태','ERROR'],['실행시작',started],['실행완료',new Date().toISOString()],['범위','2026 상반기'],['고정sourceCommit','327e007b7bfb4e751cf8141417421a0366c4ce27'],['오류',msg]]);throw e;}
}
