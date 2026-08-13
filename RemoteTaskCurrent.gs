/**
 * Issue #55 v1.0 read-only order-level delta diagnostic.
 * Compares protected production `부가세_카드매칭검증` with Issue54 corrected preview.
 * Writes only ISSUE55_* sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE55-v1.0-20260813',
  title: 'Issue54 카드매칭 PASS vs 기존 운영 상태차이 정밀진단',
  enabled: true,
  outputSheet: 'ISSUE55_카드매칭차이진단',
  statusSheet: 'ISSUE55_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var status = i55Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  i55Write_(status,[['항목','값'],['버전','v1.0-ISSUE55-CARD-DELTA-DIAGNOSTIC'],['상태','RUNNING'],['단계','LOAD'],['메시지','Issue54 preview와 기존 운영 카드검증 주문별 비교 시작'],['운영시트 변경','0']]);
  try {
    var oldSh = ss.getSheetByName('부가세_카드매칭검증');
    var newSh = ss.getSheetByName('ISSUE54_카드매칭전체PREVIEW');
    var vatSh = ss.getSheetByName('부가세_신고자료');
    if (!oldSh || oldSh.getLastRow() < 2) throw new Error('부가세_카드매칭검증이 없습니다.');
    if (!newSh || newSh.getLastRow() < 2) throw new Error('ISSUE54_카드매칭전체PREVIEW가 없습니다.');
    if (!vatSh || vatSh.getLastRow() < 2) throw new Error('부가세_신고자료가 없습니다.');
    var oldSigBefore=i55Signature_(oldSh.getDataRange().getValues());
    var vatSigBefore=i55Signature_(vatSh.getDataRange().getValues());
    var oldData=i55Load_(oldSh,true), newData=i55Load_(newSh,true);
    i55Req_(oldData.rows===1355,'기존 운영 주문수 불일치: '+oldData.rows);
    i55Req_(newData.rows===1355,'Issue54 preview 주문수 불일치: '+newData.rows);
    i55Req_(oldData.dup===0,'기존 운영 정규화키 중복: '+oldData.dup);
    i55Req_(newData.dup===0,'Issue54 정규화키 중복: '+newData.dup);
    i55Req_(oldData.counts.MATCHED===810&&oldData.counts.NON_CARD===494&&oldData.counts.AMBIGUOUS===1&&oldData.counts.NO_MATCH===50,'기존 운영 상태 기준 불일치: '+JSON.stringify(oldData.counts));
    i55Req_(newData.counts.MATCHED===808&&newData.counts.NON_CARD===498&&newData.counts.AMBIGUOUS===0&&newData.counts.NO_MATCH===49,'Issue54 상태 기준 불일치: '+JSON.stringify(newData.counts));
    i55Req_(Math.round(newData.purchase)===105762969,'Issue54 매입합계 불일치: '+Math.round(newData.purchase));
    var oldKeys=Object.keys(oldData.map),newKeys=Object.keys(newData.map),overlap=0,oldOnly=[],newOnly=[];
    oldKeys.forEach(function(k){if(newData.map[k])overlap++;else oldOnly.push(k);});
    newKeys.forEach(function(k){if(!oldData.map[k])newOnly.push(k);});
    i55Req_(overlap===1355&&oldOnly.length===0&&newOnly.length===0,'모집단 불일치 overlap/oldOnly/newOnly='+overlap+'/'+oldOnly.length+'/'+newOnly.length);
    var transitions={},changed=[],unchanged=0,changedPurchase=0,purchaseChanged=0,evidenceChangedWithinStatus=0;
    oldKeys.sort().forEach(function(k){var o=oldData.map[k],n=newData.map[k];if(o.status===n.status){unchanged++;if(i55MaterialFingerprint_(o)!==i55MaterialFingerprint_(n))evidenceChangedWithinStatus++;return;}var tk=o.status+' -> '+n.status;transitions[tk]=(transitions[tk]||0)+1;changedPurchase+=n.purchase;if(Math.round(o.purchase)!==Math.round(n.purchase))purchaseChanged++;changed.push([o.account,o.order,o.date,o.status,n.status,tk,o.purchase,n.purchase,Math.round(n.purchase-o.purchase),o.payment,n.payment,o.company,n.company,o.cardName,n.cardName,o.end4,n.end4,o.approvalDate,n.approvalDate,o.approvalNo,n.approvalNo,o.approvalAmount,n.approvalAmount,o.reason,n.reason,o.v669,n.v669,o.v670,n.v670,o.evidenceKey,n.evidenceKey]);});
    i55Req_(changed.length+unchanged===1355,'변경/동일 주문 합계 불일치');
    var out=i55Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    out.clearContents();
    var headers=['쿠팡계정ID','주문번호','주문일','기존상태','신규상태','상태이동','기존매입금액','신규매입금액','매입차액','기존롯데결제수단','신규롯데결제수단','기존카드사','신규카드사','기존카드명','신규카드명','기존끝4','신규끝4','기존승인일','신규승인일','기존승인번호','신규승인번호','기존승인금액','신규승인금액','기존매칭근거','신규매칭근거','기존v6.69','신규v6.69','기존v6.70','신규v6.70','기존canonicalKey','신규canonicalKey'];
    out.getRange(1,1,1,headers.length).setValues([headers]);
    if(changed.length)out.getRange(2,1,changed.length,headers.length).setValues(changed);
    out.getRange(1,1,1,headers.length).setFontWeight('bold');out.setFrozenRows(1);SpreadsheetApp.flush();
    i55Req_(i55Signature_(oldSh.getDataRange().getValues())===oldSigBefore,'부가세_카드매칭검증이 진단 중 변경되었습니다.');
    i55Req_(i55Signature_(vatSh.getDataRange().getValues())===vatSigBefore,'부가세_신고자료가 진단 중 변경되었습니다.');
    var rows=[['항목','값'],['버전','v1.0-ISSUE55-CARD-DELTA-DIAGNOSTIC'],['상태','PASS'],['단계','DONE'],['메시지','Issue54 preview와 기존 운영 카드검증 주문별 상태차이 진단 완료'],['운영시트 변경','0'],['기존운영주문',oldData.rows],['Issue54주문',newData.rows],['정규화overlap',overlap],['oldOnly',oldOnly.length],['newOnly',newOnly.length],['기존_MATCHED',oldData.counts.MATCHED],['기존_NON_CARD',oldData.counts.NON_CARD],['기존_AMBIGUOUS',oldData.counts.AMBIGUOUS],['기존_NO_MATCH',oldData.counts.NO_MATCH],['신규_MATCHED',newData.counts.MATCHED],['신규_NON_CARD',newData.counts.NON_CARD],['신규_AMBIGUOUS',newData.counts.AMBIGUOUS],['신규_NO_MATCH',newData.counts.NO_MATCH],['상태변경주문',changed.length],['상태동일주문',unchanged],['상태변경주문_신규매입합계',Math.round(changedPurchase)],['상태변경중_매입금액변경',purchaseChanged],['상태동일_증빙필드변경주문',evidenceChangedWithinStatus],['기존운영매입합계',Math.round(oldData.purchase)],['Issue54매입합계',Math.round(newData.purchase)]];
    Object.keys(transitions).sort().forEach(function(k){rows.push(['이동_'+k,transitions[k]]);});
    rows.push(['부가세_카드매칭검증 변경','0']);rows.push(['부가세_신고자료 변경','0']);rows.push(['완료시각',new Date().toISOString()]);
    i55Write_(status,rows);return {ok:true,done:true,changed:changed.length,transitions:transitions};
  } catch(e) {
    i55Write_(status,[['항목','값'],['버전','v1.0-ISSUE55-CARD-DELTA-DIAGNOSTIC'],['상태','ERROR'],['단계','FAILED'],['메시지','카드매칭 상태차이 진단 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0'],['완료시각',new Date().toISOString()]]);throw e;
  }
}
function runLotteonRemoteTaskContinueRemote_(){return {ok:true,done:true,reason:'NO_CONTINUE_REQUIRED'};}
function i55Load_(sheet,filterHalf){var v=sheet.getDataRange().getValues(),h=v[0].map(i55Text_),ix={year:i55Find_(h,['신고연도']),half:i55Find_(h,['반기']),date:i55Find_(h,['주문일','주문일자','날짜']),account:i55Find_(h,['쿠팡계정ID']),order:i55Find_(h,['주문번호']),purchase:i55Find_(h,['주문매입금액','매입금액']),payment:i55Find_(h,['롯데결제수단']),company:i55Find_(h,['구매카드사']),cardName:i55Find_(h,['구매카드명']),end4:i55Find_(h,['카드번호끝4']),approvalDate:i55Find_(h,['승인일']),approvalNo:i55Find_(h,['승인번호']),approvalAmount:i55Find_(h,['승인금액']),status:i55Find_(h,['카드매칭상태']),reason:i55Find_(h,['카드매칭근거']),v669:i55Find_(h,['v6.69 2차귀속']),v670:i55Find_(h,['v6.70 3차귀속']),evidenceKey:i55Find_(h,['canonicalEvidenceKey'])};i55Req_(ix.account>=0&&ix.order>=0&&ix.status>=0,'필수 헤더 누락: '+sheet.getName());var map={},rows=0,dup=0,purchase=0,counts={MATCHED:0,NON_CARD:0,AMBIGUOUS:0,NO_MATCH:0};for(var r=1;r<v.length;r++){var row=v[r];if(filterHalf&&ix.year>=0&&i55Text_(row[ix.year])&&i55Text_(row[ix.year])!=='2026')continue;if(filterHalf&&ix.half>=0&&i55Text_(row[ix.half])&&i55Text_(row[ix.half])!=='상반기')continue;var account=i55Text_(row[ix.account]).toLowerCase(),order=i55Text_(row[ix.order]),key=i55Key_(account,order);if(!key)continue;if(map[key])dup++;var rec={account:account,order:order,date:i55Val_(row,ix.date),purchase:i55Num_(i55ValRaw_(row,ix.purchase)),payment:i55Val_(row,ix.payment),company:i55Val_(row,ix.company),cardName:i55Val_(row,ix.cardName),end4:i55End4_(i55Val_(row,ix.end4)),approvalDate:i55Val_(row,ix.approvalDate),approvalNo:i55Val_(row,ix.approvalNo),approvalAmount:i55Num_(i55ValRaw_(row,ix.approvalAmount)),status:i55Status_(i55Val_(row,ix.status)),reason:i55Val_(row,ix.reason),v669:i55Val_(row,ix.v669),v670:i55Val_(row,ix.v670),evidenceKey:i55Val_(row,ix.evidenceKey)};map[key]=rec;rows++;purchase+=rec.purchase;counts[rec.status]=(counts[rec.status]||0)+1;}return {map:map,rows:rows,dup:dup,purchase:purchase,counts:counts};}
function i55MaterialFingerprint_(r){return [r.purchase,r.payment,r.company,r.cardName,r.end4,r.approvalDate,r.approvalNo,r.approvalAmount,r.reason,r.v669,r.v670,r.evidenceKey].join('|');}
function i55Status_(v){var s=i55Text_(v).toUpperCase();if(s==='MATCHED'||s==='MASTER_MATCHED')return 'MATCHED';if(s==='NON_CARD')return 'NON_CARD';if(s==='AMBIGUOUS')return 'AMBIGUOUS';return 'NO_MATCH';}
function i55Key_(a,o){var aa=i55Text_(a).toLowerCase(),oo=i55Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return aa&&oo?aa+'|'+oo:'';}
function i55End4_(v){var d=i55Text_(v).replace(/\D/g,'');return d?('0000'+d).slice(-4):'';}
function i55Find_(h,names){for(var n=0;n<names.length;n++){var q=i55Compact_(names[n]);for(var i=0;i<h.length;i++)if(i55Compact_(h[i])===q)return i;}return -1;}
function i55Val_(row,ix){return ix>=0?i55Text_(row[ix]):'';}function i55ValRaw_(row,ix){return ix>=0?row[ix]:0;}function i55Text_(v){return String(v==null?'':v).trim();}function i55Compact_(v){return i55Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}function i55Num_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(i55Text_(v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}function i55Signature_(v){var h=2166136261;for(var r=0;r<v.length;r++){for(var c=0;c<v[r].length;c++){var s=String(v[r][c]==null?'':v[r][c]);for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h+=(h<<1)+(h<<4)+(h<<7)+(h<<8)+(h<<24);}h^=31;}}return String(h>>>0)+'|'+v.length+'|'+(v[0]?v[0].length:0);}function i55Req_(ok,msg){if(!ok)throw new Error(msg);}function i55Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}function i55Write_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);SpreadsheetApp.flush();}
