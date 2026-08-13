/**
 * Issue #57 v1.0 read-only deep diagnostic for 9 blocked card-matching transitions.
 * Reads current production VAT/card evidence plus Issue54/55 diagnostics.
 * Writes only ISSUE57_* sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE57-v1.0-20260813',
  title: 'Issue56 차단 9건 카드원본·canonical 심층진단',
  enabled: true,
  outputSheet: 'ISSUE57_차단9건심층진단',
  statusSheet: 'ISSUE57_진단상태'
};
var ISSUE57_CODE_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Code.gs';
var ISSUE57_BOOT_URL = 'https://raw.githubusercontent.com/beliun1001-art/lotteon-gus-script/main/Patch_v6_24_bootstrap_auto_continue.gs';

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var status = i57Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  i57Write_(status, [
    ['항목','값'],['버전','v1.0-ISSUE57-BLOCKED9-DEEP-DIAGNOSTIC'],['상태','RUNNING'],['단계','LOAD'],
    ['메시지','Issue56 차단 9건 raw/canonical 심층진단 시작'],['운영시트 변경','0']
  ]);
  try {
    var deltaSh = ss.getSheetByName('ISSUE55_카드매칭차이진단');
    var previewSh = ss.getSheetByName('ISSUE54_카드매칭전체PREVIEW');
    var oldSh = ss.getSheetByName('부가세_카드매칭검증');
    var vatSh = ss.getSheetByName('부가세_신고자료');
    var histSh = ss.getSheetByName('카드사용내역_붙여넣기');
    var masterSh = ss.getSheetByName('카드_마스터');
    if (!deltaSh || deltaSh.getLastRow() < 2) throw new Error('ISSUE55_카드매칭차이진단이 없습니다.');
    if (!previewSh || previewSh.getLastRow() < 2) throw new Error('ISSUE54_카드매칭전체PREVIEW가 없습니다.');
    if (!oldSh || oldSh.getLastRow() < 2) throw new Error('부가세_카드매칭검증이 없습니다.');
    if (!vatSh || vatSh.getLastRow() < 2) throw new Error('부가세_신고자료가 없습니다.');
    if (!histSh || histSh.getLastRow() < 2) throw new Error('카드사용내역_붙여넣기가 없습니다.');
    if (!masterSh || masterSh.getLastRow() < 2) throw new Error('카드_마스터가 없습니다.');

    var sigs = {
      old:i57Sig_(oldSh.getDataRange().getValues()), vat:i57Sig_(vatSh.getDataRange().getValues()),
      hist:i57Sig_(histSh.getDataRange().getValues()), master:i57Sig_(masterSh.getDataRange().getValues()),
      preview:i57Sig_(previewSh.getDataRange().getValues()), delta:i57Sig_(deltaSh.getDataRange().getValues())
    };
    var targets = i57LoadTargets_(deltaSh);
    i57Req_(targets.length === 9, '차단 대상 9건 불일치: ' + targets.length);
    var trans = {};
    targets.forEach(function(t){ trans[t.transition]=(trans[t.transition]||0)+1; });
    i57Req_(trans['MATCHED -> NO_MATCH'] === 5, 'MATCHED -> NO_MATCH 대상 불일치');
    i57Req_(trans['NO_MATCH -> NON_CARD'] === 4, 'NO_MATCH -> NON_CARD 대상 불일치');

    var deep = i57RunMainEvidence_(ss, targets);
    i57Req_(deep && deep.rows && deep.rows.length === 9, '심층진단 결과 9건 불일치');
    i57Req_(deep.canonicalRows === 1990, 'canonical 증빙행 불일치: ' + deep.canonicalRows);

    var out = i57Ensure_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    out.clearContents();
    var headers = [
      '쿠팡계정ID','주문번호','상태이동','주문일','기존매입금액','신규매입금액','매입금액변경',
      '기존승인번호','기존승인금액','기존끝4','신규승인번호','신규승인금액','신규끝4',
      '기존매칭근거','신규매칭근거','기존canonicalKey','신규canonicalKey',
      'raw기존증빙건수','raw취소성값건수','canonical기존증빙건수','corrected금액canonical후보','corrected금액+끝4후보','corrected금액+끝4_주문일7일내후보',
      '분류','판정사유','canonical샘플','raw샘플'
    ];
    out.getRange(1,1,1,headers.length).setValues([headers]);
    out.getRange(2,1,deep.rows.length,headers.length).setValues(deep.rows);
    out.getRange(1,1,1,headers.length).setFontWeight('bold');
    out.setFrozenRows(1);
    SpreadsheetApp.flush();

    i57Req_(i57Sig_(oldSh.getDataRange().getValues()) === sigs.old, '부가세_카드매칭검증이 진단 중 변경되었습니다.');
    i57Req_(i57Sig_(vatSh.getDataRange().getValues()) === sigs.vat, '부가세_신고자료가 진단 중 변경되었습니다.');
    i57Req_(i57Sig_(histSh.getDataRange().getValues()) === sigs.hist, '카드사용내역_붙여넣기가 진단 중 변경되었습니다.');
    i57Req_(i57Sig_(masterSh.getDataRange().getValues()) === sigs.master, '카드_마스터가 진단 중 변경되었습니다.');
    i57Req_(i57Sig_(previewSh.getDataRange().getValues()) === sigs.preview, 'ISSUE54 preview가 진단 중 변경되었습니다.');
    i57Req_(i57Sig_(deltaSh.getDataRange().getValues()) === sigs.delta, 'ISSUE55 delta가 진단 중 변경되었습니다.');

    var counts = {EXPLAINED_SAFE:0,LIKELY_MATCHER_BUG:0,DATA_GAP_REVIEW:0,INVALID_STATE:0};
    deep.rows.forEach(function(row){ var c=String(row[23]||''); counts[c]=(counts[c]||0)+1; });
    var approve = (counts.EXPLAINED_SAFE === 9) ? 'YES' : 'NO';
    var rows = [
      ['항목','값'],['버전','v1.0-ISSUE57-BLOCKED9-DEEP-DIAGNOSTIC'],['상태','PASS'],['단계','DONE'],
      ['메시지','Issue56 차단 9건 raw/canonical 심층진단 완료'],['운영시트 변경','0'],
      ['진단대상',9],['MATCHED -> NO_MATCH',5],['NO_MATCH -> NON_CARD',4],
      ['canonical증빙행',deep.canonicalRows],['rawHistory행',deep.historyRows],
      ['EXPLAINED_SAFE',counts.EXPLAINED_SAFE||0],['LIKELY_MATCHER_BUG',counts.LIKELY_MATCHER_BUG||0],
      ['DATA_GAP_REVIEW',counts.DATA_GAP_REVIEW||0],['INVALID_STATE',counts.INVALID_STATE||0],
      ['운영반영자동승인',approve],['부가세_카드매칭검증 변경','0'],['부가세_신고자료 변경','0'],
      ['카드사용내역_붙여넣기 변경','0'],['카드_마스터 변경','0']
    ];
    deep.rows.forEach(function(r,idx){ if (String(r[23]||'') !== 'EXPLAINED_SAFE') rows.push(['확인필요_'+(idx+1), r[0]+' | '+r[1]+' | '+r[2]+' | '+r[23]+' | '+r[24]]); });
    rows.push(['완료시각',new Date().toISOString()]);
    i57Write_(status, rows);
    return {ok:true,done:true,counts:counts,approve:approve};
  } catch (e) {
    i57Write_(status, [
      ['항목','값'],['버전','v1.0-ISSUE57-BLOCKED9-DEEP-DIAGNOSTIC'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','Issue56 차단 9건 심층진단 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0'],['완료시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function runLotteonRemoteTaskContinueRemote_(){ return {ok:true,done:true,reason:'NO_CONTINUE_REQUIRED'}; }

function i57LoadTargets_(sheet) {
  var v=sheet.getDataRange().getValues(), h=v[0].map(i57Text_);
  var x={
    account:i57Find_(h,['쿠팡계정ID']),order:i57Find_(h,['주문번호']),date:i57Find_(h,['주문일']),transition:i57Find_(h,['상태이동']),
    oldPurchase:i57Find_(h,['기존매입금액']),newPurchase:i57Find_(h,['신규매입금액']),
    oldApprovalNo:i57Find_(h,['기존승인번호']),newApprovalNo:i57Find_(h,['신규승인번호']),
    oldApprovalAmount:i57Find_(h,['기존승인금액']),newApprovalAmount:i57Find_(h,['신규승인금액']),
    oldEnd4:i57Find_(h,['기존끝4']),newEnd4:i57Find_(h,['신규끝4']),
    oldReason:i57Find_(h,['기존매칭근거']),newReason:i57Find_(h,['신규매칭근거']),
    oldKey:i57Find_(h,['기존canonicalKey']),newKey:i57Find_(h,['신규canonicalKey'])
  };
  i57Req_(x.account>=0&&x.order>=0&&x.transition>=0&&x.newPurchase>=0,'ISSUE55 필수 헤더 누락');
  var out=[];
  for(var r=1;r<v.length;r++){
    var tr=i57Text_(v[r][x.transition]);
    if(tr!=='MATCHED -> NO_MATCH' && tr!=='NO_MATCH -> NON_CARD') continue;
    out.push({
      account:i57Text_(v[r][x.account]).toLowerCase(),order:i57Text_(v[r][x.order]),date:i57Val_(v[r],x.date),transition:tr,
      oldPurchase:i57Num_(i57Raw_(v[r],x.oldPurchase)),newPurchase:i57Num_(i57Raw_(v[r],x.newPurchase)),
      oldApprovalNo:i57Val_(v[r],x.oldApprovalNo),newApprovalNo:i57Val_(v[r],x.newApprovalNo),
      oldApprovalAmount:i57Num_(i57Raw_(v[r],x.oldApprovalAmount)),newApprovalAmount:i57Num_(i57Raw_(v[r],x.newApprovalAmount)),
      oldEnd4:i57End4_(i57Val_(v[r],x.oldEnd4)),newEnd4:i57End4_(i57Val_(v[r],x.newEnd4)),
      oldReason:i57Val_(v[r],x.oldReason),newReason:i57Val_(v[r],x.newReason),oldKey:i57Val_(v[r],x.oldKey),newKey:i57Val_(v[r],x.newKey)
    });
  }
  return out;
}

function i57RunMainEvidence_(ss, targets) {
  var code=i57Fetch_(ISSUE57_CODE_URL), boot=i57Fetch_(ISSUE57_BOOT_URL), sid=ss.getId();
  var invocation=[
    ';(function(){',
    'var targets='+JSON.stringify(targets)+';',
    'var ss=SpreadsheetApp.openById('+JSON.stringify(sid)+');',
    'var history=loadVatCardHistory_v660_(ss);',
    'var master=loadVatCardMaster_v660_(ss);',
    'var canonical=canonicalizeVatHistory_v664_(history,master);',
    'function T(v){return String(v==null?"":v).trim();}',
    'function K(v){return T(v).toLowerCase().replace(/[\\s._()\\[\\]{}\\-\\/]/g,"");}',
    'function N(v){if(typeof v==="number"&&isFinite(v))return v;var n=Number(T(v).replace(/[원,%\\s,]/g,""));return isFinite(n)?n:0;}',
    'function E4(v){var d=T(v).replace(/\\D/g,"");return d?("0000"+d).slice(-4):"";}',
    'function props(o,out){out=out||[];if(o==null)return out;if(typeof o!=="object"){out.push({k:"",v:o});return out;}if(Array.isArray(o)){for(var i=0;i<o.length;i++)props(o[i],out);return out;}Object.keys(o).forEach(function(k){var v=o[k];if(v!=null&&typeof v==="object")props(v,out);else out.push({k:k,v:v});});return out;}',
    'function pick(o,names){var p=props(o,[]), nn=names.map(K);for(var a=0;a<nn.length;a++){for(var i=0;i<p.length;i++){if(K(p[i].k)===nn[a]&&T(p[i].v)!=="")return p[i].v;}}return "";}',
    'function exactVal(o,val){val=T(val);if(!val)return false;var p=props(o,[]);for(var i=0;i<p.length;i++)if(T(p[i].v)===val)return true;return false;}',
    'function vals(o){return props(o,[]).map(function(p){return T(p.v);}).filter(Boolean).join(" | ");}',
    'function cancelLike(o){return /취소|환불|반품|cancel|refund|return/i.test(vals(o));}',
    'function approvalNo(o){return T(pick(o,["approvalNo","approvalNumber","승인번호"]));}',
    'function amount(o){return N(pick(o,["approvalAmount","amount","승인금액","이용금액","사용금액","거래금액"]));}',
    'function end4(o){var v=pick(o,["cardEnd4","end4","cardLast4","카드번호끝4"]);if(T(v))return E4(v);return E4(pick(o,["cardNumber","카드번호"]));}',
    'function datev(o){return T(pick(o,["approvalDate","date","승인일","승인일자","이용일","거래일","사용일"]));}',
    'function ckey(o){return T(pick(o,["canonicalEvidenceKey","evidenceKey","canonicalKey"]));}',
    'function sample(o){if(!o)return "";return [pick(o,["company","cardCompany","카드사"]),pick(o,["cardName","alias","카드명"]),end4(o),datev(o),approvalNo(o),amount(o),ckey(o),pick(o,["sourceFile","source","원본파일"]),pick(o,["cancelMemo","cancel","취소메모"])].map(T).join(" / ").slice(0,900);}',
    'function day(v){var s=T(v);if(!s)return NaN;var d=new Date(s);if(!isNaN(d.getTime()))return Math.floor(d.getTime()/86400000);var m=s.match(/(20\\d{2})[^0-9]?(\\d{1,2})[^0-9]?(\\d{1,2})/);if(!m)return NaN;return Math.floor(new Date(Number(m[1]),Number(m[2])-1,Number(m[3])).getTime()/86400000);}',
    'var rows=[];',
    'targets.forEach(function(t){',
      'var oldNo=T(t.oldApprovalNo), oldKey=T(t.oldKey), targetEnd=E4(t.newEnd4||t.oldEnd4), orderDay=day(t.date);',
      'var rawOld=history.filter(function(h){return (oldNo&&approvalNo(h)===oldNo)||(oldKey&&exactVal(h,oldKey));});',
      'var canOld=canonical.filter(function(c){return (oldNo&&approvalNo(c)===oldNo)||(oldKey&&(ckey(c)===oldKey||exactVal(c,oldKey)));});',
      'var amtCandidates=canonical.filter(function(c){return Math.round(amount(c))===Math.round(Number(t.newPurchase||0));});',
      'var endCandidates=amtCandidates.filter(function(c){return !targetEnd||end4(c)===targetEnd;});',
      'var nearCandidates=endCandidates.filter(function(c){var cd=day(datev(c));return isNaN(orderDay)||isNaN(cd)||Math.abs(cd-orderDay)<=7;});',
      'var rawCancel=rawOld.filter(cancelLike);',
      'var sameAmount=Math.round(Number(t.oldApprovalAmount||0))===Math.round(Number(t.newPurchase||0))&&Number(t.oldApprovalAmount||0)!==0;',
      'var cls="DATA_GAP_REVIEW",why="";',
      'if(t.transition==="MATCHED -> NO_MATCH"){',
        'if(rawOld.length>0&&canOld.length===0){if(rawCancel.length>0){cls="EXPLAINED_SAFE";why="기존 승인증빙이 raw에는 있으나 취소/환불성 값과 함께 canonical에서 제외됨";}else{cls="DATA_GAP_REVIEW";why="기존 승인증빙이 raw에는 있으나 canonical 제외 사유가 취소로 설명되지 않음";}}',
        'else if(canOld.length>0){if(sameAmount||endCandidates.length>0){cls="LIKELY_MATCHER_BUG";why="현재 canonical에 기존/동액 exact 증빙이 살아있는데 신규 결과가 NO_MATCH";}else if(Math.round(Number(t.oldApprovalAmount||0))!==Math.round(Number(t.newPurchase||0))&&endCandidates.length===0){cls="EXPLAINED_SAFE";why="corrected 매입금액이 기존 승인금액과 달라졌고 새 금액 exact canonical 증빙이 없음";}else{cls="DATA_GAP_REVIEW";why="canonical 기존증빙은 있으나 NO_MATCH 전환이 충분히 설명되지 않음";}}',
        'else{cls="DATA_GAP_REVIEW";why="기존 MATCHED 승인증빙을 현재 raw/canonical에서 찾지 못함";}',
      '}else if(t.transition==="NO_MATCH -> NON_CARD"){',
        'var physical=!!(T(t.newApprovalNo)||E4(t.newEnd4)||Number(t.newApprovalAmount||0));',
        'if(endCandidates.length>0){cls="LIKELY_MATCHER_BUG";why="신규 NON_CARD인데 corrected 금액/카드끝4 exact canonical 카드후보가 존재";}',
        'else if(physical){cls="INVALID_STATE";why="NON_CARD 상태에 physical card 승인정보가 잔존하며 exact canonical 후보는 없음; 상태/필드 정합성 수정 필요";}',
        'else{cls="EXPLAINED_SAFE";why="NON_CARD이고 physical card identity 및 exact canonical 후보가 없음";}',
      '}',
      'rows.push([t.account,t.order,t.transition,t.date,t.oldPurchase,t.newPurchase,Math.round(Number(t.oldPurchase||0))===Math.round(Number(t.newPurchase||0))?"N":"Y",t.oldApprovalNo,t.oldApprovalAmount,t.oldEnd4,t.newApprovalNo,t.newApprovalAmount,t.newEnd4,t.oldReason,t.newReason,t.oldKey,t.newKey,rawOld.length,rawCancel.length,canOld.length,amtCandidates.length,endCandidates.length,nearCandidates.length,cls,why,sample(endCandidates[0]||canOld[0]||null),rawOld.length?vals(rawOld[0]).slice(0,900):""]);',
    '});',
    'return {rows:rows,historyRows:history.length,canonicalRows:canonical.length};',
    '})()'
  ].join('\n');
  return eval(code+'\n\n;\n\n'+boot+'\n\n;\n\n'+invocation);
}

function i57Fetch_(url){
  var res=UrlFetchApp.fetch(url+'?ts='+new Date().getTime(),{method:'get',muteHttpExceptions:true,followRedirects:true});
  var code=res.getResponseCode(),txt=res.getContentText('UTF-8');
  if(code<200||code>=300)throw new Error('main 코드 로드 실패 HTTP '+code+' '+url);
  return txt;
}
function i57Find_(h,names){for(var n=0;n<names.length;n++){var q=i57Compact_(names[n]);for(var i=0;i<h.length;i++)if(i57Compact_(h[i])===q)return i;}return -1;}
function i57Val_(row,ix){return ix>=0?i57Text_(row[ix]):'';}
function i57Raw_(row,ix){return ix>=0?row[ix]:0;}
function i57Text_(v){return String(v==null?'':v).trim();}
function i57Compact_(v){return i57Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function i57Num_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(i57Text_(v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function i57End4_(v){var d=i57Text_(v).replace(/\D/g,'');return d?('0000'+d).slice(-4):'';}
function i57Sig_(v){var h=2166136261;for(var r=0;r<v.length;r++){for(var c=0;c<v[r].length;c++){var s=String(v[r][c]==null?'':v[r][c]);for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h+=(h<<1)+(h<<4)+(h<<7)+(h<<8)+(h<<24);}h^=31;}}return String(h>>>0)+'|'+v.length+'|'+(v[0]?v[0].length:0);}
function i57Req_(ok,msg){if(!ok)throw new Error(msg);}
function i57Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function i57Write_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);SpreadsheetApp.flush();}
