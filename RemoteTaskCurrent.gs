/**
 * Issue #51 v1.1 read-only difference diagnostic.
 * Compares the completed 1,893-order preview with the protected old 1,355-order
 * production verification and classifies the 538 new orders, especially the
 * 537 NO_MATCH rows. No production sheet is modified.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE51-v1.1-20260812',
  title: '현재 VAT 카드매칭 차이·신규 NO_MATCH 원인 진단',
  enabled: true,
  outputSheet: 'ISSUE51_차이진단',
  statusSheet: 'ISSUE51_실행상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue51v11Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue51v11Status_(state, [
    ['항목','값'],['버전','v1.1-ISSUE51-DIFF-NEW-NOMATCH-DIAGNOSTIC'],
    ['상태','RUNNING'],['단계','LOAD'],['메시지','기존 1,355 상태차이 및 신규 538 NO_MATCH 원인 진단 시작'],['운영시트 변경','0']
  ]);

  try {
    var preview = ss.getSheetByName('ISSUE51_카드매칭전체PREVIEW');
    var old = ss.getSheetByName('부가세_카드매칭검증');
    var history = ss.getSheetByName('카드사용내역_붙여넣기');
    if (!preview || preview.getLastRow() < 2) throw new Error('ISSUE51_카드매칭전체PREVIEW가 없습니다.');
    if (!old || old.getLastRow() < 2) throw new Error('부가세_카드매칭검증이 없습니다.');
    if (!history || history.getLastRow() < 2) throw new Error('카드사용내역_붙여넣기가 없습니다.');

    var pv = preview.getDataRange().getValues();
    var ph = pv[0].map(issue51v11Text_);
    var pi = issue51v11Indexes_(ph, {
      account:['쿠팡계정ID'], order:['주문번호'], date:['주문일'], purchase:['주문매입금액'],
      payment:['롯데결제수단'], origin:['주문구분'], status:['카드매칭상태'], reason:['카드매칭근거']
    });
    issue51v11Require_(pi.account>=0&&pi.order>=0&&pi.date>=0&&pi.purchase>=0&&pi.origin>=0&&pi.status>=0&&pi.reason>=0, 'preview 필수 헤더 누락');

    var ov = old.getDataRange().getValues();
    var oh = ov[0].map(issue51v11Text_);
    var oi = issue51v11Indexes_(oh, {year:['신고연도'],half:['반기'],account:['쿠팡계정ID'],order:['주문번호'],status:['카드매칭상태']});
    issue51v11Require_(oi.account>=0&&oi.order>=0&&oi.status>=0, '기존 카드검증 필수 헤더 누락');

    var oldMap = {}, oldRows = 0, oldCounts = {MATCHED:0,NON_CARD:0,AMBIGUOUS:0,NO_MATCH:0};
    for (var o=1;o<ov.length;o++) {
      var orow=ov[o];
      if (oi.year>=0 && issue51v11Text_(orow[oi.year])!=='2026') continue;
      if (oi.half>=0 && issue51v11Text_(orow[oi.half])!=='상반기') continue;
      var ok=issue51v11Key_(orow[oi.account],orow[oi.order]); if(!ok) continue;
      var os=issue51v11StatusName_(orow[oi.status]);
      if (oldMap[ok]) throw new Error('기존 카드검증 정규화키 중복: '+ok);
      oldMap[ok]={status:os}; oldRows++; oldCounts[os]=(oldCounts[os]||0)+1;
    }
    issue51v11Require_(oldRows===1355, '기존 카드검증 주문수 불일치: '+oldRows);
    issue51v11Require_(oldCounts.MATCHED===810&&oldCounts.NON_CARD===494&&oldCounts.AMBIGUOUS===1&&oldCounts.NO_MATCH===50,
      '기존 카드검증 보호 기준 불일치: '+JSON.stringify(oldCounts));

    var hv=history.getDataRange().getValues(), hh=hv[0].map(issue51v11Text_);
    var hi=issue51v11Indexes_(hh,{date:['승인일','이용일','거래일','사용일','승인일자'],amount:['승인금액','이용금액','거래금액'],status:['승인상태','승인/취소구분','상태'],lotte:['롯데계열여부']});
    issue51v11Require_(hi.date>=0&&hi.amount>=0, '카드 원본 승인일/금액 헤더 누락');
    var rawAny={}, rawLotte={};
    for(var h=1;h<hv.length;h++){
      var hr=hv[h], hd=issue51v11Date_(hr[hi.date]), ha=Math.abs(issue51v11Num_(hr[hi.amount]));
      if(!hd||!ha)continue;
      var hs=hi.status>=0?issue51v11Compact_(hr[hi.status]):'';
      if(hs.indexOf('취소')>=0||hs.indexOf('cancel')>=0)continue;
      var hk=hd+'|'+String(Math.round(ha)); rawAny[hk]=true;
      if(hi.lotte>=0 && issue51v11IsLotte_(hr[hi.lotte])) rawLotte[hk]=true;
    }

    issue51v11Require_(pv.length-1===1893, 'preview 주문수 불일치: '+(pv.length-1));
    var transitions={}, changed=0, overlap=0, existingSame=0;
    var previewCounts={MATCHED:0,NON_CARD:0,AMBIGUOUS:0,NO_MATCH:0};
    var newCounts={MATCHED:0,NON_CARD:0,AMBIGUOUS:0,NO_MATCH:0};
    var newRows=0, newNo=0, newNoPurchase=0, newNoZero=0, newNoPaymentBlank=0;
    var newNoMonths={'2026-04':0,'2026-05':0,'2026-06':0};
    var rawCats={LOTTE_EXACT_RAW:0,EXACT_RAW_NO_LOTTE_FLAG:0,NO_EXACT_RAW:0,PURCHASE_ZERO:0};
    var reasons={}, diag=[];

    for(var p=1;p<pv.length;p++){
      var pr=pv[p], key=issue51v11Key_(pr[pi.account],pr[pi.order]);
      if(!key)throw new Error('preview 주문키 공란 row '+(p+1));
      var ns=issue51v11StatusName_(pr[pi.status]);
      previewCounts[ns]=(previewCounts[ns]||0)+1;
      var origin=issue51v11Text_(pr[pi.origin]), purchase=issue51v11Num_(pr[pi.purchase]);
      var date=issue51v11Date_(pr[pi.date]), payment=pi.payment>=0?issue51v11Text_(pr[pi.payment]):'';
      var reason=issue51v11Text_(pr[pi.reason]);

      if(origin==='기존1355'){
        overlap++;
        if(!oldMap[key])throw new Error('기존1355 표시이나 보호 검증키 없음: '+key);
        var os=oldMap[key].status, tk=os+' → '+ns;
        transitions[tk]=(transitions[tk]||0)+1;
        if(os!==ns){
          changed++;
          diag.push(['기존상태변경',pr[pi.account],pr[pi.order],date,Math.round(purchase),os,ns,payment,reason,'']);
        } else existingSame++;
      } else if(origin==='신규538'){
        newRows++; newCounts[ns]=(newCounts[ns]||0)+1;
        var cat='';
        if(ns==='NO_MATCH'){
          newNo++; newNoPurchase+=purchase;
          if(!purchase){newNoZero++;cat='PURCHASE_ZERO';rawCats.PURCHASE_ZERO++;}
          else {
            var raw=issue51v11RawWindow_(date,purchase,rawAny,rawLotte);
            cat=raw;
            rawCats[raw]=(rawCats[raw]||0)+1;
          }
          if(!payment)newNoPaymentBlank++;
          var month=date?date.slice(0,7):''; if(Object.prototype.hasOwnProperty.call(newNoMonths,month))newNoMonths[month]++;
          reasons[reason||'(근거공란)']=(reasons[reason||'(근거공란)']||0)+1;
        }
        diag.push(['신규주문',pr[pi.account],pr[pi.order],date,Math.round(purchase),'',ns,payment,reason,cat]);
      } else throw new Error('알 수 없는 주문구분: '+origin);
    }

    issue51v11Require_(overlap===1355&&newRows===538, '기존/신규 주문수 불일치: '+overlap+'/'+newRows);
    issue51v11Require_(previewCounts.MATCHED===809&&previewCounts.NON_CARD===498&&previewCounts.AMBIGUOUS===0&&previewCounts.NO_MATCH===586,
      'preview 보호 기준 불일치: '+JSON.stringify(previewCounts));
    issue51v11Require_(newCounts.MATCHED===1&&newCounts.NON_CARD===0&&newCounts.AMBIGUOUS===0&&newCounts.NO_MATCH===537,
      '신규538 상태 기준 불일치: '+JSON.stringify(newCounts));

    var out=issue51v11Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    out.clearContents();
    var dh=['구분','쿠팡계정ID','주문번호','주문일','주문매입금액','기존상태','현재preview상태','롯데결제수단','현재매칭근거','raw0~+7일동일금액분류'];
    out.getRange(1,1,1,dh.length).setValues([dh]);
    if(diag.length)out.getRange(2,1,diag.length,dh.length).setValues(diag);
    out.getRange(1,1,1,dh.length).setFontWeight('bold'); out.setFrozenRows(1);

    var rows=[
      ['항목','값'],['버전','v1.1-ISSUE51-DIFF-NEW-NOMATCH-DIAGNOSTIC'],['상태','PASS'],['단계','DONE'],
      ['메시지','기존 1,355 상태차이 및 신규 538 NO_MATCH 원인 분류 완료'],['운영시트 변경','0'],
      ['기존검증주문',oldRows],['현재preview주문',pv.length-1],['기존겹침',overlap],['신규주문',newRows],
      ['기존상태동일',existingSame],['기존상태변경',changed],
      ['기존기준_MATCHED',oldCounts.MATCHED],['기존기준_NON_CARD',oldCounts.NON_CARD],['기존기준_AMBIGUOUS',oldCounts.AMBIGUOUS],['기존기준_NO_MATCH',oldCounts.NO_MATCH],
      ['현재기존1355_MATCHED',808],['현재기존1355_NON_CARD',498],['현재기존1355_AMBIGUOUS',0],['현재기존1355_NO_MATCH',49],
      ['신규538_MATCHED',newCounts.MATCHED],['신규538_NON_CARD',newCounts.NON_CARD],['신규538_AMBIGUOUS',newCounts.AMBIGUOUS],['신규538_NO_MATCH',newCounts.NO_MATCH],
      ['신규NO_MATCH_매입합계',Math.round(newNoPurchase)],['신규NO_MATCH_매입0원',newNoZero],['신규NO_MATCH_결제수단공란',newNoPaymentBlank],
      ['신규NO_MATCH_raw0~+7_롯데표시동일금액있음',rawCats.LOTTE_EXACT_RAW],
      ['신규NO_MATCH_raw0~+7_동일금액있으나롯데표시없음',rawCats.EXACT_RAW_NO_LOTTE_FLAG],
      ['신규NO_MATCH_raw0~+7_동일금액자체없음',rawCats.NO_EXACT_RAW],
      ['신규NO_MATCH_4월',newNoMonths['2026-04']],['신규NO_MATCH_5월',newNoMonths['2026-05']],['신규NO_MATCH_6월',newNoMonths['2026-06']]
    ];
    Object.keys(transitions).sort().forEach(function(k,i){rows.push(['상태전이_'+(i+1),k+' / '+transitions[k]+'건']);});
    Object.keys(reasons).sort(function(a,b){return reasons[b]-reasons[a]||a.localeCompare(b);}).slice(0,10).forEach(function(k,i){rows.push(['신규NO_MATCH사유_'+(i+1),k+' / '+reasons[k]+'건']);});
    rows.push(['완료시각',new Date().toISOString()]);
    issue51v11Status_(state,rows);
    return {ok:true,changed:changed,newNoMatch:newNo,rawCats:rawCats};
  } catch(e) {
    issue51v11Status_(state,[['항목','값'],['버전','v1.1-ISSUE51-DIFF-NEW-NOMATCH-DIAGNOSTIC'],['상태','ERROR'],['단계','FAILED'],['메시지','카드매칭 차이 진단 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0']]);
    throw e;
  }
}

function issue51v11RawWindow_(date,amount,anyMap,lotteMap){
  if(!date||!amount)return 'NO_EXACT_RAW';
  var any=false,lotte=false;
  for(var lag=0;lag<=7;lag++){
    var d=issue51v11Shift_(date,lag),k=d+'|'+String(Math.round(Math.abs(amount)));
    if(anyMap[k])any=true;if(lotteMap[k])lotte=true;
  }
  if(lotte)return 'LOTTE_EXACT_RAW';
  if(any)return 'EXACT_RAW_NO_LOTTE_FLAG';
  return 'NO_EXACT_RAW';
}
function issue51v11Shift_(s,days){var m=String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return '';var d=new Date(Date.UTC(+m[1],+m[2]-1,+m[3]+days));return Utilities.formatDate(d,'UTC','yyyy-MM-dd');}
function issue51v11IsLotte_(v){var s=issue51v11Compact_(v);return s==='y'||s==='yes'||s==='true'||s==='1'||s.indexOf('롯데')>=0;}
function issue51v11StatusName_(v){var s=issue51v11Text_(v).toUpperCase();if(s==='MASTER_MATCHED')return 'MATCHED';if(s==='MATCHED'||s==='NON_CARD'||s==='AMBIGUOUS'||s==='NO_MATCH')return s;return 'NO_MATCH';}
function issue51v11Key_(a,o){a=issue51v11Text_(a).toLowerCase();o=issue51v11NormOrder_(o);return a&&o?a+'|'+o:'';}
function issue51v11Indexes_(h,spec){var o={};Object.keys(spec).forEach(function(k){o[k]=issue51v11Find_(h,spec[k]);});return o;}
function issue51v11Find_(h,names){for(var n=0;n<names.length;n++){var x=issue51v11Compact_(names[n]);for(var i=0;i<h.length;i++)if(issue51v11Compact_(h[i])===x)return i;}return -1;}
function issue51v11Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=issue51v11Text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(m)return m[1]+'-'+issue51v11Pad_(m[2])+'-'+issue51v11Pad_(m[3]);return '';}
function issue51v11Num_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(issue51v11Text_(v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function issue51v11Text_(v){return String(v==null?'':v).trim();}
function issue51v11Compact_(v){return issue51v11Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue51v11NormOrder_(v){return issue51v11Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue51v11Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function issue51v11Require_(ok,msg){if(!ok)throw new Error(msg);}
function issue51v11Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue51v11Status_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);}
