/**
 * Issue #52 v1.2 read-only corrected Apr-Jun VAT preview.
 * Root cause under test: v6.48 status aliases omit `마켓주문상태`, so cancel rows can pass.
 * This preview explicitly resolves `마켓주문상태` first and excludes cancel/return/exchange/refund rows.
 * No production sheet is modified.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE52-v1.2-20260813',
  title: '취소상태 제외 4~6월 VAT corrected preview',
  enabled: true,
  outputSheet: 'ISSUE52_취소제외VAT_PREVIEW',
  statusSheet: 'ISSUE52_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var status = i52v12Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  i52v12Write2_(status,[['항목','값'],['버전','v1.2-ISSUE52-CANCEL-EXCLUDED-VAT-PREVIEW'],['상태','RUNNING'],['단계','LOAD'],['메시지','취소상태 제외 4~6월 VAT corrected preview 시작'],['운영시트 변경','0']]);

  try {
    var source = ss.getSheetByName('매출데이터_붙여넣기');
    var current = ss.getSheetByName('부가세_신고자료');
    var oldVerify = ss.getSheetByName('부가세_카드매칭검증');
    if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기가 없습니다.');
    if (!current || current.getLastRow() < 2) throw new Error('부가세_신고자료가 없습니다.');
    if (!oldVerify || oldVerify.getLastRow() < 2) throw new Error('부가세_카드매칭검증이 없습니다.');

    var sv=source.getDataRange().getValues(), sh=sv[0].map(i52v12Text_);
    i52v12Req_(sh.length>=29,'원천 AC열 없음');
    i52v12Req_(i52v12Compact_(sh[3])===i52v12Compact_('마켓아이디'),'D열 헤더 불일치: '+sh[3]);
    i52v12Req_(i52v12Compact_(sh[28])===i52v12Compact_('구매가격'),'AC열 헤더 불일치: '+sh[28]);

    var si=i52v12Indexes_(sh,{
      date:['마켓주문일자','주문일자','결제일자','주문일시'],
      order:['마켓주문번호','주문번호','주문ID','주문ID(마켓)'],
      sales:['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'],
      settlement:['정산예정금액(원)','정산예정금액','실제정산금액','정산금액'],
      status:['마켓주문상태','주문상태','상태','클레임상태','처리상태']
    });
    i52v12Req_(si.date>=0&&si.order>=0&&si.sales>=0&&si.status>=0,'원천 필수 헤더 누락');

    var bizMap={
      'beliun1021':'227-27-04928','1021':'227-27-04928',
      'beliun1021-1':'176-71-00758','1021-1':'176-71-00758',
      'beliun1023':'835-58-00765','1023':'835-58-00765',
      'beliun1024':'606-45-93763','1024':'606-45-93763'
    };

    var dateRows=0,cancelRows=0,salesZeroRows=0,eligibleRows=0,unmapped=0,fallbackRows=0;
    var eligibleKeys={},cancelKeys={},statusRows={},statusOrders={},monthly={},accountRows={};
    var totals={sales:0,settlement:0,purchase:0,salesVat:0,purchaseVat:0,payable:0};
    var detail=[];

    for(var r=1;r<sv.length;r++){
      var row=sv[r], iso=i52v12Date_(row[si.date]);
      if(!iso||iso<'2026-04-01'||iso>'2026-06-30') continue;
      dateRows++;
      var account=i52v12Text_(row[3]).toLowerCase();
      var rawOrder=i52v12Text_(row[si.order]), ord=i52v12NormOrder_(rawOrder);
      var key=account&&ord?account+'|'+ord:'';
      var st=i52v12Text_(row[si.status]);
      var cancelLike=/취소|반품|교환|환불/.test(st);
      if(cancelLike){
        cancelRows++;
        if(key)cancelKeys[key]=true;
        var sk=st||'(상태공란)';statusRows[sk]=(statusRows[sk]||0)+1;
        if(!statusOrders[sk])statusOrders[sk]={};if(key)statusOrders[sk][key]=true;
        continue;
      }
      var sales=i52v12Num_(row[si.sales]);
      if(!sales){salesZeroRows++;continue;}

      var business=bizMap[account]||'';
      if(!business)unmapped++;
      var settlementActual=si.settlement>=0?i52v12Num_(row[si.settlement]):0;
      var settlement=settlementActual||Math.round(sales*0.901);
      if(!settlementActual)fallbackRows++;
      var purchase=i52v12Num_(row[28]);
      var svat=i52v12Split_(sales), pvat=i52v12Split_(purchase);
      var payable=svat.vat-pvat.vat;
      eligibleRows++;
      if(key)eligibleKeys[key]=true;
      accountRows[account]=(accountRows[account]||0)+1;
      totals.sales+=sales;totals.settlement+=settlement;totals.purchase+=purchase;totals.salesVat+=svat.vat;totals.purchaseVat+=pvat.vat;totals.payable+=payable;
      var m=iso.slice(0,7);if(!monthly[m])monthly[m]={rows:0,keys:{},sales:0,purchase:0};
      monthly[m].rows++;if(key)monthly[m].keys[key]=true;monthly[m].sales+=sales;monthly[m].purchase+=purchase;
      detail.push([iso,account,business,rawOrder,st,sales,settlement,purchase,svat.vat,pvat.vat,payable]);
    }

    var eligibleOrderCount=Object.keys(eligibleKeys).length, cancelOrderCount=Object.keys(cancelKeys).length;

    var cv=current.getDataRange().getValues(), ch=cv[0].map(i52v12Text_);
    var ci=i52v12Indexes_(ch,{year:['신고연도'],half:['반기'],account:['쿠팡계정ID'],order:['주문번호'],sales:['순수매출액'],purchase:['매입금액']});
    i52v12Req_(ci.account>=0&&ci.order>=0&&ci.sales>=0&&ci.purchase>=0,'현재 VAT 필수 헤더 누락');
    var currentKeys={},currentRows=0,currentSales=0,currentPurchase=0;
    for(var c=1;c<cv.length;c++){
      var cr=cv[c];
      if(ci.year>=0&&i52v12Text_(cr[ci.year])!=='2026')continue;
      if(ci.half>=0&&i52v12Text_(cr[ci.half])!=='상반기')continue;
      currentRows++;
      var ck=i52v12Key_(cr[ci.account],cr[ci.order]);if(ck)currentKeys[ck]=true;
      currentSales+=i52v12Num_(cr[ci.sales]);currentPurchase+=i52v12Num_(cr[ci.purchase]);
    }
    var currentOnly=0,correctedOnly=0,overlap=0,currentOnlyCancel=0;
    Object.keys(currentKeys).forEach(function(k){if(eligibleKeys[k])overlap++;else{currentOnly++;if(cancelKeys[k])currentOnlyCancel++;}});
    Object.keys(eligibleKeys).forEach(function(k){if(!currentKeys[k])correctedOnly++;});

    var qv=oldVerify.getDataRange().getValues(), qh=qv[0].map(i52v12Text_);
    var qi=i52v12Indexes_(qh,{year:['신고연도'],half:['반기'],account:['쿠팡계정ID'],order:['주문번호']});
    i52v12Req_(qi.account>=0&&qi.order>=0,'기존 카드검증 필수 헤더 누락');
    var oldKeys={};
    for(var q=1;q<qv.length;q++){
      var qr=qv[q];if(qi.year>=0&&i52v12Text_(qr[qi.year])!=='2026')continue;if(qi.half>=0&&i52v12Text_(qr[qi.half])!=='상반기')continue;
      var qk=i52v12Key_(qr[qi.account],qr[qi.order]);if(qk)oldKeys[qk]=true;
    }
    var oldOverlap=0,correctedNew=0,oldOnly=0;
    Object.keys(eligibleKeys).forEach(function(k){if(oldKeys[k])oldOverlap++;else correctedNew++;});
    Object.keys(oldKeys).forEach(function(k){if(!eligibleKeys[k])oldOnly++;});

    var out=i52v12Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    var dh=['날짜','쿠팡계정ID','사업자등록번호','주문번호','마켓주문상태','순수매출액','정산기준금액','매입금액','매출부가세','매입부가세','납부예상부가세'];
    out.clearContents();out.getRange(1,1,1,dh.length).setValues([dh]);if(detail.length)out.getRange(2,1,detail.length,dh.length).setValues(detail);out.getRange(1,1,1,dh.length).setFontWeight('bold');out.setFrozenRows(1);

    var rows=[
      ['항목','값'],['버전','v1.2-ISSUE52-CANCEL-EXCLUDED-VAT-PREVIEW'],['상태','PASS'],['단계','DONE'],
      ['메시지','마켓주문상태 기반 취소 제외 corrected VAT preview 완료'],['운영시트 변경','0'],
      ['상태선택열',i52v12Col_(si.status+1)+' / '+sh[si.status]],
      ['날짜선택열',i52v12Col_(si.date+1)+' / '+sh[si.date]],['주문번호선택열',i52v12Col_(si.order+1)+' / '+sh[si.order]],
      ['매출선택열',i52v12Col_(si.sales+1)+' / '+sh[si.sales]],['정산선택열',si.settlement>=0?i52v12Col_(si.settlement+1)+' / '+sh[si.settlement]:'없음'],['매입선택열','AC / '+sh[28]],
      ['4~6월날짜범위행',dateRows],['취소/반품/교환/환불제외행',cancelRows],['취소상태고유주문',cancelOrderCount],['매출0제외행',salesZeroRows],
      ['corrected상세행',eligibleRows],['corrected고유주문',eligibleOrderCount],['사업자번호미매핑행',unmapped],['정산fallback행',fallbackRows],
      ['corrected순수매출합계',Math.round(totals.sales)],['corrected정산기준금액합계',Math.round(totals.settlement)],['corrected매입금액합계',Math.round(totals.purchase)],
      ['corrected매출부가세합계',Math.round(totals.salesVat)],['corrected매입부가세합계',Math.round(totals.purchaseVat)],['corrected납부예상부가세합계',Math.round(totals.payable)],
      ['현재VAT상세행',currentRows],['현재VAT고유주문',Object.keys(currentKeys).length],['현재VAT순수매출합계',Math.round(currentSales)],['현재VAT매입금액합계',Math.round(currentPurchase)],
      ['현재VAT↔corrected겹침주문',overlap],['현재VAT에만존재',currentOnly],['현재VAT에만존재_취소상태원천',currentOnlyCancel],['corrected에만존재',correctedOnly],
      ['기존1355검증↔corrected겹침',oldOverlap],['corrected신규주문',correctedNew],['기존1355검증에만존재',oldOnly]
    ];
    Object.keys(monthly).sort().forEach(function(m){var x=monthly[m];rows.push([m+'_상세행',x.rows]);rows.push([m+'_고유주문',Object.keys(x.keys).length]);rows.push([m+'_순수매출',Math.round(x.sales)]);rows.push([m+'_매입금액',Math.round(x.purchase)]);});
    Object.keys(accountRows).sort().forEach(function(a){rows.push(['계정_'+a+'_상세행',accountRows[a]]);});
    Object.keys(statusRows).sort(function(a,b){return statusRows[b]-statusRows[a]||a.localeCompare(b);}).slice(0,12).forEach(function(s,i){rows.push(['제외상태_'+(i+1),s+' / 행='+statusRows[s]+' / 주문='+Object.keys(statusOrders[s]||{}).length]);});
    rows.push(['완료시각',new Date().toISOString()]);
    i52v12Write2_(status,rows);
    return {ok:true,correctedRows:eligibleRows,correctedOrders:eligibleOrderCount,cancelOrders:cancelOrderCount,currentOnly:currentOnly,oldOverlap:oldOverlap,correctedNew:correctedNew};
  } catch(e) {
    i52v12Write2_(status,[['항목','값'],['버전','v1.2-ISSUE52-CANCEL-EXCLUDED-VAT-PREVIEW'],['상태','ERROR'],['단계','FAILED'],['메시지','취소 제외 corrected VAT preview 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0']]);
    throw e;
  }
}

function i52v12Indexes_(h,spec){var o={};Object.keys(spec).forEach(function(k){o[k]=i52v12Find_(h,spec[k]);});return o;}
function i52v12Find_(h,names){for(var n=0;n<names.length;n++){var q=i52v12Compact_(names[n]);for(var i=0;i<h.length;i++)if(i52v12Compact_(h[i])===q)return i;}return -1;}
function i52v12Text_(v){return String(v==null?'':v).trim();}
function i52v12Compact_(v){return i52v12Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function i52v12NormOrder_(v){return i52v12Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function i52v12Key_(a,o){var aa=i52v12Text_(a).toLowerCase(),oo=i52v12NormOrder_(o);return aa&&oo?aa+'|'+oo:'';}
function i52v12Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i52v12Text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(m)return m[1]+'-'+i52v12Pad_(m[2])+'-'+i52v12Pad_(m[3]);if(/^\d{2}[.\/-]\d{1,2}$/.test(s)){m=s.match(/^(\d{2})[.\/-](\d{1,2})$/);return '2026-'+i52v12Pad_(m[1])+'-'+i52v12Pad_(m[2]);}return '';}
function i52v12Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function i52v12Num_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(i52v12Text_(v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function i52v12Split_(n){n=i52v12Num_(n);var supply=Math.round(n/1.1);return {supply:supply,vat:n-supply};}
function i52v12Col_(n){var s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
function i52v12Req_(ok,msg){if(!ok)throw new Error(msg);}
function i52v12Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function i52v12Write2_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);}
