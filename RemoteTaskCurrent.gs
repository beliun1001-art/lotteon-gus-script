/**
 * Issue #52 v1.1 read-only robust source join diagnostic.
 * v1.0 matched 0/528 by account+order, so its AC conclusions are invalid.
 * This version traces the 528 zero-purchase preview orders against the source
 * by normalized order number first, then diagnoses account/date/status/sales differences.
 * No production sheet is modified.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE52-v1.1-20260813',
  title: '신규 매입0원 528건 원천 주문번호 재조인 진단',
  enabled: true,
  outputSheet: 'ISSUE52_원천재조인진단',
  statusSheet: 'ISSUE52_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var status = i52v11Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  i52v11Write_(status,[['항목','값'],['버전','v1.1-ISSUE52-ROBUST-SOURCE-ORDER-JOIN'],['상태','RUNNING'],['단계','LOAD'],['메시지','528건 원천 주문번호 재조인 진단 시작'],['운영시트 변경','0']]);

  try {
    var preview = ss.getSheetByName('ISSUE51_카드매칭전체PREVIEW');
    var source = ss.getSheetByName('매출데이터_붙여넣기');
    if (!preview || preview.getLastRow() < 2) throw new Error('ISSUE51_카드매칭전체PREVIEW가 없습니다.');
    if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기가 없습니다.');

    var pv=preview.getDataRange().getValues(), ph=pv[0].map(i52v11Text_);
    var pi=i52v11Indexes_(ph,{account:['쿠팡계정ID'],order:['주문번호'],date:['주문일'],purchase:['주문매입금액'],origin:['주문구분'],status:['카드매칭상태']});
    i52v11Req_(pi.account>=0&&pi.order>=0&&pi.date>=0&&pi.purchase>=0&&pi.origin>=0&&pi.status>=0,'preview 필수 헤더 누락');

    var targets={}, byOrder={}, targetCount=0;
    for(var p=1;p<pv.length;p++){
      var pr=pv[p];
      if(i52v11Text_(pr[pi.origin])!=='신규538') continue;
      if(i52v11Status_(pr[pi.status])!=='NO_MATCH') continue;
      if(i52v11Num_(pr[pi.purchase])!==0) continue;
      var acc=i52v11Text_(pr[pi.account]).toLowerCase();
      var rawOrder=i52v11Text_(pr[pi.order]);
      var ord=i52v11NormOrder_(rawOrder);
      var key=acc+'|'+ord;
      i52v11Req_(acc&&ord,'매입0원 target 주문키 공란');
      i52v11Req_(!targets[key],'매입0원 target 중복키: '+key);
      var t={key:key,account:acc,rawOrder:rawOrder,order:ord,date:i52v11Date_(pr[pi.date]),eligibleRows:[],allRows:[]};
      targets[key]=t;
      if(!byOrder[ord]) byOrder[ord]=[];
      byOrder[ord].push(t);
      targetCount++;
    }
    i52v11Req_(targetCount===528,'매입0원 target 수 불일치: '+targetCount);

    var targetOrderDup=0;
    Object.keys(byOrder).forEach(function(o){ if(byOrder[o].length>1) targetOrderDup++; });

    var sv=source.getDataRange().getValues(), sh=sv[0].map(i52v11Text_);
    i52v11Req_(sh.length>=29,'원천 AC열 없음');
    i52v11Req_(i52v11Compact_(sh[3])===i52v11Compact_('마켓아이디'),'D열 헤더 불일치: '+sh[3]);
    i52v11Req_(i52v11Compact_(sh[28])===i52v11Compact_('구매가격'),'AC열 헤더 불일치: '+sh[28]);
    var si=i52v11Indexes_(sh,{date:['마켓주문일자','주문일자','결제일자','주문일시'],order:['마켓주문번호','주문번호','주문ID','주문ID(마켓)'],sales:['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'],status:['마켓주문상태','주문상태','상태','클레임상태','처리상태']});
    i52v11Req_(si.date>=0&&si.order>=0&&si.sales>=0,'원천 날짜/주문번호/매출 헤더 누락');

    var sourceOrderAccounts={}, sourceEligibleOrderAccounts={};
    for(var r=1;r<sv.length;r++){
      var row=sv[r], ord=i52v11NormOrder_(row[si.order]);
      if(!ord || !byOrder[ord]) continue;
      var acc=i52v11Text_(row[3]).toLowerCase();
      var iso=i52v11Date_(row[si.date]);
      var st=si.status>=0?i52v11Text_(row[si.status]):'';
      var sales=i52v11Num_(row[si.sales]);
      var eligible=!!(iso&&iso>='2026-04-01'&&iso<='2026-06-30'&&!(/취소|반품|교환|환불/.test(st))&&sales!==0);
      var acRaw=row[28], ac=i52v11Maybe_(acRaw);
      var rec={rowNo:r+1,account:acc,date:iso,status:st,sales:sales,acRaw:i52v11Text_(acRaw),ac:ac,eligible:eligible};
      byOrder[ord].forEach(function(t){t.allRows.push(rec);if(eligible)t.eligibleRows.push(rec);});
      if(!sourceOrderAccounts[ord])sourceOrderAccounts[ord]={}; sourceOrderAccounts[ord][acc]=true;
      if(eligible){if(!sourceEligibleOrderAccounts[ord])sourceEligibleOrderAccounts[ord]={};sourceEligibleOrderAccounts[ord][acc]=true;}
    }

    var exactEligible=0, orderEligible=0, orderAny=0, noAny=0, eligibleAccountMismatch=0, eligibleAccountSame=0;
    var eligibleMultiAccount=0, eligibleUniqueAccount=0, anyOnlyExcluded=0;
    var acZeroOrders=0, acPositiveOrders=0, acNegativeOrders=0, acUnknownOrders=0, acPositiveSum=0;
    var acBlankCells=0, acZeroCells=0, acPositiveCells=0, acNegativeCells=0, acNonNumericCells=0;
    var accountPairs={}, exclusionCats={}, detail=[];

    Object.keys(targets).sort().forEach(function(k){
      var t=targets[k], er=t.eligibleRows, ar=t.allRows;
      if(ar.length)orderAny++;else{noAny++;exclusionCats.NO_SOURCE_ORDER=(exclusionCats.NO_SOURCE_ORDER||0)+1;}
      if(er.length){
        orderEligible++;
        var accounts={}; er.forEach(function(x){accounts[x.account]=true;});
        var accts=Object.keys(accounts);
        if(accts.length===1)eligibleUniqueAccount++;else eligibleMultiAccount++;
        var same=er.filter(function(x){return x.account===t.account;});
        if(same.length){exactEligible++;eligibleAccountSame++;}
        else{
          eligibleAccountMismatch++;
          var pair=t.account+' → '+(accts.sort().join(',')||'(없음)'); accountPairs[pair]=(accountPairs[pair]||0)+1;
        }

        var acSum=0, known=0, posCells=0, zeroCells=0, negCells=0, blankCells=0, nonCells=0;
        er.forEach(function(x){
          if(x.acRaw===''){blankCells++;acBlankCells++;return;}
          if(x.ac===null){nonCells++;acNonNumericCells++;return;}
          known++; acSum+=x.ac;
          if(x.ac>0){posCells++;acPositiveCells++;}
          else if(x.ac<0){negCells++;acNegativeCells++;}
          else{zeroCells++;acZeroCells++;}
        });
        if(!known)acUnknownOrders++;
        else if(acSum>0){acPositiveOrders++;acPositiveSum+=acSum;}
        else if(acSum<0)acNegativeOrders++;
        else acZeroOrders++;
        detail.push(['ORDER_ELIGIBLE',t.account,t.rawOrder,t.date,er.length,accts.join(','),Math.round(acSum),blankCells,zeroCells,posCells,nonCells,'']);
      } else if(ar.length){
        anyOnlyExcluded++;
        var cats={};
        ar.forEach(function(x){
          var c='OTHER_EXCLUDED';
          if(!x.date||x.date<'2026-04-01'||x.date>'2026-06-30')c='DATE_OUTSIDE';
          else if(/취소|반품|교환|환불/.test(x.status))c='CANCEL_STATUS';
          else if(x.sales===0)c='SALES_ZERO';
          cats[c]=true;
        });
        var names=Object.keys(cats).sort();
        names.forEach(function(c){exclusionCats[c]=(exclusionCats[c]||0)+1;});
        detail.push(['SOURCE_ONLY_EXCLUDED',t.account,t.rawOrder,t.date,0,Object.keys(sourceOrderAccounts[t.order]||{}).sort().join(','),0,0,0,0,0,names.join(',')]);
      }
    });

    var out=i52v11Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    var dh=['분류','preview계정','주문번호','preview주문일','원천생성대상상세행','원천계정','원천AC합계','AC공란셀','AC0셀','AC양수셀','AC비숫자셀','제외사유'];
    out.clearContents();out.getRange(1,1,1,dh.length).setValues([dh]);if(detail.length)out.getRange(2,1,detail.length,dh.length).setValues(detail);out.getRange(1,1,1,dh.length).setFontWeight('bold');out.setFrozenRows(1);

    var rows=[
      ['항목','값'],['버전','v1.1-ISSUE52-ROBUST-SOURCE-ORDER-JOIN'],['상태','PASS'],['단계','DONE'],
      ['메시지','528건 원천 주문번호 재조인 및 AC 실값 진단 완료'],['운영시트 변경','0'],
      ['매입0원대상주문',targetCount],['target주문번호중복',targetOrderDup],
      ['원천전체에서주문번호발견',orderAny],['원천전체에서도주문번호없음',noAny],
      ['4~6월생성대상주문번호매칭',orderEligible],['4~6월생성대상없고원천다른행만존재',anyOnlyExcluded],
      ['계정+주문번호정확매칭',exactEligible],['주문번호매칭_계정동일',eligibleAccountSame],['주문번호매칭_계정불일치',eligibleAccountMismatch],
      ['주문번호매칭_원천계정1개',eligibleUniqueAccount],['주문번호매칭_원천계정복수',eligibleMultiAccount],
      ['AC합계0원주문',acZeroOrders],['AC합계양수주문',acPositiveOrders],['AC합계음수주문',acNegativeOrders],['AC판정불가주문',acUnknownOrders],['AC양수주문합계',Math.round(acPositiveSum)],
      ['AC공란셀',acBlankCells],['AC0셀',acZeroCells],['AC양수셀',acPositiveCells],['AC음수셀',acNegativeCells],['AC비숫자셀',acNonNumericCells]
    ];
    Object.keys(accountPairs).sort(function(a,b){return accountPairs[b]-accountPairs[a]||a.localeCompare(b);}).forEach(function(k,i){rows.push(['계정불일치_'+(i+1),k+' / '+accountPairs[k]+'주문']);});
    Object.keys(exclusionCats).sort().forEach(function(k,i){rows.push(['원천제외사유_'+(i+1),k+' / '+exclusionCats[k]+'주문']);});
    rows.push(['완료시각',new Date().toISOString()]);
    i52v11Write_(status,rows);
    return {ok:true,target:targetCount,orderEligible:orderEligible,exactEligible:exactEligible,acZeroOrders:acZeroOrders,acPositiveOrders:acPositiveOrders};
  } catch(e) {
    i52v11Write_(status,[['항목','값'],['버전','v1.1-ISSUE52-ROBUST-SOURCE-ORDER-JOIN'],['상태','ERROR'],['단계','FAILED'],['메시지','528건 원천 재조인 진단 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0']]);
    throw e;
  }
}

function i52v11Indexes_(h,spec){var o={};Object.keys(spec).forEach(function(k){o[k]=i52v11Find_(h,spec[k]);});return o;}
function i52v11Find_(h,names){for(var n=0;n<names.length;n++){var q=i52v11Compact_(names[n]);for(var i=0;i<h.length;i++)if(i52v11Compact_(h[i])===q)return i;}return -1;}
function i52v11Text_(v){return String(v==null?'':v).trim();}
function i52v11Compact_(v){return i52v11Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function i52v11NormOrder_(v){return i52v11Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function i52v11Status_(v){var s=i52v11Text_(v).toUpperCase();return s==='MATCHED'||s==='MASTER_MATCHED'?'MATCHED':s==='NON_CARD'?'NON_CARD':s==='AMBIGUOUS'?'AMBIGUOUS':'NO_MATCH';}
function i52v11Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i52v11Text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(m)return m[1]+'-'+i52v11Pad_(m[2])+'-'+i52v11Pad_(m[3]);if(/^\d{2}[.\/-]\d{1,2}$/.test(s)){m=s.match(/^(\d{2})[.\/-](\d{1,2})$/);return '2026-'+i52v11Pad_(m[1])+'-'+i52v11Pad_(m[2]);}return '';}
function i52v11Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function i52v11Maybe_(v){if(typeof v==='number')return isFinite(v)?v:null;var s=i52v11Text_(v);if(s==='')return null;var n=Number(s.replace(/[원,%\s,]/g,''));return isFinite(n)?n:null;}
function i52v11Num_(v){var n=i52v11Maybe_(v);return n===null?0:n;}
function i52v11Req_(ok,msg){if(!ok)throw new Error(msg);}
function i52v11Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function i52v11Write_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);}
