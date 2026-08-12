/**
 * Issue #52 v1.0 read-only source completeness diagnostic.
 * Investigates 528 new NO_MATCH orders with purchase=0 and 9 positive-purchase
 * NO_MATCH orders from Issue #51. No production sheet is modified.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE52-v1.0-20260813',
  title: '신규 538주문 매입0원 원천 완전성 진단',
  enabled: true,
  outputSheet: 'ISSUE52_매입0원원천진단',
  statusSheet: 'ISSUE52_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue52Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue52WriteStatus_(state,[['항목','값'],['버전','v1.0-ISSUE52-ZERO-PURCHASE-SOURCE-COMPLETENESS'],['상태','RUNNING'],['단계','LOAD'],['메시지','신규 매입0원 원천 완전성 진단 시작'],['운영시트 변경','0']]);

  try {
    var preview = ss.getSheetByName('ISSUE51_카드매칭전체PREVIEW');
    var source = ss.getSheetByName('매출데이터_붙여넣기');
    var history = ss.getSheetByName('카드사용내역_붙여넣기');
    if (!preview || preview.getLastRow() < 2) throw new Error('ISSUE51_카드매칭전체PREVIEW가 없습니다.');
    if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기가 없습니다.');
    if (!history || history.getLastRow() < 2) throw new Error('카드사용내역_붙여넣기가 없습니다.');

    var pv=preview.getDataRange().getValues(), ph=pv[0].map(issue52Text_);
    var pi=issue52Indexes_(ph,{account:['쿠팡계정ID'],order:['주문번호'],date:['주문일'],purchase:['주문매입금액'],origin:['주문구분'],status:['카드매칭상태']});
    issue52Require_(pi.account>=0&&pi.order>=0&&pi.date>=0&&pi.purchase>=0&&pi.origin>=0&&pi.status>=0,'Issue51 preview 필수 헤더 누락');

    var zeroTargets={}, positiveTargets=[], zeroMonths={}, zeroAccounts={}, positiveSum=0, newOrders=0, newNo=0;
    for(var p=1;p<pv.length;p++){
      var pr=pv[p];
      if(issue52Text_(pr[pi.origin])!=='신규538')continue;
      newOrders++;
      if(issue52Status_(pr[pi.status])!=='NO_MATCH')continue;
      newNo++;
      var key=issue52Key_(pr[pi.account],pr[pi.order]);
      var d=issue52Date_(pr[pi.date]), amt=issue52Num_(pr[pi.purchase]);
      if(!amt){
        zeroTargets[key]={account:issue52Text_(pr[pi.account]),order:issue52Text_(pr[pi.order]),date:d,rows:0,acSum:0,blank:0,numZero:0,textZero:0,positive:0,negative:0,nonNumeric:0};
        var m=d?d.slice(0,7):'(날짜공란)'; zeroMonths[m]=(zeroMonths[m]||0)+1;
        var a=issue52Text_(pr[pi.account])||'(계정공란)'; zeroAccounts[a]=(zeroAccounts[a]||0)+1;
      } else {
        positiveTargets.push({key:key,account:issue52Text_(pr[pi.account]),order:issue52Text_(pr[pi.order]),date:d,purchase:amt});
        positiveSum+=amt;
      }
    }
    issue52Require_(newOrders===538,'신규 주문수 불일치: '+newOrders);
    issue52Require_(newNo===537,'신규 NO_MATCH 수 불일치: '+newNo);
    issue52Require_(Object.keys(zeroTargets).length===528,'매입0원 주문수 불일치: '+Object.keys(zeroTargets).length);
    issue52Require_(positiveTargets.length===9,'양수 매입 NO_MATCH 수 불일치: '+positiveTargets.length);
    issue52Require_(Math.round(positiveSum)===714440,'양수 매입 NO_MATCH 합계 불일치: '+Math.round(positiveSum));

    var sv=source.getDataRange().getValues(), sh=sv[0].map(issue52Text_);
    issue52Require_(sh.length>=29,'원천이 AC열까지 존재하지 않습니다.');
    issue52Require_(issue52Compact_(sh[3])===issue52Compact_('마켓아이디'),'D열 헤더 불일치: '+sh[3]);
    issue52Require_(issue52Compact_(sh[28])===issue52Compact_('구매가격'),'AC열 헤더 불일치: '+sh[28]);
    var si=issue52Indexes_(sh,{date:['마켓주문일자','주문일자','결제일자','주문일시'],order:['마켓주문번호','주문번호','주문ID','주문ID(마켓)'],sales:['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'],status:['마켓주문상태','주문상태','상태','클레임상태','처리상태']});
    issue52Require_(si.date>=0&&si.order>=0&&si.sales>=0,'원천 날짜/주문번호/매출 헤더 누락');

    var candidateCols=issue52CandidateCols_(sh), candidateStats={};
    candidateCols.forEach(function(c){candidateStats[c.index]={header:c.header,nonblank:0,numeric:0,positive:0,sum:0,max:0};});
    var matchedKeys={}, sourceTargetRows=0, acBlank=0, acNumZero=0, acTextZero=0, acPositive=0, acNegative=0, acNonNumeric=0;

    for(var r=1;r<sv.length;r++){
      var row=sv[r], iso=issue52Date_(row[si.date]);
      if(!iso||iso<'2026-04-01'||iso>'2026-06-30')continue;
      var st=si.status>=0?issue52Text_(row[si.status]):'';
      if(/취소|반품|교환|환불/.test(st))continue;
      if(!issue52Num_(row[si.sales]))continue;
      var key=issue52Key_(row[3],row[si.order]);
      if(!zeroTargets[key])continue;
      sourceTargetRows++; matchedKeys[key]=true;
      var z=zeroTargets[key]; z.rows++;
      var raw=row[28], rawText=issue52Text_(raw), n=issue52ParseMaybe_(raw);
      if(rawText===''){z.blank++;acBlank++;}
      else if(n===null){z.nonNumeric++;acNonNumeric++;}
      else if(n===0){if(typeof raw==='number'){z.numZero++;acNumZero++;}else{z.textZero++;acTextZero++;}}
      else if(n>0){z.positive++;z.acSum+=n;acPositive++;}
      else {z.negative++;z.acSum+=n;acNegative++;}

      candidateCols.forEach(function(c){
        var v=row[c.index], txt=issue52Text_(v), cs=candidateStats[c.index];
        if(txt!=='')cs.nonblank++;
        var cn=issue52ParseMaybe_(v);
        if(cn!==null){cs.numeric++;if(cn>0){cs.positive++;cs.sum+=cn;if(cn>cs.max)cs.max=cn;}}
      });
    }

    var zeroKeys=Object.keys(zeroTargets), matchedZero=Object.keys(matchedKeys).length, unmatchedZero=zeroKeys.length-matchedZero;
    var zeroAcPositiveOrders=0, zeroAcNonzeroSum=0;
    zeroKeys.forEach(function(k){var z=zeroTargets[k];if(z.acSum!==0){zeroAcPositiveOrders++;zeroAcNonzeroSum+=z.acSum;}});

    var hv=history.getDataRange().getValues(), hh=hv[0].map(issue52Text_);
    var hi=issue52Indexes_(hh,{date:['승인일','이용일','거래일','사용일','승인일자']});
    issue52Require_(hi.date>=0,'카드 원본 승인일 헤더 누락');
    var histMin='',histMax='';
    for(var h=1;h<hv.length;h++){
      var hd=issue52Date_(hv[h][hi.date]); if(!hd)continue;
      if(!histMin||hd<histMin)histMin=hd;if(!histMax||hd>histMax)histMax=hd;
    }
    var posMin='',posMax='',windowBeyond=0,windowCovered=0,posMonths={};
    positiveTargets.forEach(function(x){
      if(!posMin||x.date<posMin)posMin=x.date;if(!posMax||x.date>posMax)posMax=x.date;
      var m=x.date?x.date.slice(0,7):'(날짜공란)';posMonths[m]=(posMonths[m]||0)+1;
      x.windowEnd=issue52Shift_(x.date,7);
      x.coverageGap=!!(histMax&&x.windowEnd>histMax);
      if(x.coverageGap)windowBeyond++;else windowCovered++;
    });

    var out=issue52Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet), detail=[];
    zeroKeys.sort().forEach(function(k){var z=zeroTargets[k];detail.push(['매입0원',z.account,z.order,z.date,0,z.rows,Math.round(z.acSum),z.blank,z.numZero+z.textZero,z.positive,'','','']);});
    positiveTargets.forEach(function(x){detail.push(['양수매입_NO_MATCH',x.account,x.order,x.date,Math.round(x.purchase),'','','','','',x.windowEnd,histMax,x.coverageGap?'CARD_HISTORY_GAP_POSSIBLE':'HISTORY_WINDOW_COVERED']);});
    var dh=['구분','쿠팡계정ID','주문번호','주문일','preview주문매입금액','원천상세행','원천AC합계','AC공란셀','AC0셀','AC양수셀','+7일창끝','카드원본최종일','판정'];
    out.clearContents();out.getRange(1,1,1,dh.length).setValues([dh]);if(detail.length)out.getRange(2,1,detail.length,dh.length).setValues(detail);out.getRange(1,1,1,dh.length).setFontWeight('bold');out.setFrozenRows(1);

    var rows=[
      ['항목','값'],['버전','v1.0-ISSUE52-ZERO-PURCHASE-SOURCE-COMPLETENESS'],['상태','PASS'],['단계','DONE'],
      ['메시지','신규 매입0원 528건 원천 완전성 및 양수 9건 카드기간 진단 완료'],['운영시트 변경','0'],
      ['신규주문',newOrders],['신규NO_MATCH',newNo],['매입0원주문',zeroKeys.length],['양수매입NO_MATCH',positiveTargets.length],['양수매입NO_MATCH합계',Math.round(positiveSum)],
      ['원천0원주문키매칭',matchedZero],['원천0원주문키미매칭',unmatchedZero],['원천0원대상상세행',sourceTargetRows],
      ['AC헤더','AC / '+sh[28]],['AC공란셀',acBlank],['AC숫자0셀',acNumZero],['AC텍스트0셀',acTextZero],['AC양수셀',acPositive],['AC음수셀',acNegative],['AC비숫자셀',acNonNumeric],
      ['preview0원이지만원천AC합계비0주문',zeroAcPositiveOrders],['해당원천AC합계',Math.round(zeroAcNonzeroSum)],
      ['금액후보열수',candidateCols.length]
    ];
    candidateCols.forEach(function(c,i){var cs=candidateStats[c.index];rows.push(['금액후보_'+(i+1),issue52Col_(c.index+1)+' / '+c.header+' / 양수셀='+cs.positive+' / 양수합계='+Math.round(cs.sum)+' / 최대='+Math.round(cs.max)]);});
    Object.keys(zeroMonths).sort().forEach(function(k){rows.push(['매입0원_'+k,zeroMonths[k]+'주문']);});
    Object.keys(zeroAccounts).sort().forEach(function(k){rows.push(['매입0원계정_'+k,zeroAccounts[k]+'주문']);});
    rows.push(['카드원본기간',(histMin||'')+'~'+(histMax||'')],['양수9건주문기간',(posMin||'')+'~'+(posMax||'')],['양수9건_+7일창카드원본최종일초과',windowBeyond],['양수9건_카드원본기간내',windowCovered],['7월카드내역필요가능성',windowBeyond>0?'있음 / '+windowBeyond+'건의 +7일창이 현재 카드원본 최종일을 초과':'현재 기간정보상 낮음']);
    Object.keys(posMonths).sort().forEach(function(k){rows.push(['양수NO_MATCH_'+k,posMonths[k]+'주문']);});
    rows.push(['완료시각',new Date().toISOString()]);
    issue52WriteStatus_(state,rows);
    return {ok:true,zeroOrders:zeroKeys.length,matchedZero:matchedZero,positiveNoMatch:positiveTargets.length,historyGap:windowBeyond};
  } catch(e) {
    issue52WriteStatus_(state,[['항목','값'],['버전','v1.0-ISSUE52-ZERO-PURCHASE-SOURCE-COMPLETENESS'],['상태','ERROR'],['단계','FAILED'],['메시지','신규 매입0원 원천 진단 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0']]);
    throw e;
  }
}

function issue52CandidateCols_(h){var out=[], seen={};var exact=['구매가격','구매금액','매입금액','매입가','매입가격','매입원가','원가','원가금액','purchaseprice','purchaseamount','cost'];for(var i=0;i<h.length;i++){var c=issue52Compact_(h[i]);for(var j=0;j<exact.length;j++){if(c===issue52Compact_(exact[j])){if(!seen[i]){seen[i]=true;out.push({index:i,header:h[i]});}break;}}}return out;}
function issue52Indexes_(h,spec){var o={};Object.keys(spec).forEach(function(k){o[k]=issue52Find_(h,spec[k]);});return o;}
function issue52Find_(h,names){for(var n=0;n<names.length;n++){var w=issue52Compact_(names[n]);for(var i=0;i<h.length;i++)if(issue52Compact_(h[i])===w)return i;}return -1;}
function issue52Key_(a,o){var aa=issue52Text_(a).toLowerCase(),oo=issue52Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return aa&&oo?aa+'|'+oo:'';}
function issue52Status_(v){var s=issue52Text_(v).toUpperCase();return s==='MATCHED'||s==='MASTER_MATCHED'?'MATCHED':s==='NON_CARD'?'NON_CARD':s==='AMBIGUOUS'?'AMBIGUOUS':'NO_MATCH';}
function issue52Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=issue52Text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(m)return m[1]+'-'+issue52Pad_(m[2])+'-'+issue52Pad_(m[3]);if(/^\d{2}[.\/-]\d{1,2}$/.test(s)){var q=s.match(/^(\d{2})[.\/-](\d{1,2})$/);return '2026-'+issue52Pad_(q[1])+'-'+issue52Pad_(q[2]);}return '';}
function issue52Shift_(s,days){var m=String(s||'').match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return '';var d=new Date(Date.UTC(+m[1],+m[2]-1,+m[3]+days));return Utilities.formatDate(d,'UTC','yyyy-MM-dd');}
function issue52ParseMaybe_(v){if(typeof v==='number')return isFinite(v)?v:null;var s=issue52Text_(v);if(s==='')return null;var n=Number(s.replace(/[원,%\s,]/g,''));return isFinite(n)?n:null;}
function issue52Num_(v){var n=issue52ParseMaybe_(v);return n===null?0:n;}
function issue52Text_(v){return String(v==null?'':v).trim();}
function issue52Compact_(v){return issue52Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue52Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function issue52Col_(n){var s='';while(n>0){n--;s=String.fromCharCode(65+n%26)+s;n=Math.floor(n/26);}return s;}
function issue52Require_(ok,msg){if(!ok)throw new Error(msg);}
function issue52Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue52WriteStatus_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);}
