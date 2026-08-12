/**
 * Issue #48 v1.3 read-only diagnostic.
 * Finds the current purchase-amount column in 매출데이터_붙여넣기 and compares
 * it with the historical hard-coded AC source-of-truth assumption.
 * Writes only ISSUE48_* diagnostic sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE48-v1.3-20260812',
  title: '상반기 VAT 원천 매입금액 열 위치 진단',
  enabled: true,
  outputSheet: 'ISSUE48_원천매입열진단',
  statusSheet: 'ISSUE48_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue48v13Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue48v13Write_(state,[
    ['항목','값'],
    ['버전','v1.3-ISSUE48-SOURCE-PURCHASE-COLUMN-DIAGNOSTIC'],
    ['상태','RUNNING'],['단계','LOAD'],
    ['메시지','원천 매입금액 열 위치 진단 시작'],
    ['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
  ]);

  try {
    var source = ss.getSheetByName('매출데이터_붙여넣기');
    var vat = ss.getSheetByName('부가세_신고자료');
    var verify = ss.getSheetByName('부가세_카드매칭검증');
    if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트가 없습니다.');
    if (!vat || vat.getLastRow() < 2) throw new Error('부가세_신고자료 시트가 없습니다.');
    if (!verify || verify.getLastRow() < 2) throw new Error('부가세_카드매칭검증 시트가 없습니다.');

    var sourceLastRow = source.getLastRow(), sourceLastCol = source.getLastColumn();
    var sh = source.getRange(1,1,1,sourceLastCol).getValues()[0] || [];
    var sv = source.getRange(2,1,sourceLastRow-1,sourceLastCol).getValues();

    var dateIx = issue48v13FindFirst_(sh,['마켓주문일자','주문일자','결제일자','주문일시'],0);
    var orderIx = issue48v13FindFirst_(sh,['마켓주문번호','주문번호','주문ID','주문ID(마켓)'],2);
    var salesIx = issue48v13FindFirst_(sh,['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'],6);
    var statusIx = issue48v13FindFirst_(sh,['주문상태','상태','클레임상태','처리상태'],-1);
    var accountIx = 3; // production v6.48 fixed D-column account source
    var acIx = 28;
    if (sourceLastCol <= acIx) throw new Error('원천 시트에 AC열이 없습니다.');

    // Build current 2026 H1 order/account target set from VAT detail.
    var vv = vat.getDataRange().getValues(), vh = vv[0] || [];
    var vy = issue48v13FindFirst_(vh,['신고연도'],-1);
    var vhf = issue48v13FindFirst_(vh,['반기'],-1);
    var va = issue48v13FindFirst_(vh,['쿠팡계정ID','마켓아이디','계정ID'],-1);
    var vo = issue48v13FindFirst_(vh,['주문번호','마켓주문번호','주문ID','주문ID(마켓)'],-1);
    if (vy<0 || vhf<0 || va<0 || vo<0) throw new Error('VAT H1 target 헤더 누락');
    var h1Targets = {}, vatH1Rows = 0;
    for (var vr=1;vr<vv.length;vr++) {
      if (issue48v13Text_(vv[vr][vy]) !== '2026' || issue48v13Text_(vv[vr][vhf]) !== '상반기') continue;
      vatH1Rows++;
      var vk = issue48v13TargetKey_(vv[vr][va], vv[vr][vo]);
      if (vk) h1Targets[vk] = true;
    }

    // Current card-verification total is kept only as a comparison reference.
    var cv = verify.getDataRange().getValues(), ch = cv[0] || [];
    var cy = issue48v13FindFirst_(ch,['신고연도'],-1);
    var chf = issue48v13FindFirst_(ch,['반기'],-1);
    var cp = issue48v13FindFirst_(ch,['주문매입금액','매입금액'],-1);
    if (cy<0 || chf<0 || cp<0) throw new Error('카드검증 기준 헤더 누락');
    var verifyPurchase = 0, verifyOrders = 0;
    for (var cr=1;cr<cv.length;cr++) {
      if (issue48v13Text_(cv[cr][cy]) !== '2026' || issue48v13Text_(cv[cr][chf]) !== '상반기') continue;
      verifyOrders++;
      verifyPurchase += issue48v13Number_(cv[cr][cp]);
    }

    // Find exact purchase aliases and broader purchase/cost-looking headers.
    var exactAliases = ['매입금액','구매가격','매입가격','상품매입금액','결제금액(매입)','구매금액','매입가','원가'];
    var candidateMap = {};
    for (var c=0;c<sh.length;c++) {
      var h = issue48v13Text_(sh[c]);
      var compact = issue48v13Compact_(h);
      var exact = exactAliases.some(function(a){return compact === issue48v13Compact_(a);});
      var broad = /매입|구매.*가격|원가|purchase|cost/i.test(h);
      if (exact || broad) candidateMap[c] = true;
    }
    candidateMap[acIx] = true;
    var candidateIxs = Object.keys(candidateMap).map(Number).sort(function(a,b){return a-b;});

    var stats = {};
    candidateIxs.forEach(function(ix){
      stats[ix] = {
        ix:ix, col:issue48v13Col_(ix+1), header:issue48v13Text_(sh[ix]),
        allNonblank:0, allSum:0, allMax:0, allGt1e9:0,
        h1Nonblank:0, h1Sum:0, h1Max:0, h1Gt1e9:0, h1Reasonable:0, h1Zero:0,
        h1Number:0, h1String:0
      };
    });

    var h1SourceRows = 0, h1OrderTargetRows = 0, samples = [];
    for (var r=0;r<sv.length;r++) {
      var row = sv[r];
      var key = issue48v13TargetKey_(row[accountIx], row[orderIx]);
      var inTarget = !!(key && h1Targets[key]);
      if (inTarget) h1OrderTargetRows++;
      var status = statusIx>=0 ? issue48v13Text_(row[statusIx]) : '';
      var sales = issue48v13Number_(row[salesIx]);
      var productionEligible = inTarget && !/취소|반품|환불/.test(status) && !!sales;
      if (productionEligible) h1SourceRows++;

      candidateIxs.forEach(function(ix){
        var s = stats[ix], raw = row[ix], text = issue48v13Text_(raw), n = issue48v13Number_(raw), abs = Math.abs(n);
        if (text !== '') s.allNonblank++;
        s.allSum += n;
        if (abs > s.allMax) s.allMax = abs;
        if (abs > 1e9) s.allGt1e9++;
        if (!productionEligible) return;
        if (text !== '') s.h1Nonblank++;
        s.h1Sum += n;
        if (abs > s.h1Max) s.h1Max = abs;
        if (abs > 1e9) s.h1Gt1e9++;
        if (n > 0 && n <= 10000000) s.h1Reasonable++;
        if (n === 0) s.h1Zero++;
        if (typeof raw === 'number' && isFinite(raw)) s.h1Number++;
        else if (typeof raw === 'string') s.h1String++;
      });

      if (productionEligible && samples.length < 30) {
        var sample = [r+2,issue48v13Date_(row[dateIx]),issue48v13Text_(row[accountIx]),issue48v13Text_(row[orderIx]),status,sales];
        candidateIxs.forEach(function(ix){sample.push(row[ix]);});
        samples.push(sample);
      }
    }

    var candidates = candidateIxs.map(function(ix){
      var s = stats[ix];
      s.diffVerify = Math.round(s.h1Sum - verifyPurchase);
      s.plausible = s.h1Sum > 0 && s.h1Gt1e9 === 0 && s.h1Max <= 10000000;
      return s;
    });
    var plausible = candidates.filter(function(s){return s.plausible;}).sort(function(a,b){return Math.abs(a.diffVerify)-Math.abs(b.diffVerify);});
    var best = plausible.length ? plausible[0] : null;

    var outputHeaders = ['원천행','주문일','D열계정','주문번호','상태','순수매출액'];
    candidateIxs.forEach(function(ix){outputHeaders.push(issue48v13Col_(ix+1)+' / '+issue48v13Text_(sh[ix]));});
    var output = issue48v13Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet);
    output.clearContents();
    output.getRange(1,1,1,outputHeaders.length).setValues([outputHeaders]);
    if (samples.length) output.getRange(2,1,samples.length,outputHeaders.length).setValues(samples);
    output.setFrozenRows(1);
    output.getRange(1,1,1,outputHeaders.length).setBackground('#d9eaf7').setFontWeight('bold');

    var nearHeaders = [];
    for (var ni=22;ni<Math.min(sh.length,35);ni++) nearHeaders.push(issue48v13Col_(ni+1)+'='+issue48v13Text_(sh[ni]));

    var statusRows = [
      ['항목','값'],
      ['버전','v1.3-ISSUE48-SOURCE-PURCHASE-COLUMN-DIAGNOSTIC'],
      ['상태','PASS'],['단계','DONE'],
      ['메시지','원천 매입금액 열 위치 진단 완료'],
      ['운영시트 변경','0'],
      ['원천전체열수',sourceLastCol],
      ['원천전체데이터행',sourceLastRow-1],
      ['VAT상반기상세행',vatH1Rows],
      ['원천H1주문키매칭행',h1OrderTargetRows],
      ['원천H1생성대상행',h1SourceRows],
      ['검증상반기주문',verifyOrders],
      ['검증기준총매입금액',Math.round(verifyPurchase)],
      ['AC현재열','AC / '+issue48v13Text_(sh[acIx])],
      ['AC_H1합계',Math.round(stats[acIx].h1Sum)],
      ['AC_H1_10억초과',stats[acIx].h1Gt1e9],
      ['AC_H1_1천만원이하양수',stats[acIx].h1Reasonable],
      ['원천주변헤더_W_AI',nearHeaders.join(' | ')],
      ['매입후보열수',candidateIxs.length]
    ];
    candidates.forEach(function(s,i){
      statusRows.push(['후보_'+(i+1),s.col+' / '+(s.header||'(공란)')+' / H1합계='+Math.round(s.h1Sum)+' / 검증차이='+s.diffVerify+' / H1>10억='+s.h1Gt1e9+' / H1<=1천만양수='+s.h1Reasonable+' / H1최대='+Math.round(s.h1Max)]);
    });
    statusRows.push(['유력후보',best ? best.col+' / '+best.header+' / H1합계='+Math.round(best.h1Sum)+' / 검증차이='+best.diffVerify : '없음']);
    statusRows.push(['완료시각',new Date().toISOString()]);
    issue48v13Write_(state,statusRows);
    try { MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE48-v1.3',statusRows.map(function(x){return x[0]+': '+x[1];}).join('\n')); } catch(mailError) {}
    return {ok:true,best:best ? {col:best.col,header:best.header,h1Sum:best.h1Sum,diff:best.diffVerify} : null};
  } catch(e) {
    issue48v13Write_(state,[
      ['항목','값'],['버전','v1.3-ISSUE48-SOURCE-PURCHASE-COLUMN-DIAGNOSTIC'],
      ['상태','ERROR'],['단계','FAILED'],['메시지','원천 매입금액 열 위치 진단 실패'],
      ['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue48v13FindFirst_(headers,names,fallback){
  for(var n=0;n<names.length;n++){
    var wanted=issue48v13Compact_(names[n]);
    for(var i=0;i<headers.length;i++) if(issue48v13Compact_(headers[i])===wanted) return i;
  }
  return fallback;
}
function issue48v13TargetKey_(account,order){
  var a=issue48v13Text_(account).toLowerCase(), o=issue48v13OrderNorm_(order);
  return a&&o ? a+'|'+o : '';
}
function issue48v13OrderNorm_(v){return issue48v13Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue48v13Compact_(v){return issue48v13Text_(v).toLowerCase().replace(/\s/g,'');}
function issue48v13Text_(v){return v==null?'':String(v).trim();}
function issue48v13Number_(v){
  if(typeof v==='number') return isNaN(v)?0:v;
  var n=Number(String(v==null?'0':v).replace(/[원,\s]/g,''));
  return isNaN(n)?0:n;
}
function issue48v13Date_(v){
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime())) return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');
  return issue48v13Text_(v);
}
function issue48v13Col_(n){var s='';while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
function issue48v13Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function issue48v13Write_(sheet,rows){
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  SpreadsheetApp.flush();
}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
