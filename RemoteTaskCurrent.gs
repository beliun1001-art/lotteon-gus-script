/**
 * Issue #50 production apply.
 * Rebuilds ONLY 2026-04-01..2026-06-30 부가세_신고자료 from latest source.
 * Hard-guards source schema/counts/totals before any production write and
 * rolls back the prior sheet values if post-write verification fails.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE50-v1.0-20260812',
  title: '2026년 4~6월 부가세_신고자료 운영 재생성',
  enabled: true,
  statusSheet: 'ISSUE50_운영생성상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue50Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue50WriteStatus_(state, [
    ['항목','값'],
    ['버전','v1.0-ISSUE50-APR-JUN-VAT-PRODUCTION'],
    ['상태','RUNNING'],['단계','PRECHECK'],
    ['메시지','4~6월 부가세_신고자료 운영 재생성 사전검증 시작'],
    ['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
  ]);

  var production = null;
  var oldValues = null;
  var mutated = false;

  try {
    var source = ss.getSheetByName('매출데이터_붙여넣기');
    if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트가 없습니다.');
    var values = source.getDataRange().getValues();
    var headers = values[0] || [];
    if (headers.length < 29) throw new Error('원천 시트가 AC열까지 존재하지 않습니다.');

    var dHeader = issue50Text_(headers[3]);
    var acHeader = issue50Text_(headers[28]);
    if (issue50Compact_(dHeader) !== issue50Compact_('마켓아이디')) throw new Error('D열 헤더 불일치: D=' + dHeader);
    if (issue50Compact_(acHeader) !== issue50Compact_('구매가격')) throw new Error('AC열 헤더 불일치: AC=' + acHeader);

    var ix = issue50Indexes_(headers);
    var out = [];
    var orders = {};
    var accounts = {};
    var stats = {
      rows:0, missingBusiness:0, sales:0, settlement:0, purchase:0,
      salesVat:0, purchaseVat:0, payableVat:0, settlementFallback:0,
      months:{'2026-04':0,'2026-05':0,'2026-06':0}
    };

    for (var r=1; r<values.length; r++) {
      var row = values[r];
      var iso = issue50DateIso_(issue50At_(row, ix.date));
      if (!iso || iso < '2026-04-01' || iso > '2026-06-30') continue;

      var status = ix.status >= 0 ? issue50Text_(row[ix.status]) : '';
      if (/취소|반품|교환|환불/.test(status)) continue;

      var sales = issue50Number_(issue50At_(row, ix.sales));
      if (!sales) continue;

      var account = issue50Text_(row[3]);
      var business = issue50Business_(account);
      if (!business) stats.missingBusiness++;
      accounts[account.toLowerCase()] = true;

      var orderNo = issue50Text_(issue50At_(row, ix.orderNo));
      if (orderNo) orders[account.toLowerCase() + '|' + issue50OrderNorm_(orderNo)] = true;

      var settlementActual = issue50Number_(issue50At_(row, ix.settlement));
      var settlement = settlementActual || Math.round(sales * 0.901);
      if (!settlementActual) stats.settlementFallback++;

      var purchase = issue50Number_(row[28]); // AC=구매가격, prechecked above.
      var salesSplit = issue50SplitVat_(sales);
      var purchaseSplit = issue50SplitVat_(purchase);
      var fee = sales - settlement;
      var profit = settlement - purchase;
      var payable = salesSplit.vat - purchaseSplit.vat;
      var qty = issue50Number_(issue50At_(row, ix.quantity)) || 1;
      var month = iso.slice(0,7);

      stats.rows++;
      stats.sales += sales;
      stats.settlement += settlement;
      stats.purchase += purchase;
      stats.salesVat += salesSplit.vat;
      stats.purchaseVat += purchaseSplit.vat;
      stats.payableVat += payable;
      if (Object.prototype.hasOwnProperty.call(stats.months, month)) stats.months[month]++;

      out.push([
        '2026','상반기',month,iso.slice(5).replace('-','/'),
        account,business,orderNo,
        issue50Text_(issue50At_(row,ix.customer)),
        issue50Text_(issue50At_(row,ix.brand)),
        issue50Text_(issue50At_(row,ix.productNo)),
        issue50Text_(issue50At_(row,ix.productName)),
        qty,sales,salesSplit.supply,salesSplit.vat,settlement,fee,
        purchase,purchaseSplit.supply,purchaseSplit.vat,payable,profit,profit-payable,
        business ? 'ISSUE50 4~6월 최신원천' : '사업자번호 미매핑'
      ]);
    }

    var orderCount = Object.keys(orders).length;
    var accountCount = Object.keys(accounts).filter(function(k){return k;}).length;

    // Hard guards: no production write happens before all pass.
    issue50Assert_(stats.rows === 3894, '생성대상행 불일치: ' + stats.rows + ' != 3894');
    issue50Assert_(stats.missingBusiness === 0, '사업자번호 미매핑 존재: ' + stats.missingBusiness);
    issue50Assert_(accountCount === 4, '계정수 불일치: ' + accountCount + ' != 4');
    issue50Assert_(orderCount === 1893, '고유주문수 불일치: ' + orderCount + ' != 1893');
    issue50Assert_(Math.round(stats.sales) === 207301900, '순수매출합계 불일치: ' + Math.round(stats.sales));
    issue50Assert_(Math.round(stats.settlement) === 184257500, '정산기준금액합계 불일치: ' + Math.round(stats.settlement));
    issue50Assert_(Math.round(stats.purchase) === 106707957, '매입금액합계 불일치: ' + Math.round(stats.purchase));
    issue50Assert_(Math.round(stats.salesVat) === 18845564, '매출부가세합계 불일치: ' + Math.round(stats.salesVat));
    issue50Assert_(Math.round(stats.purchaseVat) === 9700694, '매입부가세합계 불일치: ' + Math.round(stats.purchaseVat));
    issue50Assert_(Math.round(stats.payableVat) === 9144870, '납부예상부가세합계 불일치: ' + Math.round(stats.payableVat));

    var outputHeaders = [
      '신고연도','반기','신고월','날짜','쿠팡계정ID','사업자등록번호','주문번호','고객명','브랜드명','상품번호','상품명','판매수량',
      '순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용','매입금액','매입공급가액','매입부가세',
      '납부예상부가세','예상이익','부가세반영예상이익','비고'
    ];

    production = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
    if (production.getLastRow() > 0 && production.getLastColumn() > 0) {
      oldValues = production.getDataRange().getValues();
    } else {
      oldValues = [];
    }

    issue50WriteStatus_(state, [
      ['항목','값'],['버전','v1.0-ISSUE50-APR-JUN-VAT-PRODUCTION'],['상태','RUNNING'],['단계','WRITE'],
      ['메시지','사전검증 PASS, 부가세_신고자료 운영 쓰기 시작'],['운영시트 변경','쓰기 진행 중'],
      ['사전검증_상세행',stats.rows],['사전검증_고유주문',orderCount],['사전검증_미매핑',stats.missingBusiness],
      ['사전검증_순수매출',Math.round(stats.sales)],['사전검증_매입금액',Math.round(stats.purchase)],['갱신시각',new Date().toISOString()]
    ]);

    production.clearContents();
    mutated = true;
    production.getRange(1,1,1,outputHeaders.length).setValues([outputHeaders]);
    production.getRange(2,1,out.length,outputHeaders.length).setValues(out);
    production.setFrozenRows(1);
    production.getRange(1,1,1,outputHeaders.length).setFontWeight('bold');
    production.getRange(2,12,out.length,1).setNumberFormat('#,##0');
    production.getRange(2,13,out.length,11).setNumberFormat('#,##0');

    // Post-write verification from the actual production sheet.
    var actualHeaders = production.getRange(1,1,1,24).getValues()[0];
    issue50Assert_(issue50Text_(actualHeaders[17]) === '매입금액', 'R 헤더 불일치: ' + issue50Text_(actualHeaders[17]));
    issue50Assert_(issue50Text_(actualHeaders[18]) === '매입공급가액', 'S 헤더 불일치: ' + issue50Text_(actualHeaders[18]));
    issue50Assert_(issue50Text_(actualHeaders[19]) === '매입부가세', 'T 헤더 불일치: ' + issue50Text_(actualHeaders[19]));
    issue50Assert_(production.getLastRow() === 3895, '운영 작성행 불일치: lastRow=' + production.getLastRow());

    var actual = production.getRange(2,1,3894,24).getValues();
    var verifyOrders = {}, verifyAccounts = {}, vSales=0, vSettlement=0, vPurchase=0, vSalesVat=0, vPurchaseVat=0, vPayable=0, vMissing=0;
    for (var i=0;i<actual.length;i++) {
      var a=actual[i];
      var acc=issue50Text_(a[4]), biz=issue50Text_(a[5]), ord=issue50Text_(a[6]);
      if (!biz) vMissing++;
      if (acc) verifyAccounts[acc.toLowerCase()] = true;
      if (ord) verifyOrders[acc.toLowerCase()+'|'+issue50OrderNorm_(ord)] = true;
      vSales += issue50Number_(a[12]);
      vSalesVat += issue50Number_(a[14]);
      vSettlement += issue50Number_(a[15]);
      vPurchase += issue50Number_(a[17]);
      vPurchaseVat += issue50Number_(a[19]);
      vPayable += issue50Number_(a[20]);
    }

    issue50Assert_(Object.keys(verifyOrders).length === 1893, '운영 고유주문 검증 실패: '+Object.keys(verifyOrders).length);
    issue50Assert_(Object.keys(verifyAccounts).length === 4, '운영 계정수 검증 실패: '+Object.keys(verifyAccounts).length);
    issue50Assert_(vMissing === 0, '운영 미매핑 검증 실패: '+vMissing);
    issue50Assert_(Math.round(vSales) === 207301900, '운영 순수매출 검증 실패: '+Math.round(vSales));
    issue50Assert_(Math.round(vSettlement) === 184257500, '운영 정산 검증 실패: '+Math.round(vSettlement));
    issue50Assert_(Math.round(vPurchase) === 106707957, '운영 매입 검증 실패: '+Math.round(vPurchase));
    issue50Assert_(Math.round(vSalesVat) === 18845564, '운영 매출부가세 검증 실패: '+Math.round(vSalesVat));
    issue50Assert_(Math.round(vPurchaseVat) === 9700694, '운영 매입부가세 검증 실패: '+Math.round(vPurchaseVat));
    issue50Assert_(Math.round(vPayable) === 9144870, '운영 납부예상부가세 검증 실패: '+Math.round(vPayable));

    var statusRows = [
      ['항목','값'],
      ['버전','v1.0-ISSUE50-APR-JUN-VAT-PRODUCTION'],['상태','PASS'],['단계','DONE'],
      ['메시지','2026년 4~6월 부가세_신고자료 운영 재생성 및 검증 완료'],
      ['운영시트 변경','부가세_신고자료 1개 재작성'],
      ['작성상세행',3894],['고유주문수',1893],['사업자번호미매핑',0],['계정수',4],
      ['순수매출합계',207301900],['정산기준금액합계',184257500],['매입금액합계',106707957],
      ['매출부가세합계',18845564],['매입부가세합계',9700694],['납부예상부가세합계',9144870],
      ['정산fallback행',stats.settlementFallback],
      ['2026-04_상세행',stats.months['2026-04']],['2026-05_상세행',stats.months['2026-05']],['2026-06_상세행',stats.months['2026-06']],
      ['R헤더',actualHeaders[17]],['S헤더',actualHeaders[18]],['T헤더',actualHeaders[19]],
      ['카드매칭검증 변경','0'],['롤백','없음'],['완료시각',new Date().toISOString()]
    ];
    issue50WriteStatus_(state,statusRows);
    try { MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE50-v1.0',statusRows.map(function(x){return x[0]+': '+x[1];}).join('\n')); } catch(ignore) {}
    return {ok:true,rows:3894,orders:1893,sales:207301900,purchase:106707957};

  } catch (e) {
    var rollback = '불필요';
    if (mutated && production) {
      try {
        production.clearContents();
        if (oldValues && oldValues.length && oldValues[0] && oldValues[0].length) {
          production.getRange(1,1,oldValues.length,oldValues[0].length).setValues(oldValues);
          rollback = '기존 값 복구 완료';
        } else {
          rollback = '기존 데이터 없음 / 빈 시트 복구';
        }
      } catch (rb) {
        rollback = '롤백 실패: ' + String(rb && rb.message ? rb.message : rb);
      }
    }
    issue50WriteStatus_(state,[
      ['항목','값'],['버전','v1.0-ISSUE50-APR-JUN-VAT-PRODUCTION'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','4~6월 부가세_신고자료 운영 재생성 실패'],['오류',String(e&&e.message?e.message:e)],
      ['운영시트 변경',mutated?'쓰기 시도 후 롤백':'0'],['롤백',rollback],['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue50Indexes_(h){
  function f(names,fallback){
    for(var n=0;n<names.length;n++){
      var want=issue50Compact_(names[n]);
      for(var i=0;i<h.length;i++) if(issue50Compact_(h[i])===want) return i;
    }
    return fallback;
  }
  return {
    date:f(['마켓주문일자','주문일자','결제일자','주문일시'],0),
    orderNo:f(['마켓주문번호','주문번호','주문ID','주문ID(마켓)'],2),
    sales:f(['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'],6),
    settlement:f(['정산예정금액(원)','정산예정금액','실제정산금액','정산금액'],-1),
    status:f(['주문상태','상태','클레임상태','처리상태'],-1),
    customer:f(['고객명','수령인','수취인','구매자','주문자'],-1),
    brand:f(['브랜드명','브랜드'],-1),
    productNo:f(['마켓상품번호','상품번호','상품코드','판매자상품코드'],4),
    productName:f(['상품명','상품명(옵션포함)','등록상품명'],-1),
    quantity:f(['판매수량','수량','구매수량'],-1)
  };
}
function issue50Business_(a){
  var s=issue50Text_(a).toLowerCase();
  if(s==='beliun1021'||s==='1021')return '227-27-04928';
  if(s==='beliun1021-1'||s==='1021-1')return '176-71-00758';
  if(s==='beliun1023'||s==='1023')return '835-58-00765';
  if(s==='beliun1024'||s==='1024')return '606-45-93763';
  return '';
}
function issue50DateIso_(v){
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime())) return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');
  var s=issue50Text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);
  if(m)return m[1]+'-'+issue50Pad_(m[2])+'-'+issue50Pad_(m[3]);
  return '';
}
function issue50SplitVat_(amount){var total=Math.round(issue50Number_(amount));var supply=Math.round(total/1.1);return {supply:supply,vat:total-supply};}
function issue50OrderNorm_(v){return issue50Text_(v).replace(/[^0-9A-Za-z가-힣]/g,'').toLowerCase();}
function issue50Number_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(String(v==null?'':v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function issue50Text_(v){return String(v==null?'':v).trim();}
function issue50Compact_(v){return issue50Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue50At_(row,ix){return ix>=0&&ix<row.length?row[ix]:'';}
function issue50Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function issue50Assert_(ok,msg){if(!ok)throw new Error(msg);}
function issue50Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue50WriteStatus_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);}
