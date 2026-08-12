/**
 * Issue #49 standalone read-only preview.
 * Rebuilds 2026-04-01 ~ 2026-06-30 VAT detail from the latest source using
 * production v6.48 semantics, but writes only ISSUE49_* preview/status sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE49-v1.0-20260812',
  title: '2026년 4~6월 VAT 신고자료 최신 원천 미리보기',
  enabled: true,
  outputSheet: 'ISSUE49_4_6월_VAT미리보기',
  statusSheet: 'ISSUE49_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue49Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue49WriteStatus_(state,[
    ['항목','값'],
    ['버전','v1.0-ISSUE49-APR-JUN-VAT-PREVIEW'],
    ['상태','RUNNING'],['단계','LOAD'],
    ['메시지','2026년 4~6월 VAT 미리보기 시작'],
    ['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
  ]);

  try {
    var source = ss.getSheetByName('매출데이터_붙여넣기');
    if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트가 없습니다.');

    var lastRow = source.getLastRow();
    var lastCol = source.getLastColumn();
    var values = source.getRange(1,1,lastRow,lastCol).getValues();
    var headers = values[0] || [];
    if (headers.length < 29) throw new Error('원천 시트가 AC열까지 존재하지 않습니다.');

    var ix = issue49Indexes_(headers);
    var dHeader = issue49Text_(headers[3]);
    var acHeader = issue49Text_(headers[28]);
    if (issue49Compact_(acHeader) !== issue49Compact_('구매가격')) {
      throw new Error('AC열 헤더가 구매가격이 아닙니다: AC=' + acHeader);
    }

    var previewHeaders = [
      '원천행','날짜','신고연도','반기','신고월','쿠팡계정ID','사업자등록번호','주문번호','주문번호정규화',
      '고객명','브랜드명','상품번호','상품명','판매수량','순수매출액','매출공급가액','매출부가세',
      '정산기준금액','정산방식','마켓수수료/비용','매입금액','매입공급가액','매입부가세',
      '납부예상부가세','예상이익','부가세반영예상이익','롯데결제수단','원천주문상태'
    ];

    var out = [];
    var uniqueOrders = {};
    var stats = {
      totalRows: values.length - 1,
      dateParsed: 0,
      rangeRows: 0,
      cancelExcluded: 0,
      salesZeroExcluded: 0,
      finalRows: 0,
      accountMissing: 0,
      businessMissing: 0,
      settlementFallback: 0,
      purchaseZero: 0,
      paymentBlank: 0,
      sales: 0,
      settlement: 0,
      purchase: 0,
      salesVat: 0,
      purchaseVat: 0,
      payableVat: 0,
      months: {
        '2026-04': {rows:0,orders:{},sales:0,purchase:0,paymentBlank:0},
        '2026-05': {rows:0,orders:{},sales:0,purchase:0,paymentBlank:0},
        '2026-06': {rows:0,orders:{},sales:0,purchase:0,paymentBlank:0}
      }
    };

    for (var r=1; r<values.length; r++) {
      var row = values[r];
      var iso = issue49DateIso_(issue49At_(row, ix.date));
      if (iso) stats.dateParsed++;
      if (!iso || iso < '2026-04-01' || iso > '2026-06-30') continue;
      stats.rangeRows++;

      var status = issue49Text_(issue49At_(row, ix.status));
      if (/취소|반품|환불/.test(status)) {
        stats.cancelExcluded++;
        continue;
      }

      var sales = issue49Number_(issue49At_(row, ix.sales));
      if (!sales) {
        stats.salesZeroExcluded++;
        continue;
      }

      // Production v6.48 semantics: D is account source, AC is purchase source-of-truth.
      var account = issue49Text_(row[3]);
      var business = issue49BusinessNo_(account);
      var orderNo = issue49Text_(issue49At_(row, ix.orderNo));
      var orderNorm = issue49OrderNorm_(orderNo);
      var settlementActual = issue49Number_(issue49At_(row, ix.settlement));
      var settlement = settlementActual || Math.round(sales * 0.901);
      var settlementMethod = settlementActual ? '원천정산금액' : '매출*0.901 fallback';
      var purchase = issue49Number_(row[28]);
      var salesSplit = issue49SplitVat_(sales);
      var purchaseSplit = issue49SplitVat_(purchase);
      var fee = sales - settlement;
      var profit = settlement - purchase;
      var payable = salesSplit.vat - purchaseSplit.vat;
      var payment = issue49Text_(issue49At_(row, ix.payment));
      var qty = issue49Number_(issue49At_(row, ix.quantity)) || 1;
      var month = iso.slice(0,7);

      stats.finalRows++;
      if (!account) stats.accountMissing++;
      if (!business) stats.businessMissing++;
      if (!settlementActual) stats.settlementFallback++;
      if (!purchase) stats.purchaseZero++;
      if (!payment) stats.paymentBlank++;
      stats.sales += sales;
      stats.settlement += settlement;
      stats.purchase += purchase;
      stats.salesVat += salesSplit.vat;
      stats.purchaseVat += purchaseSplit.vat;
      stats.payableVat += payable;

      var orderKey = [business, account.toLowerCase(), orderNorm || ('BLANK@'+(r+1))].join('|');
      uniqueOrders[orderKey] = true;
      if (stats.months[month]) {
        var ms = stats.months[month];
        ms.rows++;
        ms.orders[orderKey] = true;
        ms.sales += sales;
        ms.purchase += purchase;
        if (!payment) ms.paymentBlank++;
      }

      out.push([
        r+1,iso,'2026','상반기',month,account,business,orderNo,orderNorm,
        issue49Text_(issue49At_(row,ix.customer)),issue49Text_(issue49At_(row,ix.brand)),
        issue49Text_(issue49At_(row,ix.productNo)),issue49Text_(issue49At_(row,ix.productName)),qty,
        sales,salesSplit.supply,salesSplit.vat,settlement,settlementMethod,fee,purchase,purchaseSplit.supply,purchaseSplit.vat,
        payable,profit,profit-payable,payment,status
      ]);
    }

    if (!stats.finalRows) throw new Error('2026년 4~6월 최종 생성 대상 행이 0건입니다. 날짜/헤더를 확인하세요.');

    var output = issue49Ensure_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    output.clearContents();
    output.getRange(1,1,1,previewHeaders.length).setValues([previewHeaders]);
    output.getRange(2,1,out.length,previewHeaders.length).setValues(out);
    output.setFrozenRows(1);
    output.getRange(1,1,1,previewHeaders.length).setBackground('#d9eaf7').setFontWeight('bold');
    [15,16,17,18,20,21,22,23,24,25,26].forEach(function(c){output.getRange(2,c,out.length,1).setNumberFormat('#,##0');});
    output.getRange(2,1,out.length,1).setNumberFormat('0');

    var statusRows = [
      ['항목','값'],
      ['버전','v1.0-ISSUE49-APR-JUN-VAT-PREVIEW'],
      ['상태','PASS'],['단계','DONE'],
      ['메시지','2026년 4~6월 VAT 최신 원천 미리보기 완료'],
      ['운영시트 변경','0'],
      ['원천전체행',stats.totalRows],
      ['날짜파싱성공행',stats.dateParsed],
      ['4~6월날짜범위행',stats.rangeRows],
      ['취소/반품/환불제외행',stats.cancelExcluded],
      ['매출0제외행',stats.salesZeroExcluded],
      ['최종상세행',stats.finalRows],
      ['고유주문수',Object.keys(uniqueOrders).length],
      ['계정미확인',stats.accountMissing],
      ['사업자번호미매핑',stats.businessMissing],
      ['정산fallback행',stats.settlementFallback],
      ['매입금액0행',stats.purchaseZero],
      ['결제수단공란행',stats.paymentBlank],
      ['순수매출합계',Math.round(stats.sales)],
      ['정산기준금액합계',Math.round(stats.settlement)],
      ['매입금액합계',Math.round(stats.purchase)],
      ['매출부가세합계',Math.round(stats.salesVat)],
      ['매입부가세합계',Math.round(stats.purchaseVat)],
      ['납부예상부가세합계',Math.round(stats.payableVat)],
      ['D계정기준열','D / '+dHeader],
      ['AC매입기준열','AC / '+acHeader],
      ['날짜선택열',issue49Col_(ix.date+1)+' / '+issue49Text_(headers[ix.date])],
      ['매출선택열',issue49Col_(ix.sales+1)+' / '+issue49Text_(headers[ix.sales])],
      ['정산선택열',ix.settlement>=0 ? issue49Col_(ix.settlement+1)+' / '+issue49Text_(headers[ix.settlement]) : '없음 / fallback 사용'],
      ['주문번호선택열',issue49Col_(ix.orderNo+1)+' / '+issue49Text_(headers[ix.orderNo])],
      ['결제수단선택열',ix.payment>=0 ? issue49Col_(ix.payment+1)+' / '+issue49Text_(headers[ix.payment]) : '없음']
    ];

    ['2026-04','2026-05','2026-06'].forEach(function(m){
      var x=stats.months[m];
      statusRows.push([m+'_상세행',x.rows]);
      statusRows.push([m+'_고유주문',Object.keys(x.orders).length]);
      statusRows.push([m+'_순수매출',Math.round(x.sales)]);
      statusRows.push([m+'_매입금액',Math.round(x.purchase)]);
      statusRows.push([m+'_결제수단공란행',x.paymentBlank]);
    });
    statusRows.push(['완료시각',new Date().toISOString()]);
    issue49WriteStatus_(state,statusRows);
    try {
      MailApp.sendEmail('beliun1001@gmail.com','[LOTTEON 자동작업 결과][PASS] ISSUE49-v1.0',statusRows.map(function(x){return x[0]+': '+x[1];}).join('\n'));
    } catch (mailError) {}
    return {ok:true,rows:stats.finalRows,orders:Object.keys(uniqueOrders).length,purchase:Math.round(stats.purchase)};
  } catch(e) {
    issue49WriteStatus_(state,[
      ['항목','값'],['버전','v1.0-ISSUE49-APR-JUN-VAT-PREVIEW'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','2026년 4~6월 VAT 미리보기 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue49Indexes_(headers){
  function find(names,fallback){
    for(var n=0;n<names.length;n++){
      var wanted=issue49Compact_(names[n]);
      for(var i=0;i<headers.length;i++) if(issue49Compact_(headers[i])===wanted) return i;
    }
    return fallback;
  }
  var x={
    date:find(['마켓주문일자','주문일자','결제일자','주문일시'],0),
    orderNo:find(['마켓주문번호','주문번호','주문ID','주문ID(마켓)'],2),
    sales:find(['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'],6),
    settlement:find(['정산예정금액(원)','정산예정금액','실제정산금액','정산금액'],-1),
    status:find(['주문상태','상태','클레임상태','처리상태'],-1),
    customer:find(['고객명','수령인','수취인','구매자','주문자'],-1),
    brand:find(['브랜드명','브랜드'],-1),
    productNo:find(['마켓상품번호','상품번호','상품코드','판매자상품코드'],4),
    productName:find(['상품명','상품명(옵션포함)','등록상품명'],-1),
    quantity:find(['판매수량','수량','구매수량'],-1),
    payment:find(['결제수단','결제정보','결제방법','카드사','결제수단/카드사','결제수단(카드사)','구매결제수단'],-1)
  };
  if(x.date<0||x.orderNo<0||x.sales<0) throw new Error('날짜/주문번호/매출 필수 헤더를 찾지 못했습니다.');
  return x;
}
function issue49BusinessNo_(marketId){
  var s=issue49Text_(marketId).toLowerCase();
  if(s==='beliun1021'||s==='1021') return '227-27-04928';
  if(s==='beliun1023'||s==='1023') return '835-58-00765';
  if(s==='beliun1024'||s==='1024') return '606-45-93763';
  return '';
}
function issue49SplitVat_(amount){var total=Math.round(issue49Number_(amount));var supply=Math.round(total/1.1);return {supply:supply,vat:total-supply};}
function issue49At_(row,index){return index>=0&&index<row.length?row[index]:'';}
function issue49Number_(v){if(typeof v==='number')return isNaN(v)?0:v;var n=Number(String(v==null?'':v).replace(/[원,%\s]/g,''));return isNaN(n)?0:n;}
function issue49Text_(v){return String(v==null?'':v).trim();}
function issue49Compact_(v){return issue49Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue49OrderNorm_(v){return issue49Text_(v).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');}
function issue49DateIso_(v){
  if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime())) return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');
  var s=issue49Text_(v); if(!s)return '';
  var m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/); if(m)return m[1]+'-'+issue49Pad_(m[2])+'-'+issue49Pad_(m[3]);
  var m2=s.match(/^(\d{4})(\d{2})(\d{2})/); if(m2)return m2[1]+'-'+m2[2]+'-'+m2[3];
  return '';
}
function issue49Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function issue49Col_(n){var s='';while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
function issue49Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function issue49WriteStatus_(sheet,rows){sheet.clearContents();sheet.getRange(1,1,rows.length,2).setValues(rows);sheet.setFrozenRows(1);sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');SpreadsheetApp.flush();}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
