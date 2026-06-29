var LOTTEON_PATCH_V636_UNSETTLED_20DAY_SIMPLE_LOADED = true;
var LOTTEON_V636_MIN_DAYS = 20;

var __baseVatFull_v636 = typeof generateVatReportsFullSeparated_v622 === 'function' ? generateVatReportsFullSeparated_v622 : null;
var __baseDashFast_v636 = typeof refreshDashboardFastOnly === 'function' ? refreshDashboardFastOnly : null;
var __baseDisplay_v636 = typeof applyDisplayStandardsOnlyFast_v623 === 'function' ? applyDisplayStandardsOnlyFast_v623 : null;

buildUnsettledSettlementByAccountSheet_v629_ = function(salesAgg) { return rebuildUnsettled20DaySimple_v636_(salesAgg || buildSingleSourceSalesAgg_v628_()); };
buildUnsettledSettlementByAccountSheet_v616_ = function(salesAgg) { return rebuildUnsettled20DaySimple_v636_(salesAgg || buildSingleSourceSalesAgg_v628_()); };

generateVatReportsFullSeparated_v622 = function() { var r = __baseVatFull_v636 ? __baseVatFull_v636.apply(this, arguments) : null; rebuildUnsettled20DaySimple_v636_(buildSingleSourceSalesAgg_v628_()); return r || {ok:true}; };
refreshDashboardFastOnly = function() { var r = __baseDashFast_v636 ? __baseDashFast_v636.apply(this, arguments) : null; rebuildUnsettled20DaySimple_v636_(buildSingleSourceSalesAgg_v628_()); return r || {ok:true}; };
applyDisplayStandardsOnlyFast_v623 = function() { var r = __baseDisplay_v636 ? __baseDisplay_v636.apply(this, arguments) : null; rebuildUnsettled20DaySimple_v636_(buildSingleSourceSalesAgg_v628_()); try { SpreadsheetApp.getUi().alert('표시서식 정리 완료\n\n미정산_쿠팡계정별은 주문일 기준 20일 이상 지난 미정산 주문만 표시합니다.'); } catch(e) {} return r || {ok:true}; };

function rebuildUnsettled20DaySimple_v636_(salesAgg) {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName('미정산_쿠팡계정별') || ss.insertSheet('미정산_쿠팡계정별');
  var accountMap = sourceRowAccountMap_v636_();
  var today = new Date();
  var rows = [];
  var skipUnder20 = 0;
  var skipNoDate = 0;
  var missing = 0;

  (salesAgg && salesAgg.detailRows || []).forEach(function(d) {
    if (n636_(d.actualSettlement)) return;
    if (!d.orderDate) { skipNoDate++; return; }
    var elapsed = Math.floor((today.getTime() - d.orderDate.getTime()) / 86400000);
    if (elapsed < LOTTEON_V636_MIN_DAYS) { skipUnder20++; return; }
    var cp = s636_(d.accountId);
    if (!cp || cp === '계정미확인') cp = accountMap[d.sourceRow] || '';
    if (!cp) { cp = '계정미확인'; missing++; }
    var sales = n636_(d.netSales);
    var est = n636_(d.estimatedSettlement) || Math.round(sales * 0.901);
    var buy = n636_(d.purchase);
    var profit = n636_(d.estimatedProfit) || (est - buy);
    rows.push([cp, elapsed >= 30 ? '30일초과 미정산' : '20일이상 미정산', d.dateText || Utilities.formatDate(d.orderDate, CONFIG.TIMEZONE, 'MM/dd'), elapsed, d.orderNo || '', d.brand || '', d.customer || '고객명미확인', d.addressGroup || '', d.productNo || '', short636_(d.productName || '', 80), sales, est, buy, profit, '20일 이상 미정산 기준']);
  });

  rows.sort(function(a,b){ var da=dk636_(a[2]), db=dk636_(b[2]); if(da!==db)return da-db; var ca=String(a[0]).localeCompare(String(b[0])); if(ca!==0)return ca; return String(a[4]).localeCompare(String(b[4])); });
  write636_(sheet, ['쿠팡계정ID','구분','주문일','경과일','주문번호','브랜드명','고객명','주소구분','상품번호','상품명','순수매출액','예상정산금액','매입금액','예상이익','비고'], rows);
  fmt636_(sheet, rows.length);
  check636_(rows.length, skipUnder20, skipNoDate, missing);
  try { log_('patch_v636_unsettled_20day_simple', 'rows=' + rows.length + ' under20=' + skipUnder20 + ' nodate=' + skipNoDate + ' missing=' + missing); } catch(e) {}
  return {ok:true, rows:rows.length};
}

function sourceRowAccountMap_v636_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName(CONFIG.SHEETS.SALES_IN) || ss.getSheetByName('매출데이터_붙여넣기');
  var out = {};
  if (!sh || sh.getLastRow() < 2) return out;
  var v = sh.getDataRange().getValues();
  var h = v[0].map(function(x){ return String(x || '').trim(); });
  var idx = {
    raw: h.indexOf('원본계정ID'),
    ana: h.indexOf('분석계정ID'),
    cp: h.indexOf('쿠팡계정ID'),
    id: h.indexOf('계정ID'),
    no: h.indexOf('계정번호'),
    filter: h.indexOf('대표검색필터명')
  };
  for (var r=1; r<v.length; r++) {
    var acc = firstAccount636_([cell636_(v[r],idx.raw), cell636_(v[r],idx.ana), cell636_(v[r],idx.cp), cell636_(v[r],idx.id)]);
    if (!acc) acc = noMap636_(cell636_(v[r],idx.no));
    if (!acc) acc = filterMap636_(cell636_(v[r],idx.filter));
    if (acc) out[r+1] = acc;
  }
  return out;
}

function firstAccount636_(arr) { for (var i=0;i<arr.length;i++){ var x=s636_(arr[i]); if (/^beliun\d{4}$/i.test(x)) return x; } return ''; }
function noMap636_(x) { x=s636_(x); if(x==='1')return 'beliun1021'; if(x==='2')return 'beliun1024'; if(x==='3')return 'beliun1023'; return ''; }
function filterMap636_(x) { x=s636_(x); var m=x.match(/롯백[_\s-]*(0?1|0?2|0?3|0?4)/); if(!m)return ''; var c=String(m[1]).replace(/^0/,''); if(c==='1')return 'beliun1021'; if(c==='2')return 'beliun1024'; if(c==='3')return 'beliun1023'; if(c==='4')return 'beliun1024'; return ''; }
function cell636_(row, idx) { return idx >= 0 ? row[idx] : ''; }
function s636_(x) { return String(x == null ? '' : x).replace(/,/g,'').trim(); }
function n636_(x) { if(typeof toNumber_==='function')return toNumber_(x); if(typeof x==='number')return x; var n=Number(String(x==null?'':x).replace(/₩|,|%/g,'')); return isNaN(n)?0:n; }
function dk636_(x) { if(typeof orderDateSortKey_v635_==='function')return orderDateSortKey_v635_(x); var m=String(x||'').match(/^(\d{1,2})[-.\/](\d{1,2})$/); if(!m)return 99999999; return Number(Utilities.formatDate(new Date(), CONFIG.TIMEZONE, 'yyyy') + ('0'+Number(m[1])).slice(-2) + ('0'+Number(m[2])).slice(-2)); }
function short636_(t,n){ t=String(t||'').replace(/\s+/g,' ').trim(); return t.length>n?t.slice(0,n-1)+'…':t; }
function write636_(sh, headers, rows){ sh.clearContents(); sh.getRange(1,1,1,headers.length).setValues([headers]); if(rows.length) sh.getRange(2,1,rows.length,headers.length).setValues(rows); }
function fmt636_(sh, rows){ var c=sh.getLastColumn(); try{ sh.getRange(1,1,1,c).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center'); if(rows){ sh.getRange(2,1,rows,c).setVerticalAlignment('middle'); [1,2,5,6,7,8,9,10,15].forEach(function(x){if(x<=c)sh.getRange(2,x,rows,1).setHorizontalAlignment('left');}); [3].forEach(function(x){if(x<=c)sh.getRange(2,x,rows,1).setHorizontalAlignment('center').setNumberFormat('@');}); [4,11,12,13,14].forEach(function(x){if(x<=c)sh.getRange(2,x,rows,1).setHorizontalAlignment('right').setNumberFormat('#,##0');}); } sh.setColumnWidth(1,110); sh.setColumnWidth(2,120); sh.setColumnWidth(5,125); sh.setColumnWidth(10,260); sh.setFrozenRows(1); }catch(e){} }
function check636_(shown, under20, nodate, missing){ var sh=SpreadsheetApp.getActive().getSheetByName('미정산_정리검증')||SpreadsheetApp.getActive().insertSheet('미정산_정리검증'); var rows=[['항목','값','메모'],['표시기준','20일 이상 미정산','주문일 경과일 >= 20'],['표시행수',shown,'표시된 주문'],['20일미만 제외',under20,'미정산이나 20일 미만'],['주문일 없음 제외',nodate,'경과일 계산 불가'],['계정미확인 행수',missing,'원본 계정 컬럼 확인 필요']]; sh.clearContents(); sh.getRange(1,1,rows.length,3).setValues(rows); try{sh.getRange(1,1,1,3).setBackground('#d9eaf7').setFontWeight('bold'); sh.autoResizeColumns(1,3);}catch(e){} }
