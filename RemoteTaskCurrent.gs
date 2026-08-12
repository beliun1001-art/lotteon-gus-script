/**
 * Issue #49 v1.1 read-only diagnostic.
 * Identifies only 2026-04-01..2026-06-30 VAT-eligible rows whose D-column market ID
 * is not mapped to a business registration number. Writes only ISSUE49_* sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE49-v1.1-20260812',
  title: '2026년 4~6월 VAT 사업자번호 미매핑 계정 진단',
  enabled: true,
  outputSheet: 'ISSUE49_사업자미매핑진단',
  statusSheet: 'ISSUE49_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue49v11Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue49v11Write_(state,[['항목','값'],['버전','v1.1-ISSUE49-APR-JUN-UNMAPPED-ACCOUNT'],['상태','RUNNING'],['단계','LOAD'],['메시지','4~6월 사업자번호 미매핑 계정 진단 시작'],['운영시트 변경','0']]);
  try {
    var sh = ss.getSheetByName('매출데이터_붙여넣기');
    if (!sh || sh.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트가 없습니다.');
    var v = sh.getDataRange().getValues(), h = v[0] || [];
    if (h.length < 29) throw new Error('원천 시트가 AC열까지 존재하지 않습니다.');
    if (issue49v11Compact_(h[28]) !== issue49v11Compact_('구매가격')) throw new Error('AC열이 구매가격이 아닙니다: '+issue49v11Text_(h[28]));
    var ix = issue49v11Indexes_(h), groups = {}, eligible = 0, missing = 0;
    for (var r=1;r<v.length;r++) {
      var row=v[r], iso=issue49v11Date_(row[ix.date]);
      if (!iso || iso<'2026-04-01' || iso>'2026-06-30') continue;
      var status=ix.status>=0?issue49v11Text_(row[ix.status]):'';
      if (/취소|반품|환불/.test(status)) continue;
      var sales=issue49v11Num_(row[ix.sales]); if (!sales) continue;
      eligible++;
      var account=issue49v11Text_(row[3]);
      if (issue49v11Business_(account)) continue;
      missing++;
      var key=account||'(공란)';
      if (!groups[key]) groups[key]={account:key,rows:0,orders:{},sales:0,purchase:0,first:'',last:''};
      var g=groups[key]; g.rows++; g.sales+=sales; g.purchase+=issue49v11Num_(row[28]);
      var order=issue49v11Text_(row[ix.order]); if(order)g.orders[order]=true;
      if(!g.first||iso<g.first)g.first=iso; if(!g.last||iso>g.last)g.last=iso;
    }
    var keys=Object.keys(groups).sort();
    var out=keys.map(function(k){var g=groups[k];return [g.account,g.rows,Object.keys(g.orders).length,g.first,g.last,Math.round(g.sales),Math.round(g.purchase)];});
    var o=issue49v11Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet); o.clearContents();
    var oh=['마켓아이디','행수','고유주문수','최초일','최종일','순수매출합계','매입금액합계'];
    o.getRange(1,1,1,oh.length).setValues([oh]); if(out.length)o.getRange(2,1,out.length,oh.length).setValues(out); o.setFrozenRows(1);
    var s=[['항목','값'],['버전','v1.1-ISSUE49-APR-JUN-UNMAPPED-ACCOUNT'],['상태','PASS'],['단계','DONE'],['메시지','4~6월 사업자번호 미매핑 계정 진단 완료'],['운영시트 변경','0'],['4~6월생성대상행',eligible],['사업자번호미매핑행',missing],['미매핑계정수',keys.length]];
    keys.forEach(function(k,i){var g=groups[k];s.push(['미매핑_'+(i+1),g.account+' / 행='+g.rows+' / 주문='+Object.keys(g.orders).length+' / 기간='+g.first+'~'+g.last+' / 매출='+Math.round(g.sales)+' / 매입='+Math.round(g.purchase)]);});
    s.push(['완료시각',new Date().toISOString()]); issue49v11Write_(state,s); return {ok:true,missing:missing,accounts:keys.length};
  } catch(e) {
    issue49v11Write_(state,[['항목','값'],['버전','v1.1-ISSUE49-APR-JUN-UNMAPPED-ACCOUNT'],['상태','ERROR'],['단계','FAILED'],['메시지','미매핑 계정 진단 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0']]); throw e;
  }
}
function issue49v11Indexes_(h){function f(a,d){for(var n=0;n<a.length;n++){var w=issue49v11Compact_(a[n]);for(var i=0;i<h.length;i++)if(issue49v11Compact_(h[i])===w)return i;}return d;}return {date:f(['마켓주문일자','주문일자','결제일자','주문일시'],0),order:f(['마켓주문번호','주문번호','주문ID','주문ID(마켓)'],2),sales:f(['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'],6),status:f(['주문상태','상태','클레임상태','처리상태'],-1)};}
function issue49v11Business_(a){var s=issue49v11Text_(a).toLowerCase();if(s==='beliun1021'||s==='1021')return '227-27-04928';if(s==='beliun1023'||s==='1023')return '835-58-00765';if(s==='beliun1024'||s==='1024')return '606-45-93763';return '';}
function issue49v11Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=issue49v11Text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(m)return m[1]+'-'+issue49v11Pad_(m[2])+'-'+issue49v11Pad_(m[3]);return '';}
function issue49v11Num_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(String(v==null?'':v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function issue49v11Text_(v){return String(v==null?'':v).trim();}
function issue49v11Compact_(v){return issue49v11Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue49v11Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function issue49v11Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue49v11Write_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);}
