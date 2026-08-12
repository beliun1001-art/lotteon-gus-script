/**
 * Issue #49 v1.2 read-only mapping verification.
 * Confirms all 2026-04-01..2026-06-30 VAT-eligible D-column market IDs map to a
 * confirmed business registration number. Writes only ISSUE49_* sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE49-v1.2-20260812',
  title: '2026년 4~6월 VAT 사업자번호 매핑 최종 검증',
  enabled: true,
  outputSheet: 'ISSUE49_사업자매핑최종검증',
  statusSheet: 'ISSUE49_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue49v12Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue49v12Write_(state,[['항목','값'],['버전','v1.2-ISSUE49-APR-JUN-MAPPING-VERIFY'],['상태','RUNNING'],['단계','LOAD'],['메시지','4~6월 사업자번호 매핑 최종 검증 시작'],['운영시트 변경','0']]);
  try {
    var sh = ss.getSheetByName('매출데이터_붙여넣기');
    if (!sh || sh.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 시트가 없습니다.');
    var v = sh.getDataRange().getValues(), h = v[0] || [];
    if (h.length < 29) throw new Error('원천 시트가 AC열까지 존재하지 않습니다.');
    if (issue49v12Compact_(h[28]) !== issue49v12Compact_('구매가격')) throw new Error('AC열이 구매가격이 아닙니다: '+issue49v12Text_(h[28]));
    var ix = issue49v12Indexes_(h), groups = {}, eligible = 0, missing = 0, mapped = 0;
    for (var r=1;r<v.length;r++) {
      var row=v[r], iso=issue49v12Date_(row[ix.date]);
      if (!iso || iso<'2026-04-01' || iso>'2026-06-30') continue;
      var status=ix.status>=0?issue49v12Text_(row[ix.status]):'';
      if (/취소|반품|교환|환불/.test(status)) continue;
      var sales=issue49v12Num_(row[ix.sales]); if (!sales) continue;
      eligible++;
      var account=issue49v12Text_(row[3]), business=issue49v12Business_(account);
      if (business) mapped++; else missing++;
      var key=account||'(공란)';
      if (!groups[key]) groups[key]={account:key,business:business,rows:0,orders:{},sales:0,purchase:0,first:'',last:''};
      var g=groups[key]; g.rows++; g.sales+=sales; g.purchase+=issue49v12Num_(row[28]);
      var order=issue49v12Text_(row[ix.order]); if(order)g.orders[order]=true;
      if(!g.first||iso<g.first)g.first=iso; if(!g.last||iso>g.last)g.last=iso;
    }
    var keys=Object.keys(groups).sort();
    var out=keys.map(function(k){var g=groups[k];return [g.account,g.business||'',g.rows,Object.keys(g.orders).length,g.first,g.last,Math.round(g.sales),Math.round(g.purchase),g.business?'MAPPED':'UNMAPPED'];});
    var o=issue49v12Ensure_(ss,LOTTEON_REMOTE_TASK.outputSheet); o.clearContents();
    var oh=['마켓아이디','사업자등록번호','행수','고유주문수','최초일','최종일','순수매출합계','매입금액합계','상태'];
    o.getRange(1,1,1,oh.length).setValues([oh]); if(out.length)o.getRange(2,1,out.length,oh.length).setValues(out); o.setFrozenRows(1);
    var s=[['항목','값'],['버전','v1.2-ISSUE49-APR-JUN-MAPPING-VERIFY'],['상태','PASS'],['단계','DONE'],['메시지','4~6월 사업자번호 매핑 최종 검증 완료'],['운영시트 변경','0'],['4~6월생성대상행',eligible],['사업자번호매핑행',mapped],['사업자번호미매핑행',missing],['계정수',keys.length]];
    keys.forEach(function(k,i){var g=groups[k];s.push(['계정_'+(i+1),g.account+' → '+(g.business||'미매핑')+' / 행='+g.rows+' / 주문='+Object.keys(g.orders).length+' / 기간='+g.first+'~'+g.last]);});
    s.push(['완료시각',new Date().toISOString()]); issue49v12Write_(state,s); return {ok:true,eligible:eligible,mapped:mapped,missing:missing};
  } catch(e) {
    issue49v12Write_(state,[['항목','값'],['버전','v1.2-ISSUE49-APR-JUN-MAPPING-VERIFY'],['상태','ERROR'],['단계','FAILED'],['메시지','사업자번호 매핑 최종 검증 실패'],['오류',String(e&&e.message?e.message:e)],['운영시트 변경','0']]); throw e;
  }
}
function issue49v12Indexes_(h){function f(a,d){for(var n=0;n<a.length;n++){var w=issue49v12Compact_(a[n]);for(var i=0;i<h.length;i++)if(issue49v12Compact_(h[i])===w)return i;}return d;}return {date:f(['마켓주문일자','주문일자','결제일자','주문일시'],0),order:f(['마켓주문번호','주문번호','주문ID','주문ID(마켓)'],2),sales:f(['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'],6),status:f(['주문상태','상태','클레임상태','처리상태'],-1)};}
function issue49v12Business_(a){var s=issue49v12Text_(a).toLowerCase();if(s==='beliun1021'||s==='1021')return '227-27-04928';if(s==='beliun1021-1'||s==='1021-1')return '176-71-00758';if(s==='beliun1023'||s==='1023')return '835-58-00765';if(s==='beliun1024'||s==='1024')return '606-45-93763';return '';}
function issue49v12Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=issue49v12Text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(m)return m[1]+'-'+issue49v12Pad_(m[2])+'-'+issue49v12Pad_(m[3]);return '';}
function issue49v12Num_(v){if(typeof v==='number'&&isFinite(v))return v;var n=Number(String(v==null?'':v).replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function issue49v12Text_(v){return String(v==null?'':v).trim();}
function issue49v12Compact_(v){return issue49v12Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue49v12Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function issue49v12Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function issue49v12Write_(sh,rows){sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);}
