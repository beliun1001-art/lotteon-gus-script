var LOTTEON_PATCH_V647_DIRECT_VAT_ROWS_NO_AGG_LOADED = true;
var LOTTEON_V647_JOB_KEY = 'LOTTEON_V647_DIRECT_VAT_JOB_STATE';
var LOTTEON_V647_HANDLER = 'generateVatReportsFullSeparated_v622';
var LOTTEON_V647_DELAY_MS = 60000;
var LOTTEON_V647_CHUNK_SIZE = 25;
var LOTTEON_V647_MAX_COL = 29;

generateVatReportsFullSeparated_v622 = function() {
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (e) {}
  var state = getVatState_v647_();
  if (state && state.status === 'running') return runVatDirectChunk_v647_(state, ui);
  return startVatDirectJob_v647_(ui);
};

function startVatDirectJob_v647_(ui) {
  clearVatTriggers_v647_();
  var ss = SpreadsheetApp.getActive();
  var src = getSalesInputSheet_v647_();
  var totalSourceRows = Math.max(0, src.getLastRow() - 1);
  var state = {
    status: 'running',
    version: 'v6.47',
    phase: 'vat_detail_direct',
    sourceRow: 2,
    writtenRows: 0,
    totalSourceRows: totalSourceRows,
    skippedRows: 0,
    accountMissingRows: 0,
    startedAt: new Date().toISOString(),
    updatedAt: new Date().toISOString(),
    errors: [],
    memo: 'direct A:AC row chunks, no full aggregation'
  };
  saveVatState_v647_(state);
  writeVatStatus_v647_(state, '예약됨', 'v6.47 시작: 원본 전체 집계 없이 A:AC를 25행씩 직접 처리');
  scheduleVatTrigger_v647_();
  if (ui) {
    try { ui.alert('부가세 신고자료 직접 배치 생성 예약 완료\n\nv6.47 기준으로 원본 전체 집계 없이 A:AC를 25행씩 직접 처리합니다.\n약 1분 뒤 자동 실행됩니다.\n같은 메뉴를 한 번 더 누르면 즉시 다음 배치를 실행합니다.'); } catch (e) {}
  }
  return { ok: true, scheduled: true, state: state };
}

function runVatDirectChunk_v647_(state, ui) {
  var started = Date.now();
  try {
    writeVatStatus_v647_(state, '실행중', '현재 배치 시작: 원본 ' + state.sourceRow + '행부터');
    var src = getSalesInputSheet_v647_();
    var lastRow = src.getLastRow();
    var maxCol = Math.min(src.getLastColumn(), LOTTEON_V647_MAX_COL);
    var headers = src.getRange(1, 1, 1, maxCol).getValues()[0].map(function(h){ return String(h || '').trim(); });
    var startRow = Number(state.sourceRow || 2);

    if (startRow > lastRow) {
      finishVatDirectJob_v647_(state, ui);
      return { ok: true, done: true, state: state };
    }

    var rowCount = Math.min(LOTTEON_V647_CHUNK_SIZE, lastRow - startRow + 1);
    var values = src.getRange(startRow, 1, rowCount, maxCol).getValues();
    var detailRows = [];
    var skipped = 0;
    var accountMissing = 0;

    if (Number(state.writtenRows || 0) === 0) initVatDetailSheet_v647_();

    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      var obj = rowObj_v647_(headers, row);
      var detail = makeVatDetailFromSourceRow_v647_(obj, row, startRow + i);
      if (!detail) { skipped++; continue; }
      if (!detail[1] || detail[1] === '계정미확인' || detail[1] === '마켓아이디없음') accountMissing++;
      detailRows.push(detail);
    }

    if (detailRows.length) {
      var out = SpreadsheetApp.getActive().getSheetByName('부가세_신고자료') || SpreadsheetApp.getActive().insertSheet('부가세_신고자료');
      out.getRange(2 + Number(state.writtenRows || 0), 1, detailRows.length, getVatHeaders_v647_().length).setValues(detailRows);
      formatVatRows_v647_(out, 2 + Number(state.writtenRows || 0), detailRows.length);
    }

    state.sourceRow = startRow + rowCount;
    state.writtenRows = Number(state.writtenRows || 0) + detailRows.length;
    state.skippedRows = Number(state.skippedRows || 0) + skipped;
    state.accountMissingRows = Number(state.accountMissingRows || 0) + accountMissing;
    state.updatedAt = new Date().toISOString();
    saveVatState_v647_(state);

    var done = state.sourceRow > lastRow;
    if (done) {
      finishVatDirectJob_v647_(state, ui);
      return { ok: true, done: true, state: state };
    }

    clearVatTriggers_v647_();
    scheduleVatTrigger_v647_();
    writeVatStatus_v647_(state, '배치진행중', '이번 배치 완료: 원본 ' + startRow + '~' + (startRow + rowCount - 1) + '행 / 작성 ' + detailRows.length + '행 / 소요초 ' + Math.round((Date.now() - started)/1000));
    if (ui) {
      try { ui.alert('부가세_신고자료 배치 완료\n\n작성 누계: ' + state.writtenRows + '행\n다음 원본 행: ' + state.sourceRow + '\n\n다음 배치는 약 1분 뒤 자동 실행됩니다.'); } catch (e) {}
    }
    return { ok: true, done: false, state: state };
  } catch (e) {
    state.status = 'failed';
    state.updatedAt = new Date().toISOString();
    state.errors = state.errors || [];
    state.errors.push({ at: new Date().toISOString(), sourceRow: state.sourceRow, message: String(e && e.message ? e.message : e) });
    saveVatState_v647_(state);
    clearVatTriggers_v647_();
    writeVatStatus_v647_(state, '오류', String(e && e.message ? e.message : e));
    throw e;
  }
}

function finishVatDirectJob_v647_(state, ui) {
  state.status = 'done';
  state.updatedAt = new Date().toISOString();
  saveVatState_v647_(state);
  clearVatTriggers_v647_();
  buildAccountSummaryFromVatDetail_v647_();
  buildBusinessNoDiagnosticFromVatDetail_v647_();
  try { SpreadsheetApp.getActive().getSheetByName('부가세_신고자료').setFrozenRows(1); } catch (e) {}
  writeVatStatus_v647_(state, '완료', '부가세_신고자료 직접 생성 완료');
  if (ui) {
    try { ui.alert('부가세 신고자료 생성 완료\n\n작성행수: ' + state.writtenRows + '\n제외행수: ' + state.skippedRows + '\n계정미확인: ' + state.accountMissingRows + '\n\n확인 시트: 부가세_신고자료, 사업자별_계정별_써머리, 사업자번호_매핑검증'); } catch (e) {}
  }
}

function makeVatDetailFromSourceRow_v647_(obj, row, sourceRowNo) {
  var statusText = pickAny_v647_(obj, ['주문상태','상태','클레임상태','주문상태명','처리상태']);
  if (/취소|반품|환불/.test(String(statusText || ''))) return null;
  var accountId = clean_v647_(row[3]) || '마켓아이디없음';
  var sales = num_v647_(pickAny_v647_(obj, ['결제금액합계(원)','결제금액합계','결제금액','매출액','판매금액','순수매출액']));
  if (!sales) return null;
  var actualSettlement = num_v647_(pickAny_v647_(obj, ['정산예정금액(원)','정산예정금액','실제정산금액','정산금액']));
  var settlement = actualSettlement || Math.round(sales * 0.901);
  var purchase = num_v647_(row[28]);
  var fee = sales - settlement;
  var profit = settlement - purchase;
  var ss = splitVat_v647_(sales);
  var ps = splitVat_v647_(purchase);
  var payable = ss.vat - ps.vat;
  var dateText = shortDate_v647_(pickAny_v647_(obj, ['주문일','주문일자','결제일','결제일시','주문일시']));
  var bizNo = bizNo_v647_(accountId);
  return [
    dateText,
    accountId,
    bizNo,
    clean_v647_(pickAny_v647_(obj, ['주문번호','주문 번호','주문ID','주문ID(마켓)'])),
    clean_v647_(pickAny_v647_(obj, ['고객명','수령인','수취인','구매자','주문자'])),
    clean_v647_(pickAny_v647_(obj, ['브랜드명','브랜드'])),
    clean_v647_(pickAny_v647_(obj, ['상품번호','상품코드','판매자상품코드','마켓상품번호'])),
    clean_v647_(pickAny_v647_(obj, ['상품명','상품명(옵션포함)','등록상품명'])),
    num_v647_(pickAny_v647_(obj, ['판매수량','수량','구매수량'])) || 1,
    sales,
    ss.supply,
    ss.vat,
    settlement,
    fee,
    purchase,
    ps.supply,
    ps.vat,
    payable,
    profit,
    profit - payable,
    bizNo ? 'v6.47 직접생성' : '사업자번호 매핑 확인 필요'
  ];
}

function initVatDetailSheet_v647_() {
  var ss = SpreadsheetApp.getActive();
  ['부가세_고객별','부가세_주문번호별','고객주소_구분검증'].forEach(function(n){ var sh=ss.getSheetByName(n); if(sh){ try{ ss.deleteSheet(sh); }catch(e){ try{ sh.hideSheet(); }catch(x){} } } });
  var sh = ss.getSheetByName('부가세_신고자료') || ss.insertSheet('부가세_신고자료');
  sh.clearContents();
  sh.getRange(1,1,1,getVatHeaders_v647_().length).setValues([getVatHeaders_v647_()]);
  sh.getRange(1,1,1,getVatHeaders_v647_().length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center').setWrapStrategy(SpreadsheetApp.WrapStrategy.WRAP);
}

function getVatHeaders_v647_() {
  return ['날짜','쿠팡계정ID','사업자등록번호','주문번호','고객명','브랜드명','상품번호','상품명','판매수량','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];
}

function formatVatRows_v647_(sh, start, count) {
  try {
    sh.getRange(start,1,count,21).setVerticalAlignment('middle').setWrapStrategy(SpreadsheetApp.WrapStrategy.CLIP);
    sh.getRange(start,1,count,1).setNumberFormat('@').setHorizontalAlignment('center');
    [9,10,11,12,13,14,15,16,17,18,19,20].forEach(function(c){ sh.getRange(start,c,count,1).setNumberFormat('#,##0').setHorizontalAlignment('right'); });
  } catch(e) {}
}

function buildAccountSummaryFromVatDetail_v647_() {
  var ss = SpreadsheetApp.getActive();
  var sh = ss.getSheetByName('부가세_신고자료');
  if (!sh || sh.getLastRow() < 2) return;
  var data = sh.getRange(2,1,sh.getLastRow()-1,21).getValues();
  var map = {};
  data.forEach(function(r){
    var a = clean_v647_(r[1]) || '계정미확인';
    if(!map[a]) map[a] = {orders:{}, qty:0, sales:0, settle:0, fee:0, purchase:0, profit:0, salesVat:0, purchaseVat:0, payable:0, afterVat:0, biz: r[2]};
    if(r[3]) map[a].orders[r[3]] = true;
    map[a].qty += num_v647_(r[8]); map[a].sales += num_v647_(r[9]); map[a].settle += num_v647_(r[12]); map[a].fee += num_v647_(r[13]); map[a].purchase += num_v647_(r[14]); map[a].salesVat += num_v647_(r[11]); map[a].purchaseVat += num_v647_(r[16]); map[a].payable += num_v647_(r[17]); map[a].profit += num_v647_(r[18]); map[a].afterVat += num_v647_(r[19]);
  });
  var rows = Object.keys(map).sort().map(function(a){ var x=map[a]; return [a,x.biz||bizNo_v647_(a),Object.keys(x.orders).length,x.qty,x.sales,x.settle,x.fee,x.sales?x.fee/x.sales:0,x.purchase,x.profit,x.sales?x.profit/x.sales:0,x.salesVat,x.purchaseVat,x.payable,x.afterVat,x.sales?x.afterVat/x.sales:0,'v6.47 직접생성']; });
  var out = ss.getSheetByName('사업자별_계정별_써머리') || ss.insertSheet('사업자별_계정별_써머리');
  out.clearContents();
  var h = ['쿠팡계정ID','사업자등록번호','주문건수','판매수량','순수매출액','정산기준금액','마켓수수료/비용','마켓수수료율','매입금액','예상이익','예상이익률','매출부가세','매입부가세','납부예상부가세','부가세반영예상이익','부가세반영이익률','비고'];
  out.getRange(1,1,1,h.length).setValues([h]); if(rows.length) out.getRange(2,1,rows.length,h.length).setValues(rows);
  out.getRange(1,1,1,h.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  if(rows.length){ [8,11,16].forEach(function(c){out.getRange(2,c,rows.length,1).setNumberFormat('0.0%');}); }
}

function buildBusinessNoDiagnosticFromVatDetail_v647_() {
  var ss=SpreadsheetApp.getActive(), sh=ss.getSheetByName('부가세_신고자료'), out=ss.getSheetByName('사업자번호_매핑검증')||ss.insertSheet('사업자번호_매핑검증');
  var rows=[['계정/항목','사업자등록번호','행수/메모']];
  if(sh && sh.getLastRow()>1){ var data=sh.getRange(2,2,sh.getLastRow()-1,2).getValues(), m={}; data.forEach(function(r){var k=(r[0]||'')+'|'+(r[1]||'사업자번호없음'); m[k]=(m[k]||0)+1;}); Object.keys(m).sort().forEach(function(k){var p=k.split('|'); rows.push([p[0],p[1],m[k]]);}); }
  out.clearContents(); out.getRange(1,1,rows.length,3).setValues(rows); out.getRange(1,1,1,3).setBackground('#d9eaf7').setFontWeight('bold');
}

function writeVatStatus_v647_(state, status, memo) {
  var sh=SpreadsheetApp.getActive().getSheetByName('부가세_생성상태')||SpreadsheetApp.getActive().insertSheet('부가세_생성상태');
  var rows=[['항목','값','메모'],['상태',status,'v6.47 직접 배치'],['현재구간',state.phase||'','full agg 사용 안 함'],['원본 진행행',String(state.sourceRow||2)+' / '+String((state.totalSourceRows||0)+1),'매출데이터_붙여넣기 기준'],['작성행수',state.writtenRows||0,'부가세_신고자료'],['제외행수',state.skippedRows||0,'취소/반품/매출0 등'],['계정미확인',state.accountMissingRows||0,'D열 마켓아이디 공란'],['시작시각',state.startedAt||'',''],['갱신시각',new Date().toISOString(),''],['오류',(state.errors||[]).map(function(e){return e.message;}).join(' / '),''],['메모',memo||'','']];
  sh.clearContents(); sh.getRange(1,1,rows.length,3).setValues(rows); sh.getRange(1,1,1,3).setBackground('#d9eaf7').setFontWeight('bold'); try{sh.autoResizeColumns(1,3);}catch(e){}
}

function getSalesInputSheet_v647_(){ var sh=SpreadsheetApp.getActive().getSheetByName('매출데이터_붙여넣기'); if(!sh) throw new Error('매출데이터_붙여넣기 시트를 찾을 수 없습니다.'); return sh; }
function getVatState_v647_(){ var raw=PropertiesService.getScriptProperties().getProperty(LOTTEON_V647_JOB_KEY); if(!raw)return null; try{return JSON.parse(raw);}catch(e){return null;} }
function saveVatState_v647_(s){ PropertiesService.getScriptProperties().setProperty(LOTTEON_V647_JOB_KEY, JSON.stringify(s)); }
function scheduleVatTrigger_v647_(){ ScriptApp.newTrigger(LOTTEON_V647_HANDLER).timeBased().after(LOTTEON_V647_DELAY_MS).create(); }
function clearVatTriggers_v647_(){ try{ ScriptApp.getProjectTriggers().forEach(function(t){ try{ if(t.getHandlerFunction&&t.getHandlerFunction()===LOTTEON_V647_HANDLER) ScriptApp.deleteTrigger(t); }catch(e){} }); }catch(e){} }
function rowObj_v647_(headers,row){ var o={}; headers.forEach(function(h,i){ if(h)o[h]=row[i]; }); return o; }
function pickAny_v647_(o,names){ for(var i=0;i<names.length;i++){ if(o.hasOwnProperty(names[i]) && o[names[i]]!=='' && o[names[i]]!=null) return o[names[i]]; } return ''; }
function clean_v647_(v){ return String(v==null?'':v).replace(/,/g,'').trim(); }
function num_v647_(v){ if(typeof v==='number')return v; var n=Number(String(v==null?'':v).replace(/₩|,|%/g,'').trim()); return isNaN(n)?0:n; }
function splitVat_v647_(a){ a=Math.round(num_v647_(a)); var s=Math.round(a/1.1); return {supply:s,vat:a-s}; }
function bizNo_v647_(a){ a=clean_v647_(a); if(/1021/.test(a))return '227-27-04928'; if(/1023/.test(a))return '835-58-00765'; if(/1024/.test(a))return '606-45-93763'; return ''; }
function shortDate_v647_(v){ if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v))return Utilities.formatDate(v, Session.getScriptTimeZone()||'Asia/Seoul','MM/dd'); var s=String(v==null?'':v).trim(); var m=s.match(/(\d{1,2})[\.\/-](\d{1,2})/); if(m)return ('0'+Number(m[1])).slice(-2)+'/'+('0'+Number(m[2])).slice(-2); return s.slice(0,10); }
