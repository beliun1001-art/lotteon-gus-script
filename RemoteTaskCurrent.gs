/**
 * Issue #48 v1.2 read-only diagnostic.
 * Inspects current VAT detail Q:T columns, especially R/매입금액, because
 * production-compatible v1.1 still observed an impossible H1 purchase total.
 * Writes only ISSUE48_* diagnostic sheets.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE48-v1.2-20260812',
  title: '상반기 VAT R열 매입금액 셀 무결성 진단',
  enabled: true,
  outputSheet: 'ISSUE48_R열무결성진단',
  statusSheet: 'ISSUE48_진단상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var state = issue48v12Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  issue48v12Write_(state,[
    ['항목','값'],
    ['버전','v1.2-ISSUE48-VAT-R-COLUMN-INTEGRITY'],
    ['상태','RUNNING'],['단계','LOAD'],
    ['메시지','VAT R열 매입금액 셀 무결성 진단 시작'],
    ['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
  ]);

  try {
    var sheet = ss.getSheetByName('부가세_신고자료');
    if (!sheet || sheet.getLastRow() < 2) throw new Error('부가세_신고자료 시트가 없습니다.');
    if (sheet.getLastColumn() < 20) throw new Error('부가세_신고자료가 T열까지 존재하지 않습니다.');

    var lastRow = sheet.getLastRow();
    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1,1,1,lastCol).getValues()[0] || [];
    var yearIx = issue48v12FindHeader_(headers,['신고연도']);
    var halfIx = issue48v12FindHeader_(headers,['반기']);
    var orderIx = issue48v12FindHeader_(headers,['주문번호','마켓주문번호','주문ID','주문ID(마켓)']);
    if (yearIx < 0 || halfIx < 0 || orderIx < 0) throw new Error('신고연도/반기/주문번호 헤더를 찾지 못했습니다.');

    var values = sheet.getRange(2,1,lastRow-1,lastCol).getValues();
    var qtRange = sheet.getRange(2,17,lastRow-1,4); // Q:T
    var qtValues = qtRange.getValues();
    var qtDisplay = qtRange.getDisplayValues();
    var qtFormats = qtRange.getNumberFormats();
    var qtFormulas = qtRange.getFormulas();

    var stats = [];
    for (var c=0;c<4;c++) stats.push({
      col:17+c, letter:issue48v12Col_(17+c), header:issue48v12Text_(headers[16+c]),
      nonblank:0, numberType:0, stringType:0, formula:0, sum:0, maxAbs:0,
      gt1e12:0, gt1e9:0, reasonablePositive:0, zero:0, orderNumericEqual:0, scientificDisplay:0,
      formats:{}
    });

    var h1Rows = 0;
    var samples = [];
    var fallbackSamples = [];

    for (var r=0;r<values.length;r++) {
      var row = values[r];
      if (issue48v12Text_(row[yearIx]) !== '2026' || issue48v12Text_(row[halfIx]) !== '상반기') continue;
      h1Rows++;
      var orderRaw = row[orderIx];
      var orderNum = issue48v12OrderNumber_(orderRaw);
      var rowLarge = false;

      for (var c2=0;c2<4;c2++) {
        var s = stats[c2], raw = qtValues[r][c2], display = qtDisplay[r][c2], fmt = qtFormats[r][c2], formula = qtFormulas[r][c2];
        var text = issue48v12Text_(raw);
        var n = issue48v12Number_(raw);
        if (text !== '') s.nonblank++;
        if (typeof raw === 'number' && isFinite(raw)) s.numberType++;
        else if (typeof raw === 'string') s.stringType++;
        if (formula) s.formula++;
        if (isFinite(n)) {
          s.sum += n;
          var abs = Math.abs(n);
          if (abs > s.maxAbs) s.maxAbs = abs;
          if (abs > 1e12) { s.gt1e12++; rowLarge = true; }
          if (abs > 1e9) s.gt1e9++;
          if (n > 0 && n <= 10000000) s.reasonablePositive++;
          if (n === 0) s.zero++;
          if (isFinite(orderNum) && orderNum !== 0 && n === orderNum) s.orderNumericEqual++;
        }
        if (/[eE][+-]?\d+/.test(String(display || ''))) s.scientificDisplay++;
        s.formats[fmt || '(공란)'] = (s.formats[fmt || '(공란)'] || 0) + 1;
      }

      var sample = [r+2, issue48v12Text_(orderRaw)];
      for (var c3=0;c3<4;c3++) {
        sample.push(issue48v12Text_(qtValues[r][c3]));
        sample.push(issue48v12Text_(qtDisplay[r][c3]));
        sample.push(issue48v12Text_(qtFormats[r][c3]));
        sample.push(issue48v12Text_(qtFormulas[r][c3]));
      }
      if (fallbackSamples.length < 10) fallbackSamples.push(sample);
      if (rowLarge && samples.length < 20) samples.push(sample);
    }

    var outputHeaders = ['원본행','주문번호'];
    stats.forEach(function(s){
      outputHeaders.push(s.letter+'_raw');
      outputHeaders.push(s.letter+'_display');
      outputHeaders.push(s.letter+'_numberFormat');
      outputHeaders.push(s.letter+'_formula');
    });
    if (!samples.length) samples = fallbackSamples;
    var output = issue48v12Ensure_(ss, LOTTEON_REMOTE_TASK.outputSheet);
    output.clearContents();
    output.getRange(1,1,1,outputHeaders.length).setValues([outputHeaders]);
    if (samples.length) output.getRange(2,1,samples.length,outputHeaders.length).setValues(samples);
    output.setFrozenRows(1);
    output.getRange(1,1,1,outputHeaders.length).setBackground('#d9eaf7').setFontWeight('bold');

    var status = [
      ['항목','값'],
      ['버전','v1.2-ISSUE48-VAT-R-COLUMN-INTEGRITY'],
      ['상태','PASS'],['단계','DONE'],
      ['메시지','VAT R열 매입금액 셀 무결성 진단 완료'],
      ['운영시트 변경','0'],
      ['상반기상세행',h1Rows],
      ['진단샘플행',samples.length]
    ];

    stats.forEach(function(s){
      status.push([s.letter+'_헤더',s.header]);
      status.push([s.letter+'_비공란',s.nonblank]);
      status.push([s.letter+'_number타입',s.numberType]);
      status.push([s.letter+'_string타입',s.stringType]);
      status.push([s.letter+'_수식셀',s.formula]);
      status.push([s.letter+'_합계',s.sum]);
      status.push([s.letter+'_최대절대값',s.maxAbs]);
      status.push([s.letter+'_1조초과',s.gt1e12]);
      status.push([s.letter+'_10억초과',s.gt1e9]);
      status.push([s.letter+'_1천만원이하양수',s.reasonablePositive]);
      status.push([s.letter+'_0값',s.zero]);
      status.push([s.letter+'_주문번호숫자동일',s.orderNumericEqual]);
      status.push([s.letter+'_지수표시',s.scientificDisplay]);
      status.push([s.letter+'_주요numberFormat',issue48v12TopFormats_(s.formats)]);
    });
    status.push(['완료시각',new Date().toISOString()]);
    issue48v12Write_(state,status);
    return {ok:true,h1Rows:h1Rows,columns:stats};
  } catch(e) {
    issue48v12Write_(state,[
      ['항목','값'],['버전','v1.2-ISSUE48-VAT-R-COLUMN-INTEGRITY'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','VAT R열 매입금액 셀 무결성 진단 실패'],['오류',String(e&&e.message?e.message:e)],
      ['운영시트 변경','0'],['갱신시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function issue48v12FindHeader_(headers,names){
  for(var n=0;n<names.length;n++){
    var wanted=issue48v12Compact_(names[n]);
    for(var i=0;i<headers.length;i++) if(issue48v12Compact_(headers[i])===wanted) return i;
  }
  return -1;
}
function issue48v12Number_(v){
  if(typeof v==='number') return isNaN(v)?0:v;
  var n=Number(String(v==null?'0':v).replace(/[원,\s]/g,''));
  return isNaN(n)?0:n;
}
function issue48v12OrderNumber_(v){
  var s=issue48v12Text_(v).replace(/[\s,]/g,'');
  if(!/^\d+(?:\.\d+)?(?:[eE][+-]?\d+)?$/.test(s)) return NaN;
  var n=Number(s); return isFinite(n)?n:NaN;
}
function issue48v12TopFormats_(m){
  return Object.keys(m||{}).sort(function(a,b){return m[b]-m[a];}).slice(0,5).map(function(k){return k+'='+m[k];}).join(' | ');
}
function issue48v12Compact_(v){return issue48v12Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function issue48v12Text_(v){return v==null?'':String(v).trim();}
function issue48v12Col_(n){var s='';while(n>0){var m=(n-1)%26;s=String.fromCharCode(65+m)+s;n=Math.floor((n-1)/26);}return s;}
function issue48v12Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function issue48v12Write_(sheet,rows){sheet.clearContents();sheet.getRange(1,1,rows.length,2).setValues(rows);sheet.setFrozenRows(1);sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');SpreadsheetApp.flush();}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}
