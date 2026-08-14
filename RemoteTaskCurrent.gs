var LOTTEON_REMOTE_TASK={id:'ISSUE68-v1.1-20260814',title:'corrected VAT 부가세_기간별 guarded production rebuild',enabled:true,statusSheet:'ISSUE68_반영상태'};
var I68_VERSION='v1.1-ISSUE68-CORRECTED-VAT-PERIOD-PRODUCTION';
var I68_VAT='부가세_신고자료',I68_CARD='부가세_카드매칭검증',I68_CARD_PREVIEW='ISSUE54_카드매칭전체PREVIEW',I68_BIZ_PREVIEW='ISSUE67_사업자별부가세PREVIEW',I68_I67='ISSUE67_실행상태',I68_I66='ISSUE66_반영상태',I68_PERIOD='부가세_기간별',I68_HISTORY='카드사용내역_붙여넣기',I68_MASTER='카드_마스터',I68_CARD_BACKUP='ISSUE59_백업_부가세카드매칭검증',I68_PREP='ISSUE68_기간별준비',I68_PERIOD_BACKUP='ISSUE68_백업_부가세기간별';

function runLotteonRemoteTaskStartRemote_(){
  var ss=SpreadsheetApp.getActive();if(!ss)throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var st=i68Ensure_(ss,LOTTEON_REMOTE_TASK.statusSheet);
  i68Status_(st,'RUNNING','PRECHECK','Issue67 PASS 기반 corrected 부가세_기간별 재생성 사전검증',{});
  var wrote=false,backupReady=false,periodFilterMeta=null,before=null;
  try{
    [I68_VAT,I68_CARD,I68_CARD_PREVIEW,I68_BIZ_PREVIEW,I68_I67,I68_I66,I68_PERIOD,I68_HISTORY,I68_MASTER,I68_CARD_BACKUP].forEach(function(n){i68Need_(ss,n);});
    i68ExpectKv_(ss,I68_I66,{'버전':'v1.0-ISSUE66-FILTER-SAFE-CORRECTED-APPLY','상태':'PASS','단계':'DONE','운영주문':1355,'MATCHED':808,'NON_CARD':498,'AMBIGUOUS':0,'NO_MATCH':49,'v6.69 2차귀속':1161,'v6.70 3차귀속':81,'주문매입금액합계':105762969,'typed셀차이':0,'display셀차이':0,'롤백':0});
    i68ExpectKv_(ss,I68_I67,{'버전':'v1.0-ISSUE67-CORRECTED-BUSINESS-HALF-VAT-PREVIEW','상태':'PASS','단계':'DONE','VAT상세행':2752,'VAT주문':1355,'사업자수':4,'사업자×카드 요약행':19,'MATCHED':808,'NON_CARD':498,'AMBIGUOUS':0,'NO_MATCH':49,'VAT/카드_overlap':1355,'VAT_only':0,'카드_only':0,'주문매입금액 불일치':0,'사업자 불일치':0,'순수매출액':138432300,'정산기준금액':122495855,'마켓수수료':15936445,'매입금액':105762969,'매출부가세':12584695,'매입부가세':9614786,'납부예상부가세':2969909,'예상이익':16732886,'부가세반영예상이익':13762977,'운영카드/Issue54 typed차이':0,'운영카드/Issue54 display차이':0});

    var vat=i68Need_(ss,I68_VAT),card=i68Need_(ss,I68_CARD),cardPreview=i68Need_(ss,I68_CARD_PREVIEW),bizPreview=i68Need_(ss,I68_BIZ_PREVIEW),period=i68Need_(ss,I68_PERIOD);
    before=i68Protected_(ss);
    if(i68SheetTypedDiff_(card,cardPreview)!==0||i68SheetDisplayDiff_(card,cardPreview)!==0)throw new Error('운영 카드검증/Issue54 preview 불일치');
    if(bizPreview.getLastRow()!==21||bizPreview.getLastColumn()!==24)throw new Error('Issue67 사업자 preview 구조 불일치 '+bizPreview.getLastRow()+'x'+bizPreview.getLastColumn()+' / 기대 21x24');
    if(i68FormulaCount_(bizPreview)>0)throw new Error('Issue67 사업자 preview 수식셀 존재 '+i68FormulaCount_(bizPreview));
    if(i68FormulaCount_(period)>0)throw new Error('기존 부가세_기간별 수식셀 존재 '+i68FormulaCount_(period)+' — 자동 덮어쓰기 금지');

    var expectedBizHeaders=['신고연도','반기','사업자등록번호','연결 쿠팡계정ID','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','카드매칭상태','카드매칭근거','주문건수','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];
    var actualBizHeaders=bizPreview.getRange(2,1,1,24).getValues()[0];
    for(var bh=0;bh<24;bh++)if(i68Text_(actualBizHeaders[bh])!==expectedBizHeaders[bh])throw new Error('Issue67 preview 헤더 불일치 C'+(bh+1)+' 실제 '+actualBizHeaders[bh]+' 기대 '+expectedBizHeaders[bh]);
    var bizRecon=i68BizRecon_(bizPreview.getRange(3,1,19,24).getValues());
    i68Assert_(bizRecon,{rows:19,orders:1355,sales:138432300,settlement:122495855,fee:15936445,purchase:105762969,salesVat:12584695,purchaseVat:9614786,payable:2969909,profit:16732886,vatProfit:13762977});

    var vatValues=vat.getDataRange().getValues();
    var vatCheck=i68VatCheck_(vatValues);
    i68Assert_(vatCheck,{detailRows:2752,orders:1355,businessCount:4,businessBlank:0,periodUnknown:0,sales:138432300,settlement:122495855,fee:15936445,purchase:105762969,salesVat:12584695,purchaseVat:9614786,payable:2969909,profit:16732886,vatProfit:13762977});

    var periodAgg=i68AggregatePeriods_(vatValues);
    if(periodAgg.blankOrderDetailRows!==0)throw new Error('기간집계 주문번호 공란 상세행 '+periodAgg.blankOrderDetailRows);
    if(periodAgg.periodUnknown!==0)throw new Error('기간미확인 주문 '+periodAgg.periodUnknown);
    i68Assert_(periodAgg.halfTotals,{orders:1355,sales:138432300,settlement:122495855,fee:15936445,purchase:105762969,salesVat:12584695,purchaseVat:9614786,payable:2969909,profit:16732886,vatProfit:13762977});
    i68Assert_(periodAgg.monthTotals,{orders:1355,sales:138432300,settlement:122495855,fee:15936445,purchase:105762969,salesVat:12584695,purchaseVat:9614786,payable:2969909,profit:16732886,vatProfit:13762977});

    var matrix=i68BuildMatrix_(bizPreview.getDataRange().getValues(),periodAgg.rows);
    var oldPrep=ss.getSheetByName(I68_PREP);if(oldPrep)ss.deleteSheet(oldPrep);
    var prep=ss.insertSheet(I68_PREP);i68EnsureGrid_(prep,matrix.values.length,matrix.values[0].length);
    prep.getRange(1,1,matrix.formats.length,matrix.formats[0].length).setNumberFormats(matrix.formats);
    prep.getRange(1,1,matrix.values.length,matrix.values[0].length).setValues(matrix.values);
    i68ApplyLayout_(prep,matrix.summaryRows,matrix.detailStart,matrix.detailHeaders);
    SpreadsheetApp.flush();

    var bizSrc=bizPreview.getRange(2,1,20,24),bizDst=prep.getRange(2,1,20,24);
    var prepBizTyped=i68RangeTypedDiff_(bizSrc,bizDst),prepBizDisplay=i68RangeDisplayDiff_(bizSrc,bizDst);
    if(prepBizTyped||prepBizDisplay)throw new Error('준비시트/Issue67 사업자요약 불일치 typed='+prepBizTyped+' display='+prepBizDisplay);
    var prepPeriod=i68ReadPeriodBlock_(prep,matrix.detailStart);
    i68Assert_(prepPeriod.halfTotals,periodAgg.halfTotals);i68Assert_(prepPeriod.monthTotals,periodAgg.monthTotals);
    if(prepPeriod.periodUnknown!==0)throw new Error('준비시트 기간미확인 '+prepPeriod.periodUnknown);
    i68CheckProtected_(ss,before);

    var staleBackup=ss.getSheetByName(I68_PERIOD_BACKUP);if(staleBackup)ss.deleteSheet(staleBackup);
    var backup=period.copyTo(ss).setName(I68_PERIOD_BACKUP);backupReady=true;
    SpreadsheetApp.flush();
    var backupTyped=i68SheetTypedDiff_(period,backup),backupDisplay=i68SheetDisplayDiff_(period,backup);
    if(backupTyped||backupDisplay)throw new Error('부가세_기간별 백업 불일치 typed='+backupTyped+' display='+backupDisplay);
    periodFilterMeta=i68CaptureFilter_(period);

    i68Status_(st,'RUNNING','WRITE','준비/백업 검증 PASS; corrected 부가세_기간별 운영 반영 시작',{summaryRows:matrix.summaryRows,periodRows:periodAgg.rows.length,halfRows:periodAgg.halfRows,monthRows:periodAgg.monthRows,prepBizTyped:prepBizTyped,prepBizDisplay:prepBizDisplay,backupTyped:backupTyped,backupDisplay:backupDisplay,rollback:0});
    wrote=true;
    i68RemoveFilter_(period);
    period.clear();i68EnsureGrid_(period,matrix.values.length,matrix.values[0].length);
    period.getRange(1,1,matrix.formats.length,matrix.formats[0].length).setNumberFormats(matrix.formats);
    period.getRange(1,1,matrix.values.length,matrix.values[0].length).setValues(matrix.values);
    i68ApplyLayout_(period,matrix.summaryRows,matrix.detailStart,matrix.detailHeaders);
    i68RestoreFilter_(period,periodFilterMeta);
    SpreadsheetApp.flush();

    var prodTyped=i68SheetTypedDiff_(prep,period),prodDisplay=i68SheetDisplayDiff_(prep,period),prodFormat=i68SheetFormatDiff_(prep,period);
    if(prodTyped||prodDisplay||prodFormat)throw new Error('운영/준비 불일치 typed='+prodTyped+' display='+prodDisplay+' format='+prodFormat);
    var prodBizTyped=i68RangeTypedDiff_(bizSrc,period.getRange(2,1,20,24)),prodBizDisplay=i68RangeDisplayDiff_(bizSrc,period.getRange(2,1,20,24));
    if(prodBizTyped||prodBizDisplay)throw new Error('운영/Issue67 사업자요약 불일치 typed='+prodBizTyped+' display='+prodBizDisplay);
    var prodPeriod=i68ReadPeriodBlock_(period,matrix.detailStart);
    i68Assert_(prodPeriod.halfTotals,periodAgg.halfTotals);i68Assert_(prodPeriod.monthTotals,periodAgg.monthTotals);
    if(prodPeriod.periodUnknown!==0)throw new Error('운영 기간미확인 '+prodPeriod.periodUnknown);
    i68CheckProtected_(ss,before);

    i68Status_(st,'PASS','DONE','corrected 사업자별 요약 + 반기/월별 부가세_기간별 운영 재생성 완료',{summaryRows:matrix.summaryRows,periodRows:periodAgg.rows.length,halfRows:periodAgg.halfRows,monthRows:periodAgg.monthRows,periodUnknown:prodPeriod.periodUnknown,orders:prodPeriod.halfTotals.orders,sales:prodPeriod.halfTotals.sales,settlement:prodPeriod.halfTotals.settlement,fee:prodPeriod.halfTotals.fee,purchase:prodPeriod.halfTotals.purchase,salesVat:prodPeriod.halfTotals.salesVat,purchaseVat:prodPeriod.halfTotals.purchaseVat,payable:prodPeriod.halfTotals.payable,profit:prodPeriod.halfTotals.profit,vatProfit:prodPeriod.halfTotals.vatProfit,prepBizTyped:prepBizTyped,prepBizDisplay:prepBizDisplay,prodTyped:prodTyped,prodDisplay:prodDisplay,prodFormat:prodFormat,prodBizTyped:prodBizTyped,prodBizDisplay:prodBizDisplay,backupTyped:backupTyped,backupDisplay:backupDisplay,filterRestored:periodFilterMeta&&periodFilterMeta.exists?'YES':'NO',filterCriteria:periodFilterMeta?periodFilterMeta.criteriaCount:0,rollback:0});
    return{ok:true,done:true,summaryRows:matrix.summaryRows,periodRows:periodAgg.rows.length};
  }catch(e){
    var msg=String(e&&e.message?e.message:e),rb='0',rbErr='';
    if(wrote&&backupReady){
      try{
        var prod=ss.getSheetByName(I68_PERIOD),backupSh=ss.getSheetByName(I68_PERIOD_BACKUP);
        if(!prod||!backupSh)throw new Error('롤백 시트 누락');
        i68RemoveFilter_(prod);prod.clear();
        var br=backupSh.getLastRow(),bc=backupSh.getLastColumn();i68EnsureGrid_(prod,Math.max(br,1),Math.max(bc,1));
        if(br&&bc){
          backupSh.getRange(1,1,br,bc).copyTo(prod.getRange(1,1,br,bc),SpreadsheetApp.CopyPasteType.PASTE_FORMAT,false);
          prod.getRange(1,1,br,bc).setNumberFormats(backupSh.getRange(1,1,br,bc).getNumberFormats());
          prod.getRange(1,1,br,bc).setValues(backupSh.getRange(1,1,br,bc).getValues());
        }
        prod.setFrozenRows(backupSh.getFrozenRows());prod.setFrozenColumns(backupSh.getFrozenColumns());
        for(var c=1;c<=Math.min(prod.getMaxColumns(),backupSh.getMaxColumns());c++)prod.setColumnWidth(c,backupSh.getColumnWidth(c));
        i68RestoreFilterFromSheet_(prod,backupSh);SpreadsheetApp.flush();
        var rbt=i68SheetTypedDiff_(prod,backupSh),rbd=i68SheetDisplayDiff_(prod,backupSh);if(rbt||rbd)throw new Error('롤백 검증 실패 typed='+rbt+' display='+rbd);rb='1';
      }catch(re){rb='ERROR';rbErr=String(re&&re.message?re.message:re);}
    }
    var status=rb==='1'?'ROLLED_BACK':(rb==='ERROR'?'ROLLBACK_ERROR':'ERROR');
    i68Status_(st,status,'FAILED',rb==='1'?'부가세_기간별 반영 오류 후 기존 시트 자동 롤백 완료':'Issue68 부가세_기간별 운영 반영 실패',{error:msg+(rbErr?' / ROLLBACK: '+rbErr:''),rollback:rb});
    if(rb!=='1')throw e;
    return{ok:false,done:true,status:status,error:msg};
  }
}
function runLotteonRemoteTaskContinueRemote_(){return runLotteonRemoteTaskStartRemote_();}

function i68BizRecon_(rows){var t={rows:(rows||[]).length,orders:0,sales:0,settlement:0,fee:0,purchase:0,salesVat:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};(rows||[]).forEach(function(r){t.orders+=i68Num_(r[11]);t.sales+=i68Num_(r[12]);t.salesVat+=i68Num_(r[14]);t.settlement+=i68Num_(r[15]);t.fee+=i68Num_(r[16]);t.purchase+=i68Num_(r[17]);t.purchaseVat+=i68Num_(r[19]);t.payable+=i68Num_(r[20]);t.profit+=i68Num_(r[21]);t.vatProfit+=i68Num_(r[22]);});return t;}
function i68VatCheck_(v){
  var h=v[0]||[],ix=function(n){return i68Find_(h,n);},p={year:ix(['신고연도']),half:ix(['반기']),month:ix(['신고월']),account:ix(['쿠팡계정ID']),business:ix(['사업자등록번호']),order:ix(['주문번호','마켓주문번호','주문ID','주문ID(마켓)']),sales:ix(['순수매출액']),settlement:ix(['정산기준금액']),purchase:ix(['매입금액']),salesVat:ix(['매출부가세']),purchaseVat:ix(['매입부가세']),payable:ix(['납부예상부가세']),profit:ix(['예상이익']),vatProfit:ix(['부가세반영예상이익'])};
  Object.keys(p).forEach(function(k){if(p[k]<0)throw new Error('VAT 필수 헤더 누락 '+k);});
  var s={detailRows:0,orders:0,businessCount:0,businessBlank:0,periodUnknown:0,sales:0,settlement:0,fee:0,purchase:0,salesVat:0,purchaseVat:0,payable:0,profit:0,vatProfit:0},orders={},businesses={};
  for(var r=1;r<v.length;r++){var row=v[r],y=i68Text_(row[p.year]),half=i68Text_(row[p.half]);if(y!=='2026'||half!=='상반기')continue;s.detailRows++;var a=i68Text_(row[p.account]),b=i68Text_(row[p.business]),o=i68Text_(row[p.order]),m=i68Text_(row[p.month]);if(!a||!o)throw new Error('VAT 주문키 공란 R'+(r+1));if(!b)s.businessBlank++;else businesses[b]=true;if(!m||m==='기간미확인')s.periodUnknown++;orders[i68Key_(a,o)]=true;var sales=i68Num_(row[p.sales]),sett=i68Num_(row[p.settlement]);s.sales+=sales;s.settlement+=sett;s.fee+=sales-sett;s.purchase+=i68Num_(row[p.purchase]);s.salesVat+=i68Num_(row[p.salesVat]);s.purchaseVat+=i68Num_(row[p.purchaseVat]);s.payable+=i68Num_(row[p.payable]);s.profit+=i68Num_(row[p.profit]);s.vatProfit+=i68Num_(row[p.vatProfit]);}
  s.orders=Object.keys(orders).length;s.businessCount=Object.keys(businesses).length;return s;
}
function i68AggregatePeriods_(v){
  var h=v[0]||[],ix=function(n){return i68Find_(h,n);},p={year:ix(['신고연도']),half:ix(['반기']),month:ix(['신고월']),account:ix(['쿠팡계정ID']),business:ix(['사업자등록번호']),order:ix(['주문번호','마켓주문번호','주문ID','주문ID(마켓)']),sales:ix(['순수매출액']),settlement:ix(['정산기준금액']),purchase:ix(['매입금액']),salesVat:ix(['매출부가세']),purchaseVat:ix(['매입부가세']),payable:ix(['납부예상부가세']),profit:ix(['예상이익']),vatProfit:ix(['부가세반영예상이익'])};
  Object.keys(p).forEach(function(k){if(p[k]<0)throw new Error('기간집계 필수 헤더 누락 '+k);});
  var map={},blank=0;
  function add(kind,y,half,month,row){var account=i68Text_(row[p.account]),business=i68Text_(row[p.business]),key=[kind,y,half,month,account,business].join('|');if(!map[key])map[key]={row:[kind,y,half,month,account,business,0,0,0,0,0,0,0,0,0,0],orders:{}};var x=map[key],o=i68Text_(row[p.order]);if(o){if(!x.orders[o]){x.orders[o]=true;x.row[6]++;}}else blank++;var sales=i68Num_(row[p.sales]),sett=i68Num_(row[p.settlement]);x.row[7]+=sales;x.row[8]+=sett;x.row[9]+=sales-sett;x.row[10]+=i68Num_(row[p.purchase]);x.row[11]+=i68Num_(row[p.salesVat]);x.row[12]+=i68Num_(row[p.purchaseVat]);x.row[13]+=i68Num_(row[p.payable]);x.row[14]+=i68Num_(row[p.profit]);x.row[15]+=i68Num_(row[p.vatProfit]);}
  for(var r=1;r<v.length;r++){var row=v[r],y=i68Text_(row[p.year]),half=i68Text_(row[p.half]),m=i68Text_(row[p.month]);if(y!=='2026'||half!=='상반기')continue;if(!m||m==='기간미확인')add('기간미확인','기간미확인','기간미확인','기간미확인',row);else{add('반기',y,half,'',row);add('월별',y,half,m,row);}}
  var rank={'반기':0,'월별':1,'기간미확인':2},rows=Object.keys(map).map(function(k){return map[k].row;}).sort(function(a,b){return String(a[1]).localeCompare(String(b[1]))||(rank[a[0]]-rank[b[0]])||String(a[2]).localeCompare(String(b[2]))||String(a[3]).localeCompare(String(b[3]))||String(a[4]).localeCompare(String(b[4]));});
  var halfTotals=i68Totals_(rows.filter(function(r){return r[0]==='반기';})),monthTotals=i68Totals_(rows.filter(function(r){return r[0]==='월별';}));
  var unknown=rows.filter(function(r){return r[0]==='기간미확인';}).reduce(function(n,r){return n+i68Num_(r[6]);},0);
  return{rows:rows,blankOrderDetailRows:blank,periodUnknown:unknown,halfRows:rows.filter(function(r){return r[0]==='반기';}).length,monthRows:rows.filter(function(r){return r[0]==='월별';}).length,halfTotals:halfTotals,monthTotals:monthTotals};
}
function i68Totals_(rows){var t={orders:0,sales:0,settlement:0,fee:0,purchase:0,salesVat:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};(rows||[]).forEach(function(r){t.orders+=i68Num_(r[6]);t.sales+=i68Num_(r[7]);t.settlement+=i68Num_(r[8]);t.fee+=i68Num_(r[9]);t.purchase+=i68Num_(r[10]);t.salesVat+=i68Num_(r[11]);t.purchaseVat+=i68Num_(r[12]);t.payable+=i68Num_(r[13]);t.profit+=i68Num_(r[14]);t.vatProfit+=i68Num_(r[15]);});return t;}
function i68BuildMatrix_(bizValues,periodRows){
  if(!bizValues||bizValues.length!==21)throw new Error('Issue67 preview 행수 불일치 '+(bizValues?bizValues.length:0));
  var bizHeaders=bizValues[1],summary=bizValues.slice(2),detailHeaders=['집계구분','신고연도','반기','신고월','쿠팡계정ID','사업자등록번호','주문건수','순수매출액','정산기준금액','마켓수수료','매입금액','매출부가세','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'];
  if(!bizHeaders||bizHeaders.length!==24||summary.length!==19)throw new Error('Issue67 preview 헤더/행수 불일치');
  var maxCols=24,detailStart=summary.length+5,values=[],formats=[];
  function blankRow(){var a=[];for(var c=0;c<maxCols;c++)a.push('');return a;}
  function pad(r){var a=r.slice();while(a.length<maxCols)a.push('');return a;}
  var r1=blankRow();r1[0]='사업자별 반기 신고요약 (구매카드별)';values.push(r1);
  values.push(pad(bizHeaders));summary.forEach(function(r){values.push(pad(r));});
  while(values.length<detailStart-1)values.push(blankRow());
  values.push(pad(detailHeaders));periodRows.forEach(function(r){values.push(pad(r));});
  for(var rr=0;rr<values.length;rr++){var f=[];for(var cc=0;cc<maxCols;cc++)f.push('General');formats.push(f);}
  for(var c1=0;c1<11;c1++)for(var r2=2;r2<2+summary.length;r2++)formats[r2][c1]='@';
  for(var c2=11;c2<=22;c2++)for(var r3=2;r3<2+summary.length;r3++)formats[r3][c2]='#,##0';
  for(var r4=detailStart;r4<values.length;r4++){[1,2,3,4,5].forEach(function(ci){formats[r4][ci]='@';});for(var ci=6;ci<=15;ci++)formats[r4][ci]='#,##0';}
  return{values:values,formats:formats,summaryRows:summary.length,detailStart:detailStart,detailHeaders:detailHeaders};
}
function i68ApplyLayout_(sh,summaryRows,detailStart,detailHeaders){
  var maxCols=24;sh.getRange(1,1,1,maxCols).setBackground('#b4c6e7').setFontWeight('bold');sh.getRange(2,1,1,24).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');sh.getRange(detailStart,1,1,detailHeaders.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');sh.setFrozenRows(2);
  var headers=['신고연도','반기','사업자등록번호','연결 쿠팡계정ID','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','카드매칭상태','카드매칭근거','주문건수','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];
  for(var i=0;i<maxCols;i++){var h=headers[i]||detailHeaders[i]||'';sh.setColumnWidth(i+1,/근거|비고/.test(h)?180:(/카드|사업자|계정/.test(h)?135:(/금액|부가세|이익|수수료/.test(h)?110:90)));}
}
function i68ReadPeriodBlock_(sh,detailStart){var last=sh.getLastRow();if(last<detailStart)return{halfTotals:i68Totals_([]),monthTotals:i68Totals_([]),periodUnknown:0};var vals=sh.getRange(detailStart+1,1,last-detailStart,16).getValues(),half=vals.filter(function(r){return i68Text_(r[0])==='반기';}),month=vals.filter(function(r){return i68Text_(r[0])==='월별';}),unknown=vals.filter(function(r){return i68Text_(r[0])==='기간미확인';}).reduce(function(n,r){return n+i68Num_(r[6]);},0);return{halfTotals:i68Totals_(half),monthTotals:i68Totals_(month),periodUnknown:unknown};}
function i68CaptureFilter_(sh){var f=sh.getFilter(),o={exists:!!f,range:'',startRow:0,numRows:0,startCol:0,numCols:0,criteriaCount:0,criteria:{}};if(!f)return o;var r=f.getRange();o.range=r.getA1Notation();o.startRow=r.getRow();o.numRows=r.getNumRows();o.startCol=r.getColumn();o.numCols=r.getNumColumns();for(var c=o.startCol;c<o.startCol+o.numCols;c++){try{var cr=f.getColumnFilterCriteria(c);if(cr){o.criteria[c]=cr;o.criteriaCount++;}}catch(ignore){}}return o;}
function i68RemoveFilter_(sh){var f=sh.getFilter();if(f)f.remove();}
function i68RestoreFilter_(sh,m){if(!m||!m.exists)return;var rows=Math.min(m.numRows,sh.getMaxRows()-m.startRow+1),cols=Math.min(m.numCols,sh.getMaxColumns()-m.startCol+1);if(rows<1||cols<1)return;var f=sh.getRange(m.startRow,m.startCol,rows,cols).createFilter();Object.keys(m.criteria||{}).forEach(function(k){try{f.setColumnFilterCriteria(Number(k),m.criteria[k]);}catch(ignore){}});}
function i68RestoreFilterFromSheet_(dst,src){i68RestoreFilter_(dst,i68CaptureFilter_(src));}
function i68Protected_(ss){var names=[I68_VAT,I68_CARD,I68_CARD_PREVIEW,I68_BIZ_PREVIEW,I68_I67,I68_I66,I68_HISTORY,I68_MASTER,I68_CARD_BACKUP],o={};names.forEach(function(n){o[n]=i68Sig_(i68Need_(ss,n));});return o;}
function i68CheckProtected_(ss,before){Object.keys(before).forEach(function(n){if(i68Sig_(i68Need_(ss,n))!==before[n])throw new Error('보호시트 변경 '+n);});}
function i68Sig_(sh){var v=sh.getDataRange().getValues(),h=2166136261;for(var r=0;r<v.length;r++)for(var c=0;c<v[r].length;c++){var s=i68Cell_(v[r][c])+'\u001f';for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619);}}return sh.getLastRow()+'x'+sh.getLastColumn()+'|'+(h>>>0).toString(16);}
function i68FormulaCount_(sh){var f=sh.getDataRange().getFormulas(),n=0;f.forEach(function(r){r.forEach(function(v){if(v)n++;});});return n;}
function i68SheetTypedDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;return i68RangeTypedDiff_(a.getDataRange(),b.getDataRange());}
function i68SheetDisplayDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;return i68RangeDisplayDiff_(a.getDataRange(),b.getDataRange());}
function i68SheetFormatDiff_(a,b){if(a.getLastRow()!==b.getLastRow()||a.getLastColumn()!==b.getLastColumn())return 999999;var A=a.getDataRange().getNumberFormats(),B=b.getDataRange().getNumberFormats(),d=0;for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++)if(String(A[r][c])!==String(B[r][c]))d++;return d;}
function i68RangeTypedDiff_(a,b){var A=a.getValues(),B=b.getValues(),d=0;if(A.length!==B.length||A[0].length!==B[0].length)return 999999;for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++)if(i68Cell_(A[r][c])!==i68Cell_(B[r][c]))d++;return d;}
function i68RangeDisplayDiff_(a,b){var A=a.getDisplayValues(),B=b.getDisplayValues(),d=0;if(A.length!==B.length||A[0].length!==B[0].length)return 999999;for(var r=0;r<A.length;r++)for(var c=0;c<A[r].length;c++)if(String(A[r][c])!==String(B[r][c]))d++;return d;}
function i68Cell_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return'D:'+v.toISOString();if(typeof v==='number')return'N:'+String(v);if(typeof v==='boolean')return'B:'+String(v);return'T:'+i68Text_(v);}
function i68Assert_(a,e){Object.keys(e).forEach(function(k){if(Math.round(Number(a[k]||0))!==Math.round(Number(e[k]||0)))throw new Error(k+' 불일치 실제 '+a[k]+' 기대 '+e[k]);});}
function i68ExpectKv_(ss,n,e){var sh=i68Need_(ss,n),kv={};sh.getRange(1,1,sh.getLastRow(),Math.min(2,sh.getLastColumn())).getValues().forEach(function(r){var k=i68Text_(r[0]);if(k)kv[k]=r[1];});Object.keys(e).forEach(function(k){var w=e[k],a=kv[k];if(typeof w==='number'){if(Math.round(i68Num_(a))!==w)throw new Error(n+' '+k+' 불일치 '+a);}else if(i68Text_(a)!==String(w))throw new Error(n+' '+k+' 불일치 '+a+' / 기대 '+w);});}
function i68Status_(sh,status,stage,msg,x){x=x||{};var rows=[['항목','값'],['버전',I68_VERSION],['상태',status],['단계',stage],['메시지',msg],['사업자×카드 요약행',x.summaryRows||0],['기간집계행',x.periodRows||0],['반기집계행',x.halfRows||0],['월별집계행',x.monthRows||0],['기간미확인',x.periodUnknown||0],['주문건수',x.orders||0],['순수매출액',x.sales||0],['정산기준금액',x.settlement||0],['마켓수수료',x.fee||0],['매입금액',x.purchase||0],['매출부가세',x.salesVat||0],['매입부가세',x.purchaseVat||0],['납부예상부가세',x.payable||0],['예상이익',x.profit||0],['부가세반영예상이익',x.vatProfit||0],['준비/Issue67 typed차이',x.prepBizTyped||0],['준비/Issue67 display차이',x.prepBizDisplay||0],['운영/준비 typed차이',x.prodTyped||0],['운영/준비 display차이',x.prodDisplay||0],['운영/준비 format차이',x.prodFormat||0],['운영/Issue67 typed차이',x.prodBizTyped||0],['운영/Issue67 display차이',x.prodBizDisplay||0],['백업 typed차이',x.backupTyped||0],['백업 display차이',x.backupDisplay||0],['기존filter복원',x.filterRestored||''],['기존filterCriteria수',x.filterCriteria||0],['부가세_신고자료 변경',0],['부가세_카드매칭검증 변경',0],['Issue54preview 변경',0],['Issue67preview 변경',0],['카드사용내역_붙여넣기 변경',0],['카드_마스터 변경',0],['Issue59백업 변경',0],['롤백',x.rollback||0],['오류',x.error||''],['완료시각',(status==='PASS'||status==='ERROR'||status==='ROLLED_BACK'||status==='ROLLBACK_ERROR')?new Date().toISOString():''],['갱신시각',new Date().toISOString()]];sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setColumnWidth(1,280);sh.setColumnWidth(2,760);}
function i68Need_(ss,n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('필수 시트 없음 '+n);return sh;}
function i68Ensure_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function i68EnsureGrid_(sh,r,c){if(sh.getMaxRows()<r)sh.insertRowsAfter(sh.getMaxRows(),r-sh.getMaxRows());if(sh.getMaxColumns()<c)sh.insertColumnsAfter(sh.getMaxColumns(),c-sh.getMaxColumns());}
function i68Find_(h,names){for(var n=0;n<names.length;n++){var w=i68Compact_(names[n]);for(var i=0;i<h.length;i++)if(i68Compact_(h[i])===w)return i;}return-1;}
function i68Text_(v){return String(v==null?'':v).trim();}
function i68Compact_(v){return i68Text_(v).replace(/\s/g,'');}
function i68Num_(v){var n=Number(typeof v==='number'?v:i68Text_(v).replace(/[,원\s]/g,''));return isNaN(n)?0:n;}
function i68Key_(a,o){a=i68Text_(a).toLowerCase();o=i68Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return a&&o?a+'|'+o:'';}
