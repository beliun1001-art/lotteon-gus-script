var LOTTEON_REMOTE_TASK={id:'ISSUE79-NOMATCH22-EXTRACT-v2',title:'NO_MATCH 22 주문번호 텍스트 안전 추출',enabled:true,statusSheet:'ISSUE79_NOMATCH22_주문번호'};
function runLotteonRemoteTaskStartRemote_(){return i79NomatchExtractV2_();}
function runLotteonRemoteTaskContinueRemote_(){return i79NomatchExtractV2_();}
function i79NomatchExtractV2_(){
  var ss=SpreadsheetApp.getActive();
  var src=ss.getSheetByName('부가세_카드매칭검증');
  if(!src)throw new Error('부가세_카드매칭검증 시트 누락');
  var raw=src.getDataRange().getValues(),disp=src.getDataRange().getDisplayValues(),hr=-1,h=null;
  for(var r=0;r<Math.min(raw.length,30);r++){
    var row=raw[r]||[],oi=i79v2ix_(row,['주문번호']),si=i79v2ix_(row,['카드매칭상태']);
    if(oi>=0&&si>=0){hr=r;h=row;break;}
  }
  if(hr<0)throw new Error('주문번호/카드매칭상태 header 탐지 실패');
  var orderIx=i79v2ix_(h,['주문번호']),statusIx=i79v2ix_(h,['카드매칭상태']);
  var orders=[];
  for(var i=hr+1;i<raw.length;i++){
    var st=i79v2t_(raw[i][statusIx]).toUpperCase();
    if(st!=='NO_MATCH')continue;
    var d=i79v2t_(disp[i][orderIx]),rv=raw[i][orderIx];
    var no=i79v2normalizeOrder_(d,rv);
    if(!no)throw new Error('NO_MATCH 주문번호 빈값 row '+(i+1));
    orders.push(no);
  }
  var uniq={},dups=[];orders.forEach(function(no){if(uniq[no])dups.push(no);uniq[no]=1;});
  if(orders.length!==22)throw new Error('NO_MATCH 수 '+orders.length+' (기대 22)');
  if(dups.length)throw new Error('NO_MATCH 주문번호 중복: '+dups.join(','));
  var csv=orders.join(',');
  var out=ss.getSheetByName('ISSUE79_NOMATCH22_주문번호')||ss.insertSheet('ISSUE79_NOMATCH22_주문번호');
  out.clear();
  out.getRange(1,1,30,3).setNumberFormat('@');
  var rows=[['항목','값','비고'],['version','v2.0-ISSUE79-NOMATCH22-TEXT-SAFE-EXTRACT',''],['상태','PASS',''],['NO_MATCH',String(orders.length),''],['주문번호CSV',csv,'쉼표 구분 텍스트'],['완료시각',new Date().toISOString(),'']];
  out.getRange(1,1,rows.length,3).setValues(rows);
  out.getRange(8,1,1,2).setValues([['순번','주문번호']]);
  out.getRange(9,1,orders.length,2).setValues(orders.map(function(no,idx){return [String(idx+1),no];}));
  out.setFrozenRows(1);out.getRange(1,1,1,3).setFontWeight('bold');out.getRange(8,1,1,2).setFontWeight('bold');out.autoResizeColumns(1,3);
  return {ok:true,done:true,count:orders.length,orders:orders,csv:csv};
}
function i79v2normalizeOrder_(display,raw){
  var d=i79v2t_(display);
  if(d){
    var compact=d.replace(/\s+/g,'');
    if(/^[0-9,]+$/.test(compact))return compact.replace(/,/g,'');
    if(/^[0-9]+$/.test(compact))return compact;
    return d;
  }
  if(typeof raw==='number'){
    if(!isFinite(raw))return '';
    return Utilities.formatString('%.0f',raw);
  }
  return i79v2t_(raw);
}
function i79v2ix_(h,names){var m={};(h||[]).forEach(function(v,i){m[i79v2c_(v)]=i;});for(var j=0;j<names.length;j++){var k=i79v2c_(names[j]);if(Object.prototype.hasOwnProperty.call(m,k))return m[k];}return -1;}
function i79v2t_(v){if(v===null||v===undefined)return '';return String(v).trim();}
function i79v2c_(v){return i79v2t_(v).toLowerCase().replace(/[\s_\-\/.()\[\]:]+/g,'');}
