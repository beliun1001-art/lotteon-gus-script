var LOTTEON_REMOTE_TASK={id:'ISSUE79-NOMATCH22-EXTRACT-v1',title:'NO_MATCH 22 주문번호 읽기전용 추출',enabled:true,statusSheet:'ISSUE79_NOMATCH22_주문번호'};
function runLotteonRemoteTaskStartRemote_(){return i79NomatchExtract_();}
function runLotteonRemoteTaskContinueRemote_(){return i79NomatchExtract_();}
function i79NomatchExtract_(){
  var ss=SpreadsheetApp.getActive();
  var src=ss.getSheetByName('부가세_카드매칭검증');
  if(!src)throw new Error('부가세_카드매칭검증 시트 누락');
  var v=src.getDataRange().getValues(),hr=-1,h=null;
  for(var r=0;r<Math.min(v.length,30);r++){
    var row=v[r]||[],oi=i79ix_(row,['주문번호']),si=i79ix_(row,['카드매칭상태']);
    if(oi>=0&&si>=0){hr=r;h=row;break;}
  }
  if(hr<0)throw new Error('주문번호/카드매칭상태 header 탐지 실패');
  var orderIx=i79ix_(h,['주문번호']),statusIx=i79ix_(h,['카드매칭상태']);
  var orders=[];
  for(var i=hr+1;i<v.length;i++){
    var no=i79t_(v[i][orderIx]),st=i79t_(v[i][statusIx]).toUpperCase();
    if(no&&st==='NO_MATCH')orders.push(no);
  }
  var uniq={},dups=[];orders.forEach(function(no){if(uniq[no])dups.push(no);uniq[no]=1;});
  if(orders.length!==22)throw new Error('NO_MATCH 수 '+orders.length+' (기대 22)');
  if(dups.length)throw new Error('NO_MATCH 주문번호 중복: '+dups.join(','));
  var csv=orders.join(',');
  var out=ss.getSheetByName('ISSUE79_NOMATCH22_주문번호')||ss.insertSheet('ISSUE79_NOMATCH22_주문번호');
  out.clearContents();
  var rows=[['항목','값'],['version','v1.0-ISSUE79-NOMATCH22-READONLY-EXTRACT'],['상태','PASS'],['NO_MATCH',orders.length],['주문번호CSV',csv],['완료시각',new Date().toISOString()]];
  out.getRange(1,1,rows.length,2).setValues(rows);out.setFrozenRows(1);out.getRange(1,1,1,2).setFontWeight('bold');out.autoResizeColumns(1,2);
  return {ok:true,done:true,count:orders.length,orders:orders,csv:csv};
}
function i79ix_(h,names){var m={};(h||[]).forEach(function(v,i){m[i79c_(v)]=i;});for(var j=0;j<names.length;j++){var k=i79c_(names[j]);if(Object.prototype.hasOwnProperty.call(m,k))return m[k];}return -1;}
function i79t_(v){if(v===null||v===undefined)return '';return String(v).trim();}
function i79c_(v){return i79t_(v).toLowerCase().replace(/[\s_\-\/.()\[\]:]+/g,'');}
