/**
 * Issue #53 v1.0 corrected Apr-Jun VAT production apply.
 * Explicitly requires P=`마켓주문상태` and excludes 취소/반품/교환/환불.
 * Rebuilds only `부가세_신고자료`; protects card verification and rolls back on mismatch.
 */
var LOTTEON_REMOTE_TASK = {
  id: 'ISSUE53-v1.0-20260813',
  title: '2026년 4~6월 취소상태 제외 VAT 운영 재생성',
  enabled: true,
  outputSheet: '부가세_신고자료',
  statusSheet: 'ISSUE53_운영재생성상태'
};

function runLotteonRemoteTaskStartRemote_() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var status = i53Ensure_(ss, LOTTEON_REMOTE_TASK.statusSheet);
  i53WriteStatus_(status,[
    ['항목','값'],['버전','v1.0-ISSUE53-CORRECTED-VAT-PRODUCTION'],['상태','RUNNING'],['단계','PRECHECK'],
    ['메시지','취소상태 제외 corrected VAT 운영 재생성 사전검증 시작'],['운영시트 변경','0']
  ]);

  var detail = ss.getSheetByName('부가세_신고자료');
  var verify = ss.getSheetByName('부가세_카드매칭검증');
  if (!detail) throw new Error('부가세_신고자료 시트가 없습니다.');
  if (!verify) throw new Error('부가세_카드매칭검증 시트가 없습니다.');

  var oldDetail = detail.getDataRange().getValues();
  var oldVerify = verify.getDataRange().getValues();
  var verifySigBefore = i53Signature_(oldVerify);
  var wrote = false;

  try {
    var built = i53Build_(ss);
    i53AssertExpected_(built);
    i53Req_(oldVerify.length-1===1355,'기존 카드검증 주문수 보호 기준 불일치: '+(oldVerify.length-1));

    var headers = i53Headers_();
    detail.clearContents();
    detail.getRange(1,1,1,headers.length).setValues([headers]);
    if (built.rows.length) detail.getRange(2,1,built.rows.length,headers.length).setValues(built.rows);
    detail.getRange(1,1,1,headers.length).setFontWeight('bold');
    detail.setFrozenRows(1);
    SpreadsheetApp.flush();
    wrote = true;

    var post = i53ValidateDetail_(detail);
    i53Req_(post.rows===2752,'작성상세행 불일치: '+post.rows);
    i53Req_(post.orders===1355,'작성 고유주문수 불일치: '+post.orders);
    i53Req_(post.sales===138432300,'작성 순수매출합계 불일치: '+post.sales);
    i53Req_(post.settlement===122495855,'작성 정산기준금액합계 불일치: '+post.settlement);
    i53Req_(post.purchase===105762969,'작성 매입금액합계 불일치: '+post.purchase);
    i53Req_(post.salesVat===12584695,'작성 매출부가세합계 불일치: '+post.salesVat);
    i53Req_(post.purchaseVat===9614786,'작성 매입부가세합계 불일치: '+post.purchaseVat);
    i53Req_(post.payableVat===2969909,'작성 납부예상부가세합계 불일치: '+post.payableVat);
    i53Req_(post.unmapped===0,'작성 사업자번호 미매핑 존재: '+post.unmapped);

    var verifyAfter = verify.getDataRange().getValues();
    i53Req_(i53Signature_(verifyAfter)===verifySigBefore,'부가세_카드매칭검증이 변경되었습니다.');

    i53WriteStatus_(status,[
      ['항목','값'],['버전','v1.0-ISSUE53-CORRECTED-VAT-PRODUCTION'],['상태','PASS'],['단계','DONE'],
      ['메시지','2026년 4~6월 취소상태 제외 VAT 운영 재생성 및 검증 완료'],
      ['운영시트 변경','부가세_신고자료 1개 재작성'],
      ['상태선택열','P / 마켓주문상태'],['취소/반품/교환/환불제외행',built.excludedRows],['취소상태고유주문',built.excludedOrders],
      ['작성상세행',post.rows],['고유주문수',post.orders],['사업자번호미매핑',post.unmapped],['계정수',built.accountCount],
      ['순수매출합계',post.sales],['정산기준금액합계',post.settlement],['매입금액합계',post.purchase],
      ['매출부가세합계',post.salesVat],['매입부가세합계',post.purchaseVat],['납부예상부가세합계',post.payableVat],
      ['정산fallback행',built.fallbackRows],
      ['2026-04_상세행',built.months['2026-04'].rows],['2026-04_고유주문',built.months['2026-04'].orders],
      ['2026-05_상세행',built.months['2026-05'].rows],['2026-05_고유주문',built.months['2026-05'].orders],
      ['2026-06_상세행',built.months['2026-06'].rows],['2026-06_고유주문',built.months['2026-06'].orders],
      ['R헤더',headers[17]],['S헤더',headers[18]],['T헤더',headers[19]],
      ['카드매칭검증 변경','0'],['롤백','없음'],['완료시각',new Date().toISOString()]
    ]);
    return {ok:true,done:true,rows:post.rows,orders:post.orders};
  } catch (e) {
    var rollback = '불필요';
    if (wrote) {
      try {
        detail.clearContents();
        if (oldDetail.length && oldDetail[0].length) detail.getRange(1,1,oldDetail.length,oldDetail[0].length).setValues(oldDetail);
        SpreadsheetApp.flush();
        rollback = '기존 부가세_신고자료 복구 완료';
      } catch (rb) {
        rollback = '롤백 실패: '+String(rb&&rb.message?rb.message:rb);
      }
    }
    i53WriteStatus_(status,[
      ['항목','값'],['버전','v1.0-ISSUE53-CORRECTED-VAT-PRODUCTION'],['상태','ERROR'],['단계','FAILED'],
      ['메시지','corrected VAT 운영 재생성 실패'],['오류',String(e&&e.message?e.message:e)],
      ['운영시트 변경',wrote?'시도 후 롤백':'0'],['롤백',rollback],['완료시각',new Date().toISOString()]
    ]);
    throw e;
  }
}

function runLotteonRemoteTaskContinueRemote_(){return {ok:true,done:true,reason:'NO_CONTINUE_REQUIRED'};}

function i53Build_(ss){
  var source=ss.getSheetByName('매출데이터_붙여넣기');
  if(!source||source.getLastRow()<2)throw new Error('매출데이터_붙여넣기가 없습니다.');
  var v=source.getDataRange().getValues(), h=v[0].map(i53Text_);
  i53Req_(h.length>=39,'원천 AM열까지 존재하지 않습니다.');
  i53Req_(i53Compact_(h[0])===i53Compact_('마켓주문일자'),'A 헤더 불일치: '+h[0]);
  i53Req_(i53Compact_(h[2])===i53Compact_('마켓주문번호'),'C 헤더 불일치: '+h[2]);
  i53Req_(i53Compact_(h[3])===i53Compact_('마켓아이디'),'D 헤더 불일치: '+h[3]);
  i53Req_(i53Compact_(h[6])===i53Compact_('결제금액합계(원)'),'G 헤더 불일치: '+h[6]);
  i53Req_(i53Compact_(h[15])===i53Compact_('마켓주문상태'),'P 헤더 불일치: '+h[15]);
  i53Req_(i53Compact_(h[28])===i53Compact_('구매가격'),'AC 헤더 불일치: '+h[28]);
  i53Req_(i53Compact_(h[38])===i53Compact_('정산예정금액(원)'),'AM 헤더 불일치: '+h[38]);

  var ix={
    customer:i53Find_(h,['고객명','수령인','수취인','구매자','주문자']),
    brand:i53Find_(h,['브랜드명','브랜드']),
    productNo:i53Find_(h,['마켓상품번호','상품번호','상품코드','판매자상품코드']),
    productName:i53Find_(h,['상품명','상품명(옵션포함)','등록상품명']),
    quantity:i53Find_(h,['판매수량','수량','구매수량'])
  };
  var rows=[], excludedRows=0, excludedKeys={}, fallbackRows=0, unmapped=0, accounts={}, orders={};
  var months={'2026-04':{rows:0,keys:{}},'2026-05':{rows:0,keys:{}},'2026-06':{rows:0,keys:{}}};
  var sums={sales:0,settlement:0,purchase:0,salesVat:0,purchaseVat:0,payableVat:0};

  for(var r=1;r<v.length;r++){
    var row=v[r], date=i53Date_(row[0]);
    if(!date||date<'2026-04-01'||date>'2026-06-30')continue;
    var status=i53Text_(row[15]);
    if(/취소|반품|교환|환불/.test(status)){
      excludedRows++;
      var ek=i53OrderKey_(row[3],row[2]); if(ek)excludedKeys[ek]=true;
      continue;
    }
    var sales=i53Num_(row[6]);
    if(!sales)continue;
    var account=i53Text_(row[3]);
    var business=i53Business_(account);
    if(!business)unmapped++;
    var settlementActual=i53Num_(row[38]);
    var settlement=settlementActual||Math.round(sales*0.901);
    if(!settlementActual)fallbackRows++;
    var purchase=i53Num_(row[28]);
    var sv=i53Split_(sales), pv=i53Split_(purchase);
    var fee=sales-settlement, profit=settlement-purchase, payable=sv.vat-pv.vat;
    var month=date.slice(0,7), orderNo=i53Text_(row[2]), key=i53OrderKey_(account,orderNo);
    var qty=ix.quantity>=0?i53Num_(row[ix.quantity]):0; if(!qty)qty=1;
    rows.push([
      '2026','상반기',month,date,account,business,orderNo,
      ix.customer>=0?i53Text_(row[ix.customer]):'',ix.brand>=0?i53Text_(row[ix.brand]):'',
      ix.productNo>=0?i53Text_(row[ix.productNo]):'',ix.productName>=0?i53Text_(row[ix.productName]):'',qty,
      sales,sv.supply,sv.vat,settlement,fee,purchase,pv.supply,pv.vat,payable,profit,profit-payable,
      business?'ISSUE53 P=마켓주문상태 취소제외':'사업자번호 미매핑'
    ]);
    accounts[account]=true;if(key)orders[key]=true;
    if(months[month]){months[month].rows++;if(key)months[month].keys[key]=true;}
    sums.sales+=sales;sums.settlement+=settlement;sums.purchase+=purchase;sums.salesVat+=sv.vat;sums.purchaseVat+=pv.vat;sums.payableVat+=payable;
  }
  Object.keys(months).forEach(function(m){months[m].orders=Object.keys(months[m].keys).length;});
  return {rows:rows,orders:Object.keys(orders).length,unmapped:unmapped,accountCount:Object.keys(accounts).length,excludedRows:excludedRows,excludedOrders:Object.keys(excludedKeys).length,fallbackRows:fallbackRows,months:months,sums:sums};
}

function i53AssertExpected_(b){
  i53Req_(b.rows.length===2752,'corrected 상세행 불일치: '+b.rows.length);
  i53Req_(b.orders===1355,'corrected 고유주문 불일치: '+b.orders);
  i53Req_(b.unmapped===0,'사업자번호 미매핑: '+b.unmapped);
  i53Req_(b.accountCount===4,'계정수 불일치: '+b.accountCount);
  i53Req_(b.excludedRows===1142,'취소/반품/교환/환불 제외행 불일치: '+b.excludedRows);
  i53Req_(b.excludedOrders===553,'취소상태 고유주문 불일치: '+b.excludedOrders);
  i53Req_(b.fallbackRows===28,'정산 fallback행 불일치: '+b.fallbackRows);
  i53Req_(Math.round(b.sums.sales)===138432300,'corrected 순수매출 불일치: '+Math.round(b.sums.sales));
  i53Req_(Math.round(b.sums.settlement)===122495855,'corrected 정산 불일치: '+Math.round(b.sums.settlement));
  i53Req_(Math.round(b.sums.purchase)===105762969,'corrected 매입 불일치: '+Math.round(b.sums.purchase));
  i53Req_(Math.round(b.sums.salesVat)===12584695,'corrected 매출부가세 불일치: '+Math.round(b.sums.salesVat));
  i53Req_(Math.round(b.sums.purchaseVat)===9614786,'corrected 매입부가세 불일치: '+Math.round(b.sums.purchaseVat));
  i53Req_(Math.round(b.sums.payableVat)===2969909,'corrected 납부예상부가세 불일치: '+Math.round(b.sums.payableVat));
  i53Req_(b.months['2026-04'].rows===2&&b.months['2026-04'].orders===1,'4월 집계 불일치');
  i53Req_(b.months['2026-05'].rows===638&&b.months['2026-05'].orders===322,'5월 집계 불일치');
  i53Req_(b.months['2026-06'].rows===2112&&b.months['2026-06'].orders===1032,'6월 집계 불일치');
}

function i53ValidateDetail_(sheet){
  var v=sheet.getDataRange().getValues(), h=v[0].map(i53Text_), ix={};h.forEach(function(x,i){ix[x]=i;});
  ['쿠팡계정ID','사업자등록번호','주문번호','순수매출액','정산기준금액','매입금액','매출부가세','매입부가세','납부예상부가세'].forEach(function(k){i53Req_(ix[k]>=0,'운영 상세 헤더 누락: '+k);});
  i53Req_(h[17]==='매입금액'&&h[18]==='매입공급가액'&&h[19]==='매입부가세','R/S/T 헤더 불일치');
  var keys={}, sales=0,settlement=0,purchase=0,salesVat=0,purchaseVat=0,payableVat=0,unmapped=0;
  for(var r=1;r<v.length;r++){
    var row=v[r], key=i53OrderKey_(row[ix['쿠팡계정ID']],row[ix['주문번호']]);if(key)keys[key]=true;
    if(!i53Text_(row[ix['사업자등록번호']]))unmapped++;
    sales+=i53Num_(row[ix['순수매출액']]);settlement+=i53Num_(row[ix['정산기준금액']]);purchase+=i53Num_(row[ix['매입금액']]);
    salesVat+=i53Num_(row[ix['매출부가세']]);purchaseVat+=i53Num_(row[ix['매입부가세']]);payableVat+=i53Num_(row[ix['납부예상부가세']]);
  }
  return {rows:v.length-1,orders:Object.keys(keys).length,sales:Math.round(sales),settlement:Math.round(settlement),purchase:Math.round(purchase),salesVat:Math.round(salesVat),purchaseVat:Math.round(purchaseVat),payableVat:Math.round(payableVat),unmapped:unmapped};
}

function i53Headers_(){return ['신고연도','반기','신고월','날짜','쿠팡계정ID','사업자등록번호','주문번호','고객명','브랜드명','상품번호','상품명','판매수량','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료/비용','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];}
function i53Business_(v){var s=i53Text_(v).toLowerCase();if(s==='beliun1021'||s==='1021')return '227-27-04928';if(s==='beliun1021-1'||s==='1021-1')return '176-71-00758';if(s==='beliun1023'||s==='1023')return '835-58-00765';if(s==='beliun1024'||s==='1024')return '606-45-93763';return '';}
function i53Find_(h,names){for(var n=0;n<names.length;n++){var q=i53Compact_(names[n]);for(var i=0;i<h.length;i++)if(i53Compact_(h[i])===q)return i;}return -1;}
function i53Split_(a){var total=Math.round(i53Num_(a)), supply=Math.round(total/1.1);return {total:total,supply:supply,vat:total-supply};}
function i53Num_(v){if(typeof v==='number')return isFinite(v)?v:0;var s=i53Text_(v);if(!s)return 0;var n=Number(s.replace(/[원,%\s,]/g,''));return isFinite(n)?n:0;}
function i53Date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i53Text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(m)return m[1]+'-'+i53Pad_(m[2])+'-'+i53Pad_(m[3]);if(/^\d{2}[.\/-]\d{1,2}$/.test(s)){m=s.match(/^(\d{2})[.\/-](\d{1,2})$/);return '2026-'+i53Pad_(m[1])+'-'+i53Pad_(m[2]);}return '';}
function i53OrderKey_(a,o){var aa=i53Text_(a).toLowerCase(),oo=i53Text_(o).toLowerCase().replace(/[^0-9a-z가-힣]/g,'');return aa&&oo?aa+'|'+oo:'';}
function i53Signature_(v){var s=JSON.stringify(v),h=2166136261;for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h+=(h<<1)+(h<<4)+(h<<7)+(h<<8)+(h<<24);}return String(h>>>0)+'|'+s.length;}
function i53Text_(v){return String(v==null?'':v).trim();}
function i53Compact_(v){return i53Text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function i53Pad_(v){v=String(v);return v.length<2?'0'+v:v;}
function i53Req_(c,m){if(!c)throw new Error(m);}
function i53Ensure_(ss,name){return ss.getSheetByName(name)||ss.insertSheet(name);}
function i53WriteStatus_(sh,rows){sh.clearContents();if(rows.length)sh.getRange(1,1,rows.length,2).setValues(rows);sh.getRange(1,1,1,2).setFontWeight('bold');sh.setFrozenRows(1);SpreadsheetApp.flush();}
