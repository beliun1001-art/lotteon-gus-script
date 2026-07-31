/** PR #25 v6.66 tracking-payment standalone operating smoke runner R5.
 * No GitHub/UrlFetch calls. Reads local spreadsheet sheets only.
 * Writes only:
 * - 부가세_신고자료: 롯데결제수단 column, after full row-mapping validation
 * - PR25_카드매칭검증
 * - PR25_사업자별반기요약
 * - PR25_결제수단월별검증
 */
const LOTTEON_PR25_STANDALONE_VERSION = 'v1.16-PR25-STANDALONE-R5';

function testPr25StandaloneReady() {
  const ss = SpreadsheetApp.getActive();
  const required = ['매출데이터_붙여넣기', '부가세_신고자료', '카드사용내역_붙여넣기'];
  const missing = required.filter(function(name) { return !ss.getSheetByName(name); });
  if (missing.length) throw new Error('필수 시트 누락: ' + missing.join(', '));
  SpreadsheetApp.getUi().alert(
    'PR #25 단일형 준비 완료\n\n' +
    '버전: ' + LOTTEON_PR25_STANDALONE_VERSION + '\n' +
    'GitHub 원격 호출: 없음\n' +
    '필수 시트: 확인 완료\n\n' +
    '이제 runPr25StandaloneTrackingMatch를 실행하세요.'
  );
}

function runPr25StandaloneTrackingMatch() {
  const started = Date.now();
  const ss = SpreadsheetApp.getActive();
  const backfill = pr25r5_backfillTrackingPayment_(ss);
  const detail = ss.getSheetByName('부가세_신고자료');
  const orders = pr25r5_groupOrders_(detail.getDataRange().getValues());
  const history = pr25r5_loadHistory_(ss);
  const master = pr25r5_loadMaster_(ss);
  const canonical = pr25r5_canonicalize_(history, master);
  const stats = pr25r5_allocate_(orders, canonical, master);
  pr25r5_writeDiagnostic_(ss, orders);
  pr25r5_writeSummary_(ss, orders);
  pr25r5_writeMonthlyPayment_(ss, orders);
  SpreadsheetApp.flush();
  const elapsed = Math.round((Date.now() - started) / 1000);
  SpreadsheetApp.getUi().alert(
    'PR #25 단일형 재매칭 완료\n\n' +
    '버전: ' + LOTTEON_PR25_STANDALONE_VERSION + '\n' +
    '상세행: ' + backfill.detailRows.toLocaleString('ko-KR') + '건\n' +
    '결제수단 입력행: ' + backfill.paymentRows.toLocaleString('ko-KR') + '건\n' +
    '주문: ' + orders.length.toLocaleString('ko-KR') + '건\n' +
    'MATCHED: ' + stats.matched + '건\n' +
    'NON_CARD: ' + stats.nonCard + '건\n' +
    'AMBIGUOUS: ' + stats.ambiguous + '건\n' +
    'NO_MATCH: ' + stats.noMatch + '건\n' +
    '소요: ' + elapsed + '초\n\n' +
    'PR25_사업자별반기요약과 PR25_카드매칭검증을 확인하세요.'
  );
  return { backfill: backfill, orders: orders.length, stats: stats, seconds: elapsed };
}

function pr25r5_backfillTrackingPayment_(ss) {
  const source = ss.getSheetByName('매출데이터_붙여넣기');
  const detail = ss.getSheetByName('부가세_신고자료');
  if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 데이터가 없습니다.');
  if (!detail || detail.getLastRow() < 2) throw new Error('부가세_신고자료 데이터가 없습니다.');

  const sourceMaxCol = Math.min(29, source.getLastColumn());
  const sourceValues = source.getRange(1, 1, source.getLastRow(), sourceMaxCol).getValues();
  const sh = sourceValues[0] || [];
  const sx = {
    date: pr25r5_find_(sh, ['마켓주문일자','주문일자','결제일자','주문일시'], 0),
    account: 3,
    orderNo: pr25r5_find_(sh, ['마켓주문번호','주문번호','주문ID','주문ID(마켓)'], 2),
    customer: pr25r5_find_(sh, ['고객명','수령인','수취인','구매자','주문자'], -1),
    brand: pr25r5_find_(sh, ['브랜드명','브랜드'], -1),
    productNo: pr25r5_find_(sh, ['마켓상품번호','상품번호','상품코드','판매자상품코드'], 4),
    productName: pr25r5_find_(sh, ['상품명','상품명(옵션포함)','등록상품명'], -1),
    quantity: pr25r5_find_(sh, ['판매수량','수량','구매수량'], -1),
    sales: pr25r5_find_(sh, ['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액'], 6),
    status: pr25r5_find_(sh, ['주문상태','상태','클레임상태','처리상태'], -1),
    tracking: pr25r5_find_(sh, ['트래킹 번호','트래킹번호','tracking number','trackingnumber'], -1),
    fallback: pr25r5_find_(sh, ['결제수단','결제정보','결제방법','카드사','결제수단/카드사','결제수단(카드사)','구매결제수단'], -1),
    purchase: 28
  };
  if (sx.tracking < 0 && sx.fallback < 0) throw new Error('원본에서 트래킹 번호 또는 결제수단 열을 찾지 못했습니다.');

  const detailValues = detail.getDataRange().getValues();
  const dh = detailValues[0] || [];
  let paymentIx = pr25r5_find_(dh, ['롯데결제수단'], -1);
  if (paymentIx < 0) {
    paymentIx = dh.length;
    detail.getRange(1, paymentIx + 1).setValue('롯데결제수단');
    dh.push('롯데결제수단');
    for (let r = 1; r < detailValues.length; r++) detailValues[r].push('');
  }
  const dx = {
    date: pr25r5_find_(dh, ['날짜','주문일','주문일자','마켓주문일자'], -1),
    account: pr25r5_find_(dh, ['쿠팡계정ID'], -1),
    orderNo: pr25r5_find_(dh, ['주문번호','마켓주문번호','주문ID','주문ID(마켓)'], -1),
    customer: pr25r5_find_(dh, ['고객명'], -1),
    brand: pr25r5_find_(dh, ['브랜드명'], -1),
    productNo: pr25r5_find_(dh, ['상품번호','마켓상품번호'], -1),
    productName: pr25r5_find_(dh, ['상품명'], -1),
    quantity: pr25r5_find_(dh, ['판매수량','수량'], -1),
    sales: pr25r5_find_(dh, ['순수매출액'], -1),
    purchase: pr25r5_find_(dh, ['매입금액'], -1),
    payment: paymentIx
  };
  Object.keys(dx).forEach(function(k) { if (dx[k] < 0) throw new Error('부가세_신고자료 필수 열 누락: ' + k); });

  const queues = {};
  let generated = 0;
  let generatedNonblank = 0;
  for (let r = 1; r < sourceValues.length; r++) {
    const row = sourceValues[r];
    const status = pr25r5_text_(pr25r5_at_(row, sx.status));
    if (/취소|반품|환불/.test(status)) continue;
    const sales = pr25r5_num_(pr25r5_at_(row, sx.sales));
    if (!sales) continue;
    const generatedRow = {
      date: pr25r5_date_(pr25r5_at_(row, sx.date)),
      account: pr25r5_text_(row[sx.account]),
      orderNo: pr25r5_text_(pr25r5_at_(row, sx.orderNo)),
      customer: pr25r5_text_(pr25r5_at_(row, sx.customer)),
      brand: pr25r5_text_(pr25r5_at_(row, sx.brand)),
      productNo: pr25r5_text_(pr25r5_at_(row, sx.productNo)),
      productName: pr25r5_text_(pr25r5_at_(row, sx.productName)),
      quantity: pr25r5_num_(pr25r5_at_(row, sx.quantity)) || 1,
      sales: sales,
      purchase: pr25r5_num_(row[sx.purchase])
    };
    const payment = pr25r5_text_(pr25r5_at_(row, sx.tracking)) || pr25r5_text_(pr25r5_at_(row, sx.fallback));
    const key = pr25r5_rowKey_(generatedRow);
    if (!queues[key]) queues[key] = [];
    queues[key].push(payment);
    generated++;
    if (payment) generatedNonblank++;
  }

  const detailRows = detailValues.slice(1);
  if (generated !== detailRows.length) {
    throw new Error('안전검증 실패: 원본 생성행 ' + generated + '건 / 상세행 ' + detailRows.length + '건. 아무 값도 쓰지 않았습니다.');
  }
  const payments = [];
  const unmatched = [];
  let paymentRows = 0;
  for (let i = 0; i < detailRows.length; i++) {
    const row = detailRows[i];
    const obj = {
      date: pr25r5_date_(row[dx.date]), account: pr25r5_text_(row[dx.account]), orderNo: pr25r5_text_(row[dx.orderNo]),
      customer: pr25r5_text_(row[dx.customer]), brand: pr25r5_text_(row[dx.brand]), productNo: pr25r5_text_(row[dx.productNo]),
      productName: pr25r5_text_(row[dx.productName]), quantity: pr25r5_num_(row[dx.quantity]) || 1,
      sales: pr25r5_num_(row[dx.sales]), purchase: pr25r5_num_(row[dx.purchase])
    };
    const key = pr25r5_rowKey_(obj);
    const q = queues[key];
    if (!q || !q.length) { unmatched.push(i + 2); payments.push(['']); continue; }
    const p = q.shift();
    payments.push([p]);
    if (p) paymentRows++;
  }
  let leftovers = 0;
  Object.keys(queues).forEach(function(k) { leftovers += queues[k].length; });
  if (unmatched.length || leftovers) {
    throw new Error('안전검증 실패: 상세 미매칭 ' + unmatched.length + '건 / 원본 잔여 ' + leftovers + '건. 아무 값도 쓰지 않았습니다.');
  }
  if (paymentRows < 1 || generatedNonblank < 1) throw new Error('원본 결제수단 값이 0건입니다. 아무 값도 쓰지 않았습니다.');
  detail.getRange(2, paymentIx + 1, payments.length, 1).setNumberFormat('@').setValues(payments);
  return { detailRows: detailRows.length, paymentRows: paymentRows, blankPaymentRows: detailRows.length - paymentRows };
}

function pr25r5_groupOrders_(values) {
  const h = values[0] || [];
  const x = {
    date: pr25r5_find_(h,['날짜','주문일','주문일자','마켓주문일자'],-1), year:pr25r5_find_(h,['신고연도'],-1),
    half:pr25r5_find_(h,['반기'],-1), month:pr25r5_find_(h,['신고월'],-1), account:pr25r5_find_(h,['쿠팡계정ID'],-1),
    business:pr25r5_find_(h,['사업자등록번호'],-1), orderNo:pr25r5_find_(h,['주문번호','마켓주문번호','주문ID','주문ID(마켓)'],-1),
    payment:pr25r5_find_(h,['롯데결제수단','구매결제수단','결제수단'],-1), sales:pr25r5_find_(h,['순수매출액'],-1),
    salesSupply:pr25r5_find_(h,['매출공급가액'],-1), salesVat:pr25r5_find_(h,['매출부가세'],-1), settlement:pr25r5_find_(h,['정산기준금액'],-1),
    fee:pr25r5_find_(h,['마켓수수료/비용','마켓수수료'],-1), purchase:pr25r5_find_(h,['매입금액'],-1),
    purchaseSupply:pr25r5_find_(h,['매입공급가액'],-1), purchaseVat:pr25r5_find_(h,['매입부가세'],-1), payable:pr25r5_find_(h,['납부예상부가세'],-1),
    profit:pr25r5_find_(h,['예상이익'],-1), vatProfit:pr25r5_find_(h,['부가세반영예상이익'],-1)
  };
  ['year','half','account','business','sales','salesSupply','salesVat','settlement','purchase','purchaseSupply','purchaseVat','payable','profit','vatProfit'].forEach(function(k){
    if (x[k] < 0) throw new Error('주문집계 필수 열 누락: ' + k);
  });
  const map = {};
  for (let r=1;r<values.length;r++) {
    const row=values[r], year=pr25r5_text_(row[x.year]), half=pr25r5_text_(row[x.half]);
    if (!year || year==='기간미확인' || (half!=='상반기' && half!=='하반기')) continue;
    const account=pr25r5_text_(row[x.account]), business=pr25r5_text_(row[x.business]);
    const orderNo=x.orderNo>=0?pr25r5_text_(row[x.orderNo]):'';
    const key=orderNo?[year,half,business,account,orderNo].join('|'):[year,half,business,account,'BLANK',r].join('|');
    if (!map[key]) map[key]={key:key,year:year,half:half,month:x.month>=0?pr25r5_text_(row[x.month]):'',orderDate:pr25r5_fullDate_(row[x.date],year),business:business,account:account,orderNo:orderNo,payments:{},detailRows:0,sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};
    const o=map[key]; o.detailRows++;
    const p=x.payment>=0?pr25r5_text_(row[x.payment]):''; if(p)o.payments[p]=true;
    o.sales+=pr25r5_num_(row[x.sales]); o.salesSupply+=pr25r5_num_(row[x.salesSupply]); o.salesVat+=pr25r5_num_(row[x.salesVat]);
    o.settlement+=pr25r5_num_(row[x.settlement]); o.fee+=x.fee>=0?pr25r5_num_(row[x.fee]):pr25r5_num_(row[x.sales])-pr25r5_num_(row[x.settlement]);
    o.purchase+=pr25r5_num_(row[x.purchase]); o.purchaseSupply+=pr25r5_num_(row[x.purchaseSupply]); o.purchaseVat+=pr25r5_num_(row[x.purchaseVat]);
    o.payable+=pr25r5_num_(row[x.payable]); o.profit+=pr25r5_num_(row[x.profit]); o.vatProfit+=pr25r5_num_(row[x.vatProfit]);
  }
  return Object.keys(map).map(function(k){const o=map[k];o.lottePayment=Object.keys(o.payments).sort().join(', ');return o;}).sort(function(a,b){return a.orderDate.localeCompare(b.orderDate)||a.orderNo.localeCompare(b.orderNo);});
}

function pr25r5_loadHistory_(ss) {
  const s=ss.getSheetByName('카드사용내역_붙여넣기'); if(!s||s.getLastRow()<2)return [];
  const v=s.getDataRange().getValues(),h=v[0]||[],f=function(n){return pr25r5_find_(h,n,-1);};
  const x={company:f(['카드사']),name:f(['카드명']),number:f(['카드번호']),end4:f(['카드번호끝4']),date:f(['승인일','이용일','거래일']),time:f(['승인시각','이용시각','거래시각']),merchant:f(['가맹점명','이용가맹점']),amount:f(['승인금액','이용금액','거래금액']),approval:f(['승인번호']),status:f(['승인상태','승인/취소구분','상태']),cancelDate:f(['취소일']),cancelAmount:f(['취소금액']),orderNo:f(['가맹점주문번호','주문번호']),evidence:f(['증빙유형']),lotte:f(['롯데계열여부']),source:f(['원본파일']),memo:f(['메모'])};
  const out=[];
  for(let r=1;r<v.length;r++){
    const row=v[r],o={rowNo:r+1,company:pr25r5_text_(pr25r5_at_(row,x.company)),cardName:pr25r5_text_(pr25r5_at_(row,x.name)),cardNumber:pr25r5_text_(pr25r5_at_(row,x.number)),cardEnd4:pr25r5_text_(pr25r5_at_(row,x.end4)),date:pr25r5_fullDate_(pr25r5_at_(row,x.date),''),time:pr25r5_time_(pr25r5_at_(row,x.time)),merchant:pr25r5_text_(pr25r5_at_(row,x.merchant)),amount:pr25r5_num_(pr25r5_at_(row,x.amount)),approvalNo:pr25r5_text_(pr25r5_at_(row,x.approval)),status:pr25r5_text_(pr25r5_at_(row,x.status)),cancelDate:pr25r5_fullDate_(pr25r5_at_(row,x.cancelDate),''),cancelAmount:Math.abs(pr25r5_num_(pr25r5_at_(row,x.cancelAmount))),merchantOrderNo:pr25r5_text_(pr25r5_at_(row,x.orderNo)),evidenceType:pr25r5_text_(pr25r5_at_(row,x.evidence)),lotteFlag:pr25r5_text_(pr25r5_at_(row,x.lotte)),sourceFile:pr25r5_text_(pr25r5_at_(row,x.source)),memo:pr25r5_text_(pr25r5_at_(row,x.memo))};
    const txt=pr25r5_compact_([o.company,o.cardName,o.evidenceType,o.sourceFile,o.memo].join(' '));
    o.nonCard=/비카드|현금영수증|페이머니/.test(txt)||(/머니/.test(txt)&&!/카드/.test(txt));
    const st=pr25r5_compact_(o.status); o.cancelRow=!!st&&!/취소있음/.test(st)&&/취소|환불/.test(st);
    o.lotteEvidence=String(o.lotteFlag).toUpperCase()==='Y'||/롯데|LOTTE/i.test(o.merchant);
    o.kakaoCard=/카카오/.test(txt)&&!/머니/.test(txt)&&!o.nonCard; o.kakaoMoney=/카카오/.test(txt)&&/머니/.test(txt)&&o.nonCard;
    if(o.date||o.amount||o.approvalNo)out.push(o);
  }
  return out;
}

function pr25r5_loadMaster_(ss) {
  const s=ss.getSheetByName('카드_마스터'); if(!s||s.getLastRow()<2)return [];
  const v=s.getDataRange().getValues(),h=v[0]||[],f=function(n){return pr25r5_find_(h,n,-1);};
  const x={company:f(['카드사']),alias:f(['카드별칭']),name:f(['카드명']),number:f(['카드번호']),end4:f(['카드번호끝4'])};
  return v.slice(1).map(function(row){return{company:pr25r5_text_(pr25r5_at_(row,x.company)),alias:pr25r5_text_(pr25r5_at_(row,x.alias)),cardName:pr25r5_text_(pr25r5_at_(row,x.name)),cardNumber:pr25r5_text_(pr25r5_at_(row,x.number)),cardEnd4:pr25r5_text_(pr25r5_at_(row,x.end4))};}).filter(function(o){return o.company||o.cardName||o.cardNumber;});
}

function pr25r5_canonicalize_(history, master) {
  const groups={},singles=[];
  history.forEach(function(h){
    const issuer=pr25r5_issuer_(h.company),approval=pr25r5_text_(h.approvalNo);
    if(!h.nonCard&&issuer&&approval){const k='CARD|'+issuer+'|'+approval;(groups[k]||(groups[k]=[])).push(h);}else{const c=pr25r5_single_(h,'ROW|'+h.rowNo);if(c)singles.push(c);}
  });
  Object.keys(groups).forEach(function(k){const c=pr25r5_group_(k,groups[k],master);if(c)singles.push(c);});
  return singles;
}
function pr25r5_single_(h,key){if(h.cancelRow&&h.amount<0)return null;const original=Math.abs(h.amount||0),cancel=Math.abs(h.cancelAmount||0),effective=Math.max(original-cancel,0);const x=Object.assign({},h,{canonicalKey:key,originalAmount:original,cancelAmount:cancel,effectiveAmount:effective,fullyCanceled:original>0&&effective===0});return x;}
function pr25r5_group_(key,rows,master){const pos=rows.filter(function(h){return h.amount>0;});if(!pos.length)return null;pos.sort(function(a,b){return pr25r5_rich_(b)-pr25r5_rich_(a);});const rep=Object.assign({},pos[0]);let original=0,cancel=0,cancelDate='';rows.forEach(function(h){if(h.amount>original)original=h.amount;let c=Math.abs(h.cancelAmount||0);if(h.amount<0)c=Math.max(c,Math.abs(h.amount));if(c>cancel){cancel=c;cancelDate=h.cancelDate||h.date||cancelDate;}});cancel=Math.min(cancel,original);rep.canonicalKey=key;rep.originalAmount=original;rep.cancelAmount=cancel;rep.effectiveAmount=Math.max(original-cancel,0);rep.fullyCanceled=original>0&&rep.effectiveAmount===0;rep.cancelDate=cancelDate;rep.kakaoCard=rows.some(function(h){return h.kakaoCard;});rep.kakaoMoney=false;pr25r5_enrich_(rep,master);return rep;}
function pr25r5_rich_(h){return(h.evidenceType==='카드이용내역'?100:0)+(h.cardEnd4?20:0)+(h.cardNumber?10:0)+(h.cardName?5:0)+(h.cancelRow?0:1);}
function pr25r5_enrich_(h,master){const issuer=pr25r5_issuer_(h.company);const matches=master.filter(function(m){if(pr25r5_issuer_(m.company)!==issuer)return false;const he=pr25r5_digits_(h.cardEnd4),me=pr25r5_digits_(m.cardEnd4);if(he&&me&&he===me)return true;const hn=pr25r5_compact_(h.cardNumber),mn=pr25r5_compact_(m.cardNumber);if(hn&&mn&&(hn.indexOf(mn)>=0||mn.indexOf(hn)>=0))return true;const a=pr25r5_compact_(h.cardName),b=pr25r5_compact_(m.cardName);return a.length>=4&&b.length>=4&&(a.indexOf(b)>=0||b.indexOf(a)>=0);});if(matches.length===1){const m=matches[0];h.company=h.company||m.company;h.alias=m.alias||'';h.cardName=m.cardName||h.cardName;h.cardNumber=h.cardNumber||m.cardNumber;h.cardEnd4=h.cardEnd4||m.cardEnd4;}}

function pr25r5_allocate_(orders,evidence,master){const index={};evidence.forEach(function(e){if(!e.date||!e.effectiveAmount)return;const k=e.date+'|'+e.effectiveAmount;(index[k]||(index[k]=[])).push(e);});const used={};const stats={matched:0,nonCard:0,ambiguous:0,noMatch:0};orders.forEach(function(o){o.cardMatch=pr25r5_match_(o,index,used,master);const s=o.cardMatch.status;if(s==='MATCHED')stats.matched++;else if(s==='NON_CARD')stats.nonCard++;else if(s==='AMBIGUOUS')stats.ambiguous++;else stats.noMatch++;});return stats;}
function pr25r5_match_(order,index,used,master){if(!order.orderDate||!order.purchase)return pr25r5_no_('주문일/매입금액 없음');const rule=pr25r5_rule_(order.lottePayment);for(let lag=0;lag<=7;lag++){const date=pr25r5_shift_(order.orderDate,lag);let rows=(index[date+'|'+order.purchase]||[]).filter(function(e){return !e.fullyCanceled&&e.lotteEvidence&&!used[e.canonicalKey];});rows=pr25r5_filter_(rows,rule,order,master);if(!rows.length)continue;const ids={};rows.forEach(function(e){ids[pr25r5_identity_(e,rule,order)]=true;});if(Object.keys(ids).length!==1)return pr25r5_amb_(rows,lag,'서로 다른 구매카드 후보');rows.sort(function(a,b){return String(a.time).localeCompare(String(b.time))||String(a.canonicalKey).localeCompare(String(b.canonicalKey));});const e=rows[0];used[e.canonicalKey]=true;return pr25r5_hit_(order,e,rule,lag,rows.length);}
  const canceled=evidenceCanceledCount_(index,order);return pr25r5_no_('상반기 거래내역 0~+7일 유효증빙 매칭 없음'+(canceled?' / 완전취소 증빙 '+canceled+'건 제외':''));}
function evidenceCanceledCount_(index,order){let n=0;for(let lag=0;lag<=7;lag++){const date=pr25r5_shift_(order.orderDate,lag);Object.keys(index).forEach(function(k){});const arr=index[date+'|'+order.purchase]||[];arr.forEach(function(e){if(e.fullyCanceled)n++;});}return n;}
function pr25r5_filter_(rows,rule,order,master){if(rule.kind==='UNKNOWN')return rows;if(rule.kind==='KAKAO_MONEY')return rows.filter(function(e){return e.nonCard&&e.kakaoMoney;});if(rule.kind==='KAKAO_CARD')return rows.filter(function(e){return !e.nonCard&&e.kakaoCard;});if(rule.kind==='ISSUER_CARD')return rows.filter(function(e){return pr25r5_issuer_(e.company)===rule.issuer;});return rows;}
function pr25r5_identity_(e,rule,order){if(rule.kind==='KAKAO_MONEY')return'NON_CARD|KAKAO_MONEY';if(rule.kind==='ISSUER_CARD'){if(rule.issuer==='KB국민카드')return'KB|4091';if(rule.issuer==='우리카드')return'WOORI|7680';if(rule.issuer==='롯데카드')return order.orderDate<='2026-05-28'?'LOTTE|TRIP':'LOTTE|036';return rule.issuer+'|'+(pr25r5_digits_(e.cardEnd4)||pr25r5_compact_(e.cardName)||e.canonicalKey);}return(pr25r5_issuer_(e.company)||'UNKNOWN')+'|'+(pr25r5_digits_(e.cardEnd4)||pr25r5_compact_(e.cardName)||e.canonicalKey);}
function pr25r5_hit_(order,e,rule,lag,count){let company=e.nonCard?'비카드':(e.company||pr25r5_issuer_(e.company)),alias=e.alias||'',name=e.cardName||'',end4=e.cardEnd4||'',number=e.cardNumber||'';if(rule.kind==='KAKAO_MONEY'){company='비카드';alias='신한은행 계좌결제';name='카카오페이 페이머니';end4='';number='';}else if(rule.kind==='ISSUER_CARD'&&rule.issuer==='KB국민카드'){company='KB국민카드';name='HERITAGE Smart(할인형)';end4='4091';}else if(rule.kind==='ISSUER_CARD'&&rule.issuer==='우리카드'){company='우리카드';name='카드의정석 EVERY POINT';end4='7680';}else if(rule.kind==='ISSUER_CARD'&&rule.issuer==='롯데카드'){company='롯데카드';if(order.orderDate<='2026-05-28'){name='Trip to 로카';}else{name='LOCA LIKIT 1.2';end4='036';}}
  const prefix=pr25r5_reason_(order,rule);const reason=prefix+' / 거래내역_마켓주문일'+(lag?'+'+lag+'일':'당일')+'_실결제금액_1:1할당'+(count>1?'_동일구매수단확정('+count+'건)':'')+(e.cancelAmount?' / NET_AFTER_CANCEL':' / APPROVAL');return{status:e.nonCard?'NON_CARD':'MATCHED',reason:reason,candidateCount:count,company:company,alias:alias,cardName:name,cardNumber:number,cardEnd4:end4,approvalDate:e.date,approvalTime:e.time,approvalNo:e.approvalNo,approvalAmount:e.originalAmount,merchant:e.merchant,evidenceType:e.evidenceType,sourceFile:e.sourceFile,cancelMemo:e.cancelAmount?'취소일 '+(e.cancelDate||'-')+' / 취소금액 -'+e.cancelAmount+' / 실결제 '+e.effectiveAmount:'',candidateSummary:pr25r5_label_(e),allocationLagDays:lag};}
function pr25r5_rule_(v){const raw=pr25r5_text_(v),s=pr25r5_compact_(raw);if(!s)return{kind:'UNKNOWN',raw:raw,issuer:''};if(/카카오/.test(s)&&(/머니|계좌|현금/.test(s)))return{kind:'KAKAO_MONEY',raw:raw,issuer:'비카드'};if(/카카오/.test(s))return{kind:'KAKAO_CARD',raw:raw,issuer:''};const issuer=pr25r5_issuer_(raw);return issuer?{kind:'ISSUER_CARD',raw:raw,issuer:issuer}:{kind:'UNKNOWN',raw:raw,issuer:''};}
function pr25r5_reason_(o,r){if(r.kind==='KAKAO_MONEY')return'트래킹번호_카카오페이페이머니_신한은행계좌_현금결제_1차필터';if(r.kind==='KAKAO_CARD')return'트래킹번호_카카오페이카드_원카드승인증빙_1차필터';if(r.kind==='ISSUER_CARD'){let s='트래킹번호_'+r.issuer+'_1차필터';if(r.issuer==='KB국민카드')s+='_HERITAGE단일카드';if(r.issuer==='우리카드')s+='_EVERY_POINT단일카드';if(r.issuer==='롯데카드')s+=o.orderDate<='2026-05-28'?'_Trip_to_로카(~2026-05-28)':'_LOCA_LIKIT(2026-05-29~)';return s;}return'결제수단미확인';}
function pr25r5_amb_(rows,lag,why){return{status:'AMBIGUOUS',reason:'0~+7일 '+why+' / lag='+lag,candidateCount:rows.length,company:'',alias:'',cardName:'',cardNumber:'',cardEnd4:'',approvalDate:'',approvalTime:'',approvalNo:'',approvalAmount:0,merchant:'',evidenceType:'',sourceFile:'',cancelMemo:'',candidateSummary:rows.map(pr25r5_label_).join(' || '),allocationLagDays:lag};}
function pr25r5_no_(reason){return{status:'NO_MATCH',reason:reason,candidateCount:0,company:'',alias:'',cardName:'',cardNumber:'',cardEnd4:'',approvalDate:'',approvalTime:'',approvalNo:'',approvalAmount:0,merchant:'',evidenceType:'',sourceFile:'',cancelMemo:'',candidateSummary:'',allocationLagDays:''};}
function pr25r5_label_(e){return[e.company,e.cardName,e.cardEnd4,e.date,e.time,e.effectiveAmount,e.approvalNo].map(pr25r5_text_).join(' / ');}

function pr25r5_writeDiagnostic_(ss,orders){const headers=['신고연도','반기','신고월','사업자등록번호','쿠팡계정ID','주문번호','주문일','롯데결제수단','매입금액','순수매출액','카드매칭상태','카드매칭근거','후보수','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','승인일','승인시각','승인번호','승인금액','가맹점명','증빙유형','원본파일','취소정보','후보요약'];const rows=orders.map(function(o){const m=o.cardMatch;return[o.year,o.half,o.month,o.business,o.account,o.orderNo,o.orderDate,o.lottePayment,o.purchase,o.sales,m.status,m.reason,m.candidateCount,m.company,m.alias,m.cardName,m.cardNumber,m.cardEnd4,m.approvalDate,m.approvalTime,m.approvalNo,m.approvalAmount,m.merchant,m.evidenceType,m.sourceFile,m.cancelMemo,m.candidateSummary];});pr25r5_write_(ss,'PR25_카드매칭검증',headers,rows);}
function pr25r5_writeSummary_(ss,orders){const map={};orders.forEach(function(o){const m=o.cardMatch||pr25r5_no_('');const key=[o.year,o.half,o.business,o.account,m.company,m.alias,m.cardName,m.cardNumber,m.cardEnd4,m.status,m.reason].join('|');if(!map[key])map[key]={year:o.year,half:o.half,business:o.business,account:o.account,company:m.company,alias:m.alias,name:m.cardName,number:m.cardNumber,end4:m.cardEnd4,status:m.status,reasons:{},orders:{},sales:0,salesSupply:0,salesVat:0,settlement:0,fee:0,purchase:0,purchaseSupply:0,purchaseVat:0,payable:0,profit:0,vatProfit:0};const x=map[key];x.reasons[m.reason]=true;x.orders[o.orderNo||o.key]=true;x.sales+=o.sales;x.salesSupply+=o.salesSupply;x.salesVat+=o.salesVat;x.settlement+=o.settlement;x.fee+=o.fee;x.purchase+=o.purchase;x.purchaseSupply+=o.purchaseSupply;x.purchaseVat+=o.purchaseVat;x.payable+=o.payable;x.profit+=o.profit;x.vatProfit+=o.vatProfit;});const headers=['신고연도','반기','사업자등록번호','연결 쿠팡계정ID','구매카드사','구매카드별칭','구매카드명','카드번호','카드번호끝4','카드매칭상태','카드매칭근거','주문건수','순수매출액','매출공급가액','매출부가세','정산기준금액','마켓수수료','매입금액','매입공급가액','매입부가세','납부예상부가세','예상이익','부가세반영예상이익','비고'];const rows=Object.keys(map).map(function(k){const x=map[k];return[x.year,x.half,x.business,x.account,x.company,x.alias,x.name,x.number,x.end4,x.status,Object.keys(x.reasons).join(' / '),Object.keys(x.orders).length,x.sales,x.salesSupply,x.salesVat,x.settlement,x.fee,x.purchase,x.purchaseSupply,x.purchaseVat,x.payable,x.profit,x.vatProfit,''];}).sort(function(a,b){return String(a[0]).localeCompare(String(b[0]))||String(a[1]).localeCompare(String(b[1]))||String(a[2]).localeCompare(String(b[2]))||String(a[4]).localeCompare(String(b[4]));});pr25r5_write_(ss,'PR25_사업자별반기요약',headers,rows);}
function pr25r5_writeMonthlyPayment_(ss,orders){const map={};orders.forEach(function(o){const month=o.month||o.orderDate.slice(0,7),rule=pr25r5_rule_(o.lottePayment);let label=o.lottePayment||'공란';if(rule.kind==='KAKAO_MONEY')label='비카드-카카오페이 페이머니';else if(rule.kind==='KAKAO_CARD')label='카카오페이 카드결제';else if(rule.kind==='ISSUER_CARD'&&rule.issuer==='롯데카드')label=o.orderDate<='2026-05-28'?'롯데카드-Trip to 로카':'롯데카드-LOCA LIKIT 1.2';else if(rule.kind==='ISSUER_CARD'&&rule.issuer==='KB국민카드')label='KB국민카드-HERITAGE Smart(할인형)';else if(rule.kind==='ISSUER_CARD'&&rule.issuer==='우리카드')label='우리카드-카드의정석 EVERY POINT';else if(rule.kind==='ISSUER_CARD')label=rule.issuer;const k=[month,o.business,o.account,label].join('|');if(!map[k])map[k]=[month,o.business,o.account,label,0,0];map[k][4]++;map[k][5]+=o.purchase;});const rows=Object.keys(map).map(function(k){return map[k];}).sort(function(a,b){return String(a[0]).localeCompare(String(b[0]))||String(a[1]).localeCompare(String(b[1]))||String(a[3]).localeCompare(String(b[3]));});pr25r5_write_(ss,'PR25_결제수단월별검증',['신고월','사업자등록번호','쿠팡계정ID','트래킹 결제수단 분류','주문건수','매입금액'],rows);}
function pr25r5_write_(ss,name,headers,rows){const s=ss.getSheetByName(name)||ss.insertSheet(name);s.clearContents();s.getRange(1,1,1,headers.length).setValues([headers]).setBackground('#d9eaf7').setFontWeight('bold');if(rows.length)s.getRange(2,1,rows.length,headers.length).setValues(rows);s.setFrozenRows(1);}

function pr25r5_find_(headers,names,fallback){for(let n=0;n<names.length;n++){const t=String(names[n]).replace(/\s/g,'').toLowerCase();for(let i=0;i<headers.length;i++)if(String(headers[i]||'').replace(/\s/g,'').toLowerCase()===t)return i;}return fallback;}
function pr25r5_at_(row,i){return i>=0?row[i]:'';}
function pr25r5_text_(v){return String(v==null?'':v).trim();}
function pr25r5_num_(v){if(typeof v==='number')return Math.round(v);const n=Number(String(v==null?'':v).replace(/[원,\s]/g,''));return isNaN(n)?0:Math.round(n);}
function pr25r5_compact_(v){return pr25r5_text_(v).toLowerCase().replace(/[\s._()\[\]{}\-\/]/g,'');}
function pr25r5_digits_(v){return pr25r5_text_(v).replace(/\D/g,'');}
function pr25r5_date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','MM/dd');const s=pr25r5_text_(v),m=s.match(/(?:\d{4}[.\/-])?(\d{1,2})[.\/-](\d{1,2})/);return m?(('0'+m[1]).slice(-2)+'/'+('0'+m[2]).slice(-2)):s;}
function pr25r5_fullDate_(v,year){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');const s=pr25r5_text_(v),m=s.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);if(m)return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);const m2=s.match(/^(\d{1,2})[.\/-](\d{1,2})/);return m2&&year?year+'-'+('0'+m2[1]).slice(-2)+'-'+('0'+m2[2]).slice(-2):s;}
function pr25r5_time_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','HH:mm:ss');return pr25r5_text_(v);}
function pr25r5_rowKey_(o){return[o.date,o.account,o.orderNo,o.customer,o.brand,o.productNo,o.productName,o.quantity,o.sales,o.purchase].map(pr25r5_text_).join('|');}
function pr25r5_issuer_(v){const s=pr25r5_compact_(v);if(!s)return'';if(/kb|국민/.test(s))return'KB국민카드';if(/롯데/.test(s))return'롯데카드';if(/우리/.test(s))return'우리카드';if(/신한/.test(s))return'신한카드';if(/농협|nh/.test(s))return'NH농협카드';if(/삼성/.test(s))return'삼성카드';if(/하나/.test(s))return'하나카드';if(/현대/.test(s))return'현대카드';return'';}
function pr25r5_shift_(date,days){const m=String(date).match(/^(\d{4})-(\d{2})-(\d{2})$/);if(!m)return'';const d=new Date(Date.UTC(+m[1],+m[2]-1,+m[3]+days));return d.getUTCFullYear()+'-'+('0'+(d.getUTCMonth()+1)).slice(-2)+'-'+('0'+d.getUTCDate()).slice(-2);}
