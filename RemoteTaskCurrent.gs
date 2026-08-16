var LOTTEON_REMOTE_TASK={id:'ISSUE79-V45-ISSUE76-EXACT-SOURCE-RECON',title:'Issue79 Issue76 exact key + current source READ-ONLY reconciliation',enabled:true,statusSheet:'ISSUE79_현재상태대조상태'};
var I79V45='v4.5-ISSUE79-I76-EXACTKEY-SOURCE-RECON-READONLY';
var I79V45_CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
var I79V45_OLD22=['30100189414967','2100191293730','23100192054997','29100191903703','5100191540477','26100192791914','17100196061992','18100195889436','12100196902987','12100197310918','11100197821559','5100198099688','19100198942446','5101102615736','6101100531065','12101138407306','12101154659935','17100199506890','19100199527230','29100199660972','16101220120574','9101251100578'];
function runLotteonRemoteTaskStartRemote_(){return i79v45run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v45run_();}
function i79v45run_(){
  var ss=SpreadsheetApp.getActive(),before={},started=new Date().toISOString();
  try{
    i79v45status_(ss,[['version',I79V45],['상태','RUNNING'],['단계','READ'],['메시지','Issue76 exact key + 현재 원본 재구성 READ-ONLY 대조 중'],['실행시작',started]]);
    I79V45_CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락: '+n);before[n]=i79v45sig_(sh);});
    var cur=i79v45card_(ss.getSheetByName('부가세_카드매칭검증'),'CURRENT',true);
    var vat=i79v45vat_(ss.getSheetByName('부가세_신고자료'));
    var histSh=ss.getSheetByName('ISSUE74_카드회수PREVIEW');if(!histSh)throw new Error('ISSUE74_카드회수PREVIEW 누락');
    var hist=i79v45card_(histSh,'ISSUE74',false);
    var v2=i79v45v2_(ss.getSheetByName('ISSUE79_NOMATCH22_최신재검수'));
    var cross=i79v45crossComposite_(vat,cur);
    var orderCross=i79v45crossOrder_(vat,cur);
    var drift=i79v45history_(cur,hist);
    var source=i79v45source_(ss.getSheetByName('매출데이터_붙여넣기'));
    var sourceVat=i79v45sourceVat_(source,vat);
    var old=i79v45old22_(cur,hist,v2);
    i79v45writeDrift_(ss,drift.rows);
    i79v45writeSource_(ss,i79v45sourceRows_(drift,source));
    var changed=[];I79V45_CORE.forEach(function(n){if(i79v45sig_(ss.getSheetByName(n))!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반: '+changed.join(','));
    var cs=cur.stats,hs=hist.stats;
    var sourceExact=(source.activeDetails===vat.detailRows&&source.activeOrders===vat.orders&&source.activePurchase===vat.purchase&&sourceVat.sourceOnly===0&&sourceVat.vatOnly===0&&sourceVat.purchaseMismatch===0);
    i79v45status_(ss,[
      ['version',I79V45],['상태','PASS'],['단계','DONE'],['메시지','Issue76 exact key + 현재 원본 + Issue74 역사 READ-ONLY 대조 완료'],
      ['현재_2026H1주문',cur.orders],['현재_MATCHED',cs.MATCHED||0],['현재_NON_CARD',cs.NON_CARD||0],['현재_NO_MATCH',cs.NO_MATCH||0],['현재_AMBIGUOUS',cs.AMBIGUOUS||0],['현재_매입합계',cur.purchase],['현재_composite중복',cur.compDup],['현재_주문번호중복',cur.orderDup],
      ['VAT_2026H1상세행',vat.detailRows],['VAT_2026H1주문',vat.orders],['VAT_매입합계',vat.purchase],['VAT_composite주문',Object.keys(vat.compMap).length],['VAT_주문번호주문',Object.keys(vat.orderMap).length],
      ['Issue76_exact_overlap',cross.overlap],['Issue76_exact_VAT_ONLY',cross.vatOnly],['Issue76_exact_CARD_ONLY',cross.cardOnly],['Issue76_exact_매입불일치',cross.purchaseMismatch],['Issue76_exact_사업자불일치',cross.businessMismatch],
      ['주문번호_only_overlap',orderCross.overlap],['주문번호_only_VAT_ONLY',orderCross.vatOnly],['주문번호_only_CARD_ONLY',orderCross.cardOnly],['주문번호_only_매입불일치',orderCross.purchaseMismatch],
      ['Issue74_주문',hist.orders],['Issue74_MATCHED',hs.MATCHED||0],['Issue74_NON_CARD',hs.NON_CARD||0],['Issue74_NO_MATCH',hs.NO_MATCH||0],['Issue74_AMBIGUOUS',hs.AMBIGUOUS||0],['Issue74_매입합계',hist.purchase],
      ['현재Issue74_overlap',drift.overlap],['현재_ONLY',drift.currentOnly],['Issue74_ONLY',drift.historyOnly],['공통_매입변경',drift.purchaseMismatch],['공통_상태변경',drift.statusChanged],['현재_ONLY_매입합',drift.currentOnlyPurchase],['Issue74_ONLY_매입합',drift.historyOnlyPurchase],['공통_매입변경순차이',drift.sharedDelta],['차이분해합계',drift.currentOnlyPurchase-drift.historyOnlyPurchase+drift.sharedDelta],['현재-Issue74_매입합계차이',cur.purchase-hist.purchase],['상태전이',drift.transitionText],
      ['현재원본_2026H1활성상세행',source.activeDetails],['현재원본_2026H1활성주문',source.activeOrders],['현재원본_2026H1활성매입합계',source.activePurchase],['원본VAT_overlap',sourceVat.overlap],['원본_ONLY',sourceVat.sourceOnly],['VAT_ONLY_vs원본',sourceVat.vatOnly],['원본VAT_매입불일치',sourceVat.purchaseMismatch],['현재VAT_원본재구성exact',sourceExact?'YES':'NO'],
      ['과거22_현재MATCHED',old.stats.MATCHED||0],['과거22_현재NON_CARD',old.stats.NON_CARD||0],['과거22_현재NO_MATCH',old.stats.NO_MATCH||0],['과거22_현재AMBIGUOUS',old.stats.AMBIGUOUS||0],['과거22_현재누락',old.stats.MISSING||0],
      ['v2_SAFE건수',old.safe.total],['v2_SAFE_현재MATCHED',old.safe.MATCHED||0],['v2_SAFE_현재NON_CARD',old.safe.NON_CARD||0],['v2_SAFE_현재NO_MATCH',old.safe.NO_MATCH||0],['v2_SAFE_현재누락',old.safe.MISSING||0],
      ['v2_잔여16건수',old.remain.total],['v2_잔여16_현재MATCHED',old.remain.MATCHED||0],['v2_잔여16_현재NON_CARD',old.remain.NON_CARD||0],['v2_잔여16_현재NO_MATCH',old.remain.NO_MATCH||0],['v2_잔여16_현재누락',old.remain.MISSING||0],
      ['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
    ]);
    return{ok:true,done:true,version:I79V45,exactOverlap:cross.overlap,sourceExact:sourceExact,netDelta:cur.purchase-hist.purchase,old22Missing:old.stats.MISSING||0};
  }catch(e){var msg=String(e&&e.message?e.message:e);try{i79v45status_(ss,[['version',I79V45],['상태','ERROR'],['단계','FAILED'],['메시지','v4.5 READ-ONLY 대조 실패'],['실행시작',started],['오류',msg],['완료시각',new Date().toISOString()]]);}catch(_e){}throw e;}
}
function i79v45card_(sh,label,scope){
  var v=sh.getDataRange().getValues(),hr=i79v45header_(v,['주문번호','카드매칭상태']),h=v[hr];
  var x={year:i79v45ixExact_(h,'신고연도'),half:i79v45ixExact_(h,'반기'),business:i79v45ixExact_(h,'사업자등록번호'),account:i79v45ixExact_(h,'쿠팡계정ID'),order:i79v45ixExact_(h,'주문번호'),purchase:i79v45ixPriority_(h,['주문매입금액','매입금액']),status:i79v45ixExact_(h,'카드매칭상태'),company:i79v45ixExact_(h,'구매카드사'),name:i79v45ixExact_(h,'구매카드명'),end4:i79v45ixExact_(h,'카드번호끝4'),adate:i79v45ixExact_(h,'승인일'),approval:i79v45ixExact_(h,'승인번호'),aamount:i79v45ixExact_(h,'승인금액'),reason:i79v45ixExact_(h,'카드매칭근거')};
  ['business','account','order','purchase','status'].forEach(function(k){if(x[k]<0)throw new Error(label+' exact header 누락 '+k);});
  var compMap={},orderMap={},stats={},purchase=0,orders=0,compDup=0,orderDup=0;
  for(var r=hr+1;r<v.length;r++){
    var z=v[r],no=i79v45t_(z[x.order]);if(!no)continue;
    if(scope){if(x.year<0||x.half<0)throw new Error(label+' 신고연도/반기 header 누락');if(i79v45t_(z[x.year])!=='2026'||i79v45t_(z[x.half])!=='상반기')continue;}
    var yr=x.year>=0?i79v45t_(z[x.year]):'2026',hf=x.half>=0?i79v45t_(z[x.half]):'상반기',biz=i79v45t_(z[x.business]),acc=i79v45t_(z[x.account]),st=i79v45t_(z[x.status]).toUpperCase(),amt=i79v45n_(z[x.purchase]);
    var o={year:yr,half:hf,business:biz,account:acc,orderNo:no,purchase:amt,status:st,company:x.company>=0?i79v45t_(z[x.company]):'',cardName:x.name>=0?i79v45t_(z[x.name]):'',end4:x.end4>=0?i79v45t_(z[x.end4]):'',approvalDate:x.adate>=0?i79v45t_(z[x.adate]):'',approvalNo:x.approval>=0?i79v45t_(z[x.approval]):'',approvalAmount:x.aamount>=0?i79v45n_(z[x.aamount]):0,reason:x.reason>=0?i79v45t_(z[x.reason]):''};
    o.compKey=[yr,hf,biz,acc,no].join('|');orders++;purchase+=amt;stats[st]=(stats[st]||0)+1;if(compMap[o.compKey])compDup++;else compMap[o.compKey]=o;if(orderMap[no])orderDup++;else orderMap[no]=o;
  }
  return{orders:orders,purchase:Math.round(purchase),stats:stats,compMap:compMap,orderMap:orderMap,compDup:compDup,orderDup:orderDup};
}
function i79v45vat_(sh){
  var v=sh.getDataRange().getValues(),h=v[0]||[],x={year:i79v45ixExact_(h,'신고연도'),half:i79v45ixExact_(h,'반기'),business:i79v45ixExact_(h,'사업자등록번호'),account:i79v45ixExact_(h,'쿠팡계정ID'),order:i79v45ixExact_(h,'주문번호'),purchase:i79v45ixExact_(h,'매입금액')};
  Object.keys(x).forEach(function(k){if(x[k]<0)throw new Error('VAT exact header 누락 '+k);});
  var compMap={},orderMap={},detail=0,purchase=0;
  for(var r=1;r<v.length;r++){
    var z=v[r];if(i79v45t_(z[x.year])!=='2026'||i79v45t_(z[x.half])!=='상반기')continue;var no=i79v45t_(z[x.order]);if(!no)continue;detail++;var biz=i79v45t_(z[x.business]),acc=i79v45t_(z[x.account]),amt=i79v45n_(z[x.purchase]),ck=['2026','상반기',biz,acc,no].join('|');purchase+=amt;
    if(!compMap[ck])compMap[ck]={compKey:ck,business:biz,account:acc,orderNo:no,purchase:0,detailRows:0};compMap[ck].purchase+=amt;compMap[ck].detailRows++;
    if(!orderMap[no])orderMap[no]={orderNo:no,purchase:0,detailRows:0,businesses:{},accounts:{}};orderMap[no].purchase+=amt;orderMap[no].detailRows++;orderMap[no].businesses[biz]=1;orderMap[no].accounts[acc]=1;
  }
  Object.keys(compMap).forEach(function(k){compMap[k].purchase=Math.round(compMap[k].purchase);});Object.keys(orderMap).forEach(function(k){orderMap[k].purchase=Math.round(orderMap[k].purchase);});
  return{detailRows:detail,orders:Object.keys(compMap).length,purchase:Math.round(purchase),compMap:compMap,orderMap:orderMap};
}
function i79v45crossComposite_(vat,card){var x={overlap:0,vatOnly:0,cardOnly:0,purchaseMismatch:0,businessMismatch:0},seen={};Object.keys(vat.compMap).forEach(function(k){var a=vat.compMap[k],b=card.compMap[k];if(!b){x.vatOnly++;return;}x.overlap++;seen[k]=1;if(a.purchase!==b.purchase)x.purchaseMismatch++;if(a.business!==b.business)x.businessMismatch++;});Object.keys(card.compMap).forEach(function(k){if(!seen[k])x.cardOnly++;});return x;}
function i79v45crossOrder_(vat,card){var x={overlap:0,vatOnly:0,cardOnly:0,purchaseMismatch:0},seen={};Object.keys(vat.orderMap).forEach(function(k){var a=vat.orderMap[k],b=card.orderMap[k];if(!b){x.vatOnly++;return;}x.overlap++;seen[k]=1;if(a.purchase!==b.purchase)x.purchaseMismatch++;});Object.keys(card.orderMap).forEach(function(k){if(!seen[k])x.cardOnly++;});return x;}
function i79v45history_(cur,hist){
  var rows=[],overlap=0,co=0,ho=0,pm=0,sc=0,cop=0,hop=0,sharedDelta=0,seen={},tr={};
  Object.keys(cur.orderMap).forEach(function(no){var c=cur.orderMap[no],h=hist.orderMap[no];if(!h){co++;cop+=c.purchase;rows.push(['CURRENT_ONLY',no,'',c.status,'',c.purchase,c.purchase,c.business,c.account,c.reason]);return;}overlap++;seen[no]=1;var delta=c.purchase-h.purchase;sharedDelta+=delta;if(delta!==0){pm++;rows.push(['PURCHASE_MISMATCH',no,h.status,c.status,h.purchase,c.purchase,delta,c.business,c.account,c.reason]);}if(c.status!==h.status){sc++;var k=(h.status||'')+'→'+(c.status||'');tr[k]=(tr[k]||0)+1;rows.push(['STATUS_CHANGE',no,h.status,c.status,h.purchase,c.purchase,delta,c.business,c.account,c.reason]);}});
  Object.keys(hist.orderMap).forEach(function(no){if(!seen[no]&&!cur.orderMap[no]){var h=hist.orderMap[no];ho++;hop+=h.purchase;rows.push(['ISSUE74_ONLY',no,h.status,'',h.purchase,'',-h.purchase,h.business,h.account,h.reason]);}});
  return{rows:rows,overlap:overlap,currentOnly:co,historyOnly:ho,purchaseMismatch:pm,statusChanged:sc,currentOnlyPurchase:Math.round(cop),historyOnlyPurchase:Math.round(hop),sharedDelta:Math.round(sharedDelta),transitionText:Object.keys(tr).sort().map(function(k){return k+':'+tr[k];}).join(' / '),currentOnlySet:i79v45set_(rows,'CURRENT_ONLY'),historyOnlySet:i79v45set_(rows,'ISSUE74_ONLY')};
}
function i79v45set_(rows,type){var s={};rows.forEach(function(r){if(r[0]===type)s[r[1]]=1;});return s;}
function i79v45source_(sh){
  var v=sh.getDataRange().getValues(),h=v[0]||[];
  var x={date:i79v45ixPriority_(h,['마켓주문일자','주문일자','결제일자','주문일시']),order:i79v45ixPriority_(h,['마켓주문번호','주문번호','주문ID','주문ID(마켓)']),sales:i79v45ixPriority_(h,['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액']),status:i79v45ixPriority_(h,['마켓주문상태','주문상태','상태','클레임상태','처리상태'])};
  if(x.date<0)x.date=0;if(x.order<0)x.order=2;if(x.sales<0)x.sales=6;var pidx=28;if(v[0].length<=pidx)throw new Error('매출원본 AC 매입금액 열 누락');
  var activeMap={},allMap={},ad=0,ap=0;
  for(var r=1;r<v.length;r++){
    var z=v[r],no=i79v45t_(z[x.order]);if(!no)continue;var st=x.status>=0?i79v45t_(z[x.status]):'',dt=i79v45date_(z[x.date]),pu=i79v45n_(z[pidx]),sales=i79v45n_(z[x.sales]);
    if(!allMap[no])allMap[no]={rows:0,purchase:0,statuses:{},dates:{}};allMap[no].rows++;allMap[no].purchase+=pu;if(st)allMap[no].statuses[st]=1;if(dt)allMap[no].dates[dt]=1;
    var h1=dt&&dt>='2026-01-01'&&dt<='2026-06-30';if(!h1||/취소|반품|환불/.test(st)||!sales)continue;ad++;ap+=pu;if(!activeMap[no])activeMap[no]={orderNo:no,purchase:0,detailRows:0,statuses:{}};activeMap[no].purchase+=pu;activeMap[no].detailRows++;if(st)activeMap[no].statuses[st]=1;
  }
  Object.keys(activeMap).forEach(function(k){activeMap[k].purchase=Math.round(activeMap[k].purchase);});return{activeDetails:ad,activeOrders:Object.keys(activeMap).length,activePurchase:Math.round(ap),activeMap:activeMap,allMap:allMap,orderHeader:x.order,statusHeader:x.status,dateHeader:x.date};
}
function i79v45sourceVat_(source,vat){var x={overlap:0,sourceOnly:0,vatOnly:0,purchaseMismatch:0},seen={};Object.keys(source.activeMap).forEach(function(no){var a=source.activeMap[no],b=vat.orderMap[no];if(!b){x.sourceOnly++;return;}x.overlap++;seen[no]=1;if(a.purchase!==b.purchase)x.purchaseMismatch++;});Object.keys(vat.orderMap).forEach(function(no){if(!seen[no]&&!source.activeMap[no])x.vatOnly++;});return x;}
function i79v45sourceRows_(drift,source){var out=[];function add(type,set){Object.keys(set).sort().forEach(function(no){var a=source.allMap[no],b=source.activeMap[no];out.push([type,no,a?'Y':'N',b?'Y':'N',a?a.rows:0,b?b.detailRows:0,a?i79v45keys_(a.statuses):'',a?i79v45keys_(a.dates):'',a?Math.round(a.purchase):'',b?b.purchase:'']);});}add('CURRENT_ONLY',drift.currentOnlySet);add('ISSUE74_ONLY',drift.historyOnlySet);return out;}
function i79v45old22_(cur,hist,v2){var stats={},safe={total:0},remain={total:0};I79V45_OLD22.forEach(function(no){var c=cur.orderMap[no]||null,st=c?c.status:'MISSING',bucket=(st==='MATCHED'||st==='NON_CARD'||st==='NO_MATCH'||st==='AMBIGUOUS')?st:'MISSING',isSafe=!!v2.safe[no],b=isSafe?safe:remain;stats[bucket]=(stats[bucket]||0)+1;b.total++;b[bucket]=(b[bucket]||0)+1;});return{stats:stats,safe:safe,remain:remain};}
function i79v45v2_(sh){if(!sh)return{map:{},safe:{}};var v=sh.getDataRange().getValues(),h=v[0]||[],xo=i79v45ixExact_(h,'주문번호'),xv=i79v45ixExact_(h,'PREVIEW판정');if(xo<0||xv<0)return{map:{},safe:{}};var map={},safe={};for(var r=1;r<v.length;r++){var no=i79v45t_(v[r][xo]);if(!no)continue;var verdict=i79v45t_(v[r][xv]);map[no]=verdict;if(/^SAFE_/.test(verdict))safe[no]=1;}return{map:map,safe:safe};}
function i79v45writeDrift_(ss,rows){var sh=ss.getSheetByName('ISSUE79_v45_현재vsIssue74정밀차이')||ss.insertSheet('ISSUE79_v45_현재vsIssue74정밀차이');i79v45write_(sh,['구분','주문번호','Issue74상태','현재상태','Issue74매입','현재매입','차이','현재사업자','현재계정','현재매칭근거'],rows,2);}
function i79v45writeSource_(ss,rows){var sh=ss.getSheetByName('ISSUE79_v45_원본추적')||ss.insertSheet('ISSUE79_v45_원본추적');i79v45write_(sh,['구분','주문번호','현재원본존재','현재원본활성H1존재','원본전체행수','원본활성행수','현재원본상태','현재원본날짜','원본전체매입합','원본활성매입합'],rows,2);}
function i79v45write_(sh,h,rows,textCol){if(sh.getFilter())sh.getFilter().remove();sh.clear();sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length){sh.getRange(2,textCol,rows.length,1).setNumberFormat('@');sh.getRange(2,1,rows.length,h.length).setValues(rows);}sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}sh.autoResizeColumns(1,h.length);}
function i79v45status_(ss,pairs){var sh=ss.getSheetByName('ISSUE79_현재상태대조상태')||ss.insertSheet('ISSUE79_현재상태대조상태');sh.clearContents();var a=[['항목','값']].concat(pairs||[]);sh.getRange(1,1,a.length,2).setValues(a);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}
function i79v45sig_(sh){var rg=sh.getDataRange(),s=sh.getName()+'|'+rg.getNumRows()+'|'+rg.getNumColumns()+'|'+JSON.stringify(rg.getDisplayValues());return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,s,Utilities.Charset.UTF_8));}
function i79v45header_(v,need){for(var r=0;r<Math.min(50,v.length);r++){var ok=true;for(var j=0;j<need.length;j++)if(i79v45ixExact_(v[r],need[j])<0){ok=false;break;}if(ok)return r;}throw new Error('header 탐지 실패: '+need.join(','));}
function i79v45ixExact_(h,name){var n=i79v45c_(name);for(var i=0;i<h.length;i++)if(i79v45c_(h[i])===n)return i;return-1;}
function i79v45ixPriority_(h,names){for(var j=0;j<names.length;j++){var x=i79v45ixExact_(h,names[j]);if(x>=0)return x;}return-1;}
function i79v45t_(v){return String(v==null?'':v).trim();}
function i79v45c_(v){return i79v45t_(v).toLowerCase().replace(/\s+/g,'');}
function i79v45n_(v){if(typeof v==='number')return isFinite(v)?Math.round(v):0;var s=i79v45t_(v).replace(/,/g,'').replace(/[^0-9.\-]/g,'');var n=Number(s);return isFinite(n)?Math.round(n):0;}
function i79v45date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i79v45t_(v),m=s.match(/(20\d{2})\D?(\d{1,2})\D?(\d{1,2})/);if(!m)return'';return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);}
function i79v45keys_(o){return Object.keys(o||{}).sort().join('|');}
