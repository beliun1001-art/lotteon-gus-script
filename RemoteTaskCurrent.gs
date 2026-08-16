var LOTTEON_REMOTE_TASK={id:'ISSUE79-V47-STABLE-FINGERPRINT-ID-TRACE',title:'Issue79 안정속성 기반 주문ID 역추적 READ-ONLY',enabled:true,statusSheet:'ISSUE79_식별자역추적상태'};
var I79V47='v4.7-ISSUE79-STABLE-FINGERPRINT-ID-TRACE-READONLY';
var I79V47_CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
var I79V47_OLD22=['30100189414967','2100191293730','23100192054997','29100191903703','5100191540477','26100192791914','17100196061992','18100195889436','12100196902987','12100197310918','11100197821559','5100198099688','19100198942446','5101102615736','6101100531065','12101138407306','12101154659935','17100199506890','19100199527230','29100199660972','16101220120574','9101251100578'];
function runLotteonRemoteTaskStartRemote_(){return i79v47run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v47run_();}
function i79v47run_(){
 var ss=SpreadsheetApp.getActive(),before={},started=new Date().toISOString();
 try{
  i79v47status_(ss,[['version',I79V47],['상태','RUNNING'],['단계','READ'],['메시지','주문번호를 배제한 안정속성 fingerprint로 현재 VAT/CARD를 원본 주문에 역추적 중'],['실행시작',started]]);
  I79V47_CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락: '+n);before[n]=i79v47sig_(sh);});
  var card=i79v47card_(ss.getSheetByName('부가세_카드매칭검증'),'CURRENT',true);
  var hist=i79v47card_(ss.getSheetByName('ISSUE74_카드회수PREVIEW'),'ISSUE74',false);
  var vat=i79v47vat_(ss.getSheetByName('부가세_신고자료'));
  var src=i79v47source_(ss.getSheetByName('매출데이터_붙여넣기'));
  var v2=i79v47v2_(ss.getSheetByName('ISSUE79_NOMATCH22_최신재검수'));
  var vm=i79v47mapVatToSource_(vat,src);
  var cm=i79v47mapCardToSource_(card,vat,src,vm);
  var th=i79v47translatedHistory_(cm,hist,src,v2);
  var old=i79v47old22_(cm,v2);
  var extra=i79v47sourceExtra_(src,vm);
  i79v47writeTranslation_(ss,th.rows);
  i79v47writePurchase_(ss,th.purchaseRows);
  var changed=[];I79V47_CORE.forEach(function(n){if(i79v47sig_(ss.getSheetByName(n))!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반: '+changed.join(','));
  var cs=card.stats,hs=hist.stats;
  var snapshotExact=(vm.unique===vat.orders&&vm.ambiguous===0&&vm.unmapped===0&&vm.collision===0&&vm.purchaseMismatch===0);
  i79v47status_(ss,[
   ['version',I79V47],['상태','PASS'],['단계','DONE'],['메시지','안정속성 fingerprint 기반 VAT→원본→CARD→Issue74 주문ID 역추적 완료'],
   ['현재_CARD주문',card.orders],['현재_MATCHED',cs.MATCHED||0],['현재_NON_CARD',cs.NON_CARD||0],['현재_NO_MATCH',cs.NO_MATCH||0],['현재_AMBIGUOUS',cs.AMBIGUOUS||0],['현재_CARD매입합',card.purchase],
   ['현재_VAT상세행',vat.detailRows],['현재_VAT주문',vat.orders],['현재_VAT매입합',vat.purchase],
   ['현재_SOURCE활성상세행',src.activeDetails],['현재_SOURCE활성주문',Object.keys(src.activeMap).length],['현재_SOURCE활성매입합',src.activePurchase],['현재_SOURCE광의상세행',src.broadDetails],['현재_SOURCE광의주문',Object.keys(src.broadMap).length],['현재_SOURCE광의매입합',src.broadPurchase],
   ['VAT→SOURCE_유일매핑',vm.unique],['VAT→SOURCE_모호',vm.ambiguous],['VAT→SOURCE_미매핑',vm.unmapped],['VAT→SOURCE_충돌',vm.collision],['VAT→SOURCE_ID동일',vm.sameId],['VAT→SOURCE_ID변경',vm.changedId],['VAT→SOURCE_active매핑',vm.activeMapped],['VAT→SOURCE_broad보조매핑',vm.broadMapped],['VAT→SOURCE_매입불일치',vm.purchaseMismatch],['VAT_snapshot_SOURCE관계',snapshotExact?'1355주문 fingerprint exact subset':'추가검토'],
   ['SOURCE활성중_VAT미포함주문',extra.orders],['SOURCE활성중_VAT미포함상세행',extra.details],['SOURCE활성중_VAT미포함매입합',extra.purchase],
   ['CARD→SOURCE_해결',cm.resolved],['CARD→SOURCE_directID',cm.direct],['CARD→SOURCE_VAT번역',cm.viaVat],['CARD→SOURCE_미해결',cm.unresolved],['CARD→SOURCE_충돌',cm.collision],['CARD→SOURCE_매입불일치',cm.purchaseMismatch],
   ['번역후_Issue74_overlap',th.overlap],['번역후_현재ONLY',th.currentOnly],['번역후_Issue74_ONLY',th.historyOnly],['번역후_매입변경',th.purchaseMismatch],['번역후_상태변경',th.statusChanged],['번역후_상태전이',th.transitionText],['번역후_매입변경순차이',th.sharedDelta],['현재-Issue74_총매입차이',card.purchase-hist.purchase],
   ['매입변경_SOURCE=현재',th.sourceCurrent],['매입변경_SOURCE=Issue74',th.sourceHistory],['매입변경_SOURCE=둘다',th.sourceBoth],['매입변경_SOURCE=둘다아님',th.sourceNeither],['매입변경_SOURCE근거없음',th.sourceMissing],
   ['과거22_현재해결',old.resolved],['과거22_현재MATCHED',old.stats.MATCHED||0],['과거22_현재NON_CARD',old.stats.NON_CARD||0],['과거22_현재NO_MATCH',old.stats.NO_MATCH||0],['과거22_현재AMBIGUOUS',old.stats.AMBIGUOUS||0],['과거22_현재미해결',old.stats.MISSING||0],
   ['v2_SAFE건수',old.safe.total],['v2_SAFE_현재MATCHED',old.safe.MATCHED||0],['v2_SAFE_현재NON_CARD',old.safe.NON_CARD||0],['v2_SAFE_현재NO_MATCH',old.safe.NO_MATCH||0],['v2_SAFE_현재미해결',old.safe.MISSING||0],
   ['v2_잔여16건수',old.remain.total],['v2_잔여16_현재MATCHED',old.remain.MATCHED||0],['v2_잔여16_현재NON_CARD',old.remain.NON_CARD||0],['v2_잔여16_현재NO_MATCH',old.remain.NO_MATCH||0],['v2_잔여16_현재미해결',old.remain.MISSING||0],
   ['Issue74_주문',hist.orders],['Issue74_MATCHED',hs.MATCHED||0],['Issue74_NON_CARD',hs.NON_CARD||0],['Issue74_NO_MATCH',hs.NO_MATCH||0],['Issue74_매입합',hist.purchase],
   ['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
  ]);
  return{ok:true,done:true,version:I79V47,vatMapped:vm.unique,cardResolved:cm.resolved,translatedOverlap:th.overlap,sourceCurrent:th.sourceCurrent,sourceHistory:th.sourceHistory,old22Resolved:old.resolved};
 }catch(e){var msg=String(e&&e.message?e.message:e);try{i79v47status_(ss,[['version',I79V47],['상태','ERROR'],['단계','FAILED'],['메시지','안정속성 ID 역추적 실패'],['실행시작',started],['오류',msg],['완료시각',new Date().toISOString()]]);}catch(_e){}throw e;}
}
function i79v47card_(sh,label,scope){
 if(!sh)throw new Error(label+' 시트 누락');var v=sh.getDataRange().getValues(),hr=i79v47header_(v,['주문번호','카드매칭상태']),h=v[hr];
 var x={year:i79v47ix_(h,['신고연도']),half:i79v47ix_(h,['반기']),business:i79v47ix_(h,['사업자등록번호']),account:i79v47ix_(h,['쿠팡계정ID']),order:i79v47ix_(h,['주문번호']),purchase:i79v47ix_(h,['주문매입금액','매입금액']),status:i79v47ix_(h,['카드매칭상태']),reason:i79v47ix_(h,['카드매칭근거'])};
 ['business','account','order','purchase','status'].forEach(function(k){if(x[k]<0)throw new Error(label+' header 누락 '+k);});
 var rows=[],map={},stats={},purchase=0;
 for(var r=hr+1;r<v.length;r++){var z=v[r],no=i79v47t_(z[x.order]);if(!no)continue;if(scope){if(x.year<0||x.half<0)throw new Error(label+' 연도/반기 header 누락');if(i79v47t_(z[x.year])!=='2026'||i79v47t_(z[x.half])!=='상반기')continue;}var o={orderNo:no,business:i79v47t_(z[x.business]),account:i79v47t_(z[x.account]),purchase:i79v47n_(z[x.purchase]),status:i79v47t_(z[x.status]).toUpperCase(),reason:x.reason>=0?i79v47t_(z[x.reason]):''};if(map[no])throw new Error(label+' 주문번호 중복 '+no);map[no]=o;rows.push(o);purchase+=o.purchase;stats[o.status]=(stats[o.status]||0)+1;}
 return{rows:rows,map:map,stats:stats,orders:rows.length,purchase:Math.round(purchase)};
}
function i79v47vat_(sh){
 var v=sh.getDataRange().getValues(),h=v[0]||[],x={year:i79v47ix_(h,['신고연도']),half:i79v47ix_(h,['반기']),date:i79v47ix_(h,['날짜']),account:i79v47ix_(h,['쿠팡계정ID']),business:i79v47ix_(h,['사업자등록번호']),order:i79v47ix_(h,['주문번호']),customer:i79v47ix_(h,['고객명']),brand:i79v47ix_(h,['브랜드명']),productNo:i79v47ix_(h,['상품번호']),productName:i79v47ix_(h,['상품명']),qty:i79v47ix_(h,['판매수량']),sales:i79v47ix_(h,['순수매출액']),settle:i79v47ix_(h,['정산기준금액']),purchase:i79v47ix_(h,['매입금액'])};
 Object.keys(x).forEach(function(k){if(x[k]<0)throw new Error('VAT header 누락 '+k);});var map={},detail=0,purchase=0;
 for(var r=1;r<v.length;r++){var z=v[r];if(i79v47t_(z[x.year])!=='2026'||i79v47t_(z[x.half])!=='상반기')continue;var no=i79v47t_(z[x.order]);if(!no)continue;var line=i79v47line_({account:z[x.account],date:z[x.date],customer:z[x.customer],brand:z[x.brand],productNo:z[x.productNo],productName:z[x.productName],qty:z[x.qty],sales:z[x.sales],settle:z[x.settle],purchase:z[x.purchase]});if(!map[no])map[no]={orderNo:no,lines:[],purchase:0,detailRows:0,account:i79v47t_(z[x.account]),business:i79v47t_(z[x.business])};map[no].lines.push(line);map[no].purchase+=line.purchase;map[no].detailRows++;detail++;purchase+=line.purchase;}
 Object.keys(map).forEach(function(k){i79v47finalAgg_(map[k]);});return{map:map,orders:Object.keys(map).length,detailRows:detail,purchase:Math.round(purchase)};
}
function i79v47source_(sh){
 var v=sh.getDataRange().getValues(),h=v[0]||[],x={date:i79v47ix_(h,['마켓주문일자','주문일자','결제일자','주문일시']),order:i79v47ix_(h,['마켓주문번호','주문번호','주문ID','주문ID(마켓)']),sales:i79v47ix_(h,['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액']),settle:i79v47ix_(h,['정산예정금액(원)','정산예정금액','실제정산금액','정산금액']),status:i79v47ix_(h,['마켓주문상태','주문상태','상태','클레임상태','처리상태']),customer:i79v47ix_(h,['고객명','수령인','수취인','구매자','주문자']),brand:i79v47ix_(h,['브랜드명','브랜드']),productNo:i79v47ix_(h,['마켓상품번호','상품번호','상품코드','판매자상품코드']),productName:i79v47ix_(h,['상품명','상품명(옵션포함)','등록상품명']),qty:i79v47ix_(h,['판매수량','수량','구매수량'])};
 if(x.date<0)x.date=0;if(x.order<0)x.order=2;if(x.sales<0)x.sales=6;if(x.productNo<0)x.productNo=4;var pidx=28;if((h||[]).length<=pidx)throw new Error('SOURCE AC 매입금액 열 누락');var activeMap={},broadMap={},ad=0,bd=0,ap=0,bp=0;
 for(var r=1;r<v.length;r++){var z=v[r],no=i79v47t_(z[x.order]);if(!no)continue;var dt=i79v47date_(z[x.date]);if(!dt||dt<'2026-01-01'||dt>'2026-06-30')continue;var sales=i79v47n_(z[x.sales]);if(!sales)continue;var settle=x.settle>=0?i79v47n_(z[x.settle]):0;if(!settle)settle=Math.round(sales*0.901);var line=i79v47line_({account:z[3],date:z[x.date],customer:x.customer>=0?z[x.customer]:'',brand:x.brand>=0?z[x.brand]:'',productNo:x.productNo>=0?z[x.productNo]:'',productName:x.productName>=0?z[x.productName]:'',qty:x.qty>=0?z[x.qty]:1,sales:sales,settle:settle,purchase:z[pidx]});i79v47addAgg_(broadMap,no,line);bd++;bp+=line.purchase;var st=x.status>=0?i79v47t_(z[x.status]):'';if(/취소|반품|환불/.test(st))continue;i79v47addAgg_(activeMap,no,line);ad++;ap+=line.purchase;}
 Object.keys(activeMap).forEach(function(k){i79v47finalAgg_(activeMap[k]);});Object.keys(broadMap).forEach(function(k){i79v47finalAgg_(broadMap[k]);});return{activeMap:activeMap,broadMap:broadMap,activeDetails:ad,broadDetails:bd,activePurchase:Math.round(ap),broadPurchase:Math.round(bp)};
}
function i79v47line_(o){var q=i79v47n_(o.qty)||1,s=i79v47n_(o.sales),st=i79v47n_(o.settle),p=i79v47n_(o.purchase),line={account:i79v47norm_(o.account),date:i79v47date_(o.date),customer:i79v47norm_(o.customer),brand:i79v47norm_(o.brand),productNo:i79v47norm_(o.productNo),productName:i79v47norm_(o.productName),qty:q,sales:s,settle:st,purchase:p};line.strict=[line.account,line.date,line.customer,line.brand,line.productNo,line.productName,q,s,st,p].join('\u001f');line.relaxed=[line.account,line.date,line.productNo,line.productName,q,s,p].join('\u001f');return line;}
function i79v47addAgg_(map,no,line){if(!map[no])map[no]={orderNo:no,lines:[],purchase:0,detailRows:0};map[no].lines.push(line);map[no].purchase+=line.purchase;map[no].detailRows++;}
function i79v47finalAgg_(o){o.purchase=Math.round(o.purchase);o.strict=o.lines.map(function(x){return x.strict;}).sort().join('\u001e');o.relaxed=o.lines.map(function(x){return x.relaxed;}).sort().join('\u001e');}
function i79v47index_(map,field){var idx={};Object.keys(map).forEach(function(no){var k=map[no][field];if(!idx[k])idx[k]=[];idx[k].push(no);});return idx;}
function i79v47pick_(agg,activeStrict,activeRelax,broadStrict,broadRelax){var a=activeStrict[agg.strict]||[];if(a.length===1)return{id:a[0],method:'ACTIVE_STRICT'};var b=activeRelax[agg.relaxed]||[];if(b.length===1)return{id:b[0],method:'ACTIVE_RELAXED'};var c=broadStrict[agg.strict]||[];if(c.length===1)return{id:c[0],method:'BROAD_STRICT'};var d=broadRelax[agg.relaxed]||[];if(d.length===1)return{id:d[0],method:'BROAD_RELAXED'};var amb=Math.max(a.length,b.length,c.length,d.length);return{id:'',method:amb>1?'AMBIGUOUS':'NO_MATCH',amb:amb};}
function i79v47mapVatToSource_(vat,src){
 var as=i79v47index_(src.activeMap,'strict'),ar=i79v47index_(src.activeMap,'relaxed'),bs=i79v47index_(src.broadMap,'strict'),br=i79v47index_(src.broadMap,'relaxed'),provisional={},reverse={};
 Object.keys(vat.map).forEach(function(id){var p=i79v47pick_(vat.map[id],as,ar,bs,br);provisional[id]=p;if(p.id){if(!reverse[p.id])reverse[p.id]=[];reverse[p.id].push(id);}});
 var map={},unique=0,amb=0,un=0,col=0,same=0,chg=0,active=0,broad=0,pm=0;
 Object.keys(vat.map).forEach(function(id){var p=provisional[id];if(!p.id){if(p.method==='AMBIGUOUS')amb++;else un++;map[id]={sourceId:'',method:p.method};return;}if((reverse[p.id]||[]).length!==1){col++;map[id]={sourceId:'',method:'COLLISION'};return;}var so=src.activeMap[p.id]||src.broadMap[p.id],vo=vat.map[id];unique++;if(id===p.id)same++;else chg++;if(/^ACTIVE_/.test(p.method))active++;else broad++;if(!so||so.purchase!==vo.purchase)pm++;map[id]={sourceId:p.id,method:p.method,source:so};});
 return{map:map,unique:unique,ambiguous:amb,unmapped:un,collision:col,sameId:same,changedId:chg,activeMapped:active,broadMapped:broad,purchaseMismatch:pm};
}
function i79v47mapCardToSource_(card,vat,src,vm){
 var map={},rev={},direct=0,via=0,un=0,pm=0;
 card.rows.forEach(function(c){var sid='',method='';if(src.activeMap[c.orderNo]||src.broadMap[c.orderNo]){sid=c.orderNo;method='DIRECT_SOURCE_ID';direct++;}else if(vat.map[c.orderNo]&&vm.map[c.orderNo]&&vm.map[c.orderNo].sourceId){sid=vm.map[c.orderNo].sourceId;method='VIA_VAT_FINGERPRINT';via++;}else{un++;method='UNRESOLVED';}var so=sid?(src.activeMap[sid]||src.broadMap[sid]):null;if(so&&so.purchase!==c.purchase)pm++;map[c.orderNo]={card:c,sourceId:sid,method:method,source:so};if(sid){if(!rev[sid])rev[sid]=[];rev[sid].push(c.orderNo);}});
 var collision=0;Object.keys(rev).forEach(function(sid){if(rev[sid].length>1){collision+=rev[sid].length;rev[sid].forEach(function(cid){map[cid].sourceId='';map[cid].method='SOURCE_COLLISION';});}});return{map:map,reverse:rev,resolved:card.orders-un-collision,direct:direct,viaVat:via,unresolved:un,collision:collision,purchaseMismatch:pm};
}
function i79v47translatedHistory_(cm,hist,src,v2){
 var rows=[],prows=[],seen={},overlap=0,co=0,pm=0,sc=0,delta=0,tr={},sCur=0,sHist=0,sBoth=0,sNeither=0,sMissing=0;
 Object.keys(cm.map).forEach(function(cid){var m=cm.map[cid],c=m.card,sid=m.sourceId,h=sid?hist.map[sid]:null,so=sid?(src.activeMap[sid]||src.broadMap[sid]):null;if(!sid||!h){co++;rows.push([cid,c.status,c.purchase,sid,m.method,so?so.purchase:'','',h?h.status:'',h?h.purchase:'','',I79V47_OLD22.indexOf(sid)>=0?'Y':'',v2.safe[sid]?'Y':'']);return;}overlap++;seen[sid]=1;var d=c.purchase-h.purchase;delta+=d;if(d!==0){pm++;var eqc=so&&so.purchase===c.purchase,eqh=so&&so.purchase===h.purchase;if(!so)sMissing++;else if(eqc&&eqh)sBoth++;else if(eqc)sCur++;else if(eqh)sHist++;else sNeither++;prows.push([sid,cid,h.status,c.status,h.purchase,c.purchase,d,so?so.purchase:'',eqc?'Y':'',eqh?'Y':'',m.method]);}if(c.status!==h.status){sc++;var k=(h.status||'')+'→'+(c.status||'');tr[k]=(tr[k]||0)+1;}rows.push([cid,c.status,c.purchase,sid,m.method,so?so.purchase:'',sid,h.status,h.purchase,d,I79V47_OLD22.indexOf(sid)>=0?'Y':'',v2.safe[sid]?'Y':'']);});
 var ho=0;Object.keys(hist.map).forEach(function(id){if(!seen[id])ho++;});return{rows:rows,purchaseRows:prows,overlap:overlap,currentOnly:co,historyOnly:ho,purchaseMismatch:pm,statusChanged:sc,transitionText:Object.keys(tr).sort().map(function(k){return k+':'+tr[k];}).join(' / '),sharedDelta:Math.round(delta),sourceCurrent:sCur,sourceHistory:sHist,sourceBoth:sBoth,sourceNeither:sNeither,sourceMissing:sMissing};
}
function i79v47old22_(cm,v2){var bySource={},stats={},safe={total:0},remain={total:0},resolved=0;Object.keys(cm.map).forEach(function(cid){var m=cm.map[cid];if(m.sourceId)bySource[m.sourceId]=m.card;});I79V47_OLD22.forEach(function(id){var c=bySource[id]||null,st=c?c.status:'MISSING',b=(st==='MATCHED'||st==='NON_CARD'||st==='NO_MATCH'||st==='AMBIGUOUS')?st:'MISSING';stats[b]=(stats[b]||0)+1;if(c)resolved++;var q=v2.safe[id]?safe:remain;q.total++;q[b]=(q[b]||0)+1;});return{stats:stats,safe:safe,remain:remain,resolved:resolved};}
function i79v47sourceExtra_(src,vm){var used={};Object.keys(vm.map).forEach(function(id){var s=vm.map[id].sourceId;if(s)used[s]=1;});var o=0,d=0,p=0;Object.keys(src.activeMap).forEach(function(id){if(used[id])return;o++;d+=src.activeMap[id].detailRows;p+=src.activeMap[id].purchase;});return{orders:o,details:d,purchase:Math.round(p)};}
function i79v47v2_(sh){var safe={};if(!sh)return{safe:safe};var v=sh.getDataRange().getValues(),h=v[0]||[],xo=i79v47ix_(h,['주문번호']),xv=i79v47ix_(h,['PREVIEW판정']);if(xo<0||xv<0)return{safe:safe};for(var r=1;r<v.length;r++){var id=i79v47t_(v[r][xo]),z=i79v47t_(v[r][xv]);if(id&&/^SAFE_/.test(z))safe[id]=1;}return{safe:safe};}
function i79v47writeTranslation_(ss,rows){var sh=ss.getSheetByName('ISSUE79_v47_ID번역')||ss.insertSheet('ISSUE79_v47_ID번역');i79v47write_(sh,['현재CARD_ID','현재상태','현재매입','해결SOURCE_ID','해결방법','현재SOURCE매입','Issue74_ID','Issue74상태','Issue74매입','현재-Issue74차이','과거NO_MATCH22','v2_SAFE'],rows,1);}
function i79v47writePurchase_(ss,rows){var sh=ss.getSheetByName('ISSUE79_v47_매입변경근거')||ss.insertSheet('ISSUE79_v47_매입변경근거');i79v47write_(sh,['SOURCE_ID','현재CARD_ID','Issue74상태','현재상태','Issue74매입','현재매입','차이','현재SOURCE매입','SOURCE=현재','SOURCE=Issue74','ID해결방법'],rows,1);}
function i79v47write_(sh,h,rows,textCol){if(sh.getFilter())sh.getFilter().remove();sh.clear();sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length){sh.getRange(2,textCol,rows.length,1).setNumberFormat('@');sh.getRange(2,1,rows.length,h.length).setValues(rows);}sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}sh.autoResizeColumns(1,h.length);}
function i79v47status_(ss,pairs){var sh=ss.getSheetByName('ISSUE79_식별자역추적상태')||ss.insertSheet('ISSUE79_식별자역추적상태');sh.clearContents();var a=[['항목','값']].concat(pairs||[]);sh.getRange(1,1,a.length,2).setValues(a);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}
function i79v47header_(v,need){for(var r=0;r<Math.min(50,v.length);r++){var ok=true;for(var j=0;j<need.length;j++)if(i79v47ix_(v[r],[need[j]])<0){ok=false;break;}if(ok)return r;}throw new Error('header 탐지 실패 '+need.join(','));}
function i79v47ix_(h,names){for(var j=0;j<names.length;j++){var n=i79v47c_(names[j]);for(var i=0;i<h.length;i++)if(i79v47c_(h[i])===n)return i;}return-1;}
function i79v47t_(v){return String(v==null?'':v).trim();}
function i79v47c_(v){return i79v47t_(v).toLowerCase().replace(/\s+/g,'');}
function i79v47norm_(v){return i79v47t_(v).toLowerCase().replace(/\s+/g,' ').trim();}
function i79v47n_(v){if(typeof v==='number')return isFinite(v)?Math.round(v):0;var s=i79v47t_(v).replace(/,/g,'').replace(/[^0-9.\-]/g,'');var n=Number(s);return isFinite(n)?Math.round(n):0;}
function i79v47date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i79v47t_(v),m=s.match(/(20\d{2})\D?(\d{1,2})\D?(\d{1,2})/);if(!m)return'';return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);}
function i79v47sig_(sh){var rg=sh.getDataRange(),s=sh.getName()+'|'+rg.getNumRows()+'|'+rg.getNumColumns()+'|'+JSON.stringify(rg.getDisplayValues());return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,s,Utilities.Charset.UTF_8));}
