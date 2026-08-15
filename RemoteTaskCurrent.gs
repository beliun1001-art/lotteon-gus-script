var LOTTEON_REMOTE_TASK={id:'ISSUE79-V43-CURRENT-H1-RECON',title:'Issue79 현재 2026상반기 카드/VAT 1:1 READ-ONLY 대조',enabled:true,statusSheet:'ISSUE79_현재상태대조상태'};
var I79V43='v4.3-ISSUE79-CURRENT-2026H1-RECON-READONLY';
var I79V43_CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
var I79V43_OLD22=['30100189414967','2100191293730','23100192054997','29100191903703','5100191540477','26100192791914','17100196061992','18100195889436','12100196902987','12100197310918','11100197821559','5100198099688','19100198942446','5101102615736','6101100531065','12101138407306','12101154659935','17100199506890','19100199527230','29100199660972','16101220120574','9101251100578'];
function runLotteonRemoteTaskStartRemote_(){return i79v43run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v43run_();}
function i79v43run_(){
 var ss=SpreadsheetApp.getActive(),before={},started=new Date().toISOString();
 try{
  i79v43status_(ss,[['version',I79V43],['상태','RUNNING'],['단계','READ'],['메시지','현재 2026 상반기 카드검증/VAT 1:1 READ-ONLY 대조 중'],['실행시작',started]]);
  I79V43_CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락: '+n);before[n]=i79v43sig_(sh);});
  var card=i79v43readCard_(ss.getSheetByName('부가세_카드매칭검증'));
  var vat=i79v43readVat_(ss.getSheetByName('부가세_신고자료'));
  var cmp=i79v43compare_(card,vat);
  var pv=i79v43readV2_(ss.getSheetByName('ISSUE79_NOMATCH22_최신재검수'));
  var old=i79v43old22_(card,vat,pv);
  i79v43writeOld_(ss,old.rows);
  i79v43writeMismatch_(ss,cmp.rows);
  var changed=[];I79V43_CORE.forEach(function(n){if(i79v43sig_(ss.getSheetByName(n))!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반: '+changed.join(','));
  var st=card.stats,other=card.rows.length-(st.MATCHED||0)-(st.NON_CARD||0)-(st.NO_MATCH||0)-(st.AMBIGUOUS||0);
  i79v43status_(ss,[
   ['version',I79V43],['상태','PASS'],['단계','DONE'],['메시지','현재 2026 상반기 카드검증/VAT 및 과거 NO_MATCH 22 현재상태 READ-ONLY 대조 완료'],
   ['카드검증_2026H1주문',card.rows.length],['현재_MATCHED',st.MATCHED||0],['현재_NON_CARD',st.NON_CARD||0],['현재_NO_MATCH',st.NO_MATCH||0],['현재_AMBIGUOUS',st.AMBIGUOUS||0],['현재_기타상태',other],['현재_카드검증매입합계',card.purchase],['카드검증_중복키',card.duplicates],
   ['VAT_2026H1상세행',vat.detailRows],['VAT_2026H1주문',vat.rows.length],['VAT_매입합계',vat.purchase],['VAT_중복집계주문',vat.multiOrders],
   ['VAT_CARD_overlap',cmp.overlap],['CARD_ONLY',cmp.cardOnly],['VAT_ONLY',cmp.vatOnly],['매입금액불일치주문',cmp.purchaseMismatch],['카드-VAT_매입합계차이',card.purchase-vat.purchase],['불일치행출력',cmp.rows.length],
   ['과거22_현재MATCHED',old.stats.MATCHED||0],['과거22_현재NON_CARD',old.stats.NON_CARD||0],['과거22_현재NO_MATCH',old.stats.NO_MATCH||0],['과거22_현재AMBIGUOUS',old.stats.AMBIGUOUS||0],['과거22_현재기타',old.stats.OTHER||0],['과거22_현재누락',old.stats.MISSING||0],['과거22_매입불일치',old.purchaseMismatch],
   ['v2_PREVIEW존재',pv.exists?'YES':'NO'],['v2_SAFE건수',old.safe.total],['v2_SAFE_현재MATCHED',old.safe.MATCHED||0],['v2_SAFE_현재NON_CARD',old.safe.NON_CARD||0],['v2_SAFE_현재NO_MATCH',old.safe.NO_MATCH||0],['v2_SAFE_현재기타',old.safe.OTHER||0],
   ['v2_잔여16건수',old.remain.total],['v2_잔여16_현재MATCHED',old.remain.MATCHED||0],['v2_잔여16_현재NON_CARD',old.remain.NON_CARD||0],['v2_잔여16_현재NO_MATCH',old.remain.NO_MATCH||0],['v2_잔여16_현재기타',old.remain.OTHER||0],
   ['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
  ]);
  return{ok:true,done:true,version:I79V43,cardOrders:card.rows.length,vatOrders:vat.rows.length,noMatch:st.NO_MATCH||0,purchaseMismatch:cmp.purchaseMismatch,old22NoMatch:old.stats.NO_MATCH||0};
 }catch(e){var msg=String(e&&e.message?e.message:e);i79v43status_(ss,[['version',I79V43],['상태','ERROR'],['단계','FAILED'],['메시지','현재상태 READ-ONLY 대조 실패'],['실행시작',started],['오류',msg],['완료시각',new Date().toISOString()]]);throw e;}
}
function i79v43readCard_(sh){
 var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=i79v43header_(v,['주문번호','카드매칭상태']),h=v[hr];
 var x={year:i79v43ix_(h,['신고연도']),half:i79v43ix_(h,['반기']),business:i79v43ix_(h,['사업자등록번호']),account:i79v43ix_(h,['쿠팡계정ID']),order:i79v43ix_(h,['주문번호']),purchase:i79v43ix_(h,['주문매입금액','매입금액']),status:i79v43ix_(h,['카드매칭상태']),company:i79v43ix_(h,['구매카드사']),name:i79v43ix_(h,['구매카드명']),end4:i79v43ix_(h,['카드번호끝4']),adate:i79v43ix_(h,['승인일']),approval:i79v43ix_(h,['승인번호']),aamount:i79v43ix_(h,['승인금액']),merchant:i79v43ix_(h,['가맹점명']),source:i79v43ix_(h,['원본파일']),reason:i79v43ix_(h,['카드매칭근거'])};
 ['year','half','business','account','order','purchase','status'].forEach(function(k){if(x[k]<0)throw new Error('카드검증 header 누락 '+k);});
 var rows=[],map={},byOrder={},stats={},purchase=0,dup=0;
 for(var r=hr+1;r<v.length;r++){var no=i79v43t_(d[r][x.order]);if(!no)continue;if(i79v43c_(v[r][x.year])!=='2026'||!i79v43h1_(v[r][x.half]))continue;var biz=i79v43biz_(d[r][x.business]),acc=i79v43c_(d[r][x.account]),key=[biz,acc,i79v43c_(no)].join('|'),st=i79v43t_(v[r][x.status]).toUpperCase(),amt=i79v43n_(v[r][x.purchase]);var o={key:key,business:biz,account:i79v43t_(d[r][x.account]),orderNo:no,purchase:amt,status:st,company:x.company>=0?i79v43t_(v[r][x.company]):'',cardName:x.name>=0?i79v43t_(v[r][x.name]):'',end4:x.end4>=0?i79v43t_(d[r][x.end4]):'',approvalDate:x.adate>=0?i79v43t_(d[r][x.adate]):'',approvalNo:x.approval>=0?i79v43t_(d[r][x.approval]):'',approvalAmount:x.aamount>=0?i79v43n_(v[r][x.aamount]):0,merchant:x.merchant>=0?i79v43t_(v[r][x.merchant]):'',source:x.source>=0?i79v43t_(v[r][x.source]):'',reason:x.reason>=0?i79v43t_(v[r][x.reason]):''};rows.push(o);purchase+=amt;stats[st]=(stats[st]||0)+1;if(map[key])dup++;else map[key]=o;(byOrder[no]||(byOrder[no]=[])).push(o);}
 return{rows:rows,map:map,byOrder:byOrder,stats:stats,purchase:Math.round(purchase),duplicates:dup};
}
function i79v43readVat_(sh){
 var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=i79v43header_(v,['신고연도','주문번호']),h=v[hr];
 var x={year:i79v43ix_(h,['신고연도']),half:i79v43ix_(h,['반기']),business:i79v43ix_(h,['사업자등록번호']),account:i79v43ix_(h,['쿠팡계정ID']),order:i79v43ix_(h,['주문번호','마켓주문번호']),purchase:i79v43ix_(h,['매입금액','주문매입금액'])};
 ['year','half','business','account','order','purchase'].forEach(function(k){if(x[k]<0)throw new Error('VAT header 누락 '+k);});
 var map={},detail=0;
 for(var r=hr+1;r<v.length;r++){var no=i79v43t_(d[r][x.order]);if(!no)continue;if(i79v43c_(v[r][x.year])!=='2026'||!i79v43h1_(v[r][x.half]))continue;detail++;var biz=i79v43biz_(d[r][x.business]),acc=i79v43c_(d[r][x.account]),key=[biz,acc,i79v43c_(no)].join('|');if(!map[key])map[key]={key:key,business:biz,account:i79v43t_(d[r][x.account]),orderNo:no,purchase:0,detailRows:0};map[key].purchase+=i79v43n_(v[r][x.purchase]);map[key].detailRows++;}
 var rows=[],p=0,multi=0;Object.keys(map).forEach(function(k){map[k].purchase=Math.round(map[k].purchase);rows.push(map[k]);p+=map[k].purchase;if(map[k].detailRows>1)multi++;});return{rows:rows,map:map,detailRows:detail,purchase:Math.round(p),multiOrders:multi};
}
function i79v43compare_(card,vat){var rows=[],overlap=0,cardOnly=0,vatOnly=0,pm=0,seen={};card.rows.forEach(function(c){var v=vat.map[c.key];if(!v){cardOnly++;rows.push(['CARD_ONLY',c.business,c.account,c.orderNo,c.status,c.purchase,'',c.purchase,c.reason]);return;}overlap++;seen[c.key]=1;var delta=c.purchase-v.purchase;if(delta!==0){pm++;rows.push(['PURCHASE_MISMATCH',c.business,c.account,c.orderNo,c.status,c.purchase,v.purchase,delta,c.reason]);}});vat.rows.forEach(function(v){if(!seen[v.key]&&!card.map[v.key]){vatOnly++;rows.push(['VAT_ONLY',v.business,v.account,v.orderNo,'','',v.purchase,-v.purchase,'']);}});return{rows:rows,overlap:overlap,cardOnly:cardOnly,vatOnly:vatOnly,purchaseMismatch:pm};}
function i79v43readV2_(sh){if(!sh)return{exists:false,map:{},safe:{}};var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),h=v[0],xo=i79v43ix_(h,['주문번호']),xv=i79v43ix_(h,['PREVIEW판정']),xt=i79v43ix_(h,['재매칭기준']),xa=i79v43ix_(h,['재매칭목표금액']);if(xo<0||xv<0)return{exists:false,map:{},safe:{}};var m={},safe={};for(var r=1;r<v.length;r++){var no=i79v43t_(d[r][xo]);if(!no)continue;var z={verdict:i79v43t_(v[r][xv]),targetType:xt>=0?i79v43t_(v[r][xt]):'',targetAmount:xa>=0?i79v43n_(v[r][xa]):0};m[no]=z;if(/^SAFE_/.test(z.verdict))safe[no]=1;}return{exists:true,map:m,safe:safe};}
function i79v43old22_(card,vat,pv){var rows=[],stats={},safe={total:0},remain={total:0},pm=0;I79V43_OLD22.forEach(function(no){var ca=card.byOrder[no]||[],c=ca[0]||null,v=c?vat.map[c.key]:null,st=c?c.status:'MISSING',bucket=(st==='MATCHED'||st==='NON_CARD'||st==='NO_MATCH'||st==='AMBIGUOUS')?st:(st==='MISSING'?'MISSING':'OTHER'),v2=pv.map[no]||{},isSafe=!!pv.safe[no];stats[bucket]=(stats[bucket]||0)+1;var b=isSafe?safe:remain;b.total++;b[bucket]=(b[bucket]||0)+1;var vp=v?v.purchase:'',delta=(c&&v)?c.purchase-v.purchase:'';if(delta!==''&&delta!==0)pm++;rows.push([no,isSafe?'Y':'',v2.verdict||'',v2.targetType||'',v2.targetAmount||'',st,c?c.business:'',c?c.account:'',c?c.purchase:'',vp,delta,c?c.company:'',c?c.cardName:'',c?c.end4:'',c?c.approvalDate:'',c?c.approvalNo:'',c?c.approvalAmount:'',c?c.reason:'']);});return{rows:rows,stats:stats,safe:safe,remain:remain,purchaseMismatch:pm};}
function i79v43writeOld_(ss,rows){var sh=ss.getSheetByName('ISSUE79_기존22현재상태')||ss.insertSheet('ISSUE79_기존22현재상태');if(sh.getFilter())sh.getFilter().remove();sh.clear();var h=['주문번호','v2_SAFE','v2판정','v2재매칭기준','v2목표금액','현재상태','사업자등록번호','쿠팡계정ID','현재카드매입','VAT매입','차이','구매카드사','구매카드명','끝4','승인일','승인번호','승인금액','현재매칭근거'];sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length){sh.getRange(2,1,rows.length,1).setNumberFormat('@');sh.getRange(2,1,rows.length,h.length).setValues(rows);}sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}sh.autoResizeColumns(1,h.length);}
function i79v43writeMismatch_(ss,rows){var sh=ss.getSheetByName('ISSUE79_매입금액대조')||ss.insertSheet('ISSUE79_매입금액대조');if(sh.getFilter())sh.getFilter().remove();sh.clear();var h=['구분','사업자등록번호','쿠팡계정ID','주문번호','현재카드상태','카드검증매입','VAT매입','차이','현재매칭근거'];sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length){sh.getRange(2,4,rows.length,1).setNumberFormat('@');sh.getRange(2,1,rows.length,h.length).setValues(rows);}sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}sh.autoResizeColumns(1,h.length);}
function i79v43status_(ss,pairs){var sh=ss.getSheetByName('ISSUE79_현재상태대조상태')||ss.insertSheet('ISSUE79_현재상태대조상태');sh.clearContents();var a=[['항목','값']].concat(pairs||[]);sh.getRange(1,1,a.length,2).setValues(a);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}
function i79v43sig_(sh){var rg=sh.getDataRange(),s=sh.getName()+'|'+rg.getNumRows()+'|'+rg.getNumColumns()+'|'+JSON.stringify(rg.getDisplayValues());return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,s,Utilities.Charset.UTF_8));}
function i79v43header_(v,need){for(var r=0;r<Math.min(50,v.length);r++){var ok=true;for(var j=0;j<need.length;j++)if(i79v43ix_(v[r],need[j]==='주문번호'?['주문번호','마켓주문번호']:[need[j]])<0){ok=false;break;}if(ok)return r;}throw new Error('header 탐지 실패: '+need.join(','));}
function i79v43ix_(h,a){for(var i=0;i<h.length;i++){var x=i79v43c_(h[i]);for(var j=0;j<a.length;j++)if(x===i79v43c_(a[j]))return i;}return-1;}
function i79v43t_(v){return String(v==null?'':v).trim();}
function i79v43c_(v){return i79v43t_(v).toLowerCase().replace(/\s+/g,'');}
function i79v43n_(v){if(typeof v==='number')return isFinite(v)?Math.round(v):0;var s=i79v43t_(v).replace(/,/g,'').replace(/[^0-9.\-]/g,'');var n=Number(s);return isFinite(n)?Math.round(n):0;}
function i79v43biz_(v){var d=i79v43t_(v).replace(/\D/g,'');return d||i79v43c_(v);}
function i79v43h1_(v){var s=i79v43c_(v);return s==='상반기'||s==='1h'||s==='h1'||s==='1'||s==='1반기';}
