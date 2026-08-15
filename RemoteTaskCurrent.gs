var LOTTEON_REMOTE_TASK={id:'ISSUE79-V44-ORDERKEY-HISTORY-RECON',title:'Issue79 주문번호기준 현재/VAT/Issue74 역사대조',enabled:true,statusSheet:'ISSUE79_현재상태대조상태'};
var I79V44='v4.4-ISSUE79-ORDERKEY-CURRENT-VAT-I74-READONLY';
var I79V44_CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
var I79V44_OLD22=['30100189414967','2100191293730','23100192054997','29100191903703','5100191540477','26100192791914','17100196061992','18100195889436','12100196902987','12100197310918','11100197821559','5100198099688','19100198942446','5101102615736','6101100531065','12101138407306','12101154659935','17100199506890','19100199527230','29100199660972','16101220120574','9101251100578'];
function runLotteonRemoteTaskStartRemote_(){return i79v44run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v44run_();}
function i79v44run_(){
 var ss=SpreadsheetApp.getActive(),before={},started=new Date().toISOString();
 try{
  i79v44status_(ss,[['version',I79V44],['상태','RUNNING'],['단계','READ'],['메시지','주문번호 기준 현재 카드/VAT/Issue74 역사대조 중'],['실행시작',started]]);
  I79V44_CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락: '+n);before[n]=i79v44sig_(sh);});
  var cur=i79v44readCard_(ss.getSheetByName('부가세_카드매칭검증'),'CURRENT');
  var vat=i79v44readVat_(ss.getSheetByName('부가세_신고자료'));
  var histSh=ss.getSheetByName('ISSUE74_카드회수PREVIEW');if(!histSh)throw new Error('Issue74 역사 기준 시트 누락');
  var hist=i79v44readCard_(histSh,'ISSUE74');
  var v2=i79v44readV2_(ss.getSheetByName('ISSUE79_NOMATCH22_최신재검수'));
  var cv=i79v44compareCurrentVat_(cur,vat);
  var ch=i79v44compareHistory_(cur,hist);
  var old=i79v44old22_(cur,hist,v2);
  i79v44writeCv_(ss,cv.rows);
  i79v44writeHistory_(ss,ch.rows);
  i79v44writeOld_(ss,old.rows);
  var changed=[];I79V44_CORE.forEach(function(n){if(i79v44sig_(ss.getSheetByName(n))!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반: '+changed.join(','));
  var cs=cur.stats,hs=hist.stats;
  i79v44status_(ss,[
   ['version',I79V44],['상태','PASS'],['단계','DONE'],['메시지','주문번호 기준 현재 카드/VAT/Issue74 역사 READ-ONLY 대조 완료'],
   ['현재_2026H1주문',cur.rows.length],['현재_MATCHED',cs.MATCHED||0],['현재_NON_CARD',cs.NON_CARD||0],['현재_NO_MATCH',cs.NO_MATCH||0],['현재_AMBIGUOUS',cs.AMBIGUOUS||0],['현재_매입합계',cur.purchase],['현재_주문번호중복',cur.duplicates],
   ['VAT_2026H1상세행',vat.detailRows],['VAT_2026H1주문',vat.rows.length],['VAT_매입합계',vat.purchase],['VAT_주문번호중복집계',vat.duplicates],
   ['현재VAT_주문번호overlap',cv.overlap],['현재VAT_CARD_ONLY',cv.cardOnly],['현재VAT_VAT_ONLY',cv.vatOnly],['현재VAT_매입불일치',cv.purchaseMismatch],['현재VAT_사업자불일치',cv.businessMismatch],['현재VAT_계정불일치',cv.accountMismatch],
   ['Issue74_주문',hist.rows.length],['Issue74_MATCHED',hs.MATCHED||0],['Issue74_NON_CARD',hs.NON_CARD||0],['Issue74_NO_MATCH',hs.NO_MATCH||0],['Issue74_AMBIGUOUS',hs.AMBIGUOUS||0],['Issue74_매입합계',hist.purchase],['Issue74_주문번호중복',hist.duplicates],
   ['현재Issue74_주문번호overlap',ch.overlap],['현재_ONLY',ch.currentOnly],['Issue74_ONLY',ch.historyOnly],['현재Issue74_매입불일치',ch.purchaseMismatch],['현재Issue74_상태변경',ch.statusChanged],['현재Issue74_상태전이',ch.transitionText],['현재-Issue74_매입합계차이',cur.purchase-hist.purchase],
   ['과거22_현재MATCHED',old.stats.MATCHED||0],['과거22_현재NON_CARD',old.stats.NON_CARD||0],['과거22_현재NO_MATCH',old.stats.NO_MATCH||0],['과거22_현재AMBIGUOUS',old.stats.AMBIGUOUS||0],['과거22_현재기타',old.stats.OTHER||0],['과거22_현재누락',old.stats.MISSING||0],
   ['v2_SAFE건수',old.safe.total],['v2_SAFE_현재MATCHED',old.safe.MATCHED||0],['v2_SAFE_현재NON_CARD',old.safe.NON_CARD||0],['v2_SAFE_현재NO_MATCH',old.safe.NO_MATCH||0],['v2_SAFE_현재누락',old.safe.MISSING||0],
   ['v2_잔여16건수',old.remain.total],['v2_잔여16_현재MATCHED',old.remain.MATCHED||0],['v2_잔여16_현재NON_CARD',old.remain.NON_CARD||0],['v2_잔여16_현재NO_MATCH',old.remain.NO_MATCH||0],['v2_잔여16_현재누락',old.remain.MISSING||0],
   ['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
  ]);
  return{ok:true,done:true,version:I79V44,currentVatOverlap:cv.overlap,currentHistoryOverlap:ch.overlap,old22Missing:old.stats.MISSING||0};
 }catch(e){var msg=String(e&&e.message?e.message:e);i79v44status_(ss,[['version',I79V44],['상태','ERROR'],['단계','FAILED'],['메시지','주문번호 기준 현재상태 대조 실패'],['실행시작',started],['오류',msg],['완료시각',new Date().toISOString()]]);throw e;}
}
function i79v44readCard_(sh,label){
 var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=i79v44header_(v,['주문번호','카드매칭상태']),h=v[hr];
 var x={year:i79v44ix_(h,['신고연도']),half:i79v44ix_(h,['반기']),business:i79v44ix_(h,['사업자등록번호']),account:i79v44ix_(h,['쿠팡계정ID']),order:i79v44ix_(h,['주문번호','마켓주문번호']),purchase:i79v44ix_(h,['주문매입금액','매입금액']),status:i79v44ix_(h,['카드매칭상태']),company:i79v44ix_(h,['구매카드사']),name:i79v44ix_(h,['구매카드명']),end4:i79v44ix_(h,['카드번호끝4']),adate:i79v44ix_(h,['승인일']),approval:i79v44ix_(h,['승인번호']),aamount:i79v44ix_(h,['승인금액']),reason:i79v44ix_(h,['카드매칭근거'])};
 ['order','purchase','status'].forEach(function(k){if(x[k]<0)throw new Error(label+' header 누락 '+k);});
 var rows=[],map={},stats={},purchase=0,dup=0;
 for(var r=hr+1;r<v.length;r++){
  var no=i79v44order_(v[r][x.order],d[r][x.order]);if(!no)continue;
  if(x.year>=0&&i79v44c_(v[r][x.year])!=='2026')continue;if(x.half>=0&&!i79v44h1_(v[r][x.half]))continue;
  var st=i79v44t_(v[r][x.status]).toUpperCase(),amt=i79v44n_(v[r][x.purchase]),o={orderNo:no,business:x.business>=0?i79v44biz_(d[r][x.business]):'',account:x.account>=0?i79v44t_(d[r][x.account]):'',purchase:amt,status:st,company:x.company>=0?i79v44t_(v[r][x.company]):'',cardName:x.name>=0?i79v44t_(v[r][x.name]):'',end4:x.end4>=0?i79v44t_(d[r][x.end4]):'',approvalDate:x.adate>=0?i79v44t_(d[r][x.adate]):'',approvalNo:x.approval>=0?i79v44t_(d[r][x.approval]):'',approvalAmount:x.aamount>=0?i79v44n_(v[r][x.aamount]):0,reason:x.reason>=0?i79v44t_(v[r][x.reason]):''};
  rows.push(o);purchase+=amt;stats[st]=(stats[st]||0)+1;if(map[no])dup++;else map[no]=o;
 }
 return{rows:rows,map:map,stats:stats,purchase:Math.round(purchase),duplicates:dup};
}
function i79v44readVat_(sh){
 var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=i79v44header_(v,['신고연도','주문번호']),h=v[hr];
 var x={year:i79v44ix_(h,['신고연도']),half:i79v44ix_(h,['반기']),business:i79v44ix_(h,['사업자등록번호']),account:i79v44ix_(h,['쿠팡계정ID']),order:i79v44ix_(h,['주문번호','마켓주문번호']),purchase:i79v44ix_(h,['매입금액','주문매입금액'])};
 ['year','half','order','purchase'].forEach(function(k){if(x[k]<0)throw new Error('VAT header 누락 '+k);});
 var map={},detail=0;
 for(var r=hr+1;r<v.length;r++){
  var no=i79v44order_(v[r][x.order],d[r][x.order]);if(!no)continue;if(i79v44c_(v[r][x.year])!=='2026'||!i79v44h1_(v[r][x.half]))continue;detail++;
  if(!map[no])map[no]={orderNo:no,purchase:0,detailRows:0,businesses:{},accounts:{}};var o=map[no];o.purchase+=i79v44n_(v[r][x.purchase]);o.detailRows++;if(x.business>=0){var b=i79v44biz_(d[r][x.business]);if(b)o.businesses[b]=1;}if(x.account>=0){var a=i79v44c_(d[r][x.account]);if(a)o.accounts[a]=1;}
 }
 var rows=[],p=0,dup=0;Object.keys(map).forEach(function(k){var o=map[k];o.purchase=Math.round(o.purchase);rows.push(o);p+=o.purchase;if(o.detailRows>1)dup++;});return{rows:rows,map:map,detailRows:detail,purchase:Math.round(p),duplicates:dup};
}
function i79v44compareCurrentVat_(cur,vat){
 var rows=[],overlap=0,co=0,vo=0,pm=0,bm=0,am=0,seen={};
 cur.rows.forEach(function(c){var v=vat.map[c.orderNo];if(!v){co++;rows.push(['CARD_ONLY',c.orderNo,c.status,c.business,c.account,c.purchase,'','','','']);return;}overlap++;seen[c.orderNo]=1;var delta=c.purchase-v.purchase,badB=c.business&&Object.keys(v.businesses).length&&!v.businesses[c.business],badA=c.account&&Object.keys(v.accounts).length&&!v.accounts[i79v44c_(c.account)];if(delta!==0)pm++;if(badB)bm++;if(badA)am++;if(delta!==0||badB||badA)rows.push(['MISMATCH',c.orderNo,c.status,c.business,c.account,c.purchase,v.purchase,delta,badB?'Y':'',badA?'Y':'']);});
 vat.rows.forEach(function(v){if(!seen[v.orderNo]&&!cur.map[v.orderNo]){vo++;rows.push(['VAT_ONLY',v.orderNo,'','','','',v.purchase,-v.purchase,'','']);}});
 return{rows:rows,overlap:overlap,cardOnly:co,vatOnly:vo,purchaseMismatch:pm,businessMismatch:bm,accountMismatch:am};
}
function i79v44compareHistory_(cur,hist){
 var rows=[],overlap=0,co=0,ho=0,pm=0,sc=0,seen={},tr={};
 cur.rows.forEach(function(c){var h=hist.map[c.orderNo];if(!h){co++;rows.push(['CURRENT_ONLY',c.orderNo,'',c.status,'',c.purchase,'',c.purchase,c.business,c.account,c.reason]);return;}overlap++;seen[c.orderNo]=1;var delta=c.purchase-h.purchase;if(delta!==0){pm++;rows.push(['PURCHASE_MISMATCH',c.orderNo,h.status,c.status,h.purchase,c.purchase,delta,c.business,c.account,c.reason]);}if(c.status!==h.status){sc++;var k=(h.status||'')+'→'+(c.status||'');tr[k]=(tr[k]||0)+1;rows.push(['STATUS_CHANGE',c.orderNo,h.status,c.status,h.purchase,c.purchase,delta,c.business,c.account,c.reason]);}});
 hist.rows.forEach(function(h){if(!seen[h.orderNo]&&!cur.map[h.orderNo]){ho++;rows.push(['ISSUE74_ONLY',h.orderNo,h.status,'',h.purchase,'',-h.purchase,h.business,h.account,h.reason]);}});
 var tt=Object.keys(tr).sort().map(function(k){return k+':'+tr[k];}).join(' / ');
 return{rows:rows,overlap:overlap,currentOnly:co,historyOnly:ho,purchaseMismatch:pm,statusChanged:sc,transitionText:tt};
}
function i79v44readV2_(sh){if(!sh)return{exists:false,map:{},safe:{}};var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),h=v[0],xo=i79v44ix_(h,['주문번호']),xv=i79v44ix_(h,['PREVIEW판정']);if(xo<0||xv<0)return{exists:false,map:{},safe:{}};var m={},safe={};for(var r=1;r<v.length;r++){var no=i79v44order_(v[r][xo],d[r][xo]);if(!no)continue;var z={verdict:i79v44t_(v[r][xv])};m[no]=z;if(/^SAFE_/.test(z.verdict))safe[no]=1;}return{exists:true,map:m,safe:safe};}
function i79v44old22_(cur,hist,v2){
 var rows=[],stats={},safe={total:0},remain={total:0};I79V44_OLD22.forEach(function(raw){var no=i79v44order_(raw,raw),c=cur.map[no]||null,h=hist.map[no]||null,st=c?c.status:'MISSING',bucket=(st==='MATCHED'||st==='NON_CARD'||st==='NO_MATCH'||st==='AMBIGUOUS')?st:(st==='MISSING'?'MISSING':'OTHER'),isSafe=!!v2.safe[no],b=isSafe?safe:remain;stats[bucket]=(stats[bucket]||0)+1;b.total++;b[bucket]=(b[bucket]||0)+1;rows.push([no,isSafe?'Y':'',v2.map[no]?v2.map[no].verdict:'',h?h.status:'',st,h?h.purchase:'',c?c.purchase:'',c&&h?c.purchase-h.purchase:'',c?c.business:'',c?c.account:'',c?c.company:'',c?c.cardName:'',c?c.end4:'',c?c.approvalDate:'',c?c.approvalNo:'',c?c.reason:'']);});return{rows:rows,stats:stats,safe:safe,remain:remain};
}
function i79v44writeCv_(ss,rows){var sh=ss.getSheetByName('ISSUE79_현재VAT주문대조')||ss.insertSheet('ISSUE79_현재VAT주문대조');i79v44write_(sh,['구분','주문번호','현재상태','사업자등록번호','쿠팡계정ID','카드매입','VAT매입','차이','사업자불일치','계정불일치'],rows,2);}
function i79v44writeHistory_(ss,rows){var sh=ss.getSheetByName('ISSUE79_현재vsIssue74변경')||ss.insertSheet('ISSUE79_현재vsIssue74변경');i79v44write_(sh,['구분','주문번호','Issue74상태','현재상태','Issue74매입','현재매입','차이','현재사업자','현재계정','현재매칭근거'],rows,2);}
function i79v44writeOld_(ss,rows){var sh=ss.getSheetByName('ISSUE79_기존22현재상태')||ss.insertSheet('ISSUE79_기존22현재상태');i79v44write_(sh,['주문번호','v2_SAFE','v2판정','Issue74상태','현재상태','Issue74매입','현재매입','차이','현재사업자','현재계정','구매카드사','구매카드명','끝4','승인일','승인번호','현재매칭근거'],rows,1);}
function i79v44write_(sh,h,rows,textCol){if(sh.getFilter())sh.getFilter().remove();sh.clear();sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length){sh.getRange(2,textCol,rows.length,1).setNumberFormat('@');sh.getRange(2,1,rows.length,h.length).setValues(rows);}sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}sh.autoResizeColumns(1,h.length);}
function i79v44status_(ss,pairs){var sh=ss.getSheetByName('ISSUE79_현재상태대조상태')||ss.insertSheet('ISSUE79_현재상태대조상태');sh.clearContents();var a=[['항목','값']].concat(pairs||[]);sh.getRange(1,1,a.length,2).setValues(a);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}
function i79v44sig_(sh){var rg=sh.getDataRange(),s=sh.getName()+'|'+rg.getNumRows()+'|'+rg.getNumColumns()+'|'+JSON.stringify(rg.getDisplayValues());return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,s,Utilities.Charset.UTF_8));}
function i79v44header_(v,need){for(var r=0;r<Math.min(50,v.length);r++){var ok=true;for(var j=0;j<need.length;j++)if(i79v44ix_(v[r],need[j]==='주문번호'?['주문번호','마켓주문번호']:[need[j]])<0){ok=false;break;}if(ok)return r;}throw new Error('header 탐지 실패: '+need.join(','));}
function i79v44ix_(h,a){for(var i=0;i<h.length;i++){var x=i79v44c_(h[i]);for(var j=0;j<a.length;j++)if(x===i79v44c_(a[j]))return i;}return-1;}
function i79v44order_(raw,disp){var d=i79v44t_(disp).replace(/\s+/g,'');if(/^\d+$/.test(d))return d;if(typeof raw==='number'&&isFinite(raw))return String(Math.round(raw));var s=i79v44t_(raw).replace(/\s+/g,'');if(/^\d+$/.test(s))return s;if(/^\d+(?:\.\d+)?e[+\-]?\d+$/i.test(s)){var n=Number(s);if(isFinite(n))return String(Math.round(n));}return s.replace(/\.0+$/,'');}
function i79v44t_(v){return String(v==null?'':v).trim();}
function i79v44c_(v){return i79v44t_(v).toLowerCase().replace(/\s+/g,'');}
function i79v44n_(v){if(typeof v==='number')return isFinite(v)?Math.round(v):0;var s=i79v44t_(v).replace(/,/g,'').replace(/[^0-9.\-]/g,'');var n=Number(s);return isFinite(n)?Math.round(n):0;}
function i79v44biz_(v){var d=i79v44t_(v).replace(/\D/g,'');return d||i79v44c_(v);}
function i79v44h1_(v){var s=i79v44c_(v);return s==='상반기'||s==='1h'||s==='h1'||s==='1'||s==='1반기';}
