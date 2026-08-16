var LOTTEON_REMOTE_TASK={id:'ISSUE79-V46-ORDER-COLUMN-MATRIX',title:'Issue79 주문식별자 컬럼 전수 교차진단',enabled:true,statusSheet:'ISSUE79_키컬럼진단상태'};
var I79V46='v4.6-ISSUE79-ORDER-COLUMN-MATRIX-READONLY';
var I79V46_CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
var I79V46_OLD22=['30100189414967','2100191293730','23100192054997','29100191903703','5100191540477','26100192791914','17100196061992','18100195889436','12100196902987','12100197310918','11100197821559','5100198099688','19100198942446','5101102615736','6101100531065','12101138407306','12101154659935','17100199506890','19100199527230','29100199660972','16101220120574','9101251100578'];
function runLotteonRemoteTaskStartRemote_(){return i79v46run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v46run_();}
function i79v46run_(){
 var ss=SpreadsheetApp.getActive(),before={},started=new Date().toISOString();
 try{
  i79v46status_(ss,[['version',I79V46],['상태','RUNNING'],['단계','READ'],['메시지','주문번호 계열 컬럼 전수 교차진단 중'],['실행시작',started]]);
  I79V46_CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락: '+n);before[n]=i79v46sig_(sh);});
  var card=i79v46sheet_(ss.getSheetByName('부가세_카드매칭검증'),'CARD','CARD');
  var vat=i79v46sheet_(ss.getSheetByName('부가세_신고자료'),'VAT','VAT');
  var histSh=ss.getSheetByName('ISSUE74_카드회수PREVIEW');if(!histSh)throw new Error('ISSUE74_카드회수PREVIEW 누락');
  var hist=i79v46sheet_(histSh,'I74','CARD');
  var source=i79v46sheet_(ss.getSheetByName('매출데이터_붙여넣기'),'SOURCE','SOURCE');
  var matrix=[];
  var cv=i79v46pairs_(card,vat,matrix),ch=i79v46pairs_(card,hist,matrix),sv=i79v46pairs_(source,vat,matrix),sc=i79v46pairs_(source,card,matrix);
  var old=[];[card,vat,hist,source].forEach(function(s){s.cols.forEach(function(c){old.push([s.label,c.index+1,c.header,c.raw.size,c.disp.size,i79v46oldHit_(c.raw),i79v46oldHit_(c.disp),c.samples.join(' | ')]);});});
  i79v46writeMatrix_(ss,matrix);i79v46writeCols_(ss,old);
  var changed=[];I79V46_CORE.forEach(function(n){if(i79v46sig_(ss.getSheetByName(n))!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반: '+changed.join(','));
  i79v46status_(ss,[
   ['version',I79V46],['상태','PASS'],['단계','DONE'],['메시지','주문번호 계열 컬럼 raw/display 전수 교차진단 완료'],
   ['CARD_후보컬럼수',card.cols.length],['VAT_후보컬럼수',vat.cols.length],['Issue74_후보컬럼수',hist.cols.length],['SOURCE_후보컬럼수',source.cols.length],
   ['CARD_대상행',card.rowCount],['VAT_대상행',vat.rowCount],['Issue74_대상행',hist.rowCount],['SOURCE_활성대상행',source.rowCount],
   ['CARD↔VAT_raw최대overlap',cv.raw.overlap],['CARD↔VAT_raw최대조합',cv.raw.combo],['CARD↔VAT_display최대overlap',cv.disp.overlap],['CARD↔VAT_display최대조합',cv.disp.combo],
   ['CARD↔Issue74_raw최대overlap',ch.raw.overlap],['CARD↔Issue74_raw최대조합',ch.raw.combo],['CARD↔Issue74_display최대overlap',ch.disp.overlap],['CARD↔Issue74_display최대조합',ch.disp.combo],
   ['SOURCE↔VAT_raw최대overlap',sv.raw.overlap],['SOURCE↔VAT_raw최대조합',sv.raw.combo],['SOURCE↔VAT_display최대overlap',sv.disp.overlap],['SOURCE↔VAT_display최대조합',sv.disp.combo],
   ['SOURCE↔CARD_raw최대overlap',sc.raw.overlap],['SOURCE↔CARD_raw최대조합',sc.raw.combo],['SOURCE↔CARD_display최대overlap',sc.disp.overlap],['SOURCE↔CARD_display최대조합',sc.disp.combo],
   ['CARD_OLD22최대hit',i79v46bestOld_(card)],['VAT_OLD22최대hit',i79v46bestOld_(vat)],['Issue74_OLD22최대hit',i79v46bestOld_(hist)],['SOURCE_OLD22최대hit',i79v46bestOld_(source)],
   ['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
  ]);
  return{ok:true,done:true,version:I79V46,cardVatRaw:cv.raw.overlap,cardVatDisplay:cv.disp.overlap,sourceVatRaw:sv.raw.overlap};
 }catch(e){var msg=String(e&&e.message?e.message:e);try{i79v46status_(ss,[['version',I79V46],['상태','ERROR'],['단계','FAILED'],['메시지','주문식별자 컬럼 진단 실패'],['실행시작',started],['오류',msg],['완료시각',new Date().toISOString()]]);}catch(_e){}throw e;}
}
function i79v46sheet_(sh,label,kind){
 var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=i79v46headerRow_(v,kind),h=v[hr]||[],cols=[];
 for(var c=0;c<h.length;c++){if(i79v46isOrderHeader_(h[c]))cols.push({index:c,header:i79v46t_(h[c]),raw:i79v46setNew_(),disp:i79v46setNew_(),samples:[]});}
 if(!cols.length)throw new Error(label+' 주문번호 계열 header 없음');
 var ixYear=i79v46ix_(h,'신고연도'),ixHalf=i79v46ix_(h,'반기'),ixDate=i79v46ixAny_(h,['마켓주문일자','주문일자','결제일자','주문일시']),ixStatus=i79v46ixAny_(h,['마켓주문상태','주문상태','상태','클레임상태','처리상태']),ixSales=i79v46ixAny_(h,['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액']);
 var count=0;
 for(var r=hr+1;r<v.length;r++){
  if(kind==='CARD'||kind==='VAT'){
   if(ixYear>=0&&i79v46t_(v[r][ixYear])!=='2026')continue;if(ixHalf>=0&&i79v46t_(v[r][ixHalf])!=='상반기')continue;
  }else if(kind==='SOURCE'){
   if(ixDate<0)throw new Error('SOURCE 날짜 header 없음');var dt=i79v46date_(v[r][ixDate]);if(!dt||dt<'2026-01-01'||dt>'2026-06-30')continue;var st=ixStatus>=0?i79v46t_(v[r][ixStatus]):'';if(/취소|반품|환불/.test(st))continue;if(ixSales>=0&&!i79v46n_(v[r][ixSales]))continue;
  }
  var any=false;cols.forEach(function(o){var raw=i79v46id_(v[r][o.index]),disp=i79v46id_(d[r][o.index]);if(raw){o.raw.add(raw);any=true;if(o.samples.length<5)o.samples.push(raw);}if(disp)o.disp.add(disp);});if(any)count++;
 }
 return{label:label,kind:kind,cols:cols,rowCount:count};
}
function i79v46pairs_(a,b,out){var bestR={overlap:-1,combo:''},bestD={overlap:-1,combo:''};a.cols.forEach(function(x){b.cols.forEach(function(y){var ro=i79v46overlap_(x.raw,y.raw),dd=i79v46overlap_(x.disp,y.disp),combo=a.label+'.'+x.header+' ↔ '+b.label+'.'+y.header;out.push([a.label,x.index+1,x.header,b.label,y.index+1,y.header,x.raw.size,y.raw.size,ro,x.disp.size,y.disp.size,dd]);if(ro>bestR.overlap){bestR={overlap:ro,combo:combo};}if(dd>bestD.overlap){bestD={overlap:dd,combo:combo};}});});return{raw:bestR,disp:bestD};}
function i79v46oldHit_(s){var n=0;I79V46_OLD22.forEach(function(x){if(s.has(x))n++;});return n;}
function i79v46bestOld_(s){var best={n:-1,text:''};s.cols.forEach(function(c){var r=i79v46oldHit_(c.raw),d=i79v46oldHit_(c.disp),n=Math.max(r,d),mode=r>=d?'raw':'display';if(n>best.n)best={n:n,text:n+' / '+c.header+' / '+mode};});return best.text;}
function i79v46writeMatrix_(ss,rows){var sh=ss.getSheetByName('ISSUE79_v46_키조합매트릭스')||ss.insertSheet('ISSUE79_v46_키조합매트릭스');i79v46write_(sh,['A시트','A열','Aheader','B시트','B열','Bheader','A_raw유니크','B_raw유니크','raw_overlap','A_display유니크','B_display유니크','display_overlap'],rows);}
function i79v46writeCols_(ss,rows){var sh=ss.getSheetByName('ISSUE79_v46_키컬럼목록')||ss.insertSheet('ISSUE79_v46_키컬럼목록');i79v46write_(sh,['시트','열','header','raw유니크','display유니크','OLD22_raw_hit','OLD22_display_hit','raw샘플'],rows);}
function i79v46write_(sh,h,rows){if(sh.getFilter())sh.getFilter().remove();sh.clear();sh.getRange(1,1,1,h.length).setValues([h]);if(rows.length)sh.getRange(2,1,rows.length,h.length).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,h.length).setFontWeight('bold');try{sh.getRange(1,1,rows.length+1,h.length).createFilter();}catch(_e){}sh.autoResizeColumns(1,h.length);}
function i79v46status_(ss,pairs){var sh=ss.getSheetByName('ISSUE79_키컬럼진단상태')||ss.insertSheet('ISSUE79_키컬럼진단상태');sh.clearContents();var a=[['항목','값']].concat(pairs||[]);sh.getRange(1,1,a.length,2).setValues(a);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setFontWeight('bold');sh.autoResizeColumns(1,2);}
function i79v46headerRow_(v,kind){for(var r=0;r<Math.min(50,v.length);r++){var h=v[r]||[],hasOrder=false;for(var c=0;c<h.length;c++)if(i79v46isOrderHeader_(h[c])){hasOrder=true;break;}if(!hasOrder)continue;if(kind==='CARD'&&i79v46ix_(h,'카드매칭상태')<0)continue;if(kind==='VAT'&&i79v46ix_(h,'신고연도')<0)continue;return r;}throw new Error(kind+' header row 탐지 실패');}
function i79v46isOrderHeader_(v){var s=i79v46c_(v);return /주문/.test(s)&&(/번호/.test(s)||/id/.test(s));}
function i79v46ix_(h,name){var n=i79v46c_(name);for(var i=0;i<h.length;i++)if(i79v46c_(h[i])===n)return i;return-1;}
function i79v46ixAny_(h,names){for(var j=0;j<names.length;j++){var x=i79v46ix_(h,names[j]);if(x>=0)return x;}return-1;}
function i79v46id_(v){var s=i79v46t_(v).replace(/\s+/g,'');if(!s)return'';return s.replace(/\.0+$/,'');}
function i79v46t_(v){return String(v==null?'':v).trim();}
function i79v46c_(v){return i79v46t_(v).toLowerCase().replace(/\s+/g,'');}
function i79v46n_(v){if(typeof v==='number')return isFinite(v)?Math.round(v):0;var s=i79v46t_(v).replace(/,/g,'').replace(/[^0-9.\-]/g,'');var n=Number(s);return isFinite(n)?Math.round(n):0;}
function i79v46date_(v){if(Object.prototype.toString.call(v)==='[object Date]'&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i79v46t_(v),m=s.match(/(20\d{2})\D?(\d{1,2})\D?(\d{1,2})/);if(!m)return'';return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);}
function i79v46overlap_(a,b){var n=0;a.each(function(x){if(b.has(x))n++;});return n;}
function i79v46setNew_(){var o={},n=0;return{add:function(x){if(!Object.prototype.hasOwnProperty.call(o,x)){o[x]=1;n++;}},has:function(x){return Object.prototype.hasOwnProperty.call(o,x);},each:function(fn){Object.keys(o).forEach(fn);},get size(){return n;}};}
function i79v46sig_(sh){var rg=sh.getDataRange(),s=sh.getName()+'|'+rg.getNumRows()+'|'+rg.getNumColumns()+'|'+JSON.stringify(rg.getDisplayValues());return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.MD5,s,Utilities.Charset.UTF_8));}
