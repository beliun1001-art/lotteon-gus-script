var LOTTEON_REMOTE_TASK={id:'ISSUE79-V542-SOURCE-UID-COLLAPSE',title:'Issue79 SOURCE UID 단일버전 재구성 READ-ONLY',enabled:true,statusSheet:'ISSUE79_SOURCEUID_진단상태'};
var I79V542='v5.4.2-ISSUE79-SOURCE-UID-SINGLE-VERSION-READONLY';
var I79V542_CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
function runLotteonRemoteTaskStartRemote_(){return i79v542run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v542run_();}
function i79v542run_(){
  var ss=SpreadsheetApp.getActive(),before={},started=new Date().toISOString();
  try{
    i79v542status_(ss,[['version',I79V542],['상태','RUNNING'],['단계','UID_COLLAPSE'],['메시지','동일 UID의 복제/후속갱신을 하나의 최신 논리행으로 재구성 중'],['실행시작',started]]);
    I79V542_CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락: '+n);before[n]=i79v542sig_(sh);});
    var card=i79v542card_(ss.getSheetByName('부가세_카드매칭검증'));
    if(card.rows.length!==1355||(card.stats.MATCHED||0)!==842||(card.stats.NON_CARD||0)!==509||(card.stats.NO_MATCH||0)!==4||(card.stats.AMBIGUOUS||0)!==0||card.purchase!==105314779)throw new Error('현재 CARD baseline 불일치 '+JSON.stringify({n:card.rows.length,s:card.stats,p:card.purchase}));
    var src=i79v542source_(ss.getSheetByName('매출데이터_붙여넣기'));
    var r=i79v542collapse_(src.rows);
    if(r.crossOrderUidGroups!==0)throw new Error('UID가 서로 다른 주문번호에 재사용된 그룹 존재: '+r.crossOrderUidGroups);
    var cmp=i79v542compare_(r.active,card);
    var toss=i79v542toss_(r.active,card);
    i79v542write_(ss,'ISSUE79_SOURCEUID_버전변경',['UID','원본행수','첫행','최종행','주문번호','첫상태','최종상태','첫매입','최종매입','변경필드수'],r.versionChanges);
    i79v542write_(ss,'ISSUE79_SOURCEUID_CARD차이',['주문번호','재구성매입','현재CARD매입','차이','재구성상세행수','현재상태'],cmp.diffRows);
    i79v542write_(ss,'ISSUE79_SOURCEUID_TOSS3',['주문번호','재구성매입','현재CARD매입','차이','현재상태'],toss.rows);
    i79v542write_(ss,'ISSUE79_SOURCEUID_중복오프셋',['행오프셋','그룹수'],r.offsetRows);
    var changed=[];I79V542_CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh||i79v542sig_(sh)!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반: '+changed.join(', '));
    i79v542status_(ss,[
      ['version',I79V542],['상태','PASS'],['단계','DONE'],['메시지','동일 UID는 동일 주문의 버전으로 확인; 최신 행 1개로 H1 SOURCE 재구성 완료'],
      ['현재_CARD주문',card.rows.length],['현재_CARD_MATCHED',card.stats.MATCHED||0],['현재_CARD_NON_CARD',card.stats.NON_CARD||0],['현재_CARD_NO_MATCH',card.stats.NO_MATCH||0],['현재_CARD매입합',card.purchase],
      ['SOURCE_H1전체행',src.rows.length],['SOURCE_H1_UID공란행',r.blankUid],['SOURCE_H1_UID고유개수',r.uniqueUid],['SOURCE_H1_UID중복그룹',r.dupUidGroups],['SOURCE_H1_UID타주문재사용',r.crossOrderUidGroups],['SOURCE_H1_버전변경그룹',r.versionChanges.length],
      ['SOURCE_H1_오프셋4608그룹',r.offset4608],['SOURCE_H1_기타오프셋그룹',r.otherOffset],
      ['재구성_H1논리행',r.resolved.length],['재구성_H1활성행',r.active.length],['재구성_H1활성주문',cmp.orderCount],['재구성_H1활성매입합',cmp.purchase],
      ['CARD대비_매입차이주문',cmp.diffRows.length],['CARD대비_매입합차이',Math.round(card.purchase-cmp.purchase)],
      ['TOSS3_재구성매입합',toss.purchase],['TOSS3_현재CARD매입합',toss.cardPurchase],['TOSS3_차이',toss.cardPurchase-toss.purchase],
      ['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
    ]);
    return{ok:true,done:true,version:I79V542,logicalRows:r.resolved.length,activeRows:r.active.length,orders:cmp.orderCount,purchase:cmp.purchase,diffOrders:cmp.diffRows.length};
  }catch(e){var m=String(e&&e.stack?e.stack:(e&&e.message?e.message:e));try{i79v542status_(ss,[['version',I79V542],['상태','ERROR'],['단계','FAILED'],['메시지','Issue79 v5.4.2 SOURCE UID 재구성 실패'],['오류',m],['완료시각',new Date().toISOString()]]);}catch(_e){}throw e;}
}
function i79v542source_(sh){
  var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),h=v[0]||[];
  function ix(a){for(var i=0;i<a.length;i++){var k=a[i].replace(/\s/g,'');for(var j=0;j<h.length;j++)if(String(h[j]||'').replace(/\s/g,'')===k)return j;}return-1;}
  var x={date:ix(['마켓주문일자']),id:ix(['마켓주문번호']),acct:ix(['마켓아이디']),mp:ix(['마켓상품번호']),sales:ix(['결제금액합계(원)']),mst:ix(['마켓주문상태']),qty:ix(['결제수량']),o1:ix(['옵션1']),o2:ix(['옵션2']),sp:ix(['사이트상품번호']),so:ix(['사이트주문번호']),pur:ix(['구매가격']),tst:ix(['더망고주문상태']),uid:ix(['더망고주문고유번호']),memo:ix(['간단메모']),track:ix(['국내송장번호']),settle:ix(['정산예정금액(원)'])};
  ['date','id','sales','mst','pur','uid'].forEach(function(k){if(x[k]<0)throw new Error('SOURCE header 누락 '+k);});
  var rows=[];
  for(var r=1;r<v.length;r++){
    var dt=i79v542date_(v[r][x.date]);if(!/^2026-/.test(dt)||dt<'2026-01-01'||dt>'2026-06-30')continue;
    var id=i79v542idRaw_(v[r][x.id],d[r][x.id]);if(!id)continue;
    var uid=i79v542idRaw_(v[r][x.uid],d[r][x.uid]);
    rows.push({row:r+1,date:dt,id:id,acct:x.acct>=0?i79v542t_(v[r][x.acct]):'',mp:x.mp>=0?i79v542idRaw_(v[r][x.mp],d[r][x.mp]):'',sales:i79v542n_(v[r][x.sales]),mst:i79v542t_(v[r][x.mst]),qty:x.qty>=0?i79v542n_(v[r][x.qty]):0,o1:x.o1>=0?i79v542t_(v[r][x.o1]):'',o2:x.o2>=0?i79v542t_(v[r][x.o2]):'',sp:x.sp>=0?i79v542t_(v[r][x.sp]):'',so:x.so>=0?i79v542idRaw_(v[r][x.so],d[r][x.so]):'',pur:i79v542n_(v[r][x.pur]),tst:x.tst>=0?i79v542t_(v[r][x.tst]):'',uid:uid,memo:x.memo>=0?i79v542t_(v[r][x.memo]):'',track:x.track>=0?i79v542idRaw_(v[r][x.track],d[r][x.track]):'',settle:x.settle>=0?i79v542n_(v[r][x.settle]):0});
  }
  return{rows:rows};
}
function i79v542collapse_(rows){
  var g={},blank=0;(rows||[]).forEach(function(o){if(!o.uid){blank++;var kb='ROW|'+o.row;g[kb]=[o];return;}var k='UID|'+o.uid;if(!g[k])g[k]=[];g[k].push(o);});
  var resolved=[],dup=0,cross=0,changes=[],offs={},off4608=0,otherOff=0;
  Object.keys(g).forEach(function(k){var a=g[k].slice().sort(function(x,y){return x.row-y.row;});if(a.length>1)dup++;
    var ids=i79v542uniq_(a.map(function(z){return z.id;}));if(a[0].uid&&ids.length>1)cross++;
    var first=a[0],w=a[a.length-1];resolved.push(w);
    if(a.length>1){var delta=w.row-first.row;offs[delta]=(offs[delta]||0)+1;if(delta===4608)off4608++;else otherOff++;
      var fields=['id','acct','mp','sales','mst','qty','o1','o2','sp','so','pur','tst','memo','track','settle'],fc=0;fields.forEach(function(f){if(String(first[f])!==String(w[f]))fc++;});
      if(fc)changes.push([w.uid,a.length,first.row,w.row,w.id,first.mst+' / '+first.tst,w.mst+' / '+w.tst,Math.round(first.pur),Math.round(w.pur),fc]);
    }
  });
  var active=resolved.filter(function(o){return o.sales&&!/취소|반품|환불/.test((o.mst||'')+' '+(o.tst||''));});
  var offsetRows=Object.keys(offs).map(function(k){return[Number(k),offs[k]];}).sort(function(a,b){return b[1]-a[1]||a[0]-b[0];});
  return{resolved:resolved,active:active,blankUid:blank,uniqueUid:Object.keys(g).length,dupUidGroups:dup,crossOrderUidGroups:cross,versionChanges:changes,offsetRows:offsetRows,offset4608:off4608,otherOffset:otherOff};
}
function i79v542compare_(active,card){var m={},sum=0;(active||[]).forEach(function(o){if(!m[o.id])m[o.id]={p:0,n:0};m[o.id].p+=o.pur;m[o.id].n++;sum+=o.pur;});var dif=[];Object.keys(m).forEach(function(id){var cp=card.map[id]?card.map[id].pur:0;if(Math.round(m[id].p)!==Math.round(cp))dif.push([id,Math.round(m[id].p),Math.round(cp),Math.round(cp-m[id].p),m[id].n,card.map[id]?card.map[id].st:'MISSING_CARD']);});Object.keys(card.map).forEach(function(id){if(!m[id])dif.push([id,0,Math.round(card.map[id].pur),Math.round(card.map[id].pur),0,card.map[id].st]);});dif.sort(function(a,b){return Math.abs(b[3])-Math.abs(a[3]);});return{orderCount:Object.keys(m).length,purchase:Math.round(sum),diffRows:dif};}
function i79v542toss_(active,card){var targets={'12100196902987':1,'12100197310918':1,'5100198099688':1},m={},sum=0,cs=0;(active||[]).forEach(function(o){if(targets[o.id]){m[o.id]=(m[o.id]||0)+o.pur;sum+=o.pur;}});var rows=[];Object.keys(targets).forEach(function(id){var cp=card.map[id]?card.map[id].pur:0;cs+=cp;rows.push([id,Math.round(m[id]||0),Math.round(cp),Math.round(cp-(m[id]||0)),card.map[id]?card.map[id].st:'MISSING_CARD']);});return{rows:rows,purchase:Math.round(sum),cardPurchase:Math.round(cs)};}
function i79v542card_(sh){var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=-1;for(var r=0;r<Math.min(10,v.length);r++){var s=v[r].join('|');if(s.indexOf('주문번호')>=0&&s.indexOf('카드매칭상태')>=0){hr=r;break;}}if(hr<0)throw new Error('CARD header 실패');var h=v[hr];function ix(a){for(var i=0;i<a.length;i++){var k=a[i].replace(/\s/g,'');for(var j=0;j<h.length;j++)if(String(h[j]||'').replace(/\s/g,'')===k)return j;}return-1;}var x={y:ix(['신고연도']),hf:ix(['반기']),id:ix(['주문번호']),pur:ix(['주문매입금액','매입금액']),st:ix(['카드매칭상태'])},rows=[],map={},stats={},sum=0;for(var r=hr+1;r<v.length;r++){if(i79v542t_(v[r][x.y])!=='2026'||i79v542t_(v[r][x.hf])!=='상반기')continue;var id=i79v542idRaw_(v[r][x.id],d[r][x.id]);if(!id)continue;var st=i79v542t_(v[r][x.st]).toUpperCase(),p=i79v542n_(v[r][x.pur]),o={id:id,pur:p,st:st};rows.push(o);map[id]=o;stats[st]=(stats[st]||0)+1;sum+=p;}return{rows:rows,map:map,stats:stats,purchase:Math.round(sum)};}
function i79v542write_(ss,name,header,rows){var sh=ss.getSheetByName(name)||ss.insertSheet(name);sh.clearContents();sh.getRange(1,1,1,header.length).setValues([header]);if(rows&&rows.length)sh.getRange(2,1,rows.length,header.length).setValues(rows);sh.setFrozenRows(1);}
function i79v542status_(ss,rows){i79v542write_(ss,'ISSUE79_SOURCEUID_진단상태',['항목','값'],rows);}
function i79v542sig_(sh){var lr=sh.getLastRow(),lc=sh.getLastColumn(),a=lr&&lc?sh.getRange(1,1,Math.min(5,lr),Math.min(8,lc)).getDisplayValues():[],b=lr&&lc?sh.getRange(Math.max(1,lr-4),Math.max(1,lc-7),Math.min(5,lr),Math.min(8,lc)).getDisplayValues():[];return lr+'|'+lc+'|'+JSON.stringify(a)+'|'+JSON.stringify(b);}
function i79v542uniq_(a){var m={},o=[];(a||[]).forEach(function(x){x=String(x||'');if(!m[x]){m[x]=1;o.push(x);}});return o;}
function i79v542t_(v){return String(v==null?'':v).trim();}
function i79v542idRaw_(raw,display){if(typeof raw==='number'&&isFinite(raw)&&Math.abs(raw)<=Number.MAX_SAFE_INTEGER)return String(Math.trunc(raw));var s=String(display==null?'':display).replace(/,/g,'').replace(/\.0+$/,'').replace(/\s/g,'');if(/^[0-9]+$/.test(s))return s;var r=String(raw==null?'':raw).replace(/,/g,'').replace(/\.0+$/,'').replace(/\s/g,'');return r;}
function i79v542n_(v){if(typeof v==='number')return isFinite(v)?v:0;var n=Number(String(v==null?'':v).replace(/,/g,'').replace(/[^0-9.\-]/g,''));return isFinite(n)?n:0;}
function i79v542date_(v){if(v instanceof Date&&!isNaN(v.getTime()))return Utilities.formatDate(v,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=String(v==null?'':v).trim();var m=s.match(/(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);if(m)return m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2);if(typeof v==='number'&&v>30000){var d=new Date(Math.round((v-25569)*86400000));return Utilities.formatDate(d,'GMT','yyyy-MM-dd');}return'';}