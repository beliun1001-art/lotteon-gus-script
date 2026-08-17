var LOTTEON_REMOTE_TASK={id:'ISSUE79-V54-SOURCE-BLOCK-BOUNDARY',title:'Issue79 SOURCE 1970행 블록복제 경계/충돌 검증 READ-ONLY',enabled:true,statusSheet:'ISSUE79_SOURCE블록_진단상태'};
var I79V54='v5.4-ISSUE79-SOURCE-BLOCK-BOUNDARY-COLLISION-READONLY';
var I79V54_CORE=['매출데이터_붙여넣기','부가세_신고자료','부가세_카드매칭검증','부가세_기간별','카드사용내역_붙여넣기','카드_마스터'];
var I79V54_B1S=87,I79V54_B1E=2056,I79V54_B2S=4695,I79V54_B2E=6664,I79V54_OFFSET=4608;
function runLotteonRemoteTaskStartRemote_(){return i79v54run_();}
function runLotteonRemoteTaskContinueRemote_(){return i79v54run_();}
function i79v54run_(){
 var ss=SpreadsheetApp.getActive(),before={},started=new Date().toISOString();
 try{
  i79v54status_(ss,[['version',I79V54],['상태','RUNNING'],['단계','BLOCK_BOUNDARY'],['메시지','SOURCE 1970행 블록복제 경계와 제거 후 H1/CARD 차이 READ-ONLY 검증 중'],['실행시작',started]]);
  I79V54_CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh)throw new Error('핵심 시트 누락: '+n);before[n]=i79v54sig_(sh);});
  var sourceSh=ss.getSheetByName('매출데이터_붙여넣기');
  var src=i79v54source_(sourceSh);
  if(src.values.length<I79V54_B2E)throw new Error('SOURCE 행 부족: '+src.values.length);
  var block=i79v54block_(src);
  if(block.pairCount!==1970||block.exactPairs!==1970||block.mismatchPairs!==0)throw new Error('1970행 exact block guard 실패 '+JSON.stringify(block));
  var raw=i79v54h1_(src,false),primary=i79v54h1_(src,true);
  var primUid=i79v54uid_(primary.rows,src);
  var card=i79v54card_(ss.getSheetByName('부가세_카드매칭검증'));
  if(card.rows.length!==1355||(card.stats.MATCHED||0)!==842||(card.stats.NON_CARD||0)!==509||(card.stats.NO_MATCH||0)!==4||(card.stats.AMBIGUOUS||0)!==0||card.purchase!==105314779)throw new Error('현재 CARD baseline 불일치 '+JSON.stringify({n:card.rows.length,s:card.stats,p:card.purchase}));
  var cmp=i79v54compare_(primary,card);
  i79v54writeBlock_(ss,block,src);
  i79v54writeCollision_(ss,primUid.detail);
  i79v54writeCardDiff_(ss,cmp.diffRows);
  var changed=[];I79V54_CORE.forEach(function(n){var sh=ss.getSheetByName(n);if(!sh||i79v54sig_(sh)!==before[n])changed.push(n);});if(changed.length)throw new Error('READ-ONLY 위반: '+changed.join(', '));
  i79v54status_(ss,[
   ['version',I79V54],['상태','PASS'],['단계','DONE'],['메시지','SOURCE 87~2056 = 4695~6664 exact 블록복제와 제거 후 H1/CARD 영향 검증 완료'],
   ['SOURCE전체행',src.values.length],['블록1',I79V54_B1S+'~'+I79V54_B1E],['블록2',I79V54_B2S+'~'+I79V54_B2E],['블록offset',I79V54_OFFSET],['블록pair수',block.pairCount],['블록exactPair',block.exactPairs],['블록mismatchPair',block.mismatchPairs],
   ['RAW_H1활성행',raw.rows.length],['RAW_H1활성주문',raw.orderCount],['RAW_H1매입합',raw.purchase],
   ['복제블록_H1활성행',raw.rows.length-primary.rows.length],['복제블록_H1매입합',raw.purchase-primary.purchase],
   ['PRIMARY_H1활성행',primary.rows.length],['PRIMARY_H1활성주문',primary.orderCount],['PRIMARY_H1매입합',primary.purchase],
   ['PRIMARY_UID존재행',primUid.uidRows],['PRIMARY_UID공란행',primUid.blankRows],['PRIMARY_UID중복그룹',primUid.multiGroups],['PRIMARY_UID_exact중복그룹',primUid.exactGroups],['PRIMARY_UID_충돌그룹',primUid.conflictGroups],['PRIMARY_UID_중복추가행',primUid.extraRows],
   ['CARD_H1주문',card.rows.length],['CARD_H1매입합',card.purchase],['PRIMARY_CARD_overlap',cmp.overlap],['PRIMARY_only',cmp.primaryOnly.length],['CARD_only',cmp.cardOnly.length],['공통주문_매입차이건',cmp.diffCount],['CARD-PRIMARY_매입차이합',cmp.delta],
   ['핵심시트변경수',0],['오류',''],['완료시각',new Date().toISOString()]
  ]);
  return{ok:true,done:true,version:I79V54,blockExact:block.exactPairs,primaryRows:primary.rows.length,primaryPurchase:primary.purchase,uidConflicts:primUid.conflictGroups,cardDiff:cmp.diffCount};
 }catch(e){var m=String(e&&e.stack?e.stack:(e&&e.message?e.message:e));try{i79v54status_(ss,[['version',I79V54],['상태','ERROR'],['단계','FAILED'],['메시지','Issue79 v5.4 SOURCE 블록 경계 검증 실패'],['오류',m],['완료시각',new Date().toISOString()]]);}catch(_e){}throw e;}
}
function i79v54source_(sh){var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),h=v[0]||[],x={date:i79v54ix_(h,['마켓주문일자']),id:i79v54ix_(h,['마켓주문번호']),acct:i79v54ix_(h,['마켓아이디']),mp:i79v54ix_(h,['마켓상품번호']),sales:i79v54ix_(h,['결제금액합계(원)']),mst:i79v54ix_(h,['마켓주문상태']),qty:i79v54ix_(h,['결제수량']),o1:i79v54ix_(h,['옵션1']),o2:i79v54ix_(h,['옵션2']),sp:i79v54ix_(h,['사이트상품번호']),so:i79v54ix_(h,['사이트주문번호']),pur:i79v54ix_(h,['구매가격']),tst:i79v54ix_(h,['더망고주문상태']),uid:i79v54ix_(h,['더망고주문고유번호']),settle:i79v54ix_(h,['정산예정금액(원)'])};['date','id','sales','mst','pur','uid'].forEach(function(k){if(x[k]<0)throw new Error('SOURCE header 누락 '+k);});return{values:v,display:d,headers:h,index:x};}
function i79v54block_(src){var exact=0,mis=0,mismatch=[];for(var r1=I79V54_B1S;r1<=I79V54_B1E;r1++){var r2=r1+I79V54_OFFSET,a=src.values[r1-1],b=src.values[r2-1],diff=[];for(var c=0;c<src.headers.length;c++){if(i79v54cell_(a[c])!==i79v54cell_(b[c]))diff.push(src.headers[c]||('C'+(c+1)));}if(diff.length){mis++;if(mismatch.length<50)mismatch.push([r1,r2,diff.join('|')]);}else exact++;}return{pairCount:I79V54_B1E-I79V54_B1S+1,exactPairs:exact,mismatchPairs:mis,mismatch:mismatch};}
function i79v54h1_(src,excludeTail){var v=src.values,d=src.display,x=src.index,rows=[],orders={},sum=0;for(var r=1;r<v.length;r++){var sheetRow=r+1;if(excludeTail&&sheetRow>=I79V54_B2S&&sheetRow<=I79V54_B2E)continue;var dt=i79v54date_(v[r][x.date]);if(dt<'2026-01-01'||dt>'2026-06-30')continue;var sales=i79v54n_(v[r][x.sales]);if(!sales)continue;var mst=i79v54t_(v[r][x.mst]);if(/취소|반품|환불/.test(mst))continue;var id=i79v54id_(d[r][x.id]);if(!id)continue;var o={row:sheetRow,date:dt,id:id,acct:x.acct>=0?i79v54t_(v[r][x.acct]):'',mp:x.mp>=0?i79v54id_(d[r][x.mp]):'',sales:sales,mst:mst,qty:x.qty>=0?i79v54n_(v[r][x.qty]):0,o1:x.o1>=0?i79v54t_(v[r][x.o1]):'',o2:x.o2>=0?i79v54t_(v[r][x.o2]):'',sp:x.sp>=0?i79v54t_(v[r][x.sp]):'',so:x.so>=0?i79v54id_(d[r][x.so]):'',pur:i79v54n_(v[r][x.pur]),tst:x.tst>=0?i79v54t_(v[r][x.tst]):'',uid:i79v54id_(d[r][x.uid]),settle:x.settle>=0?i79v54n_(v[r][x.settle]):0,raw:v[r]};rows.push(o);orders[id]=1;sum+=o.pur;}var agg={};rows.forEach(function(o){if(!agg[o.id])agg[o.id]={id:o.id,date:o.date,rows:0,pur:0};agg[o.id].rows++;agg[o.id].pur+=o.pur;});return{rows:rows,orderCount:Object.keys(orders).length,purchase:Math.round(sum),agg:agg};}
function i79v54uid_(rows,src){var g={},uidRows=0,blank=0;(rows||[]).forEach(function(o){if(!o.uid){blank++;return;}uidRows++;if(!g[o.uid])g[o.uid]=[];g[o.uid].push(o);});var multi=0,exact=0,conf=0,extra=0,detail=[];Object.keys(g).forEach(function(uid){var a=g[uid];if(a.length<2)return;multi++;extra+=a.length-1;var fps={};a.forEach(function(o){fps[i79v54rowkey_(o.raw)]=1;});var typ=Object.keys(fps).length===1?'EXACT_DUP':'CONFLICT';if(typ==='EXACT_DUP')exact++;else conf++;var dif=i79v54diffCols_(a.map(function(o){return o.raw;}),src.headers);a.forEach(function(o){detail.push([uid,typ,a.length,o.row,o.id,o.date,o.pur,o.sales,o.qty,o.so,o.mp,o.o1,o.mst,o.tst,o.settle,dif.join('|')]);});});detail.sort(function(a,b){if(a[1]!==b[1])return a[1]==='CONFLICT'?-1:1;return Number(a[0])-Number(b[0]);});return{uidRows:uidRows,blankRows:blank,multiGroups:multi,exactGroups:exact,conflictGroups:conf,extraRows:extra,detail:detail};}
function i79v54card_(sh){var v=sh.getDataRange().getValues(),d=sh.getDataRange().getDisplayValues(),hr=i79v54hdr_(v,['주문번호','카드매칭상태']);if(hr<0)throw new Error('CARD header 실패');var h=v[hr],x={y:i79v54ix_(h,['신고연도']),hf:i79v54ix_(h,['반기']),id:i79v54ix_(h,['주문번호']),pur:i79v54ix_(h,['주문매입금액','매입금액']),st:i79v54ix_(h,['카드매칭상태'])};var rows=[],map={},stats={},sum=0;for(var r=hr+1;r<v.length;r++){if(i79v54t_(v[r][x.y])!=='2026'||i79v54t_(v[r][x.hf])!=='상반기')continue;var id=i79v54id_(d[r][x.id]);if(!id)continue;var st=i79v54t_(v[r][x.st]).toUpperCase(),p=i79v54n_(v[r][x.pur]),o={id:id,pur:p,st:st};rows.push(o);map[id]=o;stats[st]=(stats[st]||0)+1;sum+=p;}return{rows:rows,map:map,stats:stats,purchase:Math.round(sum)};}
function i79v54compare_(primary,card){var pmap=primary.agg,cmap=card.map,po=[],co=[],over=0,diff=[],delta=0,diffCount=0;Object.keys(pmap).forEach(function(id){if(cmap[id]){over++;var pp=Math.round(pmap[id].pur),cp=Math.round(cmap[id].pur),dd=cp-pp;if(dd!==0){diffCount++;delta+=dd;diff.push([id,pmap[id].date,pmap[id].rows,pp,cp,dd,cmap[id].st]);}}else po.push(id);});Object.keys(cmap).forEach(function(id){if(!pmap[id])co.push(id);});diff.sort(function(a,b){return Math.abs(b[5])-Math.abs(a[5]);});return{overlap:over,primaryOnly:po,cardOnly:co,diffCount:diffCount,delta:Math.round(delta),diffRows:diff};}
function i79v54writeBlock_(ss,b,src){var sh=i79v54sheet_(ss,'ISSUE79_SOURCE블록_대조');var out=[['원본행','복제행','결과','차이컬럼']];for(var r=I79V54_B1S;r<=I79V54_B1E;r++){if(r===I79V54_B1S||r===I79V54_B1E||r===87||r===2056||r%250===0)out.push([r,r+I79V54_OFFSET,'EXACT','']);}b.mismatch.forEach(function(x){out.push([x[0],x[1],'MISMATCH',x[2]]);});i79v54table_(sh,out);}
function i79v54writeCollision_(ss,rows){var sh=i79v54sheet_(ss,'ISSUE79_SOURCE블록_충돌UID');var out=[['UID','분류','그룹행수','원본행','주문번호','주문일','구매가격','결제금액','수량','사이트주문번호','마켓상품번호','옵션1','마켓상태','더망고상태','정산예정금액','차이컬럼']].concat(rows);i79v54table_(sh,out);}
function i79v54writeCardDiff_(ss,rows){var sh=i79v54sheet_(ss,'ISSUE79_SOURCE블록_CARD차이');var out=[['주문번호','주문일','PRIMARY상세행','PRIMARY매입','현재CARD매입','CARD-PRIMARY','현재상태']].concat(rows);i79v54table_(sh,out);}
function i79v54status_(ss,rows){var sh=i79v54sheet_(ss,'ISSUE79_SOURCE블록_진단상태');i79v54table_(sh,[['항목','값']].concat(rows));}
function i79v54sheet_(ss,n){return ss.getSheetByName(n)||ss.insertSheet(n);}
function i79v54table_(sh,rows){sh.clearContents();if(rows.length&&rows[0].length)sh.getRange(1,1,rows.length,rows[0].length).setValues(rows);sh.setFrozenRows(1);}
function i79v54diffCols_(rows,h){if(rows.length<2)return[];var out=[];for(var c=0;c<h.length;c++){var z={};rows.forEach(function(r){z[i79v54cell_(r[c])]=1;});if(Object.keys(z).length>1)out.push(h[c]||('C'+(c+1)));}return out;}
function i79v54rowkey_(r){var a=[];for(var c=0;c<r.length;c++)a.push(i79v54cell_(r[c]));return a.join('\u001f');}
function i79v54cell_(x){if(x===null||x===undefined)return'';if(Object.prototype.toString.call(x)==='[object Date]')return String(x.getTime());if(typeof x==='number')return String(Math.round(x*1000000)/1000000);return String(x).replace(/\r\n/g,'\n').trim();}
function i79v54sig_(sh){var v=sh.getDataRange().getDisplayValues(),h=2166136261>>>0;for(var r=0;r<v.length;r++){for(var c=0;c<v[r].length;c++){var s=String(v[r][c])+'\u001f';for(var i=0;i<s.length;i++){h^=s.charCodeAt(i);h=Math.imul(h,16777619)>>>0;}}}return v.length+'x'+(v[0]?v[0].length:0)+':'+h;}
function i79v54hdr_(v,need){for(var r=0;r<Math.min(v.length,20);r++){var h=(v[r]||[]).map(i79v54t_),ok=true;for(var i=0;i<need.length;i++)if(h.indexOf(need[i])<0){ok=false;break;}if(ok)return r;}return-1;}
function i79v54ix_(h,names){var a=(h||[]).map(function(x){return i79v54t_(x).replace(/\s/g,'');});for(var i=0;i<names.length;i++){var k=String(names[i]).replace(/\s/g,'');var p=a.indexOf(k);if(p>=0)return p;}return-1;}
function i79v54t_(x){return String(x===null||x===undefined?'':x).trim();}
function i79v54id_(x){return i79v54t_(x).replace(/,/g,'').replace(/\.0+$/,'').replace(/\s/g,'');}
function i79v54n_(x){if(typeof x==='number')return isFinite(x)?x:0;var n=Number(i79v54t_(x).replace(/,/g,'').replace(/원/g,''));return isFinite(n)?n:0;}
function i79v54date_(x){if(!x)return'';if(Object.prototype.toString.call(x)==='[object Date]')return Utilities.formatDate(x,Session.getScriptTimeZone()||'Asia/Seoul','yyyy-MM-dd');var s=i79v54t_(x).replace(/\./g,'-').replace(/\//g,'-');var m=s.match(/(20\d{2})-(\d{1,2})-(\d{1,2})/);return m?m[1]+'-'+('0'+m[2]).slice(-2)+'-'+('0'+m[3]).slice(-2):'';}
