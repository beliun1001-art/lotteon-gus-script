/** v6.71 Issue #34: remote-only safe sheet cleanup behind existing showOperationSheetsOnly menu. */
var LOTTEON_PATCH_V671_SAFE_SHEET_CLEANUP_LOADED = true;
var LOTTEON_SHEET_CLEANUP_VERSION_V671 = 'v6.71';
var LOTTEON_SHEET_CLEANUP_MANAGER_V671 = '시트정리_관리';
var LOTTEON_SHEET_CLEANUP_SMOKE_STATUS_V671 = 'PR35_운영스모크상태';
var LOTTEON_SHEET_CLEANUP_ACTIVE_STATES_V671 = {
  RUNNING:true, WRITING:true, FINALIZING:true, BATCH:true, APPLYING:true,
  ROLLBACK:true, ROLLBACK_PENDING:true, ROLLBACK_RUNNING:true
};
var LOTTEON_SHEET_CLEANUP_CORE_V671 = [
  '기준','대시보드','부가세_신고자료','부가세_기간별','부가세_상품별','부가세_카드매칭검증',
  '카드사용내역_붙여넣기','카드_마스터','매출데이터_붙여넣기',LOTTEON_SHEET_CLEANUP_MANAGER_V671
];

/** Existing local loader menu bridge calls this remote function. */
function showOperationSheetsOnly() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var manager = ss.getSheetByName(LOTTEON_SHEET_CLEANUP_MANAGER_V671);
  if (manager && lotteonCleanupSelectedCount_v671_(manager) > 0) {
    return deleteCheckedLotteonCleanupSheets_v671_();
  }
  return buildLotteonSheetCleanupList_v671_();
}

function buildLotteonSheetCleanupList_v671_(options) {
  options = options || {};
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var protection = lotteonSheetCleanupProtection_v671_(ss);
  var manager = ss.getSheetByName(LOTTEON_SHEET_CLEANUP_MANAGER_V671);
  if (!manager) manager = ss.insertSheet(LOTTEON_SHEET_CLEANUP_MANAGER_V671);

  var rows = ss.getSheets().filter(function(sheet) {
    return sheet.getName() !== LOTTEON_SHEET_CLEANUP_MANAGER_V671;
  }).map(function(sheet) {
    var name = sheet.getName();
    var decision = lotteonSheetCleanupDecision_v671_(name, protection);
    return [
      decision.recommended && !decision.protected,
      name,
      decision.category,
      decision.recommended ? '삭제 권장' : '직접 확인',
      decision.protected ? '보호' : '',
      decision.reason,
      Math.max(0, sheet.getLastRow()),
      Math.max(0, sheet.getLastColumn()),
      ''
    ];
  }).sort(function(a,b) {
    var ap=a[4]==='보호'?1:0, bp=b[4]==='보호'?1:0;
    if(ap!==bp)return ap-bp;
    var ar=a[3]==='삭제 권장'?0:1, br=b[3]==='삭제 권장'?0:1;
    if(ar!==br)return ar-br;
    return String(a[1]).localeCompare(String(b[1]),'ko');
  });

  manager.clear();
  var headers=['삭제 선택','시트명','분류','권장','보호','판단 사유','행수','열수','삭제 결과'];
  manager.getRange(1,1,1,headers.length).setValues([headers]);
  if(rows.length){
    manager.getRange(2,1,rows.length,headers.length).setValues(rows);
    manager.getRange(2,1,rows.length,1).insertCheckboxes();
    manager.getRange(2,7,rows.length,2).setNumberFormat('#,##0');
    var backgrounds=rows.map(function(row){
      if(row[4]==='보호')return new Array(headers.length).fill('#eeeeee');
      if(row[3]==='삭제 권장')return new Array(headers.length).fill('#fff2cc');
      return new Array(headers.length).fill('#ffffff');
    });
    manager.getRange(2,1,rows.length,headers.length).setBackgrounds(backgrounds);
  }
  manager.setFrozenRows(1);
  manager.getRange(1,1,1,headers.length).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
  [90,260,140,100,80,430,80,80,260].forEach(function(width,index){manager.setColumnWidth(index+1,width);});
  if(manager.getFilter())manager.getFilter().remove();
  if(rows.length)manager.getRange(1,1,rows.length+1,headers.length).createFilter();
  manager.activate();

  var active=Object.keys(protection.activePrefixes);
  if (!options.silent) {
    lotteonSheetCleanupAlert_v671_(
      '정리 대상 목록 생성 완료',
      '노란색은 기본 삭제 권장, 회색은 삭제 금지 시트입니다.\n' +
      '현재 실행 중 보호 작업: ' + (active.length?active.join(', '):'없음') + '\n\n' +
      '삭제할 시트만 체크한 뒤 같은 메뉴를 다시 누르세요.\n' +
      'LOTTEON 자동화 → ⑥ 운영 핵심 10개만 표시/불필요 삭제'
    );
  }
  return {ok:true,mode:'LIST',rows:rows.length,active:active};
}

function deleteCheckedLotteonCleanupSheets_v671_() {
  var ss=SpreadsheetApp.getActive();
  if(!ss)throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var manager=ss.getSheetByName(LOTTEON_SHEET_CLEANUP_MANAGER_V671);
  if(!manager||manager.getLastRow()<2)return buildLotteonSheetCleanupList_v671_();

  var ui=lotteonSheetCleanupUi_v671_();
  if(!ui)throw new Error('사용자 확인 UI가 없는 실행에서는 시트를 삭제할 수 없습니다.');

  var values=manager.getRange(2,1,manager.getLastRow()-1,9).getValues();
  var protection=lotteonSheetCleanupProtection_v671_(ss);
  var selected=[],blocked=[];
  values.forEach(function(row,index){
    if(row[0]!==true)return;
    var name=String(row[1]||'').trim(); if(!name)return;
    var decision=lotteonSheetCleanupDecision_v671_(name,protection);
    if(decision.protected){blocked.push({row:index+2,name:name,reason:decision.reason});return;}
    if(!ss.getSheetByName(name)){blocked.push({row:index+2,name:name,reason:'이미 존재하지 않는 시트'});return;}
    selected.push({row:index+2,name:name});
  });
  blocked.forEach(function(item){manager.getRange(item.row,9).setValue('차단: '+item.reason);manager.getRange(item.row,1).setValue(false);});
  if(!selected.length){
    ui.alert('삭제 가능한 체크 시트가 없습니다.','삭제할 시트를 체크하세요. 보호 시트는 자동 차단됩니다.',ui.ButtonSet.OK);
    return {ok:true,deleted:0,blocked:blocked.length};
  }
  if(ss.getSheets().length-selected.length<1)throw new Error('스프레드시트에는 최소 1개 시트가 남아 있어야 합니다.');
  var names=selected.map(function(x){return x.name;});
  var preview=names.slice(0,20).map(function(name){return '• '+name;}).join('\n');
  var more=names.length>20?'\n외 '+(names.length-20)+'개':'';
  var answer=ui.alert('시트 '+names.length+'개를 삭제할까요?',preview+more+'\n\n삭제 후에는 Google Sheets 버전 기록으로만 복구할 수 있습니다.',ui.ButtonSet.YES_NO);
  if(answer!==ui.Button.YES)return {ok:true,cancelled:true};

  var deleted=0,errors=[];
  selected.forEach(function(item){
    try{
      var refreshed=lotteonSheetCleanupDecision_v671_(item.name,lotteonSheetCleanupProtection_v671_(ss));
      if(refreshed.protected)throw new Error('삭제 직전 보호 상태로 변경됨: '+refreshed.reason);
      var sheet=ss.getSheetByName(item.name); if(!sheet)throw new Error('시트가 존재하지 않음');
      if(ss.getSheets().length<=1)throw new Error('마지막 시트는 삭제할 수 없음');
      ss.deleteSheet(sheet);
      manager.getRange(item.row,9).setValue('삭제 완료 '+new Date().toISOString());
      manager.getRange(item.row,1).setValue(false); deleted++;
    }catch(error){
      var message=String(error&&error.message?error.message:error);
      manager.getRange(item.row,9).setValue('실패: '+message); errors.push(item.name+': '+message);
    }
  });
  SpreadsheetApp.flush();
  ui.alert('시트 정리 완료','삭제 완료: '+deleted+'개\n차단: '+blocked.length+'개\n실패: '+errors.length+'개'+(errors.length?'\n\n'+errors.slice(0,10).join('\n'):''),ui.ButtonSet.OK);
  return {ok:errors.length===0,deleted:deleted,blocked:blocked.length,errors:errors};
}

function runSafeSheetCleanupSmoke_v671() {
  var ss=SpreadsheetApp.getActive();
  if(!ss)throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  var status=ss.getSheetByName(LOTTEON_SHEET_CLEANUP_SMOKE_STATUS_V671);
  if(!status)status=ss.insertSheet(LOTTEON_SHEET_CLEANUP_SMOKE_STATUS_V671);
  lotteonWriteSheetCleanupSmoke_v671_(status,'RUNNING','안전 시트 정리 비파괴 smoke 시작','');

  try{
    var before=ss.getSheets().map(function(sheet){return sheet.getName();}).sort();
    var result=buildLotteonSheetCleanupList_v671_({silent:true});
    var after=ss.getSheets().map(function(sheet){return sheet.getName();}).sort();
    var deleted=before.filter(function(name){return after.indexOf(name)<0;});
    if(deleted.length)throw new Error('smoke 중 시트 삭제 발생: '+deleted.join(', '));

    var manager=ss.getSheetByName(LOTTEON_SHEET_CLEANUP_MANAGER_V671);
    if(!manager||manager.getLastRow()<2)throw new Error('시트정리_관리 생성 또는 목록 작성 실패');
    var headers=manager.getRange(1,1,1,9).getDisplayValues()[0];
    var expected=['삭제 선택','시트명','분류','권장','보호','판단 사유','행수','열수','삭제 결과'];
    if(JSON.stringify(headers)!==JSON.stringify(expected))throw new Error('관리 시트 헤더 불일치');

    var protection=lotteonSheetCleanupProtection_v671_(ss);
    LOTTEON_SHEET_CLEANUP_CORE_V671.forEach(function(name){
      if(!lotteonSheetCleanupDecision_v671_(name,protection).protected)throw new Error('핵심 시트 보호 실패: '+name);
    });
    if(protection.activePrefixes.PR35!=='RUNNING')throw new Error('PR35 실행 상태 보호 인식 실패');

    var checked=lotteonCleanupSelectedCount_v671_(manager);
    lotteonWriteSheetCleanupSmoke_v671_(
      status,'PASS',
      '비파괴 smoke 완료: 목록 '+result.rows+'행 / 기본 체크 '+checked+'건 / 삭제 0건',
      ''
    );
    return {ok:true,status:'PASS',rows:result.rows,checked:checked,deleted:0};
  }catch(error){
    var message=String(error&&error.message?error.message:error);
    lotteonWriteSheetCleanupSmoke_v671_(status,'ERROR','비파괴 smoke 실패',message);
    throw error;
  }
}

function lotteonWriteSheetCleanupSmoke_v671_(sheet,status,message,error){
  var rows=[
    ['항목','값'],
    ['버전',LOTTEON_SHEET_CLEANUP_VERSION_V671],
    ['상태',status],
    ['메시지',message||''],
    ['오류',error||''],
    ['갱신시각',new Date().toISOString()]
  ];
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setColumnWidth(1,180);
  sheet.setColumnWidth(2,620);
}

function lotteonCleanupSelectedCount_v671_(manager){
  if(!manager||manager.getLastRow()<2)return 0;
  return manager.getRange(2,1,manager.getLastRow()-1,1).getValues().filter(function(row){return row[0]===true;}).length;
}
function lotteonSheetCleanupProtection_v671_(ss){
  var core={}; LOTTEON_SHEET_CLEANUP_CORE_V671.forEach(function(name){core[name]=true;});
  var activePrefixes={};
  ss.getSheets().forEach(function(sheet){
    var name=sheet.getName(),match=name.match(/^(PR\d+)_.*(?:실행상태|적용상태|운영반영상태|운영스모크상태)$/i);
    if(!match||sheet.getLastRow()<2)return;
    var status=lotteonSheetCleanupReadStatus_v671_(sheet);
    if(LOTTEON_SHEET_CLEANUP_ACTIVE_STATES_V671[status])activePrefixes[match[1].toUpperCase()]=status;
  });
  return {core:core,activePrefixes:activePrefixes};
}
function lotteonSheetCleanupReadStatus_v671_(sheet){
  var values=sheet.getRange(1,1,Math.min(sheet.getLastRow(),30),Math.min(sheet.getLastColumn(),2)).getValues();
  for(var i=0;i<values.length;i++){
    if(String(values[i][0]==null?'':values[i][0]).trim()==='상태')return String(values[i][1]==null?'':values[i][1]).trim().toUpperCase();
  }
  return '';
}
function lotteonSheetCleanupDecision_v671_(name,protection){
  if(protection.core[name])return {protected:true,recommended:false,category:'운영 핵심',reason:'항상 보호되는 운영 핵심 시트'};
  var pr=String(name||'').match(/^(PR\d+)_/i);
  var prefix=pr?pr[1].toUpperCase():'';
  if(prefix&&protection.activePrefixes[prefix])return {protected:true,recommended:false,category:'실행 중 작업',reason:prefix+' 상태가 '+protection.activePrefixes[prefix]+'이므로 관련 시트 전체 보호'};
  if(/^PR\d+_/i.test(name))return {protected:false,recommended:true,category:'종료 PR 임시자료',reason:'실행 중 상태가 아닌 PR 미리보기·진단·상태 시트'};
  if(/^script\s*backup$/i.test(name))return {protected:false,recommended:true,category:'스크립트 백업',reason:'임시 script Backup 시트'};
  var diagnostic=/(검증|진단|로그|상태|스냅샷|리포트|보고서|분석|추출경고|자가진단|삭제후보|원본대상행|분석제외|정리검증|매핑검증|원본확인)/;
  if(diagnostic.test(name))return {protected:false,recommended:false,category:'진단·로그 검토',reason:'자동 삭제하지 않음. 내용을 확인한 뒤 필요 없을 때 직접 체크'};
  return {protected:false,recommended:false,category:'기타 시트 검토',reason:'자동 삭제하지 않음. 용도를 확인한 뒤 직접 체크'};
}
function lotteonSheetCleanupUi_v671_(){
  try{return SpreadsheetApp.getUi();}catch(ignore){return null;}
}
function lotteonSheetCleanupAlert_v671_(title,message){
  var ui=lotteonSheetCleanupUi_v671_();
  if(ui)ui.alert(title,message,ui.ButtonSet.OK);
}
