/** PR #31 v1.2: resumable production apply. Run start once; continuation is automatic. */
const PR30_V12_VERSION='v1.2-PR30-BATCHED-AUTO-CONTINUE';
const PR30_V12_HANDLER='runPr30ProductionApplyV12Continue';
const PR30_V12_STATE_KEY='PR30_V12_STATE';
const PR30_V12_BATCH=150;

function runPr30ProductionApplyV12(){
  const ss=SpreadsheetApp.getActive();
  if(!ss)throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  if(typeof pr30ReadAndValidatePreview_!=='function'||typeof pr30v11RestoreBackup_!=='function'||typeof pr30v11Prepare_!=='function'){
    throw new Error('기존 PR30 원본 및 v1.1 별도 스크립트 파일이 필요합니다.');
  }
  pr30v12DeleteTriggers_();
  const previous=pr30v12StatusValue_(ss,'상태');
  const backups=!!ss.getSheetByName(PR30_BACKUP_PERIOD_SHEET)&&!!ss.getSheetByName(PR30_BACKUP_DIAG_SHEET);
  const state={spreadsheetId:ss.getId(),stage:(backups&&previous!=='PASS')?'RECOVER_PERIOD':'VALIDATE',offset:0,
    startedAt:new Date().toISOString(),error:'',stats:{matched:0,nonCard:0,ambiguous:0,noMatch:0,fallback:0}};
  pr30v12Save_(state);pr30v12Status_(ss,'RUNNING','배치 실행 시작',state);
  return runPr30ProductionApplyV12Continue();
}

function runPr30ProductionApplyV12Continue(){
  const lock=LockService.getScriptLock();if(!lock.tryLock(5000))return{ok:false,reason:'LOCK_BUSY'};
  try{
    const s=pr30v12Load_();if(!s||!s.spreadsheetId)throw new Error('v1.2 실행 상태가 없습니다.');
    const ss=SpreadsheetApp.openById(s.spreadsheetId);
    const fn={RECOVER_PERIOD:pr30v12RecoverPeriod_,RECOVER_DIAG:pr30v12RecoverDiag_,VALIDATE:pr30v12Validate_,
      BACKUP_PERIOD:pr30v12BackupPeriod_,BACKUP_DIAG:pr30v12BackupDiag_,APPLY_SUMMARY:pr30v12ApplySummary_,
      PREP_DIAG:pr30v12PrepDiag_,WRITE_DIAG:pr30v12WriteDiag_,VERIFY_SUMMARY:pr30v12VerifySummary_,
      VERIFY_DIAG:pr30v12VerifyDiag_,ROLLBACK_PERIOD:pr30v12RollbackPeriod_,ROLLBACK_DIAG:pr30v12RollbackDiag_}[s.stage];
    if(!fn){if(s.stage==='DONE')return{ok:true,done:true};throw new Error('알 수 없는 단계: '+s.stage);}
    return fn(ss,s);
  }catch(e){
    const s=pr30v12Load_()||{},ss=s.spreadsheetId?SpreadsheetApp.openById(s.spreadsheetId):SpreadsheetApp.getActive();
    s.error=String(e&&e.message?e.message:e);s.stage=ss&&ss.getSheetByName(PR30_BACKUP_PERIOD_SHEET)?'ROLLBACK_PERIOD':'DONE';
    pr30v12Save_(s);if(ss)pr30v12Status_(ss,s.stage==='DONE'?'ERROR':'ROLLBACK_PENDING',s.error,s);
    if(s.stage!=='DONE')pr30v12Schedule_();return{ok:false,error:s.error,rollbackScheduled:s.stage!=='DONE'};
  }finally{lock.releaseLock();}
}

function pr30v12RecoverPeriod_(ss,s){pr30v11RestoreBackup_(ss,PR30_BACKUP_PERIOD_SHEET,PR30_PRODUCTION_PERIOD_SHEET);return pr30v12Next_(ss,s,'RECOVER_DIAG','기존 기간별 백업 복구 완료');}
function pr30v12RecoverDiag_(ss,s){pr30v11RestoreBackup_(ss,PR30_BACKUP_DIAG_SHEET,PR30_PRODUCTION_DIAG_SHEET);return pr30v12Next_(ss,s,'VALIDATE','기존 카드검증 백업 복구 완료');}

function pr30v12Validate_(ss,s){
  const p=pr30ReadAndValidatePreview_(ss),prod=pr30ReadProduction_(ss);
  s.summaryRows=p.summaryRows.length;s.diagRows=p.diagValues.length-1;
  s.detailSignature=pr30MatrixSignature_(prod.periodDetailValues);s.offset=0;
  return pr30v12Next_(ss,s,'BACKUP_PERIOD','PR29 PASS 및 운영 상세표 검증 완료');
}
function pr30v12BackupPeriod_(ss,s){pr30v12ReplaceBackup_(ss,PR30_PRODUCTION_PERIOD_SHEET,PR30_BACKUP_PERIOD_SHEET);return pr30v12Next_(ss,s,'BACKUP_DIAG','기간별 백업 완료');}
function pr30v12BackupDiag_(ss,s){pr30v12ReplaceBackup_(ss,PR30_PRODUCTION_DIAG_SHEET,PR30_BACKUP_DIAG_SHEET);return pr30v12Next_(ss,s,'APPLY_SUMMARY','카드검증 백업 완료');}

function pr30v12ApplySummary_(ss,s){
  const preview=pr30ReadAndValidatePreview_(ss),prod=pr30ReadProduction_(ss);
  pr30v11ApplySummary_(prod.periodSheet,preview.summaryHeaders,preview.summaryRows,prod.periodDetailValues);
  SpreadsheetApp.flush();return pr30v12Next_(ss,s,'PREP_DIAG','운영 구매카드 요약 반영 완료');
}

function pr30v12PrepDiag_(ss,s){
  const src=ss.getSheetByName(PR30_PREVIEW_DIAG_SHEET),dst=ss.getSheetByName(PR30_PRODUCTION_DIAG_SHEET);
  const h=src.getRange(1,1,1,src.getLastColumn()).getValues()[0].map(pr30Text_);
  dst.clearContents();dst.getRange(1,1,1,h.length).setValues([h]);
  h.forEach(function(x,i){dst.getRange(2,i+1,s.diagRows,1).setNumberFormat(PR30_V11_DIAG_NUMERIC.indexOf(x)>=0?'#,##0':'@');});
  dst.setFrozenRows(1);dst.getRange(1,1,1,h.length).setBackground('#d9eaf7').setFontWeight('bold');s.offset=0;
  return pr30v12Next_(ss,s,'WRITE_DIAG','카드검증 배치 쓰기 준비 완료');
}

function pr30v12WriteDiag_(ss,s){
  const src=ss.getSheetByName(PR30_PREVIEW_DIAG_SHEET),dst=ss.getSheetByName(PR30_PRODUCTION_DIAG_SHEET);
  const h=src.getRange(1,1,1,src.getLastColumn()).getValues()[0].map(pr30Text_),n=Math.min(PR30_V12_BATCH,s.diagRows-s.offset);
  if(n>0){const rows=src.getRange(s.offset+2,1,n,h.length).getValues();dst.getRange(s.offset+2,1,n,h.length).setValues(pr30v11Prepare_(h,rows,PR30_V11_DIAG_NUMERIC));s.offset+=n;}
  if(s.offset>=s.diagRows){s.offset=0;return pr30v12Next_(ss,s,'VERIFY_SUMMARY','카드검증 1,355건 배치 반영 완료');}
  return pr30v12Stay_(ss,s,'카드검증 배치 반영 '+s.offset+' / '+s.diagRows);
}

function pr30v12VerifySummary_(ss,s){
  const p=pr30ReadAndValidatePreview_(ss),period=ss.getSheetByName(PR30_PRODUCTION_PERIOD_SHEET).getDataRange().getValues();
  const width=p.summaryHeaders.length,actualH=period[1].slice(0,width).map(pr30Text_),actualRows=period.slice(2,2+p.summaryRows.length).map(r=>r.slice(0,width));
  pr30v11AssertEqual_('운영 구매카드 요약',pr30v11Canonical_(p.summaryHeaders,p.summaryRows,PR30_V11_SUMMARY_NUMERIC),pr30v11Canonical_(actualH,actualRows,PR30_V11_SUMMARY_NUMERIC));
  const detailIndex=period.findIndex((r,i)=>i>=2&&pr30Text_(r[0])==='집계구분'&&r.map(pr30Text_).indexOf('신고연도')>=0);
  if(detailIndex<0||pr30MatrixSignature_(period.slice(detailIndex))!==s.detailSignature)throw new Error('기간 상세표 불변 검증 실패');
  s.offset=0;s.stats={matched:0,nonCard:0,ambiguous:0,noMatch:0,fallback:0};return pr30v12Next_(ss,s,'VERIFY_DIAG','요약 및 기간 상세표 검증 완료');
}

function pr30v12VerifyDiag_(ss,s){
  const src=ss.getSheetByName(PR30_PREVIEW_DIAG_SHEET),dst=ss.getSheetByName(PR30_PRODUCTION_DIAG_SHEET);
  const h=src.getRange(1,1,1,src.getLastColumn()).getValues()[0].map(pr30Text_),n=Math.min(PR30_V12_BATCH,s.diagRows-s.offset);
  if(n>0){
    const exp=pr30v11Prepare_(h,src.getRange(s.offset+2,1,n,h.length).getValues(),PR30_V11_DIAG_NUMERIC);
    const act=pr30v11Prepare_(h,dst.getRange(s.offset+2,1,n,h.length).getValues(),PR30_V11_DIAG_NUMERIC);
    pr30v11AssertEqual_('운영 카드검증 배치 '+s.offset,exp,act);pr30v12Count_(h,act,s.stats);s.offset+=n;
  }
  if(s.offset<s.diagRows)return pr30v12Stay_(ss,s,'카드검증 배치 검증 '+s.offset+' / '+s.diagRows);
  if(JSON.stringify(s.stats)!==JSON.stringify({matched:766,nonCard:494,ambiguous:24,noMatch:71,fallback:593}))throw new Error('최종 상태 건수 불일치: '+JSON.stringify(s.stats));
  s.stage='DONE';s.completedAt=new Date().toISOString();pr30v12Save_(s);pr30v12DeleteTriggers_();pr30v12Status_(ss,'PASS','배치 운영 반영 및 검증 완료',s);
  PropertiesService.getScriptProperties().deleteProperty(PR30_V12_STATE_KEY);ss.toast('PR30 v1.2 완료: NO_MATCH 71건','LOTTEON 자동화',10);return{ok:true,done:true,stats:s.stats};
}

function pr30v12RollbackPeriod_(ss,s){pr30v11RestoreBackup_(ss,PR30_BACKUP_PERIOD_SHEET,PR30_PRODUCTION_PERIOD_SHEET);return pr30v12Next_(ss,s,'ROLLBACK_DIAG','기간별 롤백 완료','ROLLBACK_RUNNING');}
function pr30v12RollbackDiag_(ss,s){
  pr30v11RestoreBackup_(ss,PR30_BACKUP_DIAG_SHEET,PR30_PRODUCTION_DIAG_SHEET);s.stage='DONE';s.completedAt=new Date().toISOString();
  pr30v12Save_(s);pr30v12DeleteTriggers_();pr30v12Status_(ss,'ROLLED_BACK',s.error||'오류로 자동 롤백',s);PropertiesService.getScriptProperties().deleteProperty(PR30_V12_STATE_KEY);
  return{ok:false,rolledBack:true,error:s.error||''};
}

function pr30v12Count_(h,rows,x){
  const si=h.indexOf('카드매칭상태'),fi=h.indexOf('v6.69 2차귀속');rows.forEach(function(r){const v=pr30Text_(r[si]);
    if(v==='MATCHED'||v==='MASTER_MATCHED')x.matched++;else if(v==='NON_CARD')x.nonCard++;else if(v==='AMBIGUOUS')x.ambiguous++;else x.noMatch++;if(pr30Text_(r[fi])==='Y')x.fallback++;});
}
function pr30v12ReplaceBackup_(ss,sourceName,backupName){const old=ss.getSheetByName(backupName);if(old)ss.deleteSheet(old);const src=ss.getSheetByName(sourceName);if(!src)throw new Error('백업 대상 없음: '+sourceName);src.copyTo(ss).setName(backupName);}
function pr30v12Next_(ss,s,next,msg,status){s.stage=next;pr30v12Save_(s);pr30v12Status_(ss,status||'RUNNING',msg,s);pr30v12Schedule_();return{ok:true,stage:next};}
function pr30v12Stay_(ss,s,msg){pr30v12Save_(s);pr30v12Status_(ss,'RUNNING',msg,s);pr30v12Schedule_();return{ok:true,stage:s.stage,offset:s.offset};}
function pr30v12Schedule_(){pr30v12DeleteTriggers_();ScriptApp.newTrigger(PR30_V12_HANDLER).timeBased().after(60*1000).create();}
function pr30v12DeleteTriggers_(){ScriptApp.getProjectTriggers().forEach(t=>{if(t.getHandlerFunction()===PR30_V12_HANDLER)try{ScriptApp.deleteTrigger(t);}catch(e){}});}
function pr30v12Save_(s){PropertiesService.getScriptProperties().setProperty(PR30_V12_STATE_KEY,JSON.stringify(s));}
function pr30v12Load_(){const v=PropertiesService.getScriptProperties().getProperty(PR30_V12_STATE_KEY);return v?JSON.parse(v):null;}
function pr30v12StatusValue_(ss,key){const sh=ss.getSheetByName(PR30_STATUS_SHEET);if(!sh||sh.getLastRow()<2)return'';const v=sh.getDataRange().getValues();for(let i=1;i<v.length;i++)if(String(v[i][0])===key)return String(v[i][1]||'');return'';}
function pr30v12Status_(ss,status,msg,s){
  const sh=ss.getSheetByName(PR30_STATUS_SHEET)||ss.insertSheet(PR30_STATUS_SHEET),x=s.stats||{};
  const rows=[['항목','값'],['버전',PR30_V12_VERSION],['상태',status],['단계',s.stage||''],['메시지',msg||''],['처리행',s.offset||0],['대상행',s.diagRows||0],
    ['MATCHED',x.matched||0],['NON_CARD',x.nonCard||0],['AMBIGUOUS',x.ambiguous||0],['NO_MATCH',x.noMatch||0],['2차귀속',x.fallback||0],['오류',s.error||''],
    ['시작시각',s.startedAt||''],['완료시각',s.completedAt||''],['갱신시각',new Date().toISOString()]];
  sh.clearContents();sh.getRange(1,1,rows.length,2).setValues(rows);sh.setFrozenRows(1);sh.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');sh.setColumnWidth(1,220);sh.setColumnWidth(2,600);
}
