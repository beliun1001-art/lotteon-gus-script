const assert=require('assert'),fs=require('fs'),vm=require('vm');
const props={},triggers=[]; const ctx={PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k],setProperty:(k,v)=>props[k]=v})},ScriptApp:{getProjectTriggers:()=>triggers,newTrigger:h=>({timeBased:()=>({after:delay=>({create:()=>triggers.push({getHandlerFunction:()=>h,delay})})})}),deleteTrigger:t=>{const i=triggers.indexOf(t);if(i>=0)triggers.splice(i,1);}},LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})}}; vm.createContext(ctx); vm.runInContext(fs.readFileSync('Patch_v6_59_lifecycle_safe_continuation.gs','utf8'),ctx);
ctx.CONFIG={SHEETS:{DASHBOARD:'대시보드'}};
const headers=ctx.snapshotHeaders_v659_(), sourceRows=Array.from({length:450},(_,i)=>Array.from({length:headers.length},(_,j)=>j===1?'brand-'+i:i)); const writes=[];
function range(row,col,count){return {getValues:()=>sourceRows.slice(row-2,row-2+count),setValues:v=>{writes.push({row,col,count:v.length,values:v});return range(row,col,v.length);},setNumberFormat:()=>range(row,col,count),setBackground:()=>range(row,col,count),setFontWeight:()=>range(row,col,count)};}
const source={getRange:(r,c,n)=>range(r,c,n)},dash={getLastRow:()=>10,getRange:(r,c,n)=>range(r,c,n)}; const ss={getSheetByName:n=>n===ctx.LOTTEON_V659_SNAPSHOT_SHEET?source:dash};
const state={runId:'run-new',lifecycleRows:450,nextIndex:0,dashboardStartRow:0,noProgressAttempts:0};
assert.equal(ctx.currentRunMatches_v659_('run-new'),false); ctx.saveLifecycleState_v659_(state); assert.equal(ctx.currentRunMatches_v659_('run-new'),true);
assert.equal(ctx.writeLifecycleChunk_v659_(ss,state),false); assert.equal(ctx.writeLifecycleChunk_v659_(ss,state),false); assert.equal(ctx.writeLifecycleChunk_v659_(ss,state),true);
const dataWrites=writes.filter(w=>w.col===1&&w.count>1); assert.deepStrictEqual(dataWrites.map(w=>w.count),[200,200,50]); assert.equal(state.nextIndex,450); assert.equal(new Set(dataWrites.flatMap(w=>w.values.map(r=>r[1]))).size,450); assert.equal(writes.filter(w=>w.count===1&&w.row===12).length,1);
ctx.scheduleLifecycleTrigger_v659_(state); ctx.scheduleLifecycleTrigger_v659_(state); assert.equal(triggers.length,1); ctx.clearLifecycleTriggers_v659_(); assert.equal(triggers.length,0);

// A watchdog consumed while the snapshot owns the lock must leave one delayed retry.
const busyState={runId:'busy-run',spreadsheetId:'ss',status:'pending',snapshotReady:false,triggerScheduled:true}; ctx.saveLifecycleState_v659_(busyState);
ctx.LockService={getScriptLock:()=>({tryLock:()=>false,releaseLock:()=>{}})};
const busyResult=ctx.continueBrandLifecycleDashboard_v659_(); assert.equal(busyResult.busy,true); assert.equal(busyResult.rescheduled,true); assert.equal(triggers.length,1); assert.equal(triggers[0].delay,ctx.LOTTEON_V659_SNAPSHOT_WATCHDOG_DELAY_MS); ctx.clearLifecycleTriggers_v659_();

// Model a hard timeout that left snapshotAttempts=1 and its watchdog pending. The next
// production-style parameterless handler re-enters, increments attempts, then clears the
// watchdog after a normal snapshot/publish completion.
const retryState={runId:'retry-run',spreadsheetId:'ss',status:'pending',snapshotReady:false,snapshotAttempts:1,lifecycleRows:0,nextIndex:0,noProgressAttempts:0,triggerScheduled:true}; ctx.saveLifecycleState_v659_(retryState); ctx.scheduleLifecycleTrigger_v659_(retryState,ctx.LOTTEON_V659_SNAPSHOT_WATCHDOG_DELAY_MS);
ctx.LockService={getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})}; ctx.SpreadsheetApp={openById:()=>ss,setActiveSpreadsheet:()=>{}}; ctx.saveStatus_v659_=(book,s)=>ctx.saveLifecycleState_v659_(s); ctx.runPhase_v659_=(book,s,phase,fn)=>{s.phase=phase+':start';const out=fn();s.phase=phase+':done';return out;}; ctx.prepareLifecycleSnapshot_v659_=(book,s)=>{s.snapshotReady=true;s.lifecycleRows=0;return 0;}; ctx.writeLifecycleChunk_v659_=()=>true; ctx.applyLifecycleConditionalFormats_v658_=()=>{};
const retryResult=ctx.continueBrandLifecycleDashboard_v659_(), savedRetry=ctx.getLifecycleState_v659_(); assert.equal(retryResult.done,true); assert.equal(savedRetry.snapshotAttempts,2); assert.equal(savedRetry.status,'done'); assert.equal(triggers.length,0);

const cleanupSource=fs.readFileSync('Patch_v6_53_vat_filter_and_sheet_cleanup.gs','utf8'), widthSource=fs.readFileSync('Patch_v6_52_safe_column_width_runner.gs','utf8'), lifecycleSource=fs.readFileSync('Patch_v6_59_lifecycle_safe_continuation.gs','utf8'); assert(cleanupSource.includes("'브랜드운영_자동상태'")); assert(!cleanupSource.includes("'브랜드운영_스냅샷'")); assert(widthSource.includes("'브랜드운영_자동상태'")); assert(lifecycleSource.includes('sheet.hideSheet()'));
console.log('v6.59 continuation mock: OK (actual 450 chunk offsets 200/200/50, watchdog lock-busy retry, timeout re-entry attempts=2, normal completion triggers=0, status visible/snapshot hidden allowlist)');
