const assert=require('assert'),fs=require('fs'),vm=require('vm');
const props={},triggers=[]; const ctx={PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k],setProperty:(k,v)=>props[k]=v})},ScriptApp:{getProjectTriggers:()=>triggers,newTrigger:h=>({timeBased:()=>({after:()=>({create:()=>triggers.push({getHandlerFunction:()=>h})})})}),deleteTrigger:t=>{const i=triggers.indexOf(t);if(i>=0)triggers.splice(i,1);}},LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})}}; vm.createContext(ctx); vm.runInContext(fs.readFileSync('Patch_v6_59_lifecycle_safe_continuation.gs','utf8'),ctx);
ctx.CONFIG={SHEETS:{DASHBOARD:'대시보드'}};
const headers=ctx.snapshotHeaders_v659_(), sourceRows=Array.from({length:450},(_,i)=>Array.from({length:headers.length},(_,j)=>j===1?'brand-'+i:i)); const writes=[];
function range(row,col,count){return {getValues:()=>sourceRows.slice(row-2,row-2+count),setValues:v=>{writes.push({row,col,count:v.length,values:v});return range(row,col,v.length);},setNumberFormat:()=>range(row,col,count),setBackground:()=>range(row,col,count),setFontWeight:()=>range(row,col,count)};}
const source={getRange:(r,c,n)=>range(r,c,n)},dash={getLastRow:()=>10,getRange:(r,c,n)=>range(r,c,n)}; const ss={getSheetByName:n=>n===ctx.LOTTEON_V659_SNAPSHOT_SHEET?source:dash};
const state={runId:'run-new',lifecycleRows:450,nextIndex:0,dashboardStartRow:0,noProgressAttempts:0};
assert.equal(ctx.currentRunMatches_v659_('run-new'),false); ctx.saveLifecycleState_v659_(state); assert.equal(ctx.currentRunMatches_v659_('run-new'),true);
assert.equal(ctx.writeLifecycleChunk_v659_(ss,state),false); assert.equal(ctx.writeLifecycleChunk_v659_(ss,state),false); assert.equal(ctx.writeLifecycleChunk_v659_(ss,state),true);
const dataWrites=writes.filter(w=>w.col===1&&w.count>1); assert.deepStrictEqual(dataWrites.map(w=>w.count),[200,200,50]); assert.equal(state.nextIndex,450); assert.equal(new Set(dataWrites.flatMap(w=>w.values.map(r=>r[1]))).size,450); assert.equal(writes.filter(w=>w.count===1&&w.row===12).length,1);
ctx.scheduleLifecycleTrigger_v659_(state); ctx.scheduleLifecycleTrigger_v659_(state); assert.equal(triggers.length,1); ctx.clearLifecycleTriggers_v659_(); assert.equal(triggers.length,0);
console.log('v6.59 continuation mock: OK (actual 450 chunk offsets 200/200/50, no duplicate header/rows, single trigger, run token)');

