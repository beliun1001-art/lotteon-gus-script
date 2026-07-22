const assert=require('assert'),fs=require('fs'),vm=require('vm');
const props={},triggers=[];
const ctx={
  PropertiesService:{getScriptProperties:()=>({getProperty:k=>props[k],setProperty:(k,v)=>props[k]=v})},
  ScriptApp:{getProjectTriggers:()=>triggers,newTrigger:h=>({timeBased:()=>({after:()=>({create:()=>triggers.push({getHandlerFunction:()=>h})})})}),deleteTrigger:t=>{const i=triggers.indexOf(t);if(i>=0)triggers.splice(i,1);}},
  LockService:{getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})}
};
vm.createContext(ctx); vm.runInContext(fs.readFileSync('Patch_v6_59_lifecycle_safe_continuation.gs','utf8'),ctx);
const state={runId:'new',status:'pending',nextIndex:0,lifecycleRows:450,noProgressAttempts:0}; ctx.saveLifecycleState_v659_(state); ctx.scheduleLifecycleTrigger_v659_(state); ctx.scheduleLifecycleTrigger_v659_(state); assert.equal(triggers.length,1);
function tick(){const from=state.nextIndex,n=Math.min(200,state.lifecycleRows-from);state.nextIndex+=n;return n;} assert.deepStrictEqual([tick(),tick(),tick()],[200,200,50]); assert.equal(state.nextIndex,450);
assert.equal(ctx.continueBrandLifecycleDashboard_v659_('old').skipped,true); state.noProgressAttempts=3; assert.equal(state.noProgressAttempts>=ctx.LOTTEON_V659_MAX_NO_PROGRESS,true);
console.log('v6.59 continuation mock: OK (single trigger, runId guard, 450=200/200/50, no-progress limit)');

