const assert = require('assert'), fs = require('fs'), vm = require('vm');
const stateStore={}; let triggers=[];
const ctx={PropertiesService:{getScriptProperties:()=>({getProperty:k=>stateStore[k],setProperty:(k,v)=>stateStore[k]=v})},ScriptApp:{getProjectTriggers:()=>triggers,newTrigger:handler=>({timeBased:()=>({after:()=>({create:()=>triggers.push({getHandlerFunction:()=>handler})})})}),deleteTrigger:t=>triggers=triggers.filter(x=>x!==t)}};
vm.createContext(ctx); vm.runInContext(fs.readFileSync('Patch_v6_59_lifecycle_safe_continuation.gs','utf8'),ctx);
ctx.scheduleLifecycleTrigger_v659_(); ctx.scheduleLifecycleTrigger_v659_(); assert.equal(triggers.length,1); assert.equal(triggers[0].getHandlerFunction(),'continueBrandLifecycleDashboard_v659_');
ctx.saveLifecycleState_v659_({status:'pending',spreadsheetId:'x'}); assert.equal(ctx.getLifecycleState_v659_().status,'pending');
ctx.clearLifecycleTriggers_v659_(); assert.equal(triggers.length,0);
console.log('v6.59 continuation mock: OK (single trigger, state roundtrip, trigger cleanup)');

