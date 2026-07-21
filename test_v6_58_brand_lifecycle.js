const assert = require('assert'), fs = require('fs'), vm = require('vm');
const ctx = { Utilities:{ formatDate:d => d.toISOString().slice(0,10) }, Session:{ getScriptTimeZone:()=>'Asia/Seoul' } }; vm.createContext(ctx); vm.runInContext(fs.readFileSync('Patch_v6_58_brand_lifecycle_dashboard.gs','utf8'),ctx);
assert.deepStrictEqual(JSON.parse(JSON.stringify(ctx.parseLifecycleFilterName_v658_('롯백_03_2_숲_01'))),{valid:true,brand:'숲',origin:'03',group:'2',movedTo:'01'});
assert.equal(ctx.parseLifecycleFilterName_v658_('롯백_01_4_숲').brand,'숲');
assert.equal(ctx.parseLifecycleFilterName_v658_('롯백_03_2_내부_언더스코어_02').brand,'내부_언더스코어');
const filters=ctx.aggregateLifecycleFilters_v658_({rows:[{filterName:'롯백_03_2_숲_01',brand:'2_숲_01',accountId:'a',collectCount:5,createDate:'2026-06-07',recentDate:'2026-06-08'},{filterName:'롯백_01_4_숲',brand:'4_숲',accountId:'b',collectCount:7,createDate:'2026-07-01',recentDate:'2026-07-02'}]});
assert.equal(Object.keys(filters.byBrand).length,1); assert.equal(filters.byBrand['숲'].firstCollectionDate,'2026-06-07'); assert.equal(filters.byBrand['숲'].collectCount,12);
assert.equal(ctx.decideLifecycle_v658_(0,100,9999999,'2026-07-21').status,'유지'); assert.equal(ctx.decideLifecycle_v658_(29,100,9999999,'2026-06-22').status,'유지'); assert.equal(ctx.decideLifecycle_v658_(30,0,1000000,'2026-06-21').status,'확장'); assert.equal(ctx.decideLifecycle_v658_(45,0,0,'2026-06-06').status,'퇴출'); assert.equal(ctx.decideLifecycle_v658_(null,0,0,'').status,'유지');
console.log('v6.58 lifecycle mock: OK (canonical filter parsing, move continuity, underscore, decisions)');

