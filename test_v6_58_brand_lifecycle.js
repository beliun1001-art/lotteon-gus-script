const assert = require('assert'), fs = require('fs'), vm = require('vm');
const ctx = { Utilities:{ formatDate:d => d.toISOString().slice(0,10) }, Session:{ getScriptTimeZone:()=>'Asia/Seoul' } }; vm.createContext(ctx); vm.runInContext(fs.readFileSync('Patch_v6_58_brand_lifecycle_dashboard.gs','utf8'),ctx);
assert.deepStrictEqual(JSON.parse(JSON.stringify(ctx.parseLifecycleFilterName_v658_('롯백_03_2_숲_01'))),{valid:true,brand:'숲',origin:'03',group:'2',movedTo:'01'});
assert.equal(ctx.parseLifecycleFilterName_v658_('롯백_01_4_숲').brand,'숲');
assert.equal(ctx.parseLifecycleFilterName_v658_('롯백_03_2_내부_언더스코어_02').brand,'내부_언더스코어');
const filters=ctx.aggregateLifecycleFilters_v658_({rows:[{filterName:'롯백_03_2_숲_01',brand:'2_숲_01',accountId:'a',collectCount:5,createDate:'2026-06-07',recentDate:'2026-06-08'},{filterName:'롯백_01_4_숲',brand:'4_숲',accountId:'b',collectCount:7,createDate:'2026-07-01',recentDate:'2026-07-02'}]});
assert.equal(Object.keys(filters.byBrand).length,1); assert.equal(filters.byBrand['숲'].firstCollectionDate,'2026-06-07'); assert.equal(filters.byBrand['숲'].collectCount,12);
const forest=filters.byBrand['숲']; ['latestRecentDate','earliestCreateDate','filters','accountIds','representativeFilterName','representativeAccountId','collectCount'].forEach(key=>assert.ok(Object.prototype.hasOwnProperty.call(forest,key),key));
assert.equal(forest.latestRecentDate,'2026-07-02'); assert.equal(forest.earliestCreateDate,'2026-06-07'); assert.equal(forest.filters.length,2); assert.equal(Object.keys(forest.accountIds).length,2);
// aggregateFiltersByBrand_v611_ output dates are normalized strings; canonical re-aggregation must preserve product-rotation recent date.
const baseOutput={rows:[{filterName:'롯백_03_2_숲_01',brand:'2_숲_01',accountId:'a',collectCount:1,createDate:'2026-06-07',recentDate:'2026-06-11'},{filterName:'롯백_01_4_숲',brand:'4_숲',accountId:'b',collectCount:1,createDate:'2026-07-01',recentDate:'2026-07-21'}]};
assert.equal(ctx.aggregateLifecycleFilters_v658_(baseOutput).byBrand['숲'].latestRecentDate,'2026-07-21');
const criteriaValues=[['기준키','기준값'],['확장최소경과일',10],['확장30일환산순수매출액','500,000'],['퇴출최소경과일',20]];
const criteriaSheet={getLastRow:()=>criteriaValues.length,getDataRange:()=>({getValues:()=>criteriaValues})};
const override=ctx.loadLifecycleRules_v658_({getSheetByName:name=>name==='기준'?criteriaSheet:null});
assert.equal(override.expansionMinDays,10); assert.equal(override.expansionSales30,500000); assert.equal(override.exitMinDays,20);
assert.equal(ctx.decideLifecycle_v658_(10,0,500000,'2026-07-01',override).status,'확장');
assert.deepStrictEqual(JSON.parse(JSON.stringify(ctx.lifecycleStatusFormatSpecs_v658_().map(x=>x.text))),['확장','유지','퇴출']);
assert.equal(ctx.decideLifecycle_v658_(0,100,9999999,'2026-07-21').status,'유지'); assert.equal(ctx.decideLifecycle_v658_(29,100,9999999,'2026-06-22').status,'유지'); assert.equal(ctx.decideLifecycle_v658_(30,0,1000000,'2026-06-21').status,'확장'); assert.equal(ctx.decideLifecycle_v658_(45,0,0,'2026-06-06').status,'퇴출'); assert.equal(ctx.decideLifecycle_v658_(null,0,0,'').status,'유지');
console.log('v6.58 lifecycle mock: OK (contract/recent-date, criteria override, conditional format, parser, move continuity, decisions)');

