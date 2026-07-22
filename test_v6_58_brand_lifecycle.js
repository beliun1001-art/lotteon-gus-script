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
const hyphen=ctx.aggregateLifecycleFilters_v658_({rows:[{filterName:'롯백_03_2_브랜드-A_01',brand:'2_브랜드-A_01',accountId:'a',collectCount:1,managedScore:0}]});
assert.equal(hyphen.byBrand[ctx.lifecycleBrandKey_v658_('브랜드-A')].brand,'브랜드-A');
assert.equal(ctx.lifecycleBrandKey_v658_('브랜드-A'),ctx.lifecycleBrandKey_v658_('브랜드_A'));
assert.equal(ctx.parseLifecycleFilterName_v658_('롯백_03_2_브랜드_A_01').brand,'브랜드_A'); // display underscore remains untouched
const managed=ctx.aggregateLifecycleFilters_v658_({rows:[{filterName:'롯백_03_2_동률_01',brand:'동률',accountId:'low',collectCount:10,managedScore:1},{filterName:'롯백_01_4_동률',brand:'동률',accountId:'high',collectCount:10,managedScore:9}]});
assert.equal(managed.byBrand['동률'].representativeFilterName,'롯백_01_4_동률'); assert.equal(managed.byBrand['동률'].representativeAccountId,'high');
const criteriaValues=[['기준키','기준값'],['확장최소경과일',10],['확장30일환산순수매출액','500,000'],['퇴출최소경과일',20]];
const criteriaSheet={getLastRow:()=>criteriaValues.length,getDataRange:()=>({getValues:()=>criteriaValues})};
const override=ctx.loadLifecycleRules_v658_({getSheetByName:name=>name==='기준'?criteriaSheet:null});
assert.equal(override.expansionMinDays,10); assert.equal(override.expansionSales30,500000); assert.equal(override.exitMinDays,20);
assert.equal(ctx.decideLifecycle_v658_(10,0,500000,'2026-07-01',override).status,'확장');
const prefixedValues=[['기준키','기준값'],['브랜드_확장_최소경과일',11],['브랜드_확장_30일환산매출',510000],['브랜드_확장_30일환산매출품목수',12],['브랜드_퇴출_최소경과일',21],['브랜드_퇴출_최대30일환산매출',100],['브랜드_퇴출_최대30일환산매출품목수',2]];
const prefixed=ctx.loadLifecycleRules_v658_({getSheetByName:()=>({getLastRow:()=>prefixedValues.length,getDataRange:()=>({getValues:()=>prefixedValues})})});
assert.deepStrictEqual(JSON.parse(JSON.stringify(prefixed)),{expansionMinDays:11,exitMinDays:21,expansionProducts30:12,expansionSales30:510000,exitProducts30:2,exitSales30:100});
ctx.CONFIG={MIN_OBSERVE_DAYS:31,DELETE_REVIEW_DAYS:46,GOOD_REVENUE_30D:1000001};
assert.deepStrictEqual(JSON.parse(JSON.stringify(ctx.lifecycleDefaultRules_v658_())),{expansionMinDays:31,exitMinDays:46,expansionProducts30:10,expansionSales30:1000001,exitProducts30:0,exitSales30:0});
const productCounts=ctx.lifecycleProductCounts_v658_({detailRows:[{brand:'숲',productNo:'P1',marketProductNo:'M1',productName:'이름A'},{brand:'숲',productNo:'P1',marketProductNo:'M1',productName:'이름B'},{brand:'숲',productNo:'',marketProductNo:'M2',productName:'이름C'},{brand:'숲',productNo:'',marketProductNo:'',productName:'미집계'}]});
assert.equal(productCounts['숲'],2);
const lifecycleRows=ctx.buildBrandLifecycleRows_v658_({detailRows:[{brand:'숲',productNo:'P1',productName:'A'},{brand:'숲',productNo:'P1',productName:'B'}],byBrand:{숲:{brand:'숲',salesProductCount:99,netSalesAmount:0}}},filters,'2026-07-22',ctx.lifecycleDefaultRules_v658_());
assert.equal(lifecycleRows[0].products,1);
function fakeRule(status,column){return {getBooleanCondition:()=>({getCriteriaValues:()=>[status]}),getRanges:()=>[{getColumn:()=>column}]};}
let stored=[fakeRule('기타',5)];
ctx.SpreadsheetApp={newConditionalFormatRule:()=>{const built={}; return {whenTextEqualTo:text=>{built.text=text; return {setBackground:()=>({setRanges:ranges=>({build:()=>fakeRule(built.text,ranges[0].getColumn())})})};}};}};
const formatSheet={getRange:(row,col)=>({getColumn:()=>col}),getConditionalFormatRules:()=>stored,setConditionalFormatRules:rules=>{stored=rules;}};
ctx.applyLifecycleConditionalFormats_v658_(formatSheet,10,2); const firstCount=stored.length; ctx.applyLifecycleConditionalFormats_v658_(formatSheet,20,2); assert.equal(firstCount,4); assert.equal(stored.length,4); assert.equal(stored.filter(ctx.isLifecycleStatusRule_v658_).length,3);
assert.deepStrictEqual(JSON.parse(JSON.stringify(ctx.lifecycleStatusFormatSpecs_v658_().map(x=>x.text))),['확장','유지','퇴출']);
ctx.CONFIG={};
assert.equal(ctx.decideLifecycle_v658_(0,100,9999999,'2026-07-21').status,'유지'); assert.equal(ctx.decideLifecycle_v658_(29,100,9999999,'2026-06-22').status,'유지'); assert.equal(ctx.decideLifecycle_v658_(30,0,1000000,'2026-06-21').status,'확장'); assert.equal(ctx.decideLifecycle_v658_(45,0,0,'2026-06-06').status,'퇴출'); assert.equal(ctx.decideLifecycle_v658_(null,0,0,'').status,'유지');
console.log('v6.58 lifecycle mock: OK (key contract, managed-score representative, product distinct, conditional idempotence, brand criteria keys, parser, decisions)');

