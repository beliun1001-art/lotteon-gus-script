/** v6.58 Issue #12: canonical filter brands and dashboard-only lifecycle view. */
var LOTTEON_PATCH_V658_BRAND_LIFECYCLE_LOADED = true;
var __baseAggregateFilters_v658_ = typeof aggregateFiltersByBrand_v611_ === 'function' ? aggregateFiltersByBrand_v611_ : null;
var __baseRebuildDashboardSingleSource_v658_ = typeof rebuildDashboardSingleSource_v628_ === 'function' ? rebuildDashboardSingleSource_v628_ : null;

if (__baseAggregateFilters_v658_) {
  aggregateFiltersByBrand_v611_ = function() { return aggregateLifecycleFilters_v658_(__baseAggregateFilters_v658_()); };
}
if (__baseRebuildDashboardSingleSource_v658_) {
  rebuildDashboardSingleSource_v628_ = function(salesAgg, filterAgg) {
    var result = __baseRebuildDashboardSingleSource_v658_.apply(this, arguments);
    appendBrandLifecycleDashboard_v658_(salesAgg, filterAgg || (typeof aggregateFiltersByBrand_v611_ === 'function' ? aggregateFiltersByBrand_v611_() : {}));
    return result;
  };
}

function parseLifecycleFilterName_v658_(filterName) {
  var text = String(filterName || '').trim();
  var match = text.match(/^롯백_(\d{2})_(\d+)_([\s\S]+)$/);
  if (!match) return { valid:false, brand:'', origin:'', group:'', movedTo:'' };
  var tail = match[3].split('_');
  var movedTo = tail.length > 1 && /^(01|02|03|04)$/.test(tail[tail.length - 1]) ? tail.pop() : '';
  var brand = tail.join('_').trim();
  if (!brand) return { valid:false, brand:'', origin:match[1], group:match[2], movedTo:movedTo };
  return { valid:true, brand:brand, origin:match[1], group:match[2], movedTo:movedTo };
}

function canonicalLifecycleBrand_v658_(filterName, explicitBrand) {
  var parsed = parseLifecycleFilterName_v658_(filterName);
  if (parsed.valid) return { brand:parsed.brand, parsed:parsed, memo:'' };
  var fallback = typeof canonicalBrand_v611_ === 'function' ? canonicalBrand_v611_(explicitBrand) : String(explicitBrand || '').trim();
  return { brand:fallback, parsed:parsed, memo:'검색필터명 파싱불가: 명시적 브랜드명 fallback' };
}

function lifecycleBrandKey_v658_(brand) { return String(brand || '').toLowerCase().replace(/\s+/g, '').trim(); }
function lifecycleDate_v658_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  var match = String(value || '').match(/(20\d{2})[-.\/](\d{1,2})[-.\/](\d{1,2})/);
  return match ? match[1] + '-' + ('0' + match[2]).slice(-2) + '-' + ('0' + match[3]).slice(-2) : '';
}

function aggregateLifecycleFilters_v658_(base) {
  var result = { byBrand:{}, rows:[], totalCollectCount:0 };
  (base && base.rows || []).forEach(function(row) {
    var canonical = canonicalLifecycleBrand_v658_(row.filterName, row.brand);
    if (!canonical.brand) return;
    var item = { filterName:row.filterName, brand:canonical.brand, key:lifecycleBrandKey_v658_(canonical.brand), accountId:row.accountId || '', collectCount:Number(row.collectCount || 0), recentDate:lifecycleDate_v658_(row.recentDate), createDate:lifecycleDate_v658_(row.createDate), memo:canonical.memo };
    result.rows.push(item); result.totalCollectCount += item.collectCount;
    var b = result.byBrand[item.key] || (result.byBrand[item.key] = { brand:item.brand, key:item.key, collectCount:0, filters:[], accountIds:{}, latestRecentDate:'', earliestCreateDate:'', earliestRecentDate:'', representativeFilterName:'', representativeAccountId:'', diagnosticMemo:'' });
    b.collectCount += item.collectCount; b.filters.push(item); if (item.accountId) b.accountIds[item.accountId] = true;
    if (item.createDate && (!b.earliestCreateDate || item.createDate < b.earliestCreateDate)) b.earliestCreateDate = item.createDate;
    if (item.recentDate && (!b.earliestRecentDate || item.recentDate < b.earliestRecentDate)) b.earliestRecentDate = item.recentDate;
    if (item.recentDate && (!b.latestRecentDate || item.recentDate > b.latestRecentDate)) b.latestRecentDate = item.recentDate;
    if (item.memo) b.diagnosticMemo = item.memo;
  });
  Object.keys(result.byBrand).forEach(function(key) {
    var b = result.byBrand[key];
    var rep = b.filters.slice().sort(function(a,b) { return b.collectCount - a.collectCount || String(a.filterName).localeCompare(String(b.filterName)); })[0];
    b.representativeFilterName = rep ? rep.filterName : ''; b.representativeAccountId = rep ? rep.accountId : '';
    b.firstCollectionDate = b.earliestCreateDate || b.earliestRecentDate || '';
    b.firstCollectionSource = b.earliestCreateDate ? 'API_필터생성일' : (b.earliestRecentDate ? 'API_최근수집일자 fallback' : '최초수집일 미확인');
  });
  return result;
}

function lifecycleToday_v658_() { return Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd'); }
function lifecycleElapsedDays_v658_(firstDate, today) { if (!firstDate) return null; var a = new Date(firstDate + 'T00:00:00'), b = new Date((today || lifecycleToday_v658_()) + 'T00:00:00'); if (isNaN(a) || isNaN(b) || a > b) return null; return Math.floor((b - a) / 86400000); }
function lifecycleDefaultRules_v658_() { var config = typeof CONFIG !== 'undefined' && CONFIG ? CONFIG : {}; function pick(name, fallback) { var n = Number(config[name]); return isFinite(n) && n >= 0 ? n : fallback; } return { expansionMinDays:pick('MIN_OBSERVE_DAYS',30), exitMinDays:pick('DELETE_REVIEW_DAYS',45), expansionProducts30:10, expansionSales30:pick('GOOD_REVENUE_30D',1000000), exitProducts30:0, exitSales30:0 }; }
function lifecycleRuleKey_v658_(value) { return String(value || '').replace(/[\s_]/g, '').toLowerCase(); }
function loadLifecycleRules_v658_(ss) {
  var rules = lifecycleDefaultRules_v658_(), sheet = ss && ss.getSheetByName && ss.getSheetByName('기준');
  if (!sheet || sheet.getLastRow() < 2) return rules;
  var values = sheet.getDataRange().getValues(), header = values[0] || [], keyCol = -1, valueCol = -1;
  header.forEach(function(h, i) { var key = lifecycleRuleKey_v658_(h); if (/^(key|기준키|항목|설정)$/.test(key)) keyCol = i; if (/^(value|값|기준값|설정값)$/.test(key)) valueCol = i; });
  if (keyCol < 0) keyCol = 0; if (valueCol < 0) valueCol = 1;
  var map = { 확장최소경과일:'expansionMinDays', 브랜드확장최소경과일:'expansionMinDays', 퇴출최소경과일:'exitMinDays', 브랜드퇴출최소경과일:'exitMinDays', 확장30일환산매출품목수:'expansionProducts30', 브랜드확장30일환산매출품목수:'expansionProducts30', 확장30일환산순수매출액:'expansionSales30', 확장30일환산매출:'expansionSales30', 브랜드확장30일환산매출:'expansionSales30', 퇴출최대30일환산매출품목수:'exitProducts30', 브랜드퇴출최대30일환산매출품목수:'exitProducts30', 퇴출최대30일환산순수매출액:'exitSales30', 퇴출최대30일환산매출:'exitSales30', 브랜드퇴출최대30일환산매출:'exitSales30' };
  for (var r = 1; r < values.length; r++) { var target = map[lifecycleRuleKey_v658_(values[r][keyCol])], n = Number(String(values[r][valueCol] == null ? '' : values[r][valueCol]).replace(/[,원\s]/g, '')); if (target && isFinite(n)) rules[target] = n; }
  return rules;
}
function decideLifecycle_v658_(elapsed, products30, sales30, firstDate, rules) { rules = rules || lifecycleDefaultRules_v658_(); if (!firstDate) return { status:'유지', reason:'최초수집일 미확인' }; if (elapsed == null) return { status:'유지', reason:'미래 최초수집일' }; if (elapsed < rules.expansionMinDays) return { status:'유지', reason:'관찰기간 미도달' }; if (elapsed >= rules.exitMinDays && products30 <= rules.exitProducts30 && sales30 <= rules.exitSales30) return { status:'퇴출', reason:'퇴출 기준 충족' }; if (products30 >= rules.expansionProducts30 || sales30 >= rules.expansionSales30) return { status:'확장', reason:products30 >= rules.expansionProducts30 ? '30일환산 품목수 기준 충족' : '30일환산 매출 기준 충족' }; return { status:'유지', reason:'운영 기준 유지' }; }
function lifecycleProductCounts_v658_(salesAgg) { var sets = {}; (salesAgg && salesAgg.detailRows || []).forEach(function(detail) { var key = lifecycleBrandKey_v658_(detail.brand), product = String(detail.productNo || detail.marketProductNo || '').trim(); if (!key || !product) return; if (!sets[key]) sets[key] = {}; sets[key][product] = true; }); var counts = {}; Object.keys(sets).forEach(function(key) { counts[key] = Object.keys(sets[key]).length; }); return counts; }

function buildBrandLifecycleRows_v658_(salesAgg, filterAgg, today, rules) {
  var salesByKey = {};
  var lifecycleProducts = lifecycleProductCounts_v658_(salesAgg);
  Object.keys(salesAgg && salesAgg.byBrand || {}).forEach(function(key) { var s = salesAgg.byBrand[key]; salesByKey[lifecycleBrandKey_v658_(s.brand)] = s; });
  var keys = {}; Object.keys(filterAgg && filterAgg.byBrand || {}).forEach(function(k) { keys[k] = true; }); Object.keys(salesByKey).forEach(function(k) { keys[k] = true; });
  return Object.keys(keys).map(function(key) {
    var f = filterAgg.byBrand && filterAgg.byBrand[key] || {}, s = salesByKey[key] || {};
    var first = f.firstCollectionDate || '', elapsed = lifecycleElapsedDays_v658_(first, today);
    var products = Number(lifecycleProducts[key] || 0), sales = Number(s.netSalesAmount || s.salesAmount || 0), divisor = Math.max(elapsed == null ? 1 : elapsed, 1);
    var products30 = products / divisor * 30, sales30 = sales / divisor * 30, decision = decideLifecycle_v658_(elapsed, products30, sales30, first, rules);
    return { brand:f.brand || s.brand || '브랜드미확인', filterName:f.representativeFilterName || '', accountId:f.representativeAccountId || '', collectCount:Number(f.collectCount || 0), firstDate:first, firstSource:f.firstCollectionSource || '최초수집일 미확인', elapsed:elapsed, products:products, sales:sales, products30:products30, sales30:sales30, status:decision.status, reason:decision.reason, memo:f.diagnosticMemo || '' };
  }).sort(function(a,b) { var rank={확장:0,유지:1,퇴출:2}; return rank[a.status]-rank[b.status] || b.sales30-a.sales30 || b.products30-a.products30 || String(a.brand).localeCompare(String(b.brand)); });
}

function appendBrandLifecycleDashboard_v658_(salesAgg, filterAgg) {
  var ss = SpreadsheetApp.getActive(), sheet = ss.getSheetByName(CONFIG.SHEETS.DASHBOARD); if (!sheet) return { rows:0 };
  var rules = loadLifecycleRules_v658_(ss); // one sheet read per dashboard build, never per brand.
  var rows = buildBrandLifecycleRows_v658_(salesAgg || {}, filterAgg || {}, lifecycleToday_v658_(), rules);
  var start = sheet.getLastRow() + 2, headers = ['구분','브랜드명','대표검색필터명','계정ID','수집수','최초수집일','최초수집일 출처','경과일','매출품목수','순수매출액','30일환산 매출품목수','30일환산 순수매출액','운영구분','운영구분 사유','메모'];
  sheet.getRange(start,1,1,headers.length).setValues([headers]).setBackground('#d9eaf7').setFontWeight('bold');
  if (rows.length) sheet.getRange(start+1,1,rows.length,headers.length).setValues(rows.map(function(r) { return ['브랜드운영',r.brand,r.filterName,r.accountId,r.collectCount,r.firstDate,r.firstSource,r.elapsed == null ? '' : r.elapsed,r.products,r.sales,r.products30,r.sales30,r.status,r.reason,r.memo]; }));
  if (rows.length) { sheet.getRange(start+1,10,rows.length,1).setNumberFormat('#,##0'); sheet.getRange(start+1,11,rows.length,1).setNumberFormat('0.0'); sheet.getRange(start+1,12,rows.length,1).setNumberFormat('#,##0'); }
  applyLifecycleConditionalFormats_v658_(sheet, start + 1, rows.length);
  return { rows:rows.length, expansion:rows.filter(function(r){return r.status==='확장';}).length, maintain:rows.filter(function(r){return r.status==='유지';}).length, exit:rows.filter(function(r){return r.status==='퇴출';}).length };
}
function lifecycleStatusFormatSpecs_v658_() { return [{ text:'확장', color:'#d9ead3' }, { text:'유지', color:'#fff2cc' }, { text:'퇴출', color:'#f4cccc' }]; }
function isLifecycleStatusRule_v658_(rule) { try { var condition = rule.getBooleanCondition && rule.getBooleanCondition(), values = condition && condition.getCriteriaValues ? condition.getCriteriaValues() : [], status = String(values && values[0] || ''); if (['확장','유지','퇴출'].indexOf(status) < 0) return false; return (rule.getRanges ? rule.getRanges() : []).some(function(range) { return range.getColumn && range.getColumn() === 13; }); } catch (e) { return false; } }
function applyLifecycleConditionalFormats_v658_(sheet, row, count) { if (!count || typeof SpreadsheetApp === 'undefined') return; var range = sheet.getRange(row, 13, count, 1), rules = sheet.getConditionalFormatRules().filter(function(rule) { return !isLifecycleStatusRule_v658_(rule); }); lifecycleStatusFormatSpecs_v658_().forEach(function(spec) { rules.push(SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(spec.text).setBackground(spec.color).setRanges([range]).build()); }); sheet.setConditionalFormatRules(rules); }

