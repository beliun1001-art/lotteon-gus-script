/** v6.57 Issue #11: VAT period columns and one-pass period summary only. */
var LOTTEON_PATCH_V657_VAT_PERIOD_SUMMARY_LOADED = true;
var __baseFinishVatSummaries_v657_ = typeof finishVatSummaries_v648_ === 'function' ? finishVatSummaries_v648_ : null;

if (__baseFinishVatSummaries_v657_) {
  finishVatSummaries_v648_ = function(ss, state) {
    // v6.54 marks the job done internally. Silence only its final toast until
    // the mandatory period summary succeeds, then persist a failed state on error.
    var originalToast = toastVat_v648_;
    toastVat_v648_ = function() {};
    try {
      var result = __baseFinishVatSummaries_v657_.apply(this, arguments);
      buildVatPeriodSummary_v657_(ss || SpreadsheetApp.getActive());
      originalToast(ss || SpreadsheetApp.getActive(), '부가세 자료 및 기간별 집계 생성이 완료되었습니다.');
      return result;
    } catch (e) {
      finalizeVatPeriodFailure_v657_(ss || SpreadsheetApp.getActive(), state, e);
      return { ok: false, failed: true, state: state };
    } finally {
      toastVat_v648_ = originalToast;
    }
  };
}

function buildVatPeriodSummary_v657_(ss) {
  var detail = ss.getSheetByName('부가세_신고자료');
  if (!detail || detail.getLastRow() < 1) return { rows: 0, reason: 'NO_DETAIL' };
  var values = detail.getDataRange().getValues();
  var enriched = enrichVatDetailPeriods_v657_(values, ss);
  if (enriched.changed) {
    detail.clearContents();
    detail.getRange(1, 1, enriched.values.length, enriched.values[0].length).setValues(enriched.values);
    detail.setFrozenRows(1);
    applyVatPeriodSheetFormat_v657_(detail, enriched.values[0]);
  }
  var summary = aggregateVatPeriods_v657_(enriched.values);
  var headers = ['집계구분','신고연도','반기','신고월','쿠팡계정ID','사업자등록번호','주문건수','순수매출액','정산기준금액','마켓수수료','매입금액','매출부가세','매입부가세','납부예상부가세','예상이익','부가세반영예상이익'];
  var sheet = ss.getSheetByName('부가세_기간별') || ss.insertSheet('부가세_기간별');
  sheet.clearContents(); sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (summary.length) sheet.getRange(2, 1, summary.length, headers.length).setValues(summary);
  sheet.setFrozenRows(1);
  applyVatPeriodSheetFormat_v657_(sheet, headers);
  return { rows: summary.length, periodUnknown: summary.filter(function(r) { return r[0] === '기간미확인'; }).reduce(function(n, r) { return n + r[6]; }, 0), blankOrderDetailRows: summary.blankOrderDetailRows || 0 };
}

function finalizeVatPeriodFailure_v657_(ss, state, error) {
  state = state || {};
  state.status = 'failed'; state.phase = 'period_summary'; state.nextRunScheduled = false;
  state.updatedAt = new Date().toISOString();
  state.lastError = '기간별 집계 실패: ' + String(error && error.message ? error.message : error);
  if (typeof saveVatState_v648_ === 'function') saveVatState_v648_(state);
  if (typeof writeVatStatus_v648_ === 'function') writeVatStatus_v648_(ss, state, state.lastError);
  if (typeof toastVat_v648_ === 'function') toastVat_v648_(ss, '부가세 기간별 집계가 실패했습니다. 상태 시트를 확인하세요.');
  return state;
}

function enrichVatDetailPeriods_v657_(values, ss) {
  if (!values || !values.length) return { values: values || [], changed: false };
  var header = values[0].slice();
  var dateIndex = findVatPeriodHeader_v657_(header, ['주문일','주문일자','마켓주문일자','날짜'], 0);
  var yearIndex = header.indexOf('신고연도');
  if (yearIndex >= 0) return { values: values, changed: false };
  var output = [header.slice(0, dateIndex + 1).concat(['신고연도','반기','신고월']).concat(header.slice(dateIndex + 1))];
  var sourcePeriods = vatPeriodsFromSource_v657_(ss);
  for (var i = 1; i < values.length; i++) {
    // The v6.48 detail date deliberately displays MM/dd. Use the original source
    // date when possible; never infer a reporting year from the display value.
    var key = vatDetailPeriodKey_v657_(values[i]);
    var queue = sourcePeriods[key];
    var period = queue && queue.length ? queue.shift() : vatPeriodFromDate_v657_(values[i][dateIndex]);
    output.push(values[i].slice(0, dateIndex + 1).concat([period.year, period.half, period.month]).concat(values[i].slice(dateIndex + 1)));
  }
  return { values: output, changed: true };
}
function vatPeriodsFromSource_v657_(ss) {
  var out = {};
  var source = ss && ss.getSheetByName && ss.getSheetByName(LOTTEON_V648_SOURCE_SHEET);
  if (!source || typeof vatHeaderIndexes_v648_ !== 'function' || typeof vatDetailRow_v648_ !== 'function') return out;
  var maxCol = Math.min(source.getLastColumn(), LOTTEON_V648_MAX_COL || source.getLastColumn());
  if (source.getLastRow() < 2 || !maxCol) return out;
  var headers = source.getRange(1, 1, 1, maxCol).getValues()[0];
  var indexes = vatHeaderIndexes_v648_(headers);
  var rows = source.getRange(2, 1, source.getLastRow() - 1, maxCol).getValues();
  rows.forEach(function(row, offset) {
    var result = vatDetailRow_v648_(row, indexes, offset + 2);
    if (!result || !result.row) return;
    var key = vatDetailPeriodKey_v657_(result.row);
    if (!out[key]) out[key] = [];
    out[key].push(vatPeriodFromDate_v657_(valueAt_v648_(row, indexes.date)));
  });
  return out;
}
function vatDetailPeriodKey_v657_(row) { return [row[0], row[1], row[3], row[6], row[9]].map(function(v) { return String(v == null ? '' : v).trim(); }).join('|'); }
function vatPeriodFromDate_v657_(raw) {
  if (Object.prototype.toString.call(raw) === '[object Date]' && !isNaN(raw.getTime())) return vatPeriodFromParts_v657_(raw.getFullYear(), raw.getMonth() + 1);
  var text = String(raw == null ? '' : raw).trim();
  var match = text.match(/^(\d{4})[.\/-](\d{1,2})[.\/-](\d{1,2})/);
  if (!match) return { year: '기간미확인', half: '기간미확인', month: '기간미확인' };
  return vatPeriodFromParts_v657_(Number(match[1]), Number(match[2]));
}
function vatPeriodFromParts_v657_(year, month) { if (!year || month < 1 || month > 12) return { year: '기간미확인', half: '기간미확인', month: '기간미확인' }; return { year: String(year), half: month <= 6 ? '상반기' : '하반기', month: String(year) + '-' + (month < 10 ? '0' : '') + month }; }
function aggregateVatPeriods_v657_(values) {
  var header = values[0] || [], map = {}, ix = function(names, fallback) { return findVatPeriodHeader_v657_(header, names, fallback); };
  var year = ix(['신고연도'], -1), half = ix(['반기'], -1), month = ix(['신고월'], -1), account = ix(['쿠팡계정ID'], 1), business = ix(['사업자등록번호'], 2);
  var orderNo = ix(['마켓주문번호','주문번호','주문ID','주문ID(마켓)'], -1), blankOrderDetailRows = 0;
  var sales = ix(['순수매출액'], 9), settlement = ix(['정산기준금액'], 12), fee = ix(['마켓수수료'], 13), purchase = ix(['매입금액'], 14), salesVat = ix(['매출부가세'], 11), purchaseVat = ix(['매입부가세'], 16), payable = ix(['납부예상부가세'], 17), profit = ix(['예상이익'], 18), vatProfit = ix(['부가세반영예상이익'], 19);
  function add(kind, y, h, m, row) { var key = [kind,y,h,m,row[account] || '',row[business] || ''].join('|'); if (!map[key]) map[key] = { row: [kind,y,h,m,row[account] || '',row[business] || '',0,0,0,0,0,0,0,0,0,0], orders: {} }; var item = map[key], out = item.row, order = orderNo >= 0 ? String(row[orderNo] == null ? '' : row[orderNo]).trim() : ''; if (order) { if (!item.orders[order]) { item.orders[order] = true; out[6]++; } } else blankOrderDetailRows++; var cols = [sales,settlement,fee,purchase,salesVat,purchaseVat,payable,profit,vatProfit]; for (var c = 0; c < cols.length; c++) out[7+c] += vatPeriodNumber_v657_(row[cols[c]]); }
  for (var r = 1; r < values.length; r++) { var row = values[r], y = year >= 0 ? row[year] : '기간미확인', h = half >= 0 ? row[half] : '기간미확인', m = month >= 0 ? row[month] : '기간미확인'; if (y === '기간미확인') add('기간미확인', y, h, m, row); else { add('반기', y, h, '', row); add('월별', y, h, m, row); } }
  var output = Object.keys(map).map(function(k) { return map[k].row; }).sort(function(a,b) { var rank = { '반기': 0, '월별': 1, '기간미확인': 2 }; return String(a[1]).localeCompare(String(b[1])) || (rank[a[0]]-rank[b[0]]) || String(a[2]).localeCompare(String(b[2])) || String(a[3]).localeCompare(String(b[3])) || String(a[4]).localeCompare(String(b[4])); });
  output.blankOrderDetailRows = blankOrderDetailRows; // blank order numbers are intentionally not counted as orders.
  return output;
}
function findVatPeriodHeader_v657_(headers, names, fallback) { for (var n = 0; n < names.length; n++) for (var i = 0; i < headers.length; i++) if (String(headers[i] || '').replace(/\s/g,'') === String(names[n]).replace(/\s/g,'')) return i; return fallback; }
function vatPeriodNumber_v657_(value) { var n = Number(String(value == null ? 0 : value).replace(/[원,\s]/g,'')); return isNaN(n) ? 0 : n; }
function vatPeriodFormatPlan_v657_(headers) { var plan = {}; headers.forEach(function(header, index) { var h = String(header || '').replace(/\s/g, ''); if (/순수매출액|정산기준금액|마켓수수료|매입금액|매출부가세|매입부가세|납부예상부가세|예상이익|부가세반영예상이익/.test(h)) plan[index] = '#,##0'; else if (/신고연도|반기|신고월|주문번호|상품번호|계정ID|아이디|사업자등록번호/.test(h)) plan[index] = '@'; }); return plan; }
function applyVatPeriodSheetFormat_v657_(sheet, headers) { var lastRow = Math.max(sheet.getLastRow(), 1), count = headers.length, plan = vatPeriodFormatPlan_v657_(headers); sheet.getRange(1, 1, 1, count).setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center'); Object.keys(plan).forEach(function(key) { sheet.getRange(2, Number(key) + 1, Math.max(lastRow - 1, 1), 1).setNumberFormat(plan[key]); }); for (var i = 0; i < count; i++) sheet.setColumnWidth(i + 1, plan[i] === '#,##0' ? 110 : (plan[i] === '@' ? 120 : 90)); }

