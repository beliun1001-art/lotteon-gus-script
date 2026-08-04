/**
 * Issue #15 diagnostic-only runner.
 *
 * Reads:
 * - 매출데이터_붙여넣기
 * - 부가세_신고자료
 *
 * Writes only:
 * - 부가세_기간이상_진단
 * - PR15_진단상태
 *
 * It never changes source, VAT detail, period summary, card matching, account
 * mapping, dates, purchase amounts, or financial formulas.
 */
var PR15_DIAG_VERSION = 'v1.0-ISSUE15-FUTURE-ACCOUNT-DIAGNOSTIC';
var PR15_DIAG_SHEET = '부가세_기간이상_진단';
var PR15_STATUS_SHEET = 'PR15_진단상태';
var PR15_SOURCE_SHEET = '매출데이터_붙여넣기';
var PR15_DETAIL_SHEET = '부가세_신고자료';
var PR15_TARGET_ACCOUNTS = {
  'beliun1008': true,
  'beliun1010': true,
  'beliun1008@gmail.com': true
};

function runPr15FutureAccountDiagnostic() {
  var ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('현재 스프레드시트를 찾지 못했습니다.');
  pr15WriteStatus_(ss, 'RUNNING', 'Issue #15 원본·부가세 상세 queue 진단 시작', null, '');

  try {
    if (typeof vatHeaderIndexes_v648_ !== 'function' ||
        typeof vatDetailRow_v648_ !== 'function' ||
        typeof vatDetailPeriodKey_v657_ !== 'function' ||
        typeof vatPeriodFromDate_v657_ !== 'function') {
      throw new Error('v6.57 진단 필수 함수가 main bundle에 없습니다.');
    }

    var source = ss.getSheetByName(PR15_SOURCE_SHEET);
    var detail = ss.getSheetByName(PR15_DETAIL_SHEET);
    if (!source || source.getLastRow() < 2) throw new Error('원본 시트가 없거나 비어 있습니다: ' + PR15_SOURCE_SHEET);
    if (!detail || detail.getLastRow() < 2) throw new Error('부가세 상세 시트가 없거나 비어 있습니다: ' + PR15_DETAIL_SHEET);

    var sourceRange = source.getDataRange();
    var sourceValues = sourceRange.getValues();
    var sourceDisplays = sourceRange.getDisplayValues();
    var sourceHeaders = sourceValues[0];
    var sourceIx = vatHeaderIndexes_v648_(sourceHeaders);
    var sourceExtraIx = pr15SourceExtraIndexes_(sourceHeaders);

    var sourceQueues = {};
    var sourceItems = [];
    for (var r = 1; r < sourceValues.length; r++) {
      var built = vatDetailRow_v648_(sourceValues[r], sourceIx, r + 1);
      if (!built || !built.row) continue;
      var key = vatDetailPeriodKey_v657_(built.row);
      if (!sourceQueues[key]) sourceQueues[key] = [];
      var period = vatPeriodFromDate_v657_(sourceValues[r][sourceIx.date]);
      var item = {
        sourceRow: r + 1,
        sourceKey: key,
        sourceQueueIndex: sourceQueues[key].length + 1,
        rawDate: sourceValues[r][sourceIx.date],
        displayDate: sourceDisplays[r][sourceIx.date],
        dateType: pr15ValueType_(sourceValues[r][sourceIx.date]),
        sourceDate: pr15ParseDate_(sourceValues[r][sourceIx.date]),
        sourcePeriod: period,
        account: pr15Text_(sourceValues[r][sourceIx.marketAccount]),
        business: pr15Text_(built.row[2]),
        orderNo: pr15Text_(built.row[3]),
        productNo: pr15Text_(built.row[6]),
        sales: pr15Number_(built.row[9]),
        purchase: pr15Number_(built.row[14]),
        marketStatus: pr15Cell_(sourceValues[r], sourceExtraIx.marketStatus),
        themangoStatus: pr15Cell_(sourceValues[r], sourceExtraIx.themangoStatus),
        matchedDetailRow: 0
      };
      sourceQueues[key].push(item);
      sourceItems.push(item);
    }

    var detailRange = detail.getDataRange();
    var detailValues = detailRange.getValues();
    var detailDisplays = detailRange.getDisplayValues();
    var detailIx = pr15DetailIndexes_(detailValues[0]);
    pr15RequireDetailIndexes_(detailIx);

    var consumed = {};
    var output = [];
    var matchedSourceRows = {};
    var today = pr15Today_();

    for (var d = 1; d < detailValues.length; d++) {
      var match = pr15TakeSourceMatch_(
        sourceQueues,
        consumed,
        detailValues[d],
        detailDisplays[d],
        detailIx
      );
      var sourceItem = match.item;
      if (sourceItem) {
        sourceItem.matchedDetailRow = d + 1;
        matchedSourceRows[sourceItem.sourceRow] = true;
      }

      var detailAccount = pr15Text_(detailValues[d][detailIx.account]);
      var detailMonth = pr15Text_(detailValues[d][detailIx.month]);
      var detailYear = pr15Text_(detailValues[d][detailIx.year]);
      var detailBusiness = pr15Text_(detailValues[d][detailIx.business]);
      var suspicious = pr15IsTargetAccount_(detailAccount) ||
        detailMonth === '2026-10' || detailMonth === '2026-11' ||
        (sourceItem && pr15SourceItemSuspicious_(sourceItem, today));
      if (!suspicious) continue;

      output.push(pr15BuildDiagnosticRow_({
        sourceItem: sourceItem,
        detailRow: d + 1,
        detailYear: detailYear,
        detailMonth: detailMonth,
        detailAccount: detailAccount,
        detailBusiness: detailBusiness,
        detailOrderNo: pr15Text_(detailValues[d][detailIx.orderNo]),
        detailProductNo: pr15Text_(detailValues[d][detailIx.productNo]),
        detailSales: pr15Number_(detailValues[d][detailIx.sales]),
        detailPurchase: pr15Number_(detailValues[d][detailIx.purchase]),
        matchedKey: match.key,
        queueIndex: match.queueIndex,
        today: today
      }));
    }

    // Also surface suspicious source rows that were excluded or could not be
    // paired with VAT detail. This prevents a queue mismatch from hiding evidence.
    sourceItems.forEach(function(item) {
      if (matchedSourceRows[item.sourceRow]) return;
      if (!pr15SourceItemSuspicious_(item, today)) return;
      output.push(pr15BuildDiagnosticRow_({
        sourceItem: item,
        detailRow: 0,
        detailYear: '',
        detailMonth: '',
        detailAccount: '',
        detailBusiness: '',
        detailOrderNo: '',
        detailProductNo: '',
        detailSales: 0,
        detailPurchase: 0,
        matchedKey: item.sourceKey,
        queueIndex: item.sourceQueueIndex,
        today: today
      }));
    });

    output.sort(function(a, b) {
      return Number(a[0] || 0) - Number(b[0] || 0) || Number(a[1] || 0) - Number(b[1] || 0);
    });

    if (!output.length) throw new Error('Issue #15 진단 대상 행을 찾지 못했습니다.');

    var headers = [
      '원본행','신고자료행','주문번호','원본주문일 값','원본주문일 표시','원본주문일 타입',
      '복원연도','신고월','계정ID','사업자번호','상품번호','순수매출액','매입금액',
      '마켓주문상태','더망고주문상태','매칭 key','queue index','이상사유',
      'root_cause_status','production 수정 필요 여부'
    ];
    var diag = ss.getSheetByName(PR15_DIAG_SHEET) || ss.insertSheet(PR15_DIAG_SHEET);
    diag.clear();
    diag.getRange(1, 1, 1, headers.length).setValues([headers]);
    diag.getRange(2, 1, output.length, headers.length).setValues(output);
    diag.setFrozenRows(1);
    diag.getRange(1, 1, 1, headers.length)
      .setBackground('#d9eaf7').setFontWeight('bold').setHorizontalAlignment('center');
    diag.getRange(2, 1, output.length, 3).setNumberFormat('@');
    diag.getRange(2, 7, output.length, 5).setNumberFormat('@');
    diag.getRange(2, 12, output.length, 2).setNumberFormat('#,##0');
    diag.getRange(2, 16, output.length, 1).setNumberFormat('@');
    [80,90,150,180,120,120,90,100,180,150,150,110,110,150,150,420,90,420,210,220]
      .forEach(function(width, index) { diag.setColumnWidth(index + 1, width); });

    var stats = pr15CountStatuses_(output);
    pr15WriteStatus_(ss, 'PASS',
      'Issue #15 진단 완료: ' + output.length + '행 / 운영 원본 변경 0건', stats, '');
    return {ok:true, rows:output.length, stats:stats, productionWrites:0};
  } catch (error) {
    var message = String(error && error.message ? error.message : error);
    pr15WriteStatus_(ss, 'ERROR', 'Issue #15 진단 실패', null, message);
    throw error;
  }
}

function runPr15FutureAccountDiagnosticContinue() {
  return runPr15FutureAccountDiagnostic();
}

function pr15BuildDiagnosticRow_(ctx) {
  var item = ctx.sourceItem;
  var sourceMonth = item && item.sourcePeriod ? pr15Text_(item.sourcePeriod.month) : '';
  var sourceYear = item && item.sourcePeriod ? pr15Text_(item.sourcePeriod.year) : '';
  var detailMonth = pr15Text_(ctx.detailMonth);
  var account = pr15Text_(ctx.detailAccount || (item && item.account));
  var business = pr15Text_(ctx.detailBusiness || (item && item.business));
  var reasons = [];

  if (!item) reasons.push('원본 queue 매칭 실패');
  if (item && pr15IsFutureDate_(item.sourceDate, ctx.today)) reasons.push('원본 주문일이 실행일보다 미래');
  if (detailMonth && pr15MonthAfterToday_(detailMonth, ctx.today)) reasons.push('신고월이 실행월보다 미래');
  if (item && sourceMonth && detailMonth && sourceMonth !== detailMonth) {
    reasons.push('원본 월 ' + sourceMonth + ' ↔ 신고월 ' + detailMonth + ' 불일치');
  }
  if (pr15IsTargetAccount_(account)) reasons.push('확정 사업자 매핑 외 계정');
  if (!business) reasons.push('사업자번호 공란');
  var purchase = item ? item.purchase : ctx.detailPurchase;
  if (!purchase) reasons.push('AC 매입금액 0');
  if (!ctx.detailRow) reasons.push('부가세_신고자료 상세행 미연결');

  var root = pr15ClassifyRootCause_({
    sourceItem:item,
    sourceMonth:sourceMonth,
    detailMonth:detailMonth,
    account:account,
    business:business,
    today:ctx.today
  });

  return [
    item ? String(item.sourceRow) : '',
    ctx.detailRow ? String(ctx.detailRow) : '',
    pr15Text_(ctx.detailOrderNo || (item && item.orderNo)),
    item ? pr15RawValueText_(item.rawDate) : '',
    item ? pr15Text_(item.displayDate) : '',
    item ? item.dateType : '',
    detailMonth ? pr15Text_(ctx.detailYear) : sourceYear,
    detailMonth || sourceMonth,
    account,
    business,
    pr15Text_(ctx.detailProductNo || (item && item.productNo)),
    ctx.detailRow ? ctx.detailSales : (item ? item.sales : 0),
    ctx.detailRow ? ctx.detailPurchase : (item ? item.purchase : 0),
    item ? item.marketStatus : '',
    item ? item.themangoStatus : '',
    pr15Text_(ctx.matchedKey || (item && item.sourceKey)),
    Number(ctx.queueIndex || (item && item.sourceQueueIndex) || 0),
    reasons.join(' / ') || '이상 사유 미확정',
    root,
    pr15ProductionAction_(root)
  ];
}

function pr15ClassifyRootCause_(ctx) {
  if (!ctx.sourceItem) return 'UNRESOLVED';
  if (ctx.sourceMonth && ctx.detailMonth && ctx.sourceMonth !== ctx.detailMonth) {
    return 'YEAR_RESTORE_BUG_CONFIRMED';
  }
  if (pr15IsFutureDate_(ctx.sourceItem.sourceDate, ctx.today)) {
    return 'SOURCE_DATA_ERROR';
  }
  if (pr15IsTargetAccount_(ctx.account) && !ctx.business) {
    return 'ACCOUNT_MAPPING_REQUIRED';
  }
  if (ctx.sourceMonth && ctx.detailMonth && ctx.sourceMonth === ctx.detailMonth) {
    return 'SOURCE_CONFIRMED_VALID';
  }
  return 'UNRESOLVED';
}

function pr15ProductionAction_(root) {
  if (root === 'YEAR_RESTORE_BUG_CONFIRMED') return '기간 복원 코드 수정 필요';
  if (root === 'ACCOUNT_MAPPING_REQUIRED') return '신고 대상 계정·사업자 매핑 사용자 확인 필요';
  if (root === 'SOURCE_DATA_ERROR') return '원본 주문일 정정 여부 확인 필요';
  if (root === 'SOURCE_CONFIRMED_VALID') return 'production 수정 없음';
  return '추가 원본 증거 필요';
}

function pr15TakeSourceMatch_(queues, consumed, row, displayRow, ix) {
  var candidates = pr15DetailKeys_(row, displayRow, ix);
  for (var i = 0; i < candidates.length; i++) {
    var key = candidates[i];
    var queue = queues[key];
    var used = Number(consumed[key] || 0);
    if (queue && used < queue.length) {
      consumed[key] = used + 1;
      return {item:queue[used], key:key, queueIndex:used + 1};
    }
  }
  return {item:null, key:candidates[0] || '', queueIndex:0};
}

function pr15DetailKeys_(row, displayRow, ix) {
  var dates = [pr15Text_(row[ix.date]), pr15Text_(displayRow[ix.date])];
  var account = pr15Text_(row[ix.account]);
  var orderNo = pr15Text_(row[ix.orderNo]);
  var productNo = pr15Text_(row[ix.productNo]);
  var sales = pr15Text_(row[ix.sales]);
  var out = [], seen = {};
  dates.forEach(function(date) {
    var key = [date, account, orderNo, productNo, sales].join('|');
    if (!seen[key]) { seen[key] = true; out.push(key); }
  });
  return out;
}

function pr15DetailIndexes_(headers) {
  return {
    date:pr15FindHeader_(headers,['주문일','주문일자','마켓주문일자','날짜']),
    year:pr15FindHeader_(headers,['신고연도']),
    month:pr15FindHeader_(headers,['신고월']),
    account:pr15FindHeader_(headers,['쿠팡계정ID','마켓아이디','계정ID']),
    business:pr15FindHeader_(headers,['사업자등록번호','사업자번호']),
    orderNo:pr15FindHeader_(headers,['마켓주문번호','주문번호','주문ID','주문ID(마켓)']),
    productNo:pr15FindHeader_(headers,['상품번호','사이트상품번호','마켓상품번호']),
    sales:pr15FindHeader_(headers,['순수매출액','결제금액합계','결제금액']),
    purchase:pr15FindHeader_(headers,['매입금액','구매가격','구매금액'])
  };
}

function pr15RequireDetailIndexes_(ix) {
  ['date','year','month','account','business','orderNo','productNo','sales','purchase'].forEach(function(name) {
    if (ix[name] < 0) throw new Error('부가세_신고자료 필수 헤더 없음: ' + name);
  });
}

function pr15SourceExtraIndexes_(headers) {
  return {
    marketStatus:pr15FindHeader_(headers,['마켓주문상태','주문상태','상태']),
    themangoStatus:pr15FindHeader_(headers,['더망고주문상태','더망고 상태'])
  };
}

function pr15FindHeader_(headers, names) {
  for (var n = 0; n < names.length; n++) {
    var wanted = String(names[n]).replace(/\s/g,'');
    for (var i = 0; i < headers.length; i++) {
      if (String(headers[i] == null ? '' : headers[i]).replace(/\s/g,'') === wanted) return i;
    }
  }
  return -1;
}

function pr15SourceItemSuspicious_(item, today) {
  return pr15IsTargetAccount_(item.account) ||
    item.sourcePeriod.month === '2026-10' || item.sourcePeriod.month === '2026-11' ||
    pr15IsFutureDate_(item.sourceDate, today);
}

function pr15IsTargetAccount_(value) {
  return !!PR15_TARGET_ACCOUNTS[pr15Text_(value).toLowerCase()];
}

function pr15ParseDate_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return new Date(value.getFullYear(), value.getMonth(), value.getDate());
  }
  var text = pr15Text_(value);
  var match = text.match(/^(20\d{2})[-\/.](\d{1,2})[-\/.](\d{1,2})/);
  if (!match) return null;
  var date = new Date(Number(match[1]), Number(match[2]) - 1, Number(match[3]));
  return isNaN(date.getTime()) ? null : date;
}

function pr15Today_() {
  var now = new Date();
  var text = Utilities.formatDate(now, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  var parts = text.split('-');
  return new Date(Number(parts[0]), Number(parts[1]) - 1, Number(parts[2]));
}

function pr15IsFutureDate_(date, today) {
  return !!date && date.getTime() > today.getTime();
}

function pr15MonthAfterToday_(month, today) {
  var match = pr15Text_(month).match(/^(20\d{2})-(\d{2})$/);
  if (!match) return false;
  return Number(match[1]) * 12 + Number(match[2]) > today.getFullYear() * 12 + (today.getMonth() + 1);
}

function pr15CountStatuses_(rows) {
  var out = {
    rows:rows.length,
    sourceValid:0,
    yearRestoreBug:0,
    accountMapping:0,
    sourceDataError:0,
    unresolved:0
  };
  rows.forEach(function(row) {
    var status = row[18];
    if (status === 'SOURCE_CONFIRMED_VALID') out.sourceValid++;
    else if (status === 'YEAR_RESTORE_BUG_CONFIRMED') out.yearRestoreBug++;
    else if (status === 'ACCOUNT_MAPPING_REQUIRED') out.accountMapping++;
    else if (status === 'SOURCE_DATA_ERROR') out.sourceDataError++;
    else out.unresolved++;
  });
  return out;
}

function pr15WriteStatus_(ss, status, message, stats, error) {
  var sheet = ss.getSheetByName(PR15_STATUS_SHEET) || ss.insertSheet(PR15_STATUS_SHEET);
  stats = stats || {};
  var rows = [
    ['항목','값'],
    ['버전',PR15_DIAG_VERSION],
    ['상태',status],
    ['메시지',message || ''],
    ['진단행',stats.rows || 0],
    ['SOURCE_CONFIRMED_VALID',stats.sourceValid || 0],
    ['YEAR_RESTORE_BUG_CONFIRMED',stats.yearRestoreBug || 0],
    ['ACCOUNT_MAPPING_REQUIRED',stats.accountMapping || 0],
    ['SOURCE_DATA_ERROR',stats.sourceDataError || 0],
    ['UNRESOLVED',stats.unresolved || 0],
    ['운영 원본 변경','0건'],
    ['오류',error || ''],
    ['갱신시각',new Date().toISOString()]
  ];
  sheet.clearContents();
  sheet.getRange(1,1,rows.length,2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1,1,1,2).setBackground('#d9eaf7').setFontWeight('bold');
  sheet.setColumnWidth(1,260);
  sheet.setColumnWidth(2,650);
}

function pr15ValueType_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]') return 'Date';
  if (value === null || value === '') return 'blank';
  return typeof value;
}

function pr15RawValueText_(value) {
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', "yyyy-MM-dd'T'HH:mm:ss");
  }
  return pr15Text_(value);
}

function pr15Cell_(row, index) {
  return index < 0 ? '' : pr15Text_(row[index]);
}

function pr15Text_(value) {
  return String(value == null ? '' : value).trim();
}

function pr15Number_(value) {
  var n = Number(String(value == null ? 0 : value).replace(/[,원\s]/g,''));
  return isNaN(n) ? 0 : n;
}
