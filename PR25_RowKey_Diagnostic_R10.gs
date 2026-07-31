/**
 * PR #25 read-only row-key diagnostic R10.
 *
 * Reads:
 * - 매출데이터_붙여넣기
 * - 부가세_신고자료
 *
 * Writes only:
 * - PR25_행키진단
 *
 * It does not modify 부가세_신고자료 or any card/VAT result sheet.
 */
const PR25_R10_VERSION = 'v1.21-PR25-ROWKEY-DIAGNOSTIC-R10';
const PR25_R10_OUTPUT_SHEET = 'PR25_행키진단';

function runPr25RowKeyDiagnosticR10() {
  const started = Date.now();
  const ss = SpreadsheetApp.getActive();
  const source = ss.getSheetByName('매출데이터_붙여넣기');
  const detail = ss.getSheetByName('부가세_신고자료');
  if (!source || source.getLastRow() < 2) throw new Error('매출데이터_붙여넣기 데이터가 없습니다.');
  if (!detail || detail.getLastRow() < 2) throw new Error('부가세_신고자료 데이터가 없습니다.');

  const sourceValues = source.getDataRange().getDisplayValues();
  const detailValues = detail.getDataRange().getDisplayValues();
  const sourceHeaders = sourceValues[0] || [];
  const detailHeaders = detailValues[0] || [];
  const sourceRows = sourceValues.slice(1);
  const detailRows = detailValues.slice(1);

  const detailIndexes = {
    date: pr25r10_find_(detailHeaders, ['날짜','주문일','주문일자','마켓주문일자']),
    account: pr25r10_find_(detailHeaders, ['쿠팡계정ID']),
    orderNo: pr25r10_find_(detailHeaders, ['주문번호','마켓주문번호','주문ID','주문ID(마켓)']),
    productNo: pr25r10_find_(detailHeaders, ['상품번호','마켓상품번호']),
    sales: pr25r10_find_(detailHeaders, ['순수매출액'])
  };
  Object.keys(detailIndexes).forEach(function(field) {
    if (detailIndexes[field] < 0) throw new Error('부가세_신고자료 필수 열 누락: ' + field);
  });

  const fields = ['date','account','orderNo','productNo','sales'];
  const detailNormalized = {};
  fields.forEach(function(field) {
    const index = detailIndexes[field];
    detailNormalized[field] = detailRows.map(function(row) {
      return pr25r10_normalize_(field, row[index]);
    });
  });

  const rankings = {};
  fields.forEach(function(field) {
    const detailList = detailNormalized[field];
    const detailNonblank = detailList.filter(Boolean).length;
    const candidates = [];
    for (let col = 0; col < sourceHeaders.length; col++) {
      const sourceSet = {};
      for (let r = 0; r < sourceRows.length; r++) {
        const value = pr25r10_normalize_(field, sourceRows[r][col]);
        if (value) sourceSet[value] = true;
      }
      let overlap = 0;
      for (let i = 0; i < detailList.length; i++) {
        const value = detailList[i];
        if (value && sourceSet[value]) overlap++;
      }
      candidates.push({
        field: field,
        column: col,
        header: String(sourceHeaders[col] || ''),
        overlap: overlap,
        detailNonblank: detailNonblank,
        rate: detailNonblank ? overlap / detailNonblank : 0
      });
    }
    candidates.sort(function(a, b) {
      return b.overlap - a.overlap || a.column - b.column;
    });
    rankings[field] = candidates;
  });

  const aliases = {
    date: ['마켓주문일자','주문일자','결제일자','주문일시','날짜'],
    account: ['쿠팡계정ID','계정ID','마켓아이디','마켓ID'],
    orderNo: ['마켓주문번호','주문번호','주문ID','주문ID(마켓)'],
    productNo: ['마켓상품번호','상품번호','상품코드','판매자상품코드'],
    sales: ['결제금액합계(원)','결제금액합계','결제금액','순수매출액','판매금액']
  };

  const optionLists = {};
  fields.forEach(function(field) {
    const top = rankings[field].slice(0, 3).map(function(item) { return item.column; });
    const aliasIndex = pr25r10_find_(sourceHeaders, aliases[field]);
    if (aliasIndex >= 0 && top.indexOf(aliasIndex) < 0) top.push(aliasIndex);
    optionLists[field] = top;
  });

  const combinations = [];
  optionLists.date.forEach(function(dateCol) {
    optionLists.account.forEach(function(accountCol) {
      optionLists.orderNo.forEach(function(orderCol) {
        optionLists.productNo.forEach(function(productCol) {
          optionLists.sales.forEach(function(salesCol) {
            const cols = {date:dateCol, account:accountCol, orderNo:orderCol, productNo:productCol, sales:salesCol};
            const score = pr25r10_scoreCombination_(sourceRows, detailNormalized, cols);
            combinations.push({
              matched: score.matched,
              unmatched: detailRows.length - score.matched,
              sourceResidual: sourceRows.length - score.matched,
              dateCol: dateCol,
              accountCol: accountCol,
              orderCol: orderCol,
              productCol: productCol,
              salesCol: salesCol
            });
          });
        });
      });
    });
  });
  combinations.sort(function(a, b) {
    return b.matched - a.matched || a.unmatched - b.unmatched;
  });

  const output = ss.getSheetByName(PR25_R10_OUTPUT_SHEET) || ss.insertSheet(PR25_R10_OUTPUT_SHEET);
  output.clearContents();
  const rows = [];
  rows.push(['PR25 행키 진단', PR25_R10_VERSION]);
  rows.push(['원본 행수', sourceRows.length]);
  rows.push(['원본 열수', sourceHeaders.length]);
  rows.push(['상세 행수', detailRows.length]);
  rows.push(['실행시간(초)', Math.round((Date.now() - started) / 1000)]);
  rows.push([]);
  rows.push(['필드','순위','원본 열번호','원본 헤더','상세 비공란','단일값 포함 일치행','일치율']);
  fields.forEach(function(field) {
    rankings[field].slice(0, 8).forEach(function(item, index) {
      rows.push([
        field,
        index + 1,
        item.column + 1,
        item.header,
        item.detailNonblank,
        item.overlap,
        item.detailNonblank ? Math.round(item.rate * 10000) / 100 + '%' : '0%'
      ]);
    });
  });
  rows.push([]);
  rows.push(['조합순위','매칭행','미매칭행','원본잔여','날짜 열','날짜 헤더','계정 열','계정 헤더','주문번호 열','주문번호 헤더','상품번호 열','상품번호 헤더','매출 열','매출 헤더']);
  combinations.slice(0, 20).forEach(function(item, index) {
    rows.push([
      index + 1,
      item.matched,
      item.unmatched,
      item.sourceResidual,
      item.dateCol + 1,
      sourceHeaders[item.dateCol] || '',
      item.accountCol + 1,
      sourceHeaders[item.accountCol] || '',
      item.orderCol + 1,
      sourceHeaders[item.orderCol] || '',
      item.productCol + 1,
      sourceHeaders[item.productCol] || '',
      item.salesCol + 1,
      sourceHeaders[item.salesCol] || ''
    ]);
  });

  const maxCols = rows.reduce(function(max, row) { return Math.max(max, row.length); }, 1);
  const rectangular = rows.map(function(row) {
    const copy = row.slice();
    while (copy.length < maxCols) copy.push('');
    return copy;
  });
  output.getRange(1, 1, rectangular.length, maxCols).setValues(rectangular);
  output.getRange(1, 1, 1, 2).setBackground('#d9eaf7').setFontWeight('bold');
  output.setFrozenRows(1);
  for (let c = 1; c <= maxCols; c++) output.setColumnWidth(c, c === 4 || c === 6 || c === 8 || c === 10 || c === 12 || c === 14 ? 180 : 110);
  SpreadsheetApp.flush();

  return {
    version: PR25_R10_VERSION,
    sourceRows: sourceRows.length,
    detailRows: detailRows.length,
    best: combinations[0] || null,
    outputSheet: PR25_R10_OUTPUT_SHEET
  };
}

function pr25r10_scoreCombination_(sourceRows, detailNormalized, cols) {
  const counts = {};
  for (let r = 0; r < sourceRows.length; r++) {
    const row = sourceRows[r];
    const key = [
      pr25r10_normalize_('date', row[cols.date]),
      pr25r10_normalize_('account', row[cols.account]),
      pr25r10_normalize_('orderNo', row[cols.orderNo]),
      pr25r10_normalize_('productNo', row[cols.productNo]),
      pr25r10_normalize_('sales', row[cols.sales])
    ].join('|');
    counts[key] = (counts[key] || 0) + 1;
  }
  let matched = 0;
  for (let i = 0; i < detailNormalized.date.length; i++) {
    const key = [
      detailNormalized.date[i],
      detailNormalized.account[i],
      detailNormalized.orderNo[i],
      detailNormalized.productNo[i],
      detailNormalized.sales[i]
    ].join('|');
    if (counts[key] > 0) {
      counts[key]--;
      matched++;
    }
  }
  return { matched: matched };
}

function pr25r10_find_(headers, names) {
  for (let n = 0; n < names.length; n++) {
    const target = pr25r10_compact_(names[n]);
    for (let i = 0; i < headers.length; i++) {
      if (pr25r10_compact_(headers[i]) === target) return i;
    }
  }
  return -1;
}

function pr25r10_normalize_(field, value) {
  if (field === 'date') {
    const text = String(value == null ? '' : value).trim();
    const match = text.match(/(?:\d{4}[.\/-])?(\d{1,2})[.\/-](\d{1,2})/);
    return match ? ('0' + match[1]).slice(-2) + '/' + ('0' + match[2]).slice(-2) : pr25r10_compact_(text);
  }
  if (field === 'sales') {
    const cleaned = String(value == null ? '' : value).replace(/[원,\s]/g, '');
    const number = Number(cleaned);
    return isNaN(number) ? '' : String(Math.round(number));
  }
  return pr25r10_compact_(value);
}

function pr25r10_compact_(value) {
  return String(value == null ? '' : value).trim().toLowerCase().replace(/[\s._()\[\]{}\-\/]/g, '');
}
