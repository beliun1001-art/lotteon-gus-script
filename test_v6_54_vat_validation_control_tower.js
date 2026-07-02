const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const source = fs.readFileSync('Patch_v6_54_vat_validation_control_tower.gs', 'utf8');
assert.doesNotMatch(source, /read(?:TableAmounts|DashboardSummaryAmounts)_v628_/);

const context = {};
vm.createContext(context);
vm.runInContext(source, context);

function dataSheet(values) {
  return {
    getLastRow: () => values.length,
    getDataRange: () => ({ getValues: () => values })
  };
}

const sheets = {
  '대시보드': dataSheet([
    ['구분', '항목', '값'],
    ['요약', '순수매출액', '65,805,300'],
    ['요약', '매입금액', '47,401,952'],
    ['요약', '정산기준금액', '58,635,409']
  ]),
  '브랜드별_마진율': dataSheet([
    ['브랜드명', '순수매출액', '정산기준금액', '매입금액'],
    ['A', 40000000, 35000000, 28000000],
    ['B', 25805300, 23635409, 19401952]
  ]),
  '부가세_신고자료': dataSheet([
    ['날짜', '순수매출액', '정산기준금액', '매입금액'],
    ['07/01', 40000000, 35000000, 28000000],
    ['07/02', 25805300, 23635409, 19401952]
  ]),
  '부가세_상품별': dataSheet([
    ['상품번호', '순수매출액', '정산기준금액', '매입금액'],
    ['P1', 40000000, 35000000, 28000000],
    ['P2', 25805300, 23635409, 19401952]
  ])
};

let validationRows = null;
const validationSheet = {
  clearContents() {},
  getRange(row, _column, count) {
    return {
      setValues(values) { if (row === 2 && count === 4) validationRows = values; return this; },
      setBackground() { return this; },
      setFontWeight() { return this; }
    };
  },
  setFrozenRows() {}
};
sheets['시트별_금액검증'] = validationSheet;

context.buildSheetAmountValidation_v654_({
  getSheetByName: name => sheets[name] || null,
  insertSheet: () => validationSheet
});

assert.deepStrictEqual(Array.from(validationRows, row => row[0]), [
  '대시보드', '브랜드별_마진율', '부가세_신고자료', '부가세_상품별'
]);
assert.ok(validationRows.every(row => row[10] === 'OK'));
const productRow = validationRows.find(row => row[0] === '부가세_상품별');
assert.equal(productRow[8], 58635409, '부가세_상품별 정산기준금액 합계를 읽어야 함');
assert.equal(productRow[9], 0, '부가세_상품별 정산 차이는 0이어야 함');

console.log('v6.54 buildSheetAmountValidation standalone mock: OK (4/4 sheets, product settlement=58,635,409)');
