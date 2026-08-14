const fs = require('fs');
const vm = require('vm');
const assert = require('assert');

const source = fs.readFileSync('Patch_v6_48_lightweight_vat_control_tower.gs', 'utf8');
const context = {};
vm.createContext(context);
vm.runInContext(source, context);

function headersWithStatus(name) {
  const headers = Array(29).fill('');
  headers[0] = '마켓주문일자';
  headers[2] = '마켓주문번호';
  headers[3] = '마켓아이디';
  headers[4] = '마켓상품번호';
  headers[6] = '결제금액합계(원)';
  headers[7] = name;
  return headers;
}

function cancelledRow(status) {
  const row = Array(29).fill('');
  row[0] = '2026-06-01';
  row[2] = 'ORDER-1';
  row[3] = 'beliun1021';
  row[4] = 'PRODUCT-1';
  row[6] = 100000;
  row[7] = status;
  row[28] = 80000;
  return row;
}

const marketHeaders = headersWithStatus('마켓주문상태');
const marketIndexes = context.vatHeaderIndexes_v648_(marketHeaders);
assert.equal(marketIndexes.status, 7, '마켓주문상태를 status 헤더로 인식해야 함');

['취소완료', '반품완료', '환불완료'].forEach(status => {
  const result = context.vatDetailRow_v648_(cancelledRow(status), marketIndexes, 2);
  assert.equal(result.row, null, `${status} 행은 VAT 상세에서 제외되어야 함`);
});

const legacyIndexes = context.vatHeaderIndexes_v648_(headersWithStatus('주문상태'));
assert.equal(legacyIndexes.status, 7, '기존 주문상태 alias도 계속 인식해야 함');

const bothHeaders = headersWithStatus('주문상태');
bothHeaders[8] = '마켓주문상태';
const bothIndexes = context.vatHeaderIndexes_v648_(bothHeaders);
assert.equal(bothIndexes.status, 8, '마켓주문상태가 존재하면 기존 주문상태보다 우선해야 함');

console.log('v6.48 market order status alias regression: OK');
