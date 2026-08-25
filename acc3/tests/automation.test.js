const assert = require('assert');
const fs = require('fs');
const vm = require('vm');

let storedData;
let writes = 0;
const context = {
  console,
  Date,
  JSON,
  Math,
  Number,
  String,
  Array,
  Object,
  isNaN,
  LockService: {
    getScriptLock: () => ({ waitLock() {}, hasLock: () => true, releaseLock() {} })
  },
  Utilities: {
    formatDate: date => date.toISOString().slice(0, 10)
  },
  getData: () => JSON.parse(JSON.stringify(storedData)),
  writeSharedData_: data => {
    storedData = JSON.parse(JSON.stringify(data));
    writes += 1;
  }
};
vm.createContext(context);
vm.runInContext(fs.readFileSync('src/code.js', 'utf8'), context);
context.getData = () => JSON.parse(JSON.stringify(storedData));
context.writeSharedData_ = data => {
  storedData = JSON.parse(JSON.stringify(data));
  writes += 1;
};

function baseData(values, invoiceDate) {
  return {
    columns: [
      { id: 'eday', label: 'EDAY' },
      { id: 'sm', label: 'SM' },
      { id: 'withdraw', label: 'ยอดเบิก' }
    ],
    customers: [{
      id: 9,
      name: 'ยุทธ',
      values,
      invoiceSent: true,
      invoiceDate,
      status: 'รอชำระ',
      invoiceHistory: []
    }],
    chatMessages: [],
    transactions: []
  };
}

storedData = baseData({ eday: 1000, sm: 500, withdraw: 200 }, '18/08/2026');
writes = 0;
const preview = context.processWeeklyOutstandingBalances({
  dryRun: true,
  customerName: 'ยุทธ',
  now: '2026-08-25T12:00:00+07:00'
});
assert.strictEqual(preview.success, true);
assert.strictEqual(preview.processed, 1);
assert.strictEqual(preview.results[0].newWithdraw, 1700);
assert.strictEqual(writes, 0, 'dry run must not save data');

const carried = context.processWeeklyOutstandingBalances({ now: '2026-08-25T12:00:00+07:00' });
assert.strictEqual(carried.processed, 1);
assert.deepStrictEqual(storedData.customers[0].values, { eday: 0, sm: 0, withdraw: 1700 });
assert.match(storedData.chatMessages[0].text, /EDAY: 1,000 บาท/);
assert.match(storedData.chatMessages[0].text, /SM: 500 บาท/);

const repeated = context.processWeeklyOutstandingBalances({ now: '2026-09-02T12:00:00+07:00' });
assert.strictEqual(repeated.processed, 0, 'carried balance must not be carried twice');

storedData = baseData({ eday: -800, sm: -200, withdraw: 0 }, '25 สิงหาคม 2569');
storedData.customers[0].attachedInvoiceImg = 'data:image/jpeg;base64,TEST';
writes = 0;
const refunded = context.completeCustomerRefund(9, {
  bank: { bankName: 'KBANK', accNo: '123', accName: 'ยุทธ' },
  colLabel: 'EDAY'
});
assert.strictEqual(refunded.success, true);
assert.strictEqual(refunded.refundAmount, 1000);
assert.deepStrictEqual(storedData.customers[0].values, { eday: 0, sm: 0, withdraw: 0 });
assert.strictEqual(storedData.customers[0].status, 'ชำระแล้ว');
assert.strictEqual(storedData.transactions[0].amount, 1000);
assert.strictEqual(storedData.chatMessages[0].image, 'data:image/jpeg;base64,TEST');
assert.match(storedData.chatMessages[0].text, /เคลียร์ยอดในตารางเป็น 0 แล้ว/);
assert.strictEqual(writes, 1);

console.log('automation tests passed');
