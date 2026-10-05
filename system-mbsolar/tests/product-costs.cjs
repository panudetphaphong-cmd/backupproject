const fs = require('fs');
const vm = require('vm');
const assert = require('assert/strict');

const backend = {};
vm.createContext(backend);
vm.runInContext(fs.readFileSync('src/Code.js', 'utf8'), backend);
backend.APP = vm.runInContext('APP', backend);

// In-memory tables mock
const tables = {
  ProductCosts: [],
  Expenses: [],
  Projects: [{ id: 'P01', name: 'Solar 5kW Rooftop' }]
};

backend.db_ = () => ({});
backend.readSheet_ = (_, name) => tables[name] || [];
backend.rows_ = (name) => tables[name] || [];
backend.clean_ = (s) => (s == null ? '' : String(s).trim());
backend.positive_ = (v, label) => {
  const n = Number(v);
  if (isNaN(n) || n <= 0) throw new Error((label || 'จำนวนเงิน') + ' ต้องมากกว่า 0');
  return n;
};
backend.num_ = (v) => {
  const n = Number(v);
  return isNaN(n) ? 0 : n;
};
backend.date_ = (d) => (d ? new Date(d).toISOString().slice(0, 10) : '2026-10-04');
backend.id_ = (prefix) => prefix + '_' + Math.random().toString(36).slice(2, 9);
backend.require_ = () => true;
backend.audit_ = () => {};
backend.project_ = (id) => {
  const p = tables.Projects.find(x => x.id === id);
  if (!p) throw new Error('ไม่พบโปรเจกต์ ' + id);
  return p;
};
backend.businessPayload_ = () => ({ partial: true, expenses: tables.Expenses });

backend.append_ = (sheetName, rowArray) => {
  const cols = backend.APP.sheets[sheetName];
  const obj = {};
  cols.forEach((col, idx) => {
    obj[col] = rowArray[idx];
  });
  tables[sheetName].push(obj);
};

backend.updateRow_ = (sheetName, id, changes) => {
  const row = tables[sheetName].find(x => x.id === id);
  if (row) Object.assign(row, changes);
};

backend.deleteRow_ = (sheetName, id) => {
  const idx = tables[sheetName].findIndex(x => x.id === id);
  if (idx !== -1) tables[sheetName].splice(idx, 1);
};

const user = { id: 'U1', name: 'Tester' };

// 1. Test create product cost
const res1 = backend.saveProductCost_(user, {
  name: 'แผงโซลาร์เซลล์ Tier 1 550W N-Type',
  category: 'Solar Cell',
  brand: 'Jinko Solar',
  supplier: 'ไทวัสดุ',
  unit: 'แผง',
  unitPrice: 3850
});

assert.equal(res1.partial, true);
assert.equal(res1.productCosts.length, 1);
assert.equal(res1.productCosts[0].name, 'แผงโซลาร์เซลล์ Tier 1 550W N-Type');
assert.equal(res1.productCosts[0].supplier, 'ไทวัสดุ');
assert.equal(res1.productCosts[0].unitPrice, 3850);
assert.equal(res1.productCosts[0].minPrice, 3850);
assert.equal(res1.productCosts[0].maxPrice, 3850);
console.log('PASS: saveProductCost_ create successful');

// 2. Add another supplier for the same product with a different price
const res2 = backend.saveProductCost_(user, {
  name: 'แผงโซลาร์เซลล์ Tier 1 550W N-Type',
  category: 'Solar Cell',
  brand: 'Jinko Solar',
  supplier: 'ร้านแสงทองการไฟฟ้า',
  unit: 'แผง',
  unitPrice: 3600
});

assert.equal(res2.productCosts.length, 2);
const sangthong = res2.productCosts.find(x => x.supplier === 'ร้านแสงทองการไฟฟ้า');
const thaiwatsadu = res2.productCosts.find(x => x.supplier === 'ไทวัสดุ');
assert.equal(sangthong.unitPrice, 3600);
assert.equal(thaiwatsadu.unitPrice, 3850);
// Both should have minPrice 3600 and maxPrice 3850 for comparison
assert.equal(sangthong.minPrice, 3600);
assert.equal(sangthong.maxPrice, 3850);
assert.equal(thaiwatsadu.minPrice, 3600);
assert.equal(thaiwatsadu.maxPrice, 3850);
console.log('PASS: productCostRows_ min/max price comparison across suppliers calculated correctly');

// 3. Test edit product cost
const idToEdit = sangthong.id;
const res3 = backend.saveProductCost_(user, {
  id: idToEdit,
  name: 'แผงโซลาร์เซลล์ Tier 1 550W N-Type',
  category: 'Solar Cell',
  brand: 'Jinko Solar',
  supplier: 'ร้านแสงทองการไฟฟ้า (โปรโมชั่น)',
  unit: 'แผง',
  unitPrice: 3500
});

const updated = res3.productCosts.find(x => x.id === idToEdit);
assert.equal(updated.supplier, 'ร้านแสงทองการไฟฟ้า (โปรโมชั่น)');
assert.equal(updated.unitPrice, 3500);
assert.equal(updated.minPrice, 3500);
console.log('PASS: saveProductCost_ edit successful');

// 4. Test saveCostEstimateToExpenses_
const resExp = backend.saveCostEstimateToExpenses_(user, {
  projectId: 'P01',
  date: '2026-10-04',
  items: [
    {
      name: 'แผงโซลาร์เซลล์ Tier 1 550W N-Type',
      category: 'Solar Cell',
      supplier: 'ร้านแสงทองการไฟฟ้า (โปรโมชั่น)',
      quantity: 10,
      unit: 'แผง',
      unitPrice: 3500,
      amount: 35000
    },
    {
      name: 'Inverter 5kW Grid-Tied On-Grid',
      category: 'Solar Cell',
      supplier: 'ตัวแทนจำหน่ายตรง Huawei/Solis',
      quantity: 1,
      unit: 'เครื่อง',
      unitPrice: 32000,
      amount: 32000
    }
  ]
});

assert.equal(resExp.partial, true);
assert.equal(tables.Expenses.length, 2);
assert.equal(tables.Expenses[0].projectId, 'P01');
assert.equal(tables.Expenses[0].amountExVat, 35000);
assert.equal(tables.Expenses[1].amountExVat, 32000);
console.log('PASS: saveCostEstimateToExpenses_ recorded 2 expense rows properly');

// 5. Test deleteProductCost_
const resDel = backend.deleteProductCost_(user, { id: idToEdit });
assert.equal(resDel.productCosts.length, 1);
assert.equal(resDel.productCosts[0].supplier, 'ไทวัสดุ');
console.log('PASS: deleteProductCost_ removed item successfully');

console.log('ALL PRODUCT COST TESTS PASSED 100%!');
