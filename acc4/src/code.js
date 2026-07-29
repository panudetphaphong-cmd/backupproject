const SHEETS = {
  TRANSACTIONS: 'Transactions',
  BUSINESSES: 'Businesses',
  CATEGORIES: 'Categories',
  PAYMENT_METHODS: 'PaymentMethods'
};

const HEADERS = {
  Transactions: ['transaction_id', 'transaction_date', 'business_id', 'transaction_type', 'category', 'description', 'amount', 'payment_method', 'note', 'created_at', 'updated_at', 'status'],
  Businesses: ['business_id', 'business_name', 'short_name', 'status'],
  Categories: ['category_id', 'business_id', 'transaction_type', 'category_name', 'status', 'sort_order'],
  PaymentMethods: ['payment_id', 'payment_name', 'status']
};

function doGet() {
  setupSheets_();
  return HtmlService.createTemplateFromFile('index')
    .evaluate()
    .setTitle('Toki Wash x Elexa')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function getInitialData() {
  setupSheets_();
  return {
    transactions: readObjects_(SHEETS.TRANSACTIONS).filter(row => row.status !== 'Deleted'),
    businesses: readObjects_(SHEETS.BUSINESSES).filter(row => row.status === 'Active'),
    categories: readObjects_(SHEETS.CATEGORIES).filter(row => row.status === 'Active'),
    timeZone: Session.getScriptTimeZone()
  };
}

function saveTransaction(payload) {
  setupSheets_();
  validateTransaction_(payload);
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEETS.TRANSACTIONS);
    const now = new Date();
    const id = String(payload.transaction_id || '').trim();

    if (id) {
      const rowIndex = findRowById_(sheet, id);
      if (rowIndex < 2) throw new Error('ไม่พบรายการที่ต้องการแก้ไข');
      const current = sheet.getRange(rowIndex, 1, 1, HEADERS.Transactions.length).getValues()[0];
      sheet.getRange(rowIndex, 1, 1, HEADERS.Transactions.length).setValues([[
        id,
        parseDate_(payload.transaction_date),
        payload.business_id,
        payload.transaction_type,
        payload.category,
        String(payload.description || '').trim(),
        Number(payload.amount),
        String(payload.payment_method || ''),
        String(payload.note || '').trim(),
        current[9] || now,
        now,
        'Active'
      ]]);
      return { ok: true, id: id, transaction: transactionResponse_(payload, id), message: 'บันทึกการแก้ไขเรียบร้อยแล้ว' };
    }

    const newId = 'TX-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase();
    sheet.appendRow([
      newId,
      parseDate_(payload.transaction_date),
      payload.business_id,
      payload.transaction_type,
      payload.category,
      String(payload.description || '').trim(),
      Number(payload.amount),
      String(payload.payment_method || ''),
      String(payload.note || '').trim(),
      now,
      now,
      'Active'
    ]);
    return { ok: true, id: newId, transaction: transactionResponse_(payload, newId), message: 'เพิ่มรายการเรียบร้อยแล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function saveTransactions(payloads) {
  if (!Array.isArray(payloads) || payloads.length === 0) throw new Error('กรุณาเพิ่มอย่างน้อย 1 รายการ');
  if (payloads.length > 100) throw new Error('บันทึกได้สูงสุดครั้งละ 100 รายการ');
  payloads.forEach(validateTransaction_);
  setupSheets_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEETS.TRANSACTIONS);
    const now = new Date();
    const rows = payloads.map((payload, index) => [
      'TX-' + Utilities.formatDate(now, Session.getScriptTimeZone(), 'yyyyMMdd-HHmmss') + '-' + Utilities.getUuid().slice(0, 4).toUpperCase(),
      parseDate_(payload.transaction_date),
      payload.business_id,
      payload.transaction_type,
      payload.category,
      String(payload.description || '').trim(),
      Number(payload.amount),
      String(payload.payment_method || ''),
      String(payload.note || '').trim(),
      new Date(now.getTime() + index),
      new Date(now.getTime() + index),
      'Active'
    ]);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, HEADERS.Transactions.length).setValues(rows);
    const transactions = payloads.map((payload, index) => transactionResponse_(payload, rows[index][0]));
    return { ok: true, count: rows.length, transactions: transactions, message: 'บันทึก ' + rows.length + ' รายการเรียบร้อยแล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function transactionResponse_(payload, id) {
  return {
    transaction_id: id,
    transaction_date: String(payload.transaction_date),
    business_id: payload.business_id,
    transaction_type: payload.transaction_type,
    category: payload.category,
    description: String(payload.description || '').trim(),
    amount: Number(payload.amount),
    payment_method: '',
    note: String(payload.note || '').trim(),
    status: 'Active'
  };
}

function saveCategory(payload) {
  setupSheets_();
  if (!payload || !['B001', 'B002', 'ALL'].includes(payload.business_id)) throw new Error('กรุณาเลือกธุรกิจ');
  if (!['income', 'expense'].includes(payload.transaction_type)) throw new Error('กรุณาเลือกประเภทหมวดหมู่');
  const name = String(payload.category_name || '').trim();
  if (!name) throw new Error('กรุณากรอกชื่อหมวดหมู่');
  if (name.length > 80) throw new Error('ชื่อหมวดหมู่ยาวเกินไป');
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEETS.CATEGORIES);
    const rows = readObjects_(SHEETS.CATEGORIES);
    const duplicate = rows.some(row =>
      row.status === 'Active' &&
      row.business_id === payload.business_id &&
      row.transaction_type === payload.transaction_type &&
      String(row.category_name).toLowerCase() === name.toLowerCase() &&
      row.category_id !== payload.category_id
    );
    if (duplicate) throw new Error('มีหมวดหมู่นี้อยู่แล้ว');
    if (payload.category_id) {
      const rowIndex = findRowById_(sheet, payload.category_id);
      if (rowIndex < 2) throw new Error('ไม่พบหมวดหมู่ที่ต้องการแก้ไข');
      const currentOrder = sheet.getRange(rowIndex, 6).getValue();
      sheet.getRange(rowIndex, 2, 1, 5).setValues([[
        payload.business_id, payload.transaction_type, name, 'Active', currentOrder || nextCategoryOrder_(payload.business_id, payload.transaction_type)
      ]]);
      return { ok: true, message: 'แก้ไขหมวดหมู่เรียบร้อยแล้ว' };
    }
    const newId = 'C-' + Utilities.getUuid().slice(0, 8).toUpperCase();
    sheet.appendRow([newId, payload.business_id, payload.transaction_type, name, 'Active', nextCategoryOrder_(payload.business_id, payload.transaction_type)]);
    return { ok: true, id: newId, message: 'เพิ่มหมวดหมู่เรียบร้อยแล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function reorderCategories(payload) {
  if (!payload || !Array.isArray(payload.category_ids) || !payload.category_ids.length) throw new Error('ไม่พบลำดับหมวดหมู่');
  setupSheets_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEETS.CATEGORIES);
    const rowCount = sheet.getLastRow() - 1;
    if (rowCount < 1) return { ok: true };
    const ids = sheet.getRange(2, 1, rowCount, 1).getDisplayValues();
    const orders = sheet.getRange(2, 6, rowCount, 1).getValues();
    const orderMap = new Map(payload.category_ids.map((id, index) => [String(id), index + 1]));
    ids.forEach((row, index) => {
      if (orderMap.has(row[0])) orders[index][0] = orderMap.get(row[0]);
    });
    sheet.getRange(2, 6, rowCount, 1).setValues(orders);
    return { ok: true, message: 'บันทึกลำดับหมวดหมู่แล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function nextCategoryOrder_(businessId, transactionType) {
  const rows = readObjects_(SHEETS.CATEGORIES).filter(row =>
    row.business_id === businessId && row.transaction_type === transactionType
  );
  return rows.reduce((max, row) => Math.max(max, Number(row.sort_order) || 0), 0) + 1;
}

function deleteCategory(categoryId) {
  if (!categoryId) throw new Error('ไม่พบรหัสหมวดหมู่');
  setupSheets_();
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEETS.CATEGORIES);
    const rowIndex = findRowById_(sheet, categoryId);
    if (rowIndex < 2) throw new Error('ไม่พบหมวดหมู่ที่ต้องการลบ');
    sheet.getRange(rowIndex, 5).setValue('Deleted');
    return { ok: true, message: 'ลบหมวดหมู่เรียบร้อยแล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaction(transactionId) {
  if (!transactionId) throw new Error('ไม่พบรหัสรายการ');
  const lock = LockService.getDocumentLock();
  lock.waitLock(10000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEETS.TRANSACTIONS);
    const rowIndex = findRowById_(sheet, transactionId);
    if (rowIndex < 2) throw new Error('ไม่พบรายการที่ต้องการลบ');
    sheet.getRange(rowIndex, 11).setValue(new Date());
    sheet.getRange(rowIndex, 12).setValue('Deleted');
    return { ok: true, message: 'ลบรายการเรียบร้อยแล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function deleteTransactions(transactionIds) {
  if (!Array.isArray(transactionIds) || transactionIds.length === 0) throw new Error('กรุณาเลือกรายการที่ต้องการลบ');
  if (transactionIds.length > 500) throw new Error('ลบได้สูงสุดครั้งละ 500 รายการ');
  setupSheets_();
  const uniqueIds = [...new Set(transactionIds.map(String).filter(Boolean))];
  const lock = LockService.getDocumentLock();
  lock.waitLock(15000);
  try {
    const sheet = getSpreadsheet_().getSheetByName(SHEETS.TRANSACTIONS);
    if (sheet.getLastRow() < 2) throw new Error('ไม่พบรายการบัญชี');
    const rowCount = sheet.getLastRow() - 1;
    const ids = sheet.getRange(2, 1, rowCount, 1).getDisplayValues();
    const updatedAt = sheet.getRange(2, 11, rowCount, 1).getValues();
    const statuses = sheet.getRange(2, 12, rowCount, 1).getValues();
    const targets = new Set(uniqueIds);
    let deletedCount = 0;
    ids.forEach((row, index) => {
      if (targets.has(row[0]) && statuses[index][0] !== 'Deleted') {
        updatedAt[index][0] = new Date();
        statuses[index][0] = 'Deleted';
        deletedCount++;
      }
    });
    if (!deletedCount) throw new Error('ไม่พบรายการที่ต้องการลบ');
    sheet.getRange(2, 11, rowCount, 1).setValues(updatedAt);
    sheet.getRange(2, 12, rowCount, 1).setValues(statuses);
    return { ok: true, count: deletedCount, message: 'ลบ ' + deletedCount + ' รายการเรียบร้อยแล้ว' };
  } finally {
    lock.releaseLock();
  }
}

function setupSheets_() {
  const properties = PropertiesService.getScriptProperties();
  const schemaKey = 'ACCOUNTING_SCHEMA_READY_V3';
  if (properties.getProperty(schemaKey) === 'true') return;
  const ss = getSpreadsheet_();
  Object.keys(HEADERS).forEach(name => {
    let sheet = ss.getSheetByName(name);
    if (!sheet) sheet = ss.insertSheet(name);
    const headers = HEADERS[name];
    if (sheet.getLastRow() === 0) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length)
        .setBackground('#111111')
        .setFontColor('#FFD600')
        .setFontWeight('bold');
    } else {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }
  });

  seedIfEmpty_(SHEETS.BUSINESSES, [
    ['B001', 'ร้านซักผ้า Toki Wash', 'Toki Wash', 'Active'],
    ['B002', 'สถานีชาร์จรถยนต์ไฟฟ้า Elexa', 'Elexa', 'Active']
  ]);
  seedIfEmpty_(SHEETS.CATEGORIES, [
    ['C001', 'B001', 'income', 'ค่าซักผ้า', 'Active'],
    ['C002', 'B001', 'income', 'ค่าอบผ้า', 'Active'],
    ['C003', 'B001', 'income', 'จำหน่ายสินค้า/น้ำยา', 'Active'],
    ['C004', 'B001', 'expense', 'ค่าน้ำ', 'Active'],
    ['C005', 'B001', 'expense', 'ค่าไฟ', 'Active'],
    ['C006', 'B001', 'expense', 'ค่าซ่อมบำรุง', 'Active'],
    ['C007', 'B002', 'income', 'ค่าบริการชาร์จ', 'Active'],
    ['C008', 'B002', 'expense', 'ค่าไฟสถานีชาร์จ', 'Active'],
    ['C009', 'B002', 'expense', 'ค่าบำรุงรักษา', 'Active'],
    ['C010', 'ALL', 'expense', 'ค่าเช่า', 'Active'],
    ['C011', 'ALL', 'expense', 'เงินเดือน', 'Active'],
    ['C012', 'ALL', 'expense', 'ค่าใช้จ่ายอื่น ๆ', 'Active'],
    ['C013', 'ALL', 'income', 'รายรับอื่น ๆ', 'Active']
  ]);
  seedIfEmpty_(SHEETS.PAYMENT_METHODS, [
    ['P001', 'เงินสด', 'Active'],
    ['P002', 'เงินโอน / QR', 'Active'],
    ['P003', 'บัตรเครดิต / เดบิต', 'Active'],
    ['P004', 'ระบบรับชำระ Elexa', 'Active'],
    ['P005', 'อื่น ๆ', 'Active']
  ]);

  const tx = ss.getSheetByName(SHEETS.TRANSACTIONS);
  tx.getRange(2, 2, Math.max(tx.getMaxRows() - 1, 1), 1).setNumberFormat('dd/mm/yyyy');
  tx.getRange(2, 7, Math.max(tx.getMaxRows() - 1, 1), 1).setNumberFormat('#,##0.00');
  fixIncorrectTransactionYearsOnce_(tx);
  properties.setProperty(schemaKey, 'true');
}

function fixIncorrectTransactionYearsOnce_(sheet) {
  const propertyKey = 'MIGRATION_FIX_TRANSACTION_YEARS_2026_2027_V1';
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty(propertyKey)) return;
  const lastRow = sheet.getLastRow();
  let changedCount = 0;
  if (lastRow >= 2) {
    const range = sheet.getRange(2, 2, lastRow - 1, 1);
    const dates = range.getValues();
    dates.forEach(row => {
      const value = row[0];
      if (!(value instanceof Date) || Number.isNaN(value.getTime())) return;
      const year = value.getFullYear();
      if (year === 2026 || year === 2027) {
        const corrected = new Date(value);
        corrected.setFullYear(year - 2);
        row[0] = corrected;
        changedCount++;
      }
    });
    if (changedCount) range.setValues(dates);
  }
  properties.setProperty(propertyKey, JSON.stringify({
    completed_at: new Date().toISOString(),
    changed_count: changedCount
  }));
}

function getSpreadsheet_() {
  const active = SpreadsheetApp.getActiveSpreadsheet();
  if (active) return active;
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('กรุณาผูก Apps Script กับ Google Sheet หรือตั้งค่า SPREADSHEET_ID ใน Script Properties');
  return SpreadsheetApp.openById(id);
}

function seedIfEmpty_(name, values) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  if (sheet.getLastRow() === 1 && values.length) {
    sheet.getRange(2, 1, values.length, values[0].length).setValues(values);
  }
}

function readObjects_(name) {
  const sheet = getSpreadsheet_().getSheetByName(name);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return [];
  const headers = values[0];
  return values.slice(1).filter(row => row[0] !== '').map(row => {
    const item = {};
    headers.forEach((header, index) => {
      const value = row[index];
      item[header] = value instanceof Date
        ? Utilities.formatDate(value, Session.getScriptTimeZone(), header === 'transaction_date' ? 'yyyy-MM-dd' : "yyyy-MM-dd'T'HH:mm:ss")
        : value;
    });
    return item;
  });
}

function findRowById_(sheet, id) {
  if (sheet.getLastRow() < 2) return -1;
  const ids = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues();
  const index = ids.findIndex(row => row[0] === id);
  return index === -1 ? -1 : index + 2;
}

function parseDate_(value) {
  const parts = String(value || '').split('-').map(Number);
  if (parts.length !== 3 || !parts[0] || !parts[1] || !parts[2]) throw new Error('รูปแบบวันที่ไม่ถูกต้อง');
  return new Date(parts[0], parts[1] - 1, parts[2]);
}

function validateTransaction_(payload) {
  if (!payload) throw new Error('ไม่พบข้อมูลรายการ');
  if (!payload.transaction_date) throw new Error('กรุณาเลือกวันที่');
  if (!['B001', 'B002'].includes(payload.business_id)) throw new Error('กรุณาเลือกร้าน');
  if (!['income', 'expense'].includes(payload.transaction_type)) throw new Error('กรุณาเลือกประเภทรายการ');
  if (!String(payload.category || '').trim()) throw new Error('กรุณาเลือกหมวดหมู่');
  if (!Number.isFinite(Number(payload.amount)) || Number(payload.amount) <= 0) throw new Error('ยอดเงินต้องมากกว่า 0');
}
