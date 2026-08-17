const SHEET_NAME = 'รายการบัญชี';
const HEADERS = ['รหัส', 'วันที่', 'ประเภท', 'หมวดหมู่', 'รายละเอียด', 'จำนวนเงิน', 'ผู้บันทึก', 'สถานะ', 'สร้างเมื่อ', 'แก้ไขเมื่อ', 'บัญชี', 'ลิงก์สลิป', 'เลขอ้างอิงสลิป', 'รหัสธุรกิจ'];
const CATEGORY_SHEET = 'หมวดหมู่';
const BUSINESS_SHEET = 'ธุรกิจ';
const SAVINGS_SHEET = 'เงินเก็บสะสม';
const DEFAULT_CATEGORIES = [
  ['รายรับ', 'เงินเดือน'], ['รายรับ', 'รายได้เสริม'], ['รายรับ', 'เงินรับโอน'], ['รายรับ', 'รายรับอื่น ๆ'],
  ['รายจ่าย', 'อาหาร'], ['รายจ่าย', 'เดินทาง'], ['รายจ่าย', 'บ้าน'], ['รายจ่าย', 'สาธารณูปโภค'],
  ['รายจ่าย', 'สุขภาพ'], ['รายจ่าย', 'การศึกษา'], ['รายจ่าย', 'ช้อปปิ้ง'], ['รายจ่าย', 'อื่น ๆ']
];
const BUSINESS_CATEGORIES = [
  ['รายรับ', 'ขายสินค้า'], ['รายรับ', 'ค่าบริการ'], ['รายรับ', 'รายได้ธุรกิจอื่น ๆ'],
  ['รายจ่าย', 'ต้นทุนสินค้า'], ['รายจ่าย', 'วัตถุดิบ'], ['รายจ่าย', 'ค่าขนส่ง'], ['รายจ่าย', 'ค่าโฆษณา'],
  ['รายจ่าย', 'ค่าเช่า'], ['รายจ่าย', 'ค่าน้ำค่าไฟ'], ['รายจ่าย', 'เงินเดือน'], ['รายจ่าย', 'ค่าใช้จ่ายธุรกิจอื่น ๆ']
];

function doGet() {
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('AiHouse — บัญชีบ้านของเรา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  const cache = CacheService.getScriptCache();
  const cached = cache.get('initial-data-v3');
  if (cached) return JSON.parse(cached);
  migrateBusinessData_();
  const data = { records: getRecords_(), categories: getCategories_(), businesses: getBusinesses_(), savings: getSavings_() };
  try { cache.put('initial-data-v3', JSON.stringify(data), 120); } catch (ignore) {}
  return data;
}

function invalidateDataCache_() {
  try { CacheService.getScriptCache().removeAll(['initial-data-v2', 'initial-data-v3']); } catch (ignore) {}
}

function addSavingsEntry(data) {
  const amount = Number(data && data.amount);
  const action = data && data.action === 'WITHDRAW' ? 'WITHDRAW' : 'DEPOSIT';
  if (!amount || amount <= 0) throw new Error('กรุณาใส่จำนวนเงินมากกว่า 0');
  const member = 'ส่วนกลาง';
  const row = [Utilities.getUuid(), parseDate_(data && data.date), member, action, amount, String(data && data.note || '').trim(), 'ใช้งาน', new Date()];
  ensureSavingsSheet_().appendRow(row);
  invalidateDataCache_();
  return { ok: true, message: action === 'DEPOSIT' ? 'เพิ่มเงินเก็บสะสมแล้ว' : 'บันทึกการถอนเงินเก็บแล้ว' };
}

function deleteSavingsEntry(id) {
  const sheet = ensureSavingsSheet_(), values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === String(id) && values[i][6] === 'ใช้งาน') {
    sheet.getRange(i + 1, 7).setValue('ลบแล้ว');
    invalidateDataCache_();
    return { ok: true, message: 'ลบรายการเงินเก็บแล้ว' };
  }
  throw new Error('ไม่พบรายการเงินเก็บ');
}

function addBusiness(name) {
  name = String(name || '').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อธุรกิจ');
  if (getBusinesses_().some(b => b.name.toLowerCase() === name.toLowerCase())) throw new Error('มีธุรกิจชื่อนี้อยู่แล้ว');
  const id = Utilities.getUuid();
  ensureBusinessSheet_().appendRow([id, name, 'ใช้งาน', new Date(), new Date()]);
  const categorySheet = ensureCategorySheet_();
  categorySheet.getRange(categorySheet.getLastRow() + 1, 1, BUSINESS_CATEGORIES.length, 7).setValues(BUSINESS_CATEGORIES.map(c => [Utilities.getUuid(), c[0], c[1], 'ใช้งาน', new Date(), 'BUSINESS', id]));
  invalidateDataCache_();
  return { ok: true, message: 'เพิ่มธุรกิจแล้ว', business: { id, name } };
}

function updateBusiness(id, name) {
  name = String(name || '').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อธุรกิจ');
  const sheet = ensureBusinessSheet_(), values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === String(id) && values[i][2] === 'ใช้งาน') {
    sheet.getRange(i + 1, 2).setValue(name);
    sheet.getRange(i + 1, 5).setValue(new Date());
    invalidateDataCache_();
    return { ok: true, message: 'แก้ชื่อธุรกิจแล้ว' };
  }
  throw new Error('ไม่พบธุรกิจ');
}

function deleteBusiness(id) {
  if (getRecords_().some(r => r.businessId === String(id))) throw new Error('ธุรกิจนี้มีรายการอยู่ จึงยังลบไม่ได้');
  const sheet = ensureBusinessSheet_(), values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === String(id) && values[i][2] === 'ใช้งาน') {
    sheet.getRange(i + 1, 3).setValue('ลบแล้ว');
    sheet.getRange(i + 1, 5).setValue(new Date());
    invalidateDataCache_();
    return { ok: true, message: 'ลบธุรกิจแล้ว' };
  }
  throw new Error('ไม่พบธุรกิจ');
}

function addCategory(type, name, workspace, businessId) {
  workspace = normalizeWorkspace_(workspace);
  if (!['รายรับ', 'รายจ่าย'].includes(type)) throw new Error('ประเภทหมวดหมู่ไม่ถูกต้อง');
  name = String(name || '').trim();
  if (!name) throw new Error('กรุณาใส่ชื่อหมวดหมู่');
  businessId = workspace === 'BUSINESS' ? normalizeBusinessId_(businessId) : '';
  if (getCategories_().some(c => c.workspace === workspace && c.businessId === businessId && c.type === type && c.name.toLowerCase() === name.toLowerCase())) throw new Error('มีหมวดหมู่นี้อยู่แล้ว');
  ensureCategorySheet_().appendRow([Utilities.getUuid(), type, name, 'ใช้งาน', new Date(), workspace, businessId]);
  invalidateDataCache_();
  return { ok: true, message: 'เพิ่มหมวดหมู่แล้วนะ' };
}

function updateCategory(id, type, name, workspace, businessId) {
  const sheet = ensureCategorySheet_(), values = sheet.getDataRange().getValues();
  workspace = normalizeWorkspace_(workspace);
  businessId = workspace === 'BUSINESS' ? normalizeBusinessId_(businessId) : '';
  name = String(name || '').trim();
  if (!name || !['รายรับ', 'รายจ่าย'].includes(type)) throw new Error('ข้อมูลหมวดหมู่ไม่ครบ');
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === String(id)) {
    sheet.getRange(i + 1, 2, 1, 6).setValues([[type, name, 'ใช้งาน', new Date(), workspace, businessId]]);
    invalidateDataCache_();
    return { ok: true, message: 'แก้ไขหมวดหมู่แล้วนะ' };
  }
  throw new Error('ไม่พบหมวดหมู่');
}

function deleteCategory(id) {
  const sheet = ensureCategorySheet_(), values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) if (String(values[i][0]) === String(id)) {
    sheet.getRange(i + 1, 4).setValue('ลบแล้ว');
    sheet.getRange(i + 1, 5).setValue(new Date());
    invalidateDataCache_();
    return { ok: true, message: 'นำหมวดหมู่ออกจากตัวเลือกแล้ว' };
  }
  throw new Error('ไม่พบหมวดหมู่');
}

function addRecord(data) {
  validateRecord_(data);
  const sheet = ensureSheet_();
  const now = new Date();
  const workspace = normalizeWorkspace_(data.workspace);
  const row = [Utilities.getUuid(), parseDate_(data.date), data.type, data.category || 'อื่น ๆ', data.note || '', Number(data.amount), data.user || 'ส่วนกลาง', 'ใช้งาน', now, now, workspace, data.receiptUrl || '', normalizeSlipRef_(data.slipRef), workspace === 'BUSINESS' ? normalizeBusinessId_(data.businessId) : ''];
  sheet.appendRow(row);
  invalidateDataCache_();
  return { ok: true, message: 'เก็บให้เรียบร้อยแล้วนะ', record: rowToRecord_(row) };
}

function analyzeSlip(payload, workspace, businessId) {
  if (!payload || !payload.base64) throw new Error('ไม่พบรูปสลิป กรุณาเลือกรูปอีกครั้ง');
  const bytes = Utilities.base64Decode(payload.base64);
  if (bytes.length > 8 * 1024 * 1024) throw new Error('รูปใหญ่เกิน 8 MB กรุณาลดขนาดรูปก่อนนะ');
  const mime = payload.mimeType || 'image/jpeg';
  const safeName = String(payload.name || 'slip.jpg').replace(/[^ก-๙a-zA-Z0-9._-]/g, '_');
  const blob = Utilities.newBlob(bytes, mime, safeName);

  let originalUrl = '';
  try {
    const original = getReceiptFolder_().createFile(blob).setName(Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyyMMdd-HHmmss') + '-' + safeName);
    originalUrl = original.getUrl();
  } catch (err) {
    console.error('Drive save error: ' + err.message);
  }

  let text = '';
  try {
    if (typeof Drive !== 'undefined' && Drive.Files) {
      const ocrFile = Drive.Files.create({ name: 'OCR-' + safeName, mimeType: MimeType.GOOGLE_DOCS }, blob, { ocrLanguage: 'th', fields: 'id' });
      text = readOcrDocument_(ocrFile.id);
      try { DriveApp.getFileById(ocrFile.id).setTrashed(true); } catch (ignore) {}
    }
  } catch (e) {
    console.error('OCR error: ' + e.message);
  }

  const parsed = parseSlipText_(text);
  if (normalizeWorkspace_(workspace) === 'BUSINESS') parsed.category = categorizeBusiness_(text + ' ' + parsed.note, parsed.type);
  parsed.receiptUrl = originalUrl;
  parsed.businessId = normalizeWorkspace_(workspace) === 'BUSINESS' ? normalizeBusinessId_(businessId) : '';
  const duplicate = findDuplicateSlip_(parsed, workspace, parsed.businessId);
  return {
    ok: true,
    data: parsed,
    duplicate: duplicate || null,
    ocrWarning: text.trim() ? '' : 'อ่านข้อความบนสลิปไม่สำเร็จ กรุณาตรวจและกรอกข้อมูลเอง',
    missing: [!parsed.type && 'ประเภท', !parsed.note && 'รายละเอียด', !parsed.amount && 'จำนวนเงิน', !parsed.date && 'วันที่'].filter(Boolean)
  };
}

function readOcrDocument_(id) {
  let lastError;
  for (let i = 0; i < 6; i++) {
    try {
      const text = DocumentApp.openById(id).getBody().getText();
      if (text && text.trim()) return text;
    } catch (e) { lastError = e; }
    Utilities.sleep(500);
  }
  if (lastError) throw lastError;
  return '';
}

function saveSlipRecord(data) {
  data.note = String(data.note || '').trim();
  const duplicate = findDuplicateSlip_(data, data.workspace, data.businessId);
  if (duplicate) throw new Error('สลิปนี้เคยบันทึกแล้ว: ' + duplicate.note + ' ' + duplicate.amount.toLocaleString('th-TH') + ' บาท');
  return addRecord(data);
}

function updateRecord(id, changes) {
  const sheet = ensureSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id) && values[i][7] === 'ใช้งาน') {
      const current = rowToRecord_(values[i]);
      const cleanChanges = Object.keys(changes || {}).reduce((out, key) => {
        if (changes[key] !== undefined && changes[key] !== null && changes[key] !== '') out[key] = changes[key];
        return out;
      }, {});
      const next = Object.assign({}, current, cleanChanges);
      validateRecord_(next);
      sheet.getRange(i + 1, 2, 1, 9).setValues([[
        parseDate_(next.date), next.type, next.category || 'อื่น ๆ', next.note || '', Number(next.amount),
        next.user || current.user, 'ใช้งาน', values[i][8] || new Date(), new Date()
      ]]);
      sheet.getRange(i + 1, 11).setValue(normalizeWorkspace_(next.workspace || current.workspace));
      sheet.getRange(i + 1, 14).setValue(normalizeWorkspace_(next.workspace || current.workspace) === 'BUSINESS' ? normalizeBusinessId_(next.businessId || current.businessId) : '');
      invalidateDataCache_();
      return { ok: true, message: 'แก้ไขให้แล้วนะ' };
    }
  }
  throw new Error('ไม่พบรายการที่ต้องการแก้ไข');
}

function deleteRecord(id) {
  const sheet = ensureSheet_();
  const values = sheet.getDataRange().getValues();
  for (let i = 1; i < values.length; i++) {
    if (String(values[i][0]) === String(id) && values[i][7] === 'ใช้งาน') {
      sheet.getRange(i + 1, 8).setValue('ลบแล้ว');
      sheet.getRange(i + 1, 10).setValue(new Date());
      invalidateDataCache_();
      return { ok: true, message: 'ย้ายไปถังขยะแล้ว กู้คืนได้เสมอนะ' };
    }
  }
  throw new Error('ไม่พบรายการที่ต้องการลบ');
}

function processCommand(text, member, workspace, businessId) {
  const command = parseCommand_(String(text || '').trim());
  command.member = normalizeMember_(member);
  command.workspace = normalizeWorkspace_(workspace);
  command.businessId = command.workspace === 'BUSINESS' ? normalizeBusinessId_(businessId) : '';
  if (command.workspace === 'BUSINESS') {
    if (command.action === 'add') command.data.category = categorizeBusiness_(command.text, command.data.type);
    if (command.action === 'query') command.keyword = categorizeBusiness_(command.text, command.type);
  }
  if (!command.text) throw new Error('ลองบอกน้องเฮาส์อีกครั้งนะ');
  if (command.action === 'query') return answerQuery_(command);
  if (command.action === 'add') { command.data.user = command.member; command.data.workspace = command.workspace; command.data.businessId = command.businessId; return addRecord(command.data); }

  const matches = findMatches_(command);
  if (!matches.length) throw new Error('ยังหารายการที่พูดถึงไม่เจอ ลองระบุชื่อหรือวันที่เพิ่มอีกนิดนะ');
  if (matches.length > 1) return { ok: false, needsChoice: true, action: command.action, changes: command.changes || {}, records: matches.slice(0, 8), message: 'เจอหลายรายการ เลือกรายการที่ต้องการนะ' };
  return { ok: false, needsConfirm: true, action: command.action, changes: command.changes || {}, record: matches[0], message: command.action === 'delete' ? 'ต้องการลบรายการนี้ใช่ไหม?' : 'แก้รายการนี้ใช่ไหม?' };
}

function getRecords_() {
  const values = ensureSheet_().getDataRange().getValues();
  return values.slice(1).filter(r => r[0] && r[7] === 'ใช้งาน').map(rowToRecord_).sort((a, b) => new Date(b.date) - new Date(a.date) || new Date(b.updatedAt) - new Date(a.updatedAt));
}

function ensureSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('กรุณาผูก Apps Script นี้กับ Google Sheet ก่อนใช้งาน');
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]).setFontWeight('bold').setBackground('#DCEFE6');
    sheet.setFrozenRows(1);
    sheet.getRange('B:B').setNumberFormat('dd/MM/yyyy');
    sheet.getRange('F:F').setNumberFormat('#,##0.00');
    sheet.autoResizeColumns(1, HEADERS.length);
  } else if (sheet.getLastColumn() < HEADERS.length) {
    sheet.getRange(1, 11, 1, HEADERS.length - 10).setValues([HEADERS.slice(10)]).setFontWeight('bold').setBackground('#DCEFE6');
  }
  return sheet;
}

function ensureCategorySheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(CATEGORY_SHEET);
  if (!sheet) sheet = ss.insertSheet(CATEGORY_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 7).setValues([['รหัส', 'ประเภท', 'ชื่อหมวดหมู่', 'สถานะ', 'แก้ไขเมื่อ', 'บัญชี', 'รหัสธุรกิจ']]).setFontWeight('bold').setBackground('#DCEFE6');
    sheet.getRange(2, 1, DEFAULT_CATEGORIES.length, 7).setValues(DEFAULT_CATEGORIES.map(c => [Utilities.getUuid(), c[0], c[1], 'ใช้งาน', new Date(), 'HOME', '']));
    sheet.setFrozenRows(1);
  } else if (sheet.getLastColumn() < 7) {
    sheet.getRange(1, 6, 1, 2).setValues([['บัญชี', 'รหัสธุรกิจ']]).setFontWeight('bold').setBackground('#DCEFE6');
  }
  const defaultBusiness = getBusinesses_()[0];
  const values = sheet.getDataRange().getValues();
  let migrated = false;
  values.slice(1).forEach(r => {
    if (normalizeWorkspace_(r[5]) === 'BUSINESS' && !r[6]) { r[6] = defaultBusiness.id; migrated = true; }
  });
  if (migrated) sheet.getRange(2, 1, values.length - 1, 7).setValues(values.slice(1).map(r => r.slice(0, 7)));
  const hasBusiness = values.slice(1).some(r => normalizeWorkspace_(r[5]) === 'BUSINESS' && String(r[6] || '') === defaultBusiness.id);
  if (!hasBusiness) sheet.getRange(sheet.getLastRow() + 1, 1, BUSINESS_CATEGORIES.length, 7).setValues(BUSINESS_CATEGORIES.map(c => [Utilities.getUuid(), c[0], c[1], 'ใช้งาน', new Date(), 'BUSINESS', defaultBusiness.id]));
  return sheet;
}

function getCategories_() {
  const defaultId = getBusinesses_()[0].id;
  return ensureCategorySheet_().getDataRange().getValues().slice(1).filter(r => r[0] && r[3] === 'ใช้งาน').map(r => ({ id: String(r[0]), type: r[1], name: r[2], workspace: normalizeWorkspace_(r[5]), businessId: normalizeWorkspace_(r[5]) === 'BUSINESS' ? String(r[6] || defaultId) : '' }));
}

function migrateBusinessData_() {
  const defaultId = getBusinesses_()[0].id;
  const sheet = ensureSheet_();
  if (sheet.getLastRow() < 2) return;
  const values = sheet.getRange(2, 11, sheet.getLastRow() - 1, 4).getValues();
  let changed = false;
  values.forEach(r => {
    if (normalizeWorkspace_(r[0]) === 'BUSINESS' && !r[3]) { r[3] = defaultId; changed = true; }
  });
  if (changed) sheet.getRange(2, 11, values.length, 4).setValues(values);
}

function ensureBusinessSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BUSINESS_SHEET);
  if (!sheet) sheet = ss.insertSheet(BUSINESS_SHEET);
  if (sheet.getLastRow() === 0) {
    const id = Utilities.getUuid();
    sheet.getRange(1, 1, 1, 5).setValues([['รหัส', 'ชื่อธุรกิจ', 'สถานะ', 'สร้างเมื่อ', 'แก้ไขเมื่อ']]).setFontWeight('bold').setBackground('#DCEFE6');
    sheet.appendRow([id, 'ธุรกิจหลัก', 'ใช้งาน', new Date(), new Date()]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function getBusinesses_() {
  return ensureBusinessSheet_().getDataRange().getValues().slice(1).filter(r => r[0] && r[2] === 'ใช้งาน').map(r => ({ id: String(r[0]), name: String(r[1]) }));
}

function ensureSavingsSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SAVINGS_SHEET);
  if (!sheet) sheet = ss.insertSheet(SAVINGS_SHEET);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, 8).setValues([['รหัส', 'วันที่', 'สมาชิก', 'รายการ', 'จำนวนเงิน', 'หมายเหตุ', 'สถานะ', 'สร้างเมื่อ']]).setFontWeight('bold').setBackground('#DCEFE6');
    sheet.setFrozenRows(1);
    sheet.getRange('B:B').setNumberFormat('dd/MM/yyyy');
    sheet.getRange('E:E').setNumberFormat('#,##0.00');
  }
  return sheet;
}

function getSavings_() {
  return ensureSavingsSheet_().getDataRange().getValues().slice(1).filter(r => r[0] && r[6] === 'ใช้งาน').map(r => ({
    id: String(r[0]), date: Utilities.formatDate(new Date(r[1]), 'Asia/Bangkok', 'yyyy-MM-dd'), member: normalizeMember_(r[2]), action: r[3] === 'WITHDRAW' ? 'WITHDRAW' : 'DEPOSIT', amount: Number(r[4]), note: String(r[5] || '')
  })).sort((a, b) => b.date.localeCompare(a.date));
}

function getReceiptFolder_() {
  try {
    const folders = DriveApp.getFoldersByName('AiHouse Receipts');
    return folders.hasNext() ? folders.next() : DriveApp.createFolder('AiHouse Receipts');
  } catch (e) {
    return DriveApp.getRootFolder();
  }
}

function parseSlipText_(text) {
  const rawText = String(text || '');
  const clean = rawText.replace(/,/g, '').replace(/\r/g, '');
  const inlineText = clean.replace(/\n+/g, ' ').replace(/\s+/g, ' ').trim();

  let date = '';
  const thaiMonths = {
    'ม.ค.':1, 'มกราคม':1, 'ก.พ.':2, 'กุมภาพันธ์':2, 'มี.ค.':3, 'มีนาคม':3,
    'เม.ย.':4, 'เมษายน':4, 'พ.ค.':5, 'พฤษภาคม':5, 'มิ.ย.':6, 'มิถุนายน':6,
    'ก.ค.':7, 'กรกฎาคม':7, 'ส.ค.':8, 'สิงหาคม':8, 'ก.ย.':9, 'กันยายน':9,
    'ต.ค.':10, 'ตุลาคม':10, 'พ.ย.':11, 'พฤศจิกายน':11, 'ธ.ค.':12, 'ธันวาคม':12,
    'jan':1, 'feb':2, 'mar':3, 'apr':4, 'may':5, 'jun':6, 'jul':7, 'aug':8, 'sep':9, 'oct':10, 'nov':11, 'dec':12
  };

  const thaiDateMatch = inlineText.match(/(\d{1,2})\s*(ม\.?ค\.?|มกราคม|ก\.?พ\.?|กุมภาพันธ์|มี\.?ค\.?|มีนาคม|เม\.?ย\.?|เมษายน|พ\.?ค\.?|พฤษภาคม|มิ\.?ย\.?|มิถุนายน|ก\.?ค\.?|กรกฎาคม|ส\.?ค\.?|สิงหาคม|ก\.?ย\.?|กันยายน|ต\.?ค\.?|ตุลาคม|พ\.?ย\.?|พฤศจิกายน|ธ\.?ค\.?|ธันวาคม|jan|feb|mar|apr|may|jun|jul|aug|sep|oct|nov|dec)\s*(\d{2,4})/i);
  if (thaiDateMatch) {
    const dNum = Number(thaiDateMatch[1]);
    const mStr = thaiDateMatch[2].toLowerCase().replace(/\./g, '');
    let yNum = Number(thaiDateMatch[3]);
    let mNum = 0;
    for (const key in thaiMonths) {
      const cleanKey = key.replace(/\./g, '');
      if (mStr === cleanKey || mStr.startsWith(cleanKey) || cleanKey.startsWith(mStr)) {
        mNum = thaiMonths[key];
        break;
      }
    }
    if (mNum > 0 && dNum >= 1 && dNum <= 31) {
      if (yNum > 2400) yNum -= 543;
      else if (yNum < 100) yNum += 2000;
      const parsedD = new Date(yNum, mNum - 1, dNum, 12);
      if (!isNaN(parsedD.getTime()) && parsedD.getDate() === dNum && parsedD.getMonth() === mNum - 1) {
        date = Utilities.formatDate(parsedD, 'Asia/Bangkok', 'yyyy-MM-dd');
      }
    }
  }

  if (!date) {
    const dm = inlineText.match(/(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (dm) {
      let y = Number(dm[3]);
      if (y > 2400) y -= 543;
      if (y < 100) y += 2000;
      const d = new Date(y, Number(dm[2]) - 1, Number(dm[1]), 12);
      if (!isNaN(d.getTime()) && d.getDate() === Number(dm[1]) && d.getMonth() === Number(dm[2]) - 1) date = Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
    }
  }

  let amount = 0;
  const amountPatterns = [
    /(?:จำนวนเงิน|ยอดเงิน|ยอดโอน|ยอดชำระ|จำนวน|รวมทั้งสิ้น|รวม|amount|total)\s*[:฿]?\s*([0-9]+(?:\.[0-9]{1,2})?)/i,
    /(?:฿|THB|บาท)\s*([0-9]+(?:\.[0-9]{1,2})?)/i,
    /([0-9]+(?:\.[0-9]{2}))\s*(?:บาท|THB|baht)/i,
    /(?:โอน|จ่าย|ชำระ)\s*([0-9]+(?:\.[0-9]{1,2})?)/i
  ];

  for (const pattern of amountPatterns) {
    const m = inlineText.match(pattern);
    if (m && Number(m[1]) > 0) {
      const val = Number(m[1]);
      if (val < 10000000) {
        amount = val;
        break;
      }
    }
  }

  if (!amount) {
    const decimalMatches = inlineText.match(/\b\d+\.\d{2}\b/g) || [];
    const candidates = decimalMatches.map(Number).filter(n => n > 0 && n < 1000000);
    if (candidates.length > 0) amount = Math.max(...candidates);
  }

  const incoming = /เงินเข้า|รับเงิน|ได้รับเงิน|received|เครดิตเข้า|โอนเงินเข้า/i.test(inlineText);
  const outgoing = /ชำระเงิน|โอนเงินสำเร็จ|ผู้โอน|จ่ายเงิน|payment successful|ยอดชำระ|โอนให้|สำเร็จ/i.test(inlineText);
  const type = incoming && !outgoing ? 'รายรับ' : 'รายจ่าย';

  let note = '';
  let detectedCat = '';

  const brandRules = [
    [/เซเว่น|7-eleven|7eleven|cp all/i, '7-Eleven', 'อาหาร'],
    [/amazon|อเมซอน/i, 'Café Amazon', 'อาหาร'],
    [/starbucks|สตาร์บัคส์/i, 'Starbucks', 'อาหาร'],
    [/kfc/i, 'KFC', 'อาหาร'],
    [/mcdonald|แมคโดนัลด์/i, 'McDonald\'s', 'อาหาร'],
    [/lotus|โลตัส/i, 'Lotus\'s', 'ช้อปปิ้ง'],
    [/big c|บิ๊กซี/i, 'Big C', 'ช้อปปิ้ง'],
    [/cj express|ซีเจ|cj more/i, 'CJ More', 'ช้อปปิ้ง'],
    [/makro|แม็คโคร/i, 'Makro', 'ช้อปปิ้ง'],
    [/grab/i, 'Grab', 'เดินทาง'],
    [/lineman|ไลน์แมน/i, 'LINE Man', 'อาหาร'],
    [/shopee/i, 'Shopee', 'ช้อปปิ้ง'],
    [/lazada/i, 'Lazada', 'ช้อปปิ้ง'],
    [/ptt|ปตท|bangchak|บางจาก|shell|เชลล์|caltex|เติมน้ำมัน/i, 'เติมน้ำมัน', 'เดินทาง'],
    [/ค่าไฟ|การไฟฟ้า|mea|pea/i, 'ค่าไฟฟ้า', 'สาธารณูปโภค'],
    [/ค่าน้ำ|การประปา|mwa|pwa/i, 'ค่าน้ำประปา', 'สาธารณูปโภค'],
    [/ais|true|dtac|nt|ค่าเน็ต|ค่าโทร/i, 'ค่าโทรศัพท์/อินเทอร์เน็ต', 'สาธารณูปโภค']
  ];

  for (const [pattern, name, cat] of brandRules) {
    if (pattern.test(inlineText)) {
      note = name;
      detectedCat = cat;
      break;
    }
  }

  if (!note) {
    const merchantMatch = inlineText.match(/(?:ไปยัง|ผู้รับเงิน|ผู้รับ|ถึง|รับโดย|บัญชีผู้รับ|ร้านค้า|to|receiver)\s*[:：]?\s*([^\d\n\r]{2,35})/i);
    if (merchantMatch) {
      note = merchantMatch[1].trim().replace(/(?:จำนวนเงิน|ยอดเงิน|ยอดโอน|ยอดชำระ|เลขที่|บัญชี|วันที่|เวลา|ธนาคาร|จำกัด|สาขา|ref|promptpay).*$/i, '').trim();
    }
  }

  const category = detectedCat || categorize_(inlineText + ' ' + note, type);
  const refMatch = inlineText.match(/(?:เลขที่รายการ|เลขที่อ้างอิง|รหัสอ้างอิง|หมายเลขอ้างอิง|transaction\s*(?:id|no)|reference|ref(?:erence)?\.?)[\s:#：-]*([A-Z0-9-]{6,40})/i);
  const slipRef = refMatch ? normalizeSlipRef_(refMatch[1]) : '';

  return {
    type,
    amount,
    date,
    category: category === 'อื่น ๆ' ? '' : category,
    note: note || '',
    slipRef
  };
}

function rowToRecord_(r) {
  return { id: String(r[0]), date: Utilities.formatDate(new Date(r[1]), 'Asia/Bangkok', 'yyyy-MM-dd'), type: r[2], category: r[3], note: r[4], amount: Number(r[5]), user: normalizeMember_(r[6]), status: r[7], createdAt: r[8] instanceof Date ? r[8].toISOString() : String(r[8]), updatedAt: r[9] instanceof Date ? r[9].toISOString() : String(r[9]), workspace: normalizeWorkspace_(r[10]), receiptUrl: String(r[11] || ''), slipRef: normalizeSlipRef_(r[12]), businessId: String(r[13] || '') };
}

function normalizeSlipRef_(value) {
  return String(value || '').toUpperCase().replace(/[^A-Z0-9]/g, '');
}

function findDuplicateSlip_(data, workspace, businessId) {
  const ref = normalizeSlipRef_(data && data.slipRef);
  const targetWorkspace = normalizeWorkspace_(workspace || (data && data.workspace));
  return getRecords_().find(r => {
    if (r.workspace !== targetWorkspace) return false;
    if (targetWorkspace === 'BUSINESS' && r.businessId !== normalizeBusinessId_(businessId || data.businessId)) return false;
    if (ref && r.slipRef) return ref === r.slipRef;
    return !ref && data.date && r.date === data.date && Number(r.amount) === Number(data.amount) && String(r.note || '').trim().toLowerCase() === String(data.note || '').trim().toLowerCase();
  }) || null;
}

function normalizeWorkspace_(value) {
  return String(value || '').toUpperCase() === 'BUSINESS' ? 'BUSINESS' : 'HOME';
}

function normalizeBusinessId_(value) {
  const businesses = getBusinesses_();
  const id = String(value || '');
  return businesses.some(b => b.id === id) ? id : businesses[0].id;
}

function normalizeMember_(value) {
  return ['บาส', 'แตงโม'].includes(String(value)) ? String(value) : 'ส่วนกลาง';
}

function validateRecord_(d) {
  if (!['รายรับ', 'รายจ่าย'].includes(d.type)) throw new Error('กรุณาระบุว่าเป็นรายรับหรือรายจ่าย');
  if (!Number(d.amount) || Number(d.amount) <= 0) throw new Error('กรุณาระบุจำนวนเงินมากกว่า 0 บาท');
}

function parseDate_(value) {
  if (!value) return new Date();
  const d = new Date(value + (String(value).length === 10 ? 'T12:00:00+07:00' : ''));
  return isNaN(d) ? new Date() : d;
}

function parseCommand_(text) {
  const normalized = text.replace(/,/g, '').replace(/บาท/g, '').trim();
  const isDelete = /^(ลบ|ยกเลิก)/.test(normalized);
  const isEdit = /^(แก้|แก้ไข|เปลี่ยน)/.test(normalized) || /แก้เป็น|เปลี่ยนเป็น/.test(normalized);
  const isQuery = /เท่าไหร่|เท่าไร|สรุป|ยอด|เหลือเงิน|ใช้ไป|ใช้จ่ายไป|รายรับทั้งหมด|รายจ่ายทั้งหมด/.test(normalized);
  const period = /ปีนี้|รายปี/.test(normalized) ? 'year' : /เดือนนี้|รายเดือน/.test(normalized) ? 'month' : /วันนี้|รายวัน/.test(normalized) ? 'day' : 'all';
  if (isQuery && !isDelete && !isEdit) return { text, action: 'query', period, type: /รายรับ|ได้เงิน|รับมา/.test(normalized) ? 'รายรับ' : /รายจ่าย|ใช้|จ่าย/.test(normalized) ? 'รายจ่าย' : 'สุทธิ', keyword: extractCategory_(normalized) };

  const numbers = normalized.match(/\d+(?:\.\d+)?/g) || [];
  const amount = numbers.length ? Number(numbers[numbers.length - 1]) : 0;
  const type = /รายรับ|ได้เงิน|เงินเดือน|ขาย.*ได้|รับเงิน/.test(normalized) ? 'รายรับ' : /รายจ่าย|ซื้อ|จ่าย|ค่า|เติม/.test(normalized) ? 'รายจ่าย' : '';
  const date = relativeDate_(normalized);
  const category = categorize_(normalized, type);
  const cleaned = normalized.replace(/^(เพิ่ม|บันทึก|ลบ|ยกเลิก|แก้ไข|แก้|เปลี่ยน)\s*/,'').replace(/รายรับ|รายจ่าย/g,'').replace(/\d+(?:\.\d+)?/g,'').replace(/วันนี้|เมื่อวาน|บาท|ล่าสุด|เป็น|จาก/g,' ').replace(/\s+/g,' ').trim();
  if (isDelete) return { text, action: 'delete', keyword: cleaned, date, amount: amount || null };
  if (isEdit) return { text, action: 'edit', keyword: cleaned, date, changes: { amount: amount || undefined, type: type || undefined, category: category || undefined } };
  return { text, action: 'add', data: { type, category, note: cleaned || category, amount, date } };
}

function relativeDate_(text) {
  const d = new Date();
  if (/เมื่อวาน/.test(text)) d.setDate(d.getDate() - 1);
  const match = text.match(/(?:วันที่\s*)?(\d{1,2})[\/\-](\d{1,2})(?:[\/\-](\d{2,4}))?/);
  if (match) {
    let year = match[3] ? Number(match[3]) : d.getFullYear();
    if (year > 2400) year -= 543;
    if (year < 100) year += 2000;
    d.setFullYear(year, Number(match[2]) - 1, Number(match[1]));
  }
  return Utilities.formatDate(d, 'Asia/Bangkok', 'yyyy-MM-dd');
}

function categorize_(text, type) {
  const rules = [
    ['เงินเดือน', /เงินเดือน|ค่าจ้าง/], ['รายได้เสริม', /ขาย|รายได้เสริม|โบนัส/], ['อาหาร', /ข้าว|อาหาร|กาแฟ|ชา|ขนม|ร้านอาหาร/],
    ['เดินทาง', /น้ำมัน|รถ|แท็กซี่|ทางด่วน|รถไฟ|เดินทาง/], ['บ้าน', /ค่าเช่า|ผ่อนบ้าน|ของใช้|ซ่อมบ้าน/],
    ['สาธารณูปโภค', /ค่าไฟ|ค่าน้ำ|อินเทอร์เน็ต|โทรศัพท์/], ['สุขภาพ', /ยา|หมอ|โรงพยาบาล/], ['การศึกษา', /ค่าเทอม|หนังสือ|โรงเรียน|เรียน/], ['ช้อปปิ้ง', /ซื้อเสื้อ|ช้อป|เสื้อผ้า/]
  ];
  const hit = rules.find(r => r[1].test(text));
  return hit ? hit[0] : type === 'รายรับ' ? 'รายรับอื่น ๆ' : 'อื่น ๆ';
}

function categorizeBusiness_(text, type) {
  const rules = [
    ['ขายสินค้า', /ขายสินค้า|ยอดขาย|ออเดอร์|ขายของ/], ['ค่าบริการ', /ค่าบริการ|บริการ|ค่าจ้าง/],
    ['ต้นทุนสินค้า', /ต้นทุนสินค้า|ซื้อสินค้า|สต[็อ]ก/], ['วัตถุดิบ', /วัตถุดิบ/],
    ['ค่าขนส่ง', /ค่าขนส่ง|ค่าส่ง|ส่งของ/], ['ค่าโฆษณา', /โฆษณา|แอด|facebook|tiktok|google ads/i],
    ['ค่าเช่า', /ค่าเช่า/], ['ค่าน้ำค่าไฟ', /ค่าน้ำ|ค่าไฟ/], ['เงินเดือน', /เงินเดือน|ค่าแรง/]
  ];
  const hit = rules.find(r => r[1].test(String(text || '')));
  return hit ? hit[0] : type === 'รายรับ' ? 'รายได้ธุรกิจอื่น ๆ' : type === 'รายจ่าย' ? 'ค่าใช้จ่ายธุรกิจอื่น ๆ' : '';
}

function extractCategory_(text) {
  const cat = categorize_(text, '');
  return cat === 'อื่น ๆ' ? '' : cat;
}

function findMatches_(command) {
  let records = getRecords_().filter(r => r.workspace === command.workspace);
  if (command.workspace === 'HOME') records = records.filter(r => r.user === command.member);
  else records = records.filter(r => r.businessId === command.businessId);
  if (command.amount) records = records.filter(r => r.amount === command.amount);
  if (command.keyword) {
    const words = command.keyword.split(/\s+/).filter(w => w.length > 1 && !/รายการ|ของ|ล่าสุด/.test(w));
    if (words.length) records = records.filter(r => words.some(w => (r.note + ' ' + r.category).includes(w)));
  }
  if (/ล่าสุด/.test(command.text)) return records.slice(0, 1);
  return records;
}

function answerQuery_(q) {
  const now = new Date();
  let records = getRecords_().filter(r => r.workspace === q.workspace).filter(r => q.workspace === 'BUSINESS' ? r.businessId === q.businessId : r.user === q.member).filter(r => {
    const d = new Date(r.date + 'T12:00:00');
    if (q.period === 'day') return d.toDateString() === now.toDateString();
    if (q.period === 'month') return d.getMonth() === now.getMonth() && d.getFullYear() === now.getFullYear();
    if (q.period === 'year') return d.getFullYear() === now.getFullYear();
    return true;
  });
  if (q.keyword) records = records.filter(r => r.category === q.keyword);
  const income = records.filter(r => r.type === 'รายรับ').reduce((s,r) => s + r.amount, 0);
  const expense = records.filter(r => r.type === 'รายจ่าย').reduce((s,r) => s + r.amount, 0);
  const periodName = { day: 'วันนี้', month: 'เดือนนี้', year: 'ปีนี้', all: 'ทั้งหมด' }[q.period];
  const amount = q.type === 'รายรับ' ? income : q.type === 'รายจ่าย' ? expense : income - expense;
  const label = q.type === 'รายรับ' ? 'มีรายรับ' : q.type === 'รายจ่าย' ? 'ใช้จ่ายไป' : 'มียอดสุทธิ';
  return { ok: true, query: true, amount, income, expense, message: `${periodName}${q.keyword ? ' หมวด' + q.keyword : ''} ${label} ${amount.toLocaleString('th-TH')} บาท จาก ${records.length} รายการ` };
}
