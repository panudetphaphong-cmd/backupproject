const SHEET_NAME = 'รายการบัญชี';
const HEADERS = ['รหัส', 'วันที่', 'ประเภท', 'หมวดหมู่', 'รายละเอียด', 'จำนวนเงิน', 'ผู้บันทึก', 'สถานะ', 'สร้างเมื่อ', 'แก้ไขเมื่อ'];

function doGet() {
  ensureSheet_();
  return HtmlService.createTemplateFromFile('Index').evaluate()
    .setTitle('AiHouse — บัญชีบ้านของเรา')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function getInitialData() {
  return { records: getRecords_(), user: Session.getActiveUser().getEmail() || 'ครอบครัวเรา' };
}

function addRecord(data) {
  validateRecord_(data);
  const sheet = ensureSheet_();
  const now = new Date();
  const row = [Utilities.getUuid(), parseDate_(data.date), data.type, data.category || 'อื่น ๆ', data.note || '', Number(data.amount), data.user || 'ครอบครัวเรา', 'ใช้งาน', now, now];
  sheet.appendRow(row);
  return { ok: true, message: 'เก็บให้เรียบร้อยแล้วนะ', record: rowToRecord_(row) };
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
      return { ok: true, message: 'ย้ายไปถังขยะแล้ว กู้คืนได้เสมอนะ' };
    }
  }
  throw new Error('ไม่พบรายการที่ต้องการลบ');
}

function processCommand(text) {
  const command = parseCommand_(String(text || '').trim());
  if (!command.text) throw new Error('ลองบอกน้องเฮาส์อีกครั้งนะ');
  if (command.action === 'query') return answerQuery_(command);
  if (command.action === 'add') return addRecord(command.data);

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
  }
  return sheet;
}

function rowToRecord_(r) {
  return { id: String(r[0]), date: Utilities.formatDate(new Date(r[1]), 'Asia/Bangkok', 'yyyy-MM-dd'), type: r[2], category: r[3], note: r[4], amount: Number(r[5]), user: r[6], status: r[7], createdAt: r[8] instanceof Date ? r[8].toISOString() : String(r[8]), updatedAt: r[9] instanceof Date ? r[9].toISOString() : String(r[9]) };
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

function extractCategory_(text) {
  const cat = categorize_(text, '');
  return cat === 'อื่น ๆ' ? '' : cat;
}

function findMatches_(command) {
  let records = getRecords_();
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
  let records = getRecords_().filter(r => {
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
