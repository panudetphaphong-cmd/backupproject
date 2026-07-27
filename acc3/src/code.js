/**
 * ระบบบัญชีลูกค้า & สรุปยอด - Google Apps Script Backend
 */

// แสดงผลหน้าเว็บหลัก
const ACCOUNTING_SPREADSHEET_ID = '1mPivRBQ2jIjhb-P--eddH4SyTnaTNZ3zAuxHdQe9XKc';
const ACCOUNTING_DATA_SHEET = '_SYSTEM_DATA';
const ACCOUNTING_DATA_CHUNK_SIZE = 45000;

function doGet() {
  try {
    initializeSheetDatabase_();
  } catch (error) {
    Logger.log('doGet initialize error: ' + error);
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('ระบบบันทึกบัญชีลูกค้า')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function initializeSheetDatabase_() {
  const properties = PropertiesService.getScriptProperties();
  const migrationKey = 'SHEET_DATABASE_INITIALIZED_' + ACCOUNTING_SPREADSHEET_ID;
  if (properties.getProperty(migrationKey) === 'true') return;

  try {
    let dataJson = readDataFromSheet_();
    if (!dataJson) {
      const data = getData();
      dataJson = readDataFromSheet_();
      if (!dataJson) writeDataToSheet_(data);
    }
    properties.setProperty(migrationKey, 'true');
  } catch (error) {
    Logger.log('Sheet database initialization failed: ' + error);
  }
}

function initializeSheetDatabase() {
  const properties = PropertiesService.getScriptProperties();
  const migrationKey = 'SHEET_DATABASE_INITIALIZED_' + ACCOUNTING_SPREADSHEET_ID;
  properties.deleteProperty(migrationKey);
  initializeSheetDatabase_();
  const sheet = getDataSpreadsheet_().getSheetByName(ACCOUNTING_DATA_SHEET);
  if (!sheet || sheet.getLastRow() < 3) {
    throw new Error('Database Sheet initialization did not complete');
  }
  return {
    success: true,
    spreadsheetId: ACCOUNTING_SPREADSHEET_ID,
    sheetName: ACCOUNTING_DATA_SHEET,
    chunks: Number(sheet.getRange('A2').getValue()) || 0,
    updatedAt: sheet.getRange('B2').getDisplayValue()
  };
}

// ดึงข้อมูลทั้งหมดจากระบบบันทึก
function getData() {
  const scriptProperties = PropertiesService.getScriptProperties();
  let dataJson = readDataFromSheet_();

  if (dataJson) {
    try {
      return repairCustomerRecords_(JSON.parse(dataJson));
    } catch (e) {
      Logger.log('Error parsing Sheet data: ' + e);
    }
  }

  // เก็บ JSON ในไฟล์ Drive เพราะ PropertiesService จำกัดขนาดเล็กเกินไปสำหรับรูปแนบ
  const dataFileId = scriptProperties.getProperty('ACCOUNTING_DATA_FILE_ID');
  if (dataFileId) {
    try {
      dataJson = DriveApp.getFileById(dataFileId).getBlob().getDataAsString('UTF-8');
    } catch (e) {
      Logger.log('Error reading shared data file: ' + e);
    }
  }

  // รองรับและย้ายข้อมูลจากระบบ Properties เดิม
  if (!dataJson) {
    dataJson = scriptProperties.getProperty('ACCOUNTING_SYSTEM_DATA') ||
      PropertiesService.getUserProperties().getProperty('ACCOUNTING_SYSTEM_DATA') || '';
  }
  
  if (dataJson) {
    try {
      const legacyData = JSON.parse(dataJson);
      try {
        writeDataToSheet_(legacyData);
      } catch (migrationError) {
        Logger.log('Sheet migration failed; continuing with legacy data: ' + migrationError);
      }
      return repairCustomerRecords_(legacyData);
    } catch (e) {
      Logger.log("Error parsing saved data: " + e);
    }
  }
  
  // ค่าเริ่มต้นกรณีใช้งานครั้งแรก
  return {
    columns: [
      { id: 'eday', label: 'EDAY' },
      { id: 'sm', label: 'SM' },
      { id: 'sp999', label: 'SP999' },
      { id: 'withdraw', label: 'ยอดเบิก' }
    ],
    customers: [
      { id: 1, name: 'คุณสมชาย', values: { eday: 1500, sm: 500, sp999: -200, withdraw: -300 }, status: 'ชำระแล้ว', invoiceSent: false, invoiceHistory: [{ id: 101, total: 1500, date: '25/07/2026', paidAt: '26/07/2026 15:30:00' }], banks: [{ id: 101, bankName: 'กสิกรไทย', accNo: '555-2-12345-6', accName: 'คุณสมชาย ใจดี' }] },
      { id: 2, name: 'คุณวิภา', values: { eday: 3000, sm: -1000, sp999: 0, withdraw: -500 }, status: 'รอชำระ', invoiceSent: true, invoiceDate: '27/07/2026', invoiceHistory: [], banks: [{ id: 102, bankName: 'ไทยพาณิชย์', accNo: '999-8-76543-2', accName: 'คุณวิภา สุขสันต์' }] }
    ],
    banks: [
      { id: 1, bankName: 'กสิกรไทย (KBANK)', accNo: '123-4-56789-0', accName: 'นายบัญชี รับเงิน', promptpay: '0812345678' },
      { id: 2, bankName: 'ไทยพาณิชย์ (SCB)', accNo: '987-6-54321-0', accName: 'นายบัญชี รับเงิน', promptpay: '0812345678' }
    ],
    notifications: [],
    announcements: [
      {
        id: 1,
        title: '📢 ประกาศแจ้งปิดปรับปรุงระบบชั่วคราว',
        content: 'ระบบจะทำการปิดปรับปรุงประจำสัปดาห์ในวันอาทิตย์นี้ เวลา 02:00 - 04:00 น.',
        type: 'urgent',
        createdAt: '27/07/2026 10:00:00',
        expireAt: new Date(Date.now() + 7 * 24 * 60 * 60 * 1000).toISOString()
      }
    ],
    users: [
      { username: 'admin', password: 'admin123', role: 'admin', name: 'แอดมิน' },
      { username: 'member', password: 'member123', role: 'member', name: 'เมมเบอร์' },
      { username: 'somchai', password: '1234', role: 'customer', name: 'คุณสมชาย', customerId: 1 },
      { username: 'wipa', password: '1234', role: 'customer', name: 'คุณวิภา', customerId: 2 }
    ],
    transactions: [
      {
        id: 1,
        colLabel: 'EDAY',
        type: 'income',
        amount: 50000,
        startDate: '2026-07-01',
        endDate: '2026-07-27',
        note: 'รายได้การให้บริการ EDAY ครึ่งเดือนแรก',
        createdAt: new Date().toISOString()
      },
      {
        id: 2,
        colLabel: 'EDAY',
        type: 'expense',
        amount: 12000,
        startDate: '2026-07-01',
        endDate: '2026-07-27',
        note: 'ค่าใช้จ่ายและค่าธรรมเนียม EDAY',
        createdAt: new Date().toISOString()
      },
      {
        id: 3,
        colLabel: 'SM',
        type: 'income',
        amount: 35000,
        startDate: '2026-07-01',
        endDate: '2026-07-27',
        note: 'รายได้หมวด SM',
        createdAt: new Date().toISOString()
      }
    ],
    chatMessages: [
      {
        id: 1,
        customerName: 'คุณสมชาย',
        senderName: 'คุณสมชาย',
        senderRole: 'customer',
        text: 'สวัสดีครับแอดมิน สอบถามเรื่องการส่งสลิปโอนเงินครับ',
        image: '',
        isRead: true,
        createdAt: '18:30'
      },
      {
        id: 2,
        customerName: 'คุณสมชาย',
        senderName: 'แอดมิน',
        senderRole: 'admin',
        text: 'สวัสดีครับ สามารถกดปุ่ม "💳 ชำระยอดเงิน & แนบสลิปโอน" ในใบแจ้งหนี้เพื่อส่งสลิปได้เลยครับ',
        image: '',
        isRead: true,
        createdAt: '18:31'
      }
    ]
  };
}

// Older customer clients could save a filtered snapshot containing only their
// own row. Recreate any missing customer rows from the authoritative user list
// so an existing login can never silently disappear from the admin table.
function repairCustomerRecords_(data) {
  if (!data || typeof data !== 'object') return data;
  data.customers = Array.isArray(data.customers) ? data.customers : [];
  const existingIds = {};
  data.customers.forEach(function (customer) {
    existingIds[String(customer.id)] = true;
    customer.values = customer.values || {};
    customer.invoiceHistory = customer.invoiceHistory || [];
    customer.banks = customer.banks || [];
  });
  (data.users || []).forEach(function (user) {
    if (String(user.role || '').toLowerCase() !== 'customer' || !user.customerId) return;
    const id = String(user.customerId);
    if (existingIds[id]) return;
    data.customers.push({
      id: user.customerId,
      name: user.name || user.username || ('Customer ' + id),
      values: {},
      status: 'รอชำระ',
      invoiceSent: false,
      invoiceHistory: [],
      banks: [],
      recoveredRecord: true
    });
    existingIds[id] = true;
  });
  return data;
}

// ตรวจสอบการเข้าสู่ระบบ
function authenticateUser(username, password) {
  const data = getData();
  const users = (data && data.users && data.users.length > 0) ? data.users : [
    { username: 'admin', password: 'admin123', role: 'admin', name: 'แอดมิน' },
    { username: 'member', password: 'member123', role: 'member', name: 'เมมเบอร์' },
    { username: 'somchai', password: '1234', role: 'customer', name: 'คุณสมชาย', customerId: 1 },
    { username: 'wipa', password: '1234', role: 'customer', name: 'คุณวิภา', customerId: 2 }
  ];

  const found = users.find(u => u.username.toLowerCase() === String(username).toLowerCase().trim() && u.password === password);
  if (found) {
    return {
      success: true,
      user: {
        username: found.username,
        role: found.role,
        name: found.name,
        customerId: found.customerId || null
      }
    };
  } else {
    return { success: false, message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง' };
  }
}

// บันทึกข้อมูลกลับไปยังระบบ
function saveData(data) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    // การบันทึกส่วนบัญชีอาจมาจากหน้าจอที่ยังมีแชทชุดเก่า:
    // รวมข้อความจากเซิร์ฟเวอร์ก่อนเสมอ เพื่อไม่ให้ข้อความล่าสุดถูกเขียนทับหาย
    const currentData = getData();
    data = data || {};
    const actor = data.actor || {};
    delete data.actor;

    // A customer receives only their own row in realtime snapshots. Never allow
    // that intentionally filtered payload to replace the complete shared table.
    if (String(actor.role || '').toLowerCase() === 'customer') {
      const customerId = String(actor.customerId || '');
      const customerName = String(actor.customerName || '').trim().toLowerCase();
      const incomingCustomer = (data.customers || []).find(function (customer) {
        return String(customer.id) === customerId;
      });
      data.customers = (currentData.customers || []).map(function (customer) {
        return incomingCustomer && String(customer.id) === customerId
          ? incomingCustomer
          : customer;
      });

      const notificationMap = {};
      (currentData.notifications || []).forEach(function (item) {
        notificationMap[String(item.id)] = item;
      });
      (data.notifications || []).forEach(function (item) {
        if (String(item.customerName || '').trim().toLowerCase() === customerName) {
          notificationMap[String(item.id)] = item;
        }
      });
      data.notifications = Object.keys(notificationMap).map(function (key) {
        return notificationMap[key];
      });

      // These collections are admin-owned and may be absent from a filtered
      // customer snapshot.
      data.columns = currentData.columns || [];
      data.banks = currentData.banks || [];
      data.announcements = currentData.announcements || [];
      data.users = currentData.users || [];
      data.transactions = currentData.transactions || [];
    }
    const mergedChats = {};
    (currentData.chatMessages || []).forEach(function (m) {
      mergedChats[String(m.id)] = m;
    });
    (data.chatMessages || []).forEach(function (m) {
      const key = String(m.id);
      if (mergedChats[key]) {
        m.isRead = Boolean(m.isRead || mergedChats[key].isRead);
      }
      mergedChats[key] = m;
    });
    data.chatMessages = Object.keys(mergedChats)
      .map(function (key) { return mergedChats[key]; })
      .sort(function (a, b) { return Number(a.id) - Number(b.id); });
    writeSharedData_(data);
    return { success: true, message: 'บันทึกข้อมูลเรียบร้อยแล้ว' };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// เขียนไฟล์ข้อมูลกลาง โดยผู้เรียกต้องถือ ScriptLock อยู่แล้ว
function writeSharedData_(data) {
  externalizeDataImages_(data);
  try {
    writeDataToSheet_(data);
    PropertiesService.getScriptProperties().deleteProperty('ACCOUNTING_SYSTEM_DATA');
  } catch (sheetError) {
    Logger.log('Sheet write failed; using Drive backup: ' + sheetError);
    writeLegacyBackup_(data);
  }
}

function writeLegacyBackup_(data) {
  const properties = PropertiesService.getScriptProperties();
  const json = JSON.stringify(data);
  const fileId = properties.getProperty('ACCOUNTING_DATA_FILE_ID');
  let file = null;
  if (fileId) {
    try {
      file = DriveApp.getFileById(fileId);
    } catch (error) {}
  }
  if (file) {
    file.setContent(json);
  } else {
    file = DriveApp.createFile(
      Utilities.newBlob(json, 'application/json', 'accounting-system-data.json')
    );
    properties.setProperty('ACCOUNTING_DATA_FILE_ID', file.getId());
  }
}

function getDataSpreadsheet_() {
  const properties = PropertiesService.getScriptProperties();
  const customId = properties.getProperty('ACCOUNTING_SPREADSHEET_ID') || ACCOUNTING_SPREADSHEET_ID;
  if (customId) {
    try {
      return SpreadsheetApp.openById(customId);
    } catch (e) {
      Logger.log('openById failed for ID ' + customId + ': ' + e);
    }
  }
  try {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  } catch (e) {}
  
  throw new Error('ไม่สามารถเข้าถึง Google Sheet ฐานข้อมูลได้');
}

function getOrCreateDataSheet_() {
  const spreadsheet = getDataSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(ACCOUNTING_DATA_SHEET);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(ACCOUNTING_DATA_SHEET);
    sheet.getRange('A1:B1').setValues([['DATABASE_VERSION', 'UPDATED_AT']]);
    sheet.setFrozenRows(1);
    try { sheet.hideSheet(); } catch (e) {}
  }
  return sheet;
}

function readDataFromSheet_() {
  try {
    const spreadsheet = getDataSpreadsheet_();
    const sheet = spreadsheet.getSheetByName(ACCOUNTING_DATA_SHEET);
    if (!sheet) return '';
    const lastRow = sheet.getLastRow();
    if (lastRow < 3) return '';
    
    const rows = sheet.getRange(3, 1, lastRow - 2, 1).getValues();
    return rows.map(function (row) { return row[0]; }).join('');
  } catch (error) {
    Logger.log('Error reading data from Sheet: ' + error);
    return '';
  }
}

function writeDataToSheet_(data) {
  const json = JSON.stringify(data);
  const chunks = [];
  for (let index = 0; index < json.length; index += ACCOUNTING_DATA_CHUNK_SIZE) {
    chunks.push([json.substring(index, index + ACCOUNTING_DATA_CHUNK_SIZE)]);
  }

  const sheet = getOrCreateDataSheet_();
  const requiredRows = chunks.length + 2;
  if (sheet.getMaxRows() < requiredRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), requiredRows - sheet.getMaxRows());
  }

  sheet.getRange('A1:B1').setValues([['DATABASE_VERSION', 'UPDATED_AT']]);
  sheet.getRange('A2:B2').setValues([[chunks.length, new Date()]]);
  if (chunks.length) {
    sheet.getRange(3, 1, chunks.length, 1).setValues(chunks);
  }
  const staleRows = sheet.getLastRow() - requiredRows;
  if (staleRows > 0) {
    sheet.getRange(requiredRows + 1, 1, staleRows, 1).clearContent();
  }
  SpreadsheetApp.flush();
}

// แยกรูป Base64 ออกจาก JSON เพื่อไม่ให้ไฟล์ฐานข้อมูลโตจนส่งแชทต่อไม่ได้
function externalizeDataImages_(data) {
  (data.chatMessages || []).forEach(function (m) {
    m.image = storeDataImage_(m.image, 'chat');
  });
  (data.notifications || []).forEach(function (n) {
    n.slipImage = storeDataImage_(n.slipImage, 'notification');
    n.imagePreview = storeDataImage_(n.imagePreview, 'notification');
  });
  (data.customers || []).forEach(function (c) {
    c.attachedInvoiceImg = storeDataImage_(c.attachedInvoiceImg, 'invoice');
    (c.invoiceHistory || []).forEach(function (h) {
      h.attachedInvoiceImg = storeDataImage_(h.attachedInvoiceImg, 'invoice-history');
    });
  });
}

function storeDataImage_(value, prefix) {
  if (!value || String(value).indexOf('data:image/') !== 0) return value || '';
  try {
    const match = String(value).match(/^data:(image\/[a-zA-Z0-9.+-]+);base64,(.+)$/);
    if (!match) return value;
    const mimeType = match[1];
    const extension = mimeType.indexOf('png') >= 0 ? 'png' : 'jpg';
    const blob = Utilities.newBlob(
      Utilities.base64Decode(match[2]),
      mimeType,
      prefix + '-' + new Date().getTime() + '.' + extension
    );
    const folder = getImageFolder_();
    const file = folder.createFile(blob);
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch (e) {}
    return 'https://lh3.googleusercontent.com/d/' + file.getId() + '=s1600';
  } catch (error) {
    Logger.log('Image externalization failed: ' + error);
    return value;
  }
}

function getDriveImageBase64(fileIdOrUrl) {
  try {
    const match = String(fileIdOrUrl || '').match(/(?:[?&]id=|\/d\/)([-\w]{20,})/);
    const fileId = match ? match[1] : String(fileIdOrUrl || '').trim();
    if (!fileId) return '';
    const file = DriveApp.getFileById(fileId);
    const blob = file.getBlob();
    return 'data:' + blob.getContentType() + ';base64,' + Utilities.base64Encode(blob.getBytes());
  } catch (error) {
    Logger.log('getDriveImageBase64 error: ' + error);
    return '';
  }
}

function getImageFolder_() {
  const properties = PropertiesService.getScriptProperties();
  const folderId = properties.getProperty('ACCOUNTING_IMAGE_FOLDER_ID');
  if (folderId) {
    try {
      const existingFolder = DriveApp.getFolderById(folderId);
      try {
        existingFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (e) {}
      return existingFolder;
    } catch (e) {}
  }
  const folder = DriveApp.createFolder('Accounting System Attachments');
  try {
    folder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch (e) {}
  properties.setProperty('ACCOUNTING_IMAGE_FOLDER_ID', folder.getId());
  return folder;
}

// เพิ่มข้อความแบบ atomic ป้องกันคำสั่งอ่าน/บันทึกเก่าเขียนทับข้อความใหม่
function getRealtimeSnapshot(role, customerId, customerName) {
  const data = getData();
  const normalizedRole = String(role || '').toLowerCase();
  if (normalizedRole === 'admin' || normalizedRole === 'member') {
    data.snapshotValid = true;
    return data;
  }

  if (normalizedRole === 'customer') {
    const normalizedName = String(customerName || '').trim().toLowerCase();
    return {
      snapshotValid: true,
      columns: data.columns || [],
      customers: (data.customers || []).filter(function (customer) {
        return String(customer.id) === String(customerId);
      }),
      banks: data.banks || [],
      notifications: (data.notifications || []).filter(function (item) {
        return String(item.customerName || '').trim().toLowerCase() === normalizedName;
      }),
      announcements: data.announcements || [],
      chatMessages: (data.chatMessages || []).filter(function (message) {
        return String(message.customerName || '').trim().toLowerCase() === normalizedName;
      })
    };
  }

  return {
    snapshotValid: true,
    columns: data.columns || [],
    customers: data.customers || [],
    banks: data.banks || [],
    announcements: data.announcements || []
  };
}

function appendChatMessage(message) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    data.chatMessages = data.chatMessages || [];
    const retentionDays = Number(
      PropertiesService.getScriptProperties().getProperty('CHAT_RETENTION_DAYS') || 0
    );
    if (retentionDays > 0) {
      const cutoff = new Date().getTime() - retentionDays * 86400000;
      data.chatMessages = data.chatMessages.filter(function (m) {
        return Number(m.id) >= cutoff;
      });
    }
    const exists = data.chatMessages.some(function (m) {
      return String(m.id) === String(message.id);
    });
    if (!exists) data.chatMessages.push(message);
    writeSharedData_(data);
    return { success: true, message: message };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function deleteChatConversation(customerName) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    const target = String(customerName || '').trim().toLowerCase();
    data.chatMessages = (data.chatMessages || []).filter(function (m) {
      return String(m.customerName || '').trim().toLowerCase() !== target;
    });
    writeSharedData_(data);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function setChatRetentionDays(days) {
  const value = Math.max(0, Number(days) || 0);
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    PropertiesService.getScriptProperties()
      .setProperty('CHAT_RETENTION_DAYS', String(value));
    const data = getData();
    if (value > 0) {
      const cutoff = new Date().getTime() - value * 86400000;
      data.chatMessages = (data.chatMessages || []).filter(function (m) {
        return Number(m.id) >= cutoff;
      });
      writeSharedData_(data);
    }
    return { success: true, days: value };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// อัปเดตช่องเดียวในตาราง ลด payload และป้องกันข้อมูลช่องอื่นถูกเขียนทับ
function updateCustomerCell(customerId, columnId, value) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    const customer = (data.customers || []).find(function (c) {
      return String(c.id) === String(customerId);
    });
    if (!customer) throw new Error('ไม่พบข้อมูลลูกค้า');
    customer.values = customer.values || {};
    customer.values[columnId] = Number(value) || 0;
    writeSharedData_(data);
    return { success: true, customerId: customerId, columnId: columnId, value: customer.values[columnId] };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// ส่งบิลโดยแก้ record ลูกค้าบนข้อมูลล่าสุดโดยตรง
function dispatchInvoiceToCustomer(customerId, invoice) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    const customer = (data.customers || []).find(function (c) {
      return String(c.id) === String(customerId);
    });
    if (!customer) throw new Error('ไม่พบข้อมูลลูกค้า');
    customer.invoiceSent = true;
    customer.invoiceDate = invoice.invoiceDate;
    customer.status = 'รอชำระ';
    customer.attachedInvoiceImg = invoice.attachedInvoiceImg || '';
    customer.selectedRefundBank = invoice.selectedRefundBank || null;
    writeSharedData_(data);
    return { success: true, customer: customer };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function confirmCustomerInvoice(customerId, confirmation) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    const customer = (data.customers || []).find(function (c) {
      return String(c.id) === String(customerId);
    });
    if (!customer) throw new Error('ไม่พบข้อมูลลูกค้า');
    const columns = data.columns || [];
    const total = columns.reduce(function (sum, col) {
      return sum + (Number((customer.values || {})[col.id]) || 0);
    }, 0);
    customer.status = 'ชำระแล้ว';
    customer.invoiceSent = false;
    customer.invoiceHistory = customer.invoiceHistory || [];
    customer.invoiceHistory.unshift({
      id: new Date().getTime(),
      total: total,
      date: customer.invoiceDate || confirmation.date,
      paidAt: confirmation.paidAt,
      note: confirmation.note,
      attachedInvoiceImg: customer.attachedInvoiceImg || ''
    });
    customer.attachedInvoiceImg = '';
    writeSharedData_(data);
    return { success: true, customer: customer };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// เปลี่ยนสถานะอ่านแล้วเฉพาะข้อความของคู่สนทนา โดยไม่เขียนทับรายการแชทอื่น
function markChatMessagesRead(customerName, readerRole) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    const target = String(customerName || '').trim().toLowerCase();
    data.chatMessages = (data.chatMessages || []).map(function (m) {
      const sameCustomer = String(m.customerName || '').trim().toLowerCase() === target;
      const shouldMark = readerRole === 'admin'
        ? m.senderRole === 'customer'
        : m.senderRole === 'admin';
      if (sameCustomer && shouldMark) m.isRead = true;
      return m;
    });
    writeSharedData_(data);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}
