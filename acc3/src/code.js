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
    processWeeklyOutstandingBalances({ dryRun: false });
  } catch (error) {
    Logger.log('doGet initialize error: ' + error);
  }
  try {
    ensureAccountingAutomationTrigger_();
  } catch (triggerError) {
    Logger.log('Automation trigger setup error: ' + triggerError);
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
    customers: [],
    banks: [
      { id: 1, bankName: 'กสิกรไทย (KBANK)', accNo: '123-4-56789-0', accName: 'นายบัญชี รับเงิน', promptpay: '0812345678' },
      { id: 2, bankName: 'ไทยพาณิชย์ (SCB)', accNo: '987-6-54321-0', accName: 'นายบัญชี รับเงิน', promptpay: '0812345678' }
    ],
    notifications: [],
    announcements: [],
    users: [
      { username: 'admin', password: 'admin123', role: 'admin', name: 'แอดมิน' }
    ],
    transactions: [],
    chatMessages: []
  };
}

// One-time production cleanup. Preserve system structure, banks, columns and
// non-customer staff accounts while removing all trial/customer activity.
function resetProductionCustomerData() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    data.customers = [];
    data.users = (data.users || []).filter(function (user) {
      return String(user.role || '').toLowerCase() !== 'customer';
    });
    data.notifications = [];
    data.chatMessages = [];
    data.transactions = [];
    data.announcements = [];
    writeSharedData_(data);
    return {
      success: true,
      customers: data.customers.length,
      customerUsers: data.users.filter(function (user) {
        return String(user.role || '').toLowerCase() === 'customer';
      }).length,
      notifications: data.notifications.length,
      chatMessages: data.chatMessages.length,
      transactions: data.transactions.length
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

// Older customer clients could save a filtered snapshot containing only their
// own row. Recreate any missing customer rows from the authoritative user list
// so an existing login can never silently disappear from the admin table.
function repairCustomerRecords_(data) {
  if (!data || typeof data !== 'object') return data;
  data.customers = Array.isArray(data.customers) ? data.customers : [];
  data.notifications = Array.isArray(data.notifications) ? data.notifications : [];
  normalizeOutstandingNotifications_(data.notifications);
  const existingIds = {};
  data.customers.forEach(function (customer) {
    existingIds[String(customer.id)] = true;
    customer.values = customer.values || {};
    customer.invoiceHistory = customer.invoiceHistory || [];
    customer.banks = customer.banks || [];
    const hasTableValue = (data.columns || []).some(function (column) {
      const numericValue = Number(customer.values[column.id]);
      return Number.isFinite(numericValue) && numericValue !== 0;
    });
    if (!hasTableValue) {
      customer.status = 'ยังไม่ดำเนินการ';
    } else if (hasTableValue && customer.status !== 'ชำระแล้ว') {
      customer.status = 'รอชำระ';
    }
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

// Notifications are stored newest-first. Older clients kept every approved
// partial payment in the outstanding tab, so one customer could appear there
// more than once. Treat the newest card as authoritative and archive the rest.
function normalizeOutstandingNotifications_(notifications) {
  const outstandingByCustomer = {};
  (notifications || []).forEach(function (item, index) {
    const remaining = Number(item.remainingAfterPayment);
    const isOutstanding =
      item.status === 'pending_followup' ||
      (
        item.status === 'pending' &&
        item.paymentApplied &&
        Number.isFinite(remaining) &&
        remaining > 0
      ) ||
      (
        item.paymentApplied &&
        item.supersededByNotificationId &&
        Number.isFinite(remaining) &&
        remaining > 0
      );
    if (!isOutstanding) return;

    const customerKey = String(item.customerName || '').trim().toLowerCase();
    if (!customerKey) return;
    outstandingByCustomer[customerKey] = outstandingByCustomer[customerKey] || [];
    outstandingByCustomer[customerKey].push({ item: item, index: index });
  });

  Object.keys(outstandingByCustomer).forEach(function (customerKey) {
    const entries = outstandingByCustomer[customerKey];
    if (entries.length < 2) return;

    // Payment notification IDs are generated with Date.now(). Select the
    // greatest ID, rather than relying on array order, because legacy saves
    // could reorder notifications during a realtime merge.
    const latestEntry = entries.reduce(function (latest, entry) {
      const latestId = Number(latest.item.id);
      const entryId = Number(entry.item.id);
      if (Number.isFinite(entryId) && Number.isFinite(latestId)) {
        return entryId > latestId ? entry : latest;
      }
      if (Number.isFinite(entryId)) return entry;
      if (Number.isFinite(latestId)) return latest;
      return entry.index < latest.index ? entry : latest;
    }, entries[0]);

    latestEntry.item.status = 'pending_followup';
    latestEntry.item.keepUntilPaid = true;
    delete latestEntry.item.supersededByNotificationId;

    entries.forEach(function (entry) {
      if (entry === latestEntry) return;
      entry.item.status = 'approved';
      entry.item.keepUntilPaid = false;
      entry.item.closedAt =
        entry.item.closedAt ||
        latestEntry.item.approvedAt ||
        latestEntry.item.createdAt ||
        '';
      entry.item.supersededByNotificationId = latestEntry.item.id;
    });
  });
  return notifications;
}

// ตรวจสอบการเข้าสู่ระบบ
function authenticateUser(username, password) {
  const data = getData();
  const users = (data && Array.isArray(data.users)) ? data.users : [];

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

// Delete an account atomically from the Sheet-backed database. Customer
// accounts also remove their linked customer row so realtime sync cannot bring
// a deleted member back.
function deleteUserAccountRecord(username) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const target = String(username || '').trim().toLowerCase();
    if (!target) throw new Error('กรุณาระบุบัญชีที่ต้องการลบ');
    if (target === 'admin') throw new Error('ไม่สามารถลบบัญชี Admin หลักได้');

    const data = getData();
    const existingUser = (data.users || []).find(function (user) {
      return String(user.username || '').trim().toLowerCase() === target;
    });
    if (!existingUser) return { success: true, alreadyDeleted: true };

    data.users = (data.users || []).filter(function (user) {
      return String(user.username || '').trim().toLowerCase() !== target;
    });
    if (existingUser.customerId !== null && existingUser.customerId !== undefined) {
      data.customers = (data.customers || []).filter(function (customer) {
        return String(customer.id) !== String(existingUser.customerId);
      });
    }
    writeSharedData_(data);
    return {
      success: true,
      username: existingUser.username,
      customerId: existingUser.customerId || null
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
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

// ตรวจทุกวัน และมี fallback ตอนเปิดเว็บ เผื่อ trigger ถูกลบหรือหยุดทำงาน
function ensureAccountingAutomationTrigger_() {
  const properties = PropertiesService.getScriptProperties();
  if (properties.getProperty('ACCOUNTING_AUTOMATION_TRIGGER_READY') === 'true') return;
  const handler = 'runDailyAccountingAutomation';
  const exists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === handler;
  });
  if (!exists) {
    ScriptApp.newTrigger(handler).timeBased().everyDays(1).atHour(1).create();
  }
  properties.setProperty('ACCOUNTING_AUTOMATION_TRIGGER_READY', 'true');
}

function runDailyAccountingAutomation() {
  return processWeeklyOutstandingBalances({ dryRun: false });
}

// ย้ายยอดที่ค้างครบ 7 วันเข้า "ยอดเบิก" โดยคงยอดรวมเดิม และส่งรายละเอียดเข้าแชท
// options: { dryRun: true, customerName: 'ยุทธ', now: Date|string }
function processWeeklyOutstandingBalances(options) {
  options = options || {};
  const dryRun = options.dryRun === true;
  const targetName = String(options.customerName || '').trim().toLowerCase();
  const now = options.now ? new Date(options.now) : new Date();
  if (isNaN(now.getTime())) throw new Error('วันที่ตรวจสอบไม่ถูกต้อง');

  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    const columns = data.columns || [];
    const withdrawColumn = columns.find(function (column) {
      return column.id === 'withdraw' || String(column.label || '').trim() === 'ยอดเบิก';
    });
    if (!withdrawColumn) throw new Error('ไม่พบช่องยอดเบิก');

    const results = [];
    (data.customers || []).forEach(function (customer) {
      if (targetName && String(customer.name || '').trim().toLowerCase().indexOf(targetName) < 0) return;
      const invoiceDate = parseAccountingDate_(customer.invoiceDate);
      const ageDays = invoiceDate
        ? Math.floor((startOfAccountingDay_(now).getTime() - startOfAccountingDay_(invoiceDate).getTime()) / 86400000)
        : -1;
      const values = customer.values || {};
      const detailColumns = columns.filter(function (column) {
        return column.id !== withdrawColumn.id && (Number(values[column.id]) || 0) !== 0;
      });
      const amountToCarry = detailColumns.reduce(function (sum, column) {
        return sum + (Number(values[column.id]) || 0);
      }, 0);
      const currentTotal = columns.reduce(function (sum, column) {
        return sum + (Number(values[column.id]) || 0);
      }, 0);
      if (!customer.invoiceSent || ageDays < 7 || currentTotal <= 0 || amountToCarry === 0) return;

      const details = detailColumns.map(function (column) {
        return { id: column.id, label: column.label, amount: Number(values[column.id]) || 0 };
      });
      const result = {
        customerId: customer.id,
        customerName: customer.name,
        invoiceDate: customer.invoiceDate,
        ageDays: ageDays,
        amountToCarry: amountToCarry,
        previousWithdraw: Number(values[withdrawColumn.id]) || 0,
        newWithdraw: (Number(values[withdrawColumn.id]) || 0) + amountToCarry,
        details: details
      };
      results.push(result);
      if (dryRun) return;

      details.forEach(function (detail) { values[detail.id] = 0; });
      values[withdrawColumn.id] = result.newWithdraw;
      customer.values = values;
      customer.invoiceDate = formatAccountingDate_(now);
      customer.lastWeeklyCarryoverAt = now.toISOString();
      customer.status = 'รอชำระ';
      customer.adminConfirmed = false;
      customer.customerConfirmed = false;
      data.chatMessages = data.chatMessages || [];
      data.chatMessages.push(buildSystemChatMessage_(customer.name,
        'แจ้งยอดค้างครบ 7 วัน\n' +
        details.map(function (detail) {
          return '• ' + detail.label + ': ' + formatAccountingAmount_(detail.amount) + ' บาท';
        }).join('\n') +
        '\nรวมย้ายเข้าช่องยอดเบิก: ' + formatAccountingAmount_(amountToCarry) + ' บาท\n' +
        'ยอดค้างรวมปัจจุบัน: ' + formatAccountingAmount_(currentTotal) + ' บาท',
        '', now));
    });

    if (!dryRun && results.length) writeSharedData_(data);
    return { success: true, dryRun: dryRun, processed: results.length, results: results };
  } catch (error) {
    return { success: false, dryRun: dryRun, message: error.toString(), processed: 0, results: [] };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function previewWeeklyOutstandingForCustomer(customerName) {
  return processWeeklyOutstandingBalances({ dryRun: true, customerName: customerName });
}

function previewYutWeeklyOutstanding() {
  return previewWeeklyOutstandingForCustomer('ยุทธ');
}

function completeCustomerRefund(customerId, refund) {
  refund = refund || {};
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    const customer = (data.customers || []).find(function (item) {
      return String(item.id) === String(customerId);
    });
    if (!customer) throw new Error('ไม่พบข้อมูลลูกค้า');
    const columns = data.columns || [];
    const total = columns.reduce(function (sum, column) {
      return sum + (Number((customer.values || {})[column.id]) || 0);
    }, 0);
    if (total >= 0) throw new Error('ลูกค้ารายนี้ไม่มียอดติดลบสำหรับโอนคืน');
    const slipImage = refund.slipImage || customer.attachedInvoiceImg || '';
    if (!slipImage) throw new Error('กรุณาแนบสลิปโอนเงินคืน');
    const now = new Date();
    const refundAmount = Math.abs(total);
    const bank = refund.bank || customer.selectedRefundBank || null;
    const bankText = bank
      ? [bank.bankName, bank.accNo, bank.accName].filter(Boolean).join(' - ')
      : 'ไม่ระบุบัญชี';

    customer.invoiceHistory = customer.invoiceHistory || [];
    customer.invoiceHistory.unshift({
      id: now.getTime(), total: total, paidAmount: refundAmount, remainingAmount: 0,
      date: customer.invoiceDate || formatAccountingDate_(now),
      paidAt: now.toLocaleString('th-TH'),
      note: 'แอดมินโอนเงินคืนและเคลียร์ยอด เข้าบัญชี: ' + bankText,
      attachedInvoiceImg: slipImage
    });
    customer.values = customer.values || {};
    columns.forEach(function (column) { customer.values[column.id] = 0; });
    customer.status = 'ชำระแล้ว';
    customer.invoiceSent = false;
    customer.adminConfirmed = true;
    customer.customerConfirmed = true;
    customer.pendingPaymentAmount = 0;
    customer.attachedInvoiceImg = '';
    customer.selectedRefundBank = null;

    data.transactions = data.transactions || [];
    data.transactions.unshift({
      id: now.getTime(), colLabel: refund.colLabel || 'EDAY', type: 'expense',
      amount: refundAmount,
      startDate: Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd'),
      endDate: Utilities.formatDate(now, 'Asia/Bangkok', 'yyyy-MM-dd'),
      note: 'โอนเงินคืนยอดติดลบให้คุณ ' + customer.name + ' (เข้าบัญชี: ' + bankText + ')',
      createdAt: now.toISOString()
    });
    data.chatMessages = data.chatMessages || [];
    data.chatMessages.push(buildSystemChatMessage_(customer.name,
      'โอนเงินคืนเรียบร้อยแล้ว\nจำนวน: ' + formatAccountingAmount_(refundAmount) +
      ' บาท\nบัญชีรับเงิน: ' + bankText + '\nระบบเคลียร์ยอดในตารางเป็น 0 แล้ว',
      slipImage, now));
    writeSharedData_(data);
    return { success: true, customer: customer, refundAmount: refundAmount, chatMessage: data.chatMessages[data.chatMessages.length - 1] };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function buildSystemChatMessage_(customerName, text, image, date) {
  const now = date || new Date();
  return {
    id: now.getTime(), customerName: customerName, senderName: 'ระบบ', senderRole: 'admin',
    text: text, image: image || '', isRead: false,
    createdAt: now.toLocaleTimeString('th-TH', { hour: '2-digit', minute: '2-digit' })
  };
}

function parseAccountingDate_(value) {
  if (!value) return null;
  if (value instanceof Date && !isNaN(value.getTime())) return value;
  const text = String(value).trim();
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) return parsed;
  const numericMatch = text.match(/^(\d{1,2})[\/-](\d{1,2})[\/-](\d{4})$/);
  if (numericMatch) {
    let numericYear = Number(numericMatch[3]);
    if (numericYear > 2400) numericYear -= 543;
    return new Date(numericYear, Number(numericMatch[2]) - 1, Number(numericMatch[1]));
  }
  const thaiMonths = { มกราคม:0, กุมภาพันธ์:1, มีนาคม:2, เมษายน:3, พฤษภาคม:4, มิถุนายน:5, กรกฎาคม:6, สิงหาคม:7, กันยายน:8, ตุลาคม:9, พฤศจิกายน:10, ธันวาคม:11 };
  const match = text.match(/(\d{1,2})\s+([^\s]+)\s+(\d{4})/);
  if (!match || thaiMonths[match[2]] === undefined) return null;
  let year = Number(match[3]);
  if (year > 2400) year -= 543;
  return new Date(year, thaiMonths[match[2]], Number(match[1]));
}

function startOfAccountingDay_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function formatAccountingDate_(date) {
  return date.toLocaleDateString('th-TH', { year: 'numeric', month: 'long', day: 'numeric' });
}

function formatAccountingAmount_(value) {
  return (Number(value) || 0).toLocaleString('th-TH', { minimumFractionDigits: 0, maximumFractionDigits: 2 });
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
    const hasValue = (data.columns || []).some(function (column) {
      const numericValue = Number(customer.values[column.id]);
      return Number.isFinite(numericValue) && numericValue !== 0;
    });
    customer.status = hasValue ? 'รอชำระ' : 'ยังไม่ดำเนินการ';
    customer.adminConfirmed = false;
    customer.customerConfirmed = false;
    customer.pendingPaymentAmount = 0;
    if (!hasValue) {
      customer.invoiceSent = false;
      customer.attachedInvoiceImg = '';
      customer.selectedRefundBank = null;
    }
    writeSharedData_(data);
    return {
      success: true,
      customerId: customerId,
      columnId: columnId,
      value: customer.values[columnId],
      status: customer.status,
      customer: customer
    };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function clearCustomerValues(customerId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    const customer = (data.customers || []).find(function (item) {
      return String(item.id) === String(customerId);
    });
    if (!customer) throw new Error('ไม่พบข้อมูลลูกค้า');
    customer.values = customer.values || {};
    (data.columns || []).forEach(function (column) {
      customer.values[column.id] = 0;
    });
    customer.status = 'ยังไม่ดำเนินการ';
    customer.invoiceSent = false;
    customer.adminConfirmed = false;
    customer.customerConfirmed = false;
    customer.pendingPaymentAmount = 0;
    customer.attachedInvoiceImg = '';
    customer.selectedRefundBank = null;
    writeSharedData_(data);
    return { success: true, customer: customer };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function clearAllCustomerValues() {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    (data.customers || []).forEach(function (customer) {
      customer.values = customer.values || {};
      (data.columns || []).forEach(function (column) {
        customer.values[column.id] = 0;
      });
      customer.status = 'ยังไม่ดำเนินการ';
      customer.invoiceSent = false;
      customer.adminConfirmed = false;
      customer.customerConfirmed = false;
      customer.pendingPaymentAmount = 0;
      customer.attachedInvoiceImg = '';
      customer.selectedRefundBank = null;
    });
    writeSharedData_(data);
    return { success: true, customers: data.customers || [] };
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
    customer.adminConfirmed = false;
    customer.customerConfirmed = false;
    customer.pendingPaymentAmount = 0;
    customer.attachedInvoiceImg = invoice.attachedInvoiceImg || '';
    customer.selectedRefundBank = invoice.selectedRefundBank || null;
    customer.selectedPaymentBanks = Array.isArray(invoice.selectedPaymentBanks)
      ? invoice.selectedPaymentBanks.slice(0, 1)
      : [];
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
    customer.customerConfirmed = true;
    if (!customer.adminConfirmed) {
      customer.status = 'รอชำระ';
      writeSharedData_(data);
      return {
        success: true,
        completed: false,
        message: 'บันทึกการยืนยันของลูกค้าแล้ว รอแอดมินตรวจสอบข้อมูล',
        customer: customer
      };
    }

    const requestedPayment = Number(customer.pendingPaymentAmount) || 0;
    const isPartialPayment = total > 0 && requestedPayment > 0 && requestedPayment < total;
    const paidAmount = total > 0 && requestedPayment > 0
      ? Math.min(requestedPayment, total)
      : Math.abs(total);

    customer.invoiceHistory = customer.invoiceHistory || [];
    customer.invoiceHistory.unshift({
      id: new Date().getTime(),
      total: total,
      paidAmount: paidAmount,
      remainingAmount: isPartialPayment ? total - paidAmount : 0,
      date: customer.invoiceDate || confirmation.date,
      paidAt: confirmation.paidAt,
      note: isPartialPayment
        ? 'ชำระบางส่วน ' + paidAmount + ' บาท ยอดคงค้าง ' + (total - paidAmount) + ' บาท'
        : confirmation.note,
      attachedInvoiceImg: customer.attachedInvoiceImg || ''
    });

    if (isPartialPayment) {
      let unappliedPayment = paidAmount;
      customer.values = customer.values || {};
      columns.forEach(function (col) {
        if (unappliedPayment <= 0) return;
        const currentValue = Number(customer.values[col.id]) || 0;
        if (currentValue <= 0) return;
        const deduction = Math.min(currentValue, unappliedPayment);
        customer.values[col.id] = currentValue - deduction;
        unappliedPayment -= deduction;
      });
      customer.status = 'รอชำระ';
      customer.invoiceSent = true;
      customer.adminConfirmed = false;
      customer.customerConfirmed = false;
      customer.pendingPaymentAmount = 0;
      customer.attachedInvoiceImg = '';
      writeSharedData_(data);
      return {
        success: true,
        completed: false,
        partial: true,
        paidAmount: paidAmount,
        remainingAmount: total - paidAmount,
        message: 'รับชำระบางส่วนแล้ว คงเหลือ ' + (total - paidAmount) + ' บาท',
        customer: customer
      };
    }

    customer.status = 'ชำระแล้ว';
    customer.invoiceSent = false;
    customer.attachedInvoiceImg = '';
    customer.selectedRefundBank = null;
    customer.pendingPaymentAmount = 0;
    customer.values = customer.values || {};
    columns.forEach(function (col) {
      customer.values[col.id] = 0;
    });
    writeSharedData_(data);
    return { success: true, completed: true, customer: customer };
  } catch (error) {
    return { success: false, message: error.toString() };
  } finally {
    if (lock.hasLock()) lock.releaseLock();
  }
}

function approveCustomerPayment(customerId, paymentAmount, approval, notificationId) {
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(10000);
    const data = getData();
    const customer = (data.customers || []).find(function (item) {
      return String(item.id) === String(customerId);
    });
    if (!customer) throw new Error('ไม่พบข้อมูลลูกค้า');

    const columns = data.columns || [];
    const total = columns.reduce(function (sum, column) {
      return sum + (Number((customer.values || {})[column.id]) || 0);
    }, 0);
    if (total <= 0) throw new Error('ลูกค้ารายนี้ไม่มียอดค้างชำระ');

    const paidAmount = Math.min(Math.max(0, Number(paymentAmount) || 0), total);
    if (paidAmount <= 0) throw new Error('ยอดชำระไม่ถูกต้อง');
    const remainingAmount = Math.max(0, total - paidAmount);

    customer.invoiceHistory = customer.invoiceHistory || [];
    customer.invoiceHistory.unshift({
      id: new Date().getTime(),
      total: total,
      paidAmount: paidAmount,
      remainingAmount: remainingAmount,
      date: customer.invoiceDate || (approval && approval.date),
      paidAt: (approval && approval.paidAt) || new Date().toLocaleString('th-TH'),
      note: remainingAmount > 0
        ? 'แอดมินยืนยันชำระบางส่วน ' + paidAmount + ' บาท ยอดคงค้าง ' + remainingAmount + ' บาท'
        : 'แอดมินตรวจสอบและยืนยันยอดชำระครบแล้ว',
      attachedInvoiceImg: customer.attachedInvoiceImg || ''
    });

    customer.values = customer.values || {};
    if (remainingAmount > 0) {
      let unappliedPayment = paidAmount;
      columns.forEach(function (column) {
        if (unappliedPayment <= 0) return;
        const currentValue = Number(customer.values[column.id]) || 0;
        if (currentValue <= 0) return;
        const deduction = Math.min(currentValue, unappliedPayment);
        customer.values[column.id] = currentValue - deduction;
        unappliedPayment -= deduction;
      });
      customer.status = 'รอชำระ';
      customer.invoiceSent = true;
    } else {
      columns.forEach(function (column) {
        customer.values[column.id] = 0;
      });
      customer.status = 'ชำระแล้ว';
      customer.invoiceSent = false;
      customer.attachedInvoiceImg = '';
      customer.selectedRefundBank = null;
    }
    customer.adminConfirmed = false;
    customer.customerConfirmed = false;
    customer.pendingPaymentAmount = 0;

    const targetNotification = (data.notifications || []).find(function (item) {
      return String(item.id) === String(notificationId);
    });
    const approvedAt = (approval && approval.paidAt) || new Date().toLocaleString('th-TH');
    const targetName = String(customer.name || '').trim().toLowerCase();

    // A new approved payment supersedes the customer's previous outstanding
    // notification. Keep only the latest payment in the outstanding tab so
    // earlier partial-payment cards do not repeat the same current balance.
    (data.notifications || []).forEach(function (item) {
      const isTarget = String(item.id) === String(notificationId);
      const isSameCustomer =
        String(item.customerName || '').trim().toLowerCase() === targetName;
      const isPreviousOutstanding =
        item.status === 'pending_followup' ||
        (
          item.status === 'pending' &&
          item.paymentApplied &&
          Number(item.remainingAfterPayment) > 0
        );
      if (!isTarget && isSameCustomer && isPreviousOutstanding) {
        item.status = 'approved';
        item.keepUntilPaid = false;
        item.closedAt = approvedAt;
        item.supersededByNotificationId = notificationId;
      }
    });

    if (targetNotification) {
      targetNotification.paymentApplied = true;
      targetNotification.processing = false;
      targetNotification.approvedAt = approvedAt;
      targetNotification.remainingAfterPayment = remainingAmount;
      targetNotification.keepUntilPaid = remainingAmount > 0;
      targetNotification.status = remainingAmount > 0 ? 'pending_followup' : 'approved';
    }

    if (remainingAmount === 0) {
      (data.notifications || []).forEach(function (item) {
        if (
          item.status === 'pending_followup' &&
          String(item.customerName || '').trim().toLowerCase() === targetName
        ) {
          item.status = 'approved';
          item.closedAt = approvedAt;
        }
      });
    }

    writeSharedData_(data);
    return {
      success: true,
      completed: remainingAmount === 0,
      partial: remainingAmount > 0,
      paidAmount: paidAmount,
      remainingAmount: remainingAmount,
      customer: customer,
      notification: targetNotification || null,
      notifications: data.notifications || []
    };
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
