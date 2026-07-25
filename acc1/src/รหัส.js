/**
 * AccStore Enterprise - Google Apps Script Backend
 * 
 * doGet(e) handles HTTP GET requests and renders index.html
 */

function doGet(e) {
  return HtmlService.createHtmlOutputFromFile('index')
      .setTitle('AccStore Enterprise | ระบบบัญชีและบริหารจัดการหลายร้านค้า')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0');
}

/**
 * Initialize Google Sheets Database structure and initial seed data
 * Run this function once from Apps Script Editor or from Web UI before first use
 * @return {Object} Response status
 */
function initDatabase() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // 1. Setup 'Stores' Sheet
    var storesSheet = ss.getSheetByName('Stores');
    if (!storesSheet) {
      storesSheet = ss.insertSheet('Stores');
      storesSheet.appendRow(['id', 'name', 'code', 'category', 'initialBalance', 'color', 'taxId', 'address']);
      
      // Format Header Row
      var headerRange = storesSheet.getRange(1, 1, 1, 8);
      headerRange.setBackground('#0f172a').setFontColor('#10b981').setFontWeight('bold');
      storesSheet.setFrozenRows(1);

      // Seed Initial Default Stores
      var seedStores = [
        ['store-1', 'MB Coffee & Bakery', 'ST-001', 'ร้านอาหาร/เครื่องดื่ม', 50000, '#f59e0b', '0105565000011', '123/4 ถนนสุขุมวิท กรุงเทพมหานคร โทร 081-111-2222'],
        ['store-2', 'MB Fresh Minimart', 'ST-002', 'ร้านค้าปลีก/มินิมาร์ท', 120000, '#10b981', '0105565000022', '55/9 ถนนพหลโยธิน ปทุมธานี โทร 082-333-4444'],
        ['store-3', 'MB Auto Detailing', 'ST-003', 'ศูนย์บริการ/คาร์แคร์', 85000, '#0ea5e9', '0105565000033', '88/1 ถนนราชพฤกษ์ นนทบุรี โทร 089-555-6666']
      ];
      seedStores.forEach(function(row) { storesSheet.appendRow(row); });
    }

    // 2. Setup 'Transactions' Sheet
    var txSheet = ss.getSheetByName('Transactions');
    if (!txSheet) {
      txSheet = ss.insertSheet('Transactions');
      txSheet.appendRow(['id', 'storeId', 'type', 'amount', 'date', 'category', 'paymentMethod', 'note']);
      
      // Format Header Row
      var txHeaderRange = txSheet.getRange(1, 1, 1, 8);
      txHeaderRange.setBackground('#0f172a').setFontColor('#0ea5e9').setFontWeight('bold');
      txSheet.setFrozenRows(1);

      // Seed Initial Default Transactions
      var seedTx = [
        ['tx-101', 'store-1', 'income', 14500, '2026-07-25T09:30', 'ยอดขายหน้าร้าน', 'โอนเงิน/สแกน QR', 'ยอดขายกาแฟและเบเกอรี่ประจำวัน'],
        ['tx-102', 'store-1', 'expense', 4200, '2026-07-24T14:15', 'วัตถุดิบและสต็อก', 'โอนเงิน/สแกน QR', 'สั่งซื้อเมล็ดกาแฟอาราบิก้าเกรดพรีเมียม'],
        ['tx-103', 'store-2', 'income', 28900, '2026-07-25T11:00', 'ยอดขายหน้าร้าน', 'เงินสด', 'ยอดขายสินค้าอุปโภคบริโภคประจำวัน'],
        ['tx-104', 'store-2', 'expense', 8500, '2026-07-23T16:00', 'ค่าเช่าและสาธารณูปโภค', 'โอนเงิน/สแกน QR', 'ชำระค่าน้ำค่าไฟมินิมาร์ทประจำเดือน'],
        ['tx-105', 'store-3', 'income', 19500, '2026-07-25T12:00', 'ค่าบริการคาร์แคร์', 'บัตรเครดิต', 'บริการเคลือบแก้วและล้างรถยนต์ 5 คัน'],
        ['tx-106', 'store-3', 'expense', 6000, '2026-07-22T17:30', 'เงินเดือนและค่าแรง', 'โอนเงิน/สแกน QR', 'ค่าแรงทีมช่างล้างรถล่วงเวลา']
      ];
      seedTx.forEach(function(row) { txSheet.appendRow(row); });
    }

    // 3. Setup 'Users' Sheet (User Credentials & Role Permissions)
    var usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) {
      usersSheet = ss.insertSheet('Users');
      usersSheet.appendRow(['username', 'password', 'name', 'role', 'assignedStoreId']);
      
      var usersHeaderRange = usersSheet.getRange(1, 1, 1, 5);
      usersHeaderRange.setBackground('#0f172a').setFontColor('#8b5cf6').setFontWeight('bold');
      usersSheet.setFrozenRows(1);

      var seedUsers = [
        ['admin', 'admin123', 'ผู้ดูแลระบบ (Admin)', 'admin', 'ALL'],
        ['editor', 'editor123', 'เจ้าหน้าที่บัญชี (Editor)', 'editor', 'ALL'],
        ['viewer', 'viewer123', 'ผู้ตรวจการ (Viewer)', 'viewer', 'ALL']
      ];
      seedUsers.forEach(function(row) { usersSheet.appendRow(row); });
    }

    return {
      success: true,
      message: 'เริ่มต้นฐานข้อมูล Google Sheets (Stores, Transactions & Users) สำเร็จ!'
    };
  } catch (error) {
    Logger.log('Error initDatabase: ' + error.toString());
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการตั้งค่า Sheet: ' + error.toString()
    };
  }
}

/**
 * Login Authentication Function
 */
function loginUser(username, password) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var usersSheet = ss.getSheetByName('Users');
    if (!usersSheet) {
      initDatabase();
      usersSheet = ss.getSheetByName('Users');
    }

    var users = parseSheetToObjects(usersSheet);
    var matchedUser = null;

    for (var i = 0; i < users.length; i++) {
      if (String(users[i].username).trim() === String(username).trim() && 
          String(users[i].password).trim() === String(password).trim()) {
        matchedUser = users[i];
        break;
      }
    }

    if (matchedUser) {
      return {
        success: true,
        user: {
          username: matchedUser.username,
          name: matchedUser.name,
          role: matchedUser.role, // 'admin' | 'editor' | 'viewer'
          assignedStoreId: matchedUser.assignedStoreId || 'ALL'
        }
      };
    } else {
      return {
        success: false,
        message: 'ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง!'
      };
    }
  } catch (error) {
    return {
      success: false,
      message: 'เกิดข้อผิดพลาดในการเข้าสู่ระบบ: ' + error.toString()
    };
  }
}

/**
 * Fetch all stores, transactions, and users data from Google Sheets
 * @return {Object} Application Data
 */
function getAppData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Auto-init if sheets don't exist
    if (!ss.getSheetByName('Stores') || !ss.getSheetByName('Transactions') || !ss.getSheetByName('Users')) {
      initDatabase();
    }

    var storesSheet = ss.getSheetByName('Stores');
    var txSheet = ss.getSheetByName('Transactions');
    var usersSheet = ss.getSheetByName('Users');

    var stores = parseSheetToObjects(storesSheet);
    var transactions = parseSheetToObjects(txSheet);
    var users = parseSheetToObjects(usersSheet);

    return {
      success: true,
      stores: stores,
      transactions: transactions,
      users: users
    };
  } catch (error) {
    Logger.log('Error getAppData: ' + error.toString());
    return {
      success: false,
      message: error.toString()
    };
  }
}

/**
 * Save or Update a Store in Google Sheet
 */
function saveStoreData(storeObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Stores');
    if (!sheet) {
      initDatabase();
      sheet = ss.getSheetByName('Stores');
    }

    var data = sheet.getDataRange().getValues();
    var rowIndex = -1;

    for (var i = 1; i < data.length; i++) {
      if (String(data[i][0]) === String(storeObj.id)) {
        rowIndex = i + 1;
        break;
      }
    }

    var rowValues = [
      String(storeObj.id),
      String(storeObj.name || ''),
      String(storeObj.code || ''),
      String(storeObj.category || ''),
      Number(storeObj.initialBalance) || 0,
      String(storeObj.color || '#10b981'),
      String(storeObj.taxId || ''),
      String(storeObj.address || '')
    ];

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      sheet.appendRow(rowValues);
    }

    return { success: true, message: 'บันทึกข้อมูลร้านค้าลง Google Sheet เรียบร้อยแล้ว' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Delete a Store and its transactions from Google Sheet
 */
function deleteStoreData(storeId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    
    // Delete Store Row
    var storeSheet = ss.getSheetByName('Stores');
    if (storeSheet) {
      var data = storeSheet.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(storeId)) {
          storeSheet.deleteRow(i + 1);
        }
      }
    }

    // Delete Store Transactions Rows
    var txSheet = ss.getSheetByName('Transactions');
    if (txSheet) {
      var txData = txSheet.getDataRange().getValues();
      for (var j = txData.length - 1; j >= 1; j--) {
        if (String(txData[j][1]) === String(storeId)) {
          txSheet.deleteRow(j + 1);
        }
      }
    }

    return { success: true, message: 'ลบข้อมูลร้านค้าเรียบร้อยแล้ว' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Save a Transaction in Google Sheet
 */
function saveTransactionData(txObj) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Transactions');
    if (!sheet) {
      initDatabase();
      sheet = ss.getSheetByName('Transactions');
    }

    var rowValues = [
      String(txObj.id),
      String(txObj.storeId),
      String(txObj.type),
      Number(txObj.amount) || 0,
      String(txObj.date),
      String(txObj.category),
      String(txObj.paymentMethod),
      String(txObj.note || '')
    ];

    sheet.appendRow(rowValues);
    return { success: true, message: 'บันทึกรายการลง Google Sheet เรียบร้อยแล้ว' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Delete a Transaction from Google Sheet
 */
function deleteTransactionData(txId) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('Transactions');
    if (sheet) {
      var data = sheet.getDataRange().getValues();
      for (var i = data.length - 1; i >= 1; i--) {
        if (String(data[i][0]) === String(txId)) {
          sheet.deleteRow(i + 1);
          break;
        }
      }
    }
    return { success: true, message: 'ลบรายการบัญชีเรียบร้อยแล้ว' };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

/**
 * Helper: Convert 2D Sheet array to Object Array using Row 1 Headers with clean type parsing
 */
function parseSheetToObjects(sheet) {
  if (!sheet) return [];
  var rawValues = sheet.getDataRange().getValues();
  var displayValues = sheet.getDataRange().getDisplayValues();
  if (rawValues.length <= 1) return [];

  var headers = rawValues[0];
  var result = [];

  for (var i = 1; i < rawValues.length; i++) {
    var row = rawValues[i];
    var dispRow = displayValues[i];
    var obj = {};

    for (var j = 0; j < headers.length; j++) {
      var key = String(headers[j]).trim();
      var val = row[j];

      if (key === 'amount' || key === 'initialBalance') {
        val = Number(val) || 0;
      } else if (val instanceof Date) {
        val = dispRow[j] || val.toISOString();
      } else {
        val = String(val);
      }

      obj[key] = val;
    }
    result.push(obj);
  }

  return result;
}
