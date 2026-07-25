/**
 * PROJECT ACCOUNTING WEB APP — RELEASE V15.5.0
 * Multi-entry up to 50 rows · Document types · Mobile responsive
 */
const APP = Object.freeze({
  spreadsheetProperty: 'PROJECT_ACCOUNTING_SPREADSHEET_ID',
  schemaProperty: 'PROJECT_ACCOUNTING_SCHEMA_VERSION',
  schemaVersion: '16',
  version: '15.5.0',
  projectsSheet: 'Projects',
  transactionsSheet: 'Transactions',
  usersSheet: 'Users',
  auditSheet: 'AuditLog',
  appUsersSheet: 'AppUsers',
  subprojectsSheet: 'Subprojects',
  categoriesSheet: 'Categories',
  centralTransactionsSheet: 'CentralTransactions',
  centralCategoriesSheet: 'CentralCategories',
  documentTypesSheet: 'DocumentTypes',
  timezone: 'Asia/Bangkok',
  cacheKey: 'PROJECT_ACCOUNTING_INITIAL_DATA_V16',
  authUsersCacheKey: 'PROJECT_ACCOUNTING_AUTH_USERS_V1',
  cacheSeconds: 1800,
  sessionSeconds: 21600,
  stages: ['รอเริ่มงาน', 'สำรวจและออกแบบ', 'เตรียมอุปกรณ์', 'กำลังติดตั้ง', 'ตรวจสอบและส่งมอบ', 'เสร็จสิ้น'],
  profitDistributionStatuses: ['pending', 'distributed']
});

let SPREADSHEET_INSTANCE_ = null;

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('ระบบบัญชีโปรเจกต์และงบส่วนกลาง')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

/** เรียกฟังก์ชันนี้ครั้งแรกจากหน้า Editor เพื่อสร้างฐานข้อมูล */
function setup() {
  const spreadsheet = getSpreadsheet_();
  setupSheets_(spreadsheet, true);
  warmDataCache_(spreadsheet);
  return {
    spreadsheetId: spreadsheet.getId(),
    spreadsheetUrl: spreadsheet.getUrl(),
    message: 'ตั้งค่าระบบเรียบร้อยแล้ว'
  };
}

function warmDataCache_(spreadsheet) {
  const projects = readProjects_(spreadsheet), transactions = readTransactions_(spreadsheet), subprojects = readSubprojects_(spreadsheet), categories = readCategories_(spreadsheet);
  const centralTransactions = readCentralTransactions_(spreadsheet), centralCategories = readCentralCategories_(spreadsheet), documentTypes = readDocumentTypes_(spreadsheet);
  const snapshot = { projects: projects, transactions: transactions, subprojects: subprojects, categories: categories, centralTransactions: centralTransactions, centralCategories: centralCategories, documentTypes: documentTypes, summary: buildSummary_(projects, transactions), centralSummary: buildCentralSummary_(centralTransactions) };
  const serialized = JSON.stringify(snapshot);
  if (serialized.length < 95000) CacheService.getScriptCache().put(APP.cacheKey, serialized, APP.cacheSeconds);
  return snapshot;
}

function getInitialData(authToken) {
  const authUser = requireSession_(authToken);
  const spreadsheet = getSpreadsheet_();
  setupSheets_(spreadsheet);
  return buildInitialData_(authUser, spreadsheet);
}

function buildInitialData_(authUser, spreadsheet) {
  const cache = CacheService.getScriptCache();
  let snapshot;
  const cached = cache.get(APP.cacheKey);
  if (cached) {
    try { snapshot = JSON.parse(cached); } catch (error) { snapshot = null; }
  }
  if (!snapshot) {
    snapshot = warmDataCache_(spreadsheet);
  }
  const currentUser = { email: authUser.username, name: authUser.name, role: authUser.role, status: authUser.status };
  const canViewAuditIdentity = authUser.role === 'admin';
  const userDirectory = {};
  if (canViewAuditIdentity) getAuthUserRecords_(spreadsheet).forEach(function (user) { userDirectory[user.username] = user.name; });
  return {
    appVersion: APP.version,
    projects: rowsForAuditRole_(snapshot.projects, authUser.role),
    transactions: rowsForAuditRole_(snapshot.transactions, authUser.role),
    subprojects: rowsForAuditRole_(snapshot.subprojects, authUser.role),
    categories: rowsForAuditRole_(snapshot.categories || [], authUser.role),
    centralTransactions: rowsForAuditRole_(snapshot.centralTransactions || [], authUser.role),
    centralCategories: rowsForAuditRole_(snapshot.centralCategories || [], authUser.role),
    documentTypes: rowsForAuditRole_(snapshot.documentTypes || [], authUser.role),
    summary: snapshot.summary,
    centralSummary: snapshot.centralSummary || buildCentralSummary_([]),
    spreadsheetUrl: spreadsheet.getUrl(),
    appUrl: ScriptApp.getService().getUrl() || '',
    currentUser: currentUser,
    userDirectory: userDirectory,
    users: [],
    auditLogs: [],
    adminDataLoaded: false
  };
}

/** ข้อมูลผู้บันทึกและผู้แก้ไขเปิดเผยเฉพาะ Admin เท่านั้น */
function rowsForAuditRole_(rows, role) {
  if (role === 'admin') return rows || [];
  return (rows || []).map(function (row) {
    const copy = {};
    Object.keys(row || {}).forEach(function (key) {
      if (key !== 'createdBy' && key !== 'updatedBy') copy[key] = row[key];
    });
    return copy;
  });
}

function getAdminData(authToken) {
  const authUser = requireSession_(authToken);
  if (authUser.role !== 'admin') throw new Error('เฉพาะผู้ดูแลระบบเท่านั้น');
  const spreadsheet = getSpreadsheet_();
  return { users: readAppUsers_(spreadsheet), auditLogs: readAuditLogs_(spreadsheet, 100) };
}

function getProjectActivity(projectId, authToken) {
  const authUser = requireSession_(authToken);
  if (authUser.role !== 'admin') throw new Error('เฉพาะผู้ดูแลระบบเท่านั้น');
  projectId = cleanText_(projectId, 80);
  const spreadsheet = getSpreadsheet_();
  const userNames = {};
  getAuthUserRecords_(spreadsheet).forEach(function (user) { userNames[user.username] = user.name; });
  const subprojectIds = readSubprojects_(spreadsheet).filter(function (item) { return item.projectId === projectId; }).map(function (item) { return item.id; });
  const rows = rowsWithoutHeader_(spreadsheet.getSheetByName(APP.auditSheet)).slice(-250).reverse();
  return rows.filter(function (row) {
    if (String(row[4]) === 'project' && String(row[5]) === projectId) return true;
    if (String(row[4]) === 'subproject' && subprojectIds.indexOf(String(row[5])) !== -1) return true;
    return [row[6], row[7]].some(function (value) {
      const text = String(value || '');
      return text.indexOf(projectId) !== -1 || subprojectIds.some(function (id) { return text.indexOf(id) !== -1; });
    });
  }).slice(0, 30).map(function (row) {
    const username = String(row[2] || '-');
    return { timestamp: serializeDate_(row[1]), user: userNames[username] || username, username: username, action: String(row[3] || ''), entity: String(row[4] || '') };
  });
}

/** ซ่อมฐานข้อมูลและโหลดใหม่ ใช้เมื่อ Deployment เก่าทิ้งโครงสร้างไม่ครบ */
function repairAndGetInitialData(authToken) {
  requireSession_(authToken);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const spreadsheet = getSpreadsheet_();
    setupSheets_(spreadsheet, true);
    clearDataCache_();
  } finally {
    lock.releaseLock();
  }
  return getInitialData(authToken);
}

function createProject(payload, authToken) {
  return withLock_(function () {
    const authUser = requireWriteAccess_(authToken);
    const name = cleanText_(payload && payload.name, 120);
    const customer = cleanText_(payload && payload.customer, 120);
    const budget = toAmount_(payload && payload.budget, true);
    const shareholders = parseShareholders_(payload && payload.shareholdersJson);
    const profitDistributionStatus = normalizeProfitDistributionStatus_(payload && payload.profitDistributionStatus);
    if (!name) throw new Error('กรุณาระบุชื่อโปรเจกต์');

    const sheet = getSpreadsheet_().getSheetByName(APP.projectsSheet);
    const id = Utilities.getUuid();
    const now = new Date();
    sheet.appendRow([id, name, customer, budget, 'กำลังดำเนินการ', now, now, authUser.username, authUser.username, 'รอเริ่มงาน', 0, JSON.stringify(shareholders), profitDistributionStatus]);
    writeAudit_('CREATE_PROJECT', 'project', id, '', JSON.stringify({ name: name, customer: customer, budget: budget, stage: 'รอเริ่มงาน', progress: 0, shareholders: shareholders, profitDistributionStatus: profitDistributionStatus }), authUser.username);
    clearDataCache_();
    return { success: true, project: { id: id, name: name, customer: customer, budget: budget, status: 'กำลังดำเนินการ', createdAt: serializeDate_(now), updatedAt: serializeDate_(now), createdBy: authUser.username, updatedBy: authUser.username, stage: 'รอเริ่มงาน', progress: 0, shareholders: shareholders, profitDistributionStatus: profitDistributionStatus }, message: 'เพิ่มโปรเจกต์เรียบร้อยแล้ว' };
  });
}

function updateProject(payload, authToken) {
  return withLock_(function () {
    const authUser = requireWriteAccess_(authToken);
    const id = cleanText_(payload && payload.id, 80);
    const name = cleanText_(payload && payload.name, 120);
    const customer = cleanText_(payload && payload.customer, 120);
    const budget = toAmount_(payload && payload.budget, true);
    const status = cleanText_(payload && payload.status, 30);
    const stage = cleanText_(payload && payload.stage, 40);
    const progress = toProgress_(payload && payload.progress);
    const hasShareholders = payload && Object.prototype.hasOwnProperty.call(payload, 'shareholdersJson');
    const hasDistributionStatus = payload && Object.prototype.hasOwnProperty.call(payload, 'profitDistributionStatus');
    const submittedShareholders = hasShareholders ? parseShareholders_(payload.shareholdersJson) : null;
    const submittedDistributionStatus = hasDistributionStatus ? normalizeProfitDistributionStatus_(payload.profitDistributionStatus) : '';
    const allowedStatuses = ['กำลังดำเนินการ', 'เสร็จสิ้น', 'พักงาน'];
    if (!id) throw new Error('ไม่พบรหัสโปรเจกต์');
    if (!name) throw new Error('กรุณาระบุชื่อโปรเจกต์');
    if (allowedStatuses.indexOf(status) === -1) throw new Error('สถานะไม่ถูกต้อง');
    if (APP.stages.indexOf(stage) === -1) throw new Error('ขั้นตอนงานไม่ถูกต้อง');

    const sheet = getSpreadsheet_().getSheetByName(APP.projectsSheet);
    const rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]) === id) {
        const row = index + 2;
        const previousShareholders = readStoredShareholders_(rows[index][11]);
        const previousDistributionStatus = readProfitDistributionStatus_(rows[index][12]);
        const shareholders = submittedShareholders === null ? previousShareholders : submittedShareholders;
        const profitDistributionStatus = submittedDistributionStatus || previousDistributionStatus;
        const before = { name: rows[index][1], customer: rows[index][2], budget: rows[index][3], status: rows[index][4], stage: rows[index][9] || 'รอเริ่มงาน', progress: Number(rows[index][10] || 0), shareholders: previousShareholders, profitDistributionStatus: previousDistributionStatus }, now = new Date();
        sheet.getRange(row, 2, 1, 12).setValues([[name, customer, budget, status, rows[index][5], now, rows[index][7] || authUser.username, authUser.username, stage, progress, JSON.stringify(shareholders), profitDistributionStatus]]);
        writeAudit_('UPDATE_PROJECT', 'project', id, JSON.stringify(before), JSON.stringify({ name: name, customer: customer, budget: budget, status: status, stage: stage, progress: progress, shareholders: shareholders, profitDistributionStatus: profitDistributionStatus }), authUser.username);
        clearDataCache_();
        return { success: true, project: { id: id, name: name, customer: customer, budget: budget, status: status, createdAt: serializeDate_(rows[index][5]), updatedAt: serializeDate_(now), createdBy: String(rows[index][7] || authUser.username), updatedBy: authUser.username, stage: stage, progress: progress, shareholders: shareholders, profitDistributionStatus: profitDistributionStatus }, message: 'แก้ไขโปรเจกต์เรียบร้อยแล้ว' };
      }
    }
    throw new Error('ไม่พบโปรเจกต์');
  });
}

function deleteProject(projectId, authToken) {
  return withLock_(function () {
    const user = requireAdminAccess_(authToken);
    projectId = cleanText_(projectId, 80);
    const spreadsheet = getSpreadsheet_();
    if (readTransactions_(spreadsheet).some(function (item) { return item.projectId === projectId; })) throw new Error('ไม่สามารถลบได้ เนื่องจากมีรายการรับ–จ่ายในโปรเจกต์นี้');
    if (readSubprojects_(spreadsheet).some(function (item) { return item.projectId === projectId; })) throw new Error('กรุณาลบโปรเจกต์ย่อยก่อนลบโปรเจกต์หลัก');
    const sheet = spreadsheet.getSheetByName(APP.projectsSheet), rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]) !== projectId) continue;
      const before = JSON.stringify({ name: rows[index][1], customer: rows[index][2], budget: rows[index][3], createdBy: rows[index][7], updatedBy: rows[index][8], shareholders: readStoredShareholders_(rows[index][11]), profitDistributionStatus: readProfitDistributionStatus_(rows[index][12]) });
      sheet.deleteRow(index + 2);
      writeAudit_('DELETE_PROJECT', 'project', projectId, before, '', user.username);
      clearDataCache_();
      return { success: true, id: projectId, message: 'ลบโปรเจกต์เรียบร้อยแล้ว' };
    }
    throw new Error('ไม่พบโปรเจกต์');
  });
}

function addTransaction(payload, authToken) {
  return withLock_(function () {
    const authUser = requireWriteAccess_(authToken);
    const projectId = cleanText_(payload && payload.projectId, 80);
    const subprojectId = cleanText_(payload && payload.subprojectId, 80);
    const type = cleanText_(payload && payload.type, 20);
    const category = cleanText_(payload && payload.category, 80);
    const description = cleanText_(payload && payload.description, 180);
    const amount = toAmount_(payload && payload.amount, false);
    const date = parseDate_(payload && payload.date);

    if (!projectId) throw new Error('กรุณาเลือกโปรเจกต์');
    if (['income', 'expense'].indexOf(type) === -1) throw new Error('ประเภทรายการไม่ถูกต้อง');
    if (!category) throw new Error('กรุณาระบุหมวดหมู่');
    const spreadsheet = getSpreadsheet_();
    if (!projectExists_(projectId, spreadsheet)) throw new Error('ไม่พบโปรเจกต์ที่เลือก');
    if (subprojectId && !subprojectBelongsToProject_(subprojectId, projectId, spreadsheet)) throw new Error('โปรเจกต์ย่อยไม่ตรงกับโปรเจกต์หลัก');
    if (!categoryAllowed_(category, type, spreadsheet)) throw new Error('หมวดหมู่นี้ไม่พร้อมใช้งานกับประเภทรายการที่เลือก');
    const documentType = normalizeDocumentType_(payload && payload.documentType, spreadsheet);

    const id = Utilities.getUuid();
    const now = new Date();
    spreadsheet.getSheetByName(APP.transactionsSheet)
      .appendRow([id, projectId, date, type, category, description, amount, now, subprojectId, authUser.username, authUser.username, documentType]);
    writeAudit_('CREATE_TRANSACTION', 'transaction', id, '', JSON.stringify({ projectId: projectId, subprojectId: subprojectId, type: type, category: category, amount: amount, documentType: documentType }), authUser.username);
    clearDataCache_();
    return { success: true, transaction: { id: id, projectId: projectId, subprojectId: subprojectId, date: serializeDateOnly_(date), type: type, category: category, description: description, amount: amount, documentType: documentType, createdAt: serializeDate_(now), createdBy: authUser.username, updatedBy: authUser.username }, message: 'อัปเดตรายการเรียบร้อยแล้ว' };
  });
}

function addTransactionsBatch(payloadJson, authToken) {
  return withLock_(function () {
    const authUser = requireWriteAccess_(authToken);
    let items = payloadJson;
    if (typeof items === 'string') {
      try { items = JSON.parse(items); }
      catch (error) { throw new Error('ข้อมูลรายการหลายรายการไม่ถูกต้อง'); }
    }
    if (!Array.isArray(items) || !items.length) throw new Error('กรุณาเพิ่มอย่างน้อย 1 รายการ');
    if (items.length > 50) throw new Error('บันทึกพร้อมกันได้ไม่เกิน 50 รายการ');

    const spreadsheet = getSpreadsheet_();
    const projectIds = {};
    readProjects_(spreadsheet).forEach(function (project) { projectIds[project.id] = true; });
    const subprojects = {};
    readSubprojects_(spreadsheet).forEach(function (subproject) { subprojects[subproject.id] = subproject.projectId; });
    const categories = readCategories_(spreadsheet);
    const documentTypes = {};
    readDocumentTypes_(spreadsheet).filter(function (item) { return item.status === 'active'; }).forEach(function (item) { documentTypes[item.name] = true; });
    const now = new Date(), rows = [], transactions = [], auditItems = [];

    items.forEach(function (payload, index) {
      try {
        const projectId = cleanText_(payload && payload.projectId, 80);
        const subprojectId = cleanText_(payload && payload.subprojectId, 80);
        const type = cleanText_(payload && payload.type, 20);
        const category = cleanText_(payload && payload.category, 80);
        const description = cleanText_(payload && payload.description, 180);
        const amount = toAmount_(payload && payload.amount, false);
        const date = parseDate_(payload && payload.date);
        const documentType = cleanText_(payload && payload.documentType, 80) || 'ไม่มีเอกสาร';
        if (!projectId || !projectIds[projectId]) throw new Error('ไม่พบโปรเจกต์ที่เลือก');
        if (['income', 'expense'].indexOf(type) === -1) throw new Error('ประเภทรายการไม่ถูกต้อง');
        if (!category || !categories.some(function (item) { return item.name === category && item.status === 'active' && (item.type === 'both' || item.type === type); })) throw new Error('หมวดหมู่ไม่พร้อมใช้งานกับประเภทรายการ');
        if (subprojectId && subprojects[subprojectId] !== projectId) throw new Error('โปรเจกต์ย่อยไม่ตรงกับโปรเจกต์หลัก');
        if (!documentTypes[documentType]) throw new Error('ประเภทเอกสารไม่พร้อมใช้งาน');
        const id = Utilities.getUuid();
        rows.push([id, projectId, date, type, category, description, amount, now, subprojectId, authUser.username, authUser.username, documentType]);
        transactions.push({ id: id, projectId: projectId, subprojectId: subprojectId, date: serializeDateOnly_(date), type: type, category: category, description: description, amount: amount, documentType: documentType, createdAt: serializeDate_(now), createdBy: authUser.username, updatedBy: authUser.username });
        auditItems.push({ id: id, projectId: projectId, type: type, category: category, amount: amount, documentType: documentType });
      } catch (error) {
        throw new Error('รายการที่ ' + (index + 1) + ': ' + error.message);
      }
    });

    const sheet = spreadsheet.getSheetByName(APP.transactionsSheet);
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 12).setValues(rows);
    const batchId = Utilities.getUuid();
    writeAudit_('BATCH_CREATE_TRANSACTIONS', 'transaction_batch', batchId, '', JSON.stringify({ count: transactions.length, items: auditItems }), authUser.username);
    clearDataCache_();
    return { success: true, transactions: transactions, count: transactions.length, message: 'บันทึก ' + transactions.length + ' รายการเรียบร้อยแล้ว' };
  });
}

function updateTransaction(payload, authToken) {
  return withLock_(function () {
    const authUser = requireWriteAccess_(authToken);
    const id = cleanText_(payload && payload.id, 80);
    const projectId = cleanText_(payload && payload.projectId, 80);
    const subprojectId = cleanText_(payload && payload.subprojectId, 80);
    const type = cleanText_(payload && payload.type, 20);
    const category = cleanText_(payload && payload.category, 80);
    const description = cleanText_(payload && payload.description, 180);
    const amount = toAmount_(payload && payload.amount, false);
    const date = parseDate_(payload && payload.date);
    if (!id) throw new Error('ไม่พบรายการที่ต้องการแก้ไข');
    if (['income', 'expense'].indexOf(type) === -1) throw new Error('ประเภทรายการไม่ถูกต้อง');
    if (!category) throw new Error('กรุณาระบุหมวดหมู่');
    const spreadsheet = getSpreadsheet_();
    if (!projectExists_(projectId, spreadsheet)) throw new Error('ไม่พบโปรเจกต์ที่เลือก');
    if (subprojectId && !subprojectBelongsToProject_(subprojectId, projectId, spreadsheet)) throw new Error('โปรเจกต์ย่อยไม่ตรงกับโปรเจกต์หลัก');
    if (!categoryAllowed_(category, type, spreadsheet)) throw new Error('หมวดหมู่นี้ไม่พร้อมใช้งานกับประเภทรายการที่เลือก');
    const documentType = normalizeDocumentType_(payload && payload.documentType, spreadsheet);
    const sheet = spreadsheet.getSheetByName(APP.transactionsSheet);
    const rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]) !== id) continue;
      const before = JSON.stringify({ projectId: rows[index][1], date: serializeDateOnly_(rows[index][2]), type: rows[index][3], category: rows[index][4], description: rows[index][5], amount: rows[index][6], subprojectId: rows[index][8] || '', documentType: rows[index][11] || 'ไม่มีเอกสาร' });
      sheet.getRange(index + 2, 2, 1, 11).setValues([[projectId, date, type, category, description, amount, rows[index][7], subprojectId, rows[index][9] || authUser.username, authUser.username, documentType]]);
      const after = { id: id, projectId: projectId, subprojectId: subprojectId, date: serializeDateOnly_(date), type: type, category: category, description: description, amount: amount, documentType: documentType, createdAt: serializeDate_(rows[index][7]), createdBy: String(rows[index][9] || authUser.username), updatedBy: authUser.username };
      writeAudit_('UPDATE_TRANSACTION', 'transaction', id, before, JSON.stringify(after), authUser.username);
      clearDataCache_();
      return { success: true, transaction: after, message: 'แก้ไขรายการเรียบร้อยแล้ว' };
    }
    throw new Error('ไม่พบรายการที่ต้องการแก้ไข');
  });
}

function updateProjectStatus(projectId, status, authToken) {
  return withLock_(function () {
    const user = requireWriteAccess_(authToken);
    const allowed = ['กำลังดำเนินการ', 'เสร็จสิ้น', 'พักงาน'];
    if (allowed.indexOf(status) === -1) throw new Error('สถานะไม่ถูกต้อง');
    const sheet = getSpreadsheet_().getSheetByName(APP.projectsSheet);
    const values = sheet.getDataRange().getValues();
    for (let row = 1; row < values.length; row++) {
      if (String(values[row][0]) === String(projectId)) {
        sheet.getRange(row + 1, 5, 1, 5).setValues([[status, values[row][5], new Date(), values[row][7] || user.username, user.username]]);
        writeAudit_('UPDATE_PROJECT', 'project', projectId, JSON.stringify({ status: values[row][4] }), JSON.stringify({ status: status }), user.username);
        clearDataCache_();
        return { success: true, message: 'อัปเดตสถานะเรียบร้อยแล้ว' };
      }
    }
    throw new Error('ไม่พบโปรเจกต์');
  });
}

function deleteTransaction(transactionId, authToken) {
  return withLock_(function () {
    const authUser = requireAdminAccess_(authToken);
    const sheet = getSpreadsheet_().getSheetByName(APP.transactionsSheet);
    const values = sheet.getDataRange().getValues();
    for (let row = 1; row < values.length; row++) {
      if (String(values[row][0]) === String(transactionId)) {
        const projectId = values[row][1];
        const before = JSON.stringify({ projectId: projectId, date: serializeDateOnly_(values[row][2]), type: values[row][3], category: values[row][4], description: values[row][5], amount: values[row][6], subprojectId: values[row][8] || '', documentType: values[row][11] || 'ไม่มีเอกสาร' });
        sheet.deleteRow(row + 1);
        writeAudit_('DELETE_TRANSACTION', 'transaction', transactionId, before, '', authUser.username);
        touchProject_(projectId);
        clearDataCache_();
        return { success: true, id: String(transactionId), message: 'ลบรายการเรียบร้อยแล้ว' };
      }
    }
    throw new Error('ไม่พบรายการที่ต้องการลบ');
  });
}

function createCentralTransaction(payload, authToken) {
  return withLock_(function () {
    const authUser = requireWriteAccess_(authToken);
    const type = cleanText_(payload && payload.type, 20);
    const category = cleanText_(payload && payload.category, 80);
    const description = cleanText_(payload && payload.description, 180);
    const amount = toAmount_(payload && payload.amount, false);
    const date = parseDate_(payload && payload.date);
    if (['income', 'expense'].indexOf(type) === -1) throw new Error('ประเภทรายการงบส่วนกลางไม่ถูกต้อง');
    if (!category) throw new Error('กรุณาเลือกหมวดหมู่งบส่วนกลาง');
    const spreadsheet = getSpreadsheet_();
    if (!centralCategoryAllowed_(category, type, spreadsheet)) throw new Error('หมวดหมู่นี้ไม่พร้อมใช้งานกับประเภทรายการที่เลือก');
    const documentType = normalizeDocumentType_(payload && payload.documentType, spreadsheet);
    const id = Utilities.getUuid(), now = new Date();
    spreadsheet.getSheetByName(APP.centralTransactionsSheet)
      .appendRow([id, date, type, category, description, amount, now, authUser.username, authUser.username, documentType]);
    const transaction = { id: id, date: serializeDateOnly_(date), type: type, category: category, description: description, amount: amount, documentType: documentType, createdAt: serializeDate_(now), createdBy: authUser.username, updatedBy: authUser.username };
    writeAudit_('CREATE_CENTRAL_TRANSACTION', 'central_transaction', id, '', JSON.stringify(transaction), authUser.username);
    clearDataCache_();
    return { success: true, transaction: transaction, message: 'บันทึกรายการงบส่วนกลางเรียบร้อยแล้ว' };
  });
}

function updateCentralTransaction(payload, authToken) {
  return withLock_(function () {
    const authUser = requireWriteAccess_(authToken);
    const id = cleanText_(payload && payload.id, 80);
    const type = cleanText_(payload && payload.type, 20);
    const category = cleanText_(payload && payload.category, 80);
    const description = cleanText_(payload && payload.description, 180);
    const amount = toAmount_(payload && payload.amount, false);
    const date = parseDate_(payload && payload.date);
    if (!id) throw new Error('ไม่พบรายการงบส่วนกลางที่ต้องการแก้ไข');
    if (['income', 'expense'].indexOf(type) === -1) throw new Error('ประเภทรายการงบส่วนกลางไม่ถูกต้อง');
    if (!category) throw new Error('กรุณาเลือกหมวดหมู่งบส่วนกลาง');
    const spreadsheet = getSpreadsheet_();
    if (!centralCategoryAllowed_(category, type, spreadsheet)) throw new Error('หมวดหมู่นี้ไม่พร้อมใช้งานกับประเภทรายการที่เลือก');
    const documentType = normalizeDocumentType_(payload && payload.documentType, spreadsheet);
    const sheet = spreadsheet.getSheetByName(APP.centralTransactionsSheet), rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]) !== id) continue;
      const before = { id: id, date: serializeDateOnly_(rows[index][1]), type: String(rows[index][2]), category: String(rows[index][3]), description: String(rows[index][4] || ''), amount: Number(rows[index][5] || 0), documentType: String(rows[index][9] || 'ไม่มีเอกสาร'), createdAt: serializeDate_(rows[index][6]), createdBy: String(rows[index][7] || ''), updatedBy: String(rows[index][8] || '') };
      sheet.getRange(index + 2, 2, 1, 9).setValues([[date, type, category, description, amount, rows[index][6], rows[index][7] || authUser.username, authUser.username, documentType]]);
      const transaction = { id: id, date: serializeDateOnly_(date), type: type, category: category, description: description, amount: amount, documentType: documentType, createdAt: serializeDate_(rows[index][6]), createdBy: String(rows[index][7] || authUser.username), updatedBy: authUser.username };
      writeAudit_('UPDATE_CENTRAL_TRANSACTION', 'central_transaction', id, JSON.stringify(before), JSON.stringify(transaction), authUser.username);
      clearDataCache_();
      return { success: true, transaction: transaction, message: 'แก้ไขรายการงบส่วนกลางเรียบร้อยแล้ว' };
    }
    throw new Error('ไม่พบรายการงบส่วนกลางที่ต้องการแก้ไข');
  });
}

function deleteCentralTransaction(id, authToken) {
  return withLock_(function () {
    const authUser = requireAdminAccess_(authToken);
    id = cleanText_(id, 80);
    const sheet = getSpreadsheet_().getSheetByName(APP.centralTransactionsSheet), rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]) !== id) continue;
      const before = { id: id, date: serializeDateOnly_(rows[index][1]), type: String(rows[index][2]), category: String(rows[index][3]), description: String(rows[index][4] || ''), amount: Number(rows[index][5] || 0), documentType: String(rows[index][9] || 'ไม่มีเอกสาร'), createdAt: serializeDate_(rows[index][6]), createdBy: String(rows[index][7] || ''), updatedBy: String(rows[index][8] || '') };
      sheet.deleteRow(index + 2);
      writeAudit_('DELETE_CENTRAL_TRANSACTION', 'central_transaction', id, JSON.stringify(before), '', authUser.username);
      clearDataCache_();
      return { success: true, id: id, message: 'ลบรายการงบส่วนกลางเรียบร้อยแล้ว' };
    }
    throw new Error('ไม่พบรายการงบส่วนกลางที่ต้องการลบ');
  });
}

function getSpreadsheet_() {
  if (SPREADSHEET_INSTANCE_) return SPREADSHEET_INSTANCE_;
  const properties = PropertiesService.getScriptProperties();
  const savedId = properties.getProperty(APP.spreadsheetProperty);
  if (savedId) {
    try {
      SPREADSHEET_INSTANCE_ = SpreadsheetApp.openById(savedId);
      return SPREADSHEET_INSTANCE_;
    } catch (error) { /* สร้างใหม่ด้านล่าง */ }
  }
  const active = SpreadsheetApp.getActiveSpreadsheet();
  const spreadsheet = active || SpreadsheetApp.create('ฐานข้อมูลระบบบัญชีรายโปรเจกต์');
  properties.setProperty(APP.spreadsheetProperty, spreadsheet.getId());
  SPREADSHEET_INSTANCE_ = spreadsheet;
  return SPREADSHEET_INSTANCE_;
}

function setupSheets_(spreadsheet, force) {
  spreadsheet = spreadsheet || getSpreadsheet_();
  const properties = PropertiesService.getScriptProperties();
  if (!force && properties.getProperty(APP.schemaProperty) === APP.schemaVersion) return;
  ensureSheet_(spreadsheet, APP.projectsSheet,
    ['Project ID', 'ชื่อโปรเจกต์', 'ลูกค้า', 'งบประมาณ', 'สถานะ', 'วันที่สร้าง', 'แก้ไขล่าสุด', 'สร้างโดย', 'แก้ไขล่าสุดโดย', 'ขั้นตอนปัจจุบัน', 'ความคืบหน้า (%)', 'ผู้ถือหุ้น (JSON)', 'สถานะแบ่งผลกำไร']);
  ensureSheet_(spreadsheet, APP.transactionsSheet,
    ['Transaction ID', 'Project ID', 'วันที่', 'ประเภท', 'หมวดหมู่', 'รายละเอียด', 'จำนวนเงิน', 'วันที่บันทึก', 'Subproject ID', 'สร้างโดย', 'แก้ไขล่าสุดโดย', 'ประเภทเอกสาร']);
  ensureSheet_(spreadsheet, APP.usersSheet,
    ['Email', 'ชื่อแสดงผล', 'Role', 'สถานะ', 'เข้าใช้ล่าสุด', 'วันที่สร้าง']);
  ensureSheet_(spreadsheet, APP.auditSheet,
    ['Log ID', 'วันเวลา', 'ผู้ใช้งาน', 'Action', 'ประเภทข้อมูล', 'Record ID', 'ข้อมูลเดิม', 'ข้อมูลใหม่']);
  ensureSheet_(spreadsheet, APP.appUsersSheet,
    ['Username', 'ชื่อแสดงผล', 'Password Hash', 'Salt', 'Role', 'สถานะ', 'เข้าใช้ล่าสุด', 'วันที่สร้าง', 'Auth Version']);
  ensureSheet_(spreadsheet, APP.subprojectsSheet,
    ['Subproject ID', 'Project ID', 'ชื่อโปรเจกต์ย่อย', 'รายละเอียด', 'งบประมาณ', 'สถานะ', 'วันที่สร้าง', 'แก้ไขล่าสุด', 'สร้างโดย', 'แก้ไขล่าสุดโดย', 'ขั้นตอนปัจจุบัน', 'ความคืบหน้า (%)']);
  ensureSheet_(spreadsheet, APP.categoriesSheet,
    ['Category ID', 'ชื่อหมวดหมู่', 'ประเภท', 'สถานะ', 'สร้างโดย', 'วันที่สร้าง', 'แก้ไขโดย', 'แก้ไขล่าสุด']);
  ensureSheet_(spreadsheet, APP.centralTransactionsSheet,
    ['Central Transaction ID', 'วันที่', 'ประเภท', 'หมวดหมู่', 'รายละเอียด', 'จำนวนเงิน', 'วันที่บันทึก', 'สร้างโดย', 'แก้ไขล่าสุดโดย', 'ประเภทเอกสาร']);
  ensureSheet_(spreadsheet, APP.centralCategoriesSheet,
    ['Central Category ID', 'ชื่อหมวดหมู่', 'ประเภท', 'สถานะ', 'ลำดับ']);
  ensureSheet_(spreadsheet, APP.documentTypesSheet,
    ['Document Type ID', 'ชื่อประเภทเอกสาร', 'สถานะ', 'สร้างโดย', 'วันที่สร้าง']);
  initializeDefaultAdmin_(spreadsheet);
  initializeDefaultCategories_(spreadsheet);
  initializeDefaultCentralCategories_(spreadsheet);
  initializeDefaultDocumentTypes_(spreadsheet);
  properties.setProperty(APP.schemaProperty, APP.schemaVersion);
}

function ensureSheet_(spreadsheet, name, headers) {
  let sheet = spreadsheet.getSheetByName(name);
  let needsFormatting = false;
  if (!sheet) {
    sheet = spreadsheet.insertSheet(name);
    needsFormatting = true;
  }
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
    needsFormatting = true;
  }
  // เติมคอลัมน์ใหม่เมื่ออัปเกรดระบบ โดยไม่กระทบข้อมูลแถวเดิม
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  if (needsFormatting) {
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length)
      .setBackground('#FACC15').setFontColor('#111111').setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }
  return sheet;
}

function readProjects_(spreadsheet) {
  spreadsheet = spreadsheet || getSpreadsheet_();
  const rows = rowsWithoutHeader_(spreadsheet.getSheetByName(APP.projectsSheet));
  return rows.filter(function (row) { return row[0]; }).map(function (row) {
    return {
      id: String(row[0]), name: String(row[1] || ''), customer: String(row[2] || ''),
      budget: Number(row[3] || 0), status: String(row[4] || 'กำลังดำเนินการ'),
      createdAt: serializeDate_(row[5]), updatedAt: serializeDate_(row[6]), createdBy: String(row[7] || ''), updatedBy: String(row[8] || ''), stage: String(row[9] || (row[4] === 'เสร็จสิ้น' ? 'เสร็จสิ้น' : 'รอเริ่มงาน')), progress: Number(row[10] == null || row[10] === '' ? (row[4] === 'เสร็จสิ้น' ? 100 : 0) : row[10]), shareholders: readStoredShareholders_(row[11]), profitDistributionStatus: readProfitDistributionStatus_(row[12])
    };
  });
}

function readTransactions_(spreadsheet) {
  spreadsheet = spreadsheet || getSpreadsheet_();
  const rows = rowsWithoutHeader_(spreadsheet.getSheetByName(APP.transactionsSheet));
  return rows.filter(function (row) { return row[0]; }).map(function (row) {
    return {
      id: String(row[0]), projectId: String(row[1]), date: serializeDateOnly_(row[2]),
      type: String(row[3]), category: String(row[4] || ''), description: String(row[5] || ''),
      amount: Number(row[6] || 0), createdAt: serializeDate_(row[7]), subprojectId: String(row[8] || ''), createdBy: String(row[9] || ''), updatedBy: String(row[10] || ''), documentType: String(row[11] || 'ไม่มีเอกสาร')
    };
  }).sort(function (a, b) { return b.date.localeCompare(a.date); });
}

function readCentralTransactions_(spreadsheet) {
  spreadsheet = spreadsheet || getSpreadsheet_();
  const rows = rowsWithoutHeader_(spreadsheet.getSheetByName(APP.centralTransactionsSheet));
  return rows.filter(function (row) { return row[0]; }).map(function (row) {
    return {
      id: String(row[0]), date: serializeDateOnly_(row[1]), type: String(row[2]),
      category: String(row[3] || ''), description: String(row[4] || ''), amount: Number(row[5] || 0),
      createdAt: serializeDate_(row[6]), createdBy: String(row[7] || ''), updatedBy: String(row[8] || ''), documentType: String(row[9] || 'ไม่มีเอกสาร')
    };
  }).sort(function (a, b) { return b.date.localeCompare(a.date) || String(b.createdAt).localeCompare(String(a.createdAt)); });
}

function readSubprojects_(spreadsheet) {
  const rows = rowsWithoutHeader_(spreadsheet.getSheetByName(APP.subprojectsSheet));
  return rows.filter(function (row) { return row[0]; }).map(function (row) {
    return { id: String(row[0]), projectId: String(row[1]), name: String(row[2] || ''), description: String(row[3] || ''), budget: Number(row[4] || 0), status: String(row[5] || 'กำลังดำเนินการ'), createdAt: serializeDate_(row[6]), updatedAt: serializeDate_(row[7]), createdBy: String(row[8] || ''), updatedBy: String(row[9] || ''), stage: String(row[10] || (row[5] === 'เสร็จสิ้น' ? 'เสร็จสิ้น' : 'รอเริ่มงาน')), progress: Number(row[11] == null || row[11] === '' ? (row[5] === 'เสร็จสิ้น' ? 100 : 0) : row[11]) };
  });
}

function createSubproject(payload, authToken) {
  return withLock_(function () {
    const authUser = requireWriteAccess_(authToken);
    const projectId = cleanText_(payload && payload.projectId, 80);
    const name = cleanText_(payload && payload.name, 120);
    const description = cleanText_(payload && payload.description, 200);
    const budget = toAmount_(payload && payload.budget, true);
    const status = cleanText_(payload && payload.status, 30) || 'กำลังดำเนินการ';
    const stage = cleanText_(payload && payload.stage, 40) || 'รอเริ่มงาน';
    const progress = toProgress_(payload && payload.progress);
    if (!projectId || !projectExists_(projectId, getSpreadsheet_())) throw new Error('ไม่พบโปรเจกต์หลัก');
    if (!name) throw new Error('กรุณาระบุชื่อโปรเจกต์ย่อย');
    if (['กำลังดำเนินการ', 'เสร็จสิ้น', 'พักงาน'].indexOf(status) === -1) throw new Error('สถานะไม่ถูกต้อง');
    if (APP.stages.indexOf(stage) === -1) throw new Error('ขั้นตอนงานไม่ถูกต้อง');
    const id = Utilities.getUuid();
    const now = new Date();
    getSpreadsheet_().getSheetByName(APP.subprojectsSheet).appendRow([id, projectId, name, description, budget, status, now, now, authUser.username, authUser.username, stage, progress]);
    writeAudit_('CREATE_SUBPROJECT', 'subproject', id, '', JSON.stringify({ projectId: projectId, name: name, budget: budget, status: status, stage: stage, progress: progress }), authUser.username);
    clearDataCache_();
    return { success: true, subproject: { id: id, projectId: projectId, name: name, description: description, budget: budget, status: status, createdAt: serializeDate_(now), updatedAt: serializeDate_(now), createdBy: authUser.username, updatedBy: authUser.username, stage: stage, progress: progress }, message: 'เพิ่มโปรเจกต์ย่อยเรียบร้อยแล้ว' };
  });
}

function updateSubproject(payload, authToken) {
  return withLock_(function () {
    const authUser = requireWriteAccess_(authToken);
    const id = cleanText_(payload && payload.id, 80);
    const name = cleanText_(payload && payload.name, 120);
    const description = cleanText_(payload && payload.description, 200);
    const budget = toAmount_(payload && payload.budget, true);
    const status = cleanText_(payload && payload.status, 30);
    const stage = cleanText_(payload && payload.stage, 40);
    const progress = toProgress_(payload && payload.progress);
    if (!name) throw new Error('กรุณาระบุชื่อโปรเจกต์ย่อย');
    if (['กำลังดำเนินการ', 'เสร็จสิ้น', 'พักงาน'].indexOf(status) === -1) throw new Error('สถานะไม่ถูกต้อง');
    if (APP.stages.indexOf(stage) === -1) throw new Error('ขั้นตอนงานไม่ถูกต้อง');
    const sheet = getSpreadsheet_().getSheetByName(APP.subprojectsSheet);
    const rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]) !== id) continue;
      const before = JSON.stringify({ name: rows[index][2], description: rows[index][3], budget: rows[index][4], status: rows[index][5], stage: rows[index][10] || 'รอเริ่มงาน', progress: Number(rows[index][11] || 0) });
      const now = new Date();
      sheet.getRange(index + 2, 3, 1, 10).setValues([[name, description, budget, status, rows[index][6], now, rows[index][8] || authUser.username, authUser.username, stage, progress]]);
      writeAudit_('UPDATE_SUBPROJECT', 'subproject', id, before, JSON.stringify({ name: name, description: description, budget: budget, status: status, stage: stage, progress: progress }), authUser.username);
      clearDataCache_();
      return { success: true, subproject: { id: id, projectId: String(rows[index][1]), name: name, description: description, budget: budget, status: status, createdAt: serializeDate_(rows[index][6]), updatedAt: serializeDate_(now), createdBy: String(rows[index][8] || authUser.username), updatedBy: authUser.username, stage: stage, progress: progress }, message: 'แก้ไขโปรเจกต์ย่อยเรียบร้อยแล้ว' };
    }
    throw new Error('ไม่พบโปรเจกต์ย่อย');
  });
}

function deleteSubproject(id, authToken) {
  return withLock_(function () {
    const authUser = requireAdminAccess_(authToken);
    id = cleanText_(id, 80);
    const transactions = readTransactions_(getSpreadsheet_());
    if (transactions.some(function (item) { return item.subprojectId === id; })) throw new Error('ไม่สามารถลบได้ เนื่องจากมีรายการรับ–จ่ายผูกกับโปรเจกต์ย่อยนี้');
    const sheet = getSpreadsheet_().getSheetByName(APP.subprojectsSheet);
    const rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]) !== id) continue;
      const before = JSON.stringify({ projectId: rows[index][1], name: rows[index][2], budget: rows[index][4] });
      sheet.deleteRow(index + 2);
      writeAudit_('DELETE_SUBPROJECT', 'subproject', id, before, '', authUser.username);
      clearDataCache_();
      return { success: true, id: id, message: 'ลบโปรเจกต์ย่อยเรียบร้อยแล้ว' };
    }
    throw new Error('ไม่พบโปรเจกต์ย่อย');
  });
}

function subprojectBelongsToProject_(subprojectId, projectId, spreadsheet) {
  return readSubprojects_(spreadsheet).some(function (item) { return item.id === subprojectId && item.projectId === projectId; });
}

function initializeDefaultCategories_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(APP.categoriesSheet);
  if (sheet.getLastRow() > 1) return;
  const now = new Date();
  ['ค่ามัดจำ', 'ค่างวดงาน', 'ค่าวัสดุ', 'ค่าแรง', 'ค่าเดินทาง', 'ค่าใช้จ่ายอื่นๆ'].forEach(function (name) {
    sheet.appendRow([Utilities.getUuid(), name, 'both', 'active', 'system', now, 'system', now]);
  });
}

function initializeDefaultCentralCategories_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(APP.centralCategoriesSheet);
  const defaults = [
    ['เงินทุนส่วนกลาง', 'income'],
    ['รายรับส่วนกลาง', 'income'],
    ['ค่าที่พัก', 'expense'],
    ['ค่าน้ำมัน', 'expense'],
    ['ค่าอาหาร / ค่าข้าว', 'expense'],
    ['ค่าจ้างพนักงาน', 'expense'],
    ['ค่าอุปกรณ์สำนักงาน', 'expense'],
    ['ค่าเช่าสำนักงาน', 'expense'],
    ['ค่าน้ำ–ค่าไฟ', 'expense'],
    ['ค่าบริการและซอฟต์แวร์', 'expense'],
    ['ค่าใช้จ่ายอื่นๆ', 'expense'],
    ['อื่นๆ', 'both']
  ];
  const existingNames = sheet.getLastRow() > 1
    ? sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues().map(function (row) { return String(row[0] || '').trim().toLowerCase(); })
    : [];
  defaults.forEach(function (item, index) {
    if (existingNames.indexOf(item[0].toLowerCase()) === -1) {
      sheet.appendRow([Utilities.getUuid(), item[0], item[1], 'active', index + 1]);
    }
  });
}

function readCentralCategories_(spreadsheet) {
  spreadsheet = spreadsheet || getSpreadsheet_();
  return rowsWithoutHeader_(spreadsheet.getSheetByName(APP.centralCategoriesSheet)).filter(function (row) { return row[0]; }).map(function (row) {
    return { id: String(row[0]), name: String(row[1] || ''), type: String(row[2] || 'both'), status: String(row[3] || 'active'), order: Number(row[4] || 0) };
  }).sort(function (a, b) { return a.order - b.order || a.name.localeCompare(b.name); });
}

function initializeDefaultDocumentTypes_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(APP.documentTypesSheet);
  const defaults = ['บิลเงินสด', 'ใบกำกับภาษี', 'เงินโอน', 'ใบเสร็จรับเงิน', 'ไม่มีเอกสาร'];
  const existingNames = sheet.getLastRow() > 1
    ? sheet.getRange(2, 2, sheet.getLastRow() - 1, 1).getDisplayValues().map(function (row) { return String(row[0] || '').trim().toLowerCase(); })
    : [];
  const now = new Date();
  const missingRows = defaults.filter(function (name) { return existingNames.indexOf(name.toLowerCase()) === -1; })
    .map(function (name) { return [Utilities.getUuid(), name, 'active', 'system', now]; });
  if (missingRows.length) sheet.getRange(sheet.getLastRow() + 1, 1, missingRows.length, 5).setValues(missingRows);
}

function readDocumentTypes_(spreadsheet) {
  spreadsheet = spreadsheet || getSpreadsheet_();
  return rowsWithoutHeader_(spreadsheet.getSheetByName(APP.documentTypesSheet)).filter(function (row) { return row[0]; }).map(function (row) {
    return { id: String(row[0]), name: String(row[1] || ''), status: String(row[2] || 'active'), createdBy: String(row[3] || ''), createdAt: serializeDate_(row[4]) };
  });
}

function createDocumentType(payload, authToken) {
  return withLock_(function () {
    const user = requireWriteAccess_(authToken);
    const name = cleanText_(payload && payload.name, 80);
    if (!name) throw new Error('กรุณาระบุชื่อประเภทเอกสาร');
    const spreadsheet = getSpreadsheet_(), items = readDocumentTypes_(spreadsheet);
    const existing = items.find(function (item) { return item.name.toLowerCase() === name.toLowerCase(); });
    if (existing) {
      if (existing.status !== 'active') throw new Error('ประเภทเอกสารนี้ถูกปิดใช้งาน');
      return { success: true, documentType: existing, message: 'มีประเภทเอกสารนี้อยู่แล้ว' };
    }
    const id = Utilities.getUuid(), now = new Date();
    spreadsheet.getSheetByName(APP.documentTypesSheet).appendRow([id, name, 'active', user.username, now]);
    const documentType = { id: id, name: name, status: 'active', createdBy: user.username, createdAt: serializeDate_(now) };
    writeAudit_('CREATE_DOCUMENT_TYPE', 'document_type', id, '', JSON.stringify(documentType), user.username);
    clearDataCache_();
    return { success: true, documentType: documentType, message: 'เพิ่มประเภทเอกสารเรียบร้อยแล้ว' };
  });
}

function normalizeDocumentType_(value, spreadsheet) {
  const name = cleanText_(value, 80) || 'ไม่มีเอกสาร';
  let items;
  const cached = CacheService.getScriptCache().get(APP.cacheKey);
  if (cached) { try { items = JSON.parse(cached).documentTypes; } catch (error) { items = null; } }
  items = items || readDocumentTypes_(spreadsheet || getSpreadsheet_());
  if (!items.some(function (item) { return item.name === name && item.status === 'active'; })) throw new Error('ประเภทเอกสารนี้ไม่พร้อมใช้งาน');
  return name;
}

function readCategories_(spreadsheet) {
  return rowsWithoutHeader_(spreadsheet.getSheetByName(APP.categoriesSheet)).filter(function (row) { return row[0]; }).map(function (row) {
    return { id: String(row[0]), name: String(row[1] || ''), type: String(row[2] || 'both'), status: String(row[3] || 'active'), createdBy: String(row[4] || ''), createdAt: serializeDate_(row[5]), updatedBy: String(row[6] || ''), updatedAt: serializeDate_(row[7]) };
  });
}

function createCategory(payload, authToken) {
  return withLock_(function () {
    const user = requireWriteAccess_(authToken);
    const name = cleanText_(payload && payload.name, 80);
    const type = cleanText_(payload && payload.type, 20);
    if (!name) throw new Error('กรุณาระบุชื่อหมวดหมู่');
    if (['both', 'income', 'expense'].indexOf(type) === -1) throw new Error('ประเภทหมวดหมู่ไม่ถูกต้อง');
    const spreadsheet = getSpreadsheet_();
    const categories = readCategories_(spreadsheet);
    if (categories.some(function (item) { return item.name.toLowerCase() === name.toLowerCase(); })) throw new Error('มีหมวดหมู่นี้อยู่แล้ว');
    const id = Utilities.getUuid(), now = new Date();
    spreadsheet.getSheetByName(APP.categoriesSheet).appendRow([id, name, type, 'active', user.username, now, user.username, now]);
    writeAudit_('CREATE_CATEGORY', 'category', id, '', JSON.stringify({ name: name, type: type }), user.username);
    clearDataCache_();
    return { success: true, category: { id: id, name: name, type: type, status: 'active', createdBy: user.username, createdAt: serializeDate_(now), updatedBy: user.username, updatedAt: serializeDate_(now) }, message: 'เพิ่มหมวดหมู่เรียบร้อยแล้ว' };
  });
}

function updateCategory(payload, authToken) {
  return withLock_(function () {
    const user = requireWriteAccess_(authToken);
    const id = cleanText_(payload && payload.id, 80), name = cleanText_(payload && payload.name, 80), type = cleanText_(payload && payload.type, 20), status = cleanText_(payload && payload.status, 20);
    if (!name) throw new Error('กรุณาระบุชื่อหมวดหมู่');
    if (['both', 'income', 'expense'].indexOf(type) === -1 || ['active', 'inactive'].indexOf(status) === -1) throw new Error('ข้อมูลหมวดหมู่ไม่ถูกต้อง');
    const sheet = getSpreadsheet_().getSheetByName(APP.categoriesSheet), rows = rowsWithoutHeader_(sheet);
    if (rows.some(function (row) { return String(row[0]) !== id && String(row[1]).toLowerCase() === name.toLowerCase(); })) throw new Error('มีชื่อหมวดหมู่นี้อยู่แล้ว');
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]) !== id) continue;
      const before = JSON.stringify({ name: rows[index][1], type: rows[index][2], status: rows[index][3] }), now = new Date();
      sheet.getRange(index + 2, 2, 1, 7).setValues([[name, type, status, rows[index][4], rows[index][5], user.username, now]]);
      writeAudit_('UPDATE_CATEGORY', 'category', id, before, JSON.stringify({ name: name, type: type, status: status }), user.username);
      clearDataCache_();
      return { success: true, category: { id: id, name: name, type: type, status: status, createdBy: String(rows[index][4] || ''), createdAt: serializeDate_(rows[index][5]), updatedBy: user.username, updatedAt: serializeDate_(now) }, message: 'แก้ไขหมวดหมู่เรียบร้อยแล้ว' };
    }
    throw new Error('ไม่พบหมวดหมู่');
  });
}

function deleteCategory(id, authToken) {
  return withLock_(function () {
    const user = requireAdminAccess_(authToken);
    id = cleanText_(id, 80);
    const spreadsheet = getSpreadsheet_(), categories = readCategories_(spreadsheet), category = categories.find(function (item) { return item.id === id; });
    if (!category) throw new Error('ไม่พบหมวดหมู่');
    if (readTransactions_(spreadsheet).some(function (item) { return item.category === category.name; })) throw new Error('หมวดหมู่นี้ถูกใช้งานแล้ว ให้เปลี่ยนเป็นสถานะปิดใช้งานแทน');
    const sheet = spreadsheet.getSheetByName(APP.categoriesSheet), rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) if (String(rows[index][0]) === id) { sheet.deleteRow(index + 2); break; }
    writeAudit_('DELETE_CATEGORY', 'category', id, JSON.stringify(category), '', user.username);
    clearDataCache_();
    return { success: true, id: id, message: 'ลบหมวดหมู่เรียบร้อยแล้ว' };
  });
}

function categoryAllowed_(name, transactionType, spreadsheet) {
  let categories;
  const cached = CacheService.getScriptCache().get(APP.cacheKey);
  if (cached) { try { categories = JSON.parse(cached).categories; } catch (error) { categories = null; } }
  categories = categories || readCategories_(spreadsheet);
  return categories.some(function (item) { return item.name === name && item.status === 'active' && (item.type === 'both' || item.type === transactionType); });
}

function centralCategoryAllowed_(name, transactionType, spreadsheet) {
  let categories;
  const cached = CacheService.getScriptCache().get(APP.cacheKey);
  if (cached) { try { categories = JSON.parse(cached).centralCategories; } catch (error) { categories = null; } }
  categories = categories || readCentralCategories_(spreadsheet);
  return categories.some(function (item) { return item.name === name && item.status === 'active' && (item.type === 'both' || item.type === transactionType); });
}

function buildSummary_(projects, transactions) {
  const totals = transactions.reduce(function (sum, item) {
    if (item.type === 'income') sum.income += item.amount;
    if (item.type === 'expense') sum.expense += item.amount;
    return sum;
  }, { income: 0, expense: 0 });
  return {
    projectCount: projects.length,
    activeProjectCount: projects.filter(function (p) { return p.status === 'กำลังดำเนินการ'; }).length,
    income: totals.income,
    expense: totals.expense,
    profit: totals.income - totals.expense
  };
}

function buildCentralSummary_(transactions) {
  const totals = (transactions || []).reduce(function (sum, item) {
    if (item.type === 'income') sum.income += Number(item.amount || 0);
    if (item.type === 'expense') sum.expense += Number(item.amount || 0);
    return sum;
  }, { income: 0, expense: 0 });
  return { income: totals.income, expense: totals.expense, balance: totals.income - totals.expense, transactionCount: (transactions || []).length };
}

function rowsWithoutHeader_(sheet) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
}

function projectExists_(id, spreadsheet) {
  const sheet = (spreadsheet || getSpreadsheet_()).getSheetByName(APP.projectsSheet);
  if (sheet.getLastRow() < 2) return false;
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getDisplayValues()
    .some(function (row) { return row[0] === id; });
}

function touchProject_(id, spreadsheet) {
  const sheet = (spreadsheet || getSpreadsheet_()).getSheetByName(APP.projectsSheet);
  const ids = sheet.getLastRow() > 1 ? sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues() : [];
  for (let i = 0; i < ids.length; i++) {
    if (String(ids[i][0]) === String(id)) {
      sheet.getRange(i + 2, 7).setValue(new Date());
      return;
    }
  }
}

function clearDataCache_() {
  CacheService.getScriptCache().remove(APP.cacheKey);
}

function requireWriteAccess_(authToken) {
  const user = requireSession_(authToken);
  if (user.status !== 'active') throw new Error('บัญชีนี้ถูกระงับการใช้งาน');
  if (['admin', 'editor'].indexOf(user.role) === -1) throw new Error('บัญชีนี้มีสิทธิ์ดูข้อมูลเท่านั้น');
  return user;
}

function requireAdminAccess_(authToken) {
  const user = requireSession_(authToken);
  if (user.role !== 'admin' || user.status !== 'active') throw new Error('เฉพาะผู้ดูแลระบบเท่านั้นที่ลบข้อมูลได้');
  return user;
}

function initializeDefaultAdmin_(spreadsheet) {
  const sheet = spreadsheet.getSheetByName(APP.appUsersSheet);
  if (sheet.getLastRow() > 1) return;
  const salt = Utilities.getUuid().replace(/-/g, '');
  sheet.appendRow(['admin', 'ผู้ดูแลระบบ', hashPassword_('123456', salt), salt, 'admin', 'active', '', new Date(), 1]);
  setAuthVersion_('admin', 1);
  clearAuthUsersCache_();
}

function loginUser(username, password) {
  username = cleanText_(username, 60).toLowerCase();
  password = String(password == null ? '' : password);
  if (!username || !password) throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');

  const attempts = CacheService.getScriptCache();
  const attemptKey = 'LOGIN_ATTEMPT_' + digestText_(username);
  const attemptCount = Number(attempts.get(attemptKey) || 0);
  if (attemptCount >= 8) throw new Error('เข้าสู่ระบบผิดหลายครั้ง กรุณารอ 5 นาทีแล้วลองใหม่');

  const spreadsheet = getSpreadsheet_();
  setupSheets_(spreadsheet);
  const records = getAuthUserRecords_(spreadsheet);
  for (let index = 0; index < records.length; index++) {
    const record = records[index];
    if (record.username.toLowerCase() !== username) continue;
    if (record.status !== 'active') throw new Error('บัญชีนี้ถูกระงับการใช้งาน');
    const expectedHash = record.passwordHash;
    const actualHash = hashPassword_(password, record.salt);
    if (actualHash !== expectedHash) break;
    const authVersion = Number(record.authVersion || getAuthVersion_(username) || 1);
    setAuthVersion_(username, authVersion);
    const user = { username: record.username, name: record.name, role: record.role, status: record.status, authVersion: authVersion };
    const token = Utilities.getUuid().replace(/-/g, '') + Utilities.getUuid().replace(/-/g, '');
    CacheService.getScriptCache().put('AUTH_' + token, JSON.stringify(user), APP.sessionSeconds);
    PropertiesService.getScriptProperties().setProperty('LAST_LOGIN_' + digestText_(username).slice(0, 24), new Date().toISOString());
    attempts.remove(attemptKey);
    return { success: true, token: token, expiresIn: APP.sessionSeconds, user: user, initialData: buildInitialData_(user, spreadsheet) };
  }
  attempts.put(attemptKey, String(attemptCount + 1), 300);
  throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
}

function getAuthUserRecords_(spreadsheet) {
  const cache = CacheService.getScriptCache();
  const cached = cache.get(APP.authUsersCacheKey);
  if (cached) { try { return JSON.parse(cached); } catch (error) { /* อ่านจากชีตด้านล่าง */ } }
  const records = rowsWithoutHeader_(spreadsheet.getSheetByName(APP.appUsersSheet)).map(function (row) {
    return { username: String(row[0]), name: String(row[1] || row[0]), passwordHash: String(row[2]), salt: String(row[3]), role: String(row[4] || 'viewer'), status: String(row[5] || 'active'), authVersion: Number(row[8] || 1) };
  });
  const serialized = JSON.stringify(records);
  if (serialized.length < 95000) cache.put(APP.authUsersCacheKey, serialized, APP.sessionSeconds);
  return records;
}

function clearAuthUsersCache_() { CacheService.getScriptCache().remove(APP.authUsersCacheKey); }

function logoutUser(authToken) {
  const user = requireSession_(authToken);
  CacheService.getScriptCache().remove('AUTH_' + authToken);
  writeAudit_('LOGOUT', 'auth', user.username, '', '', user.username);
  return { success: true };
}

function requireSession_(authToken) {
  const token = cleanText_(authToken, 160);
  if (!token) throw new Error('AUTH_REQUIRED');
  const value = CacheService.getScriptCache().get('AUTH_' + token);
  if (!value) throw new Error('SESSION_EXPIRED');
  let user;
  try { user = JSON.parse(value); } catch (error) { throw new Error('SESSION_EXPIRED'); }
  if (!user || user.status !== 'active') throw new Error('AUTH_REQUIRED');
  if (Number(user.authVersion || 1) !== getAuthVersion_(user.username)) throw new Error('SESSION_EXPIRED');
  CacheService.getScriptCache().put('AUTH_' + token, JSON.stringify(user), APP.sessionSeconds);
  return user;
}

function readAppUsers_(spreadsheet) {
  return rowsWithoutHeader_(spreadsheet.getSheetByName(APP.appUsersSheet)).map(function (row) {
    const username = String(row[0]);
    const fastLastLogin = PropertiesService.getScriptProperties().getProperty('LAST_LOGIN_' + digestText_(username.toLowerCase()).slice(0, 24));
    return { email: username, username: username, name: String(row[1] || row[0]), role: String(row[4] || 'viewer'), status: String(row[5] || 'active'), lastActive: fastLastLogin || serializeDate_(row[6]) };
  });
}

function hashPassword_(password, salt) {
  return digestText_(String(salt) + ':' + String(password));
}

function digestText_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, text, Utilities.Charset.UTF_8);
  return bytes.map(function (byte) { const value = byte < 0 ? byte + 256 : byte; return ('0' + value.toString(16)).slice(-2); }).join('');
}

function authVersionKey_(username) { return 'AUTH_VERSION_' + digestText_(String(username).toLowerCase()).slice(0, 24); }
function getAuthVersion_(username) { return Number(PropertiesService.getScriptProperties().getProperty(authVersionKey_(username)) || 1); }
function setAuthVersion_(username, version) { PropertiesService.getScriptProperties().setProperty(authVersionKey_(username), String(version)); }

function readAuditLogs_(spreadsheet, limit) {
  const rows = rowsWithoutHeader_(spreadsheet.getSheetByName(APP.auditSheet));
  return rows.slice(-limit).reverse().map(function (row) {
    return { id: String(row[0]), timestamp: serializeDate_(row[1]), user: String(row[2]), action: String(row[3]), entity: String(row[4]), recordId: String(row[5]), before: String(row[6] || ''), after: String(row[7] || '') };
  });
}

function writeAudit_(action, entity, recordId, before, after, actor, spreadsheet) {
  spreadsheet = spreadsheet || getSpreadsheet_();
  let sheet = spreadsheet.getSheetByName(APP.auditSheet);
  if (!sheet) {
    setupSheets_(spreadsheet, true);
    sheet = spreadsheet.getSheetByName(APP.auditSheet);
  }
  sheet.appendRow([Utilities.getUuid(), new Date(), actor || 'system', action, entity, recordId, before, after]);
}

function updateUserRole(username, role, status, authToken) {
  return withLock_(function () {
    const currentUser = requireSession_(authToken);
    const spreadsheet = getSpreadsheet_();
    if (currentUser.role !== 'admin') throw new Error('เฉพาะผู้ดูแลระบบเท่านั้น');
    if (['admin', 'editor', 'viewer'].indexOf(role) === -1) throw new Error('Role ไม่ถูกต้อง');
    if (['active', 'suspended'].indexOf(status) === -1) throw new Error('สถานะไม่ถูกต้อง');
    const sheet = spreadsheet.getSheetByName(APP.appUsersSheet);
    const rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]).toLowerCase() === String(username).toLowerCase()) {
        const before = JSON.stringify({ role: rows[index][4], status: rows[index][5] });
        sheet.getRange(index + 2, 5, 1, 2).setValues([[role, status]]);
        clearAuthUsersCache_();
        writeAudit_('UPDATE_USER', 'user', username, before, JSON.stringify({ role: role, status: status }), currentUser.username);
        clearDataCache_();
        return { success: true, message: 'อัปเดตสิทธิ์ผู้ใช้เรียบร้อยแล้ว' };
      }
    }
    throw new Error('ไม่พบผู้ใช้งาน');
  });
}

function createAppUser(payload, authToken) {
  return withLock_(function () {
    const admin = requireSession_(authToken);
    if (admin.role !== 'admin') throw new Error('เฉพาะผู้ดูแลระบบเท่านั้น');
    const username = cleanText_(payload && payload.username, 60).toLowerCase();
    const name = cleanText_(payload && payload.name, 100);
    const password = String(payload && payload.password || '');
    const role = cleanText_(payload && payload.role, 20);
    if (!/^[a-z0-9._-]{3,60}$/.test(username)) throw new Error('Username ต้องมีอย่างน้อย 3 ตัว และใช้เฉพาะ a-z, 0-9, จุด, ขีดกลาง หรือขีดล่าง');
    if (password.length < 6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัว');
    if (['admin', 'editor', 'viewer'].indexOf(role) === -1) throw new Error('Role ไม่ถูกต้อง');
    const sheet = getSpreadsheet_().getSheetByName(APP.appUsersSheet);
    const rows = rowsWithoutHeader_(sheet);
    if (rows.some(function (row) { return String(row[0]).toLowerCase() === username; })) throw new Error('Username นี้มีอยู่แล้ว');
    const salt = Utilities.getUuid().replace(/-/g, '');
    sheet.appendRow([username, name || username, hashPassword_(password, salt), salt, role, 'active', '', new Date(), 1]);
    setAuthVersion_(username, 1);
    clearAuthUsersCache_();
    writeAudit_('CREATE_USER', 'user', username, '', JSON.stringify({ name: name, role: role }), admin.username);
    return { success: true, message: 'เพิ่มสมาชิกเรียบร้อยแล้ว' };
  });
}

function updateAppUser(payload, authToken) {
  return withLock_(function () {
    const admin = requireSession_(authToken);
    if (admin.role !== 'admin') throw new Error('เฉพาะผู้ดูแลระบบเท่านั้น');
    const username = cleanText_(payload && payload.username, 60).toLowerCase();
    const name = cleanText_(payload && payload.name, 100);
    const password = String(payload && payload.password || '');
    const role = cleanText_(payload && payload.role, 20);
    const status = cleanText_(payload && payload.status, 20);
    if (password && password.length < 6) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัว');
    if (['admin', 'editor', 'viewer'].indexOf(role) === -1) throw new Error('Role ไม่ถูกต้อง');
    if (['active', 'suspended'].indexOf(status) === -1) throw new Error('สถานะไม่ถูกต้อง');
    if (username === admin.username && (role !== 'admin' || status !== 'active')) throw new Error('ไม่สามารถลดสิทธิ์หรือระงับบัญชีที่กำลังใช้งาน');
    const sheet = getSpreadsheet_().getSheetByName(APP.appUsersSheet);
    const rows = rowsWithoutHeader_(sheet);
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]).toLowerCase() !== username) continue;
      const before = JSON.stringify({ name: rows[index][1], role: rows[index][4], status: rows[index][5] });
      sheet.getRange(index + 2, 2).setValue(name || username);
      sheet.getRange(index + 2, 5, 1, 2).setValues([[role, status]]);
      if (password) {
        const salt = Utilities.getUuid().replace(/-/g, '');
        sheet.getRange(index + 2, 3, 1, 2).setValues([[hashPassword_(password, salt), salt]]);
      }
      const nextVersion = getAuthVersion_(username) + 1;
      sheet.getRange(index + 2, 9).setValue(nextVersion);
      setAuthVersion_(username, nextVersion);
      clearAuthUsersCache_();
      writeAudit_('UPDATE_USER', 'user', username, before, JSON.stringify({ name: name, role: role, status: status, passwordChanged: Boolean(password) }), admin.username);
      return { success: true, message: 'แก้ไขสมาชิกเรียบร้อยแล้ว' };
    }
    throw new Error('ไม่พบสมาชิก');
  });
}

function deleteAppUser(username, authToken) {
  return withLock_(function () {
    const admin = requireSession_(authToken);
    if (admin.role !== 'admin') throw new Error('เฉพาะผู้ดูแลระบบเท่านั้น');
    username = cleanText_(username, 60).toLowerCase();
    if (username === admin.username) throw new Error('ไม่สามารถลบบัญชีที่กำลังใช้งาน');
    const sheet = getSpreadsheet_().getSheetByName(APP.appUsersSheet);
    const rows = rowsWithoutHeader_(sheet);
    const adminCount = rows.filter(function (row) { return row[4] === 'admin' && row[5] === 'active'; }).length;
    for (let index = 0; index < rows.length; index++) {
      if (String(rows[index][0]).toLowerCase() !== username) continue;
      if (rows[index][4] === 'admin' && adminCount <= 1) throw new Error('ต้องมีผู้ดูแลระบบอย่างน้อย 1 บัญชี');
      const before = JSON.stringify({ username: rows[index][0], name: rows[index][1], role: rows[index][4] });
      sheet.deleteRow(index + 2);
      setAuthVersion_(username, getAuthVersion_(username) + 1);
      clearAuthUsersCache_();
      writeAudit_('DELETE_USER', 'user', username, before, '', admin.username);
      return { success: true, message: 'ลบสมาชิกเรียบร้อยแล้ว' };
    }
    throw new Error('ไม่พบสมาชิก');
  });
}

function normalizeProfitDistributionStatus_(value) {
  const status = cleanText_(value, 30) || 'pending';
  if (APP.profitDistributionStatuses.indexOf(status) === -1) throw new Error('สถานะแบ่งผลกำไรไม่ถูกต้อง');
  return status;
}

function readProfitDistributionStatus_(value) {
  return String(value || '').trim() === 'distributed' ? 'distributed' : 'pending';
}

function parseShareholders_(value) {
  if (value == null || value === '') return [];
  let list = value;
  if (typeof list === 'string') {
    try { list = JSON.parse(list); }
    catch (error) { throw new Error('ข้อมูลผู้ถือหุ้นไม่ถูกต้อง กรุณาตรวจสอบอีกครั้ง'); }
  }
  if (!Array.isArray(list)) throw new Error('ข้อมูลผู้ถือหุ้นต้องเป็นรายการ');
  if (list.length > 30) throw new Error('กำหนดผู้ถือหุ้นได้ไม่เกิน 30 คนต่อโปรเจกต์');
  const names = {};
  const normalized = list.map(function (item) {
    const name = cleanText_(item && item.name, 100);
    const percent = Math.round(Number(item && item.percent) * 100) / 100;
    if (!name) throw new Error('กรุณาระบุชื่อผู้ถือหุ้นให้ครบ');
    if (!Number.isFinite(percent) || percent <= 0 || percent > 100) throw new Error('สัดส่วนผู้ถือหุ้นแต่ละคนต้องมากกว่า 0 และไม่เกิน 100%');
    const key = name.toLowerCase();
    if (names[key]) throw new Error('ชื่อผู้ถือหุ้นซ้ำ: ' + name);
    names[key] = true;
    return { name: name, percent: percent };
  });
  const totalBasisPoints = normalized.reduce(function (sum, item) { return sum + Math.round(item.percent * 100); }, 0);
  if (normalized.length && totalBasisPoints !== 10000) throw new Error('สัดส่วนผู้ถือหุ้นรวมต้องเท่ากับ 100% ปัจจุบันรวม ' + (totalBasisPoints / 100).toFixed(2) + '%');
  return normalized;
}

function readStoredShareholders_(value) {
  try { return parseShareholders_(value); }
  catch (error) { return []; }
}

function cleanText_(value, maxLength) {
  return String(value == null ? '' : value).trim().slice(0, maxLength);
}

function toAmount_(value, allowZero) {
  const amount = Number(String(value == null ? '' : value).replace(/,/g, ''));
  if (!Number.isFinite(amount) || amount < 0 || (!allowZero && amount <= 0)) {
    throw new Error('จำนวนเงินไม่ถูกต้อง');
  }
  return Math.round(amount * 100) / 100;
}

function toProgress_(value) {
  if (value == null || value === '') return 0;
  const progress = Number(value);
  if (!Number.isFinite(progress) || progress < 0 || progress > 100) throw new Error('ความคืบหน้าต้องอยู่ระหว่าง 0–100%');
  return Math.round(progress);
}

function parseDate_(value) {
  const text = cleanText_(value, 10);
  const date = text ? new Date(text + 'T00:00:00+07:00') : new Date();
  if (isNaN(date.getTime())) throw new Error('วันที่ไม่ถูกต้อง');
  return date;
}

function serializeDate_(value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) return '';
  return Utilities.formatDate(value, APP.timezone, "yyyy-MM-dd'T'HH:mm:ssXXX");
}

function serializeDateOnly_(value) {
  if (!(value instanceof Date) || isNaN(value.getTime())) return String(value || '');
  return Utilities.formatDate(value, APP.timezone, 'yyyy-MM-dd');
}

function withLock_(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    setupSheets_();
    return callback();
  } finally {
    lock.releaseLock();
  }
}
