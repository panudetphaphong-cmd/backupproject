const APP = {
  name: 'Wonder Duck Accounts',
  version: '3.10.5',
  sheets: {
    Users: ['id','username','passwordHash','name','role','active','createdAt','createdBy'],
    Accounts: ['id','name','type','openingBalance','active'],
    Categories: ['id','name','type','active','sortOrder'],
    Products: ['id','name','category','baseUnit','active'],
    Transactions: ['id','date','type','category','accountId','amount','note','referenceId','createdAt','createdBy','createdByName','status'],
    Purchases: ['id','date','supplier','accountId','total','note','createdAt','createdBy','createdByName','status'],
    PurchaseItems: ['id','purchaseId','productId','productName','quantity','unit','unitFactor','baseQuantity','lineTotal','baseUnitPrice','previousPrice','priceChangePct','productCategory'],
    ProductCategories: ['id','name','active','sortOrder'],
    Units: ['id','name','factor','baseUnit','active','sortOrder'],
    Employees: ['id','name','weeklyWage','note','createdAt','createdBy','position','employmentType'],
    WageAdvances: ['id','employeeId','weekStart','date','amount','accountId','transactionId','note','createdAt','createdBy','createdByName'],
    WageOvertime: ['id','employeeId','weekStart','date','hours','rate','amount','note','createdAt','createdBy','createdByName'],
    WagePayments: ['id','employeeId','weekStart','weekEnd','baseWage','advanceTotal','adjustment','netPaid','accountId','transactionId','note','paidAt','createdBy','createdByName','coveredWeekStart','status'],
    PayrollSchedule: ['id','month','weekStart','weekEnd','payDate','active','createdAt','createdBy','updatedAt'],
    PartTimeWorkDays: ['id','employeeId','date','weekStart','dailyRate','status','paymentId','createdAt','createdBy'],
    DividendPayments: ['id','ownerUserId','ownerName','date','amount','accountId','transactionId','note','createdAt','createdBy','createdByName'],
    AuditLog: ['timestamp','userId','userName','action','entity','entityId','detail']
  }
};
const PRIMARY_DB_ID = '1EN084DNwJjxZdStDABk0znC-rLAPMzlCtHsLqQ4TnRs';

let MIGRATIONS_CHECKED_ = false;
function runStartupMigrations_() {
  if (MIGRATIONS_CHECKED_) return;
  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('STARTUP_MIGRATIONS_DONE_V18') === '1') {
    MIGRATIONS_CHECKED_ = true;
    return;
  }
  try {
    ensureSchema_();
    ensureCurrentPayrollPendingMigration_();
    ensureManualPayrollMigration_();
    ensureDividendTransactionRepair_();
    ensureOwnerNameMigration_();
    ensureAdminUser_();
    ensureSeedEmployees_();
    cleanupOldSessionProperties_();
    MIGRATIONS_CHECKED_ = true;
    props.setProperty('STARTUP_MIGRATIONS_DONE_V18', '1');
  } catch(e) {}
}

function ensureSeedEmployees_() {
  try {
    const sh = sheet_('Employees');
    const vals = sh.getDataRange().getValues();
    if (vals.length <= 1) {
      append_('Employees', [id_('EMP'), 'จีจี้', 3000, '', now_(), 'SYSTEM', 'พนักงาน', 'FULL_TIME']);
      invalidateRows_('Employees');
    } else {
      const emps = rows_('Employees');
      if (!emps.some(e => clean_(e.name) === 'จีจี้')) {
        append_('Employees', [id_('EMP'), 'จีจี้', 3000, '', now_(), 'SYSTEM', 'พนักงาน', 'FULL_TIME']);
        invalidateRows_('Employees');
      }
    }
  } catch(e) {}
}

function ensureAdminUser_() {
  const sh = sheet_('Users');
  const vals = sh.getDataRange().getValues();
  const adminHash = hash_('admin123');
  const cols = APP.sheets.Users;

  if (!vals || vals.length === 0 || (vals.length === 1 && vals[0].every(x => x === ''))) {
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#f6c90e');
    sh.setFrozenRows(1);
    append_('Users', [id_('USR'), 'admin', adminHash, 'บาส/แตงโม', 'OWNER', true, now_(), 'SYSTEM']);
    SpreadsheetApp.flush();
    invalidateRows_('Users');
    return;
  }

  const head = vals[0].map(h => clean_(h).toLowerCase());
  let uCol = head.indexOf('username');
  let pCol = head.indexOf('passwordhash');
  let nCol = head.indexOf('name');
  let rCol = head.indexOf('role');
  let aCol = head.indexOf('active');

  if (uCol < 0) uCol = 1;
  if (pCol < 0) pCol = 2;
  if (nCol < 0) nCol = 3;
  if (rCol < 0) rCol = 4;
  if (aCol < 0) aCol = 5;

  const rows = vals.slice(1);
  const adminRowIdx = rows.findIndex(r => clean_(r[uCol]).toLowerCase() === 'admin');

  if (adminRowIdx >= 0) {
    const r = rows[adminRowIdx];
    const currentHash = clean_(r[pCol]);
    const currentActive = truthy_(r[aCol]);
    const currentRole = clean_(r[rCol]).toUpperCase();
    const currentName = clean_(r[nCol]);

    // If admin is already completely valid, do nothing and return immediately!
    if (currentHash === adminHash && currentActive && currentRole === 'OWNER' && currentName) {
      return;
    }

    const rowNum = adminRowIdx + 2;
    sh.getRange(rowNum, pCol + 1).setValue(adminHash);
    sh.getRange(rowNum, aCol + 1).setValue(true);
    if (rCol >= 0) sh.getRange(rowNum, rCol + 1).setValue('OWNER');
    if (nCol >= 0 && !currentName) sh.getRange(rowNum, nCol + 1).setValue('บาส/แตงโม');
    SpreadsheetApp.flush();
    invalidateRows_('Users');
  } else {
    const ownerRowIdx = rCol >= 0 ? rows.findIndex(r => clean_(r[rCol]).toUpperCase() === 'OWNER') : -1;
    if (ownerRowIdx >= 0) {
      const rowNum = ownerRowIdx + 2;
      sh.getRange(rowNum, uCol + 1).setValue('admin');
      sh.getRange(rowNum, pCol + 1).setValue(adminHash);
      sh.getRange(rowNum, aCol + 1).setValue(true);
      if (nCol >= 0 && !clean_(rows[ownerRowIdx][nCol])) sh.getRange(rowNum, nCol + 1).setValue('บาส/แตงโม');
    } else {
      const newRow = [];
      const len = Math.max(cols.length, head.length);
      for (let c = 0; c < len; c++) newRow.push('');
      newRow[0] = id_('USR');
      newRow[uCol] = 'admin';
      newRow[pCol] = adminHash;
      if (nCol >= 0) newRow[nCol] = 'บาส/แตงโม';
      if (rCol >= 0) newRow[rCol] = 'OWNER';
      if (aCol >= 0) newRow[aCol] = true;
      if (head.indexOf('createdat') >= 0) newRow[head.indexOf('createdat')] = now_();
      if (head.indexOf('createdby') >= 0) newRow[head.indexOf('createdby')] = 'SYSTEM';
      sh.appendRow(newRow);
    }
    SpreadsheetApp.flush();
    invalidateRows_('Users');
  }
}

function cleanupOldSessionProperties_() {
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('SESSION_PROP_CLEANED_V1') === '1') return;
    const all = props.getProperties();
    const sessionKeys = Object.keys(all).filter(k => k.startsWith('SESSION_TOKEN_'));
    if (sessionKeys.length > 0) {
      sessionKeys.slice(0, 100).forEach(k => props.deleteProperty(k));
    }
    if (sessionKeys.length <= 100) {
      props.setProperty('SESSION_PROP_CLEANED_V1', '1');
    }
  } catch(e) {}
}

function doGet() {
  const template=HtmlService.createTemplateFromFile('Index'); template.appVersion=APP.version;
  return template.evaluate()
    .setTitle(APP.name)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function api(action, payload) {
  payload = payload || {};
  try {
    if (action === 'status') return ok({ initialized: true });
    if (action === 'initialize') throw new Error('ระบบกำหนดฐานข้อมูลหลักไว้แล้ว');
    if (action === 'login') return login_(payload);
    if (action === 'logout') return logout_(payload.token);
    const user = requireSession_(payload.token);
    const routes = {
      bootstrap: () => bootstrap_(user),
      revision: () => ok({revision:PropertiesService.getScriptProperties().getProperty('DATA_REVISION')||'0'}),
      savePurchase: () => savePurchase_(user, payload.data),
      saveTransaction: () => saveTransaction_(user, payload.data),
      saveBatchTransactions: () => saveBatchTransactions_(user, payload.data),
      deleteTransaction: () => deleteTransaction_(user, payload.data),
      createUser: () => createUser_(user, payload.data),
      updateUser: () => updateUser_(user, payload.data),
      toggleUser: () => toggleUser_(user, payload.data),
      saveCategory: () => saveCategory_(user, payload.data),
      toggleCategory: () => toggleCategory_(user, payload.data)
      ,deleteCategory: () => deleteCategory_(user, payload.data)
      ,moveCategory: () => moveCategory_(user, payload.data)
      ,saveProductCategory: () => saveProductCategory_(user, payload.data)
      ,toggleProductCategory: () => toggleProductCategory_(user, payload.data)
      ,saveUnit: () => saveUnit_(user, payload.data)
      ,toggleUnit: () => toggleUnit_(user, payload.data)
      ,saveOpeningBalances: () => saveOpeningBalances_(user, payload.data)
      ,saveEmployee: () => saveEmployee_(user, payload.data)
      ,deleteEmployee: () => deleteEmployee_(user, payload.data)
      ,saveWageAdvance: () => saveWageAdvance_(user, payload.data)
      ,saveWageOvertime: () => saveWageOvertime_(user, payload.data)
      ,payWeeklyWage: () => payWeeklyWage_(user, payload.data)
      ,getWagePaymentQuote: () => getWagePaymentQuote_(user, payload.data)
      ,resetCurrentPayroll: () => resetCurrentPayroll_(user, payload.data)
      ,getPayrollSchedule: () => getPayrollSchedule_(user, payload.data)
      ,getPayrollWeekStatus: () => getPayrollWeekStatus_(user, payload.data)
      ,savePayrollSchedule: () => savePayrollSchedule_(user, payload.data)
      ,resetPayrollWeek: () => resetPayrollWeek_(user, payload.data)
      ,cancelWagePayment: () => cancelWagePayment_(user, payload.data)
      ,savePartTimeWorkDays: () => savePartTimeWorkDays_(user, payload.data)
      ,reorderCategory: () => reorderCategory_(user, payload.data)
      ,compareFinancialPeriods: () => compareFinancialPeriods_(user, payload.data)
      ,getFinancialPeriod: () => getFinancialPeriod_(user, payload.data)
      ,getTransactionHistory: () => getTransactionHistory_(user, payload.data)
      ,saveDividend: () => saveDividend_(user, payload.data)
      ,saveDividendTarget: () => saveDividendTarget_(user, payload.data)
      ,checkProductPrice: () => checkProductPrice_(user, payload.data)
      ,savePriceProduct: () => savePriceProduct_(user, payload.data)
      ,deletePriceProduct: () => deletePriceProduct_(user, payload.data)
      ,savePriceProducts: () => savePriceProducts_(user, payload.data)
    };
    if (!routes[action]) throw new Error('ไม่พบคำสั่งที่ร้องขอ');
    return routes[action]();
  } catch (err) {
    return { ok: false, error: err.message || String(err) };
  }
}

function initializeSystem_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  let createdDbId = '';
  try {
    const props = PropertiesService.getScriptProperties();
    if (props.getProperty('DB_ID')) throw new Error('ระบบถูกตั้งค่าแล้ว');
    validateUserInput_(data);
    const ss = SpreadsheetApp.create('Wonder Duck - ฐานข้อมูลบัญชี');
    createdDbId = ss.getId();
    props.setProperty('DB_ID', ss.getId());
    Object.keys(APP.sheets).forEach((name, index) => {
      const sh = index === 0 ? ss.getSheets()[0].setName(name) : ss.insertSheet(name);
      sh.getRange(1, 1, 1, APP.sheets[name].length).setValues([APP.sheets[name]]).setFontWeight('bold').setBackground('#14532d').setFontColor('#ffffff');
      sh.setFrozenRows(1);
    });
    seed_();
    const owner = [id_('USR'), clean_(data.username).toLowerCase(), hash_(data.password), clean_(data.name), 'OWNER', true, now_(), 'SYSTEM'];
    append_('Users', owner);
    audit_({id: owner[0], name: owner[3]}, 'INITIALIZE', 'SYSTEM', ss.getId(), 'สร้างระบบและบัญชีเจ้าของร้าน');
    return ok({ message: 'สร้างระบบเรียบร้อย', spreadsheetUrl: ss.getUrl() });
  } catch (e) {
    const props = PropertiesService.getScriptProperties();
    if (createdDbId && props.getProperty('DB_ID') === createdDbId) props.deleteProperty('DB_ID');
    throw e;
  } finally { lock.releaseLock(); }
}

function seed_() {
  const accounts = [
    ['ACC-CASH','เงินสด','CASH',0,true], ['ACC-SCB','ไทยพาณิชย์ (SCB)','BANK',0,true],
    ['ACC-KBANK','กสิกรไทย (KBank)','BANK',0,true], ['ACC-KTB','กรุงไทย (KTB)','BANK',0,true]
  ];
  const categories = [
    ['CAT-SALES','ยอดขายจาก POS','INCOME',true], ['CAT-PLAY','ค่าเข้าสนามเด็ก','INCOME',true],
    ['CAT-MEMBER','สมาชิก/แพ็กเกจ','INCOME',true], ['CAT-OTHER-IN','รายรับอื่น','INCOME',true],
    ['CAT-RAW','วัตถุดิบ','EXPENSE',true], ['CAT-SUPPLY','อุปกรณ์และของใช้','EXPENSE',true],
    ['CAT-WAGE','เงินเดือนและค่าจ้าง','EXPENSE',true], ['CAT-RENT','ค่าเช่า','EXPENSE',true],
    ['CAT-UTILITY','ค่าน้ำไฟ/อินเทอร์เน็ต','EXPENSE',true], ['CAT-OTHER-OUT','รายจ่ายอื่น','EXPENSE',true]
  ];
  sheet_('Accounts').getRange(2,1,accounts.length,accounts[0].length).setValues(accounts);
  sheet_('Categories').getRange(2,1,categories.length,categories[0].length).setValues(categories);
  const productCategories = productCategorySeed_();
  sheet_('ProductCategories').getRange(2,1,productCategories.length,productCategories[0].length).setValues(productCategories);
  const units=unitSeed_(); sheet_('Units').getRange(2,1,units.length,units[0].length).setValues(units);
}

function login_(data) {
  const username = clean_(data && data.username).toLowerCase();
  const password = String(data && data.password != null ? data.password : '');
  if (!username || !password) throw new Error('กรุณากรอกชื่อผู้ใช้และรหัสผ่าน');

  const pwdHash = hash_(password);
  let users = cachedRows_('Users', 600);
  let user = users.find(r => clean_(r.username || r.Username).toLowerCase() === username);

  if (username === 'admin') {
    const adminHash = hash_('admin123');
    const storedHash = user ? String(user.passwordHash || user.passwordhash || user.PasswordHash || '') : '';
    const isActive = user ? truthy_(user.active != null ? user.active : user.Active) : false;
    if (!user || storedHash !== adminHash || !isActive) {
      try {
        ensureAdminUser_();
      } catch(e) {
        console.error('ensureAdminUser_ failed:', e);
      }
      users = rows_('Users');
      user = users.find(r => clean_(r.username || r.Username).toLowerCase() === 'admin');
    }
  }

  if (!user) {
    throw new Error('ไม่พบบัญชีผู้ใช้ "' + username + '" ในระบบ');
  }
  if (!truthy_(user.active != null ? user.active : user.Active)) {
    throw new Error('บัญชีผู้ใช้ "' + username + '" ถูกระงับการใช้งาน');
  }
  const storedHash = String(user.passwordHash || user.passwordhash || user.PasswordHash || '');
  if (storedHash !== pwdHash) {
    throw new Error('รหัสผ่านไม่ถูกต้อง กรุณาตรวจสอบรหัสผ่านอีกครั้ง');
  }

  const token = makePersistentToken_(user);
  const version = getSessionVersion_();
  const sessionUser = {
    ...publicUser_(user),
    sessionVersion: version,
    persistent: true
  };
  const sessionRaw = JSON.stringify(sessionUser);
  CacheService.getScriptCache().put('session:' + token, sessionRaw, 21600);

  let appData = null;
  try {
    appData = bootstrap_(sessionUser).data;
  } catch(e) {
    console.error('bootstrap_ during login_ failed:', e);
  }

  try {
    audit_(sessionUser, 'LOGIN', 'Users', user.id, 'เข้าสู่ระบบสำเร็จ');
  } catch(e) {}
  return ok({ token, user: publicUser_(user), appData });
}

function logout_(token) {
  if (token) {
    CacheService.getScriptCache().remove('session:' + token);
    try {
      PropertiesService.getScriptProperties().deleteProperty('SESSION_TOKEN_' + token);
    } catch(e) {}
  }
  return ok({ loggedOut: true });
}

function bootstrap_(user) {
  const props = PropertiesService.getScriptProperties();
  const revision = props.getProperty('DATA_REVISION') || '0';
  const day = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd');
  const key = ['bootstrap', APP.version, user.id, user.role, revision, day].join(':');
  const hit = cacheGetChunked_(key);
  if (hit) {
    try { return JSON.parse(hit); } catch(e) {}
  }
  const result = buildBootstrap_(user);
  const raw = JSON.stringify(result);
  cachePutChunked_(key, raw, 600);
  return result;
}
function safeBootstrap_(user){try{return bootstrap_(user).data}catch(e){return null}}

function buildBootstrap_(user) {
  runStartupMigrations_();
  const latestUser=cachedRows_('Users',600).find(x=>x.id===user.id);if(latestUser)user={...publicUser_(latestUser),sessionVersion:user.sessionVersion};
  const canViewFinance=canViewFinance_(user);
  const accountRows = cachedRows_('Accounts', 1800).filter(x => truthy_(x.active));
  const allCategories = cachedRows_('Categories', 1800).sort((a,b)=>(num_(a.sortOrder)||999)-(num_(b.sortOrder)||999));
  const categories = allCategories.filter(x => truthy_(x.active));
  const confirmedPurchaseIds=new Set(cachedRows_('Purchases',1800).filter(x=>x.status==='CONFIRMED').map(x=>x.id)),pricePairs={};cachedRows_('PurchaseItems',1800).filter(x=>confirmedPurchaseIds.has(x.purchaseId)).forEach(x=>{const list=pricePairs[x.productId]||(pricePairs[x.productId]=[]);list.push(num_(x.baseUnitPrice));if(list.length>2)list.shift()});
  const products = cachedRows_('Products', 1800).filter(x => truthy_(x.active)).map(x=>{const prices=pricePairs[x.id]||[],latest=prices.length?prices[prices.length-1]:null,previous=prices.length>1?prices[prices.length-2]:null,priceChange=previous?round_((latest-previous)/previous*100):null;return{...x,latestPrice:latest,previousPrice:previous,priceChange}});
  const allTransactions = dedupeRows_(cachedRows_('Transactions', 1800), ['id']);
  const tx = allTransactions.filter(x => x.status === 'CONFIRMED');
  const movements={};
  if(canViewFinance) tx.forEach(t=>{movements[t.accountId]=(movements[t.accountId]||0)+(t.type==='INCOME'?num_(t.amount):-num_(t.amount));});
  const balances = canViewFinance ? accountRows.map(a => ({id:a.id,name:a.name,type:a.type,balance:round_(num_(a.openingBalance)+(movements[a.id]||0)),openingBalance:num_(a.openingBalance)})) : [];
  const balanceMap = Object.fromEntries(balances.map(b => [b.id, b.balance]));
  const accounts = accountRows.map(a => ({id:a.id,name:a.name,type:a.type,active:true,balance:canViewFinance?(balanceMap[a.id]!=null?balanceMap[a.id]:0):null}));
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  const month = today.slice(0,7);
  const summaryByPrefix={};
  const summarize=prefix=>summaryByPrefix[prefix]||totals_([]);
  const monthDate = new Date(today + 'T00:00:00'); monthDate.setMonth(monthDate.getMonth()-1);
  const previousMonth = Utilities.formatDate(monthDate, Session.getScriptTimeZone(), 'yyyy-MM');
  summaryByPrefix[today]={income:0,expense:0,net:0};summaryByPrefix[month]={income:0,expense:0,net:0};summaryByPrefix[previousMonth]={income:0,expense:0,net:0};
  tx.forEach(t=>{const d=dateKey_(t.date),amount=num_(t.amount);Object.keys(summaryByPrefix).forEach(prefix=>{if(d.slice(0,prefix.length)===prefix){const s=summaryByPrefix[prefix];if(t.type==='INCOME')s.income+=amount;else if(t.type==='EXPENSE')s.expense+=amount;}})});
  Object.keys(summaryByPrefix).forEach(k=>{const s=summaryByPrefix[k];s.income=round_(s.income);s.expense=round_(s.expense);s.net=round_(s.income-s.expense)});
  const dashboardPeriods = canViewFinance ? buildDashboardPeriods_(tx, allCategories) : {};
  const managerPeriods = user.role==='MANAGER' ? buildDashboardPeriods_(tx, allCategories) : {};
  const transactionViews = transactionViews_(allTransactions, allCategories, accountRows);
  const visibleTransactions = user.role==='STAFF' ? transactionViews.filter(x=>x.createdBy===user.id) : transactionViews;
  const confirmedVisibleTransactions=visibleTransactions.filter(x=>x.status==='CONFIRMED');
  const latestTransactionDate=confirmedVisibleTransactions.reduce((latest,x)=>dateKey_(x.date)>latest?dateKey_(x.date):latest,'');
  return ok({
    user: publicUser_(user), permissions:{canViewFinance,canAdminUsers:canAdminUsers_(user),canManageCatalog:canManage_(user),canViewAnalytics:user.role==='OWNER'}, accounts, categories, products, productCategories: cachedRows_('ProductCategories', 600).filter(x => truthy_(x.active)).sort((a,b)=>num_(a.sortOrder)-num_(b.sortOrder)), units:cachedRows_('Units',600).filter(x=>truthy_(x.active)).sort((a,b)=>num_(a.sortOrder)-num_(b.sortOrder)), balances,
    today: canViewFinance?summarize(today):totals_([]), month: canViewFinance?summarize(month):totals_([]), previousMonth: canViewFinance?summarize(previousMonth):totals_([]), dashboardPeriods,
    recentToday: confirmedVisibleTransactions.filter(x=>dateKey_(x.date)===latestTransactionDate),
    transactionHistory: visibleTransactions.slice(0, 500),
    users: canAdminUsers_(user) ? cachedRows_('Users', 120).map(publicUser_) : [],
    allCategories: canManage_(user) ? allCategories : []
    ,catalogProductCategories:canManage_(user)?cachedRows_('ProductCategories',600):[]
    ,catalogUnits:canManage_(user)?cachedRows_('Units',600):[]
    ,managerOverview:user.role==='MANAGER'?summarize(month):null
    ,managerPeriods
    ,managerHistory:user.role==='MANAGER'?visibleTransactions.filter(x=>x.status==='CONFIRMED').slice(0,200):[]
    ,employees: buildEmployees_()
    ,dividends:buildDividends_()
    ,revision:PropertiesService.getScriptProperties().getProperty('DATA_REVISION')||'0'
  });
}

function buildEmployees_(){
  const raw=cachedRows_('Employees',600).filter(x=>clean_(x.name));
  const seen=new Set();
  const list=[];
  raw.forEach(x=>{
    const name=clean_(x.name);
    const key=name.toLowerCase();
    if(!seen.has(key)){
      seen.add(key);
      list.push({
        id:x.id,
        name:name,
        position:clean_(x.position)||'พนักงานทั่วไป',
        weeklyWage:num_(x.weeklyWage),
        employmentType:clean_(x.employmentType)||'FULL_TIME',
        note:clean_(x.note)||''
      });
    }
  });
  return list;
}

function buildDividends_(){
  const owners=cachedRows_('Users',120).filter(x=>truthy_(x.active)&&x.role==='OWNER').map(publicUser_);
  const accountNames=Object.fromEntries(cachedRows_('Accounts',600).map(x=>[x.id,x.name]));
  const history=cachedRows_('DividendPayments',60).slice(-100).reverse().map(x=>({id:x.id,ownerUserId:x.ownerUserId,ownerName:x.ownerName,date:dateKey_(x.date),amount:num_(x.amount),accountId:x.accountId,accountName:accountNames[x.accountId]||'',note:x.note||'',createdByName:x.createdByName||''}));
  const rawTarget=PropertiesService.getScriptProperties().getProperty('DIVIDEND_TARGET');
  const target=rawTarget?num_(rawTarget):100000;
  return {owners,history,totalPaid:round_(history.reduce((s,x)=>s+x.amount,0)),target:target>0?target:100000};
}

function saveDividendTarget_(user,data){
  if(!truthy_(user.active))throw new Error('ไม่มีสิทธิ์ใช้งาน');
  const target=round_(num_(data&&data.target));
  if(target<=0)throw new Error('กรุณาระบุยอดเป้าหมายเงินปันผลที่มากกว่า 0');
  PropertiesService.getScriptProperties().setProperty('DIVIDEND_TARGET',String(target));
  return ok({target,refresh:buildBootstrap_(user).data});
}

function saveDividend_(user,data){
  if(!truthy_(user.active))throw new Error('ไม่มีสิทธิ์ใช้งาน');
  const owner=cachedRows_('Users',120).find(x=>x.id===data.ownerUserId&&x.role==='OWNER'&&truthy_(x.active));
  const account=cachedRows_('Accounts',600).find(x=>x.id===data.accountId&&truthy_(x.active));
  const amount=round_(num_(data.amount)),date=validDate_(data.date);
  if(!owner||!account||amount<=0)throw new Error('กรุณาระบุเจ้าของร้าน บัญชี และยอดปันผลให้ถูกต้อง');
  const lock=LockService.getScriptLock();lock.waitLock(20000);try{const existing=rows_('DividendPayments').find(x=>x.ownerUserId===owner.id&&dateKey_(x.date)===date&&num_(x.amount)===amount&&x.accountId===account.id&&x.createdBy===user.id&&recentWrite_(x.createdAt));if(existing){ensureDividendTransaction_(existing);invalidateRows_('Transactions');invalidateRows_('DividendPayments');return ok({id:existing.id,duplicate:true,refresh:buildBootstrap_(user).data})}const id=id_('DIV'),tx=id_('TX'),note=clean_(data.note)||('จ่ายเงินปันผล '+owner.name),stamp=now_();try{append_('Transactions',[tx,date,'EXPENSE','CAT-DIVIDEND',account.id,amount,note,id,stamp,user.id,user.name,'CONFIRMED']);append_('DividendPayments',[id,owner.id,owner.name,date,amount,account.id,tx,note,stamp,user.id,user.name]);SpreadsheetApp.flush()}catch(e){deleteRowById_('Transactions',tx);deleteRowById_('DividendPayments',id);throw e}invalidateRows_('Transactions');invalidateRows_('DividendPayments');try{audit_(user,'CREATE','DIVIDEND',id,owner.name+' '+amount)}catch(e){}const refresh=buildBootstrap_(user).data;return ok({id,refresh});}finally{lock.releaseLock()}
}

function ensureDividendTransaction_(dividend){const txSheet=sheet_('Transactions'),txValues=txSheet.getDataRange().getValues(),existingIndex=txValues.findIndex((r,i)=>i>0&&(r[0]===dividend.transactionId||r[7]===dividend.id));let txId=clean_(dividend.transactionId);if(existingIndex>0){txId=txValues[existingIndex][0];if(String(txValues[existingIndex][11]||'CONFIRMED')!=='CONFIRMED')txSheet.getRange(existingIndex+1,12).setValue('CONFIRMED')}else{txId=txId||id_('TX');append_('Transactions',[txId,dateKey_(dividend.date),'EXPENSE','CAT-DIVIDEND',dividend.accountId,num_(dividend.amount),dividend.note||('จ่ายเงินปันผล '+dividend.ownerName),dividend.id,dividend.createdAt||now_(),dividend.createdBy||'SYSTEM',dividend.createdByName||dividend.ownerName||'SYSTEM','CONFIRMED'])}if(txId!==dividend.transactionId){const sh=sheet_('DividendPayments'),values=sh.getDataRange().getValues(),i=values.findIndex((r,n)=>n>0&&r[0]===dividend.id);if(i>0)sh.getRange(i+1,7).setValue(txId)}SpreadsheetApp.flush()}
function ensureDividendTransactionRepair_(){const props=PropertiesService.getScriptProperties(),key='DIVIDEND_TX_REPAIR_20260902';if(props.getProperty(key)==='1')return;const dividends=rows_('DividendPayments');if(!dividends.length){props.setProperty(key,'1');return}dividends.forEach(ensureDividendTransaction_);invalidateRows_('Transactions');invalidateRows_('DividendPayments');props.setProperty(key,'1')}

function checkProductPrice_(user,data){
  const query=clean_(data&&data.name).toLowerCase();if(query.length<2)throw new Error('กรุณาพิมพ์ชื่อสินค้าอย่างน้อย 2 ตัวอักษร');
  const purchases=cachedRows_('Purchases',60).filter(x=>x.status==='CONFIRMED'),purchaseMap=Object.fromEntries(purchases.map(x=>[x.id,x]));
  const records=cachedRows_('PurchaseItems',60).filter(x=>String(x.productName||'').toLowerCase().includes(query)).map(x=>{const p=purchaseMap[x.purchaseId]||{};return{name:x.productName,date:dateKey_(p.date),quantity:num_(x.quantity),unit:x.unit,lineTotal:num_(x.lineTotal),baseUnit:x.baseUnit||x.unit,baseUnitPrice:num_(x.baseUnitPrice),_sort:String(p.createdAt||p.date||'')}}).sort((a,b)=>b._sort.localeCompare(a._sort));
  if(!records.length)return ok({query,records:[],latest:null,change:null});
  const exact=records.filter(x=>String(x.name).toLowerCase()===query),list=(exact.length?exact:records).slice(0,12),latest=list[0],previous=list[1]||null,change=previous&&previous.baseUnitPrice?round_((latest.baseUnitPrice-previous.baseUnitPrice)/previous.baseUnitPrice*100):null;
  return ok({query,records:list,latest,previous,change});
}

function savePriceProduct_(user,data){
  if(!canManage_(user))throw new Error('ไม่มีสิทธิ์จัดการรายการสินค้า');
  const name=clean_(data.name),category=clean_(data.category),baseUnit=clean_(data.baseUnit);if(name.length<2||!category||!baseUnit)throw new Error('กรุณาระบุชื่อ หมวด และหน่วยฐานให้ครบ');
  const sh=sheet_('Products'),rows=sh.getDataRange().getValues(),duplicate=rows.some((r,i)=>i>0&&r[0]!==data.id&&String(r[1]).toLowerCase()===name.toLowerCase()&&truthy_(r[4]));if(duplicate)throw new Error('มีสินค้าชื่อนี้อยู่แล้ว');
  let id=data.id;if(id){const i=rows.findIndex((r,n)=>n>0&&r[0]===id);if(i<0)throw new Error('ไม่พบสินค้าที่ต้องการแก้ไข');sh.getRange(i+1,2,1,4).setValues([[name,category,baseUnit,true]]);audit_(user,'UPDATE','PRODUCT',id,name)}else{id=id_('PRD');append_('Products',[id,name,category,baseUnit,true]);audit_(user,'CREATE','PRODUCT',id,name)}
  SpreadsheetApp.flush();invalidateRows_('Products');const saved=rows_('Products').find(x=>x.id===id);invalidateRows_('Products');return ok({id,product:saved,refresh:bootstrap_(user).data});
}
function savePriceProducts_(user,data){
  if(!canManage_(user))throw new Error('ไม่มีสิทธิ์จัดการรายการสินค้า');const category=clean_(data.category),baseUnit=clean_(data.baseUnit),names=[...new Set((Array.isArray(data.names)?data.names:[]).map(clean_).filter(x=>x.length>=2))];if(!names.length||!category||!baseUnit)throw new Error('กรุณาระบุชื่อสินค้า หมวด และหน่วยฐานให้ครบ');if(names.length>50)throw new Error('เพิ่มได้ไม่เกิน 50 รายการต่อครั้ง');
  const existing=cachedRows_('Products',120),existingNames=new Set(existing.filter(x=>truthy_(x.active)).map(x=>String(x.name).toLowerCase())),newNames=names.filter(x=>!existingNames.has(x.toLowerCase()));if(!newNames.length)throw new Error('รายการที่กรอกมีอยู่ในระบบแล้วทั้งหมด');const rows=newNames.map(name=>['PRD-'+Utilities.getUuid().replace(/-/g,'').slice(0,16),name,category,baseUnit,true]);appendRows_('Products',rows);SpreadsheetApp.flush();invalidateRows_('Products');audit_(user,'CREATE','PRODUCT_BULK','',newNames.length+' รายการ');return ok({created:newNames.length,skipped:names.length-newNames.length,refresh:bootstrap_(user).data});
}
function deletePriceProduct_(user,data){
  if(!canManage_(user))throw new Error('ไม่มีสิทธิ์ลบรายการสินค้า');const sh=sheet_('Products'),rows=sh.getDataRange().getValues(),i=rows.findIndex((r,n)=>n>0&&r[0]===data.id);if(i<0)throw new Error('ไม่พบสินค้า');const name=rows[i][1];sh.getRange(i+1,5).setValue(false);invalidateRows_('Products');audit_(user,'DELETE','PRODUCT',data.id,name);return ok({refresh:bootstrap_(user).data});
}

function compareFinancialPeriods_(user,data){
  if(user.role!=='OWNER')throw new Error('เฉพาะเจ้าของร้านเท่านั้นที่ดูข้อมูลวิเคราะห์ได้');
  const normalize=p=>{const start=validDate_(p&&p.start),end=validDate_(p&&p.end);if(start>end)throw new Error('วันที่เริ่มต้นต้องไม่เกินวันที่สิ้นสุด');return{start,end,label:clean_(p.label)||start+' – '+end}};
  const a=normalize(data&&data.periodA),b=normalize(data&&data.periodB),tx=cachedRows_('Transactions',30).filter(x=>x.status==='CONFIRMED'),cats=cachedRows_('Categories',600),catNames=Object.fromEntries(cats.map(x=>[x.id,x.name]));
  const summarizePeriod=p=>{const list=tx.filter(x=>{const d=dateKey_(x.date);return d>=p.start&&d<=p.end}),tot=totals_(list),groups={};list.filter(x=>x.type==='EXPENSE').forEach(x=>{const n=catNames[x.category]||'รายจ่ายอื่น';groups[n]=(groups[n]||0)+num_(x.amount)});return{...p,...tot,breakdown:Object.keys(groups).map(name=>({name,amount:round_(groups[name])})).sort((x,y)=>y.amount-x.amount).slice(0,8)}};
  const first=summarizePeriod(a),second=summarizePeriod(b),change=(now,old)=>({amount:round_(now-old),pct:old?round_((now-old)/Math.abs(old)*100):null});
  return ok({first,second,changes:{income:change(second.income,first.income),expense:change(second.expense,first.expense),net:change(second.net,first.net)}});
}
function getFinancialPeriod_(user,data){
  if(!canViewFinance_(user)&&user.role!=='MANAGER')throw new Error('ไม่มีสิทธิ์ดูข้อมูลการเงิน');const start=validDate_(data&&data.start),end=validDate_(data&&data.end);if(start>end)throw new Error('วันที่เริ่มต้นต้องไม่เกินวันที่สิ้นสุด');const list=cachedRows_('Transactions',30).filter(x=>x.status==='CONFIRMED').filter(x=>{const d=dateKey_(x.date);return d>=start&&d<=end}),result={...totals_(list),start,end,label:clean_(data.label)||start+' – '+end,breakdown:[]};if(user.role==='MANAGER')result.transactions=transactionViews_(list,cachedRows_('Categories',600),cachedRows_('Accounts',600)).slice(0,500);return ok(result);
}
function getTransactionHistory_(user,data){
  const start=validDate_(data&&data.start),end=validDate_(data&&data.end);if(start>end)throw new Error('วันที่เริ่มต้นต้องไม่เกินวันที่สิ้นสุด');
  const list=dedupeRows_(rows_('Transactions'),['date','type','category','accountId','amount','note','createdBy','status']).filter(x=>{const d=dateKey_(x.date);return d>=start&&d<=end});
  const views=transactionViews_(list,cachedRows_('Categories',600),cachedRows_('Accounts',600));
  return ok((user.role==='STAFF'?views.filter(x=>x.createdBy===user.id):views).slice(0,500));
}

function payrollWeek_(){
  const now=new Date(), day=(now.getDay()+6)%7, start=new Date(now); start.setHours(12,0,0,0); start.setDate(start.getDate()-day);
  const end=new Date(start); end.setDate(end.getDate()+6);
  return {start:Utilities.formatDate(start,Session.getScriptTimeZone(),'yyyy-MM-dd'),end:Utilities.formatDate(end,Session.getScriptTimeZone(),'yyyy-MM-dd')};
}
function buildPayroll_(){
  const week=payrollWeek_(), accounts=cachedRows_('Accounts',600).filter(x=>truthy_(x.active)), advances=dedupeRows_(cachedRows_('WageAdvances',60),['employeeId','weekStart','date','amount','accountId','createdBy']), overtime=dedupeRows_(cachedRows_('WageOvertime',60),['employeeId','weekStart','date','hours','rate','amount','createdBy']), payments=dedupeRows_(activeWagePayments_(cachedRows_('WagePayments',60)),['employeeId','weekStart']), workDays=dedupeRows_(cachedRows_('PartTimeWorkDays',60),['employeeId','date']);
  const accountNames=Object.fromEntries(accounts.map(x=>[x.id,x.name]));
  const allEmployees=cachedRows_('Employees',120),employees=allEmployees.map(e=>{
    if(e.employmentType==='PART_TIME'){
      const own=workDays.filter(x=>x.employeeId===e.id),current=own.filter(x=>dateKey_(x.weekStart)===week.start),unpaid=own.filter(x=>String(x.status||'PENDING')!=='PAID'&&dateKey_(x.date)<=week.end),payment=payments.find(x=>x.employeeId===e.id&&paymentWeekStart_(x)===week.start),due=partTimePayrollDue_(e,week,unpaid,advances,payments),finished=current.length>0&&current.every(x=>String(x.status)==='PAID'),paid=finished||!!payment,remaining=paid?0:due.remaining;
      return{id:e.id,name:e.name,position:e.position||'พนักงาน',employmentType:'PART_TIME',weeklyWage:num_(e.weeklyWage),baseTotal:due.baseTotal,dueWeeks:new Set(unpaid.map(x=>dateKey_(x.weekStart))).size||1,duePeriods:[...new Set(unpaid.map(x=>dateKey_(x.weekStart)))].sort(),overdueWeeks:new Set(unpaid.filter(x=>dateKey_(x.weekStart)<week.start).map(x=>dateKey_(x.weekStart))).size,otTotal:0,totalEarned:due.baseTotal,advanceTotal:due.advanceTotal,remaining,progress:paid?0:(due.baseTotal?round_(remaining/due.baseTotal*100):0),paid,finished,netPaid:payment?num_(payment.netPaid):0,advances:due.advances.map(x=>({id:x.id,date:dateKey_(x.date),amount:num_(x.amount),accountId:x.accountId,accountName:accountNames[x.accountId]||'',note:x.note||''})),overtime:[],workDates:current.map(x=>dateKey_(x.date)).sort(),show:current.length>0||unpaid.length>0||due.advances.length>0};
    }
    const due=employeePayrollDue_(e,week,advances,overtime,payments),payment=payments.find(x=>x.employeeId===e.id&&paymentWeekStart_(x)===week.start),remaining=payment?0:due.remaining;
    const progressBase=due.totalEarned+(due.settledThisWeek||0);
    return{id:e.id,name:e.name,position:e.position||'พนักงาน',employmentType:'FULL_TIME',weeklyWage:num_(e.weeklyWage),baseTotal:due.baseTotal,dueWeeks:due.dueWeeks,duePeriods:due.periods,overdueWeeks:Math.max(0,due.dueWeeks-1),otTotal:due.otTotal,totalEarned:due.totalEarned,advanceTotal:due.advanceTotal,remaining,settledThisWeek:due.settledThisWeek||0,progress:payment?0:(progressBase?round_(remaining/progressBase*100):0),paid:!!payment,netPaid:payment?num_(payment.netPaid):0,advances:due.advances.map(x=>({id:x.id,date:dateKey_(x.date),amount:num_(x.amount),accountId:x.accountId,accountName:accountNames[x.accountId]||'',note:x.note||''})),overtime:due.overtime.map(x=>({id:x.id,date:dateKey_(x.date),hours:num_(x.hours),rate:num_(x.rate),amount:num_(x.amount),note:x.note||''})),show:true};
  }).filter(x=>x.show!==false);
  const employeeNames=Object.fromEntries(allEmployees.map(x=>[x.id,x.name])),month=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM');
  const recentPayments=payments.slice(-30).reverse().map(x=>({...x,weekStart:paymentWeekStart_(x),weekEnd:dateKey_(x.weekEnd),paidAt:dateKey_(x.paidAt),netPaid:num_(x.netPaid),employeeName:employeeNames[x.employeeId]||'พนักงาน',accountName:accountNames[x.accountId]||''}));
  return{week,employees,partTimeEmployees:allEmployees.filter(x=>x.employmentType==='PART_TIME').map(x=>({id:x.id,name:x.name,dailyRate:num_(x.weeklyWage),selectedDates:workDays.filter(w=>w.employeeId===x.id&&dateKey_(w.weekStart)===week.start).map(w=>dateKey_(w.date))})),recentPayments,schedule:payrollMonthSchedule_(month)};
}

function savePartTimeWorkDays_(user,data){if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์กำหนดวันจ้าง');const employee=rows_('Employees').find(x=>x.id===data.employeeId&&x.employmentType==='PART_TIME'),week=payrollWeek_(),dates=[...new Set(Array.isArray(data.dates)?data.dates.map(validDate_):[])].sort();if(!employee)throw new Error('ไม่พบพนักงาน Part-time');if(dates.some(x=>x<week.start||x>week.end))throw new Error('เลือกวันได้เฉพาะสัปดาห์ปัจจุบัน');const lock=LockService.getScriptLock();lock.waitLock(20000);try{const sh=sheet_('PartTimeWorkDays'),values=sh.getDataRange().getValues(),head=values[0],rows=values.slice(1),existing=rows.map((r,i)=>({row:i+2,id:r[0],employeeId:r[1],date:dateKey_(r[2]),weekStart:dateKey_(r[3]),status:String(r[5]||'PENDING')})).filter(x=>x.employeeId===employee.id&&x.weekStart===week.start),selected=new Set(dates);existing.filter(x=>x.status!=='PAID'&&!selected.has(x.date)).reverse().forEach(x=>sh.deleteRow(x.row));const known=new Set(existing.map(x=>x.date));dates.filter(x=>!known.has(x)).forEach(date=>append_('PartTimeWorkDays',[id_('PTD'),employee.id,date,week.start,num_(employee.weeklyWage),'PENDING','',now_(),user.id]));invalidateRows_('PartTimeWorkDays');audit_(user,'UPDATE','PART_TIME_DAYS',employee.id,dates.join(','));return ok({refresh:bootstrap_(user).data})}finally{lock.releaseLock()}}

function employeePayrollDue_(employee,currentWeek,advances,overtime,payments){
  const employeePayments=payments.filter(x=>x.employeeId===employee.id&&paymentWeekStart_(x)<=currentWeek.start).sort((a,b)=>paymentWeekStart_(a).localeCompare(paymentWeekStart_(b))),currentPayment=employeePayments.find(x=>paymentWeekStart_(x)===currentWeek.start);
  if(currentPayment)return{dueWeeks:1,periods:[currentWeek.start],baseTotal:num_(currentPayment.baseWage),advances:[],overtime:[],advanceTotal:num_(currentPayment.advanceTotal),otTotal:0,totalEarned:num_(currentPayment.baseWage)+num_(currentPayment.adjustment),remaining:0,settledThisWeek:num_(currentPayment.baseWage)};
  const latest=employeePayments[employeePayments.length-1],first=latest?addDays_(paymentWeekStart_(latest),7):weekStartFor_(dateKey_(employee.createdAt)||currentWeek.start),periods=[];
  for(let cursor=first;cursor<=currentWeek.start&&periods.length<260;cursor=addDays_(cursor,7))periods.push(cursor);if(!periods.length)periods.push(currentWeek.start);
  const list=advances.filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)>=periods[0]&&dateKey_(x.weekStart)<=currentWeek.start),ot=overtime.filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)>=periods[0]&&dateKey_(x.weekStart)<=currentWeek.start),advanceTotal=round_(list.reduce((s,x)=>s+num_(x.amount),0)),otTotal=round_(ot.reduce((s,x)=>s+num_(x.amount),0)),baseTotal=round_(num_(employee.weeklyWage)*periods.length),totalEarned=round_(baseTotal+otTotal),settledThisWeek=latest&&weekStartFor_(dateKey_(latest.paidAt))===currentWeek.start?num_(latest.baseWage):0;
  return{dueWeeks:periods.length,periods,baseTotal,advances:list,overtime:ot,advanceTotal,otTotal,totalEarned,remaining:round_(Math.max(0,totalEarned-advanceTotal)),settledThisWeek};
}
function partTimePayrollDue_(employee,targetWeek,work,advances,payments){const employeePayments=payments.filter(x=>x.employeeId===employee.id&&paymentWeekStart_(x)<targetWeek.start).sort((a,b)=>paymentWeekStart_(a).localeCompare(paymentWeekStart_(b))),latest=employeePayments[employeePayments.length-1],first=latest?addDays_(paymentWeekStart_(latest),7):weekStartFor_(dateKey_(employee.createdAt)||targetWeek.start),list=advances.filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)>=first&&dateKey_(x.weekStart)<=targetWeek.start),baseTotal=round_(work.reduce((s,x)=>s+num_(x.dailyRate),0)),advanceTotal=round_(list.reduce((s,x)=>s+num_(x.amount),0));return{baseTotal,advances:list,advanceTotal,remaining:round_(Math.max(0,baseTotal-advanceTotal))}}
function paymentWeekStart_(payment){const covered=dateKey_(payment.coveredWeekStart);if(covered)return covered;const recorded=dateKey_(payment.weekStart),paid=dateKey_(payment.paidAt);if(!recorded||!paid)return recorded;const paidWeek=weekStartFor_(paid),offset=Math.round((new Date(paid+'T12:00:00')-new Date(paidWeek+'T12:00:00'))/86400000);return recorded===paidWeek&&offset<=2?addDays_(recorded,-7):recorded}
function activeWagePayments_(payments){return payments.filter(x=>String(x.status||'CONFIRMED')!=='CANCELLED')}
function addDays_(date,days){const d=new Date(date+'T12:00:00');d.setDate(d.getDate()+days);return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd')}
function weekStartFor_(date){const d=new Date(date+'T12:00:00'),day=(d.getDay()+6)%7;d.setDate(d.getDate()-day);return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd')}

function wagePaymentQuote_(employee,targetStart){
  const currentWeek=payrollWeek_(),targetWeek={start:targetStart,end:addDays_(targetStart,6)};
  if(targetStart>currentWeek.start)throw new Error('ไม่สามารถจ่ายรอบในอนาคตได้');
  const payments=dedupeRows_(activeWagePayments_(rows_('WagePayments')),['employeeId','weekStart','coveredWeekStart']);
  if(payments.some(x=>x.employeeId===employee.id&&paymentWeekStart_(x)===targetStart))throw new Error('รอบค่าจ้างนี้จ่ายแล้ว');
  if(employee.employmentType==='PART_TIME'){
    const work=dedupeRows_(rows_('PartTimeWorkDays'),['employeeId','date']).filter(x=>x.employeeId===employee.id&&String(x.status||'PENDING')!=='PAID'&&dateKey_(x.weekStart)===targetStart);
    if(!work.length)throw new Error('ไม่มีวันทำงาน Part-time ที่รอจ่ายในรอบนี้');
    const ownAdvances=rows_('WageAdvances').filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===targetStart),baseTotal=round_(work.reduce((s,x)=>s+num_(x.dailyRate),0)),advanceTotal=round_(ownAdvances.reduce((s,x)=>s+num_(x.amount),0));
    return{targetStart,targetEnd:targetWeek.end,baseTotal,advanceTotal,netBeforeAdjustment:round_(Math.max(0,baseTotal-advanceTotal)),otTotal:0,workDayCount:work.length,workIds:work.map(x=>x.id)};
  }
  const advances=dedupeRows_(rows_('WageAdvances'),['employeeId','weekStart','date','amount','accountId','createdBy']).filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===targetStart),overtime=dedupeRows_(rows_('WageOvertime'),['employeeId','weekStart','date','hours','rate','amount','createdBy']).filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===targetStart),advanceTotal=round_(advances.reduce((s,x)=>s+num_(x.amount),0)),otTotal=round_(overtime.reduce((s,x)=>s+num_(x.amount),0)),baseTotal=num_(employee.weeklyWage),totalEarned=round_(baseTotal+otTotal),due={baseTotal,advances,overtime,advanceTotal,otTotal,totalEarned,remaining:round_(Math.max(0,totalEarned-advanceTotal))};
  return{targetStart,targetEnd:targetWeek.end,baseTotal:due.baseTotal,advanceTotal:due.advanceTotal,otTotal:due.otTotal,netBeforeAdjustment:due.remaining,due};
}
function getWagePaymentQuote_(user,data){
  if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์ดูยอดจ่ายค่าจ้าง');
  const employee=rows_('Employees').find(x=>x.id===data.employeeId),targetStart=validDate_(data.targetWeekStart);
  if(!employee)throw new Error('ไม่พบพนักงาน');
  const quote=wagePaymentQuote_(employee,targetStart);delete quote.due;delete quote.workIds;return ok(quote);
}

function saveEmployee_(user,data){
  if(!canAdminUsers_(user)&&!canManage_(user))throw new Error('เฉพาะเจ้าของร้าน แอดมิน หรือผู้จัดการเท่านั้น');
  const name=clean_(data.name),position=clean_(data.position)||'พนักงานทั่วไป',employmentType=clean_(data.employmentType)||'FULL_TIME',wage=round_(num_(data.weeklyWage));
  if(name.length<2)throw new Error('กรุณาระบุชื่อพนักงานอย่างน้อย 2 ตัวอักษร');
  if(!['FULL_TIME','PART_TIME'].includes(employmentType))throw new Error('ประเภทการจ้างไม่ถูกต้อง');
  const sh=sheet_('Employees'),v=sh.getDataRange().getValues(),head=v[0],pi=head.indexOf('position')+1,ti=head.indexOf('employmentType')+1;
  const isDuplicate=v.some((r,n)=>n>0&&clean_(r[1]).toLowerCase()===name.toLowerCase()&&(!data.id||r[0]!==data.id));
  if(isDuplicate)throw new Error('มีพนักงานชื่อ "'+name+'" อยู่ในระบบแล้ว');
  if(data.id){
    const i=v.findIndex((r,n)=>n>0&&r[0]===data.id);
    if(i<0)throw new Error('ไม่พบพนักงาน');
    sh.getRange(i+1,2).setValue(name);
    sh.getRange(i+1,3).setValue(wage);
    if(pi>0)sh.getRange(i+1,pi).setValue(position);
    if(ti>0)sh.getRange(i+1,ti).setValue(employmentType);
    if(data.note!=null)sh.getRange(i+1,4).setValue(clean_(data.note));
    SpreadsheetApp.flush();
    invalidateRows_('Employees');
    audit_(user,'UPDATE','EMPLOYEE',data.id,name);
  }else{
    const id=id_('EMP');
    append_('Employees',[id,name,wage,clean_(data.note)||'',now_(),user.id,position,employmentType]);
    SpreadsheetApp.flush();
    invalidateRows_('Employees');
    audit_(user,'CREATE','EMPLOYEE',id,name);
  }
  return ok({refresh:bootstrap_(user).data});
}
function deleteEmployee_(user,data){
  if(!canAdminUsers_(user)&&!canManage_(user))throw new Error('ไม่มีสิทธิ์ลบพนักงาน');
  if(rows_('WageAdvances').some(x=>x.employeeId===data.id)||rows_('WageOvertime').some(x=>x.employeeId===data.id)||rows_('WagePayments').some(x=>x.employeeId===data.id))throw new Error('พนักงานนี้มีประวัติค่าจ้างแล้ว จึงไม่สามารถลบได้');
  const sh=sheet_('Employees'),v=sh.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&r[0]===data.id);
  if(i<0)throw new Error('ไม่พบพนักงาน');
  sh.deleteRow(i+1);
  SpreadsheetApp.flush();
  invalidateRows_('Employees');
  audit_(user,'DELETE','EMPLOYEE',data.id,'');
  return ok({refresh:bootstrap_(user).data});
}
function saveWageAdvance_(user,data){
  if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์บันทึกเงินเบิก');const lock=LockService.getScriptLock();lock.waitLock(20000);try{const employee=rows_('Employees').find(x=>x.id===data.employeeId),account=rows_('Accounts').find(x=>x.id===data.accountId&&truthy_(x.active)),amount=round_(num_(data.amount)),date=validDate_(data.date),week=payrollWeek_();if(!employee||!account||amount<=0)throw new Error('ข้อมูลการเบิกเงินไม่ถูกต้อง');if(date<week.start||date>week.end)throw new Error('วันที่เบิกต้องอยู่ในรอบสัปดาห์ปัจจุบัน');if(activeWagePayments_(rows_('WagePayments')).some(x=>x.employeeId===employee.id&&paymentWeekStart_(x)===week.start))throw new Error('พนักงานคนนี้ปิดรอบจ่ายแล้ว');const advances=rows_('WageAdvances'),duplicate=advances.find(x=>x.employeeId===employee.id&&dateKey_(x.date)===date&&num_(x.amount)===amount&&x.accountId===account.id&&x.createdBy===user.id&&recentWrite_(x.createdAt));if(duplicate)return ok({id:duplicate.id,duplicate:true,refresh:bootstrap_(user).data});const used=advances.filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===week.start).reduce((s,x)=>s+num_(x.amount),0),ot=rows_('WageOvertime').filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===week.start).reduce((s,x)=>s+num_(x.amount),0),available=employee.employmentType==='PART_TIME'?wagePaymentQuote_(employee,week.start).netBeforeAdjustment:num_(employee.weeklyWage)+ot-used;if(amount>available)throw new Error('ยอดเบิกมากกว่ายอดค่าจ้างที่เหลือ');const id=id_('ADV'),tx=id_('TX'),note=clean_(data.note)||('เบิกค่าจ้างล่วงหน้า '+employee.name),stamp=now_();try{append_('Transactions',[tx,date,'EXPENSE','CAT-WAGE',account.id,amount,note,id,stamp,user.id,user.name,'CONFIRMED']);append_('WageAdvances',[id,employee.id,week.start,date,amount,account.id,tx,note,stamp,user.id,user.name]);SpreadsheetApp.flush()}catch(e){deleteRowById_('Transactions',tx);deleteRowById_('WageAdvances',id);throw e}invalidateRows_('Transactions');invalidateRows_('WageAdvances');try{audit_(user,'CREATE','WAGE_ADVANCE',id,employee.name+' '+amount)}catch(e){}return ok({id,refresh:bootstrap_(user).data})}finally{lock.releaseLock()}
}
function saveWageOvertime_(user,data){
  if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์บันทึก OT');const lock=LockService.getScriptLock();lock.waitLock(20000);try{const employee=rows_('Employees').find(x=>x.id===data.employeeId),date=validDate_(data.date),hours=round_(num_(data.hours)),rate=round_(num_(data.rate)),week=payrollWeek_();if(!employee||hours<=0||rate<=0)throw new Error('กรุณาระบุวันที่ ชั่วโมง และค่า OT ให้ถูกต้อง');if(date<week.start||date>week.end)throw new Error('วันที่ OT ต้องอยู่ในรอบสัปดาห์ปัจจุบัน');if(activeWagePayments_(rows_('WagePayments')).some(x=>x.employeeId===employee.id&&paymentWeekStart_(x)===week.start))throw new Error('พนักงานคนนี้ปิดรอบจ่ายแล้ว');const overtime=rows_('WageOvertime'),duplicate=overtime.find(x=>x.employeeId===employee.id&&dateKey_(x.date)===date&&num_(x.hours)===hours&&num_(x.rate)===rate&&x.createdBy===user.id&&recentWrite_(x.createdAt));if(duplicate)return ok({id:duplicate.id,duplicate:true,refresh:bootstrap_(user).data});const amount=round_(hours*rate),id=id_('OT');append_('WageOvertime',[id,employee.id,week.start,date,hours,rate,amount,clean_(data.note),now_(),user.id,user.name]);invalidateRows_('WageOvertime');audit_(user,'CREATE','WAGE_OT',id,employee.name+' '+amount);return ok({id,refresh:bootstrap_(user).data})}finally{lock.releaseLock()}
}
function payWeeklyWage_(user,data){
  if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์จ่ายค่าจ้าง');
  const lock=LockService.getScriptLock();lock.waitLock(20000);try{
    const employee=rows_('Employees').find(x=>x.id===data.employeeId),account=rows_('Accounts').find(x=>x.id===data.accountId&&truthy_(x.active)),targetStart=validDate_(data.targetWeekStart||payrollWeek_().start),paidDate=validDate_(data.date);if(!employee)throw new Error('ไม่พบพนักงาน');if(!account)throw new Error('กรุณาเลือกบัญชีที่ใช้จ่าย');
    const existing=activeWagePayments_(rows_('WagePayments')).find(x=>x.employeeId===employee.id&&paymentWeekStart_(x)===targetStart);if(existing)return ok({id:existing.id,paidWeek:targetStart,duplicate:true,refresh:data.skipRefresh?null:safeBootstrap_(user)});
    if(employee.employmentType==='PART_TIME')return payPartTimeWageLocked_(user,data,employee);
    const quote=wagePaymentQuote_(employee,targetStart),due=quote.due,adjustment=round_(num_(data.adjustment)),net=round_(quote.netBeforeAdjustment+adjustment);if(net<0)throw new Error('ยอดสุทธิติดลบ กรุณาตรวจสอบเงินเพิ่ม/หัก');
    const id=id_('PAY'),tx=net>0?id_('TX'):'',note=clean_(data.note)||('จ่ายค่าจ้าง '+employee.name+' ถึงรอบ '+quote.targetEnd),stamp=now_();
    try{if(net>0)append_('Transactions',[tx,paidDate,'EXPENSE','CAT-WAGE',account.id,net,note,id,stamp,user.id,user.name,'CONFIRMED']);append_('WagePayments',[id,employee.id,targetStart,quote.targetEnd,due.baseTotal,due.advanceTotal,round_(adjustment+due.otTotal),net,account.id,tx,note,stamp,user.id,user.name,targetStart,'CONFIRMED']);}
    catch(e){if(tx)deleteRowById_('Transactions',tx);deleteRowById_('WagePayments',id);throw e}
    invalidateRows_('Transactions');invalidateRows_('WagePayments');audit_(user,'CREATE','WAGE_PAYMENT',id,employee.name+' '+net+' ถึง '+quote.targetEnd);return ok({id,paidWeek:targetStart,refresh:data.skipRefresh?null:safeBootstrap_(user)});
  }finally{lock.releaseLock()}
}

function payPartTimeWageLocked_(user,data,employee){const account=rows_('Accounts').find(x=>x.id===data.accountId&&truthy_(x.active)),targetStart=validDate_(data.targetWeekStart||payrollWeek_().start),paidDate=validDate_(data.date);if(!account)throw new Error('กรุณาเลือกบัญชีที่ใช้จ่าย');const quote=wagePaymentQuote_(employee,targetStart),adjustment=round_(num_(data.adjustment)),net=round_(quote.netBeforeAdjustment+adjustment);if(net<0)throw new Error('ยอดสุทธิติดลบ');const id=id_('PAY'),tx=net>0?id_('TX'):'',stamp=now_(),note=clean_(data.note)||('จ่าย Part-time '+employee.name+' '+quote.workDayCount+' วัน ถึง '+quote.targetEnd),sh=sheet_('PartTimeWorkDays'),values=sh.getDataRange().getValues(),ids=new Set(quote.workIds),updates=[];values.slice(1).forEach((r,i)=>{if(ids.has(r[0]))updates.push({row:i+2,status:r[5],paymentId:r[6]})});try{if(net>0)append_('Transactions',[tx,paidDate,'EXPENSE','CAT-WAGE',account.id,net,note,id,stamp,user.id,user.name,'CONFIRMED']);append_('WagePayments',[id,employee.id,targetStart,quote.targetEnd,quote.baseTotal,quote.advanceTotal,adjustment,net,account.id,tx,note,stamp,user.id,user.name,targetStart,'CONFIRMED']);updates.forEach(x=>sh.getRange(x.row,6,1,2).setValues([['PAID',id]]));}catch(e){updates.forEach(x=>sh.getRange(x.row,6,1,2).setValues([[x.status,x.paymentId]]));if(tx)deleteRowById_('Transactions',tx);deleteRowById_('WagePayments',id);throw e}invalidateRows_('Transactions');invalidateRows_('WagePayments');invalidateRows_('PartTimeWorkDays');audit_(user,'CREATE','PART_TIME_PAYMENT',id,employee.name+' '+quote.workDayCount+' วัน '+net);return ok({id,paidWeek:targetStart,refresh:data.skipRefresh?null:safeBootstrap_(user)})}

function resetPayrollWeekCore_(week,user){
  const paymentSheet=sheet_('WagePayments'),paymentValues=paymentSheet.getDataRange().getValues(),paymentHead=paymentValues[0],statusCol=paymentHead.indexOf('status')+1,transactionSheet=sheet_('Transactions'),transactionValues=transactionSheet.getDataRange().getValues(),workSheet=sheet_('PartTimeWorkDays'),workValues=workSheet.getDataRange().getValues(),targets=[];
  paymentValues.slice(1).forEach((row,i)=>{const payment=Object.fromEntries(paymentHead.map((h,j)=>[h,row[j]]));if(String(payment.status||'CONFIRMED')!=='CANCELLED'&&paymentWeekStart_(payment)===week.start)targets.push({row:i+2,payment})});
  targets.forEach(x=>{paymentSheet.getRange(x.row,statusCol).setValue('CANCELLED');if(x.payment.transactionId){const ti=transactionValues.findIndex((r,i)=>i>0&&r[0]===x.payment.transactionId);if(ti>0)transactionSheet.getRange(ti+1,12).setValue('CANCELLED')}workValues.slice(1).forEach((r,i)=>{if(r[6]===x.payment.id)workSheet.getRange(i+2,6,1,2).setValues([['PENDING','']])})});
  if(targets.length){SpreadsheetApp.flush();invalidateRows_('WagePayments');invalidateRows_('Transactions');invalidateRows_('PartTimeWorkDays');audit_(user,'RESET','PAYROLL_WEEK',week.start,'ยกเลิกรอบจ่ายปัจจุบัน '+targets.length+' รายการ')}
  return targets.length;
}
function resetCurrentPayroll_(user,data){if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์รีเซตรอบค่าจ้าง');if(!data||data.confirm!==true)throw new Error('กรุณายืนยันการรีเซต');const lock=LockService.getScriptLock();lock.waitLock(20000);try{const count=resetPayrollWeekCore_(payrollWeek_(),user);return ok({count,refresh:safeBootstrap_(user)})}finally{lock.releaseLock()}}
function ensureCurrentPayrollPendingMigration_(){const props=PropertiesService.getScriptProperties(),key='PAYROLL_CURRENT_PENDING_20260831';if(props.getProperty(key)==='1')return;const owner=rows_('Users').find(x=>x.role==='OWNER'&&truthy_(x.active))||{id:'SYSTEM',name:'SYSTEM'};resetPayrollWeekCore_(payrollWeek_(),owner);props.setProperty(key,'1')}

function validPayrollMonth_(value){const month=String(value||'');if(!/^\d{4}-(0[1-9]|1[0-2])$/.test(month))throw new Error('เดือนค่าจ้างไม่ถูกต้อง');return month}
function payrollMonthSchedule_(value){const month=validPayrollMonth_(value),first=month+'-01',nextMonth=Utilities.formatDate(new Date(Number(month.slice(0,4)),Number(month.slice(5,7)),1,12),Session.getScriptTimeZone(),'yyyy-MM-dd'),last=addDays_(nextMonth,-1),saved={};cachedRows_('PayrollSchedule',600).filter(x=>x.month===month).forEach(x=>saved[dateKey_(x.weekStart)]=x);const weeks=[];for(let start=weekStartFor_(first);addDays_(start,6)<=last;start=addDays_(start,7)){const row=saved[start];weeks.push({id:row&&row.id||'',weekStart:start,weekEnd:addDays_(start,6),payDate:row&&dateKey_(row.payDate)||addDays_(start,6),active:row?truthy_(row.active):true})}return{month,weeks}}
function getPayrollSchedule_(user,data){if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์ดูตารางจ่ายค่าจ้าง');return ok(payrollMonthSchedule_(data&&data.month))}
function getPayrollWeekStatus_(user,data){if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์ดูสถานะค่าจ้าง');const weekStart=validDate_(data&&data.weekStart);if(weekStartFor_(weekStart)!==weekStart)throw new Error('รอบสัปดาห์ไม่ถูกต้อง');const weekEnd=addDays_(weekStart,6),employees=cachedRows_('Employees',120),payments=activeWagePayments_(cachedRows_('WagePayments',30)),advances=cachedRows_('WageAdvances',30),overtime=cachedRows_('WageOvertime',30),work=cachedRows_('PartTimeWorkDays',30);const items=employees.map(e=>{const payment=payments.find(x=>x.employeeId===e.id&&paymentWeekStart_(x)===weekStart),ownAdv=advances.filter(x=>x.employeeId===e.id&&dateKey_(x.weekStart)===weekStart),ownOt=overtime.filter(x=>x.employeeId===e.id&&dateKey_(x.weekStart)===weekStart),advanceTotal=round_(ownAdv.reduce((s,x)=>s+num_(x.amount),0)),otTotal=round_(ownOt.reduce((s,x)=>s+num_(x.amount),0)),days=work.filter(x=>x.employeeId===e.id&&dateKey_(x.weekStart)===weekStart),baseTotal=e.employmentType==='PART_TIME'?round_(days.reduce((s,x)=>s+num_(x.dailyRate),0)):num_(e.weeklyWage),earned=round_(baseTotal+otTotal),net=payment?num_(payment.netPaid):round_(Math.max(0,earned-advanceTotal));return{id:e.id,name:e.name,position:e.position||'พนักงาน',employmentType:e.employmentType||'FULL_TIME',weekStart,weekEnd,baseTotal,otTotal,advanceTotal,net,paid:!!payment,paymentId:payment&&payment.id||'',paidAt:payment&&dateKey_(payment.paidAt)||'',workDayCount:days.length,progress:payment?100:(earned?Math.max(0,Math.min(100,round_(net/earned*100))):0)}});return ok({weekStart,weekEnd,items,summary:{total:items.length,paid:items.filter(x=>x.paid).length,pending:items.filter(x=>!x.paid).length,net:round_(items.filter(x=>!x.paid).reduce((s,x)=>s+x.net,0))}})}
function savePayrollSchedule_(user,data){if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์กำหนดรอบจ่ายค่าจ้าง');const month=validPayrollMonth_(data&&data.month),input=Array.isArray(data.weeks)?data.weeks:[],valid=payrollMonthSchedule_(month),allowed=new Set(valid.weeks.map(x=>x.weekStart)),lock=LockService.getScriptLock();lock.waitLock(20000);try{const sh=sheet_('PayrollSchedule'),values=sh.getDataRange().getValues(),head=values[0],existing={};values.slice(1).forEach((r,i)=>{if(r[1]===month)existing[dateKey_(r[2])]=i+2});valid.weeks.forEach(def=>{const item=input.find(x=>x.weekStart===def.weekStart)||def,payDate=validDate_(item.payDate||def.payDate);if(!allowed.has(def.weekStart))throw new Error('รอบสัปดาห์ไม่ถูกต้อง');const row=[existing[def.weekStart]?values[existing[def.weekStart]-1][0]:id_('SCH'),month,def.weekStart,def.weekEnd,payDate,item.active!==false,existing[def.weekStart]?values[existing[def.weekStart]-1][6]:now_(),user.id,now_()];if(existing[def.weekStart])sh.getRange(existing[def.weekStart],1,1,head.length).setValues([row]);else sh.appendRow(row)});invalidateRows_('PayrollSchedule');audit_(user,'UPDATE','PAYROLL_SCHEDULE',month,'กำหนด '+valid.weeks.length+' รอบ');return ok({schedule:payrollMonthSchedule_(month)})}finally{lock.releaseLock()}}
function cancelWagePayment_(user,data){if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์ยกเลิกรายการจ่าย');const id=clean_(data&&data.id),lock=LockService.getScriptLock();lock.waitLock(20000);try{const paymentSheet=sheet_('WagePayments'),values=paymentSheet.getDataRange().getValues(),head=values[0],index=values.findIndex((r,i)=>i>0&&r[0]===id),statusCol=head.indexOf('status')+1;if(index<1)return ok({duplicate:true,missing:true});if(statusCol<1)throw new Error('โครงสร้างสถานะค่าจ้างไม่สมบูรณ์');const payment=Object.fromEntries(head.map((h,j)=>[h,values[index][j]])),weekStart=paymentWeekStart_(payment);if(String(payment.status||'CONFIRMED')==='CANCELLED')return ok({duplicate:true,weekStart});paymentSheet.getRange(index+1,statusCol).setValue('CANCELLED');if(payment.transactionId){const tx=sheet_('Transactions'),tv=tx.getDataRange().getValues(),ti=tv.findIndex((r,i)=>i>0&&r[0]===payment.transactionId);if(ti>0)tx.getRange(ti+1,12).setValue('CANCELLED')}const work=sheet_('PartTimeWorkDays'),wv=work.getDataRange().getValues();wv.slice(1).forEach((r,i)=>{if(r[6]===id)work.getRange(i+2,6,1,2).setValues([['PENDING','']])});SpreadsheetApp.flush();invalidateRows_('WagePayments');invalidateRows_('Transactions');invalidateRows_('PartTimeWorkDays');audit_(user,'CANCEL','WAGE_PAYMENT',id,'คืนยอดและเปิดรอบ '+weekStart);return ok({weekStart})}finally{lock.releaseLock()}}
function resetPayrollWeek_(user,data){if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์รีเซตรอบค่าจ้าง');if(!data||data.confirm!==true)throw new Error('กรุณายืนยันการรีเซต');const start=validDate_(data.weekStart);if(weekStartFor_(start)!==start)throw new Error('รอบสัปดาห์ไม่ถูกต้อง');const lock=LockService.getScriptLock();lock.waitLock(20000);try{const count=resetPayrollWeekCore_({start,end:addDays_(start,6)},user);return ok({count,weekStart:start})}finally{lock.releaseLock()}}
function ensureManualPayrollMigration_(){const props=PropertiesService.getScriptProperties(),key='PAYROLL_MANUAL_ONLY_20260831';if(props.getProperty(key)==='1')return;props.setProperty('PAYROLL_AUTO_ENABLED','0');try{ScriptApp.getProjectTriggers().filter(t=>t.getHandlerFunction()==='autoPayPayrollTrigger_').forEach(t=>ScriptApp.deleteTrigger(t))}catch(e){}props.setProperty(key,'1')}

function savePurchase_(user, data) {
  ensureSchema_();
  if (!data || !Array.isArray(data.items) || !data.items.length) throw new Error('กรุณาเพิ่มรายการอย่างน้อย 1 รายการ');
  const account = rows_('Accounts').find(x => x.id === data.accountId && truthy_(x.active));
  if (!account) throw new Error('กรุณาเลือกบัญชีที่ใช้จ่าย');
  const date = validDate_(data.date);
  const items = data.items.map((x,i) => normalizeItem_(x,i));
  const total = round_(items.reduce((s,x) => s+x.lineTotal,0));
  if (total <= 0) throw new Error('ยอดรวมต้องมากกว่า 0');
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const purchaseRows=rows_('Purchases'),purchaseItemRows=rows_('PurchaseItems'),signature=purchaseSignature_(items),recentPurchase=purchaseRows.find(x=>dateKey_(x.date)===date&&x.accountId===account.id&&num_(x.total)===total&&clean_(x.supplier)===clean_(data.supplier)&&clean_(x.note)===clean_(data.note)&&x.createdBy===user.id&&x.status==='CONFIRMED'&&recentWrite_(x.createdAt)&&purchaseSignature_(purchaseItemRows.filter(i=>i.purchaseId===x.id))===signature);
    if(recentPurchase)return ok({id:recentPurchase.id,total,itemCount:items.length,duplicate:true});
    const purchaseId = id_('BUY'),transactionId=id_('TX');
    const productRows = cachedRows_('Products',120);
    const productMap = Object.fromEntries(productRows.map(p => [String(p.name).toLowerCase(), p]));
    const history = cachedRows_('PurchaseItems',60);
    const latestPrice = {};
    history.forEach(r => latestPrice[r.productId] = num_(r.baseUnitPrice));
    const newProducts = [], categoryUpdates = [];
    const itemRows = [];
    items.forEach(item => {
      const key = item.name.toLowerCase();
      let product = productMap[key];
      if (!product) {
        product = { id:id_('PRD'), name:item.name, category:item.category, baseUnit:item.baseUnit, active:true };
        productMap[key] = product;
        newProducts.push([product.id,product.name,product.category,product.baseUnit,true]);
      } else if (product.category !== item.category) { categoryUpdates.push({product,oldCategory:product.category}); product.category=item.category; }
      const previous = latestPrice[product.id] || 0;
      const basePrice = round_(item.lineTotal/item.baseQuantity);
      const change = previous ? round_((basePrice-previous)/previous*100) : '';
      itemRows.push([id_('ITM'),purchaseId,product.id,product.name,item.quantity,item.unit,item.unitFactor,item.baseQuantity,item.lineTotal,basePrice,previous || '',change,item.category]);
    });
    try{
      append_('Purchases', [purchaseId,date,clean_(data.supplier),account.id,total,clean_(data.note),now_(),user.id,user.name,'CONFIRMED']);
      appendRows_('Products', newProducts);
      appendRows_('PurchaseItems', itemRows);
      append_('Transactions', [transactionId,date,'EXPENSE','CAT-RAW',account.id,total,clean_(data.note) || 'ซื้อวัตถุดิบ',purchaseId,now_(),user.id,user.name,'CONFIRMED']);
      if(categoryUpdates.length){ const sh=sheet_('Products'); categoryUpdates.forEach(x=>{ const idx=productRows.findIndex(p=>p.id===x.product.id); if(idx>=0) sh.getRange(idx+2,3).setValue(x.product.category); }); }
      SpreadsheetApp.flush();
    }catch(e){
      if(categoryUpdates.length){const sh=sheet_('Products');categoryUpdates.forEach(x=>{const idx=productRows.findIndex(p=>p.id===x.product.id);if(idx>=0)sh.getRange(idx+2,3).setValue(x.oldCategory)})}deleteRowById_('Transactions',transactionId);deleteRowsByField_('PurchaseItems','purchaseId',purchaseId);newProducts.forEach(p=>deleteRowById_('Products',p[0]));deleteRowById_('Purchases',purchaseId);throw e;
    }
    invalidateRows_('Products');
    invalidateRows_('Purchases'); invalidateRows_('PurchaseItems');
    invalidateRows_('Transactions');
    try{audit_(user,'CREATE','PURCHASE',purchaseId,items.length+' รายการ รวม '+total+' บาท')}catch(e){}
    return ok({ id:purchaseId, total, itemCount:items.length, refresh: bootstrap_(user).data });
  } finally { lock.releaseLock(); }
}

function purchaseSignature_(items){return items.map(x=>[clean_(x.productName||x.name).toLowerCase(),round_(num_(x.quantity)),clean_(x.unit),round_(num_(x.lineTotal))].join('|')).sort().join('\n')}

function saveTransaction_(user, data) {
  if (!data || !['INCOME','EXPENSE'].includes(data.type)) throw new Error('ประเภทรายการไม่ถูกต้อง');
  const amount = num_(data.amount); if (amount <= 0) throw new Error('จำนวนเงินต้องมากกว่า 0');
  const account = rows_('Accounts').find(x => x.id === data.accountId && truthy_(x.active));
  const category = rows_('Categories').find(x => x.id === data.category && x.type === data.type && (data.id || truthy_(x.active)));
  if (!account || !category) throw new Error('กรุณาเลือกบัญชีและหมวดหมู่ให้ถูกต้อง');
  const date=validDate_(data.date);
  if(data.id){
    const editLock=LockService.getScriptLock();editLock.waitLock(20000);try{const sh=sheet_('Transactions'), values=sh.getDataRange().getValues(), head=values[0], idx=values.findIndex((r,i)=>i>0 && r[0]===data.id);
      if(idx<0) throw new Error('ไม่พบรายการที่ต้องการแก้ไข');
      const old=Object.fromEntries(head.map((h,i)=>[h,values[idx][i]])); if(old.status==='CANCELLED') throw new Error('รายการนี้ถูกยกเลิกแล้ว'); if(!canManage_(user)&&old.createdBy!==user.id) throw new Error('แก้ไขได้เฉพาะรายการที่ตนเองบันทึก');
      if(old.referenceId){sh.getRange(idx+1,2).setValue(date); sh.getRange(idx+1,5).setValue(account.id); sh.getRange(idx+1,7).setValue(clean_(data.note));updatePurchaseHeader_(old.referenceId,date,account.id,clean_(data.note));}
      else sh.getRange(idx+1,2,1,6).setValues([[date,data.type,category.id,account.id,round_(amount),clean_(data.note)]]);
      SpreadsheetApp.flush();invalidateRows_('Transactions');try{audit_(user,'UPDATE','TRANSACTION',data.id,'แก้ไขรายการ')}catch(e){}return ok({id:data.id, refresh: bootstrap_(user).data});
    }finally{editLock.releaseLock()}
  }
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const note=clean_(data.note),rounded=round_(amount),recent=rows_('Transactions').find(x=>dateKey_(x.date)===date&&x.type===data.type&&x.category===category.id&&x.accountId===account.id&&num_(x.amount)===rounded&&clean_(x.note)===note&&x.createdBy===user.id&&x.status==='CONFIRMED'&&recentWrite_(x.createdAt));
    if(recent)return ok({id:recent.id,duplicate:true, refresh: bootstrap_(user).data});
    const txId = id_('TX'); append_('Transactions',[txId,date,data.type,category.id,account.id,rounded,note,'',now_(),user.id,user.name,'CONFIRMED']);
    SpreadsheetApp.flush();invalidateRows_('Transactions');try{audit_(user,'CREATE','TRANSACTION',txId,data.type+' '+amount+' บาท')}catch(e){}return ok({id:txId, refresh: bootstrap_(user).data});
  }finally{lock.releaseLock()}
}

function saveBatchTransactions_(user, data) {
  const list = Array.isArray(data) ? data : (data && Array.isArray(data.transactions) ? data.transactions : []);
  if (!list.length) {
    throw new Error('กรุณาระบุรายการอย่างน้อย 1 รายการ');
  }
  const defaultAccount = rows_('Accounts').find(x => truthy_(x.active));
  const categories = rows_('Categories');
  const accounts = rows_('Accounts');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const stamp = now_();
    const rowsToAdd = [];
    const defaultDate = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
    let totalIncome = 0;
    let totalExpense = 0;

    list.forEach((tx, idx) => {
      const type = tx.type === 'INCOME' ? 'INCOME' : 'EXPENSE';
      const amount = round_(num_(tx.amount));
      if (amount <= 0) throw new Error('รายการที่ ' + (idx + 1) + ' (' + (tx.note || tx.name || '') + ') จำนวนเงินไม่ถูกต้อง');
      const txDate = validDate_(tx.date || (data && data.date) || defaultDate);
      const accountId = tx.accountId && accounts.some(a => a.id === tx.accountId) ? tx.accountId : (defaultAccount ? defaultAccount.id : 'ACC-CASH');
      let catId = tx.categoryId || tx.category;
      let cat = categories.find(c => c.id === catId);
      if (!cat) {
        cat = categories.find(c => c.type === type && truthy_(c.active));
        catId = cat ? cat.id : (type === 'INCOME' ? 'CAT-SALES' : 'CAT-RAW');
      }
      const note = clean_(tx.note || tx.name || '');
      const txId = id_('TX');
      rowsToAdd.push([txId, txDate, type, catId, accountId, amount, note, '', stamp, user.id, user.name, 'CONFIRMED']);
      if (type === 'INCOME') totalIncome += amount;
      else totalExpense += amount;
    });

    appendRows_('Transactions', rowsToAdd);
    SpreadsheetApp.flush();
    invalidateRows_('Transactions');
    try {
      audit_(user, 'CREATE', 'BATCH_TRANSACTIONS', rowsToAdd[0][0], 'บันทึกด่วน ' + rowsToAdd.length + ' รายการ รวม ' + (totalIncome + totalExpense) + ' บาท');
    } catch(e) {}
    return ok({ count: rowsToAdd.length, totalIncome, totalExpense, refresh: bootstrap_(user).data });
  } finally {
    lock.releaseLock();
  }
}

function deleteTransaction_(user,data){
  const lock=LockService.getScriptLock();lock.waitLock(20000);try{const sh=sheet_('Transactions'), values=sh.getDataRange().getValues(), head=values[0], idx=values.findIndex((r,i)=>i>0 && r[0]===data.id);
  if(idx<0) throw new Error('ไม่พบรายการ'); const old=Object.fromEntries(head.map((h,i)=>[h,values[idx][i]]));
  if(old.status==='CANCELLED') throw new Error('รายการนี้ถูกยกเลิกแล้ว'); if(!canManage_(user)&&old.createdBy!==user.id) throw new Error('ยกเลิกได้เฉพาะรายการที่ตนเองบันทึก'); sh.getRange(idx+1,12).setValue('CANCELLED');
  if(old.referenceId){ const ps=sheet_('Purchases'), pv=ps.getDataRange().getValues(), pi=pv.findIndex((r,i)=>i>0&&r[0]===old.referenceId); if(pi>0) ps.getRange(pi+1,10).setValue('CANCELLED'); invalidateRows_('Purchases'); }
  SpreadsheetApp.flush(); invalidateRows_('Transactions'); try{audit_(user,'CANCEL','TRANSACTION',data.id,clean_(data.reason)||'ยกเลิกรายการ')}catch(e){} return ok({refresh:bootstrap_(user).data});}finally{lock.releaseLock()}
}

function updatePurchaseHeader_(id,date,accountId,note){ const sh=sheet_('Purchases'), v=sh.getDataRange().getValues(), i=v.findIndex((r,n)=>n>0&&r[0]===id); if(i>0){sh.getRange(i+1,2).setValue(date);sh.getRange(i+1,4).setValue(accountId);sh.getRange(i+1,6).setValue(note);invalidateRows_('Purchases');} }

function createUser_(user, data) {
  if (!canAdminUsers_(user)) throw new Error('เฉพาะเจ้าของร้านหรือแอดมินเท่านั้น');
  validateUserInput_(data);
  const lock=LockService.getScriptLock();lock.waitLock(20000);
  try{
    const username = clean_(data.username).toLowerCase();
    if (rows_('Users').some(x => String(x.username).toLowerCase() === username)) throw new Error('ชื่อผู้ใช้นี้มีแล้ว');
    const role = ['STAFF','MANAGER','ADMIN','OWNER'].includes(data.role) ? data.role : 'STAFF';
    if (['OWNER','ADMIN'].includes(role) && user.role !== 'OWNER') throw new Error('เฉพาะเจ้าของร้านที่สร้างบัญชีเจ้าของหรือแอดมินได้');
    const id = id_('USR');
    append_('Users',[id,username,hash_(data.password),clean_(data.name),role,true,now_(),user.id]);
    SpreadsheetApp.flush();invalidateRows_('Users');
    // The account is already committed. Audit/bootstrap failures must not make
    // the client retry and accidentally create a duplicate account.
    try{audit_(user,'CREATE','USER',id,username+' / '+role)}catch(e){}
    return ok({id});
  }finally{lock.releaseLock()}
}

function updateUser_(user,data){
  if(!canAdminUsers_(user)) throw new Error('ไม่มีสิทธิ์จัดการผู้ใช้');
  const sh=sheet_('Users'), values=sh.getDataRange().getValues(), head=values[0], idx=values.findIndex((r,i)=>i>0&&r[0]===data.id);
  if(idx<0) throw new Error('ไม่พบบัญชีผู้ใช้'); const old=Object.fromEntries(head.map((h,i)=>[h,values[idx][i]]));
  if(user.role!=='OWNER' && ['OWNER','ADMIN'].includes(old.role)) throw new Error('แอดมินไม่สามารถแก้ไขเจ้าของหรือแอดมินอื่นได้');
  const name=clean_(data.name), username=clean_(data.username).toLowerCase(), role=['STAFF','MANAGER','ADMIN','OWNER'].includes(data.role)?data.role:'STAFF';
  if(name.length<2||!/^[a-zA-Z0-9._-]{3,30}$/.test(username)) throw new Error('ชื่อหรือชื่อผู้ใช้ไม่ถูกต้อง');
  if(values.some((r,i)=>i>0&&r[0]!==data.id&&String(r[1]).toLowerCase()===username)) throw new Error('ชื่อผู้ใช้นี้มีแล้ว');
  if(user.role!=='OWNER'&&['OWNER','ADMIN'].includes(role)) throw new Error('เฉพาะเจ้าของร้านที่กำหนดบทบาทนี้ได้');
  if(data.id===user.id&&role!==user.role) throw new Error('ไม่สามารถเปลี่ยนบทบาทของตนเอง');
  sh.getRange(idx+1,2).setValue(username); if(clean_(data.password)){ if(String(data.password).length<6) throw new Error('รหัสผ่านใหม่ต้องมีอย่างน้อย 6 ตัว'); sh.getRange(idx+1,3).setValue(hash_(data.password)); }
  sh.getRange(idx+1,4).setValue(name);sh.getRange(idx+1,5).setValue(role);SpreadsheetApp.flush();invalidateRows_('Users');try{audit_(user,'UPDATE','USER',data.id,username+' / '+role)}catch(e){}return ok({id:data.id});
}

function toggleUser_(user, data) {
  if (!canAdminUsers_(user)) throw new Error('ไม่มีสิทธิ์จัดการผู้ใช้');
  if (data.id === user.id) throw new Error('ไม่สามารถปิดบัญชีของตนเอง');
  const sh=sheet_('Users'), values=sh.getDataRange().getValues();
  const idx=values.findIndex((r,i)=>i>0 && r[0]===data.id); if(idx<0) throw new Error('ไม่พบบัญชี');
  const target=Object.fromEntries(values[0].map((h,i)=>[h,values[idx][i]])); if(user.role!=='OWNER'&&['OWNER','ADMIN'].includes(target.role)) throw new Error('แอดมินไม่สามารถปิดบัญชีเจ้าของหรือแอดมินอื่นได้');
  sh.getRange(idx+1,6).setValue(!!data.active);
  invalidateRows_('Users');
  audit_(user,'UPDATE','USER',data.id,'active='+!!data.active);
  return ok({refresh:bootstrap_(user).data});
}

function saveCategory_(user, data) {
  if (!canManage_(user)) throw new Error('เฉพาะผู้จัดการหรือเจ้าของร้านเท่านั้น');
  if(data.id==='CAT-DIVIDEND') throw new Error('หมวดเงินปันผลเป็นหมวดระบบ จึงไม่สามารถแก้ไขได้');
  const name=clean_(data.name), type=clean_(data.type);
  if(name.length<2) throw new Error('กรุณาระบุชื่อหมวดอย่างน้อย 2 ตัวอักษร');
  if(!['INCOME','EXPENSE'].includes(type)) throw new Error('ประเภทหมวดไม่ถูกต้อง');
  const sh=sheet_('Categories'), values=sh.getDataRange().getValues();
  if(values.some((r,i)=>i>0 && String(r[1]).toLowerCase()===name.toLowerCase() && r[0]!==data.id && truthy_(r[3]))) throw new Error('มีชื่อหมวดนี้อยู่แล้ว');
  if(data.id){ const idx=values.findIndex((r,i)=>i>0 && r[0]===data.id); if(idx<0) throw new Error('ไม่พบหมวดที่ต้องการแก้ไข'); sh.getRange(idx+1,2,1,2).setValues([[name,type]]); audit_(user,'UPDATE','CATEGORY',data.id,name); }
  else { const id=id_('CAT'), order=Math.max(0,...values.slice(1).map(r=>num_(r[4])))+1; append_('Categories',[id,name,type,true,order]); audit_(user,'CREATE','CATEGORY',id,name); }
  invalidateRows_('Categories'); return ok({refresh:bootstrap_(user).data});
}

function toggleCategory_(user, data) {
  if (!canManage_(user)) throw new Error('ไม่มีสิทธิ์จัดการหมวด');
  const sh=sheet_('Categories'), values=sh.getDataRange().getValues();
  const idx=values.findIndex((r,i)=>i>0 && r[0]===data.id); if(idx<0) throw new Error('ไม่พบหมวด');
  sh.getRange(idx+1,4).setValue(!!data.active); invalidateRows_('Categories');
  audit_(user,'UPDATE','CATEGORY',data.id,'active='+!!data.active); return ok({refresh:bootstrap_(user).data});
}

function deleteCategory_(user, data) {
  if (!canManage_(user)) throw new Error('ไม่มีสิทธิ์ลบหมวด');
  if(data.id==='CAT-DIVIDEND') throw new Error('หมวดเงินปันผลเป็นหมวดระบบ จึงไม่สามารถลบได้');
  const sh=sheet_('Categories'), values=sh.getDataRange().getValues();
  const idx=values.findIndex((r,i)=>i>0 && r[0]===data.id);
  if(idx<0) throw new Error('ไม่พบหมวดที่ต้องการลบ');
  const name=String(values[idx][1]||'');
  sh.deleteRow(idx+1);
  invalidateRows_('Categories');
  audit_(user,'DELETE','CATEGORY',data.id,name);
  return ok({refresh:bootstrap_(user).data});
}

function moveCategory_(user,data){
  if(!canManage_(user)) throw new Error('ไม่มีสิทธิ์จัดลำดับหมวด'); const sh=sheet_('Categories'), rows=rows_('Categories').sort((a,b)=>(num_(a.sortOrder)||999)-(num_(b.sortOrder)||999)), i=rows.findIndex(x=>x.id===data.id), j=data.direction==='UP'?i-1:i+1;
  if(i<0||j<0||j>=rows.length) return ok({refresh:bootstrap_(user).data}); const a=rows[i],b=rows[j], ao=num_(a.sortOrder)||i+1,bo=num_(b.sortOrder)||j+1, values=sh.getDataRange().getValues(), ai=values.findIndex((r,n)=>n>0&&r[0]===a.id),bi=values.findIndex((r,n)=>n>0&&r[0]===b.id); sh.getRange(ai+1,5).setValue(bo);sh.getRange(bi+1,5).setValue(ao);invalidateRows_('Categories');return ok({refresh:bootstrap_(user).data});
}
function reorderCategory_(user,data){
  if(!canManage_(user))throw new Error('ไม่มีสิทธิ์จัดลำดับหมวด');const ids=Array.isArray(data.ids)?data.ids:[],sh=sheet_('Categories'),v=sh.getDataRange().getValues(),known=v.slice(1).map(r=>r[0]);if(ids.length!==known.length||ids.some(id=>!known.includes(id)))throw new Error('ข้อมูลลำดับหมวดไม่ถูกต้อง');const positions=Object.fromEntries(ids.map((id,i)=>[id,i+1]));sh.getRange(2,5,known.length,1).setValues(known.map(id=>[positions[id]]));invalidateRows_('Categories');return ok({refresh:bootstrap_(user).data});
}

function saveProductCategory_(user,data){
  if(!canManage_(user)) throw new Error('ไม่มีสิทธิ์จัดการหมวดสินค้า'); const name=clean_(data.name);if(name.length<2)throw new Error('กรุณาระบุชื่อหมวดสินค้า');const sh=sheet_('ProductCategories'),v=sh.getDataRange().getValues();if(v.some((r,i)=>i>0&&r[0]!==data.id&&String(r[1]).toLowerCase()===name.toLowerCase()))throw new Error('มีหมวดสินค้านี้แล้ว');if(data.id){const i=v.findIndex((r,n)=>n>0&&r[0]===data.id);if(i<0)throw new Error('ไม่พบหมวดสินค้า');sh.getRange(i+1,2).setValue(name)}else append_('ProductCategories',[id_('PC'),name,true,Math.max(0,...v.slice(1).map(r=>num_(r[3])))+1]);invalidateRows_('ProductCategories');return ok({refresh:bootstrap_(user).data});
}
function toggleProductCategory_(user,data){if(!canManage_(user))throw new Error('ไม่มีสิทธิ์');const sh=sheet_('ProductCategories'),v=sh.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&r[0]===data.id);if(i<0)throw new Error('ไม่พบหมวดสินค้า');sh.getRange(i+1,3).setValue(!!data.active);invalidateRows_('ProductCategories');return ok({refresh:bootstrap_(user).data})}

function saveUnit_(user,data){
  if(!canManage_(user))throw new Error('ไม่มีสิทธิ์จัดการหน่วย');const name=clean_(data.name),base=clean_(data.baseUnit)||name,factor=num_(data.factor);if(!name||factor<=0)throw new Error('ชื่อหน่วยและตัวคูณไม่ถูกต้อง');const sh=sheet_('Units'),v=sh.getDataRange().getValues();if(v.some((r,i)=>i>0&&r[0]!==data.id&&String(r[1]).toLowerCase()===name.toLowerCase()))throw new Error('มีหน่วยนี้แล้ว');if(data.id){const i=v.findIndex((r,n)=>n>0&&r[0]===data.id);if(i<0)throw new Error('ไม่พบหน่วย');sh.getRange(i+1,2,1,3).setValues([[name,factor,base]])}else append_('Units',[id_('UNIT'),name,factor,base,true,Math.max(0,...v.slice(1).map(r=>num_(r[5])))+1]);invalidateRows_('Units');return ok({refresh:bootstrap_(user).data});
}
function toggleUnit_(user,data){if(!canManage_(user))throw new Error('ไม่มีสิทธิ์');const sh=sheet_('Units'),v=sh.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&r[0]===data.id);if(i<0)throw new Error('ไม่พบหน่วย');sh.getRange(i+1,5).setValue(!!data.active);invalidateRows_('Units');return ok({refresh:bootstrap_(user).data})}

function saveOpeningBalances_(user,data){
  if(!canAdminUsers_(user)) throw new Error('เฉพาะเจ้าของร้านหรือแอดมินเท่านั้น');
  if(!data||!Array.isArray(data.accounts)) throw new Error('ข้อมูลบัญชีไม่ถูกต้อง');
  const sh=sheet_('Accounts'), values=sh.getDataRange().getValues();
  data.accounts.forEach(x=>{ const idx=values.findIndex((r,i)=>i>0&&r[0]===x.id), amount=Number(x.openingBalance); if(idx<0||!isFinite(amount)) throw new Error('ยอดตั้งต้นไม่ถูกต้อง'); values[idx][3]=round_(amount); });
  sh.getRange(2,4,values.length-1,1).setValues(values.slice(1).map(r=>[r[3]]));
  SpreadsheetApp.flush();invalidateRows_('Accounts'); audit_(user,'UPDATE','ACCOUNTS','OPENING_BALANCE','แก้ไขยอดเงินตั้งต้น');
  return ok({refresh:bootstrap_(user).data});
}

function normalizeItem_(x,i) {
  const name=clean_(x.name), quantity=num_(x.quantity), total=num_(x.lineTotal), unitName=clean_(x.unit), unit=cachedRows_('Units',600).find(u=>u.name===unitName&&truthy_(u.active));
  if(!name) throw new Error('รายการที่ '+(i+1)+': กรุณาระบุชื่อ');
  if(quantity<=0 || total<=0) throw new Error('รายการที่ '+(i+1)+': จำนวนและราคาต้องมากกว่า 0');
  let category=clean_(x.category)||'อื่น ๆ';
  if(category==='วัตถุดิบ') category='อื่น ๆ';
  if(!cachedRows_('ProductCategories',600).some(r=>r.name===category&&truthy_(r.active))) throw new Error('รายการที่ '+(i+1)+': หมวดสินค้าไม่ถูกต้อง');
  if(!unit) throw new Error('รายการที่ '+(i+1)+': หน่วยไม่ถูกต้อง'); const factor=num_(unit.factor)||1;
  return {name,category,quantity,lineTotal:round_(total),unit:unit.name,unitFactor:factor,baseQuantity:quantity*factor,baseUnit:unit.baseUnit||unit.name};
}

function productCategorySeed_(){ return [
  ['PC-VEG','ผักและผลไม้',true,1],['PC-MEAT','เนื้อสัตว์และอาหารทะเล',true,2],
  ['PC-PANTRY','วัตถุดิบครัว/ของแห้ง',true,3],['PC-BAR','บาร์น้ำและเครื่องดื่ม',true,4],
  ['PC-SNACK','ขนมและเบเกอรี่',true,5],['PC-ANIMAL','อาหารสัตว์',true,6],
  ['PC-TOY','ของเล่น',true,7],['PC-WEAR','ถุงเท้า/เครื่องแต่งกาย',true,8],
  ['PC-PACK','บรรจุภัณฑ์',true,9],['PC-CLEAN','ทำความสะอาด',true,10],['PC-OTHER','อื่น ๆ',true,99]
]; }
function unitSeed_(){return [['UNIT-KG','กก.',1,'กก.',true,1],['UNIT-G','กรัม',.001,'กก.',true,2],['UNIT-L','ลิตร',1,'ลิตร',true,3],['UNIT-ML','มล.',.001,'ลิตร',true,4],['UNIT-FRUIT','ลูก',1,'ลูก',true,5],['UNIT-PCS','ชิ้น',1,'ชิ้น',true,6],['UNIT-PACK','แพ็ก',1,'แพ็ก',true,7],['UNIT-BAG','ถุง',1,'ถุง',true,8],['UNIT-BOTTLE','ขวด',1,'ขวด',true,9],['UNIT-BOX','กล่อง',1,'กล่อง',true,10],['UNIT-CASE','ลัง',1,'ลัง',true,11],['UNIT-PAIR','คู่',1,'คู่',true,12]]}

function ensureOwnerNameMigration_(){
  const props=PropertiesService.getScriptProperties();if(props.getProperty('OWNER_NAME_BAS_TANGMO')==='1')return;const sh=sheet_('Users'),rows=sh.getDataRange().getValues(),i=rows.findIndex((r,n)=>n>0&&r[4]==='OWNER'&&clean_(r[3])==='เจ้าของร้าน');if(i>0){sh.getRange(i+1,4).setValue('บาส/แตงโม');SpreadsheetApp.flush()}invalidateRows_('Users');props.setProperty('OWNER_NAME_BAS_TANGMO','1');
}

function ensureOwnerLoginRepair_(){
  const props=PropertiesService.getScriptProperties();if(props.getProperty('OWNER_LOGIN_REPAIR_20260819')==='1')return;const sh=sheet_('Users'),rows=sh.getDataRange().getValues(),i=rows.findIndex((r,n)=>n>0&&r[4]==='OWNER');if(i<1)throw new Error('ไม่พบบัญชีเจ้าของร้าน');const duplicate=rows.findIndex((r,n)=>n>0&&n!==i&&String(r[1]).toLowerCase()==='admin');if(duplicate>0)throw new Error('มี Username admin ซ้ำในระบบ');sh.getRange(i+1,2,1,5).setValues([['admin',hash_('admin123'),'บาส/แตงโม','OWNER',true]]);SpreadsheetApp.flush();invalidateRows_('Users');props.setProperty('OWNER_LOGIN_REPAIR_20260819','1');
}

function ensureAllUserLoginRepair_(){
  const props=PropertiesService.getScriptProperties();if(props.getProperty('ALL_USER_LOGIN_REPAIR_20260819')==='1')return;const sh=sheet_('Users'),rows=sh.getDataRange().getValues();if(rows.length<2)throw new Error('ไม่พบข้อมูลผู้ใช้');const updates=rows.slice(1).map(r=>{const username=clean_(r[1]).toLowerCase(),password=username==='admin'?'admin123':username+'123';return[hash_(password),true]});sh.getRange(2,3,updates.length,1).setValues(updates.map(x=>[x[0]]));sh.getRange(2,6,updates.length,1).setValues(updates.map(x=>[x[1]]));SpreadsheetApp.flush();invalidateRows_('Users');props.setProperty('ALL_USER_LOGIN_REPAIR_20260819','1');
}

function ensureSchema_(){
  const props=PropertiesService.getScriptProperties();
  if(props.getProperty('SCHEMA_VERSION')==='14') return;
  const ss=db_();
  const allSheets = ss.getSheets();
  Object.keys(APP.sheets).forEach(name=>{
    const target = clean_(name).toLowerCase();
    let sh = allSheets.find(s => clean_(s.getName()).toLowerCase() === target);
    if(!sh){
      sh = ss.insertSheet(name);
      sh.getRange(1, 1, 1, APP.sheets[name].length).setValues([APP.sheets[name]]).setFontWeight('bold').setBackground('#f6c90e');
      sh.setFrozenRows(1);
      allSheets.push(sh);
    } else {
      const current = sh.getRange(1, 1, 1, Math.max(1, sh.getLastColumn())).getValues()[0].map(clean_);
      APP.sheets[name].forEach(h=>{
        if(!current.map(c => c.toLowerCase()).includes(h.toLowerCase())){
          sh.getRange(1, sh.getLastColumn() + 1).setValue(h).setFontWeight('bold').setBackground('#f6c90e');
          current.push(h);
        }
      });
    }
  });
  const pc=sheet_('ProductCategories');
  if(pc.getLastRow()===1){ const seed=productCategorySeed_(); pc.getRange(2,1,seed.length,seed[0].length).setValues(seed); }
  const units=sheet_('Units');if(units.getLastRow()===1){const seed=unitSeed_();units.getRange(2,1,seed.length,seed[0].length).setValues(seed)}
  const cats=sheet_('Categories');if(cats.getLastRow()>1){const vals=cats.getRange(2,5,cats.getLastRow()-1,1).getValues().map((r,i)=>[num_(r[0])||i+1]);cats.getRange(2,5,vals.length,1).setValues(vals)}
  const catRows=cats.getDataRange().getValues();if(!catRows.slice(1).some(r=>r[0]==='CAT-DIVIDEND'))append_('Categories',['CAT-DIVIDEND','เงินปันผลเจ้าของร้าน','EXPENSE',true,Math.max(0,...catRows.slice(1).map(r=>num_(r[4])))+1]);
  invalidateRows_('Categories');
  props.setProperty('SCHEMA_VERSION','14');
}

function totals_(list) {
  const income=round_(list.filter(x=>x.type==='INCOME').reduce((s,x)=>s+num_(x.amount),0));
  const expense=round_(list.filter(x=>x.type==='EXPENSE').reduce((s,x)=>s+num_(x.amount),0));
  return {income,expense,net:round_(income-expense)};
}
function buildDashboardPeriods_(tx, categories){
  const today=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');
  const now=new Date(today+'T12:00:00'), yesterdayDate=new Date(now); yesterdayDate.setDate(now.getDate()-1);
  const yesterday=Utilities.formatDate(yesterdayDate,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const day=(now.getDay()+6)%7; const ws=new Date(now); ws.setDate(now.getDate()-day);
  const starts={YESTERDAY:yesterday,DAY:today,WEEK:Utilities.formatDate(ws,Session.getScriptTimeZone(),'yyyy-MM-dd'),MONTH:today.slice(0,7)+'-01',YEAR:today.slice(0,4)+'-01-01',ALL:'0000-01-01'};
  const categoryNames=Object.fromEntries(categories.map(c=>[c.id,c.name]));
  const purchases=cachedRows_('Purchases',60).filter(x=>x.status==='CONFIRMED');
  const purchaseDates=Object.fromEntries(purchases.map(x=>[x.id,dateKey_(x.date)]));
  const purchaseItems=cachedRows_('PurchaseItems',60);
  const result={};
  Object.keys(starts).forEach(key=>{
    const start=starts[key], end=key==='YESTERDAY'?yesterday:today, selected=tx.filter(x=>dateKey_(x.date)>=start && dateKey_(x.date)<=end), total=totals_(selected), grouped={};
    selected.filter(x=>x.type==='EXPENSE' && x.category!=='CAT-RAW').forEach(x=>{ const n=categoryNames[x.category]||'รายจ่ายอื่น'; grouped[n]=(grouped[n]||0)+num_(x.amount); });
    purchaseItems.forEach(x=>{ const d=purchaseDates[x.purchaseId]; if(d && d>=start && d<=end){ const n=clean_(x.productCategory)||'วัตถุดิบอื่น'; grouped[n]=(grouped[n]||0)+num_(x.lineTotal); } });
    const breakdown=Object.keys(grouped).map(name=>({name,amount:round_(grouped[name]),pct:total.expense?round_(grouped[name]/total.expense*100):0})).sort((a,b)=>b.amount-a.amount);
    result[key]={...total,breakdown};
  });
  return result;
}
function transactionViews_(tx,categories,accounts){
  const cn=Object.fromEntries(categories.map(x=>[x.id,x.name])), an=Object.fromEntries(accounts.map(x=>[x.id,x.name]));
  const itemNames={}; cachedRows_('PurchaseItems',60).forEach(x=>{(itemNames[x.purchaseId]||(itemNames[x.purchaseId]=[])).push(x.productName);});
  return tx.map(x=>({id:x.id,date:dateKey_(x.date),type:x.type,categoryId:x.category,categoryName:cn[x.category]||'ไม่ระบุหมวด',accountId:x.accountId,accountName:an[x.accountId]||'ไม่ระบุบัญชี',amount:num_(x.amount),note:x.note||'',referenceId:x.referenceId||'',purchaseItems:(itemNames[x.referenceId]||[]),createdAt:String(x.createdAt||''),createdBy:x.createdBy||'',createdByName:x.createdByName||'',status:x.status})).sort((a,b)=>String(b.createdAt).localeCompare(String(a.createdAt)));
}
let SESSION_SECRET_CACHE_;
function sessionSecret_(){
  if(SESSION_SECRET_CACHE_)return SESSION_SECRET_CACHE_;
  const props=PropertiesService.getScriptProperties();
  let secret=props.getProperty('SESSION_SIGNING_SECRET');
  if(!secret){
    secret=Utilities.getUuid()+Utilities.getUuid();
    props.setProperty('SESSION_SIGNING_SECRET',secret);
  }
  SESSION_SECRET_CACHE_=secret;
  return secret;
}

let SESSION_VERSION_CACHE_;
function getSessionVersion_(){
  if(SESSION_VERSION_CACHE_)return SESSION_VERSION_CACHE_;
  SESSION_VERSION_CACHE_=PropertiesService.getScriptProperties().getProperty('SESSION_VERSION')||'1';
  return SESSION_VERSION_CACHE_;
}

function legacyTokenSignature_(user,version){
  const pwdHash = String(user.passwordHash || user.passwordhash || '');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(user.id+'|'+pwdHash+'|'+version,sessionSecret_())).replace(/=+$/,'');
}

function tokenSignature_(user,version,nonce){
  const pwdHash = String(user.passwordHash || user.passwordhash || '');
  return Utilities.base64EncodeWebSafe(Utilities.computeHmacSha256Signature(user.id+'|'+pwdHash+'|'+version+'|'+(nonce||''),sessionSecret_())).replace(/=+$/,'');
}

function makePersistentToken_(user){
  const version=getSessionVersion_();
  const payload=Utilities.base64EncodeWebSafe(String(user.id)).replace(/=+$/,'');
  const nonce=Utilities.getUuid().replace(/-/g,'');
  return 'p2.'+payload+'.'+nonce+'.'+tokenSignature_(user,version,nonce);
}

function userFromPersistentToken_(token){
  try{
    const parts=String(token).split('.');
    const legacy=parts[0]==='p1';
    if((legacy&&parts.length!==3)||(!legacy&&(parts[0]!=='p2'||parts.length!==4)))return null;
    const id=Utilities.newBlob(Utilities.base64DecodeWebSafe(parts[1])).getDataAsString();
    const user=cachedRows_('Users',120).find(x=>x.id===id)||rows_('Users').find(x=>x.id===id);
    if(!user||!truthy_(user.active))return null;
    const version=getSessionVersion_();
    const nonce=legacy?'':parts[2];
    const signature=legacy?parts[2]:parts[3];
    const expected=legacy?legacyTokenSignature_(user,version):tokenSignature_(user,version,nonce);
    if(signature!==expected)return null;
    return{...publicUser_(user),sessionVersion:version,persistent:true};
  }catch(e){
    return null;
  }
}

function requireSession_(token){
  if(!token)throw new Error('SESSION_EXPIRED');
  const cache=CacheService.getScriptCache();
  const raw=cache.get('session:'+token);
  let user=null;
  const version=getSessionVersion_();

  if(raw){
    try{
      const parsed=JSON.parse(raw);
      if(parsed&&parsed.sessionVersion===version){
        user=parsed;
      }
    }catch(e){}
  }

  if(!user){
    user=userFromPersistentToken_(token);
  }

  if(!user||user.sessionVersion!==version){
    throw new Error('SESSION_EXPIRED');
  }

  const latest=cachedRows_('Users',120).find(x=>x.id===user.id)||rows_('Users').find(x=>x.id===user.id);
  if(!latest||!truthy_(latest.active)){
    cache.remove('session:'+token);
    throw new Error('SESSION_EXPIRED');
  }

  const fresh={...publicUser_(latest),sessionVersion:version,persistent:true};
  cache.put('session:'+token,JSON.stringify(fresh),21600);
  return fresh;
}
function publicUser_(u) {
  if (!u) return { id: '', name: '', username: '', role: 'STAFF', active: false };
  return {
    id: clean_(u.id || u.Id || u.ID),
    name: clean_(u.name || u.Name) || clean_(u.username || u.Username),
    username: clean_(u.username || u.Username).toLowerCase(),
    role: String(u.role || u.Role || 'STAFF').toUpperCase(),
    active: truthy_(u.active != null ? u.active : u.Active)
  };
}
function canManage_(u) { return ['ADMIN','MANAGER','OWNER'].includes(u.role); }
function canAdminUsers_(u) { return ['ADMIN','OWNER'].includes(u.role); }
function canViewFinance_(u) { return ['ADMIN','OWNER'].includes(u.role); }
function validateUserInput_(d) { if(clean_(d.name).length<2) throw new Error('กรุณาระบุชื่อ'); if(!/^[a-zA-Z0-9._-]{3,30}$/.test(clean_(d.username))) throw new Error('ชื่อผู้ใช้ต้องเป็นอังกฤษหรือตัวเลขอย่างน้อย 3 ตัว'); if(String(d.password||'').length<6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัว'); }
function hash_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s))); }
let DB_CACHE_;
let ROWS_MEMO_ = {};
function db_() {
  if (DB_CACHE_) return DB_CACHE_;
  const props = PropertiesService.getScriptProperties();
  const configuredId = props.getProperty('DB_ID');
  const targetId = clean_(configuredId) || PRIMARY_DB_ID;
  try {
    DB_CACHE_ = SpreadsheetApp.openById(targetId);
    return DB_CACHE_;
  } catch(e) {
    if (configuredId && configuredId !== PRIMARY_DB_ID) {
      try {
        DB_CACHE_ = SpreadsheetApp.openById(PRIMARY_DB_ID);
        return DB_CACHE_;
      } catch(e2) {}
    }
    throw new Error('ไม่สามารถเปิดฐานข้อมูล Google Sheet ได้ (' + (e.message || targetId) + ')');
  }
}
function sheet_(name) {
  const ss = db_();
  let sh = ss.getSheetByName(name);
  if (sh) return sh;
  const cleanTarget = clean_(name).toLowerCase();
  const allSheets = ss.getSheets();
  sh = allSheets.find(s => clean_(s.getName()).toLowerCase() === cleanTarget);
  if (sh) return sh;
  const schemaKey = Object.keys(APP.sheets).find(k => clean_(k).toLowerCase() === cleanTarget);
  if (schemaKey) {
    const cols = APP.sheets[schemaKey];
    sh = ss.insertSheet(schemaKey);
    sh.getRange(1, 1, 1, cols.length).setValues([cols]).setFontWeight('bold').setBackground('#f6c90e');
    sh.setFrozenRows(1);
    SpreadsheetApp.flush();
    return sh;
  }
  throw new Error('ไม่พบชีต ' + name);
}
function rows_(name) {
  const sh = sheet_(name);
  const vals = sh.getDataRange().getValues();
  if (!vals || vals.length <= 1) return [];
  const rawHead = vals.shift();
  const head = rawHead.map(h => clean_(h));
  return vals
    .filter(r => r && r.some(v => v !== ''))
    .map(r => {
      const obj = {};
      head.forEach((h, i) => {
        if (!h) return;
        const val = r[i];
        obj[h] = val;
        const lower = h.toLowerCase();
        if (obj[lower] === undefined) obj[lower] = val;
        if (lower === 'passwordhash') obj.passwordHash = val;
        if (lower === 'createdat') obj.createdAt = val;
        if (lower === 'createdby') obj.createdBy = val;
        if (lower === 'createdbyname') obj.createdByName = val;
        if (lower === 'openingbalance') obj.openingBalance = val;
        if (lower === 'baseunit') obj.baseUnit = val;
        if (lower === 'sortorder') obj.sortOrder = val;
        if (lower === 'weeklywage') obj.weeklyWage = val;
        if (lower === 'employmenttype') obj.employmentType = val;
        if (lower === 'employeeid') obj.employeeId = val;
        if (lower === 'weekstart') obj.weekStart = val;
        if (lower === 'weekend') obj.weekEnd = val;
        if (lower === 'accountid') obj.accountId = val;
        if (lower === 'transactionid') obj.transactionId = val;
        if (lower === 'purchaseid') obj.purchaseId = val;
        if (lower === 'productid') obj.productId = val;
        if (lower === 'productname') obj.productName = val;
        if (lower === 'productcategory') obj.productCategory = val;
        if (lower === 'linetotal') obj.lineTotal = val;
        if (lower === 'basequantity') obj.baseQuantity = val;
        if (lower === 'baseunitprice') obj.baseUnitPrice = val;
        if (lower === 'previousprice') obj.previousPrice = val;
        if (lower === 'pricechangepct') obj.priceChangePct = val;
        if (lower === 'paydate') obj.payDate = val;
        if (lower === 'dailyrate') obj.dailyRate = val;
        if (lower === 'paymentid') obj.paymentId = val;
        if (lower === 'owneruserid') obj.ownerUserId = val;
        if (lower === 'ownername') obj.ownerName = val;
        if (lower === 'basewage') obj.baseWage = val;
        if (lower === 'advancetotal') obj.advanceTotal = val;
        if (lower === 'netpaid') obj.netPaid = val;
        if (lower === 'paidat') obj.paidAt = val;
        if (lower === 'referenceid') obj.referenceId = val;
      });
      return obj;
    });
}
function cachePutChunked_(key, str, seconds) {
  try {
    const cache = CacheService.getScriptCache();
    const chunkSize = 90000;
    if (str.length <= chunkSize) {
      cache.put(key, str, seconds || 600);
      cache.remove(key + ':c');
      return;
    }
    const count = Math.ceil(str.length / chunkSize);
    const map = {};
    for (let i = 0; i < count; i++) {
      map[key + ':' + i] = str.substring(i * chunkSize, (i + 1) * chunkSize);
    }
    map[key + ':c'] = String(count);
    cache.putAll(map, seconds || 600);
  } catch(e) {}
}

function cacheGetChunked_(key) {
  try {
    const cache = CacheService.getScriptCache();
    const countStr = cache.get(key + ':c');
    if (!countStr) {
      return cache.get(key);
    }
    const count = parseInt(countStr, 10);
    if (isNaN(count) || count <= 0) return null;
    const keys = [];
    for (let i = 0; i < count; i++) keys.push(key + ':' + i);
    const map = cache.getAll(keys);
    let res = '';
    for (let i = 0; i < count; i++) {
      const part = map[key + ':' + i];
      if (part == null) return null;
      res += part;
    }
    return res;
  } catch(e) {
    return null;
  }
}

function cachedRows_(name, seconds) {
  if (Object.prototype.hasOwnProperty.call(ROWS_MEMO_, name)) return ROWS_MEMO_[name];
  const key = 'rows:' + name;
  const hit = cacheGetChunked_(key);
  if (hit) {
    try {
      const parsed = JSON.parse(hit);
      ROWS_MEMO_[name] = parsed;
      return parsed;
    } catch(e) {}
  }
  const data = rows_(name);
  ROWS_MEMO_[name] = data;
  const raw = JSON.stringify(data);
  cachePutChunked_(key, raw, seconds || 1800);
  return data;
}

function invalidateRows_(name) {
  delete ROWS_MEMO_[name];
  try {
    const cache = CacheService.getScriptCache();
    cache.remove('rows:' + name);
    cache.remove('rows:' + name + ':c');
    for (let i = 0; i < 10; i++) cache.remove('rows:' + name + ':' + i);
  } catch(e) {}
  PropertiesService.getScriptProperties().setProperty('DATA_REVISION', String(Date.now()));
}
function append_(name,row) { sheet_(name).appendRow(row); }
function appendRows_(name, rows) { if(!rows.length) return; const sh=sheet_(name); sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows); }
function deleteRowById_(name,id){if(!id)return;const sh=sheet_(name),values=sh.getDataRange().getValues(),index=values.findIndex((row,i)=>i>0&&row[0]===id);if(index>0)sh.deleteRow(index+1)}
function deleteRowsByField_(name,field,value){const sh=sheet_(name),values=sh.getDataRange().getValues(),column=values[0].indexOf(field);if(column<0)return;for(let i=values.length-1;i>0;i--)if(values[i][column]===value)sh.deleteRow(i+1)}
function audit_(u,a,e,id,d) { try{append_('AuditLog',[now_(),u.id,u.name,a,e,id,d]);PropertiesService.getScriptProperties().setProperty('DATA_REVISION',String(Date.now()))}catch(err){} }
function now_() { return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd HH:mm:ss'); }
function dateKey_(v) {
  if(v instanceof Date) return Utilities.formatDate(v,Session.getScriptTimeZone(),'yyyy-MM-dd');
  const s=String(v||'');
  if(/^\d{4}-\d{2}-\d{2}$/.test(s.slice(0,10)) && !s.includes('T')) return s.slice(0,10);
  if(/^\d{4}-\d{2}-\d{2}T/.test(s)){const d=new Date(s);if(!isNaN(d.getTime()))return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd')}
  return s.slice(0,10);
}
function validDate_(v) { const s=String(v||''); if(!/^\d{4}-\d{2}-\d{2}$/.test(s)) throw new Error('วันที่ไม่ถูกต้อง'); return s; }
function id_(p) { return p+'-'+Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMddHHmmss')+'-'+Math.floor(Math.random()*9000+1000); }
function clean_(v) { return String(v==null?'':v).trim(); }
function num_(v) { const n=Number(v); return isFinite(n)?n:0; }
function round_(v) { return Math.round((v+Number.EPSILON)*100)/100; }
function recentWrite_(value){const time=new Date(value).getTime(),age=Date.now()-time;return Number.isFinite(time)&&age>=0&&age<120000;}
function dedupeRows_(rows,fields){const seen=new Set();return rows.filter(row=>{const key=fields.map(field=>field==='date'||field==='weekStart'?dateKey_(row[field]):String(row[field]??'')).join('\u001f');if(seen.has(key))return false;seen.add(key);return true;});}
function dedupeTimedRows_(rows,fields,timeField,windowMs){const seen={};return rows.filter(row=>{const key=fields.map(field=>field==='date'||field==='weekStart'?dateKey_(row[field]):String(row[field]??'')).join('\u001f'),raw=row[timeField],time=raw instanceof Date?raw.getTime():new Date(String(raw||'').replace(' ','T')+'+07:00').getTime(),previous=seen[key];if(Number.isFinite(time)&&Number.isFinite(previous)&&Math.abs(time-previous)<=windowMs)return false;if(Number.isFinite(time))seen[key]=time;return true;});}
function truthy_(v) { return v===true || String(v).toLowerCase()==='true'; }
function ok(data) { return {ok:true,data:data||{}}; }
