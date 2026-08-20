const APP = {
  name: 'Wonder Duck Accounts',
  version: '2.4.0',
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
    WagePayments: ['id','employeeId','weekStart','weekEnd','baseWage','advanceTotal','adjustment','netPaid','accountId','transactionId','note','paidAt','createdBy','createdByName'],
    DividendPayments: ['id','ownerUserId','ownerName','date','amount','accountId','transactionId','note','createdAt','createdBy','createdByName'],
    AuditLog: ['timestamp','userId','userName','action','entity','entityId','detail']
  }
};
const PRIMARY_DB_ID = '1EN084DNwJjxZdStDABk0znC-rLAPMzlCtHsLqQ4TnRs';

function doGet() {
  try{if(PropertiesService.getScriptProperties().getProperty('DB_ID')){ensureOwnerNameMigration_();ensureOwnerLoginRepair_();ensureAllUserLoginRepair_()}}catch(e){}
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
      ,reorderCategory: () => reorderCategory_(user, payload.data)
      ,compareFinancialPeriods: () => compareFinancialPeriods_(user, payload.data)
      ,getFinancialPeriod: () => getFinancialPeriod_(user, payload.data)
      ,getTransactionHistory: () => getTransactionHistory_(user, payload.data)
      ,saveDividend: () => saveDividend_(user, payload.data)
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
  const username = clean_(data.username).toLowerCase();
  const user = rows_('Users').find(r => String(r.username).toLowerCase() === username);
  if (!user || !truthy_(user.active) || String(user.passwordHash) !== hash_(data.password)) throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  const token = Utilities.getUuid().replace(/-/g, '');
  const sessionUser=publicUser_(user); sessionUser.sessionVersion=PropertiesService.getScriptProperties().getProperty('SESSION_VERSION')||'1';sessionUser.persistent=true;
  const sessionRaw=JSON.stringify(sessionUser);CacheService.getScriptCache().put('session:' + token,sessionRaw,21600);PropertiesService.getScriptProperties().setProperty('SESSION_TOKEN_'+token,sessionRaw);
  return ok({ token, user: publicUser_(user), appData:bootstrap_(sessionUser).data });
}

function logout_(token){if(token){CacheService.getScriptCache().remove('session:'+token);PropertiesService.getScriptProperties().deleteProperty('SESSION_TOKEN_'+token)}return ok({loggedOut:true})}

function bootstrap_(user) {
  const props=PropertiesService.getScriptProperties();
  const revision=props.getProperty('DATA_REVISION')||'0';
  const day=Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyyMMdd');
  const cache=CacheService.getScriptCache();
  const key=['bootstrap',APP.version,user.id,user.role,revision,day].join(':');
  const hit=cache.get(key);
  if(hit){try{return JSON.parse(hit)}catch(e){cache.remove(key)}}
  const result=buildBootstrap_(user);
  const raw=JSON.stringify(result);
  if(raw.length<95000)cache.put(key,raw,120);
  return result;
}

function buildBootstrap_(user) {
  ensureSchema_();
  ensureOwnerNameMigration_();
  ensureOwnerLoginRepair_();
  ensureAllUserLoginRepair_();
  const latestUser=cachedRows_('Users',120).find(x=>x.id===user.id);if(latestUser)user={...publicUser_(latestUser),sessionVersion:user.sessionVersion};
  const canViewFinance=canViewFinance_(user);
  const accountRows = cachedRows_('Accounts', 600).filter(x => truthy_(x.active));
  const accounts = canViewFinance ? accountRows : accountRows.map(x=>({id:x.id,name:x.name,type:x.type,active:true}));
  const allCategories = cachedRows_('Categories', 600).sort((a,b)=>(num_(a.sortOrder)||999)-(num_(b.sortOrder)||999));
  const categories = allCategories.filter(x => truthy_(x.active));
  const confirmedPurchaseIds=new Set(cachedRows_('Purchases',60).filter(x=>x.status==='CONFIRMED').map(x=>x.id)),pricePairs={};cachedRows_('PurchaseItems',60).filter(x=>confirmedPurchaseIds.has(x.purchaseId)).forEach(x=>{const list=pricePairs[x.productId]||(pricePairs[x.productId]=[]);list.push(num_(x.baseUnitPrice));if(list.length>2)list.shift()});
  const products = cachedRows_('Products', 120).filter(x => truthy_(x.active)).map(x=>{const prices=pricePairs[x.id]||[],latest=prices.length?prices[prices.length-1]:null,previous=prices.length>1?prices[prices.length-2]:null,priceChange=previous?round_((latest-previous)/previous*100):null;return{...x,latestPrice:latest,previousPrice:previous,priceChange}});
  const allTransactions = cachedRows_('Transactions', 30);
  const tx = allTransactions.filter(x => x.status === 'CONFIRMED');
  const movements={};
  if(canViewFinance) tx.forEach(t=>{movements[t.accountId]=(movements[t.accountId]||0)+(t.type==='INCOME'?num_(t.amount):-num_(t.amount));});
  const balances = canViewFinance ? accountRows.map(a => ({id:a.id,name:a.name,type:a.type,balance:round_(num_(a.openingBalance)+(movements[a.id]||0)),openingBalance:num_(a.openingBalance)})) : [];
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
  return ok({
    user: publicUser_(user), permissions:{canViewFinance,canAdminUsers:canAdminUsers_(user),canManageCatalog:canManage_(user),canViewAnalytics:user.role==='OWNER'}, accounts, categories, products, productCategories: cachedRows_('ProductCategories', 600).filter(x => truthy_(x.active)).sort((a,b)=>num_(a.sortOrder)-num_(b.sortOrder)), units:cachedRows_('Units',600).filter(x=>truthy_(x.active)).sort((a,b)=>num_(a.sortOrder)-num_(b.sortOrder)), balances,
    today: canViewFinance?summarize(today):totals_([]), month: canViewFinance?summarize(month):totals_([]), previousMonth: canViewFinance?summarize(previousMonth):totals_([]), dashboardPeriods,
    recentToday: visibleTransactions.filter(x=>x.date===today && x.status==='CONFIRMED').slice(0,10),
    transactionHistory: visibleTransactions.slice(0,50),
    users: canAdminUsers_(user) ? cachedRows_('Users', 120).map(publicUser_) : [],
    allCategories: canManage_(user) ? allCategories : []
    ,catalogProductCategories:canManage_(user)?cachedRows_('ProductCategories',600):[]
    ,catalogUnits:canManage_(user)?cachedRows_('Units',600):[]
    ,managerOverview:user.role==='MANAGER'?summarize(month):null
    ,managerPeriods
    ,managerHistory:user.role==='MANAGER'?visibleTransactions.filter(x=>x.status==='CONFIRMED').slice(0,200):[]
    ,payroll:canAdminUsers_(user)?buildPayroll_():null
    ,dividends:user.role==='OWNER'?buildDividends_():null
    ,revision:PropertiesService.getScriptProperties().getProperty('DATA_REVISION')||'0'
  });
}

function buildDividends_(){
  const owners=cachedRows_('Users',120).filter(x=>truthy_(x.active)&&x.role==='OWNER').map(publicUser_);
  const accountNames=Object.fromEntries(cachedRows_('Accounts',600).map(x=>[x.id,x.name]));
  const history=cachedRows_('DividendPayments',60).slice(-100).reverse().map(x=>({id:x.id,ownerUserId:x.ownerUserId,ownerName:x.ownerName,date:dateKey_(x.date),amount:num_(x.amount),accountId:x.accountId,accountName:accountNames[x.accountId]||'',note:x.note||'',createdByName:x.createdByName||''}));
  return {owners,history,totalPaid:round_(history.reduce((s,x)=>s+x.amount,0))};
}

function saveDividend_(user,data){
  if(user.role!=='OWNER')throw new Error('เฉพาะเจ้าของร้านเท่านั้นที่บันทึกเงินปันผลได้');
  const owner=cachedRows_('Users',120).find(x=>x.id===data.ownerUserId&&x.role==='OWNER'&&truthy_(x.active));
  const account=cachedRows_('Accounts',600).find(x=>x.id===data.accountId&&truthy_(x.active));
  const amount=round_(num_(data.amount)),date=validDate_(data.date);
  if(!owner||!account||amount<=0)throw new Error('กรุณาระบุเจ้าของร้าน บัญชี และยอดปันผลให้ถูกต้อง');
  const id=id_('DIV'),tx=id_('TX'),note=clean_(data.note)||('จ่ายเงินปันผล '+owner.name);
  append_('Transactions',[tx,date,'EXPENSE','CAT-DIVIDEND',account.id,amount,note,id,now_(),user.id,user.name,'CONFIRMED']);
  append_('DividendPayments',[id,owner.id,owner.name,date,amount,account.id,tx,note,now_(),user.id,user.name]);
  SpreadsheetApp.flush();invalidateRows_('Transactions');invalidateRows_('DividendPayments');audit_(user,'CREATE','DIVIDEND',id,owner.name+' '+amount);
  return ok({id,refresh:bootstrap_(user).data});
}

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
  const list=rows_('Transactions').filter(x=>{const d=dateKey_(x.date);return d>=start&&d<=end});
  const views=transactionViews_(list,cachedRows_('Categories',600),cachedRows_('Accounts',600));
  return ok((user.role==='STAFF'?views.filter(x=>x.createdBy===user.id):views).slice(0,500));
}

function payrollWeek_(){
  const now=new Date(), day=(now.getDay()+6)%7, start=new Date(now); start.setHours(12,0,0,0); start.setDate(start.getDate()-day);
  const end=new Date(start); end.setDate(end.getDate()+6);
  return {start:Utilities.formatDate(start,Session.getScriptTimeZone(),'yyyy-MM-dd'),end:Utilities.formatDate(end,Session.getScriptTimeZone(),'yyyy-MM-dd')};
}
function buildPayroll_(){
  const week=payrollWeek_(), accounts=cachedRows_('Accounts',600).filter(x=>truthy_(x.active)), advances=cachedRows_('WageAdvances',60).filter(x=>dateKey_(x.weekStart)===week.start), overtime=cachedRows_('WageOvertime',60).filter(x=>dateKey_(x.weekStart)===week.start), payments=cachedRows_('WagePayments',60), paid=Object.fromEntries(payments.filter(x=>dateKey_(x.weekStart)===week.start).map(x=>[x.employeeId,x]));
  const accountNames=Object.fromEntries(accounts.map(x=>[x.id,x.name]));
  return {week,employees:cachedRows_('Employees',120).map(e=>{const list=advances.filter(x=>x.employeeId===e.id),ot=overtime.filter(x=>x.employeeId===e.id), advanceTotal=round_(list.reduce((s,x)=>s+num_(x.amount),0)),otTotal=round_(ot.reduce((s,x)=>s+num_(x.amount),0)), wage=num_(e.weeklyWage),totalEarned=round_(wage+otTotal), payment=paid[e.id];const remaining=round_(Math.max(0,wage+otTotal-advanceTotal));return {id:e.id,name:e.name,position:e.position||'พนักงาน',employmentType:e.employmentType||'FULL_TIME',weeklyWage:wage,otTotal,totalEarned,advanceTotal,remaining,progress:wage+otTotal?round_(remaining/(wage+otTotal)*100):0,paid:!!payment,netPaid:payment?num_(payment.netPaid):0,advances:list.map(x=>({id:x.id,date:dateKey_(x.date),amount:num_(x.amount),accountId:x.accountId,accountName:accountNames[x.accountId]||'',note:x.note||''})),overtime:ot.map(x=>({id:x.id,date:dateKey_(x.date),hours:num_(x.hours),rate:num_(x.rate),amount:num_(x.amount),note:x.note||''}))}}),recentPayments:payments.slice(-20).reverse()};
}

function saveEmployee_(user,data){
  if(!canAdminUsers_(user))throw new Error('เฉพาะเจ้าของร้านหรือแอดมินเท่านั้น');const name=clean_(data.name),position=clean_(data.position),employmentType=clean_(data.employmentType),wage=round_(num_(data.weeklyWage));if(name.length<2||position.length<2||wage<=0)throw new Error('กรุณาระบุชื่อ ตำแหน่ง และค่าจ้างให้ถูกต้อง');if(!['FULL_TIME','PART_TIME'].includes(employmentType))throw new Error('ประเภทการจ้างไม่ถูกต้อง');const sh=sheet_('Employees'),v=sh.getDataRange().getValues(),head=v[0],pi=head.indexOf('position')+1,ti=head.indexOf('employmentType')+1;if(data.id){const i=v.findIndex((r,n)=>n>0&&r[0]===data.id);if(i<0)throw new Error('ไม่พบพนักงาน');sh.getRange(i+1,2).setValue(name);sh.getRange(i+1,3).setValue(wage);sh.getRange(i+1,pi).setValue(position);sh.getRange(i+1,ti).setValue(employmentType);audit_(user,'UPDATE','EMPLOYEE',data.id,name)}else{const id=id_('EMP');append_('Employees',[id,name,wage,'',now_(),user.id,position,employmentType]);audit_(user,'CREATE','EMPLOYEE',id,name)}invalidateRows_('Employees');return ok({refresh:bootstrap_(user).data});
}
function deleteEmployee_(user,data){
  if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์ลบพนักงาน');if(rows_('WageAdvances').some(x=>x.employeeId===data.id)||rows_('WageOvertime').some(x=>x.employeeId===data.id)||rows_('WagePayments').some(x=>x.employeeId===data.id))throw new Error('พนักงานนี้มีประวัติค่าจ้างแล้ว จึงไม่สามารถลบได้');const sh=sheet_('Employees'),v=sh.getDataRange().getValues(),i=v.findIndex((r,n)=>n>0&&r[0]===data.id);if(i<0)throw new Error('ไม่พบพนักงาน');sh.deleteRow(i+1);invalidateRows_('Employees');audit_(user,'DELETE','EMPLOYEE',data.id,'');return ok({refresh:bootstrap_(user).data});
}
function saveWageAdvance_(user,data){
  if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์บันทึกเงินเบิก');const employee=rows_('Employees').find(x=>x.id===data.employeeId),account=rows_('Accounts').find(x=>x.id===data.accountId&&truthy_(x.active)),amount=round_(num_(data.amount)),date=validDate_(data.date),week=payrollWeek_();if(!employee||!account||amount<=0)throw new Error('ข้อมูลการเบิกเงินไม่ถูกต้อง');if(date<week.start||date>week.end)throw new Error('วันที่เบิกต้องอยู่ในรอบสัปดาห์ปัจจุบัน');if(rows_('WagePayments').some(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===week.start))throw new Error('พนักงานคนนี้ปิดรอบจ่ายแล้ว');const used=rows_('WageAdvances').filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===week.start).reduce((s,x)=>s+num_(x.amount),0),ot=rows_('WageOvertime').filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===week.start).reduce((s,x)=>s+num_(x.amount),0);if(amount>num_(employee.weeklyWage)+ot-used)throw new Error('ยอดเบิกมากกว่ายอดค่าจ้างที่เหลือ');const id=id_('ADV'),tx=id_('TX'),note=clean_(data.note)||('เบิกค่าจ้างล่วงหน้า '+employee.name);append_('Transactions',[tx,date,'EXPENSE','CAT-WAGE',account.id,amount,note,id,now_(),user.id,user.name,'CONFIRMED']);append_('WageAdvances',[id,employee.id,week.start,date,amount,account.id,tx,note,now_(),user.id,user.name]);invalidateRows_('Transactions');invalidateRows_('WageAdvances');audit_(user,'CREATE','WAGE_ADVANCE',id,employee.name+' '+amount);return ok({refresh:bootstrap_(user).data});
}
function saveWageOvertime_(user,data){
  if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์บันทึก OT');const employee=rows_('Employees').find(x=>x.id===data.employeeId),date=validDate_(data.date),hours=round_(num_(data.hours)),rate=round_(num_(data.rate)),week=payrollWeek_();if(!employee||hours<=0||rate<=0)throw new Error('กรุณาระบุวันที่ ชั่วโมง และค่า OT ให้ถูกต้อง');if(date<week.start||date>week.end)throw new Error('วันที่ OT ต้องอยู่ในรอบสัปดาห์ปัจจุบัน');if(rows_('WagePayments').some(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===week.start))throw new Error('พนักงานคนนี้ปิดรอบจ่ายแล้ว');const amount=round_(hours*rate),id=id_('OT');append_('WageOvertime',[id,employee.id,week.start,date,hours,rate,amount,clean_(data.note),now_(),user.id,user.name]);invalidateRows_('WageOvertime');audit_(user,'CREATE','WAGE_OT',id,employee.name+' '+amount);return ok({refresh:bootstrap_(user).data});
}
function payWeeklyWage_(user,data){
  if(!canAdminUsers_(user))throw new Error('ไม่มีสิทธิ์จ่ายค่าจ้าง');const employee=rows_('Employees').find(x=>x.id===data.employeeId),account=rows_('Accounts').find(x=>x.id===data.accountId&&truthy_(x.active)),week=payrollWeek_();if(!employee||!account)throw new Error('ข้อมูลการจ่ายไม่ถูกต้อง');if(rows_('WagePayments').some(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===week.start))throw new Error('รอบนี้จ่ายแล้ว');const advance=round_(rows_('WageAdvances').filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===week.start).reduce((s,x)=>s+num_(x.amount),0)),ot=round_(rows_('WageOvertime').filter(x=>x.employeeId===employee.id&&dateKey_(x.weekStart)===week.start).reduce((s,x)=>s+num_(x.amount),0)),adjustment=round_(num_(data.adjustment)),net=round_(num_(employee.weeklyWage)+ot-advance+adjustment);if(net<0)throw new Error('ยอดสุทธิติดลบ กรุณาตรวจสอบเงินเพิ่ม/หัก');const id=id_('PAY'),tx=net>0?id_('TX'):'',note=clean_(data.note)||('จ่ายค่าจ้าง '+employee.name+' รอบ '+week.start);if(net>0)append_('Transactions',[tx,validDate_(data.date),'EXPENSE','CAT-WAGE',account.id,net,note,id,now_(),user.id,user.name,'CONFIRMED']);append_('WagePayments',[id,employee.id,week.start,week.end,num_(employee.weeklyWage),advance,round_(adjustment+ot),net,account.id,tx,note,now_(),user.id,user.name]);invalidateRows_('Transactions');invalidateRows_('WagePayments');audit_(user,'CREATE','WAGE_PAYMENT',id,employee.name+' '+net);return ok({refresh:bootstrap_(user).data});
}

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
    const purchaseId = id_('BUY');
    append_('Purchases', [purchaseId,date,clean_(data.supplier),account.id,total,clean_(data.note),now_(),user.id,user.name,'CONFIRMED']);
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
      } else if (product.category !== item.category) { product.category=item.category; categoryUpdates.push(product); }
      const previous = latestPrice[product.id] || 0;
      const basePrice = round_(item.lineTotal/item.baseQuantity);
      const change = previous ? round_((basePrice-previous)/previous*100) : '';
      itemRows.push([id_('ITM'),purchaseId,product.id,product.name,item.quantity,item.unit,item.unitFactor,item.baseQuantity,item.lineTotal,basePrice,previous || '',change,item.category]);
    });
    appendRows_('Products', newProducts);
    if(categoryUpdates.length){ const sh=sheet_('Products'); categoryUpdates.forEach(p=>{ const idx=productRows.findIndex(x=>x.id===p.id); if(idx>=0) sh.getRange(idx+2,3).setValue(p.category); }); }
    appendRows_('PurchaseItems', itemRows);
    invalidateRows_('Products');
    invalidateRows_('Purchases'); invalidateRows_('PurchaseItems');
    append_('Transactions', [id_('TX'),date,'EXPENSE','CAT-RAW',account.id,total,clean_(data.note) || 'ซื้อวัตถุดิบ',purchaseId,now_(),user.id,user.name,'CONFIRMED']);
    SpreadsheetApp.flush();
    invalidateRows_('Transactions');
    audit_(user,'CREATE','PURCHASE',purchaseId,items.length+' รายการ รวม '+total+' บาท');
    return ok({ id:purchaseId, total, itemCount:items.length, refresh:bootstrap_(user).data });
  } finally { lock.releaseLock(); }
}

function saveTransaction_(user, data) {
  if (!data || !['INCOME','EXPENSE'].includes(data.type)) throw new Error('ประเภทรายการไม่ถูกต้อง');
  const amount = num_(data.amount); if (amount <= 0) throw new Error('จำนวนเงินต้องมากกว่า 0');
  const account = rows_('Accounts').find(x => x.id === data.accountId && truthy_(x.active));
  const category = rows_('Categories').find(x => x.id === data.category && x.type === data.type && (data.id || truthy_(x.active)));
  if (!account || !category) throw new Error('กรุณาเลือกบัญชีและหมวดหมู่ให้ถูกต้อง');
  const date=validDate_(data.date);
  if(data.id){
    const sh=sheet_('Transactions'), values=sh.getDataRange().getValues(), head=values[0], idx=values.findIndex((r,i)=>i>0 && r[0]===data.id);
    if(idx<0) throw new Error('ไม่พบรายการที่ต้องการแก้ไข');
    const old=Object.fromEntries(head.map((h,i)=>[h,values[idx][i]])); if(old.status==='CANCELLED') throw new Error('รายการนี้ถูกยกเลิกแล้ว'); if(!canManage_(user)&&old.createdBy!==user.id) throw new Error('แก้ไขได้เฉพาะรายการที่ตนเองบันทึก');
    if(old.referenceId){
      sh.getRange(idx+1,2).setValue(date); sh.getRange(idx+1,5).setValue(account.id); sh.getRange(idx+1,7).setValue(clean_(data.note));
      updatePurchaseHeader_(old.referenceId,date,account.id,clean_(data.note));
    } else sh.getRange(idx+1,2,1,6).setValues([[date,data.type,category.id,account.id,round_(amount),clean_(data.note)]]);
    SpreadsheetApp.flush(); invalidateRows_('Transactions'); audit_(user,'UPDATE','TRANSACTION',data.id,'แก้ไขรายการ'); return ok({id:data.id,refresh:bootstrap_(user).data});
  }
  const txId = id_('TX'); append_('Transactions',[txId,date,data.type,category.id,account.id,round_(amount),clean_(data.note),'',now_(),user.id,user.name,'CONFIRMED']);
  SpreadsheetApp.flush(); invalidateRows_('Transactions'); audit_(user,'CREATE','TRANSACTION',txId,data.type+' '+amount+' บาท'); return ok({ id:txId,refresh:bootstrap_(user).data });
}

function deleteTransaction_(user,data){
  const sh=sheet_('Transactions'), values=sh.getDataRange().getValues(), head=values[0], idx=values.findIndex((r,i)=>i>0 && r[0]===data.id);
  if(idx<0) throw new Error('ไม่พบรายการ'); const old=Object.fromEntries(head.map((h,i)=>[h,values[idx][i]]));
  if(old.status==='CANCELLED') throw new Error('รายการนี้ถูกยกเลิกแล้ว'); if(!canManage_(user)&&old.createdBy!==user.id) throw new Error('ยกเลิกได้เฉพาะรายการที่ตนเองบันทึก'); sh.getRange(idx+1,12).setValue('CANCELLED');
  if(old.referenceId){ const ps=sheet_('Purchases'), pv=ps.getDataRange().getValues(), pi=pv.findIndex((r,i)=>i>0&&r[0]===old.referenceId); if(pi>0) ps.getRange(pi+1,10).setValue('CANCELLED'); invalidateRows_('Purchases'); }
  SpreadsheetApp.flush(); invalidateRows_('Transactions'); audit_(user,'CANCEL','TRANSACTION',data.id,clean_(data.reason)||'ยกเลิกรายการ'); return ok({refresh:bootstrap_(user).data});
}

function updatePurchaseHeader_(id,date,accountId,note){ const sh=sheet_('Purchases'), v=sh.getDataRange().getValues(), i=v.findIndex((r,n)=>n>0&&r[0]===id); if(i>0){sh.getRange(i+1,2).setValue(date);sh.getRange(i+1,4).setValue(accountId);sh.getRange(i+1,6).setValue(note);invalidateRows_('Purchases');} }

function createUser_(user, data) {
  if (!canAdminUsers_(user)) throw new Error('เฉพาะเจ้าของร้านหรือแอดมินเท่านั้น');
  validateUserInput_(data);
  const username = clean_(data.username).toLowerCase();
  if (rows_('Users').some(x => String(x.username).toLowerCase() === username)) throw new Error('ชื่อผู้ใช้นี้มีแล้ว');
  const role = ['STAFF','MANAGER','ADMIN','OWNER'].includes(data.role) ? data.role : 'STAFF';
  if (['OWNER','ADMIN'].includes(role) && user.role !== 'OWNER') throw new Error('เฉพาะเจ้าของร้านที่สร้างบัญชีเจ้าของหรือแอดมินได้');
  const id = id_('USR');
  append_('Users',[id,username,hash_(data.password),clean_(data.name),role,true,now_(),user.id]);
  invalidateRows_('Users');
  audit_(user,'CREATE','USER',id,username+' / '+role);
  return ok({ id,refresh:bootstrap_(user).data });
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
  sh.getRange(idx+1,4).setValue(name); sh.getRange(idx+1,5).setValue(role); SpreadsheetApp.flush();invalidateRows_('Users'); const updated=rows_('Users').find(x=>x.id===data.id);invalidateRows_('Users');audit_(user,'UPDATE','USER',data.id,username+' / '+role); return ok({refresh:bootstrap_(data.id===user.id?{...publicUser_(updated),sessionVersion:user.sessionVersion}:user).data});
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
  if(props.getProperty('SCHEMA_VERSION')==='8') return;
  const ss=db_();
  Object.keys(APP.sheets).forEach(name=>{
    let sh=ss.getSheetByName(name);
    if(!sh){ sh=ss.insertSheet(name); sh.getRange(1,1,1,APP.sheets[name].length).setValues([APP.sheets[name]]).setFontWeight('bold').setBackground('#f6c90e'); sh.setFrozenRows(1); }
    else { const current=sh.getRange(1,1,1,Math.max(1,sh.getLastColumn())).getValues()[0]; APP.sheets[name].forEach(h=>{ if(!current.includes(h)){ sh.getRange(1,sh.getLastColumn()+1).setValue(h).setFontWeight('bold').setBackground('#f6c90e'); current.push(h); } }); }
  });
  const pc=sheet_('ProductCategories');
  if(pc.getLastRow()===1){ const seed=productCategorySeed_(); pc.getRange(2,1,seed.length,seed[0].length).setValues(seed); }
  const units=sheet_('Units');if(units.getLastRow()===1){const seed=unitSeed_();units.getRange(2,1,seed.length,seed[0].length).setValues(seed)}
  const cats=sheet_('Categories');if(cats.getLastRow()>1){const vals=cats.getRange(2,5,cats.getLastRow()-1,1).getValues().map((r,i)=>[num_(r[0])||i+1]);cats.getRange(2,5,vals.length,1).setValues(vals)}
  const catRows=cats.getDataRange().getValues();if(!catRows.slice(1).some(r=>r[0]==='CAT-DIVIDEND'))append_('Categories',['CAT-DIVIDEND','เงินปันผลเจ้าของร้าน','EXPENSE',true,Math.max(0,...catRows.slice(1).map(r=>num_(r[4])))+1]);
  invalidateRows_('Categories');
  props.setProperty('SCHEMA_VERSION','8');
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
function requireSession_(token) { if(!token)throw new Error('SESSION_EXPIRED');const cache=CacheService.getScriptCache(),props=PropertiesService.getScriptProperties(),key='SESSION_TOKEN_'+token,raw=cache.get('session:'+token)||props.getProperty(key);if(!raw)throw new Error('SESSION_EXPIRED');let user=JSON.parse(raw);const version=props.getProperty('SESSION_VERSION')||'1';if(user.sessionVersion!==version){props.deleteProperty(key);throw new Error('SESSION_EXPIRED')}const latest=cachedRows_('Users',120).find(x=>x.id===user.id);if(!latest||!truthy_(latest.active)){props.deleteProperty(key);throw new Error('SESSION_EXPIRED')}const needsPersist=!user.persistent,fresh={...publicUser_(latest),sessionVersion:version,persistent:true};if(needsPersist||fresh.name!==user.name||fresh.username!==user.username||fresh.role!==user.role){user=fresh;props.setProperty(key,JSON.stringify(user))}cache.put('session:'+token,JSON.stringify(user),21600);return user; }
function publicUser_(u) { return {id:u.id,name:u.name,username:u.username,role:u.role,active:truthy_(u.active)}; }
function canManage_(u) { return ['ADMIN','MANAGER','OWNER'].includes(u.role); }
function canAdminUsers_(u) { return ['ADMIN','OWNER'].includes(u.role); }
function canViewFinance_(u) { return ['ADMIN','OWNER'].includes(u.role); }
function validateUserInput_(d) { if(clean_(d.name).length<2) throw new Error('กรุณาระบุชื่อ'); if(!/^[a-zA-Z0-9._-]{3,30}$/.test(clean_(d.username))) throw new Error('ชื่อผู้ใช้ต้องเป็นอังกฤษหรือตัวเลขอย่างน้อย 3 ตัว'); if(String(d.password||'').length<6) throw new Error('รหัสผ่านต้องมีอย่างน้อย 6 ตัว'); }
function hash_(s) { return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,String(s))); }
let DB_CACHE_;
let ROWS_MEMO_ = {};
function db_() { if(DB_CACHE_) return DB_CACHE_; DB_CACHE_=SpreadsheetApp.openById(PRIMARY_DB_ID); return DB_CACHE_; }
function sheet_(name) { const sh=db_().getSheetByName(name); if(!sh) throw new Error('ไม่พบชีต '+name); return sh; }
function rows_(name) { const sh=sheet_(name), vals=sh.getDataRange().getValues(), head=vals.shift(); return vals.filter(r=>r.some(v=>v!=='' )).map(r=>Object.fromEntries(head.map((h,i)=>[h,r[i]]))); }
function cachedRows_(name, seconds) { if(Object.prototype.hasOwnProperty.call(ROWS_MEMO_,name))return ROWS_MEMO_[name];const cache=CacheService.getScriptCache(), key='rows:'+name, hit=cache.get(key); if(hit)return ROWS_MEMO_[name]=JSON.parse(hit); const data=rows_(name);ROWS_MEMO_[name]=data; const raw=JSON.stringify(data); if(raw.length<95000) cache.put(key,raw,seconds||300); return data; }
function invalidateRows_(name) { delete ROWS_MEMO_[name];CacheService.getScriptCache().remove('rows:'+name);PropertiesService.getScriptProperties().setProperty('DATA_REVISION',String(Date.now())); }
function append_(name,row) { sheet_(name).appendRow(row); }
function appendRows_(name, rows) { if(!rows.length) return; const sh=sheet_(name); sh.getRange(sh.getLastRow()+1,1,rows.length,rows[0].length).setValues(rows); }
function audit_(u,a,e,id,d) { append_('AuditLog',[now_(),u.id,u.name,a,e,id,d]);PropertiesService.getScriptProperties().setProperty('DATA_REVISION',String(Date.now())); }
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
function truthy_(v) { return v===true || String(v).toLowerCase()==='true'; }
function ok(data) { return {ok:true,data:data||{}}; }
