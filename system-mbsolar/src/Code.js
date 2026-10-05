const APP = {
  version: '2.33.0',
  maintenanceMode: true,
  spreadsheetId: '15wNWYPRNU3ozpuaQI4r_g2rYfwKSRwvWZV93jZGrDl4',
  sheets: {
    Users: ['id','name','username','passwordHash','role','permissions','active','createdAt'],
    Projects: ['id','name','customer','contractValue','vatRate','status','startDate','endDate','note','createdAt','createdBy','siteLocation','province','projectType','deliveryDate'],
    Income: ['id','projectId','date','description','amountExVat','vat','total','account','createdAt','createdBy'],
    Expenses: ['id','projectId','date','category','description','amountExVat','vat','total','paidBy','taxInvoice','createdAt','createdBy'],
    Employees: ['id','name','type','wageRate','otRate','active','note','createdAt','position'],
    Overtime: ['id','employeeId','projectId','date','hours','rate','amount','note','paid','createdAt','createdBy'],
    SalaryAdvances: ['id','employeeId','monthKey','date','amount','note','createdAt','createdBy'],
    SalaryPayments: ['id','employeeId','monthKey','dueDate','baseSalary','advanceTotal','netPaid','paidDate','note','createdAt','createdBy','otTotal'],
    Banks: ['id','name','plan','interestType','annualRate','minYears','maxYears','active'],
    Categories: ['id','name','active','sortOrder','createdAt','createdBy'],
    PaymentMethods: ['id','name','active','sortOrder','createdAt','createdBy'],
    BalanceAccounts: ['id','name','type','openingBalance','active','sortOrder','createdAt','createdBy'],
    Partners: ['id','name','active','createdAt'],
    Dividends: ['id','partnerId','date','amount','paidFrom','note','createdAt','createdBy','projectId'],
    Quotations: ['id','quoteNo','issueDate','customerName','customerPhone','customerAddress','customerTaxId','workType','itemsJson','vatRate','paymentTerms','deliveryTerms','warrantyText','scopeText','note','status','subtotal','vatAmount','grandTotal','createdAt','updatedAt','createdBy','logoData','noteColor','companyJson','projectName'],
    QuotePresets: ['id','kind','workType','name','dataJson','updatedAt','createdBy'],
    Invoices: ['id','invoiceNo','issueDate','quotationId','quoteNo','installment','percent','subtotal','vatRate','vatAmount','grandTotal','paymentTerms','deliveryTerms','poNo','snapshotJson','itemsJson','logoData','createdAt','createdBy'],
    Receipts: ['id','receiptNo','issueDate','invoiceId','invoiceNo','quotationId','quoteNo','installment','subtotal','vatRate','vatAmount','grandTotal','snapshotJson','itemsJson','logoData','createdAt','createdBy'],
    Attendance: ['id','employeeId','employeeName','date','checkInTime','checkOutTime','locationName','latitude','longitude','photoData','penaltyAmount','penaltyReason','otHours','otAmount','status','note','createdAt','createdBy'],
    ProductCosts: ['id','code','name','category','subCategory','brand','model','supplier','unit','unitPrice','minPrice','maxPrice','lastPurchaseDate','note','createdAt','createdBy'],
    AuditLog: ['id','date','userId','action','entity','entityId','detail']
  }
};

function doGet() {
  try {
    ensureSchemaOnce_();
    const initialData = JSON.stringify(initialClientPayload_(getBootstrapData_())).replace(/</g, '\\u003c');
    const source = HtmlService.createTemplateFromFile('Index').getRawContent().replace('<?!= initialData ?>', () => initialData);
    const html = source.replace(/<script>([\s\S]*?)<\/script>/g, (_, script) => {
      const encoded = Utilities.base64Encode(script, Utilities.Charset.UTF_8);
      return '<script>const appScript=document.createElement("script");appScript.textContent=new TextDecoder().decode(Uint8Array.from(atob("' + encoded + '"),c=>c.charCodeAt(0)));document.head.appendChild(appScript);</script>';
    });
    return HtmlService.createHtmlOutput(html).setTitle('MB Solar Project Profit')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  } catch (err) {
    return HtmlService.createHtmlOutput('<div style="padding:24px;font-family:sans-serif;color:#c00;"><h2>เกิดข้อผิดพลาดในการโหลดหน้าเว็บ</h2><p>' + (err.stack || err.message || String(err)) + '</p></div>')
      .setTitle('MB Solar - Error')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
  }
}

function getBootstrapData_() {
  const cacheKey = 'MB_BOOTSTRAP_CACHE_V3';
  try {
    if (typeof CacheService !== 'undefined' && CacheService.getScriptCache()) {
      const cached = CacheService.getScriptCache().get(cacheKey);
      if (cached) return JSON.parse(cached);
    }
  } catch (e) {}
  const data = bootstrap_(maintenanceUser_());
  try {
    if (typeof CacheService !== 'undefined' && CacheService.getScriptCache()) {
      const payloadStr = JSON.stringify(data);
      if (payloadStr.length < 95000) {
        CacheService.getScriptCache().put(cacheKey, payloadStr, 21600);
      }
    }
  } catch (e) {}
  return data;
}

function clearBootstrapCache_() {
  try {
    if (typeof CacheService !== 'undefined' && CacheService.getScriptCache()) {
      CacheService.getScriptCache().remove('MB_BOOTSTRAP_CACHE_V3');
    }
  } catch (e) {}
}

function api(action, payload) {
  ensureSchemaOnce_();
  payload = payload || {};
  if (action === 'login') return clientSafe_(login_(payload));
  const user = APP.maintenanceMode && !payload.token ? maintenanceUser_() : sessionUser_(payload.token);
  if (action === 'bootstrap') return clientSafe_(bootstrap_(user));
  if (action === 'createInvoice') { const res = createInvoice_(user, payload.data || {}); clearBootstrapCache_(); return clientSafe_(res); }
  if (action === 'updateInvoice') { const res = updateInvoice_(user, payload.data || {}); clearBootstrapCache_(); return clientSafe_(res); }
  if (action === 'deleteInvoice') { const res = deleteInvoice_(user, payload.data || {}); clearBootstrapCache_(); return clientSafe_(res); }
  if (action === 'createReceipt') { const res = createReceipt_(user, payload.data || {}); clearBootstrapCache_(); return clientSafe_(res); }
  if (action === 'updateReceipt') { const res = updateReceipt_(user, payload.data || {}); clearBootstrapCache_(); return clientSafe_(res); }
  if (action === 'deleteReceipt') { const res = deleteReceipt_(user, payload.data || {}); clearBootstrapCache_(); return clientSafe_(res); }
  const routes = {saveProject:saveProject_,deleteProject:deleteProject_,saveIncome:saveIncome_,saveExpense:saveExpense_,saveExpensesBatch:saveExpensesBatch_,saveTransactionsBatch:saveTransactionsBatch_,updateTransaction:updateTransaction_,deleteTransaction:deleteTransaction_,saveEmployee:saveEmployee_,deleteEmployee:deleteEmployee_,saveOvertime:saveOvertime_,saveSalaryAdvance:saveSalaryAdvance_,payMonthlySalary:payMonthlySalary_,distributeProjectDividend:distributeProjectDividend_,saveQuotation:saveQuotation_,saveQuotePreset:saveQuotePreset_,deleteQuotePreset:deleteQuotePreset_,deleteQuotation:deleteQuotation_,saveUser:saveUser_,saveCategory:saveCategory_,deleteCategory:deleteCategory_,reorderCategories:reorderCategories_,savePaymentMethod:savePaymentMethod_,deletePaymentMethod:deletePaymentMethod_,reorderPaymentMethods:reorderPaymentMethods_,saveBalanceAccount:saveBalanceAccount_,deleteBalanceAccount:deleteBalanceAccount_,saveAttendance:saveAttendance_,checkInAttendance:checkInAttendance_,checkOutAttendance:checkOutAttendance_,deleteAttendance:deleteAttendance_,saveProductCost:saveProductCost_,deleteProductCost:deleteProductCost_,saveCostEstimateToExpenses:saveCostEstimateToExpenses_};
  if (!routes[action]) throw new Error('ไม่พบคำสั่งที่ร้องขอ');
  const result = routes[action](user, payload.data || {});
  clearBootstrapCache_();
  return clientSafe_(result);
}

function ensureSchemaOnce_(){
  const props=PropertiesService.getScriptProperties();if(props.getProperty('MB_SCHEMA_VERSION')==='29')return;
  const ss=db_();Object.keys(APP.sheets).forEach(name=>{let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);const expected=APP.sheets[name];if(sh.getLastRow()===0){sh.getRange(1,1,1,expected.length).setValues([expected]);sh.setFrozenRows(1)}else{const lastCol=Math.max(1,sh.getLastColumn()),existing=sh.getRange(1,1,1,lastCol).getValues()[0];expected.filter(h=>existing.indexOf(h)<0).forEach(h=>{sh.getRange(1,sh.getLastColumn()+1).setValue(h)})}sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#dff5e8')});
  if(readSheet_(ss,'Partners').length===0){append_('Partners',['PART-BAS','บาส',true,new Date()]);append_('Partners',['PART-GOLF','กอล์ฟ',true,new Date()])}
  if(readSheet_(ss,'BalanceAccounts').length===0){append_('BalanceAccounts',['BAL-CASH','เงินสด','CASH',0,true,1,new Date(),'SYSTEM']);append_('BalanceAccounts',['BAL-COMPANY','บัญชีบริษัท','COMPANY',0,true,2,new Date(),'SYSTEM'])}
  seedEvPriceSets_();
  fillEvPackageDescriptions_();
  splitEvPackageRows_();
  seedSolarPriceSets_();
  migrateSolarCompany_();
  migrateSolarPanelCounts_();
  migrateSolarBank_();
  migrateSavedSolarBanks_();
  migrateInvoiceNumbers_();
  seedProductCosts_();
  props.setProperty('MB_SCHEMA_VERSION','29');
}

function migrateSavedSolarBanks_(){const props=PropertiesService.getScriptProperties(),key='SAVED_SOLAR_BANKS_V1',lock=LockService.getScriptLock();lock.waitLock(30000);try{if(props.getProperty(key)==='done')return;const bank=solarCompanyDefaults_();rows_('Quotations').filter(x=>x.workType==='Solar Cell').forEach(row=>{const company=JSON.parse(row.companyJson||'{}');if(String(company.bankAccount||'').trim())return;append_('AuditLog',[id_('AUD'),new Date(),'SYSTEM','BACKUP','QUOTATION_COMPANY',row.id,row.companyJson||'{}']);const updated={...bank,...company,bankName:bank.bankName,bankAccount:bank.bankAccount,accountName:bank.accountName};updateRow_('Quotations',row.id,{companyJson:JSON.stringify(updated),updatedAt:new Date()})});props.setProperty(key,'done')}finally{lock.releaseLock()}}

function solarPanelCount_(kw,panel=''){const fixed={3:5,5:8,10:16}[Number(kw)],watt=String(panel).match(/(\d+(?:\.\d+)?)\s*W\b/i);return fixed||(Number(kw)>0&&watt&&Number(watt[1])>0?Math.ceil(Number(kw)*1000/Number(watt[1])):0)}
function migrateSolarPanelCounts_(){const props=PropertiesService.getScriptProperties(),key='SOLAR_PANEL_COUNTS_V2',lock=LockService.getScriptLock();lock.waitLock(30000);try{if(props.getProperty(key)==='done')return;rows_('QuotePresets').filter(x=>x.kind==='SET'&&x.workType==='Solar Cell').forEach(row=>{const data=JSON.parse(row.dataJson||'{}');let changed=false;(data.items||[]).forEach(item=>{const description=String(item.description||''),match=description.split('\n')[0].match(/Inverter\s+\S+\s+(\d+)\s*kW/i),count=match?solarPanelCount_(match[1],(description.match(/PV Module ([^\n]*)/)||[])[1]||''):0;if(!count)return;const next=description.replace(/(PV Module [^\n]*?) \(จำนวนแผงตามแบบที่ยืนยัน\)/g,'$1 — '+count+' แผง');if(next!==description){item.description=next;changed=true}});if(changed){append_('AuditLog',[id_('AUD'),new Date(),'SYSTEM','BACKUP','QUOTE_PRESET',row.id,row.dataJson]);updateRow_('QuotePresets',row.id,{dataJson:JSON.stringify(data),updatedAt:new Date()})}});props.setProperty(key,'done')}finally{lock.releaseLock()}}

function solarCompanyDefaults_(){return {name:'ห้างหุ้นส่วนจำกัด เอ็มบี โซล่า (สำนักงานใหญ่)',phone:'095-238-3016',email:'',taxId:'0403566004363',address:'87 หมู่ที่ 1 ตำบลหนองแวงนางเบ้า อำเภอพล จ.ขอนแก่น 40120',bankName:'ธนาคารกสิกรไทย',bankAccount:'232-2-79355-6',accountName:'ห้างหุ้นส่วนจำกัด เอ็ม บี โซล่า',signerName:'นาย ณัฐฤทธิ์ นนทโคตร',signatureData:''}}
function migrateSolarBank_(){const props=PropertiesService.getScriptProperties(),key='SOLAR_BANK_KBANK_2322793556_V1',lock=LockService.getScriptLock();lock.waitLock(30000);try{if(props.getProperty(key)==='done')return;const template=rows_('QuotePresets').find(x=>x.kind==='TEMPLATE'&&x.workType==='Solar Cell');if(!template)throw new Error('ไม่พบฟอร์ม Solar Cell สำหรับเพิ่มบัญชี');const data=JSON.parse(template.dataJson||'{}'),defaults=solarCompanyDefaults_();append_('AuditLog',[id_('AUD'),new Date(),'SYSTEM','BACKUP','QUOTE_PRESET',template.id,template.dataJson]);data.company={...(data.company||defaults),bankName:defaults.bankName,bankAccount:defaults.bankAccount,accountName:defaults.accountName};updateRow_('QuotePresets',template.id,{dataJson:JSON.stringify(data),updatedAt:new Date()});props.setProperty(key,'done')}finally{lock.releaseLock()}}
function migrateSolarCompany_(){const props=PropertiesService.getScriptProperties(),key='SOLAR_COMPANY_MB_V1',lock=LockService.getScriptLock();lock.waitLock(30000);try{if(props.getProperty(key)==='done')return;const template=rows_('QuotePresets').find(x=>x.kind==='TEMPLATE'&&x.workType==='Solar Cell');if(template){const data=JSON.parse(template.dataJson||'{}');append_('AuditLog',[id_('AUD'),new Date(),'SYSTEM','BACKUP','QUOTE_PRESET',template.id,String(template.dataJson||'{}')]);data.company=solarCompanyDefaults_();updateRow_('QuotePresets',template.id,{dataJson:JSON.stringify(data),updatedAt:new Date()})}else append_('QuotePresets',['QPR-TEMPLATE-SOLAR-ROOF','TEMPLATE','Solar Cell','Solar Cell',JSON.stringify({...solarTemplateDefaults_(),company:solarCompanyDefaults_()}),new Date(),'SYSTEM']);props.setProperty(key,'done')}finally{lock.releaseLock()}}

function solarPriceCatalog_(){return [
  ...[[3,1,99500],[5,1,129000],[5,3,149000],[10,1,229000],[10,3,239000],[15,3,329000],[20,3,369000],[30,3,429000],[50,3,899000],[100,3,1690000],[199,3,2990000]].map(([kw,phase,price])=>({type:'On-grid',brand:'Huawei',kw,phase,price,panel:kw>=30?'Trina 720 W':'AIKO 670 W'})),
  ...[[6,1,249000],[8,1,259000],[10,1,329000],[10,3,359000],[16,3,429000],[20,3,539000]].map(([kw,phase,price])=>({type:'Hybrid On/Off-grid',brand:'Solis',kw,phase,price,panel:'AIKO 670 W'}))
]}
function solarPackageItem_(x){const hybrid=x.type==='Hybrid On/Off-grid',title='งานติดตั้งระบบ Solar Roof '+x.type+' Inverter '+x.brand+' '+x.kw+' kW '+x.phase+' Phase'+(hybrid?' + Battery Solis 16 kWh':''),details=[
  'PV Module '+x.panel+(solarPanelCount_(x.kw,x.panel)?' — '+solarPanelCount_(x.kw,x.panel)+' แผง':' (จำนวนแผงตามแบบที่ยืนยัน)'),
  'Inverter '+x.brand+' '+x.type+' '+x.kw+' kW '+x.phase+' Phase — 1 เครื่อง',
  ...(hybrid?['Battery Solis 314Ah 51.2V IntelliHome-16kWh-OD (IP66) ขนาด 16 kWh — 1 ชุด']:[]),
  'DC Cable (Link) และ AC Cable (Yazaki - Thai)',
  'Mounting Structure, Raceway และ Wireway',
  'DC Fuse Cabinet และ Solar AC Cabinet',
  'Ground Systems ระบบสายดิน DC และ AC',
  'Zero Export System ระบบกันย้อนตามมาตรฐาน PEA',
  'วิศวกรควบคุมงาน จัดทำ SLD และลงนามรับรองตามมาตรฐาน MEA / PEA',
  'ดำเนินการขออนุญาตการไฟฟ้า และแอปติดตามสถานะ Solar Cell',
  'ติดตั้ง ทดสอบระบบ และส่งมอบงาน'
];return {description:title+'\nรายการย่อยที่ลูกค้าจะได้รับ\n'+details.map(t=>'• '+t).join('\n'),quantity:1,unit:'งาน',unitPrice:x.price,packagePrice:true,packageIncluded:false}}
function solarTemplateDefaults_(){return {company:solarCompanyDefaults_(),vatRate:7,noteColor:'black',paymentTerms:'งวดที่ 1: 50% เมื่อออกใบ PO และนัดวันติดตั้ง · งวดที่ 2: 50% หลังติดตั้งออนระบบและส่งมอบงานแล้วเสร็จ ก่อนยื่นขนานไฟจาก PEA',deliveryTerms:'ระยะดำเนินการตามตกลงของหน้างานและสัญญาโครงการ',warrantyText:'รับประกันงานติดตั้ง 3 ปี ฟรีล้างแผงและ PM 3 ปี (ปีละ 1 ครั้ง) กรณีเร่งด่วนที่แก้ไขทางโทรศัพท์ไม่ได้ เข้าบริการ Onsite ภายใน 48 ชั่วโมง บริการออนไลน์ผ่านโทรศัพท์ตลอดอายุการใช้งาน การรับประกันแผง Inverter และแบตเตอรี่ตามรุ่นและเงื่อนไขผู้ผลิต',scopeText:'อุปกรณ์และงานเดินสายตามมาตรฐานการไฟฟ้าและวิศวกรรม ภายนอกใช้ท่อ IMC หรือท่ออ่อนกันน้ำ ภายในใช้ท่อ UPVC ราง UPVC หรือ Wireway อลูมิเนียม จัดทำ SLD พร้อมวิศวกรลงนามรับรองตาม MEA / PEA ฟรีค่าดำเนินการขออนุญาตการไฟฟ้า บริษัทจัดหาอุปกรณ์ ติดตั้ง และทดสอบระบบก่อนส่งมอบ ใบเสนอราคามีอายุ 7 วัน ห้ามนำข้อมูลไปใช้หรือเผยแพร่โดยไม่ได้รับอนุญาต',note:'หลังพ้นรับประกันงานติดตั้ง 3 ปี การเข้า Service Onsite หรือเปลี่ยนอุปกรณ์มีค่าดำเนินการ'}}
function seedSolarPriceSets_(){const props=PropertiesService.getScriptProperties(),key='SOLAR_PRICE_SETS_20260912_V1',lock=LockService.getScriptLock();lock.waitLock(30000);try{if(props.getProperty(key)==='done')return;const existing=rows_('QuotePresets'),ids=new Set(existing.map(x=>x.id));solarPriceCatalog_().forEach(x=>{const hybrid=x.type==='Hybrid On/Off-grid',id='QPR-SOLAR-'+(hybrid?'HYBRID':'ONGRID')+'-'+x.kw+'-'+x.phase;if(ids.has(id))return;const name=x.type+' · '+x.brand+' '+x.kw+' kW · '+x.phase+' Phase'+(hybrid?' + Battery 16 kWh':'');append_('QuotePresets',[id,'SET','Solar Cell',name,JSON.stringify({items:[solarPackageItem_(x)]}),new Date(),'SYSTEM'])});if(!existing.some(x=>x.kind==='TEMPLATE'&&x.workType==='Solar Cell'))append_('QuotePresets',['QPR-TEMPLATE-SOLAR-ROOF','TEMPLATE','Solar Cell','Solar Cell',JSON.stringify(solarTemplateDefaults_()),new Date(),'SYSTEM']);props.setProperty(key,'done')}finally{lock.releaseLock()}}

function quotationCompany_(d){const result={};['name','phone','email','taxId','address','bankName','bankAccount','accountName','signerName'].forEach(k=>{const field='company_'+k;if(Object.prototype.hasOwnProperty.call(d,field))result[k]=String(d[field]??'').trim().slice(0,k==='address'?1000:200)});if(Object.prototype.hasOwnProperty.call(d,'company_signatureData')){const image=String(d.company_signatureData||'');if(image&&(image.length>20000||!/^data:image\/(png|webp|jpeg);base64,[A-Za-z0-9+/=]+$/.test(image)))throw new Error('ไฟล์ลายเซ็นไม่ถูกต้องหรือมีขนาดใหญ่เกินไป');result.signatureData=image}return result}

function quotationLineAmount_(x){return round_(x.packagePrice===true?x.unitPrice:x.quantity*x.unitPrice)}
function evPackageItems_(kw,count,total){
  const lines=evPackageDescription_(kw,count).split('\n').slice(1);
  const units=['เครื่อง','ชุด','ชุด','งาน','ระบบ','งาน','งาน'];
  return lines.map((text,i)=>({description:text.replace(/^\d+\. /,'').replace(/ — \d+ \S+$/,''),quantity:i===0?count:1,unit:units[i],unitPrice:i===0?total:0,packagePrice:i===0,packageIncluded:i!==0}));
}
function splitEvPackageRows_(){
  const props=PropertiesService.getScriptProperties(),key='EV_PACKAGE_ROWS_V1',lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    if(props.getProperty(key)==='done')return;
    rows_('QuotePresets').forEach(row=>{
      const m=String(row.id).match(/^QPR-EV-FLEXXFAST-(120|150|180|240)-([1-4])$/);
      if(!m||row.kind!=='SET'||row.workType!=='EV Charger')return;
      const data=JSON.parse(row.dataJson||'{}');
      if(data.evRowsVersion===1)return;
      if(!Array.isArray(data.items)||!data.items.length)throw new Error('ไม่พบรายการในชุด '+row.name);
      const total=round_(data.items.reduce((s,x)=>s+quotationLineAmount_(x),0));
      append_('AuditLog',[id_('AUD'),new Date(),'SYSTEM','BACKUP_BEFORE_SPLIT','QUOTE_PRESET',row.id,row.dataJson]);
      data.items=evPackageItems_(Number(m[1]),Number(m[2]),total);data.evRowsVersion=1;
      updateRow_('QuotePresets',row.id,{dataJson:JSON.stringify(data),updatedAt:new Date()});
    });
    props.setProperty(key,'done');
  }finally{lock.releaseLock()}
}

function evPackageDescription_(kw,count){
  if(![120,150,180,240].includes(kw)||![1,2,3,4].includes(count))throw new Error('รุ่นหรือจำนวนเครื่องไม่ถูกต้อง');
  const heads=count*2;
  return [
    'Flexxfast '+kw+' kW จำนวน '+count+' เครื่อง',
    '1. ตู้ชาร์จ Flexxfast '+kw+' kW ต่อเครื่อง แบบ Dual CCS2 รวม '+heads+' หัวจ่าย พร้อมเราเตอร์ 4G/5G และเชื่อมต่อ CMS ตามสเปกที่ยืนยัน — '+count+' เครื่อง',
    '2. หม้อแปลงเฉพาะสถานี 3 Phase พร้อมอุปกรณ์แรงสูง ขนาดตามโหลดรวม '+(kw*count)+' kW และแบบวิศวกรอนุมัติ — 1 ชุด',
    '3. ตู้ Outdoor MDB (IP65) พร้อม Main Breaker, SPD และ RCD Type B ตามแบบระบบป้องกัน — 1 ชุด',
    '4. งานเดินสายกำลังใต้ดิน XLPE/NYY ท่อ HDPE และระบบสายดินไม่เกิน 5 โอห์ม หรือตามเกณฑ์ที่เข้มงวดกว่าของแบบอนุมัติ — 1 งาน',
    '5. ระบบ CCTV IP Camera พร้อมอุปกรณ์บันทึกภาพ ตามแบบสถานี — 1 ระบบ',
    '6. งานโยธาและช่องจอด EV: ฐานรากตู้ชาร์จ '+count+' ฐาน เสากันชนเหล็กเบื้องต้น '+heads+' ต้น ยางกันล้อ '+heads+' ชิ้น และตีเส้นพร้อมสัญลักษณ์ช่องจอด '+heads+' ช่อง ตามผังที่ตกลง — 1 งาน',
    '7. ค่าแรงติดตั้ง ทดสอบระบบ Commissioning และเอกสารรับรองโดยวิศวกรไฟฟ้า กว. ตามขอบเขตงาน — 1 งาน'
  ].join('\n');
}

// Previous generated text retained only to recognize safe migration targets.
function evPackageDescriptionV1_(kw,count){
  if(![120,150,180,240].includes(kw)||![1,2,3,4].includes(count))throw new Error('รุ่นหรือจำนวนเครื่องไม่ถูกต้อง');
  const heads=count*2,load=kw*count,kva=Math.ceil(load/0.95/0.95*1.25);
  return [
    'แพ็กเกจ Turnkey Flexxfast '+kw+' kW จำนวน '+count+' เครื่อง (ราคาเหมารวม 1 ชุด)',
    '1. ตู้ชาร์จ DC Flexxfast กำลังรวม '+kw+' kW ต่อเครื่อง จำนวน '+count+' เครื่อง หัวจ่าย Dual CCS2 รวม '+heads+' หัว พร้อมเราเตอร์ 4G/5G และตั้งค่าเชื่อมต่อ CMS ตามรุ่นอุปกรณ์ที่ยืนยัน',
    '2. หม้อแปลงเฉพาะสถานี 3 Phase จำนวน 1 ชุด รองรับกำลังชาร์จรวม '+load+' kW พร้อมอุปกรณ์แรงสูงตามแบบการไฟฟ้าอนุมัติ; กำลังหม้อแปลงคำนวณเบื้องต้นประมาณ '+kva+' kVA (สมมติประสิทธิภาพ 95%, PF 0.95, เผื่อ 25%; ไม่ใช่สเปกยืนยันของเครื่อง) ขนาดพิกัดติดตั้งและโหลดประกอบต้องให้วิศวกรยืนยัน',
    '3. ตู้ Outdoor MDB (IP65) จำนวน 1 ชุด พร้อม Main Breaker, SPD และระบบป้องกันไฟรั่ว RCD Type B ตามแบบและข้อกำหนดผู้ผลิต',
    '4. งานสายไฟฟ้ากำลังใต้ดิน XLPE/NYY พร้อมท่อ HDPE และอุปกรณ์ จำนวน 1 งาน ขนาดสายและระยะทางตามแบบและขอบเขตที่ตกลง',
    '5. ระบบสายดินและการต่อประสานศักย์ จำนวน 1 ระบบ พร้อมตรวจวัดความต้านทานดินไม่เกิน 5 โอห์ม หรือตามเกณฑ์ที่เข้มงวดกว่าของแบบอนุมัติ',
    '6. CCTV IP Camera พร้อมอุปกรณ์บันทึกภาพ จำนวน 1 ระบบ จำนวนกล้องและตำแหน่งตามแบบสถานี',
    '7. โคมไฟสปอร์ตไลท์ LED พร้อมสายไฟและอุปกรณ์ติดตั้ง จำนวน 1 ชุด จำนวนโคมและตำแหน่งตามแบบแสงสว่าง',
    '8. ถังดับเพลิงประจำสถานีพร้อมป้าย จำนวน 1 ชุด ชนิด ขนาด และจำนวนถังตามแผนความปลอดภัย',
    '9. ฐานรากคอนกรีตสำหรับตู้ชาร์จ จำนวน '+count+' ฐาน พร้อมอุปกรณ์ยึดตามแบบโครงสร้าง',
    '10. เสากันชนเหล็ก เบื้องต้น '+heads+' ต้น (2 ต้นต่อเครื่อง) ยืนยันจำนวนและตำแหน่งตามผังหน้างาน',
    '11. ยางกันล้อ เบื้องต้น '+heads+' ชิ้น (1 ชิ้นต่อช่องจอด)',
    '12. งานทาสีตีเส้นและสัญลักษณ์ช่องจอด EV จำนวน '+heads+' ช่อง รองรับหัวชาร์จตามผังที่ตกลง',
    '13. ค่าแรงติดตั้งอุปกรณ์และระบบประกอบตามขอบเขตแพ็กเกจ จำนวน 1 งาน',
    '14. ทดสอบและเดินระบบ Commissioning ทดสอบการชาร์จและ CMS แนะนำการใช้งาน พร้อมเอกสารรับรอง/ผลทดสอบโดยวิศวกรไฟฟ้าผู้มีใบอนุญาต กว. ตามขอบเขตงาน จำนวน 1 งาน'
  ].join('\n');
}

// Fill only the original imported placeholder. Preserve custom text, prices,
// deleted packages, Solar templates and all previously saved quotations.
function fillEvPackageDescriptions_(){
  const props=PropertiesService.getScriptProperties(),key='EV_PACKAGE_DESCRIPTIONS_V2',lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    if(props.getProperty(key)==='done')return;
    rows_('QuotePresets').forEach(row=>{
      const match=String(row.id).match(/^QPR-EV-FLEXXFAST-(120|150|180|240)-([1-4])$/);
      if(!match||row.kind!=='SET'||row.workType!=='EV Charger')return;
      let data;try{data=JSON.parse(row.dataJson)}catch(e){return}
      const kw=Number(match[1]),count=Number(match[2]),placeholder='Flexxfast '+kw+' kW · '+count+' เครื่อง';
      if(!data||!Array.isArray(data.items)||data.items.length!==1)return;
      if(![placeholder,evPackageDescriptionV1_(kw,count)].includes(data.items[0].description))return;
      data.items[0].description=evPackageDescription_(kw,count);
      updateRow_('QuotePresets',row.id,{dataJson:JSON.stringify(data),updatedAt:new Date()});
    });
    props.setProperty(key,'done');
  }finally{lock.releaseLock()}
}

// User-supplied Canva DAHJ_hxMMwo: total package prices, excluding VAT.
// One-time import: subsequent edits/deletions are never restored on page load.
function seedEvPriceSets_(){
  const props=PropertiesService.getScriptProperties(),key='EV_PRICE_SETS_DAHJ_hxMMwo_V1';
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    if(props.getProperty(key)==='done')return;
    const prices={120:[1590000,2590000,3690000,4690000],150:[1790000,2890000,4190000,5290000],180:[2290000,3490000,4990000,5990000],240:[2490000,3990000,5590000,6990000]};
    const existing=new Set(rows_('QuotePresets').map(x=>x.id));
    Object.entries(prices).forEach(([kw,totals])=>totals.forEach((total,i)=>{
      const count=i+1,id='QPR-EV-FLEXXFAST-'+kw+'-'+count;
      if(existing.has(id))return;
      const name='Flexxfast '+kw+' kW · '+count+' เครื่อง';
      const data={items:[{description:name,quantity:1,unit:'ชุด',unitPrice:total}]};
      append_('QuotePresets',[id,'SET','EV Charger',name,JSON.stringify(data),new Date(),'SYSTEM']);
    }));
    props.setProperty(key,'done');
  }finally{lock.releaseLock()}
}

function maintenanceUser_(){
  const user=readSheet_(db_(),'Users').find(x=>String(x.username).toLowerCase()==='owner'&&truthy_(x.active));
  if(!user)throw new Error('ไม่พบบัญชีเจ้าของระบบ');
  return user;
}

function ensureSetup_() {
  const ss = db_();
  Object.keys(APP.sheets).forEach(name => {
    let sh = ss.getSheetByName(name); if (!sh) sh = ss.insertSheet(name);
    const headers = APP.sheets[name];
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,headers.length).setValues([headers]); sh.setFrozenRows(1);
      sh.getRange(1,1,1,headers.length).setFontWeight('bold').setBackground('#dff5e8');
    }
  });
  if (!rows_('Users').length) append_('Users',[id_('USR'),'เจ้าของระบบ','owner',hash_('1234'),'OWNER','*',true,new Date()]);
  if (!rows_('Banks').length) {
    append_('Banks',['BNK-01','ธนาคารตัวอย่าง A','สินเชื่อโซลาร์','REDUCING',4.5,1,10,true]);
    append_('Banks',['BNK-02','ธนาคารตัวอย่าง B','ผ่อนสบาย','FLAT',3.5,1,7,true]);
  }
  if (!rows_('Categories').length) ['ซื้อของ','ค่าแรง','ค่าที่พัก','ค่าน้ำมัน','ค่าอาหาร','อื่น ๆ'].forEach((name,i)=>append_('Categories',[id_('CAT'),name,true,i+1,new Date(),'SYSTEM']));
  if (!rows_('PaymentMethods').length) ['บัญชีบริษัท','เงินสด','เจ้าของสำรอง','ผู้ร่วมงานสำรอง'].forEach((name,i)=>append_('PaymentMethods',[id_('PAYBY'),name,true,i+1,new Date(),'SYSTEM']));
}

function db_() {
  return SpreadsheetApp.openById(APP.spreadsheetId);
}

function login_(data) {
  const username=clean_(data.username).toLowerCase();
  const user=readSheet_(db_(),'Users').find(x=>String(x.username).toLowerCase()===username&&truthy_(x.active));
  if(!user||user.passwordHash!==hash_(String(data.password||'')))throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  const token=Utilities.getUuid();
  CacheService.getScriptCache().put('session:'+token,JSON.stringify({id:user.id,name:user.name,username:user.username,role:user.role,permissions:user.permissions,active:true}),21600);
  return {token,data:bootstrap_(user)};
}

function sessionUser_(token) {
  if(!token)throw new Error('กรุณาเข้าสู่ระบบ');
  const raw=CacheService.getScriptCache().get('session:'+token); if(!raw)throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบอีกครั้ง');
  const user=JSON.parse(raw); CacheService.getScriptCache().put('session:'+token,raw,21600); return user;
}

function bootstrap_(user, startupOnly=false) {
  migrateLegacyFinanceChannelNames_(user);
  const ss=db_(),projects=readSheet_(ss,'Projects'),income=readSheet_(ss,'Income'),expenses=readSheet_(ss,'Expenses'),allEmployees=startupOnly?[]:readSheet_(ss,'Employees'),employeeRows=allEmployees.filter(x=>truthy_(x.active)),overtime=readSheet_(ss,'Overtime'),salaryAdvances=startupOnly?[]:readSheet_(ss,'SalaryAdvances'),salaryPayments=startupOnly?[]:readSheet_(ss,'SalaryPayments'),month=monthInfo_(),dividends=readSheet_(ss,'Dividends'),partnerRows=startupOnly?[]:readSheet_(ss,'Partners').filter(x=>truthy_(x.active));
  const partners=partnerRows.map(p=>Object.assign({},p,{totalPaid:round_(dividends.filter(x=>x.partnerId===p.id).reduce((s,x)=>s+num_(x.amount),0)),payments:dividends.filter(x=>x.partnerId===p.id).sort((a,b)=>String(b.date).localeCompare(String(a.date)))}));
  const employees=employeeRows.map(e=>{const advances=salaryAdvances.filter(x=>x.employeeId===e.id&&x.monthKey===month.key),advanceTotal=round_(advances.reduce((s,x)=>s+num_(x.amount),0)),employeeOt=overtime.filter(x=>x.employeeId===e.id&&String(x.date).slice(0,7)===month.key),otTotal=round_(employeeOt.reduce((s,x)=>s+num_(x.amount),0)),payment=salaryPayments.find(x=>x.employeeId===e.id&&x.monthKey===month.key),salary=num_(e.wageRate),earned=round_(salary+otTotal),remaining=payment?0:round_(Math.max(0,earned-advanceTotal)),progress=earned?round_(remaining/earned*100):0;return Object.assign({},e,{salary:{monthKey:month.key,dueDate:month.dueDate,baseSalary:salary,otTotal,totalEarned:earned,advanceTotal,remaining,progress,paid:!!payment,netPaid:payment?num_(payment.netPaid):0,advances}})});
  const projectSummaries=projects.map(p=>{
    const inc=income.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amountExVat),0);
    const exp=expenses.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amountExVat),0);
    const ot=overtime.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amount),0);
    const dividendPaid=dividends.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amount),0);
    return Object.assign({},p,{income:round_(inc),expenses:round_(exp+ot),profit:round_(inc-exp-ot),otCost:round_(ot),dividendPaid:round_(dividendPaid),profitAfterDividend:round_(inc-exp-ot-dividendPaid)});
  });
  const totals=projectSummaries.reduce((a,p)=>({contract:a.contract+num_(p.contractValue),income:a.income+p.income,expenses:a.expenses+p.expenses,profit:a.profit+p.profit}),{contract:0,income:0,expenses:0,profit:0});
  const banks=startupOnly?[]:readSheet_(ss,'Banks').filter(x=>truthy_(x.active)),categories=startupOnly?[]:sortActiveRows_(readSheet_(ss,'Categories')),paymentMethods=startupOnly?[]:sortActiveRows_(readSheet_(ss,'PaymentMethods')),balanceAccounts=sortActiveRows_(readSheet_(ss,'BalanceAccounts'));
  const employeeNames=allEmployees.reduce((o,e)=>(o[e.id]=e.name,o),{}),wageHistory=[...salaryPayments.map(x=>({id:x.id,type:'SALARY',date:x.paidDate||x.dueDate,monthKey:x.monthKey,employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',baseSalary:num_(x.baseSalary),otTotal:num_(x.otTotal),advanceTotal:num_(x.advanceTotal),amount:num_(x.netPaid),note:x.note||''})),...overtime.map(x=>({id:x.id,type:'OT',date:x.date,monthKey:String(x.date).slice(0,7),employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',projectId:x.projectId,hours:num_(x.hours),rate:num_(x.rate),amount:num_(x.amount),note:x.note||''})),...salaryAdvances.map(x=>({id:x.id,type:'ADVANCE',date:x.date,monthKey:x.monthKey,employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',amount:num_(x.amount),note:x.note||''}))].sort((a,b)=>String(b.date).localeCompare(String(a.date)));
  return {startupOnly,appVersion:APP.version,user:publicUser_(user),permissions:permissions_(user),totals,projects:projectSummaries,income,expenses,employees,overtime,wageHistory,banks,categories,paymentMethods,balanceAccounts,payrollMonth:month,partners,dividends,users:!startupOnly&&allowed_(user,'USERS_MANAGE')?readSheet_(ss,'Users').map(publicUser_):[],quotations:startupOnly?[]:quotationRows_(ss),quotePresets:startupOnly?[]:quotePresetRows_(ss),invoices:startupOnly?[]:invoiceRows_(ss),receipts:startupOnly?[]:receiptRows_(ss),attendance:startupOnly?[]:attendanceRows_(ss),productCosts:startupOnly?[]:productCostRows_(ss)};
}

function saveProject_(user,d) {
  require_(user,'PROJECT_EDIT'); const name=clean_(d.name),value=num_(d.contractValue),siteLocation=clean_(d.siteLocation),province=clean_(d.province),projectType=clean_(d.projectType)==='OTHER'?clean_(d.projectTypeOther):clean_(d.projectType),startDate=date_(d.startDate),deliveryDate=clean_(d.deliveryDate); if(!name||value<0)throw new Error('กรุณากรอกชื่อโปรเจกต์และมูลค่างานให้ถูกต้อง');if(!projectType)throw new Error('กรุณาเลือกหรือระบุประเภทงาน');if(deliveryDate&&deliveryDate<startDate)throw new Error('วันที่ส่งมอบต้องไม่ก่อนวันเริ่มงาน');
  if(d.id){const old=project_(d.id);updateRow_('Projects',d.id,{name,customer:clean_(d.customer),contractValue:value,status:clean_(d.status)||'PLANNING',startDate,endDate:old.endDate||'',note:clean_(d.note),siteLocation,province,projectType,deliveryDate});audit_(user,'UPDATE','PROJECT',old.id,name);return businessPayload_(db_())}
  const id=id_('PRJ'); append_('Projects',[id,name,clean_(d.customer),value,7,clean_(d.status)||'PLANNING',startDate,'',clean_(d.note),new Date(),user.id,siteLocation,province,projectType,deliveryDate]); audit_(user,'CREATE','PROJECT',id,name); return businessPayload_(db_());
}
function deleteProject_(user,d){require_(user,'PROJECT_EDIT');const p=project_(d.id),used=rows_('Income').some(x=>x.projectId===p.id)||rows_('Expenses').some(x=>x.projectId===p.id)||rows_('Overtime').some(x=>x.projectId===p.id);if(used)throw new Error('ลบไม่ได้ เพราะโปรเจกต์นี้มีรายการการเงินหรือ OT แล้ว');deleteRow_('Projects',p.id);audit_(user,'DELETE','PROJECT',p.id,p.name);return businessPayload_(db_())}
function saveIncome_(user,d) {
  require_(user,'FINANCE_EDIT'); project_(d.projectId); const amount=positive_(d.amount,'จำนวนเงิน'),vat=round_(amount*.07),id=id_('INC');
  append_('Income',[id,d.projectId,date_(d.date),clean_(d.description)||'รับเงินโครงการ',amount,vat,round_(amount+vat),clean_(d.account),new Date(),user.id]); audit_(user,'CREATE','INCOME',id,String(amount)); return businessPayload_(db_());
}
function saveExpense_(user,d) {
  require_(user,'EXPENSE_EDIT'); project_(d.projectId); const raw=positive_(d.amount,'จำนวนเงิน'),included=d.vatMode==='INCLUDED',ex=included?round_(raw/1.07):raw,vat=d.vatMode==='NONE'?0:round_(included?raw-ex:ex*.07),id=id_('EXP');
  append_('Expenses',[id,d.projectId,date_(d.date),clean_(d.category)||'อื่น ๆ',clean_(d.description),ex,vat,round_(ex+vat),clean_(d.paidBy),truthy_(d.taxInvoice),new Date(),user.id]); audit_(user,'CREATE','EXPENSE',id,String(ex)); return businessPayload_(db_());
}
function saveExpensesBatch_(user,d) {
  require_(user,'EXPENSE_EDIT');
  const items=Array.isArray(d.items)?d.items:[];
  if(!items.length)throw new Error('กรุณากรอกค่าใช้จ่ายอย่างน้อย 1 รายการ');
  if(items.length>200)throw new Error('บันทึกได้สูงสุดครั้งละ 200 รายการ');
  items.forEach((x,i)=>{
    project_(x.projectId);
    if(!clean_(x.description))throw new Error('กรุณากรอกรายละเอียดแถวที่ '+(i+1));
    const raw=positive_(x.amount,'ยอดเงินแถวที่ '+(i+1)),included=x.vatMode==='INCLUDED',ex=included?round_(raw/1.07):raw,vat=x.vatMode==='NONE'?0:round_(included?raw-ex:ex*.07),id=id_('EXP');
    append_('Expenses',[id,x.projectId,date_(x.date),clean_(x.category)||'อื่น ๆ',clean_(x.description),ex,vat,round_(ex+vat),clean_(x.paidBy),false,new Date(),user.id]);
    audit_(user,'CREATE','EXPENSE_BATCH',id,String(ex));
  });
  return businessPayload_(db_());
}
function saveTransactionsBatch_(user,d) {
  const items=Array.isArray(d.items)?d.items:[];
  if(!items.length)throw new Error('กรุณากรอกรายการรับหรือจ่ายอย่างน้อย 1 รายการ');
  if(items.length>200)throw new Error('บันทึกได้สูงสุดครั้งละ 200 รายการ');
  items.forEach((x,i)=>{
    project_(x.projectId);
    if(!clean_(x.description))throw new Error('กรุณากรอกรายละเอียดแถวที่ '+(i+1));
    const amount=positive_(x.amount,'ยอดเงินแถวที่ '+(i+1)),type=clean_(x.type)==='INCOME'?'INCOME':'EXPENSE',stamp=new Date();
    require_(user,type==='INCOME'?'FINANCE_EDIT':'EXPENSE_EDIT');
    if(x.id){const sheet=type==='INCOME'?'Income':'Expenses',row=rows_(sheet).find(r=>r.id===x.id);if(!row)throw new Error('ไม่พบรายการเดิมในแถวที่ '+(i+1));const changes={projectId:x.projectId,date:date_(x.date),description:clean_(x.description),amountExVat:amount,vat:0,total:amount};if(type==='INCOME')changes.account=clean_(x.paidBy);else{changes.category=clean_(x.category)||'อื่น ๆ';changes.paidBy=clean_(x.paidBy)}updateRow_(sheet,x.id,changes);audit_(user,'UPDATE',type+'_GRID',x.id,String(amount));return}
    const id=id_(type==='INCOME'?'INC':'EXP');if(type==='INCOME')append_('Income',[id,x.projectId,date_(x.date),clean_(x.description)||'รับเงินโครงการ',amount,0,amount,clean_(x.paidBy),stamp,user.id]);else append_('Expenses',[id,x.projectId,date_(x.date),clean_(x.category)||'อื่น ๆ',clean_(x.description),amount,0,amount,clean_(x.paidBy),false,stamp,user.id]);audit_(user,'CREATE',type+'_BATCH',id,String(amount));
  });
  return businessPayload_(db_());
}
function updateTransaction_(user,d){const type=clean_(d.type)==='INCOME'?'INCOME':'EXPENSE',sheet=type==='INCOME'?'Income':'Expenses';require_(user,type==='INCOME'?'FINANCE_EDIT':'EXPENSE_EDIT');project_(d.projectId);const row=rows_(sheet).find(x=>x.id===d.id);if(!row)throw new Error('ไม่พบรายการที่ต้องการแก้ไข');const amount=positive_(d.amount,'ยอดเงิน'),changes={projectId:d.projectId,date:date_(d.date),description:clean_(d.description),amountExVat:amount,vat:0,total:amount};if(!changes.description)throw new Error('กรุณากรอกรายละเอียด');if(type==='INCOME')changes.account=clean_(d.paidBy);else{changes.category=clean_(d.category)||'อื่น ๆ';changes.paidBy=clean_(d.paidBy)}updateRow_(sheet,d.id,changes);audit_(user,'UPDATE',type,d.id,'ย้ายไป '+d.projectId+' · '+amount);return businessPayload_(db_())}
function deleteTransaction_(user,d){const type=clean_(d.type)==='INCOME'?'INCOME':'EXPENSE',sheet=type==='INCOME'?'Income':'Expenses';require_(user,type==='INCOME'?'FINANCE_EDIT':'EXPENSE_EDIT');const row=rows_(sheet).find(x=>x.id===d.id);if(!row)throw new Error('ไม่พบรายการที่ต้องการลบ');deleteRow_(sheet,d.id);audit_(user,'DELETE',type,d.id,String(row.description||row.amountExVat||''));return businessPayload_(db_())}
function saveEmployee_(user,d) {
  require_(user,'EMPLOYEE_MANAGE');if(!clean_(d.name)||!clean_(d.position))throw new Error('กรุณากรอกชื่อและตำแหน่งงาน');const ss=db_();
  if(d.id){updateRow_('Employees',d.id,{name:clean_(d.name),position:clean_(d.position),type:clean_(d.type)||'PART_TIME',wageRate:num_(d.wageRate),otRate:num_(d.otRate),note:clean_(d.note)});return Object.assign({},payrollPayload_(ss),attendancePayload_(ss))}
  const id=id_('EMP');ss.getSheetByName('Employees').appendRow([id,clean_(d.name),clean_(d.type)||'PART_TIME',num_(d.wageRate),num_(d.otRate),true,clean_(d.note),new Date(),clean_(d.position)]);return Object.assign({},payrollPayload_(ss),attendancePayload_(ss));
}
function deleteEmployee_(user,d){require_(user,'EMPLOYEE_MANAGE');const ss=db_(),employee=readSheet_(ss,'Employees').find(x=>x.id===d.id);if(!employee)throw new Error('ไม่พบพนักงาน');const hasHistory=readSheet_(ss,'Overtime').some(x=>x.employeeId===d.id)||readSheet_(ss,'SalaryAdvances').some(x=>x.employeeId===d.id)||readSheet_(ss,'SalaryPayments').some(x=>x.employeeId===d.id)||readSheet_(ss,'Attendance').some(x=>x.employeeId===d.id);if(hasHistory)updateRow_('Employees',d.id,{active:false});else deleteRow_('Employees',d.id);return Object.assign({},payrollPayload_(ss),attendancePayload_(ss))}
function attendanceRows_(ss){return readSheet_(ss,'Attendance').map(row=>({...row,penaltyAmount:num_(row.penaltyAmount),otHours:num_(row.otHours),otAmount:num_(row.otAmount)})).sort((a,b)=>String(b.date||'').localeCompare(String(a.date||''))||String(b.checkInTime||'').localeCompare(String(a.checkInTime||''))||String(b.createdAt||'').localeCompare(String(a.createdAt||'')))}
function attendancePayload_(ss){return {partial:true,attendance:attendanceRows_(ss),employees:readSheet_(ss,'Employees').filter(x=>truthy_(x.active))}}
function saveAttendance_(user,d){
  const ss=db_(),emp=rows_('Employees').find(x=>x.id===d.employeeId),employeeName=emp?.name||clean_(d.employeeName)||'พนักงาน',date=date_(d.date||new Date()),checkInTime=clean_(d.checkInTime),checkOutTime=clean_(d.checkOutTime),otHours=num_(d.otHours),otRate=emp?num_(emp.otRate):num_(d.otRate),otAmount=d.otAmount!=null?round_(num_(d.otAmount)):round_(otHours*otRate),penaltyAmount=round_(num_(d.penaltyAmount)),penaltyReason=clean_(d.penaltyReason),status=clean_(d.status)||(checkOutTime?'CHECKED_OUT':'CHECKED_IN'),locationName=clean_(d.locationName),latitude=clean_(d.latitude),longitude=clean_(d.longitude),photoData=String(d.photoData||'').slice(0,45000),note=clean_(d.note);
  if(d.id){
    const existing=rows_('Attendance').find(x=>x.id===d.id);if(!existing)throw new Error('ไม่พบข้อมูลการลงเวลาที่ต้องการแก้ไข');
    const changes={employeeId:d.employeeId||existing.employeeId,employeeName,date,checkInTime,checkOutTime,locationName,latitude,longitude,penaltyAmount,penaltyReason,otHours,otAmount,status,note};
    if(photoData)changes.photoData=photoData;
    updateRow_('Attendance',d.id,changes);
    audit_(user,'UPDATE','ATTENDANCE',d.id,employeeName+' '+date+' '+checkInTime);
  }else{
    const id=id_('ATT');
    append_('Attendance',[id,clean_(d.employeeId),employeeName,date,checkInTime,checkOutTime,locationName,latitude,longitude,photoData,penaltyAmount,penaltyReason,otHours,otAmount,status,note,new Date(),user.id]);
    audit_(user,'CREATE','ATTENDANCE',id,employeeName+' '+date+' '+checkInTime);
  }
  return attendancePayload_(ss);
}
function checkInAttendance_(user,d){
  if(!clean_(d.employeeId))throw new Error('กรุณาเลือกพนักงาน');
  const ss=db_(),emp=rows_('Employees').find(x=>x.id===d.employeeId&&truthy_(x.active));
  if(!emp)throw new Error('ไม่พบข้อมูลพนักงานที่เลือก');
  const date=date_(d.date||new Date()),now=new Date(),timeStr=clean_(d.checkInTime)||Utilities.formatDate(now,Session.getScriptTimeZone(),'HH:mm:ss');
  const allAtt=rows_('Attendance'),existingToday=allAtt.find(x=>x.employeeId===d.employeeId&&x.date===date);
  if(existingToday&&!d.allowMultiple)throw new Error(emp.name+' เช็คชื่อเข้างานวันนี้ไปแล้ว ('+existingToday.checkInTime+' น.)');
  const id=id_('ATT'),penaltyAmount=round_(num_(d.penaltyAmount)),penaltyReason=clean_(d.penaltyReason),locationName=clean_(d.locationName)||'หน้างาน',latitude=clean_(d.latitude),longitude=clean_(d.longitude),photoData=String(d.photoData||'').slice(0,45000),note=clean_(d.note);
  append_('Attendance',[id,emp.id,emp.name,date,timeStr,'',locationName,latitude,longitude,photoData,penaltyAmount,penaltyReason,0,0,'CHECKED_IN',note,new Date(),user.id]);
  audit_(user,'CHECKIN','ATTENDANCE',id,emp.name+' เช็คอิน '+date+' '+timeStr);
  return attendancePayload_(ss);
}
function checkOutAttendance_(user,d){
  const ss=db_(),allAtt=rows_('Attendance'),existing=d.id?allAtt.find(x=>x.id===d.id):null;
  const empId=clean_(d.employeeId)||(existing?existing.employeeId:'');
  if(!empId)throw new Error('กรุณาเลือกพนักงาน');
  const emp=rows_('Employees').find(x=>x.id===empId);
  if(!emp)throw new Error('ไม่พบข้อมูลพนักงาน');
  const date=date_(d.date||(existing?existing.date:new Date())),now=new Date(),timeStr=clean_(d.checkOutTime)||Utilities.formatDate(now,Session.getScriptTimeZone(),'HH:mm:ss');
  const att=existing||allAtt.find(x=>x.employeeId===empId&&x.date===date);
  if(!att)throw new Error('ไม่พบรายการเช็คอินของ '+emp.name+' ในวันนี้ กรุณาเช็คอินก่อน');
  const otHours=num_(d.otHours),otRate=num_(emp.otRate)||0,otAmount=d.otAmount!=null?round_(num_(d.otAmount)):round_(otHours*otRate),penaltyAmount=d.penaltyAmount!=null?round_(num_(d.penaltyAmount)):num_(att.penaltyAmount),penaltyReason=clean_(d.penaltyReason)||att.penaltyReason||'';
  const changes={checkOutTime:timeStr,otHours,otAmount,penaltyAmount,penaltyReason,status:'CHECKED_OUT'};
  if(clean_(d.locationName))changes.locationName=clean_(d.locationName);
  if(clean_(d.latitude))changes.latitude=clean_(d.latitude);
  if(clean_(d.longitude))changes.longitude=clean_(d.longitude);
  if(d.photoData)changes.photoData=String(d.photoData).slice(0,45000);
  if(clean_(d.note))changes.note=clean_(d.note);
  updateRow_('Attendance',att.id,changes);
  audit_(user,'CHECKOUT','ATTENDANCE',att.id,emp.name+' เช็คเอาท์ '+date+' '+timeStr+' (OT: '+otHours+' ชม.)');
  return attendancePayload_(ss);
}
function deleteAttendance_(user,d){
  require_(user,'EMPLOYEE_MANAGE');
  const ss=db_(),item=rows_('Attendance').find(x=>x.id===d.id);
  if(!item)throw new Error('ไม่พบรายการลงเวลาที่ต้องการลบ');
  deleteRow_('Attendance',d.id);
  audit_(user,'DELETE','ATTENDANCE',d.id,item.employeeName+' '+item.date);
  return attendancePayload_(ss);
}
function productCostRows_(ss){
  const rows=readSheet_(ss,'ProductCosts').map(r=>({...r,unitPrice:num_(r.unitPrice),minPrice:num_(r.minPrice),maxPrice:num_(r.maxPrice)}));
  const priceMap={};
  rows.forEach(r=>{
    const key=(r.name||'').trim().toLowerCase();
    if(!key)return;
    if(!priceMap[key])priceMap[key]={min:r.unitPrice,max:r.unitPrice};
    else{
      if(r.unitPrice<priceMap[key].min)priceMap[key].min=r.unitPrice;
      if(r.unitPrice>priceMap[key].max)priceMap[key].max=r.unitPrice;
    }
  });
  return rows.map(r=>{
    const key=(r.name||'').trim().toLowerCase();
    const stats=priceMap[key]||{min:r.unitPrice,max:r.unitPrice};
    return {...r,minPrice:stats.min,maxPrice:stats.max};
  }).sort((a,b)=>String(a.category||'').localeCompare(String(b.category||''),'th')||String(a.name||'').localeCompare(String(b.name||''),'th')||a.unitPrice-b.unitPrice);
}
function productCostsPayload_(ss){
  return {partial:true,productCosts:productCostRows_(ss)};
}
function saveProductCost_(user,d){
  require_(user,'EXPENSE_EDIT');
  const ss=db_(),name=clean_(d.name),category=clean_(d.category)||'Solar Cell',unitPrice=positive_(d.unitPrice,'ราคาต่อหน่วย');
  if(!name)throw new Error('กรุณากรอกชื่อสินค้าหรืออุปกรณ์');
  const code=clean_(d.code)||id_('PRC').slice(4),subCategory=clean_(d.subCategory),brand=clean_(d.brand),model=clean_(d.model),supplier=clean_(d.supplier)||'ร้านค้าทั่วไป',unit=clean_(d.unit)||'ชิ้น',note=clean_(d.note),lastPurchaseDate=date_(d.lastPurchaseDate||new Date());
  if(d.id){
    const existing=rows_('ProductCosts').find(x=>x.id===d.id);
    if(!existing)throw new Error('ไม่พบรายการสินค้าที่ต้องการแก้ไข');
    updateRow_('ProductCosts',d.id,{code,name,category,subCategory,brand,model,supplier,unit,unitPrice,minPrice:Math.min(num_(existing.minPrice)||unitPrice,unitPrice),maxPrice:Math.max(num_(existing.maxPrice)||unitPrice,unitPrice),lastPurchaseDate,note});
    audit_(user,'UPDATE','PRODUCT_COST',d.id,name+' ('+supplier+') '+unitPrice);
  }else{
    const id=id_('PRC');
    append_('ProductCosts',[id,code,name,category,subCategory,brand,model,supplier,unit,unitPrice,unitPrice,unitPrice,lastPurchaseDate,note,new Date(),user.id]);
    audit_(user,'CREATE','PRODUCT_COST',id,name+' ('+supplier+') '+unitPrice);
  }
  return productCostsPayload_(ss);
}
function deleteProductCost_(user,d){
  require_(user,'EXPENSE_EDIT');
  const ss=db_(),existing=rows_('ProductCosts').find(x=>x.id===d.id);
  if(!existing)throw new Error('ไม่พบรายการสินค้าที่ต้องการลบ');
  deleteRow_('ProductCosts',d.id);
  audit_(user,'DELETE','PRODUCT_COST',d.id,existing.name);
  return productCostsPayload_(ss);
}
function saveCostEstimateToExpenses_(user,d){
  require_(user,'EXPENSE_EDIT');
  const ss=db_();
  project_(d.projectId);
  const items=Array.isArray(d.items)?d.items:[];
  if(!items.length)throw new Error('ไม่มีรายการสำหรับบันทึกลงโปรเจกต์');
  items.forEach((item,i)=>{
    const desc=clean_(item.name||item.description);
    if(!desc)return;
    const amount=positive_(item.amount||(num_(item.quantity)*num_(item.unitPrice)),'ยอดเงินรายการที่ '+(i+1));
    const id=id_('EXP'),cat=clean_(item.category)||'อุปกรณ์และวัสดุ',supplier=clean_(item.supplier)||'เงินสด/โอน';
    append_('Expenses',[id,d.projectId,date_(d.date||new Date()),cat,desc+(item.quantity?' ('+item.quantity+' '+(item.unit||'หน่วย')+' @ '+item.unitPrice+')':''),amount,0,amount,supplier,false,new Date(),user.id]);
    audit_(user,'CREATE','EXPENSE_ESTIMATE',id,String(amount));
  });
  return businessPayload_(ss);
}
function seedProductCosts_(){
  const ss=db_();
  if(readSheet_(ss,'ProductCosts').length>0)return;
  const initial=[
    ['PRC-001','SLR-PV-01','แผงโซล่าเซลล์ AIKO 670W N-Type ABC','Solar Cell','แผงโซล่าเซลล์','AIKO','AK-A670-N','ตัวแทนจำหน่ายตรง','แผง',3450,3450,3800,new Date(),'ประสิทธิภาพ 23.8% รับประกันกำลังผลิต 30 ปี'],
    ['PRC-002','SLR-PV-02','แผงโซล่าเซลล์ AIKO 670W N-Type ABC','Solar Cell','แผงโซล่าเซลล์','AIKO','AK-A670-N','ไทวัสดุ','แผง',3800,3450,3800,new Date(),'ซื้อหน้าร้าน มีของพร้อมส่งทันที'],
    ['PRC-003','SLR-PV-03','แผงโซล่าเซลล์ Trina Solar 720W Vertex N','Solar Cell','แผงโซล่าเซลล์','Trina','TSM-NEG21C.20','ตัวแทนจำหน่ายตรง','แผง',3600,3600,3950,new Date(),'Bifacial สองหน้า เหมาะสำหรับงานอุตสาหกรรม'],
    ['PRC-004','SLR-PV-04','แผงโซล่าเซลล์ Trina Solar 720W Vertex N','Solar Cell','แผงโซล่าเซลล์','Trina','TSM-NEG21C.20','ร้านแสงทองการไฟฟ้า','แผง',3950,3600,3950,new Date(),'สั่งยกล็อต 50 แผงขึ้นไปลดเพิ่ม 3%'],
    ['PRC-005','SLR-INV-01','Inverter Huawei SUN2000-5KTL-L1 (1 Phase)','Solar Cell','อินเวอร์เตอร์','Huawei','SUN2000-5KTL-L1','Huawei Distributor','เครื่อง',32500,32500,35000,new Date(),'On-grid 5kW 1P รับประกันศูนย์ไทย 10 ปี'],
    ['PRC-006','SLR-INV-02','Inverter Huawei SUN2000-5KTL-L1 (1 Phase)','Solar Cell','อินเวอร์เตอร์','Huawei','SUN2000-5KTL-L1','เมกาโฮม','เครื่อง',35000,32500,35000,new Date(),'ศูนย์บริการใกล้บ้าน เคลมด่วนได้'],
    ['PRC-007','SLR-INV-03','Inverter Huawei SUN2000-10KTL-M1 (3 Phase)','Solar Cell','อินเวอร์เตอร์','Huawei','SUN2000-10KTL-M1','Huawei Distributor','เครื่อง',54000,54000,57500,new Date(),'On-grid 10kW 3P ประสิทธิภาพ 98.6%'],
    ['PRC-008','SLR-INV-04','Inverter Solis Hybrid 10kW (3 Phase)','Solar Cell','อินเวอร์เตอร์','Solis','S6-EH3P10K-H','Solis Direct','เครื่อง',62000,62000,66500,new Date(),'Hybrid รองรับ Battery High-Voltage'],
    ['PRC-009','SLR-BAT-01','Battery Solis 16 kWh IntelliHome (IP66)','Solar Cell','แบตเตอรี่กักเก็บพลังงาน','Solis','314Ah 51.2V','Solis Direct','ชุด',125000,125000,132000,new Date(),'LiFePO4 ปลอดภัยสูง รอบชาร์จมากกว่า 6,000 ครั้ง'],
    ['PRC-010','SLR-MTR-01','Smart Power Sensor Huawei DDSU666-H (1P)','Solar Cell','สมาร์ทมิเตอร์','Huawei','DDSU666-H','Huawei Distributor','ชุด',2800,2800,3100,new Date(),'มิเตอร์กันย้อน 1 เฟส เชื่อมต่อ RS485'],
    ['PRC-011','SLR-MTR-02','Smart Power Sensor Huawei DTSU666-H (3P)','Solar Cell','สมาร์ทมิเตอร์','Huawei','DTSU666-H','Huawei Distributor','ชุด',4500,4500,4900,new Date(),'มิเตอร์กันย้อน 3 เฟส พร้อม CT ในชุด'],
    ['PRC-012','EVC-CHG-01','เครื่องชาร์จ EV Charger 7.4 kW 1 Phase','EV Charger','เครื่องชาร์จรถยนต์ไฟฟ้า','ABB','Terra AC Wallbox','ABB Direct','เครื่อง',24500,24500,27900,new Date(),'Type 2 สายยาว 5 เมตร มาตรฐานสากล'],
    ['PRC-013','EVC-CHG-02','เครื่องชาร์จ EV Charger 7.4 kW 1 Phase','EV Charger','เครื่องชาร์จรถยนต์ไฟฟ้า','ABB','Terra AC Wallbox','ไทวัสดุ','เครื่อง',27900,24500,27900,new Date(),'รวมโปรโมชั่นรูดบัตรผ่อน 0% 10 เดือน'],
    ['PRC-014','EVC-CHG-03','เครื่องชาร์จ EV Charger 22 kW 3 Phase','EV Charger','เครื่องชาร์จรถยนต์ไฟฟ้า','ABB','Terra AC Wallbox 22k','ABB Direct','เครื่อง',36000,36000,39500,new Date(),'Type 2 รองรับ RFID, Wi-Fi, Ethernet'],
    ['PRC-015','EVC-CHG-04','เครื่องชาร์จ EV Wallbox Pulsar Plus 22kW','EV Charger','เครื่องชาร์จรถยนต์ไฟฟ้า','Wallbox','Pulsar Plus 22kW','ตัวแทนจำหน่าย','เครื่อง',41000,41000,44500,new Date(),'ขนาดกะทัดรัด ดีไซน์หรู จัดการผ่านแอป'],
    ['PRC-016','EVC-RCD-01','เบรกเกอร์กันดูด RCD Type B 40A 2P','EV Charger','อุปกรณ์ป้องกันไฟฟ้า','CNC','Type B 40A 30mA','ร้านแสงทองการไฟฟ้า','ตัว',2400,2400,2750,new Date(),'มาตรฐานสำหรับติดตั้ง EV Charger ป้องกันไฟรั่ว DC'],
    ['PRC-017','EVC-COL-01','แท่นเสาตั้งเครื่องชาร์จ EV Pedestal สแตนเลส','EV Charger','เสาติดตั้งและแท่นวาง','MB Fabricate','MB-EV-STAND-01','โรงงานสั่งผลิต','ต้น',3500,3500,4200,new Date(),'สแตนเลส 304 ทนแดด ทนฝน ไม่เป็นสนิม'],
    ['PRC-018','STR-RAL-01','รางอลูมิเนียม Solar Mounting 4.2 เมตร','โครงสร้างและท่อร้อยสาย','รางยึดแผงโซล่าเซลล์','MB Aluminum','RAIL-4200-AL','โรงงานอลูมิเนียม','เส้น',320,320,410,new Date(),'อลูมิเนียมเกรด AL6005-T5 หนา แข็งแรง'],
    ['PRC-019','STR-RAL-02','รางอลูมิเนียม Solar Mounting 4.2 เมตร','โครงสร้างและท่อร้อยสาย','รางยึดแผงโซล่าเซลล์','MB Aluminum','RAIL-4200-AL','ไทวัสดุ','เส้น',390,320,410,new Date(),'มีสต็อกพร้อมขนย้ายหน้างาน'],
    ['PRC-020','STR-CLP-01','L-Feet สแตนเลสพร้อมยางกันซึม EPDM','โครงสร้างและท่อร้อยสาย','อุปกรณ์ยึดหลังคา','Sunfix','L-FEET-SS','ตัวแทนจำหน่าย','ชุด',35,35,48,new Date(),'ใช้กับหลังคา Metal Sheet กันน้ำซึม 100%'],
    ['PRC-021','STR-CLP-02','L-Feet สแตนเลสพร้อมยางกันซึม EPDM','โครงสร้างและท่อร้อยสาย','อุปกรณ์ยึดหลังคา','Sunfix','L-FEET-SS','ไทวัสดุ','ชุด',48,35,48,new Date(),'ซื้อปลีกหน้าสาขา'],
    ['PRC-022','STR-CLP-03','Mid Clamp อลูมิเนียม 35/40 มม.','โครงสร้างและท่อร้อยสาย','อุปกรณ์ยึดแผง','Sunfix','MC-3540-AL','ตัวแทนจำหน่าย','ตัว',16,16,25,new Date(),'ตัวยึดระหว่างแผง พร้อมสกรูสแตนเลส'],
    ['PRC-023','STR-CLP-04','End Clamp อลูมิเนียม 35/40 มม.','โครงสร้างและท่อร้อยสาย','อุปกรณ์ยึดแผง','Sunfix','EC-3540-AL','ตัวแทนจำหน่าย','ตัว',16,16,25,new Date(),'ตัวยึดริมแผง พร้อมสกรูสแตนเลส'],
    ['PRC-024','STR-PIP-01','ท่อเหล็กร้อยสายไฟ EMT 1/2 นิ้ว (ยาว 3 ม.)','โครงสร้างและท่อร้อยสาย','ท่อร้อยสายไฟฟ้า','PAT','EMT-050','ร้านแสงทองการไฟฟ้า','ท่อน',125,125,145,new Date(),'ท่อเหล็กชุบสังกะสี มาตรฐาน มอก.'],
    ['PRC-025','STR-PIP-02','ท่อเหล็กร้อยสายไฟ EMT 1/2 นิ้ว (ยาว 3 ม.)','โครงสร้างและท่อร้อยสาย','ท่อร้อยสายไฟฟ้า','PAT','EMT-050','ไทวัสดุ','ท่อน',145,125,145,new Date(),'ซื้อเฉพาะงานด่วนขาดของ'],
    ['PRC-026','STR-PIP-03','ท่อเหล็กร้อยสายไฟ IMC 3/4 นิ้ว (ยาว 3 ม.)','โครงสร้างและท่อร้อยสาย','ท่อร้อยสายไฟฟ้า','PAT','IMC-075','ร้านแสงทองการไฟฟ้า','ท่อน',280,280,320,new Date(),'ท่อหนาเกลียวหัวท้าย สำหรับภายนอกอาคาร'],
    ['PRC-027','STR-CBL-01','สายไฟ Solar DC Cable 1x4 sq.mm (100 ม.)','โครงสร้างและท่อร้อยสาย','สายไฟฟ้าโซล่าเซลล์','Link','CB-1040-BK','Link Direct','ม้วน',2100,2100,2450,new Date(),'ทนแรงดัน 1500VDC ฉนวน 2 ชั้น ทน UV'],
    ['PRC-028','STR-CBL-02','สายไฟ Solar DC Cable 1x4 sq.mm (100 ม.)','โครงสร้างและท่อร้อยสาย','สายไฟฟ้าโซล่าเซลล์','Link','CB-1040-BK','ไทวัสดุ','ม้วน',2450,2100,2450,new Date(),'มีสต็อกหน้าสาขา'],
    ['PRC-029','STR-CBL-03','สายไฟ Solar DC Cable 1x6 sq.mm (100 ม.)','โครงสร้างและท่อร้อยสาย','สายไฟฟ้าโซล่าเซลล์','Link','CB-1060-BK','Link Direct','ม้วน',2900,2900,3350,new Date(),'สำหรับระยะสายไกลกว่า 30 เมตร ลด Loss'],
    ['PRC-030','SRV-LAB-01','ค่าแรงติดตั้งระบบ Solar Cell (รวมโครงสร้างและระบบไฟ)','ค่าแรงและงานบริการ','งานติดตั้งระบบ Solar','ทีมช่าง MB','LABOR-SLR-KW','ทีมช่างประจำ MB','กิโลวัตต์ (kW)',1200,1200,1500,new Date(),'ติดตั้งตามมาตรฐานวิศวกรรมของ MB Solar'],
    ['PRC-031','SRV-LAB-02','ค่าแรงติดตั้งระบบ Solar Cell (รวมโครงสร้างและระบบไฟ)','ค่าแรงและงานบริการ','งานติดตั้งระบบ Solar','Subcontract','LABOR-SLR-SUB','ผู้รับเหมาช่วง','กิโลวัตต์ (kW)',1500,1200,1500,new Date(),'สำหรับกรณีคิวงานช่างประจำเต็ม'],
    ['PRC-032','SRV-LAB-03','ค่าแรงติดตั้งเครื่องชาร์จ EV Charger มาตรฐาน','ค่าแรงและงานบริการ','งานติดตั้ง EV Charger','ทีมช่าง MB','LABOR-EVC-PT','ทีมช่างประจำ MB','จุด',3500,3500,4500,new Date(),'รวมเดินสายท่อ EMT ระยะไม่เกิน 15 เมตร'],
    ['PRC-033','SRV-ENG-01','ค่าวิศวกรจัดทำแบบ SLD และลงนามรับรองโครงสร้าง/ไฟฟ้า','ค่าแรงและงานบริการ','งานวิศวกรรมรับรอง','วิศวกรวิชาชีพ','ENG-SLD-SIGN','วิศวกรภาคี/สามัญ','โครงการ',3000,3000,4500,new Date(),'สำหรับยื่นขอขนานไฟ PEA และ MEA'],
    ['PRC-034','SRV-FEE-01','ค่าธรรมเนียมยื่นขอขนานไฟ PEA/MEA','ค่าแรงและงานบริการ','ค่าธรรมเนียมราชการ','การไฟฟ้า','FEE-GRID-TIE','การไฟฟ้าส่วนภูมิภาค/นครหลวง','โครงการ',3500,3500,3500,new Date(),'ค่าตรวจสอบระบบและการเปลี่ยนมิเตอร์ไฟฟ้าแบบสองทิศทาง'],
    ['PRC-035','MSC-GRD-01','แท่งกราวด์ร็อดทองแดง 5/8 นิ้ว ยาว 2.4 เมตร','เบ็ดเตล็ดและสิ้นเปลือง','ระบบสายดิน Grounding','Kumwell','GR-58-24','ร้านแสงทองการไฟฟ้า','ต้น',220,220,260,new Date(),'แท่งกราวด์ทองแดงแท้ มอก.'],
    ['PRC-036','MSC-GRD-02','แท่งกราวด์ร็อดทองแดง 5/8 นิ้ว ยาว 2.4 เมตร','เบ็ดเตล็ดและสิ้นเปลือง','ระบบสายดิน Grounding','Kumwell','GR-58-24','ไทวัสดุ','ต้น',260,220,260,new Date(),'ซื้อหน้าร้าน'],
    ['PRC-037','MSC-CON-01','เทปพันสายไฟ 3M Scotch 33+ ทนความร้อนสูง','เบ็ดเตล็ดและสิ้นเปลือง','วัสดุสิ้นเปลือง','3M','Scotch 33+','ร้านแสงทองการไฟฟ้า','ม้วน',65,65,79,new Date(),'เกรดงานอุตสาหกรรม ทนอุณหภูมิ 105 องศาเซลเซียส'],
    ['PRC-038','MSC-CON-02','ซิลิโคนกันน้ำ Weatherproof Sealant สเปคงานกลางแจ้ง','เบ็ดเตล็ดและสิ้นเปลือง','วัสดุสิ้นเปลือง','Dowsil','791-WEATHER','ไทวัสดุ','หลอด',135,130,140,new Date(),'กันรังสี UV ยึดเกาะโลหะและคอนกรีตดีเยี่ยม'],
    ['PRC-039','MSC-WNG-01','ชุดป้ายเตือน Warning Signs ระบบ Solar ตามแบบ PEA','เบ็ดเตล็ดและสิ้นเปลือง','ป้ายเตือนความปลอดภัย','MB Sign','SIGN-PEA-SET','โรงพิมพ์ป้าย','ชุด',450,450,550,new Date(),'แผ่นอะคริลิกสะท้อนแสง ทนแดด 5 ปี']
  ];
  initial.forEach(row=>{
    append_('ProductCosts',[row[0],row[1],row[2],row[3],row[4],row[5],row[6],row[7],row[8],row[9],row[10],row[11],row[12],row[13],new Date(),'SYSTEM']);
  });
}
function saveOvertime_(user,d) {
  require_(user,'OT_EDIT'); project_(d.projectId); const employee=rows_('Employees').find(x=>x.id===d.employeeId); if(!employee)throw new Error('ไม่พบพนักงาน');
  const hours=positive_(d.hours,'จำนวนชั่วโมง'),rate=positive_(d.rate,'อัตรา OT'),amount=round_(hours*rate),id=id_('OT');
  append_('Overtime',[id,employee.id,d.projectId,date_(d.date),hours,rate,amount,clean_(d.note),false,new Date(),user.id]); audit_(user,'CREATE','OVERTIME',id,employee.name+' '+amount); return Object.assign({},businessPayload_(db_()),payrollPayload_(db_()));
}
function saveSalaryAdvance_(user,d){require_(user,'EMPLOYEE_MANAGE');const employee=rows_('Employees').find(x=>x.id===d.employeeId&&truthy_(x.active));if(!employee)throw new Error('ไม่พบพนักงาน');const month=monthInfo_(),amount=positive_(d.amount,'ยอดเบิก'),payments=rows_('SalaryPayments');if(payments.some(x=>x.employeeId===employee.id&&x.monthKey===month.key))throw new Error('พนักงานคนนี้ปิดรอบเงินเดือนแล้ว');const used=rows_('SalaryAdvances').filter(x=>x.employeeId===employee.id&&x.monthKey===month.key).reduce((s,x)=>s+num_(x.amount),0);if(amount>num_(employee.wageRate)-used)throw new Error('ยอดเบิกมากกว่าเงินเดือนคงเหลือ');const id=id_('ADV');append_('SalaryAdvances',[id,employee.id,month.key,date_(d.date),amount,clean_(d.note),new Date(),user.id]);audit_(user,'CREATE','SALARY_ADVANCE',id,employee.name+' '+amount);return payrollPayload_(db_())}
function payMonthlySalary_(user,d){require_(user,'EMPLOYEE_MANAGE');const employee=rows_('Employees').find(x=>x.id===d.employeeId&&truthy_(x.active));if(!employee)throw new Error('ไม่พบพนักงาน');const month=monthInfo_(),payments=rows_('SalaryPayments');if(payments.some(x=>x.employeeId===employee.id&&x.monthKey===month.key))throw new Error('ปิดรอบเงินเดือนนี้แล้ว');const advanceTotal=round_(rows_('SalaryAdvances').filter(x=>x.employeeId===employee.id&&x.monthKey===month.key).reduce((s,x)=>s+num_(x.amount),0)),otTotal=round_(rows_('Overtime').filter(x=>x.employeeId===employee.id&&String(x.date).slice(0,7)===month.key).reduce((s,x)=>s+num_(x.amount),0)),base=num_(employee.wageRate),net=round_(Math.max(0,base+otTotal-advanceTotal)),id=id_('SAL');append_('SalaryPayments',[id,employee.id,month.key,month.dueDate,base,advanceTotal,net,date_(d.paidDate),clean_(d.note),new Date(),user.id,otTotal]);audit_(user,'CREATE','SALARY_PAYMENT',id,employee.name+' '+net);return payrollPayload_(db_())}
function distributeProjectDividend_(user,d){require_(user,'FINANCE_EDIT');const partners=rows_('Partners').filter(x=>truthy_(x.active)&&(x.id==='PART-BAS'||x.id==='PART-GOLF'));if(partners.length!==2)throw new Error('ไม่พบข้อมูลหุ้นส่วนบาสและกอล์ฟครบทั้งสองคน');const expenses=rows_('Expenses').reduce((s,x)=>s+num_(x.amountExVat),0),ot=rows_('Overtime').reduce((s,x)=>s+num_(x.amount),0),income=rows_('Income').reduce((s,x)=>s+num_(x.amountExVat),0),alreadyPaid=rows_('Dividends').reduce((s,x)=>s+num_(x.amount),0),available=round_(income-expenses-ot-alreadyPaid),amountByPartner={'PART-BAS':Math.max(0,round_(d.basAmount)),'PART-GOLF':Math.max(0,round_(d.golfAmount))},total=round_(amountByPartner['PART-BAS']+amountByPartner['PART-GOLF']);if(available<=0)throw new Error('ไม่มีกำไรคงเหลือรวมสำหรับปันผล');if(total<=0)throw new Error('กรุณากรอกยอดที่จ่ายจริงอย่างน้อย 1 คน');if(total>available)throw new Error('ยอดจ่ายรวมมากกว่ากำไรคงเหลือรวม');const paidFrom=clean_(d.paidFrom),source=rows_('BalanceAccounts').find(x=>truthy_(x.active)&&x.name===paidFrom);if(!source)throw new Error('กรุณาเลือกบัญชีหรือเงินสดที่ใช้งานอยู่');const sourceBalance=round_(num_(source.openingBalance)+rows_('Income').filter(x=>x.account===paidFrom).reduce((s,x)=>s+num_(x.total!=null?x.total:x.amountExVat),0)-rows_('Expenses').filter(x=>x.paidBy===paidFrom).reduce((s,x)=>s+num_(x.total!=null?x.total:x.amountExVat),0)-rows_('Dividends').filter(x=>x.paidFrom===paidFrom).reduce((s,x)=>s+num_(x.amount),0));if(total>sourceBalance)throw new Error('ยอดเงินใน '+paidFrom+' ไม่พอจ่ายปันผล (คงเหลือ '+sourceBalance+' บาท)');partners.forEach(partner=>{const amount=amountByPartner[partner.id];if(amount<=0)return;const id=id_('DIV');append_('Dividends',[id,partner.id,date_(d.date),amount,paidFrom,clean_(d.note)||'ปันผลกำไรรวมตามจริง',new Date(),user.id,'']);audit_(user,'CREATE','GLOBAL_DIVIDEND_ACTUAL',id,'กำไรรวม · '+partner.name+' '+amount)});return dividendPayload_(db_())}
function quotePresetRows_(ss){return readSheet_(ss,'QuotePresets').map(x=>{let data={};try{data=JSON.parse(x.dataJson||'{}')}catch(e){}return Object.assign({},x,{data})})}
function saveQuotePreset_(user,d){
  require_(user,'FINANCE_EDIT');
  if(!['TEMPLATE','SET','FORM'].includes(d.kind)||!['EV Charger','Solar Cell','Other'].includes(d.workType))throw new Error('ประเภทฟอร์มไม่ถูกต้อง');
  const name=clean_(d.name),long=v=>String(v==null?'':v).trim().slice(0,5000),all=rows_('QuotePresets');
  if(!name)throw new Error('กรุณาระบุชื่อชุด');
  const items=(Array.isArray(d.items)?d.items:[]).map(x=>({description:long(x.description),quantity:Number(x.quantity),unit:clean_(x.unit)||'ชุด',unitPrice:Number(x.unitPrice),packagePrice:x.packagePrice===true,packageIncluded:x.packageIncluded===true}));
  if((d.kind==='SET'||d.kind==='FORM')&&((d.kind==='SET'&&!items.length)||items.length>100||items.some(x=>!x.description||!Number.isFinite(x.quantity)||x.quantity<=0||!Number.isFinite(x.unitPrice)||x.unitPrice<0)))throw new Error('กรุณากรอกรายละเอียด จำนวน และราคาให้ถูกต้อง (สูงสุด 100 รายการ)');
  const data=d.kind==='SET'?{items}:Object.fromEntries(['paymentTerms','deliveryTerms','warrantyText','scopeText','note'].map(k=>[k,long(d[k])]));
  if(d.kind==='FORM')data.items=items;if(['TEMPLATE','FORM'].includes(d.kind)){data.company=quotationCompany_(d);data.logoData=String(d.logoData||'').slice(0,45000);data.noteColor=d.noteColor==='red'?'red':'black';data.vatRate=Number(d.vatRate);if(!Number.isFinite(data.vatRate)||data.vatRate<0)throw new Error('VAT ไม่ถูกต้อง')}
  const dataJson=JSON.stringify(data);if(dataJson.length>49000)throw new Error('ข้อมูลชุดหรือรูปโลโก้มีขนาดใหญ่เกินไป กรุณาลดขนาด');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{const fresh=rows_('QuotePresets'),existing=d.id?fresh.find(x=>x.id===d.id):d.kind==='TEMPLATE'?fresh.find(x=>x.kind===d.kind&&x.workType===d.workType):null;
    if(d.id&&!existing)throw new Error('ไม่พบชุดที่ต้องการแก้ไข');
    if(existing&&(existing.kind!==d.kind||existing.workType!==d.workType))throw new Error('ไม่สามารถเปลี่ยนประเภทชุดเดิมได้');
    const id=existing?existing.id:id_('QPR'),values={kind:d.kind,workType:d.workType,name,dataJson,updatedAt:new Date(),createdBy:user.id};
    if(existing)updateRow_('QuotePresets',id,values);else append_('QuotePresets',[id,values.kind,values.workType,name,dataJson,values.updatedAt,user.id]);
    audit_(user,existing?'UPDATE':'CREATE','QUOTE_PRESET',id,name);
  }finally{lock.releaseLock()}
  return {partial:true,quotePresets:quotePresetRows_(db_())};
}
function deleteQuotePreset_(user,d){require_(user,'FINANCE_EDIT');const x=rows_('QuotePresets').find(x=>x.id===d.id);if(!x||!['SET','FORM'].includes(x.kind))throw new Error('ไม่พบชุดรายการ');deleteRow_('QuotePresets',d.id);audit_(user,'DELETE','QUOTE_PRESET',d.id,x.name);return {partial:true,quotePresets:quotePresetRows_(db_())}}
function quotationRows_(ss){return readSheet_(ss,'Quotations').map(x=>{let items=[],company={};try{company=JSON.parse(x.companyJson||'{}')}catch(e){}try{items=JSON.parse(String(x.itemsJson||'[]'))}catch(e){}return Object.assign({},x,{items:Array.isArray(items)?items:[],company})}).sort((a,b)=>String(b.issueDate).localeCompare(String(a.issueDate))||String(b.createdAt).localeCompare(String(a.createdAt)))}
function documentNumber_(prefix,sequence,issueDate){const date=String(issueDate||'').match(/^(\d{4})-(\d{2})-(\d{2})$/),number=Number(sequence);if(!['QT','IV','RC'].includes(prefix)||!date||!Number.isSafeInteger(number)||number<1)throw new Error('ข้อมูลเลขเอกสารไม่ถูกต้อง');return prefix+String(number).padStart(2,'0')+date[3]+date[2]+String(Number(date[1])+543)}
function quotationSequence_(quoteNo){const text=String(quoteNo||''),current=text.match(/^QT(\d+)\d{8}$/),legacy=text.match(/^QT(?:SO|EV)-\d{6}-(\d+)$/);return Number(current?.[1]||legacy?.[1]||0)}
function relatedDocumentNumber_(prefix,q,installment,issueDate){if(!['IV','RC'].includes(prefix))throw new Error('ประเภทเอกสารไม่ถูกต้อง');const reference=String(q.quoteNo||''),modern=/^QT\d{9,}$/.test(reference),sequence=quotationSequence_(reference);if(!sequence)throw new Error('กรุณาปรับเลขใบเสนอราคาเป็นรูปแบบ QT เช่น QT8013092569 ก่อนออกเอกสาร');const base=issueDate?documentNumber_(prefix,sequence,issueDate):modern?prefix+reference.slice(2):documentNumber_(prefix,sequence,q.issueDate);if(installment!==undefined&&![1,2].includes(Number(installment)))throw new Error('งวดเอกสารไม่ถูกต้อง');return base+(installment===undefined?'':'-'+Number(installment))}
function nextQuotationNo_(workType,issueDate){const props=PropertiesService.getScriptProperties(),max=rows_('Quotations').reduce((n,q)=>Math.max(n,quotationSequence_(q.quoteNo)),Number(props.getProperty('DOCUMENT_QT_SEQUENCE'))||0),next=max+1;props.setProperty('DOCUMENT_QT_SEQUENCE',String(next));return documentNumber_('QT',next,issueDate)}
function saveQuotation_(user,d){require_(user,'FINANCE_EDIT');const lock=LockService.getScriptLock();lock.waitLock(30000);try{const result=saveQuotationLocked_(user,d),props=PropertiesService.getScriptProperties(),sequence=quotationSequence_(d.quoteNo);if(sequence>(Number(props.getProperty('DOCUMENT_QT_SEQUENCE'))||0))props.setProperty('DOCUMENT_QT_SEQUENCE',String(sequence));return result}finally{lock.releaseLock()}}
function saveQuotationLocked_(user,d){require_(user,'FINANCE_EDIT');const long=v=>String(v==null?'':v).trim().slice(0,5000),customerName=clean_(d.customerName),issueDate=date_(d.issueDate),workType=['EV Charger','Solar Cell','Other'].includes(clean_(d.workType))?clean_(d.workType):'Other',rawItems=Array.isArray(d.items)?d.items:[],items=rawItems.map(x=>({description:long(x.description),quantity:round_(x.quantity),unit:clean_(x.unit)||'ชุด',unitPrice:round_(x.unitPrice),packagePrice:x.packagePrice===true,packageIncluded:x.packageIncluded===true})).filter(x=>x.description||x.quantity||x.unitPrice);if(!customerName)throw new Error('กรุณากรอกชื่อลูกค้า');if(!items.length)throw new Error('กรุณาเพิ่มรายละเอียดงานอย่างน้อย 1 รายการ');if(items.length>100)throw new Error('เพิ่มรายละเอียดได้สูงสุด 100 รายการต่อเอกสาร');items.forEach((x,i)=>{if(!x.description)throw new Error('กรุณากรอกรายละเอียดรายการที่ '+(i+1));if(x.quantity<=0)throw new Error('จำนวนรายการที่ '+(i+1)+' ต้องมากกว่า 0');if(x.unitPrice<0)throw new Error('ราคาต่อหน่วยไม่ถูกต้อง')});const vatRate=Math.max(0,round_(d.vatRate==null?7:d.vatRate)),subtotal=round_(items.reduce((s,x)=>s+quotationLineAmount_(x),0)),vatAmount=round_(subtotal*vatRate/100),grandTotal=round_(subtotal+vatAmount),all=rows_('Quotations'),quoteNo=clean_(d.quoteNo)||all.find(x=>x.id===d.id)?.quoteNo||nextQuotationNo_(workType,issueDate);if(all.some(x=>x.id!==d.id&&String(x.quoteNo).toLowerCase()===quoteNo.toLowerCase()))throw new Error('เลขที่ใบเสนอราคานี้มีอยู่แล้ว');const company=quotationCompany_(d),existing=all.find(x=>x.id===d.id),companyJson=Object.keys(company).length?JSON.stringify(company):existing?.companyJson||'{}';const values={projectName:clean_(d.projectName),companyJson,quoteNo,issueDate,customerName,customerPhone:clean_(d.customerPhone),customerAddress:long(d.customerAddress),customerTaxId:clean_(d.customerTaxId),workType,itemsJson:JSON.stringify(items),vatRate,paymentTerms:long(d.paymentTerms),deliveryTerms:long(d.deliveryTerms),warrantyText:long(d.warrantyText),scopeText:long(d.scopeText),note:long(d.note),noteColor:d.noteColor==='red'?'red':'black',logoData:String(d.logoData||'').slice(0,45000),status:['DRAFT','SENT','APPROVED'].indexOf(clean_(d.status))>=0?clean_(d.status):'DRAFT',subtotal,vatAmount,grandTotal,updatedAt:new Date()};if(d.id){if(!all.some(x=>x.id===d.id))throw new Error('ไม่พบใบเสนอราคาที่ต้องการแก้ไข');updateRow_('Quotations',d.id,values);audit_(user,'UPDATE','QUOTATION',d.id,quoteNo)}else{const id=id_('QTN');append_('Quotations',[id,values.quoteNo,values.issueDate,values.customerName,values.customerPhone,values.customerAddress,values.customerTaxId,values.workType,values.itemsJson,values.vatRate,values.paymentTerms,values.deliveryTerms,values.warrantyText,values.scopeText,values.note,values.status,values.subtotal,values.vatAmount,values.grandTotal,new Date(),values.updatedAt,user.id,values.logoData,values.noteColor,values.companyJson,values.projectName]);audit_(user,'CREATE','QUOTATION',id,quoteNo)}return quotationPayload_(db_())}
function deleteQuotation_(user,d){require_(user,'FINANCE_EDIT');const item=rows_('Quotations').find(x=>x.id===d.id);if(!item)throw new Error('ไม่พบใบเสนอราคา');deleteRow_('Quotations',d.id);audit_(user,'DELETE','QUOTATION',d.id,item.quoteNo);return quotationPayload_(db_())}
function saveUser_(user,d) {
  require_(user,'USERS_MANAGE'); if(!clean_(d.name)||!clean_(d.username)||String(d.password||'').length<4)throw new Error('กรุณากรอกข้อมูลผู้ใช้และรหัสผ่านอย่างน้อย 4 ตัว');
  if(rows_('Users').some(x=>String(x.username).toLowerCase()===clean_(d.username).toLowerCase()))throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้ว');
  const id=id_('USR'),perms=Array.isArray(d.permissions)?d.permissions.join(','):''; append_('Users',[id,clean_(d.name),clean_(d.username),hash_(String(d.password)),clean_(d.role)||'MANAGER',perms,true,new Date()]); audit_(user,'CREATE','USER',id,d.name); return {partial:true,users:readSheet_(db_(),'Users').map(publicUser_)};
}

function saveCategory_(user,d){return saveLookup_(user,d,'Categories','CATEGORY','CAT')}
function deleteCategory_(user,d){return deleteLookup_(user,d,'Categories','CATEGORY')}
function reorderCategories_(user,d){return reorderLookup_(user,d,'Categories','CATEGORY')}
function savePaymentMethod_(user,d){return saveLookup_(user,d,'PaymentMethods','PAYMENT_METHOD','PAYBY')}
function deletePaymentMethod_(user,d){return deleteLookup_(user,d,'PaymentMethods','PAYMENT_METHOD')}
function reorderPaymentMethods_(user,d){return reorderLookup_(user,d,'PaymentMethods','PAYMENT_METHOD')}
function saveBalanceAccount_(user,d){require_(user,'SETTINGS_MANAGE');const name=clean_(d.name),type=clean_(d.type);if(!name)throw new Error('กรุณากรอกชื่อบัญชี');if(type!=='CASH'&&type!=='COMPANY')throw new Error('ประเภทบัญชีไม่ถูกต้อง');const accounts=rows_('BalanceAccounts');if(accounts.some(x=>x.id!==d.id&&String(x.name).toLowerCase()===name.toLowerCase()&&truthy_(x.active)))throw new Error('มีชื่อบัญชีนี้อยู่แล้ว');if(d.id){const old=accounts.find(x=>x.id===d.id);if(!old)throw new Error('ไม่พบบัญชีที่ต้องการแก้ไข');updateRow_('BalanceAccounts',d.id,{name,type,openingBalance:round_(d.openingBalance),sortOrder:num_(d.sortOrder)||99,active:true});if(old.name!==name){['Income','Expenses'].forEach(sheet=>{const field=sheet==='Income'?'account':'paidBy';rows_(sheet).filter(x=>x[field]===old.name).forEach(x=>updateRow_(sheet,x.id,{[field]:name}))});rows_('Dividends').filter(x=>x.paidFrom===old.name).forEach(x=>updateRow_('Dividends',x.id,{paidFrom:name}));const method=rows_('PaymentMethods').find(x=>x.name===old.name&&truthy_(x.active));if(method)updateRow_('PaymentMethods',method.id,{name})}audit_(user,'UPDATE','BALANCE_ACCOUNT',d.id,name)}else{const id=id_('BAL');append_('BalanceAccounts',[id,name,type,round_(d.openingBalance),true,num_(d.sortOrder)||99,new Date(),user.id]);if(!rows_('PaymentMethods').some(x=>String(x.name).toLowerCase()===name.toLowerCase()&&truthy_(x.active)))append_('PaymentMethods',[id_('PAYBY'),name,true,num_(d.sortOrder)||99,new Date(),user.id]);audit_(user,'CREATE','BALANCE_ACCOUNT',id,name)}return lookupPayload_(db_())}
function deleteBalanceAccount_(user,d){require_(user,'SETTINGS_MANAGE');const item=rows_('BalanceAccounts').find(x=>x.id===d.id);if(!item)throw new Error('ไม่พบบัญชี');const used=rows_('Income').some(x=>x.account===item.name)||rows_('Expenses').some(x=>x.paidBy===item.name)||rows_('Dividends').some(x=>x.paidFrom===item.name);if(used)throw new Error('ลบบัญชีนี้ไม่ได้ เพราะมีรายการรับ–จ่ายหรือปันผลใช้งานอยู่ กรุณาแก้ไขชื่อแทน');updateRow_('BalanceAccounts',d.id,{active:false});const method=rows_('PaymentMethods').find(x=>x.name===item.name&&truthy_(x.active));if(method)updateRow_('PaymentMethods',method.id,{active:false});audit_(user,'DELETE','BALANCE_ACCOUNT',d.id,item.name);return lookupPayload_(db_())}
function saveLookup_(user,d,sheet,entity,prefix){require_(user,'SETTINGS_MANAGE');const name=clean_(d.name);if(!name)throw new Error('กรุณากรอกชื่อรายการ');const all=rows_(sheet);if(all.some(x=>x.id!==d.id&&String(x.name).toLowerCase()===name.toLowerCase()&&truthy_(x.active)))throw new Error('มีชื่อนี้อยู่แล้ว');if(d.id){const old=all.find(x=>x.id===d.id);if(!old)throw new Error('ไม่พบรายการที่ต้องการแก้ไข');const oldName=String(old.name);updateRow_(sheet,d.id,{name,sortOrder:num_(d.sortOrder)||99,active:true});if(oldName!==name){if(sheet==='Categories')rows_('Expenses').filter(x=>x.category===oldName).forEach(x=>updateRow_('Expenses',x.id,{category:name}));if(sheet==='PaymentMethods'){rows_('Income').filter(x=>x.account===oldName).forEach(x=>updateRow_('Income',x.id,{account:name}));rows_('Expenses').filter(x=>x.paidBy===oldName).forEach(x=>updateRow_('Expenses',x.id,{paidBy:name}))}}audit_(user,'UPDATE',entity,d.id,name)}else{const id=id_(prefix);append_(sheet,[id,name,true,num_(d.sortOrder)||99,new Date(),user.id]);audit_(user,'CREATE',entity,id,name)}return lookupPayload_(db_())}
function deleteLookup_(user,d,sheet,entity){require_(user,'SETTINGS_MANAGE');const item=rows_(sheet).find(x=>x.id===d.id);if(!item)throw new Error('ไม่พบรายการ');updateRow_(sheet,d.id,{active:false});audit_(user,'DELETE',entity,d.id,item.name);return lookupPayload_(db_())}
function migrateLegacyFinanceChannelNames_(user){if(typeof PropertiesService==='undefined')return;const props=PropertiesService.getScriptProperties(),key='FINANCE_CHANNEL_MB_LABEL_V1';if(props.getProperty(key)==='done')return;const methods=rows_('PaymentMethods'),target=methods.find(x=>truthy_(x.active)&&x.name==='บัญชีบริษัท MB');if(!target)return;const oldName='บัญชีบริษัท',updates=[...rows_('Income').filter(x=>x.account===oldName).map(x=>['Income',x.id,'account']),...rows_('Expenses').filter(x=>x.paidBy===oldName).map(x=>['Expenses',x.id,'paidBy'])];updates.forEach(([sheet,id,field])=>updateRow_(sheet,id,{[field]:target.name}));if(updates.length)audit_(user,'UPDATE','FINANCE_HISTORY_CHANNEL','',oldName+' → '+target.name+' ('+updates.length+')');props.setProperty(key,'done')}
function reorderLookup_(user,d,sheet,entity){require_(user,'SETTINGS_MANAGE');const active=sortActiveRows_(rows_(sheet)),ids=Array.isArray(d.ids)?d.ids.map(String):[];if(ids.length!==active.length||new Set(ids).size!==ids.length||active.some(x=>ids.indexOf(String(x.id))<0))throw new Error('ลำดับรายการไม่ถูกต้อง กรุณาลองใหม่');ids.forEach((id,index)=>updateRow_(sheet,id,{sortOrder:index+1}));audit_(user,'REORDER',entity,'',ids.join(','));return lookupPayload_(db_())}

function permissions_(u){const all=['DASHBOARD_VIEW','PROFIT_VIEW','PROJECT_EDIT','FINANCE_EDIT','EXPENSE_EDIT','EMPLOYEE_MANAGE','OT_EDIT','LOAN_USE','USERS_MANAGE','SETTINGS_MANAGE'];if(u.role==='OWNER'||u.permissions==='*')return all;return String(u.permissions||'').split(',').filter(Boolean)}
function allowed_(u,p){return permissions_(u).indexOf(p)>=0} function require_(u,p){if(!allowed_(u,p))throw new Error('คุณไม่มีสิทธิ์ทำรายการนี้')}
function publicUser_(u){return{id:u.id,name:u.name,username:u.username,role:u.role,permissions:String(u.permissions||''),active:truthy_(u.active)}}
function project_(id){const p=rows_('Projects').find(x=>x.id===id);if(!p)throw new Error('กรุณาเลือกโปรเจกต์');return p}
function sortedActive_(name){return rows_(name).filter(x=>truthy_(x.active)).sort((a,b)=>num_(a.sortOrder)-num_(b.sortOrder)||String(a.name).localeCompare(String(b.name),'th'))}
function sortActiveRows_(rows){return rows.filter(x=>truthy_(x.active)).sort((a,b)=>num_(a.sortOrder)-num_(b.sortOrder)||String(a.name).localeCompare(String(b.name),'th'))}
function sheetValue_(key,value){if(value instanceof Date&&(key==='date'||/Date$/.test(String(key))))return Utilities.formatDate(value,Session.getScriptTimeZone(),'yyyy-MM-dd');return value}
function getSheetByNameFast_(ss,name){
  if(!ss)return null;
  if(!ss._sheetMap){
    ss._sheetMap={};
    const sheets=typeof ss.getSheets==='function'?ss.getSheets():[];
    (sheets||[]).forEach(sh=>{if(sh&&typeof sh.getName==='function')ss._sheetMap[sh.getName()]=sh;});
  }
  if(ss._sheetMap&&ss._sheetMap[name])return ss._sheetMap[name];
  if(typeof ss.getSheetByName==='function')return ss.getSheetByName(name);
  return null;
}
function readSheet_(ss,name){let sh=getSheetByNameFast_(ss,name);if(!sh){const expected=APP.sheets&&APP.sheets[name];if(expected&&ss&&typeof ss.insertSheet==='function'){try{sh=ss.insertSheet(name);sh.getRange(1,1,1,expected.length).setValues([expected]);sh.setFrozenRows(1);if(!ss._sheetMap)ss._sheetMap={};ss._sheetMap[name]=sh;return[]}catch(e){}}if(expected)return[];throw new Error('ไม่พบตาราง '+name)}const values=sh.getDataRange().getValues();if(values.length<2)return[];const h=values[0];return values.slice(1).filter(r=>r.some(v=>v!=='')).map(r=>h.reduce((o,k,i)=>(o[k]=sheetValue_(k,r[i]),o),{}))}
function rows_(name){return readSheet_(db_(),name)}
function append_(name,row){const ss=db_();let sh=getSheetByNameFast_(ss,name);if(!sh){const expected=APP.sheets&&APP.sheets[name];if(expected&&ss&&typeof ss.insertSheet==='function'){try{sh=ss.insertSheet(name);sh.getRange(1,1,1,expected.length).setValues([expected]);sh.setFrozenRows(1);if(!ss._sheetMap)ss._sheetMap={};ss._sheetMap[name]=sh}catch(e){}}}if(!sh)throw new Error('ไม่พบตาราง '+name);sh.appendRow(row)} function audit_(u,a,e,id,d){append_('AuditLog',[id_('LOG'),new Date(),u.id,a,e,id,d])}
function updateRow_(name,id,changes){const sh=getSheetByNameFast_(db_(),name);if(!sh)throw new Error('ไม่พบตาราง '+name);const values=sh.getDataRange().getValues(),h=values[0],index=values.findIndex((r,i)=>i>0&&r[0]===id);if(index<1)throw new Error('ไม่พบข้อมูลที่ต้องการแก้ไข');Object.keys(changes).forEach(k=>{const col=h.indexOf(k);if(col>=0)values[index][col]=changes[k]});sh.getRange(index+1,1,1,h.length).setValues([values[index]])}
function deleteRow_(name,id){const sh=getSheetByNameFast_(db_(),name);if(!sh)throw new Error('ไม่พบตาราง '+name);const values=sh.getDataRange().getValues(),index=values.findIndex((r,i)=>i>0&&r[0]===id);if(index<1)throw new Error('ไม่พบข้อมูลที่ต้องการลบ');sh.deleteRow(index+1)}
function id_(p){return p+'-'+Utilities.getUuid().slice(0,8).toUpperCase()} function hash_(s){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8))}
function clean_(v){return String(v==null?'':v).trim().slice(0,500)} function num_(v){const n=Number(v);return isFinite(n)?n:0} function round_(n){return Math.round((num_(n)+Number.EPSILON)*100)/100}
function positive_(v,label){const n=num_(v);if(n<=0)throw new Error(label+'ต้องมากกว่า 0');return round_(n)} function truthy_(v){return v===true||String(v).toLowerCase()==='true'||v===1}
function date_(v){if(!v)return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');const d=new Date(v);if(isNaN(d))throw new Error('วันที่ไม่ถูกต้อง');return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd')}
function monthInfo_(){const now=new Date(),year=Number(Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy')),month=Number(Utilities.formatDate(now,Session.getScriptTimeZone(),'MM')),last=new Date(year,month,0);return{key:year+'-'+String(month).padStart(2,'0'),dueDate:Utilities.formatDate(last,Session.getScriptTimeZone(),'yyyy-MM-dd')}}
function businessPayload_(ss){const projects=readSheet_(ss,'Projects'),income=readSheet_(ss,'Income'),expenses=readSheet_(ss,'Expenses'),overtime=readSheet_(ss,'Overtime'),dividends=readSheet_(ss,'Dividends'),projectSummaries=projects.map(p=>{const inc=income.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amountExVat),0),exp=expenses.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amountExVat),0),ot=overtime.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amount),0),dividendPaid=dividends.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amount),0);return Object.assign({},p,{income:round_(inc),expenses:round_(exp+ot),profit:round_(inc-exp-ot),otCost:round_(ot),dividendPaid:round_(dividendPaid),profitAfterDividend:round_(inc-exp-ot-dividendPaid)})}),totals=projectSummaries.reduce((a,p)=>({contract:a.contract+num_(p.contractValue),income:a.income+p.income,expenses:a.expenses+p.expenses,profit:a.profit+p.profit}),{contract:0,income:0,expenses:0,profit:0});return{partial:true,projects:projectSummaries,income,expenses,totals}}
function payrollPayload_(ss){const month=monthInfo_(),allEmployees=readSheet_(ss,'Employees'),employeeRows=allEmployees.filter(x=>truthy_(x.active)),advances=readSheet_(ss,'SalaryAdvances'),payments=readSheet_(ss,'SalaryPayments'),overtime=readSheet_(ss,'Overtime');const employees=employeeRows.map(e=>{const own=advances.filter(x=>x.employeeId===e.id&&x.monthKey===month.key),advanceTotal=round_(own.reduce((s,x)=>s+num_(x.amount),0)),employeeOt=overtime.filter(x=>x.employeeId===e.id&&String(x.date).slice(0,7)===month.key),otTotal=round_(employeeOt.reduce((s,x)=>s+num_(x.amount),0)),payment=payments.find(x=>x.employeeId===e.id&&x.monthKey===month.key),salary=num_(e.wageRate),earned=round_(salary+otTotal),remaining=payment?0:round_(Math.max(0,earned-advanceTotal)),progress=earned?round_(remaining/earned*100):0;return Object.assign({},e,{salary:{monthKey:month.key,dueDate:month.dueDate,baseSalary:salary,otTotal,totalEarned:earned,advanceTotal,remaining,progress,paid:!!payment,netPaid:payment?num_(payment.netPaid):0,advances:own}})}),employeeNames=allEmployees.reduce((o,e)=>(o[e.id]=e.name,o),{}),wageHistory=[...payments.map(x=>({id:x.id,type:'SALARY',date:x.paidDate||x.dueDate,monthKey:x.monthKey,employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',baseSalary:num_(x.baseSalary),otTotal:num_(x.otTotal),advanceTotal:num_(x.advanceTotal),amount:num_(x.netPaid),note:x.note||''})),...overtime.map(x=>({id:x.id,type:'OT',date:x.date,monthKey:String(x.date).slice(0,7),employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',projectId:x.projectId,hours:num_(x.hours),rate:num_(x.rate),amount:num_(x.amount),note:x.note||''})),...advances.map(x=>({id:x.id,type:'ADVANCE',date:x.date,monthKey:x.monthKey,employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',amount:num_(x.amount),note:x.note||''}))].sort((a,b)=>String(b.date).localeCompare(String(a.date)));return{partial:true,employees,overtime,wageHistory,payrollMonth:month}}
function dividendPayload_(ss){const dividends=readSheet_(ss,'Dividends'),partners=readSheet_(ss,'Partners').filter(x=>truthy_(x.active)).map(p=>Object.assign({},p,{totalPaid:round_(dividends.filter(x=>x.partnerId===p.id).reduce((s,x)=>s+num_(x.amount),0)),payments:dividends.filter(x=>x.partnerId===p.id).sort((a,b)=>String(b.date).localeCompare(String(a.date)))}));return Object.assign({},businessPayload_(ss),{partners,dividends})}
function lookupPayload_(ss){return{partial:true,categories:sortActiveRows_(readSheet_(ss,'Categories')),paymentMethods:sortActiveRows_(readSheet_(ss,'PaymentMethods')),balanceAccounts:sortActiveRows_(readSheet_(ss,'BalanceAccounts')),income:readSheet_(ss,'Income'),expenses:readSheet_(ss,'Expenses')}}
function migrateInvoiceNumbers_(){const props=PropertiesService.getScriptProperties(),key='INVOICE_ISSUE_DATE_NUMBERS_V1',lock=LockService.getScriptLock();lock.waitLock(30000);try{if(props.getProperty(key)==='done')return;const all=rows_('Invoices'),plan=all.filter(x=>/^(?:INVSO-\d{6}-\d+|IV\d{9,}-[12])$/.test(String(x.invoiceNo))&&quotationSequence_(x.quoteNo)>0).map(x=>({row:x,invoiceNo:relatedDocumentNumber_('IV',{quoteNo:x.quoteNo},Number(x.installment),x.issueDate)})),names=new Set();all.forEach(x=>{const number=(plan.find(p=>p.row.id===x.id)?.invoiceNo||String(x.invoiceNo)).toUpperCase();if(names.has(number))throw new Error('เลขใบแจ้งหนี้ซ้ำในการปรับรูปแบบ กรุณาตรวจสอบเอกสาร');names.add(number)});plan.filter(x=>x.row.invoiceNo!==x.invoiceNo).forEach(x=>{append_('AuditLog',[id_('AUD'),new Date(),'SYSTEM','BACKUP','INVOICE_NUMBER',x.row.id,JSON.stringify({invoiceNo:x.row.invoiceNo,newInvoiceNo:x.invoiceNo,quoteNo:x.row.quoteNo,issueDate:x.row.issueDate})]);updateRow_('Invoices',x.row.id,{invoiceNo:x.invoiceNo})});props.setProperty(key,'done')}finally{lock.releaseLock()}}
function invoiceRows_(ss){return readSheet_(ss,'Invoices').map(row=>({...row,snapshot:{...JSON.parse(row.snapshotJson||'{}'),items:JSON.parse(row.itemsJson||'[]'),logoData:row.logoData||''}})).sort((a,b)=>String(b.issueDate).localeCompare(String(a.issueDate))||String(b.createdAt).localeCompare(String(a.createdAt)))}
function receiptRows_(ss){return readSheet_(ss,'Receipts').map(row=>({...row,snapshot:{...JSON.parse(row.snapshotJson||'{}'),items:JSON.parse(row.itemsJson||'[]'),logoData:row.logoData||''}})).sort((a,b)=>String(b.issueDate).localeCompare(String(a.issueDate))||String(b.createdAt).localeCompare(String(a.createdAt)))}
function createReceipt_(user,d){
  require_(user,'FINANCE_EDIT');
  const paidDate=String(d.paidDate||''),paidTime=new Date(paidDate+'T00:00:00Z');if(!/^\d{4}-\d{2}-\d{2}$/.test(paidDate)||!Number.isFinite(paidTime.getTime())||paidTime.toISOString().slice(0,10)!==paidDate)throw new Error('กรุณาระบุวันที่รับชำระเงินให้ถูกต้อง');
  if(d.confirmPaid!==true)throw new Error('กรุณายืนยันว่าได้รับชำระเงินแล้ว');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=db_(),all=receiptRows_(ss),existing=all.find(x=>x.invoiceId===d.invoiceId);if(existing)return {partial:true,receipts:all,issuedReceiptId:existing.id};
    const invoice=readSheet_(ss,'Invoices').find(x=>x.id===d.invoiceId);if(!invoice)throw new Error('ไม่พบใบแจ้งหนี้ที่ต้องการออกใบเสร็จ');
    if(!Number.isFinite(Number(invoice.grandTotal))||Number(invoice.grandTotal)<=0||round_(Number(invoice.subtotal)+Number(invoice.vatAmount))!==Number(invoice.grandTotal))throw new Error('ยอดใบแจ้งหนี้ไม่ถูกต้อง');
    const receiptNo=relatedDocumentNumber_('RC',{quoteNo:invoice.quoteNo},Number(invoice.installment),paidDate);if(all.some(x=>String(x.receiptNo).toUpperCase()===receiptNo.toUpperCase()))throw new Error('เลขใบเสร็จนี้มีอยู่แล้ว กรุณาตรวจสอบใบแจ้งหนี้อ้างอิง');
    const id=id_('RCT'),row={...invoice,id,receiptNo,issueDate:paidDate,invoiceId:invoice.id,createdAt:new Date(),createdBy:user.id};
    append_('Receipts',APP.sheets.Receipts.map(key=>row[key]));audit_(user,'CREATE','RECEIPT',id,receiptNo+' / '+invoice.invoiceNo+' / '+paidDate);
    return {partial:true,receipts:receiptRows_(ss),issuedReceiptId:id};
  }finally{lock.releaseLock()}
}
function invoiceAmounts_(q,installment){const cents=key=>{const amount=Number(q[key]);if(!Number.isFinite(amount)||amount<0)throw new Error('ยอดใบเสนอราคาไม่ถูกต้อง');return Math.round(amount*100)},base=cents('subtotal'),vat=cents('vatAmount'),total=cents('grandTotal');if(base+vat!==total||total<=0)throw new Error('ยอดใบเสนอราคาไม่สมบูรณ์');const firstTotal=Math.round(total/2),firstBase=Math.round(base/2),firstVat=firstTotal-firstBase;return installment===1?{subtotal:firstBase/100,vatAmount:firstVat/100,grandTotal:firstTotal/100}:{subtotal:(base-firstBase)/100,vatAmount:(vat-firstVat)/100,grandTotal:(total-firstTotal)/100}}
function createInvoice_(user,d){
  require_(user,'FINANCE_EDIT');
  const installment=Number(d.installment);if(![1,2].includes(installment))throw new Error('เลือกงวดที่ 1 หรือ 2');
  const issueTime=new Date(String(d.issueDate)+'T00:00:00Z');if(!/^\d{4}-\d{2}-\d{2}$/.test(String(d.issueDate||''))||!Number.isFinite(issueTime.getTime())||issueTime.toISOString().slice(0,10)!==d.issueDate)throw new Error('กรุณาระบุวันที่ออกเอกสารให้ถูกต้อง');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=db_(),all=invoiceRows_(ss),related=all.filter(x=>x.quotationId===d.quotationId),existing=related.find(x=>Number(x.installment)===installment);
    if(existing)return {partial:true,invoices:all,issuedInvoiceId:existing.id};
    const q=related.length?related[0].snapshot:quotationRows_(ss).find(x=>x.id===d.quotationId);
    if(!q)throw new Error('กรุณาเลือกใบเสนอราคา');
    const amounts=invoiceAmounts_(q,installment),terms=clean_(d.paymentTerms)||(q.workType==='Solar Cell'?(installment===1?'งวดที่ 1 50% เมื่อออกใบ PO และนัดวันติดตั้ง':'งวดที่ 2 50% หลังติดตั้งออนระบบและส่งมอบงานแล้วเสร็จ ก่อนยื่นขนานไฟจาก PEA'):('งวดที่ '+installment+' 50%')),delivery=clean_(d.deliveryTerms)||'ระยะดำเนินการตามตกลงของหน้างานและสัญญาโครงการ';
    const invoiceNo=relatedDocumentNumber_('IV',q,installment,d.issueDate);if(all.some(x=>String(x.invoiceNo).toUpperCase()===invoiceNo.toUpperCase()))throw new Error('เลขใบแจ้งหนี้นี้มีอยู่แล้ว กรุณาตรวจเลขใบเสนอราคาอ้างอิง');
    const snapshot={...q};delete snapshot.items;delete snapshot.logoData;delete snapshot.itemsJson;delete snapshot.companyJson;
    const snapshotJson=JSON.stringify(snapshot),itemsJson=JSON.stringify(q.items||[]);if(snapshotJson.length>49000||itemsJson.length>49000)throw new Error('รายละเอียดใบเสนอราคายาวเกินไปสำหรับออกเอกสาร');
    const id=id_('INV'),row={id,invoiceNo,issueDate:d.issueDate,quotationId:d.quotationId,quoteNo:q.quoteNo,installment,percent:50,...amounts,vatRate:q.vatRate,paymentTerms:terms,deliveryTerms:delivery,poNo:clean_(d.poNo),snapshotJson,itemsJson,logoData:q.logoData||'',createdAt:new Date(),createdBy:user.id};
    append_('Invoices',APP.sheets.Invoices.map(key=>row[key]));audit_(user,'CREATE','INVOICE',id,invoiceNo+' / '+q.quoteNo+' / '+installment);
    return {partial:true,invoices:invoiceRows_(ss),issuedInvoiceId:id};
  }finally{lock.releaseLock()}
}

function updateInvoice_(user,d){
  require_(user,'FINANCE_EDIT');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=db_(),invoices=invoiceRows_(ss),inv=invoices.find(x=>x.id===d.id);
    if(!inv)throw new Error('ไม่พบใบแจ้งหนี้ที่ต้องการแก้ไข');
    const changes={};
    if(d.issueDate&&/^\d{4}-\d{2}-\d{2}$/.test(d.issueDate))changes.issueDate=d.issueDate;
    if(d.invoiceNo)changes.invoiceNo=clean_(d.invoiceNo).toUpperCase();
    if(d.poNo!==undefined)changes.poNo=clean_(d.poNo);
    if(d.paymentTerms!==undefined)changes.paymentTerms=clean_(d.paymentTerms);
    if(d.deliveryTerms!==undefined)changes.deliveryTerms=clean_(d.deliveryTerms);
    if(d.subtotal!==undefined)changes.subtotal=round_(d.subtotal);
    if(d.vatRate!==undefined)changes.vatRate=num_(d.vatRate);
    if(d.vatAmount!==undefined)changes.vatAmount=round_(d.vatAmount);
    if(d.grandTotal!==undefined)changes.grandTotal=round_(d.grandTotal);
    const snapshot={...inv.snapshot};let snapChanged=false;
    ['customerName','customerPhone','customerTaxId','customerAddress','projectName'].forEach(k=>{
      if(d[k]!==undefined&&d[k]!==snapshot[k]){snapshot[k]=clean_(d[k]);snapChanged=true}
    });
    if(snapChanged){delete snapshot.items;delete snapshot.logoData;changes.snapshotJson=JSON.stringify(snapshot)}
    updateRow_('Invoices',d.id,changes);
    audit_(user,'UPDATE','INVOICE',d.id,(changes.invoiceNo||inv.invoiceNo)+' แก้ไขข้อมูล');
    return {partial:true,invoices:invoiceRows_(ss)};
  }finally{lock.releaseLock()}
}

function deleteInvoice_(user,d){
  require_(user,'FINANCE_EDIT');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=db_(),receipts=receiptRows_(ss);
    if(receipts.some(r=>r.invoiceId===d.id))throw new Error('ไม่สามารถลบใบแจ้งหนี้นี้ได้ เนื่องจากมีใบเสร็จรับเงินอ้างอิงอยู่ กรุณาลบใบเสร็จก่อน');
    const inv=invoiceRows_(ss).find(x=>x.id===d.id);
    if(!inv)throw new Error('ไม่พบใบแจ้งหนี้ที่ต้องการลบ');
    deleteRow_('Invoices',d.id);
    audit_(user,'DELETE','INVOICE',d.id,inv.invoiceNo);
    return {partial:true,invoices:invoiceRows_(ss)};
  }finally{lock.releaseLock()}
}

function updateReceipt_(user,d){
  require_(user,'FINANCE_EDIT');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=db_(),receipts=receiptRows_(ss),rc=receipts.find(x=>x.id===d.id);
    if(!rc)throw new Error('ไม่พบใบเสร็จรับเงินที่ต้องการแก้ไข');
    const changes={};
    if(d.issueDate&&/^\d{4}-\d{2}-\d{2}$/.test(d.issueDate))changes.issueDate=d.issueDate;
    if(d.receiptNo)changes.receiptNo=clean_(d.receiptNo).toUpperCase();
    if(d.subtotal!==undefined)changes.subtotal=round_(d.subtotal);
    if(d.vatRate!==undefined)changes.vatRate=num_(d.vatRate);
    if(d.vatAmount!==undefined)changes.vatAmount=round_(d.vatAmount);
    if(d.grandTotal!==undefined)changes.grandTotal=round_(d.grandTotal);
    const snapshot={...rc.snapshot};let snapChanged=false;
    ['customerName','customerPhone','customerTaxId','customerAddress','projectName'].forEach(k=>{
      if(d[k]!==undefined&&d[k]!==snapshot[k]){snapshot[k]=clean_(d[k]);snapChanged=true}
    });
    if(snapChanged){delete snapshot.items;delete snapshot.logoData;changes.snapshotJson=JSON.stringify(snapshot)}
    updateRow_('Receipts',d.id,changes);
    audit_(user,'UPDATE','RECEIPT',d.id,(changes.receiptNo||rc.receiptNo)+' แก้ไขข้อมูล');
    return {partial:true,receipts:receiptRows_(ss)};
  }finally{lock.releaseLock()}
}

function deleteReceipt_(user,d){
  require_(user,'FINANCE_EDIT');
  const lock=LockService.getScriptLock();lock.waitLock(30000);
  try{
    const ss=db_(),rc=receiptRows_(ss).find(x=>x.id===d.id);
    if(!rc)throw new Error('ไม่พบใบเสร็จรับเงินที่ต้องการลบ');
    deleteRow_('Receipts',d.id);
    audit_(user,'DELETE','RECEIPT',d.id,rc.receiptNo);
    return {partial:true,receipts:receiptRows_(ss)};
  }finally{lock.releaseLock()}
}

function quotationPayload_(ss){return{partial:true,quotations:quotationRows_(ss),quotePresets:quotePresetRows_(ss)}}
function clientSafe_(value){return JSON.parse(JSON.stringify(value))}

// Parsed fields are already present; omit their duplicate storage representation in the initial response.
function initialClientPayload_(data){const result={...data};for(const key of ['quotations','quotePresets','invoices','receipts'])result[key]=(data[key]||[]).map(row=>{const copy={...row};delete copy.itemsJson;delete copy.companyJson;delete copy.dataJson;delete copy.snapshotJson;return copy});return result}
