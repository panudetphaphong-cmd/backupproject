const APP = {
  version: '2.6.0',
  maintenanceMode: true,
  spreadsheetId: '15wNWYPRNU3ozpuaQI4r_g2rYfwKSRwvWZV93jZGrDl4',
  sheets: {
    Users: ['id','name','username','passwordHash','role','permissions','active','createdAt'],
    Projects: ['id','name','customer','contractValue','vatRate','status','startDate','endDate','note','createdAt','createdBy'],
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
    AuditLog: ['id','date','userId','action','entity','entityId','detail']
  }
};

function doGet() {
  ensureSchemaOnce_();
  const template=HtmlService.createTemplateFromFile('Index');
  template.initialData=JSON.stringify(bootstrap_(maintenanceUser_())).replace(/</g,'\\u003c');
  return template.evaluate().setTitle('MB Solar Project Profit')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1, viewport-fit=cover');
}

function api(action, payload) {
  ensureSchemaOnce_();
  payload = payload || {};
  if (action === 'login') return clientSafe_(login_(payload));
  const user = APP.maintenanceMode && !payload.token ? maintenanceUser_() : sessionUser_(payload.token);
  if (action === 'bootstrap') return clientSafe_(bootstrap_(user));
  const routes = {saveProject:saveProject_,deleteProject:deleteProject_,saveIncome:saveIncome_,saveExpense:saveExpense_,saveExpensesBatch:saveExpensesBatch_,saveTransactionsBatch:saveTransactionsBatch_,updateTransaction:updateTransaction_,deleteTransaction:deleteTransaction_,saveEmployee:saveEmployee_,deleteEmployee:deleteEmployee_,saveOvertime:saveOvertime_,saveSalaryAdvance:saveSalaryAdvance_,payMonthlySalary:payMonthlySalary_,saveDividend:saveDividend_,distributeProjectDividend:distributeProjectDividend_,saveUser:saveUser_,saveCategory:saveCategory_,deleteCategory:deleteCategory_,savePaymentMethod:savePaymentMethod_,deletePaymentMethod:deletePaymentMethod_,saveBalanceAccount:saveBalanceAccount_,deleteBalanceAccount:deleteBalanceAccount_};
  if (!routes[action]) throw new Error('ไม่พบคำสั่งที่ร้องขอ');
  return clientSafe_(routes[action](user, payload.data || {}));
}

function ensureSchemaOnce_(){
  const props=PropertiesService.getScriptProperties();if(props.getProperty('MB_SCHEMA_VERSION')==='6')return;
  const ss=db_();Object.keys(APP.sheets).forEach(name=>{let sh=ss.getSheetByName(name);if(!sh)sh=ss.insertSheet(name);const expected=APP.sheets[name];if(sh.getLastRow()===0){sh.getRange(1,1,1,expected.length).setValues([expected]);sh.setFrozenRows(1)}else{const lastCol=Math.max(1,sh.getLastColumn()),existing=sh.getRange(1,1,1,lastCol).getValues()[0];expected.filter(h=>existing.indexOf(h)<0).forEach(h=>{sh.getRange(1,sh.getLastColumn()+1).setValue(h)})}sh.getRange(1,1,1,sh.getLastColumn()).setFontWeight('bold').setBackground('#dff5e8')});
  if(readSheet_(ss,'Partners').length===0){append_('Partners',['PART-BAS','บาส',true,new Date()]);append_('Partners',['PART-GOLF','กอล์ฟ',true,new Date()])}
  if(readSheet_(ss,'BalanceAccounts').length===0){append_('BalanceAccounts',['BAL-CASH','เงินสด','CASH',0,true,1,new Date(),'SYSTEM']);append_('BalanceAccounts',['BAL-COMPANY','บัญชีบริษัท','COMPANY',0,true,2,new Date(),'SYSTEM'])}
  props.setProperty('MB_SCHEMA_VERSION','6');
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

function bootstrap_(user) {
  const ss=db_(),projects=readSheet_(ss,'Projects'),income=readSheet_(ss,'Income'),expenses=readSheet_(ss,'Expenses'),employeeRows=readSheet_(ss,'Employees').filter(x=>truthy_(x.active)),overtime=readSheet_(ss,'Overtime'),salaryAdvances=readSheet_(ss,'SalaryAdvances'),salaryPayments=readSheet_(ss,'SalaryPayments'),month=monthInfo_(),dividends=readSheet_(ss,'Dividends'),partnerRows=readSheet_(ss,'Partners').filter(x=>truthy_(x.active));
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
  const banks=readSheet_(ss,'Banks').filter(x=>truthy_(x.active)),categories=sortActiveRows_(readSheet_(ss,'Categories')),paymentMethods=sortActiveRows_(readSheet_(ss,'PaymentMethods')),balanceAccounts=sortActiveRows_(readSheet_(ss,'BalanceAccounts'));
  const employeeNames=readSheet_(ss,'Employees').reduce((o,e)=>(o[e.id]=e.name,o),{}),wageHistory=[...salaryPayments.map(x=>({id:x.id,type:'SALARY',date:x.paidDate||x.dueDate,monthKey:x.monthKey,employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',baseSalary:num_(x.baseSalary),otTotal:num_(x.otTotal),advanceTotal:num_(x.advanceTotal),amount:num_(x.netPaid),note:x.note||''})),...overtime.map(x=>({id:x.id,type:'OT',date:x.date,monthKey:String(x.date).slice(0,7),employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',projectId:x.projectId,hours:num_(x.hours),rate:num_(x.rate),amount:num_(x.amount),note:x.note||''})),...salaryAdvances.map(x=>({id:x.id,type:'ADVANCE',date:x.date,monthKey:x.monthKey,employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',amount:num_(x.amount),note:x.note||''}))].sort((a,b)=>String(b.date).localeCompare(String(a.date)));
  return {appVersion:APP.version,user:publicUser_(user),permissions:permissions_(user),totals,projects:projectSummaries,income,expenses,employees,overtime,wageHistory,banks,categories,paymentMethods,balanceAccounts,payrollMonth:month,partners,dividends,users:allowed_(user,'USERS_MANAGE')?readSheet_(ss,'Users').map(publicUser_):[]};
}

function saveProject_(user,d) {
  require_(user,'PROJECT_EDIT'); const name=clean_(d.name),value=num_(d.contractValue); if(!name||value<0)throw new Error('กรุณากรอกชื่อโปรเจกต์และมูลค่างานให้ถูกต้อง');
  if(d.id){const old=project_(d.id);updateRow_('Projects',d.id,{name,customer:clean_(d.customer),contractValue:value,status:clean_(d.status)||'PLANNING',startDate:date_(d.startDate),endDate:date_(d.endDate),note:clean_(d.note)});audit_(user,'UPDATE','PROJECT',old.id,name);return businessPayload_(db_())}
  const id=id_('PRJ'); append_('Projects',[id,name,clean_(d.customer),value,7,clean_(d.status)||'PLANNING',date_(d.startDate),date_(d.endDate),clean_(d.note),new Date(),user.id]); audit_(user,'CREATE','PROJECT',id,name); return businessPayload_(db_());
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
  if(d.id){updateRow_('Employees',d.id,{name:clean_(d.name),position:clean_(d.position),type:clean_(d.type)||'PART_TIME',wageRate:num_(d.wageRate)});return payrollPayload_(ss)}
  const id=id_('EMP');ss.getSheetByName('Employees').appendRow([id,clean_(d.name),clean_(d.type)||'PART_TIME',num_(d.wageRate),num_(d.otRate),true,clean_(d.note),new Date(),clean_(d.position)]);return payrollPayload_(ss);
}
function deleteEmployee_(user,d){require_(user,'EMPLOYEE_MANAGE');const ss=db_(),employee=readSheet_(ss,'Employees').find(x=>x.id===d.id);if(!employee)throw new Error('ไม่พบพนักงาน');const hasHistory=readSheet_(ss,'Overtime').some(x=>x.employeeId===d.id)||readSheet_(ss,'SalaryAdvances').some(x=>x.employeeId===d.id)||readSheet_(ss,'SalaryPayments').some(x=>x.employeeId===d.id);if(hasHistory)updateRow_('Employees',d.id,{active:false});else deleteRow_('Employees',d.id);return payrollPayload_(ss)}
function saveOvertime_(user,d) {
  require_(user,'OT_EDIT'); project_(d.projectId); const employee=rows_('Employees').find(x=>x.id===d.employeeId); if(!employee)throw new Error('ไม่พบพนักงาน');
  const hours=positive_(d.hours,'จำนวนชั่วโมง'),rate=positive_(d.rate,'อัตรา OT'),amount=round_(hours*rate),id=id_('OT');
  append_('Overtime',[id,employee.id,d.projectId,date_(d.date),hours,rate,amount,clean_(d.note),false,new Date(),user.id]); audit_(user,'CREATE','OVERTIME',id,employee.name+' '+amount); return Object.assign({},businessPayload_(db_()),payrollPayload_(db_()));
}
function saveSalaryAdvance_(user,d){require_(user,'EMPLOYEE_MANAGE');const employee=rows_('Employees').find(x=>x.id===d.employeeId&&truthy_(x.active));if(!employee)throw new Error('ไม่พบพนักงาน');const month=monthInfo_(),amount=positive_(d.amount,'ยอดเบิก'),payments=rows_('SalaryPayments');if(payments.some(x=>x.employeeId===employee.id&&x.monthKey===month.key))throw new Error('พนักงานคนนี้ปิดรอบเงินเดือนแล้ว');const used=rows_('SalaryAdvances').filter(x=>x.employeeId===employee.id&&x.monthKey===month.key).reduce((s,x)=>s+num_(x.amount),0);if(amount>num_(employee.wageRate)-used)throw new Error('ยอดเบิกมากกว่าเงินเดือนคงเหลือ');const id=id_('ADV');append_('SalaryAdvances',[id,employee.id,month.key,date_(d.date),amount,clean_(d.note),new Date(),user.id]);audit_(user,'CREATE','SALARY_ADVANCE',id,employee.name+' '+amount);return payrollPayload_(db_())}
function payMonthlySalary_(user,d){require_(user,'EMPLOYEE_MANAGE');const employee=rows_('Employees').find(x=>x.id===d.employeeId&&truthy_(x.active));if(!employee)throw new Error('ไม่พบพนักงาน');const month=monthInfo_(),payments=rows_('SalaryPayments');if(payments.some(x=>x.employeeId===employee.id&&x.monthKey===month.key))throw new Error('ปิดรอบเงินเดือนนี้แล้ว');const advanceTotal=round_(rows_('SalaryAdvances').filter(x=>x.employeeId===employee.id&&x.monthKey===month.key).reduce((s,x)=>s+num_(x.amount),0)),otTotal=round_(rows_('Overtime').filter(x=>x.employeeId===employee.id&&String(x.date).slice(0,7)===month.key).reduce((s,x)=>s+num_(x.amount),0)),base=num_(employee.wageRate),net=round_(Math.max(0,base+otTotal-advanceTotal)),id=id_('SAL');append_('SalaryPayments',[id,employee.id,month.key,month.dueDate,base,advanceTotal,net,date_(d.paidDate),clean_(d.note),new Date(),user.id,otTotal]);audit_(user,'CREATE','SALARY_PAYMENT',id,employee.name+' '+net);return payrollPayload_(db_())}
function saveDividend_(user,d){require_(user,'FINANCE_EDIT');const partner=rows_('Partners').find(x=>x.id===d.partnerId&&truthy_(x.active));if(!partner)throw new Error('ไม่พบหุ้นส่วน');const project=project_(d.projectId);if(String(project.status).toUpperCase()!=='COMPLETED')throw new Error('เลือกปันผลได้เฉพาะโปรเจกต์ที่เสร็จแล้ว');const expenses=rows_('Expenses').filter(x=>x.projectId===project.id).reduce((s,x)=>s+num_(x.amountExVat),0),ot=rows_('Overtime').filter(x=>x.projectId===project.id).reduce((s,x)=>s+num_(x.amount),0),income=rows_('Income').filter(x=>x.projectId===project.id).reduce((s,x)=>s+num_(x.amountExVat),0),alreadyPaid=rows_('Dividends').filter(x=>x.projectId===project.id).reduce((s,x)=>s+num_(x.amount),0),available=round_(income-expenses-ot-alreadyPaid),amount=positive_(d.amount,'ยอดปันผล');if(amount>available)throw new Error('ยอดปันผลมากกว่ากำไรคงเหลือของโปรเจกต์');const id=id_('DIV');append_('Dividends',[id,partner.id,date_(d.date),amount,clean_(d.paidFrom),clean_(d.note),new Date(),user.id,project.id]);audit_(user,'CREATE','DIVIDEND',id,project.name+' · '+partner.name+' '+amount);return dividendPayload_(db_())}
function distributeProjectDividend_(user,d){require_(user,'FINANCE_EDIT');const project=project_(d.projectId);if(String(project.status).toUpperCase()!=='COMPLETED')throw new Error('ปันผลได้เฉพาะโปรเจกต์ที่เสร็จแล้ว');const partners=rows_('Partners').filter(x=>truthy_(x.active)&&(x.id==='PART-BAS'||x.id==='PART-GOLF'));if(partners.length!==2)throw new Error('ไม่พบข้อมูลหุ้นส่วนบาสและกอล์ฟครบทั้งสองคน');const expenses=rows_('Expenses').filter(x=>x.projectId===project.id).reduce((s,x)=>s+num_(x.amountExVat),0),ot=rows_('Overtime').filter(x=>x.projectId===project.id).reduce((s,x)=>s+num_(x.amount),0),income=rows_('Income').filter(x=>x.projectId===project.id).reduce((s,x)=>s+num_(x.amountExVat),0),alreadyPaid=rows_('Dividends').filter(x=>x.projectId===project.id).reduce((s,x)=>s+num_(x.amount),0),available=round_(income-expenses-ot-alreadyPaid);if(available<=0)throw new Error('โปรเจกต์นี้ไม่มีกำไรคงเหลือสำหรับปันผล');const first=round_(available/2),amounts=[first,round_(available-first)],paidFrom=clean_(d.paidFrom),source=rows_('BalanceAccounts').find(x=>truthy_(x.active)&&x.name===paidFrom);if(!source)throw new Error('กรุณาเลือกบัญชีเงินคงเหลือที่ใช้งานอยู่');partners.forEach((partner,i)=>{const id=id_('DIV'),amount=amounts[i];append_('Dividends',[id,partner.id,date_(d.date),amount,paidFrom,clean_(d.note)||'ปันผลกำไรโครงการ 50%',new Date(),user.id,project.id]);audit_(user,'CREATE','PROJECT_DIVIDEND_50',id,project.name+' · '+partner.name+' '+amount)});return dividendPayload_(db_())}
function saveUser_(user,d) {
  require_(user,'USERS_MANAGE'); if(!clean_(d.name)||!clean_(d.username)||String(d.password||'').length<4)throw new Error('กรุณากรอกข้อมูลผู้ใช้และรหัสผ่านอย่างน้อย 4 ตัว');
  if(rows_('Users').some(x=>String(x.username).toLowerCase()===clean_(d.username).toLowerCase()))throw new Error('ชื่อผู้ใช้นี้มีอยู่แล้ว');
  const id=id_('USR'),perms=Array.isArray(d.permissions)?d.permissions.join(','):''; append_('Users',[id,clean_(d.name),clean_(d.username),hash_(String(d.password)),clean_(d.role)||'MANAGER',perms,true,new Date()]); audit_(user,'CREATE','USER',id,d.name); return {partial:true,users:readSheet_(db_(),'Users').map(publicUser_)};
}

function saveCategory_(user,d){return saveLookup_(user,d,'Categories','CATEGORY','CAT')}
function deleteCategory_(user,d){return deleteLookup_(user,d,'Categories','CATEGORY')}
function savePaymentMethod_(user,d){return saveLookup_(user,d,'PaymentMethods','PAYMENT_METHOD','PAYBY')}
function deletePaymentMethod_(user,d){return deleteLookup_(user,d,'PaymentMethods','PAYMENT_METHOD')}
function saveBalanceAccount_(user,d){require_(user,'SETTINGS_MANAGE');const name=clean_(d.name),type=clean_(d.type);if(!name)throw new Error('กรุณากรอกชื่อบัญชี');if(type!=='CASH'&&type!=='COMPANY')throw new Error('ประเภทบัญชีไม่ถูกต้อง');const accounts=rows_('BalanceAccounts');if(accounts.some(x=>x.id!==d.id&&String(x.name).toLowerCase()===name.toLowerCase()&&truthy_(x.active)))throw new Error('มีชื่อบัญชีนี้อยู่แล้ว');if(d.id){const old=accounts.find(x=>x.id===d.id);if(!old)throw new Error('ไม่พบบัญชีที่ต้องการแก้ไข');updateRow_('BalanceAccounts',d.id,{name,type,openingBalance:round_(d.openingBalance),sortOrder:num_(d.sortOrder)||99,active:true});if(old.name!==name){['Income','Expenses'].forEach(sheet=>{const field=sheet==='Income'?'account':'paidBy';rows_(sheet).filter(x=>x[field]===old.name).forEach(x=>updateRow_(sheet,x.id,{[field]:name}))});rows_('Dividends').filter(x=>x.paidFrom===old.name).forEach(x=>updateRow_('Dividends',x.id,{paidFrom:name}));const method=rows_('PaymentMethods').find(x=>x.name===old.name&&truthy_(x.active));if(method)updateRow_('PaymentMethods',method.id,{name})}audit_(user,'UPDATE','BALANCE_ACCOUNT',d.id,name)}else{const id=id_('BAL');append_('BalanceAccounts',[id,name,type,round_(d.openingBalance),true,num_(d.sortOrder)||99,new Date(),user.id]);if(!rows_('PaymentMethods').some(x=>String(x.name).toLowerCase()===name.toLowerCase()&&truthy_(x.active)))append_('PaymentMethods',[id_('PAYBY'),name,true,num_(d.sortOrder)||99,new Date(),user.id]);audit_(user,'CREATE','BALANCE_ACCOUNT',id,name)}return lookupPayload_(db_())}
function deleteBalanceAccount_(user,d){require_(user,'SETTINGS_MANAGE');const item=rows_('BalanceAccounts').find(x=>x.id===d.id);if(!item)throw new Error('ไม่พบบัญชี');const used=rows_('Income').some(x=>x.account===item.name)||rows_('Expenses').some(x=>x.paidBy===item.name)||rows_('Dividends').some(x=>x.paidFrom===item.name);if(used)throw new Error('ลบบัญชีนี้ไม่ได้ เพราะมีรายการรับ–จ่ายหรือปันผลใช้งานอยู่ กรุณาแก้ไขชื่อแทน');updateRow_('BalanceAccounts',d.id,{active:false});const method=rows_('PaymentMethods').find(x=>x.name===item.name&&truthy_(x.active));if(method)updateRow_('PaymentMethods',method.id,{active:false});audit_(user,'DELETE','BALANCE_ACCOUNT',d.id,item.name);return lookupPayload_(db_())}
function saveLookup_(user,d,sheet,entity,prefix){require_(user,'SETTINGS_MANAGE');const name=clean_(d.name);if(!name)throw new Error('กรุณากรอกชื่อรายการ');if(rows_(sheet).some(x=>x.id!==d.id&&String(x.name).toLowerCase()===name.toLowerCase()&&truthy_(x.active)))throw new Error('มีชื่อนี้อยู่แล้ว');if(d.id){updateRow_(sheet,d.id,{name,sortOrder:num_(d.sortOrder)||99,active:true});audit_(user,'UPDATE',entity,d.id,name)}else{const id=id_(prefix);append_(sheet,[id,name,true,num_(d.sortOrder)||99,new Date(),user.id]);audit_(user,'CREATE',entity,id,name)}return lookupPayload_(db_())}
function deleteLookup_(user,d,sheet,entity){require_(user,'SETTINGS_MANAGE');const item=rows_(sheet).find(x=>x.id===d.id);if(!item)throw new Error('ไม่พบรายการ');updateRow_(sheet,d.id,{active:false});audit_(user,'DELETE',entity,d.id,item.name);return lookupPayload_(db_())}

function permissions_(u){const all=['DASHBOARD_VIEW','PROFIT_VIEW','PROJECT_EDIT','FINANCE_EDIT','EXPENSE_EDIT','EMPLOYEE_MANAGE','OT_EDIT','LOAN_USE','USERS_MANAGE','SETTINGS_MANAGE'];if(u.role==='OWNER'||u.permissions==='*')return all;return String(u.permissions||'').split(',').filter(Boolean)}
function allowed_(u,p){return permissions_(u).indexOf(p)>=0} function require_(u,p){if(!allowed_(u,p))throw new Error('คุณไม่มีสิทธิ์ทำรายการนี้')}
function publicUser_(u){return{id:u.id,name:u.name,username:u.username,role:u.role,permissions:String(u.permissions||''),active:truthy_(u.active)}}
function project_(id){const p=rows_('Projects').find(x=>x.id===id);if(!p)throw new Error('กรุณาเลือกโปรเจกต์');return p}
function sortedActive_(name){return rows_(name).filter(x=>truthy_(x.active)).sort((a,b)=>num_(a.sortOrder)-num_(b.sortOrder)||String(a.name).localeCompare(String(b.name),'th'))}
function sortActiveRows_(rows){return rows.filter(x=>truthy_(x.active)).sort((a,b)=>num_(a.sortOrder)-num_(b.sortOrder)||String(a.name).localeCompare(String(b.name),'th'))}
function readSheet_(ss,name){const sh=ss.getSheetByName(name);if(!sh)throw new Error('ไม่พบตาราง '+name);const lastRow=sh.getLastRow(),lastCol=sh.getLastColumn();if(lastRow<2||lastCol<1)return[];const values=sh.getRange(1,1,lastRow,lastCol).getValues(),h=values[0];return values.slice(1).filter(r=>r.some(v=>v!=='')).map(r=>h.reduce((o,k,i)=>(o[k]=r[i],o),{}))}
function rows_(name){return readSheet_(db_(),name)}
function append_(name,row){db_().getSheetByName(name).appendRow(row)} function audit_(u,a,e,id,d){append_('AuditLog',[id_('LOG'),new Date(),u.id,a,e,id,d])}
function updateRow_(name,id,changes){const sh=db_().getSheetByName(name),values=sh.getDataRange().getValues(),h=values[0],index=values.findIndex((r,i)=>i>0&&r[0]===id);if(index<1)throw new Error('ไม่พบข้อมูลที่ต้องการแก้ไข');Object.keys(changes).forEach(k=>{const col=h.indexOf(k);if(col>=0)values[index][col]=changes[k]});sh.getRange(index+1,1,1,h.length).setValues([values[index]])}
function deleteRow_(name,id){const sh=db_().getSheetByName(name),values=sh.getDataRange().getValues(),index=values.findIndex((r,i)=>i>0&&r[0]===id);if(index<1)throw new Error('ไม่พบข้อมูลที่ต้องการลบ');sh.deleteRow(index+1)}
function id_(p){return p+'-'+Utilities.getUuid().slice(0,8).toUpperCase()} function hash_(s){return Utilities.base64Encode(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256,s,Utilities.Charset.UTF_8))}
function clean_(v){return String(v==null?'':v).trim().slice(0,500)} function num_(v){const n=Number(v);return isFinite(n)?n:0} function round_(n){return Math.round((num_(n)+Number.EPSILON)*100)/100}
function positive_(v,label){const n=num_(v);if(n<=0)throw new Error(label+'ต้องมากกว่า 0');return round_(n)} function truthy_(v){return v===true||String(v).toLowerCase()==='true'||v===1}
function date_(v){if(!v)return Utilities.formatDate(new Date(),Session.getScriptTimeZone(),'yyyy-MM-dd');const d=new Date(v);if(isNaN(d))throw new Error('วันที่ไม่ถูกต้อง');return Utilities.formatDate(d,Session.getScriptTimeZone(),'yyyy-MM-dd')}
function monthInfo_(){const now=new Date(),year=Number(Utilities.formatDate(now,Session.getScriptTimeZone(),'yyyy')),month=Number(Utilities.formatDate(now,Session.getScriptTimeZone(),'MM')),last=new Date(year,month,0);return{key:year+'-'+String(month).padStart(2,'0'),dueDate:Utilities.formatDate(last,Session.getScriptTimeZone(),'yyyy-MM-dd')}}
function businessPayload_(ss){const projects=readSheet_(ss,'Projects'),income=readSheet_(ss,'Income'),expenses=readSheet_(ss,'Expenses'),overtime=readSheet_(ss,'Overtime'),dividends=readSheet_(ss,'Dividends'),projectSummaries=projects.map(p=>{const inc=income.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amountExVat),0),exp=expenses.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amountExVat),0),ot=overtime.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amount),0),dividendPaid=dividends.filter(x=>x.projectId===p.id).reduce((s,x)=>s+num_(x.amount),0);return Object.assign({},p,{income:round_(inc),expenses:round_(exp+ot),profit:round_(inc-exp-ot),otCost:round_(ot),dividendPaid:round_(dividendPaid),profitAfterDividend:round_(inc-exp-ot-dividendPaid)})}),totals=projectSummaries.reduce((a,p)=>({contract:a.contract+num_(p.contractValue),income:a.income+p.income,expenses:a.expenses+p.expenses,profit:a.profit+p.profit}),{contract:0,income:0,expenses:0,profit:0});return{partial:true,projects:projectSummaries,income,expenses,totals}}
function payrollPayload_(ss){const month=monthInfo_(),allEmployees=readSheet_(ss,'Employees'),employeeRows=allEmployees.filter(x=>truthy_(x.active)),advances=readSheet_(ss,'SalaryAdvances'),payments=readSheet_(ss,'SalaryPayments'),overtime=readSheet_(ss,'Overtime');const employees=employeeRows.map(e=>{const own=advances.filter(x=>x.employeeId===e.id&&x.monthKey===month.key),advanceTotal=round_(own.reduce((s,x)=>s+num_(x.amount),0)),employeeOt=overtime.filter(x=>x.employeeId===e.id&&String(x.date).slice(0,7)===month.key),otTotal=round_(employeeOt.reduce((s,x)=>s+num_(x.amount),0)),payment=payments.find(x=>x.employeeId===e.id&&x.monthKey===month.key),salary=num_(e.wageRate),earned=round_(salary+otTotal),remaining=payment?0:round_(Math.max(0,earned-advanceTotal)),progress=earned?round_(remaining/earned*100):0;return Object.assign({},e,{salary:{monthKey:month.key,dueDate:month.dueDate,baseSalary:salary,otTotal,totalEarned:earned,advanceTotal,remaining,progress,paid:!!payment,netPaid:payment?num_(payment.netPaid):0,advances:own}})}),employeeNames=allEmployees.reduce((o,e)=>(o[e.id]=e.name,o),{}),wageHistory=[...payments.map(x=>({id:x.id,type:'SALARY',date:x.paidDate||x.dueDate,monthKey:x.monthKey,employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',baseSalary:num_(x.baseSalary),otTotal:num_(x.otTotal),advanceTotal:num_(x.advanceTotal),amount:num_(x.netPaid),note:x.note||''})),...overtime.map(x=>({id:x.id,type:'OT',date:x.date,monthKey:String(x.date).slice(0,7),employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',projectId:x.projectId,hours:num_(x.hours),rate:num_(x.rate),amount:num_(x.amount),note:x.note||''})),...advances.map(x=>({id:x.id,type:'ADVANCE',date:x.date,monthKey:x.monthKey,employeeId:x.employeeId,employeeName:employeeNames[x.employeeId]||'พนักงานเดิม',amount:num_(x.amount),note:x.note||''}))].sort((a,b)=>String(b.date).localeCompare(String(a.date)));return{partial:true,employees,overtime,wageHistory,payrollMonth:month}}
function dividendPayload_(ss){const dividends=readSheet_(ss,'Dividends'),partners=readSheet_(ss,'Partners').filter(x=>truthy_(x.active)).map(p=>Object.assign({},p,{totalPaid:round_(dividends.filter(x=>x.partnerId===p.id).reduce((s,x)=>s+num_(x.amount),0)),payments:dividends.filter(x=>x.partnerId===p.id).sort((a,b)=>String(b.date).localeCompare(String(a.date)))}));return Object.assign({},businessPayload_(ss),{partners,dividends})}
function lookupPayload_(ss){return{partial:true,categories:sortActiveRows_(readSheet_(ss,'Categories')),paymentMethods:sortActiveRows_(readSheet_(ss,'PaymentMethods')),balanceAccounts:sortActiveRows_(readSheet_(ss,'BalanceAccounts'))}}
function clientSafe_(value){return JSON.parse(JSON.stringify(value))}
