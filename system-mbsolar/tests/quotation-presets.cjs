const fs = require('fs');
const vm = require('vm');
const assert = require('assert/strict');
const html = fs.readFileSync('src/Index.html', 'utf8');
const code = fs.readFileSync('src/Code.js', 'utf8');
new vm.Script(html.match(/<script>([\s\S]*?)<\/script>/)[1].replace('<?!= initialData ?>', '{}'));
const db = {QuotePresets: [], Quotations: [], AuditLog: []};
let serial = 0, releases = 0;
const server = {console, LockService:{getScriptLock:()=>({waitLock(){},releaseLock(){releases++;}})}};
vm.createContext(server);
vm.runInContext(code, server);
Object.assign(server, {
 require_(){}, db_:()=>db, rows_:name=>db[name], readSheet_:(_,name)=>db[name],
 id_:()=>`P${++serial}`, audit_(){},
 append_(name, values){const headers=vm.runInContext(`APP.sheets.${name}`,server);db[name].push(Object.fromEntries(headers.map((h,i)=>[h,values[i]])));},
 updateRow_(name,id,values){Object.assign(db[name].find(x=>x.id===id),values);},
 deleteRow_(name,id){db[name]=db[name].filter(x=>x.id!==id);}
});
const user={id:'TEST'};
const item={description:'Test <panel>',quantity:2,unit:'unit',unitPrice:1234.5};
const save=d=>server.saveQuotePreset_(user,d);
save({kind:'SET',workType:'EV Charger',name:'EV set',items:[item]});
save({kind:'SET',workType:'Solar Cell',name:'Solar set',items:[{...item,unitPrice:3000}]});
assert.equal(db.QuotePresets.length,2);
assert.equal(server.quotePresetRows_(db)[0].data.items[0].unitPrice,1234.5);
assert.throws(()=>save({kind:'SET',workType:'EV Charger',name:'Invalid',items:[{...item,quantity:0}]}));
assert.throws(()=>save({kind:'SET',workType:'EV Charger',name:'Invalid',items:[{...item,unitPrice:Infinity}]}));
assert.throws(()=>save({kind:'SET',workType:'EV Charger',name:'Invalid',items:Array(101).fill(item)}));
assert.throws(()=>save({id:'missing',kind:'SET',workType:'EV Charger',name:'Missing',items:[item]}));
assert.throws(()=>save({id:db.QuotePresets[0].id,kind:'SET',workType:'Solar Cell',name:'Wrong type',items:[item]}));
for(const type of ['EV Charger','Solar Cell'])save({kind:'TEMPLATE',workType:type,name:type,vatRate:7,warrantyText:type,paymentTerms:'',customerName:'MUST NOT SAVE',items:[item]});
save({kind:'TEMPLATE',workType:'Solar Cell',name:'Solar Cell',vatRate:0,warrantyText:'Solar updated'});
assert.equal(db.QuotePresets.length,4);
const templates=server.quotePresetRows_(db).filter(x=>x.kind==='TEMPLATE');
assert.equal(templates.find(x=>x.workType==='EV Charger').data.warrantyText,'EV Charger');
assert.equal(templates.find(x=>x.workType==='Solar Cell').data.warrantyText,'Solar updated');
assert(!('customerName' in templates[0].data));assert(!('items' in templates[0].data));
assert(releases>0);
let editor='',printed='';
const client={S:{data:{quotePresets:server.quotePresetRows_(db),quotations:[]}},QUOTE_DEFAULTS:{warranty:'EV WARRANTY',scope:'EV SCOPE'},structuredClone,
 esc:s=>String(s??'').replace(/&/g,'&amp;').replace(/</g,'&lt;').replace(/"/g,'&quot;'),
 money:n=>Number(n).toFixed(2),thaiDocumentDate:s=>s||'-',thaiBahtText:n=>'WORDS-'+n,today:()=> '2026-09-12',
 modal:(title,body)=>{editor=body;},setTimeout(){},window:{open:()=>({document:{write:s=>{printed=s;},close(){}},addEventListener(){},focus(){}})}
};
client.documentPanel=client.modal;vm.createContext(client);
const names=['quotationSolarStamp','quotationWarrantyText','quotationDocumentRows','quotationPackageRows','quotationDescriptionHtml','quotationCompanyAddress','quotationSignatureUpload','quotationProjectTitle','quotationCompany','quotationCompanyForm','quotationCompanyFromForm','fitQuotationPrint','quotationColumns','quotationLineAmount','quotationRowData','quotationPriceLabel','quotationFormDefaults','quotationSetPicker','quotationItemRow','openQuotation','quotationTotals','quotationNote','quotationHeader','quotationTerms','quotationSignatures','nl2br','printQuotation'];
vm.runInContext(html.split(/\r?\n/).filter(l=>names.some(n=>l.startsWith('function '+n+'('))).join('\n'),client);
client.openQuotation(null,'EV Charger');assert(editor.includes('EV set'));assert(!editor.includes('Solar set'));
assert(editor.includes('name="workType" value="EV Charger"'));assert(editor.includes('Test')===false);
client.openQuotation(null,'Solar Cell');assert(editor.includes('Solar set'));assert(!editor.includes('EV set'));assert(!editor.includes('EV WARRANTY'));
const old={id:'OLD',workType:'EV Charger',quoteNo:'QT-OLD',items:[item],customerName:'Original',warrantyText:'',scopeText:'',note:'<Note>',noteColor:'red',subtotal:2469,vatRate:7,vatAmount:172.83,grandTotal:2641.83};
client.S.data.quotations=[old];client.openQuotation('OLD');assert(editor.includes('Original'));assert(!editor.includes('EV WARRANTY'));
client.printQuotation('OLD');assert(!printed.includes('<section class="meta"'));assert(!printed.includes('quotationStatus'));
assert(printed.includes('WORDS-2641.83'));assert(printed.includes('grid-template-columns:5% 47% 7% 7% 16% 18%'));
assert(printed.indexOf('&lt;Note>')>printed.indexOf('WORDS-2641.83'));assert(printed.includes('text-align:left;margin:5px 0 6px;padding:0'));
assert(!printed.includes('<div class="words">'));
const before=JSON.stringify(old);server.deleteQuotePreset_(user,{id:db.QuotePresets[0].id});assert.equal(JSON.stringify(old),before);
client.S.data.quotePresets=[];assert.equal(client.quotationFormDefaults('Solar Cell').warrantyText,'');assert.equal(client.quotationFormDefaults('EV Charger').warrantyText,'EV WARRANTY');
const properties=new Map();server.PropertiesService={getScriptProperties:()=>({getProperty:k=>properties.get(k),setProperty:(k,v)=>properties.set(k,v)})};
const priorCount=db.QuotePresets.length;
server.seedEvPriceSets_();
const imported=()=>server.quotePresetRows_(db).filter(x=>x.id.startsWith('QPR-EV-FLEXXFAST-'));
assert.equal(imported().length,16);assert.equal(db.QuotePresets.length,priorCount+16);
const expected={120:[1590000,2590000,3690000,4690000],150:[1790000,2890000,4190000,5290000],180:[2290000,3490000,4990000,5990000],240:[2490000,3990000,5590000,6990000]};
for(const [kw,totals] of Object.entries(expected))totals.forEach((price,i)=>{
 const set=imported().find(x=>x.id===`QPR-EV-FLEXXFAST-${kw}-${i+1}`);
 assert.equal(set.workType,'EV Charger');assert.equal(set.data.items.length,1);
 assert.equal(set.data.items[0].quantity*set.data.items[0].unitPrice,price);
});
server.seedEvPriceSets_();assert.equal(imported().length,16);
const edited=imported()[0];save({id:edited.id,kind:'SET',workType:'EV Charger',name:'Edited package',items:[{...item,unitPrice:500}]});
const removed=imported()[1].id;server.deleteQuotePreset_(user,{id:removed});
server.seedEvPriceSets_();assert.equal(imported().length,15);
assert.equal(imported().find(x=>x.id===edited.id).data.items[0].unitPrice,500);
assert(!imported().some(x=>x.id===removed));
const pricesBefore=new Map(imported().map(x=>[x.id,x.data.items[0].unitPrice]));
// Simulate an existing production package using the previous 14-item text.
const legacy=imported().find(x=>x.id!==edited.id);
const legacyParts=legacy.id.match(/FLEXXFAST-(\d+)-(\d+)/);
legacy.data.items[0].description=server.evPackageDescriptionV1_(Number(legacyParts[1]),Number(legacyParts[2]));
server.updateRow_('QuotePresets',legacy.id,{dataJson:JSON.stringify(legacy.data)});
server.fillEvPackageDescriptions_();
for(const set of imported()){
 assert.equal(set.data.items[0].unitPrice,pricesBefore.get(set.id));
 if(set.id===edited.id){assert.equal(set.data.items[0].description,item.description);continue;}
 const [,kw,count]=set.id.match(/FLEXXFAST-(\d+)-(\d+)/);
 const text=set.data.items[0].description;
 assert(text.includes(kw+' kW'));assert(text.includes((Number(kw)*Number(count))+' kW'));
 assert.equal(text.split('\n').length,8);assert(text.length<5000);
 for(let i=1;i<=7;i++)assert(text.includes('\n'+i+'. '));
 assert(!text.includes('LED'));assert(!text.includes('14. '));
 assert(text.includes('CCTV'));assert(text.includes('Commissioning'));
}
const descriptionState=JSON.stringify(db);server.fillEvPackageDescriptions_();assert.equal(JSON.stringify(db),descriptionState);
assert(!imported().some(x=>x.id===removed));assert.equal(JSON.stringify(old),before);
for(const kw of [120,150,180,240])for(const count of [1,2,3,4])assert(server.evPackageDescription_(kw,count).length<5000);
console.log('PASS: preset CRUD, 16 package calculations/descriptions, prices unchanged, custom text/deletions/old quotes preserved');
// Explicit user-requested conversion includes formerly customized 120 kW text.
const totalsBefore=new Map(imported().map(x=>[x.id,x.data.items.reduce((s,i)=>s+server.quotationLineAmount_(i),0)]));
server.splitEvPackageRows_();
for(const set of imported()){
 assert.equal(set.data.items.length,7);
 const count=Number(set.id.split('-').at(-1));assert.equal(set.data.items[0].quantity,count);
 assert.equal(set.data.items.reduce((s,i)=>s+server.quotationLineAmount_(i),0),totalsBefore.get(set.id));
 assert.equal(set.data.items.reduce((s,i)=>s+client.quotationLineAmount(i),0),totalsBefore.get(set.id));
 assert(set.data.items.every(x=>!x.description.includes('\n')));
 save({id:set.id,kind:'SET',workType:'EV Charger',name:set.name,items:set.data.items});
 const saved=server.quotePresetRows_(db).find(x=>x.id===set.id);
 assert.equal(saved.data.items[0].packagePrice,true);
 assert.equal(saved.data.items.reduce((s,i)=>s+server.quotationLineAmount_(i),0),totalsBefore.get(set.id));
}
assert.equal(db.AuditLog.length,15);assert(!imported().some(x=>x.id===removed));
const splitState=JSON.stringify(db);server.splitEvPackageRows_();assert.equal(JSON.stringify(db),splitState);
const example=server.evPackageItems_(150,3,4190000);
assert.equal(example.reduce((s,i)=>s+client.quotationLineAmount(i),0),4190000);
assert.equal(client.quotationLineAmount({quantity:3,unitPrice:100}),300);
const rowHtml=client.quotationItemRow(example[0]);assert(rowHtml.includes('data-package-price="true"'));assert(rowHtml.includes('text-align:left'));
client.S.data.quotations=[{...old,items:example,subtotal:4190000,vatAmount:293300,grandTotal:4483300}];client.printQuotation('OLD');
assert.equal((printed.match(/class="quotation-subitem"/g)||[]).length,7);
console.log('PASS: 7 nested package details, exact totals, preset round-trip, legacy 120 conversion backup, left-aligned PDF');
server.date_=v=>v;
server.saveQuotation_(user,{quoteNo:'TEST-PACKAGE',issueDate:'2026-09-12',customerName:'Test',workType:'EV Charger',items:example,vatRate:7});
const savedQuote=server.quotationRows_(db).find(x=>x.quoteNo==='TEST-PACKAGE');
assert.equal(savedQuote.subtotal,4190000);assert.equal(savedQuote.vatAmount,293300);assert.equal(savedQuote.grandTotal,4483300);
assert(savedQuote.items[0].packagePrice);assert(savedQuote.items[1].packageIncluded);
assert.equal(savedQuote.items.length,7);console.log('PASS: saved quotation package total and VAT round-trip');
server.saveQuotation_(user,{id:savedQuote.id,quoteNo:savedQuote.quoteNo,issueDate:'2026-09-12',customerName:'Test',workType:'EV Charger',items:example,vatRate:7,company_name:'Custom <Company>',company_phone:'0999999999',company_email:'',company_taxId:'TAX',company_address:'Address',company_accountName:'Custom Account',company_bankName:'Bank',company_bankAccount:'123'});
const companyQuote=server.quotationRows_(db).find(x=>x.id===savedQuote.id);
assert.equal(companyQuote.company.name,'Custom <Company>');assert.equal(companyQuote.company.email,'');
save({kind:'TEMPLATE',workType:'EV Charger',name:'EV Charger',vatRate:7,company_name:'New template company'});
assert.equal(server.quotationRows_(db).find(x=>x.id===savedQuote.id).company.name,'Custom <Company>');
client.S.data.quotations=[companyQuote];client.printQuotation(companyQuote.id);
assert(printed.includes('Custom &lt;Company>'));assert(printed.includes('Custom Account'));assert(printed.includes('Bank'));assert(printed.includes('quotation-company-name'));assert(printed.includes('white-space:nowrap'));
assert(printed.includes('.num{text-align:center;white-space:nowrap;vertical-align:top}'));
client.openQuotation(companyQuote.id);assert(editor.includes('name="company_name"'));assert(editor.includes('Custom &lt;Company>'));
console.log('PASS: editable company fields, blank values, template/quote isolation and centered prices');
for(const kw of [120,150,180,240])for(const count of [1,2,3,4]){
 const title=client.quotationProjectTitle({workType:'EV Charger',items:server.evPackageItems_(kw,count,100)});
 assert.equal(title,'เสนอโครงการติดตั้งเครื่องอัดประจุความเร็วสูง '+kw+' kW จำนวน '+count+' ตู้');
}
assert.equal(client.quotationProjectTitle({workType:'Solar Cell',items:example}),'เสนอโครงการติดตั้งระบบ Solar Cell');
assert(client.quotationProjectTitle({items:[...server.evPackageItems_(120,1,100),...server.evPackageItems_(120,3,100)]}).includes('120 kW จำนวน 4 ตู้'));
const tableBody=printed.match(/<tbody>([\s\S]*?)<\/tbody>/)[1];
assert.equal((tableBody.match(/<td class="num"><\/td>/g)||[]).length,0);
assert(printed.includes('.center{text-align:center;white-space:nowrap;vertical-align:top}'));
assert(!printed.includes('<div class="part-label">ข้อมูลบริษัทและลูกค้า</div>'));
assert(!printed.includes('<div class="part-label">รายละเอียดและราคา</div>'));
assert(printed.includes('เสนอโครงการติดตั้งเครื่องอัดประจุความเร็วสูง 150 kW จำนวน 3 ตู้'));
assert(printed.includes('WORDS-4483300'));
console.log('PASS: 16 dynamic project headings, mixed sets, package row prices, centered cells and preserved totals');
const customerSignature=client.quotationSignatures(false,{customerName:'Customer <Name>'});
assert(customerSignature.includes('Customer &lt;Name>'));assert(!customerSignature.includes('ลงชื่อ'));assert(!customerSignature.includes('ชื่อ ....'));
assert(customerSignature.includes('max-width:180px'));assert(customerSignature.includes('height:30px'));
assert(printed.includes('<th style="text-align:center">รายละเอียดงาน</th>'));
const totalsHtml=client.quotationTotals({subtotal:100,vatRate:7,vatAmount:7,grandTotal:107});
assert(totalsHtml.includes('grid-column:2;text-align:center'));assert(totalsHtml.includes('grid-column:5;text-align:center'));
console.log('PASS: centered description header, aligned totals, compact signatures and escaped automatic customer name');
assert(client.quotationCompany().address.endsWith('44140'));
const oldCompanyAddress='เลขที่ 80 ม.10 ต.หัวขวาง อ.โกสุมพิสัย จ.มหาสารคาม';
assert.equal(client.quotationCompany({address:oldCompanyAddress}).address,oldCompanyAddress+' 44140');
assert.equal(client.quotationCompany({address:'Custom address 12345'}).address,'Custom address 12345');
assert.equal(client.quotationCompany({address:oldCompanyAddress+' 44140'}).address,oldCompanyAddress+' 44140');
assert(client.quotationHeader('',false,{}).includes('grid-template-columns:minmax(0,35fr) minmax(0,30fr) minmax(0,35fr)'));
console.log('PASS: postcode correction preserves custom addresses; customer block starts at price-column boundary');
assert(client.quotationCompanyAddress(oldCompanyAddress+' 44140').includes(' จ.มหาสารคาม 44140'));assert(!client.quotationCompanyAddress(oldCompanyAddress+' 44140').includes('<br>'));
assert.equal(client.quotationCompanyAddress('อำเภอพล\nจ.ขอนแก่น 40120'),'อำเภอพล จ.ขอนแก่น 40120');
assert.equal(client.quotationCompanyAddress('Custom address'),'Custom address');
assert.equal(client.quotationCompanyAddress('เลขที่ 80 ม.10\r\nต.หัวขวาง\nอ.โกสุมพิสัย\nจ.มหาสารคาม\n44140'),oldCompanyAddress+' 44140');
const signature='data:image/png;base64,iVBORw0KGgo=';
assert.equal(server.quotationCompany_({company_signatureData:signature}).signatureData,signature);
assert.throws(()=>server.quotationCompany_({company_signatureData:'https://example.com/image.png'}));
assert.throws(()=>server.quotationCompany_({company_signatureData:'data:image/svg+xml;base64,AAAA'}));
const signed=client.quotationSignatures(false,{customerName:'Test',company:{signatureData:signature}});
assert(signed.includes(signature));assert(signed.includes('margin-top:38px'));
save({kind:'TEMPLATE',workType:'EV Charger',name:'EV Charger',vatRate:7,company_signatureData:signature});
assert.equal(server.quotePresetRows_(db).find(x=>x.kind==='TEMPLATE'&&x.workType==='EV Charger').data.company.signatureData,signature);
assert.equal(server.quotationCompany_({company_signatureData:''}).signatureData,'');
console.log('PASS: province line break, centered header geometry, signature validation/render/template persistence and removal');
assert(!client.quotationSetPicker('EV Charger').includes('type="search"'));
assert(client.quotationSetPicker('EV Charger').includes('clearQuotationItems()'));
const clearFunctions=['clearQuotationItems','removeQuotationItem'];
vm.runInContext(html.split(/\r?\n/).filter(l=>clearFunctions.some(n=>l.startsWith('function '+n+'('))).join('\n'),client);
const body={innerHTML:'old rows',children:[{}]},picker={value:'old'};let refreshed=0;
client.$=id=>id==='quotationItems'?body:picker;client.window.confirm=()=>false;client.updateQuotationTotal=()=>refreshed++;client.renderQuotationPreview=()=>refreshed++;
client.clearQuotationItems();assert.equal(body.innerHTML,'old rows');
client.window.confirm=()=>true;client.clearQuotationItems();assert.equal(picker.value,'');assert(body.innerHTML.includes('<textarea'));assert.equal(refreshed,2);
client.removeQuotationItem({});assert.equal(refreshed,4);
assert(printed.includes('<footer class="paper-footer">'));assert(printed.includes('.paper-footer{margin-top:auto;'));
assert(client.quotationSignatures(false,{}).includes('วันที่ —'));
for(const type of ['EV Charger','Solar Cell'])for(const preview of [false,true]){const date='2026-09-12',signed=client.quotationSignatures(preview,{workType:type,issueDate:date});assert.equal(signed.split('วันที่ '+client.thaiDocumentDate(date)).length-1,2);assert(!signed.includes('............'));}
console.log('PASS: no search input, confirmed clear/cancel, removable final row, bottom footer and extended date blanks');
for(const kw of [120,150,180,240])for(let count=1;count<=4;count++){
 const source=server.evPackageItems_(kw,count,expected[kw][count-1]),snapshot=JSON.stringify(source),grouped=client.quotationPackageRows(source,'EV Charger');
 assert.equal(grouped.length,1);assert.equal(grouped[0].quantity,1);assert.equal(grouped[0].unit,'งาน');
 assert(grouped[0].description.startsWith('งานติดตั้งสถานีอัดประจุไฟฟ้าความเร็วสูง (DC Charge) ขนาด '+kw+' kW'));
 assert.equal(grouped[0].description.split('\n• ').length-1,7);
 assert(grouped[0].description.includes('— '+count+' เครื่อง'));
 assert.equal(client.quotationLineAmount(grouped[0]),expected[kw][count-1]);
 assert(client.quotationProjectTitle({items:grouped}).includes(kw+' kW จำนวน '+count+' ตู้'));
 assert.equal(JSON.stringify(source),snapshot);assert.equal(JSON.stringify(client.quotationPackageRows(grouped,'EV Charger')),JSON.stringify(grouped));
 assert.equal(client.quotationPackageRows(source,'Solar Cell'),source);
}
const nested=client.quotationPackageRows(example,'EV Charger');
server.saveQuotation_(user,{quoteNo:'TEST-NESTED',issueDate:'2026-09-12',customerName:'Nested',workType:'EV Charger',items:nested,vatRate:7});
const nestedSaved=server.quotationRows_(db).find(x=>x.quoteNo==='TEST-NESTED');
assert(nestedSaved);assert.equal(nestedSaved.items.length,1);assert.equal(nestedSaved.items[0].description,nested[0].description);assert.equal(nestedSaved.subtotal,4190000);assert.equal(nestedSaved.grandTotal,4483300);
assert.equal(client.quotationPackageRows([...example,...example],'EV Charger').length,2);
console.log('PASS: all 16 nested packages, quantities, prices, idempotence, original data, Solar isolation and saved nested quotation totals');
const legacyIP=example.map(x=>({...x,description:x.description.replace(/IP65/g,'IP55')}));
const revised=client.quotationPackageRows(legacyIP,'EV Charger'),rendered=client.quotationDocumentRows(revised);
assert(rendered.includes('IP65'));assert(!rendered.includes('IP55'));assert(legacyIP.some(x=>x.description.includes('IP55')));
assert.equal((rendered.match(/class="quotation-subitem"/g)||[]).length,7);
assert(!rendered.includes('— 3 เครื่อง'));assert(rendered.includes('>3</td>'));assert(rendered.includes('>เครื่อง</td>'));
assert.equal((rendered.match(/class="num"/g)||[]).length,2);
const continuousTerms=client.quotationTerms({warrantyText:'บรรทัดแรก\nบรรทัดถัดไป',scopeText:'ขอบเขต\r\nงาน'});
assert(continuousTerms.includes('บรรทัดแรก บรรทัดถัดไป'));assert(!continuousTerms.includes('<br>'));
assert(printed.includes('padding-bottom:12mm'));
console.log('PASS: IP65 legacy rendering, separate subitem quantities/units, continuous terms and raised signatures');
vm.runInContext(html.split(/\r?\n/).find(l=>l.startsWith('const QUOTE_DEFAULTS=')),client);
const warrantyDefault=vm.runInContext('QUOTE_DEFAULTS.warranty',client);
assert(warrantyDefault.includes('Onsite Service ภายในไม่เกิน 3 วัน หรือ 72 ชั่วโมง'));
assert.equal(client.quotationWarrantyText(warrantyDefault.split('\nOnsite Service')[0]),warrantyDefault);
assert.equal(client.quotationWarrantyText(warrantyDefault),warrantyDefault);
assert.equal(client.quotationWarrantyText('เงื่อนไขเฉพาะ'), 'เงื่อนไขเฉพาะ');
const newHeader=client.quotationHeader('data:image/png;base64,AAAA',false,{customerName:'Header customer'});
assert(newHeader.includes('class="quotation-parties"'));
assert(!newHeader.includes('margin-top:40px'));assert(!client.quotationHeader('',true,{}).includes('margin-top:20px'));
assert(newHeader.includes('quotation-logo-fit'));assert(newHeader.includes('position:absolute;inset:0'));assert(newHeader.includes('align-self:stretch'));assert(newHeader.includes('grid-template-columns:minmax(0,35fr) minmax(0,30fr) minmax(0,35fr)'));assert(newHeader.includes('Header customer'));
assert(newHeader.indexOf('quotation-company-block')<newHeader.indexOf('quotation-brand-heading'));
assert(newHeader.indexOf('quotation-brand-heading')<newHeader.indexOf('quotation-customer-block'));
assert(newHeader.includes('text-align:center;display:flex'));assert(newHeader.includes(';text-align:right'));assert(newHeader.includes('object-position:center top'));assert(!newHeader.includes('quotation-customer-strip'));
const compactHeader=client.quotationHeader('',false,{customerAddress:'Street\nProvince',customerName:'<Customer>'});
assert(compactHeader.includes('Street Province'));assert(compactHeader.includes('&lt;Customer>'));
console.log('PASS: onsite warranty preserved; left company, centered title/logo, right-aligned customer');
const beforeSolar=JSON.stringify(db.QuotePresets),solarCatalog=server.solarPriceCatalog_();
assert.equal(solarCatalog.length,17);assert.equal(solarCatalog.filter(x=>x.type==='On-grid').length,11);
const inverter16=solarCatalog.find(x=>x.type==='Hybrid On/Off-grid'&&x.kw===16);assert.equal(inverter16.price,429000);assert.equal(inverter16.phase,3);
assert(!solarCatalog.some(x=>x.type==='Hybrid On/Off-grid'&&x.kw===15));
const existingSolarTemplate=JSON.stringify(db.QuotePresets.find(x=>x.kind==='TEMPLATE'&&x.workType==='Solar Cell'));
server.seedSolarPriceSets_();const solarSets=server.quotePresetRows_(db).filter(x=>x.id.startsWith('QPR-SOLAR-'));assert.equal(solarSets.length,17);
assert.equal(JSON.stringify(db.QuotePresets.find(x=>x.kind==='TEMPLATE'&&x.workType==='Solar Cell')),existingSolarTemplate);
const countSolar=db.QuotePresets.length;server.seedSolarPriceSets_();assert.equal(db.QuotePresets.length,countSolar);
for(const spec of solarCatalog){const item=server.solarPackageItem_(spec);assert.equal(server.quotationLineAmount_(item),spec.price);assert.equal(client.quotationLineAmount(item),spec.price);assert(item.description.includes(spec.kw+' kW '+spec.phase+' Phase'));assert.equal(item.description.includes('Battery Solis'),spec.type==='Hybrid On/Off-grid');assert(item.description.length<5000);}
client.S.data.quotePresets=server.quotePresetRows_(db);const solarPicker=client.quotationSetPicker('Solar Cell');assert(solarPicker.includes('<optgroup label="On-grid">'));assert(solarPicker.includes('<optgroup label="Hybrid On/Off-grid">'));
const solarItem=server.solarPackageItem_(inverter16),solarRows=client.quotationDocumentRows([solarItem]);assert.equal((solarRows.match(/class="quotation-subitem"/g)||[]).length,11);assert(solarRows.includes('16 kW'));assert(!solarRows.includes('15 kW'));
server.saveQuotation_(user,{quoteNo:'TEST-SOLAR16',issueDate:'2026-09-12',customerName:'Solar Test',workType:'Solar Cell',items:[solarItem],vatRate:7});const solarQuote=server.quotationRows_(db).find(x=>x.quoteNo==='TEST-SOLAR16');assert.equal(solarQuote.subtotal,429000);assert.equal(solarQuote.vatAmount,30030);assert.equal(solarQuote.grandTotal,459030);
const deletedSolar=solarSets[0].id;server.deleteQuotePreset_(user,{id:deletedSolar});server.seedSolarPriceSets_();assert(!db.QuotePresets.some(x=>x.id===deletedSolar));
console.log('PASS: 17 Solar sets, 16 kW correction, before-VAT pricing, On-grid/Hybrid groups, template isolation, save totals, no reseed after deletion');
const evCompanyTemplateBefore=JSON.stringify(db.QuotePresets.filter(x=>x.kind==='TEMPLATE'&&x.workType==='EV Charger')),quotesBeforeCompany=JSON.stringify(db.Quotations);
const priorSolarData=JSON.parse(db.QuotePresets.find(x=>x.kind==='TEMPLATE'&&x.workType==='Solar Cell').dataJson),auditBeforeCompany=db.AuditLog.length;
server.migrateSolarCompany_();const solarCompanyTemplate=server.quotePresetRows_(db).find(x=>x.kind==='TEMPLATE'&&x.workType==='Solar Cell');
assert.equal(solarCompanyTemplate.data.company.name,'ห้างหุ้นส่วนจำกัด เอ็มบี โซล่า (สำนักงานใหญ่)');assert.equal(solarCompanyTemplate.data.company.taxId,'0403566004363');assert.equal(solarCompanyTemplate.data.company.phone,'095-238-3016');assert.equal(solarCompanyTemplate.data.company.bankAccount,'232-2-79355-6');assert.equal(solarCompanyTemplate.data.company.email,'');
assert.equal(solarCompanyTemplate.data.warrantyText,priorSolarData.warrantyText);assert.equal(JSON.stringify(db.Quotations),quotesBeforeCompany);assert.equal(JSON.stringify(db.QuotePresets.filter(x=>x.kind==='TEMPLATE'&&x.workType==='EV Charger')),evCompanyTemplateBefore);assert.equal(db.AuditLog.length,auditBeforeCompany+1);
server.migrateSolarCompany_();assert.equal(db.AuditLog.length,auditBeforeCompany+1);
client.S.data.quotePresets=server.quotePresetRows_(db);const newSolarDefaults=client.quotationFormDefaults('Solar Cell');assert.equal(newSolarDefaults.company.taxId,'0403566004363');
assert(client.quotationSignatures(false,newSolarDefaults).includes('นาย ณัฐฤทธิ์ นนทโคตร'));assert(!client.quotationSignatures(false,newSolarDefaults).includes('สุภาวดี'));
assert.equal(server.quotationCompany_({company_signerName:'Test signer'}).signerName,'Test signer');
console.log('PASS: MB Solar Excel company and signer, blank bank/email, backed-up template migration, preserved EV and existing quotations');
for(const [kw,count] of [[3,5],[5,8],[10,16]]){const spec=server.solarPriceCatalog_().find(x=>x.kw===kw);const generated=server.solarPackageItem_(spec);assert(generated.description.includes('— '+count+' แผง'));assert(client.quotationDocumentRows([generated]).includes('>'+count+'</td>'));}
assert.equal(server.solarPanelCount_(16),0);
const panelPreset=db.QuotePresets.find(x=>x.id==='QPR-SOLAR-HYBRID-10-1'),panelData=JSON.parse(panelPreset.dataJson);panelData.items[0].description=panelData.items[0].description.replace('— 16 แผง','(จำนวนแผงตามแบบที่ยืนยัน)');panelPreset.dataJson=JSON.stringify(panelData);
const panelOriginalQuotes=JSON.stringify(db.Quotations),panelPrice=panelData.items[0].unitPrice;
server.migrateSolarPanelCounts_();assert(JSON.parse(panelPreset.dataJson).items[0].description.includes('— 16 แผง'));assert.equal(JSON.parse(panelPreset.dataJson).items[0].unitPrice,panelPrice);assert.equal(JSON.stringify(db.Quotations),panelOriginalQuotes);
const panelAudit=db.AuditLog.length;server.migrateSolarPanelCounts_();assert.equal(db.AuditLog.length,panelAudit);
console.log('PASS: specified panel counts 3/5/10 kW, quantity columns, backed-up preset upgrade, unchanged prices and historical quotes');
const countsByKw={3:5,5:8,6:9,8:12,10:16,15:23,16:24,20:30,30:42,50:70,100:139,199:277};
for(const spec of server.solarPriceCatalog_()){const count=countsByKw[spec.kw],row=server.solarPackageItem_(spec);assert.equal(server.solarPanelCount_(spec.kw,spec.panel),count);assert(row.description.includes('PV Module '+spec.panel+' — '+count+' แผง'));assert.equal(row.unitPrice,spec.price);assert(!row.description.includes('จำนวนแผงตามแบบที่ยืนยัน'));}
assert.equal(server.solarPanelCount_(30,'Trina 720 W'),42);assert.equal(server.solarPanelCount_(199,'Trina 720 W'),277);
console.log('PASS: quotation panel counts for all 17 sets, correct AIKO/Trina wattages, prices unchanged');
const solarBankRow=db.QuotePresets.find(x=>x.kind==='TEMPLATE'&&x.workType==='Solar Cell'),solarBankData=JSON.parse(solarBankRow.dataJson);solarBankData.company.bankAccount='';solarBankData.company.phone='custom phone';solarBankRow.dataJson=JSON.stringify(solarBankData);
const bankQuotesBefore=JSON.stringify(db.Quotations),bankEvBefore=JSON.stringify(db.QuotePresets.filter(x=>x.workType==='EV Charger'));
server.migrateSolarBank_();const bankAfter=JSON.parse(solarBankRow.dataJson);assert.equal(bankAfter.company.bankAccount,'232-2-79355-6');assert.equal(bankAfter.company.bankName,'ธนาคารกสิกรไทย');assert.equal(bankAfter.company.accountName,'ห้างหุ้นส่วนจำกัด เอ็ม บี โซล่า');assert.equal(bankAfter.company.phone,'custom phone');assert.equal(JSON.stringify(db.Quotations),bankQuotesBefore);assert.equal(JSON.stringify(db.QuotePresets.filter(x=>x.workType==='EV Charger')),bankEvBefore);
const bankAudit=db.AuditLog.length;server.migrateSolarBank_();assert.equal(db.AuditLog.length,bankAudit);
console.log('PASS: Solar bank details from supplied image, other company fields, EV and historical quotes preserved');
const oldSolar=db.Quotations.find(x=>x.workType==='Solar Cell');oldSolar.companyJson=JSON.stringify({name:'MB saved',phone:'saved phone',bankAccount:''});const oldSolarTotal=oldSolar.grandTotal;
db.Quotations.push({...oldSolar,id:'KEEP-CUSTOM-BANK',companyJson:JSON.stringify({bankAccount:'CUSTOM-ACCOUNT'})});
const evBeforeSavedBank=JSON.stringify(db.Quotations.filter(x=>x.workType==='EV Charger'));server.migrateSavedSolarBanks_();assert.equal(JSON.parse(oldSolar.companyJson).bankAccount,'232-2-79355-6');assert.equal(JSON.parse(oldSolar.companyJson).phone,'saved phone');assert.equal(oldSolar.grandTotal,oldSolarTotal);assert.equal(JSON.parse(db.Quotations.find(x=>x.id==='KEEP-CUSTOM-BANK').companyJson).bankAccount,'CUSTOM-ACCOUNT');assert.equal(JSON.stringify(db.Quotations.filter(x=>x.workType==='EV Charger')),evBeforeSavedBank);
const savedBankAudit=db.AuditLog.length;server.migrateSavedSolarBanks_();assert.equal(db.AuditLog.length,savedBankAudit);
const solarStampHtml=client.quotationSignatures(false,{workType:'Solar Cell',issueDate:'2026-09-12',company:server.solarCompanyDefaults_()});assert(solarStampHtml.includes('ตราประทับ ห้างหุ้นส่วนจำกัด เอ็มบี โซล่า'));assert(!solarStampHtml.includes('บริษัท / ผู้มีอำนาจลงนาม'));assert(solarStampHtml.indexOf('ตราประทับ')>solarStampHtml.indexOf('<b>ห้างหุ้นส่วนจำกัด เอ็มบี โซล่า</b>'));assert(!client.quotationSignatures(false,{workType:'EV Charger'}).includes('ตราประทับ'));
console.log('PASS: saved Solar bank backfill with backup and custom-bank protection, stamp below company signature, EV unchanged');
for(const preview of [false,true]){const terms=client.quotationTerms({deliveryTerms:'ระยะดำเนินการตามตกลง',scopeText:'ระยะ\nดำเนินการ <script>'},preview);assert.equal(terms.split('white-space:nowrap;word-break:normal;overflow-wrap:normal">ระยะดำเนินการ</span>').length-1,2);assert(!terms.includes('<script>'));assert(terms.includes('overflow-wrap:break-word;word-break:normal'));}
console.log('PASS: Thai duration phrase kept together and term headings do not split in preview/PDF');
vm.runInContext(html.split(/\r?\n/).filter(l=>['quotationTypeFolder'].some(n=>l.startsWith('function '+n+'('))).join('\n'),client);
client.displayDate=v=>v||'';
client.quotationStatus=s=>s;const folderQuotes=[{id:'E1',quoteNo:'EV-ONLY',customerName:'Shared <customer>',workType:'EV Charger',grandTotal:100,items:[]},{id:'S1',quoteNo:'SOLAR-ONLY',customerName:'Shared <customer>',workType:'Solar Cell',grandTotal:200,items:[]}],folderSnapshot=JSON.stringify(folderQuotes);
const evFolder=client.quotationTypeFolder(folderQuotes,'EV Charger'),solarFolder=client.quotationTypeFolder(folderQuotes,'Solar Cell');assert(evFolder.includes('EV-ONLY'));assert(!evFolder.includes('SOLAR-ONLY'));assert(solarFolder.includes('SOLAR-ONLY'));assert(!solarFolder.includes('EV-ONLY'));assert(evFolder.includes('Shared &lt;customer>'));assert(solarFolder.includes('Shared &lt;customer>'));assert(evFolder.includes("printQuotation('E1')"));assert(solarFolder.includes("openQuotation('S1')"));assert.equal(JSON.stringify(folderQuotes),folderSnapshot);assert(client.quotationTypeFolder([],'Solar Cell').includes('ยังไม่มีใบเสนอราคาในโฟลเดอร์นี้'));
// Customer archive rendering is verified in document-archive.cjs.
console.log('PASS: EV and Solar parent folders, customer subfolders, separate totals/actions, empty states and unchanged documents');
