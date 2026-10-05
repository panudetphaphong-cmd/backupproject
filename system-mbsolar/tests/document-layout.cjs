const fs=require('fs'),vm=require('vm'),http=require('http'),path=require('path'),{spawn}=require('child_process');
const source=fs.readFileSync('src/Index.html','utf8');
const styles=(source.replace(/<script>[\s\S]*?<\/script>/g,'').match(/<style[^>]*>[\s\S]*?<\/style>/g)||[]).join('\n');
let content='',popup='';
const ctx={S:{data:{quotations:[],quotePresets:[]}},QUOTE_DEFAULTS:{warranty:'Warranty',scope:'Scope'},structuredClone,setTimeout(){},today:()=> '2026-09-12',money:n=>Number(n).toLocaleString('en-US',{minimumFractionDigits:2}),esc:s=>String(s??'').replace(/</g,'&lt;').replace(/"/g,'&quot;'),modal:(t,h,c)=>{content=h;popup=c;}};
ctx.documentPanel=ctx.modal;vm.createContext(ctx);
const names=['quotationSolarStamp','quotationWarrantyText','quotationDocumentRows','quotationPackageRows','quotationDescriptionHtml','quotationCompanyAddress','quotationSignatureUpload','quotationProjectTitle','quotationCompany','quotationCompanyForm','quotationCompanyFromForm','fitQuotationPrint','quotationColumns','quotationLineAmount','quotationRowData','quotationPriceLabel','quotationFormDefaults','quotationSetPicker','quotationItemRow','openQuotation','quotationLibraryTools','openQuotationSets'];
vm.runInContext(source.split(/\r?\n/).filter(l=>names.some(n=>l.startsWith('function '+n+'('))).join('\n'),ctx);
const backend={};vm.createContext(backend);vm.runInContext(fs.readFileSync('src/Code.js','utf8'),backend);
for(let i=1;i<=16;i++)ctx.S.data.quotePresets.push({id:'TEST'+i,kind:'SET',workType:'EV Charger',name:'Flexxfast 120 kW / '+i,data:{items:backend.evPackageItems_(120,3,3690000)}});
const server=http.createServer((req,res)=>{
 const mode=new URL(req.url,'http://localhost').searchParams.get('mode');
 if(mode==='sets')ctx.openQuotationSets('EV Charger');else ctx.openQuotation(null,'EV Charger',mode==='editset'?'SET':'',mode==='editset'?'TEST1':'');
 const body=mode==='library'?'<main id="page-documents">'+ctx.quotationLibraryTools()+'</main>':'<div class="swal2-popup '+popup+'" style="display:block;margin:16px auto"><div class="swal2-html-container">'+content+'</div></div>';
 res.end('<!doctype html><html><head><meta charset="utf-8"><meta name="viewport" content="width=device-width,initial-scale=1">'+styles+'</head><body style="margin:0;background:#eee">'+body+'<script>window.addEventListener("load",()=>{const f=document.querySelector(".quotation-editor");const controls=[...document.querySelectorAll("input:not([type=hidden]),select,textarea,button")].filter(x=>x.getBoundingClientRect().width);const bad=controls.filter(x=>{const r=x.getBoundingClientRect();return r.left<0||r.right>innerWidth+2});document.title=JSON.stringify({width:innerWidth,overflow:document.documentElement.scrollWidth>innerWidth+2,bad:bad.length,columns:f?getComputedStyle(f).gridTemplateColumns:null,fields:!!document.querySelector(".quotation-fields")});});</script></body></html>');
});
server.listen(0,'127.0.0.1',async()=>{
 try{for(const [mode,width] of [['editor',1400],['editor',600],['sets',1000],['editset',1000],['library',1000]]){
 const args=['--headless=new','--disable-gpu','--no-first-run','--disable-extensions','--user-data-dir='+path.resolve('tests/.edge-layout-profile'),'--window-size='+width+',1000','--virtual-time-budget=1500','--screenshot='+path.resolve('tests/layout-'+mode+'-'+width+'.png'),'--dump-dom','http://127.0.0.1:'+server.address().port+'/?mode='+mode];
 const child=spawn('C:/Program Files (x86)/Microsoft/Edge/Application/msedge.exe',args,{windowsHide:true});let output='';child.stdout.on('data',b=>output+=b);child.stderr.resume();await new Promise((resolve,reject)=>{child.on('error',reject);child.on('close',resolve)});
 const result=output.match(/<title>(.*?)<\/title>/)?.[1];console.log(mode,width,result||'NO METRICS');if(!result||JSON.parse(result).overflow||JSON.parse(result).bad)process.exitCode=1;
 }}catch(e){console.error(e);process.exitCode=1}finally{server.close()}
});
