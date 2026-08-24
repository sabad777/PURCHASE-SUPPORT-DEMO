(function(){
'use strict';
const E=window.PurchaseEngine;
let state={products:[],computed:[],filtered:[],reportDate:null,fileName:'',map:null,settings:loadSettings(),page:1,pageSize:50,sortKey:'priority',sortDir:1,selected:new Set(),selectedBrands:new Set(),selectedConditions:new Set(),activeView:'dashboard'};

const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
const fmt=(x,d=0)=>Number.isFinite(x)?x.toLocaleString(undefined,{maximumFractionDigits:d}):'—';
const dateFmt=d=>d instanceof Date&&!isNaN(d)?d.toISOString().slice(0,10):'—';
const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function loadSettings(){try{return Object.assign({},E.DEFAULT_SETTINGS,JSON.parse(localStorage.getItem('purchaseSettingsV2')||'{}'));}catch(_){return {...E.DEFAULT_SETTINGS};}}
function saveSettings(){localStorage.setItem('purchaseSettingsV2',JSON.stringify(state.settings));}
function toast(msg){const t=$('toast');t.textContent=msg;t.classList.remove('hidden');clearTimeout(toast.timer);toast.timer=setTimeout(()=>t.classList.add('hidden'),2600);}
function cover(x){return x===Infinity?'∞':fmt(x,1);}
function ratio(x){return x===Infinity?'∞':Number.isFinite(x)?x.toFixed(2):'—';}

function setView(name){
 state.activeView=name;
 document.querySelectorAll('.view').forEach(v=>v.classList.toggle('active',v.id===`view-${name}`));
 document.querySelectorAll('.nav button').forEach(b=>b.classList.toggle('active',b.dataset.view===name));
 const titles={dashboard:'Purchase Support',planner:'Purchase Planner',brands:'Brand Analysis',quality:'Data Quality',settings:'Calculation Settings'};
 $('pageTitle').textContent=titles[name]||'Purchase Support';
 if(name==='planner')renderPlanner();if(name==='brands')renderBrands();if(name==='quality')renderQuality();if(name==='settings')syncSettingsUI();
}
function showUploadZone(show){$('uploadZone').classList.toggle('hidden',!show);}

async function readFile(file){
 if(!window.XLSX)throw new Error('Excel reader library did not load. Check your internet connection and refresh.');
 const buf=await file.arrayBuffer();
 const wb=XLSX.read(buf,{type:'array',cellDates:true});
 const ws=wb.Sheets[wb.SheetNames[0]];
 const matrix=XLSX.utils.sheet_to_json(ws,{header:1,defval:'',raw:false,dateNF:'yyyy-mm-dd'});
 return E.parseMatrix(matrix);
}
async function handleFile(file){
 if(!file)return;
 try{
  $('fileChip').textContent='Reading '+file.name+'…';
  const parsed=await readFile(file);
  state.products=parsed.products;state.reportDate=parsed.reportDate;state.map=parsed.map;state.fileName=file.name;state.selected.clear();state.page=1;
  recompute();
  $('fileChip').textContent=file.name;
  $('reportMeta').textContent=`${fmt(state.products.length)} products • Data as of ${dateFmt(state.reportDate)} • Smart month-by-month + purchase-history analysis`;
  showUploadZone(false);$('exportTop').disabled=false;renderQuality();
  toast(`Loaded ${fmt(state.products.length)} products`);
 }catch(err){console.error(err);$('fileChip').textContent='No file loaded';alert(err.message||String(err));}
}

function demoProducts(){
 const d=new Date(2026,7,23);
 const make=(id,ref,oem,brand,desc,cat,stock,ow,ow2,vals,last,purchase=0,pCount=0,lastPurchase=null,salesCount=null)=>({
   _id:String(id),productId:String(id),internalRef:ref,oem,oemKey:E.normalizeOEM(oem),brand,brandKey:brand.toUpperCase(),brandPartNo:ref,description:desc,category:cat,
   motorlineStock:Math.floor(stock*.5),bavariaStock:Math.floor(stock*.2),tibaoStock:stock-Math.floor(stock*.5)-Math.floor(stock*.2),allCompany:stock,tibaoADRef:0,tibaoDXBRef:0,
   monthly:vals,totalSales:vals.slice(0,8).reduce((a,b)=>a+b,0),salesCount:salesCount??Math.max(1,Math.round(vals.slice(0,8).reduce((a,b)=>a+b,0)/3)),lastSaleDate:last?new Date(last):null,
   totalPurchase:purchase,purchaseCount:pCount,lastPurchaseDate:lastPurchase?new Date(lastPurchase):null,
   motorlineOnWay:ow,bavariaOnWay:0,tibaoOnWay:0,onWay:ow,motorlineOnWay2:0,bavariaOnWay2:0,tibaoOnWay2:ow2,onWay2:ow2,reportDate:d
 });
 return [
  make(1,'FEBI-1001','1K0698151','FEBI','BRAKE PAD SET FRONT','Brake System',2,0,0,[18,20,17,22,19,21,18,16,0,0,0,0],'2026-08-20',100,2,'2026-05-10',28),
  make(2,'TB-1001','1K0698151','TIBAO','BRAKE PAD SET FRONT','Brake System',28,10,0,[9,8,11,10,12,9,10,8,0,0,0,0],'2026-08-18',80,3,'2026-06-02',22),
  make(3,'BREM-1001','1K0698151','BREMBO','BRAKE PAD SET FRONT','Brake System',7,0,0,[12,13,12,14,11,15,14,9,0,0,0,0],'2026-08-22',60,2,'2026-04-18',31),
  make(4,'LEM-2200','8K0199381','LEMFORDER','ENGINE MOUNT RIGHT','Engine Mounting',0,0,0,[8,7,9,10,8,9,9,8,0,0,0,0],'2026-08-21',70,2,'2026-05-01',19),
  make(5,'TB-2200','8K0199381','TIBAO','ENGINE MOUNT RIGHT','Engine Mounting',3,6,10,[4,5,5,6,5,4,5,4,0,0,0,0],'2026-08-10',40,2,'2026-05-01',15),
  make(6,'MANN-03C','03C115561H','MANN FILTER','OIL FILTER','Filters',55,20,10,[22,24,28,26,25,24,27,20,0,0,0,0],'2026-08-23',200,4,'2026-07-10',55),
  make(7,'HENGST-03C','03C115561H','HENGST','OIL FILTER','Filters',12,0,0,[7,8,9,8,9,8,7,6,0,0,0,0],'2026-08-17',70,3,'2026-06-14',27),
  make(8,'FEBI-7788','4G0407183','FEBI','CONTROL ARM BUSH','Suspension',120,0,0,[1,1,0,1,0,0,1,0,0,0,0,0],'2025-04-18',160,2,'2026-01-20',4),
  make(9,'TB-7788','4G0407183','TIBAO','CONTROL ARM BUSH','Suspension',18,0,0,[3,2,4,3,4,3,4,3,0,0,0,0],'2026-08-11',35,2,'2026-05-11',14),
  make(10,'BREM-5555','5Q0615301F','BREMBO','BRAKE DISC FRONT','Brake System',4,12,0,[15,18,16,19,17,18,20,13,0,0,0,0],'2026-08-22',120,3,'2026-07-01',35),
  make(11,'FEBI-5555','5Q0615301F','FEBI','BRAKE DISC FRONT','Brake System',0,0,18,[10,9,11,12,10,11,12,8,0,0,0,0],'2026-08-19',90,2,'2026-05-12',26),
  make(12,'VDO-9090','A2C59506225','VDO','FUEL PUMP','Fuel System',0,0,0,[3,3,4,4,5,5,6,5,0,0,0,0],'2026-08-16',30,2,'2026-05-21',18),
  make(13,'TB-9090','A2C59506225','TIBAO','FUEL PUMP','Fuel System',14,0,0,[2,2,2,3,3,3,3,2,0,0,0,0],'2026-08-12',25,2,'2026-05-19',14),
  make(14,'FEBI-DEAD','7L0407182E','FEBI','SUSPENSION BUSH','Suspension',26,0,0,[2,0,0,0,0,0,0,0,0,0,0,0],'2025-07-01',50,1,'2026-01-10',1),
  make(15,'FEBI-SPIKE','03L115562','FEBI','OIL FILTER SPECIAL','Filters',3,0,0,[50,0,0,0,0,0,0,0,0,0,0,0],'2026-01-18',50,1,'2026-01-02',1),
  make(16,'TRUCK-8888','06H103495','TRUCKTEC','PCV VALVE','Engine',6,0,8,[6,7,6,8,7,8,9,7,0,0,0,0],'2026-08-15',60,2,'2026-05-15',20)
 ];
}
function loadDemo(){state.products=demoProducts();state.reportDate=new Date(2026,7,23);state.map={bavariaOnWay:1,tibaoOnWay:1,reportDate:1,groupOnWay:1,groupOnWay2:1,totalPurchase:1,purchaseCount:1,lastPurchaseDate:1};state.fileName='Demo data';state.selected.clear();recompute();$('fileChip').textContent='Demo data';$('reportMeta').textContent=`${state.products.length} demo products • Data as of ${dateFmt(state.reportDate)} • Smart logic v2.2`;showUploadZone(false);$('exportTop').disabled=false;toast('Demo data loaded');}

function recompute(){state.computed=E.calculate(state.products,state.settings,state.reportDate||new Date());populateFilters();applyFilters();renderDashboard();renderPlanner();renderBrands();renderQuality();}
function uniqueSorted(key){return [...new Set(state.computed.map(x=>x[key]).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b)));}
function fillSelect(id,values,label){const el=$(id),cur=el.value;el.innerHTML=`<option value="">${label}</option>`+values.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');if(values.includes(cur))el.value=cur;}

function pruneSelection(set,values){const allowed=new Set(values);[...set].forEach(v=>{if(!allowed.has(v))set.delete(v);});}
function updateMultiTrigger(type,values){
 const set=type==='brand'?state.selectedBrands:state.selectedConditions;
 const trigger=$(type==='brand'?'brandFilterTrigger':'conditionFilterTrigger');
 const allLabel=type==='brand'?'All Brands':'All Conditions';
 if(!set.size)trigger.textContent=allLabel;
 else if(set.size===1)trigger.textContent=[...set][0];
 else trigger.textContent=`${set.size} selected`;
 trigger.classList.toggle('has-selection',set.size>0);
}
function renderMultiMenu(type,values){
 const set=type==='brand'?state.selectedBrands:state.selectedConditions;
 const menu=$(type==='brand'?'brandFilterMenu':'conditionFilterMenu');
 const searchable=type==='brand';
 menu.innerHTML=`
  <div class="multi-head">
   ${searchable?'<input class="multi-search" type="text" placeholder="Search brands..." autocomplete="off">':''}
   <div class="multi-actions"><button type="button" data-action="all">Select all</button><button type="button" data-action="clear">Clear</button></div>
  </div>
  <div class="multi-options">${values.map(v=>`<label class="multi-option" data-label="${esc(String(v).toLowerCase())}"><input type="checkbox" value="${esc(v)}" ${set.has(v)?'checked':''}><span>${esc(v)}</span></label>`).join('')}</div>`;
 menu.querySelectorAll('.multi-option input').forEach(cb=>cb.onchange=()=>{cb.checked?set.add(cb.value):set.delete(cb.value);updateMultiTrigger(type,values);applyFilters();});
 menu.querySelector('[data-action="all"]').onclick=()=>{values.forEach(v=>set.add(v));renderMultiMenu(type,values);updateMultiTrigger(type,values);applyFilters();};
 menu.querySelector('[data-action="clear"]').onclick=()=>{set.clear();renderMultiMenu(type,values);updateMultiTrigger(type,values);applyFilters();};
 const search=menu.querySelector('.multi-search');
 if(search)search.oninput=()=>{const q=search.value.trim().toLowerCase();menu.querySelectorAll('.multi-option').forEach(x=>x.classList.toggle('hidden',q&&!x.dataset.label.includes(q)));};
 updateMultiTrigger(type,values);
}
function populateFilters(){
 const brands=uniqueSorted('brand');pruneSelection(state.selectedBrands,brands);renderMultiMenu('brand',brands);
 pruneSelection(state.selectedConditions,E.CONDITION_OPTIONS);renderMultiMenu('condition',E.CONDITION_OPTIONS);
 fillSelect('makeFilter',uniqueSorted('make'),'All Makes');
 fillSelect('movementFilter',E.MOVEMENT_OPTIONS,'All Movement');
 fillSelect('patternFilter',E.PATTERN_OPTIONS,'All Patterns');
 fillSelect('categoryFilter',uniqueSorted('category'),'All Categories');
}
function applyFilters(){
 const q=$('search').value.trim().toLowerCase(),
       desc=$('descriptionFilter').value.trim().toLowerCase(),
       make=$('makeFilter').value,
       mov=$('movementFilter').value,pat=$('patternFilter').value,cat=$('categoryFilter').value;
 state.filtered=state.computed.filter(p=>{
  if(state.selectedBrands.size&&!state.selectedBrands.has(p.brand))return false;
  if(state.selectedConditions.size&&!state.selectedConditions.has(p.condition))return false;
  if(make&&p.make!==make)return false;if(mov&&p.movement!==mov)return false;if(pat&&p.demandPattern!==pat)return false;if(cat&&p.category!==cat)return false;
  if(desc&&!String(p.description||'').toLowerCase().includes(desc))return false;
  if(q){const hay=[p.internalRef,p.oem,p.brand,p.brandPartNo,p.make,p.category].join(' ').toLowerCase();if(!hay.includes(q))return false;}
  return true;
 });
 state.page=1;renderDashboard();renderPlanner();
}
function count(fn){return state.filtered.filter(fn).length;}

function renderDashboard(){
 if(!state.computed.length){['kpiCritical','kpiReorder','kpiWait','kpiFast','kpiDead','kpiSuggested'].forEach(id=>$(id).textContent='0');$('statusBars').innerHTML='<div class="empty">Upload data to see purchasing intelligence.</div>';$('dashboardTable').innerHTML='';return;}
 $('kpiCritical').textContent=fmt(count(p=>p.condition==='CRITICAL ORDER'));
 $('kpiReorder').textContent=fmt(count(p=>p.condition==='REORDER SOON'));
 $('kpiWait').textContent=fmt(count(p=>p.condition==='WAIT INCOMING'));
 $('kpiFast').textContent=fmt(count(p=>p.movement==='FAST MOVING'));
 $('kpiDead').textContent=fmt(count(p=>p.condition==='DEAD STOCK'));
 $('kpiSuggested').textContent=fmt(state.filtered.reduce((s,p)=>s+p.suggestedQty,0));
 const statuses=['CRITICAL ORDER','REORDER SOON','WAIT INCOMING','OK','OVERSTOCK','DORMANT / REVIEW','ONE-TIME / REVIEW','DEAD STOCK'];
 const max=Math.max(1,...statuses.map(s=>count(p=>p.condition===s)));
 $('statusBars').innerHTML=statuses.map(s=>{const c=count(p=>p.condition===s);return `<div class="status-line"><span>${esc(s)}</span><div class="bar"><span class="bar-${conditionClass(s)}" style="width:${(c/max)*100}%"></span></div><strong class="num">${fmt(c)}</strong></div>`}).join('');
 const rows=[...state.filtered].sort((a,b)=>a.priority-b.priority||b.suggestedQty-a.suggestedQty||b.demandRate-a.demandRate).slice(0,12);renderCompactTable($('dashboardTable'),rows);
}

function conditionClass(s){if(s==='CRITICAL ORDER')return'critical';if(s==='REORDER SOON')return'reorder';if(s==='WAIT INCOMING')return'wait';if(s==='OK')return'ok';if(s==='DEAD STOCK')return'dead';if(s==='OVERSTOCK')return'overstock';if(s==='DORMANT / REVIEW')return'dormant';if(s==='ONE-TIME / REVIEW')return'spike';return'review';}
function movementClass(s){if(s==='FAST MOVING')return'fast';if(s==='MEDIUM MOVING')return'medium';if(s==='SLOW MOVING')return'slow';if(s==='DEAD STOCK')return'dead';return'review';}
function confidenceClass(s){return s==='HIGH'?'confidence-high':s==='MEDIUM'?'confidence-medium':'confidence-review';}
function patternClass(s){if(s==='REGULAR')return'ok';if(s==='RISING')return'fast';if(s==='FALLING')return'reorder';if(s==='INTERMITTENT')return'wait';if(s==='DEAD STOCK')return'dead';if(s==='DORMANT')return'dormant';if(s==='ONE-TIME SPIKE')return'spike';return'review';}

function renderCompactTable(el,rows){
 el.innerHTML=`<thead><tr><th>Internal Ref</th><th>Add Description</th><th>Make</th><th>OEM</th><th>Brand</th><th class="num">All Company</th><th class="num">On Way</th><th class="num">On Way 2</th><th>Demand Pattern</th><th class="num">Demand / Mo</th><th class="num">Active Months</th><th class="num">Purchase Qty</th><th>Same OEM Other Brands</th><th class="num">Cover</th><th>Condition</th><th class="num">Suggested Qty</th><th>Action</th></tr></thead><tbody>${rows.map(p=>`<tr><td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(p.internalRef||'—')}</button></td><td class="desc" title="${esc(p.description)}">${esc(p.description||'—')}</td><td>${esc(p.make||'UNKNOWN')}</td><td>${esc(p.oem||'—')}</td><td>${esc(p.brand)}</td><td class="num">${fmt(p.allCompany)}</td><td class="num">${fmt(p.onWay)}</td><td class="num">${fmt(p.onWay2)}</td><td><span class="badge ${patternClass(p.demandPattern)}">${esc(p.demandPattern)}</span></td><td class="num">${fmt(p.demandRate,1)}</td><td class="num">${p.activeMonths}/${p.monthCount}</td><td class="num">${fmt(p.totalPurchase)}</td><td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(p.equivalentNote)}</button></td><td class="num">${cover(p.currentCover)}</td><td><span class="badge ${conditionClass(p.condition)}">${esc(p.condition)}</span></td><td class="num qty">${fmt(p.suggestedQty)}</td><td>${esc(p.action)}</td></tr>`).join('')||'<tr><td colspan="17" class="empty">No items match the current filters.</td></tr>'}</tbody>`;
 bindDetailLinks(el);
}

const plannerColumns=[
 ['internalRef','Internal Ref'],['description','Add Description'],['make','Make'],['oem','OEM'],['brand','Brand'],['brandPartNo','Brand Part No.'],
 ['allCompany','All Company Qty'],['onWay','On Way'],['onWay2','On Way 2'],['totalSales','Sales Qty'],['salesCount','Sales Count'],['activeMonths','Sales Months'],
 ['avgMonthlySales','Overall Avg'],['recent3Avg','Recent 3M'],['recent6Avg','Recent 6M'],['demandRate','Smart Demand'],['demandPattern','Demand Pattern'],['demandConfidence','Confidence'],
 ['totalPurchase','Purchase Qty'],['purchaseCount','Purchase Count'],['purchaseSalesRatio','Purch/Sales'],['purchaseSignal','Purchase Signal'],
 ['equivalentNote','Same OEM Other Brands'],['currentCover','Current Cover'],['pipelineCover','Pipeline Cover'],['condition','Stock Condition'],['priority','Priority'],['suggestedQty','Suggested Purchase'],['action','Purchase Action']
];
function plannerCell(p,k){
 let v=p[k];
 if(k==='internalRef')return`<td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(v||'—')}</button></td>`;
 if(k==='description')return`<td class="desc" title="${esc(v)}">${esc(v||'—')}</td>`;
 if(['allCompany','onWay','onWay2','totalSales','salesCount','totalPurchase','purchaseCount','suggestedQty'].includes(k))return`<td class="num ${k==='suggestedQty'?'qty':''}">${fmt(v)}</td>`;
 if(k==='activeMonths')return`<td class="num">${p.activeMonths}/${p.monthCount}</td>`;
 if(['avgMonthlySales','recent3Avg','recent6Avg','demandRate'].includes(k))return`<td class="num">${fmt(v,1)}</td>`;
 if(['currentCover','pipelineCover'].includes(k))return`<td class="num">${cover(v)}</td>`;
 if(k==='purchaseSalesRatio')return`<td class="num">${ratio(v)}</td>`;
 if(k==='movement')return`<td><span class="badge ${movementClass(v)}">${esc(v)}</span></td>`;
 if(k==='demandPattern')return`<td><span class="badge ${patternClass(v)}">${esc(v)}</span></td>`;
 if(k==='demandConfidence')return`<td><span class="badge ${confidenceClass(v)}">${esc(v)}</span></td>`;
 if(k==='condition')return`<td><span class="badge ${conditionClass(v)}">${esc(v)}</span></td>`;
 if(k==='priority')return`<td><span class="priority p${v}">${v}</span></td>`;
 if(k==='equivalentNote')return`<td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(v)}</button></td>`;
 if(k==='purchaseSignal')return`<td><span class="signal ${v==='OVERBUYING RISK'?'signal-warn':''}">${esc(v)}</span></td>`;
 return`<td>${esc(v||'—')}</td>`;
}
function renderPlanner(){
 if(!state.computed.length){$('plannerTable').innerHTML='<tbody><tr><td class="empty">Upload data first.</td></tr></tbody>';$('plannerResults').textContent='0 items';return;}
 const sorted=[...state.filtered].sort((a,b)=>sortCompare(a,b,state.sortKey,state.sortDir));
 const pages=Math.max(1,Math.ceil(sorted.length/state.pageSize));state.page=Math.min(state.page,pages);
 const start=(state.page-1)*state.pageSize,rows=sorted.slice(start,start+state.pageSize);
 $('plannerResults').textContent=`${fmt(sorted.length)} items • ${fmt(sorted.reduce((s,p)=>s+p.suggestedQty,0))} pcs suggested`;
 $('pageInfo').textContent=`Page ${state.page} of ${pages} • showing ${rows.length}`;
 $('prevPage').disabled=state.page<=1;$('nextPage').disabled=state.page>=pages;
 $('plannerTable').innerHTML=`<thead><tr><th><input type="checkbox" id="selectPage"></th>${plannerColumns.map(([k,l])=>`<th data-sort="${k}" class="${['allCompany','onWay','onWay2','totalSales','salesCount','activeMonths','avgMonthlySales','recent3Avg','recent6Avg','demandRate','totalPurchase','purchaseCount','purchaseSalesRatio','currentCover','pipelineCover','priority','suggestedQty'].includes(k)?'num':''}">${esc(l)}${state.sortKey===k?(state.sortDir>0?' ↑':' ↓'):''}</th>`).join('')}</tr></thead><tbody>${rows.map(p=>`<tr><td><input class="row-select" type="checkbox" data-id="${esc(p._id)}" ${state.selected.has(p._id)?'checked':''}></td>${plannerColumns.map(([k])=>plannerCell(p,k)).join('')}</tr>`).join('')||'<tr><td colspan="30" class="empty">No items match filters.</td></tr>'}</tbody>`;
 $('plannerTable').querySelectorAll('th[data-sort]').forEach(th=>th.onclick=()=>{const k=th.dataset.sort;if(state.sortKey===k)state.sortDir*=-1;else{state.sortKey=k;state.sortDir=1;}renderPlanner();});
 $('plannerTable').querySelectorAll('.row-select').forEach(cb=>cb.onchange=()=>{cb.checked?state.selected.add(cb.dataset.id):state.selected.delete(cb.dataset.id);updateSelectedButton();});
 const sp=$('selectPage');if(sp)sp.onchange=()=>{rows.forEach(p=>sp.checked?state.selected.add(p._id):state.selected.delete(p._id));renderPlanner();updateSelectedButton();};
 bindDetailLinks($('plannerTable'));updateSelectedButton();
}
function sortCompare(a,b,k,d){let av=a[k],bv=b[k];if(typeof av==='string'||typeof bv==='string')return String(av??'').localeCompare(String(bv??''))*d;av=Number.isFinite(av)?av:-Infinity;bv=Number.isFinite(bv)?bv:-Infinity;return(av-bv)*d;}
function updateSelectedButton(){$('exportSelected').disabled=state.selected.size===0;$('exportSelected').textContent=`Export Selected (${state.selected.size})`;}

function brandSummary(){
 const m=new Map();state.computed.forEach(p=>{if(!m.has(p.brand))m.set(p.brand,{brand:p.brand,products:0,stock:0,fast:0,critical:0,reorder:0,wait:0,overstock:0,dead:0,review:0,suggested:0});const b=m.get(p.brand);b.products++;b.stock+=p.allCompany;b.fast+=p.movement==='FAST MOVING';b.critical+=p.condition==='CRITICAL ORDER';b.reorder+=p.condition==='REORDER SOON';b.wait+=p.condition==='WAIT INCOMING';b.overstock+=p.condition==='OVERSTOCK';b.dead+=p.condition==='DEAD STOCK';b.review+=['DORMANT / REVIEW','ONE-TIME / REVIEW','NO SALES / REVIEW'].includes(p.condition);b.suggested+=p.suggestedQty;});return[...m.values()].sort((a,b)=>b.suggested-a.suggested||b.critical-a.critical||a.brand.localeCompare(b.brand));
}
function renderBrands(){
 if(!state.computed.length){$('brandGrid').innerHTML='<div class="empty">Upload data to analyse brands.</div>';$('brandTable').innerHTML='';return;}
 const rows=brandSummary();
 $('brandGrid').innerHTML=rows.slice(0,12).map(b=>`<div class="brand-card" data-brand="${esc(b.brand)}"><div class="brand-name">${esc(b.brand)}</div><div class="brand-meta">${fmt(b.products)} products • ${fmt(b.stock)} pcs company stock</div><div class="brand-numbers"><div><strong>${fmt(b.critical)}</strong><span>Critical</span></div><div><strong>${fmt(b.dead)}</strong><span>Dead Stock</span></div><div><strong>${fmt(b.suggested)}</strong><span>Suggest Qty</span></div></div></div>`).join('');
 $('brandGrid').querySelectorAll('.brand-card').forEach(c=>c.onclick=()=>selectBrand(c.dataset.brand));
 $('brandTable').innerHTML=`<thead><tr><th>Brand</th><th class="num">Products</th><th class="num">Company Stock</th><th class="num">Fast</th><th class="num">Critical</th><th class="num">Reorder</th><th class="num">Wait</th><th class="num">Overstock</th><th class="num">Review</th><th class="num">Dead Stock</th><th class="num">Suggested Purchase</th></tr></thead><tbody>${rows.map(b=>`<tr class="brand-row" data-brand="${esc(b.brand)}"><td><button class="mini-link">${esc(b.brand)}</button></td><td class="num">${fmt(b.products)}</td><td class="num">${fmt(b.stock)}</td><td class="num">${fmt(b.fast)}</td><td class="num">${fmt(b.critical)}</td><td class="num">${fmt(b.reorder)}</td><td class="num">${fmt(b.wait)}</td><td class="num">${fmt(b.overstock)}</td><td class="num">${fmt(b.review)}</td><td class="num">${fmt(b.dead)}</td><td class="num qty">${fmt(b.suggested)}</td></tr>`).join('')}</tbody>`;
 $('brandTable').querySelectorAll('.brand-row').forEach(r=>r.onclick=()=>selectBrand(r.dataset.brand));
}
function selectBrand(brand){state.selectedBrands.clear();state.selectedBrands.add(brand);renderMultiMenu('brand',uniqueSorted('brand'));applyFilters();setView('dashboard');toast(`Filtered to ${brand}`);}

function renderQuality(){
 if(!state.products.length){$('qualityGrid').innerHTML='<div class="empty">Upload data first.</div>';$('columnChecks').innerHTML='';return;}
 const q=E.dataQuality(state.products,state.map||{}),total=state.products.length;
 $('qualityGrid').innerHTML=`<div class="quality-card"><strong>${fmt(total)}</strong><span>Total product rows</span></div><div class="quality-card"><strong>${fmt(q.missing.oem)}</strong><span>Missing OEM (${((q.missing.oem/Math.max(1,total))*100).toFixed(1)}%)</span></div><div class="quality-card"><strong>${fmt(q.missing.brand)}</strong><span>Missing brand</span></div><div class="quality-card"><strong>${fmt(q.missing.lastSaleDate)}</strong><span>Sales rows missing Last Sale Date</span></div>`;
 const checks=[['Odoo Make field (otherwise derived from Internal Ref)',q.columns.make],['Bavaria On Way',q.columns.bavariaOnWay],['Tibao SHJ On Way',q.columns.tibaoOnWay],['Report Till Date column',q.columns.reportDate],['Group On Way fallback',q.columns.groupOnWay],['Group On Way 2',q.columns.groupOnWay2],['Total Purchase Qty',q.columns.totalPurchase],['Purchase Count',q.columns.purchaseCount],['Last Purchase Date',q.columns.lastPurchaseDate]];
 $('columnChecks').innerHTML=checks.map(([l,ok])=>`<div class="check-item"><span>${esc(l)}</span><span class="${ok?'yes':'no'}">${ok?'✓ Available':'○ Missing / fallback'}</span></div>`).join('');
}

function bindDetailLinks(root){root.querySelectorAll('.open-detail').forEach(b=>b.onclick=()=>openDetail(b.dataset.id));}
function openDetail(id){
 const p=state.computed.find(x=>x._id===id);if(!p)return;
 $('modalTitle').textContent=p.internalRef||p.brandPartNo||'Product analysis';
 $('modalSub').textContent=`${p.make||'UNKNOWN'} • ${p.brand} • ${p.oem||'No OEM'} • ${p.description||''}`;
 const monthCount=p.monthCount,max=Math.max(1,...p.monthly.slice(0,monthCount));
 const monthBars=p.monthly.slice(0,monthCount).map((v,i)=>`<div class="month-cell"><div class="month-value">${fmt(v)}</div><div class="month-bar"><span style="height:${Math.max(2,(v/max)*100)}%"></span></div><div class="month-label">${months[i]}</div></div>`).join('');
 const altRows=p.alternatives.length?p.alternatives.map(a=>`<tr><td>${esc(a.brand)}</td><td>${esc(a.partNumbers.join(', ')||'—')}</td><td class="num">${fmt(a.stock)}</td><td class="num">${fmt(a.onWay)}</td><td class="num">${fmt(a.onWay2)}</td><td class="num">${fmt(a.sales)}</td></tr>`).join(''):'<tr><td colspan="6" class="empty">No exact same-OEM match in another brand.</td></tr>';
 const lastSale=p.lastSaleDate?`${dateFmt(p.lastSaleDate)} (${fmt(p.lastSaleAgeMonths,1)} mo ago)`:'—';
 const lastPurchase=p.lastPurchaseDate?`${dateFmt(p.lastPurchaseDate)} (${fmt(p.lastPurchaseAgeMonths,1)} mo ago)`:'—';
 $('modalBody').innerHTML=`
  <div class="detail-grid">
   <div class="metric"><div class="m-label">All Company Qty</div><div class="m-value">${fmt(p.allCompany)}</div></div>
   <div class="metric"><div class="m-label">On Way</div><div class="m-value">${fmt(p.onWay)}</div></div>
   <div class="metric"><div class="m-label">On Way 2</div><div class="m-value">${fmt(p.onWay2)}</div></div>
   <div class="metric"><div class="m-label">Suggested Purchase</div><div class="m-value purple-text">${fmt(p.suggestedQty)}</div></div>
   <div class="metric"><div class="m-label">Sales Qty</div><div class="m-value">${fmt(p.totalSales)}</div></div>
   <div class="metric"><div class="m-label">Active Sales Months</div><div class="m-value">${p.activeMonths}/${p.monthCount}</div></div>
   <div class="metric"><div class="m-label">Smart Demand / Mo</div><div class="m-value">${fmt(p.demandRate,1)}</div></div>
   <div class="metric"><div class="m-label">Demand Confidence</div><div class="m-value"><span class="badge ${confidenceClass(p.demandConfidence)}">${esc(p.demandConfidence)}</span></div></div>
  </div>
  <div class="section-block"><div class="section-block-title">Monthly Sales Pattern</div><div class="month-chart">${monthBars}</div><div class="micro-grid"><span>Overall Avg <strong>${fmt(p.avgMonthlySales,1)}</strong></span><span>Recent 3M <strong>${fmt(p.recent3Avg,1)}</strong></span><span>Recent 6M <strong>${fmt(p.recent6Avg,1)}</strong></span><span>Pattern <strong>${esc(p.demandPattern)}</strong></span><span>Largest month share <strong>${fmt(p.spikeSharePct,0)}%</strong></span><span>Last sale <strong>${esc(lastSale)}</strong></span><span>Zero-sales streak <strong>${fmt(p.zeroSalesStreakMonths,0)} months</strong></span><span>Inactivity basis <strong>${esc(p.inactivityBasis||'—')}</strong></span></div></div>
  <div class="section-block"><div class="section-block-title">Purchase History Signal</div><div class="micro-grid"><span>Total Purchase Qty <strong>${fmt(p.totalPurchase)}</strong></span><span>Purchase Count <strong>${fmt(p.purchaseCount)}</strong></span><span>Last Purchase <strong>${esc(lastPurchase)}</strong></span><span>Purchase / Sales <strong>${ratio(p.purchaseSalesRatio)}</strong></span><span>Signal <strong>${esc(p.purchaseSignal)}</strong></span></div><div class="note-box">Historical purchases are <strong>not subtracted again</strong> from the suggestion because receipts are already reflected in current stock. They are used to detect overbuying, recent replenishment and unusual purchase/sales patterns.</div></div>
  <div class="section-block"><div class="section-block-title">Suggested Purchase Calculation</div><div class="formula-box"><strong>${esc(p.condition)} — ${esc(p.action)}</strong><br>${esc(p.reason)}<br><br>Current cover: <strong>${cover(p.currentCover)} mo</strong> • With On Way + On Way 2: <strong>${cover(p.pipelineCover)} mo</strong> • Target gap before reorder rule: <strong>${fmt(p.targetGapQty)} pcs</strong><br>Final suggested quantity: <strong>${fmt(p.suggestedQty)} pcs</strong></div></div>
  <div class="section-block"><div class="section-block-title">Exact Same OEM — Other Brands</div><table class="alt-table"><thead><tr><th>Brand</th><th>Brand Part No.</th><th class="num">Stock</th><th class="num">On Way</th><th class="num">On Way 2</th><th class="num">Sales</th></tr></thead><tbody>${altRows}</tbody></table></div>`;
 $('modalBackdrop').classList.remove('hidden');
}

function exportRows(rows,name){
 if(!rows.length){toast('Nothing to export');return;}
 const data=rows.map(p=>{
   const out={
    'Internal Reference':p.internalRef,'Add Description':p.description,'Make':p.make,'OEM':p.oem,'Brand':p.brand,'Brand Part No.':p.brandPartNo,'Category':p.category,
    'All Company Qty':p.allCompany,'On Way':p.onWay,'On Way 2':p.onWay2,
    'Sales Qty':p.totalSales,'Sales Count':p.salesCount,'Active Sales Months':`${p.activeMonths}/${p.monthCount}`,'Overall Avg / Mo':+p.avgMonthlySales.toFixed(2),'Recent 3M Avg':+p.recent3Avg.toFixed(2),'Recent 6M Avg':+p.recent6Avg.toFixed(2),'Smart Demand / Mo':+p.demandRate.toFixed(2),'Demand Pattern':p.demandPattern,'Demand Confidence':p.demandConfidence,'Last Sale Date':dateFmt(p.lastSaleDate),'Last Sale Age Months':p.lastSaleAgeMonths===null?'':+p.lastSaleAgeMonths.toFixed(2),'Zero Sales Streak Months':p.zeroSalesStreakMonths,'Inactivity Age Months':p.inactivityAgeMonths===null?'':+p.inactivityAgeMonths.toFixed(2),'Inactivity Basis':p.inactivityBasis,
    'Historical Purchase Qty':p.totalPurchase,'Purchase Count':p.purchaseCount,'Last Purchase Date':dateFmt(p.lastPurchaseDate),'Purchase/Sales Ratio':p.purchaseSalesRatio===Infinity?'INF':+p.purchaseSalesRatio.toFixed(2),'Purchase Signal':p.purchaseSignal,
    'Same OEM Other Brand Stock':p.otherStock,'Same OEM Other Brand On Way':p.otherOnWay,'Same OEM Other Brand On Way 2':p.otherOnWay2,'Same OEM Detail':p.equivalentNote,
    'Current Months Cover':p.currentCover===Infinity?'INF':+p.currentCover.toFixed(2),'Pipeline Months Cover':p.pipelineCover===Infinity?'INF':+p.pipelineCover.toFixed(2),'Stock Condition':p.condition,'Priority':p.priority,'Target Gap Qty':p.targetGapQty,'Suggested Purchase Qty':p.suggestedQty,'Purchase Action':p.action,'Reason':p.reason
   };
   months.forEach((m,i)=>out[`${m} Sales`]=p.monthly[i]||0);return out;
  });
 const ws=XLSX.utils.json_to_sheet(data),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Purchase Suggestions');XLSX.writeFile(wb,name);toast(`Exported ${fmt(rows.length)} rows`);
}

function syncSettingsUI(){
 const s=state.settings;$('sTarget').value=s.targetCover;$('sSafety').value=s.safetyCover;$('sCritical').value=s.criticalCover;$('sReorder').value=s.reorderCover;$('sOverstock').value=s.overstockCover;$('sDormant').value=s.dormantMonths;$('sDead').value=s.deadStockMonths;$('sFast').value=s.fastRate;$('sMedium').value=s.mediumRate;$('sSpike').value=s.spikeDominancePct;$('sRegular').value=s.regularActivePct;$('sEquivalent').value=s.equivalentCreditPct;$('sDemand').value=s.demandMethod;$('sSuggestionMode').value=s.suggestionMode;
}
function applySettingsFromUI(){
 state.settings.targetCover=+$('sTarget').value;state.settings.safetyCover=+$('sSafety').value;state.settings.criticalCover=+$('sCritical').value;state.settings.reorderCover=+$('sReorder').value;state.settings.overstockCover=+$('sOverstock').value;state.settings.dormantMonths=+$('sDormant').value;state.settings.deadStockMonths=+$('sDead').value;state.settings.fastRate=+$('sFast').value;state.settings.mediumRate=+$('sMedium').value;state.settings.spikeDominancePct=+$('sSpike').value;state.settings.regularActivePct=+$('sRegular').value;state.settings.equivalentCreditPct=+$('sEquivalent').value;state.settings.demandMethod=$('sDemand').value;state.settings.suggestionMode=$('sSuggestionMode').value;saveSettings();recompute();toast('Calculation settings applied');
}

// Events
['uploadBtn','uploadTop'].forEach(id=>$(id).onclick=()=>$('fileInput').click());
$('fileInput').onchange=e=>handleFile(e.target.files[0]);$('demoBtn').onclick=loadDemo;
const zone=$('uploadZone');['dragenter','dragover'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.add('drag');}));['dragleave','drop'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.remove('drag');}));zone.addEventListener('drop',e=>handleFile(e.dataTransfer.files[0]));
document.querySelectorAll('.nav button').forEach(b=>b.onclick=()=>setView(b.dataset.view));document.querySelectorAll('[data-go]').forEach(b=>b.onclick=()=>setView(b.dataset.go));
['search','descriptionFilter'].forEach(id=>$(id).addEventListener('input',applyFilters));
['makeFilter','movementFilter','patternFilter','categoryFilter'].forEach(id=>$(id).addEventListener('change',applyFilters));
function toggleMultiMenu(type){
 const menu=$(type==='brand'?'brandFilterMenu':'conditionFilterMenu');
 const other=$(type==='brand'?'conditionFilterMenu':'brandFilterMenu');
 other.classList.add('hidden');menu.classList.toggle('hidden');
}
$('brandFilterTrigger').onclick=e=>{e.stopPropagation();toggleMultiMenu('brand');};
$('conditionFilterTrigger').onclick=e=>{e.stopPropagation();toggleMultiMenu('condition');};
['brandFilterMenu','conditionFilterMenu'].forEach(id=>$(id).onclick=e=>e.stopPropagation());
document.addEventListener('click',()=>{$('brandFilterMenu').classList.add('hidden');$('conditionFilterMenu').classList.add('hidden');});
$('clearFilters').onclick=()=>{state.selectedBrands.clear();state.selectedConditions.clear();['search','descriptionFilter','makeFilter','movementFilter','patternFilter','categoryFilter'].forEach(id=>$(id).value='');populateFilters();applyFilters();};
document.querySelectorAll('.kpi[data-condition]').forEach(k=>k.onclick=()=>{state.selectedConditions.clear();state.selectedConditions.add(k.dataset.condition);renderMultiMenu('condition',E.CONDITION_OPTIONS);applyFilters();});
document.querySelectorAll('.kpi[data-movement]').forEach(k=>k.onclick=()=>{$('movementFilter').value=k.dataset.movement;applyFilters();});
$('pageSize').onchange=()=>{state.pageSize=+$('pageSize').value;state.page=1;renderPlanner();};$('prevPage').onclick=()=>{if(state.page>1){state.page--;renderPlanner();}};$('nextPage').onclick=()=>{state.page++;renderPlanner();};
$('exportTop').onclick=$('exportFiltered').onclick=()=>exportRows(state.filtered,`Purchase_Suggestions_v2_2_${dateFmt(state.reportDate||new Date())}.xlsx`);$('exportSelected').onclick=()=>exportRows(state.computed.filter(p=>state.selected.has(p._id)),`Selected_Purchase_Suggestions_v2_2_${dateFmt(state.reportDate||new Date())}.xlsx`);
$('modalClose').onclick=()=>$('modalBackdrop').classList.add('hidden');$('modalBackdrop').onclick=e=>{if(e.target===$('modalBackdrop'))$('modalBackdrop').classList.add('hidden');};
$('applySettings').onclick=applySettingsFromUI;$('resetSettings').onclick=()=>{state.settings={...E.DEFAULT_SETTINGS};saveSettings();syncSettingsUI();recompute();toast('Defaults restored');};

populateFilters();renderDashboard();renderPlanner();
})();
