(function(){
'use strict';
const E=window.PurchaseEngine;
let state={products:[],computed:[],filtered:[],shjRows:[],shjFiltered:[],reportDate:null,fileName:'',map:null,settings:loadSettings(),page:1,pageSize:50,shjPage:1,shjPageSize:100,sortKey:'priority',sortDir:1,selected:new Set(),selectedBrands:new Set(),selectedConditions:new Set(),qtyOverrides:new Map(),shjSelected:new Set(),shjSelectedBrands:new Set(),shjSelectedMakes:new Set(),shjSelectedStatuses:new Set(),shjQtyOverrides:new Map(),activeView:'dashboard',brandGroupDrafts:[]};

const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
const fmt=(x,d=0)=>Number.isFinite(x)?x.toLocaleString(undefined,{maximumFractionDigits:d}):'—';
const dateFmt=d=>d instanceof Date&&!isNaN(d)?d.toISOString().slice(0,10):'—';
const months=['Jan','Feb','Mar','Apr','May','Jun','Jul','Aug','Sep','Oct','Nov','Dec'];

function normalizeSettings(raw){
 const out=Object.assign({},E.DEFAULT_SETTINGS,raw||{});
 out.brandRules=(raw&&raw.brandRules&&typeof raw.brandRules==='object')?raw.brandRules:{};
 out.brandGroups=Array.isArray(raw&&raw.brandGroups)?raw.brandGroups.map((g,i)=>({id:String(g.id||`group_${i+1}`),name:String(g.name||`Brand Group ${i+1}`),enabled:g.enabled!==false,brands:Array.isArray(g.brands)?[...g.brands]:[],stockCreditPct:Number.isFinite(+g.stockCreditPct)?+g.stockCreditPct:100,onWayCreditPct:Number.isFinite(+g.onWayCreditPct)?+g.onWayCreditPct:100,onWay2CreditPct:Number.isFinite(+g.onWay2CreditPct)?+g.onWay2CreditPct:100})):[];
 return out;
}
function loadSettings(){try{return normalizeSettings(JSON.parse(localStorage.getItem('purchaseSettingsV2')||'{}'));}catch(_){return normalizeSettings({});}}
function saveSettings(){localStorage.setItem('purchaseSettingsV2',JSON.stringify(state.settings));}
function toast(msg){const t=$('toast');t.textContent=msg;t.classList.remove('hidden');clearTimeout(toast.timer);toast.timer=setTimeout(()=>t.classList.add('hidden'),2600);}
function cover(x){return x===Infinity?'∞':fmt(x,1);}
function ratio(x){return x===Infinity?'∞':Number.isFinite(x)?x.toFixed(2):'—';}
function stockoutFmt(d,demand,supply){
 if(!(d instanceof Date)||isNaN(d))return demand>0?'—':'No demand forecast';
 if(demand>0&&supply<=0)return 'OUT NOW';
 return dateFmt(d);
}
function plannerQty(p){const v=state.qtyOverrides.get(p._id);return v===undefined?p.suggestedQty:v;}
function setPlannerQty(p,value){
 const parsed=Number(value);const qty=Number.isFinite(parsed)?Math.max(0,Math.round(parsed)):p.suggestedQty;
 if(qty===p.suggestedQty)state.qtyOverrides.delete(p._id);else state.qtyOverrides.set(p._id,qty);
 return qty;
}
function totalPlannerQty(rows){return rows.reduce((s,p)=>s+plannerQty(p),0);}

function setView(name){
 state.activeView=name;
 document.querySelectorAll('.view').forEach(v=>v.classList.toggle('active',v.id===`view-${name}`));
 document.querySelectorAll('.nav button').forEach(b=>b.classList.toggle('active',b.dataset.view===name));
 const titles={dashboard:'Purchase Support',planner:'Purchase Planner',shj:'SHJ Replenishment',brands:'Brand Analysis',quality:'Data Quality',settings:'Calculation Settings'};
 $('pageTitle').textContent=titles[name]||'Purchase Support';
 if(name==='planner')renderPlanner();if(name==='shj')renderShj();if(name==='brands')renderBrands();if(name==='quality')renderQuality();if(name==='settings')syncSettingsUI();
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
  state.products=parsed.products;state.reportDate=parsed.reportDate;state.map=parsed.map;state.fileName=file.name;state.selected.clear();state.qtyOverrides.clear();state.shjSelected.clear();state.shjSelectedBrands.clear();state.shjSelectedMakes.clear();state.shjSelectedStatuses.clear();state.shjQtyOverrides.clear();state.page=1;state.shjPage=1;
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
function loadDemo(){state.products=demoProducts();state.reportDate=new Date(2026,7,23);state.map={bavariaOnWay:1,tibaoOnWay:1,reportDate:1,groupOnWay:1,groupOnWay2:1,totalPurchase:1,purchaseCount:1,lastPurchaseDate:1};state.fileName='Demo data';state.selected.clear();state.qtyOverrides.clear();state.shjSelected.clear();state.shjSelectedBrands.clear();state.shjSelectedMakes.clear();state.shjSelectedStatuses.clear();state.shjQtyOverrides.clear();state.shjPage=1;recompute();$('fileChip').textContent='Demo data';$('reportMeta').textContent=`${state.products.length} demo products • Data as of ${dateFmt(state.reportDate)} • Smart logic v3.3`;showUploadZone(false);$('exportTop').disabled=false;toast('Demo data loaded');}

function recompute(){state.computed=E.calculate(state.products,state.settings,state.reportDate||new Date());state.shjRows=calculateShjRows();populateFilters();populateShjFilters();applyFilters();applyShjFilters();renderDashboard();renderPlanner();renderShj();renderBrands();renderQuality();}
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
 // Keep selections strictly inside the current filtered result.
 const visibleIds=new Set(state.filtered.map(p=>p._id));
 [...state.selected].forEach(id=>{if(!visibleIds.has(id))state.selected.delete(id);});
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
 $('kpiSuggested').textContent=fmt(totalPlannerQty(state.filtered));
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
 el.innerHTML=`<thead><tr><th>Internal Ref</th><th>Add Description</th><th>Make</th><th>OEM</th><th>Brand</th><th class="num">All Company</th><th class="num">On Way</th><th class="num">On Way 2</th><th>Current Stock-Out</th><th>With Incoming</th><th>Demand Pattern</th><th class="num">Demand / Mo</th><th class="num">Active Months</th><th class="num">Purchase Qty</th><th>Same OEM Other Brands</th><th class="num">Cover</th><th>Condition</th><th class="num">Suggested Qty</th><th>Action</th></tr></thead><tbody>${rows.map(p=>`<tr><td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(p.internalRef||'—')}</button></td><td class="desc" title="${esc(p.description)}">${esc(p.description||'—')}</td><td>${esc(p.make||'UNKNOWN')}</td><td>${esc(p.oem||'—')}</td><td>${esc(p.brand)}</td><td class="num">${fmt(p.allCompany)}</td><td class="num">${fmt(p.onWay)}</td><td class="num">${fmt(p.onWay2)}</td><td>${esc(stockoutFmt(p.currentStockoutDate,p.demandRate,p.allCompany))}</td><td>${esc(stockoutFmt(p.pipelineStockoutDate,p.demandRate,p.directSupply))}</td><td><span class="badge ${patternClass(p.demandPattern)}">${esc(p.demandPattern)}</span></td><td class="num">${fmt(p.demandRate,1)}</td><td class="num">${p.activeMonths}/${p.monthCount}</td><td class="num">${fmt(p.totalPurchase)}</td><td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(p.equivalentNote)}</button></td><td class="num">${cover(p.currentCover)}</td><td><span class="badge ${conditionClass(p.condition)}">${esc(p.condition)}</span></td><td class="num qty">${fmt(plannerQty(p))}</td><td>${esc(p.action)}</td></tr>`).join('')||'<tr><td colspan="19" class="empty">No items match the current filters.</td></tr>'}</tbody>`;
 bindDetailLinks(el);
}

const plannerColumns=[
 ['internalRef','Internal Ref'],['description','Add Description'],['make','Make'],['oem','OEM'],['brand','Brand'],['brandPartNo','Brand Part No.'],['equivalentNote','Same OEM Other Brands'],
 ['suggestedQty','Suggested Qty (Editable)'],['action','Purchase Action'],
 ['allCompany','All Company Qty'],['onWay','On Way'],['onWay2','On Way 2'],['currentStockoutDate','Current Stock-Out'],['pipelineStockoutDate','With Incoming Stock-Out'],['leadTimeDays','Lead Time Days'],['totalSales','Sales Qty'],['salesCount','Sales Count'],['activeMonths','Sales Months'],
 ['avgMonthlySales','Overall Avg'],['recent3Avg','Recent 3M'],['recent6Avg','Recent 6M'],['demandRate','Smart Demand'],['demandPattern','Demand Pattern'],['demandConfidence','Confidence'],
 ['totalPurchase','Purchase Qty'],['purchaseCount','Purchase Count'],['purchaseSalesRatio','Purch/Sales'],['purchaseSignal','Purchase Signal'],
 ['currentCover','Current Cover'],['pipelineCover','Pipeline Cover'],['condition','Stock Condition'],['priority','Priority']
];
function plannerCell(p,k){
 let v=p[k];
 if(k==='internalRef')return`<td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(v||'—')}</button></td>`;
 if(k==='description')return`<td class="desc" title="${esc(v)}">${esc(v||'—')}</td>`;
 if(k==='suggestedQty'){const q=plannerQty(p),modified=state.qtyOverrides.has(p._id);return`<td class="num qty-cell"><input class="qty-edit ${modified?'modified':''}" type="number" min="0" step="1" value="${q}" data-id="${esc(p._id)}" data-calculated="${p.suggestedQty}" title="Calculated suggestion: ${p.suggestedQty}. Edit this quantity before exporting."></td>`;}
 if(['allCompany','onWay','onWay2','leadTimeDays','totalSales','salesCount','totalPurchase','purchaseCount','leadTimeDays'].includes(k))return`<td class="num">${fmt(v)}</td>`;
 if(k==='currentStockoutDate')return`<td>${esc(stockoutFmt(v,p.demandRate,p.allCompany))}</td>`;
 if(k==='pipelineStockoutDate')return`<td>${esc(stockoutFmt(v,p.demandRate,p.directSupply))}</td>`;
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
 $('plannerResults').textContent=`${fmt(sorted.length)} items • ${fmt(totalPlannerQty(sorted))} pcs planned`;
 $('pageInfo').textContent=`Page ${state.page} of ${pages} • showing ${rows.length}`;
 $('prevPage').disabled=state.page<=1;$('nextPage').disabled=state.page>=pages;
 $('plannerTable').innerHTML=`<thead><tr><th><input type="checkbox" id="selectPage" title="Select all currently filtered items"></th><th class="num line-no">#</th>${plannerColumns.map(([k,l])=>`<th data-sort="${k}" class="${['allCompany','onWay','onWay2','leadTimeDays','totalSales','salesCount','activeMonths','avgMonthlySales','recent3Avg','recent6Avg','demandRate','totalPurchase','purchaseCount','purchaseSalesRatio','currentCover','pipelineCover','priority','suggestedQty'].includes(k)?'num':''}">${esc(l)}${state.sortKey===k?(state.sortDir>0?' ↑':' ↓'):''}</th>`).join('')}</tr></thead><tbody>${rows.map((p,i)=>`<tr><td><input class="row-select" type="checkbox" data-id="${esc(p._id)}" ${state.selected.has(p._id)?'checked':''}></td><td class="num line-no">${start+i+1}</td>${plannerColumns.map(([k])=>plannerCell(p,k)).join('')}</tr>`).join('')||'<tr><td colspan="32" class="empty">No items match filters.</td></tr>'}</tbody>`;
 $('plannerTable').querySelectorAll('th[data-sort]').forEach(th=>th.onclick=()=>{const k=th.dataset.sort;if(state.sortKey===k)state.sortDir*=-1;else{state.sortKey=k;state.sortDir=1;}renderPlanner();});
 $('plannerTable').querySelectorAll('.row-select').forEach(cb=>cb.onchange=()=>{cb.checked?state.selected.add(cb.dataset.id):state.selected.delete(cb.dataset.id);updateSelectedButton();});
 $('plannerTable').querySelectorAll('.qty-edit').forEach(inp=>{inp.onfocus=()=>inp.select();inp.onchange=()=>{const p=state.computed.find(x=>x._id===inp.dataset.id);if(!p)return;const q=setPlannerQty(p,inp.value);inp.value=q;inp.classList.toggle('modified',state.qtyOverrides.has(p._id));$('plannerResults').textContent=`${fmt(state.filtered.length)} items • ${fmt(totalPlannerQty(state.filtered))} pcs planned`;$('kpiSuggested').textContent=fmt(totalPlannerQty(state.filtered));toast(q===p.suggestedQty?'Quantity reset to calculated suggestion':`Purchase quantity changed to ${fmt(q)}`);};});
 const sp=$('selectPage');if(sp){
 const selectedInFilter=state.filtered.reduce((n,p)=>n+(state.selected.has(p._id)?1:0),0);
 sp.title='Select all currently filtered items';
 sp.checked=state.filtered.length>0&&selectedInFilter===state.filtered.length;
 sp.indeterminate=selectedInFilter>0&&selectedInFilter<state.filtered.length;
 sp.onchange=()=>{
  state.selected.clear();
  if(sp.checked)state.filtered.forEach(p=>state.selected.add(p._id));
  renderPlanner();updateSelectedButton();
 };
}
 bindDetailLinks($('plannerTable'));updateSelectedButton();
}
function sortCompare(a,b,k,d){let av=k==='suggestedQty'?plannerQty(a):a[k],bv=k==='suggestedQty'?plannerQty(b):b[k];if(av instanceof Date||bv instanceof Date){av=av instanceof Date?av.getTime():-Infinity;bv=bv instanceof Date?bv.getTime():-Infinity;return(av-bv)*d;}if(typeof av==='string'||typeof bv==='string')return String(av??'').localeCompare(String(bv??''))*d;av=Number.isFinite(av)?av:-Infinity;bv=Number.isFinite(bv)?bv:-Infinity;return(av-bv)*d;}
function updateSelectedButton(){
 const n=state.selected.size;
 $('exportSelected').disabled=n===0;
 $('exportSelected').textContent=`Export Selected (${n})`;
 if($('clearSelection')){
  $('clearSelection').disabled=false;
  $('clearSelection').textContent=n?`Clear Selection (${n})`:'Clear Selection';
 }
}

function shjTransferQty(p){const v=state.shjQtyOverrides.get(p._id);return v===undefined?p.fromMotorline:v;}
function setShjTransferQty(p,value){const parsed=Number(value);const qty=Number.isFinite(parsed)?Math.max(0,Math.round(parsed)):p.fromMotorline;if(qty===p.fromMotorline)state.shjQtyOverrides.delete(p._id);else state.shjQtyOverrides.set(p._id,qty);return qty;}
function shjShortageAfterEdit(p){return Math.max(0,p.shjGap-p.fromBavaria-shjTransferQty(p));}
function shjActionAfterEdit(p){const q=shjTransferQty(p),shortage=shjShortageAfterEdit(p);if(q>0&&shortage>0)return `SEND ${q} FROM MOTORLINE • SHORT ${shortage}`;if(q>0)return `SEND ${q} FROM MOTORLINE`;if(p.fromBavaria>0&&shortage<=0)return `USE ${p.fromBavaria} FROM BAVARIA`;if(shortage>0)return 'PURCHASE / REVIEW';return 'NO TRANSFER';}
function updateShjSelectionButtons(){const n=state.shjSelected.size;const ex=$('shjExportSelected'),cl=$('shjClearSelection');if(ex){ex.disabled=n===0;ex.textContent=`Export Selected (${n})`;}if(cl)cl.textContent=n?`Clear Selection (${n})`:'Clear Selection';}

function calculateShjRows(){
 const s=state.settings,share=Math.max(0,Math.min(100,+s.shjSharePct||0))/100;
 const weeksPerMonth=4.345;
 return state.computed.map(p=>{
  const demand=Math.max(0,p.demandRate||0);
  const estShjDemand=demand*share;
  const estMotorlineDemand=demand*Math.max(0,1-share);
  const targetWeeks=p.movement==='FAST MOVING'?Math.max(0,+s.shjFastCoverWeeks||0):Math.max(0,+s.shjNormalCoverWeeks||0);
  const targetQty=estShjDemand>0?Math.ceil(estShjDemand*(targetWeeks/weeksPerMonth)):0;
  const shjGap=Math.max(0,targetQty-Math.max(0,p.tibaoStock||0));
  const bavariaAvailable=s.shjUseBavariaBackup?Math.max(0,p.bavariaStock||0):0;
  const fromBavaria=Math.min(shjGap,bavariaAvailable);
  const afterBavaria=Math.max(0,shjGap-fromBavaria);
  const mlReserveQty=Math.ceil(estMotorlineDemand*Math.max(0,+s.shjMotorlineReserveMonths||0));
  const mlExcess=Math.max(0,Math.max(0,p.motorlineStock||0)-mlReserveQty);
  const fromMotorline=Math.min(afterBavaria,mlExcess);
  const remainingShortage=Math.max(0,afterBavaria-fromMotorline);
  const reviewPattern=['DEAD STOCK','DORMANT','ONE-TIME SPIKE','NO SALES'].includes(p.demandPattern)||demand<Math.max(0,+s.shjMinGroupDemand||0);
  let shjStatus='SHJ STOCK OK',shjAction='NO TRANSFER';
  if(reviewPattern){shjStatus='REVIEW DEMAND';shjAction='NO AUTOMATIC TRANSFER';}
  else if(shjGap<=0){shjStatus='SHJ STOCK OK';shjAction='NO TRANSFER';}
  else if(fromBavaria>=shjGap){shjStatus='USE BAVARIA STOCK';shjAction=`MOVE ${fromBavaria} FROM BAVARIA`;}
  else if(fromMotorline>0&&remainingShortage>0){shjStatus='PARTIAL TRANSFER / REVIEW';shjAction=`ML ${fromMotorline}${fromBavaria?` + BAV ${fromBavaria}`:''}; SHORT ${remainingShortage}`;}
  else if(fromMotorline>0&&p.tibaoStock<=0&&p.movement==='FAST MOVING'){shjStatus='URGENT TRANSFER';shjAction=`SEND ${fromMotorline} FROM MOTORLINE`;}
  else if(fromMotorline>0){shjStatus='TRANSFER TO SHJ';shjAction=`SEND ${fromMotorline} FROM MOTORLINE`;}
  else if(remainingShortage>0){shjStatus='PURCHASE / REVIEW';shjAction=fromBavaria?`USE BAVARIA ${fromBavaria}; SHORT ${remainingShortage}`:`NO SAFE MOTORLINE EXCESS`;}
  const actionable=['URGENT TRANSFER','TRANSFER TO SHJ','USE BAVARIA STOCK','PARTIAL TRANSFER / REVIEW','PURCHASE / REVIEW'].includes(shjStatus);
  return Object.assign({},p,{estShjDemand,estMotorlineDemand,targetWeeks,targetShjQty:targetQty,shjGap,fromBavaria,mlReserveQty,mlExcess,fromMotorline,remainingShortage,shjStatus,shjAction,shjActionable:actionable});
 });
}
function shjFilterConfig(type){
 const configs={
  brand:{set:state.shjSelectedBrands,trigger:'shjBrandFilterTrigger',menu:'shjBrandFilterMenu',all:'All Brands',search:'Search brands...'},
  make:{set:state.shjSelectedMakes,trigger:'shjMakeFilterTrigger',menu:'shjMakeFilterMenu',all:'All Makes',search:'Search makes...'},
  status:{set:state.shjSelectedStatuses,trigger:'shjStatusFilterTrigger',menu:'shjStatusFilterMenu',all:'All Statuses',search:null}
 };
 return configs[type];
}
function updateShjMultiTrigger(type){
 const c=shjFilterConfig(type),set=c.set,trigger=$(c.trigger);
 if(!set.size)trigger.textContent=c.all;
 else if(set.size===1)trigger.textContent=[...set][0];
 else trigger.textContent=`${set.size} selected`;
 trigger.classList.toggle('has-selection',set.size>0);
}
function renderShjMultiMenu(type,values){
 const c=shjFilterConfig(type),set=c.set,menu=$(c.menu);
 menu.innerHTML=`
  <div class="multi-head">
   ${c.search?`<input class="multi-search" type="text" placeholder="${c.search}" autocomplete="off">`:''}
   <div class="multi-actions"><button type="button" data-action="all">Select all</button><button type="button" data-action="clear">Clear</button></div>
  </div>
  <div class="multi-options">${values.map(v=>`<label class="multi-option" data-label="${esc(String(v).toLowerCase())}"><input type="checkbox" value="${esc(v)}" ${set.has(v)?'checked':''}><span>${esc(v)}</span></label>`).join('')}</div>`;
 menu.querySelectorAll('.multi-option input').forEach(cb=>cb.onchange=()=>{cb.checked?set.add(cb.value):set.delete(cb.value);updateShjMultiTrigger(type);applyShjFilters();});
 menu.querySelector('[data-action="all"]').onclick=()=>{values.forEach(v=>set.add(v));renderShjMultiMenu(type,values);updateShjMultiTrigger(type);applyShjFilters();};
 menu.querySelector('[data-action="clear"]').onclick=()=>{set.clear();renderShjMultiMenu(type,values);updateShjMultiTrigger(type);applyShjFilters();};
 const search=menu.querySelector('.multi-search');
 if(search)search.oninput=()=>{const q=search.value.trim().toLowerCase();menu.querySelectorAll('.multi-option').forEach(x=>x.classList.toggle('hidden',q&&!x.dataset.label.includes(q)));};
 updateShjMultiTrigger(type);
}
function populateShjFilters(){
 const brands=[...new Set(state.shjRows.map(x=>x.brand).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b)));
 const makes=[...new Set(state.shjRows.map(x=>x.make).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b)));
 const statuses=['URGENT TRANSFER','TRANSFER TO SHJ','USE BAVARIA STOCK','PARTIAL TRANSFER / REVIEW','PURCHASE / REVIEW','SHJ STOCK OK','REVIEW DEMAND'];
 pruneSelection(state.shjSelectedBrands,brands);pruneSelection(state.shjSelectedMakes,makes);pruneSelection(state.shjSelectedStatuses,statuses);
 renderShjMultiMenu('brand',brands);renderShjMultiMenu('make',makes);renderShjMultiMenu('status',statuses);
}
function applyShjFilters(){
 if(!$('shjSearch'))return;
 const q=$('shjSearch').value.trim().toLowerCase(),actionOnly=$('shjActionableOnly').checked;
 state.shjFiltered=state.shjRows.filter(p=>{
  if(state.shjSelectedBrands.size&&!state.shjSelectedBrands.has(p.brand))return false;
  if(state.shjSelectedMakes.size&&!state.shjSelectedMakes.has(p.make))return false;
  if(state.shjSelectedStatuses.size&&!state.shjSelectedStatuses.has(p.shjStatus))return false;
  if(actionOnly&&!p.shjActionable)return false;
  if(q){const hay=[p.internalRef,p.oem,p.description,p.brand,p.brandPartNo,p.make].join(' ').toLowerCase();if(!hay.includes(q))return false;}
  return true;
 });
 const visibleIds=new Set(state.shjFiltered.map(p=>p._id));[...state.shjSelected].forEach(id=>{if(!visibleIds.has(id))state.shjSelected.delete(id);});
 state.shjPage=1;renderShj();updateShjSelectionButtons();
}
function shjStatusClass(s){if(s==='URGENT TRANSFER')return'critical';if(s==='TRANSFER TO SHJ')return'wait';if(s==='USE BAVARIA STOCK')return'ok';if(s==='PURCHASE / REVIEW')return'reorder';if(s==='PARTIAL TRANSFER / REVIEW')return'spike';if(s==='SHJ STOCK OK')return'ok';return'review';}
function renderShj(){
 if(!$('shjTable'))return;
 const rows=state.shjFiltered||[];
 if(!state.computed.length){$('shjTable').innerHTML='<tbody><tr><td class="empty">Upload data first.</td></tr></tbody>';['shjKpiUrgent','shjKpiTransfer','shjKpiBavaria','shjKpiPurchase','shjKpiQty'].forEach(id=>$(id).textContent='0');$('shjResults').textContent='0 items';if($('shjPageInfo'))$('shjPageInfo').textContent='Page 1';updateShjSelectionButtons();return;}
 const count=s=>rows.filter(p=>p.shjStatus===s).length;
 $('shjKpiUrgent').textContent=fmt(count('URGENT TRANSFER'));$('shjKpiTransfer').textContent=fmt(rows.filter(p=>shjTransferQty(p)>0).length);$('shjKpiBavaria').textContent=fmt(rows.filter(p=>p.fromBavaria>0).length);$('shjKpiPurchase').textContent=fmt(rows.filter(p=>['PURCHASE / REVIEW','PARTIAL TRANSFER / REVIEW'].includes(p.shjStatus)).length);$('shjKpiQty').textContent=fmt(rows.reduce((s,p)=>s+shjTransferQty(p),0));
 $('shjResults').textContent=`${fmt(rows.length)} items • ${fmt(rows.reduce((s,p)=>s+shjTransferQty(p),0))} pcs planned from Motorline • ${fmt(rows.reduce((s,p)=>s+p.fromBavaria,0))} pcs from Bavaria`;
 const sorted=[...rows].sort((a,b)=>{const rank={'URGENT TRANSFER':1,'TRANSFER TO SHJ':2,'USE BAVARIA STOCK':3,'PARTIAL TRANSFER / REVIEW':4,'PURCHASE / REVIEW':5,'SHJ STOCK OK':6,'REVIEW DEMAND':7};return (rank[a.shjStatus]||9)-(rank[b.shjStatus]||9)||shjTransferQty(b)-shjTransferQty(a)||b.estShjDemand-a.estShjDemand;});
 const totalPages=Math.max(1,Math.ceil(sorted.length/state.shjPageSize));if(state.shjPage>totalPages)state.shjPage=totalPages;if(state.shjPage<1)state.shjPage=1;
 const start=(state.shjPage-1)*state.shjPageSize,pageRows=sorted.slice(start,start+state.shjPageSize);
 $('shjTable').innerHTML=`<thead><tr><th><input type="checkbox" id="shjSelectAll" title="Select all currently filtered items"></th><th class="num line-no">#</th><th>Old Number</th><th>Add Description</th><th>Make</th><th>OEM</th><th>Brand</th><th>Brand Part No.</th><th>Status</th><th>Action</th><th class="num">Suggested From Motorline (Editable)</th><th class="num">Motorline Stock</th><th class="num">Bavaria Nearby</th><th class="num">SHJ Stock</th><th class="num">Group Demand / Mo</th><th class="num">Est. SHJ Demand / Mo</th><th class="num">Target Weeks</th><th class="num">SHJ Target Qty</th><th class="num">Use Bavaria</th><th class="num">ML Safe Excess</th><th class="num">Shortage After Transfer</th></tr></thead><tbody>${pageRows.map((p,i)=>{const q=shjTransferQty(p),modified=state.shjQtyOverrides.has(p._id);return `<tr><td><input class="shj-row-select" type="checkbox" data-id="${esc(p._id)}" ${state.shjSelected.has(p._id)?'checked':''}></td><td class="num line-no">${start+i+1}</td><td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(p.internalRef||'—')}</button></td><td class="desc" title="${esc(p.description)}">${esc(p.description||'—')}</td><td>${esc(p.make||'UNKNOWN')}</td><td>${esc(p.oem||'—')}</td><td>${esc(p.brand)}</td><td>${esc(p.brandPartNo||'—')}</td><td><span class="badge ${shjStatusClass(p.shjStatus)}">${esc(p.shjStatus)}</span></td><td class="shj-action">${esc(shjActionAfterEdit(p))}</td><td class="num qty-cell"><input class="qty-edit shj-qty-edit ${modified?'modified':''}" type="number" min="0" step="1" value="${q}" data-id="${esc(p._id)}" data-calculated="${p.fromMotorline}" title="Calculated from Motorline: ${p.fromMotorline}. Edit before export."></td><td class="num">${fmt(p.motorlineStock)}</td><td class="num">${fmt(p.bavariaStock)}</td><td class="num">${fmt(p.tibaoStock)}</td><td class="num">${fmt(p.demandRate,1)}</td><td class="num">${fmt(p.estShjDemand,1)}</td><td class="num">${fmt(p.targetWeeks,1)}</td><td class="num">${fmt(p.targetShjQty)}</td><td class="num">${fmt(p.fromBavaria)}</td><td class="num">${fmt(p.mlExcess)}</td><td class="num">${fmt(shjShortageAfterEdit(p))}</td></tr>`}).join('')||'<tr><td colspan="21" class="empty">No items match the current SHJ replenishment filters.</td></tr>'}</tbody>`;
 $('shjTable').querySelectorAll('.shj-row-select').forEach(cb=>cb.onchange=()=>{cb.checked?state.shjSelected.add(cb.dataset.id):state.shjSelected.delete(cb.dataset.id);updateShjSelectionButtons();});
 $('shjTable').querySelectorAll('.shj-qty-edit').forEach(inp=>{inp.onfocus=()=>inp.select();inp.onchange=()=>{const p=state.shjRows.find(x=>x._id===inp.dataset.id);if(!p)return;const q=setShjTransferQty(p,inp.value);toast(q===p.fromMotorline?'Transfer quantity reset to calculated suggestion':`Motorline transfer changed to ${fmt(q)}`);renderShj();};});
 const all=$('shjSelectAll');if(all){const selectedInFilter=state.shjFiltered.reduce((n,p)=>n+(state.shjSelected.has(p._id)?1:0),0);all.checked=state.shjFiltered.length>0&&selectedInFilter===state.shjFiltered.length;all.indeterminate=selectedInFilter>0&&selectedInFilter<state.shjFiltered.length;all.onchange=()=>{state.shjSelected.clear();if(all.checked)state.shjFiltered.forEach(p=>state.shjSelected.add(p._id));renderShj();updateShjSelectionButtons();};}
 if($('shjPageInfo'))$('shjPageInfo').textContent=`Page ${state.shjPage} of ${totalPages} • showing ${fmt(pageRows.length)} of ${fmt(rows.length)}`;
 if($('shjPrevPage'))$('shjPrevPage').disabled=state.shjPage<=1;if($('shjNextPage'))$('shjNextPage').disabled=state.shjPage>=totalPages;
 bindDetailLinks($('shjTable'));updateShjSelectionButtons();
}
function exportShjRows(rows,name){
 if(!rows.length){toast('Nothing to export');return;}
 const data=rows.map((p,i)=>({'Line no':i+1,'Old Number':p.internalRef,'OEM No Space':p.oemKey||E.normalizeOEM(p.oem),'Add Description':p.description,'Brand No':p.brandPartNo||'','Brand':p.brand,'Suggested From Motorline':shjTransferQty(p),'Motorline Stock':p.motorlineStock,'Bavaria Nearby Stock':p.bavariaStock,'SHJ Stock':p.tibaoStock,'Status':p.shjStatus}));
 const ws=XLSX.utils.json_to_sheet(data),wb=XLSX.utils.book_new();ws['!cols']=[{wch:9},{wch:28},{wch:22},{wch:48},{wch:20},{wch:18},{wch:24},{wch:18},{wch:22},{wch:14},{wch:24}];XLSX.utils.book_append_sheet(wb,ws,'SHJ Replenishment');XLSX.writeFile(wb,name);toast(`Exported ${fmt(rows.length)} SHJ replenishment items`);
}

function brandSummary(){
 const m=new Map();state.computed.forEach(p=>{if(!m.has(p.brand))m.set(p.brand,{brand:p.brand,products:0,stock:0,fast:0,critical:0,reorder:0,wait:0,overstock:0,dead:0,review:0,suggested:0});const b=m.get(p.brand);b.products++;b.stock+=p.allCompany;b.fast+=p.movement==='FAST MOVING';b.critical+=p.condition==='CRITICAL ORDER';b.reorder+=p.condition==='REORDER SOON';b.wait+=p.condition==='WAIT INCOMING';b.overstock+=p.condition==='OVERSTOCK';b.dead+=p.condition==='DEAD STOCK';b.review+=['DORMANT / REVIEW','ONE-TIME / REVIEW','NO SALES / REVIEW'].includes(p.condition);b.suggested+=plannerQty(p);});return[...m.values()].sort((a,b)=>b.suggested-a.suggested||b.critical-a.critical||a.brand.localeCompare(b.brand));
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
 $('qualityGrid').innerHTML=`<div class="quality-card"><strong>${fmt(total)}</strong><span>Total product rows</span></div><div class="quality-card"><strong>${fmt(q.missing.oem)}</strong><span>Missing OEM (${((q.missing.oem/Math.max(1,total))*100).toFixed(1)}%)</span></div><div class="quality-card"><strong>${fmt(q.missing.brand)}</strong><span>Missing brand</span></div><div class="quality-card"><strong>${fmt(q.missing.lastSaleDate)}</strong><span>Sales rows missing Last Sale Date</span></div><div class="quality-card"><strong>${fmt(q.missing.negativeIncoming)}</strong><span>Negative incoming rows (calculation protected)</span></div>`;
 const checks=[
  ['Odoo Make field (otherwise derived from Internal Ref)',q.columns.make],
  ['Motorline On Way',q.columns.motorlineOnWay],['Tibao On Way',q.columns.tibaoOnWay],['Bavaria On Way',q.columns.bavariaOnWay],['Group On Way',q.columns.groupOnWay],
  ['Motorline On Way 2',q.columns.motorlineOnWay2],['Tibao On Way 2',q.columns.tibaoOnWay2],['Bavaria On Way 2',q.columns.bavariaOnWay2],['Group On Way 2',q.columns.groupOnWay2],
  ['Monthly sales Jan-Aug',q.columns.monthlySales],['Report Till Date column (title fallback supported)',q.columns.reportDate],
  ['Purchase Qty / Received Qty',q.columns.totalPurchase],['Purchase Count / Receipts',q.columns.purchaseCount],['Last Purchase Date',q.columns.lastPurchaseDate]
 ];
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
   <div class="metric"><div class="m-label">Current Stock-Out</div><div class="m-value small-date">${esc(stockoutFmt(p.currentStockoutDate,p.demandRate,p.allCompany))}</div></div>
   <div class="metric"><div class="m-label">With Incoming Stock-Out</div><div class="m-value small-date">${esc(stockoutFmt(p.pipelineStockoutDate,p.demandRate,p.directSupply))}</div></div>
   <div class="metric"><div class="m-label">Brand Lead Time</div><div class="m-value">${fmt(p.leadTimeDays)} days</div></div>
   <div class="metric"><div class="m-label">Expected New Order Arrival</div><div class="m-value small-date">${esc(dateFmt(p.expectedNewOrderArrivalDate))}</div></div>
  </div>
  <div class="section-block"><div class="section-block-title">Monthly Sales Pattern</div><div class="month-chart">${monthBars}</div><div class="micro-grid"><span>Overall Avg <strong>${fmt(p.avgMonthlySales,1)}</strong></span><span>Recent 3M <strong>${fmt(p.recent3Avg,1)}</strong></span><span>Recent 6M <strong>${fmt(p.recent6Avg,1)}</strong></span><span>Pattern <strong>${esc(p.demandPattern)}</strong></span><span>Largest month share <strong>${fmt(p.spikeSharePct,0)}%</strong></span><span>Last sale <strong>${esc(lastSale)}</strong></span><span>Zero-sales streak <strong>${fmt(p.zeroSalesStreakMonths,0)} months</strong></span><span>Inactivity basis <strong>${esc(p.inactivityBasis||'—')}</strong></span></div></div>
  <div class="section-block"><div class="section-block-title">Purchase History Signal</div><div class="micro-grid"><span>Total Purchase Qty <strong>${fmt(p.totalPurchase)}</strong></span><span>Purchase Count <strong>${fmt(p.purchaseCount)}</strong></span><span>Last Purchase <strong>${esc(lastPurchase)}</strong></span><span>Purchase / Sales <strong>${ratio(p.purchaseSalesRatio)}</strong></span><span>Signal <strong>${esc(p.purchaseSignal)}</strong></span></div><div class="note-box">Historical purchases are <strong>not subtracted again</strong> from the suggestion because receipts are already reflected in current stock. They are used to detect overbuying, recent replenishment and unusual purchase/sales patterns.</div></div>
  <div class="section-block"><div class="section-block-title">Suggested Purchase Calculation</div><div class="formula-box"><strong>${esc(p.condition)} — ${esc(p.action)}</strong><br>${esc(p.reason)}<br><br>Own current cover: <strong>${cover(p.currentCover)} mo</strong> • Effective current cover: <strong>${cover(p.effectiveCurrentCover)} mo</strong> • Effective cover with incoming: <strong>${cover(p.effectiveCover)} mo</strong> • Lead time: <strong>${fmt(p.leadTimeDays)} days</strong><br>Target gap before MOQ/multiple: <strong>${fmt(p.rawSuggestedQty)} pcs</strong> • MOQ: <strong>${fmt(p.moq)}</strong> • Order multiple: <strong>${fmt(p.orderMultiple)}</strong><br>Calculated suggested quantity: <strong>${fmt(p.suggestedQty)} pcs</strong>${state.qtyOverrides.has(p._id)?` • Purchaser override: <strong>${fmt(plannerQty(p))} pcs</strong>`:''}</div>${p.purchaseGroupName?`<div class="note-box group-credit-note"><strong>Purchasing Brand Group: ${esc(p.purchaseGroupName)}</strong><br>Exact same OEM only. Credited stock: <strong>${fmt(p.groupStockCredit)}</strong> • Credited On Way: <strong>${fmt(p.groupOnWayCredit)}</strong> • Credited On Way 2: <strong>${fmt(p.groupOnWay2Credit)}</strong>. Outside-group brands follow the general Other-brand stock credit setting.</div>`:''}</div>
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
    'Purchasing Brand Group':p.purchaseGroupName,'Group Credited Stock':p.groupStockCredit,'Group Credited On Way':p.groupOnWayCredit,'Group Credited On Way 2':p.groupOnWay2Credit,'Same OEM Other Brand Stock':p.otherStock,'Same OEM Other Brand On Way':p.otherOnWay,'Same OEM Other Brand On Way 2':p.otherOnWay2,'Same OEM Detail':p.equivalentNote,
    'Current Months Cover':p.currentCover===Infinity?'INF':+p.currentCover.toFixed(2),'Pipeline Months Cover':p.pipelineCover===Infinity?'INF':+p.pipelineCover.toFixed(2),
    'Current Stock-Out Date':stockoutFmt(p.currentStockoutDate,p.demandRate,p.allCompany),'With Incoming Stock-Out Date':stockoutFmt(p.pipelineStockoutDate,p.demandRate,p.directSupply),
    'Brand Lead Time Days':p.leadTimeDays,'Expected New Order Arrival':dateFmt(p.expectedNewOrderArrivalDate),'MOQ Setting':p.moq,'Order Multiple Setting':p.orderMultiple,
    'Stock Condition':p.condition,'Priority':p.priority,'Target Gap Qty':p.targetGapQty,'System Suggested Qty':p.suggestedQty,'Suggested Purchase Qty':plannerQty(p),'Purchase Action':p.action,'Reason':p.reason
   };
   months.forEach((m,i)=>out[`${m} Sales`]=p.monthly[i]||0);return out;
  });
 const ws=XLSX.utils.json_to_sheet(data),wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Purchase Suggestions');XLSX.writeFile(wb,name);toast(`Exported ${fmt(rows.length)} rows`);
}

function exportSelectedPurchase(rows,name){
 if(!rows.length){toast('Nothing selected to export');return;}
 const data=rows.map(p=>({
  'Old Number':p.internalRef,
  'OEM No Space':p.oemKey||E.normalizeOEM(p.oem),
  'Add Description':p.description,
  'Brand':p.brand,
  'Available Quantity':p.allCompany,
  'Total Incoming Quantity':p.onWay+p.onWay2,
  'Suggested Quantity':plannerQty(p)
 }));
 const ws=XLSX.utils.json_to_sheet(data),wb=XLSX.utils.book_new();
 ws['!cols']=[{wch:28},{wch:22},{wch:48},{wch:20},{wch:18},{wch:22},{wch:18}];
 XLSX.utils.book_append_sheet(wb,ws,'Selected Purchase');
 XLSX.writeFile(wb,name);toast(`Exported ${fmt(rows.length)} selected items`);
}

function groupBrandKey(v){return String(v??'').toUpperCase().replace(/[_\-]+/g,' ').replace(/\s+/g,' ').trim();}
function groupDraftCopy(groups){return (groups||[]).map((g,i)=>({id:String(g.id||`group_${Date.now()}_${i}`),name:String(g.name||`Brand Group ${i+1}`),enabled:g.enabled!==false,brands:[...(g.brands||[])],stockCreditPct:Number.isFinite(+g.stockCreditPct)?+g.stockCreditPct:100,onWayCreditPct:Number.isFinite(+g.onWayCreditPct)?+g.onWayCreditPct:100,onWay2CreditPct:Number.isFinite(+g.onWay2CreditPct)?+g.onWay2CreditPct:100}));}
function availableGroupBrands(){
 const m=new Map();state.products.forEach(p=>{if(p.brand)m.set(groupBrandKey(p.brand),p.brand);});
 state.brandGroupDrafts.forEach(g=>(g.brands||[]).forEach(k=>{if(!m.has(groupBrandKey(k)))m.set(groupBrandKey(k),k);}));
 return [...m.entries()].sort((a,b)=>String(a[1]).localeCompare(String(b[1])));
}
function groupTriggerLabel(g,brandMap){const n=(g.brands||[]).length;if(!n)return 'Select brands';if(n===1)return brandMap.get(groupBrandKey(g.brands[0]))||g.brands[0];return `${n} selected`;}
function closeBrandGroupMenus(exceptId=''){document.querySelectorAll('.brand-group-menu').forEach(m=>{if(m.id!==exceptId)m.classList.add('hidden');});}
function renderBrandGroupsUI(){
 const root=$('brandGroupsContainer');if(!root)return;
 const brands=availableGroupBrands(),brandMap=new Map(brands);
 if(!state.brandGroupDrafts.length){root.innerHTML='<div class="brand-group-empty">No purchasing brand group yet. Click <strong>Add TiBAO Family</strong> for your related brands, or create a custom group.</div>';return;}
 root.innerHTML=state.brandGroupDrafts.map(g=>{const selected=new Set((g.brands||[]).map(groupBrandKey));return `<div class="brand-group-card" data-group-id="${esc(g.id)}">
  <div class="brand-group-card-head"><input class="group-name" value="${esc(g.name)}" placeholder="Group name"><label class="group-enabled"><input type="checkbox" class="group-enabled-input" ${g.enabled?'checked':''}> Enabled</label><button type="button" class="btn group-remove">Remove</button></div>
  <div class="brand-group-grid">
   <div class="group-field"><label>Brands in group</label><div class="multi-select"><button type="button" class="multi-trigger group-brand-trigger">${esc(groupTriggerLabel(g,brandMap))}</button><div class="multi-menu brand-group-menu hidden" id="brandGroupMenu_${esc(g.id)}"><div class="multi-head"><input class="multi-search group-brand-search" placeholder="Search brands..." autocomplete="off"><div class="multi-actions"><button type="button" data-action="all">Select all</button><button type="button" data-action="clear">Clear</button></div></div><div class="multi-options">${brands.map(([key,name])=>`<label class="multi-option" data-label="${esc(String(name).toLowerCase())}"><input type="checkbox" value="${esc(key)}" ${selected.has(key)?'checked':''}><span>${esc(name)}</span></label>`).join('')}</div></div></div><p>Only exact normalized OEM matches are combined.</p></div>
   <div class="group-field"><label>Stock credit %</label><input class="group-pct" data-field="stockCreditPct" type="number" min="0" max="100" step="5" value="${g.stockCreditPct}"><p>Current same-OEM stock in grouped brands.</p></div>
   <div class="group-field"><label>On Way credit %</label><input class="group-pct" data-field="onWayCreditPct" type="number" min="0" max="100" step="5" value="${g.onWayCreditPct}"><p>Already shipped same-OEM stock.</p></div>
   <div class="group-field"><label>On Way 2 credit %</label><input class="group-pct" data-field="onWay2CreditPct" type="number" min="0" max="100" step="5" value="${g.onWay2CreditPct}"><p>Ordered / supplier preparing.</p></div>
  </div></div>`;}).join('');
 root.querySelectorAll('.brand-group-card').forEach(card=>{
  const id=card.dataset.groupId,g=state.brandGroupDrafts.find(x=>x.id===id);if(!g)return;
  card.querySelector('.group-name').oninput=e=>g.name=e.target.value;card.querySelector('.group-enabled-input').onchange=e=>g.enabled=e.target.checked;
  card.querySelectorAll('.group-pct').forEach(inp=>inp.oninput=()=>g[inp.dataset.field]=Math.max(0,Math.min(100,+inp.value||0)));
  card.querySelector('.group-remove').onclick=()=>{state.brandGroupDrafts=state.brandGroupDrafts.filter(x=>x.id!==id);renderBrandGroupsUI();};
  const trigger=card.querySelector('.group-brand-trigger'),menu=card.querySelector('.brand-group-menu');
  trigger.onclick=e=>{e.stopPropagation();closeBrandGroupMenus(menu.id);menu.classList.toggle('hidden');};menu.onclick=e=>e.stopPropagation();
  const refreshLabel=()=>trigger.textContent=groupTriggerLabel(g,brandMap);
  menu.querySelectorAll('.multi-option input').forEach(cb=>cb.onchange=()=>{const set=new Set((g.brands||[]).map(groupBrandKey));cb.checked?set.add(groupBrandKey(cb.value)):set.delete(groupBrandKey(cb.value));g.brands=[...set];refreshLabel();});
  menu.querySelector('[data-action="all"]').onclick=()=>{g.brands=brands.map(x=>x[0]);renderBrandGroupsUI();};menu.querySelector('[data-action="clear"]').onclick=()=>{g.brands=[];renderBrandGroupsUI();};
  const search=menu.querySelector('.group-brand-search');search.oninput=()=>{const q=search.value.trim().toLowerCase();menu.querySelectorAll('.multi-option').forEach(x=>x.classList.toggle('hidden',q&&!x.dataset.label.includes(q)));};
 });
}
function addBrandGroup(preset=false){
 const brands=availableGroupBrands();let selected=[];
 if(preset){const wanted=new Set(['TIBAO','TIBAO HD','TIBAO EXTRA','TG AUTOTEILE','SHOWORLD','SHOWORLD NEW']);selected=brands.filter(([k])=>wanted.has(groupBrandKey(k))).map(([k])=>k);}
 state.brandGroupDrafts.push({id:`group_${Date.now()}_${Math.random().toString(36).slice(2,6)}`,name:preset?'TIBAO FAMILY':`Brand Group ${state.brandGroupDrafts.length+1}`,enabled:true,brands:selected,stockCreditPct:100,onWayCreditPct:100,onWay2CreditPct:100});
 renderBrandGroupsUI();
 if(preset&&!selected.length)toast('TiBAO Family added — select the brands manually');else if(preset)toast(`TiBAO Family added with ${selected.length} matching brands`);
}
function validateBrandGroups(groups){const used=new Map();for(const g of groups){if(!g.enabled)continue;for(const raw of g.brands||[]){const k=groupBrandKey(raw);if(used.has(k))return `Brand ${k} is enabled in both “${used.get(k)}” and “${g.name}”. A brand can belong to only one enabled purchasing group.`;used.set(k,g.name);}}return '';}

function renderBrandRulesUI(){
 const table=$('brandRulesTable');if(!table)return;
 const brandNames=new Map();
 state.products.forEach(p=>{if(p.brand)brandNames.set(String(p.brand).toUpperCase(),p.brand);});
 Object.entries(state.settings.brandRules||{}).forEach(([k,r])=>brandNames.set(k,(r&&r.brand)||k));
 const rows=[...brandNames.entries()].sort((a,b)=>String(a[1]).localeCompare(String(b[1])));
 if(!rows.length){table.innerHTML='<tbody><tr><td colspan="4" class="empty">Upload an Excel file to configure brand-specific purchasing rules.</td></tr></tbody>';return;}
 const s=state.settings;
 table.innerHTML=`<thead><tr><th>Brand</th><th class="num">Lead Time Days</th><th class="num">MOQ (optional)</th><th class="num">Order Multiple (optional)</th></tr></thead><tbody>${rows.map(([key,name])=>{
  const r=(s.brandRules||{})[key]||{};
  const lead=Number.isFinite(+r.leadTimeDays)?+r.leadTimeDays:s.defaultLeadTimeDays;
  const moq=Number.isFinite(+r.moq)?+r.moq:s.defaultMOQ;
  const mult=Number.isFinite(+r.orderMultiple)&&+r.orderMultiple>0?+r.orderMultiple:s.defaultOrderMultiple;
  return `<tr class="brand-rule-row" data-search="${esc(String(name).toLowerCase())}"><td><strong>${esc(name)}</strong></td><td class="num"><input class="rule-input" data-brand-key="${esc(key)}" data-brand-name="${esc(name)}" data-field="leadTimeDays" type="number" min="0" step="1" value="${lead}"></td><td class="num"><input class="rule-input" data-brand-key="${esc(key)}" data-brand-name="${esc(name)}" data-field="moq" type="number" min="0" step="1" value="${moq}"></td><td class="num"><input class="rule-input" data-brand-key="${esc(key)}" data-brand-name="${esc(name)}" data-field="orderMultiple" type="number" min="1" step="1" value="${mult}"></td></tr>`;
 }).join('')}</tbody>`;
}
function syncSettingsUI(){
 const s=state.settings;$('sTarget').value=s.targetCover;$('sSafety').value=s.safetyCover;$('sCritical').value=s.criticalCover;$('sReorder').value=s.reorderCover;$('sOverstock').value=s.overstockCover;$('sDormant').value=s.dormantMonths;$('sDead').value=s.deadStockMonths;$('sFast').value=s.fastRate;$('sMedium').value=s.mediumRate;$('sSpike').value=s.spikeDominancePct;$('sRegular').value=s.regularActivePct;$('sEquivalent').value=s.equivalentCreditPct;$('sDemand').value=s.demandMethod;$('sSuggestionMode').value=s.suggestionMode;
 $('sUseLeadTime').value=s.useLeadTime?'on':'off';$('sDefaultLead').value=s.defaultLeadTimeDays;$('sApplyConstraints').value=s.applyOrderConstraints?'on':'off';$('sDefaultMOQ').value=s.defaultMOQ;$('sDefaultMultiple').value=s.defaultOrderMultiple;
 $('sShjShare').value=s.shjSharePct;$('sShjNormalWeeks').value=s.shjNormalCoverWeeks;$('sShjFastWeeks').value=s.shjFastCoverWeeks;$('sShjMLReserve').value=s.shjMotorlineReserveMonths;$('sShjBavaria').value=s.shjUseBavariaBackup?'on':'off';$('sShjMinDemand').value=s.shjMinGroupDemand;
 state.brandGroupDrafts=groupDraftCopy(s.brandGroups||[]);renderBrandGroupsUI();renderBrandRulesUI();
}
function applySettingsFromUI(){
 state.qtyOverrides.clear();
 state.settings.targetCover=+$('sTarget').value;state.settings.safetyCover=+$('sSafety').value;state.settings.criticalCover=+$('sCritical').value;state.settings.reorderCover=+$('sReorder').value;state.settings.overstockCover=+$('sOverstock').value;state.settings.dormantMonths=+$('sDormant').value;state.settings.deadStockMonths=+$('sDead').value;state.settings.fastRate=+$('sFast').value;state.settings.mediumRate=+$('sMedium').value;state.settings.spikeDominancePct=+$('sSpike').value;state.settings.regularActivePct=+$('sRegular').value;state.settings.equivalentCreditPct=+$('sEquivalent').value;state.settings.demandMethod=$('sDemand').value;state.settings.suggestionMode=$('sSuggestionMode').value;
 state.settings.useLeadTime=$('sUseLeadTime').value==='on';state.settings.defaultLeadTimeDays=Math.max(0,+$('sDefaultLead').value||0);state.settings.applyOrderConstraints=$('sApplyConstraints').value==='on';state.settings.defaultMOQ=Math.max(0,+$('sDefaultMOQ').value||0);state.settings.defaultOrderMultiple=Math.max(1,+$('sDefaultMultiple').value||1);
 state.settings.shjSharePct=Math.max(0,Math.min(100,+$('sShjShare').value||0));state.settings.shjNormalCoverWeeks=Math.max(0,+$('sShjNormalWeeks').value||0);state.settings.shjFastCoverWeeks=Math.max(0,+$('sShjFastWeeks').value||0);state.settings.shjMotorlineReserveMonths=Math.max(0,+$('sShjMLReserve').value||0);state.settings.shjUseBavariaBackup=$('sShjBavaria').value==='on';state.settings.shjMinGroupDemand=Math.max(0,+$('sShjMinDemand').value||0);
 const rules={...(state.settings.brandRules||{})};
 document.querySelectorAll('#brandRulesTable .rule-input').forEach(inp=>{
  const key=inp.dataset.brandKey,name=inp.dataset.brandName,field=inp.dataset.field;
  if(!rules[key])rules[key]={brand:name};
  rules[key].brand=name;
  if(field==='leadTimeDays')rules[key][field]=Math.max(0,+inp.value||0);
  else if(field==='moq')rules[key][field]=Math.max(0,+inp.value||0);
  else rules[key][field]=Math.max(1,+inp.value||1);
 });
 state.settings.brandRules=rules;
 const groups=groupDraftCopy(state.brandGroupDrafts).map((g,i)=>({id:g.id||`group_${i+1}`,name:(g.name||`Brand Group ${i+1}`).trim(),enabled:g.enabled!==false,brands:Array.from(new Set((g.brands||[]).map(groupBrandKey).filter(Boolean))),stockCreditPct:Math.max(0,Math.min(100,+g.stockCreditPct||0)),onWayCreditPct:Math.max(0,Math.min(100,+g.onWayCreditPct||0)),onWay2CreditPct:Math.max(0,Math.min(100,+g.onWay2CreditPct||0))})).filter(g=>g.brands.length>=2);
 const groupError=validateBrandGroups(groups);if(groupError){alert(groupError);return;}state.settings.brandGroups=groups;
 saveSettings();recompute();toast('Calculation settings applied');
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
document.addEventListener('click',()=>{$('brandFilterMenu').classList.add('hidden');$('conditionFilterMenu').classList.add('hidden');['shjBrandFilterMenu','shjMakeFilterMenu','shjStatusFilterMenu'].forEach(id=>$(id).classList.add('hidden'));closeBrandGroupMenus();});
$('clearFilters').onclick=()=>{state.selectedBrands.clear();state.selectedConditions.clear();['search','descriptionFilter','makeFilter','movementFilter','patternFilter','categoryFilter'].forEach(id=>$(id).value='');populateFilters();applyFilters();};
document.querySelectorAll('.kpi[data-condition]').forEach(k=>k.onclick=()=>{state.selectedConditions.clear();state.selectedConditions.add(k.dataset.condition);renderMultiMenu('condition',E.CONDITION_OPTIONS);applyFilters();});
document.querySelectorAll('.kpi[data-movement]').forEach(k=>k.onclick=()=>{$('movementFilter').value=k.dataset.movement;applyFilters();});
$('pageSize').onchange=()=>{state.pageSize=+$('pageSize').value;state.page=1;renderPlanner();};$('prevPage').onclick=()=>{if(state.page>1){state.page--;renderPlanner();}};$('nextPage').onclick=()=>{state.page++;renderPlanner();};
$('exportTop').onclick=$('exportFiltered').onclick=()=>exportRows(state.filtered,`Purchase_Suggestions_v3_3_${dateFmt(state.reportDate||new Date())}.xlsx`);$('exportSelected').onclick=()=>exportSelectedPurchase(state.filtered.filter(p=>state.selected.has(p._id)),`Selected_Purchase_Order_v3_3_${dateFmt(state.reportDate||new Date())}.xlsx`);
$('clearSelection').onclick=()=>{const had=state.selected.size;state.selected.clear();renderPlanner();updateSelectedButton();toast(had?'Selection cleared':'No items were selected');};
['shjSearch'].forEach(id=>$(id).addEventListener('input',applyShjFilters));$('shjActionableOnly').addEventListener('change',applyShjFilters);
function closeShjMultiMenus(){['shjBrandFilterMenu','shjMakeFilterMenu','shjStatusFilterMenu'].forEach(id=>$(id).classList.add('hidden'));}
[['brand','shjBrandFilterTrigger','shjBrandFilterMenu'],['make','shjMakeFilterTrigger','shjMakeFilterMenu'],['status','shjStatusFilterTrigger','shjStatusFilterMenu']].forEach(([type,triggerId,menuId])=>{
 $(triggerId).onclick=e=>{e.stopPropagation();$('brandFilterMenu').classList.add('hidden');$('conditionFilterMenu').classList.add('hidden');['shjBrandFilterMenu','shjMakeFilterMenu','shjStatusFilterMenu'].forEach(id=>{if(id!==menuId)$(id).classList.add('hidden');});$(menuId).classList.toggle('hidden');};
 $(menuId).onclick=e=>e.stopPropagation();
});
$('shjClearFilters').onclick=()=>{state.shjSelectedBrands.clear();state.shjSelectedMakes.clear();state.shjSelectedStatuses.clear();$('shjSearch').value='';$('shjActionableOnly').checked=true;populateShjFilters();applyShjFilters();};
$('shjExportFiltered').onclick=()=>exportShjRows(state.shjFiltered,`SHJ_Replenishment_Filtered_v3_3_${dateFmt(state.reportDate||new Date())}.xlsx`);$('shjExportSelected').onclick=()=>exportShjRows(state.shjFiltered.filter(p=>state.shjSelected.has(p._id)),`SHJ_Replenishment_Selected_v3_3_${dateFmt(state.reportDate||new Date())}.xlsx`);$('shjClearSelection').onclick=()=>{const had=state.shjSelected.size;state.shjSelected.clear();renderShj();updateShjSelectionButtons();toast(had?'SHJ selection cleared':'No SHJ items were selected');};
$('shjPageSize').onchange=()=>{state.shjPageSize=+$('shjPageSize').value;state.shjPage=1;renderShj();};$('shjPrevPage').onclick=()=>{if(state.shjPage>1){state.shjPage--;renderShj();}};$('shjNextPage').onclick=()=>{state.shjPage++;renderShj();};
$('modalClose').onclick=()=>$('modalBackdrop').classList.add('hidden');$('modalBackdrop').onclick=e=>{if(e.target===$('modalBackdrop'))$('modalBackdrop').classList.add('hidden');};
$('brandRuleSearch').oninput=()=>{const q=$('brandRuleSearch').value.trim().toLowerCase();document.querySelectorAll('#brandRulesTable .brand-rule-row').forEach(r=>r.classList.toggle('hidden',q&&!r.dataset.search.includes(q)));};
$('addBrandGroup').onclick=()=>addBrandGroup(false);$('addTibaoFamily').onclick=()=>addBrandGroup(true);
$('applySettings').onclick=applySettingsFromUI;$('resetSettings').onclick=()=>{state.settings=normalizeSettings({});state.brandGroupDrafts=[];state.qtyOverrides.clear();state.shjQtyOverrides.clear();state.shjSelected.clear();saveSettings();syncSettingsUI();recompute();toast('Defaults restored');};

populateFilters();renderDashboard();renderPlanner();
})();
