(function(){
'use strict';
const E = window.PurchaseEngine;
let state={products:[],computed:[],filtered:[],reportDate:null,fileName:'',map:null,settings:loadSettings(),page:1,pageSize:50,sortKey:'priority',sortDir:1,selected:new Set(),activeView:'dashboard'};

const $=id=>document.getElementById(id);
const esc=s=>String(s??'').replace(/[&<>"']/g,c=>({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#039;'}[c]));
const fmt=(x,d=0)=>Number.isFinite(x)?x.toLocaleString(undefined,{maximumFractionDigits:d}):'—';
const pct=(n,d)=>d?`${((n/d)*100).toFixed(1)}%`:'0%';
const dateFmt=d=>d instanceof Date&&!isNaN(d)?d.toISOString().slice(0,10):'—';

function loadSettings(){try{return Object.assign({},E.DEFAULT_SETTINGS,JSON.parse(localStorage.getItem('purchaseSettings')||'{}'));}catch(_){return {...E.DEFAULT_SETTINGS};}}
function saveSettings(){localStorage.setItem('purchaseSettings',JSON.stringify(state.settings));}
function toast(msg){const t=$('toast');t.textContent=msg;t.classList.remove('hidden');clearTimeout(toast.timer);toast.timer=setTimeout(()=>t.classList.add('hidden'),2600);}

function setView(name){state.activeView=name;document.querySelectorAll('.view').forEach(v=>v.classList.toggle('active',v.id===`view-${name}`));document.querySelectorAll('.nav button').forEach(b=>b.classList.toggle('active',b.dataset.view===name));const titles={dashboard:'Purchase Support',planner:'Purchase Planner',brands:'Brand Analysis',quality:'Data Quality',settings:'Calculation Settings'};$('pageTitle').textContent=titles[name]||'Purchase Support';if(name==='planner')renderPlanner();if(name==='brands')renderBrands();if(name==='quality')renderQuality();if(name==='settings')syncSettingsUI();}

function showUploadZone(show){$('uploadZone').classList.toggle('hidden',!show);}

async function readFile(file){
  if(!window.XLSX){throw new Error('Excel reader library did not load. Please check your internet connection and refresh.');}
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
    $('reportMeta').textContent=`${fmt(state.products.length)} products • Data as of ${dateFmt(state.reportDate)} • Browser-local analysis`;
    showUploadZone(false);$('exportTop').disabled=false;
    renderQuality();
    toast(`Loaded ${fmt(state.products.length)} products from ${file.name}`);
  }catch(err){console.error(err);$('fileChip').textContent='No file loaded';alert(err.message||String(err));}
}

function demoProducts(){
 const d=new Date(2026,7,23);
 const make=(id,ref,oem,brand,desc,cat,stock,ow,ow2,months,last,purchase=0)=>({
   _id:String(id),productId:String(id),internalRef:ref,oem,oemKey:E.normalizeOEM(oem),brand,brandKey:brand.toUpperCase(),brandPartNo:ref,description:desc,category:cat,motorlineStock:Math.floor(stock*.5),bavariaStock:Math.floor(stock*.2),tibaoStock:stock-Math.floor(stock*.5)-Math.floor(stock*.2),allCompany:stock,tibaoADRef:0,tibaoDXBRef:0,
   monthly:months,totalSales:months.reduce((a,b)=>a+b,0),salesCount:Math.max(1,Math.round(months.reduce((a,b)=>a+b,0)/3)),lastSaleDate:last?new Date(last):null,totalPurchase:purchase,purchaseCount:purchase?2:0,lastPurchaseDate:purchase?new Date(2026,4,1):null,motorlineOnWay:ow,bavariaOnWay:0,tibaoOnWay:0,onWay:ow,motorlineOnWay2:0,bavariaOnWay2:0,tibaoOnWay2:ow2,onWay2:ow2,reportDate:d
 });
 return [
  make(1,'FEBI-1001','1K0698151','FEBI','BRAKE PAD SET FRONT','Brake System',2,0,0,[18,20,17,22,19,21,18,16,0,0,0,0],'2026-08-20',50),
  make(2,'TB-1001','1K0698151','TIBAO','BRAKE PAD SET FRONT','Brake System',28,10,0,[9,8,11,10,12,9,10,8,0,0,0,0],'2026-08-18',80),
  make(3,'BREM-1001','1K0698151','BREMBO','BRAKE PAD SET FRONT','Brake System',7,0,0,[12,13,12,14,11,15,14,9,0,0,0,0],'2026-08-22',60),
  make(4,'LEM-2200','8K0199381','LEMFORDER','ENGINE MOUNT RIGHT','Engine Mounting',0,0,0,[8,7,9,10,8,9,9,8,0,0,0,0],'2026-08-21',36),
  make(5,'TB-2200','8K0199381','TIBAO','ENGINE MOUNT RIGHT','Engine Mounting',3,6,10,[4,5,5,6,5,4,5,4,0,0,0,0],'2026-08-10',40),
  make(6,'MANN-03C','03C115561H','MANN FILTER','OIL FILTER','Filters',55,20,10,[22,24,28,26,25,24,27,20,0,0,0,0],'2026-08-23',120),
  make(7,'HENGST-03C','03C115561H','HENGST','OIL FILTER','Filters',12,0,0,[7,8,9,8,9,8,7,6,0,0,0,0],'2026-08-17',50),
  make(8,'FEBI-7788','4G0407183','FEBI','CONTROL ARM BUSH','Suspension',120,0,0,[1,1,0,1,0,0,1,0,0,0,0,0],'2025-04-18',160),
  make(9,'TB-7788','4G0407183','TIBAO','CONTROL ARM BUSH','Suspension',18,0,0,[3,2,4,3,4,3,4,3,0,0,0,0],'2026-08-11',35),
  make(10,'BREM-5555','5Q0615301F','BREMBO','BRAKE DISC FRONT','Brake System',4,12,0,[15,18,16,19,17,18,20,13,0,0,0,0],'2026-08-22',65),
  make(11,'FEBI-5555','5Q0615301F','FEBI','BRAKE DISC FRONT','Brake System',0,0,18,[10,9,11,12,10,11,12,8,0,0,0,0],'2026-08-19',45),
  make(12,'VDO-9090','A2C59506225','VDO','FUEL PUMP','Fuel System',0,0,0,[3,3,4,4,5,5,6,5,0,0,0,0],'2026-08-16',20),
  make(13,'TB-9090','A2C59506225','TIBAO','FUEL PUMP','Fuel System',14,0,0,[2,2,2,3,3,3,3,2,0,0,0,0],'2026-08-12',25),
  make(14,'FEBI-3333','7L0407182E','FEBI','SUSPENSION BUSH','Suspension',26,0,0,[0,0,0,0,0,0,0,0,0,0,0,0],'2024-11-02',50),
  make(15,'TRUCK-8888','06H103495','TRUCKTEC','PCV VALVE','Engine',6,0,8,[6,7,6,8,7,8,9,7,0,0,0,0],'2026-08-15',35)
 ];
}

function loadDemo(){state.products=demoProducts();state.reportDate=new Date(2026,7,23);state.map={bavariaOnWay:1,tibaoOnWay:1,reportDate:1,groupOnWay:1,groupOnWay2:1};state.fileName='Demo data';state.selected.clear();recompute();$('fileChip').textContent='Demo data';$('reportMeta').textContent=`${state.products.length} demo products • Data as of ${dateFmt(state.reportDate)} • Same-OEM matching enabled`;showUploadZone(false);$('exportTop').disabled=false;toast('Demo data loaded');}

function recompute(){
 state.computed=E.calculate(state.products,state.settings,state.reportDate||new Date());
 populateFilters();applyFilters();renderDashboard();renderPlanner();renderBrands();renderQuality();
}

function uniqueSorted(key){return [...new Set(state.computed.map(x=>x[key]).filter(Boolean))].sort((a,b)=>String(a).localeCompare(String(b)));}
function fillSelect(id,values,label){const el=$(id),cur=el.value;el.innerHTML=`<option value="">${label}</option>`+values.map(v=>`<option value="${esc(v)}">${esc(v)}</option>`).join('');if(values.includes(cur))el.value=cur;}
function populateFilters(){fillSelect('brandFilter',uniqueSorted('brand'),'All Brands');fillSelect('conditionFilter',uniqueSorted('condition'),'All Conditions');fillSelect('movementFilter',uniqueSorted('movement'),'All Movement');fillSelect('categoryFilter',uniqueSorted('category'),'All Categories');}

function applyFilters(){
 const q=$('search').value.trim().toLowerCase(),brand=$('brandFilter').value,cond=$('conditionFilter').value,mov=$('movementFilter').value,cat=$('categoryFilter').value;
 state.filtered=state.computed.filter(p=>{
  if(brand&&p.brand!==brand)return false;if(cond&&p.condition!==cond)return false;if(mov&&p.movement!==mov)return false;if(cat&&p.category!==cat)return false;
  if(q){const hay=[p.internalRef,p.oem,p.brand,p.brandPartNo,p.description,p.category].join(' ').toLowerCase();if(!hay.includes(q))return false;}
  return true;
 });
 state.page=1;renderDashboard();renderPlanner();
}

function count(cond){return state.filtered.filter(cond).length;}
function renderDashboard(){
 if(!state.computed.length){['kpiCritical','kpiReorder','kpiWait','kpiFast','kpiDead','kpiSuggested'].forEach(id=>$(id).textContent='0');$('statusBars').innerHTML='<div class="empty">Upload data to see purchasing intelligence.</div>';$('dashboardTable').innerHTML='';return;}
 $('kpiCritical').textContent=fmt(count(p=>p.condition==='CRITICAL ORDER'));
 $('kpiReorder').textContent=fmt(count(p=>p.condition==='REORDER SOON'));
 $('kpiWait').textContent=fmt(count(p=>p.condition==='WAIT INCOMING'));
 $('kpiFast').textContent=fmt(count(p=>p.movement==='FAST MOVING'));
 $('kpiDead').textContent=fmt(count(p=>p.condition==='DEAD STOCK'));
 $('kpiSuggested').textContent=fmt(state.filtered.reduce((s,p)=>s+p.suggestedQty,0));
 const statuses=['CRITICAL ORDER','REORDER SOON','WAIT INCOMING','OK','DEAD STOCK'];const max=Math.max(1,...statuses.map(s=>count(p=>p.condition===s)));
 $('statusBars').innerHTML=statuses.map(s=>{const c=count(p=>p.condition===s);return `<div class="status-line"><span>${esc(s)}</span><div class="bar"><span style="width:${(c/max)*100}%"></span></div><strong class="num">${fmt(c)}</strong></div>`}).join('');
 const rows=[...state.filtered].sort((a,b)=>a.priority-b.priority||b.suggestedQty-a.suggestedQty).slice(0,12);renderCompactTable($('dashboardTable'),rows);
}

function conditionClass(s){if(s==='CRITICAL ORDER')return'critical';if(s==='REORDER SOON')return'reorder';if(s==='WAIT INCOMING')return'wait';if(s==='OK')return'ok';if(s==='DEAD STOCK')return'dead';if(s==='OVERSTOCK')return'overstock';return'review';}
function movementClass(s){if(s==='FAST MOVING')return'fast';if(s==='MEDIUM MOVING')return'medium';if(s==='SLOW MOVING')return'slow';return'review';}
function cover(x){return x===Infinity?'∞':fmt(x,1);}

function renderCompactTable(el,rows){
 el.innerHTML=`<thead><tr><th>Internal Ref</th><th>Add Description</th><th>OEM</th><th>Brand</th><th class="num">All Company</th><th class="num">On Way</th><th class="num">On Way 2</th><th>Same OEM Other Brands</th><th class="num">Demand / Mo</th><th class="num">Cover</th><th>Condition</th><th class="num">Suggested Qty</th><th>Action</th></tr></thead><tbody>${rows.map(p=>`<tr data-id="${esc(p._id)}"><td class="nowrap"><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(p.internalRef||'—')}</button></td><td class="desc" title="${esc(p.description)}">${esc(p.description||'—')}</td><td>${esc(p.oem||'—')}</td><td><strong>${esc(p.brand)}</strong></td><td class="num">${fmt(p.allCompany)}</td><td class="num">${fmt(p.onWay)}</td><td class="num">${fmt(p.onWay2)}</td><td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(p.equivalentNote)}</button></td><td class="num">${fmt(p.demandRate,1)}</td><td class="num">${cover(p.pipelineCover)}</td><td><span class="badge ${conditionClass(p.condition)}">${esc(p.condition)}</span></td><td class="num qty">${fmt(p.suggestedQty)}</td><td>${esc(p.action)}</td></tr>`).join('')}</tbody>`;
 bindDetailLinks(el);
}

const PLANNER_COLS=[
 ['internalRef','Internal Ref'],['description','Add Description'],['oem','OEM'],['brand','Brand'],['brandPartNo','Brand Part No.'],['allCompany','All Company Qty'],['onWay','On Way'],['onWay2','On Way 2'],['totalPurchase','Purchase Qty'],['totalSales','Sales Qty'],['demandRate','Demand/Mo'],['movement','Movement'],['equivalentNote','Same OEM Other Brands'],['currentCover','Current Cover'],['pipelineCover','Pipeline Cover'],['condition','Stock Condition'],['priority','Priority'],['suggestedQty','Suggested Purchase'],['action','Purchase Action']
];
function sortValue(p,key){const v=p[key];if(v===Infinity)return 999999999;if(typeof v==='number')return v;return String(v||'').toLowerCase();}
function sortRows(rows){const {sortKey,sortDir}=state;return [...rows].sort((a,b)=>{const x=sortValue(a,sortKey),y=sortValue(b,sortKey);if(x<y)return-sortDir;if(x>y)return sortDir;return a.priority-b.priority;});}
function renderPlanner(){
 const el=$('plannerTable');if(!el)return;if(!state.computed.length){el.innerHTML='<tbody><tr><td class="empty">Upload an Excel file first.</td></tr></tbody>';$('plannerResults').textContent='0 items';return;}
 const sorted=sortRows(state.filtered);const total=sorted.length;const pages=Math.max(1,Math.ceil(total/state.pageSize));state.page=Math.min(state.page,pages);const start=(state.page-1)*state.pageSize;const rows=sorted.slice(start,start+state.pageSize);
 $('plannerResults').textContent=`${fmt(total)} items • ${fmt(state.filtered.reduce((s,p)=>s+p.suggestedQty,0))} pcs suggested`;$('pageInfo').textContent=`Page ${state.page} of ${pages} • showing ${total?start+1:0}-${Math.min(start+state.pageSize,total)} of ${total}`;$('prevPage').disabled=state.page<=1;$('nextPage').disabled=state.page>=pages;
 el.innerHTML=`<thead><tr><th><input type="checkbox" id="selectPage"></th>${PLANNER_COLS.map(([k,l])=>`<th data-sort="${k}">${esc(l)}${state.sortKey===k?(state.sortDir>0?' ▲':' ▼'):''}</th>`).join('')}</tr></thead><tbody>${rows.map(p=>`<tr><td><input class="row-select" type="checkbox" data-id="${esc(p._id)}" ${state.selected.has(p._id)?'checked':''}></td>${PLANNER_COLS.map(([k])=>plannerCell(p,k)).join('')}</tr>`).join('')}</tbody>`;
 el.querySelectorAll('th[data-sort]').forEach(th=>th.onclick=()=>{const k=th.dataset.sort;if(state.sortKey===k)state.sortDir*=-1;else{state.sortKey=k;state.sortDir=k==='priority'?1:-1;}renderPlanner();});
 const sp=$('selectPage');if(sp)sp.onchange=e=>{rows.forEach(p=>e.target.checked?state.selected.add(p._id):state.selected.delete(p._id));renderPlanner();updateSelection();};
 el.querySelectorAll('.row-select').forEach(cb=>cb.onchange=e=>{e.target.checked?state.selected.add(e.target.dataset.id):state.selected.delete(e.target.dataset.id);updateSelection();});bindDetailLinks(el);updateSelection();
}
function plannerCell(p,k){let v=p[k];if(k==='internalRef')return`<td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(v||'—')}</button></td>`;if(k==='description')return`<td class="desc" title="${esc(v)}">${esc(v||'—')}</td>`;if(['allCompany','onWay','onWay2','totalPurchase','totalSales','suggestedQty'].includes(k))return`<td class="num ${k==='suggestedQty'?'qty':''}">${fmt(v)}</td>`;if(['demandRate','currentCover','pipelineCover'].includes(k))return`<td class="num">${k.includes('Cover')?cover(v):fmt(v,1)}</td>`;if(k==='movement')return`<td><span class="badge ${movementClass(v)}">${esc(v)}</span></td>`;if(k==='condition')return`<td><span class="badge ${conditionClass(v)}">${esc(v)}</span></td>`;if(k==='priority')return`<td><span class="priority p${v}">${v}</span></td>`;if(k==='equivalentNote')return`<td><button class="mini-link open-detail" data-id="${esc(p._id)}">${esc(v)}</button></td>`;return`<td>${esc(v||'—')}</td>`;}
function updateSelection(){$('exportSelected').textContent=`Export Selected (${state.selected.size})`;$('exportSelected').disabled=state.selected.size===0;}

function brandSummary(){const m=new Map();state.computed.forEach(p=>{if(!m.has(p.brand))m.set(p.brand,{brand:p.brand,products:0,stock:0,critical:0,reorder:0,fast:0,dead:0,overstock:0,suggested:0});const b=m.get(p.brand);b.products++;b.stock+=p.allCompany;b.suggested+=p.suggestedQty;if(p.condition==='CRITICAL ORDER')b.critical++;if(p.condition==='REORDER SOON')b.reorder++;if(p.movement==='FAST MOVING')b.fast++;if(p.condition==='DEAD STOCK')b.dead++;if(p.condition==='OVERSTOCK')b.overstock++;});return [...m.values()].sort((a,b)=>b.suggested-a.suggested||b.critical-a.critical);}
function renderBrands(){if(!state.computed.length){$('brandGrid').innerHTML='<div class="empty">Upload data to analyse brands.</div>';$('brandTable').innerHTML='';return;}const rows=brandSummary();$('brandGrid').innerHTML=rows.slice(0,12).map(b=>`<div class="brand-card" data-brand="${esc(b.brand)}"><div class="brand-name">${esc(b.brand)}</div><div class="brand-meta">${fmt(b.products)} products • ${fmt(b.stock)} pcs company stock</div><div class="brand-numbers"><div><strong>${fmt(b.critical)}</strong><span>Critical</span></div><div><strong>${fmt(b.dead)}</strong><span>Dead Stock</span></div><div><strong>${fmt(b.suggested)}</strong><span>Suggest Qty</span></div></div></div>`).join('');$('brandGrid').querySelectorAll('.brand-card').forEach(c=>c.onclick=()=>selectBrand(c.dataset.brand));$('brandTable').innerHTML=`<thead><tr><th>Brand</th><th class="num">Products</th><th class="num">Company Stock</th><th class="num">Fast Moving</th><th class="num">Critical</th><th class="num">Reorder Soon</th><th class="num">Overstock</th><th class="num">Dead Stock</th><th class="num">Suggested Purchase</th></tr></thead><tbody>${rows.map(b=>`<tr class="brand-row" data-brand="${esc(b.brand)}"><td><button class="mini-link">${esc(b.brand)}</button></td><td class="num">${fmt(b.products)}</td><td class="num">${fmt(b.stock)}</td><td class="num">${fmt(b.fast)}</td><td class="num">${fmt(b.critical)}</td><td class="num">${fmt(b.reorder)}</td><td class="num">${fmt(b.overstock)}</td><td class="num">${fmt(b.dead)}</td><td class="num qty">${fmt(b.suggested)}</td></tr>`).join('')}</tbody>`;$('brandTable').querySelectorAll('.brand-row').forEach(r=>r.onclick=()=>selectBrand(r.dataset.brand));}
function selectBrand(brand){$('brandFilter').value=brand;applyFilters();setView('dashboard');toast(`Dashboard filtered to ${brand}`);}

function renderQuality(){if(!state.products.length){$('qualityGrid').innerHTML='<div class="empty">Upload data to run checks.</div>';$('columnChecks').innerHTML='';return;}const q=E.dataQuality(state.products,state.map||{}),total=state.products.length;$('qualityGrid').innerHTML=`<div class="quality-card"><strong>${fmt(q.missing.oem)}</strong><span>Missing OEM (${pct(q.missing.oem,total)})</span></div><div class="quality-card"><strong>${fmt(q.missing.brand)}</strong><span>Missing Brand (${pct(q.missing.brand,total)})</span></div><div class="quality-card"><strong>${fmt(q.missing.description)}</strong><span>Missing Add Description (${pct(q.missing.description,total)})</span></div><div class="quality-card"><strong>${fmt(q.missing.internalRef)}</strong><span>Missing Internal Reference (${pct(q.missing.internalRef,total)})</span></div>`;const checks=[['Bavaria On Way Qty',q.columns.bavariaOnWay],['Tibao SHJ On Way Qty',q.columns.tibaoOnWay],['Report Till Date column',q.columns.reportDate],['Group On Way total',q.columns.groupOnWay],['Group On Way 2 total',q.columns.groupOnWay2]];$('columnChecks').innerHTML=checks.map(([l,ok])=>`<div class="check-item"><span>${esc(l)}</span><span class="${ok?'yes':'no'}">${ok?'FOUND':'NOT FOUND'}</span></div>`).join('');}

function openDetail(id){const p=state.computed.find(x=>x._id===id);if(!p)return;$('modalTitle').textContent=`${p.internalRef || p.brandPartNo || 'Product'} — ${p.description||''}`;$('modalSub').textContent=`${p.brand} • OEM ${p.oem||'—'} • ${p.condition}`;
 const altRows=p.alternatives.length?p.alternatives.map(a=>`<tr><td><strong>${esc(a.brand)}</strong></td><td>${esc(a.partNumbers.join(', ')||'—')}</td><td class="num">${fmt(a.stock)}</td><td class="num">${fmt(a.onWay)}</td><td class="num">${fmt(a.onWay2)}</td><td class="num">${fmt(a.sales)}</td></tr>`).join(''):`<tr><td colspan="6" class="empty">No same-OEM product found in a different brand.</td></tr>`;
 $('modalBody').innerHTML=`<div class="detail-grid"><div class="metric"><div class="m-label">All Company Qty</div><div class="m-value">${fmt(p.allCompany)}</div></div><div class="metric"><div class="m-label">On Way</div><div class="m-value">${fmt(p.onWay)}</div></div><div class="metric"><div class="m-label">On Way 2</div><div class="m-value">${fmt(p.onWay2)}</div></div><div class="metric"><div class="m-label">Suggested Purchase</div><div class="m-value" style="color:var(--purple)">${fmt(p.suggestedQty)}</div></div><div class="metric"><div class="m-label">Overall Avg / Month</div><div class="m-value">${fmt(p.avgMonthlySales,1)}</div></div><div class="metric"><div class="m-label">Recent 3-Month Avg</div><div class="m-value">${fmt(p.recent3Avg,1)}</div></div><div class="metric"><div class="m-label">Pipeline Cover</div><div class="m-value">${cover(p.pipelineCover)} mo</div></div><div class="metric"><div class="m-label">Other Brand Stock</div><div class="m-value">${fmt(p.otherStock)}</div></div></div>
 <div class="formula-box"><strong>Why this recommendation?</strong><br>${esc(p.reason)}.<br><br><strong>Calculation:</strong> demand rate ${fmt(p.demandRate,1)} × target+buffer ${(state.settings.targetCover+state.settings.safetyCover).toFixed(1)} months = ${fmt(p.targetQty,1)} target pcs. Direct supply = ${fmt(p.allCompany)} company stock + ${fmt(p.onWay)} On Way + ${fmt(p.onWay2)} On Way 2. Same-OEM other-brand supply = ${fmt(p.otherSupply)}; management-approved credit currently ${state.settings.equivalentCreditPct}%, so ${fmt(p.equivalentCredit,1)} pcs are deducted from the suggestion.</div>
 <h3 style="font-size:13px;margin:18px 0 6px">Same OEM in other brands</h3><div class="panel-note">Matching is by normalized OEM number only — no supersession logic in this prototype.</div><table class="alt-table"><thead><tr><th>Brand</th><th>Brand Part No.</th><th class="num">Stock</th><th class="num">On Way</th><th class="num">On Way 2</th><th class="num">Report Sales</th></tr></thead><tbody>${altRows}</tbody></table>`;
 $('modalBackdrop').classList.remove('hidden');}
function bindDetailLinks(root){root.querySelectorAll('.open-detail').forEach(x=>x.onclick=e=>{e.stopPropagation();openDetail(x.dataset.id);});}

function exportRows(rows,name){if(!window.XLSX){alert('Excel export library is not loaded.');return;}const data=rows.map(p=>({
 'Odoo Product ID':p.productId,'Internal Reference':p.internalRef,'Main OEM No Space':p.oem,'Brand':p.brand,'Brand Part No.':p.brandPartNo,'Add Description':p.description,'Product Category':p.category,
 'All Company Qty':p.allCompany,'On Way':p.onWay,'On Way 2':p.onWay2,'Historical Purchase Qty':p.totalPurchase,'Report Sales Qty':p.totalSales,'Avg Monthly Sales':+p.avgMonthlySales.toFixed(2),'Recent 3M Avg':+p.recent3Avg.toFixed(2),'Demand Rate':+p.demandRate.toFixed(2),'Movement':p.movement,
 'Other Brand Stock':p.otherStock,'Other Brand On Way':p.otherOnWay,'Other Brand On Way 2':p.otherOnWay2,'Same OEM Other Brands':p.alternatives.map(a=>`${a.brand}: ${a.stock} stock + ${a.onWay+a.onWay2} incoming`).join(' | '),
 'Current Months Cover':p.currentCover===Infinity?'INF':+p.currentCover.toFixed(2),'Pipeline Months Cover':p.pipelineCover===Infinity?'INF':+p.pipelineCover.toFixed(2),'Stock Condition':p.condition,'Priority':p.priority,'Suggested Purchase Qty':p.suggestedQty,'Purchase Action':p.action,'Reason':p.reason
 }));const ws=XLSX.utils.json_to_sheet(data);const wb=XLSX.utils.book_new();XLSX.utils.book_append_sheet(wb,ws,'Purchase Suggestions');XLSX.writeFile(wb,name);toast(`Exported ${fmt(rows.length)} rows`);}

function syncSettingsUI(){const s=state.settings;$('sTarget').value=s.targetCover;$('sSafety').value=s.safetyCover;$('sCritical').value=s.criticalCover;$('sReorder').value=s.reorderCover;$('sOverstock').value=s.overstockCover;$('sDead').value=s.deadStockMonths;$('sFast').value=s.fastRate;$('sMedium').value=s.mediumRate;$('sEquivalent').value=s.equivalentCreditPct;$('sDemand').value=s.demandMethod;}
function readSettingsUI(){return Object.assign({},state.settings,{targetCover:+$('sTarget').value,safetyCover:+$('sSafety').value,criticalCover:+$('sCritical').value,reorderCover:+$('sReorder').value,overstockCover:+$('sOverstock').value,deadStockMonths:+$('sDead').value,fastRate:+$('sFast').value,mediumRate:+$('sMedium').value,equivalentCreditPct:+$('sEquivalent').value,demandMethod:$('sDemand').value});}

// Events
document.querySelectorAll('.nav button').forEach(b=>b.onclick=()=>setView(b.dataset.view));document.querySelectorAll('[data-go]').forEach(b=>b.onclick=()=>setView(b.dataset.go));
$('uploadBtn').onclick=$('uploadTop').onclick=()=>$('fileInput').click();$('fileInput').onchange=e=>handleFile(e.target.files[0]);$('demoBtn').onclick=loadDemo;
['search','brandFilter','conditionFilter','movementFilter','categoryFilter'].forEach(id=>$(id).addEventListener(id==='search'?'input':'change',applyFilters));$('clearFilters').onclick=()=>{['search','brandFilter','conditionFilter','movementFilter','categoryFilter'].forEach(id=>$(id).value='');applyFilters();};
document.querySelectorAll('.kpi[data-condition]').forEach(k=>k.onclick=()=>{$('conditionFilter').value=k.dataset.condition;applyFilters();});document.querySelectorAll('.kpi[data-movement]').forEach(k=>k.onclick=()=>{$('movementFilter').value=k.dataset.movement;applyFilters();});
$('pageSize').onchange=e=>{state.pageSize=+e.target.value;state.page=1;renderPlanner();};$('prevPage').onclick=()=>{if(state.page>1){state.page--;renderPlanner();}};$('nextPage').onclick=()=>{state.page++;renderPlanner();};
$('exportTop').onclick=$('exportFiltered').onclick=()=>exportRows(state.filtered,`Purchase_Suggestions_${dateFmt(state.reportDate||new Date())}.xlsx`);$('exportSelected').onclick=()=>exportRows(state.computed.filter(p=>state.selected.has(p._id)),`Selected_Purchase_Suggestions_${dateFmt(state.reportDate||new Date())}.xlsx`);
$('modalClose').onclick=()=>$('modalBackdrop').classList.add('hidden');$('modalBackdrop').onclick=e=>{if(e.target.id==='modalBackdrop')$('modalBackdrop').classList.add('hidden');};
$('applySettings').onclick=()=>{state.settings=readSettingsUI();saveSettings();recompute();toast('Calculation settings applied');};$('resetSettings').onclick=()=>{state.settings={...E.DEFAULT_SETTINGS};saveSettings();syncSettingsUI();recompute();toast('Settings reset to prototype defaults');};
const zone=$('uploadZone');['dragenter','dragover'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.add('drag');}));['dragleave','drop'].forEach(ev=>zone.addEventListener(ev,e=>{e.preventDefault();zone.classList.remove('drag');}));zone.addEventListener('drop',e=>handleFile(e.dataTransfer.files[0]));

syncSettingsUI();renderDashboard();
})();
