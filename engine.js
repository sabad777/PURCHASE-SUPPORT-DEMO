/* TiBAO Purchase Intelligence v2.4 - calculation engine
 * Pure browser-side logic. No server calls and no data persistence.
 * Purchase history is used as a decision signal / warning, not double-counted as stock.
 */
(function (global) {
  'use strict';

  const MONTH_KEYS = [
    [['Jan Ext Sales','Jan'], 'jan', 'Jan'], [['Feb Ext Sales','Feb'], 'feb', 'Feb'], [['Mar Ext Sales','Mar'], 'mar', 'Mar'],
    [['Apr Ext Sales','Apr'], 'apr', 'Apr'], [['May Ext Sales','May'], 'may', 'May'], [['Jun Ext Sales','Jun'], 'jun', 'Jun'],
    [['Jul Ext Sales','Jul'], 'jul', 'Jul'], [['Aug Ext Sales','Aug'], 'aug', 'Aug'], [['Sep Ext Sales','Sep'], 'sep', 'Sep'],
    [['Oct Ext Sales','Oct'], 'oct', 'Oct'], [['Nov Ext Sales','Nov'], 'nov', 'Nov'], [['Dec Ext Sales','Dec'], 'dec', 'Dec']
  ];

  const CONDITION_OPTIONS = [
    'CRITICAL ORDER','REORDER SOON','WAIT INCOMING','OK','OVERSTOCK','DORMANT / REVIEW',
    'DEAD STOCK','ONE-TIME / REVIEW','NO SALES / REVIEW','NO DEMAND'
  ];
  const MOVEMENT_OPTIONS = ['FAST MOVING','MEDIUM MOVING','SLOW MOVING','NO SALES','DEAD STOCK'];
  const PATTERN_OPTIONS = ['REGULAR','RISING','FALLING','INTERMITTENT','ONE-TIME SPIKE','DORMANT','DEAD STOCK','NO SALES'];

  const DEFAULT_SETTINGS = {
    targetCover: 3.0,
    safetyCover: 0.5,
    criticalCover: 0.5,
    reorderCover: 1.5,
    overstockCover: 6.0,
    dormantMonths: 6,
    deadStockMonths: 12,
    fastRate: 5.0,
    mediumRate: 1.5,
    equivalentCreditPct: 0,
    demandMethod: 'smart',
    suggestionMode: 'proactive',
    currentMonthNormalize: true,
    spikeDominancePct: 70,
    spikeMaxActiveMonths: 2,
    regularActivePct: 60
  };

  const HEADER_ALIASES = {
    productId: ['Odoo Product ID', 'Product ID'],
    internalRef: ['Internal Reference', 'Internal Ref', 'Item Code', 'Default Code'],
    oem: ['Main OEM No Space', 'Main OEM', 'OEM', 'OEM No Space'],
    brand: ['Brand'],
    brandPartNo: ['Brand Part No.', 'Brand Part No', 'Brand Part Number'],
    description: ['Add Description', 'Description'],
    category: ['Product Category', 'Category'],
    make: ['Make', 'Vehicle Make', 'Car Make', 'Vehicle Brand', 'Car Brand', 'Vehicle Manufacturer'],
    motorlineStock: ['Motorline Stock', 'Motorline'],
    bavariaStock: ['Bavaria Stock', 'Bavaria'],
    tibaoStock: ['Tibao SHJ Stock', 'Tibao Stock', 'Tibao SHJ'],
    allCompany: ['All Company Qty (Core)', 'All Company Qty', 'All Company Quantity', 'All Company (Core)'],
    tibaoADRef: ['Tibao AD Stock (Ref)', 'Tibao AD Stock', 'Tibao AD *'],
    tibaoDXBRef: ['Tibao DXB Stock (Ref)', 'Tibao DXB Stock', 'Tibao DXB *'],
    totalSales: ['Total External Sales Qty', 'Total External Sales', 'Sales Qty', 'Total Qty'],
    salesCount: ['External Sales Invoice Count', 'Sales Invoice Count', 'Sales Count', 'Invoices'],
    lastSaleDate: ['Last External Sale Date', 'Last Sale Date', 'Last Sale'],
    totalPurchase: ['Total External Purchase Qty', 'Total Purchase Qty', 'Purchase Qty', 'Received Qty'],
    purchaseCount: ['External Purchase Count', 'Purchase Count', 'Receipts'],
    lastPurchaseDate: ['Last External Purchase Date', 'Last Purchase Date', 'Last Purchase'],
    motorlineOnWay: ['Motorline On Way', 'ML On Way'],
    bavariaOnWay: ['Bavaria On Way'],
    tibaoOnWay: ['Tibao On Way', 'Tibao SHJ On Way'],
    groupOnWay: ['Group On Way', 'All Company On Way', 'Total On Way'],
    motorlineOnWay2: ['Motorline On Way 2', 'ML On Way 2'],
    bavariaOnWay2: ['Bavaria On Way 2'],
    tibaoOnWay2: ['Tibao On Way 2', 'Tibao SHJ On Way 2'],
    groupOnWay2: ['Group On Way 2', 'All Company On Way 2', 'Total On Way 2'],
    reportDate: ['Report Till Date', 'Data As Of', 'As Of Date', 'Report Date']
  };

  function n(v) {
    if (v === null || v === undefined || v === '') return 0;
    if (typeof v === 'number' && Number.isFinite(v)) return v;
    const cleaned = String(v).replace(/,/g, '').trim();
    const out = Number(cleaned);
    return Number.isFinite(out) ? out : 0;
  }
  function text(v) { return v === null || v === undefined ? '' : String(v).trim(); }
  function normalizeKey(v) { return text(v).toLowerCase().replace(/[\s._()\-\/]+/g, ''); }
  function normalizeOEM(v) { return text(v).toUpperCase().replace(/[^A-Z0-9]/g, ''); }

  // Vehicle make mapping from the confirmed old/internal-reference prefixes used by the business.
  // Exact token matching prevents BE (Bentley) from ever being confused with BENZ (Mercedes-Benz).
  const MAKE_CODE_MAP = {
    'VW':'VW','WV':'VW','VOLKSWAGEN':'VW',
    'BENZ':'BENZ','MERC':'BENZ','MERCEDES':'BENZ','MERCEDES BENZ':'BENZ',
    'BMW':'BMW',
    'ROV':'LAND ROVER','LR':'LAND ROVER','LAND ROVER':'LAND ROVER','LANDROVER':'LAND ROVER',
    'TY':'TOYOTA','TOYOTA':'TOYOTA',
    'BE':'BENTLEY','BENTLEY':'BENTLEY',
    'P':'PORSCHE','PO':'PORSCHE','POR':'PORSCHE','PORSCHE':'PORSCHE',
    'SKD':'SKODA','SKODA':'SKODA',
    'AD':'AUDI','AU':'AUDI','AUDI':'AUDI',
    'UN':'UNIVERSAL','UNIVERSAL':'UNIVERSAL',
    'HY':'HYUNDAI','HYUNDAI':'HYUNDAI',
    'JE':'JEEP','JEEP':'JEEP',
    'JA':'JAGUAR','JAG':'JAGUAR','JAGUAR':'JAGUAR',
    'HO':'HONDA','HONDA':'HONDA',
    'OP':'OPEL','OPEL':'OPEL',
    'SK':'SKODA',
    'MZ':'MAZDA','MAZDA':'MAZDA',
    'MI':'MITSUBISHI','MITSUBISHI':'MITSUBISHI',
    'MAS':'MASERATI','MASERATI':'MASERATI',
    'SEAT':'SEAT/CUPRA','CUPRA':'SEAT/CUPRA',
    'MINI':'MINI','VOLVO':'VOLVO','SUBARU':'SUBARU','FORD':'FORD'
  };
  function canonicalMake(v) {
    const raw = text(v).toUpperCase().replace(/[_-]+/g,' ').replace(/\s+/g,' ').trim();
    return MAKE_CODE_MAP[raw] || '';
  }
  function deriveMake(explicitMake, internalRef) {
    const directRaw = text(explicitMake);
    if (directRaw) return canonicalMake(directRaw) || directRaw.toUpperCase();
    const ref = text(internalRef).toUpperCase().trim();
    if (!ref) return 'UNKNOWN';
    const firstToken = ref.split(/\s+/)[0].replace(/^[-_]+|[-_]+$/g,'');
    if (MAKE_CODE_MAP[firstToken]) return MAKE_CODE_MAP[firstToken];
    // Small fallback for legacy references where VW/BENZ/BMW were attached directly to the number.
    if (/^VW(?=\d)/.test(firstToken)) return 'VW';
    if (/^BENZ(?=\d)/.test(firstToken)) return 'BENZ';
    if (/^BMW(?=\d)/.test(firstToken)) return 'BMW';
    return 'UNKNOWN';
  }

  function parseDate(v) {
    if (!v) return null;
    if (v instanceof Date && !isNaN(v)) return v;
    if (typeof v === 'number' && global.XLSX && global.XLSX.SSF) {
      const p = global.XLSX.SSF.parse_date_code(v);
      if (p) return new Date(p.y, p.m - 1, p.d);
    }
    const s = text(v);
    const iso = s.match(/(20\d{2})[-\/]([01]?\d)[-\/]([0-3]?\d)/);
    if (iso) return new Date(+iso[1], +iso[2] - 1, +iso[3]);
    const d = new Date(s);
    return isNaN(d) ? null : d;
  }
  function daysBetween(a, b) { return (!a || !b) ? null : Math.floor((b.getTime() - a.getTime()) / 86400000); }
  function monthsBetween(a, b) { const d = daysBetween(a, b); return d === null ? null : Math.max(0, d / 30.4375); }

  function headerIndex(headerRow) {
    const normalized = headerRow.map(normalizeKey), map = {};
    Object.entries(HEADER_ALIASES).forEach(([field, aliases]) => {
      const aliasKeys = aliases.map(normalizeKey);
      map[field] = normalized.findIndex(x => aliasKeys.includes(x));
    });
    MONTH_KEYS.forEach(([headers, key]) => {
      const aliasKeys = (Array.isArray(headers) ? headers : [headers]).map(normalizeKey);
      map[key] = normalized.findIndex(x => aliasKeys.includes(x));
    });
    return map;
  }

  function findHeaderRow(matrix) {
    const max = Math.min(matrix.length, 15);
    for (let i = 0; i < max; i++) {
      const candidate = headerIndex(matrix[i] || []);
      if (candidate.internalRef >= 0 && (candidate.productId >= 0 || candidate.oem >= 0)) return i;
    }
    return 0;
  }

  function detectReportDate(matrix, headerRowIdx, map) {
    if (map.reportDate >= 0) {
      for (let i = headerRowIdx + 1; i < Math.min(matrix.length, headerRowIdx + 30); i++) {
        const d = parseDate((matrix[i] || [])[map.reportDate]); if (d) return d;
      }
    }
    for (let i = 0; i < Math.min(matrix.length, Math.max(4, headerRowIdx)); i++) {
      const s = (matrix[i] || []).map(text).join(' | ');
      const m = s.match(/(?:Data\s*As\s*Of|Report\s*Till\s*Date|As\s*Of)\s*[:|\-]?\s*(20\d{2}[-\/]\d{1,2}[-\/]\d{1,2})/i);
      if (m) { const d = parseDate(m[1]); if (d) return d; }
    }
    return new Date();
  }

  function elapsedMonths(reportDate) {
    if (!reportDate) return 12;
    const m = reportDate.getMonth();
    const days = new Date(reportDate.getFullYear(), m + 1, 0).getDate();
    return Math.max(0.05, m + Math.min(1, reportDate.getDate() / days));
  }
  function valueAt(row, idx) { return idx >= 0 ? row[idx] : ''; }
  function resolveQty(row, map, groupField, componentFields) {
    // Prefer the three company-level fields when all are present. This makes the engine
    // resilient even if a Group On Way total is accidentally wrong in a future export.
    const allComponentsPresent = componentFields.every(f => map[f] >= 0);
    if (allComponentsPresent) return componentFields.reduce((sum, f) => sum + n(valueAt(row, map[f])), 0);
    const groupIdx = map[groupField];
    if (groupIdx >= 0) return n(valueAt(row, groupIdx));
    return componentFields.reduce((sum, f) => sum + n(valueAt(row, map[f])), 0);
  }

  function rowToProduct(row, map, reportDate, sourceIndex) {
    const monthly = MONTH_KEYS.map(([, key]) => n(valueAt(row, map[key])));
    const currentMonth = reportDate ? reportDate.getMonth() : 11;
    const monthlySumToDate = monthly.slice(0, currentMonth + 1).reduce((a, b) => a + b, 0);
    const reportedTotal = n(valueAt(row, map.totalSales));
    const totalSales = reportedTotal || monthlySumToDate;
    const motorlineStock = n(valueAt(row, map.motorlineStock));
    const bavariaStock = n(valueAt(row, map.bavariaStock));
    const tibaoStock = n(valueAt(row, map.tibaoStock));
    const explicitAll = n(valueAt(row, map.allCompany));
    const allCompany = map.allCompany >= 0 ? explicitAll : (motorlineStock + bavariaStock + tibaoStock);
    const rawOnWay = resolveQty(row, map, 'groupOnWay', ['motorlineOnWay', 'bavariaOnWay', 'tibaoOnWay']);
    const rawOnWay2 = resolveQty(row, map, 'groupOnWay2', ['motorlineOnWay2', 'bavariaOnWay2', 'tibaoOnWay2']);
    // Incoming shipment quantities should never be negative. Keep the raw value for Data Quality,
    // but never let a negative export value artificially increase a purchase recommendation.
    const onWay = Math.max(0, rawOnWay);
    const onWay2 = Math.max(0, rawOnWay2);
    const internalRef = text(valueAt(row, map.internalRef));
    const oemRaw = text(valueAt(row, map.oem));
    const brand = text(valueAt(row, map.brand)) || 'Unbranded';

    return {
      _id: text(valueAt(row, map.productId)) || `${sourceIndex}-${internalRef}`,
      sourceIndex, productId: text(valueAt(row, map.productId)), internalRef,
      oem: oemRaw, oemKey: normalizeOEM(oemRaw), brand, brandKey: brand.toUpperCase(),
      brandPartNo: text(valueAt(row, map.brandPartNo)), description: text(valueAt(row, map.description)),
      category: text(valueAt(row, map.category)) || 'Uncategorized',
      make: deriveMake(valueAt(row, map.make), internalRef),
      motorlineStock, bavariaStock, tibaoStock, allCompany,
      tibaoADRef: n(valueAt(row, map.tibaoADRef)), tibaoDXBRef: n(valueAt(row, map.tibaoDXBRef)),
      monthly, totalSales, salesCount: n(valueAt(row, map.salesCount)), lastSaleDate: parseDate(valueAt(row, map.lastSaleDate)),
      totalPurchase: n(valueAt(row, map.totalPurchase)), purchaseCount: n(valueAt(row, map.purchaseCount)), lastPurchaseDate: parseDate(valueAt(row, map.lastPurchaseDate)),
      motorlineOnWay: n(valueAt(row, map.motorlineOnWay)), bavariaOnWay: n(valueAt(row, map.bavariaOnWay)), tibaoOnWay: n(valueAt(row, map.tibaoOnWay)), rawOnWay, onWay,
      motorlineOnWay2: n(valueAt(row, map.motorlineOnWay2)), bavariaOnWay2: n(valueAt(row, map.bavariaOnWay2)), tibaoOnWay2: n(valueAt(row, map.tibaoOnWay2)), rawOnWay2, onWay2,
      reportDate
    };
  }

  function projectMonthValue(v, monthIndex, reportDate, normalizeCurrentMonth) {
    if (!normalizeCurrentMonth || monthIndex !== reportDate.getMonth() || v <= 0) return v;
    const days = new Date(reportDate.getFullYear(), monthIndex + 1, 0).getDate();
    const fraction = Math.max(0.25, Math.min(1, reportDate.getDate() / days));
    return v / fraction;
  }
  function windowAverage(product, reportDate, back, normalizeCurrentMonth) {
    const end = reportDate.getMonth(), start = Math.max(0, end - back + 1), vals = [];
    for (let m = start; m <= end; m++) vals.push(projectMonthValue(product.monthly[m] || 0, m, reportDate, normalizeCurrentMonth));
    return vals.length ? vals.reduce((a,b)=>a+b,0) / vals.length : 0;
  }
  function priorWindowAverage(product, reportDate, recentBack) {
    const end = reportDate.getMonth() - recentBack, start = Math.max(0, end - recentBack + 1), vals=[];
    for (let m=start; m<=end; m++) if (m>=0) vals.push(product.monthly[m] || 0);
    return vals.length ? vals.reduce((a,b)=>a+b,0)/vals.length : 0;
  }

  function consecutiveZeroSalesMonths(product, reportDate) {
    const end = reportDate.getMonth();
    let count = 0;
    for (let m=end; m>=0; m--) {
      if ((product.monthly[m] || 0) > 0) break;
      count++;
    }
    return count;
  }

  function inferredLastSaleAgeMonths(product, reportDate) {
    for (let m=reportDate.getMonth(); m>=0; m--) {
      if ((product.monthly[m] || 0) > 0) {
        const monthEnd = new Date(reportDate.getFullYear(), m + 1, 0);
        return monthsBetween(monthEnd, reportDate);
      }
    }
    return null;
  }

  function salesProfile(product, settings, reportDate) {
    const elapsed = elapsedMonths(reportDate);
    const monthCount = reportDate.getMonth() + 1;
    const raw = product.monthly.slice(0, monthCount);
    const activeMonths = raw.filter(v => v > 0).length;
    const activePct = monthCount ? (activeMonths / monthCount) * 100 : 0;
    const totalFromMonths = raw.reduce((a,b)=>a+b,0);
    const total = product.totalSales || totalFromMonths;
    const overall = total / elapsed;
    const recent3 = windowAverage(product, reportDate, 3, settings.currentMonthNormalize);
    const recent6 = windowAverage(product, reportDate, 6, settings.currentMonthNormalize);
    const prior3 = priorWindowAverage(product, reportDate, 3);
    const maxMonth = raw.length ? Math.max(...raw) : 0;
    const maxMonthIndex = raw.indexOf(maxMonth);
    const spikeSharePct = total > 0 ? (maxMonth / total) * 100 : 0;
    const explicitLastSaleAgeMonths = product.lastSaleDate ? monthsBetween(product.lastSaleDate, reportDate) : null;
    const inferredSaleAge = explicitLastSaleAgeMonths === null && total > 0 ? inferredLastSaleAgeMonths(product, reportDate) : null;
    const lastSaleAgeMonths = explicitLastSaleAgeMonths !== null ? explicitLastSaleAgeMonths : inferredSaleAge;
    const lastPurchaseAgeForInactivity = product.lastPurchaseDate ? monthsBetween(product.lastPurchaseDate, reportDate) : null;
    const zeroSalesStreakMonths = consecutiveZeroSalesMonths(product, reportDate);

    // Dead-stock logic must also work for items that have NEVER sold.
    // If an item never sold, Last Sale Date is naturally blank, so we use the last purchase age when available.
    // If neither sale nor purchase date exists, the consecutive zero-sales calendar months in the report are the fallback.
    let inactivityAgeMonths = lastSaleAgeMonths;
    let inactivityBasis = lastSaleAgeMonths !== null ? (product.lastSaleDate ? 'Last sale date' : 'Last non-zero sales month') : '';
    if (total <= 0) {
      if (lastPurchaseAgeForInactivity !== null) {
        inactivityAgeMonths = lastPurchaseAgeForInactivity;
        inactivityBasis = 'Last purchase date (never sold)';
      } else {
        inactivityAgeMonths = zeroSalesStreakMonths;
        inactivityBasis = 'Consecutive zero-sales months';
      }
    }

    let pattern = 'NO SALES';
    const isDead = product.allCompany > 0 && inactivityAgeMonths !== null && inactivityAgeMonths >= settings.deadStockMonths;
    const isDormant = product.allCompany > 0 && !isDead && inactivityAgeMonths !== null && inactivityAgeMonths >= settings.dormantMonths;
    const isSpike = total > 0 && activeMonths <= settings.spikeMaxActiveMonths && spikeSharePct >= settings.spikeDominancePct &&
      (lastSaleAgeMonths === null || lastSaleAgeMonths >= 1.5);

    if (isDead) pattern = 'DEAD STOCK';
    else if (isSpike) pattern = 'ONE-TIME SPIKE';
    else if (isDormant) pattern = 'DORMANT';
    else if (total <= 0) pattern = 'NO SALES';
    else if (activePct >= settings.regularActivePct) {
      if (prior3 > 0 && recent3 >= prior3 * 1.25) pattern = 'RISING';
      else if (prior3 > 0 && recent3 <= prior3 * 0.75) pattern = 'FALLING';
      else pattern = 'REGULAR';
    } else pattern = 'INTERMITTENT';

    let confidence = 'REVIEW';
    if (['DEAD STOCK','DORMANT','ONE-TIME SPIKE','NO SALES'].includes(pattern)) confidence = 'REVIEW';
    else if (activePct >= 70 && product.salesCount >= 4) confidence = 'HIGH';
    else if (activeMonths >= 3 || product.salesCount >= 3) confidence = 'MEDIUM';

    let demand = 0;
    if (settings.demandMethod === 'overall') demand = overall;
    else if (settings.demandMethod === 'recent') demand = recent3;
    else if (settings.demandMethod === 'max') demand = Math.max(overall, recent3);
    else {
      if (pattern === 'RISING') demand = (recent3 * 0.65) + (recent6 * 0.25) + (overall * 0.10);
      else if (pattern === 'FALLING') demand = (recent3 * 0.50) + (recent6 * 0.30) + (overall * 0.20);
      else if (pattern === 'REGULAR') demand = (recent3 * 0.45) + (recent6 * 0.30) + (overall * 0.25);
      else if (pattern === 'INTERMITTENT') demand = (overall * 0.60) + (recent6 * 0.40);
      else demand = 0; // conservative guard for dormant / dead / one-time spike / no sales
    }

    return {
      elapsed, monthCount, total, overall, recent3, recent6, prior3, activeMonths, activePct,
      maxMonth, maxMonthIndex, spikeSharePct, lastSaleAgeMonths, zeroSalesStreakMonths,
      inactivityAgeMonths, inactivityBasis, pattern, confidence,
      demand: Math.max(0, demand), isDead, isDormant, isSpike
    };
  }

  function purchaseProfile(product, reportDate, salesProfileObj) {
    const sales = salesProfileObj.total;
    const purchases = product.totalPurchase;
    const ratio = sales > 0 ? purchases / sales : (purchases > 0 ? Infinity : 0);
    const lastPurchaseAgeMonths = product.lastPurchaseDate ? monthsBetween(product.lastPurchaseDate, reportDate) : null;
    let signal = 'NO PURCHASE HISTORY';
    if (purchases > 0 && sales <= 0) signal = 'PURCHASED / NO SALES';
    else if (ratio >= 2.0) signal = 'OVERBUYING RISK';
    else if (ratio >= 1.25) signal = 'PURCHASE ABOVE SALES';
    else if (ratio > 0 && ratio < 0.75) signal = 'SALES ABOVE PURCHASE';
    else if (purchases > 0) signal = 'BALANCED';
    return { purchaseSalesRatio: ratio, lastPurchaseAgeMonths, purchaseSignal: signal };
  }

  function formatNumber(x, digits = 1) {
    if (!Number.isFinite(x)) return '—';
    return x.toLocaleString(undefined, { maximumFractionDigits: digits });
  }

  function buildFamilies(products) {
    const families = new Map();
    products.forEach(p => { if (p.oemKey) { if (!families.has(p.oemKey)) families.set(p.oemKey, []); families.get(p.oemKey).push(p); } });
    return families;
  }

  function calculate(products, settingsInput, reportDate) {
    const settings = Object.assign({}, DEFAULT_SETTINGS, settingsInput || {});
    const families = buildFamilies(products);

    return products.map(p => {
      const sp = salesProfile(p, settings, reportDate);
      const pp = purchaseProfile(p, reportDate, sp);
      const family = p.oemKey ? (families.get(p.oemKey) || []) : [];
      const otherBrandRows = family.filter(x => x.brandKey !== p.brandKey);
      const otherByBrand = new Map();
      otherBrandRows.forEach(x => {
        if (!otherByBrand.has(x.brandKey)) otherByBrand.set(x.brandKey, { brand:x.brand, stock:0, onWay:0, onWay2:0, sales:0, partNumbers:new Set() });
        const b = otherByBrand.get(x.brandKey);
        b.stock += x.allCompany; b.onWay += x.onWay; b.onWay2 += x.onWay2; b.sales += x.totalSales;
        if (x.brandPartNo) b.partNumbers.add(x.brandPartNo);
      });
      const alternatives = Array.from(otherByBrand.values()).map(b => ({
        brand:b.brand, stock:b.stock, onWay:b.onWay, onWay2:b.onWay2, sales:b.sales, partNumbers:Array.from(b.partNumbers)
      })).sort((a,b)=>(b.stock+b.onWay+b.onWay2)-(a.stock+a.onWay+a.onWay2));

      const otherStock = alternatives.reduce((s,x)=>s+x.stock,0);
      const otherOnWay = alternatives.reduce((s,x)=>s+x.onWay,0);
      const otherOnWay2 = alternatives.reduce((s,x)=>s+x.onWay2,0);
      const otherSupply = otherStock + otherOnWay + otherOnWay2;
      const equivalentCredit = otherSupply * Math.max(0,Math.min(100,settings.equivalentCreditPct))/100;

      const directSupply = p.allCompany + p.onWay + p.onWay2;
      const effectiveSupply = directSupply + equivalentCredit;
      const demand = sp.demand;
      const currentCover = demand > 0 ? p.allCompany/demand : (p.allCompany>0?Infinity:0);
      const pipelineCover = demand > 0 ? directSupply/demand : (directSupply>0?Infinity:0);
      const effectiveCover = demand > 0 ? effectiveSupply/demand : (effectiveSupply>0?Infinity:0);
      const targetMonths = settings.targetCover + settings.safetyCover;
      const targetQty = demand * targetMonths;
      const targetGapQty = Math.max(0, Math.ceil(targetQty - effectiveSupply));

      let movement = 'NO SALES';
      if (sp.isDead) movement = 'DEAD STOCK';
      else if (demand >= settings.fastRate) movement = 'FAST MOVING';
      else if (demand >= settings.mediumRate) movement = 'MEDIUM MOVING';
      else if (demand > 0) movement = 'SLOW MOVING';

      let condition = 'OK', suggested = 0, priority = 4, action = 'NO ACTION';
      if (sp.isDead) condition = 'DEAD STOCK';
      else if (sp.isSpike) condition = 'ONE-TIME / REVIEW';
      else if (sp.isDormant) condition = 'DORMANT / REVIEW';
      else if (sp.total <= 0) condition = p.allCompany > 0 ? 'NO SALES / REVIEW' : 'NO DEMAND';
      else if (currentCover >= settings.overstockCover) condition = 'OVERSTOCK';
      else if (demand <= 0) condition = 'NO SALES / REVIEW';
      else if (effectiveCover < settings.criticalCover) condition = 'CRITICAL ORDER';
      else if (settings.suggestionMode === 'trigger') {
        if (effectiveCover < settings.reorderCover) condition = 'REORDER SOON';
        else if (currentCover < settings.reorderCover && (p.onWay + p.onWay2 + equivalentCredit) > 0) condition = 'WAIT INCOMING';
        else condition = 'OK';
      } else {
        // Proactive mode: recommend a top-up whenever projected cover is below the normal target.
        if (currentCover < settings.targetCover && effectiveCover >= settings.targetCover && (p.onWay + p.onWay2 + equivalentCredit) > 0) condition = 'WAIT INCOMING';
        else if (effectiveCover < settings.targetCover) condition = 'REORDER SOON';
        else condition = 'OK';
      }

      if (condition === 'CRITICAL ORDER' || condition === 'REORDER SOON') suggested = targetGapQty;
      if (condition === 'CRITICAL ORDER') { priority=1; action='ORDER NOW'; }
      else if (condition === 'REORDER SOON') { priority=2; action='ADD TO NEXT PURCHASE'; }
      else if (condition === 'WAIT INCOMING') { priority=3; action=p.onWay>0?'WAIT / EXPEDITE ON WAY':'FOLLOW UP ON WAY 2'; }
      else if (condition === 'OVERSTOCK') { priority=5; action='DO NOT ORDER'; }
      else if (condition === 'DEAD STOCK') { priority=5; action='STOP BUYING / CLEAR STOCK'; }
      else if (condition === 'DORMANT / REVIEW') { priority=5; action='HOLD / REVIEW DEMAND'; }
      else if (condition === 'ONE-TIME / REVIEW') { priority=5; action='MANUAL REVIEW - POSSIBLE ONE-OFF SALE'; }
      else if (condition === 'NO SALES / REVIEW') { priority=5; action='REVIEW ITEM'; }
      else if (condition === 'NO DEMAND') { priority=5; action='NO PURCHASE'; }

      const equivalentNote = alternatives.length
        ? `${alternatives.length} other brand${alternatives.length>1?'s':''} / ${formatNumber(otherStock,0)} stock / ${formatNumber(otherOnWay+otherOnWay2,0)} incoming`
        : 'No other brand match';

      const reasonParts = [
        `${sp.pattern} demand`, `${sp.activeMonths}/${sp.monthCount} sales months`, `smart demand ${formatNumber(demand,1)}/mo`,
        `stock ${formatNumber(p.allCompany,0)}`
      ];
      if (sp.isDead || sp.isDormant || sp.total <= 0) reasonParts.push(`zero-sales streak ${sp.zeroSalesStreakMonths} month${sp.zeroSalesStreakMonths===1?'':'s'}`);
      if ((sp.isDead || sp.isDormant) && sp.inactivityBasis) reasonParts.push(`inactivity based on ${sp.inactivityBasis.toLowerCase()}`);
      if (p.onWay) reasonParts.push(`On Way ${formatNumber(p.onWay,0)}`);
      if (p.onWay2) reasonParts.push(`On Way 2 ${formatNumber(p.onWay2,0)}`);
      if (p.totalPurchase) reasonParts.push(`purchased ${formatNumber(p.totalPurchase,0)} in report`);
      if (pp.purchaseSignal !== 'BALANCED' && pp.purchaseSignal !== 'NO PURCHASE HISTORY') reasonParts.push(pp.purchaseSignal.toLowerCase());
      if (alternatives.length) reasonParts.push(`same OEM other brands ${formatNumber(otherStock,0)} stock`);
      if (settings.equivalentCreditPct > 0 && alternatives.length) reasonParts.push(`${settings.equivalentCreditPct}% equivalent credit used`);
      if (suggested > 0) reasonParts.push(`replenish toward ${targetMonths.toFixed(1)} months`);

      return Object.assign({}, p, sp, pp, {
        avgMonthlySales:sp.overall, recent3Avg:sp.recent3, recent6Avg:sp.recent6, demandRate:demand,
        demandPattern:sp.pattern, demandConfidence:sp.confidence, movement,
        currentCover,pipelineCover,effectiveCover,directSupply,effectiveSupply,
        alternatives,otherStock,otherOnWay,otherOnWay2,otherSupply,equivalentCredit,equivalentNote,
        condition,priority,targetQty,targetGapQty,suggestedQty:suggested,action,reason:reasonParts.join(' • ')
      });
    });
  }

  function parseMatrix(matrix) {
    if (!Array.isArray(matrix) || !matrix.length) throw new Error('The workbook appears to be empty.');
    const headerRowIdx = findHeaderRow(matrix), headerRow = matrix[headerRowIdx] || [], map = headerIndex(headerRow);
    if (map.internalRef < 0) throw new Error('Could not find the Internal Reference column. Please use the Odoo Purchase Support export.');
    const reportDate = detectReportDate(matrix, headerRowIdx, map), products=[];
    for (let i=headerRowIdx+1;i<matrix.length;i++) {
      const row=matrix[i]||[]; if (!row.some(v=>text(v)!=='')) continue;
      const p=rowToProduct(row,map,reportDate,i+1);
      if (!p.internalRef&&!p.productId&&!p.oem&&!p.description) continue;
      products.push(p);
    }
    return {products,reportDate,headerRow,map,headerRowIdx};
  }

  function dataQuality(products, map) {
    const missing={
      oem:products.filter(p=>!p.oemKey).length,
      brand:products.filter(p=>!p.brand||p.brand==='Unbranded').length,
      description:products.filter(p=>!p.description).length,
      internalRef:products.filter(p=>!p.internalRef).length,
      lastSaleDate:products.filter(p=>p.totalSales>0&&!p.lastSaleDate).length,
      negativeIncoming:products.filter(p=>p.rawOnWay<0||p.rawOnWay2<0||p.motorlineOnWay<0||p.bavariaOnWay<0||p.tibaoOnWay<0||p.motorlineOnWay2<0||p.bavariaOnWay2<0||p.tibaoOnWay2<0).length
    };
    const columns={
      make:map.make>=0,
      motorlineOnWay:map.motorlineOnWay>=0,bavariaOnWay:map.bavariaOnWay>=0,tibaoOnWay:map.tibaoOnWay>=0,groupOnWay:map.groupOnWay>=0,
      motorlineOnWay2:map.motorlineOnWay2>=0,bavariaOnWay2:map.bavariaOnWay2>=0,tibaoOnWay2:map.tibaoOnWay2>=0,groupOnWay2:map.groupOnWay2>=0,
      reportDate:map.reportDate>=0,totalPurchase:map.totalPurchase>=0,purchaseCount:map.purchaseCount>=0,lastPurchaseDate:map.lastPurchaseDate>=0,
      monthlySales:map.jan>=0&&map.aug>=0
    };
    return {missing,columns};
  }

  global.PurchaseEngine={DEFAULT_SETTINGS,MONTH_KEYS,CONDITION_OPTIONS,MOVEMENT_OPTIONS,PATTERN_OPTIONS,parseMatrix,calculate,dataQuality,normalizeOEM,parseDate,elapsedMonths,deriveMake};
})(window);
