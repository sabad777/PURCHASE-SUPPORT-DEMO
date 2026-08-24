/* TiBAO Purchase Intelligence - calculation engine
 * Pure browser-side logic. No server calls and no data persistence.
 */
(function (global) {
  'use strict';

  const MONTH_KEYS = [
    ['Jan Ext Sales', 'jan'], ['Feb Ext Sales', 'feb'], ['Mar Ext Sales', 'mar'],
    ['Apr Ext Sales', 'apr'], ['May Ext Sales', 'may'], ['Jun Ext Sales', 'jun'],
    ['Jul Ext Sales', 'jul'], ['Aug Ext Sales', 'aug'], ['Sep Ext Sales', 'sep'],
    ['Oct Ext Sales', 'oct'], ['Nov Ext Sales', 'nov'], ['Dec Ext Sales', 'dec']
  ];

  const DEFAULT_SETTINGS = {
    targetCover: 3.0,
    safetyCover: 0.5,
    criticalCover: 0.5,
    reorderCover: 1.5,
    overstockCover: 6.0,
    deadStockMonths: 12,
    fastRate: 5.0,
    mediumRate: 1.5,
    equivalentCreditPct: 0,
    demandMethod: 'weighted',
    currentMonthNormalize: true
  };

  const HEADER_ALIASES = {
    productId: ['Odoo Product ID', 'Product ID'],
    internalRef: ['Internal Reference', 'Item Code', 'Default Code'],
    oem: ['Main OEM No Space', 'Main OEM', 'OEM', 'OEM No Space'],
    brand: ['Brand'],
    brandPartNo: ['Brand Part No.', 'Brand Part No', 'Brand Part Number'],
    description: ['Add Description', 'Description'],
    category: ['Product Category', 'Category'],
    motorlineStock: ['Motorline Stock'],
    bavariaStock: ['Bavaria Stock'],
    tibaoStock: ['Tibao SHJ Stock', 'Tibao Stock'],
    allCompany: ['All Company Qty (Core)', 'All Company Qty', 'All Company Quantity'],
    tibaoADRef: ['Tibao AD Stock (Ref)', 'Tibao AD Stock'],
    tibaoDXBRef: ['Tibao DXB Stock (Ref)', 'Tibao DXB Stock'],
    totalSales: ['Total External Sales Qty', 'Total External Sales', 'Sales Qty'],
    salesCount: ['External Sales Invoice Count', 'Sales Invoice Count', 'Sales Count'],
    lastSaleDate: ['Last External Sale Date', 'Last Sale Date'],
    totalPurchase: ['Total External Purchase Qty', 'Total Purchase Qty', 'Purchase Qty'],
    purchaseCount: ['External Purchase Count', 'Purchase Count'],
    lastPurchaseDate: ['Last External Purchase Date', 'Last Purchase Date'],
    motorlineOnWay: ['Motorline On Way'],
    bavariaOnWay: ['Bavaria On Way'],
    tibaoOnWay: ['Tibao On Way', 'Tibao SHJ On Way'],
    groupOnWay: ['Group On Way', 'All Company On Way', 'Total On Way'],
    motorlineOnWay2: ['Motorline On Way 2'],
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

  function text(v) {
    return v === null || v === undefined ? '' : String(v).trim();
  }

  function normalizeKey(v) {
    return text(v).toLowerCase().replace(/[\s._()\-\/]+/g, '');
  }

  function normalizeOEM(v) {
    return text(v).toUpperCase().replace(/[^A-Z0-9]/g, '');
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

  function daysBetween(a, b) {
    if (!a || !b) return null;
    return Math.floor((b.getTime() - a.getTime()) / 86400000);
  }

  function monthsBetween(a, b) {
    const d = daysBetween(a, b);
    return d === null ? null : d / 30.4375;
  }

  function headerIndex(headerRow) {
    const normalized = headerRow.map(normalizeKey);
    const map = {};
    Object.entries(HEADER_ALIASES).forEach(([field, aliases]) => {
      const aliasKeys = aliases.map(normalizeKey);
      const idx = normalized.findIndex(x => aliasKeys.includes(x));
      map[field] = idx;
    });
    MONTH_KEYS.forEach(([header, key]) => {
      const idx = normalized.findIndex(x => x === normalizeKey(header));
      map[key] = idx;
    });
    return map;
  }

  function findHeaderRow(matrix) {
    const max = Math.min(matrix.length, 15);
    for (let i = 0; i < max; i++) {
      const row = (matrix[i] || []).map(normalizeKey);
      if (row.includes(normalizeKey('Internal Reference')) &&
          (row.includes(normalizeKey('Odoo Product ID')) || row.includes(normalizeKey('Main OEM No Space')))) {
        return i;
      }
    }
    return 0;
  }

  function detectReportDate(matrix, headerRowIdx, map) {
    // 1) explicit report-date column
    if (map.reportDate >= 0) {
      for (let i = headerRowIdx + 1; i < Math.min(matrix.length, headerRowIdx + 30); i++) {
        const d = parseDate((matrix[i] || [])[map.reportDate]);
        if (d) return d;
      }
    }
    // 2) title / first rows
    for (let i = 0; i < Math.min(matrix.length, Math.max(4, headerRowIdx)); i++) {
      const s = (matrix[i] || []).map(text).join(' | ');
      const m = s.match(/(?:Data\s*As\s*Of|Report\s*Till\s*Date|As\s*Of)\s*[:|\-]?\s*(20\d{2}[-\/]\d{1,2}[-\/]\d{1,2})/i);
      if (m) {
        const d = parseDate(m[1]);
        if (d) return d;
      }
    }
    return new Date();
  }

  function elapsedMonths(reportDate) {
    if (!reportDate) return 12;
    const m = reportDate.getMonth();
    const days = new Date(reportDate.getFullYear(), m + 1, 0).getDate();
    return Math.max(0.05, m + Math.min(1, reportDate.getDate() / days));
  }

  function valueAt(row, idx) {
    return idx >= 0 ? row[idx] : '';
  }

  function resolveQty(row, map, groupField, componentFields) {
    const groupIdx = map[groupField];
    if (groupIdx >= 0) return n(valueAt(row, groupIdx));
    return componentFields.reduce((sum, f) => sum + n(valueAt(row, map[f])), 0);
  }

  function rowToProduct(row, map, reportDate, sourceIndex) {
    const monthly = MONTH_KEYS.map(([, key]) => n(valueAt(row, map[key])));
    const monthlySum = monthly.reduce((a, b) => a + b, 0);
    const reportedTotal = n(valueAt(row, map.totalSales));
    const totalSales = reportedTotal || monthlySum;

    const motorlineStock = n(valueAt(row, map.motorlineStock));
    const bavariaStock = n(valueAt(row, map.bavariaStock));
    const tibaoStock = n(valueAt(row, map.tibaoStock));
    const explicitAll = n(valueAt(row, map.allCompany));
    const allCompany = map.allCompany >= 0 ? explicitAll : (motorlineStock + bavariaStock + tibaoStock);

    const onWay = resolveQty(row, map, 'groupOnWay', ['motorlineOnWay', 'bavariaOnWay', 'tibaoOnWay']);
    const onWay2 = resolveQty(row, map, 'groupOnWay2', ['motorlineOnWay2', 'bavariaOnWay2', 'tibaoOnWay2']);

    const internalRef = text(valueAt(row, map.internalRef));
    const oemRaw = text(valueAt(row, map.oem));
    const brand = text(valueAt(row, map.brand)) || 'Unbranded';

    return {
      _id: text(valueAt(row, map.productId)) || `${sourceIndex}-${internalRef}`,
      sourceIndex,
      productId: text(valueAt(row, map.productId)),
      internalRef,
      oem: oemRaw,
      oemKey: normalizeOEM(oemRaw),
      brand,
      brandKey: brand.toUpperCase(),
      brandPartNo: text(valueAt(row, map.brandPartNo)),
      description: text(valueAt(row, map.description)),
      category: text(valueAt(row, map.category)) || 'Uncategorized',
      motorlineStock,
      bavariaStock,
      tibaoStock,
      allCompany,
      tibaoADRef: n(valueAt(row, map.tibaoADRef)),
      tibaoDXBRef: n(valueAt(row, map.tibaoDXBRef)),
      monthly,
      totalSales,
      salesCount: n(valueAt(row, map.salesCount)),
      lastSaleDate: parseDate(valueAt(row, map.lastSaleDate)),
      totalPurchase: n(valueAt(row, map.totalPurchase)),
      purchaseCount: n(valueAt(row, map.purchaseCount)),
      lastPurchaseDate: parseDate(valueAt(row, map.lastPurchaseDate)),
      motorlineOnWay: n(valueAt(row, map.motorlineOnWay)),
      bavariaOnWay: n(valueAt(row, map.bavariaOnWay)),
      tibaoOnWay: n(valueAt(row, map.tibaoOnWay)),
      onWay,
      motorlineOnWay2: n(valueAt(row, map.motorlineOnWay2)),
      bavariaOnWay2: n(valueAt(row, map.bavariaOnWay2)),
      tibaoOnWay2: n(valueAt(row, map.tibaoOnWay2)),
      onWay2,
      reportDate
    };
  }

  function recent3Average(product, reportDate, normalizeCurrentMonth) {
    const currentMonth = reportDate.getMonth();
    const start = Math.max(0, currentMonth - 2);
    const values = [];
    for (let m = start; m <= currentMonth; m++) {
      let v = product.monthly[m] || 0;
      if (m === currentMonth && normalizeCurrentMonth) {
        const days = new Date(reportDate.getFullYear(), m + 1, 0).getDate();
        const fraction = Math.max(0.15, Math.min(1, reportDate.getDate() / days));
        v = v / fraction;
      }
      values.push(v);
    }
    if (!values.length) return 0;
    return values.reduce((a, b) => a + b, 0) / values.length;
  }

  function demandRate(product, settings, reportDate) {
    const elapsed = elapsedMonths(reportDate);
    const overall = product.totalSales / elapsed;
    const recent = recent3Average(product, reportDate, settings.currentMonthNormalize);
    let demand;
    if (settings.demandMethod === 'overall') demand = overall;
    else if (settings.demandMethod === 'max') demand = Math.max(overall, recent);
    else demand = (overall * 0.4) + (recent * 0.6);
    return { overall, recent, demand: Math.max(0, demand) };
  }

  function formatNumber(x, digits = 1) {
    if (!Number.isFinite(x)) return '—';
    return x.toLocaleString(undefined, { maximumFractionDigits: digits });
  }

  function buildFamilies(products) {
    const families = new Map();
    products.forEach(p => {
      if (!p.oemKey) return;
      if (!families.has(p.oemKey)) families.set(p.oemKey, []);
      families.get(p.oemKey).push(p);
    });
    return families;
  }

  function calculate(products, settingsInput, reportDate) {
    const settings = Object.assign({}, DEFAULT_SETTINGS, settingsInput || {});
    const families = buildFamilies(products);

    return products.map(p => {
      const rates = demandRate(p, settings, reportDate);
      const family = p.oemKey ? (families.get(p.oemKey) || []) : [];
      const otherBrandRows = family.filter(x => x.brandKey !== p.brandKey);
      const otherByBrand = new Map();
      otherBrandRows.forEach(x => {
        if (!otherByBrand.has(x.brandKey)) {
          otherByBrand.set(x.brandKey, { brand: x.brand, stock: 0, onWay: 0, onWay2: 0, sales: 0, partNumbers: new Set() });
        }
        const b = otherByBrand.get(x.brandKey);
        b.stock += x.allCompany;
        b.onWay += x.onWay;
        b.onWay2 += x.onWay2;
        b.sales += x.totalSales;
        if (x.brandPartNo) b.partNumbers.add(x.brandPartNo);
      });
      const alternatives = Array.from(otherByBrand.values()).map(b => ({
        brand: b.brand,
        stock: b.stock,
        onWay: b.onWay,
        onWay2: b.onWay2,
        sales: b.sales,
        partNumbers: Array.from(b.partNumbers)
      })).sort((a,b) => (b.stock + b.onWay + b.onWay2) - (a.stock + a.onWay + a.onWay2));

      const otherStock = alternatives.reduce((s, x) => s + x.stock, 0);
      const otherOnWay = alternatives.reduce((s, x) => s + x.onWay, 0);
      const otherOnWay2 = alternatives.reduce((s, x) => s + x.onWay2, 0);
      const otherSupply = otherStock + otherOnWay + otherOnWay2;
      const equivalentCredit = otherSupply * Math.max(0, Math.min(100, settings.equivalentCreditPct)) / 100;

      const directSupply = p.allCompany + p.onWay + p.onWay2;
      const effectiveSupply = directSupply + equivalentCredit;
      const currentCover = rates.demand > 0 ? p.allCompany / rates.demand : (p.allCompany > 0 ? Infinity : 0);
      const pipelineCover = rates.demand > 0 ? directSupply / rates.demand : (directSupply > 0 ? Infinity : 0);
      const effectiveCover = rates.demand > 0 ? effectiveSupply / rates.demand : (effectiveSupply > 0 ? Infinity : 0);
      const lastSaleAgeMonths = p.lastSaleDate ? monthsBetween(p.lastSaleDate, reportDate) : null;

      let movement = 'NO SALES';
      if (rates.demand >= settings.fastRate) movement = 'FAST MOVING';
      else if (rates.demand >= settings.mediumRate) movement = 'MEDIUM MOVING';
      else if (rates.demand > 0) movement = 'SLOW MOVING';

      const isDead = p.allCompany > 0 && p.lastSaleDate && lastSaleAgeMonths >= settings.deadStockMonths && p.totalSales === 0;
      if (isDead) movement = 'DEAD STOCK';

      let condition = 'OK';
      if (isDead) condition = 'DEAD STOCK';
      else if (rates.demand <= 0) condition = p.allCompany > 0 ? 'NO SALES / REVIEW' : 'NO DEMAND';
      else if (currentCover >= settings.overstockCover) condition = 'OVERSTOCK';
      else if (effectiveCover < settings.criticalCover) condition = 'CRITICAL ORDER';
      else if (currentCover < settings.criticalCover && pipelineCover >= settings.criticalCover) condition = 'WAIT INCOMING';
      else if (effectiveCover < settings.reorderCover) condition = 'REORDER SOON';

      const targetQty = rates.demand * (settings.targetCover + settings.safetyCover);
      let suggested = Math.max(0, Math.ceil(targetQty - effectiveSupply));
      if (['DEAD STOCK','OVERSTOCK','NO SALES / REVIEW','NO DEMAND'].includes(condition)) suggested = 0;

      let priority = 4;
      if (condition === 'CRITICAL ORDER') priority = 1;
      else if (condition === 'REORDER SOON') priority = 2;
      else if (condition === 'WAIT INCOMING') priority = 3;
      else if (['DEAD STOCK','OVERSTOCK','NO SALES / REVIEW','NO DEMAND'].includes(condition)) priority = 5;

      let action = 'NO ACTION';
      if (condition === 'CRITICAL ORDER') action = 'ORDER NOW';
      else if (condition === 'REORDER SOON') action = 'ADD TO NEXT PURCHASE';
      else if (condition === 'WAIT INCOMING') action = 'WAIT / EXPEDITE INCOMING';
      else if (condition === 'OVERSTOCK') action = 'DO NOT ORDER';
      else if (condition === 'DEAD STOCK') action = 'STOP BUYING / CLEAR STOCK';
      else if (condition === 'NO SALES / REVIEW') action = 'REVIEW ITEM';
      else if (condition === 'NO DEMAND') action = 'NO PURCHASE';

      const equivalentNote = alternatives.length
        ? `${alternatives.length} other brand${alternatives.length > 1 ? 's' : ''} / ${formatNumber(otherStock,0)} stock / ${formatNumber(otherOnWay + otherOnWay2,0)} incoming`
        : 'No other brand match';

      const reasonParts = [];
      reasonParts.push(`Demand ${formatNumber(rates.demand,1)}/month`);
      reasonParts.push(`company stock ${formatNumber(p.allCompany,0)}`);
      if (p.onWay) reasonParts.push(`On Way ${formatNumber(p.onWay,0)}`);
      if (p.onWay2) reasonParts.push(`On Way 2 ${formatNumber(p.onWay2,0)}`);
      if (alternatives.length) reasonParts.push(`same OEM other brands ${formatNumber(otherStock,0)} stock`);
      if (settings.equivalentCreditPct > 0 && alternatives.length) reasonParts.push(`${settings.equivalentCreditPct}% equivalent credit used`);
      if (suggested > 0) reasonParts.push(`target ${(settings.targetCover + settings.safetyCover).toFixed(1)} months`);

      return Object.assign({}, p, {
        avgMonthlySales: rates.overall,
        recent3Avg: rates.recent,
        demandRate: rates.demand,
        movement,
        currentCover,
        pipelineCover,
        effectiveCover,
        directSupply,
        effectiveSupply,
        equivalentCredit,
        alternatives,
        otherStock,
        otherOnWay,
        otherOnWay2,
        otherSupply,
        equivalentNote,
        condition,
        priority,
        targetQty,
        suggestedQty: suggested,
        action,
        lastSaleAgeMonths,
        reason: reasonParts.join(' • ')
      });
    });
  }

  function parseMatrix(matrix) {
    if (!Array.isArray(matrix) || !matrix.length) throw new Error('The workbook appears to be empty.');
    const headerRowIdx = findHeaderRow(matrix);
    const headerRow = matrix[headerRowIdx] || [];
    const map = headerIndex(headerRow);
    if (map.internalRef < 0) throw new Error('Could not find the Internal Reference column. Please use the Odoo Purchase Support export.');
    const reportDate = detectReportDate(matrix, headerRowIdx, map);
    const products = [];
    for (let i = headerRowIdx + 1; i < matrix.length; i++) {
      const row = matrix[i] || [];
      if (!row.some(v => text(v) !== '')) continue;
      const p = rowToProduct(row, map, reportDate, i + 1);
      if (!p.internalRef && !p.productId && !p.oem && !p.description) continue;
      products.push(p);
    }
    return { products, reportDate, headerRow, map, headerRowIdx };
  }

  function dataQuality(products, map) {
    const missing = {
      oem: products.filter(p => !p.oemKey).length,
      brand: products.filter(p => !p.brand || p.brand === 'Unbranded').length,
      description: products.filter(p => !p.description).length,
      internalRef: products.filter(p => !p.internalRef).length
    };
    const columns = {
      bavariaOnWay: map.bavariaOnWay >= 0,
      tibaoOnWay: map.tibaoOnWay >= 0,
      reportDate: map.reportDate >= 0,
      groupOnWay: map.groupOnWay >= 0,
      groupOnWay2: map.groupOnWay2 >= 0
    };
    return { missing, columns };
  }

  global.PurchaseEngine = {
    DEFAULT_SETTINGS,
    MONTH_KEYS,
    parseMatrix,
    calculate,
    dataQuality,
    normalizeOEM,
    parseDate,
    elapsedMonths
  };
})(window);
