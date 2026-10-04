/*
 * match.js —— BOM 解析 + 装车计划解析 + 匹配（浏览器/Node 通用）
 * 依赖全局 XLSX（浏览器）或 require('xlsx')（Node）。
 * 通过 window.SlipMatch / module.exports 双重导出。
 */
(function (global) {
  'use strict';

  function toStr(v) {
    return (v === null || v === undefined) ? '' : String(v).trim();
  }
  function toInt(v, d) {
    d = (d === undefined) ? 0 : d;
    const n = parseInt(v, 10);
    return isNaN(n) ? d : n;
  }
  function toFloat(v, d) {
    d = (d === undefined) ? 0.0 : d;
    const n = parseFloat(v);
    return isNaN(n) ? d : n;
  }
  // 去掉表头里的单位括号，如 "总重量（kg）" -> "总重量"
  function stripUnit(h) {
    return String(h).replace(/[（(].*?[)）]/g, '').trim();
  }
  function fmtNum(v) {
    if (v === null || v === undefined || v === '') return '';
    if (typeof v === 'number') {
      return Number.isInteger(v) ? String(v) : String(v);
    }
    return String(v).trim();
  }
  function fmtDate(v) {
    if (v === null || v === undefined || v === '') return '';
    if (v instanceof Date) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, '0');
      const d = String(v.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    const s = String(v).trim();
    for (const f of ['YYYY-MM-DD', 'YYYY/MM/DD', 'YYYYMMDD']) {
      // 简化：只做常见格式尝试
    }
    return s;
  }

  // ---------------- BOM 数据库 ----------------
  // 扁平 BOM 行模型：{ parentCode, parentName, subCode, subName, subQty, weight, volume }
  // 表头支持中英文别名（不区分大小写、忽略单位括号和空格）
  const HEADER_ALIASES = {
    parentCode: ['父件编码', '父件代码', 'parentcode', 'parent_code'],
    parentName: ['父件名称', 'parentname'],
    subCode: ['子件编码', '子件代码', 'subcode', 'sub_code'],
    subName: ['子件名称', 'subname'],
    subQty: ['子件数量', '数量', 'subqty', 'qty'],
    weight: ['重量', 'weight'],
    volume: ['体积', 'volume'],
  };

  function buildHeaderIdx(header) {
    const idx = {};
    header.forEach((h, i) => {
      const s = stripUnit(h).toLowerCase().replace(/\s+/g, '');
      if (!s) return;
      for (const field in HEADER_ALIASES) {
        if (idx[field] === undefined && HEADER_ALIASES[field].indexOf(s) >= 0) idx[field] = i;
      }
    });
    return idx;
  }

  function extractBOMRows(wb) {
    const names = wb.SheetNames || [];
    const sheetName = names.find(n => n.indexOf('BOM数据库') >= 0) || names[0];
    if (!sheetName) throw new Error('BOM 工作簿中找不到工作表');
    const sheet = wb.Sheets[sheetName];
    const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, cellDates: true });
    if (!rows.length) throw new Error('BOM 数据库为空');
    const header = (rows[0] || []).map(h => (h === null ? '' : String(h).trim()));
    const idx = buildHeaderIdx(header);
    // 只强制要求 父件编码 + 子件编码；其余列缺失时用默认值
    const required = ['parentCode', 'subCode'];
    const missing = required.filter(k => idx[k] === undefined);
    if (missing.length) {
      throw new Error('BOM 表缺少必需列: ' + missing.join(', ') +
        '（表头需包含 父件编码/子件编码 或 parentCode/subCode；当前表头: ' +
        header.filter(h => h).join(' | ') + '）');
    }
    const get = (row, field) => (idx[field] !== undefined && idx[field] < row.length) ? row[idx[field]] : null;

    const out = [];
    for (let r = 1; r < rows.length; r++) {
      const row = rows[r];
      if (!row) continue;
      if (row.every(v => v === null || v === undefined || v === '')) continue;
      const parentCode = toStr(get(row, 'parentCode'));
      const parentName = toStr(get(row, 'parentName'));
      const subCode = toStr(get(row, 'subCode'));
      const subName = toStr(get(row, 'subName'));
      const subQty = toFloat(get(row, 'subQty'), 1.0);
      const weight = toFloat(get(row, 'weight'), 0.0);
      const volume = toFloat(get(row, 'volume'), 0.0);
      if (!parentCode || !subCode) continue;
      out.push({ parentCode, parentName, subCode, subName, subQty, weight, volume });
    }
    return out;
  }

  // 从扁平行（无论来自 Excel 还是手动添加）构建查询表
  function buildBomMapFromRows(bomRows) {
    const grouped = {};
    const namesMap = {};
    for (const item of bomRows) {
      const { parentCode, parentName, subCode, subName, subQty, weight, volume } = item;
      if (!parentCode || !subCode) continue;
      const sub = { subCode, subName, subQty, weight, volume };
      if (!grouped[parentCode]) grouped[parentCode] = [];
      grouped[parentCode].push(sub);
      namesMap[parentCode] = parentName;
    }
    const productsByCode = {};
    const productsByName = {};
    for (const code in grouped) {
      const product = { parentCode: code, parentName: namesMap[code] || '', subs: grouped[code] };
      productsByCode[code] = product;
      if (product.parentName) productsByName[product.parentName] = product;
    }
    return { productsByCode, productsByName };
  }

  function parseBOM(wb) {
    return buildBomMapFromRows(extractBOMRows(wb));
  }

  // 行去重键：父件编码 + 子件编码
  function bomRowKey(row) {
    return toStr(row.parentCode) + '|' + toStr(row.subCode);
  }

  // ---------------- 手动批量文本解析 ----------------
  function splitLine(line) {
    if (line.indexOf('\t') >= 0) return line.split('\t');
    if (line.indexOf(',') >= 0) return line.split(',');
    return line.split(/\s+/);
  }

  function detectHeaderMap(cells) {
    const map = {};
    const specs = [
      ['parentCode', ['父件编码', 'parentcode', 'parent_code']],
      ['parentName', ['父件名称', 'parentname']],
      ['subCode', ['子件编码', 'subcode', 'sub_code']],
      ['subName', ['子件名称', 'subname']],
      ['subQty', ['子件数量', '数量', 'subqty']],
      ['weight', ['重量', 'weight']],
      ['volume', ['体积', 'volume']],
    ];
    cells.forEach((c, i) => {
      const s = toStr(c).toLowerCase();
      specs.forEach(sp => {
        if (map[sp[0]] === undefined && sp[1].indexOf(s) >= 0) map[sp[0]] = i;
      });
    });
    return map;
  }

  // 解析批量粘贴文本：每行一条，逗号/制表符/空格分隔，可带表头。
  // 返回 { rows: [{parentCode,parentName,subCode,subName,subQty,weight,volume}], skipped }
  function parseBomBatch(text) {
    let lines = String(text).split(/\r?\n/).map(l => l.trim()).filter(l => l.length);
    if (!lines.length) return { rows: [], skipped: 0 };

    let useMap = false, headerMap = null;
    const first = lines[0].toLowerCase();
    if (first.indexOf('父件编码') >= 0 || first.indexOf('子件编码') >= 0 ||
        first.indexOf('parentcode') >= 0 || first.indexOf('subcode') >= 0) {
      headerMap = detectHeaderMap(splitLine(lines[0]));
      useMap = headerMap.parentCode !== undefined && headerMap.subCode !== undefined;
      lines = lines.slice(1);
    }

    const colVal = (cols, field, pos) => {
      if (useMap) { const i = headerMap[field]; return (i !== undefined) ? (cols[i] || '') : ''; }
      return cols[pos] || '';
    };

    let rows = [], skipped = 0;
    for (let i = 0; i < lines.length; i++) {
      const cols = splitLine(lines[i]);
      if (cols.length < 4) { skipped++; continue; }
      const pc = toStr(colVal(cols, 'parentCode', 0));
      const sc = toStr(colVal(cols, 'subCode', 2));
      if (!pc || !sc) { skipped++; continue; }
      rows.push({
        parentCode: pc,
        parentName: toStr(colVal(cols, 'parentName', 1)),
        subCode: sc,
        subName: toStr(colVal(cols, 'subName', 3)),
        subQty: toFloat(colVal(cols, 'subQty', 4), 1.0),
        weight: toFloat(colVal(cols, 'weight', 5), 0),
        volume: toFloat(colVal(cols, 'volume', 6), 0),
      });
    }
    return { rows, skipped };
  }

  function piecesPerPackage(product) {
    const name = product.parentName || '';
    const m = name.match(/-(\d+)(?:件|张|个|套|只|条|包|支)\/包/);
    if (m) return parseInt(m[1], 10);
    const subs = product.subs || [];
    if (subs.length === 1 || subs.every(s => s.subCode === product.parentCode)) {
      const q = subs.length ? subs[0].subQty : 0;
      if (q && 0 < q && q < 1) return Math.round(1 / q);
    }
    const q = subs.length ? subs[0].subQty : 0;
    if (q && q >= 1) return Math.round(q);
    return 1;
  }

  function matchOne(bomMap, materialCode, materialName, matchByCodeFirst) {
    matchByCodeFirst = (matchByCodeFirst !== false);
    materialCode = toStr(materialCode);
    materialName = toStr(materialName);
    if (matchByCodeFirst && materialCode && bomMap.productsByCode[materialCode]) {
      return { product: bomMap.productsByCode[materialCode], note: '物料编码匹配 (' + materialCode + ')' };
    }
    if (materialName && bomMap.productsByName[materialName]) {
      return { product: bomMap.productsByName[materialName], note: '物料名称匹配 (' + materialName + ')' };
    }
    if (!matchByCodeFirst && materialCode && bomMap.productsByCode[materialCode]) {
      return { product: bomMap.productsByCode[materialCode], note: '物料编码匹配 (' + materialCode + ')' };
    }
    return { product: null, note: null };
  }

  // ---------------- 装车计划 ----------------
  function loadPlan(wb, skipSheets) {
    skipSheets = skipSheets || ['取消'];
    const skip = (skipSheets || []).map(s => String(s).toLowerCase());
    const items = [];
    (wb.SheetNames || []).forEach(sheetName => {
      if (skip.some(s => sheetName.toLowerCase().indexOf(s) >= 0)) return;
      const sheet = wb.Sheets[sheetName];
      const rows = XLSX.utils.sheet_to_json(sheet, { header: 1, defval: null, cellDates: true });
      if (!rows.length) return;
      const header = (rows[0] || []).map(h => (h === null ? '' : String(h).trim()));
      const idx = {};
      header.forEach((h, i) => { idx[stripUnit(h)] = i; });
      for (let r = 1; r < rows.length; r++) {
        const row = rows[r];
        if (!row) continue;
        if (row.every(v => v === null || v === undefined || v === '')) continue;
        const get = (k, d) => {
          const i = idx[k];
          return (i !== undefined && i < row.length) ? row[i] : (d === undefined ? null : d);
        };
        const item = {
          seq: toInt(get('序号', 0)),
          purchaseOrder: toStr(get('采购订单', '')),
          erpOrder: get('ERP采购订单'),
          orderOu: toStr(get('订单OU', '')),
          customerPO: toStr(get('客户PO', '')),
          province: toStr(get('省份', '')),
          dealerName: toStr(get('经销商名称', '')),
          dealerAccount: toStr(get('经销商账号', '')),
          materialCode: toStr(get('物料编码', '')),
          materialName: toStr(get('物料名称', '')),
          requiredDate: get('用户需求日期'),
          shipDate: get('更新计划发运日期'),
          qty: toFloat(get('需求数量', 0)),
          unit: toStr(get('单位', '件')),
          totalPackages: toFloat(get('总件数', 0)),
          totalWeight: toFloat(get('总重量', 0)),
          totalVolume: toFloat(get('总体积', 0)),
          receiver: toStr(get('收货联系人', '')),
          receiverPhone: get('收货电话'),
          receiverAddress: toStr(get('收货地址', '')),
          logistics: toStr(get('物流商', '')),
          warehouseAddress: toStr(get('仓库地址', '')),
          warehouseContact: toStr(get('仓库联系人', '')),
          warehousePhone: get('仓库电话'),
          deliveryMode: toStr(get('配送模式', '')),
          carSeq: get('装车序号'),
          carNo: get('车次'),
          sheetName: sheetName,
        };
        if (!item.materialCode && !item.materialName) continue;
        item.qtyPerPackage = (item.totalPackages && item.totalPackages > 0)
          ? (item.qty / item.totalPackages) : item.qty;
        item.pkgTotal = (item.totalPackages && item.totalPackages > 0) ? Math.round(item.totalPackages) : 1;
        items.push(item);
      }
    });
    return items;
  }

  function matchItems(planItems, bomMap, matchByCodeFirst) {
    const rows = [];
    let matched = 0;
    for (const item of planItems) {
      const res = matchOne(bomMap, item.materialCode, item.materialName, matchByCodeFirst);
      const sub = res.product ? res.product.subs[0] : null;
      rows.push({ plan: item, bom: res.product, sub: sub, matchNote: res.note || '未匹配' });
      if (res.product) matched++;
    }
    return { rows, matched };
  }

  function isCombo(product) {
    if (!product || !product.subs || !product.subs.length) return false;
    return product.subs.some(s => s.subCode !== product.parentCode);
  }

  function expandRows(row, expandComboSubs) {
    const p = row.plan;
    const product = row.bom;
    if (!expandComboSubs || !isCombo(product)) {
      return [{ p: p, sub: row.sub, isFirst: true, qty: p.qty }];
    }
    const out = [];
    product.subs.forEach((sub, i) => {
      out.push({ p: p, sub: sub, isFirst: (i === 0), qty: p.qty * sub.subQty });
    });
    return out;
  }

  function groupForDelivery(rows) {
    const groups = {};
    for (const row of rows) {
      const key = row.plan.purchaseOrder + '||' + row.plan.customerPO;
      if (!groups[key]) {
        groups[key] = {
          purchaseOrder: row.plan.purchaseOrder,
          customerPO: row.plan.customerPO,
          items: [],
        };
      }
      groups[key].items.push(row);
    }
    return Object.keys(groups).map(k => groups[k]);
  }

  const API = {
    toStr, toInt, toFloat, stripUnit, fmtNum, fmtDate,
    parseBOM, extractBOMRows, buildBomMapFromRows, bomRowKey, parseBomBatch,
    piecesPerPackage, matchOne, matchItems,
    loadPlan, isCombo, expandRows, groupForDelivery,
  };

  global.SlipMatch = Object.assign(global.SlipMatch || {}, API);
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
})(typeof window !== 'undefined' ? window : globalThis);
