/*
 * pdfgen.js —— 生成送货单 / 物流单 PDF（浏览器/Node 通用）
 * 依赖全局 PDFLib、fontkit（浏览器）或 require（Node）。
 * fonts 参数：{ regular: Uint8Array, bold: Uint8Array }
 */
(function (global) {
  'use strict';

  const MM = 2.834645669; // 1mm = 2.834645669 pt
  const BROWN = null; // 延迟初始化（需要 PDFLib.rgb）

  function hexToRGB(hex) {
    const h = String(hex).replace('#', '');
    const n = parseInt(h, 16);
    return PDFLib.rgb(((n >> 16) & 255) / 255, ((n >> 8) & 255) / 255, (n & 255) / 255);
  }

  function fmt(v) {
    if (v === null || v === undefined) return '';
    if (v instanceof Date) {
      const y = v.getFullYear();
      const m = String(v.getMonth() + 1).padStart(2, '0');
      const d = String(v.getDate()).padStart(2, '0');
      return `${y}-${m}-${d}`;
    }
    if (typeof v === 'number') {
      if (Number.isInteger(v)) return String(v);
      return (Math.round(v * 100) / 100).toString();
    }
    return String(v).trim();
  }

  // 按宽度折行（逐字符，适配中文）
  function wrapText(text, font, size, maxWidth) {
    text = (text === null || text === undefined) ? '' : String(text);
    if (text === '') return [''];
    if (font.widthOfTextAtSize(text, size) <= maxWidth + 0.01) return [text];
    const lines = [];
    let cur = '';
    for (const ch of text) {
      const test = cur + ch;
      if (font.widthOfTextAtSize(test, size) <= maxWidth + 0.01) {
        cur = test;
      } else {
        if (cur) lines.push(cur);
        cur = ch;
      }
    }
    if (cur) lines.push(cur);
    return lines;
  }

  // 判断是否为编码/数字串：这类串用窄体西文字体渲染（更紧凑好看）
  const ALNUM = /^[0-9A-Za-z\-_./ ]+$/;
  function isCodeLike(text) {
    return ALNUM.test(String(text || '').replace(/\s/g, ''));
  }

  // 为编码类文字挑一款窄体西文字体（若已提供 code 字体），否则回退原字体
  function pickCodeFont(fonts, text) {
    if (!fonts.codeBold && !fonts.codeRegular) return null;
    if (!isCodeLike(text)) return null;
    return fonts.codeBold || fonts.codeRegular;
  }

  function drawCentered(page, text, cx, baselineY, size, font, color) {
    const w = font.widthOfTextAtSize(text, size);
    page.drawText(text, { x: cx - w / 2, y: baselineY, size, font, color });
  }

  function drawRight(page, text, rightX, baselineY, size, font, color) {
    const w = font.widthOfTextAtSize(text, size);
    page.drawText(text, { x: rightX - w, y: baselineY, size, font, color });
  }

  // 嵌入字体：中文（regular/bold）必选，窄体西文（codeRegular/codeBold）可选
  async function embedFonts(pdf, fontsBytes) {
    const f = {
      regular: await pdf.embedFont(fontsBytes.regular, { subset: false }),
      bold: await pdf.embedFont(fontsBytes.bold, { subset: false }),
    };
    if (fontsBytes.codeRegular) {
      f.codeRegular = await pdf.embedFont(fontsBytes.codeRegular, { subset: false });
    }
    if (fontsBytes.codeBold) {
      f.codeBold = await pdf.embedFont(fontsBytes.codeBold, { subset: false });
    }
    return f;
  }

  // ---------------- 送货单 ----------------
  function buildDeliveryPDF(groups, fontsBytes, opts) {
    opts = opts || {};
    const PDFDocument = PDFLib.PDFDocument;
    return (async () => {
      const pdf = await PDFDocument.create();
      await pdf.registerFontkit(fontkit);
      const fonts = await embedFonts(pdf, fontsBytes);
      const expandComboSubs = !!opts.expandComboSubs;
      const senderCfg = opts.sender || {};
      for (const g of groups) {
        const page = pdf.addPage([842, 595]); // A4 横向
        drawDeliveryPage(page, g, fonts, { expandComboSubs, senderCfg });
      }
      return await pdf.save();
    })();
  }

  function drawDeliveryPage(page, g, fonts, cfg) {
    const reg = fonts.regular, bold = fonts.bold;
    const BLACK = PDFLib.rgb(0, 0, 0);
    const GRAY = hexToRGB('F0F0F0');
    const margin = 24;
    const top = 595 - margin;

    page.drawText('供应商送货单', { x: 842 / 2 - reg.widthOfTextAtSize('供应商送货单', 18) / 2, y: top - 10, size: 18, font: reg, color: BLACK });

    const first = g.items[0].plan;
    const cfgSender = (cfg && cfg.senderCfg) || {};
    const sender = g.senderName || cfgSender.senderName || '索菲亚家居股份有限公司';
    const deliveryPerson = g.deliveryPerson || cfgSender.deliveryPerson || '';
    const deliveryPhone = g.deliveryPhone || cfgSender.deliveryPhone || '';

    let y = top - 40;
    const lx = margin, lv = margin + 80, rx = 842 / 2 + 10, rv = 842 / 2 + 90;
    // 右侧值区可用宽度（到右边距）
    const rAvail = 842 - margin - rv;
    const lAvail = rx - lv - 10;
    function hline(ly, labX, labV, label, value, avail) {
      page.drawText(label, { x: labX, y: ly, size: 10, font: reg, color: BLACK });
      const txt = fmt(value);
      // 编码/数字串用窄体西文字体，字距更紧凑
      const cf = pickCodeFont(fonts, txt);
      page.drawText(txt, { x: labV, y: ly, size: 10, font: cf || reg, color: BLACK });
    }
    hline(y, lx, lv, '采购单号：', first.purchaseOrder, lAvail);
    hline(y, rx, rv, '收货方：', first.dealerAccount + first.dealerName, rAvail);
    y -= 18;
    hline(y, lx, lv, '送货方：', sender, lAvail);
    hline(y, rx, rv, '物流商：', first.logistics, rAvail);
    y -= 18;
    hline(y, lx, lv, '仓库联系人：', first.warehouseContact);
    hline(y, rx, rv, '仓库电话：', fmt(first.warehousePhone));
    y -= 18;
    page.drawText('仓库地址：', { x: lx, y, size: 10, font: reg, color: BLACK });
    page.drawText(fmt(first.warehouseAddress), { x: lv, y, size: 10, font: reg, color: BLACK });
    y -= 26;

    // 表格
    const tableTop = y;
    const headers = ['订单号', '物料编码', '物料名称', '数量', '发运计划', '分包编码', '分包物料名称', '单位', '总件数', '总重量', '总体积'];
    const data = [headers];
    let totalPackages = 0, totalWeight = 0, totalVolume = 0;
    const matchApi = global.SlipMatch;
    for (const row of g.items) {
      const expanded = matchApi.expandRows(row, cfg.expandComboSubs);
      for (const ex of expanded) {
        const p = ex.p;
        const product = row.bom;
        const sub = ex.sub;
        const subCode = sub ? sub.subCode : p.materialCode;
        let subName;
        if (product && sub && sub.subCode === product.parentCode) subName = product.parentName;
        else subName = sub ? sub.subName : p.materialName;
        const shipDate = fmt(p.requiredDate);
        const pkg = ex.isFirst ? fmt(p.totalPackages) : '';
        const wt = ex.isFirst ? fmt(p.totalWeight) : '';
        const vol = ex.isFirst ? fmt(p.totalVolume) : '';
        data.push([
          fmt(p.customerPO), fmt(p.materialCode), fmt(p.materialName), fmt(ex.qty),
          shipDate, fmt(subCode), fmt(subName), fmt(p.unit), pkg, wt, vol,
        ]);
        if (ex.isFirst) {
          totalPackages += p.totalPackages;
          totalWeight += p.totalWeight;
          totalVolume += p.totalVolume;
        }
      }
    }

    // 构建单元格对象
    const colWidthsRaw = [70, 70, 100, 40, 60, 70, 100, 35, 45, 55, 55];
    const available = 842 - 2 * margin;
    const scale = available / colWidthsRaw.reduce((a, b) => a + b, 0);
    const colWidths = colWidthsRaw.map(w => w * scale);

    const cellRows = data.map((r, ri) =>
      r.map((text, ci) => ({
        text: fmt(text),
        font: ri === 0 ? 'bold' : 'regular',
        size: 9,
        align: 'center',
        color: BLACK,
      }))
    );

    // 计算行高（padX 需与绘制时一致，确保文字不贴/越过网格线）
    const lineH = 13, padX = 5, padY = 3;

    // 编码/数字类单元格：改用窄体西文字体，字距更紧凑（不再缩字号）
    cellRows.forEach((row, ri) => {
      row.forEach(cell => {
        if (ri === 0) return; // 表头用中文粗体
        const cf = pickCodeFont(fonts, cell.text);
        if (cf) cell.cf = cf;
      });
    });

    const rowHeights = cellRows.map(row => {
      let maxLines = 1;
      row.forEach((cell, ci) => {
        const font = cell.cf || fonts[cell.font];
        const lines = wrapText(cell.text, font, cell.size, colWidths[ci] - padX * 2);
        cell._lines = lines;
        if (lines.length > maxLines) maxLines = lines.length;
      });
      return maxLines * lineH + padY * 2;
    });

    let yy = tableTop;
    // 表头底色
    page.drawRectangle({ x: margin, y: yy - rowHeights[0], width: available, height: rowHeights[0], color: GRAY });
    for (let ri = 0; ri < cellRows.length; ri++) {
      const rh = rowHeights[ri];
      const row = cellRows[ri];
      let cx = margin;
      const totalTextH = Math.max.apply(null, row.map(c => c._lines.length)) * lineH;
      const startY = yy - (rh - totalTextH) / 2 - 9;
      for (let ci = 0; ci < row.length; ci++) {
        const cell = row[ci];
        const cw = colWidths[ci];
        const font = cell.cf || fonts[cell.font];
        for (let li = 0; li < cell._lines.length; li++) {
          const line = cell._lines[li];
          const w = font.widthOfTextAtSize(line, cell.size);
          let tx;
          if (cell.align === 'center') tx = cx + (cw - w) / 2;
          else if (cell.align === 'right') tx = cx + cw - padX - w;
          else tx = cx + padX;
          page.drawText(line, { x: tx, y: startY - li * lineH, size: cell.size, font, color: cell.color });
        }
        cx += cw;
      }
      yy -= rh;
    }
    // 网格线
    let gx = margin;
    for (let ci = 0; ci <= colWidths.length; ci++) {
      page.drawLine({ start: { x: gx, y: tableTop }, end: { x: gx, y: yy }, thickness: 0.5, color: BLACK });
      if (ci < colWidths.length) gx += colWidths[ci];
    }
    // 顶边框 + 各行下边框
    page.drawLine({ start: { x: margin, y: tableTop }, end: { x: margin + available, y: tableTop }, thickness: 0.5, color: BLACK });
    let hy = tableTop;
    for (let ri = 0; ri < rowHeights.length; ri++) {
      hy -= rowHeights[ri];
      page.drawLine({ start: { x: margin, y: hy }, end: { x: margin + available, y: hy }, thickness: 0.5, color: BLACK });
    }

    // 汇总 + 签字
    let sy = yy - 20;
    page.drawText('汇总：', { x: margin, y: sy, size: 10, font: reg, color: BLACK });
    drawRight(page, fmt(totalPackages), 842 - margin - 140, sy, 10, reg, BLACK);
    drawRight(page, fmt(totalWeight), 842 - margin - 80, sy, 10, reg, BLACK);
    drawRight(page, fmt(totalVolume), 842 - margin - 20, sy, 10, reg, BLACK);

    sy -= 25;
    page.drawText('收货人签字：', { x: margin, y: sy, size: 10, font: reg, color: BLACK });
    page.drawText('送货人签字：' + deliveryPerson, { x: margin + 180, y: sy, size: 10, font: reg, color: BLACK });
    const today = new Date();
    const ds = `${today.getFullYear()}-${String(today.getMonth() + 1).padStart(2, '0')}-${String(today.getDate()).padStart(2, '0')}`;
    page.drawText('送货日期：' + ds, { x: margin + 400, y: sy, size: 10, font: reg, color: BLACK });
    sy -= 20;
    page.drawText('收货日期：', { x: margin, y: sy, size: 10, font: reg, color: BLACK });
    page.drawText('送货人电话：' + deliveryPhone, { x: margin + 180, y: sy, size: 10, font: reg, color: BLACK });
  }

  // ---------------- 物流单 ----------------
  function buildLogisticsPDF(rows, fontsBytes, opts) {
    opts = opts || {};
    const PDFDocument = PDFLib.PDFDocument;
    return (async () => {
      const pdf = await PDFDocument.create();
      await pdf.registerFontkit(fontkit);
      const fonts = await embedFonts(pdf, fontsBytes);
      const w = 100 * MM, h = 150 * MM;
      const matchApi = global.SlipMatch;
      for (const row of rows) {
        const p = row.plan;
        const product = row.bom;
        const sub = row.sub;
        const subCode = sub ? sub.subCode : p.materialCode;
        let subName;
        if (product && sub && sub.subCode === product.parentCode) subName = product.parentName;
        else subName = sub ? sub.subName : p.materialName;
        const pkgTotal = p.pkgTotal || 1;
        for (let pkg = 1; pkg <= pkgTotal; pkg++) {
          const page = pdf.addPage([w, h]);
          drawLogisticsLabel(page, {
            fonts, w, h, p, product, sub, subCode, subName, pkgNo: pkg, pkgTotal,
          });
        }
      }
      return await pdf.save();
    })();
  }

  function drawLogisticsLabel(page, ctx) {
    const fonts = ctx.fonts, w = ctx.w, h = ctx.h, p = ctx.p;
    const product = ctx.product, sub = ctx.sub;
    const subCode = ctx.subCode, subName = ctx.subName;
    const pkgNo = ctx.pkgNo, pkgTotal = ctx.pkgTotal;
    const reg = fonts.regular, bold = fonts.bold;
    const BLACK = PDFLib.rgb(0, 0, 0);
    const BROWN = hexToRGB('8B4513');
    const GRAY = hexToRGB('666666');

    let qtyPerPkg;
    if (product && sub && sub.subCode !== product.parentCode) {
      qtyPerPkg = (sub.subQty * p.qty) / pkgTotal;
    } else {
      qtyPerPkg = p.qtyPerPackage;
    }
    const qtyText = `${fmt(qtyPerPkg)}（行号${p.seq}，第${pkgNo}包，共${pkgTotal}包）`;

    const margin = 3 * MM;
    const keyW = 28 * MM;
    const valW = w - margin * 2 - keyW;
    // 单元格左右内边距：文字必须离边框留出间距，否则长编码会贴线/越界
    const cellPadX = 3;

    const data = [
      ['分包物料编码', fmt(subCode), 'regular', 9, false],
      ['分包物料名称', fmt(subName), 'regular', 9, false],
      ['产品数量', qtyText, 'regular', 9, false],
      ['整包物料编码', fmt(p.materialCode), 'regular', 9, false],
      ['整包物料名称', fmt(p.materialName), 'regular', 9, false],
      ['销售单号', fmt(p.customerPO), 'regular', 9, false],
      ['经销商编码', fmt(p.dealerAccount), 'regular', 9, false],
      ['经销商名称', fmt(p.dealerName), 'regular', 9, false],
      ['收货人', fmt(p.receiver), 'regular', 9, false],
      ['联系电话', fmt(p.receiverPhone), 'regular', 9, false],
      ['收货地址', fmt(p.receiverAddress), 'regular', 9, false],
      ['物流商', fmt(p.logistics), 'regular', 9, false],
      ['仓库联系人', fmt(p.warehouseContact), 'regular', 9, true],
      ['仓库电话', fmt(p.warehousePhone), 'regular', 9, true],
      ['仓库地址', fmt(p.warehouseAddress), 'regular', 9, true],
    ];

    const lineH = 12, padX = 2, padY = 4;
    const rows = data.map(([k, v, f, size, kb]) => {
      const font = fonts[f];
      const avail = valW - cellPadX * 2;
      // 编码/数字串改用窄体西文字体（数字更紧凑），字号统一
      let useFont = font;
      if (isCodeLike(v)) {
        const cf = pickCodeFont(fonts, v);
        if (cf) useFont = cf;
      }
      const lines = wrapText(v, useFont, size, avail);
      const linesK = wrapText(k, fonts.regular, 9, keyW - cellPadX * 2);
      const maxLines = Math.max(lines.length, linesK.length);
      const rh = maxLines * (size + 3) + padY * 2;
      return { k, v, f, size, font: useFont, kb, lines, linesK, rh };
    });
    let tableH = rows.reduce((s, r) => s + r.rh, 0);

    const titleSize = 22, subSize = 8;
    const titleAbove = titleSize / 72 * 25.4 + 1;
    const titleToSub = 9 * MM;
    const subDescent = subSize / 72 * 25.4 * 0.35;
    const subToTable = 5 * MM;
    const titleBlock = titleAbove + titleToSub + subDescent + subToTable;
    const minMargin = 4 * MM;

    let equalMargin;
    if (tableH + titleBlock <= h - 2 * minMargin) {
      equalMargin = (h - (tableH + titleBlock)) / 2;
    } else {
      equalMargin = minMargin;
      const avail = h - 2 * minMargin - titleBlock;
      const sc = avail / tableH;
      rows.forEach(r => { r.rh = Math.max(6 * MM, r.rh * sc); });
      tableH = rows.reduce((s, r) => s + r.rh, 0);
    }

    const titleBaseline = h - equalMargin - titleAbove;
    drawCentered(page, '索菲亚', w / 2, titleBaseline, titleSize, reg, BROWN);
    drawCentered(page, '家具 | 整家定制', w / 2, titleBaseline - titleToSub, subSize, reg, GRAY);

    const tableTop = titleBaseline - titleToSub - subDescent - subToTable;
    // 画表格
    let ty = tableTop;
    for (let ri = 0; ri < rows.length; ri++) {
      const r = rows[ri];
      const rh = r.rh;
      // key
      const keyLines = r.linesK;
      const keyTextH = keyLines.length * 12;
      const keyY = ty - (rh - keyTextH) / 2 - 9;
      for (let li = 0; li < keyLines.length; li++) {
        const ww = fonts.regular.widthOfTextAtSize(keyLines[li], 9);
        page.drawText(keyLines[li], { x: margin + cellPadX + (keyW - cellPadX * 2 - ww) / 2, y: keyY - li * 12, size: 9, font: fonts.regular, color: BLACK });
      }
      // value
      const valFont = r.font;
      const valTextH = r.lines.length * (r.size + 3);
      const valY = ty - (rh - valTextH) / 2 - r.size;
      for (let li = 0; li < r.lines.length; li++) {
        const ww = valFont.widthOfTextAtSize(r.lines[li], r.size);
        const cx = margin + keyW + cellPadX + (valW - cellPadX * 2 - ww) / 2;
        page.drawText(r.lines[li], { x: cx, y: valY - li * (r.size + 3), size: r.size, font: valFont, color: BLACK });
      }
      ty -= rh;
    }
    // 网格
    for (let ci = 0; ci <= 2; ci++) {
      const xLine = (ci === 0) ? margin : (ci === 1) ? margin + keyW : margin + keyW + valW;
      page.drawLine({ start: { x: xLine, y: tableTop }, end: { x: xLine, y: ty }, thickness: 0.5, color: BLACK });
    }
    // 顶边框 + 各行下边框
    page.drawLine({ start: { x: margin, y: tableTop }, end: { x: margin + keyW + valW, y: tableTop }, thickness: 0.5, color: BLACK });
    let hy = tableTop;
    for (let ri = 0; ri < rows.length; ri++) {
      hy -= rows[ri].rh;
      page.drawLine({ start: { x: margin, y: hy }, end: { x: margin + keyW + valW, y: hy }, thickness: 0.5, color: BLACK });
    }
  }

  const API = { buildDeliveryPDF, buildLogisticsPDF };
  global.SlipPDF = Object.assign(global.SlipPDF || {}, API);
  if (typeof module !== 'undefined' && module.exports) module.exports = API;
})(typeof window !== 'undefined' ? window : globalThis);
