/*
 * app.js —— 索菲亚自动打单工具 前端逻辑
 * 依赖全局：XLSX、PDFLib、fontkit、SlipMatch、SlipPDF（均由页面 <script> 注入）
 *
 * BOM 数据库：保存在浏览器 localStorage，Excel 导入 / 导出导入 JSON，点「浏览数据库」查看。
 */
(function () {
  'use strict';

  var FONT_REGULAR = 'fonts/NotoSansSC-subset.ttf';
  var FONT_BOLD = 'fonts/NotoSansSC-Bold-subset.ttf';
  var DB_KEY = 'sofia_bom_db_v1';
  var state = {
    bomRows: [],
    planWb: null,
    planFileName: '',
    fonts: null,
    fontsLoading: false,
    blobs: { delivery: null, logistics: null },
    history: [],
  };

  // ---------------- 工具 ----------------
  function $(id) { return document.getElementById(id); }
  function clean(v) { return (v === null || v === undefined) ? '' : String(v).trim(); }
  function num(v, d) { var n = parseFloat(clean(v)); return isNaN(n) ? d : n; }

  var toastTimer = null;
  function toast(msg, type) {
    var el = $('toast');
    el.textContent = msg;
    el.className = 'toast' + (type ? ' ' + type : '');
    el.classList.remove('hidden');
    if (toastTimer) clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.classList.add('hidden'); }, 4000);
  }

  function sleep(ms) { return new Promise(function (r) { setTimeout(r, ms); }); }
  function pad(n) { return String(n).padStart(2, '0'); }
  function todayStr() {
    var t = new Date();
    return t.getFullYear() + '-' + String(t.getMonth() + 1).padStart(2, '0') + '-' + String(t.getDate()).padStart(2, '0');
  }
  function stampStr(ts) {
    var d = new Date(ts);
    return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate()) + ' ' +
      pad(d.getHours()) + ':' + pad(d.getMinutes());
  }

  // ---------------- 字体加载 ----------------
  // 优先用内嵌 base64 字体（fonts-data.js），file:// 直接打开也能用；
  // 没有内嵌数据时回退到 fetch 字体文件（需 http 环境）。
  function b64ToBytes(b64) {
    var bin = atob(b64);
    var len = bin.length;
    var bytes = new Uint8Array(len);
    for (var i = 0; i < len; i++) bytes[i] = bin.charCodeAt(i);
    return bytes;
  }

  async function loadFonts() {
    if (state.fonts) return state.fonts;
    if (state.fontsLoading) { while (state.fontsLoading) await sleep(100); return state.fonts; }
    state.fontsLoading = true;
    try {
      var embedded = (typeof EMBEDDED_FONTS !== 'undefined') ? EMBEDDED_FONTS : null;
      if (embedded && embedded.regular && embedded.bold) {
        state.fonts = { regular: b64ToBytes(embedded.regular), bold: b64ToBytes(embedded.bold) };
        // 窄体西文字体（编码用），缺失时不影响主流程
        if (embedded.codeRegular) state.fonts.codeRegular = b64ToBytes(embedded.codeRegular);
        if (embedded.codeBold) state.fonts.codeBold = b64ToBytes(embedded.codeBold);
        return state.fonts;
      }
      if (location.protocol === 'file:') throw new Error('FONT_MISSING');
      var names = ['NotoSansSC-subset.ttf', 'NotoSansSC-Bold-subset.ttf',
        'RobotoCondensed-Regular.ttf', 'RobotoCondensed-Bold.ttf'];
      var buffers = await Promise.all(names.map(async function (n) {
        try {
          var res = await fetch('fonts/' + n);
          if (!res.ok) return null;
          return await res.arrayBuffer();
        } catch (e) { return null; }
      }));
      state.fonts = {
        regular: new Uint8Array(buffers[0]),
        bold: new Uint8Array(buffers[1]),
      };
      if (buffers[2]) state.fonts.codeRegular = new Uint8Array(buffers[2]);
      if (buffers[3]) state.fonts.codeBold = new Uint8Array(buffers[3]);
      return state.fonts;
    } finally { state.fontsLoading = false; }
  }

  // ---------------- 文件读取 ----------------
  function readWorkbook(file) {
    return new Promise(function (resolve, reject) {
      var reader = new FileReader();
      reader.onerror = function () { reject(new Error('文件读取失败: ' + file.name)); };
      reader.onload = function (e) {
        try {
          var data = new Uint8Array(e.target.result);
          var wb = XLSX.read(data, { type: 'array', cellDates: true });
          resolve(wb);
        } catch (err) { reject(new Error('Excel 解析失败: ' + err.message)); }
      };
      reader.readAsArrayBuffer(file);
    });
  }

  // ---------------- 历史记录（IndexedDB，存 PDF 二进制） ----------------
  var HDB_NAME = 'sofia_slip_history', HDB_STORE = 'records';
  function openHistoryDB() {
    return new Promise(function (resolve, reject) {
      var req = indexedDB.open(HDB_NAME, 1);
      req.onupgradeneeded = function () {
        var db = req.result;
        if (!db.objectStoreNames.contains(HDB_STORE)) {
          db.createObjectStore(HDB_STORE, { keyPath: 'id', autoIncrement: true });
        }
      };
      req.onsuccess = function () { resolve(req.result); };
      req.onerror = function () { reject(req.error); };
    });
  }
  function historyStore(mode) {
    return openHistoryDB().then(function (db) {
      return db.transaction(HDB_STORE, mode).objectStore(HDB_STORE);
    });
  }
  function historyPut(rec) {
    return new Promise(function (resolve, reject) {
      historyStore('readwrite').then(function (s) {
        var r = s.add(rec);
        r.onsuccess = function () { resolve(r.result); };
        r.onerror = function () { reject(r.error); };
      }).catch(reject);
    });
  }
  function historyGetAll() {
    return new Promise(function (resolve, reject) {
      historyStore('readonly').then(function (s) {
        var r = s.getAll();
        r.onsuccess = function () { resolve(r.result || []); };
        r.onerror = function () { reject(r.error); };
      }).catch(reject);
    });
  }
  function historyDelete(id) {
    return new Promise(function (resolve, reject) {
      historyStore('readwrite').then(function (s) {
        var r = s.delete(id);
        r.onsuccess = function () { resolve(); };
        r.onerror = function () { reject(r.error); };
      }).catch(reject);
    });
  }
  function toArrayBuffer(bytes) {
    // 复制出一份精确长度的 buffer，避免共享底层 buffer 被后续覆盖
    var copy = bytes.slice();
    return copy.buffer;
  }
  async function saveHistoryRecord(deliveryBytes, logisticsBytes, meta) {
    if (!deliveryBytes && !logisticsBytes) return;
    try {
      var rec = {
        ts: Date.now(),
        planName: clean(meta.planName),
        senderName: clean(meta.senderName),
        stats: meta.stats || {},
        delivery: deliveryBytes ? toArrayBuffer(deliveryBytes) : null,
        logistics: logisticsBytes ? toArrayBuffer(logisticsBytes) : null,
      };
      await historyPut(rec);
      await loadHistoryList();
      toast('已存入历史记录', 'ok');
    } catch (e) {
      console.error(e);
      toast('历史记录保存失败（浏览器可能不支持 IndexedDB）', 'err');
    }
  }
  async function loadHistoryList() {
    try {
      var recs = await historyGetAll();
      recs.sort(function (a, b) { return b.ts - a.ts; });
      state.history = recs;
    } catch (e) { state.history = []; }
    var n = (state.history || []).length;
    $('history-count').textContent = n ? ('· 共 ' + n + ' 条') : '';
    $('history-empty').classList.toggle('hidden', n > 0);
    if (!$('history-list').classList.contains('hidden')) renderHistoryList();
  }
  function renderHistoryList() {
    var recs = state.history || [];
    if (!recs.length) { $('history-list').innerHTML = ''; return; }
    $('history-list').innerHTML = recs.map(function (r) {
      var s = r.stats || {};
      var meta = '<div class="hc-meta"><b>' + esc(r.planName || '未命名装车计划') + '</b>' +
        (r.senderName ? '<br><span class="muted">' + esc(r.senderName) + '</span>' : '') + '</div>';
      var stat = '<div class="hc-stat">父件 ' + (s.parent || 0) + ' · 计划 ' + (s.plan || 0) + ' 行<br>' +
        '匹配 ' + (s.matched || 0) + '/' + (s.plan || 0) + ' · 分组 ' + (s.groups || 0) + '</div>';
      var acts = [];
      if (r.delivery) acts.push('<button class="btn-mini" data-act="dl" data-kind="delivery" data-id="' + r.id + '">⬇ 送货单</button>');
      if (r.logistics) acts.push('<button class="btn-mini" data-act="dl" data-kind="logistics" data-id="' + r.id + '">⬇ 物流单</button>');
      acts.push('<button class="btn-mini hc-del" data-act="del" data-id="' + r.id + '">删除</button>');
      return '<div class="hc-row">' +
        '<div class="hc-time">' + stampStr(r.ts) + '</div>' +
        meta + stat +
        '<div class="hc-acts">' + acts.join('') + '</div>' +
        '</div>';
    }).join('');
  }
  function dlHistory(id, kind) {
    var r = (state.history || []).find(function (x) { return x.id === id; });
    if (!r || !r[kind]) return;
    var blob = new Blob([r[kind]], { type: 'application/pdf' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url;
    a.download = (kind === 'delivery' ? '送货单_' : '物流单_') + stampStr(r.ts).replace(/[:\s]/g, '') + '.pdf';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }
  async function delHistory(id) {
    if (!confirm('删除这条历史记录？此操作不可撤销。')) return;
    try { await historyDelete(id); await loadHistoryList(); toast('已删除', 'ok'); }
    catch (e) { toast('删除失败', 'err'); }
  }

  // ---------------- BOM 数据库持久化 ----------------
  function loadDb() {
    try {
      var raw = localStorage.getItem(DB_KEY);
      state.bomRows = raw ? JSON.parse(raw) : [];
      if (!Array.isArray(state.bomRows)) state.bomRows = [];
    } catch (e) { state.bomRows = []; }
  }
  function saveDb() {
    try { localStorage.setItem(DB_KEY, JSON.stringify(state.bomRows)); }
    catch (e) { toast('保存失败（可能超出浏览器存储上限）', 'err'); }
  }

  function renderBom() {
    var tbody = $('db-tbody');
    if (!state.bomRows.length) {
      tbody.innerHTML = '';
      $('db-empty').classList.remove('hidden');
    } else {
      $('db-empty').classList.add('hidden');
      var html = state.bomRows.map(function (r, i) {
        return '<tr>' +
          '<td>' + esc(r.parentCode) + '</td>' +
          '<td>' + esc(r.parentName) + '</td>' +
          '<td>' + esc(r.subCode) + '</td>' +
          '<td>' + esc(r.subName) + '</td>' +
          '<td>' + esc(r.subQty) + '</td>' +
          '<td>' + esc(r.weight) + '</td>' +
          '<td>' + esc(r.volume) + '</td>' +
          '<td><button class="row-del" data-idx="' + i + '">删除</button></td>' +
          '</tr>';
      }).join('');
      tbody.innerHTML = html;
    }
    var parents = {};
    state.bomRows.forEach(function (r) { parents[r.parentCode] = 1; });
    $('db-stat').textContent = '父件 ' + Object.keys(parents).length + ' · 条目 ' + state.bomRows.length;
    refreshGenerateBtn();
  }

  function esc(v) {
    return String(v === undefined ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;');
  }

  // ---------------- 批量解析 ----------------
  function mergeRows(newRows) {
    var existing = {};
    state.bomRows.forEach(function (r, idx) { existing[SlipMatch.bomRowKey(r)] = idx; });
    var added = 0, updated = 0;
    newRows.forEach(function (r) {
      var k = SlipMatch.bomRowKey(r);
      if (existing[k] !== undefined) { state.bomRows[existing[k]] = r; updated++; }
      else { state.bomRows.push(r); added++; }
    });
    saveDb(); renderBom();
    return { added: added, updated: updated };
  }

  async function importBomExcel(file) {
    try {
      var wb = await readWorkbook(file);
      var rows = SlipMatch.extractBOMRows(wb);
      var res = mergeRows(rows);
      toast('从 Excel 导入：新增 ' + res.added + '、更新 ' + res.updated, 'ok');
    } catch (err) { toast(err.message, 'err'); }
  }

  function exportJson() {
    if (!state.bomRows.length) { toast('数据库为空，无可导出', 'err'); return; }
    var blob = new Blob([JSON.stringify(state.bomRows, null, 2)], { type: 'application/json' });
    var url = URL.createObjectURL(blob);
    var a = document.createElement('a');
    a.href = url; a.download = 'BOM数据库_' + todayStr() + '.json';
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 1000);
  }

  async function importJson(file) {
    try {
      var text = await file.text();
      var arr = JSON.parse(text);
      if (!Array.isArray(arr)) throw new Error('JSON 格式不正确（应为数组）');
      var valid = arr.filter(function (r) {
        return r && clean(r.parentCode) && clean(r.subCode);
      }).map(function (r) {
        return {
          parentCode: clean(r.parentCode),
          parentName: clean(r.parentName),
          subCode: clean(r.subCode),
          subName: clean(r.subName),
          subQty: num(r.subQty, 1.0),
          weight: num(r.weight, 0),
          volume: num(r.volume, 0),
        };
      });
      if (!valid.length) throw new Error('JSON 中没有有效的 BOM 条目');
      var res = mergeRows(valid);
      toast('导入 JSON：新增 ' + res.added + '、更新 ' + res.updated, 'ok');
    } catch (err) { toast('导入失败: ' + err.message, 'err'); }
  }

  function toggleDbView() {
    var wrap = document.querySelector('.db-table-wrap');
    var btn = $('btn-toggle-db');
    var show = wrap.classList.contains('hidden');
    wrap.classList.toggle('hidden', !show);
    btn.textContent = show ? '收起数据库' : '浏览数据库';
  }

  function deleteRow(idx) {
    state.bomRows.splice(idx, 1);
    saveDb(); renderBom();
  }

  // ---------------- 装车计划 ----------------
  async function handlePlan(file) {
    if (!file) return;
    try {
      state.planWb = await readWorkbook(file);
      state.planFileName = file.name;
      $('plan-name').textContent = file.name;
      toast('已载入 ' + file.name, 'ok');
      refreshGenerateBtn();
    } catch (err) { toast(err.message, 'err'); }
  }

  function refreshGenerateBtn() {
    $('btn-generate').disabled = !(state.bomRows.length && state.planWb);
  }

  // ---------------- 生成 PDF ----------------
  async function generate() {
    if (!(state.bomRows.length && state.planWb)) {
      toast('请先准备 BOM 数据库并上传装车计划', 'err'); return;
    }
    var btn = $('btn-generate');
    btn.disabled = true;
    $('gen-spinner').classList.remove('hidden');
    $('gen-progress').classList.remove('hidden');
    $('gen-status').textContent = '正在加载字体…';
    var fonts;
    try { fonts = await loadFonts(); }
    catch (err) {
      btn.disabled = false;
      $('gen-spinner').classList.add('hidden');
      $('gen-progress').classList.add('hidden');
      if (err.message === 'FONT_MISSING' || err.message === 'FILE_PROTOCOL') {
        toast('字体未能加载。若是本地 file:// 打开，请确认 fonts/fonts-data.js 存在；或用静态服务器打开。', 'err');
      } else { toast(err.message || '字体加载失败', 'err'); }
      $('gen-status').textContent = '';
      return;
    }

    try {
      $('gen-status').textContent = '正在解析与匹配…';
      var bomMap = SlipMatch.buildBomMapFromRows(state.bomRows);
      var planItems = SlipMatch.loadPlan(state.planWb, ['取消']);
      var matchByCodeFirst = document.querySelector('input[name="matchmode"]:checked').value === 'code';
      var matched = SlipMatch.matchItems(planItems, bomMap, matchByCodeFirst);
      var groups = SlipMatch.groupForDelivery(matched.rows);

      var sender = {
        senderName: $('sender-name').value.trim() || '索菲亚家居股份有限公司',
        deliveryPerson: $('sender-person').value.trim(),
        deliveryPhone: $('sender-phone').value.trim(),
      };
      var opts = { expandComboSubs: $('expand-combo').checked, sender: sender };

      $('gen-status').textContent = '正在生成 PDF…';
      var deliveryBytes = null, logisticsBytes = null;
      if ($('gen-delivery').checked) deliveryBytes = await SlipPDF.buildDeliveryPDF(groups, fonts, opts);
      if ($('gen-logistics').checked) logisticsBytes = await SlipPDF.buildLogisticsPDF(matched.rows, fonts, opts);

      await showResult(bomMap, planItems, matched, groups, deliveryBytes, logisticsBytes);

      // 存入历史记录（PDF 二进制写入 IndexedDB）
      saveHistoryRecord(deliveryBytes, logisticsBytes, {
        planName: state.planFileName,
        senderName: $('sender-name').value.trim(),
        stats: {
          parent: Object.keys(bomMap.productsByCode).length,
          plan: planItems.length,
          matched: matched.matched,
          groups: groups.length,
        },
      });

      $('gen-status').textContent = '完成 ✓';
      toast('生成完成', 'ok');
    } catch (err) {
      console.error(err);
      toast('生成失败: ' + err.message, 'err');
      $('gen-status').textContent = '';
    } finally {
      btn.disabled = false;
      $('gen-spinner').classList.add('hidden');
      $('gen-progress').classList.add('hidden');
      refreshGenerateBtn();
    }
  }

  async function showResult(bomMap, planItems, matched, groups, deliveryBytes, logisticsBytes) {
    var deliveryPages = deliveryBytes ? await countPdfPages(deliveryBytes) : '—';
    var logisticsPages = logisticsBytes ? await countPdfPages(logisticsBytes) : '—';
    var stats = [
      { num: Object.keys(bomMap.productsByCode).length, lbl: 'BOM 父件数' },
      { num: planItems.length, lbl: '计划行数' },
      { num: matched.matched + ' / ' + matched.rows.length, lbl: '匹配成功' },
      { num: groups.length, lbl: '送货单分组' },
      { num: deliveryPages, lbl: '送货单页数' },
      { num: logisticsPages, lbl: '物流单页数' },
    ];
    $('stats').innerHTML = stats.map(function (s) {
      return '<div class="stat"><div class="num">' + s.num + '</div><div class="lbl">' + s.lbl + '</div></div>';
    }).join('');
    $('stats').classList.remove('hidden');

    // ---------------- 未匹配提醒（可填写并保存进 BOM） ----------------
    (function renderUnmatched() {
      var seen = {}, list = [];
      matched.rows.forEach(function (r) {
        if (r.bom) return;
        var code = clean(r.plan.materialCode);
        var name = clean(r.plan.materialName);
        var key = code || name;
        if (!key) return;
        if (!seen[key]) { seen[key] = 1; list.push({ code: code, name: name }); }
      });
      var box = $('unmatched-warn');
      if (!list.length) { box.classList.add('hidden'); $('unmatched-list').innerHTML = ''; return; }
      $('unmatched-count').textContent = list.length;
      function umField(label, fld, val, required, type) {
        return '<label class="um-f"><span>' + label + (required ? ' *' : '') + '</span>' +
          '<input class="um-input" data-fld="' + fld + '" type="' + (type || 'text') +
          '" value="' + esc(val) + '"></label>';
      }
      $('unmatched-list').innerHTML = list.map(function (it) {
        var label = it.code ? esc(it.code) : '';
        if (it.name) label += (label ? ' · ' : '') + esc(it.name);
        return '<div class="um-card">' +
          '<div class="um-head">未匹配项：<b>' + label + '</b></div>' +
          '<div class="um-grid">' +
            umField('父件编码', 'parentCode', it.code, true) +
            umField('父件名称', 'parentName', it.name, false) +
            umField('子件编码', 'subCode', it.code, true) +
            umField('子件名称', 'subName', it.name, false) +
            umField('子件数量', 'subQty', '1', false, 'number') +
            umField('重量', 'weight', '', false) +
            umField('体积', 'volume', '', false) +
          '</div>' +
          '<div class="um-actions">' +
            '<button class="btn-mini um-save">保存到数据库</button>' +
            '<span class="um-ok hidden">✓ 已保存到 BOM</span>' +
          '</div>' +
        '</div>';
      }).join('');
      box.classList.remove('hidden');
    })();

    revoke('delivery'); revoke('logistics');

    if (deliveryBytes) {
      var dBlob = new Blob([deliveryBytes], { type: 'application/pdf' });
      var dUrl = URL.createObjectURL(dBlob);
      state.blobs.delivery = { blob: dBlob, url: dUrl };
      $('preview-delivery').src = dUrl;
      $('dl-delivery').href = dUrl;
      $('dl-delivery').download = '送货单_汇总.pdf';
      $('empty-delivery').classList.add('hidden');
    } else { $('preview-delivery').src = ''; $('empty-delivery').classList.remove('hidden'); }

    if (logisticsBytes) {
      var lBlob = new Blob([logisticsBytes], { type: 'application/pdf' });
      var lUrl = URL.createObjectURL(lBlob);
      state.blobs.logistics = { blob: lBlob, url: lUrl };
      $('preview-logistics').src = lUrl;
      $('dl-logistics').href = lUrl;
      $('dl-logistics').download = '物流单_汇总.pdf';
      $('empty-logistics').classList.add('hidden');
    } else { $('preview-logistics').src = ''; $('empty-logistics').classList.remove('hidden'); }

    $('result-card').classList.remove('hidden');
    switchTab(deliveryBytes ? 'delivery' : 'logistics');
  }

  function revoke(kind) {
    var b = state.blobs[kind];
    if (b && b.url) { URL.revokeObjectURL(b.url); state.blobs[kind] = null; }
  }

  async function countPdfPages(bytes) {
    try { return (await PDFLib.PDFDocument.load(bytes)).getPageCount(); }
    catch (e) { return '?'; }
  }

  function switchTab(name) {
    document.querySelectorAll('.tab').forEach(function (t) { t.classList.toggle('active', t.dataset.tab === name); });
    $('pane-delivery').classList.toggle('active', name === 'delivery');
    $('pane-logistics').classList.toggle('active', name === 'logistics');
  }

  function openPrint(kind) {
    var entry = state.blobs[kind];
    if (!entry || !entry.blob) { toast('尚未生成该 PDF', 'err'); return; }
    var url = URL.createObjectURL(entry.blob);
    var w = window.open(url, '_blank');
    if (!w) { toast('浏览器拦截了新窗口，请允许弹窗', 'err'); return; }
    setTimeout(function () { try { w.focus(); w.print(); } catch (e) {} }, 600);
  }

  // ---------------- 事件绑定 ----------------
  function bind() {
    $('file-bom').addEventListener('change', function (e) {
      if (e.target.files[0]) importBomExcel(e.target.files[0]);
      e.target.value = '';
    });
    $('file-bom-json').addEventListener('change', function (e) {
      if (e.target.files[0]) importJson(e.target.files[0]);
      e.target.value = '';
    });
    $('file-plan').addEventListener('change', function (e) { handlePlan(e.target.files[0]); });
    $('btn-export-json').addEventListener('click', exportJson);
    $('btn-toggle-db').addEventListener('click', toggleDbView);
    $('btn-toggle-history').addEventListener('click', function () {
      var list = $('history-list');
      var show = list.classList.contains('hidden');
      list.classList.toggle('hidden', !show);
      $('btn-toggle-history').textContent = show ? '收起历史记录' : '查看历史记录';
      if (show) renderHistoryList();
    });
    $('history-list').addEventListener('click', function (e) {
      var btn = e.target.closest('button[data-act]');
      if (!btn) return;
      var id = Number(btn.dataset.id);
      if (btn.dataset.act === 'dl') dlHistory(id, btn.dataset.kind);
      else if (btn.dataset.act === 'del') delHistory(id);
    });
    $('unmatched-list').addEventListener('click', function (e) {
      var btn = e.target.closest('.um-save');
      if (!btn) return;
      var card = btn.closest('.um-card');
      var getVal = function (f) {
        var el = card.querySelector('input[data-fld="' + f + '"]');
        return el ? el.value.trim() : '';
      };
      var pCode = getVal('parentCode'), sCode = getVal('subCode');
      if (!pCode || !sCode) { toast('父件编码和子件编码不能为空', 'err'); return; }
      var row = {
        parentCode: pCode, parentName: getVal('parentName'),
        subCode: sCode, subName: getVal('subName'),
        subQty: parseFloat(getVal('subQty')) || 1.0,
        weight: parseFloat(getVal('weight')) || 0,
        volume: parseFloat(getVal('volume')) || 0,
      };
      mergeRows([row]);
      card.classList.add('um-saved');
      btn.disabled = true;
      card.querySelector('.um-ok').classList.remove('hidden');
      toast('已保存到 BOM 数据库', 'ok');
    });
    $('db-tbody').addEventListener('click', function (e) {
      if (e.target && e.target.classList.contains('row-del')) {
        deleteRow(parseInt(e.target.dataset.idx, 10));
      }
    });

    var drop = $('drop-plan');
    drop.addEventListener('dragover', function (e) { e.preventDefault(); drop.classList.add('drag'); });
    drop.addEventListener('dragleave', function () { drop.classList.remove('drag'); });
    drop.addEventListener('drop', function (e) {
      e.preventDefault(); drop.classList.remove('drag');
      if (e.dataTransfer.files && e.dataTransfer.files[0]) handlePlan(e.dataTransfer.files[0]);
    });

    $('btn-generate').addEventListener('click', generate);
    document.querySelectorAll('.tab').forEach(function (t) {
      t.addEventListener('click', function () { switchTab(t.dataset.tab); });
    });
    $('print-delivery').addEventListener('click', function () { openPrint('delivery'); });
    $('print-logistics').addEventListener('click', function () { openPrint('logistics'); });
  }

  loadDb();
  loadHistoryList();
  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { bind(); renderBom(); });
  } else {
    bind(); renderBom();
  }
})();
