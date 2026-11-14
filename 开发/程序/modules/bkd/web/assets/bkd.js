/* 货币资金 BKD 生成器 · 前端逻辑 */
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var $$ = function (s) { return Array.prototype.slice.call(document.querySelectorAll(s)); };

  var state = { tb: null, prior: null, priorPick: null, leadPick: null, result: null, expPriorFid: '' };

  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function fm(n) {
    if (n == null || n === '') return '';
    var v = Number(n);
    if (!isFinite(v)) return '';
    return v.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function oneLine(t) { return String(t == null ? '' : t).replace(/\s*\n\s*/g, ''); }
  function toast(msg, err) {
    var d = document.createElement('div');
    d.className = 'toast' + (err ? ' err' : '');
    d.textContent = msg;
    $('#toastWrap').appendChild(d);
    setTimeout(function () { d.classList.add('on'); }, 10);
    setTimeout(function () { d.classList.remove('on'); setTimeout(function () { d.remove(); }, 220); }, 2800);
  }

  function setStatus(el, cls, html) {
    var e = $(el);
    if (!html) { e.className = 'status'; e.innerHTML = ''; return; }
    e.className = 'status ' + cls;
    e.innerHTML = html;
  }

  /* ---------- 自动取数：科目余额表 ---------- */
  function loadTb() {
    setStatus('#tbStatus', 'loading', '正在从PTW平台读取清洗后的科目余额表…');
    fetch('/api/tb').then(function (r) { return r.json(); }).then(function (d) {
      if (d.error) {
        setStatus('#tbStatus', 'err', '读取失败：' + esc(d.error));
        return;
      }
      state.tb = d;
      var n = (d.records || []).length;
      setStatus('#tbStatus', 'ok', '已从PTW平台读取 <b>' + n + '</b> 个货币资金账户' +
        (d.meta && d.meta.rows ? '（余额表共 ' + d.meta.rows + ' 行）' : ''));
      renderTb();
      updateGating();
      // 读完之后顺手跑一次预览（不落盘），把折算率主体、待补格数备好，
      // 用户在「关键信息」那一步就能看到该填什么
      if (state.onTbReady) state.onTbReady();
    }).catch(function (e) {
      setStatus('#tbStatus', 'err', '无法连接本地服务：' + esc(e.message));
    });
  }

  function renderTb() {
    var d = state.tb, recs = d.records || [];
    var sumLocal = 0, sumOrig = 0;
    recs.forEach(function (r) { sumLocal += r.local; sumOrig += r.orig; });
    var sum = $('#tbSummary');
    sum.hidden = false;
    sum.innerHTML = '' +
      '<div class="tile"><div class="k">账户数</div><div class="v">' + recs.length + '</div></div>' +
      '<div class="tile"><div class="k">期末本币合计</div><div class="v">' + fm(sumLocal) + '</div></div>' +
      '<div class="tile"><div class="k">期末原币合计</div><div class="v">' + fm(sumOrig) + '</div></div>';
    var h = ['<table><thead><tr><th>公司名称</th><th>科目编码</th><th>科目名称</th><th>银行/机构</th><th>账号</th><th>币种</th><th class="num">期末原币</th><th class="num">期末本币</th></tr></thead><tbody>'];
    recs.slice(0, 10).forEach(function (r) {
      h.push('<tr><td>' + esc(r.company) + '</td><td class="mono">' + esc(r.acct_code || '') + '</td><td>' + esc(r.subject) + '</td><td>' + esc(r.bank) + '</td><td class="mono">' + esc(r.acctNo || '—') + '</td><td>' + esc(r.currency) + '</td><td class="num">' + fm(r.orig) + '</td><td class="num">' + fm(r.local) + '</td></tr>');
    });
    h.push('</tbody></table>');
    if (recs.length > 10) h.push('<div class="tbl-note">仅显示前 10 行，共 ' + recs.length + ' 行。</div>');
    var t = $('#tbTable');
    t.hidden = false;
    t.innerHTML = h.join('');
  }

  /* ---------- 上年 BKD 底稿 ---------- */
  function initDz() {
    var dz = $('#dz-prior');
    var input = document.createElement('input');
    input.type = 'file';
    input.accept = '.xlsx,.xlsm,.xls';
    input.style.display = 'none';
    document.body.appendChild(input);
    dz.addEventListener('click', function () { input.value = ''; input.click(); });
    input.addEventListener('change', function () {
      if (input.files && input.files[0]) handlePrior(input.files[0]);
    });
    dz.addEventListener('dragover', function (e) { e.preventDefault(); dz.classList.add('over'); });
    dz.addEventListener('dragleave', function () { dz.classList.remove('over'); });
    dz.addEventListener('drop', function (e) {
      e.preventDefault(); dz.classList.remove('over');
      var f = e.dataTransfer.files && e.dataTransfer.files[0];
      if (f) handlePrior(f);
    });
  }

  function handlePrior(file) {
    state._priorFile = file;
    var dz = $('#dz-prior');
    dz.querySelector('.dz-file').textContent = '正在解析 ' + file.name + ' …';
    var fd = new FormData();
    fd.append('prior', file, file.name);
    fetch('/api/prior/plan', { method: 'POST', body: fd }).then(function (r) { return r.json(); }).then(function (d) {
      if (d.error) { toast('解析失败：' + d.error, true); return; }
      state.prior = d;
      state.priorPick = null;      // 换文件时重新按自动识别结果选中
      state.leadPick = null;
      dz.querySelector('.dz-file').textContent = file.name + ' · ' + d.sheets.length + ' 个工作表';
      dz.classList.add('loaded');
      renderPrior();
      updateGating();
    }).catch(function (e) { toast('上传失败：' + e.message, true); });
  }

  function pickIdx() {
    var d = state.prior;
    if (!d || !d.sheets || !d.sheets.length) return -1;
    var i = state.priorPick;
    if (i == null || i < 0 || i >= d.sheets.length) {
      // 后端会给出推荐工作表（C.00 BKD），没给就退回第一个被识别的
      var k = d.recommended;
      if (k == null || k < 0 || k >= d.sheets.length) {
        k = -1;
        for (var j = 0; j < d.sheets.length; j++) {
          if (d.sheets[j].usable !== false) { k = j; break; }
        }
      }
      i = state.priorPick = (k >= 0 ? k : 0);
    }
    return i;
  }

  // 上年 Lead 的选表：与 BKD 完全同构——推荐项优先，没有推荐就用第一个能用的
  function leadPickIdx() {
    var d = state.prior;
    if (!d || !d.sheets || !d.sheets.length) return -1;
    var i = state.leadPick;
    if (i == null || i < 0 || i >= d.sheets.length) {
      var k = d.leadRecommended;
      if (k == null || k < 0 || k >= d.sheets.length) {
        k = -1;
        for (var j = 0; j < d.sheets.length; j++) {
          if (d.sheets[j].leadUsable) { k = j; break; }
        }
      }
      i = state.leadPick = (k >= 0 ? k : 0);
    }
    return i;
  }

  // 生成 BKD 时把上年文件一起带上：后端按 {sheet, header_row} 重新解析，
  // 拿到上期审定数、性质、是否函证。没上传就返回 null，接口不传这个字段即可。
  function priorSpec() {
    var d = state.prior, i = pickIdx();
    if (!d || i < 0 || !d.fid) return null;
    var s = d.sheets[i];
    var hd = Number(($('#priorHd') || {}).value) || s.header_row;
    // 上年 Lead：与 BKD 同一个文件、同一套 {sheet, header_row} 结构
    var li = leadPickIdx();
    var lead = null;
    if (li >= 0) {
      var ls = d.sheets[li];
      var lhd = Number(($('#leadHd') || {}).value) || ls.leadHeaderRow || 1;
      lead = { sheet: ls.sheet, header_row: lhd };
    }
    return { fid: d.fid, spec: { sheet: s.sheet, header_row: hd, lead: lead } };
  }

  function renderPrior() {
    var d = state.prior, i = pickIdx();
    if (i < 0) return;
    var s = d.sheets[i];
    var reset = !!s._mapReset;                 // 用户改过表头行 → 列的识别结果待后端重算
    var hasMap = !!(s.map && Object.keys(s.map).length);
    var usable = reset ? true : (s.usable !== false);
    var acctCol = '—', amtCol = '—';
    if (hasMap && s.map.acctNo) {
      var idxs = Array.isArray(s.map.acctNo) ? s.map.acctNo : [s.map.acctNo];
      acctCol = idxs.map(function (c) { return colName(c + 1) + ' ' + oneLine(s.headers[c] || ''); }).join('、');
    }
    if (hasMap && s.map.localEnd != null) amtCol = colName(s.map.localEnd + 1) + ' 本币';
    else if (hasMap && s.map.endAny != null) amtCol = colName(s.map.endAny + 1) + ' 期末';
    else if (hasMap && s.map.origEnd != null) amtCol = colName(s.map.origEnd + 1) + ' 原币';
    if (reset) { acctCol = '生成时按第 ' + s.header_row + ' 行重新识别'; amtCol = '同上'; }

    var opts = d.sheets.map(function (x, k) {
      var tag = (k === d.recommended) ? '（推荐）'
        : (x.usable === false ? '（未识别）' : '');
      return '<option value="' + k + '"' + (k === i ? ' selected' : '') + '>' +
        esc(x.sheet) + tag + '</option>';
    }).join('');

    var h = [];
    h.push('<div class="alert ok">已识别上年底稿。下面两组各选一次：哪张表是「货币资金 BKD 明细」、' +
      '哪张表是「上年 Lead」（通常是 C.00 BKD 与 C.00 Lead）。不选就按推荐项走。</div>');
    h.push('<div class="rowgrid">' +
      '<label>上年 BKD 明细工作表<select id="priorPick">' + opts + '</select></label>' +
      '<label>该表表头所在行<input type="number" min="1" id="priorHd" value="' + s.header_row + '"/></label>' +
      '</div>');
    h.push('<div class="sheets-pick mt">' +
      '<span class="k">账号列</span><span class="mono v">' + esc(acctCol) + '</span>' +
      '<span class="k">金额列</span><span class="mono v">' + esc(amtCol) + '</span>' +
      '<span class="k">状态</span><span class="v">' +
        (reset ? '<span class="tag-off">表头行已调整，生成时重新识别</span>'
               : (usable ? '<span class="tag-on">已识别为 BKD 明细</span>'
                         : '<span class="tag-off">未识别到账号/银行/科目列</span>')) + '</span>' +
      '</div>');
    if (!usable) {
      h.push('<div class="alert warn mt">当前工作表没能识别出「账号 / 银行 / 科目」列，' +
        '请确认是否选错工作表，或调整上方的表头行号后重试。</div>');
    }
    // ---- 上年 Lead：与上面 BKD 完全同构的一组控件（选表 + 表头行 + 状态） ----
    var li = leadPickIdx();
    var ls = (li >= 0) ? d.sheets[li] : null;
    var lOpts = d.sheets.map(function (x, k) {
      var tag = (k === d.leadRecommended) ? '（推荐）' : (x.leadUsable ? '' : '（未识别）');
      return '<option value="' + k + '"' + (k === li ? ' selected' : '') + '>' +
        esc(x.sheet) + tag + '</option>';
    }).join('');
    h.push('<div class="rowgrid mt">' +
      '<label>上年 Lead 工作表<select id="leadPick">' + lOpts + '</select></label>' +
      '<label>该表表头所在行<input type="number" min="1" id="leadHd" value="' +
      ((ls && ls.leadHeaderRow) || 1) + '"/></label>' +
      '</div>');
    h.push('<div class="sheets-pick mt">' +
      '<span class="k">科目行</span><span class="mono v">' + (ls ? (ls.leadRows || 0) : 0) + ' 行</span>' +
      '<span class="k">读到审定数</span><span class="mono v">' + (ls ? (ls.leadAmountCells || 0) : 0) + ' 格</span>' +
      '<span class="k">状态</span><span class="v">' +
        (ls && ls.leadUsable
          ? '<span class="tag-on">已识别为 Lead</span>'
          : '<span class="tag-off">未识别到「科目名称＋期末审定数」</span>') +
      '</span></div>');
    if (ls && !ls.leadUsable) {
      h.push('<div class="alert warn mt">当前工作表没有「科目名称 + 期末审定数」表头，' +
        '本年表1 的上期数与表2 的上年数会取不到数。请换一张工作表，或调整表头行号。</div>');
    }

    $('#priorPane').innerHTML = h.join('');

    $('#priorPick').addEventListener('change', function () {
      state.priorPick = Number(this.value);
      renderPrior();
      updateGating();
    });
    $('#priorHd').addEventListener('change', function () {
      s.header_row = Math.max(1, Number(this.value) || 1);
      // 表头行改了，让后端按新表头行重新识别列（前端不缓存 map）
      s.map = {};
      s._mapReset = true;
      renderPrior();
      updateGating();
    });

    if ($('#leadPick')) $('#leadPick').addEventListener('change', function () {
      state.leadPick = Number(this.value);
      renderPrior();
    });
    if ($('#leadHd')) $('#leadHd').addEventListener('change', function () {
      var j = leadPickIdx();
      if (j >= 0) d.sheets[j].leadHeaderRow = Math.max(1, Number(this.value) || 1);
      renderPrior();
      updateGating();
    });
  }

  function colName(n) {
    var s = '';
    while (n > 0) { var m = (n - 1) % 26; s = String.fromCharCode(65 + m) + s; n = Math.floor((n - 1) / 26); }
    return s;
  }

  /* ---------- 门控：流程条 + 生成前置检查 ---------- */
  /* 四步流程是这套界面唯一的进度语言，两个科目共用同一套：
     done 表示这一步的资料已经齐了，active 表示当前该看哪一步。 */
  function flowSet(pane, done, active) {
    var bar = $('#flow' + (pane === 'exp' ? 'Exp' : 'Cash'));
    if (!bar) return;
    [].forEach.call(bar.querySelectorAll('.fl'), function (el) {
      var n = Number(el.getAttribute('data-f'));
      el.classList.toggle('done', n <= done);
      el.classList.toggle('active', n === active && n > done);
    });
  }

  function tbReady() {
    return !!(state.tb && state.tb.records && state.tb.records.length);
  }

  function priorReady() {
    return !!(state.prior && state.prior.sheets && state.prior.sheets.length && pickIdx() >= 0);
  }

  function renderPreflight() {
    var box = $('#preflight');
    if (!box) return;
    var n = tbReady() ? state.tb.records.length : 0;
    var tpl = state.tpl || {};
    var ps = priorReady() ? state.prior.sheets[pickIdx()] : null;
    var items = [
      { ok: n > 0, must: true,
        txt: n ? '余额表已读取 <b>' + n + '</b> 个货币资金账户'
               : '还没读到货币资金账户 —— 先去「导入清洗」合并科目余额表' },
      { ok: !!tpl.exists, must: true,
        txt: tpl.exists ? '底稿模板已就绪：<b>' + esc(tpl.filename || '') + '</b>'
                        : '没找到底稿模板，去第 2 步上传一份' },
      { ok: !!ps, must: false,
        txt: ps ? '上年底稿已选中工作表 <b>' + esc(ps.sheet) + '</b>（第 ' +
                  (ps.header_row || 1) + ' 行作表头），上年数可回捞'
                : '未上传上年底稿（可选）—— 上年数将留空标黄' }
    ];
    box.innerHTML = items.map(function (it) {
      var cls = it.ok ? 'ok' : (it.must ? 'no' : 'opt');
      var ic = it.ok ? '✓' : (it.must ? '✕' : '—');
      return '<div class="pf ' + cls + '"><span class="ic">' + ic + '</span><span>' + it.txt + '</span></div>';
    }).join('');
    var b = $('#btnGenAll');
    if (b) b.disabled = !(n > 0 && tpl.exists);
    flowSet('cash', (n > 0 ? 1 : 0) + (tpl.exists ? 1 : 0) + (ps ? 1 : 0), 4);
  }

  function updateGating() {
    renderPreflight();
  }

  function genConfig() {
    var on = $('#optCutoff') && $('#optCutoff').checked;
    var pick = null;
    var pi = pickIdx();
    if (pi >= 0) {
      var ps = state.prior.sheets[pi];
      // 只传选中的那一张工作表；map 不缓存，由后端按最终表头行重新识别列
      pick = { sheet: ps.sheet, header_row: ps.header_row, map: {} };
    }
    return {
      keyword: $('#optKeyword').value || '货币资金',
      rate: $('#optRate').value || 50,
      amt: $('#optAmt').value || 100000,
      highlight: $('#optHighlight').checked ? '1' : '0',
      mark: $('#optMark').checked ? '1' : '0',
      cutoff: on ? '1' : '0',
      // 截止期间跟着资产负债表日走，不接受单独填的日期
      cutoff_date: ($('#optC00Date') || {}).value || '',
      cutoff_before: ($('#optCutoffBefore') || {}).value || 1,
      cutoff_after: ($('#optCutoffAfter') || {}).value || 1,
      cutoff_threshold: ($('#optCutoffAmt') || {}).value || 100000,
      customer: ($('#optC00Customer') || {}).value || '',
      balance_date: ($('#optC00Date') || {}).value || '',
      standard: ($('#optStandard') || {}).value || '',
      currency: ($('#optCurrency') || {}).value || '',
      te: ($('#optTE') || {}).value || '',
      sad: ($('#optSAD') || {}).value || '',
      sheets: JSON.stringify(pick ? [pick] : [])
    };
  }

  /* ---------- 基础信息：全底稿共用（客户 / 日期 / 准则 / 本位币 / TE / SAD） ---------- */
  function syncCutoffDate() {
    var d = $('#optC00Date'), c = $('#optCutoffDate');
    if (d && c) c.value = d.value || '';
  }
  function fillBaseInfo(i) {
    i = i || {};
    var set = function (sel, v) {
      var el = $(sel);
      if (el && v != null && v !== '') el.value = v;
    };
    set('#optC00Customer', i.customer);
    set('#optC00Date', i.balance_date);
    set('#optStandard', i.standard);
    set('#optCurrency', i.currency);
    set('#optTE', i.te);
    set('#optSAD', i.sad);
    var ed = $('#optExpDate');
    if (ed && !ed.value && i.balance_date) ed.value = i.balance_date;
    syncCutoffDate();
  }
  /* ---------- CRA 配置（按科目） ----------
     百分比存百分比数字（70 表示 70%），提交时换成 0.70 存盘；
     Lead 的认定阈值公式用 VLOOKUP 查这张表，所以改这里就等于改公式里的百分比。 */
  var CRA_STATE = { modules: [], types: [], assertions: [], data: {} };

  function craPctEls() {
    return { Minimal: $('#craPctMinimal'), Low: $('#craPctLow'),
             Moderate: $('#craPctModerate'), High: $('#craPctHigh') };
  }

  function renderCraAssertions() {
    var box = $('#craAssertRows');
    if (!box) return;
    box.innerHTML = '';
    CRA_STATE.assertions.forEach(function (a) {
      var lab = document.createElement('label');
      lab.textContent = a.label;
      var sel = document.createElement('select');
      sel.id = 'craA_' + a.key;
      var opt = document.createElement('option');
      opt.value = ''; opt.textContent = '（不写，保留手选）';
      sel.appendChild(opt);
      CRA_STATE.types.forEach(function (t) {
        var o = document.createElement('option');
        o.value = t; o.textContent = t;
        sel.appendChild(o);
      });
      lab.appendChild(sel);
      box.appendChild(lab);
    });
  }

  function fillCra(cra, types, assertions, modules) {
    CRA_STATE.types = types || [];
    CRA_STATE.assertions = assertions || [];
    CRA_STATE.modules = modules || [];
    CRA_STATE.data = cra || {};
    var msel = $('#craModule');
    if (msel) {
      msel.innerHTML = '';
      CRA_STATE.modules.forEach(function (m) {
        var o = document.createElement('option');
        o.value = m; o.textContent = m;
        msel.appendChild(o);
      });
      msel.onchange = function () { paintCra(msel.value); };
    }
    renderCraAssertions();
    paintCra(CRA_STATE.modules[0] || '');
  }

  function paintCra(module) {
    var cfg = (CRA_STATE.data || {})[module] || {};
    var pcts = cfg.pcts || {};
    var els = craPctEls();
    Object.keys(els).forEach(function (k) {
      if (!els[k]) return;
      var v = pcts[k];
      els[k].value = (v == null) ? '' : Math.round(v * 1000) / 10;
    });
    var as = cfg.assertions || {};
    CRA_STATE.assertions.forEach(function (a) {
      var el = $('#craA_' + a.key);
      if (el) el.value = as[a.key] || '';
    });
  }

  function craBody() {
    var module = ($('#craModule') || {}).value || (CRA_STATE.modules[0] || '');
    var pcts = {}, els = craPctEls();
    Object.keys(els).forEach(function (k) {
      if (!els[k]) return;
      var raw = String(els[k].value || '').replace('%', '').trim();
      if (raw === '') return;
      var num = parseFloat(raw);
      if (!isNaN(num)) pcts[k] = num / 100;
    });
    var assertions = {};
    CRA_STATE.assertions.forEach(function (a) {
      var el = $('#craA_' + a.key);
      if (el && el.value) assertions[a.key] = el.value;
    });
    var out = {};
    out[module] = { pcts: pcts, assertions: assertions };
    return out;
  }

  function loadBaseInfo() {
    fetch('/api/build/baseinfo').then(function (r) { return r.json(); }).then(function (d) {
      fillBaseInfo(d && d.info);
      fillCra(d && d.cra, d && d.craTypes, d && d.assertions, d && d.modules);
    }).catch(function () { });
  }

  function exportBaseInfo() {
    fetch('/api/build/baseinfo/export').then(function (r) { return r.json(); }).then(function (d) {
      if (!d || !d.ok) { toast((d && d.error) || '导出失败', true); return; }
      var blob = new Blob([d.text], { type: 'application/json;charset=utf-8' });
      var a = document.createElement('a');
      a.href = URL.createObjectURL(blob);
      a.download = d.filename || 'PTW关键信息.json';
      document.body.appendChild(a); a.click();
      setTimeout(function () { URL.revokeObjectURL(a.href); a.remove(); }, 500);
      toast('已导出：' + (d.filename || ''));
    }).catch(function () { toast('导出失败（网络）', true); });
  }

  function importBaseInfo(file) {
    if (!file) return;
    var fr = new FileReader();
    fr.onload = function () {
      fetch('/api/build/baseinfo/import', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ text: String(fr.result || '') })
      }).then(function (r) { return r.json(); }).then(function (d) {
        if (d && d.ok) { toast('已导入关键信息'); loadBaseInfo(); }
        else toast((d && d.error) || '导入失败', true);
      }).catch(function () { toast('导入失败（网络）', true); });
    };
    fr.readAsText(file, 'utf-8');
  }
  function baseInfoBody() {
    return {
      customer: ($('#optC00Customer') || {}).value || '',
      balance_date: ($('#optC00Date') || {}).value || '',
      standard: ($('#optStandard') || {}).value || '',
      currency: ($('#optCurrency') || {}).value || '',
      te: ($('#optTE') || {}).value || '',
      sad: ($('#optSAD') || {}).value || '',
      cra: craBody()
    };
  }
  function saveBaseInfo() {
    fetch('/api/build/baseinfo/save', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(baseInfoBody())
    }).then(function (r) { return r.json(); }).then(function (d) {
      if (d && d.ok) toast('已保存：全底稿共用这一份基础信息');
      else toast((d && d.error) || '保存失败', true);
    }).catch(function () { toast('保存失败（网络）', true); });
  }

  /* ---------- Cutoff 设置 ---------- */
  function initCutoff() {
    var cb = $('#optCutoff');
    if (!cb) return;
    var box = box_(cb), dateEl = $('#optCutoffDate');
    // 截止期间恒等于资产负债表日，不再给「本年 12-31」的默认值。
    // 那个默认值曾让 C.03 按 12-31 找样本，一条都抓不到、还不报错。
    if (dateEl) dateEl.readOnly = true;
    var dEl = $('#optC00Date');
    if (dEl) dEl.addEventListener('change', syncCutoffDate);
    syncCutoffDate();
    function sync() {
      var on = cb.checked;
      if (box) box.classList.toggle('off', !on);
      ['#optCutoffDate', '#optCutoffBefore', '#optCutoffAfter', '#optCutoffAmt'].forEach(function (sel) {
        var el = $(sel); if (el) el.disabled = !on;
      });
    }
    cb.addEventListener('change', sync);
    sync();
  }
  function box_(el) { return el && el.closest ? el.closest('.cutoff-box') : null; }

  function run() {
    var btn = $('#btnRun');
    var cfg = genConfig();
    if (cfg.cutoff === '1' && !cfg.cutoff_date) {
      toast('请先填写截止期间（资产负债表日）', true); return;
    }
    btn.disabled = true; btn.textContent = '生成中…';
    var fd = new FormData();
    // 重新上传上年底稿（后端不持久化文件）
    if (state._priorFile) fd.append('prior', state._priorFile, state._priorFile.name);
    Object.keys(cfg).forEach(function (k) { fd.append(k, cfg[k]); });
    fetch('/api/generate', { method: 'POST', body: fd }).then(function (r) { return r.json(); }).then(function (d) {
      btn.disabled = false; btn.textContent = '生成 BKD 底稿';
      if (d.error) { toast('生成失败：' + d.error, true); return; }
      state.result = d;
      renderResult();
      $('#cardRes').hidden = false;
      $('#cardRes').scrollIntoView({ behavior: 'smooth', block: 'start' });
    }).catch(function (e) {
      btn.disabled = false; btn.textContent = '生成 BKD 底稿';
      toast('生成失败：' + e.message, true);
    });
  }

  function renderResult() {
    var st = state.result.stat;
    var tiles = [
      { k: '上年 BKD 账户', v: st.priorCount, s: '折合本币 ' + fm(st.sumPrior) },
      { k: '本年余额表账户', v: st.tbCount, s: '折合本币 ' + fm(st.sumTb) },
      { k: '匹配成功', v: st.matched, s: '精确 ' + st.exact + ' · 仅账号 ' + st.onlyAcct + ' · 公司+银行 ' + st.byBank, cls: 'ok' },
      { k: '本年新增（上年无）', v: st.newCount, s: '本币 ' + fm(sumLocal(state.result.news)), cls: st.newCount ? 'warn' : '' },
      { k: '上年有、本年无', v: st.missCount, s: '本币 ' + fm(sumLocal(state.result.miss)), cls: st.missCount ? 'danger' : '' },
      { k: '期末余额差异', v: fm(st.diff), s: '本年 ' + fm(st.sumTb) + ' − 上年 ' + fm(st.sumPrior), cls: Math.abs(st.diff) < 0.005 ? 'ok' : '' }
    ];
    $('#tiles').innerHTML = tiles.map(function (t) {
      return '<div class="tile ' + (t.cls || '') + '"><div class="k">' + t.k + '</div><div class="v">' + esc(t.v) + '</div><div class="s">' + esc(t.s || '') + '</div></div>';
    }).join('');
    renderBkdPreview();
    renderCutoff();
    renderCheckPreview();
  }

  function sumLocal(list) { var s = 0; (list || []).forEach(function (r) { s += r.local; }); return s; }

  function renderBkdPreview() {
    var rows = state.result.rows || [];
    var shown = rows.slice(0, 10);
    var h = ['<div class="scroll"><table><thead><tr><th>公司名称</th><th>科目名称</th><th>银行/机构</th><th>账号</th><th>币种</th><th class="num">期末原币</th><th class="num">期末本币</th><th>性质用途</th><th>是否函证</th></tr></thead><tbody>'];
    shown.forEach(function (r) {
      var isNew = r._matched === false;
      h.push('<tr' + (isNew ? ' class="row-new"' : '') + '><td>' + esc(r.company) + '</td><td>' + esc(r.subject) + '</td><td>' + esc(r.bankName) + '</td><td class="mono">' + esc(r.acctNo || '—') + '</td><td>' + esc(r.currency) + '</td><td class="num">' + fm(r.orig) + '</td><td class="num">' + fm(r.local) + '</td><td>' + esc(r.nature) + '</td><td>' + esc(r.confirm) + '</td></tr>');
    });
    h.push('</tbody></table></div>');
    h.push('<div class="tbl-note">浅黄底色为「上年底稿未找到」的新增账户，需补录性质用途与函证安排。' +
      (rows.length > 10 ? '仅显示前 10 行，共 ' + rows.length + ' 行。' : '') + '</div>');
    $('#resPane-bkd').innerHTML = h.join('');
  }

  /* ---------- C.03 Cutoff 截止性测试预览 ---------- */
  function renderCutoff() {
    var box = $('#resPane-cutoff');
    if (!box) return;
    var c = (state.result || {}).cutoff;
    if (!c || c.enabled === false) {
      box.innerHTML = '<div class="tbl-note">本次生成未启用截止性测试，「C.03 Cutoff」保持模板原样。</div>';
      return;
    }
    var st = c.stat || {};
    var h = [];
    if (c.warn) h.push('<div class="alert warn">' + esc(c.warn) + '</div>');

    function tile(k, v, s, cls) {
      return '<div class="tile ' + (cls || '') + '"><div class="k">' + esc(k) + '</div>' +
        '<div class="v">' + esc(v) + '</div><div class="s">' + esc(s || '') + '</div></div>';
    }
    h.push('<div class="tiles">' +
      tile('截止期间', st.endDate || '—', '前后各 ' + st.beforeDays + ' / ' + st.afterDays + ' 天（' +
        (st.windowStart || '') + ' 至 ' + (st.windowEnd || '') + '）') +
      tile('测试阈值', fm(st.threshold), '单笔转账金额下限') +
      tile('数据来源', c.source === 'flow' ? '银行流水' : '未接入流水',
        c.source === 'flow' ? ('流水 ' + (st.flowRows || 0) + ' 行') : '样本区留空，待取得流水') +
      tile('识别转账', st.transferPairs, '期末前 ' + st.beforeCount + ' · 期末后 ' + st.afterCount) +
      tile('跨期转账', st.crossCount, st.crossCount ? '需重点关注' : '无', st.crossCount ? 'warn' : 'ok') +
      tile('未配对单边记录', st.singleCount, st.singleCount ? '请复核是否漏记到账' : '无',
        st.singleCount ? 'warn' : '') +
      '</div>');

    function subTable(title, rows, note) {
      var out = ['<div class="subhd">' + title + '（' + rows.length + ' 笔）</div>'];
      if (!rows.length) {
        out.push('<div class="tbl-note">' +
          (c.source === 'flow' ? '期间内未识别到符合条件的转账。'
                               : '本期未接入银行流水，样本区留空；取得流水后重跑即自动填充。') +
          '</div>');
        return out.join('');
      }
      out.push('<div class="scroll"><table><thead><tr>' +
        '<th>编号</th><th>公司名称</th><th>付款银行</th><th>付款账号</th><th>付款日期</th>' +
        '<th>凭证号</th><th>币种</th><th class="num">付款金额</th>' +
        '<th>收款银行</th><th>收款账号</th><th>收款日期</th><th class="num">收款金额</th><th>跨期</th>' +
        '</tr></thead><tbody>');
      rows.forEach(function (s) {
        out.push('<tr>' +
          '<td>' + esc(s.no) + '</td><td>' + esc(s.company) + '</td>' +
          '<td>' + esc(s.payBank) + '</td><td class="mono">' + esc(s.payAcct || '—') + '</td>' +
          '<td>' + esc(s.payDate) + '</td><td>' + esc(s.vno) + '</td><td>' + esc(s.currency) + '</td>' +
          '<td class="num">' + fm(s.payAmt) + '</td>' +
          '<td>' + esc(s.recvBank) + '</td><td class="mono">' + esc(s.recvAcct || '—') + '</td>' +
          '<td>' + esc(s.recvDate) + '</td><td class="num">' + fm(s.recvAmt) + '</td>' +
          '<td' + (s.cross === '是' ? ' class="row-new"' : '') + '>' + esc(s.cross) + '</td>' +
          '</tr>');
      });
      out.push('</tbody></table></div>');
      if (note) out.push('<div class="tbl-note">' + note + '</div>');
      return out.join('');
    }

    var noteB = (st.insertedRows > 0)
      ? '模板每表预置 5 行，本次已在 C.03 自动插入共 ' + st.insertedRows +
        ' 行，全部样本均已写入底稿。' : '';
    var noteA = '';

    h.push(subTable('一、期末前银行间转账测试', c.before || [], noteB));
    h.push(subTable('二、期末后银行间转账测试', c.after || [], noteA));

    if (st.belowThreshold) {
      h.push('<div class="tbl-note">另有 ' + st.belowThreshold + ' 笔转账低于阈值 ' +
        fm(st.threshold) + ' 元，未纳入测试范围。</div>');
    }
    if (c.singleTotal) {
      h.push('<div class="subhd">三、未配对单边记录（' + c.singleTotal + ' 笔，需人工复核）</div>');
      var sg = ['<div class="scroll"><table><thead><tr><th>方向</th><th>日期</th><th>凭证号</th>' +
        '<th>科目</th><th>银行</th><th>账号</th><th class="num">金额</th><th>摘要</th></tr></thead><tbody>'];
      (c.single || []).forEach(function (r) {
        sg.push('<tr><td>' + esc(r._side) + '</td><td>' + esc(r.vdate || '') + '</td>' +
          '<td>' + esc(r.vno) + '</td><td>' + esc(r.name) + '</td>' +
          '<td>' + esc(r.bank) + '</td><td class="mono">' + esc(r.acctNo || '—') + '</td>' +
          '<td class="num">' + fm(r._amt) + '</td><td>' + esc(r.summary) + '</td></tr>');
      });
      sg.push('</tbody></table></div>');
      if (c.singleTotal > (c.single || []).length) {
        sg.push('<div class="tbl-note">仅列出前 ' + (c.single || []).length + ' 笔，共 ' + c.singleTotal + ' 笔。</div>');
      }
      h.push(sg.join(''));
    }
    h.push('<div class="tbl-note">配对口径：同一凭证号 + 同一主体 + 金额相等的一借一贷；跨期转账为付款日与到账日分处资产负债表日两侧且金额相等。' +
      '付款方填 I 列为负、收款方填 M 列为正，模板 P 列 Check（=M+I）应归零。</div>');
    box.innerHTML = h.join('');
  }

  function renderCheckPreview() {
    var res = state.result, st = res.stat;
    var h = [];

    h.push('<div class="subhd">一、匹配情况汇总</div>');
    h.push(simpleTable(
      [{ t: '项目' }, { t: '数量', num: true }, { t: '金额（本币）', num: true }, { t: '说明' }],
      [
        ['上年 BKD 底稿账户数', st.priorCount, fm(st.sumPrior), '去重后账户数'],
        ['本年科目余额表账户数', st.tbCount, fm(st.sumTb), 'PTW平台自动取数'],
        ['匹配成功', st.matched, '', '精确 ' + st.exact + ' · 仅账号 ' + st.onlyAcct + ' · 公司+银行 ' + st.byBank],
        ['本年新增（上年未找到）', st.newCount, fm(sumLocal(res.news)), '详见第三部分'],
        ['上年有、本年无', st.missCount, fm(sumLocal(res.miss)), '详见第四部分'],
        ['期末余额差异', '', fm(st.diff), Math.abs(st.diff) < 0.005 ? '一致' : '系新增/销户及余额变动所致']
      ]));

    h.push('<div class="subhd">二、本年有、上年 BKD 未找到（' + res.news.length + ' 户）</div>');
    h.push(simpleTable(
      [{ t: '公司名称' }, { t: '账号' }, { t: '银行/机构名称' }, { t: '币种' }, { t: '期末本币', num: true }, { t: '处理建议' }],
      res.news.map(function (r) {
        return [r.company, r.acctNo || '（未识别到账号）', r.bankName, r.currency, fm(r.local),
          r.acctNo ? '确认为新增账户，补录性质用途/函证安排' : '检查余额表账号列或手工补录'];
      }), 'row-new'));

    h.push('<div class="subhd">三、上年有、本年余额表未找到（' + res.miss.length + ' 户）</div>');
    h.push(simpleTable(
      [{ t: '公司名称' }, { t: '账号' }, { t: '银行/机构名称' }, { t: '币种' }, { t: '上年本币', num: true }, { t: '处理建议' }],
      res.miss.map(function (r) {
        return [r.company, r.acctNo || '（无账号）', r.bank, r.currency, fm(r.local), '如已销户确认余额转出；如未销户核实是否漏记'];
      }), 'row-miss'));

    h.push('<div class="subhd">四、非精确匹配（' + res.nonexact.length + ' 户，需人工复核）</div>');
    h.push(simpleTable(
      [{ t: '本年公司名称' }, { t: '本年账号' }, { t: '匹配方式' }, { t: '上年底稿公司' }, { t: '上年底稿账号' }, { t: '本年余额', num: true }, { t: '上年余额', num: true }],
      res.nonexact.map(function (r) {
        return [r.company, r.acctNo, r._how, r.priorCompany, r.priorAcct || '', fm(r.local), fm(r.priorLocal)];
      })));

    h.push('<div class="subhd">五、金额变动较大账户（' + res.changes.length + ' 户）</div>');
    h.push(simpleTable(
      [{ t: '公司名称' }, { t: '账号' }, { t: '银行/机构名称' }, { t: '上年余额', num: true }, { t: '本年余额', num: true }, { t: '变动额', num: true }, { t: '变动率', num: true }],
      res.changes.map(function (x) {
        return [x.r.company, x.r.acctNo, x.r.bankName, fm(x.r.priorLocal), fm(x.r.local), fm(x.delta), (x.rate * 100).toFixed(1) + '%'];
      })));

    $('#resPane-check').innerHTML = h.join('');
  }

  function simpleTable(cols, rows, cls) {
    var shown = rows.slice(0, 10);
    var h = ['<div class="scroll"><table><thead><tr>'];
    cols.forEach(function (c) { h.push('<th' + (c.num ? ' class="num"' : '') + '>' + esc(c.t) + '</th>'); });
    h.push('</tr></thead><tbody>');
    if (!shown.length) h.push('<tr><td colspan="' + cols.length + '">无此类账户。</td></tr>');
    shown.forEach(function (r) {
      h.push('<tr' + (cls ? ' class="' + cls + '"' : '') + '>');
      r.forEach(function (v, i) {
        var klass = (cols[i] && cols[i].num) ? 'num' : '';
        h.push('<td' + (klass ? ' class="' + klass + '"' : '') + '>' + esc(v) + '</td>');
      });
      h.push('</tr>');
    });
    h.push('</tbody></table></div>');
    if (rows.length > 10) h.push('<div class="tbl-note">仅显示前 10 行，共 ' + rows.length + ' 行。</div>');
    return h.join('');
  }

  /* ---------- 导出 ---------- */
  function exportXlsx() {
    if (!state.result || !state.result.url) { toast('请先生成', true); return; }
    window.location.href = state.result.url;
    toast('已开始下载');
  }

  /* ---------- 底稿模板管理 ---------- */
  function loadTpl() {
    fetch('/api/tpl').then(function (r) { return r.json(); }).then(function (d) {
      state.tpl = d;
      var el = $('#tplStatus');
      if (d.exists) {
        setStatus('#tplStatus', 'ok',
          '当前模板：<b>' + esc(d.filename) + '</b>（' + esc(d.version) + '） · ' + formatBytes(d.size));
      } else {
        setStatus('#tplStatus', 'err', '未找到内置模板，请点击「上传替换」上传 C SWP 模板');
      }
      if (d.note) $('#tplNote').textContent = d.note;
      var nm = $('#tplName');
      if (nm) nm.value = d.exists ? (d.filename || '') : '';
      var tag = $('#tplTag');
      if (tag) { tag.textContent = d.exists ? '已就绪' : '缺失'; tag.className = d.exists ? 'tag-on' : 'tag-off'; }
      updateGating();
    }).catch(function (e) {
      setStatus('#tplStatus', 'err', '无法连接本地服务：' + esc(e.message));
    });
  }
  function formatBytes(n) {
    if (n == null) return '';
    if (n < 1024) return n + ' B';
    if (n < 1048576) return (n / 1024).toFixed(1) + ' KB';
    return (n / 1048576).toFixed(1) + ' MB';
  }
  function initTpl() {
    var up = document.createElement('input');
    up.type = 'file'; up.accept = '.xlsx,.xlsm'; up.style.display = 'none';
    document.body.appendChild(up);
    $('#btnTplUpload').addEventListener('click', function () { up.value = ''; up.click(); });
    up.addEventListener('change', function () {
      if (!up.files || !up.files[0]) return;
      var fd = new FormData(); fd.append('tpl', up.files[0], up.files[0].name);
      setStatus('#tplStatus', 'loading', '正在上传替换模板…');
      fetch('/api/tpl/upload', { method: 'POST', body: fd }).then(function (r) { return r.json(); }).then(function (d) {
        if (d.error) { setStatus('#tplStatus', 'err', '上传失败：' + esc(d.error)); return; }
        toast('模板已替换'); loadTpl();
      }).catch(function (e) { setStatus('#tplStatus', 'err', '上传失败：' + esc(e.message)); });
    });
    $('#btnTplDownload').addEventListener('click', function () {
      window.location.href = '/api/tpl/download';
    });
  }

  /* ---------- 序时账自检 ---------- */
  function initProbe() {
    var b = $('#btnProbe');
    if (!b) return;
    b.addEventListener('click', function () {
      var out = $('#probeOut');
      b.disabled = true;
      out.innerHTML = '<div class="hint">正在读取序时账…</div>';
      fetch('/api/journal-probe').then(function (r) { return r.json(); }).then(function (d) {
        b.disabled = false;
        out.innerHTML = renderProbe(d);
      }).catch(function (e) {
        b.disabled = false;
        out.innerHTML = '<div class="tbl-note">自检失败：' + esc(e.message) + '</div>';
      });
    });
  }

  function renderProbe(d) {
    if (d.error) return '<div class="alert warn">' + esc(d.error) + '</div>';
    var h = ['<div class="tiles">' +
      '<div class="tile"><div class="k">序时账总行数</div><div class="v">' + fm(d.total || 0) + '</div><div class="s">journal 表</div></div>' +
      '<div class="tile"><div class="k">日期范围</div><div class="v" style="font-size:15px">' +
      esc((d.dateMin || '—') + ' ~ ' + (d.dateMax || '—')) + '</div><div class="s">用于截止期间过滤</div></div>' +
      '<div class="tile' + ((d.moneyRows || 0) ? ' ok' : ' warn') + '"><div class="k">识别为货币资金的行</div><div class="v">' +
      fm(d.moneyRows || 0) + '</div><div class="s">来自 ' + (d.moneyAccounts || 0) + ' 个科目</div></div>' +
      '<div class="tile' + (d.badDate ? ' warn' : '') + '"><div class="k">日期解析不了的行</div><div class="v">' +
      fm(d.badDate || 0) + '</div><div class="s">这些行会被跳过</div></div>' +
      '</div>'];

    var accs = d.topAccounts || [];
    h.push('<div class="subhd">科目分布（按行数前 ' + accs.length + ' 个，✓ 表示会被纳入截止性测试）</div>');
    if (!accs.length) {
      h.push('<div class="tbl-note">没有读到科目编码/名称列。</div>');
    } else {
      h.push('<div class="scroll"><table><thead><tr><th>科目编码</th><th>科目名称</th><th>一级科目</th>' +
        '<th class="num">行数</th><th>是否货币资金</th></tr></thead><tbody>');
      accs.forEach(function (a) {
        h.push('<tr><td class="mono">' + esc(a.code || '—') + '</td><td>' + esc(a.name || '—') + '</td>' +
          '<td>' + esc(a.l1 || '—') + '</td><td class="num">' + fm(a.count) + '</td>' +
          '<td>' + (a.money ? '<span class="tag-on">✓ 纳入</span>' : '—') + '</td></tr>');
      });
      h.push('</tbody></table></div>');
    }
    if (!(d.moneyRows || 0)) {
      h.push('<div class="tbl-note">没有一行被识别为货币资金科目，这就是样本为空的直接原因：' +
        '科目编码需以 1001/1002/1012 开头，或科目名称含「库存现金 / 银行存款 / 其他货币资金」。</div>');
    }
    return h.join('');
  }

  /* ---------- 账面信息表（按账套填官方「账面模板」） ---------- */
  var bookState = null;

  function initBook() {
    // 「账面信息表」这一块在新版页面里没有放，元素不存在。
    // 这里必须先短路：以前直接绑定，一旦有人调用 initBook，
    // 第一句就抛 null 错误，后面的初始化全部中断。
    if (!$('#btnBookGen')) return;
    $('#btnBookPreview').addEventListener('click', function () { bookPlan(); });
    $('#btnBookGen').addEventListener('click', bookGen);
    $('#btnBookTplDownload').addEventListener('click', function () {
      window.location.href = '/api/book/tpl/download';
    });
    var up = $('#bookTplFile');
    $('#btnBookTplUpload').addEventListener('click', function () { up.value = ''; up.click(); });
    up.addEventListener('change', function () {
      if (!up.files || !up.files[0]) return;
      var fd = new FormData();
      fd.append('tpl', up.files[0], up.files[0].name);
      setStatus('#bookStatus', 'loading', '正在上传替换账面模板…');
      fetch('/api/book/tpl/upload', { method: 'POST', body: fd })
        .then(function (r) { return r.json(); })
        .then(function (d) {
          if (d.error) { setStatus('#bookStatus', 'err', '上传失败：' + esc(d.error)); return; }
          setStatus('#bookStatus', 'ok', '模板已替换：' + esc(d.filename) + '（工作表 ' + esc(d.sheet) + '）');
        })
        .catch(function (e) { setStatus('#bookStatus', 'err', '上传失败：' + esc(e.message)); });
    });
    var d = new Date();
    $('#optBookDate').value = d.toISOString().slice(0, 10);
  }

  function bookPlan() {
    var kw = ($('#optBookKeyword').value || '银行存款').trim();
    setStatus('#bookStatus', 'loading', '正在读取科目余额表…');
    return fetch('/api/book/plan?kw=' + encodeURIComponent(kw))
      .then(function (r) { return r.json(); })
      .then(function (d) {
        if (d.error) {
          setStatus('#bookStatus', 'err', esc(d.error));
          $('#bookSummary').hidden = true;
          $('#bookMiss').hidden = true;
          bookState = null;
          return null;
        }
        bookState = d;
        renderBookPlan(d);
        return d;
      })
      .catch(function (e) {
        setStatus('#bookStatus', 'err', '读取失败：' + esc(e.message));
        return null;
      });
  }

  function renderBookPlan(d) {
    var m = d.meta || {}, rows = d.rows || [];
    setStatus('#bookStatus', 'ok',
      '可生成 <b>' + (m.rows || 0) + '</b> 行 · ' + (m.groups || 0) + ' 个账户 · '
      + '本币合计 <b>' + fm(m.balanced) + '</b>');
    var cur = {};
    rows.forEach(function (r) { cur[r.currency] = (cur[r.currency] || 0) + 1; });
    var curTxt = Object.keys(cur).map(function (k) { return k + ' ' + cur[k]; }).join(' / ');
    var sum = $('#bookSummary');
    sum.hidden = false;
    sum.innerHTML =
      '<div class="tile"><div class="k">可生成行数</div><div class="v">' + (m.rows || 0) + '</div></div>' +
      '<div class="tile"><div class="k">账户数</div><div class="v">' + (m.groups || 0) + '</div></div>' +
      '<div class="tile"><div class="k">本币合计</div><div class="v">' + fm(m.balanced) + '</div></div>' +
      '<div class="tile"><div class="k">缺银行名 / 账号</div><div class="v">'
        + (m.missBank || 0) + ' / ' + (m.missAcct || 0) + '</div></div>';
    var box = $('#bookMiss');
    var miss = rows.filter(function (r) { return r._missBank || r._missAcct; });
    if (!miss.length) { box.hidden = true; box.innerHTML = ''; return; }
    box.hidden = false;
    var h = ['<div class="tbl-note" style="color:var(--danger);font-size:12px">'
      + miss.length + ' 行的银行名称/账号取不到（账套里没有明细科目，也没有「银行账号」辅助核算），'
      + '生成后这几列会留空，需要人工补。下面只列前 5 行：</div>'
      + '<table><thead><tr>'
      + '<th>公司</th><th>科目</th><th>币种</th><th class="num">本币</th><th>缺</th>'
      + '</tr></thead><tbody>'];
    miss.slice(0, 5).forEach(function (r) {
      h.push('<tr><td>' + esc(r.company) + '</td><td>' + esc(r.subject) + '</td><td>'
        + esc(r.currency) + '</td><td class="num">' + fm(r.local) + '</td><td>'
        + (r._missBank ? '银行名 ' : '') + (r._missAcct ? '账号' : '') + '</td></tr>');
    });
    h.push('</tbody></table>');
    box.innerHTML = h.join('');
  }

  function bookGen() {
    var bd = $('#optBookDate').value;
    if (!bd) { toast('请先选「资产负债表日」', true); $('#optBookDate').focus(); return; }
    var kw = ($('#optBookKeyword').value || '银行存款').trim();
    setStatus('#bookStatus', 'loading', '正在生成账面信息表…');
    fetch('/api/book/generate', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ balance_date: bd, keyword: kw })
    }).then(function (r) { return r.json(); }).then(function (d) {
      if (d.error) { setStatus('#bookStatus', 'err', '生成失败：' + esc(d.error)); return; }
      renderBookPlan({ meta: d.stat, rows: d.rows });   // 生成接口把统计放在 stat 里
      setStatus('#bookStatus', 'ok',
        '已生成 <b>' + (d.stat.rows) + '</b> 行 → ' + esc(d.file) + '（浏览器会自动下载）');
      toast('账面信息表已生成');
      window.location.href = d.url;
    }).catch(function (e) {
      setStatus('#bookStatus', 'err', '生成失败：' + esc(e.message));
    });
  }

  /* ---------- 初始化 ---------- */
  document.addEventListener('DOMContentLoaded', function () {
    initDz();
    initTpl();
    initCutoff();
    loadBaseInfo();
    if ($('#btnSaveBase')) $('#btnSaveBase').addEventListener('click', saveBaseInfo);
    if ($('#btnExportBase')) $('#btnExportBase').addEventListener('click', exportBaseInfo);
    if ($('#btnImportBase')) $('#btnImportBase').addEventListener('click', function () {
      var f = $('#fileImportBase'); if (f) f.click();
    });
    if ($('#fileImportBase')) $('#fileImportBase').addEventListener('change', function (e) {
      var f = e.target.files && e.target.files[0];
      importBaseInfo(f);
      e.target.value = '';
    });
    initProbe();
    loadTpl();
    $('#btnReloadTb').addEventListener('click', loadTb);
    $('#btnReset').addEventListener('click', function () { location.reload(); });
    $$('#resTabs .tab').forEach(function (t) {
      t.addEventListener('click', function () {
        $$('#resTabs .tab').forEach(function (x) { x.classList.remove('active'); });
        this.classList.add('active');
        var p = this.dataset.pane;
        ['bkd', 'gaps', 'cutoff'].forEach(function (k) {
          var el = $('#resPane-' + k);
          if (el) el.hidden = (p !== k);
        });
      });
    });
    loadTb();

    /* ================= 页签：银行存款 / 费用 ================= */
    $$('#mainTabs .tab').forEach(function (t) {
      t.addEventListener('click', function () {
        $$('#mainTabs .tab').forEach(function (x) { x.classList.remove('active'); });
        this.classList.add('active');
        var p = this.dataset.mpane;
        ['wp_cash', 'wp_exp', 'wp_fa'].forEach(function (k) {
          var el = $('#mpane-' + k);
          if (el) el.hidden = (p !== k);
        });
        // 费用取数延后到这里：后端是一个 DuckDB 连接，两个 pane 同时请求会
        // 报 "connection is busy with another call"。切过去再拉，天然串行。
        if (p === 'wp_exp' && !state.expLoaded) loadExpTb();
        // 长期资产同理：切过去才请求，避免和货币资金的取数抢同一个连接
        if (p === 'wp_fa' && window.FA && !window.FA.loaded()) window.FA.load();
      });
    });

    // 从主框架的 ?sub=wp_cash / wp_exp 决定初始页签（导航里点「货币资金」直接到位）
    (function initSub() {
      var m = /[?&]sub=([^&]+)/.exec(location.search);
      var sub = m ? decodeURIComponent(m[1]) : '';
      if (!sub) return;
      var btn = document.querySelector('#mainTabs .tab[data-mpane="' + sub + '"]');
      if (btn) btn.click();
    })();

    /* ================= 一键上数：银行存款 ================= */
    var c00File = null;

    function todayStr() {
      var d = new Date();
      return d.getFullYear() + '-' + String(d.getMonth() + 1).padStart(2, '0') + '-' + String(d.getDate()).padStart(2, '0');
    }
    if ($('#optExpDate') && !$('#optExpDate').value) $('#optExpDate').value = todayStr();
    // 货币资金的资产负债表日不再默认「今天」：那是生成日期，不是资产负债表日，
    // 曾把 Lead 表头、C.02a 汇率日期、C.03 截止期间一起带偏。改为读共用基础信息。

    // 预审截止月：默认取资产负债表日前 3 个月（模板示例也是「预审 1-9 / 剩余 10-12」），
    // 只填空的时候补一次，用户改过就不动。
    function guessPreEnd() {
      var el = $('#optExpPre');
      var d = ($('#optExpDate') || {}).value || '';
      if (!el || el.value) return;
      var m = parseInt(String(d).slice(5, 7), 10);
      if (m > 3) el.value = String(m - 3);
    }
    guessPreEnd();
    if ($('#optExpDate')) $('#optExpDate').addEventListener('change', function () {
      var el = $('#optExpPre');
      if (el) el.value = '';
      guessPreEnd();
    });

    function renderGaps(boxSel, gaps, title) {
      var box = $(boxSel);
      if (!box) return;
      if (!gaps || !gaps.length) {
        box.hidden = true;
        box.innerHTML = '';
        return;
      }
      box.hidden = false;
      var byField = {};
      gaps.forEach(function (g) { byField[g.field] = (byField[g.field] || 0) + 1; });
      var chips = Object.keys(byField).map(function (k) {
        return '<span style="display:inline-block;background:#fff7e0;border:1px solid #f0dfc0;border-radius:12px;' +
          'padding:2px 10px;margin:0 6px 6px 0;font-size:12px">' + esc(k) + ' ' + byField[k] + '</span>';
      }).join('');
      var rows = gaps.slice(0, 200).map(function (g) {
        return '<tr><td>' + esc(g.company || '') + '</td><td>' + esc(g.subject || '') + '</td>' +
          '<td>' + esc(g.acct || '') + '</td><td>' + esc(g.field) + '</td>' +
          '<td>' + (g.col ? String.fromCharCode(64 + Number(g.col)) + g.row : '') + '</td>' +
          '<td>' + esc(g.reason || '') + '</td></tr>';
      }).join('');
      box.innerHTML = '<div style="font-weight:600;margin-bottom:8px">' + esc(title) +
        '（共 ' + gaps.length + ' 处）</div><div style="margin-bottom:8px">' + chips + '</div>' +
        '<div class="scroll"><table><thead><tr>' +
        '<th>公司</th><th>科目</th><th>账号/账簿</th><th>字段</th><th>单元格</th><th>原因</th>' +
        '</tr></thead><tbody>' + rows + '</tbody></table></div>' +
        (gaps.length > 200 ? '<div class="tbl-note">只显示前 200 条，完整清单见 Excel「检查」表。</div>' : '');
    }

    var c00Companies = [];

    // 折算率输入：列出全部主体（本位币为外币的可能币种标的是人民币，靠 hasFx 判断会漏），
    // 有外币账户的主体高亮提示，其余保持 1 即可
    function renderFx(d) {
      c00Companies = (d && d.companies) || c00Companies || [];
      var box = $('#c00Fx');
      if (!box) return;
      if (!c00Companies.length) {
        box.hidden = true;
        $('#c00FxRows').innerHTML = '';
        return;
      }
      box.hidden = false;
      $('#c00FxRows').innerHTML = c00Companies.map(function (c, i) {
        var cur = (c.currencies || []).join('/');
        return '<label' + (c.hasFx ? ' style="font-weight:600"' : '') + '>折算率 · ' +
          esc(c.company) +
          (cur ? ' <span class="hint">' + esc(cur) + (c.hasFx ? ' ⚠ 含外币' : '') + '</span>' : '') +
          '<input type="number" step="0.000001" min="0" data-fx="' + i +
          '" data-company="' + esc(c.company) + '" value="1"/></label>';
      }).join('');
    }

    function fxRates() {
      var out = {};
      var rows = document.querySelectorAll('#c00FxRows input[data-company]');
      [].forEach.call(rows, function (el) {
        var v = Number(el.value);
        out[el.getAttribute('data-company')] = (isFinite(v) && v > 0) ? v : 1;
      });
      return out;
    }

    function c00Body() {
      var ps = priorSpec() || {};
      return JSON.stringify({
        balance_date: ($('#optC00Date') || {}).value || '',
        customer: ($('#optC00Customer') || {}).value || '',
        rates: fxRates(),
        prior_fid: ps.fid || undefined,
        prior_spec: ps.spec || undefined,
      });
    }

    function c00Tiles(d) {
      return tile('账户数', d.rows, 'ok') +
        tile('标黄待补', d.gaps, d.gaps ? 'warn' : 'ok') +
        tile('本币合计', fm(d.amount), '') +
        (d.fxCount ? tile('折算人民币合计', fm(d.rmbAmount), 'ok') : '') +
        (d.tpl ? tile('模板', esc(d.tpl), '') : '');
    }

    /* 读 TB 之后顺手跑一次预览（不落盘）：只为把折算率主体、待补格数备好 */
    function c00Plan() {
      return fetch('/api/build/c00/plan').then(function (r) { return r.json(); }).then(function (d) {
        if (d.error) { state.c00plan = null; return null; }
        state.c00plan = d;
        renderFx(d);
        renderPreflight();
        if (tbReady()) flowSet('cash', 2, 3);
        return d;
      }).catch(function () { return null; });
    }

    /* 一键生成：BKD + Lead + Cutoff + 检查表，主底稿失败就停（用户必须知道）。 */
    function genAll() {
      var btn = $('#btnGenAll');
      var old = btn.textContent;
      btn.disabled = true;
      btn.textContent = '生成中…';
      setStatus('#genStatus', 'loading', '正在生成 C.00 BKD + Lead + Cutoff…');
      $('#genWarn').innerHTML = '';
      flowSet('cash', 4, 4);

      var body = JSON.parse(c00Body());
      body.subject = 'cash';
      body.cutoff = ($('#optCutoff') || {}).checked === false ? { enabled: false } : {
        enabled: true,
        date: ($('#optCutoffDate') || {}).value || '',
        before: ($('#optCutoffBefore') || {}).value || 1,
        after: ($('#optCutoffAfter') || {}).value || 1,
        threshold: ($('#optCutoffAmt') || {}).value || 100000
      };
      if (body.cutoff.enabled && !body.cutoff.date) {
        body.cutoff.enabled = false;
      }

      fetch('/api/build/wp/run', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      }).then(function (r) { return r.json(); }).then(function (d) {
        btn.disabled = false; btn.textContent = old;
        if (d.error) {
          setStatus('#genStatus', 'err', '生成失败：' + esc(d.error));
          flowSet('cash', 3, 4);
          return;
        }
        renderGenResult(d);
        var st = d.main || {};
        setStatus('#genStatus', 'ok',
          '已生成 <b>' + esc(st.file || '') + '</b>：' + st.rows + ' 个账户，<b>' +
          st.gaps + '</b> 格待人工确认' +
          (st.fxCount ? '，' + st.fxCount + ' 个账户已按折算率换算' : ''));
        toast('底稿已生成');
        $('#genResult').scrollIntoView({ behavior: 'smooth', block: 'start' });
      }).catch(function (e) {
        btn.disabled = false; btn.textContent = old;
        setStatus('#genStatus', 'err', '请求失败：' + esc(e.message));
      });
    }

    function renderGenResult(d) {
      var st = d.main || {};
      $('#genResult').hidden = false;
      $('#c00Tiles').innerHTML = c00Tiles(st);

      var files = [];
      if (st.url) {
        files.push('<a class="filebtn" href="' + st.url + '">↓ 主底稿 ' + esc(st.file || '') +
          ' <span class="sz">BKD + Lead + Cutoff</span></a>');
      }
      $('#genFiles').innerHTML = files.join('');

      var rows = st.previewRows || [];
      if (rows.length) {
        var h = ['<div class="scroll"><table><thead><tr><th>公司</th><th>科目</th><th>银行/机构</th>' +
          '<th>账号</th><th>币种</th><th class="num">期末原币</th><th class="num">期末本币</th>' +
          '<th class="num">折算人民币</th><th class="num">上年余额</th></tr></thead><tbody>'];
        rows.forEach(function (r) {
          h.push('<tr><td>' + esc(r.company) + '</td><td>' + esc(r.subject) + '</td><td>' +
            esc(r.bank || '') + '</td><td class="mono">' + esc(r.acctNo || '—') + '</td><td>' +
            esc(r.currency || '') + '</td><td class="num">' + fm(r.orig) + '</td><td class="num">' +
            fm(r.local) + '</td><td class="num">' + fm(r.rmb) + '</td><td class="num">' +
            (r.priorLocal == null ? '<span class="hint">—</span>' : fm(r.priorLocal)) + '</td></tr>');
        });
        h.push('</tbody></table></div>');
        if ((st.rows || 0) > rows.length) {
          h.push('<div class="tbl-note">仅显示前 10 行，共 ' + st.rows + ' 行；完整内容在 Excel 里。</div>');
        }
        $('#resPane-bkd').innerHTML = h.join('');
      } else {
        $('#resPane-bkd').innerHTML = '<div class="hint">没有可预览的行。</div>';
      }

      renderGaps('#c00Gaps', st.gapList || [], '待确认清单（标黄格）');
      renderCutoffBox(st.cutoff);

      if ((d.warnings || []).length) {
        $('#genWarn').innerHTML = d.warnings.map(function (w) {
          return '<div class="alert warn">' + esc(w) + '</div>';
        }).join('');
      }
    }

    function renderCutoffBox(c) {
      var box = $('#cutoutOut');
      if (!box) return;
      if (!c || !c.enabled) {
        box.innerHTML = '<div class="hint">本次未执行截止性测试（第 3 步可以打开）。</div>';
        return;
      }
      var s = c.stat || {};
      box.innerHTML = '<div class="tiles">' +
        tile('期末前样本', (c.before || []).length, '') +
        tile('期末后样本', (c.after || []).length, '') +
        tile('跨期转账', s.crossCount || 0, (s.crossCount ? 'warn' : 'ok')) +
        tile('低于阈值未纳入', s.belowThreshold || 0, '') +
        '</div>' + (c.warn ? '<div class="alert warn">' + esc(c.warn) + '</div>' : '');
    }

    state.onTbReady = c00Plan;
    $('#btnGenAll').addEventListener('click', genAll);

    /* ================= 费用底稿 ================= */
    var expFile = null;

    function tile(k, v, cls) {
      return '<div class="tile ' + (cls || '') + '"><div class="k">' + k + '</div><div class="v">' + v + '</div></div>';
    }

    /* ---------- 费用 CRA（销管费用）：四档百分比 + 四项认定 ---------- */
    var EXP_CRA = { types: [], asserts: [] };

    function expCraPctEls() {
      return { Minimal: $('#expCraPctMinimal'), Low: $('#expCraPctLow'),
               Moderate: $('#expCraPctModerate'), High: $('#expCraPctHigh') };
    }

    function renderExpCraAssertions(cur) {
      var box = $('#expCraAssertRows');
      if (!box) return;
      box.innerHTML = '';
      EXP_CRA.asserts.forEach(function (a) {
        var lab = document.createElement('label');
        lab.textContent = a.label;
        var sel = document.createElement('select');
        sel.id = 'expCraA_' + a.key;
        var o0 = document.createElement('option');
        o0.value = ''; o0.textContent = '（不写，保留手选）';
        sel.appendChild(o0);
        EXP_CRA.types.forEach(function (t) {
          var o = document.createElement('option');
          o.value = t; o.textContent = t;
          sel.appendChild(o);
        });
        sel.value = (cur || {})[a.key] || '';
        lab.appendChild(sel);
        box.appendChild(lab);
      });
    }

    function loadExpCra() {
      fetch('/api/build/baseinfo').then(function (r) { return r.json(); }).then(function (d) {
        var mod = '销管费用';
        var cfg = (d && d.cra && d.cra[mod]) || {};
        var dflt = (d && d.craDefaults && d.craDefaults[mod]) || {};
        EXP_CRA.types = (d && d.craTypes) || ['Minimal', 'Low', 'Moderate', 'High', 'N/A'];
        EXP_CRA.asserts = ((d && d.assertionsByModule && d.assertionsByModule[mod]) || []);
        var pcts = cfg.pcts || {};
        var els = expCraPctEls();
        Object.keys(els).forEach(function (k) {
          if (!els[k]) return;
          var v = (pcts[k] != null) ? pcts[k] : dflt[k];
          els[k].value = (v == null) ? '' : Math.round(v * 1000) / 10;
        });
        renderExpCraAssertions(cfg.assertions || {});
        var ed = $('#optExpDate');
        if (ed && !ed.value && d && d.info && d.info.balance_date) ed.value = d.info.balance_date;
      }).catch(function () { });
    }

    function expCraBody() {
      var pcts = {}, els = expCraPctEls();
      Object.keys(els).forEach(function (k) {
        if (!els[k]) return;
        var raw = String(els[k].value || '').replace('%', '').trim();
        if (raw === '') return;
        var num = parseFloat(raw);
        if (!isNaN(num)) pcts[k] = num / 100;
      });
      var assertions = {};
      EXP_CRA.asserts.forEach(function (a) {
        var el = $('#expCraA_' + a.key);
        if (el && el.value) assertions[a.key] = el.value;
      });
      return { pcts: pcts, assertions: assertions };
    }

    function saveExpCra() {
      fetch('/api/build/baseinfo/save', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({ cra: { '销管费用': expCraBody() } })
      }).then(function (r) { return r.json(); }).then(function (d) {
        toast((d && d.ok) ? '已保存：销管费用的 CRA 配置' : ((d && d.error) || '保存失败'), !(d && d.ok));
      }).catch(function () { toast('保存失败（网络）', true); });
    }

    function expBody() {
      return JSON.stringify({
        period_from: ($('#optExpFrom') || {}).value || '1',
        period_to: ($('#optExpTo') || {}).value || '6',
        balance_date: ($('#optExpDate') || {}).value || '',
        customer: ($('#optExpCustomer') || {}).value || '',
        te: ($('#optExpTe') || {}).value || '',
        sad: ($('#optExpSad') || {}).value || '',
        // 预审截止月：TOD「预审+剩余期间」页按它切两段；留空则剩余期间为空
        pre_end: ($('#optExpPre') || {}).value || '',
        // 会计准则 / 本位币：项目组最清楚，平台直接收，省得 AI 每次猜
        standard: ($('#optExpStd') || {}).value || '',
        currency: ($('#optExpCcy') || {}).value || '',
        // 四项认定等级：从右侧 CRA 面板取（空着走 Minimal）
        // 空着会让 BKD E8=0 → C22=0 → 「进一步调查」恒为"是"
        assurance: (function () {
          var a = expCraBody().assertions || {};
          return [a.completeness, a.occurrence, a.measurement, a.presentation]
            .map(function (x) { return x || 'Minimal'; });
        })(),
        // CRA 百分比：后端按它重写两张 BKD 的 D8:D11 公式
        cra: expCraBody(),
        prior_fid: state.expPriorFid || '',
        // 上年底稿的 VC / VD 各用哪张表（两个下拉选的）
        prior_sheets: state.expPriorSheets || [],
        // 上一期 TOD：文件 + 手动指定的工作表（预审段引用它，两期重要性水平不同）
        prior_tod_fid: state.expPriorTodFid || '',
        prior_tod_sheet: state.expPriorTodSheet || ''
      });
    }

    function expQuery() {
      var b = JSON.parse(expBody());
      return '?from=' + encodeURIComponent(b.period_from) + '&to=' + encodeURIComponent(b.period_to);
    }

    function expTilesFrom(d) {
      var o = d.oar || {};
      var bl = (o.blocks || []).map(function (b) { return b.label + ' ' + b.rows + ' 项'; }).join(' / ');
      return tile('销售费用行数', d.vc, 'ok') +
        tile('管理费用行数', d.vd, 'ok') +
        tile('本期账面数合计', fm((d.vc_sum || 0) + (d.vd_sum || 0)), '') +
        tile('OAR 口径', o.items != null ? o.items : '—',
          bl || '薪酬/折旧摊销/广告费/单列大额/其他', o.items ? 'ok' : '') +
        tile('需分析口径', o.analyzed != null ? o.analyzed : '—',
          o.wave ? '波动幅度 ' + fm(o.wave) : '', o.analyzed ? 'warn' : 'ok') +
        tile('标黄待补', d.gaps, d.gaps ? 'warn' : 'ok');
    }

    /* 第 1 步：读费用发生额，展示前 10 行 */
    function loadExpTb() {
      setStatus('#expTbStatus', 'loading', '正在从序时账汇总费用发生额…');
      fetch('/api/build/exp/tb' + expQuery()).then(function (r) { return r.json(); }).then(function (d) {
        if (d.error) {
          setStatus('#expTbStatus', 'err', '读取失败：' + esc(d.error));
          return;
        }
        setStatus('#expTbStatus', 'ok', '已汇总 <b>' + d.rows + '</b> 行，涉及 <b>' +
          d.books + '</b> 个账簿，期间 ' + esc(d.period || ''));
        $('#expTbSummary').innerHTML =
          '<div class="tile"><div class="k">科目×账簿行数</div><div class="v">' + d.rows + '</div></div>' +
          '<div class="tile"><div class="k">本期发生额合计</div><div class="v">' + fm(d.sum) + '</div></div>' +
          '<div class="tile"><div class="k">已剔除结转损益</div><div class="v">' + fm(d.closeOut) + '</div></div>';
        var rows = d.sample || [];
        var h = ['<table><thead><tr><th>账簿</th><th>科目编码</th><th>科目名称</th>' +
          '<th class="num">本期发生额</th></tr></thead><tbody>'];
        rows.forEach(function (r) {
          h.push('<tr><td>' + esc(r.book) + '</td><td class="mono">' + esc(r.code) +
            '</td><td>' + esc(r.name) + '</td><td class="num">' + fm(r.amount) + '</td></tr>');
        });
        h.push('</tbody></table>');
        if (d.rows > rows.length) {
          h.push('<div class="tbl-note">仅显示前 10 行，共 ' + d.rows + ' 行。</div>');
        }
        $('#expTbTable').innerHTML = h.join('');
        state.expRows = d.rows || 0;
        state.expLoaded = true;
        flowSet('exp', 1, 2);
        renderExpPreflight();
      }).catch(function (e) {
        setStatus('#expTbStatus', 'err', '请求失败：' + esc(e.message));
      });
    }

    function renderExpPreflight() {
      var box = $('#expPreflight');
      if (!box) return;
      var has = state.expRows > 0;
      var items = [
        { ok: has, must: true,
          txt: has ? '费用发生额已读取 <b>' + state.expRows + '</b> 行'
                   : '还没读到费用发生额 —— 检查期间设置与序时账是否已合并' },
        { ok: true, must: true, txt: '底稿模板已就绪：<b>U_exp SWP VC&amp;VD 模板.xlsx</b>' },
        { ok: true, must: false,
          txt: '账簿→公司映射取自导入阶段的主体 mapping（不用再传对照表）' },
        { ok: !!(state.expPriorSheets || []).length, must: false,
          txt: (state.expPriorSheets || []).length
            ? '上年底稿已选表：<b>' + esc((state.expPriorSheets || []).join(' / ')) + '</b>'
            : '未上传上一期底稿（可选）—— K 列留空标黄' },
        { ok: !!(state.expPriorTodFid && state.expPriorTodSheet), must: false,
          txt: state.expPriorTodSheet
            ? '上一期 TOD：<b>' + esc(state.expPriorTodSheet) + '</b>（预审段引用它）'
            : '上一期底稿里没用 TOD 页（可选）—— 预审段按本期阈值从序时账算' }
      ];
      box.innerHTML = items.map(function (it) {
        var cls = it.ok ? 'ok' : (it.must ? 'no' : 'opt');
        var ic = it.ok ? '✓' : (it.must ? '✕' : '—');
        return '<div class="pf ' + cls + '"><span class="ic">' + ic + '</span><span>' + it.txt + '</span></div>';
      }).join('');
      var b = $('#btnExpGen');
      if (b) b.disabled = !has;
    }

    /* 第 4 步：一键生成费用底稿 */
    function expGen() {
      var btn = $('#btnExpGen');
      var old = btn.textContent;
      btn.disabled = true;
      btn.textContent = '生成中…';
      setStatus('#expGenStatus', 'loading', '正在生成 VC.00 / VD.00 费用底稿…');
      flowSet('exp', 4, 4);
      var body = JSON.parse(expBody());
      body.subject = 'exp';
      fetch('/api/build/wp/run', {
        method: 'POST', headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify(body)
      }).then(function (r) { return r.json(); }).then(function (d) {
        btn.disabled = false; btn.textContent = old;
        if (d.error) {
          setStatus('#expGenStatus', 'err', '生成失败：' + esc(d.error));
          flowSet('exp', 3, 4);
          return;
        }
        var m = d.main || {};
        $('#expGenResult').hidden = false;
        var _extra = m.priorError
          ? '，<b style="color:#c00">上期数未回填：' + esc(m.priorError) + '</b>'
          : (state.expPriorFid ? '，上期审定数已回填' : '');
        setStatus('#expGenStatus', 'ok', '已生成 <b>' + esc(m.file || '') + '</b>：VC ' +
          m.vc + ' 行 / VD ' + m.vd + ' 行，<b>' + m.gaps + '</b> 格待人工确认' + _extra);
        $('#expTiles').innerHTML = expTilesFrom(m);
        $('#expFiles').innerHTML = m.url
          ? '<a class="filebtn" href="' + m.url + '">↓ 费用底稿 ' + esc(m.file || '') +
            ' <span class="sz">VC + VD + 检查表</span></a>' : '';
        renderGaps('#expGaps', m.gapList || [], '待确认清单（标黄格）');
        toast('费用底稿已生成');
        $('#expGenResult').scrollIntoView({ behavior: 'smooth', block: 'start' });
      }).catch(function (e) {
        btn.disabled = false; btn.textContent = old;
        setStatus('#expGenStatus', 'err', '请求失败：' + esc(e.message));
      });
    }

    if ($('#btnReloadExp')) $('#btnReloadExp').addEventListener('click', loadExpTb);
    if ($('#btnExpGen')) $('#btnExpGen').addEventListener('click', expGen);

    /* 账簿 → 公司：直接读导入阶段写进 _books 的主体 mapping，不再让项目组传对照表 */

    /* ---------- 上年度费用底稿：拖拽 + VC/VD 两个下拉 ---------- */
    function initExpPrior() {
      var dz = $('#dz-exp-prior');
      if (!dz) return;
      var input = document.createElement('input');
      input.type = 'file';
      input.accept = '.xlsx,.xlsm,.xls';
      input.style.display = 'none';
      document.body.appendChild(input);
      dz.addEventListener('click', function () { input.value = ''; input.click(); });
      input.addEventListener('change', function () {
        if (input.files && input.files[0]) handleExpPrior(input.files[0]);
      });
      dz.addEventListener('dragover', function (e) { e.preventDefault(); dz.classList.add('over'); });
      dz.addEventListener('dragleave', function () { dz.classList.remove('over'); });
      dz.addEventListener('drop', function (e) {
        e.preventDefault(); dz.classList.remove('over');
        var f = e.dataTransfer.files && e.dataTransfer.files[0];
        if (f) handleExpPrior(f);
      });
    }

    function handleExpPrior(file) {
      var dz = $('#dz-exp-prior');
      if (dz) dz.querySelector('.dz-file').textContent = '正在解析 ' + file.name + ' …';
      var fd = new FormData();
      fd.append('prior', file, file.name);
      fetch('/api/build/exp/prior/plan', { method: 'POST', body: fd })
        .then(function (r) { return r.json(); })
        .then(function (d) {
          if (d.error) {
            state.expPriorFid = '';
            state.expPriorData = null;
            state.expPriorTodFid = '';
            state.expPriorTodData = null;
            state.expPriorTodSheet = '';
            toast('解析失败：' + d.error, true);
            if (dz) dz.querySelector('.dz-file').textContent = '尚未上传';
            return;
          }
          state.expPriorFid = d.fid || '';
          state.expPriorData = d;
          state.expPriorVC = d.recVC;
          state.expPriorVD = d.recVD;
          // 上一期 TOD 就在这同一份文件里：不用再传一次，直接把探测结果接上
          state.expPriorTodFid = d.fid || '';
          state.expPriorTodData = d.tod || null;
          state.expPriorTodSheet = '';
          var _tod = state.expPriorTodData || {};
          var _rec = (_tod.sheets || [])[_tod.recommended];
          if (_rec && _rec.usable) state.expPriorTodSheet = _rec.sheet;
          if (dz) {
            dz.querySelector('.dz-file').textContent = file.name + ' · ' + d.sheets.length + ' 个工作表';
            dz.classList.add('loaded');
          }
          var tag = $('#expPriorTag');
          if (tag) { tag.className = 'tag-on'; tag.textContent = '已载入'; }
          renderExpPriorPane();
        })
        .catch(function (e) { toast('上传失败：' + e.message, true); });
    }

    function renderExpPriorPane() {
      var box = $('#expPriorPane');
      var d = state.expPriorData;
      if (!box || !d) return;
      var usable = [];
      (d.sheets || []).forEach(function (s, i) { if (s.usable) usable.push({ s: s, i: i }); });
      var h = [];
      function kindTxt(k) {
        return k === 'vd' ? '管理费用' : (k === 'vc' ? '销售费用' : (k === 'both' ? '两类都有' : ''));
      }
      if (usable.length) {
        function opts(pick) {
          return usable.map(function (x) {
            var s = x.s;
            return '<option value="' + esc(s.sheet) + '"' + (x.i === pick ? ' selected' : '') + '>' +
              esc(s.sheet) + '（' + s.rows + ' 行' + (kindTxt(s.kind) ? ' · ' + kindTxt(s.kind) : '') + '）</option>';
          }).join('');
        }
        h.push(
          '<div class="pripsel">' +
          '<label>销售费用用哪张表<select id="expPriorVC">' + opts(d.recVC) + '</select>' +
          '<span class="meta" id="expPriorVCMeta"></span></label>' +
          '<label>管理费用用哪张表<select id="expPriorVD">' + opts(d.recVD) + '</select>' +
          '<span class="meta" id="expPriorVDMeta"></span></label>' +
          '</div>');
      } else {
        h.push('<div class="hint">这份文件里没找到能读出「上期审定数」的 BKD 明细表，' +
          '上期数将留空标黄；若文件里只有 TOD 页，下面的 TOD 仍可直接用。</div>');
        state.expPriorSheets = [];
      }
      // TOD 页：与 BKD 同一个文件，选表即可
      var tod = d.tod || {};
      var tus = [];
      (tod.sheets || []).forEach(function (s, i) { if (s.usable) tus.push({ s: s, i: i }); });
      if (tus.length) {
        h.push('<div class="pripsel"><label>上一期 TOD 用哪张表<select id="expPriorTodSheet">');
        tus.forEach(function (x) {
          h.push('<option value="' + esc(x.s.sheet) + '"' +
            (x.i === tod.recommended ? ' selected' : '') + '>' +
            esc(x.s.sheet) + '（' +
            (x.s.kind === 'split' ? '预审+剩余（取合计列）' : '年审（单列）') +
            ' · 有数 ' + x.s.hasValues + ' 格）</option>');
        });
        h.push('</select><span class="meta" id="expPriorTodMeta"></span></label></div>');
      } else {
        h.push('<div class="hint">这份文件里没有可识别的 TOD 页（要有「本期全部费用交易」标签），' +
          '本期预审段按本期阈值从序时账算。</div>');
      }
      box.innerHTML = h.join('');
      if ($('#expPriorVC')) {
        $('#expPriorVC').onchange = syncExpPriorSheets;
        $('#expPriorVD').onchange = syncExpPriorSheets;
        syncExpPriorSheets();
      } else {
        renderExpPreflight();
      }
      if ($('#expPriorTodSheet')) {
        $('#expPriorTodSheet').onchange = syncExpPriorTodSheet;
        syncExpPriorTodSheet();
      }
    }

    function syncExpPriorSheets() {
      var vc = $('#expPriorVC'), vd = $('#expPriorVD');
      state.expPriorSheets = [vc && vc.value, vd && vd.value].filter(Boolean);
      var byName = {};
      ((state.expPriorData || {}).sheets || []).forEach(function (s) { byName[s.sheet] = s; });
      function meta(sel, mid) {
        var el = $(mid);
        var s = byName[(sel && sel.value) || ''];
        if (el) el.textContent = s ? ('能读到上期审定数 ' + s.rows + ' 行') : '';
      }
      meta(vc, '#expPriorVCMeta');
      meta(vd, '#expPriorVDMeta');
      renderExpPreflight();
    }

    function syncExpPriorTodSheet() {
      var sel = $('#expPriorTodSheet');
      state.expPriorTodSheet = (sel && sel.value) || '';
      var meta = $('#expPriorTodMeta');
      var s = null;
      ((state.expPriorTodData || {}).sheets || []).forEach(function (x) {
        if (x.sheet === state.expPriorTodSheet) s = x;
      });
      if (meta) {
        meta.textContent = s
          ? (s.kind === 'split'
              ? '本期预审段将取该页「合计」列（预审 + 剩余 = 全期）'
              : '本期预审段将取该页金额列（覆盖全期）')
          : '';
      }
      renderExpPreflight();
    }

    initExpPrior();
    loadExpCra();
    if ($('#btnExpCraSave')) $('#btnExpCraSave').addEventListener('click', saveExpCra);


    /* ================= AI 能力：上年 WP 结构探测包 ================= */
    if ($('#btnPriorProbe')) {
      $('#btnPriorProbe').addEventListener('click', function () {
        var f = state._priorFile;
        var box = $('#priorAiBox');
        if (!f) { toast('先在上面上传上年底稿', true); return; }
        box.hidden = false;
        box.innerHTML = '<div class="hint">正在生成结构探测包…</div>';
        fetch('/api/upload', {
          method: 'POST', body: f,
          headers: { 'X-Filename': encodeURIComponent(f.name), 'Content-Type': 'application/octet-stream' }
        })
          .then(function (r) { return r.json(); })
          .then(function (up) {
            if (!up.file_id) throw new Error(up.error || '上传失败');
            return fetch('/api/build/prior/probe', {
              method: 'POST', headers: { 'Content-Type': 'application/json' },
              body: JSON.stringify({ fid: up.file_id })
            });
          })
          .then(function (r) { return r.json(); })
          .then(function (d) {
            if (d.error) {
              box.innerHTML = '<div class="status err">' + esc(d.error) + '</div>';
              return;
            }
            var tb = (d.sheets || []).slice(0, 10).map(function (s) {
              return '<tr><td>' + esc(s.name) + '</td><td>' + (s.rows || '') + '</td><td>' + (s.cols || '') +
                '</td><td>' + (s.score || 0) + '</td><td>' + (s.header_row || '') + '</td></tr>';
            }).join('');
            box.innerHTML =
              '<div style="font-weight:600;margin-bottom:8px">结构探测结果（分数越高越像 BKD 明细表）</div>' +
              '<div class="scroll"><table><thead><tr><th>工作表</th><th>行数</th><th>列数</th>' +
              '<th>表头匹配分</th><th>推测表头行</th></tr></thead><tbody>' + tb + '</tbody></table></div>' +
              '<div class="hint mt">把下面整段复制给 AI，让它回一份 <code>{sheet, header_row}</code>，' +
              '再带回平台解析。规则见 <code>03_文档/03_底稿_费用.md</code>。</div>' +
              '<textarea id="probeMd" readonly style="width:100%;height:260px;margin-top:6px;' +
              'font-family:Consolas,Menlo,monospace;font-size:12px;line-height:1.5">' +
              esc(d.markdown || '') + '</textarea>' +
              '<div class="mt"><button class="btn" id="btnCopyMd">复制全部</button></div>';
            var btn = $('#btnCopyMd');
            if (btn) {
              btn.onclick = function () {
                var ta = $('#probeMd');
                ta.select();
                try { document.execCommand('copy'); toast('已复制'); }
                catch (e) { toast('复制失败，请手动全选', true); }
              };
            }
          })
          .catch(function (e) {
            box.innerHTML = '<div class="status err">' + esc(e.message) + '</div>';
          });
      });
    }
  });
})();
