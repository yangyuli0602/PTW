/* 固定资产（长期资产）· SWP 底稿前端
   与货币资金/费用共用 bkd.html，独立文件便于单独改动。
   V1.2.98 起跑的是 SWP 引擎（modules/fa_swp），旧的 K1_100Breakdown 链路已下线。 */
(function () {
  'use strict';
  var $ = function (s) { return document.querySelector(s); };
  var S = { plan: null, loaded: false, busy: false };

  function esc(s) {
    return String(s == null ? '' : s).replace(/&/g, '&amp;').replace(/</g, '&lt;')
      .replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  }
  function fm(n) {
    if (n == null || n === '') return '—';
    var v = Number(n);
    if (!isFinite(v)) return '—';
    return v.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  function wan(n) {
    var v = Number(n || 0);
    if (!isFinite(v)) return '—';
    return (v / 10000).toLocaleString('zh-CN', { maximumFractionDigits: 2 }) + ' 万';
  }
  function setStatus(sel, cls, html) {
    var e = $(sel);
    if (!e) return;
    if (!html) { e.className = 'status'; e.innerHTML = ''; return; }
    e.className = 'status ' + cls;
    e.innerHTML = html;
  }
  function toast(msg, err) {
    var w = $('#toastWrap');
    if (!w) return;
    var d = document.createElement('div');
    d.className = 'toast' + (err ? ' err' : '');
    d.textContent = msg;
    w.appendChild(d);
    setTimeout(function () { d.classList.add('on'); }, 10);
    setTimeout(function () {
      d.classList.remove('on');
      setTimeout(function () { d.remove(); }, 220);
    }, 2800);
  }
  function tile(k, v, s, cls) {
    return '<div class="tile ' + (cls || '') + '"><div class="k">' + esc(k) + '</div>' +
      '<div class="v">' + esc(v) + '</div>' +
      (s ? '<div class="s">' + esc(s) + '</div>' : '') + '</div>';
  }
  function sum(a) {
    var t = 0;
    (a || []).forEach(function (x) { t += Number(x.amt || 0); });
    return t;
  }

  /* ---------- ① 读取 ---------- */
  function load() {
    if (S.busy) return;
    S.busy = true;
    setStatus('#faStatus', 'loading', '正在读取科目余额表与序时账…');
    fetch('/api/build/fa_swp/plan').then(function (r) { return r.json(); }).then(function (d) {
      S.busy = false;
      S.loaded = true;
      if (d.error) {
        setStatus('#faStatus', 'err', esc(d.error));
        if ($('#faTiles')) $('#faTiles').innerHTML = '';
        if ($('#faCats')) $('#faCats').innerHTML = '';
        if ($('#faWarn')) $('#faWarn').innerHTML = '';
        return;
      }
      S.plan = d;
      setStatus('#faStatus', 'ok',
        '已读取：TB <b>' + d.rows.tb + '</b> 行 / 固定资产序时账 <b>' + d.rows.journal +
        '</b> 行 / 识别出 <b>' + d.cats.length + '</b> 个资产类别');
      render(d);
    }).catch(function (e) {
      S.busy = false;
      setStatus('#faStatus', 'err', '读取失败：' + esc(e.message));
    });
  }

  function render(d) {
    var dedupeN = (d.dedupe && d.dedupe.n) || 0;
    $('#faTiles').innerHTML =
      tile('资产类别', d.cats.length, '决定表1 有几个列块',
        d.cats.length > 7 ? 'warn' : 'ok') +
      tile('原值期末', wan(d.totals.fa_end), '期初 ' + wan(d.totals.fa_beg)) +
      tile('净值期末', wan(d.net.end), '期初 ' + wan(d.net.beg)) +
      tile('变动方式', (d.ways || []).length, '决定表1 有几行变动') +
      tile('TB 重复行', dedupeN, dedupeN ? '已按七元组去重' : '未发现重复',
        dedupeN ? 'warn' : 'ok');

    var rows = d.cats || [];
    var h = ['<table class="tbl"><thead><tr>',
      '<th>资产类别</th><th class="r">期初原值</th><th class="r">期末原值</th>',
      '<th class="r">期末累计折旧</th><th class="r">期末减值</th><th class="r">期末净值</th>',
      '</tr></thead><tbody>'].join('');
    rows.forEach(function (c) {
      h += '<tr><td>' + esc(c.cat) + '</td>' +
        '<td class="r">' + fm(c.faBeg) + '</td>' +
        '<td class="r">' + fm(c.faEnd) + '</td>' +
        '<td class="r">' + fm(c.depEnd) + '</td>' +
        '<td class="r">' + fm(c.impEnd) + '</td>' +
        '<td class="r">' + fm(c.nbvEnd) + '</td></tr>';
    });
    if (rows.length) {
      h += '<tr><td><b>合计</b></td>' +
        '<td class="r"><b>' + fm(d.totals.fa_beg) + '</b></td>' +
        '<td class="r"><b>' + fm(d.totals.fa_end) + '</b></td>' +
        '<td class="r"><b>' + fm(d.totals.dep_end) + '</b></td>' +
        '<td class="r"><b>' + fm(d.totals.imp_end) + '</b></td>' +
        '<td class="r"><b>' + fm(d.net.end) + '</b></td></tr>';
    }
    h += '</tbody></table>';
    $('#faCats').innerHTML = rows.length ? h : '<div class="hint">没有识别到资产类别</div>';

    warn(d);
  }

  function warn(d) {
    var box = $('#faWarn');
    var parts = [];

    var u = d.catUnion || {};
    if (!u.ok && ((u.onlyJnl || []).length)) {
      parts.push('<div class="status err"><b>序时账里有 ' + u.onlyJnl.length +
        ' 个类别，科目余额表里没有</b>：' + esc((u.onlyJnl || []).join('、')) +
        '<br/>这类分叉会让期末必然对不上（年初来自 TB、变动来自序时账），' +
        '而表现只是「差异若干元」。先回清洗把两边分类口径对齐，再生成。</div>');
    }
    if (!u.ok && ((u.onlyTb || []).length)) {
      parts.push('<div class="status warn">科目余额表里有 <b>' + u.onlyTb.length +
        '</b> 个类别本期没有任何发生额：' + esc((u.onlyTb || []).join('、')) +
        '<br/>这是正常业务（不提折旧、不购置、不处置），表1 该类别的变动行会是 0。</div>');
    }

    var dd = d.dedupe || {};
    if (dd.n) {
      var samp = (dd.sample || []).slice(0, 3).map(function (x) {
        return esc(x.company) + ' ' + esc(x.acct) + '（源行 ' + esc(x.srcRow) + '）';
      }).join('、');
      parts.push('<div class="status warn"><b>TB 里有 ' + dd.n + ' 行是整块重复导入</b>：' +
        samp + ' 等。<br/>已按「主体 + 科目 + 辅助核算 + 四个金额」去重，' +
        '去重后与序时账净额精确勾稽；不去重的话期初期末会翻倍。</div>');
    }
    (dd.adjusts || []).forEach(function (a) {
      parts.push('<div class="status warn">期初修正：' + esc(a.company) + ' ' + esc(a.acct) +
        ' TB 期初 ' + fm(a.tb_open) + ' → 反推 ' + fm(a.expect_open) +
        '（' + esc(a.reason) + '）</div>');
    });

    var t4 = d.table4 || {};
    if (t4.diff != null && Math.abs(Number(t4.diff)) > 0.005) {
      parts.push('<div class="status warn">表4 折旧费用与「累计折旧本年计提」差 <b>' +
        fm(t4.diff) + '</b>，生成时会再校一次，届时不平会直接拦下。</div>');
    }

    if ((d.cats || []).length > 7) {
      parts.push('<div class="status warn">类别有 ' + d.cats.length +
        ' 个，超过模板预留的 7 块，生成时会自动插列。</div>');
    }

    box.innerHTML = parts.join('');
  }

  /* ---------- ④ 生成 ---------- */
  function gen() {
    if (S.busy) return;
    var dateEl = $('#optFaDate');
    if (!dateEl || !dateEl.value) {
      setStatus('#faGenStatus', 'err', '请先填资产负债表日');
      toast('请先填资产负债表日', true);
      return;
    }
    S.busy = true;
    $('#faGenResult').hidden = true;
    setStatus('#faGenStatus', 'loading',
      '正在生成：去重 → 归类 → 排变动行 → 排类别列 → 三张清单 → K.01 → Lead …（约 10~30 秒）');
    var body = {
      customer: ($('#optFaCustomer') || {}).value || '',
      balance_date: dateEl.value || ''
    };
    fetch('/api/build/fa_swp/run', {
      method: 'POST', headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify(body)
    }).then(function (r) { return r.json(); }).then(function (d) {
      S.busy = false;
      if (d.error) {
        setStatus('#faGenStatus', 'err', esc(d.error));
        toast(d.error, true);
        return;
      }
      setStatus('#faGenStatus', 'ok',
        '已生成：<b>' + esc(d.latestFile || d.file) + '</b>　耗时 ' + esc(d.elapsed) + 's　' +
        '校验 ' + esc(d.checkItems) + ' 项' +
        ((d.checkFails || []).length ? '（' + d.checkFails.length + ' 项未过）' : '（全过）'));
      $('#faGenTiles').innerHTML =
        tile('资产类别', d.cats, (d.catOrder || []).join('、'), d.cats > 7 ? 'warn' : 'ok') +
        tile('原值期末', wan(Number(d.totals.fa_end)), '期初 ' + wan(Number(d.totals.fa_beg))) +
        tile('净值期末', wan(Number(d.net.end)), '期初 ' + wan(Number(d.net.beg))) +
        tile('插入/删除行', d.rowsInserted, d.rowsInserted ? '按清单方式动态插行' : '模板原有行') +
        tile('TB 重复行', d.tbDup, '已去重', 'warn') +
        tile('TB 页行数', d.tbRows, '原始记录沉进底稿', 'ok');

      var files = '<a class="btn primary" href="' + d.url + '">下载底稿</a>' +
        '<a class="btn ghost" href="' + d.latestUrl + '">下载「最新」副本</a>';
      if (d.prior) {
        files += '<span class="hint" style="align-self:center">上年 K.01（模拟）：' +
          esc(d.prior) + '</span>';
      }
      $('#faFiles').innerHTML = files;
      $('#faGenResult').hidden = false;
      gaps(d);
      if ((d.warnings || []).length) toast(d.warnings[0], true);
    }).catch(function (e) {
      S.busy = false;
      setStatus('#faGenStatus', 'err', '请求失败：' + esc(e.message));
    });
  }

  function gaps(d) {
    var box = $('#faGaps');
    var rows = [];

    (d.checkFails || []).forEach(function (n) { rows.push(['校验未过', esc(n)]); });

    var rp = d.rowPlan || [];
    rp.forEach(function (x) {
      rows.push(['插入变动行', '第 ' + x.sec + ' 段第 ' + x.at + ' 行前插入：' +
        esc((x.labels || []).join('、'))]);
    });

    var bp = d.blockPlan || {};
    if (bp.n_new > 0) {
      rows.push(['插入类别块', '模板 ' + bp.n_have + ' 块 → 实际 ' + bp.n_need +
        ' 块，在第 ' + bp.insert_at + ' 列前插 ' + (bp.n_new * 3) + ' 列']);
    } else if (bp.n_need < bp.n_have) {
      rows.push(['删除空类别块', '模板 ' + bp.n_have + ' 块 → 实际 ' + bp.n_need +
        ' 块，删掉 ' + ((bp.n_have - bp.n_need) * 3) + ' 列空块']);
    }

    var k = d.k01 || {};
    Object.keys(k).forEach(function (key) {
      var p = key.split('|');
      if (p.length === 2) {
        rows.push(['表1 ' + (p[0] === 'fa' ? '原值' : p[0] === 'dep' ? '折旧' : '减值') +
          ' · ' + p[1], fm(Number(k[key]))]);
      }
    });

    var t4 = d.table4 || {};
    if (t4 && Object.keys(t4).length) {
      Object.keys(t4).forEach(function (b) {
        if (Number(t4[b]) !== 0) rows.push(['表4 折旧费用 · ' + b, fm(Number(t4[b]))]);
      });
    }

    (d.hidden || []).length && rows.push(['隐藏辅助页', (d.hidden || []).join('、') +
      '（右键取消隐藏可查来源）']);

    (d.warnings || []).forEach(function (w) { rows.push(['提示', esc(w)]); });

    if (!rows.length) {
      box.innerHTML = '<div class="status ok">没有待确认项。</div>';
      return;
    }
    var h = '<table class="tbl"><thead><tr><th style="width:150px">类型</th><th>说明</th>' +
      '</tr></thead><tbody>';
    rows.forEach(function (r) { h += '<tr><td>' + r[0] + '</td><td>' + r[1] + '</td></tr>'; });
    box.innerHTML = h + '</tbody></table>';
  }

  function bind() {
    if (!$('#btnFaReload')) return;
    $('#btnFaReload').addEventListener('click', load);

    var dt = $('#optFaDate');
    if (dt && !dt.value) {
      var n = new Date();
      dt.value = n.getFullYear() + '-' + String(n.getMonth() + 1).padStart(2, '0') +
        '-' + String(n.getDate()).padStart(2, '0');
    }

    $('#btnFaGen').addEventListener('click', gen);
  }

  window.FA = {
    load: function () { bind(); load(); },
    loaded: function () { return S.loaded; }
  };

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', function () { if ($('#btnFaReload')) bind(); });
  } else if ($('#btnFaReload')) {
    bind();
  }
})();
