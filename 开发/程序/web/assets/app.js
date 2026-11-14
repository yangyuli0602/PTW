'use strict';
/* PTW平台 · 前端（无框架、无外网依赖） */

const $ = (s) => document.querySelector(s);
const esc = (s) => String(s === null || s === undefined ? '' : s)
  .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;')
  .replace(/"/g, '&quot;').replace(/'/g, '&#39;');
const fmtInt = (n) => (n === null || n === undefined) ? '-' : Number(n).toLocaleString('zh-CN');
const fmtSize = (b) => {
  if (b === null || b === undefined) return '-';
  const u = ['B', 'KB', 'MB', 'GB', 'TB']; let i = 0, v = b;
  while (v >= 1024 && i < u.length - 1) { v /= 1024; i++; }
  return (i === 0 ? v : v.toFixed(v >= 100 ? 0 : 1)) + ' ' + u[i];
};
const fmtCell = (v) => {
  if (v === null || v === undefined) return '';
  if (typeof v === 'number') {
    if (Number.isInteger(v)) return v.toLocaleString('zh-CN');
    return v.toLocaleString('zh-CN', { minimumFractionDigits: 2, maximumFractionDigits: 2 });
  }
  return String(v);
};

const S = {
  state: null, tab: 'import',
  modules: [],               // 板块清单，来自 /api/modules
  appInfo: null,
  files: [],                 // [{fid, name, size, kind, data}]
  mapSid: null, mapDraft: {}, auxDraft: {}, previewCache: {},
  rules: [], ruleId: '__free', mode: 'sql', sqlDraft: {}, paramDraft: {},
  result: null, page: 0, pageSize: 200,
  schema: null, running: false,
  mergeStat: {}, shareOut: null, busyT: null, importing: false,
  cleanSid: null, cleanSteps: [], stepOpen: -1, cleanTypes: null, presets: [],
  pv: null, dropped: null, recon: null, reconDiffOnly: true,
  books: null, bookSid: null,
  skip: new Set(), pgGroups: [], recipes: [], recipeSel: ''
};

/* ------------------------------------------------------------------ 通讯 */
async function req(method, url, body) {
  const opt = { method, headers: {} };
  if (body !== undefined) { opt.headers['Content-Type'] = 'application/json'; opt.body = JSON.stringify(body); }
  const r = await fetch(url, opt);
  const t = await r.text();
  let j = null; try { j = t ? JSON.parse(t) : null; } catch (e) { /* 非 JSON */ }
  if (!r.ok) {
    // 把响应体挂到 error 上：门禁这类「有结构的失败」要把清单渲染出来，不能只弹一句
    const err = new Error((j && (j.error || j.verdict)) || ('HTTP ' + r.status + ' ' + t.slice(0, 200)));
    err.payload = j; err.status = r.status;
    throw err;
  }
  return j;
}
const apiGet = (u) => req('GET', u);
const apiPost = (u, b) => req('POST', u, b !== undefined ? b : {});

function uploadFile(file, onPct) {
  return new Promise((res, rej) => {
    const xhr = new XMLHttpRequest();
    xhr.open('POST', '/api/upload');
    xhr.setRequestHeader('X-Filename', encodeURIComponent(file.name));
    xhr.upload.onprogress = (e) => { if (e.lengthComputable && onPct) onPct(e.loaded / e.total); };
    xhr.onload = () => {
      let j = null;
      try { j = JSON.parse(xhr.responseText); } catch (e) { /* ignore */ }
      if (xhr.status >= 400) rej(new Error((j && j.error) || ('上传失败 ' + xhr.status)));
      else res(j);
    };
    xhr.onerror = () => rej(new Error('上传失败（网络中断）'));
    xhr.send(file);
  });
}

/* ------------------------------------------------------------------ 提示 */
function toast(msg, kind) {
  const d = document.createElement('div');
  d.className = 'toast' + (kind ? ' ' + kind : '');
  d.textContent = msg;
  $('#toastWrap').appendChild(d);
  setTimeout(() => d.remove(), kind === 'err' ? 4200 : 2400);
}
function busy(msg, sub) {
  $('#busyMsg').textContent = msg || '处理中';
  const t0 = Date.now();
  $('#busySub').textContent = sub || '';
  $('#busyBar').hidden = true;
  $('#busyPct').hidden = true;
  $('#busy').classList.add('on');
  clearInterval(S.busyT);
  S.busyT = setInterval(() => {
    if (!$('#busyPct').hidden) return;          // 进度模式下不再显示计时
    const s = (Date.now() - t0) / 1000;
    $('#busySub').textContent = (sub ? sub + ' · ' : '') + '已用 ' + s.toFixed(s < 10 ? 1 : 0) + 's';
  }, 100);
}
/* 导入进度模式：显示百分比进度条（后端按阶段 + 字节量推进），不显示计时 */
function busyProgress(msg, sub) {
  busy(msg, sub);
  $('#busyBar').hidden = false;
  $('#busyPct').hidden = false;
  $('#busyFill').style.width = '0%';
  $('#busyPct').textContent = '0%';
}
function setProgress(pct, phase) {
  const p = Math.max(0, Math.min(100, Math.round(pct || 0)));
  $('#busyFill').style.width = p + '%';
  $('#busyPct').textContent = phase ? (p + '% · ' + phase) : (p + '%');
}
function idle() { clearInterval(S.busyT); $('#busy').classList.remove('on'); }
async function guard(msg, fn) {
  busy(msg); try { return await fn(); } finally { idle(); }
}
function modal(title, bodyHtml, footHtml) {
  $('#modalTitle').textContent = title;
  $('#modalBody').innerHTML = bodyHtml;
  $('#modalFoot').innerHTML = footHtml || '<button class="btn" data-close>关闭</button>';
  $('#modal').classList.add('on');
}
function closeModal() { $('#modal').classList.remove('on'); }

/* ------------------------------------------------------------------ 启动 */
async function boot() {
  // 板块清单：决定左侧导航，拿不到就用内置兜底
  try {
    const info = await apiGet('/api/modules');
    S.appInfo = info;
    S.modules = (info && info.modules) || [];
    const sub = $('#brandSub');
    if (sub && info && info.version) {
      sub.textContent = (info.app || 'PTW平台') + ' · PBC to WP · v' + info.version;
    }
    const vt = $('#verTag');
    if (vt && info && info.version) vt.textContent = 'v' + info.version;
  } catch (e) { S.modules = []; }
  if (!navItems().length) S.modules = NAV_FALLBACK;
  try {
    S.rules = (await apiGet('/api/rules')).rules || [];
  } catch (e) { toast('规则加载失败：' + e.message, 'err'); S.rules = []; }
  try {
    const t = await apiGet('/api/clean/types');
    S.cleanTypes = t.types; S.presets = t.presets;
  } catch (e) { S.cleanTypes = []; S.presets = []; }
  await refresh();
  renderNav();
  go(S.tab);
  wireGlobals();
}

function wireGlobals() {
  // 子页签（清洗↔映射、合并↔检查）用事件委托：页面重绘多少次都不用重新绑定
  const mainEl = $('#main');
  if (mainEl && !mainEl.dataset.subBound) {
    mainEl.dataset.subBound = '1';
    mainEl.addEventListener('click', (e) => {
      const t = e.target.closest ? e.target.closest('.subtab') : null;
      if (t && t.dataset.sub) go(t.dataset.sub);
    });
  }
  $('#modalClose').onclick = closeModal;
  $('#modal').addEventListener('click', (e) => { if (e.target.id === 'modal') closeModal(); });
  $('#btnHelp').onclick = showHelp;
  $('#btnBackup').onclick = async () => {
    try {
      const r = await guard('正在备份数据库…', () => apiPost('/api/backup'));
      toast('已备份：' + r.file, 'ok');
      window.location.href = r.url;
    } catch (e) { toast(e.message, 'err'); }
  };
}

async function refresh() {
  S.state = await apiGet('/api/state');
  if (S.state.sheets.length && !S.mapSid) S.mapSid = S.state.sheets[0].sid;
  await loadRecipes();
  renderTop();
  renderNav();
}

async function loadRecipes() {
  try {
    const r = await apiGet('/api/recipes');
    S.recipes = r.recipes || [];
  } catch (e) { S.recipes = S.recipes || []; }
}

function renderTop() {
  const st = S.state || { stats: {}, sheets: [], db_size: 0 };
  const jr = (st.stats.journal || {}).rows || 0;
  const tr = (st.stats.tb || {}).rows || 0;
  const mapped = st.sheets.filter((s) => s.target && Object.keys(s.mapping || {}).length).length;
  $('#topStatus').innerHTML = [
    `<span class="chip">序时账 <b>${fmtInt(jr)}</b> 行</span>`,
    `<span class="chip">科目余额表 <b>${fmtInt(tr)}</b> 行</span>`,
    `<span class="chip">工作表 <b>${st.sheets.length}</b> / 已映射 <b>${mapped}</b></span>`,
    `<span class="chip">库 <b>${fmtSize(st.db_size)}</b></span>`,
    st.excel_ok === false ? '<span class="chip" style="color:#c62828">Excel 扩展未加载</span>' : ''
  ].join('');
}

/* ---------------------------------------------------------- 导航（由板块清单驱动）
 * 板块清单来自后端 /api/modules（扫 modules 下各目录的 manifest.json 得到）。
 * 新增一个板块：只加目录 + manifest.json，本文件不用动。
 * 例外：type=native 的板块需要在这里的 VIEWS 里登记渲染函数（复用主页面框架）。
 */
const VIEWS = {
  import: renderImport,
  clean: renderClean,
  map: renderMap,
  merge: renderMerge,
  check: renderCheck,
  query: renderQuery,
  share: renderShare
};
// 渲染后需要跑一次的副作用（如查账要拉 schema）
const VIEW_HOOKS = {
  query: () => { S.schema = null; loadSchema().then(() => { if (S.tab === 'query') renderQuery(); }); },
  check: () => { if (!S.recon) restoreRecon(); }
};

// 检查页刷新后恢复上次结果（结果已持久化在库里）
async function restoreRecon() {
  try {
    const last = await req('GET', '/api/recon/last');
    if (last && last.summary && last.summary.items) {
      const p = last.params || {};
      const tbBase = p.tb_base || 'ytd';
      S.recon = {
        summary: last.summary,
        period: p.period || {},
        tol: p.tol || 0.01,
        tb_base: tbBase,
        tb_label: (TB_BASES.find((b) => b[0] === tbBase) || [])[1] || '',
        _restored: true, ran_at: last.ran_at
      };
      if (S.tab === 'check') renderCheck();
    }
  } catch (e) { /* 没有历史结果就静默跳过 */ }
}

// 一步里并了两个页面时，页内子页签（与 manifest 的 subs 对应）
const SUB_TABS = {
  clean: [{ k: 'clean', label: '数据清洗' }, { k: 'map', label: '字段映射' }],
  merge: [{ k: 'merge', label: '合并成表' }, { k: 'check', label: '数据检查' }]
};

// 导入与清洗板块的 4 步；第 2、3 步各带页内子页签（数据清洗↔字段映射、合并成表↔数据检查）
const INGEST_STEPS = [
  { k: 'import', no: '1', label: '导入数据' },
  { k: 'clean', no: '2', label: '清洗与映射', subs: [
    { k: 'clean', label: '数据清洗' }, { k: 'map', label: '字段映射' }] },
  { k: 'merge', no: '3', label: '合并与检查', subs: [
    { k: 'merge', label: '合并成表' }, { k: 'check', label: '数据检查' }] },
  { k: 'share', no: '4', label: '分享导出' }
];

// 后端拿不到清单时的兜底，保证界面不会空掉
const NAV_FALLBACK = [
  { key: 'ingest', label: '一、导入与清洗', type: 'native', children: INGEST_STEPS },
  { key: 'analysis', label: '二、查账与分析', type: 'native', children: [
    { k: 'query', no: '1', label: '查账分析' }
  ]},
  { key: 'bkd', label: '三、WP初稿生成', short: 'WP初稿生成', type: 'embed', entry: 'bkd.html', icon: '◆', children: [
    { k: 'wp_cash', no: '1', label: '货币资金' },
    { k: 'wp_exp', no: '2', label: '销管费用' }
  ] }
];

const navItems = () => (S.modules && S.modules.length) ? S.modules : NAV_FALLBACK;

function moduleOf(tab) {
  for (const m of navItems()) {
    if (m.key === tab) return m;
    const kids = m.children || [];
    if (kids.some((c) => c.k === tab)) return m;
    // 子页签（map / check）也算属于该板块
    if (kids.some((c) => (c.subs || []).some((s) => s.k === tab))) return m;
  }
  return null;
}

const shortLabel = (label) => String(label || '').replace(/^[一二三四五六七八九十]+、\s*/, '');

/** 某一步是否已完成（导航对勾、子页签对勾共用） */
function stepDone(k) {
  const st = S.state || { stats: {}, sheets: [] };
  const sheets = st.sheets || [];
  if (k === 'import') return sheets.length > 0;
  if (k === 'clean') return sheets.some((s) => s.stg || (s.clean || []).length);
  if (k === 'map') return sheets.some((s) => s.target && Object.keys(s.mapping || {}).length > 1);
  if (k === 'merge') return (((st.stats.journal || {}).rows || 0) + ((st.stats.tb || {}).rows || 0)) > 0;
  if (k === 'check') return !!(S.recon && S.recon.summary);
  return false;
}

/** 页内子页签条：把两步并成一步后，靠它在本步内部切换 */
function subTabBar(group, cur) {
  const items = SUB_TABS[group] || [];
  if (!items.length) return '';
  return `<div class="subtabs">${items.map((it) => {
    const on = it.k === cur;
    return `<div class="subtab${on ? ' on' : ''}" data-sub="${esc(it.k)}">`
      + `${esc(it.label)}${!on && stepDone(it.k) ? '<i>✓</i>' : ''}</div>`;
  }).join('')}</div>`;
}

function renderNav() {
  const html = [];
  for (const grp of navItems()) {
    const kids = grp.children || [];
    html.push(`<div class="nav-group">${esc(grp.label)}</div>`);
    if (!kids.length) {
      // 独立页板块：点击直接进入 iframe 页
      const cls = ['nav-item'];
      if (S.tab === grp.key) cls.push('on');
      html.push(`<div class="${cls.join(' ')}" data-tab="${esc(grp.key)}">`
        + `<span class="no">${esc(grp.icon || '◆')}</span>${esc(grp.short || shortLabel(grp.label))}</div>`);
    } else {
      for (const n of kids) {
        const subs = n.subs || [];
        // 停在子页签（map / check）上时，父步骤同样要高亮
        const active = S.tab === n.k || subs.some((s) => s.k === S.tab);
        const cls = ['nav-item'];
        if (active) cls.push('on');
        else if (stepDone(n.k)) cls.push('done');
        html.push(`<div class="${cls.join(' ')}" data-tab="${esc(n.k)}">`
          + `<span class="no">${!active && stepDone(n.k) ? '✓' : esc(n.no)}</span>${esc(n.label)}</div>`);
      }
    }
  }
  $('#nav').innerHTML = html.join('');
  $('#nav').querySelectorAll('.nav-item').forEach((el) => {
    el.onclick = () => go(el.dataset.tab);
  });
}
function go(tab) {
  S.tab = tab;
  renderNav();
  const m = moduleOf(tab);
  if (m && m.type === 'embed') return renderEmbed(m);
  const fn = VIEWS[tab];
  if (!fn) {
    $('#main').innerHTML = `<div class="card"><div class="card-b">`
      + `板块「${esc(tab)}」是原生类型，但 app.js 的 VIEWS 里没有登记它的渲染函数。`
      + `<br>要么把 manifest 的 type 改成 embed 走独立页，要么在 VIEWS 里补一个函数。</div></div>`;
    return;
  }
  fn();
  const hook = VIEW_HOOKS[tab];
  if (hook) hook();
}

/* ================================================================== 独立页板块（iframe 嵌入，如 BKD） */
function renderEmbed(m) {
  // 独立页板块如果带子级（如 WP初稿生成 → 货币资金 / 销管费用），
  // 用 ?sub= 告诉内嵌页当前该打开哪一块，页面自己切页签。
  const kids = m.children || [];
  const sub = kids.some((c) => c.k === S.tab) ? '?sub=' + encodeURIComponent(S.tab) : '';
  const src = '/' + m.key + '/' + (m.entry || 'index.html') + sub;
  $('#main').innerHTML = `
    <div class="card" style="padding:0;overflow:hidden;flex:1;display:flex;flex-direction:column">
      <iframe id="modFrame" src="${esc(src)}" style="width:100%;height:calc(100vh - 140px);border:none;display:block"
        title="${esc(m.label)}"></iframe>
    </div>`;
}


/* ================================================================== 1 导入 */
function renderImport() {
  const st = S.state;
  const imported = st.sheets;
  $('#main').innerHTML = `
    <div class="imp3">
      <div class="card imp-card">
        <div class="card-h">
          <h3>导入原始数据</h3>
          <span class="sub">xlsx / xls / csv / 网页表格，可整个文件夹</span>
        </div>
        <div class="card-b">
          <div class="drop drop-sq" id="drop">
            <b>把文件或文件夹拖到这里</b>
            <span>或点击选择文件</span>
            <div class="hint">支持 .xlsx .xls .csv（.xlsm / .txt / .tsv 同样可读）；
              ERP / 网银「导出 Excel」得到的网页表格（.html/.htm，或后缀写成 .xls 的 HTML）会自动转换；
              文件夹自动递归子目录，临时文件与非表格文件自动跳过</div>
          </div>
          <div class="row" style="gap:6px;margin-top:8px">
            <button class="btn sm" id="btnPickFiles">选择文件</button>
            <button class="btn sm" id="btnPickFolder">选择文件夹</button>
          </div>
        </div>
      </div>
      <div class="card imp-card">
        <div class="card-h">
          <h3>导入数据包（zip）</h3>
          <span class="sub">含标准表 + 清洗规则 + 字段映射 + 账簿对照表</span>
        </div>
        <div class="card-b">
          <div class="drop drop-sq" id="dropZip">
            <b>把 zip 数据包拖到这里</b>
            <span>或点击选择文件</span>
            <div class="hint">同事打包发来的数据包，一键还原成标准表与全部配置</div>
          </div>
          <div class="row" style="gap:14px;margin:8px 0 6px">
            <label class="ck"><input type="radio" name="impMode" value="replace" checked> 替换</label>
            <label class="ck"><input type="radio" name="impMode" value="append"> 追加</label>
          </div>
          <div id="impOut"></div>
        </div>
      </div>
      <div class="card imp-card">
        <div class="card-h">
          <h3 style="white-space:nowrap">账簿对照表</h3>
          <span class="sub">账簿 → 公司 / 简称</span>
          <span style="flex:1"></span>
          ${st.books_rows ? '<button class="btn sm ghost danger" id="btnBooksClear">清空</button>' : ''}
        </div>
        <div class="card-b">
          <div class="drop drop-sq" id="dropBooks">
            <b>把对照表拖到这里</b>
            <span>或点击选择文件</span>
            <div class="hint">账簿名称、公司名称、简称三列即可；合并时按账簿名称自动带进余额表和序时账</div>
          </div>
          <div class="small muted" style="margin-top:6px">已有 ${fmtInt(st.books_rows || 0)} 条</div>
          <div class="card-b" id="booksBox" style="padding:8px 0 0"></div>
        </div>
      </div>
    </div>
    <input type="file" id="booksFile" accept=".xlsx,.xlsm,.xls,.csv,.txt" hidden>
    <div id="impSummary"></div>
    <div id="pending"></div>
    <div class="card">
      <div class="card-h"><h3>已入库的原始表</h3><span class="sub">共 ${imported.length} 张</span></div>
      <div class="card-b" id="importedBox"></div>
    </div>`;

  const drop = $('#drop');
  drop.onclick = () => $('#fileInput').click();
  bindDrop(drop, (fs, meta) => handleFiles(fs, {
    fromFolder: !!(meta && meta.hasDir),
    preSkipped: (meta && meta.skipped) || []      // 扫描阶段就跳过的，也要进汇总
  }));
  const bpFiles = $('#btnPickFiles');
  if (bpFiles) bpFiles.onclick = (e) => { e.stopPropagation(); $('#fileInput').click(); };
  const bpFolder = $('#btnPickFolder');
  if (bpFolder) bpFolder.onclick = (e) => { e.stopPropagation(); $('#folderInput').click(); };
  $('#fileInput').onchange = (e) => {
    const fs = Array.from(e.target.files || []);
    e.target.value = '';
    if (fs.length) handleFiles(fs);
  };
  // 选文件夹：webkitdirectory 会带上相对路径，显示和排查都靠它
  const fIn = $('#folderInput');
  if (fIn) {
    fIn.onchange = (e) => {
      const fs = Array.from(e.target.files || []);
      e.target.value = '';
      if (!fs.length) { toast('这个文件夹里没有可读的文件', 'err'); return; }
      fs.forEach((f) => { try { f.rel = f.webkitRelativePath || f.name; } catch (_) { /* 只读属性，忽略 */ } });
      handleFiles(fs, { fromFolder: true });
    };
  }
  // 数据包与对照表：拖到哪个方块就按哪个用途处理，不用先想清楚再找按钮
  const dz = $('#dropZip');
  if (dz) bindDrop(dz, (fs) => {
    const f = fs.filter((x) => /\.zip$/i.test(x.name))[0];
    if (!f) { toast('数据包只接 zip；原始表请拖到左边的方块', 'err'); return; }
    doImportZip(f);
  });
  const db = $('#dropBooks');
  if (db) bindDrop(db, (fs) => {
    const f = fs.filter((x) => /\.(xlsx|xlsm|xls|csv|txt)$/i.test(x.name))[0];
    if (!f) { toast('对照表只接 Excel / CSV', 'err'); return; }
    doImportBooks(f);
  });
  renderPending();
  renderImported();
  renderImportSummary();
  bindBooks();
  bindZipImport();
}

/* 统一的拖拽绑定：高亮 + 落盘（文件和文件夹都接） */
function bindDrop(el, onFiles) {
  if (!el) return;
  ['dragenter', 'dragover'].forEach((ev) => el.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation(); el.classList.add('hover');
  }));
  ['dragleave', 'drop'].forEach((ev) => el.addEventListener(ev, (e) => {
    e.preventDefault(); e.stopPropagation(); el.classList.remove('hover');
  }));
  el.addEventListener('drop', async (e) => {
    const dt = e.dataTransfer;
    // 拖文件夹时 dataTransfer.files 只有一个 0 字节占位，必须走 entry 递归
    const hasDir = Array.from((dt && dt.items) || []).some((it) => {
      if (it.kind !== 'file' || typeof it.webkitGetAsEntry !== 'function') return false;
      const en = it.webkitGetAsEntry();
      return !!(en && en.isDirectory);
    });
    if (hasDir) busy('扫描文件夹…', '正在递归子目录');
    try {
      const r = await filesFromDrop(dt);
      if (!r.files.length && !r.skipped.length) { toast('没有可导入的文件', 'err'); return; }
      onFiles(r.files, { skipped: r.skipped, hasDir });
    } catch (err) {
      toast('读取拖入内容失败：' + ((err && err.message) || err), 'err');
    } finally {
      if (hasDir) idle();
    }
  });
}

/* 导入范围：这几种后缀才进库，其余在扫描阶段就归到「跳过」并说明原因
   .zip 也收：客户 ERP 常按账导出成 zip（里面是明细账 xlsx），上传后由服务端解开，
   不当成"平台数据包"打发走（数据包另有 manifest.json，会在展开时被识别出来）。 */
const IMPORT_EXT = ['.xlsx', '.xlsm', '.xls', '.csv', '.txt', '.tsv', '.html', '.htm', '.zip'];
const TEMP_EXT = ['.tmp', '.temp', '.crdownload', '.part', '.partial', '.download', '.!ut', '.bak', '.swp'];
const TEMP_NAMES = ['thumbs.db', '.ds_store', 'desktop.ini'];
const MAX_DEPTH = 12;            // 防异常目录层级/软链接把浏览器卡死

function extOf(name) {
  const m = String(name || '').toLowerCase().match(/\.[^.\/\\]+$/);
  return m ? m[0] : '';
}

function relOf(f) {
  return String((f && (f.rel || f.name)) || '');
}

function byRelPath(a, b) {
  return relOf(a).localeCompare(relOf(b), 'zh', { numeric: true });
}

/* 为什么这个文件不导入；返回空串表示可以导 */
function skipReason(f) {
  const base = String((f && f.name) || '');
  const low = base.toLowerCase();
  const rel = relOf(f);
  const ext = extOf(low);
  if (!base) return '文件名为空';
  if (/^~\$/.test(base)) return 'Office 临时锁文件（~$ 开头）';
  if (TEMP_NAMES.indexOf(low) >= 0) return '系统文件';
  if (TEMP_EXT.indexOf(ext) >= 0) return '临时文件或未下载完';
  if (/(^|\/)\._/.test(rel)) return 'macOS 资源分支文件';
  if (!f.size) return '空文件（0 字节）';
  if (IMPORT_EXT.indexOf(ext) < 0) return '非目标格式（' + (ext || '无后缀') + '）';
  return '';
}

function isIgnoredDir(name) {
  const n = String(name || '').toLowerCase();
  if (!n) return true;
  if (n === '__macosx' || n === '.git' || n === '.svn' || n === 'node_modules'
      || n === '$recycle.bin' || n === 'system volume information') return true;
  return n.charAt(0) === '.';    // .idea / .vscode / .cache 之类
}

/* readEntries 单次最多给 100 条，必须循环读到返回空数组为止，否则大目录会漏文件 */
function readAllEntries(reader) {
  return new Promise((resolve, reject) => {
    const acc = [];
    const step = () => reader.readEntries((batch) => {
      if (!batch || !batch.length) return resolve(acc);
      acc.push(...batch);
      step();
    }, reject);
    step();
  });
}

async function walkEntry(entry, prefix, files, skipped, depth) {
  const rel = prefix ? prefix + '/' + entry.name : entry.name;
  if (depth > MAX_DEPTH) { skipped.push({ name: rel, why: '目录层级过深，已跳过' }); return; }
  if (entry.isFile) {
    let f = null;
    try {
      f = await new Promise((res, rej) => entry.file(res, rej));
    } catch (e) {
      skipped.push({ name: rel, why: '读取失败：' + ((e && e.message) || '无权限') });
      return;
    }
    try { f.rel = rel; } catch (_) { /* File.name 只读，挂了就退回文件名 */ }
    const why = skipReason(f);
    if (why) skipped.push({ name: rel, why });
    else files.push(f);
  } else if (entry.isDirectory) {
    if (isIgnoredDir(entry.name)) { skipped.push({ name: rel, why: '系统 / 隐藏目录' }); return; }
    let kids = [];
    try {
      kids = await readAllEntries(entry.createReader());
    } catch (e) {
      skipped.push({ name: rel, why: '目录读取失败：' + ((e && e.message) || '无权限') });
      return;
    }
    kids.sort((a, b) => String(a.name).localeCompare(String(b.name), 'zh', { numeric: true }));
    for (const k of kids) await walkEntry(k, rel, files, skipped, depth + 1);
  }
}

/* 把一次拖拽展开成「可导入文件 + 跳过清单」 */
async function filesFromDrop(dt) {
  const items = Array.from((dt && dt.items) || []);
  const files = [], skipped = [], roots = [];
  for (const it of items) {
    if (it.kind !== 'file' || typeof it.webkitGetAsEntry !== 'function') continue;
    const en = it.webkitGetAsEntry();
    if (en) roots.push(en);
  }
  if (!roots.length) {                     // 老浏览器没有 entry API，退回 files
    return { files: Array.from((dt && dt.files) || []), skipped };
  }
  for (const r of roots) await walkEntry(r, '', files, skipped, 0);
  files.sort(byRelPath);
  return { files, skipped };
}

/* 把后端/网络的原始报错换成能看懂的话 */
function friendlyImportErr(e) {
  const m = String((e && e.message) || e || '').trim();
  // 加密必须排在「损坏」前面：加密的 xlsx/.xls 底层都是 OLE2，报错文本里常带
  // not a zip / 损坏 字样，先匹配损坏会把人往「文件坏了」的方向带。
  if (/加密|密码|受保护|password|encrypt/i.test(m)) {
    return '该文件已加密（需要密码），请在 Excel 里另存为不加密版本后重试';
  }
  // 下面三条同样必须排在「损坏」之前——它们描述的是"格式不对"，不是"文件坏了"。
  // 尤其 HTML 伪 xls：以前就是被当成损坏，实际是内容为网页表格。
  if (/spreadsheetml|2003\s*xml|office\s*xml|mso-application/i.test(m)) {
    return '这是 Excel 2003 XML 表格（SpreadsheetML），请在 Excel 里另存为 .xlsx 再导入';
  }
  if (/网页|html|<table>|\.mht/i.test(m)) {
    return '这是网页表格，没能转成 Excel：请用 Excel 打开后另存为 .xlsx 再导入';
  }
  if (/不是 Excel 工作簿|zip 压缩包/i.test(m)) {
    return '这不是 Excel 文件（是压缩包，或别的 Office 文档改了后缀）；压缩包请直接拖进导入区';
  }
  if (/bad ?zip|not a zip|corrupt|CRC error|损坏|truncated/i.test(m)) {
    return '无法解析：文件可能已损坏，或后缀与实际格式不符';
  }
  if (/utf|decode|codec|编码/i.test(m)) return '无法解析：文本编码识别失败（可另存为 UTF-8 后重试）';
  if (/permission|denied|拒绝|占用/i.test(m)) return '文件被占用或无权限读取（请先关闭 Excel）';
  if (/timeout|timed out|network|网络|中断/i.test(m)) return '上传中断（网络或服务异常）';
  return m.slice(0, 160) || '未知错误';
}

/* ---- 数据包导入（zip）：按钮与拖拽共用同一个动作 ---- */
function bindZipImport() {
  const dz = $('#dropZip');
  if (dz && !dz.dataset.bound) {
    dz.dataset.bound = '1';
    dz.addEventListener('click', () => $('#zipInput').click());
  }
  const zi = $('#zipInput');
  if (zi) {
    zi.onchange = (e) => {
      const f = e.target.files[0];
      e.target.value = '';
      if (f) doImportZip(f);
    };
  }
}

async function doImportZip(f) {
  const mode = (document.querySelector('input[name=impMode]:checked') || {}).value || 'replace';
  const out = $('#impOut');
  if (out) out.innerHTML = '<div class="small muted">正在上传数据包…</div>';
  try {
    const info = await guard('上传数据包 ' + f.name, () => uploadFile(f));
    const r = await guard('正在导入数据包…', () => apiPost('/api/share/import', { fid: info.file_id, mode }));
    const lines = Object.entries(r.tables || {}).map(([k, v]) =>
      (TGT_LABEL[k] || k) + ' ' + fmtInt(v) + ' 行').join('，');
    const cfg = r.config || {};
    const extra = [];
    if (cfg.books) extra.push('账簿对照表 ' + fmtInt(cfg.books) + ' 条');
    if (cfg.sheetConfigs) extra.push('字段映射/清洗规则 ' + fmtInt(cfg.sheetConfigs) + ' 张表');
    const sk = r.skipped || [];
    if (out) {
      out.innerHTML = `<div class="row" style="flex-wrap:wrap"><span class="tag green">完成</span>
        <span class="small">${esc(lines || '无标准表')}</span></div>`
        + (extra.length ? `<div class="small muted" style="margin-top:4px">同时恢复：${esc(extra.join('，'))}</div>` : '')
        + (sk.length ? `<div class="small" style="color:var(--danger);margin-top:4px">
            包内有 ${sk.length} 个文件读取失败，其余已导入：${esc(sk.slice(0, 3).map((x) => x.name + '（' + x.why + '）').join('；'))}${sk.length > 3 ? ' …' : ''}</div>` : '');
    }
    await refresh(); renderNav();
    toast('数据包导入完成', 'ok');
  } catch (err) {
    if (out) out.innerHTML = `<div class="small" style="color:var(--danger)">${esc(err.message)}</div>`;
    toast(err.message, 'err');
  }
}

/* ---- 账簿对照表：导入 / 预览 / 清空 ---- */
function bindBooks() {
  const db = $('#dropBooks');
  if (db && !db.dataset.bound) {
    db.dataset.bound = '1';
    db.addEventListener('click', () => $('#booksFile').click());
  }
  const bf = $('#booksFile');
  if (bf) {
    bf.onchange = async (e) => {
      const f = e.target.files && e.target.files[0];
      e.target.value = '';
      if (f) doImportBooks(f);
    };
  }
  const bc = $('#btnBooksClear');
  if (bc) {
    bc.onclick = async () => {
      if (!confirm('清空账簿对照表？已合并表里的公司名称不会跟着变。')) return;
      await apiPost('/api/books/clear');
      await refresh();
      renderImport();
      toast('已清空');
    };
  }
  if (S.state.books_rows) loadBooks();
}

async function doImportBooks(f) {
  try {
    const up = await guard('上传账簿对照表…', () => uploadFile(f));
    const pv = await guard('识别对照表…', () => apiPost('/api/books/preview', { fid: up.file_id }));
    showBooksPicker(up.file_id, pv);
  } catch (err) { toast(err.message, 'err'); }
}

async function loadBooks() {
  const box = $('#booksBox');
  if (!box) return;
  try {
    const r = await apiGet('/api/books/list?limit=2000');
    S.books = r;
    if (!r.total) return;
    box.innerHTML = `
      <div class="tbl-wrap mini" style="max-height:220px"><table class="grid book-tbl">
        <thead><tr>${(r.columns || []).map((c) => `<th>${esc(c)}</th>`).join('')}</tr></thead>
        <tbody>${(r.rows || []).slice(0, 300).map((row) => `<tr>
          ${row.map((v) => `<td>${esc(v)}</td>`).join('')}</tr>`).join('')}</tbody>
      </table></div>
      <div class="hint">共 ${fmtInt(r.total)} 条${r.total > 300 ? '，这里只显示前 300 条' : ''}。
        合并时按<b>账簿名称</b>一对一匹配，匹配不上的会在合并结果里列出来。</div>`;
  } catch (e) { /* 忽略 */ }
}

function showBooksPicker(fid, pv) {
  const sheets = (pv.sheets || []).filter((s) => !s.empty && (s.preview || []).length);
  if (!sheets.length) {
    modal('导入账簿对照表', '<div class="empty">这个文件里没有可用的数据表</div>',
      '<button class="btn" data-close>关闭</button>');
    bindClose(); return;
  }
  const guess = (cols, kws, excl) => {
    for (const c of cols || []) {
      const n = String(c);
      if (excl && excl.some((e) => n.includes(e))) continue;
      if (kws.some((k) => n.includes(k))) return c;
    }
    return '';
  };
  const build = (si) => {
    const s = sheets[si];
    const cols = s.columns || [];
    const bk = S.bookGuess && S.bookGuess.si === si ? S.bookGuess : null;
    const b = (bk && bk.book) || guess(cols, ['账簿', '账套', '核算主体'], null)
      || guess(cols, ['名称', '单位'], null);
    const c = (bk && bk.company) || guess(cols, ['公司名称', '企业名称', '全称', '主体名称', '单位名称'], ['简称'])
      || guess(cols, ['公司', '企业', '主体', '单位'], ['简称']);
    const sh = (bk && bk.short) || guess(cols, ['简称', '短称', '缩写'], null)
      || guess(cols, ['公司', '企业'], ['名称']);
    const opt = (cur) => ['<option value="">— 不用 —</option>']
      .concat(cols.map((x) => `<option value="${esc(x)}" ${x === cur ? 'selected' : ''}>${esc(x)}</option>`)).join('');
    return `
      ${sheets.length > 1 ? `<div class="row small" style="gap:6px;margin-bottom:8px">
        <span class="muted">工作表</span>
        <select id="bkSheet">${sheets.map((x, i) => `<option value="${i}" ${i === si ? 'selected' : ''}>${esc(x.sheet)}</option>`).join('')}</select>
      </div>` : ''}
      <div class="row small" style="gap:6px;margin-bottom:8px">
        <span class="muted">表头在第</span>
        <input id="bkHr" type="number" min="1" value="${s.header_row || 1}" style="width:70px">
        <span class="muted">行</span>
      </div>
      <table class="map-tbl" style="margin-top:6px">
        <thead><tr><th style="width:120px">对照表字段</th><th>对应列</th></tr></thead>
        <tbody>
          <tr><td><b>账簿名称</b> <span class="tag red">必</span></td><td><select id="bkBook">${opt(b)}</select></td></tr>
          <tr><td><b>公司名称</b></td><td><select id="bkCompany">${opt(c)}</select></td></tr>
          <tr><td><b>公司简称</b></td><td><select id="bkShort">${opt(sh)}</select></td></tr>
        </tbody>
      </table>
      <div class="hint">账簿名称必须能和科目余额表 / 序时账里的账簿名称对得上（一模一样，含空格差异会自动去掉）。
        没有简称列可以留空，但公司名称建议有。</div>`;
  };
  let si = 0;
  const paint = () => {
    modal('导入账簿对照表（' + fmtInt(S.state.books_rows || 0) + ' 条已有）',
      `<div class="small muted" style="margin-bottom:8px">${esc(pv.file.file_name)}</div>${build(si)}`,
      '<button class="btn" data-close>取消</button><button class="btn primary" id="bkOk">导入</button>');
    bindClose();
    const sel = $('#bkSheet');
    if (sel) {
      sel.onchange = () => {
        S.bookGuess = { si: +sel.value, book: $('#bkBook') ? $('#bkBook').value : '',
          company: $('#bkCompany') ? $('#bkCompany').value : '',
          short: $('#bkShort') ? $('#bkShort').value : '' };
        si = +sel.value;
        paint();
      };
    }
    $('#bkOk').onclick = async () => {
      const book = $('#bkBook').value;
      if (!book) { toast('必须指定「账簿名称」是哪一列', 'err'); return; }
      try {
        const r = await guard('导入账簿对照表…', () => apiPost('/api/books/import', {
          fid, sheet: sheets[si].sheet,
          header_row: parseInt($('#bkHr').value, 10) || 1,
          book_col: book,
          company_col: $('#bkCompany').value || null,
          short_col: $('#bkShort').value || null
        }));
        closeModal();
        await refresh();
        renderImport();
        toast(`已导入 ${fmtInt(r.rows)} 行，对照表共 ${fmtInt(r.total)} 条`, 'ok');
        (r.conflicts || []).slice(0, 2).forEach((c) =>
          toast(`账簿「${c.book_name}」在表里有 ${c.n} 个不同公司名称，已取最后一个`, 'err'));
      } catch (e) { toast(e.message, 'err'); }
    };
  };
  paint();
}

/* 批量导入：先分类（可导/跳过），再逐个上传识别，最后给一份结果汇总。
   文件夹递归后的清单也走这里，所以数量可能上百个，进度条按「第 i/n 个」推进。 */
async function handleFiles(files, opts) {
  const list = (files || []).filter(Boolean);
  const pre = (opts && opts.preSkipped) || [];
  if (!list.length && !pre.length) return;
  list.sort(byRelPath);

  const ok = [], skipped = pre.slice(), failed = [], todo = [], zips = [];
  const seen = new Set();
  for (const f of list) {
    const rel = relOf(f);
    if (seen.has(rel)) { skipped.push({ name: rel, why: '同一批里重复出现' }); continue; }
    seen.add(rel);
    const why = skipReason(f);
    if (why) { skipped.push({ name: rel, why }); continue; }
    // 压缩包不直接进 todo：先上传让服务端解开，包内每个表格各算一个待导入文件
    if (extOf(f.name) === '.zip') zips.push(f);
    else todo.push(f);
  }

  try {
    /* ---- 第一步：展开压缩包（客户 ERP 按账导出的 zip，里面是账表） ---- */
    for (let zi = 0; zi < zips.length; zi++) {
      const zf = zips[zi];
      const zlabel = relOf(zf);
      busy(`解压 ${zi + 1}/${zips.length}`, zlabel);
      try {
        const info = await uploadFile(zf);
        const r = await apiPost('/api/zip/expand', { fid: info.file_id });
        if (r && r.data_pack) {
          skipped.push({ name: zlabel,
            why: '这是平台数据包（含 manifest.json），请用中间「导入数据包（zip）」入口' });
          continue;
        }
        ((r && r.skipped) || []).forEach((x) =>
          skipped.push({ name: zlabel + ' / ' + x.name, why: x.why }));
        const got = (r && r.files) || [];
        if (!got.length) { failed.push({ name: zlabel, why: '压缩包里没有可导入的表格' }); continue; }
        got.forEach((x) => todo.push({
          ...x, from_zip: true,
          rel: zlabel + ' / ' + (x.rel || x.file_name || '')
        }));
      } catch (e) {
        failed.push({ name: zlabel, why: friendlyImportErr(e) });
      }
    }

    /* ---- 第二步：逐个识别并放进待导入 ---- */
    for (let i = 0; i < todo.length; i++) {
      const f = todo[i];
      const label = relOf(f);
      busy(`导入 ${i + 1}/${todo.length}`, label);
      try {
        // 压缩包里的文件已经在服务端登记过了，不用再传一次
        const info = f.from_zip
          ? { file_id: f.file_id, file_name: f.file_name, size: f.size, kind: f.kind }
          : await uploadFile(f, (p) => {
              $('#busySub').textContent = label + ' · ' + Math.round(p * 100) + '%';
            });
        $('#busyMsg').textContent = `识别 ${i + 1}/${todo.length}`;
        const data = await apiPost('/api/inspect', { fid: info.file_id });
        const sheets = (data && data.sheets) || [];
        if (!sheets.length) { failed.push({ name: label, why: '文件里没有可读的工作表' }); continue; }
        if (S.files.some((x) => x.file_id === info.file_id)) {
          skipped.push({ name: label, why: '已在待导入列表里（同一份文件）' });
          continue;
        }
        // 后端可能把「网页表格」就地转成了 xlsx，标记要带到待导入列表里给人看
        const fmeta = (data && data.file) || {};
        S.files.push({ ...info, data, rel: label,
          converted_from: fmeta.converted_from || '',
          file_note: fmeta.html_note || '' });
        ok.push({ name: label, sheets: sheets.length });
        if (i % 10 === 9) renderPending();      // 太多时别每张都重绘
      } catch (e) {
        failed.push({ name: label, why: friendlyImportErr(e) });
      }
    }
  } finally {
    idle();
  }

  renderPending();
  S.lastImport = { ok, skipped, failed, at: Date.now(), fromFolder: !!(opts && opts.fromFolder) };
  renderImportSummary();

  const tail = skipped.length ? `，跳过 ${skipped.length} 个` : '';
  if (!todo.length) toast(`没有可导入的文件：跳过 ${skipped.length} 个，见下方汇总`, 'err');
  else if (failed.length) toast(`成功 ${ok.length} 个，失败 ${failed.length} 个${tail}`, 'err');
  else toast(`导入完成：成功 ${ok.length} 个${tail}`, 'ok');
}

/* 上一次导入的结果汇总：成功 / 失败（逐条原因）/ 跳过（按原因归类） */
function renderImportSummary() {
  const box = $('#impSummary');
  if (!box) return;
  const r = S.lastImport;
  if (!r) { box.innerHTML = ''; return; }
  const total = r.ok.length + r.skipped.length + r.failed.length;
  if (!total) { box.innerHTML = ''; return; }

  const esc1 = (x) => esc(x);
  const sec = [];

  if (r.failed.length) {
    sec.push(`<div class="imp-sum-sec">
      <div class="imp-sum-h" style="color:var(--danger)">无法导入 ${fmtInt(r.failed.length)} 个</div>
      <ul class="imp-sum-list">${r.failed.slice(0, 100).map((x) =>
        `<li><code>${esc1(x.name)}</code><span class="muted"> — ${esc1(x.why)}</span></li>`).join('')}</ul>
      ${r.failed.length > 100 ? `<div class="small muted">仅列出前 100 个，共 ${fmtInt(r.failed.length)} 个</div>` : ''}
    </div>`);
  }
  if (r.skipped.length) {
    const by = {};
    r.skipped.forEach((x) => { (by[x.why] = by[x.why] || []).push(x); });
    const groups = Object.keys(by).sort((a, b) => by[b].length - by[a].length);
    sec.push(`<div class="imp-sum-sec">
      <div class="imp-sum-h">跳过 ${fmtInt(r.skipped.length)} 个</div>
      ${groups.map((g) => {
        const arr = by[g];
        const sample = arr.slice(0, 3).map((x) => esc1(x.name)).join('、');
        return `<div class="small" style="margin:3px 0">
          <span class="tag gray">${fmtInt(arr.length)}</span> ${esc1(g)}
          <span class="muted">（如 ${sample}${arr.length > 3 ? ' 等' : ''}）</span></div>`;
      }).join('')}
    </div>`);
  }
  if (r.ok.length) {
    sec.push(`<div class="imp-sum-sec">
      <div class="imp-sum-h" style="color:var(--ok)">已加入待导入 ${fmtInt(r.ok.length)} 个</div>
      <div class="small muted">${r.ok.slice(0, 8).map((x) => esc1(x.name)).join('、')}${r.ok.length > 8 ? ' 等' : ''}</div>
    </div>`);
  }

  box.innerHTML = `<div class="card">
    <div class="card-h">
      <h3>上次扫描结果</h3>
      <span class="sub">${r.fromFolder ? '文件夹（含子目录）' : '文件'}</span>
      <span style="flex:1"></span>
      <span class="small muted">${new Date(r.at).toLocaleTimeString('zh-CN', { hour12: false })}</span>
      <button class="btn sm ghost" id="btnSumClear">关闭</button>
    </div>
    <div class="card-b">
      <div class="imp-sum-stats">
        <span>共扫描 <b>${fmtInt(total)}</b> 个文件</span>
        <span style="color:var(--ok)">成功 <b>${fmtInt(r.ok.length)}</b></span>
        <span style="color:var(--danger)">失败 <b>${fmtInt(r.failed.length)}</b></span>
        <span class="muted">跳过 <b>${fmtInt(r.skipped.length)}</b></span>
      </div>
      ${sec.join('')}
    </div>
  </div>`;
  const bc = $('#btnSumClear');
  if (bc) bc.onclick = () => { S.lastImport = null; renderImportSummary(); };
}

/* 表头指纹：与后端 core/recipe.py 的 head_key 同一套规则
   —— 全角转半角、去空白标点、转小写、排序后拼接。
   同一套表头（列名相同、顺序不同）算同一组。 */
function normColName(s) {
  return String(s == null ? '' : s)
    .replace(/[\uFF01-\uFF5E]/g, (c) => String.fromCharCode(c.charCodeAt(0) - 0xFEE0))
    .toLowerCase()
    .replace(/[\s_\-()\[\]{}<>:：,，.。、/\\'"]/g, '');
}
function headKeyOf(cols) {
  return (cols || []).map(normColName).filter(Boolean).sort().join('|');
}

function pendingGroups() {
  const groups = [], idx = {};
  let empties = 0;
  const broken = [];
  const skip = S.skip || (S.skip = new Set());
  // 每次渲染都会重建分组对象，识别结果/人工选择必须存在外面（按表头指纹索引），
  // 否则「识别完 → 重绘 → 结果丢了 → 再识别」会打成死循环。
  const meta = S.pgMeta || (S.pgMeta = {});
  S.files.forEach((f) => {
    (f.data.sheets || []).forEach((s, si) => {
      if (s.empty) {
        // 疑似损坏不能混进"空表"里，否则用户以为只是没数据，实际是数据丢了
        if (s.broken) {
          broken.push({ name: (f.rel || f.file_name || '') + (s.sheet ? ' / ' + s.sheet : ''),
                        why: s.note || '' });
        } else empties++;
        return;
      }
      if (s.done || skip.has(f.file_id + '|' + s.sheet)) return;
      const cols = s.columns || [];
      const key = headKeyOf(cols);
      let g = idx[key];
      if (!g) {
        const m = meta[key] || {};
        g = idx[key] = {
          key, cols, items: [],
          target: m.target || '', detected: m.detected || null, touched: !!m.touched,
          header_row: m.header_row || s.header_row || 1,
          header_rows: m.header_rows || s.header_rows || 1
        };
        groups.push(g);
      }
      g.items.push({ f, si, s });
    });
  });
  S.pgGroups = groups;
  return { groups, empties, broken };
}

/* 把分组上的选择写回指纹缓存 */
function saveGroupMeta(g) {
  if (!g || !g.key) return;
  (S.pgMeta || (S.pgMeta = {}))[g.key] = {
    target: g.target || '', detected: g.detected || null, touched: !!g.touched,
    header_row: g.header_row, header_rows: g.header_rows
  };
}

/* 疑似损坏的工作表：单独红色提示，绝不并入"空表已跳过" */
function brokenWarn(list) {
  const shown = list.slice(0, 6)
    .map((x) => `<li>${esc(x.name)}${x.why ? '：' + esc(x.why) : ''}</li>`).join('');
  return `<div style="border:1px solid var(--danger);border-radius:8px;padding:10px 12px;margin-bottom:10px">
    <b style="color:var(--danger)">${list.length} 张工作表疑似损坏，没有进入待导入</b>
    <div class="small muted" style="margin:4px 0 6px">
      这些表有内容但读不出任何单元格，多是导出中断或文件被工具改过。
      请让客户<b>重新导出对应文件</b>，不要当成空表忽略——那样数据会静悄悄少一块。
    </div>
    <ul class="small" style="margin:0;padding-left:18px">${shown}${list.length > 6 ? `<li>…另 ${list.length - 6} 张</li>` : ''}</ul>
  </div>`;
}

function renderPending() {
  const box = $('#pending');
  if (!box) return;
  if (!S.files.length) { box.innerHTML = ''; return; }
  const { groups, empties, broken } = pendingGroups();
  const totalSheets = groups.reduce((a, g) => a + g.items.length, 0);
  const tail = [
    empties ? `另有 ${empties} 张空表已自动跳过` : '',
    broken.length ? `${broken.length} 张疑似损坏` : ''
  ].filter(Boolean).join('，');
  box.innerHTML = `<div class="card">
    <div class="card-h"><h3>待导入</h3>
      <span class="sub">按表头分成 ${groups.length} 组，一组只需选一次目标表${tail ? `（${tail}）` : ''}</span>
      <span style="flex:1"></span>
      <button class="btn sm primary" id="btnImpAll" ${totalSheets ? '' : 'disabled'}>全部导入（${totalSheets} 张）</button>
    </div>
    <div class="card-b">${broken.length ? brokenWarn(broken) : ''}${groups.map(groupCard).join('')}</div></div>`;
  $('#btnImpAll').onclick = () => importAllPending();
  bindPending();
  autoDetectPending();
}

/* 自动识别表类型：先把「这组是什么」猜好填进去，人只负责复核改错。
   只在还没猜过、且用户没手动选过的组上跑，避免把人的选择覆盖掉。 */
let _detecting = false;
async function autoDetectPending() {
  if (_detecting) return;
  const groups = S.pgGroups || [];
  const todo = [];
  groups.forEach((g, i) => {
    if (!g.detected && !g.touched && g.items.length) todo.push(i);
  });
  if (!todo.length) return;
  _detecting = true;
  try {
    const items = todo.map((i) => ({
      key: String(i),
      file_name: groups[i].items[0].f.rel || groups[i].items[0].f.file_name || '',
      sheet_name: groups[i].items[0].s.sheet || '',
      columns: groups[i].cols || [],
    }));
    const r = await apiPost('/api/detect', { items });
    (r.items || []).forEach((it) => {
      const g = groups[+it.key];
      if (!g) return;
      g.detected = it;
      if (!g.target && !g.touched && it.target) g.target = it.target;
      saveGroupMeta(g);
    });
    renderPending();
  } catch (e) {
    /* 识别失败不挡导入，人照旧手选 */
    groups.forEach((g, i) => { if (todo.includes(i)) g.detected = { target: '', reason: '' }; });
    renderPending();
  } finally { _detecting = false; }
}

function groupCard(g, gi) {
  const nf = new Set(g.items.map((x) => x.f.file_id)).size;
  const head = g.cols.length
    ? esc(g.cols.slice(0, 5).join('、')) + (g.cols.length > 5 ? ' …' : '')
    : '（表头为空）';
  const rows = g.items.map((it, ii) => {
    const s = it.s;
    const dim = it.f.kind === 'csv'
      ? `${s.cols || 0} 列`
      : `${s.rows ? '约 ' + fmtInt(s.rows) + ' 行' : ''}${s.cols ? ' × ' + s.cols + ' 列' : ''}`;
    // 文件夹导入时同名文件可能分布在不同子目录，必须显示相对路径才分得清
    const name = it.f.rel || it.f.file_name || '';
    const sheetName = s.sheet || '';
    const showSheet = sheetName && sheetName !== it.f.file_name;   // csv 的工作表名就是文件名，别再重复一遍
    return `<div class="sheet-row gsub" data-gi="${gi}" data-ii="${ii}" data-fi="${it.f.file_id}" data-si="${it.si}">
      <span class="sn" title="${esc(name)} / ${esc(sheetName)}">${esc(name)}</span>
      ${it.f.converted_from === 'html'
        ? '<span class="tag amber" title="原文件其实是网页表格（HTML），已自动转成 Excel 工作表；表名取自表题或序号">网页表格 · 已转换</span>'
        : ''}
      ${it.f.file_note
        ? `<span class="small muted" title="${esc(it.f.file_note)}">⚠</span>`
        : ''}
      <span class="muted small">${showSheet ? esc(sheetName) : ''}</span>
      <span class="muted small">${dim}</span>
      <button class="btn sm" data-act="preview">预览</button>
      <button class="btn sm" data-act="doimport">只导这张</button>
      <span class="st small"></span>
    </div>`;
  }).join('');
  return `<div class="fcard gcard" data-gi="${gi}">
    <div class="fcard-h">
      <span class="tag green">同表头 ×${g.items.length}</span>
      <span class="nm" title="${esc(g.cols.join('、'))}">${head}<span class="muted small">（${g.cols.length} 列）</span></span>
      <span class="muted small">${nf} 个文件 / ${g.items.length} 张表</span>
      <span class="sp" style="flex:1"></span>
      <button class="btn sm" data-act="toggle">明细</button>
      <button class="btn sm ghost danger" data-act="rmgroup">移除这组</button>
    </div>
    <div class="fcard-b">
      <div class="row" style="gap:12px;flex-wrap:wrap;align-items:center;margin-bottom:8px">
        <span class="small muted">目标表</span>
        <select class="gtgt" data-gi="${gi}">
          <option value="">未指定</option>
          ${TGT_OPTIONS.map(([v, lb]) =>
            `<option value="${v}" ${g.target === v ? 'selected' : ''}>${lb}</option>`).join('')}
        </select>
        ${g.detected && g.detected.target
          ? `<span class="tag ${g.target === g.detected.target ? 'green' : 'amber'}" title="${esc(g.detected.reason || '')}">
              自动识别：${esc(g.detected.label || '')}</span>
             <span class="small muted auto-why">${esc(g.detected.reason || '')}</span>`
          : (g.detected
            ? `<span class="tag gray" title="${esc(g.detected.reason || '')}">未识别，请选</span>
               <span class="small muted auto-why">${esc(g.detected.reason || '')}</span>`
            : '<span class="small muted">识别中…</span>')}
        <span class="small muted">表头行</span>
        <input type="number" min="1" class="ghr" value="${g.header_row}" style="width:56px">
        <span class="small muted">占</span>
        <input type="number" min="1" class="ghrs" value="${g.header_rows}" style="width:46px">
        <span class="small muted">行</span>
        <span class="sep"></span>
        <button class="btn sm primary" data-act="impgroup">导入这组（${g.items.length} 张）</button>
        <span class="gst small muted"></span>
      </div>
      <div class="grp-body" ${g.items.length > 1 ? '' : ''}>${rows}</div>
    </div>
  </div>`;
}

function bindPending() {
  const box = $('#pending');
  // 人一旦手选过目标表，就别再被自动识别覆盖
  box.querySelectorAll('.gtgt').forEach((sel) => {
    sel.onchange = () => {
      const g = (S.pgGroups || [])[+sel.closest('.gcard').dataset.gi];
      if (g) { g.target = sel.value || ''; g.touched = true; saveGroupMeta(g); }
    };
  });
  box.querySelectorAll('[data-act]').forEach((btn) => {
    btn.onclick = async () => {
      const act = btn.dataset.act;
      const card = btn.closest('.gcard');
      const gi = +card.dataset.gi;
      const g = (S.pgGroups || [])[gi];
      if (!g) return;
      if (act === 'toggle') {
        const body = card.querySelector('.grp-body');
        body.hidden = !body.hidden;
        return;
      }
      if (act === 'rmgroup') {
        g.items.forEach((it) => S.skip.add(it.f.file_id + '|' + it.s.sheet));
        S.pgDirty = true;
        renderPending(); return;
      }
      const cfg = groupSettings(card);
      if (act === 'impgroup') return importItems(g.items, cfg, card.querySelector('.gst'));
      const row = btn.closest('.sheet-row');
      const it = g.items[+row.dataset.ii];
      if (!it) return;
      if (act === 'preview') return showPreview(it.f, it.si);
      if (act === 'doimport') {
        if (!cfg.target) { toast('先给这组选好目标表（序时账 / 科目余额表）', 'err'); return; }
        return importItems([it], cfg, row.querySelector('.st'));
      }
    };
  });
}

function groupSettings(card) {
  const cfg = {
    header_row: +card.querySelector('.ghr').value || 1,
    header_rows: +card.querySelector('.ghrs').value || 1,
    target: card.querySelector('.gtgt').value || ''
  };
  const g = (S.pgGroups || [])[+card.dataset.gi];
  if (g) { g.target = cfg.target; g.header_row = cfg.header_row; g.header_rows = cfg.header_rows; saveGroupMeta(g); }
  return cfg;
}

async function importAllPending() {
  if (S.importing) { toast('还有导入任务在进行中，请稍等', 'err'); return; }
  const groups = S.pgGroups || [];
  const jobs = [];
  let noTarget = 0;
  document.querySelectorAll('#pending .gcard').forEach((card) => {
    const g = groups[+card.dataset.gi];
    if (!g || !g.items.length) return;
    const cfg = groupSettings(card);
    if (!cfg.target) { noTarget++; return; }
    jobs.push({ g, cfg, card });
  });
  if (!jobs.length) {
    toast(noTarget ? `有 ${noTarget} 组还没选目标表，选好再导` : '没有待导入的工作表', 'err');
    return;
  }
  for (const job of jobs) {
    await importItems(job.g.items, job.cfg, job.card.querySelector('.gst'));
  }
}

async function importItems(items, cfg, stEl) {
  if (S.importing) { toast('还有导入任务在进行中，请稍等', 'err'); return; }
  S.importing = true;
  let ok = 0, rows = 0, ms = 0;
  try {
    for (const it of items) {
      const s = it.s;
      if (stEl) { stEl.textContent = `导入中…（${s.sheet}）`; }
      const pkey = it.f.file_id + '|' + s.sheet;
      busyProgress('导入 ' + s.sheet, it.f.name || it.f.file_name);
      let poll = null;
      const t0 = Date.now();
      try {
        // 导入是一次长请求，期间轮询后端进度，把百分比喂给进度条
        poll = setInterval(async () => {
          try {
            const p = await apiGet('/api/import/progress?key=' + encodeURIComponent(pkey));
            setProgress(p.pct, p.phase);
          } catch (e) { /* 进度取不到不影响导入本身 */ }
        }, 200);
        const r = await apiPost('/api/import', {
          fid: it.f.file_id, sheet: s.sheet,
          header_row: cfg.header_row, header_rows: cfg.header_rows, target: cfg.target || null
        });
        s.done = true;
        ok++; rows += r.n_rows || 0; ms += Date.now() - t0;
        if (stEl) stEl.innerHTML = `<span class="tag green">已导入 ${fmtInt(r.n_rows)} 行 · ${r.elapsed}s</span>`;
      } catch (e) {
        toast(esc(s.sheet) + '：' + e.message, 'err');
        if (stEl) stEl.innerHTML = `<span class="tag red">失败</span>`;
      } finally {
        if (poll) clearInterval(poll);
        setProgress(100, '完成');
        await new Promise((done) => setTimeout(done, 260));   // 让 100% 露一下再关
        idle();
      }
    }
  } finally { idle(); S.importing = false; }
  if (ok) {
    await refresh();
    renderNav(); renderPending(); renderImported();
    toast(items.length > 1
      ? `已导入 ${ok} 张工作表，共 ${fmtInt(rows)} 行，用时 ${(ms / 1000).toFixed(1)} 秒`
      : `已导入 ${fmtInt(rows)} 行，用时 ${(ms / 1000).toFixed(1)} 秒`, 'ok');
  }
}

async function showPreview(f, si) {
  const s = f.data.sheets[si];
  const prev = s.preview || [];
  const body = `<div class="muted small" style="margin-bottom:8px">
      下面是从文件里读到的前 ${prev.length} 行原始内容。<b>点某一行</b>即可把它设为表头行。</div>
    <div class="tbl-wrap" style="max-height:60vh">
      <table class="grid"><thead><tr><th class="rn">#</th>
      ${(prev[0] || []).map((_, i) => `<th>${String.fromCharCode(65 + (i % 26))}</th>`).join('')}
      </tr></thead><tbody>
      ${prev.map((r, i) => `<tr data-hr="${i + 1}" style="cursor:pointer">
        <td class="rn">${i + 1}</td>
        ${r.map((c) => `<td>${esc(c)}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div>`;
  modal('预览 · ' + s.sheet + '（' + f.name + '）', body,
    `<span class="muted small" style="margin-right:auto">当前表头行：${s.header_row}</span>
     <button class="btn" data-close>关闭</button>`);
  $('#modalBody').querySelectorAll('tr[data-hr]').forEach((tr) => {
    tr.onclick = () => {
      s.header_row = +tr.dataset.hr;
      closeModal(); renderPending();
      toast('表头行已设为第 ' + s.header_row + ' 行');
    };
  });
}

function renderImported() {
  const box = $('#importedBox');
  if (!box) return;
  const sheets = S.state.sheets;
  if (!sheets.length) {
    box.innerHTML = '<div class="empty">还没有导入数据</div>';
    return;
  }
  const picked = S.pickedSids || (S.pickedSids = new Set());
  // 表被删掉后，选中集合里可能残留旧 sid，先清掉
  [...picked].forEach((sid) => { if (!sheets.some((s) => s.sid === sid)) picked.delete(sid); });
  box.innerHTML = `
    <div class="row" style="gap:10px;align-items:center;margin-bottom:8px">
      <span class="small muted">勾选后可批量删除（不影响已合并的标准表）</span>
      <span style="flex:1"></span>
      <button class="btn sm" id="btnPickAll">全选</button>
      <button class="btn sm" id="btnPickNone">取消全选</button>
      <button class="btn sm ghost danger" id="btnDropPicked" ${picked.size ? '' : 'disabled'}>
        删除选中（${picked.size}）</button>
    </div>
    <div class="tbl-wrap" style="max-height:none"><table class="grid">
    <thead><tr><th style="width:34px"></th><th>来源文件</th><th>工作表</th><th>表头行</th><th>列数</th><th>行数</th>
    <th>清洗</th><th>目标表</th><th>映射</th><th>操作</th></tr></thead><tbody>
    ${sheets.map((s) => {
      const m = s.mapping || {};
      const n = Object.keys(m).filter((k) => !k.startsWith('_')).length;
      return `<tr class="${picked.has(s.sid) ? 'picked' : ''}">
        <td><input type="checkbox" class="pk" data-sid="${s.sid}" ${picked.has(s.sid) ? 'checked' : ''}></td>
        <td>${esc(s.file_name)}</td><td>${esc(s.sheet_name)}</td>
        <td>${s.header_row}${s.header_rows > 1 ? '（' + s.header_rows + ' 行）' : ''}</td>
        <td>${s.columns.length}</td>
        <td class="num">${s.stg ? fmtInt(s.n_stg) : fmtInt(s.n_rows)}${s.stg ? `<span class="small muted"> / ${fmtInt(s.n_rows)}</span>` : ''}</td>
        <td>${s.stg ? '<span class="tag green">已清洗</span>'
          : ((s.clean || []).length ? `<span class="tag amber">${s.clean.length} 个动作未应用</span>` : '<span class="tag gray">未清洗</span>')}</td>
        <td>${s.target ? `<span class="tag blue">${TGT_LABEL[s.target] || '未指定'}</span>` : '<span class="tag gray">未指定</span>'}</td>
        <td>${n ? `<span class="tag green">${n} 个字段</span>` : '<span class="tag gray">未映射</span>'}</td>
        <td><button class="btn sm link" data-go-clean="${s.sid}">去清洗</button>
            <button class="btn sm link" data-go-map="${s.sid}">去映射</button>
            <button class="btn sm link" data-drop="${s.sid}">删除</button></td>
      </tr>`;
    }).join('')}
    </tbody></table></div>`;

  const syncBtn = () => {
    const b = $('#btnDropPicked');
    if (b) { b.disabled = !picked.size; b.textContent = `删除选中（${picked.size}）`; }
  };
  box.querySelectorAll('.pk').forEach((cb) => {
    cb.onchange = () => {
      const sid = cb.dataset.sid;
      if (cb.checked) picked.add(sid); else picked.delete(sid);
      const tr = cb.closest('tr');
      if (tr) tr.classList.toggle('picked', cb.checked);
      syncBtn();
    };
  });
  const ba = $('#btnPickAll');
  if (ba) ba.onclick = () => { sheets.forEach((s) => picked.add(s.sid)); renderImported(); };
  const bn = $('#btnPickNone');
  if (bn) bn.onclick = () => { picked.clear(); renderImported(); };
  const bd = $('#btnDropPicked');
  if (bd) bd.onclick = () => dropSheets([...picked]);

  box.querySelectorAll('[data-go-clean]').forEach((b) => {
    b.onclick = () => { loadCleanFor(b.dataset.goClean); go('clean'); };
  });
  box.querySelectorAll('[data-drop]').forEach((b) => {
    b.onclick = () => dropSheets([b.dataset.drop]);
  });
  box.querySelectorAll('[data-go-map]').forEach((b) => {
    b.onclick = () => { S.mapSid = b.dataset.goMap; go('map'); };
  });
}

async function dropSheets(sids) {
  sids = (sids || []).filter(Boolean);
  if (!sids.length) return;
  const msg = sids.length === 1
    ? '删除这张原始表？（不影响已合并的结果表）'
    : `删除选中的 ${sids.length} 张原始表？（不影响已合并的结果表）`;
  if (!confirm(msg)) return;
  try {
    const r = sids.length === 1
      ? await apiPost('/api/drop', { sid: sids[0] })
      : await apiPost('/api/drop', { sids });
    (S.pickedSids || new Set()).clear();
    await refresh(); renderNav(); renderImported();
    toast(`已删除 ${fmtInt((r && r.deleted) || sids.length)} 张`, 'ok');
  } catch (e) { toast(e.message, 'err'); }
}

/* ================================================================== 2 清洗 */
function typeSpec(t) { return (S.cleanTypes || []).find((x) => x.type === t); }
function cloneSteps(a) { return JSON.parse(JSON.stringify(a || [])); }

function cleanSheet() {
  return (S.state.sheets || []).find((s) => s.sid === S.cleanSid) || null;
}

function loadCleanFor(sid) {
  const sh = (S.state.sheets || []).find((s) => s.sid === sid);
  S.cleanSid = sid;
  S.cleanSteps = cloneSteps(sh ? (sh.clean || []) : []);
  S.stepOpen = -1;
  S.pv = null;
  S.cleanLoadedFor = sid;
}

function availCols(sheet) {
  const cols = (sheet.columns || []).filter((c) => c !== '_src_row').slice();
  (S.cleanSteps || []).forEach((st) => {
    if (st.enabled === false) return;
    if (st.type === 'derive_level') {
      const n = st.params.name || '科目级次';
      if (!cols.includes(n)) cols.push(n);
    }
    if (st.type === 'add_l1') {
      const n = st.params.name || '一级科目';
      if (!cols.includes(n)) cols.push(n);
    }
    if (st.type === 'split_aux') {
      ['辅助核算1', '辅助核算2', '辅助核算3'].forEach((n) => { if (!cols.includes(n)) cols.push(n); });
    }
    if (st.type === 'derive_l1') {
      // 规则必须与后端 clean._uniq_col 一字不差，否则下拉里会出现后端产不出的列名
      [st.params.out_code || '一级代码', st.params.out_name || '一级科目'].forEach((base) => {
        let n = base, k = 2;
        while (cols.includes(n)) { n = base + k; k += 1; }
        cols.push(n);
      });
    }
  });
  return cols;
}

function stepSummary(st) {
  const spec = typeSpec(st.type);
  if (!spec) return '';
  const ps = (st.params || {});
  if (st.type === 'drop_empty') {
    const picked = Array.isArray(ps.cols) ? ps.cols : [];
    if (!picked.length) return '所有列都为空才删';
    const body = picked.length > 3 ? picked.slice(0, 3).join('、') + ' 等 ' + picked.length + ' 列' : picked.join('、');
    return (ps.mode === 'any' ? '任一为空就删：' : '这些列全空才删：') + body;
  }
  if (st.type === 'fill_const') {
    if (!ps.name || !ps.value) return '（列名/取值待填）';
    return ps.name + ' = ' + ps.value;
  }
  if (st.type === 'add_l1') {
    if (!ps.col) return '（未选源列）';
    return `${ps.col} 按「${ps.sep || '默认符号'}」取第一段 → ${ps.name || '一级科目'}`;
  }
  if (st.type === 'negate_col') {
    if (!ps.col) return '（未选列）';
    const m = { flip: '整列取反', to_negative: '统一成负数', to_positive: '统一成正数' }[ps.mode || 'flip'];
    return ps.col + ' · ' + (m || ps.mode);
  }
  if (st.type === 'derive_l1') {
    if (!ps.col) return '（未选科目编码列）';
    const pos = ps.length ? ('按前 ' + ps.length + ' 位派生') : '按编码自动推断位数派生';
    return `${ps.col} ${pos} → ${ps.out_code || '一级代码'} / ${ps.out_name || '一级科目'}`;
  }
  const bits = [];
  if (ps.col) bits.push('列=' + ps.col);
  if (Array.isArray(ps.cols) && ps.cols.length) bits.push('列=' + ps.cols.join('/'));
  if (ps.kw) bits.push('关键词=' + ps.kw);
  if (ps.mode) bits.push(ps.mode);
  if (ps.length) bits.push('补到' + ps.length + '位');
  if (ps.sep !== undefined && st.type === 'split_aux') bits.push('分隔' + (ps.sep || '不拆'));
  if (st.type === 'check_balance') bits.push('容差' + (ps.tol ?? 0.01));
  return bits.join('，') || '（默认）';
}

function paramInput(st, i, p, cols) {
  const v = (st.params || {})[p.key];
  if (p.kind === 'col') {
    return `<select data-i="${i}" data-k="${p.key}"><option value="">— 不指定 —</option>
      ${cols.map((c) => `<option ${c === v ? 'selected' : ''}>${esc(c)}</option>`).join('')}</select>`;
  }
  if (p.kind === 'cols') {
    const arr = Array.isArray(v) ? v.slice() : [];
    const n = arr.length;
    return `<div class="cp" data-ci="${i}" data-ck="${p.key}">
      <div class="cp-bar">
        <span class="cp-btn" data-cp="all">全选</span>
        <span class="cp-btn" data-cp="none">清空</span>
        <span class="cp-tip">${n ? '已选 ' + n + ' 列' : '未选 = 全部列'}</span>
      </div>
      <div class="cp-list">${cols.map((c) => `<label class="cp-i"><input type="checkbox" value="${esc(c)}"${arr.indexOf(c) >= 0 ? ' checked' : ''}><span>${esc(c)}</span></label>`).join('')}</div>
    </div>`;
  }
  if (p.kind === 'select') {
    return `<select data-i="${i}" data-k="${p.key}">${(p.options || []).map((o) => {
      const s = String(o);
      const val = s.includes(':') ? s.split(':')[0] : s;
      const lb = s.includes(':') ? s.split(':').slice(1).join(':') : s;
      return `<option value="${esc(val)}" ${String(v) === val ? 'selected' : ''}>${esc(lb)}</option>`;
    }).join('')}</select>`;
  }
  if (p.kind === 'number') {
    return `<input type="number" data-i="${i}" data-k="${p.key}" value="${esc(v === undefined || v === null ? '' : v)}" style="width:90px">`;
  }
  return `<input type="text" data-i="${i}" data-k="${p.key}" value="${esc(v === undefined || v === null ? '' : v)}" style="width:230px">`;
}

function paramsHtml(st, i, cols) {
  const spec = typeSpec(st.type);
  if (!spec) return '';
  return `<div class="small muted" style="margin-bottom:6px">${esc(spec.desc || '')}</div>
    <div class="row" style="gap:14px">
    ${(spec.params || []).map((p) => `<div class="row small" style="gap:6px">
      <span class="muted">${esc(p.label)}</span>${paramInput(st, i, p, cols)}
      ${p.hint ? `<span class="small muted">${esc(p.hint)}</span>` : ''}
    </div>`).join('')}
    </div>`;
}


/* ============================================================ 清洗达标标准
   清洗是整条链路的分水岭：洗过头（删了真实行）和洗不够（留着父科目/汇总层）
   都会让后面的检查、BKD 全盘失真，但两种错误在页面上都只是「数字对不上」。
   所以这里把标准直接摆在操作者眼前，并给一个可点的自检和一段可复制的提示词。
   提示词存在后端（core/audit.py 的 PROMPT），页面只负责展示和复制，避免两处写法不一致。 */
/* 收起状态存 localStorage：以前只记在 DOM 上，切到别的模块再回来就重新渲染成展开，
   用户每次都要再收一次。这个开关属于「使用习惯」，跟数据无关，留在本地最合适。 */
const GUIDE_KEY = 'ptw.cleanGuide.closed';
function guideClosed() {
  try { return localStorage.getItem(GUIDE_KEY) === '1'; } catch (e) { return false; }
}
function setGuideClosed(v) {
  try { localStorage.setItem(GUIDE_KEY, v ? '1' : '0'); } catch (e) { /* 隐私模式忽略 */ }
}

function cleanGuide() {
  const a = S.audit;
  const closed = guideClosed();
  const badge = a
    ? (a.ok ? '<span class="guide-badge" style="background:var(--ok)">已达标</span>'
            : `<span class="guide-badge" style="background:var(--danger)">未达标 ${a.stat.failed} 项</span>`)
    : '';
  return `<div class="card guide" style="margin-bottom:12px">
    <div class="guide-hd">
      <span class="guide-badge">清洗达标标准</span>
      ${badge}
      <b style="font-size:13px">目标是「数据完整 + 最细颗粒度」，达不到不要往下走</b>
      <span style="flex:1"></span>
      <button class="btn sm primary" id="btnAudit">检查清洗达标</button>
      <button class="btn sm" id="btnPrompt">复制 AI 提示词</button>
      <button class="btn sm ghost" id="btnGuide">${closed ? '展开' : '收起'}</button>
    </div>
    <div class="guide-bd" ${closed ? 'hidden' : ''}>
      <div class="guide-cols">
        <div>
          <div class="guide-t">科目余额表必须做到</div>
          <ul>
            <li>只留<b>末级科目</b>，父科目不能和明细并存</li>
            <li>没有合计 / 小计 / 总计行，也没有「核算维度=合计」的汇总行</li>
            <li>银行存款要带<b>账号</b>，往来科目要带<b>客商</b></li>
            <li>一级科目已派生，编码和名称不为空</li>
          </ul>
        </div>
        <div>
          <div class="guide-t">序时账必须做到</div>
          <ul>
            <li>日期、凭证号、摘要、科目编码都不为空</li>
            <li>借贷金额不能出现两边都为空的行</li>
            <li>一级科目已派生，科目全称尽量有</li>
            <li>客商 / 部门 / 项目辅助核算不要整列空</li>
          </ul>
        </div>
        <div>
          <div class="guide-t">怎么验收</div>
          <ul>
            <li>点「检查清洗达标」，逐项过，不达标按提示改</li>
            <li>合并后在「数据检查」跑五组，全部一致</li>
            <li>输出每项结果 + 整体通过/不通过结论</li>
            <li>差额必须能解释，禁止为了凑平去改数据</li>
          </ul>
        </div>
      </div>
      <div id="auditOut">${a ? auditHtml(a) : ''}</div>
    </div>
  </div>`;
}

function auditHtml(a) {
  if (!a) return '';
  const dot = (i) => i.ok
    ? '<span class="audit-dot" style="background:var(--ok)"></span>'
    : (i.level === 'warn'
      ? '<span class="audit-dot" style="background:var(--warn)"></span>'
      : '<span class="audit-dot" style="background:var(--danger)"></span>');
  const groups = {};
  (a.items || []).forEach((i) => { (groups[i.group] = groups[i.group] || []).push(i); });
  const body = Object.keys(groups).map((g) => `<div style="margin-top:10px">
      <div class="guide-t">${esc(g)}</div>
      ${groups[g].map((i) => `<div class="audit-row">
        ${dot(i)}
        <div style="width:168px;flex:none">${esc(i.title)}</div>
        <div style="width:132px;flex:none;color:var(--muted)">${esc(i.value)}</div>
        <div style="flex:1">${esc(i.detail)}${i.ok ? '' : (i.fix
          ? `<div class="small" style="color:var(--accent);margin-top:2px">改法：${esc(i.fix)}</div>` : '')}</div>
      </div>`).join('')}
    </div>`).join('');
  const st = a.stat || {};
  return `<div style="margin-top:12px;padding:10px 12px;border-radius:var(--radius);
      background:${a.ok ? 'var(--ok-soft)' : 'var(--danger-soft)'}">
      <b style="color:${a.ok ? 'var(--ok)' : 'var(--danger)'}">${a.ok ? '达标' : '未达标'}</b>
      <span class="small" style="margin-left:8px">${esc(a.verdict)}</span>
      <span class="small" style="float:right;color:var(--muted2)">${st.passed}/${st.total} 项通过</span>
    </div>${body}`;
}

function bindCleanGuide() {
  const tog = $('#btnGuide');
  if (tog) tog.onclick = () => {
    const bd = document.querySelector('.guide-bd');
    if (!bd) return;
    bd.hidden = !bd.hidden;
    tog.textContent = bd.hidden ? '展开' : '收起';
    setGuideClosed(bd.hidden);
  };
  const bA = $('#btnAudit');
  if (bA) bA.onclick = async () => {
    const box = $('#auditOut');
    if (box) box.innerHTML = '<div class="small muted" style="margin-top:10px">正在判定…</div>';
    try {
      const r = await apiGet('/api/clean/audit');
      S.audit = r;
      if (box) box.innerHTML = auditHtml(r);
      toast(r.ok ? '清洗达标，可以合并' : `有 ${r.stat.failed} 项未达标`, r.ok ? 'ok' : 'err');
    } catch (e) {
      if (box) box.innerHTML = `<div class="small" style="color:var(--danger)">${esc(e.message)}</div>`;
    }
  };
  const bP = $('#btnPrompt');
  if (bP) bP.onclick = async () => {
    let p = (S.audit || {}).prompt;
    if (!p) {
      try { const r = await apiGet('/api/clean/audit'); S.audit = r; p = r.prompt; }
      catch (e) { p = null; }
    }
    if (p) copyText(p); else toast('取提示词失败，请再点一次', 'err');
  };
}

function copyText(t) {
  const done = () => toast('已复制，可直接粘给 AI', 'ok');
  const fallback = () => {
    const ta = document.createElement('textarea');
    ta.value = t;
    ta.style.cssText = 'position:fixed;top:0;left:0;opacity:0';
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand('copy'); done(); }
    catch (e) {
      modal('自动复制失败，请手动全选复制', `<pre class="small" style="white-space:pre-wrap;max-height:52vh;overflow:auto">${esc(t)}</pre>`, '');
    }
    document.body.removeChild(ta);
  };
  if (navigator.clipboard && navigator.clipboard.writeText) {
    navigator.clipboard.writeText(t).then(done).catch(fallback);
  } else fallback();
}

function renderClean() {
  const st = S.state;
  if (!st.sheets.length) {
    $('#main').innerHTML = subTabBar('clean', 'clean') + '<div class="card"><div class="empty"><h2>还没有导入数据</h2>先去「导入数据」把文件拖进来。</div></div>';
    return;
  }
  if (!S.cleanSid || S.cleanLoadedFor !== S.cleanSid) loadCleanFor(S.cleanSid || st.sheets[0].sid);
  const cur = cleanSheet() || st.sheets[0];
  const cols = availCols(cur);

  $('#main').innerHTML = subTabBar('clean', 'clean') + cleanGuide() + `
    <div class="split">
      <div class="col-list">
        <div class="card">
          <div class="card-h"><h3>原始工作表</h3><span class="sub">${st.sheets.length} 张</span></div>
          <div class="card-b" style="padding:6px">
            ${st.sheets.map((s) => `<div class="ritem ${s.sid === S.cleanSid ? 'on' : ''}" data-csid="${s.sid}">
              <div style="flex:1;min-width:0">
                <div style="overflow:hidden;text-overflow:ellipsis">${esc(s.sheet_name)}</div>
                <div class="small" style="color:${s.sid === S.cleanSid ? 'var(--accent)' : 'var(--muted2)'}">
                  ${esc(s.file_name)} · ${fmtInt(s.stg ? s.n_stg : s.n_rows)} 行${s.stg ? '（已清洗）' : ''}</div>
              </div>
              ${s.stg ? '<span class="tag green">清</span>' : ''}
            </div>`).join('')}
          </div>
        </div>
        <div class="card" style="margin-top:10px">
          <div class="card-b small muted">
            清洗只写「清洗表」，<b>原始表一列不动</b>。剔除的行全部记录在案，可在下方查看。
          </div>
        </div>
      </div>
      <div class="col-main" id="cleanMain"></div>
    </div>`;

  $('#main').querySelectorAll('[data-csid]').forEach((el) => {
    el.onclick = () => { loadCleanFor(el.dataset.csid); renderClean(); };
  });
  bindCleanGuide();
  renderCleanMain(cur, cols);
}

function renderCleanMain(sheet, cols) {
  const box = $('#cleanMain');
  const applied = sheet.stg;
  // 有多少张表配了动作——「全部清洗」只在有意义时可选
  const readyN = ((S.state || {}).sheets || [])
    .filter((s) => (s.clean || []).some((x) => x.enabled !== false)).length;
  box.innerHTML = `
    <div class="card">
      <div class="card-h">
        <h3>${esc(sheet.sheet_name)}</h3>
        <span class="sub">${fmtInt(sheet.n_rows)} 行原始数据${applied ? ` → 已清洗为 ${fmtInt(sheet.n_stg)} 行` : ''}</span>
        <span style="flex:1"></span>
        <select id="presetSel" style="max-width:190px">
          ${(S.presets || []).map((p) => `<option value="${p.key}" title="${esc(p.desc || '')}">${esc(p.name)}</option>`).join('')}
        </select>
        <button class="btn sm" id="btnPreset">套用配方</button>
        <span class="sep"></span>
        <select id="recipeSel" style="max-width:190px" title="已保存的清洗方案，可直接套到这张表">
          <option value="">${(S.recipes || []).length ? '— 我的方案 —' : '— 还没有存过方案 —'}</option>
          ${(S.recipes || []).map((r) => `<option value="${esc(r.name)}" ${S.recipeSel === r.name ? 'selected' : ''}>${esc(r.name)}（${r.n} 步）</option>`).join('')}
        </select>
        <button class="btn sm" id="btnRecipeUse" ${(S.recipes || []).length ? '' : 'disabled'}>套用</button>
        <button class="btn sm primary" id="btnRecipeSave" ${S.cleanSteps.length ? '' : 'disabled'}>存为方案</button>
        ${(S.recipes || []).length ? '<button class="btn sm ghost danger" id="btnRecipeDel">删除</button>' : ''}
      </div>
      <div class="card-b">
        <div id="stepBox">${(S.cleanSteps.length ? S.cleanSteps : []).map((s, i) => `
          <div class="step-row ${S.stepOpen === i ? 'on' : ''}" data-row="${i}">
            <label class="ck" onclick="event.stopPropagation()"><input type="checkbox" data-en="${i}" ${s.enabled === false ? '' : 'checked'}></label>
            <span class="sno">${i + 1}</span>
            <span class="snm">${esc((typeSpec(s.type) || {}).name || s.type)}</span>
            <span class="small muted ssum">${esc(stepSummary(s))}</span>
            <span style="flex:1"></span>
            <button class="btn sm link" data-up="${i}" ${i === 0 ? 'disabled' : ''}>↑</button>
            <button class="btn sm link" data-down="${i}" ${i === S.cleanSteps.length - 1 ? 'disabled' : ''}>↓</button>
            <button class="btn sm link danger" data-del="${i}">删除</button>
          </div>
          <div class="step-params" ${S.stepOpen === i ? '' : 'hidden'}>${paramsHtml(s, i, cols)}</div>`).join('')
        || `<div class="empty" style="padding:20px;line-height:1.9">
              这张表还没有任何清洗动作。<b>系统不预置默认动作</b>——每张表的脏法都不一样，
              预设的配方删错一行比少删一行危险得多。<br>
              三条路任选：① 上面下拉框选一个配方，点「套用配方」；
              ② 右边「导入清洗方案」拿以前验证过的方案（.json）；
              ③ 从下面「添加动作」一个个配。<br>
              配得对不对不用猜：上面「检查清洗达标」逐项告诉你还差什么，最后拿「数据检查」五组兜底。
            </div>`}
        </div>
        <div class="row" style="margin-top:10px">
          <span class="small muted">添加动作</span>
          <select id="addSel" style="max-width:220px">
            ${(S.cleanTypes || []).map((t) => `<option value="${t.type}">${esc(t.group)} · ${esc(t.name)}</option>`).join('')}
          </select>
          <button class="btn sm" id="btnAdd">添加</button>
          <span class="sep"></span>
          <button class="btn sm primary" id="btnDry">试运行</button>
          <button class="btn sm primary" id="btnApply" ${S.cleanSteps.length ? '' : 'disabled'}>应用清洗</button>
          <button class="btn sm" id="btnApplyAll" ${readyN ? '' : 'disabled'}
            title="把所有配置了动作的工作表一次清洗完，不用一张张点；只算已启用的动作">全部清洗${readyN ? `（${readyN} 表）` : ''}</button>
          <button class="btn sm" id="btnDropRows">看被剔除的行</button>
          <span class="sep"></span>
          <button class="btn sm" id="btnRecipeOut" title="把所有工作表的清洗动作导出成 xlsx，一种模式一个 sheet，换版本后导入即可复用">导出清洗方案</button>
          <button class="btn sm" id="btnRecipeIn" title="导入之前导出的清洗方案 xlsx，套用到当前工作表">导入清洗方案</button>
          <input type="file" id="recipeFile" accept=".xlsx,.xlsm" hidden>
          ${applied ? '<button class="btn sm ghost danger" id="btnReset">恢复到原始</button>' : ''}
        </div>
        <div class="hint">「试运行」只算不落库，会显示每步删了几行；「应用清洗」把结果写成清洗表，后面映射/合并用的就是它。</div>
      </div>
    </div>
    <div id="pvBox">${S.pv ? pvHtml(S.pv) : ''}</div>`;

  bindClean(sheet, cols);
  if (S.pv) bindPv();
}

function bindClean(sheet, cols) {
  const box = $('#cleanMain');
  box.querySelectorAll('[data-row]').forEach((el) => {
    el.onclick = () => {
      const i = +el.dataset.row;
      S.stepOpen = (S.stepOpen === i ? -1 : i);
      renderCleanMain(sheet, cols);
    };
  });
  box.querySelectorAll('[data-en]').forEach((el) => {
    el.onchange = () => {
      S.cleanSteps[+el.dataset.en].enabled = el.checked;
      saveClean();
    };
  });
  box.querySelectorAll('[data-up]').forEach((b) => {
    b.onclick = (e) => { e.stopPropagation(); moveStep(+b.dataset.up, -1); };
  });
  box.querySelectorAll('[data-down]').forEach((b) => {
    b.onclick = (e) => { e.stopPropagation(); moveStep(+b.dataset.down, 1); };
  });
  box.querySelectorAll('[data-del]').forEach((b) => {
    b.onclick = (e) => {
      e.stopPropagation();
      S.cleanSteps.splice(+b.dataset.del, 1);
      S.stepOpen = -1; saveClean(); renderCleanMain(sheet, cols);
    };
  });
  box.querySelectorAll('.cp').forEach((el) => {
    const i = +el.dataset.ci;
    const k = el.dataset.ck;
    const sync = () => {
      const arr = Array.from(el.querySelectorAll('input')).filter((x) => x.checked).map((x) => x.value);
      S.cleanSteps[i].params[k] = arr;
      const tip = el.querySelector('.cp-tip');
      if (tip) tip.textContent = arr.length ? '已选 ' + arr.length + ' 列' : '未选 = 全部列';
      // 只刷新这一行的摘要，不整页重绘，免得滚动位置被顶回去
      const sum = box.querySelector(`.step-row[data-row="${i}"] .ssum`);
      if (sum) sum.textContent = stepSummary(S.cleanSteps[i]);
      saveClean();
    };
    el.querySelectorAll('input').forEach((x) => { x.onchange = sync; });
    el.querySelectorAll('[data-cp]').forEach((b) => {
      b.onclick = (e) => {
        e.preventDefault();
        const on = b.dataset.cp === 'all';
        el.querySelectorAll('input').forEach((x) => { x.checked = on; });
        sync();
      };
    });
  });
  box.querySelectorAll('[data-k]').forEach((el) => {
    el.onchange = () => {
      const i = +el.dataset.i; const k = el.dataset.k;
      if (el.dataset.multi) {
        S.cleanSteps[i].params[k] = Array.from(el.selectedOptions).map((o) => o.value);
      } else {
        S.cleanSteps[i].params[k] = el.value;
      }
      saveClean();
      renderCleanMain(sheet, cols);
    };
  });
  $('#btnAdd').onclick = () => {
    const type = $('#addSel').value;
    const spec = typeSpec(type);
    const params = {};
    ((spec || {}).params || []).forEach((p) => {
      params[p.key] = p.kind === 'cols' ? [] : (p.default === undefined ? '' : p.default);
    });
    S.cleanSteps.push({ id: 's' + Date.now().toString(36), type, enabled: true, params });
    S.stepOpen = S.cleanSteps.length - 1;
    saveClean(); renderCleanMain(sheet, cols);
  };
  $('#btnPreset').onclick = async () => {
    const key = $('#presetSel').value;
    if (S.cleanSteps.length && !confirm('套用配方会覆盖当前的动作列表，继续？')) return;
    try {
      const r = await apiPost('/api/clean/preset', { sid: sheet.sid, key });
      S.cleanSteps = r.steps || [];
      S.stepOpen = -1; S.pv = null;
      await saveClean();
      renderCleanMain(sheet, cols);
      toast('已套用配方，共 ' + S.cleanSteps.length + ' 步', 'ok');
    } catch (e) { toast(e.message, 'err'); }
  };

  // ---- 方案库（右上角下拉框）
  const rsel = $('#recipeSel');
  if (rsel) rsel.onchange = () => { S.recipeSel = rsel.value; };
  const rUse = $('#btnRecipeUse');
  if (rUse) rUse.onclick = async () => {
    const name = $('#recipeSel').value;
    if (!name) { toast('先在下拉框里选一个方案', 'err'); return; }
    if (S.cleanSteps.length && !confirm(`套用方案「${name}」会覆盖当前的动作列表，继续？`)) return;
    try {
      const r = await guard('正在套用方案…', () => apiPost('/api/recipes/apply', { name, sid: sheet.sid }));
      S.recipeSel = name;
      await refresh();
      S.cleanLoadedFor = null;
      renderClean();
      (r.warnings || []).slice(0, 3).forEach((w) => toast(w, 'err'));
      toast(`已套用「${name}」：${r.n} 步` + (r.mapping ? `，同时套用 ${r.mapping} 项字段映射` : ''), 'ok');
    } catch (e) { toast(e.message, 'err'); }
  };
  const rSave = $('#btnRecipeSave');
  if (rSave) rSave.onclick = () => saveRecipeModal(sheet);
  const rDel = $('#btnRecipeDel');
  if (rDel) rDel.onclick = async () => {
    const name = $('#recipeSel').value;
    if (!name) { toast('先在下拉框里选一个方案', 'err'); return; }
    if (!confirm(`从方案库里删除「${name}」？（不影响已经配好的工作表）`)) return;
    try {
      await apiPost('/api/recipes/delete', { name });
      S.recipeSel = '';
      await loadRecipes();
      renderClean();
      toast('已删除方案', 'ok');
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#btnDry').onclick = async () => {
    try {
      const r = await guard('试运行清洗…', () => apiPost('/api/clean/preview', { sid: sheet.sid, sample: 20 }));
      S.pv = r;
      renderCleanMain(sheet, cols);
      toast(`试运行完成：${fmtInt(r.total_before)} → ${fmtInt(r.total_after)} 行`, r.error ? 'err' : 'ok');
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#btnApply').onclick = async () => {
    let r;
    try {
      r = await guard('正在清洗并落库…', () => apiPost('/api/clean/apply', { sid: sheet.sid }));
    } catch (e) { toast(e.message, 'err'); return; }
    if (r.error) {
      S.pv = r;
      try { renderCleanMain(sheet, cols); } catch (e) { console.error(e); }
      toast(r.error, 'err');
      return;
    }
    // 落库已经成功，后面任何渲染问题都不该让用户以为清洗失败
    S.pv = r;
    delete S.previewCache[sheet.sid];
    try { await refresh(); renderNav(); } catch (e) { console.error(e); }
    // 跑通了就自动进方案库，右上角下拉框以后能直接选
    if (r.recipe && r.recipe.name) {
      try { await loadRecipes(); S.recipeSel = r.recipe.name; } catch (e) { /* 忽略 */ }
    }
    const sh = cleanSheet() || sheet;
    try { renderCleanMain(sh, availCols(sh)); } catch (e) { console.error(e); }
    toast(`已应用清洗：${fmtInt(r.total_after)} 行（剔除 ${fmtInt(r.dropped)} 行）`, 'ok');
    if (r.recipe && r.recipe.name && !r.recipe.exists) {
      toast(`这套方案已存为「${r.recipe.name}」，以后可在右上角直接套用`, 'ok');
    }
  };
  const bAll = $('#btnApplyAll');
  if (bAll) bAll.onclick = () => applyAllClean();
  $('#btnDropRows').onclick = showDropped;
  $('#btnRecipeOut').onclick = async () => {
    try {
      const r = await guard('正在导出清洗方案…', () => apiPost('/api/clean/recipe/export'));
      toast(`已导出 ${r.plans} 个方案 / ${r.sheets} 张表`, 'ok');
      window.location.href = r.url;
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#btnRecipeIn').onclick = () => $('#recipeFile').click();
  $('#recipeFile').onchange = async (ev) => {
    const f = ev.target.files && ev.target.files[0];
    ev.target.value = '';
    if (!f) return;
    try {
      const up = await guard('上传清洗方案…', () => uploadFile(f));
      const r = await guard('解析清洗方案…', () => apiPost('/api/clean/recipe/parse', { fid: up.file_id }));
      showRecipePicker(r, sheet);
    } catch (e) { toast(e.message, 'err'); }
  };
  const rst = $('#btnReset');
  if (rst) {
    rst.onclick = async () => {
      if (!confirm('放弃清洗结果，恢复成原始数据？（动作配置会保留）')) return;
      await apiPost('/api/clean/reset', { sid: sheet.sid });
      delete S.previewCache[sheet.sid];
      await refresh();
      const sh = cleanSheet();
      renderCleanMain(sh || sheet, availCols(sh || sheet));
      toast('已恢复到原始数据');
    };
  }
}

/* 一键清洗：把所有配了动作的工作表逐张洗一遍。
   走 /api/clean/apply-all，后端逐表复用「应用清洗」的逻辑（含被剔除行留痕、方案自动登记），
   单张出错不中断其余表。原来的逐张「应用清洗」完全保留。 */
async function applyAllClean() {
  const sheets = (S.state || {}).sheets || [];
  const ready = sheets.filter((s) => (s.clean || []).some((x) => x.enabled !== false));
  const idle = sheets.length - ready.length;
  if (!ready.length) {
    toast('还没有任何工作表配置了清洗动作。先套用配方，或从右上角下拉框选方案', 'err');
    return;
  }
  const tip = `将对 ${ready.length} 张表按各自配置的动作为依次清洗`
    + (idle ? `，另外 ${idle} 张没配动作会跳过` : '')
    + '。\n\n已清洗过的会按最新配置重算，原始数据不受影响。继续？';
  if (!confirm(tip)) return;
  let r;
  try {
    r = await guard(`正在批量清洗 ${ready.length} 张表…`, () => apiPost('/api/clean/apply-all'));
  } catch (e) { toast(e.message, 'err'); return; }
  try { await refresh(); } catch (e) { console.error(e); }
  S.pv = null;
  S.previewCache = {};
  try { renderClean(); } catch (e) { console.error(e); }
  showApplyAll(r);
  if (r.failed) toast(`批量清洗完成，${r.failed} 张失败，详见结果列表`, 'err');
  else toast(`批量清洗完成：${r.ok_n} 张表，用时 ${r.seconds}s`, 'ok');
}

function showApplyAll(r) {
  const rows = (r.items || []).map((it) => {
    const tag = it.status === 'ok' ? '<span class="tag green">已清洗</span>'
      : it.status === 'skip' ? '<span class="tag">跳过</span>'
        : '<span class="tag red">失败</span>';
    const desc = it.status === 'ok'
      ? `${fmtInt(it.before)} → ${fmtInt(it.after)} 行（剔除 ${fmtInt(it.dropped)}）`
      : esc(it.reason || it.error || '');
    return `<tr><td>${esc(it.sheet || '—')}</td>
      <td class="small muted">${esc(it.file || '—')}</td>
      <td>${tag}</td>
      <td class="small">${desc}</td></tr>`;
  }).join('');
  modal('批量清洗结果', `
    <div class="row" style="gap:12px;margin-bottom:10px;flex-wrap:wrap">
      <span class="tag green">成功 ${r.ok_n}</span>
      <span class="tag">跳过 ${r.skipped}</span>
      ${r.failed ? `<span class="tag red">失败 ${r.failed}</span>` : ''}
      <span class="small muted">共 ${r.total} 张表 / ${r.seconds}s</span>
    </div>
    <div class="tbl-wrap" style="max-height:420px"><table class="grid">
      <thead><tr><th>工作表</th><th>文件</th><th>结果</th><th>行数变化 / 说明</th></tr></thead>
      <tbody>${rows || '<tr><td colspan="4" class="empty">没有可清洗的表</td></tr>'}</tbody>
    </table></div>
    <div class="hint" style="margin-top:8px">清洗结果已落库，后面「字段映射」和「合并成表」用到的就是它。
      想改某张表的动作，左边选那张表再单独「应用清洗」即可。</div>`,
    '<button class="btn primary" data-close>知道了</button>');
  bindClose();
}

function moveStep(i, d) {
  const j = i + d;
  if (j < 0 || j >= S.cleanSteps.length) return;
  const a = S.cleanSteps;
  [a[i], a[j]] = [a[j], a[i]];
  S.stepOpen = j;
  saveClean();
  const sheet = cleanSheet();
  renderCleanMain(sheet, availCols(sheet));
}

let cleanTimer = null;
function saveClean() {
  clearTimeout(cleanTimer);
  cleanTimer = setTimeout(async () => {
    try { await apiPost('/api/clean/save', { sid: S.cleanSid, steps: S.cleanSteps }); }
    catch (e) { toast('保存清洗配置失败：' + e.message, 'err'); }
  }, 350);
}

function saveRecipeModal(sheet) {
  if (!S.cleanSteps.length) { toast('当前还没有动作，先配好再存', 'err'); return; }
  const def = String(sheet.sheet_name || '清洗方案')
    .replace(/\.(xlsx|xlsm|xls|csv|txt)$/i, '').slice(0, 30);
  modal('存为清洗方案', `
    <div class="small muted" style="margin-bottom:10px">
      当前 ${S.cleanSteps.length} 个动作会存进方案库，以后在右上角下拉框里直接套用；同名方案会被覆盖。</div>
    <input id="recipeName" style="width:100%" value="${esc(def)}" placeholder="方案名，例如 余额表-通用">
    <div class="hint">方案只记录动作，不绑表头。套到别的表时，表里没有的列对应的动作会自动跳过并提示。</div>`,
    `<button class="btn" data-close>取消</button>
     <button class="btn primary" id="btnRecipeOk">保存</button>`);
  const submit = async () => {
    const name = ($('#recipeName').value || '').trim();
    if (!name) { toast('方案名不能为空', 'err'); return; }
    try {
      const r = await apiPost('/api/recipes/save',
        { name, steps: S.cleanSteps, sid: sheet.sid });
      await loadRecipes();
      S.recipeSel = r.name;
      closeModal();
      renderClean();
      toast('已存入方案库：' + r.name, 'ok');
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#btnRecipeOk').onclick = submit;
  $('#recipeName').onkeydown = (e) => { if (e.key === 'Enter') submit(); };
  setTimeout(() => {
    const i = $('#recipeName');
    if (i) { i.focus(); i.select(); }
  }, 30);
}

function pvHtml(pv) {
  pv = pv || {};
  const pvRows = pv.rows || [];
  const rawRows = pv.raw_rows || [];
  const rows = (pv.steps || []).map((s) => `<tr>
      <td class="rn">${s.i}</td><td>${esc(s.name)}</td>
      <td class="num">${s.before === null || s.before === undefined ? '-' : fmtInt(s.before)}</td>
      <td class="num">${s.after === null || s.after === undefined ? '-' : fmtInt(s.after)}</td>
      <td class="num">${s.dropped ? '<b style="color:var(--warn)">−' + fmtInt(s.dropped) + '</b>' : (s.after === null ? '' : '0')}</td>
      <td class="small muted">${esc(s.error || s.note || '')}</td></tr>`).join('');
  const checks = (pv.checks || []).map((c) => `<span class="tag ${c.bad ? 'red' : 'green'}">
      ${esc(c.name)}：${fmtInt(c.rows)} 行，不平 ${fmtInt(c.bad)} 行，差额合计 ${fmtCell(c.diff)}</span>`).join(' ');
  return `<div class="card">
    <div class="card-h"><h3>${(pv.stg || pv.mode === 'apply') ? '清洗已应用' : '试运行结果'}</h3>
      <span class="sub">${fmtInt(pv.total_before)} → <b>${fmtInt(pv.total_after)}</b> 行（剔除 ${fmtInt(pv.dropped)}）</span>
      <span style="flex:1"></span>${checks}</div>
    <div class="card-b">
      ${pv.error ? `<div class="tag red" style="margin-bottom:8px">${esc(pv.error)}</div>` : ''}
      <table class="grid" style="width:100%">
        <thead><tr><th class="rn">#</th><th>动作</th><th>前</th><th>后</th><th>剔除</th><th>说明</th></tr></thead>
        <tbody>${rows || '<tr><td colspan="6" class="empty">没有动作</td></tr>'}</tbody></table>
      <div class="row" style="margin-top:12px;align-items:flex-start;gap:14px">
        <div style="flex:1;min-width:0">
          <div class="small muted" style="margin-bottom:4px">原始前 ${fmtInt(rawRows.length)} 行</div>
          <div class="tbl-wrap mini">${tableInner(pv.raw_columns || [], rawRows)}</div>
        </div>
        <div style="flex:1;min-width:0">
          <div class="small muted" style="margin-bottom:4px">清洗后前 ${fmtInt(pvRows.length)} 行</div>
          <div class="tbl-wrap mini">${tableInner(pv.columns || [], pvRows)}</div>
        </div>
      </div>
    </div></div>`;
}

function bindPv() { /* 结果表无交互 */ }

async function showDropped() {
  const sheet = cleanSheet();
  if (!sheet) return;
  try {
    const r = await guard('读取剔除记录…', () => apiPost('/api/clean/dropped', { sid: sheet.sid, limit: 300 }));
    if (!r.total) {
      modal('被剔除的行', '<div class="empty">还没有剔除记录。先「试运行」或「应用清洗」。</div>');
      bindClose(); return;
    }
    const head = (r.cols || []).map((c) => `<th>${esc(c)}</th>`).join('');
    modal('被剔除的行（共 ' + fmtInt(r.total) + ' 行）', `
      <div class="hint" style="margin-bottom:8px">每条都记录了是第几步、按什么规则剔除的，可回溯。</div>
      <div class="tbl-wrap" style="max-height:64vh"><table class="grid"><thead><tr>
        <th class="rn">原始行</th><th>步骤</th><th>动作</th><th>原因</th>${head}</tr></thead><tbody>
        ${r.rows.map((row, i) => `<tr><td class="rn">${row[0]}</td>
          <td>${row[1]}</td><td>${esc(row[2])}</td><td>${esc(row[3])}</td>
          ${(r.preview[i] || []).map((v) => `<td>${esc(v)}</td>`).join('')}</tr>`).join('')}
      </tbody></table></div>`, '<button class="btn" data-close>关闭</button>');
    bindClose();
  } catch (e) { toast(e.message, 'err'); }
}

/* ---- 清洗方案：挑一个方案套到当前工作表 */
function showRecipePicker(r, sheet) {
  const plans = r.plans || [];
  if (!plans.length) {
    modal('导入清洗方案', `<div class="empty">这个文件里没有解析出任何方案。
      ${(r.warnings || []).map((w) => `<div class="tag amber" style="margin-top:8px">${esc(w)}</div>`).join('')}
      </div>`, '<button class="btn" data-close>关闭</button>');
    bindClose(); return;
  }
  // 映射按表头匹配，不按文件名：同一套表头在任何文件里都能自动认出来
  const byHead = r.maps_hk || {};
  const hk = headKeyOf(sheet.columns || []);
  const hit = byHead[hk] || null;
  const hitN = hit ? Object.keys(hit.mapping || {}).length : 0;
  const warns = (r.warnings || []).map((w) => `<div class="tag amber" style="margin-bottom:6px">${esc(w)}</div>`).join('');
  const body = `
    <div class="hint" style="margin-bottom:10px">选一个方案套用到当前工作表「${esc(sheet.sheet_name)}」，
      会<b>覆盖</b>这张表现有的动作清单（原始数据不受影响）。</div>
    ${warns}
    <div class="rc-list">${plans.map((p, i) => `
      <label class="rc-i">
        <input type="radio" name="rc" value="${i}" ${i === 0 ? 'checked' : ''}>
        <div style="flex:1;min-width:0">
          <div><b>${esc(p.name)}</b> <span class="small muted">${p.n} 个动作</span></div>
          <div class="small muted">原用于：${(p.uses || []).map((u) => esc(u.file + ' / ' + u.sheet)).join('；') || '—'}</div>
          <div class="small" style="color:var(--muted2)">${p.steps.map((s) => esc((typeSpec(s.type) || {}).name || s.type)).join(' → ')}</div>
        </div>
      </label>`).join('')}</div>
    ${hitN ? `<label class="row small" style="gap:6px;margin-top:10px">
        <input type="checkbox" id="rcMap" checked>同时套用字段映射（这个表头下有 ${hitN} 个字段的映射备份，只对同名列生效）
      </label>` : `<div class="small muted" style="margin-top:10px">
        这个文件里没有和当前表头（${esc((sheet.columns || []).slice(0, 3).join('、') || '—')}…）对得上的字段映射，
        只套用清洗动作。</div>`}
    <label class="row small" style="gap:6px;margin-top:8px">
      <input type="checkbox" id="rcSaveLib" checked>把文件里的 ${plans.length} 个方案存入方案库
      <span class="muted">（以后任何工作表都能在右上角下拉框直接选，不用再导这个文件）</span>
    </label>`;
  modal('导入清洗方案', body,
    '<button class="btn" data-close>取消</button><button class="btn primary" id="rcOk">套用</button>');
  bindClose();
  $('#rcOk').onclick = async () => {
    const sel = document.querySelector('input[name="rc"]:checked');
    if (!sel) return;
    const p = plans[+sel.value];
    const useMap = $('#rcMap') && $('#rcMap').checked;
    const saveLib = $('#rcSaveLib') && $('#rcSaveLib').checked;
    // 先存进方案库：存一次，其余工作表在右上角下拉框直接选，不用再导一遍文件
    let savedN = 0;
    if (saveLib) {
      try {
        const sv = await guard('存入方案库…', () => apiPost('/api/recipes/save-many', {
          plans: plans.map((x) => ({ name: x.name, steps: x.steps })),
        }));
        savedN = sv.n || 0;
        await loadRecipes();
        S.recipeSel = p.name;
      } catch (e) { toast('存入方案库失败：' + e.message, 'err'); }
    }
    try {
      const res = await guard('套用清洗方案…', () => apiPost('/api/clean/recipe/apply', {
        sid: sheet.sid, steps: p.steps,
        mapping: useMap && hit ? hit.mapping : null,
        target: useMap && hit ? (hit.target || null) : null
      }));
      closeModal();
      await refresh();
      const sh = (S.state.sheets || []).find((x) => x.sid === sheet.sid) || sheet;
      S.cleanSteps = cloneSteps(sh.clean || []);
      S.stepOpen = -1; S.pv = null;
      delete S.previewCache[sheet.sid];
      if (saveLib && savedN) S.recipeSel = p.name;
      renderClean();
      toast(`已套用「${p.name}」：${res.n} 个动作`
        + (res.mapping ? `，${res.mapping} 个字段映射` : '')
        + (savedN ? `；${savedN} 个方案已存入方案库` : ''), 'ok');
      (res.warnings || []).slice(0, 2).forEach((w) => toast(w, 'err'));
    } catch (e) { toast(e.message, 'err'); }
  };
}

/* ================================================================== 3 映射 */
function renderMap() {
  const st = S.state;
  if (!st.sheets.length) {
    $('#main').innerHTML = subTabBar('clean', 'map') + '<div class="card"><div class="empty"><h2>还没有导入数据</h2>先去「导入数据」把文件拖进来。</div></div>';
    return;
  }
  if (!S.mapSid || !st.sheets.some((s) => s.sid === S.mapSid)) S.mapSid = st.sheets[0].sid;
  const cur = st.sheets.find((s) => s.sid === S.mapSid);

  $('#main').innerHTML = subTabBar('clean', 'map') + cleanGuide() + `
    <div class="split">
      <div class="col-list">
        <div class="card">
          <div class="card-h"><h3>原始工作表</h3><span class="sub">${st.sheets.length} 张</span></div>
          <div class="card-b" style="padding:6px">
            ${st.sheets.map((s) => {
              const n = Object.keys(s.mapping || {}).length;
              return `<div class="ritem ${s.sid === S.mapSid ? 'on' : ''}" data-sid="${s.sid}">
                <div style="flex:1;min-width:0">
                  <div style="overflow:hidden;text-overflow:ellipsis">${esc(s.sheet_name)}</div>
                  <div class="small" style="color:${s.sid === S.mapSid ? 'var(--accent)' : 'var(--muted2)'}">
                    ${esc(s.file_name)} · ${fmtInt(s.stg ? s.n_stg : s.n_rows)} 行${s.stg ? ' · 已清洗' : ''}</div>
                </div>
                ${s.target ? `<span class="tag blue">${TGT_SHORT[s.target] || '?'}</span>` : ''}
                ${n ? `<span class="tag green">${n}</span>` : ''}
              </div>`;
            }).join('')}
          </div>
        </div>
      </div>
      <div class="col-main" id="mapPanel"></div>
    </div>`;

  $('#main').querySelectorAll('[data-sid]').forEach((el) => {
    el.onclick = () => { S.mapSid = el.dataset.sid; renderMap(); };
  });
  bindCleanGuide();
  renderMapPanel(cur);
}

async function renderMapPanel(sheet) {
  const box = $('#mapPanel');
  const fields = S.state.fields;
  const target = sheet.target || '';
  const mapping = sheet.mapping || {};
  S.mapDraft = { ...mapping };
  S.auxDraft = { ...(mapping._aux || {}) };
  const sig = sheet.columns.join('\u0001');
  const sameSheets = S.state.sheets.filter((s) =>
    s.sid !== sheet.sid && s.columns.join('\u0001') === sig);

  // 预览取「清洗后的表」，所以清洗新增的列（辅助核算1/2/3、科目级次）也能参与映射
  let cols = sheet.columns.slice();
  const samples = {};
  try {
    if (!S.previewCache[sheet.sid]) {
      const p = await apiPost('/api/preview', { sid: sheet.sid, limit: 80 });
      S.previewCache[sheet.sid] = p;
    }
    const p = S.previewCache[sheet.sid];
    if (p.columns && p.columns.length) cols = p.columns;
    p.columns.forEach((c, i) => {
      const vals = [];
      for (const r of p.rows) {
        const v = r[i];
        if (v !== null && String(v).trim() !== '' && !vals.includes(String(v))) vals.push(String(v));
        if (vals.length >= 2) break;
      }
      samples[c] = vals.join(' / ');
    });
  } catch (e) { /* 预览失败不阻塞映射 */ }

  const head = `<div class="card-h">
      <h3>${esc(sheet.sheet_name)}</h3>
      <span class="sub">${esc(sheet.file_name)} · ${fmtInt(sheet.stg ? sheet.n_stg : sheet.n_rows)} 行 · ${cols.length} 列 · ${sheet.stg ? '<b>清洗后</b>' : '原始'}</span>
      <span style="flex:1"></span>
      <span class="small muted">归入</span>
      <select id="mapTarget">
        <option value="">未指定</option>
        <option value="journal" ${target === 'journal' ? 'selected' : ''}>序时账</option>
        <option value="tb" ${target === 'tb' ? 'selected' : ''}>科目余额表</option>
        <option value="flow" ${target === 'flow' ? 'selected' : ''}>银行流水</option>
      </select>
      <button class="btn sm" id="btnAuto" ${target ? '' : 'disabled'}>重新自动匹配</button>
      ${sameSheets.length ? `<button class="btn sm" id="btnApplySame" ${target ? '' : 'disabled'}>套用到 ${sameSheets.length} 张同表头表</button>` : ''}
    </div>`;

  if (!target) {
    box.innerHTML = `<div class="card">${head}<div class="card-b">
      <div class="empty" style="padding:28px">
        <h2>先选这张表属于什么</h2>
        <div class="muted">同样表头的表会被合并成一张标准表。选好目标后系统会按列名自动匹配，再人工核对。</div>
      </div></div></div>`;
    bindMapHead(sheet);
    return;
  }

  const fs = fields[target];
  // 外币金额不单独分区：它就是一列金额，标签带「(外币)」后缀，混在同一个列表里更好找
  const mainF = fs.filter((f) => f.group !== 'aux');
  const auxF = fs.filter((f) => f.group === 'aux');
  // 必备字段单独成块：以前只是在字段名后面挂一个小「必」字，夹在二十几行里根本看不见，
  // 漏了要等后面 BKD 数字对不上才发现。这里先摆必备、再摆其余，缺几个一眼就知道。
  const reqF = mainF.filter((f) => f.common);
  const optF = mainF.filter((f) => !f.common);
  const opts = (cur) => ['<option value="">— 未映射 —</option>']
    .concat(cols.map((c) => `<option value="${esc(c)}" ${c === cur ? 'selected' : ''}>${esc(c)}</option>`)).join('');
  const roles = S.state.aux_roles || [];
  const roleSel = (slot) => `<select data-slot="${slot}" style="width:100%">
      <option value="">— 未指定 —</option>
      ${roles.map((r) => `<option value="${r.key}" ${S.auxDraft[slot] === r.key ? 'selected' : ''}>${esc(r.label)}</option>`).join('')}
    </select>`;
  const count = () => Object.keys(S.mapDraft).filter((k) => !k.startsWith('_')).length;
  // 账簿名称 / 公司名称可以「不给列」：靠文件名、工作表名、手填值、对照表来定。
  // 所以判「必备是否满足」要看来源配置，不能只看下拉框选没选列。
  const bookCfg = () => (S.mapDraft._book || {});
  const compCfg = () => (S.mapDraft._company || {});
  const satisfied = (f) => {
    if (S.mapDraft[f.key]) return true;
    if (f.key === 'book_name') return !!bookCfg().mode;
    if (f.key === 'company_name') {
      const c = compCfg();
      if (c.mode === 'book') return !!bookCfg().mode;
      if (c.mode === 'col') return !!c.col;
      return (S.state.books_rows || 0) > 0;      // 默认：靠对照表匹配
    }
    return false;
  };
  const srcNote = (key) => {
    if (key === 'book_name') {
      const m = bookCfg().mode;
      const label = { file: '文件名', sheet: '工作表名', const: '手工填写', col: '指定列' }[m] || '';
      return m ? `已用「${label}」作为账簿名称` : '';
    }
    if (key === 'company_name') {
      const c = compCfg();
      if (c.mode === 'book') return '将取「账簿名称」的值';
      if (c.mode === 'col') return c.col ? `将取列「${c.col}」` : '';
      return (S.state.books_rows || 0) > 0 ? '将按账簿对照表匹配' : '';
    }
    return '';
  };
  const reqMiss = reqF.filter((f) => !satisfied(f)).length;
  const row = (f) => {
    const cur = S.mapDraft[f.key] || '';
    const slot = /^aux[123]$/.test(f.key) ? f.key : '';
    // 辅助核算1 在货币资金里就是「账户」这一维：BKD 按它区分账户，不填会把多个户并成一行
    const tip = f.key === 'aux1'
      ? '<div class="small" style="color:var(--accent);margin-top:2px">银行账户唯一标识（货币资金按它分户）</div>'
      : (f.key === 'bank_name'
        ? '<div class="small" style="color:var(--muted2);margin-top:2px">开户行/支行，排在账号上面便于识别</div>'
        : '');
    const note = (!cur && srcNote(f.key))
      ? `<div class="small" style="color:var(--ok);margin-top:2px">${esc(srcNote(f.key))}</div>` : '';
    return `<tr class="${f.common && !satisfied(f) ? 'req-miss' : ''}">
      <td><b>${esc(f.label)}</b> ${f.common ? '<span class="tag red">必</span>' : ''}
          <div class="small" style="color:var(--muted2)">${esc(f.key)}</div>${tip}</td>
      <td class="small muted">${f.type}</td>
      <td><select data-field="${f.key}" style="width:100%">${opts(cur)}</select>${note}</td>
      <td>${slot ? roleSel(slot) : ''}</td>
      <td class="sample">${esc(samples[cur] || '')}</td>
    </tr>`;
  };
  const thead = `<thead><tr><th style="width:190px">标准字段</th><th style="width:56px">类型</th>
        <th style="width:220px">对应原始列</th><th style="width:170px">辅助核算类型</th><th>样例值</th></tr></thead>`;

  box.innerHTML = `<div class="card">${head}
    <div class="card-b" style="padding-top:8px">
      <div class="row small muted" style="margin-bottom:8px;flex-wrap:wrap">
        共 ${fs.length} 个标准字段，已映射 <b id="mapCount">${count()}</b> 个。
        金额自动去千分位与全角符号，日期支持多种写法；标签带「(外币)」的是原币金额，没有外币核算就整列留空。
      </div>
      ${bookBlock(sheet, cols)}
      <div class="map-sec req">
        <div class="map-sec-h">必备字段
          <span class="small ${reqMiss ? 'req-bad' : 'muted'}">
            已映射 ${reqF.length - reqMiss}/${reqF.length}${reqMiss ? `，还差 ${reqMiss} 个` : '，齐了'}</span>
          <span class="small muted">这些列后面的 BKD / 检查都直接用，缺一个整条链路就少一块</span>
        </div>
        <table class="map-tbl">${thead}<tbody>${reqF.map(row).join('')}</tbody></table>
      </div>
      <div class="map-sec">
        <div class="map-sec-h">其他字段
          <span class="small muted">按客户导出习惯选填，不影响前四张核心表的生成</span></div>
        <table class="map-tbl">${thead}<tbody>${optF.map(row).join('')}</tbody></table>
      </div>
      ${auxF.length ? `<div class="aux-sec">
        <div class="aux-h">辅助核算
          <span class="small muted">位置槽承接原始列，类型决定派生到哪个语义列；全部可以不填，不影响合并</span></div>
        <table class="map-tbl">${thead}
          <tbody>${auxF.map(row).join('')}</tbody>
        </table>
        <div class="hint">一格多类型（如「客户:A/部门:B」）的类型选「混合」，系统按前缀自动分到 客商 / 部门 / 项目 等列。
          <b>辅助核算1 在货币资金里是银行账户的唯一标识</b>：BKD 靠它区分「同一个主体下的不同账户」，
          填「开户行&账号」最稳；只填账户名的话，同名账户会被并成一行。</div>
      </div>` : ''}
      <div class="hint">改动会自动保存。映射结果在「合并成表」页统一生成标准表。</div>
    </div></div>`;

  bindMapHead(sheet);
  const persist = async () => {
    const m = { ...S.mapDraft };
    if (Object.keys(S.auxDraft).length) m._aux = { ...S.auxDraft };
    await apiPost('/api/mapping', { sid: sheet.sid, target: sheet.target, mapping: m });
    sheet.mapping = m;
  };
  box.querySelectorAll('select[data-field]').forEach((sel) => {
    sel.onchange = async () => {
      const k = sel.dataset.field;
      const isAux = /^aux[123]$/.test(k);
      if (sel.value) S.mapDraft[k] = sel.value; else delete S.mapDraft[k];
      if (sel.value) {
        for (const kk of Object.keys(S.mapDraft)) {
          // 辅助核算位置槽允许和上面的主字段（以及其它辅助核算槽）重复选同一列：
          // 同一列既是「银行账号」又作为「辅助核算1」原样保留，是常见需求。
          const kkIsAux = /^aux[123]$/.test(kk);
          if (kk !== k && !kk.startsWith('_') && S.mapDraft[kk] === sel.value) {
            // 当前选的是辅助核算槽 → 不抢别人（允许重复）；否则只清掉其它主字段的重复，
            // 保留用户已选好的辅助核算槽映射。
            if (!isAux && !kkIsAux) delete S.mapDraft[kk];
          }
        }
      }
      const rowEl = sel.closest('tr');
      rowEl.querySelector('.sample').textContent = sel.value ? (samples[sel.value] || '') : '';
      // 给辅助核算位置槽选列时，顺手猜一次类型
      if (/^aux[123]$/.test(k) && sel.value && !S.auxDraft[k]) {
        try {
          const r = await apiPost('/api/aux/detect', { sid: sheet.sid, col: sel.value });
          const hit = (r.roles || {})[sel.value];
          if (hit && hit.role) {
            S.auxDraft[k] = hit.role;
            await persist();
            renderMapPanel(sheet);
            return;
          }
        } catch (e) { /* 猜不出来就人工选 */ }
      }
      await persist();
      $('#mapCount').textContent = count();
      if (sel.value) renderMapPanel(sheet);
    };
  });
  box.querySelectorAll('select[data-slot]').forEach((sel) => {
    sel.onchange = async () => {
      const slot = sel.dataset.slot;
      if (sel.value) S.auxDraft[slot] = sel.value; else delete S.auxDraft[slot];
      await persist();
    };
  });
  bindBook(sheet);
}


function withAux(mp, aux) {
  const o = { ...(mp || {}) };
  if (aux && Object.keys(aux).length) o._aux = { ...aux };
  return o;
}

/* ---- 映射页：账簿名称来源 ---- */
function bookBlock(sheet, cols) {
  const cfg = (S.mapDraft._book || {});
  const mode = cfg.mode || '';
  const modes = S.state.book_modes || [];
  const cm = (S.mapDraft._company || {});
  const cMode = cm.mode || 'books';          // 默认按对照表匹配
  const cModes = S.state.company_modes || [];
  const hitTxt = S.state.books_rows
    ? `账簿对照表已有 <b>${fmtInt(S.state.books_rows)}</b> 条`
    : '还没导入账簿对照表（在「导入数据」页导入）';
  let extra = '';
  if (mode === 'col') {
    extra = `<div class="row small" style="gap:6px"><span class="muted">用哪一列</span>
      <select id="bookCol"><option value="">— 选列 —</option>
      ${cols.map((c) => `<option value="${esc(c)}" ${c === (cfg.col || '') ? 'selected' : ''}>${esc(c)}</option>`).join('')}
      </select></div>`;
  } else if (mode === 'const') {
    extra = `<div class="row small" style="gap:6px"><span class="muted">账簿名称</span>
      <input id="bookVal" type="text" style="width:230px" value="${esc(cfg.value || '')}"
        placeholder="例如 某某公司2026年账"></div>`;
  } else if (mode === 'file' || mode === 'sheet') {
    const v = mode === 'file'
      ? String(sheet.file_name || '').replace(/\.[^.]+$/, '')
      : String(sheet.sheet_name || '');
    extra = `<span class="small muted">取值：<b>${esc(v)}</b></span>`;
  }
  let cExtra = '';
  if (cMode === 'col') {
    cExtra = `<div class="row small" style="gap:6px"><span class="muted">用哪一列</span>
      <select id="compCol"><option value="">— 选列 —</option>
      ${cols.map((c) => `<option value="${esc(c)}" ${c === (cm.col || '') ? 'selected' : ''}>${esc(c)}</option>`).join('')}
      </select></div>`;
  } else if (cMode === 'book') {
    cExtra = `<span class="small muted">取上面那个<b>账簿名称</b>的值，不再匹配对照表</span>`;
  } else {
    cExtra = `<span class="small muted">${hitTxt}</span>`;
  }
  return `<div class="aux-sec" style="margin-top:0;padding-top:0;border-top:0">
    <div class="aux-h">账簿与公司
      <span class="small muted">这两个是必备字段：账簿名称决定和「账簿对照表」怎么对上，公司名称决定下游底稿按谁分户</span></div>
    <div class="row small" style="gap:14px;flex-wrap:wrap;align-items:center">
      <div class="row" style="gap:6px"><span class="muted">账簿名称来源</span>
        <select id="bookMode"><option value="">— 不填 —</option>
        ${modes.map((m) => `<option value="${m.key}" ${mode === m.key ? 'selected' : ''}>${esc(m.label)}</option>`).join('')}
        </select></div>
      ${extra}
    </div>
    <div class="row small" style="gap:14px;flex-wrap:wrap;align-items:center;margin-top:6px">
      <div class="row" style="gap:6px"><span class="muted">公司名称来源</span>
        <select id="compMode">
        ${cModes.map((m) => `<option value="${m.key}" ${cMode === m.key ? 'selected' : ''}>${esc(m.label)}</option>`).join('')}
        </select></div>
      ${cExtra}
    </div>
  </div>`;
}

function bindBook(sheet) {
  const box = $('#mapPanel');
  const put = async (patch) => {
    const m = { ...S.mapDraft };
    const cur = { ...(m._book || {}) };
    Object.assign(cur, patch);
    if (!cur.mode) delete m._book; else m._book = cur;
    S.mapDraft = m;
    const out = withAux(m, S.auxDraft);
    await apiPost('/api/mapping', { sid: sheet.sid, target: sheet.target, mapping: out });
    sheet.mapping = out;
  };
  const sel = box.querySelector('#bookMode');
  if (sel) {
    sel.onchange = async () => {
      await put({ mode: sel.value });
      renderMapPanel(sheet);
    };
  }
  const bc = box.querySelector('#bookCol');
  if (bc) bc.onchange = async () => { await put({ mode: 'col', col: bc.value }); };
  const bv = box.querySelector('#bookVal');
  if (bv) {
    bv.onchange = async () => { await put({ mode: 'const', value: bv.value.trim() }); };
    bv.onkeydown = (e) => { if (e.key === 'Enter') bv.blur(); };
  }

  // 公司名称来源：默认「按账簿对照表匹配」，也可以指一列，或直接等于账簿名称
  const putCompany = async (patch) => {
    const m = { ...S.mapDraft };
    const cur = { ...(m._company || {}) };
    Object.assign(cur, patch);
    if (!cur.mode || cur.mode === 'books') delete m._company;   // 默认值不落盘，省得 mapping 里多条噪声
    else m._company = cur;
    S.mapDraft = m;
    const out = withAux(m, S.auxDraft);
    await apiPost('/api/mapping', { sid: sheet.sid, target: sheet.target, mapping: out });
    sheet.mapping = out;
  };
  const csel = box.querySelector('#compMode');
  if (csel) {
    csel.onchange = async () => {
      await putCompany({ mode: csel.value });
      renderMapPanel(sheet);
    };
  }
  const ccol = box.querySelector('#compCol');
  if (ccol) ccol.onchange = async () => { await putCompany({ mode: 'col', col: ccol.value }); };
}

function bindMapHead(sheet) {
  const sel = $('#mapTarget');
  if (sel) {
    sel.onchange = async () => {
      const t = sel.value || null;
      await apiPost('/api/mapping', { sid: sheet.sid, target: t, mapping: {} });
      sheet.target = t; sheet.mapping = {};
      if (t) {
        try {
          const r = await apiPost('/api/suggest', { sid: sheet.sid, target: t });
          const mp = withAux(r.mapping, r.aux);
          await apiPost('/api/mapping', { sid: sheet.sid, target: t, mapping: mp });
          sheet.mapping = mp;
        } catch (e) { toast(e.message, 'err'); }
      }
      await refresh(); renderNav(); renderMap();
    };
  }
  const btn = $('#btnAuto');
  if (btn) {
    btn.onclick = async () => {
      if (!sheet.target) return;
      const r = await apiPost('/api/suggest', { sid: sheet.sid, target: sheet.target });
      const mp = withAux(r.mapping, r.aux);
      await apiPost('/api/mapping', { sid: sheet.sid, target: sheet.target, mapping: mp });
      sheet.mapping = mp;
      await refresh(); renderMap();
      const nAux = r.aux ? Object.keys(r.aux).length : 0;
      toast('已按列名自动匹配 ' + Object.keys(r.mapping).length + ' 个字段' +
        (nAux ? '，识别到 ' + nAux + ' 个辅助核算类型' : ''), 'ok');
    };
  }
  const btnSame = $('#btnApplySame');
  if (btnSame) {
    btnSame.onclick = async () => {
      if (!sheet.target) return;
      const sig = sheet.columns.join('\u0001');
      const same = S.state.sheets.filter((s) =>
        s.sid !== sheet.sid && s.columns.join('\u0001') === sig);
      if (!same.length) return;
      if (!confirm(`把当前映射（含辅助核算类型）套用到另外 ${same.length} 张同表头的工作表？`)) return;
      busy('套用映射…');
      try {
        const mp = withAux(S.mapDraft, S.auxDraft);
        for (const s of same) {
          await apiPost('/api/mapping', { sid: s.sid, target: sheet.target, mapping: mp });
        }
        await refresh(); renderMap();
        toast(`已套用到 ${same.length} 张表`, 'ok');
      } catch (e) { toast(e.message, 'err'); } finally { idle(); }
    };
  }
}

/* 三张标准表：新增数据源时只改这里 + fields.py 的 TARGETS */
const TGT_LABEL = { journal: '序时账', tb: '科目余额表', flow: '银行流水',
  fa: '固定资产台账(期末)', fa_beg: '固定资产台账(期初)' };
const TGT_SHORT = { journal: '序', tb: '余', flow: '流', fa: '资', fa_beg: '资初' };
// 导入时下拉框与自动识别共用一份选项，避免两处写法不一致
const TGT_OPTIONS = [
  ['journal', '序时账'], ['tb', '科目余额表'], ['flow', '银行流水'],
  ['fa', '固定资产台账（期末）'], ['fa_beg', '固定资产台账（期初）'],
];

/* ================================================================== 3 合并 */
/* 合并被清洗门禁拦下：摊开不达标清单，给「去修」和「强制合并」两条路 */
function showMergeBlocked(t, p) {
  const box = $('#mergeResult');
  if (!box) { toast(p.verdict || '清洗未达标，已阻止合并', 'err'); return; }
  const items = (p.items || []).map((i) => `
    <li style="margin:7px 0">
      <b>${esc(i.title || '')}</b> —— ${esc(i.detail || '')}
      ${i.fix ? `<div class="muted" style="margin-top:2px">怎么修：${esc(i.fix)}</div>` : ''}
    </li>`).join('');
  box.innerHTML = `
    <div class="card" style="margin-top:12px;border-left:4px solid var(--danger)">
      <div style="font-weight:600;margin-bottom:6px">
        清洗未达标，已阻止合并（${(p.items || []).length} 项）
      </div>
      <div class="muted" style="margin-bottom:8px">${esc(p.hint || '')}</div>
      <ul style="padding-left:18px;margin:0 0 10px">${items}</ul>
      <div class="row" style="gap:8px">
        <button class="btn sm primary" id="mergeGoClean">去清洗页修</button>
        <button class="btn sm ghost" id="mergeDoForce">仍要强制合并</button>
      </div>
    </div>`;
  const b1 = $('#mergeGoClean');
  if (b1) b1.onclick = () => go('clean');
  const b2 = $('#mergeDoForce');
  if (b2) b2.onclick = async () => {
    if (!confirm('强制合并会把未达标数据带进标准表，后续数据检查会出现大面积假不符。\n确定继续？')) return;
    try {
      const r = await guard('强制合并 ' + (TGT_LABEL[t] || t) + '…',
        () => apiPost('/api/merge', { target: t, force: true }));
      toast(`已强制合并：${fmtInt(r.rows)} 行（未达标项已记入 warnings）`, 'err');
      await refresh(); renderNav(); await loadStat(t); renderMerge();
    } catch (e) { toast(e.message, 'err'); }
  };
}

function renderMerge() {
  const st = S.state;
  const targets = [
    { k: 'journal', label: '序时账', desc: '凭证分录明细：日期、凭证号、摘要、科目、借贷金额' },
    { k: 'tb', label: '科目余额表', desc: '科目级余额：期初 / 本期发生 / 期末' },
    { k: 'flow', label: '银行流水', desc: '银行账户交易明细：本方账号、对方账号、收付、余额' }
  ];
  const card = (t) => {
    const sheets = st.sheets.filter((s) => s.target === t.k);
    const sumRows = sheets.reduce((a, s) => a + (s.stg ? (s.n_stg || 0) : (s.n_rows || 0)), 0);
    const unmapped = sheets.filter((s) => !Object.keys(s.mapping || {}).filter((k) => !k.startsWith('_')).length).length;
    const cleaned = sheets.filter((s) => s.stg).length;
    const have = (st.stats[t.k] || {}).rows || 0;
    const ms = S.mergeStat[t.k];
    return `<div class="card" style="flex:1;min-width:300px">
      <div class="card-h"><h3 style="white-space:nowrap">${t.label}</h3><span class="sub">${t.desc}</span></div>
      <div class="card-b">
        <div class="kpi" style="margin-bottom:12px">
          <div><span>已映射工作表</span><b>${sheets.length}</b></div>
          <div><span>已清洗</span><b>${cleaned} / ${sheets.length}</b></div>
          <div><span>待合并行数</span><b>${fmtInt(sumRows)}</b></div>
          <div><span>现有标准表</span><b>${fmtInt(have)}</b></div>
        </div>
        ${unmapped ? `<div class="tag amber" style="margin-bottom:10px">有 ${unmapped} 张表还没做字段映射，将被跳过</div>` : ''}
        <div class="row">
          <button class="btn primary" data-merge="${t.k}" ${sheets.length ? '' : 'disabled'}>
            ${have ? '重建标准表' : '生成标准表'}</button>
          ${have ? `<button class="btn" data-peek="${t.k}">预览前 100 行</button>
                    <button class="btn" data-stat="${t.k}">统计概览</button>` : ''}
        </div>
        <div id="stat-${t.k}">${ms ? statHtml(t.k, ms) : ''}</div>
        ${t.k === 'flow' ? `<div class="row" style="margin-top:10px">
          <button class="btn sm" data-acc>重建账户清单</button>
          <button class="btn sm" data-link>重新配对</button>
          <span class="small muted" style="flex:1">清单取自科目余额表的货币资金账户；补完账号后重跑配对，不用重导流水</span>
        </div>
        <div id="flowInfo" class="mt"></div>` : ''}
      </div>
    </div>`;
  };
  $('#main').innerHTML = subTabBar('merge', 'merge') + `
    <div class="card">
      <div class="card-h"><h3>合并成标准表</h3>
        <span class="sub">把所有已映射的同目标工作表按统一字段 UNION 成一张表，重复执行会整表重建</span></div>
      <div class="card-b small muted">
        合并规则：字段按映射取值 → 类型转换（文本去首尾空格、金额转 DECIMAL(20,2)、日期转 DATE）→ 追加来源文件与原始行号，便于回溯。
      </div>
    </div>
    <div class="split" style="flex-wrap:wrap">${targets.map(card).join('')}</div>
    <div id="mergeResult"></div>`;

  $('#main').querySelectorAll('[data-merge]').forEach((b) => {
    b.onclick = async () => {
      const t = b.dataset.merge;
      try {
        const r = await guard('正在合并 ' + (TGT_LABEL[t] || t) + '…（千万行约需 30–90 秒）',
          () => apiPost('/api/merge', { target: t }));
        toast(`合并完成：${fmtInt(r.rows)} 行，用时 ${r.elapsed}s`, 'ok');
        if (r.warnings && r.warnings.length) {
          toast(r.warnings[0] + (r.warnings.length > 1 ? `（共 ${r.warnings.length} 处）` : ''), 'err');
        }
        await refresh(); renderNav();
        await loadStat(t);
        renderMerge();
        if (t === 'flow') renderFlowInfo(r.link, r.accounts || (r.link || {}).accounts);
        else renderBookResult(r.books, t);
        toast('已生成标准表 ' + fmtInt(r.rows) + ' 行', 'ok');
        if (r.books && (r.books.unmatched || []).length) {
          toast(`有 ${r.books.unmatched.length} 个账簿名称在对照表里匹配不上`, 'err');
        }
      } catch (e) {
        if (e.payload && e.payload.blocked) { showMergeBlocked(t, e.payload); }
        else { toast(e.message, 'err'); }
      }
    };
  });
  $('#main').querySelectorAll('[data-peek]').forEach((b) => {
    b.onclick = () => peekTable(b.dataset.peek);
  });
  $('#main').querySelectorAll('[data-stat]').forEach((b) => {
    b.onclick = async () => { await loadStat(b.dataset.stat); renderMerge(); };
  });
  $('#main').querySelectorAll('[data-acc]').forEach((b) => {
    b.onclick = async () => {
      try {
        const r = await guard('正在从科目余额表重建账户清单…',
          () => apiPost('/api/flow/accounts', {}));
        toast(`账户清单 ${fmtInt(r.total)} 个：可用 ${fmtInt(r.ok)}、缺账号 ${fmtInt(r.no_acct)}`,
          r.ok ? 'ok' : 'err');
        renderFlowInfo(null, r);
      } catch (e) { toast(e.message, 'err'); }
    };
  });
  $('#main').querySelectorAll('[data-link]').forEach((b) => {
    b.onclick = async () => {
      try {
        const r = await guard('正在重新配对流水与账户清单…', () => apiPost('/api/flow/link', {}));
        renderFlowInfo(r, r.accounts);
        toast(`配对完成：内部 ${fmtInt(r.internal)} / 对外 ${fmtInt(r.external)} / 待查 ${fmtInt(r.unknown)}`,
          r.unknown ? 'err' : 'ok');
      } catch (e) { toast(e.message, 'err'); }
    };
  });
}

/* ---- 银行流水：账户清单与配对结果 ---- */
function renderFlowInfo(link, accounts) {
  const box = $('#flowInfo');
  if (!box) return;
  if (!link) {
    if (!accounts) return;
    box.innerHTML = `<div class="tag ${accounts.ok ? 'green' : 'amber'}">
      账户清单 ${fmtInt(accounts.total)} 个：可用 ${fmtInt(accounts.ok)}、
      缺账号 ${fmtInt(accounts.no_acct)}、非银行 ${fmtInt(accounts.non_bank)}</div>
      ${accounts.no_acct ? `<div class="hint">缺账号的账户无法与流水配对。
        到「清洗与映射」把银行账号映射到辅助核算（aux 槽或 bank_account 列）后重建即可。</div>` : ''}`;
    return;
  }
  const un = link.unknown || 0;
  box.innerHTML = `<div class="kpi" style="margin-top:8px">
      <div><span>内部划转</span><b>${fmtInt(link.internal)}</b></div>
      <div><span>对外收付</span><b>${fmtInt(link.external)}</b></div>
      <div><span>无法判定</span><b style="color:${un ? 'var(--danger)' : 'var(--ok)'}">${fmtInt(un)}</b></div>
      <div><span>已配对账户</span><b>${fmtInt(link.acctLinked)} / ${fmtInt(link.rows)}</b></div>
      <div><span>账户清单可用</span><b>${fmtInt(link.okAccounts)} 个</b></div>
    </div>
    ${un ? `<div class="hint">「无法判定」= 本方账号没配上账户清单，或对方既无账号也无户名。
      这些行不能用于内部转账判断，要先补账号。</div>` : ''}`;
}

/* ---- 合并后：账簿对照匹配情况 ---- */
function renderBookResult(books, t) {
  const box = $('#mergeResult');
  if (!box) return;
  const lab = TGT_LABEL[t] || t;
  if (!books) {
    box.innerHTML = `<div class="card" style="margin-top:10px"><div class="card-b small muted">
      合并完成。没有导入账簿对照表，公司名称 / 简称没有回填——需要的话到「导入数据」页导入对照表后重新合并。</div></div>`;
    return;
  }
  const un = books.unmatched || [];
  box.innerHTML = `<div class="card" style="margin-top:10px">
    <div class="card-h"><h3>账簿对照匹配结果（${lab}）</h3>
      <span class="sub">按账簿名称从对照表带出公司名称 / 简称；源数据里已有公司名称的不覆盖</span></div>
    <div class="card-b">
      <div class="kpi" style="margin-bottom:${un.length ? '10px' : '0'}">
        <div><span>对照表</span><b>${fmtInt(books.books)} 条</b></div>
        <div><span>有账簿名称的行</span><b>${fmtInt(books.rows_with_book)}</b></div>
        <div><span>匹配上的行</span><b style="color:${books.matched ? 'var(--ok)' : 'var(--muted)'}">${fmtInt(books.matched)}</b></div>
        <div><span>对不上的账簿</span><b style="color:${un.length ? 'var(--danger)' : 'var(--muted)'}">${un.length}</b></div>
      </div>
      ${un.length ? `<div class="tag amber" style="margin-bottom:6px">
        下面这些账簿名称在对照表里找不到，公司名称 / 简称留空：</div>
        <div class="small" style="color:var(--danger)">${un.map(esc).join('、')}</div>
        <div class="hint">别名不一致是最常见的原因（多一个「2026年」、全角空格、简称 vs 全称）。在对照表里补一条同名记录，或改映射页的「账簿名称来源」。</div>`
        : (books.rows_with_book ? '<div class="tag green">所有账簿名称都匹配上了</div>'
          : `<div class="tag amber">${esc(books.note || '这张表没有账簿名称，未回填')}</div>`)}
    </div></div>`;
}

function statHtml(t, ms) {  if (!ms) return '';
  if (t === 'journal') {
    return `<div class="kpi" style="margin-top:12px">
      <div><span>行数</span><b>${fmtInt(ms.rows)}</b></div>
      <div><span>借方合计</span><b>${fmtCell(ms.dr)}</b></div>
      <div><span>贷方合计</span><b>${fmtCell(ms.cr)}</b></div>
      <div><span>借贷差</span><b style="color:${Math.abs((ms.dr || 0) - (ms.cr || 0)) > 0.01 ? 'var(--danger)' : 'var(--ok)'}">${fmtCell((ms.dr || 0) - (ms.cr || 0))}</b></div>
      <div><span>日期范围</span><b class="small">${esc(ms.d1 || '')} ~ ${esc(ms.d2 || '')}</b></div>
      <div><span>科目数</span><b>${fmtInt(ms.accts)}</b></div>
      <div><span>期间数</span><b>${fmtInt(ms.periods)}</b></div>
    </div>`;
  }
  if (t === 'flow') {
    return `<div class="kpi" style="margin-top:12px">
      <div><span>行数</span><b>${fmtInt(ms.rows)}</b></div>
      <div><span>收款合计</span><b>${fmtCell(ms.inflow)}</b></div>
      <div><span>付款合计</span><b>${fmtCell(ms.outflow)}</b></div>
      <div><span>涉及账户</span><b>${fmtInt(ms.accts)}</b></div>
      <div><span>内部划转</span><b>${fmtInt(ms.internal)}</b></div>
      <div><span>日期范围</span><b class="small">${esc(ms.d1 || '')} ~ ${esc(ms.d2 || '')}</b></div>
    </div>`;
  }
  return `<div class="kpi" style="margin-top:12px">
    <div><span>行数</span><b>${fmtInt(ms.rows)}</b></div>
    <div><span>本期借方</span><b>${fmtCell(ms.dr)}</b></div>
    <div><span>本期贷方</span><b>${fmtCell(ms.cr)}</b></div>
    <div><span>借贷差</span><b style="color:${Math.abs((ms.dr || 0) - (ms.cr || 0)) > 0.01 ? 'var(--danger)' : 'var(--ok)'}">${fmtCell((ms.dr || 0) - (ms.cr || 0))}</b></div>
    <div><span>科目数</span><b>${fmtInt(ms.accts)}</b></div>
  </div>`;
}

async function loadStat(t) {
  const sql = t === 'journal'
    ? `SELECT count(*) AS rows, sum(debit) AS dr, sum(credit) AS cr, count(DISTINCT acct_code) AS accts,
              count(DISTINCT period) AS periods, min(vdate)::VARCHAR AS d1, max(vdate)::VARCHAR AS d2 FROM journal`
    : t === 'flow'
      ? `SELECT count(*) AS rows,
                sum(CASE WHEN amount > 0 THEN amount ELSE 0 END) AS inflow,
                sum(CASE WHEN amount < 0 THEN -amount ELSE 0 END) AS outflow,
                count(DISTINCT acct_key) AS accts,
                count(*) FILTER (WHERE is_internal = '1') AS internal,
                min(txn_date)::VARCHAR AS d1, max(txn_date)::VARCHAR AS d2 FROM flow`
      : `SELECT count(*) AS rows, sum(cur_debit) AS dr, sum(cur_credit) AS cr, count(DISTINCT acct_code) AS accts FROM tb`;
  try {
    const r = await apiPost('/api/query', { sql, limit: 5 });
    const o = {};
    r.columns.forEach((c, i) => { o[c] = r.rows[0][i]; });
    S.mergeStat[t] = o;
  } catch (e) { /* 空表时会失败，忽略 */ }
}

async function peekTable(t) {
  try {
    const r = await apiPost('/api/query', { sql: `SELECT * FROM ${t} LIMIT 100`, limit: 100 });
    modal('预览 · ' + (TGT_LABEL[t] || t), tableHtml(r.columns, r.rows),
      '<button class="btn" data-close>关闭</button>');
    bindClose();
  } catch (e) { toast(e.message, 'err'); }
}
function bindClose() {
  $('#modalFoot').querySelectorAll('[data-close]').forEach((b) => { b.onclick = closeModal; });
}

/* ================================================================== 5 数据检查 */
// 四组检查：一 余额表自平衡 / 二 序时账自平衡 / 三 公司+一级科目 / 四 公司+科目全称
// 每组可能有多张明细表（part），后端把每张都落成 _recon_* 表，导出 Excel 时读表。
const RECON_ROW_LIMIT = 500;

const TB_BASES = [
  ['ytd', '本年累计(本币)'],
  ['cur', '本期发生(本币)'],
  ['open', '期初(本币)'],
  ['end', '期末(本币)']
];

// 数据检查 · 按列值删行：两张标准表可选删除的列
/* 「按列值删除行」的可选列：正常从 /api/delete-rows/modes 取（中文名来自 core/fields.py
   的标准字段定义）。下面这份只是接口不可用时的兜底，别再维护第二套中文名。 */
const DEL_COLS = {
  journal: ['vdate', 'period', 'vtype', 'vno', 'voucher_key', 'summary', 'acct_code',
            'acct_name', 'acct_full', 'acct_l1', 'debit', 'credit', 'currency',
            'book_name', 'company_name', 'dept', 'counterparty', 'person', 'aux1', 'aux2'],
  tb: ['company_name', 'book_name', 'acct_code', 'acct_name', 'acct_full', 'acct_l1',
       'level', 'currency', 'end_bal', 'end_debit', 'end_credit', 'bank_account',
       'counterparty_name', 'dept_name', 'project_name', 'person_name', 'item_name',
       'aux1', 'aux2']
};

function isOkStatus(v) { return v === '一致' || v === '平衡'; }

/** 渲染一张检查明细表。表头用后端给的中文（part.headers），数字右对齐。 */
function partTable(part, onlyDiff) {
  const allCols = part.columns || [];
  const allHeads = (part.headers && part.headers.length === allCols.length) ? part.headers : allCols;
  const allRows = part.rows || [];
  const view = part.view_cols || null;
  // 整列全空的原币列（没外币核算的账套常见）不显示，表太宽反而看不清
  const keep = allCols.map((c, i) => {
    if (view && view.indexOf(c) < 0) return false;
    if (!/_fx$/.test(c)) return true;
    return allRows.some((r) => r[i] !== null && r[i] !== undefined && Number(r[i]) !== 0);
  });
  const cols = allCols.filter((_, i) => keep[i]);
  const heads = allHeads.filter((_, i) => keep[i]);
  const si = part.status_col ? allCols.indexOf(part.status_col) : -1;
  const diffIdx = new Set((part.diff_cols || []).map((k) => allCols.indexOf(k)).filter((i) => i >= 0));
  const tol = Number((S.recon || {}).tol) || 0.01;
  let rows = allRows.map((r) => r.filter((_, i) => keep[i]));
  const si2 = si >= 0 ? cols.indexOf(allCols[si]) : -1;
  // 这一组本来就没问题时不过滤，否则「全部平衡」会显示成一张空表
  if (onlyDiff && si2 >= 0 && (part.bad || 0) > 0) rows = rows.filter((r) => !isOkStatus(r[si2]));
  const total = rows.length;
  const shown = rows.slice(0, RECON_ROW_LIMIT);

  const head = heads.map((h, i) => {
    const first = shown[0] || rows[0] || [];
    const num = typeof first[i] === 'number' && cols[i] !== 'vno';
    return `<th${num ? ' class="num"' : ''}>${esc(h)}</th>`;
  }).join('');

  const body = shown.map((row) => '<tr>' + cols.map((c, i) => {
    const v = row[i];
    if (i === si2) {
      const ok = isOkStatus(v);
      return `<td><span class="tag ${ok ? 'green' : 'red'}">${esc(v)}</span></td>`;
    }
    if (typeof v === 'number') {
      const bad = diffIdx.has(allCols.indexOf(c)) && Math.abs(v) > tol;
      return `<td class="num"${bad ? ' style="color:var(--danger);font-weight:600"' : ''}>${fmtCell(v)}</td>`;
    }
    if (v === null || v === undefined || v === '') return '<td class="muted">—</td>';
    return `<td>${esc(v)}</td>`;
  }).join('') + '</tr>').join('');

  const empty = onlyDiff
    ? '<tr><td colspan="' + cols.length + '" class="empty">这一项全部平衡，没有需要关注的</td></tr>'
    : '<tr><td colspan="' + cols.length + '" class="empty">没有数据</td></tr>';

  const more = total > shown.length
    ? `<div class="small muted" style="padding:6px 0">共 ${fmtInt(total)} 行，这里只显示前 ${RECON_ROW_LIMIT} 行；完整数据请点「导出 Excel」</div>`
    : '';

  return `<div class="tbl-wrap" style="max-height:520px"><table class="grid">
    <thead><tr>${head}</tr></thead><tbody>${body || empty}</tbody></table></div>${more}`;
}

/** 对账组（三、四）的按主体汇总：一眼看出是哪家公司对不上 */
function entSummary(part) {
  const ents = part.entities || [];
  if (ents.length <= 1) return '';
  return `<div class="aux-h">按公司汇总
      <span class="small muted">对不上时先看这张，能立刻定位到是哪家公司</span></div>
    <table class="grid" style="width:100%;margin-bottom:14px">
      <thead><tr><th>主体（公司/账簿）</th><th class="num">科目数</th><th class="num">一致</th>
        <th class="num">不符</th><th class="num">单边</th><th class="num">净额差异</th></tr></thead>
      <tbody>${ents.map((e) => `<tr>
        <td><b>${esc(e.entity)}</b></td>
        <td class="num">${fmtInt(e.n)}</td>
        <td class="num">${fmtInt(e.ok)}</td>
        <td class="num"${e.bad ? ' style="color:var(--danger);font-weight:600"' : ''}>${fmtInt(e.bad)}</td>
        <td class="num"${(e.only_tb + e.only_jr) ? ' style="color:var(--warn);font-weight:600"' : ''}>${fmtInt(e.only_tb + e.only_jr)}</td>
        <td class="num"${Math.abs(e.diff_net || 0) > 0.01 ? ' style="color:var(--danger);font-weight:600"' : ''}>${fmtCell(e.diff_net)}</td>
      </tr>`).join('')}</tbody>
    </table>`;
}

/** 一组检查的卡片 */
function checkCard(chk) {
  const err = chk.error;
  const parts = chk.parts || [];
  const bad = chk.bad || 0;
  const tag = err ? '<span class="tag red">无法检查</span>'
    : (chk.ok ? '<span class="tag green">通过</span>'
      : `<span class="tag red">${fmtInt(bad)} 项待查</span>`);
  const notes = (chk.notes || []).map((n) => `<span class="chip">${esc(n)}</span>`).join('');
  const warns = (chk.warnings || []).map((w) => `<div class="tag amber" style="margin-bottom:6px">${esc(w)}</div>`).join('');
  let body = '';
  if (err) {
    body = `<div class="tag red" style="margin-bottom:8px">${esc(err)}</div>
      <div class="small muted">解决办法：${chk.key === 'tb_self' || chk.key === 'jr_self'
        ? '先到第 3 步「合并与检查 → 合并成表」生成标准表，再到第 2 步「清洗与映射 → 字段映射」把借贷金额、公司名称、会计期间等列映射上。'
        : '先在第 2 步「清洗与映射 → 数据清洗」用「提取一级科目」或「生成一级科目」拆出科目列，再到同一步的「字段映射」把它映射到「一级科目」/「科目全称」，两张表都要。'}</div>`;
  } else {
    body = `${notes ? `<div class="row" style="gap:6px;margin-bottom:8px;flex-wrap:wrap">${notes}</div>` : ''}
      ${warns}
      ${parts.map((p) => `${parts.length > 1 ? `<div class="aux-h">${esc(p.title)}${p.note ? `<span class="small muted">${esc(p.note)}</span>` : ''}</div>` : ''}
        ${entSummary(p)}
        ${partTable(p, S.reconDiffOnly)}`).join('')}`;
  }
  return `<div class="card" style="margin-top:12px" id="chk-${esc(chk.key)}">
    <div class="card-h">
      <h3>${esc(chk.no ? chk.no + '、' : '')}${esc(chk.title || '')}</h3>
      <span class="sub">${esc(chk.desc || '')}</span>
      <span style="flex:1"></span>${tag}
    </div>
    <div class="card-b">${body}</div>
  </div>`;
}

function renderCheck() {
  const st = S.state;
  const hasTb = ((st.stats.tb || {}).rows || 0) > 0;
  const hasJr = ((st.stats.journal || {}).rows || 0) > 0;
  const r = S.recon;
  const tbN = (st.stats.tb || {}).rows || 0;
  const jrN = (st.stats.journal || {}).rows || 0;
  const pr = st.period_range || null;
  const pFromVal = (r && r.period && r.period.from) || (pr ? pr.min : '');
  const pToVal = (r && r.period && r.period.to) || (pr ? pr.max : '');

  $('#main').innerHTML = subTabBar('merge', 'check') + `
    <div class="card">
      <div class="card-h">
        <h3>数据检查 · 四组平衡</h3>
        <span class="sub">余额表自平衡 · 序时账自平衡 · 两表一级科目对账 · 两表科目全称对账</span>
      </div>
      <div class="card-b">
        <div class="kpi" style="margin-bottom:12px">
          <div><span>科目余额表</span><b>${fmtInt(tbN)} 行</b></div>
          <div><span>序时账</span><b>${fmtInt(jrN)} 行</b></div>
          <div><span>账簿对照</span><b>${fmtInt(st.books_rows || 0)} 条</b></div>
        </div>
        ${hasTb && hasJr ? '' : `<div class="tag amber" style="margin-bottom:10px">
          还没生成${hasTb ? '序时账' : (hasJr ? '科目余额表' : '标准表')}，先到第 3 步「合并与检查 → 合并成表」生成（部分检查可以先跑）</div>`}
        <div class="row" style="gap:14px;flex-wrap:wrap;align-items:flex-end">
          <div class="row small" style="gap:6px">
            <span class="muted">期间从</span>
            <input id="pFrom" type="text" placeholder="全部" style="width:110px"
              value="${esc(pFromVal)}">
            <span class="muted">到</span>
            <input id="pTo" type="text" placeholder="全部" style="width:110px"
              value="${esc(pToVal)}">
          </div>
          <div class="row small" style="gap:6px">
            <span class="muted">余额表口径</span>
            <select id="tbBase" style="width:128px">
              ${TB_BASES.map((b) => `<option value="${b[0]}" ${(r && r.tb_base === b[0]) || (!r && b[0] === 'ytd') ? 'selected' : ''}>${b[1]}</option>`).join('')}
            </select>
          </div>
          <div class="row small" style="gap:6px">
            <span class="muted">容差</span>
            <input id="rTol" type="number" step="0.01" style="width:84px"
              value="${r ? (r.tol ?? 0.01) : 0.01}">
          </div>
          <button class="btn primary" id="btnRecon">开始检查</button>
          ${r && r.summary ? `<button class="btn" id="btnReconX">导出 Excel</button>` : ''}
          <button class="btn" id="btnDiag">数据体检</button>
        </div>
        <div class="hint">
          <b>一、科目余额表</b>：按公司名称分组，查本期/本年累计借贷发生额是否相等、期初+本期借−本期贷是否等于期末。<br>
          <b>二、序时账</b>：按公司名称分组，查总额、每月、每张凭证的借贷是否相等。凭证唯一性靠「唯一凭证号」<br>
          &nbsp;&nbsp;&nbsp;&nbsp;（清洗里的「生成唯一凭证号」＝公司名称-期间-凭证号）；没映射就现场拼一个，结果一样。
          <b>三、四</b>：两表按「公司名称 + 一级科目」和「公司名称 + 科目全称」逐项比发生额。<br>
          <b>期间</b>只影响第三、四项对账（第一、二项是自平衡，永远看全量）。原始数据变动后请重新合并并重跑检查。
        </div>
      </div>
    </div>
    <div class="card" style="margin-top:12px">
      <div class="card-h">
        <h3>按列值删除行</h3>
        <span class="sub">对已合并的标准表直接删行；多条条件<b>全部满足</b>才删（且）</span>
      </div>
      <div class="card-b">
        <div class="row" style="gap:10px;flex-wrap:wrap;align-items:flex-end">
          <div class="col" style="min-width:170px">
            <label class="small muted">表</label>
            <select id="delTable">
              <option value="journal" ${S.delTable === 'tb' ? '' : 'selected'}>序时账 journal</option>
              <option value="tb" ${S.delTable === 'tb' ? 'selected' : ''}>科目余额表 tb</option>
            </select>
          </div>
          <button class="btn sm" id="btnDelAdd">＋ 添加条件</button>
          <span style="flex:1"></span>
          <button class="btn danger" id="btnDelRows" ${hasTb || hasJr ? '' : 'disabled'}>删除匹配行</button>
        </div>
        <div id="delConds">${delCondRows()}</div>
        <div class="hint" style="margin-top:8px">
          「等于」忽略首尾空格和大小写；「包含」是子串匹配；「不等于」= 留下等于该值的行；
          「为空 / 不为空」不用填值。
          <b>多条条件之间是「且」</b>，例如「公司名称 等于 甲公司」+「科目名称 等于 银行存款」，只删同时满足的行。
          删除<b>直接生效、不可撤销</b>，执行前会把整张表快照到 work 目录，误删可找回来。
        </div>
      </div>
    </div>
    <div id="reconBox">${r ? (r._restored ? `<div class="tag amber" style="margin-bottom:10px">这是上次检查的结果${r.ran_at ? '（' + esc(r.ran_at) + '）' : ''}，数据可能已变动，建议重新点「开始检查」</div>` : '') + reconHtml(r) : ''}</div>
    <div id="diagBox"></div>`;

  const btn = $('#btnRecon');
  if (btn) btn.onclick = doRecon;
  const btnDiag = $('#btnDiag');
  if (btnDiag) btnDiag.onclick = doDiagnose;
  const btnX = $('#btnReconX');
  if (btnX) {
    btnX.onclick = async () => {
      try {
        const stamp = new Date().toISOString().slice(0, 10).replace(/-/g, '');
        const out = await guard('正在导出 Excel…',
          () => apiPost('/api/recon/export', { name: '数据检查_' + stamp }));
        toast(`已导出 ${out.file}（${out.sheets} 张工作表）`, 'ok');
        window.location.href = out.url;
      } catch (e) { toast(e.message, 'err'); }
    };
  }
  const ck = $('#rDiffOnly');
  if (ck) ck.onchange = () => { S.reconDiffOnly = ck.checked; renderCheck(); };

  // 按列值删行：条件行是动态的，统一交给 delWire() 绑
  const delTable = $('#delTable');
  if (delTable) {
    delTable.onchange = () => {
      S.delTable = delTable.value;
      S.delConds = [{ col: '', mode: '等于', value: '' }];   // 换表后列不一样，条件清空
      const box = $('#delConds');
      if (box) { box.innerHTML = delCondRows(); delWire(); }
    };
  }
  const btnDelAdd = $('#btnDelAdd');
  if (btnDelAdd) btnDelAdd.onclick = () => {
    S.delConds = delCondsState();
    S.delConds.push({ col: '', mode: '等于', value: '' });
    const box = $('#delConds');
    if (box) { box.innerHTML = delCondRows(); delWire(); }
  };
  const btnDel = $('#btnDelRows');
  if (btnDel) btnDel.onclick = doDeleteRows;
  delWire();
  if (!S.delMeta && !S.delMetaLoading) loadDelMeta();
}

/* ---- 按列值删行：条件行 ---- */
function delMetaCols(table) {
  const m = S.delMeta && S.delMeta.tables && S.delMeta.tables[table];
  if (m && m.length) return m;
  // 接口没回来时先用内部列名兜底，别让下拉空着
  return (DEL_COLS[table] || []).map((k) => ({ key: k, label: k }));
}

function delColLabel(table, key) {
  const hit = delMetaCols(table).find((c) => c.key === key);
  return hit ? hit.label : key;
}

function delCondsState() {
  if (!S.delConds || !S.delConds.length) {
    S.delConds = [{ col: '', mode: '等于', value: '' }];
  }
  return S.delConds;
}

/** 这两个模式不需要填值 */
function delNeedNoValue(mode) { return mode === '为空' || mode === '不为空'; }

function delCondRows() {
  const table = S.delTable || 'journal';
  const cols = delMetaCols(table);
  const modes = (S.delMeta && S.delMeta.modes) || ['等于', '包含', '不等于', '为空', '不为空'];
  const list = delCondsState();
  return `<div class="del-conds">${list.map((c, i) => `
    <div class="del-cond">
      <span class="del-join">${i === 0 ? '条件' : '且'}</span>
      <select data-dc="col" data-i="${i}" style="min-width:200px">
        <option value="">— 选择列 —</option>
        ${cols.map((x) => `<option value="${x.key}" ${c.col === x.key ? 'selected' : ''}>${esc(x.label)}</option>`).join('')}
      </select>
      <select data-dc="mode" data-i="${i}" style="width:100px">
        ${modes.map((m) => `<option value="${m}" ${c.mode === m ? 'selected' : ''}>${m}</option>`).join('')}
      </select>
      <input data-dc="value" data-i="${i}" type="text" style="flex:1;min-width:180px"
        placeholder="${delNeedNoValue(c.mode) ? '（此条件不用填值）' : '要匹配的值'}"
        value="${esc(c.value || '')}" ${delNeedNoValue(c.mode) ? 'disabled' : ''}>
      <button class="btn sm ghost" data-dc="rm" data-i="${i}" title="删掉这条条件"
        ${list.length <= 1 ? 'disabled' : ''}>✕</button>
    </div>`).join('')}</div>`;
}

function delWire() {
  const box = $('#delConds');
  if (!box) return;
  const rerender = () => { box.innerHTML = delCondRows(); delWire(); };
  box.querySelectorAll('[data-dc]').forEach((el) => {
    const i = Number(el.dataset.i);
    const f = el.dataset.dc;
    if (f === 'rm') {
      el.onclick = () => {
        S.delConds = delCondsState();
        if (S.delConds.length <= 1) return;
        S.delConds.splice(i, 1);
        rerender();
      };
    } else if (f === 'mode') {
      el.onchange = () => {
        S.delConds = delCondsState();
        S.delConds[i].mode = el.value;
        rerender();                       // 换模式要切换「值」输入框是否可用
      };
    } else if (f === 'col') {
      el.onchange = () => { S.delConds = delCondsState(); S.delConds[i].col = el.value; };
    } else {
      el.oninput = () => { S.delConds = delCondsState(); S.delConds[i].value = el.value; };
    }
  });
}

async function loadDelMeta() {
  S.delMetaLoading = true;
  try {
    S.delMeta = await apiGet('/api/delete-rows/modes');
  } catch (e) {
    S.delMeta = { modes: ['等于', '包含', '不等于', '为空', '不为空'], tables: {} };
  }
  S.delMetaLoading = false;
  const box = $('#delConds');
  if (box) { box.innerHTML = delCondRows(); delWire(); }
}

async function doDeleteRows() {
  const table = ($('#delTable') ? $('#delTable').value : '') || 'journal';
  const conds = (S.delConds || []).filter((c) => c && c.col);
  if (!conds.length) { toast('请先选一条判断列', 'err'); return; }
  const bad = conds.find((c) => !delNeedNoValue(c.mode) && !String(c.value || '').trim());
  if (bad) { toast(`「${delColLabel(table, bad.col)}」这一条没填值`, 'err'); return; }

  // 先试算：说清楚会删多少行、删掉的是哪几行，再让人确认
  let pre;
  try {
    pre = await apiPost('/api/delete-rows', { table, conditions: conds, dry_run: true });
  } catch (e) { toast(e.message, 'err'); return; }
  if (!pre.matched) { toast('没有匹配的行，未做任何改动', 'warn'); return; }

  const txt = conds.map((c, i) => `${i ? '且 ' : ''}${delColLabel(table, c.col)} ${c.mode}`
    + (delNeedNoValue(c.mode) ? '' : `「${c.value}」`)).join('\n');
  const sample = (pre.sample || []).map(
    (r) => '  ' + r.map((x) => `${x.label}=${x.value}`).join('，')).join('\n');
  const msg = `将从「${TGT_LABEL[table] || table}」删除 ${fmtInt(pre.matched)} 行`
    + `（当前 ${fmtInt(pre.before)} 行）。\n\n${txt}`
    + (sample ? `\n\n将被删的前几行：\n${sample}` : '')
    + '\n\n删除直接生效、不可撤销（会先把整张表备份到 work 目录）。确定吗？';
  if (!confirm(msg)) return;

  try {
    const r = await guard('正在删除匹配行…',
      () => apiPost('/api/delete-rows', { table, conditions: conds }));
    toast(`已删除 ${fmtInt(r.deleted)} 行（${fmtInt(r.before)} → ${fmtInt(r.after)}）`, 'ok');
    if (r.wiped) toast('注意：这张表已被删空，请确认条件是不是写宽了', 'err');
    await refresh();
    renderCheck();
  } catch (e) { toast(e.message, 'err'); }
}

async function doRecon() {
  const body = {
    period_from: ($('#pFrom').value || '').trim() || null,
    period_to: ($('#pTo').value || '').trim() || null,
    tol: Number($('#rTol').value) || 0.01,
    tb_base: ($('#tbBase').value) || 'ytd'
  };
  try {
    const r = await guard('正在跑四组平衡检查…', () => apiPost('/api/recon', body));
    S.recon = r;
    renderCheck();
    const sm = r.summary || {};
    const items = sm.items || [];
    const bad = items.filter((x) => !x.ok);
    if (bad.length) {
      toast(`检查完成：${fmtInt(sm.passed)} 项通过，${fmtInt(sm.failed)} 项有问题 —— ${bad.map((x) => x.title.split('·')[0].trim()).join('、')}`,
        'err');
    } else {
      toast(`检查完成：${fmtInt(sm.total)} 项全部通过`, 'ok');
    }
  } catch (e) { toast(e.message, 'err'); }
}

async function doDiagnose() {
  try {
    const d = await guard('正在做数据体检…', () => apiPost('/api/diagnose', {}));
    const box = $('#diagBox');
    if (!box) return;
    const fs = d.findings || [];
    const lv = { ok: 'green', info: '', warn: 'amber', error: 'red' };
    box.innerHTML = `<div class="card" style="margin-top:12px">
      <div class="card-h"><h3>数据体检</h3>
        <span class="sub">自动检查重复层 / 汇总层 / 小计行 / 页脚 / 分卷 / 期间错位</span></div>
      <div class="card-b">${fs.map((f) => `<div class="row" style="gap:8px;align-items:flex-start;margin-bottom:10px">
        <span class="tag ${lv[f.level] || ''}" style="flex:none">${esc(f.title)}</span>
        <div style="flex:1"><div class="small">${esc(f.detail)}</div>
        ${f.suggest ? `<div class="small muted">→ ${f.suggest.map(esc).join('；')}</div>` : ''}</div>
      </div>`).join('')}</div></div>`;
  } catch (e) { toast(e.message, 'err'); }
}

function reconHtml(r) {
  if (!r.summary) return '';
  const sm = r.summary;
  const items = sm.items || [];
  const rowHtml = items.map((it) => {
    const ok = it.ok;
    return `<tr style="cursor:pointer" onclick="document.getElementById('chk-${esc(it.key)}').scrollIntoView({behavior:'smooth'})">
      <td><b>${esc(it.no || '')}、${esc(it.title || '')}</b></td>
      <td class="num">${it.error ? '—' : fmtInt(it.total)}</td>
      <td class="num"${it.bad ? ' style="color:var(--danger);font-weight:600"' : ''}>${it.error ? '—' : fmtInt(it.bad)}</td>
      <td>${it.error ? '<span class="tag red">无法检查</span>'
        : (ok ? '<span class="tag green">通过</span>' : '<span class="tag red">有问题</span>')}</td>
    </tr>`;
  }).join('');

  return `
  <div class="card" style="margin-top:12px">
    <div class="card-h">
      <h3>检查结论</h3>
      <span class="sub">容差 ${r.tol} · 余额表口径「${esc(r.tb_label || '')}」· 期间 ${esc((r.period || {}).from || '全部')} ~ ${esc((r.period || {}).to || '全部')}</span>
      <span style="flex:1"></span>
      <label class="small muted" style="display:flex;align-items:center;gap:5px">
        <input type="checkbox" id="rDiffOnly" ${S.reconDiffOnly ? 'checked' : ''}>只看有问题的</label>
    </div>
    <div class="card-b">
      <div class="kpi" style="margin-bottom:12px">
        <div><span>检查项</span><b>${fmtInt(sm.total)}</b></div>
        <div><span>通过</span><b style="color:var(--ok)">${fmtInt(sm.passed)}</b></div>
        <div><span>有问题</span><b style="color:${sm.failed ? 'var(--danger)' : 'var(--muted)'}">${fmtInt(sm.failed)}</b></div>
        <div><span>问题合计</span><b style="color:${sm.bad_items ? 'var(--danger)' : 'var(--muted)'}">${fmtInt(sm.bad_items)}</b></div>
      </div>
      <div class="tag ${sm.failed ? 'red' : 'green'}" style="margin-bottom:10px">
        ${sm.failed
          ? `有 ${fmtInt(sm.failed)} 组检查发现问题，共 ${fmtInt(sm.bad_items)} 项。点下面任意一行可直接跳到该组明细。`
          : '四组平衡全部通过：余额表与序时账内部平衡，两表按一级科目和科目全称都核对一致。'}
      </div>
      <table class="grid" style="width:100%">
        <thead><tr><th>检查项</th><th class="num">检查数量</th><th class="num">问题数量</th><th>结论</th></tr></thead>
        <tbody>${rowHtml}</tbody>
      </table>
    </div>
  </div>
  ${(r.checks || []).map(checkCard).join('')}`;
}

/* ================================================================== 6 查账 */
function renderQuery() {
  const cats = [];
  S.rules.forEach((r, i) => {
    let g = cats.find((c) => c.cat === r.cat);
    if (!g) { g = { cat: r.cat, items: [] }; cats.push(g); }
    g.items.push({ r, i });
  });
  const cur = currentRule();

  // 规则下拉框选项（含「自由 SQL」）
  const ruleOpts = ['<option value="__free">— 自由 SQL —</option>']
    .concat(S.rules.map((r) => `<option value="${esc(r.id)}">${esc(r.cat)} / ${esc(r.name)}</option>`))
    .join('');

  $('#main').innerHTML = `
    <div class="card">
      <div class="card-h">
        <h3 id="ruleName">${cur ? esc(cur.name) : '自由 SQL'}</h3>
        <span class="sub" id="ruleCat">${cur ? esc(cur.cat) : ''}</span>
        <span style="flex:1"></span>
        <span class="seg" style="display:flex;gap:0;border:1px solid var(--line);border-radius:6px;overflow:hidden">
          <button class="btn sm ghost" data-mode="form" style="border:none;border-radius:0">表单</button>
          <button class="btn sm ghost" data-mode="sql" style="border:none;border-radius:0">SQL</button>
        </span>
        <button class="btn sm ghost" id="btnImportRule" title="从 xlsx 导入规则">导入规则</button>
        <button class="btn sm ghost" id="btnExportRule" title="导出全部规则为 xlsx">导出规则</button>
      </div>
      <div class="card-b">
        <div class="row small" style="gap:8px;align-items:center;margin-bottom:8px">
          <span class="muted">选择规则</span>
          <select id="ruleSel" style="flex:1;max-width:340px">${ruleOpts}</select>
          <span style="flex:1"></span>
          <button class="btn sm ghost" id="btnToggleSchema">${S.schema ? '收起' : '展开'}表结构</button>
        </div>
        <div class="small muted" id="ruleDesc">${cur ? esc(cur.desc || '') : '直接写 SQL，对 journal / tb 两张标准表查询，结果可导出。'}</div>
        <div class="row" id="paramBox" style="margin-top:10px"></div>
        <div id="schemaBox" style="margin-top:10px" ${S.schema ? '' : 'hidden'}>${schemaHtml()}</div>
        <div id="sqlBox" style="margin-top:10px" hidden>
          <textarea id="sqlText" spellcheck="false"></textarea>
          <div class="hint">可用 <code>{{参数名}}</code> 占位；表：<code>journal</code>（序时账）、<code>tb</code>（科目余额表）${S.schema && S.schema._recon ? '、<code>_recon</code>（核对结果）' : ''}。改完直接运行，不会影响原规则。</div>
        </div>
        <div class="row" style="margin-top:12px">
          <button class="btn primary" id="btnRun">运行</button>
          <button class="btn" id="btnCount">统计总行数</button>
          <span class="sep"></span>
          <button class="btn" id="btnCsv">导出 CSV</button>
          <button class="btn" id="btnXlsx">导出 Excel</button>
          <span class="muted small">导出时表头自动换成中文</span>
          <span class="sep"></span>
          <button class="btn ghost" id="btnSaveRule">保存为规则</button>
          <button class="btn ghost" id="btnReset">还原</button>
          <span style="flex:1"></span>
          <button class="btn danger ghost" id="btnInterrupt" ${S.running ? '' : 'disabled'}>中断查询</button>
        </div>
      </div>
    </div>
    <div class="card" id="resultCard">
      <div class="card-h"><h3>查询结果</h3><span class="sub" id="resSub">尚未运行</span></div>
      <div id="resultBody"><div class="empty">点「运行」查看结果</div></div>
    </div>`;

  // 规则下拉
  const sel = $('#ruleSel');
  if (sel) {
    sel.value = S.ruleId;
    sel.onchange = () => {
      S.ruleId = sel.value; S.mode = sel.value === '__free' ? 'sql' : 'form'; S.result = null;
      renderQuery();
    };
  }
  // 表结构折叠
  const tg = $('#btnToggleSchema');
  if (tg) tg.onclick = () => {
    if (S.schema) { S.schema = null; } else { loadSchema(); }
    renderQuery();
  };
  // 规则导入/导出
  const bi = $('#btnImportRule');
  if (bi) bi.onclick = () => { openImportRules(); };
  const be = $('#btnExportRule');
  if (be) be.onclick = async () => {
    try {
      const r = await guard('导出规则…', () => apiPost('/api/rules/export'));
      toast('已导出 ' + r.count + ' 条规则', 'ok');
      window.location.href = r.url;
    } catch (e) { toast(e.message, 'err'); }
  };

  $('#main').querySelectorAll('[data-mode]').forEach((b) => {
    b.onclick = () => { S.mode = b.dataset.mode; renderQuery(); };
  });
  syncModeUi();
  wireQueryActions();
}

async function loadSchema() {
  try {
    S.schema = await apiGet('/api/schema');
  } catch (e) { S.schema = null; }
}

function schemaHtml() {
  if (!S.schema) return '<div class="empty" style="padding:10px">加载中…</div>';
  const tables = Object.keys(S.schema);
  if (!tables.length) return '<div class="empty" style="padding:10px">还没有合并标准表</div>';
  return `<div class="schema-grid">
    ${tables.map((t) => {
      const info = S.schema[t];
      return `<div class="schema-card">
        <div class="schema-t">${esc(t)} <span class="muted small">${esc(info.label)} · ${fmtInt(info.rows)} 行</span></div>
        <div class="schema-cols">${info.columns.map((c) =>
          `<span class="schema-f" title="${esc(c.key)}"><b>${esc(c.label)}</b><i>${esc(c.key)}</i></span>`).join('')}</div>
      </div>`;
    }).join('')}
  </div>`;
}

function currentRule() {
  if (S.ruleId === '__free') return { id: '__free', name: '自由 SQL', cat: '自定义', desc: '', params: [], sql: '' };
  return S.rules.find((r) => r.id === S.ruleId) || null;
}

function syncModeUi() {
  const rule = currentRule();
  const box = $('#paramBox');
  const sqlBox = $('#sqlBox');
  const ta = $('#sqlText');
  const formBtn = $('#main').querySelector('[data-mode=form]');
  const sqlBtn = $('#main').querySelector('[data-mode=sql]');
  formBtn.classList.toggle('primary', S.mode === 'form');
  sqlBtn.classList.toggle('primary', S.mode === 'sql');
  formBtn.disabled = S.ruleId === '__free';

  if (!rule) return;
  const ps = rule.params || [];
  if (!S.paramDraft[rule.id]) {
    S.paramDraft[rule.id] = {};
    ps.forEach((p) => { S.paramDraft[rule.id][p.key] = p.default; });
  }
  const vals = S.paramDraft[rule.id];
  if (S.mode === 'form') {
    box.hidden = false; sqlBox.hidden = true;
    box.innerHTML = ps.length ? ps.map((p) => `
      <label class="row small" style="gap:6px">
        <span class="muted">${esc(p.label)}</span>
        <input type="${p.type === 'number' ? 'number' : 'text'}" data-param="${esc(p.key)}"
               value="${esc(vals[p.key] !== undefined ? vals[p.key] : '')}" style="width:${p.type === 'number' ? 110 : 200}px">
      </label>`).join('') : '<span class="small muted">此规则无参数</span>';
    box.querySelectorAll('[data-param]').forEach((inp) => {
      inp.oninput = () => { vals[inp.dataset.param] = inp.value; };
    });
  } else {
    box.hidden = true; sqlBox.hidden = false;
    ta.value = S.sqlDraft[rule.id] !== undefined ? S.sqlDraft[rule.id] : (rule.sql || '');
    ta.oninput = () => { S.sqlDraft[rule.id] = ta.value; };
    ta.style.height = Math.min(360, Math.max(140, ta.value.split('\n').length * 19 + 20)) + 'px';
  }
}

function currentSql() {
  const rule = currentRule();
  if (S.mode === 'sql') {
    return S.sqlDraft[rule.id] !== undefined ? S.sqlDraft[rule.id] : (rule.sql || '');
  }
  return rule.sql || '';
}

function wireQueryActions() {
  const rule = currentRule();
  const params = () => (S.mode === 'form' ? (S.paramDraft[rule.id] || {}) : (S.paramDraft[rule.id] || {}));

  const setRunning = (on) => {
    S.running = on;
    const b = $('#btnInterrupt');
    if (b) b.disabled = !on;
  };

  $('#btnRun').onclick = async () => {
    const sql = currentSql();
    if (!sql.trim()) { toast('SQL 为空', 'err'); return; }
    setRunning(true);
    try {
      const r = await guard('查询中…（可点「中断查询」）', () => apiPost('/api/query', { sql, params: params(), limit: 5000 }));
      S.result = r; S.page = 0;
      renderResult();
    } catch (e) { toast(e.message, 'err'); }
    finally { setRunning(false); }
  };
  $('#btnCount').onclick = async () => {
    const sql = currentSql();
    setRunning(true);
    try {
      const r = await guard('统计中…（可点「中断查询」）', () => apiPost('/api/count', { sql, params: params() }));
      toast('结果共 ' + fmtInt(r.rows) + ' 行（' + r.elapsed_ms + 'ms）');
    } catch (e) { toast(e.message, 'err'); }
    finally { setRunning(false); }
  };
  const doExport = async (fmt) => {
    const sql = currentSql();
    const name = (rule ? rule.name : '查询结果') + '_' + new Date().toISOString().slice(0, 10);
    setRunning(true);
    try {
      const r = await guard('正在导出 ' + fmt.toUpperCase() + '…（可点「中断查询」）',
        () => apiPost('/api/export', { sql, params: params(), fmt, name }));
      toast('已导出 ' + r.file + '（' + fmtSize(r.size) + '）', 'ok');
      window.location.href = r.url;
    } catch (e) { toast(e.message, 'err'); }
    finally { setRunning(false); }
  };
  $('#btnCsv').onclick = () => doExport('csv');
  $('#btnXlsx').onclick = () => doExport('xlsx');
  $('#btnInterrupt').onclick = async () => {
    try {
      await apiPost('/api/interrupt');
      toast('已发送中断信号', 'ok');
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#btnReset').onclick = () => {
    if (!rule) return;
    delete S.sqlDraft[rule.id];
    S.paramDraft[rule.id] = {};
    renderQuery();
  };
  $('#btnSaveRule').onclick = () => {
    if (!rule) return;
    const sql = currentSql();
    const nm = prompt('规则名称', rule.name === '自由 SQL' ? '我的规则' : rule.name + '（改）');
    if (!nm) return;
    const nr = {
      id: 'u_' + Date.now().toString(36),
      cat: rule.id === '__free' ? '自定义' : (rule.cat || '自定义'),
      name: nm,
      desc: rule.desc || '自定义规则',
      params: (rule.params || []).map((p) => ({ ...p, default: (S.paramDraft[rule.id] || {})[p.key] ?? p.default })),
      sql
    };
    S.rules.push(nr);
    apiPost('/api/rules', { rules: S.rules }).then(() => {
      S.ruleId = nr.id;
      renderQuery();
      toast('已保存到规则库', 'ok');
    }).catch((e) => toast(e.message, 'err'));
  };
}

function openImportRules() {
  modal('从 xlsx 导入查询规则',
    `<div class="small muted" style="margin-bottom:10px">
      格式与「导出规则」一致：分类 / 名称 / 描述 / 参数(JSON) / SQL。同名规则会被覆盖，其余追加。</div>
     <div class="drop" id="ruleDrop" style="padding:18px">
       <b>把规则 xlsx 拖到这里，或点击选择</b>
       <div class="hint">只支持 .xlsx</div>
     </div>
     <input type="file" id="ruleFile" accept=".xlsx" hidden>`,
    '<button class="btn" data-close>关闭</button>');
  const dz = $('#ruleDrop');
  const inp = $('#ruleFile');
  const go = (f) => {
    if (!/\.xlsx$/i.test(f.name)) { toast('只支持 .xlsx', 'err'); return; }
    guard('导入规则…', () => uploadFile(f)).then(async (up) => {
      try {
        const r = await apiPost('/api/rules/import', { fid: up.file_id, mode: 'append' });
        closeModal();
        S.rules = (await apiGet('/api/rules')).rules || [];
        renderQuery();
        toast('已导入 ' + r.count + ' 条规则', 'ok');
      } catch (e) { toast(e.message, 'err'); }
    }).catch((e) => toast(e.message, 'err'));
  };
  dz.onclick = () => inp.click();
  inp.onchange = () => { if (inp.files[0]) go(inp.files[0]); };
  dz.ondragover = (e) => { e.preventDefault(); dz.classList.add('over'); };
  dz.ondragleave = () => dz.classList.remove('over');
  dz.ondrop = (e) => { e.preventDefault(); dz.classList.remove('over'); if (e.dataTransfer.files[0]) go(e.dataTransfer.files[0]); };
}

function renderResult() {
  const r = S.result;
  const box = $('#resultBody');
  if (!r) { box.innerHTML = '<div class="empty">点「运行」查看结果</div>'; $('#resSub').textContent = '尚未运行'; return; }
  $('#resSub').textContent = `${fmtInt(r.rows.length)} 行${r.truncated ? '（已截断）' : ''} · ${r.elapsed_ms}ms`;
  const pages = Math.max(1, Math.ceil(r.rows.length / S.pageSize));
  S.page = Math.min(S.page, pages - 1);
  const slice = r.rows.slice(S.page * S.pageSize, (S.page + 1) * S.pageSize);
  box.innerHTML = `
    <div class="resbar">
      <div class="kpi">
        <div><span>返回行数</span><b>${fmtInt(r.rows.length)}</b></div>
        <div><span>耗时</span><b>${r.elapsed_ms}<span class="small muted">ms</span></b></div>
      </div>
      ${r.truncated ? '<span class="tag amber">结果超过 5000 行已截断，导出可拿全量</span>' : ''}
      <span style="flex:1"></span>
      <button class="btn sm" id="pgPrev" ${S.page === 0 ? 'disabled' : ''}>上一页</button>
      <span class="small muted">${S.page + 1} / ${pages}</span>
      <button class="btn sm" id="pgNext" ${S.page >= pages - 1 ? 'disabled' : ''}>下一页</button>
    </div>
    <div class="tbl-wrap">${tableInner(r.columns, slice)}</div>`;
  const prev = $('#pgPrev'), next = $('#pgNext');
  if (prev) prev.onclick = () => { S.page--; renderResult(); };
  if (next) next.onclick = () => { S.page++; renderResult(); };
}

function tableInner(cols, rows) {
  cols = cols || []; rows = rows || [];
  if (!cols.length) return '<div class="empty" style="padding:10px">无数据</div>';
  const numeric = cols.map((_, i) => {
    let n = 0, tot = 0;
    for (const r of rows.slice(0, 30)) {
      const v = r[i];
      if (v === null || v === undefined || v === '') continue;
      tot++;
      if (typeof v === 'number') n++;
    }
    return tot > 0 && n === tot;
  });
  return `<table class="grid"><thead><tr><th class="rn">#</th>
    ${cols.map((c) => `<th>${esc(c)}</th>`).join('')}</tr></thead><tbody>
    ${rows.map((r, i) => `<tr><td class="rn">${i + 1}</td>
      ${cols.map((_, j) => `<td class="${numeric[j] ? 'num' : ''}">${esc(fmtCell(r[j]))}</td>`).join('')}
    </tr>`).join('')}
  </tbody></table>`;
}
function tableHtml(cols, rows) {
  return `<div class="tbl-wrap" style="max-height:64vh">${tableInner(cols, rows)}</div>`;
}

/* ================================================================== 5 分享 */
function renderShare() {
  const st = S.state;
  const jr = (st.stats.journal || {}).rows || 0;
  const tr = (st.stats.tb || {}).rows || 0;
  $('#main').innerHTML = `
    <div class="card">
      <div class="card-h"><h3>导出数据包（发给同事）</h3>
        <span class="sub">打包成 zip，对方在自己的工具里导入即可得到同一份标准表</span></div>
      <div class="card-b">
        <div class="row" style="gap:16px;margin-bottom:10px">
          <label class="ck"><input type="checkbox" id="ckJ" ${jr ? 'checked' : ''} ${jr ? '' : 'disabled'}>
            序时账 <span class="muted small">${fmtInt(jr)} 行</span></label>
          <label class="ck"><input type="checkbox" id="ckT" ${tr ? 'checked' : ''} ${tr ? '' : 'disabled'}>
            科目余额表 <span class="muted small">${fmtInt(tr)} 行</span></label>
        </div>
        <div class="row" style="margin-bottom:10px">
          <span class="small muted">备注</span>
          <input type="text" id="shareNote" placeholder="例如：某某公司 2025 年 1-12 月序时账" style="flex:1;max-width:420px">
        </div>
        <button class="btn primary" id="btnPack">打包并下载</button>
        <div class="hint">包内含 <code>journal.parquet</code> / <code>tb.parquet</code> + <code>manifest.json</code> + 当前规则库；
          千万行 zstd 压缩后通常在 100–300MB。接收方在「导入数据」页导入数据包即可。</div>
        <div id="packOut" style="margin-top:12px"></div>
      </div>
    </div>
    <div class="card">
      <div class="card-h"><h3>其他</h3></div>
      <div class="card-b row">
        <button class="btn" id="btnBackup2">备份整库（.duckdb 文件）</button>
        <button class="btn" id="btnCsvJ">导出序时账全量 CSV</button>
        <button class="btn" id="btnCsvT">导出科目余额表全量 CSV</button>
        <span class="muted small">整库备份可原样拷给同事，放进 data 目录即可接着用。导出的 CSV 表头是中文，可直接回灌本工具。</span>
      </div>
    </div>`;

  $('#btnPack').onclick = async () => {
    const tables = [];
    if ($('#ckJ').checked) tables.push('journal');
    if ($('#ckT').checked) tables.push('tb');
    if (!tables.length) { toast('至少选一张表', 'err'); return; }
    try {
      const r = await guard('正在打包（大表需要几十秒）…',
        () => apiPost('/api/share/export', { note: $('#shareNote').value, tables }));
      $('#packOut').innerHTML = `<div class="row">
        <span class="tag green">完成</span>
        <b>${esc(r.file)}</b><span class="muted small">${fmtSize(r.size)}</span>
        <a class="btn sm primary" href="${r.url}">下载</a></div>`;
      toast('打包完成：' + r.file, 'ok');
    } catch (e) { toast(e.message, 'err'); }
  };
  const bk = async () => {
    try {
      const r = await guard('备份中…', () => apiPost('/api/backup'));
      toast('已备份 ' + r.file, 'ok');
      window.location.href = r.url;
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#btnBackup2').onclick = bk;
  const expT = async (t, nm) => {
    try {
      const r = await guard('导出中…', () => apiPost('/api/export', { sql: `SELECT * FROM ${t}`, fmt: 'csv', name: nm }));
      toast('已导出 ' + r.file, 'ok');
      window.location.href = r.url;
    } catch (e) { toast(e.message, 'err'); }
  };
  $('#btnCsvJ').onclick = () => expT('journal', '序时账全量');
  $('#btnCsvT').onclick = () => expT('tb', '科目余额表全量');
}

/* ------------------------------------------------------------------ 说明 */
function showHelp() {
  const f = S.state ? S.state.fields : { journal: [], tb: [] };
  const doc = (list) => `<div class="field-doc">${list.map((x) =>
    `<div><b>${x.key}</b><span class="muted">${esc(x.label)}</span></div>`).join('')}</div>`;
  modal('使用说明', `
    <h3 style="margin:0 0 6px">标准流程</h3>
    <ol class="steps">
      <li><b>导入</b>：拖入 xlsx / csv，系统把每个工作表原样读成一张原始表（全部文本，不清洗），并自动猜表头行。<b>同一套表头的表自动归成一组</b>，一组只选一次目标表。同一页可以导入<b>账簿对照表</b>（账簿名称 → 公司名称 / 简称）</li>
      <li><b>清洗</b>：用可配置的动作清掉合计行、重复表头、页脚、缩进空格，或把辅助核算明细行补齐。只写清洗表，原始表一列不动。调好的方案会自动进<b>方案库</b>（右上角下拉框随时套用），也能<b>导出 xlsx 存档</b>（含字段映射，按表头关联）</li>
      <li><b>字段映射</b>：指定这张表属于「序时账」还是「科目余额表」，系统按列名自动匹配，人工核对一遍；余额表的辅助核算还要标一下类型；可指定<b>账簿名称来源</b>（某列 / 文件名 / 工作表名 / 手填）</li>
      <li><b>合并</b>：把所有映射到同一目标的工作表拼成一张标准表，附加来源文件与原始行号，并按账簿名称带出公司名称、简称</li>
      <li><b>数据检查</b>：<b>科目余额表 × 序时账 核对</b>，按「主体（公司名称，无则账簿名称）+ 一级科目」分组算本年累计发生额（借 / 贷 / 净额，本币与原币同时核），不符和单边科目单独列出，并给出<b>按主体汇总</b>定位是哪家公司</li>
      <li><b>查账</b>：用现成规则（表单填参数）或直接写 SQL，结果可导出 CSV / Excel（表头自动中文）</li>
      <li><b>分享</b>：打包成 zip 发给同事，对方在自己的工具里「导入数据包」即可</li>
    </ol>
    <h3 style="margin:18px 0 6px">清洗能做什么</h3>
    <ul class="steps">
      <li><b>选行</b>：删空行（<b>可勾选任意多列判空</b>，这几列全空才删）、按关键词删合计/小计/重复表头/页脚、整行去重</li>
      <li><b>改值</b>：文本归一（去全角空格缩进）、编码补零（1001 → 001001）</li>
      <li><b>填值</b>：辅助核算明细成行时，把科目编码/名称向下填满</li>
      <li><b>派生</b>：按编码长度推断科目级次；<b>提取一级科目</b>（把「1001-库存现金」「管理费用/办公费」按分隔符取第一段，生成一级科目列）；把「客户:A/部门:B」拆成辅助核算1/2/3；<b>填充固定值</b>（整列填同一个值，用于补公司名称、币种）</li>
      <li><b>校验</b>：期初 + 本期借 − 本期贷 = 期末，只出报告不改数据</li>
      <li>每步都能单独关掉；「试运行」显示每步删了几行、前后对照；被剔除的行全部留档，可查「第几步、什么原因」</li>
    </ul>
    <h3 style="margin:18px 0 6px">辅助核算怎么映射</h3>
    <div class="small muted" style="line-height:1.9">
      余额表的辅助核算分三层：<b>位置槽</b>（辅助核算1/2/3，原样承接导出列）→ <b>类型</b>（银行账号 / 客商 / 部门 / 项目 / 职员 / 存货 / 混合）→ <b>语义列</b>（bank_account、counterparty_name、dept_name…）。<br>
      一格多类型（「客户:阿里巴巴/部门:销售部」）把类型选「混合」，系统按前缀自动分到对应语义列；整列就是一个名字的，选对应类型即可。<br>
      全部不填也没关系，<code>aux1/2/3</code> 会原样保留，还能用 <code>aux_text</code> 做全文模糊搜。
    </div>
    <h3 style="margin:18px 0 6px">序时账 journal 字段</h3>${doc(f.journal)}
    <h3 style="margin:18px 0 6px">科目余额表 tb 字段</h3>${doc(f.tb)}
    <h3 style="margin:18px 0 6px">SQL 示例</h3>
    <div class="mono-box">-- 某科目全年逐笔
SELECT vdate, vno, summary, debit, credit
FROM journal WHERE acct_code LIKE '1002%' ORDER BY vdate LIMIT 200;

-- 按摘要关键词找分录
SELECT vdate, vno, summary, acct_code, debit
FROM journal WHERE summary LIKE '%咨询费%';

-- 科目余额表里期末余额最大的 20 个科目
SELECT acct_code, acct_name, end_bal FROM tb ORDER BY abs(coalesce(end_bal,0)) DESC LIMIT 20;

-- 凭证借贷不平
SELECT period, vno, sum(debit)-sum(credit) d FROM journal
GROUP BY 1,2 HAVING abs(d) > 0.005;</div>
    <h3 style="margin:18px 0 6px">几个说明</h3>
    <ul class="steps">
      <li>所有数据都在本机 <code>data</code> 目录的 DuckDB 文件里，不联网、不上传</li>
      <li>原始表全部按文本存，所以科目编码 <code>1001</code> 的前导零、<code>记-0001</code> 都不会丢</li>
      <li>金额转换会自动去掉千分位与全角符号；<code>(1,000.00)</code> 视为 −1000</li>
      <li>日期支持 2025-01-31 / 2025/1/31 / 20250131 / 2025年1月31日 / Excel 序列号</li>
      <li>改表头行：在导入页点「预览」，点某一行即可设为表头；两层表头把「表头行数」改成 2</li>
      <li>规则存于 <code>data/rules.json</code>，可手工编辑或通过「保存为规则」新增</li>
    </ul>`);
  bindClose();
}

boot();
