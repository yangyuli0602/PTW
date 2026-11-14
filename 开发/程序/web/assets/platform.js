/* 平台级能力：日志 / 内置 AI / 检查更新
   独立文件，不改业务代码；依赖 app.js 里的 apiGet/apiPost/esc/toast/modal/closeModal。 */

/* ---------------------------------------------------------------- 日志 */
async function openLogs() {
  modal('运行日志', '<div id="logBox" class="muted">读取中…</div>',
    '<button class="btn sm ghost" id="logRefresh">刷新</button>'
    + '<button class="btn sm" id="logClear">清空界面</button>'
    + '<button class="btn sm primary" id="logDl">下载日志包</button>');
  const box = document.getElementById('logBox');
  const paint = (r) => {
    const items = r.items || [];
    if (!items.length) { box.innerHTML = '<div class="muted">暂无日志</div>'; return; }
    const color = { ERROR: 'var(--danger)', WARN: 'var(--warn)', INFO: 'var(--ink2)' };
    box.innerHTML = '<div class="muted" style="margin-bottom:6px">' + esc(r.paths || '')
      + ' | 目录：' + esc(r.dir || '') + ' | 共 ' + items.length + ' 条</div>'
      + '<div style="max-height:52vh;overflow:auto;font:12px/1.55 Consolas,monospace;'
      + 'background:#0f1626;color:#c8d3e5;padding:10px;border-radius:8px;white-space:pre-wrap">'
      + items.map((x) => '<div style="color:' + (color[x.level] || '#c8d3e5') + '">'
        + esc(x.ts + ' [' + x.level + ']' + (x.tag ? ' ' + x.tag : '') + ' ' + x.msg) + '</div>').join('')
      + '</div>';
  };
  const load = async () => {
    try { paint(await apiGet('/api/logs?n=400')); }
    catch (e) { box.innerHTML = '<div class="muted">读取失败：' + esc(e.message) + '</div>'; }
  };
  await load();
  document.getElementById('logRefresh').onclick = load;
  document.getElementById('logDl').onclick = () => { location.href = '/api/logs/download'; };
  document.getElementById('logClear').onclick = async () => {
    await apiPost('/api/logs/clear', {}); await load(); toast('界面日志已清空（文件保留）', 'ok');
  };
}

/* ---------------------------------------------------------------- 内置 AI */
async function openAI() {
  modal('AI 助手（内置）', '<div id="aiBox" class="muted">读取中…</div>', '');
  const box = document.getElementById('aiBox');
  let st;
  try { st = await apiGet('/api/ai/status'); }
  catch (e) { box.innerHTML = '读取失败：' + esc(e.message); return; }

  const opts = (st.providers || []).map((p) =>
    `<option value="${esc(p.key)}" ${p.key === st.provider ? 'selected' : ''}>${esc(p.label)}</option>`).join('');
  box.innerHTML = `
    <div class="card" style="padding:12px;margin-bottom:10px">
      <div class="row" style="gap:10px;flex-wrap:wrap;align-items:center">
        <label><input type="checkbox" id="aiEn" ${st.enabled ? 'checked' : ''}> 启用内置 AI</label>
        <select id="aiProv" class="inp" style="min-width:260px">${opts}</select>
      </div>
      <div class="row" style="gap:10px;margin-top:8px;flex-wrap:wrap">
        <input id="aiKey" class="inp" type="password" placeholder="API Key（只存本机 data/settings.json）"
               value="" style="flex:1;min-width:300px">
        <button class="btn sm" id="aiTest">测试连接</button>
        <button class="btn sm primary" id="aiSave">保存</button>
      </div>
      <div class="row muted" style="gap:16px;margin-top:6px;font-size:12px">
        <span>模型：<b id="aiModel">${esc(st.model)}</b></span>
        <span>端点：${esc(st.base_url)}</span>
        <span>${st.has_key ? '已保存 Key' : '未填 Key'}</span>
      </div>
      <div id="aiTestOut" class="muted" style="margin-top:6px"></div>
    </div>
    <div class="muted" style="margin:4px 0 6px">
      三条通道与「外部 AI 旁路」完全等价（写同一批 *_result.json），可只跑其中一条，也可都不跑由人工填。
    </div>
    <div id="aiTasks"></div>`;
  document.getElementById('aiProv').onchange = async (ev) => {
    const p = (st.providers || []).find((x) => x.key === ev.target.value);
    if (p) { document.getElementById('aiModel').textContent = p.model; }
  };
  document.getElementById('aiSave').onclick = async () => {
    const body = { ai_enabled: document.getElementById('aiEn').checked,
                   ai_provider: document.getElementById('aiProv').value };
    const k = document.getElementById('aiKey').value.trim();
    if (k) body.ai_key = k;
    try {
      const r = await apiPost('/api/ai/save', body);
      document.getElementById('aiModel').textContent = r.model;
      toast('已保存', 'ok');
    } catch (e) { toast(e.message, 'err'); }
  };
  document.getElementById('aiTest').onclick = async () => {
    const out = document.getElementById('aiTestOut');
    out.textContent = '测试中…';
    try {
      const r = await apiPost('/api/ai/test', {});
      out.innerHTML = r.ok
        ? '✅ 连通（' + r.ms + 'ms）：' + esc(r.reply || '')
        : '❌ ' + esc(r.error || '失败') + (r.detail ? '<br><span class="muted">' + esc(r.detail) + '</span>' : '');
    } catch (e) { out.textContent = '❌ ' + e.message; }
  };

  const tb = document.getElementById('aiTasks');
  tb.innerHTML = (st.tasks || []).map((t) => `
    <div class="card" style="padding:10px 12px;margin-bottom:8px">
      <div class="row" style="justify-content:space-between;align-items:center;gap:10px">
        <div><b>${esc(t.label)}</b>
          <div class="muted" style="font-size:12px">待办 ${t.pending === null ? '未生成清单' : t.pending + ' 条'}
            ${t.done ? ' · 已有结果（可重跑覆盖）' : ''}</div></div>
        <button class="btn sm primary" data-ai-run="${esc(t.id)}"
          ${t.pending ? '' : 'disabled'}>运行</button>
      </div>
      <div class="muted ai-run-out" data-out="${esc(t.id)}" style="margin-top:4px;font-size:12px"></div>
    </div>`).join('');
  tb.querySelectorAll('[data-ai-run]').forEach((b) => {
    b.onclick = async () => {
      const id = b.dataset.aiRun;
      const out = tb.querySelector('[data-out="' + id + '"]');
      b.disabled = true; out.textContent = 'AI 处理中（可能 30–60 秒）…';
      try {
        const r = await apiPost('/api/ai/run', { task: id, limit: 60 });
        out.innerHTML = r.ok
          ? '✅ 写入 ' + r.written + ' 条 → ' + esc((r.result_file || '').replace(/^.*[\\\/]/, ''))
          : '❌ ' + esc(r.error || '失败');
        toast(r.ok ? 'AI 任务完成' : 'AI 任务失败', r.ok ? 'ok' : 'err');
      } catch (e) { out.textContent = '❌ ' + e.message; }
      b.disabled = false;
    };
  });
}

/* ---------------------------------------------------------------- 检查更新 */
async function checkUpdate() {
  toast('正在检查更新…', 'ok');
  let r;
  try { r = await apiGet('/api/update/check'); }
  catch (e) { modal('检查更新', '<div>检查失败：' + esc(e.message) + '</div>', ''); return; }

  if (!r.ok) {
    const att = (r.attempts || []).map((a) => ' · ' + esc(a[0]) + '：' + esc(a[2])).join('<br>');
    modal('检查更新', `<div>当前版本 <b>${esc(r.current)}</b></div>
      <div class="muted" style="margin-top:8px">${esc(r.error || '')}</div>
      ${att ? '<div class="muted" style="margin-top:6px;font-size:12px">' + att + '</div>' : ''}`, '');
    return;
  }
  const i = r.info || {};
  if (!i.is_new) {
    modal('检查更新', `<div>已是最新版本：<b>${esc(r.current)}</b>
      ${i.version && i.version !== r.current ? '（远端 ' + esc(i.version) + '）' : ''}</div>`, '');
    return;
  }
  const mb = i.size ? (i.size / 1048576).toFixed(1) + ' MB' : '未知大小';
  modal('发现新版本', `
    <div style="font-size:14px;margin-bottom:8px">
      当前 <b>${esc(r.current)}</b> → 新版 <b style="color:var(--ok)">${esc(i.version)}</b>
      <span class="muted">（${mb}${i.date ? ' · ' + esc(i.date) : ''}）</span>
    </div>
    <div style="max-height:280px;overflow:auto;background:var(--bg2,#f6f8fb);
         padding:10px;border-radius:8px;white-space:pre-wrap;font-size:12.5px">${esc(i.notes || '（无更新说明）')}</div>
    <div class="muted" style="margin-top:8px;font-size:12px">
      点「立即更新」后自动下载、校验、替换并重启，无需其他操作。数据目录不受影响。</div>`,
    '<button class="btn sm ghost" id="upLater">稍后</button>'
    + '<button class="btn sm primary" id="upNow">立即更新</button>');
  document.getElementById('upLater').onclick = () => closeModal();
  document.getElementById('upNow').onclick = async () => {
    const btn = document.getElementById('upNow');
    btn.disabled = true; btn.textContent = '下载中…';
    const body = document.getElementById('modalBody');
    try {
      const a = await apiPost('/api/update/apply', {});
      if (!a.ok) { toast(a.error || '更新失败', 'err'); btn.disabled = false; btn.textContent = '立即更新'; return; }
      document.getElementById('modalFoot').innerHTML = '';
      // 后台在下载，前端只负责显示进度（下载可能要几分钟，不能靠一个长连接死等）
      const t0 = Date.now();
      const tick = async () => {
        if (Date.now() - t0 > 30 * 60000) return;
        let p = {};
        try { p = await apiGet('/api/update/progress'); } catch (e) { /* 重启瞬间会失败 */ }
        const ph = p.phase || '';
        if (ph === 'failed') {
          body.innerHTML = '<div style="color:var(--danger)">更新失败：'
            + esc(p.error || '未知原因') + '</div><div class="muted" style="margin-top:6px">'
            + '可以稍后重试；当前版本仍可正常使用。</div>';
          btn.disabled = false; btn.textContent = '立即更新';
          return;
        }
        if (ph === 'restarting') {
          body.innerHTML = `<div style="padding:8px 0">正在替换程序并重启（${esc(p.version || '')}），窗口会自动关闭…</div>`;
          setTimeout(waitBack, 3000);
          return;
        }
        const total = p.total || 0, got = p.got || 0;
        const pct = total ? Math.min(99, Math.round(got * 100 / total)) : 0;
        body.innerHTML = `<div style="padding:4px 0">正在下载 ${esc(p.version || '')}　${total ? pct + '%' : ''}</div>
          <div style="height:6px;background:var(--bg2,#e8eef6);border-radius:4px;overflow:hidden;margin-top:8px">
            <div style="height:100%;width:${pct}%;background:var(--ok,#2e9e5b);transition:width .3s"></div></div>
          <div class="muted" style="margin-top:8px;font-size:12px">${total
            ? (got / 1048576).toFixed(1) + ' / ' + (total / 1048576).toFixed(1) + ' MB'
            : esc(p.note || '')}</div>`;
        setTimeout(tick, 800);
      };
      const waitBack = async () => {
        try { await apiGet('/api/modules'); location.reload(); return; } catch (e) { /* 还没起来 */ }
        setTimeout(waitBack, 1500);
      };
      setTimeout(tick, 600);
    } catch (e) {
      toast('更新失败：' + e.message, 'err');
      btn.disabled = false; btn.textContent = '立即更新';
    }
  };
}

/* ---------------------------------------------------------------- 退出 */
async function quitApp() {
  if (!confirm('退出 PTW 平台？\n（未保存的操作会中断，数据在退出前会自动落盘）')) return;
  try {
    await apiPost('/api/app/quit', {});
    document.body.innerHTML = '<div style="display:flex;align-items:center;justify-content:center;'
      + 'height:100vh;font:15px/1.8 system-ui;color:#9fb0c8">'
      + 'PTW 平台已退出，可以关闭本窗口。</div>';
    setTimeout(() => { try { window.close(); } catch (e) {} }, 600);
  } catch (e) {
    toast('退出失败：' + e.message, 'err');
  }
}

/* 顶栏按钮绑定（DOM 就绪后） */
document.addEventListener('DOMContentLoaded', () => {
  const bind = (id, fn) => { const b = document.getElementById(id); if (b) b.onclick = fn; };
  bind('btnLogs', openLogs);
  bind('btnAI', openAI);
  bind('btnUpdate', checkUpdate);
  bind('btnQuit', quitApp);
});
