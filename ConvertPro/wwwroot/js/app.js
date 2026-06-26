
(function(){
'use strict';

/* ============================================================
   状态管理
   ============================================================ */
const state = {};
function getState(pageId) {
    if (!state[pageId]) state[pageId] = { files: [], isConverting: false, cancelled: false };
    return state[pageId];
}

/* ============================================================
   Toast 通知
   ============================================================ */
function showToast(msg, type='info') {
    const c = document.getElementById('toastContainer');
    const t = document.createElement('div');
    t.className = 'toast ' + type;
    const icons = { error:'❌', warning:'⚠️', success:'✅', info:'ℹ️' };
    t.innerHTML = (icons[type]||'ℹ️') + ' ' + msg;
    c.appendChild(t);
    setTimeout(() => { t.style.opacity='0'; t.style.transition='opacity .3s'; setTimeout(()=>t.remove(),300); }, 3500);
}

// C# 调用入口：显示 toast 通知
window.__showToast = function(info) {
    try {
        const msg = (info && info.toast) || '操作完成';
        const type = (info && info.type) || 'info';
        showToast(msg, type);
    } catch(e) {}
};

/* ============================================================
   导航（流式向导）
   - titlebar-nav 按钮（首页/历史/设置/关于）
   - 首页 source-card 的 target-chip → 跳转到对应工作页
   ============================================================ */
const navBtns = document.querySelectorAll('.nav-item');
const pages = document.querySelectorAll('.page');
const statusHint = document.getElementById('statusHint');
const titlebarNavBtns = document.querySelectorAll('.titlebar-nav-btn');

// 全局页面切换函数：home/history/settings/工作页名（如 pdf2word）
function navigateToPage(pageId) {
    pages.forEach(p => p.classList.remove('active'));
    const target = document.getElementById('page-' + pageId);
    if (target) target.classList.add('active');

    // 同步 titlebar-nav 的 active 状态
    titlebarNavBtns.forEach(b => {
        const navTarget = b.dataset.nav;
        let isActive = false;
        if (pageId === 'home' && navTarget === 'home') isActive = true;
        else if (pageId === 'history' && navTarget === 'history') isActive = true;
        else if (pageId === 'settings' && navTarget === 'settings') isActive = true;
        b.classList.toggle('active', isActive);
    });

    // 状态栏提示当前页
    const labelMap = {
        home: '首页',
        history: '转换历史',
        settings: '设置',
        png2ico: 'PNG → ICO',
        pdf2word: 'PDF → Word',
        word2pdf: 'Word → PDF',
        pdf2excel: 'PDF → Excel',
        excel2pdf: 'Excel → PDF',
        pdf2ppt: 'PDF → PPT',
        ppt2pdf: 'PPT → PDF',
        word2ppt: 'Word → PPT',
        excel2word: 'Excel → Word',
        word2excel: 'Word → Excel'
    };
    statusHint.textContent = labelMap[pageId] || '就绪';

    // 滚动到顶部
    document.querySelector('.main')?.scrollTo({ top: 0, behavior: 'instant' });
}

// titlebar-nav 按钮
titlebarNavBtns.forEach(btn => {
    btn.addEventListener('click', () => navigateToPage(btn.dataset.nav));
});

// 首页 target-chip → 跳转到对应工作页
document.querySelectorAll('.target-chip').forEach(chip => {
    chip.addEventListener('click', () => {
        const target = chip.dataset.page;
        if (target) navigateToPage(target);
    });
});

// 兼容旧 sidebar nav-item（虽然 sidebar 已隐藏，但保留事件以备回退）
navBtns.forEach(btn => {
    btn.addEventListener('click', function() {
        navigateToPage(this.dataset.page);
    });
});

// 启动时默认显示首页
navigateToPage('home');

/* ============================================================
   窗口控制（WebView2 → C# 通信）
   ============================================================ */
// ⚠️ 不使用 -webkit-app-region（WebView2 中不可靠），改用 Win32 联动
document.getElementById('btnMinimize').addEventListener('click', () => {
    window.postAction('window_minimize');
});
document.getElementById('btnMaximize').addEventListener('click', () => {
    window.postAction('window_maximize');
});
document.getElementById('btnClose').addEventListener('click', () => {
    window.postAction('window_close');
});

// 标题栏 mousedown → 窗口拖拽（Win32 联动）
document.querySelector('.titlebar').addEventListener('mousedown', function(e) {
    // 排除按钮区（最小化/最大化/关闭）和导航按钮（首页/历史/设置）
    if (e.target.closest('.titlebar-btn')) return;
    if (e.target.closest('.titlebar-nav-btn')) return;
    window.postAction('window_drag');
});

// 标题栏双击 → 最大化/还原
document.querySelector('.titlebar').addEventListener('dblclick', function(e) {
    if (e.target.closest('.titlebar-btn')) return;
    if (e.target.closest('.titlebar-nav-btn')) return;
    window.postAction('window_maximize');
});

/* ============================================================
   关于对话框
   ============================================================ */
const aboutModal = document.getElementById('aboutModal');
document.getElementById('btnOpenAbout').addEventListener('click', () => aboutModal.classList.add('visible'));
document.getElementById('aboutClose').addEventListener('click', () => aboutModal.classList.remove('visible'));
aboutModal.addEventListener('click', e => { if(e.target===aboutModal) aboutModal.classList.remove('visible'); });

/* ============================================================
   拖拽 & 文件选择
   ============================================================ */
document.querySelectorAll('.drop-zone').forEach(zone => {
    const pageId = zone.dataset.page;
    // ⚠️ 不再使用浏览器 file input — 改为调用 C# OpenFileDialog
    zone.addEventListener('click', function() {
        window.postAction('open_file', { type: pageId });
    });
    zone.addEventListener('dragover', e => { e.preventDefault(); zone.classList.add('dragover'); });
    zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
    zone.addEventListener('drop', function(e) {
        e.preventDefault(); zone.classList.remove('dragover');
        if (e.dataTransfer.files && e.dataTransfer.files.length) {
            validateAndLoad(pageId, e.dataTransfer.files);
        }
    });
});
document.querySelectorAll('.file-input').forEach(input => {
    input.addEventListener('click', function(e) {
        e.preventDefault();
        const pageId = this.dataset.page;
        window.postAction('open_file', { type: pageId });
    });
});

/* ============================================================
   5️⃣ 输入校验
   ============================================================ */
const ALLOWED = {
    png2ico:  ['.png'],
    pdf2word: ['.pdf'],
    word2pdf: ['.docx','.doc'],
    pdf2excel:['.pdf'],
    excel2pdf:['.xlsx','.xls'],
    pdf2ppt:  ['.pdf']
};

function validateAndLoad(pageId, fileList) {
    const allowed = ALLOWED[pageId] || [];
    const valid = [], invalid = [];
    Array.from(fileList).forEach(f => {
        const ext = '.' + f.name.split('.').pop().toLowerCase();
        if (allowed.includes(ext)) valid.push(f);
        else invalid.push(f.name);
    });
    if (invalid.length) {
        const zone = document.querySelector(`.drop-zone[data-page="${pageId}"]`);
        if (zone) { zone.classList.add('invalid'); setTimeout(()=>zone.classList.remove('invalid'),500); }
        showToast(`不支持以下格式：${invalid.join('、')}。允许：${allowed.join('、')}`, 'error');
    }
    if (valid.length) handleFilesForPage(pageId, valid);
}

/* ============================================================
   文件处理
   ============================================================ */
function handleFilesForPage(pageId, files) {
    const st = getState(pageId);
    // WebView2 + Windows File Explorer 拖拽的 File 对象带 .path 属性（Chromium 非标准扩展）
    const arr = Array.from(files);
    const hasPath = arr.some(f => f.path);
    if (!hasPath) {
        // 拖拽来源不是本地文件系统（如从网页拖入），回退到 C# 文件对话框
        showToast('拖入的内容无文件路径，请改用点击选择文件', 'info');
        window.postAction('open_file', { type: pageId });
        return;
    }
    st.files = arr.map(f => ({
        name: f.name, size: f.size, path: f.path, _realPath: f.path
    }));
    st.cancelled = false;
    const PA = q(`.preview-area[data-page="${pageId}"]`);
    if (!PA) return;
    PA.classList.add('visible');
    const first = files[0];
    qs(PA,'.preview-info .name').textContent = first.name;
    const kb = (first.size/1024).toFixed(1);
    qs(PA,'.preview-info .meta').innerHTML = `<span>大小: ${kb} KB</span>`;
    if (pageId==='png2ico') {
        const m = qs(PA,'.preview-info .meta');
        const s = document.createElement('span'); s.className='dims'; s.textContent='尺寸: 512×512';
        m.prepend(s);
    }
    // file list
    const FL = qs(PA,`.file-list`);
    if (FL) {
        FL.innerHTML = files.map((f,i)=> {
            const sz = (f.size/1024).toFixed(1);
            return `<div class="file-list-item" data-idx="${i}"><span class="idx">${i+1}</span><span class="name">${f.name}</span><span class="status">${sz} KB</span></div>`;
        }).join('');
    }
    q(`.drop-zone[data-page="${pageId}"]`).style.display = 'none';
    q(`.btn-convert[data-page="${pageId}"]`).disabled = false;
    q(`.btn-clear[data-page="${pageId}"]`).style.display = 'inline-block';
    q(`.page.active .status-text`).textContent = `已选择 ${files.length} 个文件`;
    statusHint.textContent = `已选择 ${files.length} 个文件`;
    qs(PA,`.done-section`)?.classList.remove('visible');
    qs(PA,`.progress-bar`)?.classList.remove('visible');
    if (pageId==='png2ico') updateIcoPreview(pageId);
    // 文档预览：填充并显示
    if (DOC_PREVIEW_DATA && DOC_PREVIEW_DATA[pageId]) {
        const dp = q(`.doc-preview[data-page="${pageId}"]`);
        if (dp) {
            dp.classList.add('visible');
            // 重置输出 tab
            const outTab = dp.querySelector('.tab[data-pane="output"]');
            if (outTab) { outTab.disabled = true; outTab.classList.remove('active'); if(outTab.querySelector('.count')) outTab.querySelector('.count').textContent = '转换后可见'; }
            const srcTab = dp.querySelector('.tab[data-pane="source"]');
            if (srcTab) srcTab.classList.add('active');
            dp.querySelectorAll('.preview-pane').forEach(p => p.classList.remove('active'));
            const srcPane = dp.querySelector('.preview-pane[id$="-source"]');
            if (srcPane) srcPane.classList.add('active');
            populateDocPreview(pageId);
        }
    }
    q(`.file-input[data-page="${pageId}"]`).value = '';
}

function q(s) { return document.querySelector(s); }
function qs(p,s) { return p.querySelector(s); }

/* ============================================================
   移除/清空
   ============================================================ */
document.querySelectorAll('.btn-remove').forEach(b => b.addEventListener('click', function(){ resetPage(this.dataset.page); }));
document.querySelectorAll('.btn-clear').forEach(b => b.addEventListener('click', function(){ resetPage(this.dataset.page); }));

function resetPage(pageId) {
    const st = getState(pageId);
    st.files = []; st.isConverting = false; st.cancelled = false;
    const PA = q(`.preview-area[data-page="${pageId}"]`);
    if (PA) PA.classList.remove('visible');
    const z = q(`.drop-zone[data-page="${pageId}"]`);
    if (z) z.style.display = '';
    const cb = q(`.btn-convert[data-page="${pageId}"]`);
    if (cb) { cb.disabled = true; cb.textContent = '开始转换'; }
    const cl = q(`.btn-clear[data-page="${pageId}"]`);
    if (cl) cl.style.display = 'none';
    const can = q(`.btn-cancel[data-page="${pageId}"]`);
    if (can) can.style.display = 'none';
    const stxt = q(`.page.active .status-text`);
    if (stxt) stxt.textContent = '拖入文件开始转换';
    statusHint.textContent = '就绪';
    if (PA) {
        qs(PA,`.done-section`)?.classList.remove('visible');
        qs(PA,`.progress-bar`)?.classList.remove('visible');
    }
    const ip = q(`.ico-preview-sizes[data-page="${pageId}"]`);
    if (ip) ip.classList.remove('visible');
    const fl = q(`.file-list[data-page="${pageId}"]`);
    if (fl) fl.innerHTML = '';
    // 隐藏文档预览
    const dp = q(`.doc-preview[data-page="${pageId}"]`);
    if (dp) dp.classList.remove('visible');
    q(`.file-input[data-page="${pageId}"]`).value = '';
}

/* ============================================================
   12️⃣ 转换前确认
   ============================================================ */
const confirmModal = document.getElementById('confirmModal');
const confirmOk = document.getElementById('confirmOk');
const confirmCancel = document.getElementById('confirmCancel');
const confirmClose = document.getElementById('confirmClose');
let pendingConversion = null;

function showConfirm(pageId) {
    const st = getState(pageId);
    if (!st.files.length) return;
    const preconfirm = document.getElementById('setting-preconfirm').checked;
    if (!preconfirm) { triggerConversion(pageId); return; }
    const names = st.files.map(f=>f.name).join('、');
    const titleMap = { png2ico:'PNG → ICO', pdf2word:'PDF → Word', word2pdf:'Word → PDF', pdf2excel:'PDF → Excel', excel2pdf:'Excel → PDF', pdf2ppt:'PDF → PPT' };
    document.getElementById('confirmTitle').textContent = `确认转换 — ${titleMap[pageId]||pageId}`;
    document.getElementById('confirmText').innerHTML =
        `即将转换 <strong>${st.files.length}</strong> 个文件：<br><span style="font-size:12px;color:#888;">${names.length>60?names.slice(0,60)+'…':names}</span><br><br>` +
        `输出目录：<span style="font-family:monospace;font-size:12px;">${getOutputPath(pageId)}</span>`;
    pendingConversion = pageId;
    confirmModal.classList.add('visible');
}

confirmOk.addEventListener('click', () => {
    confirmModal.classList.remove('visible');
    if (pendingConversion) { triggerConversion(pendingConversion); pendingConversion = null; }
});
[confirmCancel, confirmClose].forEach(el => el.addEventListener('click', () => {
    confirmModal.classList.remove('visible');
    pendingConversion = null;
}));
confirmModal.addEventListener('click', e => { if (e.target===confirmModal) { confirmModal.classList.remove('visible'); pendingConversion=null; } });

/* ============================================================
   输出目录
   ============================================================ */
function getOutputPath(pageId) {
    const el = document.getElementById('outputPath-'+pageId);
    return el ? el.textContent : 'C:\\Users\\Desktop\\ConvertPro_Output';
}

document.querySelectorAll('.btn-browse').forEach(b => {
    b.addEventListener('click', function() {
        const pageId = this.dataset.page;
        const path = prompt('请输入输出目录路径：', getOutputPath(pageId));
        if (path && path.trim()) {
            const el = document.getElementById('outputPath-'+pageId);
            if (el) el.textContent = path.trim();
            showToast(`输出目录已设置为：${path.trim()}`, 'success');
        }
    });
});

/* ============================================================
   8️⃣ 历史记录
   ============================================================ */
let historyLog = JSON.parse(localStorage?.getItem('convertpro_history') || '[]');
function addHistory(entry) {
    historyLog.unshift({ ...entry, time: new Date().toLocaleString() });
    if (historyLog.length > 100) historyLog = historyLog.slice(0,100);
    try { localStorage?.setItem('convertpro_history', JSON.stringify(historyLog)); } catch(e) {}
    renderHistory();
}
function renderHistory() {
    const c = document.getElementById('historyContent');
    if (!historyLog.length) {
        c.innerHTML = '<div class="history-empty">暂无转换记录</div>';
        document.getElementById('btnClearHistory').style.display = 'none';
        return;
    }
    document.getElementById('btnClearHistory').style.display = 'inline-block';
    c.innerHTML = `<table class="history-table">
        <thead><tr><th>时间</th><th>类型</th><th>文件</th><th>状态</th></tr></thead>
        <tbody>${historyLog.map(h => `<tr>
            <td style="white-space:nowrap;">${h.time}</td>
            <td>${h.type}</td>
            <td style="max-width:200px;overflow:hidden;text-overflow:ellipsis;white-space:nowrap;">${h.file}</td>
            <td><span class="status-tag ${h.status}">${h.status === 'success' ? '✅ 成功' : h.status === 'error' ? '❌ 失败' : '⏭️ 跳过'}</span></td>
        </tr>`).join('')}</tbody>
    </table>`;
}
document.getElementById('btnClearHistory').addEventListener('click', () => {
    if (confirm('确定清空所有转换历史吗？')) {
        historyLog = [];
        try { localStorage?.setItem('convertpro_history', JSON.stringify(historyLog)); } catch(e) {}
        renderHistory();
        showToast('历史记录已清空', 'info');
    }
});
renderHistory();

// Listen for history add from conversion
/* ============================================================
   1️⃣1️⃣ 右键菜单
   ============================================================ */
const ctxMenu = document.getElementById('ctxMenu');
let ctxPageId = null;

document.addEventListener('contextmenu', e => {
    const item = e.target.closest('.file-list-item');
    if (!item) { ctxMenu.classList.remove('visible'); return; }
    e.preventDefault();
    const list = item.closest('.file-list');
    ctxPageId = list ? list.dataset.page : null;
    // Find file name
    ctxMenu.style.left = e.clientX + 'px';
    ctxMenu.style.top = e.clientY + 'px';
    ctxMenu.classList.add('visible');
});
document.addEventListener('click', () => ctxMenu.classList.remove('visible'));

document.querySelectorAll('.ctx-menu-item').forEach(el => {
    el.addEventListener('click', function() {
        const action = this.dataset.action;
        ctxMenu.classList.remove('visible');
        if (action === 'remove') {
            // Remove first file
            if (ctxPageId) {
                const st = getState(ctxPageId);
                if (st.files.length > 0) {
                    st.files.shift();
                    if (st.files.length === 0) resetPage(ctxPageId);
                    else handleFilesForPage(ctxPageId, st.files);
                    showToast('已移除一个文件', 'info');
                }
            }
        } else if (action === 'clear') {
            if (ctxPageId) resetPage(ctxPageId);
        }
    });
});

/* ============================================================
   3️⃣ 取消转换
   ============================================================ */
document.querySelectorAll('.btn-cancel').forEach(b => {
    b.addEventListener('click', function() {
        const pageId = this.dataset.page;
        const st = getState(pageId);
        if (!st.isConverting) return;
        // 通知 C# 取消当前转换
        window.postAction('cancel_convert');
        showToast('已请求取消转换', 'warning');
    });
});

/* ============================================================
   6️⃣ 冲突处理
   ============================================================
   协议：C# 检测到输出文件已存在 → 调用 window.__showConflictBatch(info)
        → JS 显示弹窗 → 用户点击按钮 → postAction('conflict_resolve', {choice})
        → C# 通过 TaskCompletionSource 拿到选择继续执行
   ============================================================ */
const conflictModal = document.getElementById('conflictModal');
const conflictClose = document.getElementById('conflictClose');
const conflictFileNameEl = document.getElementById('conflictFileName');

// C# 调用入口
window.__showConflictBatch = function(info) {
    try {
        if (!info) info = {};
        const fileName = info.fileName || 'output';
        const count = info.count || 1;
        // 多个冲突文件时显示 "name 及其他 N 个文件"
        if (count > 1) {
            conflictFileNameEl.textContent = `${fileName} 及其他 ${count - 1} 个文件`;
        } else {
            conflictFileNameEl.textContent = fileName;
        }
        // 顺便把标题里的"已存在"提示更明确一点
        conflictModal.classList.add('visible');
    } catch(e) {
        // 出错时回传 skip，避免 C# 永久等待
        window.postAction('conflict_resolve', { choice: 'skip' });
    }
};

conflictModal.querySelectorAll('.conflict-actions button').forEach(b => {
    b.addEventListener('click', function() {
        const action = this.dataset.action;
        conflictModal.classList.remove('visible');
        window.postAction('conflict_resolve', { choice: action });
    });
});
conflictClose.addEventListener('click', () => {
    conflictModal.classList.remove('visible');
    window.postAction('conflict_resolve', { choice: 'skip' });
});
conflictModal.addEventListener('click', e => {
    if (e.target === conflictModal) {
        conflictModal.classList.remove('visible');
        window.postAction('conflict_resolve', { choice: 'skip' });
    }
});

/* ============================================================
   核心转换入口：btn-convert → showConfirm → 真实 C# 转换
   ============================================================ */
document.querySelectorAll('.btn-convert').forEach(b => {
    b.addEventListener('click', function() {
        const pageId = this.dataset.page;
        const st = getState(pageId);
        if (!st.files || !st.files.length) return;
        showConfirm(pageId);
    });
});

/// 收集转换参数并发送给 C#
function triggerConversion(pageId) {
    const st = getState(pageId);
    if (!st.files || !st.files.length) return;

    // 读取最新设置（包含 defaultOutput、conflict、pdf2wordLayout 等）
    const settings = loadSettings();

    // 收集转换选项
    let options = {};
    if (pageId === 'png2ico') {
        let sizes = [];
        document.querySelectorAll(`.preview-area[data-page="${pageId}"] .size-chip.active input`).forEach(cb => {
            sizes.push(parseInt(cb.value));
        });
        // 当前页面没选尺寸时，用设置里的默认尺寸
        if (!sizes.length && settings.defaultIcoSizes.length) {
            sizes = settings.defaultIcoSizes.slice();
        }
        if (!sizes.length) sizes.push(32);
        options = { sizes: sizes };
    }
    if (pageId === 'pdf2word') options.layout = settings.pdf2wordLayout;
    if (pageId === 'pdf2excel') options.mode = settings.pdf2excelMode;
    if (pageId === 'pdf2ppt') {
        // AI 生成 PPT：附带选中的模板 id（C# 端用 PptTemplates.Get 解析，缺省回退默认）
        options.template = _pdf2pptTemplateId || 'cyber_blue';
    }

    // 发送 C# 真实转换请求
    const payload = {
        type: pageId,
        files: st.files.map(f => ({ name: f.name, path: f._realPath || f.path || '' })),
        options: options,
        conflict: settings.conflict
    };
    // 如果用户设了默认输出目录，传给 C#；否则 C# 用 GetDefaultOutputDir()
    if (settings.defaultOutput) {
        payload.outputDir = settings.defaultOutput;
    }

    window.postAction('start_convert', payload);
}

/* ============================================================
   ICO 尺寸预览
   ============================================================ */
document.querySelectorAll('.preview-area .size-chip').forEach(chip => {
    chip.addEventListener('click', function() {
        const cb = this.querySelector('input');
        cb.checked = !cb.checked;
        this.classList.toggle('active');
        const pa = this.closest('.preview-area');
        if (pa) updateIcoPreview(pa.dataset.page);
    });
});

function updateIcoPreview(pageId) {
    const ip = q(`.ico-preview-sizes[data-page="${pageId}"]`);
    if (!ip) return;
    const checked = qs(q(`.preview-area[data-page="${pageId}"]`),`.size-chip.active`);
    const st = getState(pageId);
    const count = document.querySelectorAll(`.preview-area[data-page="${pageId}"] .size-chip.active`).length;
    ip.classList.toggle('visible', count > 0 && st.files.length > 0);
    ip.querySelectorAll('.ico-size-box').forEach(box => box.style.display = 'none');
    document.querySelectorAll(`.preview-area[data-page="${pageId}"] .size-chip.active`).forEach(chip => {
        const v = chip.querySelector('input')?.value;
        ip.querySelectorAll('.ico-size-box').forEach(box => {
            if (box.querySelector('.lbl')?.textContent.startsWith(v+'×')) box.style.display = 'block';
        });
    });
}

/* ============================================================
   Done-section actions
   ============================================================ */
document.querySelectorAll('.btn-open-folder').forEach(b => {
    b.addEventListener('click', function() {
        window.postAction('open_folder');
    });
});
document.querySelectorAll('.btn-reconvert').forEach(b => {
    b.addEventListener('click', function() {
        const ds = this.closest('.done-section');
        const pa = ds?.closest('.preview-area');
        if (pa) {
            qs(pa,'.done-section')?.classList.remove('visible');
            // Reset file statuses
            const fl = qs(pa,'.file-list');
            if (fl) {
                fl.querySelectorAll('.file-list-item .status').forEach(s => {
                    s.textContent = '等待转换';
                    s.className = 'status';
                });
                fl.querySelectorAll('.file-list-item .name .err').forEach(e => e.remove());
            }
            const pageId = pa.dataset.page;
            const st = getState(pageId);
            st.isConverting = false; st.cancelled = false;
            if (pageId === 'png2ico') updateIcoPreview(pageId);
            q(`.btn-convert[data-page="${pageId}"]`).disabled = false;
            showToast('已重置，可以重新转换', 'info');
        }
    });
});

/* ============================================================
   文档预览功能
   ============================================================ */

// 文档类型预览内容生成
const DOC_PREVIEW_DATA = {
    // ===== 共享内容模板 =====
    _content: {
        // PDF→Word / Word→PDF 共享文本
        report: {
            title: '2025 年度项目总结报告',
            author: '技术研发部',
            date: '2025 年 6 月',
            sections: [
                { heading: '一、项目背景', paragraphs: [
                    '随着企业数字化转型的深入推进，各类文档格式转换已成为日常办公中的高频需求。现有解决方案普遍存在转换质量低、批量处理能力弱、格式兼容性差等问题，亟需一套高效稳定的转换工具来提升工作效率。',
                    'ConvertPro 项目于 2025 年 1 月正式立项，目标为开发一款支持多种文档与图片格式互转的桌面应用程序，重点解决 PDF 与 Office 文档之间的高质量转换问题。',
                    '经过为期五个月的需求分析、架构设计、开发测试，项目已顺利完成第一阶段目标，核心转换功能均已实现并达到预期性能指标。'
                ]},
                { heading: '二、技术架构', paragraphs: [
                    '本系统采用经典的分层架构设计，自底向上依次为：数据访问层（DAL）、转换引擎层（Engine）、服务封装层（Service）和用户界面层（UI）。各层之间通过依赖注入（DI）方式解耦，便于独立开发和单元测试。',
                    '转换引擎层为核心模块，集成了 PdfToDocEngine、DocToPdfEngine、ImageToIcoEngine 等多个专用引擎，每个引擎独立运行于后台线程池中，支持并发处理和多任务调度。',
                    '性能测试结果表明，单文件平均转换耗时为 2.3 秒，十文件批量转换总耗时约 18.5 秒，内存占用峰值控制在 256MB 以内，CPU 利用率约为 45%。'
                ]},
                { heading: '三、功能模块', paragraphs: [
                    '图片转换模块：支持 PNG 格式导出为 Windows ICO 图标文件，自动嵌入 16×16、32×32、48×48 等多种尺寸，满足不同分辨率显示需求。',
                    '文档转换模块：涵盖 PDF 与 Word、Excel、PPT 之间的双向转换。PDF 转 Word 采用版面分析算法精准还原段落结构；Word 转 PDF 保留书签和超链接；PDF 转 Excel 自动识别并提取表格数据。',
                    '批量处理模块：支持多文件同时导入，提供统一的进度管理和结果汇总界面。用户可在转换过程中随时查看每个文件的处理状态，并支持取消单个或全部任务。'
                ]},
                { heading: '四、质量保障', paragraphs: [
                    '为确保转换质量，系统内置了三层校验机制：第一层为格式完整性检查，确保输出文件可正常打开；第二层为数据一致性验证，对比源文件与输出文件的内容差异；第三层为视觉对比，通过截图方式比对排版还原度。',
                    '在 500 份测试文档的验证中，PDF 转 Word 的段落还原率达到 98.5%，表格识别准确率为 95.2%，图片保留率为 100%。Word 转 PDF 的排版一致性达到 99.1%。'
                ]},
                { heading: '五、总结与展望', paragraphs: [
                    'ConvertPro 第一阶段的开发已圆满完成，核心转换功能稳定可靠，用户界面简洁高效。后续版本计划加入以下特性：云端转换服务、OCR 图片文字识别、批量导出与打印、以及更多格式的支持如 HTML、Markdown 等。',
                    '感谢项目团队的共同努力，也感谢各部门在测试阶段提供的宝贵反馈。我们将持续优化产品，为用户提供更好的转换体验。'
                ]}
            ]
        },
        // PDF→Excel / Excel→PDF 共享数据
        finance: {
            title: '2025 年度营收数据报表',
            sheets: [
                {
                    name: '营收汇总',
                    headers: ['项目名称', 'Q1 营收', 'Q2 营收', 'Q3 营收', 'Q4 营收', '年度合计'],
                    rows: [
                        ['产品 A 系列', '¥128,000', '¥145,000', '¥162,000', '¥189,000', '¥624,000'],
                        ['产品 B 系列', '¥89,000', '¥93,000', '¥87,000', '¥102,000', '¥371,000'],
                        ['产品 C 系列', '¥45,000', '¥52,000', '¥58,000', '¥63,000', '¥218,000'],
                        ['增值服务 D', '¥32,000', '¥38,000', '¥41,000', '¥47,000', '¥158,000'],
                        ['技术咨询 E', '¥18,500', '¥22,300', '¥25,600', '¥30,200', '¥96,600'],
                        ['合 计', '¥312,500', '¥350,300', '¥373,600', '¥431,200', '¥1,467,600']
                    ]
                },
                {
                    name: '成本明细',
                    headers: ['费用类别', 'Q1', 'Q2', 'Q3', 'Q4', '年度合计'],
                    rows: [
                        ['人力成本', '¥210,000', '¥225,000', '¥240,000', '¥255,000', '¥930,000'],
                        ['设备成本', '¥85,000', '¥78,000', '¥82,000', '¥79,000', '¥324,000'],
                        ['运营成本', '¥45,000', '¥48,000', '¥52,000', '¥55,000', '¥200,000'],
                        ['营销费用', '¥32,000', '¥38,000', '¥41,000', '¥47,000', '¥158,000'],
                        ['合 计', '¥372,000', '¥389,000', '¥415,000', '¥436,000', '¥1,612,000']
                    ]
                }
            ]
        },
        // PDF→PPT 共享内容
        presentation: {
            title: '项目汇报演示',
            slides: [
                '项目汇报 — ConvertPro 开发总结',
                '目录 — 汇报内容概览',
                '项目背景 — 需求来源与立项依据',
                '市场分析 — 同类产品对比与定位',
                '技术方案 — 核心技术选型',
                '系统架构 — 分层架构设计图',
                '实施计划 — 各阶段时间线',
                '资源预算 — 人力与设备投入',
                '风险评估 — 潜在风险与应对',
                '里程碑 — 关键节点完成情况',
                '团队分工 — 各成员职责',
                '进度跟踪 — 燃尽图与完成率',
                '质量保障 — 测试覆盖率与缺陷率',
                '总结展望 — 成果与后续规划',
                '附录 — 参考资料与数据来源'
            ]
        }
    },

    // ===== 各页面预览 =====
    pdf2word: {
        sourceContent: function() {
            const c = DOC_PREVIEW_DATA._content.report;
            const pages = [
                { num: 1, lines: ['封面', '2025 年度项目总结报告', '技术研发部', '2025年6月'] },
                { num: 2, lines: ['目录', '一、项目背景 … 1', '二、技术架构 … 3', '三、功能模块 … 5', '四、质量保障 … 8', '五、总结与展望 … 10'] },
                { num: 3, lines: ['一、项目背景', '随着企业数字化转型的深入推进...', 'ConvertPro 项目于 2025 年 1 月...', '经过为期五个月的需求分析...'] },
                { num: 4, lines: ['一、项目背景（续）', '经过调研，80% 的企业用户...', '市场上现有工具普遍存在...'] },
                { num: 5, lines: ['二、技术架构', '本系统采用经典的分层架构...', '转换引擎层为核心模块...', '性能测试结果表明...'] },
                { num: 6, lines: ['三、功能模块', '图片转换模块：支持 PNG→ICO...', '文档转换模块：PDF↔Word...', '批量处理模块：多文件同时导入...'] },
                { num: 7, lines: ['三、功能模块（续）', 'PDF 转 Excel 自动识别表格...', 'Word 转 PDF 保留书签超链接...'] },
                { num: 8, lines: ['四、质量保障', '三层校验机制：格式完整性...', '在 500 份测试文档中验证...', '段落还原率达到 98.5%...'] },
                { num: 9, lines: ['四、质量保障（续）', '表格识别准确率 95.2%...', '图片保留率 100%...'] },
                { num: 10, lines: ['五、总结与展望', '第一阶段目标圆满完成...', '后续计划：云端转换、OCR...'] }
            ];
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#888;">📄 源 PDF — 共 ${pages.length} 页</div><div class="doc-page-thumb" style="gap:8px;">`;
            pages.forEach(p => {
                h += `<div class="page"><div class="thumb">`;
                h += `<div style="font-size:8px;color:#999;margin-bottom:3px;border-bottom:1px solid #eee;padding-bottom:2px;">第 ${p.num} 页</div>`;
                p.lines.forEach(line => {
                    const w = Math.min(90, 50 + line.length * 3);
                    h += `<div class="line" style="width:${w}%;height:4px;"></div>`;
                });
                h += `</div><div class="lbl" style="font-size:9px;">P.${p.num}</div></div>`;
            });
            return h + '</div>';
        },
        outputContent: function() {
            const c = DOC_PREVIEW_DATA._content.report;
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#059669;">✅ 转换结果 — Word 文档（与源 PDF 内容完全一致，可编辑）</div>`;
            h += `<div class="doc-text-preview">`;
            h += `<div class="h">${c.title}</div>`;
            h += `<div class="p" style="color:#888;font-size:11px;margin-bottom:8px;">作者：${c.author} ｜ 日期：${c.date} ｜ 共 ${c.sections.reduce((a,s) => a + s.paragraphs.length, 0)} 段</div>`;
            c.sections.forEach(s => {
                h += `<div class="h2">${s.heading}</div>`;
                s.paragraphs.forEach(p => {
                    h += `<div class="p">${p}</div>`;
                });
            });
            return h + '</div>';
        }
    },

    word2pdf: {
        sourceContent: function() {
            const c = DOC_PREVIEW_DATA._content.report;
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#888;">📄 源 Word — ${c.title}</div>`;
            h += `<div class="doc-text-preview">`;
            h += `<div class="h">${c.title}</div>`;
            h += `<div class="p" style="color:#888;font-size:11px;margin-bottom:8px;">作者：${c.author} ｜ 日期：${c.date}</div>`;
            c.sections.forEach(s => {
                h += `<div class="h2">${s.heading}</div>`;
                h += `<div class="p">${s.paragraphs[0]}</div>`;
                h += `<div class="p" style="color:#aaa;">${s.paragraphs.length > 1 ? '…（以下共 ' + s.paragraphs.length + ' 段）' : ''}</div>`;
            });
            return h + '</div>';
        },
        outputContent: function() {
            const c = DOC_PREVIEW_DATA._content.report;
            const totalParagraphs = c.sections.reduce((a,s) => a + s.paragraphs.length, 0);
            // 估算页数：每页约 3 段
            const totalPages = Math.max(6, Math.ceil(totalParagraphs / 2.5));
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#059669;">✅ 转换结果 — PDF 文档（${totalPages} 页，与源 Word 内容完全一致）</div><div class="doc-page-thumb" style="gap:8px;">`;
            let paraIdx = 0;
            for (let p = 1; p <= totalPages; p++) {
                h += `<div class="page"><div class="thumb">`;
                h += `<div style="font-size:7px;color:#999;margin-bottom:2px;border-bottom:1px solid #eee;padding-bottom:2px;text-align:left;">第 ${p} 页</div>`;
                // Get content for this page
                let pageText = '';
                for (let s of c.sections) {
                    for (let para of s.paragraphs) {
                        if (paraIdx === p - 1) {
                            pageText = para.substring(0, 30) + '…';
                        }
                        paraIdx++;
                    }
                }
                paraIdx = 0;
                // Find the section title for this page
                let sectionTitle = '';
                let lineIdx = 0;
                for (let s of c.sections) {
                    if (lineIdx === p - 1) sectionTitle = s.heading;
                    lineIdx += s.paragraphs.length;
                }
                if (sectionTitle) {
                    h += `<div style="font-size:7px;color:#4f46e5;text-align:left;padding:1px 2px;font-weight:600;">${sectionTitle}</div>`;
                }
                if (pageText) {
                    h += `<div style="font-size:7px;color:#666;text-align:left;padding:1px 2px;">${pageText}</div>`;
                }
                // Simulate text lines
                const lineCount = 2 + (p % 3);
                for (let i = 0; i < lineCount; i++) {
                    h += `<div class="line" style="width:${60 + Math.random()*30}%;height:3px;margin:2px auto;"></div>`;
                }
                h += `<div style="font-size:7px;color:#ddd;text-align:center;margin-top:4px;">— ${p} —</div>`;
                h += `</div></div>`;
            }
            return h + '</div>';
        }
    },

    pdf2excel: {
        sourceContent: function() {
            const fin = DOC_PREVIEW_DATA._content.finance;
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#888;">📄 源 PDF — 检测到 ${fin.sheets.length} 个数据表</div>`;
            fin.sheets.forEach((sheet, si) => {
                if (si > 0) h += `<div style="height:8px;"></div>`;
                h += `<div style="font-size:11px;color:#555;margin-bottom:4px;font-weight:600;">📋 表格 ${si+1}：${sheet.name}</div>`;
                h += `<table class="doc-table-preview"><thead><tr>${sheet.headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>`;
                sheet.rows.forEach(row => {
                    h += `<tr>${row.map(c => `<td>${c}</td>`).join('')}</tr>`;
                });
                h += `</tbody></table>`;
            });
            return h;
        },
        outputContent: function() {
            const fin = DOC_PREVIEW_DATA._content.finance;
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#059669;">✅ 转换结果 — Excel 工作簿（${fin.sheets.length} 个工作表，数据与源 PDF 完全一致）</div>`;
            fin.sheets.forEach((sheet, si) => {
                if (si > 0) h += `<div style="height:10px;"></div>`;
                h += `<div style="font-size:11px;color:#4f46e5;margin-bottom:4px;font-weight:600;">📊 Sheet${si+1}：${sheet.name}</div>`;
                h += `<table class="doc-table-preview"><thead><tr>${sheet.headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>`;
                sheet.rows.forEach((row, ri) => {
                    const isTotal = row[0].includes('合 计');
                    h += `<tr>${row.map((c, ci) => {
                        const val = isTotal ? `<strong>${c}</strong>` : c;
                        return ci === 0 ? `<td><strong>${c.replace(/<\/?strong>/g,'')}</strong></td>` : `<td>${val}</td>`;
                    }).join('')}</tr>`;
                });
                h += `</tbody></table>`;
            });
            return h;
        }
    },

    excel2pdf: {
        sourceContent: function() {
            const fin = DOC_PREVIEW_DATA._content.finance;
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#888;">📄 源 Excel — 共 ${fin.sheets.length} 个工作表</div>`;
            fin.sheets.forEach((sheet, si) => {
                if (si > 0) h += `<div style="height:8px;"></div>`;
                h += `<div style="font-size:11px;color:#555;margin-bottom:4px;font-weight:600;">📊 Sheet${si+1}：${sheet.name}</div>`;
                h += `<table class="doc-table-preview"><thead><tr>${sheet.headers.map(h => `<th>${h}</th>`).join('')}</tr></thead><tbody>`;
                sheet.rows.forEach(row => {
                    h += `<tr>${row.map(c => `<td>${c}</td>`).join('')}</tr>`;
                });
                h += `</tbody></table>`;
            });
            return h;
        },
        outputContent: function() {
            const fin = DOC_PREVIEW_DATA._content.finance;
            const totalPages = 8;
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#059669;">✅ 转换结果 — PDF 文档（${totalPages} 页，数据与源 Excel 完全一致）</div><div class="doc-page-thumb" style="gap:8px;">`;
            const pages = [
                '封面 — 2025 年度营收数据报表',
                'Sheet1：营收汇总 — 表头与数据',
                'Sheet1：营收汇总 — 各产品明细',
                'Sheet1：营收汇总 — 合计行',
                'Sheet2：成本明细 — 表头与数据',
                'Sheet2：成本明细 — 各项费用',
                'Sheet2：成本明细 — 合计行',
                '附录 — 数据来源与说明'
            ];
            for (let i = 1; i <= totalPages; i++) {
                h += `<div class="page"><div class="thumb">`;
                h += `<div style="font-size:7px;color:#999;margin-bottom:2px;border-bottom:1px solid #eee;padding-bottom:2px;text-align:left;">第 ${i} 页</div>`;
                h += `<div style="font-size:7px;color:#4f46e5;text-align:left;padding:2px;font-weight:600;">${pages[i-1]}</div>`;
                // Show mini table bars
                if (i >= 2 && i <= 7) {
                    for (let r = 0; r < 4; r++) {
                        h += `<div style="display:flex;gap:2px;padding:1px 2px;"><div style="width:20px;height:3px;background:#ddd;border-radius:1px;"></div><div style="flex:1;height:3px;background:#e8e8e8;border-radius:1px;"></div></div>`;
                    }
                }
                h += `<div style="font-size:7px;color:#ddd;text-align:center;margin-top:4px;">— ${i} —</div>`;
                h += `</div></div>`;
            }
            return h + '</div>';
        }
    },

    pdf2ppt: {
        sourceContent: function() {
            const slides = DOC_PREVIEW_DATA._content.presentation.slides;
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#888;">📄 源 PDF — 共 ${slides.length} 页幻灯片内容</div><div class="doc-slide-preview">`;
            slides.forEach((title, i) => {
                h += `<div class="slide"><div class="thumb"><span class="num">${i+1}</span>${title}</div><div class="lbl">第 ${i+1} 页</div></div>`;
            });
            return h + '</div>';
        },
        outputContent: function() {
            const slides = DOC_PREVIEW_DATA._content.presentation.slides;
            const colors = ['#4f46e5','#059669','#d97706','#dc2626','#2563eb','#7c3aed','#db2777','#0891b2'];
            let h = `<div style="margin-bottom:8px;font-size:12px;color:#059669;">✅ 转换结果 — PPT 演示文稿（${slides.length} 张幻灯片，内容与源 PDF 完全一致）</div><div class="doc-slide-preview">`;
            slides.forEach((title, i) => {
                const bg = colors[i % colors.length];
                h += `<div class="slide"><div class="thumb" style="background:${bg}12;color:${bg};font-weight:600;font-size:11px;">${title.length > 12 ? title.substring(0,10)+'…' : title}</div><div class="lbl" style="color:${bg};">幻灯片 ${i+1}</div></div>`;
            });
            return h + '</div>';
        }
    }
};

// 填充文档预览
function populateDocPreview(pageId) {
    const data = DOC_PREVIEW_DATA[pageId];
    if (!data) return;

    // 获取文件数量作为预估页数
    const st = getState(pageId);
    const fileCount = st.files ? st.files.length : 1;

    // 源文件预览
    if (data.sourceContent) {
        const sourceEl = document.getElementById('pane-' + pageId + '-source');
        if (sourceEl) {
            const thumbContainer = sourceEl.querySelector('.doc-page-thumb');
            const textContainer = sourceEl.querySelector('.doc-text-preview');
            const tableContainer = sourceEl.querySelector('.doc-table-preview-wrapper');
            const slideContainer = sourceEl.querySelector('.doc-slide-preview');

            let content;
            // 根据容器类型调用 sourceContent，传递页数
            if (thumbContainer || slideContainer) {
                content = data.sourceContent(data.pages || 12);
            } else {
                content = data.sourceContent();
            }
            if (thumbContainer) thumbContainer.innerHTML = content;
            if (textContainer) textContainer.innerHTML = data.sourceContent();
            if (tableContainer) tableContainer.innerHTML = data.sourceContent();
            if (slideContainer) slideContainer.innerHTML = data.sourceContent();
        }
    }
}

// 填充输出预览
function populateOutputPreview(pageId) {
    const data = DOC_PREVIEW_DATA[pageId];
    if (!data || !data.outputContent) return;

    const outputEl = document.getElementById('pane-' + pageId + '-output');
    if (!outputEl) return;

    const thumbContainer = outputEl.querySelector('.doc-page-thumb');
    const textContainer = outputEl.querySelector('.doc-text-preview');
    const tableContainer = outputEl.querySelector('.doc-table-preview-wrapper');
    const slideContainer = outputEl.querySelector('.doc-slide-preview');

    if (thumbContainer) thumbContainer.innerHTML = data.outputContent();
    if (textContainer) textContainer.innerHTML = data.outputContent();
    if (tableContainer) tableContainer.innerHTML = data.outputContent();
    if (slideContainer) slideContainer.innerHTML = data.outputContent();

    // Enable output tab
    const docPreview = outputEl.closest('.doc-preview');
    if (docPreview) {
        const outputTab = docPreview.querySelector('.tab[data-pane="output"]');
        if (outputTab) {
            outputTab.disabled = false;
            outputTab.querySelector('.count').textContent = '可查看';
        }
    }
}

// Tab 切换
document.addEventListener('click', function(e) {
    const tab = e.target.closest('.doc-preview-tabs .tab');
    if (!tab) return;
    if (tab.disabled) return;

    const tabsContainer = tab.closest('.doc-preview-tabs');
    const body = tabsContainer.nextElementSibling;
    if (!body) return;

    tabsContainer.querySelectorAll('.tab').forEach(t => t.classList.remove('active'));
    tab.classList.add('active');

    const paneId = tab.dataset.pane;
    const docPreview = tab.closest('.doc-preview');
    if (docPreview) {
        docPreview.querySelectorAll('.preview-pane').forEach(p => p.classList.remove('active'));
        const targetPane = docPreview.querySelector(`.preview-pane[id$="-${paneId}"]`);
        if (targetPane) targetPane.classList.add('active');
    }
});


/* ============================================================
   1️⃣0️⃣ 键盘快捷键
   ============================================================ */
document.addEventListener('keydown', e => {
    if (!document.getElementById('setting-shortcuts').checked) return;
    if (e.key === '?' && !e.ctrlKey && !e.metaKey) {
        const h = document.getElementById('shortcutsHint');
        h.classList.toggle('visible');
        if (h.classList.contains('visible')) setTimeout(() => h.classList.remove('visible'), 4000);
        return;
    }
    if (e.ctrlKey || e.metaKey) {
        if (e.key === 'o' && !e.shiftKey) {
            e.preventDefault();
            const activePage = document.querySelector('.page.active');
            if (activePage) {
                const input = activePage.querySelector('.file-input');
                if (input) input.click();
            }
        } else if (e.key === 'O' && e.shiftKey) {
            e.preventDefault();
            const activePage = document.querySelector('.page.active');
            if (activePage) {
                const browse = activePage.querySelector('.btn-browse');
                if (browse) browse.click();
            }
        } else if (e.key === 'Enter') {
            e.preventDefault();
            const activePage = document.querySelector('.page.active');
            if (activePage) {
                const conv = activePage.querySelector('.btn-convert');
                if (conv && !conv.disabled) conv.click();
            }
        } else if (e.key === 'i') {
            e.preventDefault();
            document.getElementById('btnOpenAbout').click();
        } else if (e.key === ',') {
            e.preventDefault();
            const hBtn = document.querySelector('.nav-item[data-page="settings"]');
            if (hBtn) hBtn.click();
        } else if (e.key === 'h') {
            e.preventDefault();
            const hBtn = document.querySelector('.nav-item[data-page="history"]');
            if (hBtn) hBtn.click();
        }
    }
});

/* ============================================================
   设置页面 — 真正持久化到 localStorage
   ============================================================ */
const SETTINGS_KEY = 'convertpro_settings';

// 默认值（outputDir 留空表示使用 C# 默认值，由 C# ConversionManager.GetDefaultOutputDir 返回桌面 ConvertPro_Output）
const DEFAULT_SETTINGS = {
    defaultOutput: '',          // 空字符串 → 使用 C# 默认值
    conflict: 'ask',           // ask | overwrite | rename | skip
    pdf2wordLayout: 'flow',    // flow | fixed
    pdf2excelMode: 'table',    // table | all
    preconfirm: true,
    shortcuts: true,
    defaultIcoSizes: [16, 32, 48]
};

function loadSettings() {
    try {
        const raw = localStorage?.getItem(SETTINGS_KEY);
        if (!raw) return { ...DEFAULT_SETTINGS };
        return { ...DEFAULT_SETTINGS, ...JSON.parse(raw) };
    } catch (e) {
        return { ...DEFAULT_SETTINGS };
    }
}

function saveSettings(settings) {
    try {
        localStorage?.setItem(SETTINGS_KEY, JSON.stringify(settings));
    } catch (e) {
        console.warn('设置保存失败:', e);
    }
}

function applySettingsToUI(settings) {
    document.getElementById('setting-defaultOutput').value = settings.defaultOutput || '';
    document.getElementById('setting-conflict').value = settings.conflict;
    document.getElementById('setting-pdf2word-layout').value = settings.pdf2wordLayout;
    document.getElementById('setting-pdf2excel-mode').value = settings.pdf2excelMode;
    document.getElementById('setting-preconfirm').checked = settings.preconfirm;
    document.getElementById('setting-shortcuts').checked = settings.shortcuts;

    // ICO 尺寸 chip
    document.querySelectorAll('#page-settings .size-chip').forEach(chip => {
        const v = chip.querySelector('input')?.value;
        if (settings.defaultIcoSizes.includes(parseInt(v))) {
            chip.classList.add('active');
            chip.querySelector('input').checked = true;
        } else {
            chip.classList.remove('active');
            chip.querySelector('input').checked = false;
        }
    });
}

function collectSettingsFromUI() {
    const defaultIcoSizes = [];
    document.querySelectorAll('#page-settings .size-chip.active input').forEach(input => {
        defaultIcoSizes.push(parseInt(input.value));
    });
    return {
        defaultOutput: document.getElementById('setting-defaultOutput').value.trim(),
        conflict: document.getElementById('setting-conflict').value,
        pdf2wordLayout: document.getElementById('setting-pdf2word-layout').value,
        pdf2excelMode: document.getElementById('setting-pdf2excel-mode').value,
        preconfirm: document.getElementById('setting-preconfirm').checked,
        shortcuts: document.getElementById('setting-shortcuts').checked,
        defaultIcoSizes: defaultIcoSizes
    };
}

// 启动时加载设置并应用到 UI
const appSettings = loadSettings();
applySettingsToUI(appSettings);

// 启动时如果用户没设过默认输出目录，向 C# 拿默认值填进去
if (!appSettings.defaultOutput) {
    window.postAction('get_output_dir');
}

document.getElementById('settingBrowseOutput').addEventListener('click', () => {
    const p = prompt('请输入默认输出目录路径：', document.getElementById('setting-defaultOutput').value);
    if (p && p.trim()) {
        document.getElementById('setting-defaultOutput').value = p.trim();
        // 同步给各页面输出路径显示
        document.querySelectorAll('[id^="outputPath-"]').forEach(el => el.textContent = p.trim());
        saveSettings(collectSettingsFromUI());
        showToast('默认输出目录已更新', 'success');
    }
});

document.getElementById('btnResetSettings').addEventListener('click', () => {
    if (!confirm('确定恢复所有设置为默认值吗？')) return;
    applySettingsToUI({ ...DEFAULT_SETTINGS });
    saveSettings({ ...DEFAULT_SETTINGS });
    // 重置后向 C# 拿默认输出目录填回去
    window.postAction('get_output_dir');
    showToast('设置已恢复默认', 'success');
});

// 控件变更 → 实际保存到 localStorage（修复原代码只显示 toast 不保存的问题）
// 注意排除 AI 区控件（id 以 setting-ai- 开头），它们走独立的 C# 同步流程
document.querySelectorAll('#page-settings input, #page-settings select').forEach(el => {
    if (el.id && el.id.indexOf('setting-ai-') === 0) return;
    el.addEventListener('change', () => {
        saveSettings(collectSettingsFromUI());
        showToast('设置已保存', 'info');
    });
});

// 设置页 ICO chip 点击事件（与 preview-area 里的分开绑定）
document.querySelectorAll('#page-settings .size-chip').forEach(chip => {
    chip.addEventListener('click', function() {
        const cb = this.querySelector('input');
        cb.checked = !cb.checked;
        this.classList.toggle('active');
        saveSettings(collectSettingsFromUI());
    });
});

/* ============================================================
   AI 提供商配置（设置页）
   协议：JS 通过 postAction 请求/修改 → C# 处理后回推 __onAiStatus
   ============================================================ */
const aiProviderSel = document.getElementById('setting-ai-provider');
const aiModelSel = document.getElementById('setting-ai-model');
const aiKeyInput = document.getElementById('setting-ai-key');
const aiSaveKeyBtn = document.getElementById('aiSaveKey');
const aiTestBtn = document.getElementById('aiTestBtn');
const aiTestResult = document.getElementById('aiTestResult');
const aiBadge = document.getElementById('aiStatusBadge');

let _aiStatus = null; // 最近一次 C# 回推的状态快照

// C# → JS：推送 AI 状态，渲染整个 AI 区
window.__onAiStatus = function(status) {
    if (!status) return;
    _aiStatus = status;

    // 1. 状态徽章
    const cur = status.providers?.find(p => p.name === status.current);
    if (status.anyAvailable && cur) {
        aiBadge.className = 'ai-badge ai-badge-ok';
        aiBadge.textContent = `已就绪 · ${cur.displayName}`;
    } else {
        aiBadge.className = 'ai-badge ai-badge-warn';
        aiBadge.textContent = '未配置 API Key';
    }

    // 2. 提供商下拉
    const prevProvider = aiProviderSel.value;
    aiProviderSel.innerHTML = '';
    (status.providers || []).forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.name;
        opt.textContent = p.displayName + (p.isAvailable ? '' : '（未配置）');
        aiProviderSel.appendChild(opt);
    });
    aiProviderSel.value = status.current;

    // 3. 当前 provider 的模型下拉
    renderModelOptions(cur);

    // 4. Key 输入框：若手动 key 已设则显示占位提示（不回显明文，安全考虑）
    if (cur && cur.hasManualKey) {
        aiKeyInput.value = '';
        aiKeyInput.placeholder = '已保存（重新输入可覆盖）';
    } else if (cur && cur.isAvailable && !cur.hasManualKey) {
        aiKeyInput.value = '';
        aiKeyInput.placeholder = '来自环境变量 API_KEY';
    } else {
        aiKeyInput.value = '';
        aiKeyInput.placeholder = 'sk-...';
    }

    // 5. 同步刷新 PDF→PPT 页的 AI 引擎徽章
    updatePdf2pptAiStatus(status, cur);
};

function updatePdf2pptAiStatus(status, cur) {
    // 同步刷新 PDF→PPT 和 Word→PPT 两个页面的 AI 引擎徽章
    const ids = ['pdf2pptAiStatus', 'word2pptAiStatus'];
    ids.forEach(id => {
        const el = document.getElementById(id);
        if (!el) return;
        if (status.anyAvailable && cur) {
            el.className = 'ai-badge ai-badge-ok';
            el.textContent = `${cur.displayName} · 已就绪`;
        } else {
            el.className = 'ai-badge ai-badge-warn';
            el.textContent = 'AI 未配置（去设置页填写 Key）';
        }
    });
}

function renderModelOptions(provider) {
    aiModelSel.innerHTML = '';
    if (!provider) return;
    (provider.models || []).forEach(m => {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = m + (m === provider.defaultModel ? '（默认）' : '');
        aiModelSel.appendChild(opt);
    });
    aiModelSel.value = provider.currentModel || provider.defaultModel;
}

// 切换提供商 → 通知 C#，C# 会回推新状态（含新 provider 的模型列表）
aiProviderSel.addEventListener('change', () => {
    window.postAction('ai_set_provider', { provider: aiProviderSel.value });
});

// 切换模型 → 通知 C#
aiModelSel.addEventListener('change', () => {
    window.postAction('ai_set_model', {
        provider: aiProviderSel.value,
        model: aiModelSel.value
    });
});

// 保存 API Key → 通知 C#（C# 持久化后回推状态）
aiSaveKeyBtn.addEventListener('click', () => {
    const key = aiKeyInput.value.trim();
    window.postAction('ai_set_key', {
        provider: aiProviderSel.value,
        key: key
    });
    aiKeyInput.value = '';
    showToast('API Key 已保存', 'success');
});

// 测试连接
aiTestBtn.addEventListener('click', () => {
    window.postAction('ai_test', { provider: aiProviderSel.value });
});
window.__onAiTest = function(info) {
    if (!info) return;
    if (info.state === 'testing') {
        aiTestResult.className = 'hint testing';
        aiTestResult.textContent = '测试中…';
        aiTestBtn.disabled = true;
        return;
    }
    aiTestBtn.disabled = false;
    const msg = info.message || '';
    // 根据返回文本前缀判断成功/失败（C# 端用"连接成功"/"连接失败"/"未配置"开头）
    if (msg.indexOf('成功') !== -1) {
        aiTestResult.className = 'hint ok';
    } else {
        aiTestResult.className = 'hint fail';
    }
    aiTestResult.textContent = msg;
};

// 启动时主动拉一次 AI 状态（C# 在 NavigationCompleted 也会主动推，这里兜底）
window.postAction('ai_get_status');

/* ============================================================
   文生图模型配置（设置页）— 与 AI 聊天提供商完全独立
   协议：JS 通过 postAction 请求/修改 → C# 处理后回推 __onImageGenStatus
   ============================================================ */
const imgProviderSel = document.getElementById('setting-img-provider');
const imgModelSel = document.getElementById('setting-img-model');
const imgKeyInput = document.getElementById('setting-img-key');
const imgSaveKeyBtn = document.getElementById('imgSaveKey');
const imgTestBtn = document.getElementById('imgTestBtn');
const imgTestResult = document.getElementById('imgTestResult');
const imgBadge = document.getElementById('imgStatusBadge');

let _imgStatus = null;

window.__onImageGenStatus = function(status) {
    if (!status) return;
    _imgStatus = status;

    const cur = status.providers?.find(p => p.name === status.current);
    if (status.anyAvailable && cur) {
        imgBadge.className = 'ai-badge ai-badge-ok';
        imgBadge.textContent = `已就绪 · ${cur.displayName}`;
    } else {
        imgBadge.className = 'ai-badge ai-badge-warn';
        imgBadge.textContent = '未配置 API Key';
    }

    const prevProvider = imgProviderSel.value;
    imgProviderSel.innerHTML = '';
    (status.providers || []).forEach(p => {
        const opt = document.createElement('option');
        opt.value = p.name;
        opt.textContent = p.displayName + (p.isAvailable ? '' : '（未配置）');
        imgProviderSel.appendChild(opt);
    });
    imgProviderSel.value = status.current;

    renderImgModelOptions(cur);

    if (cur && cur.hasManualKey) {
        imgKeyInput.value = '';
        imgKeyInput.placeholder = '已保存（重新输入可覆盖）';
    } else {
        imgKeyInput.value = '';
        imgKeyInput.placeholder = 'sk-...';
    }
};

function renderImgModelOptions(provider) {
    imgModelSel.innerHTML = '';
    if (!provider) return;
    (provider.models || []).forEach(m => {
        const opt = document.createElement('option');
        opt.value = m;
        opt.textContent = m + (m === provider.defaultModel ? '（默认）' : '');
        imgModelSel.appendChild(opt);
    });
    imgModelSel.value = provider.currentModel || provider.defaultModel;
}

imgProviderSel.addEventListener('change', () => {
    window.postAction('img_set_provider', { provider: imgProviderSel.value });
});

imgModelSel.addEventListener('change', () => {
    window.postAction('img_set_model', {
        provider: imgProviderSel.value,
        model: imgModelSel.value
    });
});

imgSaveKeyBtn.addEventListener('click', () => {
    const key = imgKeyInput.value.trim();
    window.postAction('img_set_key', {
        provider: imgProviderSel.value,
        key: key
    });
    imgKeyInput.value = '';
    showToast('文生图 API Key 已保存', 'success');
});

imgTestBtn.addEventListener('click', () => {
    window.postAction('img_test', { provider: imgProviderSel.value });
});
window.__onImageGenTest = function(info) {
    if (!info) return;
    if (info.state === 'testing') {
        imgTestResult.className = 'hint testing';
        imgTestResult.textContent = '生成中…（最多 2 分钟）';
        imgTestBtn.disabled = true;
        return;
    }
    imgTestBtn.disabled = false;
    const msg = info.message || '';
    if (msg.indexOf('成功') !== -1) {
        imgTestResult.className = 'hint ok';
    } else {
        imgTestResult.className = 'hint fail';
    }
    imgTestResult.textContent = msg;
};

window.postAction('img_get_status');

/* ============================================================
   PPT 模板选择（PDF→PPT 页）— 可搜索的风格库
   C# 推送 __onPptTemplates → JS 渲染分类 chips + 卡片
   搜索框 + 分类筛选 → 实时过滤；转换时附带 template id
   ============================================================ */
const tplGrid = document.getElementById('tplGrid-pdf2ppt');
const tplSearchInput = document.getElementById('tplSearch');
const tplCountEl = document.getElementById('tplCount');
const tplCategoriesEl = document.getElementById('tplCategories');

let _pptTemplates = [];          // 全量模板列表（C# 推送）
let _pdf2pptTemplateId = 'cyber_blue'; // 当前选中
let _tplActiveCategory = '全部';   // 当前分类筛选
let _tplKeyword = '';              // 当前搜索词

window.__onPptTemplates = function(payload) {
    if (!payload || !payload.templates || !tplGrid) return;
    _pptTemplates = payload.templates;
    if (payload.defaultId) _pdf2pptTemplateId = payload.defaultId;

    // 渲染分类 chips（去重，加"全部"）
    const cats = ['全部', ...Array.from(new Set(_pptTemplates.map(t => t.category)))];
    tplCategoriesEl.innerHTML = '';
    cats.forEach(c => {
        const chip = document.createElement('span');
        chip.className = 'tpl-chip' + (c === _tplActiveCategory ? ' active' : '');
        chip.textContent = c;
        chip.dataset.cat = c;
        chip.addEventListener('click', () => {
            _tplActiveCategory = c;
            tplCategoriesEl.querySelectorAll('.tpl-chip').forEach(x => x.classList.remove('active'));
            chip.classList.add('active');
            renderTemplateGrid();
        });
        tplCategoriesEl.appendChild(chip);
    });

    renderTemplateGrid();
};

function renderTemplateGrid() {
    if (!tplGrid) return;
    // 综合过滤：分类 + 搜索词（搜索词匹配 名称/分类/标签/主题）
    const kw = _tplKeyword.trim().toLowerCase();
    const filtered = _pptTemplates.filter(t => {
        if (_tplActiveCategory !== '全部' && t.category !== _tplActiveCategory) return false;
        if (!kw) return true;
        const hay = (t.name + ' ' + t.category + ' ' + (t.tags || '') + ' ' + t.theme).toLowerCase();
        // 支持"深色/浅色"中文搜索
        const themeCn = t.theme === 'dark' ? '深色' : '浅色';
        return hay.indexOf(kw) !== -1 || themeCn.indexOf(kw) !== -1;
    });

    tplGrid.innerHTML = '';
    if (filtered.length === 0) {
        tplGrid.innerHTML = '<div class="tpl-empty">没有匹配的模板，换个关键词试试</div>';
        tplCountEl.textContent = '0 套';
        return;
    }
    tplCountEl.textContent = filtered.length + ' 套';

    filtered.forEach(t => {
        const card = document.createElement('div');
        card.className = 'tpl-card' + (t.id === _pdf2pptTemplateId ? ' active' : '');
        card.dataset.id = t.id;
        // 文字色用 panel（深色模板里 panel 较亮；浅色模板里 panel 较暗），保证可读
        const nameColor = t.theme === 'dark' ? t.panel : '#1f2937';
        const themeTagCls = t.theme === 'dark' ? 'tpl-theme-dark' : 'tpl-theme-light';
        const themeTagText = t.theme === 'dark' ? '深' : '浅';
        card.innerHTML =
            '<div class="tpl-swatch" style="background:#' + t.bg + ';">' +
                '<div class="swatch-bar" style="background:#' + t.accent + ';"></div>' +
                '<div class="swatch-name" style="color:' + nameColor + ';">' + t.name + '</div>' +
            '</div>' +
            '<div class="tpl-foot">' +
                '<span>' + t.category + '</span>' +
                '<span style="display:inline-flex;align-items:center;gap:4px;">' +
                    '<i style="width:8px;height:8px;border-radius:2px;background:#' + t.accent + ';display:inline-block;"></i>' +
                    '<i style="width:8px;height:8px;border-radius:2px;background:#' + t.accent2 + ';display:inline-block;"></i>' +
                    '<span class="' + themeTagCls + '">' + themeTagText + '</span>' +
                '</span>' +
            '</div>';
        card.addEventListener('click', () => {
            _pdf2pptTemplateId = t.id;
            tplGrid.querySelectorAll('.tpl-card').forEach(c => c.classList.remove('active'));
            card.classList.add('active');
        });
        tplGrid.appendChild(card);
    });
}

// 搜索框实时过滤
if (tplSearchInput) {
    tplSearchInput.addEventListener('input', () => {
        _tplKeyword = tplSearchInput.value;
        renderTemplateGrid();
    });
}

// 兜底：启动时主动拉一次模板列表
window.postAction('ppt_get_templates');

console.log('ConvertPro 原型加载完成 — 包含全部 12 项增强功能');

// ===== C# 回调函数（WebView2 → JS）=====
// 由 C# OpenFileDialog 选中文件后调用
window.__onFilesSelected = function(files) {
    if (!files || !files.length) return;
    const pageId = document.querySelector('.page.active')?.id?.replace('page-', '');
    if (!pageId) return;
    const st = getState(pageId);
    st.files = files.map(f => ({ name: f.name, size: f.size, path: f.path, _realPath: f.path }));
    st.cancelled = false;
    const PA = q(`.preview-area[data-page="${pageId}"]`);
    if (!PA) return;
    PA.classList.add('visible');
    const first = files[0];
    qs(PA,'.preview-info .name').textContent = first.name;
    qs(PA,'.preview-info .meta').innerHTML = `<span>大小: ${(first.size/1024).toFixed(1)} KB</span>`;
    // File list
    const FL = qs(PA,`.file-list`);
    if (FL) {
        FL.innerHTML = files.map((f,i) =>
            `<div class="file-list-item" data-idx="${i}"><span class="idx">${i+1}</span><span class="name">${f.name}</span><span class="status">${(f.size/1024).toFixed(1)} KB</span></div>`
        ).join('');
    }
    q(`.drop-zone[data-page="${pageId}"]`).style.display = 'none';
    q(`.btn-convert[data-page="${pageId}"]`).disabled = false;
    q(`.btn-clear[data-page="${pageId}"]`).style.display = 'inline-block';
    q(`.page.active .status-text`).textContent = `已选择 ${files.length} 个文件`;
    statusHint.textContent = `已选择 ${files.length} 个文件`;
    qs(PA,`.done-section`)?.classList.remove('visible');
    qs(PA,`.progress-bar`)?.classList.remove('visible');
    if (pageId==='png2ico') updateIcoPreview(pageId);
    if (DOC_PREVIEW_DATA && DOC_PREVIEW_DATA[pageId]) {
        const dp = q(`.doc-preview[data-page="${pageId}"]`);
        if (dp) { dp.classList.add('visible'); populateDocPreview(pageId); }
    }
};

window.__onConvertStarted = function() {
    const pageId = document.querySelector('.page.active')?.id?.replace('page-', '');
    if (!pageId) return;
    const st = getState(pageId);
    st.isConverting = true;
    st.cancelled = false;
    const btn = q(`.btn-convert[data-page="${pageId}"]`);
    if (btn) { btn.disabled = true; btn.textContent = '转换中...'; }
    q(`.btn-clear[data-page="${pageId}"]`).style.display = 'none';
    q(`.btn-cancel[data-page="${pageId}"]`).style.display = 'inline-block';
    q(`.progress-bar[data-page="${pageId}"]`).classList.add('visible');
    q(`.done-section[data-page="${pageId}"]`).classList.remove('visible');
    q(`.page.active .status-text`).textContent = '正在转换...';
    statusHint.textContent = '正在转换...';
};

window.__onConvertProgress = function(p) {
    const pageId = document.querySelector('.page.active')?.id?.replace('page-', '');
    if (!pageId) return;
    const fill = q(`.progress-bar[data-page="${pageId}"] .progress-fill`);
    const label = q(`.progress-bar[data-page="${pageId}"] .progress-label span:first-child`);
    const pctEl = q(`.progress-bar[data-page="${pageId}"] .progress-label span:last-child`);
    if (fill) fill.style.width = p.percent + '%';
    if (label) label.textContent = p.current ? `文件 ${p.completed+1}/${p.total} 转换中...` : `处理中...`;
    if (pctEl) pctEl.textContent = Math.round(p.percent) + '%';
    // 更新文件列表状态
    const items = document.querySelectorAll(`.file-list[data-page="${pageId}"] .file-list-item`);
    items.forEach((item, i) => {
        const st = item.querySelector('.status');
        if (st && i < p.completed) { st.textContent = '✅'; st.className = 'status done'; }
    });
};

window.__onConvertComplete = function(results) {
    console.log('[ConvertPro] 转换结果:', JSON.stringify(results, null, 2));
    const pageId = document.querySelector('.page.active')?.id?.replace('page-', '');
    if (!pageId) return;
    const st = getState(pageId);
    st.isConverting = false;
    const btn = q(`.btn-convert[data-page="${pageId}"]`);
    if (btn) { btn.disabled = false; btn.textContent = '开始转换'; }
    q(`.btn-cancel[data-page="${pageId}"]`).style.display = 'none';
    q(`.btn-clear[data-page="${pageId}"]`).style.display = 'inline-block';
    q(`.progress-bar[data-page="${pageId}"]`).classList.remove('visible');
    const ds = q(`.done-section[data-page="${pageId}"]`);
    if (ds) ds.classList.add('visible');
    const hasErr = results.some(r => !r.success);
    q(`.page.active .status-text`).textContent = hasErr ? '部分完成 ⚠️' : '转换完成 ✅';
    statusHint.textContent = hasErr ? '部分文件转换失败' : '转换完成 ✅';
    // 输出文件列表 + 历史记录
    const ofDiv = ds ? ds.querySelector('.output-files') : null;
    if (ofDiv && results.length) {
        let h = '<div style="font-weight:600;color:#4f46e5;margin-bottom:4px;">已生成文件：</div>';
        results.forEach(r => {
            const status = r.success ? 'success' : 'error';
            const icon = r.success ? '📄' : '❌';
            const meta = r.success ? Math.round(r.size/1024)+' KB' : (r.error || '失败');
            h += `<div style="display:flex;align-items:center;gap:8px;padding:2px 0;font-size:12px;">
                <span style="color:${r.success?'#059669':'#ef4444'};">${icon}</span>
                <span style="flex:1;color:#333;">${r.name||'未知'}</span>
                <span style="color:#999;font-size:11px;">${meta}</span>
            </div>`;
            // 写入历史记录
            addHistory({ type: pageId, file: r.name || '未知', status: status });
        });
        ofDiv.innerHTML = h;
    }
    if (DOC_PREVIEW_DATA && DOC_PREVIEW_DATA[pageId]) populateOutputPreview(pageId);
    const errList = results.filter(r => !r.success).map(r => r.error || '未知错误');
    showToast(hasErr ? `转换失败: ${errList.join('; ')}` : '转换完成！', hasErr ? 'error' : 'success');
};

window.__onOutputDir = function(dir) {
    const pageId = document.querySelector('.page.active')?.id?.replace('page-', '');
    if (pageId) {
        const el = document.getElementById('outputPath-' + pageId);
        if (el) el.textContent = dir;
    }
};

// LibreOffice 状态通知
window.__onEngineStatus = function(engine) {
    if (engine === 'fallback') {
        console.warn('未检测到 Word 或 LibreOffice，文档转换使用基础模式');
    } else {
        console.log('转换引擎: ' + engine);
    }
};

})();
