/**
 * 智训页面 JS — 对话式训练
 */

// ==================== 状态 ====================
let _currentTenantId = null;
let _currentSessionId = null;
let _sessions = [];
let _chatMessages = [];
let _isStreaming = false;
let _currentAccuracy = null;
let _currentCode = null;
let _filePasswordsMap = {};
let _chatStreamEl = null;   // AI 对话流式输出的 DOM 元素
let _chatStreamBuf = '';    // AI 对话流式输出的文本缓冲
let _chatThinkingBuf = '';  // DeepSeek 推理模型思考过程缓冲（灰色区显示，不进入正式内容）
let _chatThinkingTimer = null;  // 思考渲染节流器（100ms 合并一次 DOM 更新，防 O(n²) 卡死页面）
let _pendingScriptName = null;  // 新建训练时由用户命名的脚本名（仅新建会话首次提交时使用）

// ===== 对话历史分页（B 方案：初始最近 5 轮，上划加载更早）=====
let _historyOldestId = null;        // 当前已加载最早消息的 page_start_id（上划游标）
let _historyHasMore = false;        // 是否还有更早的消息可加载
let _historyLoading = false;        // 上划加载中标志（防抖）
let _historyRenderedIters = new Set();  // 已渲染的迭代号（跨页去重孤儿迭代）
window._historyRenderedIters = _historyRenderedIters;
let _historyScrollBound = false;    // scroll 监听是否已绑定（只绑一次）

function _resetHistoryPaging() {
    _historyOldestId = null;
    _historyHasMore = false;
    _historyLoading = false;
    _historyRenderedIters = new Set();
    window._historyRenderedIters = _historyRenderedIters;
}

// 重置上传相关状态：密码映射、文件 input、文件列表 UI、附件徽标
// 在切租户/切会话/新建/删除/提交完成等"开始下一轮操作"前调用
function _resetUploadState() {
    _filePasswordsMap = {};
    ['source-files', 'target-file', 'rule-files'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.value = '';
    });
    ['source-file-list', 'target-file-list', 'rule-file-list'].forEach(id => {
        const el = document.getElementById(id);
        if (el) el.innerHTML = '';
    });
    try { _updateAttachBadge(); } catch (e) {}
}

// ==================== 初始化 ====================
document.addEventListener('DOMContentLoaded', function () {
    AUTH.requireAuth();
    AUTH.renderUserInfo(document.querySelector('header'));
    if (AUTH.isAdmin()) {
        const el = document.getElementById('nav-admin');
        if (el) el.style.display = '';
    }
    loadAiProviderOptions();

    // 租户输入框
    const tenantInput = document.getElementById('tenant-input');
    tenantInput.addEventListener('focus', () => _showTenantDropdown());
    tenantInput.addEventListener('input', () => {
        _filterTenantDropdown();
        // 同步手动输入的值到 _currentTenantId
        _currentTenantId = tenantInput.value.trim() || null;
        _applyTenantPermission();
    });
    document.addEventListener('click', (e) => {
        const combo = document.getElementById('tenant-combo');
        if (!combo.contains(e.target)) {
            document.getElementById('tenant-dropdown').style.display = 'none';
        }
    });

    // 文件监听
    document.getElementById('source-files').addEventListener('change', function () {
        _renderFileList('source-file-list', this.files);
        _checkEncryption(Array.from(this.files));
        _updateAttachBadge();
    });
    document.getElementById('target-file').addEventListener('change', function () {
        _renderFileList('target-file-list', this.files);
        _checkEncryption(Array.from(this.files));
        _updateAttachBadge();
    });
    document.getElementById('rule-files').addEventListener('change', function () {
        _renderFileList('rule-file-list', this.files);
        _updateAttachBadge();
    });

    // 初次加载根据当前 mode 同步 label
    if (typeof onModeChange === 'function') {
        try { onModeChange(); } catch (e) { /* 忽略：某些会话路径 mode 控件不存在 */ }
    }

    // 加载租户列表
    _loadTenants();
});

// ==================== 生成模式切换 ====================
function onModeChange() {
    const mode = (document.getElementById('mode') || {}).value || 'formula';
    const labelEl = document.getElementById('target-file-label');
    const hintEl = document.getElementById('target-file-hint');
    if (!labelEl) return;
    if (mode === 'template') {
        labelEl.textContent = '模板文件:';
        if (hintEl) {
            hintEl.style.display = 'block';
            hintEl.textContent = '模板模式：保留此文件的公式与格式，仅按规则填入指定列（智算时自动复用此模板，无需重传）';
        }
    } else if (mode === 'auto') {
        labelEl.textContent = '目标文件 (可选):';
        if (hintEl) {
            hintEl.style.display = 'block';
            hintEl.textContent = '自动模式：目标文件可选；提供则作为列结构软参考，AI 仍按规则自由设计输出';
        }
    } else {
        labelEl.textContent = '目标文件:';
        if (hintEl) { hintEl.style.display = 'none'; hintEl.textContent = ''; }
    }
}

// ==================== 租户列表 ====================
let _tenantList = [];
let _permittedTenantIds = new Set();   // 当前用户有权操作的租户

async function _loadTenants() {
    try {
        const resp = await AUTH.authFetch('/api/tenants');
        if (resp.ok) {
            const data = await resp.json();
            _tenantList = data.tenants || data || [];
        }
    } catch (e) {
        console.warn('加载租户列表失败:', e);
    }
    // 加载当前用户可访问的租户（用于按钮灰化）
    try {
        const resp2 = await AUTH.authFetch('/api/dashboard/tenants');
        if (resp2.ok) {
            const data2 = await resp2.json();
            const items = data2.items || data2.tenants || data2 || [];
            _permittedTenantIds = new Set(items.map(t => t.tenant_id || t.id || t.name).filter(Boolean));
        }
    } catch (e) {
        console.warn('加载可访问租户失败:', e);
    }
    _applyTenantPermission();
}

function _applyTenantPermission() {
    // 智训页允许为新租户起名(后端在租户已有 Script 时才校验 operable),
    // 故前端不做灰化,避免误把"全新租户名"判定为无权。
    const sendBtn = document.getElementById('send-btn');
    const genBtn = document.getElementById('generate-btn');
    if (sendBtn && !_isStreaming) { sendBtn.disabled = false; sendBtn.title = ''; }
    if (genBtn && !_isStreaming) { genBtn.disabled = false; genBtn.title = ''; }
}

function _showTenantDropdown() {
    _filterTenantDropdown();
    document.getElementById('tenant-dropdown').style.display = 'block';
}

function _filterTenantDropdown() {
    const input = document.getElementById('tenant-input');
    const dropdown = document.getElementById('tenant-dropdown');
    const filter = input.value.trim().toLowerCase();

    const filtered = _tenantList.filter(t => {
        const id = (t.tenant_id || t.name || '').toLowerCase();
        return !filter || id.includes(filter);
    });

    dropdown.innerHTML = filtered.map(t => {
        const tid = t.tenant_id || t.name || '';
        const score = t.best_score;
        const scoreHtml = score != null
            ? `<span class="combo-score">${(score * 100).toFixed(0)}%</span>`
            : `<span class="combo-score untrained">未训练</span>`;
        return `<div class="combo-item" onclick="_selectTenant('${tid}')">
            <span class="combo-id">${tid}</span>
            ${scoreHtml}
        </div>`;
    }).join('');

    if (filtered.length === 0) {
        dropdown.innerHTML = '<div style="padding:10px;color:#999;text-align:center;font-size:13px;">无匹配租户</div>';
    }
    dropdown.style.display = 'block';
}

function _selectTenant(tenantId) {
    document.getElementById('tenant-input').value = tenantId;
    document.getElementById('tenant-dropdown').style.display = 'none';
    _currentTenantId = tenantId;
    _currentSessionId = null;
    _chatMessages = [];
    _currentAccuracy = null;
    _currentCode = null;
    _resetUploadState();
    _clearChatUI();
    _hideActionButtons();
    _applyTenantPermission();
    loadSessions();
}

// ==================== AI 模型下拉框（按后台启用状态过滤） ====================
async function loadAiProviderOptions() {
    const sel = document.getElementById('ai-provider');
    if (!sel) return;
    try {
        const resp = await AUTH.authFetch('/api/admin/ai-providers');
        if (!resp.ok) return;
        const data = await resp.json();
        const items = (data.items || []).filter(p => p.enabled);
        if (!items.length) return;   // 未配置时保持默认
        sel.innerHTML = items.map(p => `<option value="${p.key}">${p.label}${p.key === 'claude' ? ' (推荐)' : ''}</option>`).join('');
    } catch (e) {
        console.warn('加载AI模型列表失败:', e);
    }
}

// ==================== 会话管理 ====================
async function loadSessions() {
    if (!_currentTenantId) return;
    try {
        const resp = await AUTH.authFetch(`/api/training/chat/sessions?tenant_id=${encodeURIComponent(_currentTenantId)}`);
        if (!resp.ok) return;
        const data = await resp.json();
        _sessions = data.sessions || [];
        _renderSessionList();
    } catch (e) {
        console.warn('加载会话列表失败:', e);
    }
}

function _renderSessionList() {
    const container = document.getElementById('session-list');
    const emptyHint = document.getElementById('session-list-empty');
    container.querySelectorAll('.session-item').forEach(el => el.remove());

    if (_sessions.length === 0) {
        if (emptyHint) {
            emptyHint.style.display = '';
            emptyHint.textContent = '暂无训练记录，点击 "+ 新建" 开始';
        }
        return;
    }
    if (emptyHint) emptyHint.style.display = 'none';

    _sessions.forEach(s => {
        const div = document.createElement('div');
        div.className = 'session-item' + (s.id === _currentSessionId ? ' active' : '');
        div.onclick = () => selectSession(s.id);

        const titleRow = document.createElement('div');
        titleRow.className = 'session-item-title-row';

        const title = document.createElement('div');
        title.className = 'session-item-title';
        // 优先展示用户命名的脚本名 + 时间，没有则退回 session_key
        const versionLabel = s.session_key || '';
        const time = s.started_at ? new Date(s.started_at).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : '';
        const scriptLabel = s.script_name && !s.script_name.startsWith('script_') ? s.script_name : '';
        title.textContent = scriptLabel ? `${scriptLabel} · ${time}` : (versionLabel || `${s.mode || 'formula'} - ${time}`);
        title.title = `脚本: ${scriptLabel || '(默认)'} | 版本: ${versionLabel} | 会话 #${s.id}`;

        // 重命名按钮
        const renameBtn = document.createElement('button');
        renameBtn.className = 'session-item-rename';
        renameBtn.innerHTML = '&#x270E;';  // ✎
        renameBtn.title = '重命名';
        renameBtn.onclick = (e) => {
            e.stopPropagation();
            _renameSession(s.id, versionLabel);
        };

        titleRow.appendChild(title);
        titleRow.appendChild(renameBtn);

        const meta = document.createElement('div');
        meta.className = 'session-item-meta';

        const accSpan = document.createElement('span');
        accSpan.className = 'session-item-accuracy';
        accSpan.textContent = s.best_accuracy != null ? `${(s.best_accuracy * 100).toFixed(1)}%` : '—';

        const statusSpan = document.createElement('span');
        statusSpan.className = `session-item-status ${s.status}`;
        statusSpan.textContent = _statusText(s.status);

        const iterSpan = document.createElement('span');
        iterSpan.style.cssText = 'font-size:11px;color:#999;';
        iterSpan.textContent = `${s.total_iterations || 0}轮`;

        meta.appendChild(accSpan);
        meta.appendChild(statusSpan);
        meta.appendChild(iterSpan);

        const delBtn = document.createElement('button');
        delBtn.className = 'session-item-delete';
        delBtn.innerHTML = '&#x2715;';
        delBtn.title = '删除';
        delBtn.onclick = (e) => { e.stopPropagation(); deleteSession(s.id); };

        div.appendChild(titleRow);
        div.appendChild(meta);
        div.appendChild(delBtn);
        container.appendChild(div);
    });
}

function _statusText(status) {
    const map = { running: '进行中', completed: '已完成', failed: '失败', cancelled: '已取消' };
    return map[status] || status || '';
}

function createNewSession() {
    // 允许手动输入租户名称
    if (!_currentTenantId) {
        const typed = document.getElementById('tenant-input').value.trim();
        if (typed) {
            _currentTenantId = typed;
        } else {
            alert('请先输入或选择租户');
            return;
        }
    }
    _promptScriptName(_currentTenantId).then(scriptName => {
        if (!scriptName) return;  // 用户取消
        _pendingScriptName = scriptName;

        _currentSessionId = null;
        _chatMessages = [];
        _currentAccuracy = null;
        _currentCode = null;
        _resetUploadState();
        _clearChatUI();
        _hideActionButtons();
        _highlightActiveSession();
        document.getElementById('chat-title').textContent = `新训练 · ${scriptName}`;
        document.getElementById('chat-status').textContent = '';
        document.getElementById('chat-status').className = 'chat-status';

        // 提示用户上传文件
        _addSystemMessage(`脚本名称：${scriptName}\n请通过左下角 📎 上传源文件和目标文件，然后发送消息开始训练。`);
    });
}

// 弹出小模态框让用户输入脚本名（也展示当前租户已有名称作为参考/选择）
async function _promptScriptName(tenantId) {
    // 哈希 ID 格式: script_ + 12 位 hex(参见 storage_manager.save_script)
    const _hashIdPattern = /^script_[0-9a-f]{12}$/i;
    let existingNames = [];
    try {
        const resp = await AUTH.authFetch(`/api/tenant-scripts/${encodeURIComponent(tenantId)}`);
        if (resp.ok) {
            const data = await resp.json();
            // 仅展示有真实 name 的脚本,过滤掉系统生成的哈希 ID
            existingNames = (data.scripts || [])
                .map(s => (s.name || '').trim())
                .filter(n => n && !_hashIdPattern.test(n))
                .filter((v, i, a) => a.indexOf(v) === i);
        }
    } catch (_) {}

    return new Promise((resolve) => {
        const overlay = document.createElement('div');
        overlay.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,0.45);z-index:9999;display:flex;align-items:center;justify-content:center;';
        const existingHtml = existingNames.length
            ? `<div style="margin-top:10px;font-size:12px;color:#666;">租户已有脚本（点击复用同名训练，将覆盖原版本）：</div>
               <div id="_sn_existing" style="display:flex;flex-wrap:wrap;gap:6px;margin-top:6px;">
                   ${existingNames.map(n => `<span class="_sn-chip" style="padding:3px 10px;border:1px solid #cfd8dc;border-radius:14px;cursor:pointer;font-size:12px;background:#f5f7fa;">${n.replace(/</g,'&lt;')}</span>`).join('')}
               </div>`
            : '';
        overlay.innerHTML = `
            <div style="background:#fff;border-radius:10px;padding:22px 24px;width:420px;max-width:90vw;box-shadow:0 4px 18px rgba(0,0,0,0.18);">
                <h3 style="margin:0 0 4px;font-size:16px;color:#333;">为本次训练命名</h3>
                <div style="font-size:12px;color:#888;margin-bottom:14px;">租户「${tenantId}」可同时拥有多个脚本（如：基础数据整合 / 考勤整合 / 算薪）</div>
                <input id="_sn_input" type="text" placeholder="例如：考勤整合" maxlength="80"
                       style="width:100%;box-sizing:border-box;padding:8px 10px;border:1.5px solid #d0d7de;border-radius:6px;font-size:14px;">
                ${existingHtml}
                <div style="display:flex;justify-content:flex-end;gap:8px;margin-top:18px;">
                    <button id="_sn_cancel" style="padding:6px 16px;border:1px solid #ccc;background:#fff;border-radius:6px;cursor:pointer;">取消</button>
                    <button id="_sn_ok" style="padding:6px 16px;border:none;background:#1976d2;color:#fff;border-radius:6px;cursor:pointer;">确定</button>
                </div>
            </div>`;
        document.body.appendChild(overlay);
        const input = overlay.querySelector('#_sn_input');
        input.focus();

        overlay.querySelectorAll('._sn-chip').forEach(el => {
            el.addEventListener('click', () => { input.value = el.textContent; input.focus(); });
        });
        const close = (val) => { document.body.removeChild(overlay); resolve(val); };
        overlay.querySelector('#_sn_cancel').onclick = () => close(null);
        overlay.querySelector('#_sn_ok').onclick = () => {
            const v = input.value.trim();
            if (!v) { input.style.borderColor = '#e53935'; return; }
            close(v);
        };
        input.addEventListener('keydown', (e) => {
            if (e.key === 'Enter') overlay.querySelector('#_sn_ok').click();
            if (e.key === 'Escape') close(null);
        });
    });
}

async function selectSession(sessionId) {
    if (_isStreaming) return;
    _pendingScriptName = null;  // 切换到已存在会话，不再需要新建命名
    _resetUploadState();
    _resetHistoryPaging();
    try {
        const resp = await AUTH.authFetch(`/api/training/chat/sessions/${sessionId}/messages?limit_turns=5`);
        if (!resp.ok) return;
        const data = await resp.json();

        _currentSessionId = data.session.id;
        _currentAccuracy = data.current_accuracy;
        _currentCode = data.current_code;
        // 从会话数据同步租户，防止 _currentTenantId 丢失
        if (data.session.tenant_id) {
            _currentTenantId = data.session.tenant_id;
        }

        _clearChatUI();

        // 显示原始训练文件信息
        _showSessionFilesInfo(data);

        // 合并消息 + 迭代记录，构建完整时间线（首屏最近 5 轮）
        _renderFullHistory(data.messages || [], data.iterations || []);

        // 记录分页游标，绑定上划加载更早
        _historyOldestId = data.page_start_id;
        _historyHasMore = !!data.has_more;
        _bindHistoryScroll();

        // 显示历史训练产物的下载按钮（脚本/输出/差异 + 提示词）
        const latestFiles = data.latest_files || {};
        const hasAnyFile = latestFiles.script_file || latestFiles.output_file || latestFiles.diff_file;
        if (hasAnyFile || data.has_rules || data.session.total_iterations > 0) {
            _showHistoryDownloadBar(_currentSessionId, latestFiles, data.has_rules);
        }

        // 如果训练文件存在且准确率未达100%，显示"分析差异"操作按钮
        const canRetrain = data.session.has_source_files && data.session.has_expected_file;
        if (_currentCode && _currentAccuracy != null && _currentAccuracy < 1.0 && canRetrain) {
            _showAnalyzeDiffButton();
        }
        // 如果训练文件已丢失，提示用户
        if (!canRetrain && data.session.total_iterations > 0) {
            _addSystemMessage('训练源文件已丢失，如需继续训练请创建新会话并重新上传文件。', 'status', { error: true });
        }

        // 更新头部
        const _scriptLabel = data.session.script_name ? ` · ${data.session.script_name}` : '';
        document.getElementById('chat-title').textContent = `训练 #${data.session.id}${_scriptLabel}`;
        _updateChatStatus(data.session.status);
        _updateActionButtons(data.session);

        // 显示"重新生成"按钮（已进入会话且非流式中）
        const regenBtn = document.getElementById('regenerate-btn');
        if (regenBtn) regenBtn.style.display = '';

        _highlightActiveSession();
    } catch (e) {
        console.error('加载会话失败:', e);
    }
}

async function deleteSession(sessionId) {
    if (!confirm('确定删除此训练会话？\n删除后该对话历史、生成的脚本、训练文件将一并删除，且不可恢复。')) return;
    try {
        const resp = await AUTH.authFetch(`/api/training/chat/sessions/${sessionId}`, { method: 'DELETE' });
        if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            alert('删除失败: ' + (err.detail || '未知错误'));
            return;
        }
    } catch (e) {
        console.error('删除失败:', e);
        alert('删除失败: ' + (e.message || '网络错误'));
        return;
    }
    if (_currentSessionId === sessionId) {
        _currentSessionId = null;
        _chatMessages = [];
        _resetUploadState();
        _clearChatUI();
        _hideActionButtons();
        const regenBtn = document.getElementById('regenerate-btn');
        if (regenBtn) regenBtn.style.display = 'none';
    }
    await loadSessions();
}

function _highlightActiveSession() {
    document.querySelectorAll('.session-item').forEach(el => el.classList.remove('active'));
    if (_currentSessionId) {
        const idx = _sessions.findIndex(s => s.id === _currentSessionId);
        const items = document.querySelectorAll('.session-item');
        if (idx >= 0 && items[idx]) items[idx].classList.add('active');
    }
}

async function _renameSession(sessionId, currentName) {
    const newName = prompt('修改版本名称:', currentName || '');
    if (newName === null || newName.trim() === '') return;

    try {
        const resp = await AUTH.authFetch(`/api/training/chat/sessions/${sessionId}/rename`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ name: newName.trim() }),
        });
        if (resp.ok) {
            loadSessions();
        } else {
            const err = await resp.json().catch(() => ({}));
            alert('重命名失败: ' + (err.detail || '未知错误'));
        }
    } catch (e) {
        alert('重命名失败: ' + e.message);
    }
}

// ==================== 附件弹窗 ====================
function toggleAttachPopover() {
    const popover = document.getElementById('attach-popover');
    popover.style.display = popover.style.display === 'none' ? '' : 'none';
}

function _updateAttachBadge() {
    const btn = document.getElementById('attach-btn');
    const has = document.getElementById('source-files').files.length > 0
             || document.getElementById('target-file').files.length > 0;
    btn.classList.toggle('has-files', has);
}

// ==================== 文件列表 ====================
function _renderFileList(containerId, files) {
    const container = document.getElementById(containerId);
    if (!files || files.length === 0) {
        container.innerHTML = '';
        return;
    }
    const arr = Array.from(files);
    const first = arr[0].name;
    const allNames = arr.map(f => f.name).join('\n');
    if (arr.length === 1) {
        container.innerHTML = `<span class="file-item" title="${allNames}">${first}</span>`;
    } else {
        container.innerHTML = `<span class="file-item" title="${allNames}">${first} 等 ${arr.length} 个文件</span>`;
    }
}

// ==================== 加密检测 ====================
async function _probeModernExcelEncryption(file) {
    const name = (file?.name || '').toLowerCase();
    const modern = ['.xlsx', '.xlsm', '.xltx', '.xltm'].some(ext => name.endsWith(ext));
    if (!modern) return null; // 老 .xls 仍交给服务端/Aspose 判断
    try {
        const bytes = new Uint8Array(await file.slice(0, 8).arrayBuffer());
        // 未加密 OOXML 是 ZIP(PK)；加密后的 OOXML 是 OLE Compound(D0 CF 11 E0...)。
        if (bytes[0] === 0x50 && bytes[1] === 0x4b) return false;
        if (bytes[0] === 0xd0 && bytes[1] === 0xcf && bytes[2] === 0x11 && bytes[3] === 0xe0) return true;
    } catch (e) {
        console.warn('本地读取 Excel 文件头失败，回退服务端检测:', file?.name, e);
    }
    return null;
}

async function _checkEncryption(filesToCheck) {
    if (!filesToCheck || filesToCheck.length === 0) return;
    try {
        const encrypted = [];
        const needServerCheck = [];
        for (const file of filesToCheck) {
            if (_filePasswordsMap[file.name]) continue;
            const local = await _probeModernExcelEncryption(file);
            if (local === true) encrypted.push(file.name);
            else if (local === null) needServerCheck.push(file);
        }

        // 普通 xlsx/xlsm 不再重复上传；仅老 .xls 或无法识别的文件请求服务端检测。
        if (needServerCheck.length > 0) {
            const formData = new FormData();
            needServerCheck.forEach(f => formData.append('files', f));
            const resp = await AUTH.authFetch('/api/files/check-encrypted', { method: 'POST', body: formData });
            if (resp.ok) {
                const data = await resp.json();
                (data.encrypted_files || []).forEach(name => {
                    if (!encrypted.includes(name)) encrypted.push(name);
                });
            }
        }
        if (encrypted.length === 0) return;
        const passwords = await _promptFilePasswords(encrypted);
        if (passwords) _filePasswordsMap = { ..._filePasswordsMap, ...passwords };
    } catch (e) {
        console.warn('加密检测失败:', e);
    }
}

function _promptFilePasswords(encryptedFiles) {
    return new Promise((resolve) => {
        const inputs = encryptedFiles.map((name, i) =>
            `<div style="margin-bottom:10px;">
                <label style="display:block;font-size:13px;margin-bottom:4px;color:#333;">${name}</label>
                <input id="_enc_pwd_${i}" type="password" placeholder="请输入打开密码"
                    style="width:100%;padding:8px 12px;border:1.5px solid #d1d5db;border-radius:8px;box-sizing:border-box;font-size:13px;">
            </div>`
        ).join('');

        const overlay = document.createElement('div');
        overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:9999;display:flex;align-items:center;justify-content:center;';
        overlay.innerHTML = `
            <div style="background:#fff;border-radius:16px;padding:28px;width:400px;max-width:90vw;box-shadow:0 10px 40px rgba(0,0,0,0.2);">
                <h3 style="margin:0 0 6px;font-size:16px;font-weight:600;">检测到加密文件</h3>
                <p style="margin:0 0 18px;font-size:13px;color:#6b7280;">以下文件有密码保护，请输入密码：</p>
                ${inputs}
                <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:18px;">
                    <button id="_enc_cancel" style="padding:8px 20px;border:1.5px solid #d1d5db;border-radius:8px;background:#fff;cursor:pointer;font-size:13px;">取消</button>
                    <button id="_enc_confirm" style="padding:8px 20px;border:none;border-radius:8px;background:#1976d2;color:#fff;cursor:pointer;font-size:13px;">确认</button>
                </div>
            </div>`;
        document.body.appendChild(overlay);

        document.getElementById('_enc_cancel').onclick = () => { document.body.removeChild(overlay); resolve(null); };
        document.getElementById('_enc_confirm').onclick = () => {
            const passwords = {};
            encryptedFiles.forEach((name, i) => {
                const pwd = document.getElementById(`_enc_pwd_${i}`).value;
                if (pwd) passwords[name] = pwd;
            });
            const missing = encryptedFiles.filter(name => !passwords[name]);
            if (missing.length > 0) { alert('请为所有加密文件输入密码'); return; }
            document.body.removeChild(overlay);
            resolve(passwords);
        };
        setTimeout(() => document.getElementById('_enc_pwd_0')?.focus(), 100);
    });
}

// ==================== Markdown 渲染 ====================
function _renderMarkdown(text) {
    let html = text.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
    html = html.replace(/```(\w*)\n([\s\S]*?)```/g, (_, lang, code) => `<pre><code>${code}</code></pre>`);
    html = html.replace(/`([^`]+)`/g, '<code>$1</code>');
    html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
    html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
    html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');
    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');
    html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>\n?)+/g, m => '<ul>' + m + '</ul>');
    html = html.replace(/^\|(.+)\|$/gm, function (line) {
        const cells = line.split('|').filter(c => c.trim() !== '');
        if (cells.every(c => /^[\s\-:]+$/.test(c))) return '';
        return '<tr>' + cells.map(c => `<td>${c.trim()}</td>`).join('') + '</tr>';
    });
    html = html.replace(/(<tr>.*<\/tr>\n?)+/g, m => '<table>' + m + '</table>');
    html = html.replace(/\n\n/g, '</p><p>');
    html = html.replace(/\n/g, '<br>');
    return html;
}

// ==================== 对话 UI ====================
// 实时聊天容器；传入其它 target（如 DocumentFragment）时渲染到该目标且不强制滚底（用于上划分页）
function _liveChatContainer() { return document.getElementById('chat-messages'); }

function _addMessage(role, content, isStreaming, target) {
    const live = _liveChatContainer();
    const container = target || live;
    const placeholder = document.getElementById('chat-placeholder');
    if (placeholder) placeholder.style.display = 'none';

    const msgDiv = document.createElement('div');
    msgDiv.className = `message ${role}`;

    const label = document.createElement('div');
    label.className = 'message-label';
    label.textContent = role === 'user' ? '你' : 'AI';

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    if (isStreaming) contentDiv.classList.add('streaming-cursor');

    if (role === 'assistant') {
        contentDiv.innerHTML = _renderMarkdown(content);
    } else {
        contentDiv.textContent = content;
    }

    msgDiv.appendChild(label);
    msgDiv.appendChild(contentDiv);
    container.appendChild(msgDiv);
    if (container === live) live.scrollTop = live.scrollHeight;
    return contentDiv;
}

function _addSystemMessage(content, msgType, metadata, target) {
    const live = _liveChatContainer();
    const container = target || live;
    const placeholder = document.getElementById('chat-placeholder');
    if (placeholder) placeholder.style.display = 'none';

    const msgDiv = document.createElement('div');
    msgDiv.className = 'message system';

    // 根据内容类型添加样式
    if (metadata) {
        if (metadata.error) msgDiv.classList.add('error');
        else if (metadata.rollback) msgDiv.classList.add('warning');
        else if (metadata.accuracy >= 1.0) msgDiv.classList.add('success');
    }
    if (msgType === 'status' && content.includes('失败')) msgDiv.classList.add('error');

    const label = document.createElement('div');
    label.className = 'message-label';
    label.textContent = '系统';

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';
    contentDiv.innerHTML = _renderMarkdown(content);

    // 添加准确率徽章
    if (metadata && metadata.accuracy != null) {
        const badge = document.createElement('span');
        badge.className = 'accuracy-badge';
        const acc = metadata.accuracy;
        badge.classList.add(acc >= 0.95 ? 'high' : acc >= 0.7 ? 'medium' : 'low');
        badge.textContent = `${(acc * 100).toFixed(1)}%`;
        label.appendChild(badge);
    }

    msgDiv.appendChild(label);
    msgDiv.appendChild(contentDiv);
    container.appendChild(msgDiv);
    if (container === live) live.scrollTop = live.scrollHeight;
    return contentDiv;
}

function _updateStreamingMessage(contentDiv, text, thinkingText, thinkingDirty) {
    // 思考过程用【增量 textContent】更新（不重解析 HTML/markdown）：DeepSeek 思考 token 逐块
    // 到达且思考可能长达数分钟/数万字，每 token 全量 _escapeHtml(全部思考)+markdown 重建是
    // O(n²)，思考越长页面越卡（表现为无响应）。textContent 赋值 O(1) 且自动转义。
    let tb = contentDiv.querySelector('.thinking-block');
    if (thinkingDirty && thinkingText) {
        if (!tb) {
            tb = document.createElement('div');
            tb.className = 'thinking-block';
            contentDiv.prepend(tb);
        }
        tb.textContent = thinkingText;
    }
    if (text !== undefined) {
        let body = contentDiv.querySelector('.msg-content-body');
        if (!body) {
            body = document.createElement('div');
            body.className = 'msg-content-body';
            contentDiv.textContent = '';      // 清除占位文本（连带移除子元素，重建顺序）
            if (tb) contentDiv.appendChild(tb);
            contentDiv.appendChild(body);
        }
        body.innerHTML = _renderMarkdown(text || '');
    }
    const container = document.getElementById('chat-messages');
    container.scrollTop = container.scrollHeight;
}

function _finishStreamingMessage(contentDiv) {
    contentDiv.classList.remove('streaming-cursor');
}

// 代码流式追加：将 AI 生成的代码片段逐步显示。
// 性能优化：每 chunk 全量 _renderMarkdown(全部代码) 是 O(n²)（长代码/多轮修正时页面卡死），
// 改为累积 + 100ms 节流渲染；_finishCodeStream 时 flush 最终完整代码。
let _codeStreamEl = null;
let _codeStreamBuf = '';
let _codeStreamTimer = null;

function _appendCodeStream(chunk) {
    const container = document.getElementById('chat-messages');
    if (!_codeStreamEl) {
        // 创建一个 assistant 类型的消息用于代码流
        const placeholder = document.getElementById('chat-placeholder');
        if (placeholder) placeholder.style.display = 'none';

        const msgDiv = document.createElement('div');
        msgDiv.className = 'message assistant';

        const label = document.createElement('div');
        label.className = 'message-label';
        label.textContent = 'AI';

        const contentDiv = document.createElement('div');
        contentDiv.className = 'message-content streaming-cursor';

        msgDiv.appendChild(label);
        msgDiv.appendChild(contentDiv);
        container.appendChild(msgDiv);

        _codeStreamEl = contentDiv;
        _codeStreamBuf = '';
    }

    _codeStreamBuf += chunk;
    if (_codeStreamTimer) return;   // 节流：已有排程
    _codeStreamTimer = setTimeout(() => {
        _codeStreamTimer = null;
        _renderCodeStreamOnce();
    }, 100);
}

function _renderCodeStreamOnce() {
    if (!_codeStreamEl) return;
    _codeStreamEl.innerHTML = _renderMarkdown('```python\n' + _codeStreamBuf + '\n```');
    const container = document.getElementById('chat-messages');
    container.scrollTop = container.scrollHeight;
}

function _finishCodeStream() {
    if (_codeStreamEl) {
        if (_codeStreamTimer) { clearTimeout(_codeStreamTimer); _codeStreamTimer = null; }
        _renderCodeStreamOnce();   // flush：渲染完整最终代码
        _codeStreamEl.classList.remove('streaming-cursor');
        _codeStreamEl = null;
        _codeStreamBuf = '';
    }
}

function _clearChatUI() {
    const container = document.getElementById('chat-messages');
    container.innerHTML = `
        <div class="chat-placeholder" id="chat-placeholder">
            <div class="placeholder-content">
                <div class="placeholder-icon">&#x1F9E0;</div>
                <p class="placeholder-title">智能训练对话</p>
                <p class="placeholder-hint">选择租户 → 新建或选择训练会话 → 上传文件开始训练</p>
            </div>
        </div>
    `;
}

function _updateChatStatus(status) {
    const el = document.getElementById('chat-status');
    el.textContent = _statusText(status);
    el.className = `chat-status ${status || ''}`;
}

function _updateActionButtons(session) {
    const btnSetBest = document.getElementById('btn-set-best');
    const btnUploadCode = document.getElementById('btn-upload-code');
    const btnDownloadCode = document.getElementById('btn-download-code');

    const hasSession = !!_currentSessionId;
    const hasCode = !!_currentCode;
    const isCompleted = session && session.status === 'completed';

    btnSetBest.style.display = hasSession && hasCode ? '' : 'none';
    btnUploadCode.style.display = hasSession ? '' : 'none';
    btnDownloadCode.style.display = hasCode ? '' : 'none';

    if (isCompleted && session.has_script) {
        btnSetBest.disabled = true;
        btnSetBest.textContent = '已设置';
    } else {
        btnSetBest.disabled = false;
        btnSetBest.textContent = '设为最佳';
    }
}

function _hideActionButtons() {
    document.getElementById('btn-set-best').style.display = 'none';
    document.getElementById('btn-upload-code').style.display = 'none';
    document.getElementById('btn-download-code').style.display = 'none';
}

// ==================== 发送消息 ====================
function sendMessage(action) {
    const input = document.getElementById('chat-input');
    const text = input.value.trim();
    if (_isStreaming) return;

    // action: undefined/'chat' = 对话讨论, 'generate' = 执行代码修正, 'regenerate' = 上传新文件后重新生成
    action = action || 'chat';

    // 已在会话中 → 直接发消息，无需再选租户
    if (_currentSessionId) {
        if (action === 'regenerate') {
            _sendRegenerateMessage(text);
            input.value = '';
            input.style.height = '';
            return;
        }
        if (!text && action === 'chat') return;
        _sendChatMessage(text || '请根据之前的讨论修正代码', action);
        input.value = '';
        input.style.height = '';
        return;
    }

    // 新会话：需要租户
    if (!_currentTenantId) {
        const typed = document.getElementById('tenant-input').value.trim();
        if (typed) {
            _currentTenantId = typed;
        } else {
            alert('请先输入或选择租户');
            return;
        }
    }

    const sourceFiles = document.getElementById('source-files').files;
    const targetFile = document.getElementById('target-file').files[0];
    const hasFiles = sourceFiles.length > 0 && targetFile;

    if (hasFiles) {
        _startTraining(text || '请根据文件生成数据处理脚本');
    } else {
        alert('请先通过 📎 上传源文件和目标文件，然后开始训练');
        return;
    }

    input.value = '';
    input.style.height = '';
}

function _startTraining(userText) {
    if (_isStreaming) return;  // 防止重复触发
    const sourceFiles = document.getElementById('source-files').files;
    const targetFile = document.getElementById('target-file').files[0];
    const ruleFiles = document.getElementById('rule-files').files;
    const aiProvider = document.getElementById('ai-provider').value;
    const mode = document.getElementById('mode').value;
    const salaryMonth = document.getElementById('salary-month').value.trim();
    const standardHours = document.getElementById('standard-hours').value.trim();

    // 模板/自动模式：先轻量解析目标文件，弹出 sheet 多选框，让用户勾选目标 sheet
    const _needsPicker = (mode === 'template' || mode === 'auto') && targetFile;
    const _go = (selectedSheets) => {
        _submitStartTraining(userText, sourceFiles, targetFile, ruleFiles,
                             aiProvider, mode, salaryMonth, standardHours,
                             selectedSheets || []);
    };
    if (_needsPicker) {
        console.log('[训练] 模板/自动模式 → 调用 peek-template-sheets');
        _pickTargetSheets(targetFile).then(selectedSheets => {
            if (selectedSheets === null) {
                console.log('[训练] 用户取消 sheet 选择，跳过本次训练');
                return;  // 取消才中止
            }
            _go(selectedSheets);
        }).catch(err => {
            console.warn('[sheet 选择] 解析失败，fallback 走全量 sheet:', err);
            try { _addMessage('system', '目标文件 sheet 解析失败：' + (err && err.message ? err.message : err) + '；将默认使用全部 sheet 进入训练'); } catch (e) {}
            _go([]);  // 软失败：仍进入训练，由后端按全量 sheet 处理
        });
    } else {
        _go([]);  // 公式模式/直接导入：不需选择
    }
}

function _submitStartTraining(userText, sourceFiles, targetFile, ruleFiles,
                              aiProvider, mode, salaryMonth, standardHours,
                              selectedSheets) {
    const formData = new FormData();
    formData.append('tenant_id', _currentTenantId);
    if (_pendingScriptName) formData.append('script_name', _pendingScriptName);
    formData.append('ai_provider', aiProvider);
    formData.append('mode', mode);
    if (salaryMonth) formData.append('salary_year_month', salaryMonth);
    if (standardHours) formData.append('monthly_standard_hours', standardHours);
    const manualHeaders = document.getElementById('manual-headers').value.trim();
    if (manualHeaders) formData.append('manual_headers', manualHeaders);
    const multiSheetSource = document.getElementById('multi-sheet-source').checked;
    if (multiSheetSource) formData.append('multi_sheet_source', 'true');
    const _useHistEl = document.getElementById('use-history');
    formData.append('use_history', (_useHistEl && _useHistEl.checked) ? 'true' : 'false');
    if (_currentSessionId) formData.append('session_id', _currentSessionId);
    if (Object.keys(_filePasswordsMap).length > 0) {
        formData.append('file_passwords', JSON.stringify(_filePasswordsMap));
        console.log('[训练] file_passwords:', JSON.stringify(_filePasswordsMap));
    }
    if (selectedSheets && selectedSheets.length > 0) {
        formData.append('target_sheets', JSON.stringify(selectedSheets));
        console.log('[训练] target_sheets:', selectedSheets);
    }
    // 文件字段放在所有文本字段之后，避免 python-multipart 旧版本解析丢失后续字段
    Array.from(sourceFiles).forEach(f => formData.append('source_files', f));
    if (targetFile) formData.append('target_file', targetFile);
    if (ruleFiles.length > 0) {
        Array.from(ruleFiles).forEach(f => formData.append('rule_files', f));
    }

    _addMessage('user', userText);
    document.getElementById('attach-popover').style.display = 'none';
    _setUIStreaming(true);

    _fetchTrainingSSE('/api/training/chat/start', { method: 'POST', body: formData });
    _resetUploadState();
}

// ==================== 目标 Sheet 选择弹窗 ====================

let _sheetPickerResolver = null;  // Promise resolve 函数

function _pickTargetSheets(targetFile) {
    return new Promise((resolve, reject) => {
        // 先检查 modal DOM 是否存在（应对浏览器缓存了旧版 training.html 的情况）
        const modal = document.getElementById('sheet-picker-modal');
        const list = document.getElementById('sheet-picker-list');
        const title = document.getElementById('sheet-picker-title');
        if (!modal || !list || !title) {
            return reject(new Error('sheet 选择弹窗 DOM 不存在；请按 Ctrl+F5 强制刷新页面再试'));
        }

        const fd = new FormData();
        fd.append('target_file', targetFile);
        const _pwd = _filePasswordsMap[targetFile.name];
        if (_pwd) fd.append('file_password', _pwd);

        // 强制 inline 样式，避免 CSS 缓存（.modal-overlay 类样式没刷新到时）导致弹窗肉眼看不见
        Object.assign(modal.style, {
            display: 'flex',
            position: 'fixed', top: '0', left: '0', right: '0', bottom: '0',
            background: 'rgba(0,0,0,0.45)', zIndex: '9999',
            justifyContent: 'center', alignItems: 'flex-start',
            paddingTop: '8vh', overflowY: 'auto',
        });
        // 内层 dialog 也强制
        const dialog = modal.querySelector('.modal-dialog');
        if (dialog) {
            Object.assign(dialog.style, {
                background: '#fff', borderRadius: '12px',
                boxShadow: '0 8px 32px rgba(0,0,0,0.18)',
                maxWidth: '560px', width: '90vw', padding: '0',
            });
        }
        // header 加点 padding
        const header = modal.querySelector('.modal-header');
        if (header) {
            Object.assign(header.style, {
                padding: '14px 18px', borderBottom: '1px solid #eee',
                fontWeight: '600', display: 'flex',
                justifyContent: 'space-between', alignItems: 'center',
            });
        }
        const body = modal.querySelector('.modal-body');
        if (body) Object.assign(body.style, { padding: '14px 18px' });

        title.textContent = '正在分析目标文件...';
        list.innerHTML = '<div style="padding:16px;text-align:center;color:#888;">解析中，请稍候...</div>';
        console.log('[sheet picker] modal 已显示，开始 peek',
                    'computed:', getComputedStyle(modal).display, getComputedStyle(modal).position);

        // 用 AUTH.authFetch 携带登录 token
        AUTH.authFetch('/api/training/chat/peek-template-sheets', {
            method: 'POST', body: fd,
        })
        .then(r => r.ok ? r.json() : r.json().then(j => Promise.reject(new Error(j.detail || ('HTTP ' + r.status)))))
        .then(data => {
            const sheets = (data && data.sheets) || [];
            if (sheets.length === 0) {
                modal.style.display = 'none';
                return reject(new Error('目标文件没有可用 sheet'));
            }
            title.textContent = `选择需要处理的 Sheet（共 ${sheets.length} 个）`;
            list.innerHTML = sheets.map((s, i) => {
                const previewStr = (s.preview || []).slice(0, 2).map(row => row.filter(x => x).slice(0, 6).join(' | ')).join(' / ') || '(空)';
                const safeName = String(s.name).replace(/"/g, '&quot;');
                return `
                <label style="display:block;padding:6px 4px;border-bottom:1px solid #f0f0f0;cursor:pointer;">
                    <input type="checkbox" class="sheet-picker-chk" data-name="${safeName}" checked style="width:auto;margin-right:8px;">
                    <span style="font-weight:500;">${safeName}</span>
                    <div style="margin-left:24px;font-size:12px;color:#888;white-space:nowrap;overflow:hidden;text-overflow:ellipsis;">${previewStr}</div>
                </label>`;
            }).join('');
            _sheetPickerResolver = resolve;
        })
        .catch(err => {
            modal.style.display = 'none';
            reject(err);
        });
    });
}

function sheetPickerToggleAll(checked) {
    document.querySelectorAll('#sheet-picker-list .sheet-picker-chk').forEach(c => { c.checked = !!checked; });
}

function closeSheetPicker(cancelled) {
    const modal = document.getElementById('sheet-picker-modal');
    let result = null;
    if (!cancelled) {
        const checked = document.querySelectorAll('#sheet-picker-list .sheet-picker-chk:checked');
        result = Array.from(checked).map(c => c.getAttribute('data-name'));
        if (result.length === 0) {
            alert('请至少勾选一个 Sheet');
            return;
        }
    }
    modal.style.display = 'none';
    if (typeof _sheetPickerResolver === 'function') {
        const r = _sheetPickerResolver;
        _sheetPickerResolver = null;
        r(result);
    }
}

function _sendChatMessage(text, action) {
    const ruleFiles = document.getElementById('rule-files').files;

    const formData = new FormData();
    formData.append('message', text);
    formData.append('action', action || 'chat');
    if (ruleFiles.length > 0) {
        Array.from(ruleFiles).forEach(f => formData.append('rule_files', f));
    }

    _addMessage('user', text);
    _setUIStreaming(true);

    _fetchTrainingSSE(`/api/training/chat/sessions/${_currentSessionId}/message`, {
        method: 'POST',
        body: formData,
    });
}

function _sendRegenerateMessage(text) {
    const sourceFiles = document.getElementById('source-files').files;
    const targetFile = document.getElementById('target-file').files[0];
    const ruleFiles = document.getElementById('rule-files').files;

    const hasAny = (sourceFiles && sourceFiles.length > 0) || !!targetFile || (ruleFiles && ruleFiles.length > 0);
    if (!hasAny) {
        const ok = confirm('未选择新文件。继续将使用原有文件重新生成代码（追加为新一轮迭代）。是否继续？');
        if (!ok) return;
    }

    const formData = new FormData();
    formData.append('message', text || '使用新文件重新生成');
    formData.append('action', 'regenerate');

    if (sourceFiles && sourceFiles.length > 0) {
        Array.from(sourceFiles).forEach(f => formData.append('source_files', f));
    }
    if (targetFile) {
        formData.append('expected_result', targetFile);
    }
    if (ruleFiles && ruleFiles.length > 0) {
        Array.from(ruleFiles).forEach(f => formData.append('rule_files', f));
    }

    // 描述文本带上文件标签，方便用户在历史中识别这一轮
    const tagParts = [];
    if (sourceFiles && sourceFiles.length > 0) tagParts.push(`源文件×${sourceFiles.length}`);
    if (targetFile) tagParts.push(`目标=${targetFile.name}`);
    if (ruleFiles && ruleFiles.length > 0) tagParts.push(`规则×${ruleFiles.length}`);
    const tag = tagParts.length > 0 ? `\n[新文件: ${tagParts.join(', ')}]` : '\n[未上传新文件]';
    _addMessage('user', (text || '使用新文件重新生成') + tag);

    // 关闭附件 popover
    const pop = document.getElementById('attach-popover');
    if (pop) pop.style.display = 'none';

    _setUIStreaming(true);
    _fetchTrainingSSE(`/api/training/chat/sessions/${_currentSessionId}/message`, {
        method: 'POST',
        body: formData,
    });
    _resetUploadState();
}

// ==================== SSE ====================
async function _fetchTrainingSSE(url, options) {
    let _gotAnyEvent = false;
    try {
        // 使用 AUTH.authFetch 保证 token 正确携带 + 401 自动跳转登录
        const response = await AUTH.authFetch(url, options);
        if (!response.ok) {
            let errMsg = `HTTP ${response.status}`;
            try { errMsg = (await response.json()).detail || errMsg; } catch (e) {}
            throw new Error(errMsg);
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop() || '';

            for (const line of lines) {
                if (!line.startsWith('data: ')) continue;
                const jsonStr = line.slice(6).trim();
                if (!jsonStr) continue;
                try {
                    const event = JSON.parse(jsonStr);
                    _gotAnyEvent = true;
                    _handleSSEEvent(event);
                } catch (parseErr) {
                    if (parseErr.message && !parseErr.message.includes('JSON')) throw parseErr;
                }
            }
        }
    } catch (e) {
        console.error('SSE error:', e);
        // 网络中断 / 服务热重载：fetch 抛 TypeError("Failed to fetch")
        // 给出可操作的提示，而不是裸的 "Failed to fetch"
        const isNetErr = (e && e.name === 'TypeError') ||
                         (e && typeof e.message === 'string' && /Failed to fetch|NetworkError|ERR_/i.test(e.message));
        if (isNetErr) {
            const hint = _gotAnyEvent
                ? '与服务器的连接已断开（可能服务正在重启），请稍后重试或刷新页面。'
                : '无法连接到服务器，请确认服务在线后重试。';
            _addSystemMessage(hint, 'status', { error: true });
        } else {
            _addSystemMessage('请求失败: ' + e.message, 'status', { error: true });
        }
    } finally {
        _setUIStreaming(false);
    }
}

function _handleSSEEvent(event) {
    switch (event.type) {
        case 'session_created':
            _currentSessionId = event.session_id;
            document.getElementById('chat-title').textContent = `训练 #${event.session_id}`;
            // 刷新租户列表（新租户目录可能刚被创建）
            _loadTenants();
            break;

        case 'status':
            _addSystemMessage(event.message, 'status');
            break;

        case 'iteration_complete': {
            _finishCodeStream();  // 结束代码流式显示
            const acc = event.accuracy;
            _currentAccuracy = acc;

            // 设置 _currentCode 标记，使 "设为最佳"/"下载代码" 按钮可见
            if (event.success) {
                _currentCode = 'generated';
            }

            if (event.rollback) {
                const msg = `本轮修改导致准确率从 ${(event.accuracy * 100).toFixed(1)}% 下降到 ${(event.attempted_accuracy * 100).toFixed(1)}%，已自动回滚到之前的最佳代码。`;
                _addSystemMessage(msg, 'status', { rollback: true, accuracy: event.accuracy });
            } else if (event.success) {
                const accPct = (acc * 100).toFixed(1);
                if (acc >= 1.0) {
                    _addSystemMessage(
                        `第 ${event.iteration} 轮完成，准确率 ${accPct}%，所有数据匹配！`,
                        'status', { accuracy: acc }
                    );
                } else {
                    let msg = `第 ${event.iteration} 轮完成，准确率 ${accPct}%`;
                    if (event.diff_details) {
                        msg += '\n\n' + _formatDiffDetails(event.diff_details);
                    }
                    msg += '\n\n请描述需要调整的逻辑，AI 将根据反馈修正代码。';
                    _addSystemMessage(msg, 'diff', { accuracy: acc, diff_details: event.diff_details });
                }
            } else {
                _addSystemMessage(
                    `第 ${event.iteration} 轮执行失败: ${event.error || '未知错误'}`,
                    'status', { error: event.error }
                );
            }

            // 显示下载按钮区
            if ((event.success || event.rollback) && _currentSessionId) {
                _showDownloadBar(_currentSessionId, event.files);
            }

            _updateActionButtons({ status: event.success ? 'running' : 'failed', has_script: false });
            _updateChatStatus(event.success ? 'running' : 'failed');
            loadSessions();
            _loadTenants();   // 刷新租户训练分数
            break;
        }

        case 'assistant_message':
            _addMessage('assistant', event.content);
            break;

        case 'chat_chunk':
            // AI 对话流式输出
            if (!_chatStreamEl) {
                _chatStreamEl = _addMessage('assistant', '', true);
                _chatStreamBuf = '';
                _chatThinkingBuf = '';
            }
            _chatStreamBuf += event.content;
            _updateStreamingMessage(_chatStreamEl, _chatStreamBuf, _chatThinkingBuf);
            break;

        case 'thinking':
            // DeepSeek 推理模型思考过程：灰色区流式显示（不进入正式内容）。
            // 思考 token 逐块到达（可能数分钟/数万字），每 token 全量重渲染是 O(n²) 页面卡死源；
            // 改为累积 + 100ms 节流 + 增量 textContent（_updateStreamingMessage 已优化）。
            if (!_chatStreamEl) {
                _chatStreamEl = _addMessage('assistant', '', true);
                _chatStreamBuf = '';
                _chatThinkingBuf = '';
            }
            _chatThinkingBuf += event.content;
            if (!_chatThinkingTimer) {
                _chatThinkingTimer = setTimeout(() => {
                    _chatThinkingTimer = null;
                    _updateStreamingMessage(_chatStreamEl, undefined, _chatThinkingBuf, true);
                }, 100);
            }
            break;

        case 'chat_done':
            // AI 对话完成
            if (_chatStreamEl) {
                if (_chatThinkingTimer) { clearTimeout(_chatThinkingTimer); _chatThinkingTimer = null; }
                _finishStreamingMessage(_chatStreamEl);
                // 用最终完整内容重新渲染（确保 markdown 完整；思考过程保留在灰色区）
                _updateStreamingMessage(_chatStreamEl, _chatStreamBuf || event.content, _chatThinkingBuf, true);
                _chatStreamEl = null;
                _chatStreamBuf = '';
                _chatThinkingBuf = '';
            } else {
                _addMessage('assistant', event.content);
            }
            break;

        case 'error':
            _finishCodeStream();
            _addSystemMessage(event.message, 'status', { error: true });
            break;

        case 'log':
            // 训练引擎的日志输出（含代码流）— 追加到对话框
            if (event.message) {
                // 匹配 [HH:MM:SS] [CODE] chunk 格式
                const codeMatch = event.message.match(/\[CODE\]\s*([\s\S]*)/);
                if (codeMatch) {
                    _appendCodeStream(codeMatch[1]);
                }
                // 其他日志不显示在对话框（避免刷屏），但可以 console.log
            }
            break;
    }
}

function _formatDiffDetails(diff) {
    if (!diff) return '';
    let text = '**差异详情:**\n';

    if (diff.field_diff_samples) {
        const samples = diff.field_diff_samples;
        if (Array.isArray(samples)) {
            text += '\n| 字段 | 期望值 | 实际值 |\n|---|---|---|\n';
            samples.slice(0, 10).forEach(s => {
                text += `| ${s.field || s.column || '—'} | ${s.expected ?? '—'} | ${s.actual ?? '—'} |\n`;
            });
            if (samples.length > 10) {
                text += `\n... 共 ${samples.length} 处差异\n`;
            }
        } else if (typeof samples === 'object') {
            for (const [field, details] of Object.entries(samples)) {
                if (Array.isArray(details)) {
                    text += `\n**${field}**: ${details.length} 处差异\n`;
                    details.slice(0, 3).forEach(d => {
                        text += `  - 行${d.row || '?'}: 期望 \`${d.expected ?? ''}\` / 实际 \`${d.actual ?? ''}\`\n`;
                    });
                }
            }
        }
    }

    if (diff.total_cells != null && diff.matched_cells != null) {
        text += `\n总单元格: ${diff.total_cells}, 匹配: ${diff.matched_cells}\n`;
    }

    return text;
}

async function _downloadFile(sessionId, fileType) {
    try {
        const resp = await AUTH.authFetch(`/api/training/chat/sessions/${sessionId}/download/${fileType}`);
        if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            alert('下载失败: ' + (err.detail || `HTTP ${resp.status}`));
            return;
        }
        const blob = await resp.blob();
        const disposition = resp.headers.get('content-disposition') || '';
        const fnMatch = disposition.match(/filename[*]?=(?:UTF-8'')?["']?([^"';\n]+)/i);
        const filename = fnMatch ? decodeURIComponent(fnMatch[1]) : `${fileType}_${sessionId}`;

        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = filename;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } catch (e) {
        alert('下载失败: ' + e.message);
    }
}

function _showDownloadBar(sessionId, files) {
    const container = document.getElementById('chat-messages');
    const bar = document.createElement('div');
    bar.className = 'message system download-bar';

    const label = document.createElement('div');
    label.className = 'message-label';
    label.textContent = '下载';

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content download-buttons';

    const baseUrl = `/api/training/chat/sessions/${sessionId}/download`;

    const items = [
        { type: 'output', label: '生成结果 (.xlsx)', icon: '📊' },
        { type: 'diff', label: '差异对比 (.xlsx)', icon: '📋' },
        { type: 'script', label: '脚本 (.py)', icon: '📄' },
    ];

    items.forEach(item => {
        // 检查是否有对应文件
        const hasFile = files && (
            (item.type === 'script' && files.script_file) ||
            (item.type === 'output' && files.output_file) ||
            (item.type === 'diff' && files.diff_file)
        );
        if (!hasFile) return;

        const btn = document.createElement('button');
        btn.className = 'download-btn';
        btn.textContent = `${item.icon} ${item.label}`;
        btn.onclick = () => _downloadFile(sessionId, item.type);
        contentDiv.appendChild(btn);
    });

    // 提示词/上下文下载
    const promptBtn = document.createElement('button');
    promptBtn.className = 'download-btn prompt-btn';
    promptBtn.textContent = '📝 提示词/上下文';
    promptBtn.onclick = () => _downloadOriginalFile(sessionId, 'prompt');
    contentDiv.appendChild(promptBtn);

    // 规则文件下载
    if (files && files.has_rules) {
        const rulesBtn = document.createElement('button');
        rulesBtn.className = 'download-btn';
        rulesBtn.textContent = '📖 规则文件';
        rulesBtn.onclick = () => _downloadOriginalFile(sessionId, 'rules');
        contentDiv.appendChild(rulesBtn);
    }

    // 最终规则（AI 整理：原始规则 + 多轮对话 + 当前最佳代码）
    const finalBtn1 = document.createElement('button');
    finalBtn1.className = 'download-btn';
    finalBtn1.textContent = '📋 最终规则';
    finalBtn1.title = '根据原始规则+多轮对话+当前最佳代码，整理出可直接用于下次训练的最终规则';
    finalBtn1.onclick = () => _downloadFinalRules(sessionId, finalBtn1);
    contentDiv.appendChild(finalBtn1);

    if (contentDiv.children.length === 0) return;  // 没有可下载的文件

    bar.appendChild(label);
    bar.appendChild(contentDiv);
    container.appendChild(bar);
    container.scrollTop = container.scrollHeight;
}

function _showAnalyzeDiffButton() {
    const container = document.getElementById('chat-messages');
    const bar = document.createElement('div');
    bar.className = 'message system analyze-bar';

    const label = document.createElement('div');
    label.className = 'message-label';
    label.textContent = '操作';

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content';

    const btn = document.createElement('button');
    btn.className = 'download-btn analyze-btn';
    btn.textContent = '分析差异并修正';
    btn.onclick = () => {
        if (_isStreaming) return;
        const input = document.getElementById('chat-input');
        input.value = '请分析上次执行结果与预期的差异，找出代码中的问题并给出修改建议';
        sendMessage();
    };

    contentDiv.appendChild(btn);
    bar.appendChild(label);
    bar.appendChild(contentDiv);
    container.appendChild(bar);
    container.scrollTop = container.scrollHeight;
}

function _showSessionFilesInfo(data) {
    const container = document.getElementById('chat-messages');
    const placeholder = document.getElementById('chat-placeholder');
    if (placeholder) placeholder.style.display = 'none';

    const sourceNames = data.source_file_names || [];
    const expectedName = data.expected_file_name;
    if (sourceNames.length === 0 && !expectedName) return;

    const msgDiv = document.createElement('div');
    msgDiv.className = 'message system files-info';

    const label = document.createElement('div');
    label.className = 'message-label';
    label.textContent = '训练文件';

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content session-files-content';

    // 源文件列表 + 下载
    if (sourceNames.length > 0) {
        const srcSection = document.createElement('div');
        srcSection.className = 'files-section';
        srcSection.innerHTML = '<strong>源文件:</strong> ';
        sourceNames.forEach(fn => {
            const link = document.createElement('a');
            link.className = 'file-download-link';
            link.textContent = fn;
            link.href = '#';
            link.onclick = (e) => {
                e.preventDefault();
                _downloadOriginalFile(_currentSessionId, 'source', fn);
            };
            srcSection.appendChild(link);
            srcSection.appendChild(document.createTextNode(' '));
        });
        contentDiv.appendChild(srcSection);
    }

    // 预期文件 + 下载
    if (expectedName) {
        const expSection = document.createElement('div');
        expSection.className = 'files-section';
        expSection.innerHTML = '<strong>预期文件:</strong> ';
        const link = document.createElement('a');
        link.className = 'file-download-link';
        link.textContent = expectedName;
        link.href = '#';
        link.onclick = (e) => {
            e.preventDefault();
            _downloadOriginalFile(_currentSessionId, 'expected');
        };
        expSection.appendChild(link);
        contentDiv.appendChild(expSection);
    }

    msgDiv.appendChild(label);
    msgDiv.appendChild(contentDiv);
    container.appendChild(msgDiv);
}

/**
 * 合并消息 + 迭代记录，构建完整的对话时间线
 * 迭代记录包含代码生成详情，补充消息中缺失的信息
 */
function _renderFullHistory(messages, iterations, target) {
    // 构建迭代按 iteration_num 索引
    const iterMap = {};
    iterations.forEach(it => { iterMap[it.iteration_num] = it; });

    // 记录哪些迭代已经通过消息 metadata 提及（避免重复）
    const mentionedIters = new Set();
    messages.forEach(msg => {
        if (msg.metadata && msg.metadata.iteration) {
            mentionedIters.add(msg.metadata.iteration);
        }
    });

    // 按时间线渲染消息，在合适位置插入迭代详情
    let lastRenderedIter = 0;

    messages.forEach(msg => {
        // 在此消息之前，检查是否有未展示的迭代信息需要插入
        const msgIter = msg.metadata && msg.metadata.iteration;

        // 渲染原始消息
        if (msg.role === 'user') {
            _addMessage('user', msg.content, false, target);
        } else if (msg.role === 'assistant') {
            _addMessage('assistant', msg.content, false, target);
        } else if (msg.role === 'system') {
            _addSystemMessage(msg.content, msg.msg_type, msg.metadata, target);
        }

        // 如果这条消息是迭代结果消息，在其后追加代码摘要
        if (msgIter && iterMap[msgIter]) {
            const it = iterMap[msgIter];
            if (it.generated_code) {
                _renderIterationCodeSummary(it, target);
                if (window._historyRenderedIters) _historyRenderedIters.add(msgIter);
            }
            lastRenderedIter = msgIter;
        }
    });

    // 如果有迭代没有对应的消息（例如首轮训练后没保存足够消息），补充渲染
    iterations.forEach(it => {
        if (mentionedIters.has(it.iteration_num)) return;
        if (window._historyRenderedIters && _historyRenderedIters.has(it.iteration_num)) return;
        _renderOrphanIteration(it, target);
        if (window._historyRenderedIters) _historyRenderedIters.add(it.iteration_num);
    });
}

/**
 * 绑定聊天容器的上划滚动监听（只绑一次），滚到顶部自动加载更早对话
 */
function _bindHistoryScroll() {
    if (_historyScrollBound) return;
    const c = _liveChatContainer();
    if (!c) return;
    c.addEventListener('scroll', () => {
        if (c.scrollTop < 40 && _historyHasMore && !_historyLoading) {
            _loadEarlierHistory();
        }
    });
    _historyScrollBound = true;
}

/**
 * 加载更早的一页对话，从顶部插入并保持视口位置
 */
async function _loadEarlierHistory() {
    if (_historyLoading || !_historyHasMore || _historyOldestId == null) return;
    _historyLoading = true;
    const c = _liveChatContainer();

    // 顶部临时加载指示
    const spinner = document.createElement('div');
    spinner.className = 'message system';
    spinner.innerHTML = '<div class="message-content" style="text-align:center;opacity:.6;">加载更早对话…</div>';
    c.insertBefore(spinner, c.firstChild);

    try {
        const resp = await AUTH.authFetch(
            `/api/training/chat/sessions/${_currentSessionId}/messages?limit_turns=5&before_id=${_historyOldestId}`
        );
        if (!resp.ok) return;
        const data = await resp.json();

        // 渲染进离屏 fragment，避免逐条 reflow
        const frag = document.createDocumentFragment();
        _renderFullHistory(data.messages || [], data.iterations || [], frag);

        // 保持滚动位置：插入前后 scrollHeight 差补偿
        if (spinner.parentNode) c.removeChild(spinner);
        const prevH = c.scrollHeight;
        c.insertBefore(frag, c.firstChild);
        c.scrollTop += c.scrollHeight - prevH;

        _historyOldestId = data.page_start_id;
        _historyHasMore = !!data.has_more;
    } catch (e) {
        // 忽略，下方 finally 兜底清理
    } finally {
        if (spinner.parentNode) c.removeChild(spinner);
        _historyLoading = false;
    }
}

/**
 * 渲染迭代代码摘要（折叠式，可展开查看代码）
 */
function _renderIterationCodeSummary(iteration, target) {
    const container = target || _liveChatContainer();
    const msgDiv = document.createElement('div');
    msgDiv.className = 'message system';

    const label = document.createElement('div');
    label.className = 'message-label';
    label.textContent = `第 ${iteration.iteration_num} 轮代码`;

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content iteration-code-summary';

    const codeLines = iteration.generated_code.split('\n');
    const lineCount = codeLines.length;
    const preview = codeLines.slice(0, 6).join('\n');

    const toggleId = `code-toggle-${iteration.iteration_num}`;
    contentDiv.innerHTML = `
        <div class="code-summary-header" onclick="document.getElementById('${toggleId}').classList.toggle('expanded')">
            <span>生成代码（${lineCount} 行）</span>
            <span class="code-toggle-icon">▶ 展开/收起</span>
        </div>
        <div id="${toggleId}" class="code-collapse">
            <pre><code>${_escapeHtml(iteration.generated_code)}</code></pre>
        </div>
    `;

    msgDiv.appendChild(label);
    msgDiv.appendChild(contentDiv);
    container.appendChild(msgDiv);
}

/**
 * 渲染没有对应消息的孤立迭代记录
 */
function _renderOrphanIteration(iteration, target) {
    // 显示迭代结果
    const acc = iteration.accuracy;
    let resultText = `第 ${iteration.iteration_num} 轮`;
    if (acc != null) {
        resultText += ` — 准确率 ${(acc * 100).toFixed(1)}%`;
    }
    if (iteration.status === 'failed') {
        resultText += '（执行失败）';
    }

    const metadata = {
        iteration: iteration.iteration_num,
        accuracy: acc
    };
    _addSystemMessage(resultText, acc != null ? 'status' : 'status', metadata, target);

    // 如果有代码，显示折叠摘要
    if (iteration.generated_code) {
        _renderIterationCodeSummary(iteration, target);
    }
}

function _escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}

function _showHistoryDownloadBar(sessionId, latestFiles, hasRules) {
    const container = document.getElementById('chat-messages');
    const bar = document.createElement('div');
    bar.className = 'message system download-bar';

    const label = document.createElement('div');
    label.className = 'message-label';
    label.textContent = '下载';

    const contentDiv = document.createElement('div');
    contentDiv.className = 'message-content download-buttons';

    const items = [
        { type: 'output', label: '生成结果 (.xlsx)', icon: '\uD83D\uDCCA', has: latestFiles.output_file },
        { type: 'diff', label: '差异对比 (.xlsx)', icon: '\uD83D\uDCCB', has: latestFiles.diff_file },
        { type: 'script', label: '脚本 (.py)', icon: '\uD83D\uDCC4', has: latestFiles.script_file },
    ];

    items.forEach(item => {
        if (!item.has) return;
        const btn = document.createElement('button');
        btn.className = 'download-btn';
        btn.textContent = `${item.icon} ${item.label}`;
        btn.onclick = () => _downloadFile(sessionId, item.type);
        contentDiv.appendChild(btn);
    });

    // 提示词下载
    const promptBtn = document.createElement('button');
    promptBtn.className = 'download-btn prompt-btn';
    promptBtn.textContent = '\uD83D\uDCDD 提示词/上下文';
    promptBtn.onclick = () => _downloadOriginalFile(sessionId, 'prompt');
    contentDiv.appendChild(promptBtn);

    // 规则下载
    if (hasRules) {
        const rulesBtn = document.createElement('button');
        rulesBtn.className = 'download-btn';
        rulesBtn.textContent = '\uD83D\uDCD6 规则文件';
        rulesBtn.onclick = () => _downloadOriginalFile(sessionId, 'rules');
        contentDiv.appendChild(rulesBtn);
    }

    // 最终规则（AI 整理：原始规则 + 多轮对话 + 当前最佳代码）
    const finalBtn2 = document.createElement('button');
    finalBtn2.className = 'download-btn';
    finalBtn2.textContent = '📋 最终规则';
    finalBtn2.title = '根据原始规则+多轮对话+当前最佳代码，整理出可直接用于下次训练的最终规则';
    finalBtn2.onclick = () => _downloadFinalRules(sessionId, finalBtn2);
    contentDiv.appendChild(finalBtn2);

    if (contentDiv.children.length === 0) return;

    bar.appendChild(label);
    bar.appendChild(contentDiv);
    container.appendChild(bar);
    container.scrollTop = container.scrollHeight;
}

async function _downloadOriginalFile(sessionId, category, filename) {
    try {
        let url = `/api/training/chat/sessions/${sessionId}/original-files/${category}`;
        if (filename) url += `?filename=${encodeURIComponent(filename)}`;
        const resp = await AUTH.authFetch(url);
        if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            alert('下载失败: ' + (err.detail || `HTTP ${resp.status}`));
            return;
        }
        const blob = await resp.blob();
        const disposition = resp.headers.get('content-disposition') || '';
        const fnMatch = disposition.match(/filename[*]?=(?:UTF-8'')?["']?([^"';\n]+)/i);
        const dlName = fnMatch ? decodeURIComponent(fnMatch[1]) : (filename || `${category}_${sessionId}`);

        const objUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = objUrl;
        a.download = dlName;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(objUrl);
    } catch (e) {
        alert('下载失败: ' + e.message);
    }
}

// 下载"最终规则"：SSE 流式（两次 AI 调用耗时长，用心跳保活避免 504/502），完成后下载为 .md
async function _downloadFinalRules(sessionId, btn) {
    const oldText = btn ? btn.textContent : '';
    if (btn) { btn.disabled = true; btn.textContent = '⏳ 整理中...'; }
    try {
        const resp = await AUTH.authFetch(`/api/training/chat/sessions/${sessionId}/final-rules`);
        if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            alert('生成最终规则失败: ' + (err.detail || `HTTP ${resp.status}`));
            return;
        }

        const reader = resp.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';
        let done = null;      // 完成事件
        let errMsg = null;    // 错误事件

        while (true) {
            const { done: streamDone, value } = await reader.read();
            if (streamDone) break;
            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop() || '';
            for (const line of lines) {
                if (!line.startsWith('data: ')) continue;   // 忽略心跳 ": heartbeat"
                const jsonStr = line.slice(6).trim();
                if (!jsonStr) continue;
                let event;
                try { event = JSON.parse(jsonStr); } catch (e) { continue; }
                if (event.type === 'status') {
                    if (btn && event.message) btn.textContent = '⏳ ' + event.message;
                } else if (event.type === 'done') {
                    done = event;
                } else if (event.type === 'error') {
                    errMsg = event.message || '未知错误';
                }
            }
        }

        if (errMsg) { alert('生成最终规则失败: ' + errMsg); return; }
        const rules = done && done.rules ? done.rules : '';
        if (!rules.trim()) { alert('未生成有效规则内容'); return; }

        const blob = new Blob([rules], { type: 'text/markdown;charset=utf-8' });
        const objUrl = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = objUrl;
        a.download = (done && done.filename) || `最终规则_${sessionId}.md`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(objUrl);
    } catch (e) {
        alert('生成最终规则失败: ' + e.message);
    } finally {
        if (btn) { btn.disabled = false; btn.textContent = oldText || '📋 最终规则'; }
    }
}

// ==================== UI 状态 ====================
function _setUIStreaming(streaming) {
    _isStreaming = streaming;
    const sendBtn = document.getElementById('send-btn');
    const genBtn = document.getElementById('generate-btn');
    const regenBtn = document.getElementById('regenerate-btn');
    sendBtn.disabled = streaming;
    if (genBtn) genBtn.disabled = streaming;
    if (regenBtn) regenBtn.disabled = streaming;
    if (streaming) {
        sendBtn.textContent = '处理中...';
    } else {
        sendBtn.textContent = '发送';
        // 对话流式状态也要重置
        _chatStreamEl = null;
        _chatStreamBuf = '';
        _chatThinkingBuf = '';
        // 流式结束后，重新应用租户权限灰化
        _applyTenantPermission();
        // 已进入会话时，显示"重新生成"按钮
        if (regenBtn) regenBtn.style.display = _currentSessionId ? '' : 'none';
    }
}

// ==================== 动作按钮 ====================
async function setBestCode() {
    if (!_currentSessionId) return;
    if (!confirm('确定将当前最佳代码设为正式脚本？')) return;

    try {
        const resp = await AUTH.authFetch(`/api/training/chat/sessions/${_currentSessionId}/set-best`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({}),
        });
        if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            throw new Error(err.detail || `HTTP ${resp.status}`);
        }
        const data = await resp.json();
        _addSystemMessage(
            `已设为最佳脚本 (v${data.version}，准确率 ${(data.accuracy * 100).toFixed(1)}%)`,
            'status', { accuracy: data.accuracy }
        );
        _updateChatStatus('completed');
        document.getElementById('btn-set-best').disabled = true;
        document.getElementById('btn-set-best').textContent = '已设置';
        loadSessions();
    } catch (e) {
        alert('设为最佳失败: ' + e.message);
    }
}

function showUploadCode() {
    document.getElementById('upload-code-input').click();
}

// 打开上传代码弹窗（代码必选 + 模板可选）
function openUploadCodeModal() {
    if (!_currentSessionId) { alert('请先选择一个训练会话'); return; }
    const modal = document.getElementById('upload-code-modal');
    if (!modal) {
        // 页面 DOM 仍是旧缓存（无弹窗节点）：回退到旧的直传文件方式，避免点击无反应
        console.warn('[上传代码] 弹窗 DOM 不存在，回退旧方式；请 Ctrl+F5 强制刷新页面');
        const legacy = document.getElementById('upload-code-input');
        if (legacy) { legacy.click(); return; }
        alert('页面未更新，请按 Ctrl+F5 强制刷新后重试');
        return;
    }
    const cf = document.getElementById('uc-code-file');
    const tf = document.getElementById('uc-template-file');
    if (cf) cf.value = '';
    if (tf) tf.value = '';
    // 强制 inline 样式，避免 .modal-overlay 类样式未刷新时弹窗肉眼看不见（与 sheet-picker 一致）
    Object.assign(modal.style, {
        display: 'flex', position: 'fixed', top: '0', left: '0', right: '0', bottom: '0',
        background: 'rgba(0,0,0,0.45)', zIndex: '9999',
        justifyContent: 'center', alignItems: 'flex-start', paddingTop: '8vh', overflowY: 'auto',
    });
    const dialog = modal.querySelector('.modal-dialog');
    if (dialog) Object.assign(dialog.style, {
        background: '#fff', borderRadius: '12px', boxShadow: '0 8px 32px rgba(0,0,0,0.18)',
        width: '90vw', maxWidth: '480px', minHeight: 'auto', padding: '0',
    });
}

function closeUploadCodeModal() {
    const modal = document.getElementById('upload-code-modal');
    if (modal) modal.style.display = 'none';
}

// 提交弹窗：代码文件 + 可选模板文件一起上传
async function submitUploadCode() {
    const codeFile = document.getElementById('uc-code-file').files[0];
    const templateFile = document.getElementById('uc-template-file').files[0];
    if (!codeFile) { alert('请选择代码文件（.py）'); return; }
    closeUploadCodeModal();
    await _doUploadCode(codeFile, templateFile);
}

// 兼容旧的隐藏 input 直传（仅代码）
async function handleUploadCode(event) {
    const file = event.target.files[0];
    event.target.value = '';
    if (!file || !_currentSessionId) return;
    await _doUploadCode(file, null);
}

async function _doUploadCode(codeFile, templateFile) {
    if (!codeFile || !_currentSessionId) return;

    const formData = new FormData();
    formData.append('code_file', codeFile);
    if (templateFile) formData.append('template_file', templateFile);

    try {
        _setUIStreaming(true);
        const tip = templateFile
            ? `正在上传并验证代码文件: ${codeFile.name}（含模板: ${templateFile.name}）`
            : `正在上传并验证代码文件: ${codeFile.name}`;
        _addSystemMessage(tip, 'status');

        const resp = await AUTH.authFetch(`/api/training/chat/sessions/${_currentSessionId}/upload-code`, {
            method: 'POST',
            body: formData,
        });
        if (!resp.ok) {
            const err = await resp.json().catch(() => ({}));
            throw new Error(err.detail || `HTTP ${resp.status}`);
        }
        const data = await resp.json();

        if (data.success) {
            const accPct = (data.accuracy * 100).toFixed(1);
            _currentAccuracy = data.accuracy;
            _currentCode = 'uploaded';  // 标记有代码，使按钮可见
            let msg = `代码上传成功，准确率 ${accPct}%`;
            if (data.accuracy < 1.0 && data.diff_details) {
                msg += '\n\n' + _formatDiffDetails(data.diff_details);
            }
            _addSystemMessage(msg, data.accuracy < 1.0 ? 'diff' : 'status', { accuracy: data.accuracy });
        } else {
            _addSystemMessage(`代码上传执行失败: ${data.error || '未知错误'}`, 'status', { error: data.error });
        }
        _updateActionButtons({ status: 'running', has_script: false });
        loadSessions();
    } catch (e) {
        _addSystemMessage('上传代码失败: ' + e.message, 'status', { error: true });
    } finally {
        _setUIStreaming(false);
    }
}

async function downloadCode() {
    if (!_currentSessionId) return;
    try {
        const resp = await AUTH.authFetch(`/api/training/chat/sessions/${_currentSessionId}/code`);
        if (!resp.ok) {
            alert('获取代码失败');
            return;
        }
        const data = await resp.json();
        if (!data.code) {
            alert('暂无可下载的代码');
            return;
        }
        const blob = new Blob([data.code], { type: 'text/x-python; charset=utf-8' });
        const url = URL.createObjectURL(blob);
        const a = document.createElement('a');
        a.href = url;
        a.download = `script_${_currentTenantId || 'unknown'}.py`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    } catch (e) {
        alert('下载失败: ' + e.message);
    }
}

// ==================== 输入处理 ====================
function handleInputKeydown(event) {
    if (event.key === 'Enter' && !event.shiftKey) {
        event.preventDefault();
        if (event.ctrlKey || event.metaKey) {
            // Ctrl+Enter → 执行修正
            sendMessage('generate');
        } else {
            // Enter → 对话
            sendMessage();
        }
    }
}

function clearChat() {
    createNewSession();
}
