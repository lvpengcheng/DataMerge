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

// ==================== 初始化 ====================
document.addEventListener('DOMContentLoaded', function () {
    AUTH.requireAuth();
    AUTH.renderUserInfo(document.querySelector('header'));
    if (AUTH.isAdmin()) {
        const el = document.getElementById('nav-admin');
        if (el) el.style.display = '';
    }

    // 租户输入框
    const tenantInput = document.getElementById('tenant-input');
    tenantInput.addEventListener('focus', () => _showTenantDropdown());
    tenantInput.addEventListener('input', () => {
        _filterTenantDropdown();
        // 同步手动输入的值到 _currentTenantId
        _currentTenantId = tenantInput.value.trim() || null;
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

    // 加载租户列表
    _loadTenants();
});

// ==================== 租户列表 ====================
let _tenantList = [];

async function _loadTenants() {
    try {
        const resp = await AUTH.authFetch('/api/tenants');
        if (!resp.ok) return;
        const data = await resp.json();
        _tenantList = data.tenants || data || [];
    } catch (e) {
        console.warn('加载租户列表失败:', e);
    }
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
    _clearChatUI();
    _hideActionButtons();
    loadSessions();
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
        // 显示版本号（session_key，格式 {tenant_id}_yyyyMMddHHmmss 或自定义名称）
        const versionLabel = s.session_key || '';
        const time = s.started_at ? new Date(s.started_at).toLocaleString('zh-CN', { month: '2-digit', day: '2-digit', hour: '2-digit', minute: '2-digit' }) : '';
        title.textContent = versionLabel || `${s.mode || 'formula'} - ${time}`;
        title.title = `版本: ${versionLabel} | 会话 #${s.id}`;

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
    _currentSessionId = null;
    _chatMessages = [];
    _currentAccuracy = null;
    _currentCode = null;
    _clearChatUI();
    _hideActionButtons();
    _highlightActiveSession();
    document.getElementById('chat-title').textContent = '新训练';
    document.getElementById('chat-status').textContent = '';
    document.getElementById('chat-status').className = 'chat-status';

    // 提示用户上传文件
    _addSystemMessage('请通过左下角 📎 上传源文件和目标文件，然后发送消息开始训练。');
}

async function selectSession(sessionId) {
    if (_isStreaming) return;
    try {
        const resp = await AUTH.authFetch(`/api/training/chat/sessions/${sessionId}/messages`);
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

        // 合并消息 + 迭代记录，构建完整时间线
        _renderFullHistory(data.messages || [], data.iterations || []);

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
        document.getElementById('chat-title').textContent = `训练 #${data.session.id}`;
        _updateChatStatus(data.session.status);
        _updateActionButtons(data.session);

        _highlightActiveSession();
    } catch (e) {
        console.error('加载会话失败:', e);
    }
}

async function deleteSession(sessionId) {
    if (!confirm('确定删除此训练会话？相关的迭代数据也将被删除。')) return;
    if (_currentSessionId === sessionId) {
        _currentSessionId = null;
        _chatMessages = [];
        _clearChatUI();
        _hideActionButtons();
    }
    _sessions = _sessions.filter(s => s.id !== sessionId);
    _renderSessionList();
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
    container.innerHTML = Array.from(files)
        .map(f => `<span class="file-item">${f.name}</span>`)
        .join('');
}

// ==================== 加密检测 ====================
async function _checkEncryption(filesToCheck) {
    if (!filesToCheck || filesToCheck.length === 0) return;
    try {
        const formData = new FormData();
        filesToCheck.forEach(f => formData.append('files', f));
        const resp = await AUTH.authFetch('/api/files/check-encrypted', { method: 'POST', body: formData });
        if (!resp.ok) return;
        const data = await resp.json();
        const encrypted = data.encrypted_files || [];
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
function _addMessage(role, content, isStreaming) {
    const container = document.getElementById('chat-messages');
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
    container.scrollTop = container.scrollHeight;
    return contentDiv;
}

function _addSystemMessage(content, msgType, metadata) {
    const container = document.getElementById('chat-messages');
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
    container.scrollTop = container.scrollHeight;
    return contentDiv;
}

function _updateStreamingMessage(contentDiv, text) {
    contentDiv.innerHTML = _renderMarkdown(text);
    const container = document.getElementById('chat-messages');
    container.scrollTop = container.scrollHeight;
}

function _finishStreamingMessage(contentDiv) {
    contentDiv.classList.remove('streaming-cursor');
}

// 代码流式追加：将 AI 生成的代码片段逐步显示
let _codeStreamEl = null;
let _codeStreamBuf = '';

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
    // 渲染为 markdown 代码块
    _codeStreamEl.innerHTML = _renderMarkdown('```python\n' + _codeStreamBuf + '\n```');
    container.scrollTop = container.scrollHeight;
}

function _finishCodeStream() {
    if (_codeStreamEl) {
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

    // action: undefined/'chat' = 对话讨论, 'generate' = 执行代码修正
    action = action || 'chat';

    // 已在会话中 → 直接发消息，无需再选租户
    if (_currentSessionId) {
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

    const formData = new FormData();
    formData.append('tenant_id', _currentTenantId);
    formData.append('ai_provider', aiProvider);
    formData.append('mode', mode);
    if (salaryMonth) formData.append('salary_year_month', salaryMonth);
    if (standardHours) formData.append('monthly_standard_hours', standardHours);
    const manualHeaders = document.getElementById('manual-headers').value.trim();
    if (manualHeaders) formData.append('manual_headers', manualHeaders);
    const multiSheetSource = document.getElementById('multi-sheet-source').checked;
    if (multiSheetSource) formData.append('multi_sheet_source', 'true');
    if (_currentSessionId) formData.append('session_id', _currentSessionId);
    if (Object.keys(_filePasswordsMap).length > 0) {
        formData.append('file_passwords', JSON.stringify(_filePasswordsMap));
        console.log('[训练] file_passwords:', JSON.stringify(_filePasswordsMap));
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

// ==================== SSE ====================
async function _fetchTrainingSSE(url, options) {
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
                    _handleSSEEvent(event);
                } catch (parseErr) {
                    if (parseErr.message && !parseErr.message.includes('JSON')) throw parseErr;
                }
            }
        }
    } catch (e) {
        console.error('SSE error:', e);
        _addSystemMessage('请求失败: ' + e.message, 'status', { error: true });
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
            }
            _chatStreamBuf += event.content;
            _updateStreamingMessage(_chatStreamEl, _chatStreamBuf);
            break;

        case 'chat_done':
            // AI 对话完成
            if (_chatStreamEl) {
                _finishStreamingMessage(_chatStreamEl);
                // 用最终完整内容重新渲染（确保 markdown 完整）
                _chatStreamEl.innerHTML = _renderMarkdown(_chatStreamBuf || event.content);
                _chatStreamEl = null;
                _chatStreamBuf = '';
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
function _renderFullHistory(messages, iterations) {
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
            _addMessage('user', msg.content);
        } else if (msg.role === 'assistant') {
            _addMessage('assistant', msg.content);
        } else if (msg.role === 'system') {
            _addSystemMessage(msg.content, msg.msg_type, msg.metadata);
        }

        // 如果这条消息是迭代结果消息，在其后追加代码摘要
        if (msgIter && iterMap[msgIter]) {
            const it = iterMap[msgIter];
            if (it.generated_code) {
                _renderIterationCodeSummary(it);
            }
            lastRenderedIter = msgIter;
        }
    });

    // 如果有迭代没有对应的消息（例如首轮训练后没保存足够消息），补充渲染
    iterations.forEach(it => {
        if (!mentionedIters.has(it.iteration_num)) {
            _renderOrphanIteration(it);
        }
    });
}

/**
 * 渲染迭代代码摘要（折叠式，可展开查看代码）
 */
function _renderIterationCodeSummary(iteration) {
    const container = document.getElementById('chat-messages');
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
function _renderOrphanIteration(iteration) {
    const container = document.getElementById('chat-messages');

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
    _addSystemMessage(resultText, acc != null ? 'status' : 'status', metadata);

    // 如果有代码，显示折叠摘要
    if (iteration.generated_code) {
        _renderIterationCodeSummary(iteration);
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

// ==================== UI 状态 ====================
function _setUIStreaming(streaming) {
    _isStreaming = streaming;
    const sendBtn = document.getElementById('send-btn');
    const genBtn = document.getElementById('generate-btn');
    sendBtn.disabled = streaming;
    if (genBtn) genBtn.disabled = streaming;
    if (streaming) {
        sendBtn.textContent = '处理中...';
    } else {
        sendBtn.textContent = '发送';
        // 对话流式状态也要重置
        _chatStreamEl = null;
        _chatStreamBuf = '';
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

async function handleUploadCode(event) {
    const file = event.target.files[0];
    if (!file || !_currentSessionId) return;

    const formData = new FormData();
    formData.append('code_file', file);

    try {
        _setUIStreaming(true);
        _addSystemMessage(`正在上传并验证代码文件: ${file.name}`, 'status');

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
        event.target.value = '';
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
