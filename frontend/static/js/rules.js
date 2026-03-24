/**
 * 规则整理页面 JS — 会话记忆 + 附件弹窗
 */

// ==================== 状态 ====================
let _chatMessages = [];
let _lastAIContent = '';
let _isStreaming = false;
let _currentSessionId = null;
let _filePasswordsMap = {};
let _sessions = [];

// ==================== 初始化 ====================
document.addEventListener('DOMContentLoaded', function () {
    AUTH.requireAuth();
    AUTH.renderUserInfo(document.querySelector('header'));
    if (AUTH.isAdmin()) {
        const el = document.getElementById('nav-admin');
        if (el) el.style.display = '';
    }

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
    document.getElementById('design-docs').addEventListener('change', function () {
        _renderFileList('design-doc-list', this.files);
        _updateAttachBadge();
    });

    // 点击外部关闭弹窗
    document.addEventListener('click', function (e) {
        const popover = document.getElementById('attach-popover');
        const btn = document.getElementById('attach-btn');
        if (popover.style.display !== 'none'
            && !popover.contains(e.target)
            && !btn.contains(e.target)) {
            popover.style.display = 'none';
        }
    });

    loadSessions();
});

// ==================== 会话管理 ====================
async function loadSessions() {
    try {
        const resp = await AUTH.authFetch('/api/rules/sessions');
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
        if (emptyHint) emptyHint.style.display = '';
        return;
    }
    if (emptyHint) emptyHint.style.display = 'none';

    _sessions.forEach(s => {
        const div = document.createElement('div');
        div.className = 'session-item' + (s.id === _currentSessionId ? ' active' : '');
        div.onclick = () => selectSession(s.id);

        const title = document.createElement('span');
        title.className = 'session-item-title';
        title.textContent = s.title;
        title.title = s.title;

        const delBtn = document.createElement('button');
        delBtn.className = 'session-item-delete';
        delBtn.innerHTML = '✕';
        delBtn.title = '删除';
        delBtn.onclick = (e) => { e.stopPropagation(); deleteSession(s.id); };

        div.appendChild(title);
        div.appendChild(delBtn);
        container.appendChild(div);
    });
}

function createNewSession() {
    _currentSessionId = null;
    _chatMessages = [];
    _lastAIContent = '';
    _clearChatUI();
    _highlightActiveSession();
    document.getElementById('chat-title').textContent = '';
}

async function selectSession(sessionId) {
    if (_isStreaming) return;
    try {
        const resp = await AUTH.authFetch(`/api/rules/sessions/${sessionId}`);
        if (!resp.ok) return;
        const data = await resp.json();

        _currentSessionId = data.id;
        _chatMessages = data.messages || [];
        _lastAIContent = data.final_result || '';

        _clearChatUI();
        _chatMessages.forEach(msg => {
            if (msg.role === 'user' || msg.role === 'assistant') {
                _addMessage(msg.role, msg.content);
            }
        });

        if (_lastAIContent) {
            document.getElementById('download-btn').style.display = '';
        }

        document.getElementById('chat-title').textContent = data.title || '';
        _highlightActiveSession();
    } catch (e) {
        console.error('加载会话失败:', e);
    }
}

async function deleteSession(sessionId) {
    if (!confirm('确定删除此对话？')) return;
    try {
        await AUTH.authFetch(`/api/rules/sessions/${sessionId}`, { method: 'DELETE' });
        if (_currentSessionId === sessionId) createNewSession();
        await loadSessions();
    } catch (e) {
        console.error('删除失败:', e);
    }
}

function _highlightActiveSession() {
    document.querySelectorAll('.session-item').forEach(el => el.classList.remove('active'));
    if (_currentSessionId) {
        const idx = _sessions.findIndex(s => s.id === _currentSessionId);
        const items = document.querySelectorAll('.session-item');
        if (idx >= 0 && items[idx]) items[idx].classList.add('active');
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
function _promptFilePasswords(encryptedFiles) {
    return new Promise((resolve) => {
        const inputs = encryptedFiles.map((name, i) =>
            `<div style="margin-bottom:10px;">
                <label style="display:block;font-size:13px;margin-bottom:4px;color:#333;">🔒 ${name}</label>
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
                    <button id="_enc_cancel" style="padding:8px 20px;border:1.5px solid #d1d5db;border-radius:8px;background:#fff;cursor:pointer;font-size:13px;color:#374151;">取消</button>
                    <button id="_enc_confirm" style="padding:8px 20px;border:none;border-radius:8px;background:#6366f1;color:#fff;cursor:pointer;font-size:13px;">确认</button>
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

function _updateStreamingMessage(contentDiv, text) {
    contentDiv.innerHTML = _renderMarkdown(text);
    const container = document.getElementById('chat-messages');
    container.scrollTop = container.scrollHeight;
}

function _finishStreamingMessage(contentDiv) {
    contentDiv.classList.remove('streaming-cursor');
}

function _clearChatUI() {
    const container = document.getElementById('chat-messages');
    container.innerHTML = `
        <div class="chat-placeholder" id="chat-placeholder">
            <p>点击下方 📎 上传文件，输入需求后发送</p>
        </div>
    `;
    document.getElementById('download-btn').style.display = 'none';
}

// ==================== 发送 ====================
function sendMessage() {
    const input = document.getElementById('chat-input');
    const text = input.value.trim();
    if (!text || _isStreaming) return;

    const sourceFiles = document.getElementById('source-files').files;
    const targetFile = document.getElementById('target-file').files[0];
    const isFirstOrganize = sourceFiles.length > 0 && targetFile && _chatMessages.length === 0;

    if (isFirstOrganize) {
        _startOrganizeWithMessage(text, sourceFiles, targetFile);
    } else if (_chatMessages.length > 0 || _currentSessionId) {
        _sendChatMessage(text);
    } else {
        alert('请先通过左下角📎上传源文件和目标文件');
        return;
    }
    input.value = '';
    input.style.height = '';
}

function _startOrganizeWithMessage(userText, sourceFiles, targetFile) {
    const designDocs = document.getElementById('design-docs').files;
    const aiProvider = document.getElementById('ai-provider').value;

    const formData = new FormData();
    Array.from(sourceFiles).forEach(f => formData.append('source_files', f));
    formData.append('target_file', targetFile);
    if (designDocs.length > 0) {
        Array.from(designDocs).forEach(f => formData.append('design_docs', f));
    }
    formData.append('ai_provider', aiProvider);
    if (Object.keys(_filePasswordsMap).length > 0) {
        formData.append('file_passwords', JSON.stringify(_filePasswordsMap));
    }
    if (_currentSessionId) {
        formData.append('session_id', _currentSessionId);
    }

    _chatMessages = [{ role: 'system', content: _getSystemPromptSummary() }];
    _addMessage('user', userText);
    _chatMessages.push({ role: 'user', content: userText });

    document.getElementById('attach-popover').style.display = 'none';
    _setUIStreaming(true);

    _fetchSSE('/api/rules/organize/stream', { method: 'POST', body: formData });
}

function startOrganize() {
    const sourceFiles = document.getElementById('source-files').files;
    const targetFile = document.getElementById('target-file').files[0];
    if (!sourceFiles.length || !targetFile) {
        alert('请先通过📎选择源文件和目标文件');
        return;
    }
    if (_isStreaming) return;
    const input = document.getElementById('chat-input');
    if (!input.value.trim()) input.value = '请根据文件整理数据处理规则';
    sendMessage();
}

function _sendChatMessage(text) {
    _addMessage('user', text);
    _chatMessages.push({ role: 'user', content: text });
    _setUIStreaming(true);

    _fetchSSE('/api/rules/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            messages: _chatMessages,
            ai_provider: document.getElementById('ai-provider').value,
            session_id: _currentSessionId,
        }),
    });
}

// ==================== SSE ====================
async function _fetchSSE(url, options) {
    let streamingDiv = null;
    let fullContent = '';

    try {
        if (!options.headers) options.headers = {};
        Object.assign(options.headers, AUTH.getAuthHeaders());

        const response = await fetch(url, options);
        if (!response.ok) {
            let errMsg = `HTTP ${response.status}`;
            try { errMsg = (await response.json()).detail || errMsg; } catch (e) {}
            throw new Error(errMsg);
        }

        _removeLastAssistantMessage();
        streamingDiv = _addMessage('assistant', '', true);

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
                    const data = JSON.parse(jsonStr);
                    if (data.type === 'chunk') {
                        fullContent += data.content;
                        _updateStreamingMessage(streamingDiv, fullContent);
                    } else if (data.type === 'complete') {
                        fullContent = data.content;
                        _updateStreamingMessage(streamingDiv, fullContent);
                        _finishStreamingMessage(streamingDiv);
                        if (data.session_id) {
                            _currentSessionId = data.session_id;
                            loadSessions();
                        }
                    } else if (data.type === 'error') {
                        throw new Error(data.message);
                    }
                } catch (parseErr) {
                    if (parseErr.message && !parseErr.message.includes('JSON')) throw parseErr;
                }
            }
        }
    } catch (e) {
        console.error('SSE error:', e);
        if (streamingDiv) _finishStreamingMessage(streamingDiv);
        _addMessage('assistant', '请求失败: ' + e.message);
    } finally {
        if (streamingDiv) _finishStreamingMessage(streamingDiv);
        if (fullContent) {
            _chatMessages.push({ role: 'assistant', content: fullContent });
            _lastAIContent = fullContent;
            document.getElementById('download-btn').style.display = '';
        }
        _setUIStreaming(false);
    }
}

function _removeLastAssistantMessage() {
    const container = document.getElementById('chat-messages');
    const msgs = container.querySelectorAll('.message.assistant');
    if (msgs.length > 0) {
        const last = msgs[msgs.length - 1];
        const c = last.querySelector('.message-content');
        if (c && c.classList.contains('streaming-cursor')) container.removeChild(last);
    }
}

// ==================== UI 状态 ====================
function _setUIStreaming(streaming) {
    _isStreaming = streaming;
    document.getElementById('send-btn').disabled = streaming;
}

// ==================== 下载 / 辅助 ====================
function downloadRules() {
    if (!_lastAIContent) { alert('暂无可下载的规则内容'); return; }
    const blob = new Blob([_lastAIContent], { type: 'text/markdown; charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url; a.download = 'rules.md';
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function clearChat() { createNewSession(); }

function handleInputKeydown(event) {
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
        event.preventDefault();
        sendMessage();
    }
}

function _getSystemPromptSummary() {
    return '你是一位专业的数据处理规则分析师。用户已经上传了源文件和目标文件的结构信息。请根据用户的追问继续调整和完善规则文档。保持 Markdown 格式输出。';
}
