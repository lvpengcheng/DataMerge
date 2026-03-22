/**
 * 规则整理页面 JS
 */

// ==================== 状态管理 ====================

// 对话历史 (OpenAI 格式 messages)
let _chatMessages = [];
// 最后一条 AI 回复的完整内容（用于下载）
let _lastAIContent = '';
// 是否正在流式接收
let _isStreaming = false;
// 当前 EventSource
let _currentEventSource = null;
// 文件密码映射
let _filePasswordsMap = {};

// ==================== 初始化 ====================

document.addEventListener('DOMContentLoaded', function () {
    AUTH.requireAuth();
    AUTH.renderUserInfo(document.querySelector('header'));
    if (AUTH.isAdmin()) {
        const adminNav = document.getElementById('nav-admin');
        if (adminNav) adminNav.style.display = '';
    }

    // 文件选择监听：只检测当次上传的文件
    document.getElementById('source-files').addEventListener('change', function () {
        _renderFileList('source-file-list', this.files);
        _checkEncryption(Array.from(this.files));
    });
    document.getElementById('target-file').addEventListener('change', function () {
        _renderFileList('target-file-list', this.files);
        _checkEncryption(Array.from(this.files));
    });
    document.getElementById('design-docs').addEventListener('change', function () {
        _renderFileList('design-doc-list', this.files);
    });
});

// ==================== 文件列表显示 ====================

function _renderFileList(containerId, files) {
    const container = document.getElementById(containerId);
    if (!files || files.length === 0) {
        container.innerHTML = '';
        return;
    }
    container.innerHTML = Array.from(files)
        .map(f => `<span class="file-item">${f.name}</span>`)
        .join(' ');
}

// ==================== 加密检测（弹窗方式） ====================

/**
 * 弹出密码输入对话框，为加密文件输入密码
 * @param {string[]} encryptedFiles - 加密文件名列表
 * @returns {Promise<object|null>} 文件名→密码映射，取消返回null
 */
function _promptFilePasswords(encryptedFiles) {
    return new Promise((resolve) => {
        const inputs = encryptedFiles.map((name, i) =>
            `<div style="margin-bottom:10px;">
                <label style="display:block;font-size:13px;margin-bottom:4px;color:#333;">
                    <span style="color:#e65100;">🔒</span> ${name}
                </label>
                <input id="_enc_pwd_${i}" type="password" placeholder="请输入打开密码"
                    style="width:100%;padding:8px;border:1px solid #ddd;border-radius:4px;box-sizing:border-box;">
            </div>`
        ).join('');

        const overlay = document.createElement('div');
        overlay.style.cssText = 'position:fixed;top:0;left:0;right:0;bottom:0;background:rgba(0,0,0,0.5);z-index:9999;display:flex;align-items:center;justify-content:center;';
        overlay.innerHTML = `
            <div style="background:#fff;border-radius:10px;padding:24px;width:400px;max-width:90vw;box-shadow:0 4px 20px rgba(0,0,0,0.2);">
                <h3 style="margin:0 0 6px;font-size:16px;">检测到加密文件</h3>
                <p style="margin:0 0 16px;font-size:13px;color:#666;">以下文件有密码保护，请输入密码后继续：</p>
                ${inputs}
                <div style="display:flex;gap:10px;justify-content:flex-end;margin-top:16px;">
                    <button id="_enc_cancel" style="padding:8px 20px;border:1px solid #ddd;border-radius:4px;background:#fff;cursor:pointer;">取消</button>
                    <button id="_enc_confirm" style="padding:8px 20px;border:none;border-radius:4px;background:#1976d2;color:#fff;cursor:pointer;">确认解锁</button>
                </div>
            </div>`;
        document.body.appendChild(overlay);

        document.getElementById('_enc_cancel').onclick = () => {
            document.body.removeChild(overlay);
            resolve(null);
        };
        document.getElementById('_enc_confirm').onclick = () => {
            const passwords = {};
            encryptedFiles.forEach((name, i) => {
                const pwd = document.getElementById(`_enc_pwd_${i}`).value;
                if (pwd) passwords[name] = pwd;
            });
            const missing = encryptedFiles.filter(name => !passwords[name]);
            if (missing.length > 0) {
                alert('请为所有加密文件输入密码：\n' + missing.join('\n'));
                return;
            }
            document.body.removeChild(overlay);
            resolve(passwords);
        };

        setTimeout(() => document.getElementById('_enc_pwd_0')?.focus(), 100);
    });
}

/**
 * 文件选择后，调用服务端检测加密，有加密则立即弹窗
 * @param {File[]} filesToCheck - 当次选择的文件数组
 */
async function _checkEncryption(filesToCheck) {
    if (!filesToCheck || filesToCheck.length === 0) return;

    const btn = document.getElementById('organize-btn');
    btn.disabled = true;
    btn.textContent = '检测文件中...';

    try {
        const formData = new FormData();
        filesToCheck.forEach(f => formData.append('files', f));

        const resp = await AUTH.authFetch('/api/files/check-encrypted', {
            method: 'POST',
            body: formData,
        });
        if (!resp.ok) return;
        const data = await resp.json();
        const encrypted = data.encrypted_files || [];

        if (encrypted.length === 0) {
            return;
        }

        // 有加密文件，弹窗输入密码
        const passwords = await _promptFilePasswords(encrypted);
        if (passwords) {
            _filePasswordsMap = { ..._filePasswordsMap, ...passwords };
        }
    } catch (e) {
        console.warn('加密检测失败:', e);
    } finally {
        btn.disabled = false;
        btn.textContent = '开始整理';
    }
}

// ==================== Markdown 简易渲染 ====================

function _renderMarkdown(text) {
    // 转义 HTML
    let html = text
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;');

    // 代码块 ```
    html = html.replace(/```(\w*)\n([\s\S]*?)```/g, function (_, lang, code) {
        return `<pre><code>${code}</code></pre>`;
    });

    // 行内代码
    html = html.replace(/`([^`]+)`/g, '<code>$1</code>');

    // 标题
    html = html.replace(/^### (.+)$/gm, '<h3>$1</h3>');
    html = html.replace(/^## (.+)$/gm, '<h2>$1</h2>');
    html = html.replace(/^# (.+)$/gm, '<h1>$1</h1>');

    // 粗体 / 斜体
    html = html.replace(/\*\*(.+?)\*\*/g, '<strong>$1</strong>');
    html = html.replace(/\*(.+?)\*/g, '<em>$1</em>');

    // 无序列表
    html = html.replace(/^- (.+)$/gm, '<li>$1</li>');
    html = html.replace(/(<li>.*<\/li>\n?)+/g, function (match) {
        return '<ul>' + match + '</ul>';
    });

    // 表格 (简单支持)
    html = html.replace(/^\|(.+)\|$/gm, function (line) {
        const cells = line.split('|').filter(c => c.trim() !== '');
        // 跳过分隔行
        if (cells.every(c => /^[\s\-:]+$/.test(c))) return '';
        const tag = 'td';
        const row = cells.map(c => `<${tag}>${c.trim()}</${tag}>`).join('');
        return `<tr>${row}</tr>`;
    });
    html = html.replace(/(<tr>.*<\/tr>\n?)+/g, function (match) {
        return '<table>' + match + '</table>';
    });

    // 段落：连续非空行
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

// ==================== 开始整理 ====================

function startOrganize() {
    const sourceFiles = document.getElementById('source-files').files;
    const targetFile = document.getElementById('target-file').files[0];

    if (!sourceFiles.length || !targetFile) {
        alert('请先选择源文件和目标文件');
        return;
    }

    if (_isStreaming) {
        alert('正在处理中，请等待完成');
        return;
    }

    const designDocs = document.getElementById('design-docs').files;
    const aiProvider = document.getElementById('ai-provider').value;

    // 构建 FormData
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

    // 重置对话历史，开始新的整理
    _chatMessages = [
        { role: 'system', content: _getSystemPromptSummary() },
    ];

    // 添加系统消息提示
    _addMessage('assistant', '正在分析文件结构，请稍候...\n\n这可能需要 1-2 分钟。', true);

    _setUIStreaming(true);

    // 使用 fetch 发起 SSE 请求（因为 EventSource 不支持 POST）
    _fetchSSE('/api/rules/organize/stream', {
        method: 'POST',
        body: formData,
    });
}

// ==================== 发送追问消息 ====================

function sendMessage() {
    const input = document.getElementById('chat-input');
    const text = input.value.trim();
    if (!text || _isStreaming) return;

    if (_chatMessages.length === 0) {
        alert('请先点击"开始整理"生成规则');
        return;
    }

    // 添加用户消息
    _addMessage('user', text);
    _chatMessages.push({ role: 'user', content: text });
    input.value = '';

    _setUIStreaming(true);

    // 流式发送对话
    _fetchSSE('/api/rules/chat', {
        method: 'POST',
        headers: { 'Content-Type': 'application/json' },
        body: JSON.stringify({
            messages: _chatMessages,
            ai_provider: document.getElementById('ai-provider').value,
        }),
    });
}

// ==================== SSE 流式请求 ====================

async function _fetchSSE(url, options) {
    let streamingDiv = null;
    let fullContent = '';

    try {
        // 注入 Authorization header（使用 AUTH 工具）
        if (!options.headers) options.headers = {};
        const authHeaders = AUTH.getAuthHeaders();
        Object.assign(options.headers, authHeaders);

        const response = await fetch(url, options);

        if (!response.ok) {
            let errMsg = `HTTP ${response.status}`;
            try {
                const errData = await response.json();
                errMsg = errData.detail || errMsg;
            } catch (e) {}
            throw new Error(errMsg);
        }

        // 移除"正在分析"的占位消息（如果存在）
        _removeLastAssistantMessage();

        // 创建流式输出 div
        streamingDiv = _addMessage('assistant', '', true);

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });

            // 解析 SSE 行
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
                    } else if (data.type === 'error') {
                        throw new Error(data.message);
                    }
                } catch (parseErr) {
                    if (parseErr.message && !parseErr.message.includes('JSON')) {
                        throw parseErr;
                    }
                }
            }
        }
    } catch (e) {
        console.error('SSE 请求失败:', e);
        if (streamingDiv) {
            _finishStreamingMessage(streamingDiv);
        }
        _addMessage('assistant', '请求失败: ' + e.message);
    } finally {
        if (streamingDiv) {
            _finishStreamingMessage(streamingDiv);
        }

        // 记录 AI 回复到对话历史
        if (fullContent) {
            _chatMessages.push({ role: 'assistant', content: fullContent });
            _lastAIContent = fullContent;
            document.getElementById('download-btn').style.display = 'inline-block';
        }

        _setUIStreaming(false);
    }
}

function _removeLastAssistantMessage() {
    const container = document.getElementById('chat-messages');
    const messages = container.querySelectorAll('.message.assistant');
    if (messages.length > 0) {
        const last = messages[messages.length - 1];
        const content = last.querySelector('.message-content');
        if (content && content.classList.contains('streaming-cursor')) {
            container.removeChild(last);
        }
    }
}

// ==================== UI 状态控制 ====================

function _setUIStreaming(streaming) {
    _isStreaming = streaming;
    document.getElementById('organize-btn').disabled = streaming;
    document.getElementById('send-btn').disabled = streaming;

    if (streaming) {
        document.getElementById('organize-btn').textContent = '整理中...';
    } else {
        document.getElementById('organize-btn').textContent = '开始整理';
        document.getElementById('send-btn').disabled = false;
    }
}

// ==================== 下载 / 清空 ====================

function downloadRules() {
    if (!_lastAIContent) {
        alert('暂无可下载的规则内容');
        return;
    }
    const blob = new Blob([_lastAIContent], { type: 'text/markdown; charset=utf-8' });
    const url = URL.createObjectURL(blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = 'rules.md';
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
    URL.revokeObjectURL(url);
}

function clearChat() {
    _chatMessages = [];
    _lastAIContent = '';

    const container = document.getElementById('chat-messages');
    container.innerHTML = `
        <div class="chat-placeholder" id="chat-placeholder">
            <div class="placeholder-icon">📋</div>
            <p>上传源文件和目标文件后，点击"开始整理"</p>
            <p class="hint">AI 将分析文件结构和设计文档，生成结构化的规则文件</p>
        </div>
    `;

    document.getElementById('download-btn').style.display = 'none';
}

// ==================== 辅助函数 ====================

function handleInputKeydown(event) {
    // Ctrl+Enter 或 Cmd+Enter 发送
    if ((event.ctrlKey || event.metaKey) && event.key === 'Enter') {
        event.preventDefault();
        sendMessage();
    }
}

function _getSystemPromptSummary() {
    // 返回系统提示的简要描述，用于对话历史上下文
    return '你是一位专业的数据处理规则分析师。用户已经上传了源文件和目标文件的结构信息。请根据用户的追问继续调整和完善规则文档。保持 Markdown 格式输出。';
}
