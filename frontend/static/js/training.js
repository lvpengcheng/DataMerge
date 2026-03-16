// 全局变量
let trainingStartTime = null;
let timerInterval = null;
let eventSource = null;
let tenantListData = [];
let trainingHistoryData = {};
let codeStreamBuffer = '';

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    initializeForm();
    loadTenantList();
    loadTrainingHistory();

    // 输入框失焦时加载状态
    const input = document.getElementById('tenant-input');
    if (input) {
        input.addEventListener('blur', function() {
            setTimeout(() => {
                document.getElementById('tenant-dropdown').style.display = 'none';
                if (input.value.trim()) {
                    onTenantSelected(input.value.trim());
                }
            }, 200);
        });

        // 添加focus和input事件
        input.addEventListener('focus', showTenantDropdown);
        input.addEventListener('input', filterTenantDropdown);
    }

    // 监听文件选择
    const sourceFiles = document.getElementById('source-files');
    const targetFile = document.getElementById('target-file');
    const ruleFiles = document.getElementById('rule-files');

    console.log('File inputs found:', {
        sourceFiles: !!sourceFiles,
        targetFile: !!targetFile,
        ruleFiles: !!ruleFiles
    });

    if (sourceFiles) {
        sourceFiles.addEventListener('change', () => {
            console.log('source-files changed');
            updateFileList('source-files', 'source-file-list');
        });
    }

    if (targetFile) {
        targetFile.addEventListener('change', () => {
            console.log('target-file changed');
            updateFileList('target-file', 'target-file-list');
        });
    }

    if (ruleFiles) {
        ruleFiles.addEventListener('change', () => {
            console.log('rule-files changed');
            updateFileList('rule-files', 'rule-file-list');
        });
    }
});

// 初始化表单
function initializeForm() {
    const form = document.getElementById('training-form');
    const sourceFiles = document.getElementById('source-files');
    const targetFile = document.getElementById('target-file');
    const ruleFiles = document.getElementById('rule-files');

    sourceFiles.addEventListener('change', function(e) {
        displayFileList(e.target.files, 'source-file-list');
    });

    targetFile.addEventListener('change', function(e) {
        displayFileList(e.target.files, 'target-file-list');
    });

    ruleFiles.addEventListener('change', function(e) {
        displayFileList(e.target.files, 'rule-file-list');
    });

    form.addEventListener('submit', function(e) {
        e.preventDefault();
        startTraining();
    });
}

function displayFileList(files, containerId) {
    const container = document.getElementById(containerId);
    container.innerHTML = '';
    Array.from(files).forEach(file => {
        const div = document.createElement('div');
        div.textContent = `${file.name} (${formatFileSize(file.size)})`;
        container.appendChild(div);
    });
}

function formatFileSize(bytes) {
    if (bytes < 1024) return bytes + ' B';
    if (bytes < 1024 * 1024) return (bytes / 1024).toFixed(2) + ' KB';
    return (bytes / (1024 * 1024)).toFixed(2) + ' MB';
}

// ==================== 租户下拉框 ====================

async function loadTenantList() {
    try {
        const response = await fetch('/api/tenants');
        if (!response.ok) return;
        const data = await response.json();
        tenantListData = data.tenants || [];
        renderTenantDropdown(tenantListData);
    } catch (e) {
        console.error('加载租户列表失败:', e);
    }
}

function renderTenantDropdown(tenants) {
    const dropdown = document.getElementById('tenant-dropdown');
    dropdown.innerHTML = '';
    if (tenants.length === 0) {
        dropdown.innerHTML = '<div class="combo-item disabled">暂无租户</div>';
        return;
    }
    tenants.forEach(tenant => {
        const item = document.createElement('div');
        item.className = 'combo-item';
        const score = tenant.best_score !== null
            ? `<span class="combo-score">${(tenant.best_score * 100).toFixed(1)}%</span>`
            : '<span class="combo-score untrained">未训练</span>';
        item.innerHTML = `<span class="combo-id">${tenant.tenant_id}</span>${score}`;
        item.addEventListener('mousedown', function(e) {
            e.preventDefault();
            document.getElementById('tenant-input').value = tenant.tenant_id;
            dropdown.style.display = 'none';
            onTenantSelected(tenant.tenant_id);
        });
        dropdown.appendChild(item);
    });
}

function showTenantDropdown() {
    filterTenantDropdown();
    document.getElementById('tenant-dropdown').style.display = 'block';
}

function filterTenantDropdown() {
    const input = document.getElementById('tenant-input').value.trim().toLowerCase();
    const filtered = input
        ? tenantListData.filter(t => t.tenant_id.toLowerCase().includes(input))
        : tenantListData;
    renderTenantDropdown(filtered);
    document.getElementById('tenant-dropdown').style.display = 'block';
}

document.addEventListener('click', function(e) {
    const combo = document.getElementById('tenant-combo');
    const dropdown = document.getElementById('tenant-dropdown');
    if (combo && dropdown && !combo.contains(e.target)) {
        dropdown.style.display = 'none';
    }
});

// 选中租户后
function onTenantSelected(tenantId) {
    if (!tenantId) return;
    document.getElementById('tenant-id').textContent = tenantId;
    loadTenantStatus();
    loadTenantScripts(tenantId);

    // 重置右侧
    resetProgress();
    clearResultSection();

    // 如果有训练历史，显示结果
    const tenantHistory = trainingHistoryData[tenantId];
    if (tenantHistory && tenantHistory.sessions && tenantHistory.sessions.length > 0) {
        showHistoryResult(tenantId, tenantHistory);
    }
}

// ==================== 训练ID下拉 ====================

async function loadTenantScripts(tenantId) {
    const group = document.getElementById('script-selector-group');
    const select = document.getElementById('script-select');

    try {
        const response = await fetch(`/api/tenant-scripts/${tenantId}`);
        if (!response.ok) { group.style.display = 'none'; return; }
        const data = await response.json();
        const scripts = data.scripts || [];

        if (scripts.length === 0) {
            group.style.display = 'none';
            return;
        }

        select.innerHTML = '';
        scripts.forEach(s => {
            const opt = document.createElement('option');
            opt.value = s.script_id;
            const scoreText = s.score != null ? ` (${(s.score * 100).toFixed(1)}%)` : '';
            const activeText = s.is_active ? ' [当前]' : '';
            const shortId = s.script_id.length > 20 ? s.script_id.substring(0, 20) + '...' : s.script_id;
            opt.textContent = `${shortId}${scoreText}${activeText}`;
            if (s.is_active) opt.selected = true;
            select.appendChild(opt);
        });

        group.style.display = 'block';
    } catch (e) {
        console.error('加载训练ID失败:', e);
        group.style.display = 'none';
    }
}

// ==================== 训练状态面板 ====================

async function loadTenantStatus() {
    const tenantId = document.getElementById('tenant-input').value.trim();
    const card = document.getElementById('tenant-status-card');
    if (!tenantId) { card.style.display = 'none'; return; }

    try {
        const response = await fetch(`/api/training-status/${tenantId}`);
        if (!response.ok) {
            card.style.display = 'block';
            document.getElementById('tenant-training-status').textContent = '未找到';
            document.getElementById('tenant-best-score').textContent = '-';
            document.getElementById('tenant-script-id').textContent = '-';
            document.getElementById('tenant-last-training').textContent = '-';
            return;
        }

        const data = await response.json();
        card.style.display = 'block';

        const statusEl = document.getElementById('tenant-training-status');
        if (data.is_training) {
            statusEl.textContent = '训练中...';
            statusEl.style.color = '#ffc107';
        } else if (data.has_trained) {
            statusEl.textContent = '已完成';
            statusEl.style.color = '#28a745';
        } else {
            statusEl.textContent = '未训练';
            statusEl.style.color = '#6c757d';
        }

        const scoreEl = document.getElementById('tenant-best-score');
        if (data.best_score != null) {
            const pct = (data.best_score * 100).toFixed(2) + '%';
            scoreEl.textContent = pct;
            scoreEl.style.color = data.best_score >= 0.95 ? '#28a745' : data.best_score >= 0.8 ? '#ffc107' : '#dc3545';
        } else {
            scoreEl.textContent = '-';
        }

        document.getElementById('tenant-script-id').textContent = data.script_id || '-';

        if (data.last_training && data.last_training.completed) {
            document.getElementById('tenant-last-training').textContent =
                new Date(data.last_training.completed).toLocaleString('zh-CN');
        } else {
            document.getElementById('tenant-last-training').textContent = '-';
        }
    } catch (e) {
        console.error('查询训练状态失败:', e);
    }
}

// ==================== 训练历史 ====================

async function loadTrainingHistory() {
    try {
        const response = await fetch('/api/training-history');
        if (!response.ok) return;
        const data = await response.json();
        trainingHistoryData = data.history || {};
    } catch (e) {
        console.error('加载训练历史失败:', e);
    }
}

// 显示已训练租户的结果
function showHistoryResult(tenantId, tenantHistory) {
    const sessions = tenantHistory.sessions;
    const lastSession = sessions[sessions.length - 1];
    const files = lastSession.files || [];

    const tenant = tenantListData.find(t => t.tenant_id === tenantId);
    const bestScore = tenant && tenant.best_score != null
        ? (tenant.best_score * 100).toFixed(2) + '%'
        : 'N/A';

    const ts = lastSession.timestamp;
    const displayTime = ts !== 'unknown'
        ? `${ts.substring(0,4)}-${ts.substring(4,6)}-${ts.substring(6,8)} ${ts.substring(9,11)}:${ts.substring(11,13)}:${ts.substring(13,15)}`
        : '未知时间';

    // 更新进度区汇总
    document.getElementById('current-status').textContent = '训练已完成';
    document.getElementById('current-iteration').textContent = sessions.length;
    document.getElementById('max-iteration').textContent = sessions.length;
    document.getElementById('accuracy').textContent = bestScore;
    updateProgressBar(100);

    // 填充结果区
    const resultSection = document.getElementById('result-section');
    const resultCard = document.getElementById('result-card');
    const resultDownloads = document.getElementById('result-downloads');

    resultCard.className = 'result-card result-success';
    resultCard.innerHTML = `
        <div class="result-row">
            <div class="result-item">
                <div class="label">训练状态</div>
                <div class="value" style="color: #28a745; font-size: 20px;">已完成</div>
            </div>
            <div class="result-item">
                <div class="label">最佳准确率</div>
                <div class="value score">${bestScore}</div>
            </div>
            <div class="result-item">
                <div class="label">训练次数</div>
                <div class="value">${sessions.length}</div>
            </div>
            <div class="result-item">
                <div class="label">最后训练</div>
                <div class="value" style="font-size: 13px;">${displayTime}</div>
            </div>
        </div>
    `;

    const codeFile = files.find(f => f.type === 'code');
    const outputFile = files.find(f => f.type === 'output');
    const compFile = files.find(f => f.type === 'comparison');

    let downloadHtml = '<div class="download-row">';
    downloadHtml += `<button class="btn btn-download" onclick="downloadScript()">下载脚本 (.py)</button>`;
    if (outputFile) {
        downloadHtml += `<button class="btn btn-download" onclick="window.location.href='/api/download-log/${tenantId}/${outputFile.filename}'">下载输出结果 (.xlsx)</button>`;
    } else {
        downloadHtml += `<button class="btn btn-download" onclick="downloadTrainingFiles('${tenantId}', 'output')">下载输出结果 (.xlsx)</button>`;
    }
    if (compFile) {
        downloadHtml += `<button class="btn btn-download" onclick="window.location.href='/api/download-log/${tenantId}/${compFile.filename}'">下载差异对比 (.xlsx)</button>`;
    } else {
        downloadHtml += `<button class="btn btn-download" onclick="downloadTrainingFiles('${tenantId}', 'comparison')">下载差异对比 (.xlsx)</button>`;
    }
    downloadHtml += `<button class="btn btn-adjust" onclick="openAdjustModal('${tenantId}')">调整逻辑</button>`;
    downloadHtml += '</div>';

    resultDownloads.innerHTML = downloadHtml;
}

// ==================== 训练流程 ====================

async function startTraining() {
    const tenantId = document.getElementById('tenant-input').value;
    const sourceFiles = document.getElementById('source-files').files;
    const targetFile = document.getElementById('target-file').files[0];
    const ruleFiles = document.getElementById('rule-files').files;
    const aiProvider = document.getElementById('ai-provider').value;
    const mode = document.getElementById('mode').value;
    const maxIterations = parseInt(document.getElementById('max-iterations').value);
    const forceRetrain = document.getElementById('force-retrain').checked;

    if (!tenantId || sourceFiles.length === 0 || !targetFile) {
        alert('请填写所有必填项');
        return;
    }

    document.getElementById('tenant-id').textContent = tenantId;
    document.getElementById('max-iteration').textContent = maxIterations;
    document.getElementById('start-training-btn').disabled = true;
    document.getElementById('start-training-btn').textContent = '训练中...';

    clearResultSection();
    resetProgress();

    const formData = new FormData();
    formData.append('tenant_id', tenantId);
    formData.append('ai_provider', aiProvider);
    formData.append('mode', mode);
    formData.append('max_iterations', maxIterations);
    formData.append('force_retrain', forceRetrain);
    Array.from(sourceFiles).forEach(file => formData.append('source_files', file));
    formData.append('target_file', targetFile);
    Array.from(ruleFiles).forEach(file => formData.append('rule_files', file));

    // 添加可选参数
    const salaryMonth = document.getElementById('salary-month').value.trim();
    const standardHours = document.getElementById('standard-hours').value.trim();
    const manualHeaders = document.getElementById('manual-headers').value.trim();

    // 解析薪资年月（格式：2026-03）
    if (salaryMonth) {
        const parts = salaryMonth.split('-');
        if (parts.length === 2) {
            const year = parseInt(parts[0]);
            const month = parseInt(parts[1]);
            if (!isNaN(year) && !isNaN(month)) {
                formData.append('salary_year', year);
                formData.append('salary_month', month);
            }
        }
    }

    if (standardHours) formData.append('monthly_standard_hours', standardHours);
    if (manualHeaders) formData.append('manual_headers', manualHeaders);

    trainingStartTime = Date.now();
    startTimer();

    addLog('info', '开始训练...');
    addLog('info', `租户: ${tenantId}  模型: ${aiProvider}  模式: ${mode}  最大迭代: ${maxIterations}`);

    try {
        const response = await fetch('/api/train/stream', {
            method: 'POST',
            body: formData
        });

        if (!response.ok) {
            throw new Error(`HTTP error! status: ${response.status}`);
        }

        const reader = response.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const parts = buffer.split('\n\n');
            buffer = parts.pop();

            for (const part of parts) {
                const lines = part.split('\n');
                for (const line of lines) {
                    if (line.startsWith('data: ')) {
                        const data = line.substring(6).trim();
                        if (data) {
                            try {
                                handleTrainingEvent(JSON.parse(data));
                            } catch (e) {
                                addLog('info', data);
                            }
                        }
                    }
                }
            }
        }

        if (buffer.trim()) {
            const lines = buffer.split('\n');
            for (const line of lines) {
                if (line.startsWith('data: ')) {
                    const data = line.substring(6).trim();
                    if (data) {
                        try { handleTrainingEvent(JSON.parse(data)); } catch (e) { addLog('info', data); }
                    }
                }
            }
        }

    } catch (error) {
        console.error('Training error:', error);
        addLog('error', `训练失败: ${error.message}`);
        showError(error.message);
    }
}

function handleTrainingEvent(event) {
    switch (event.type) {
        case 'status':
            updateStatus(event.message);
            addLog('info', event.message);
            break;

        case 'iteration_start':
            startIteration(event.iteration, event.total);
            addLog('info', `开始第 ${event.iteration}/${event.total} 轮`);
            break;

        case 'iteration_progress':
            updateIterationProgress(event.iteration, event.message);
            addLog('info', event.message);
            break;

        case 'iteration_complete':
            flushCodeStreamBuffer();
            completeIteration(event.iteration, event.success, event.accuracy, event.message);
            if (event.success) {
                addLog('success', `第 ${event.iteration} 轮 - 准确率: ${event.accuracy ? (event.accuracy * 100).toFixed(2) + '%' : 'N/A'}`);
            } else {
                addLog('warning', `第 ${event.iteration} 轮: ${event.message}`);
            }
            break;

        case 'complete':
            flushCodeStreamBuffer();
            completeTraining(event.success, event.data);
            break;

        case 'error':
            addLog('error', event.message);
            showError(event.message);
            break;

        case 'log':
            addLog(event.level || 'info', event.message);
            break;

        case 'code_stream':
            codeStreamBuffer += (event.chunk || '');
            const codeLines = codeStreamBuffer.split('\n');
            codeStreamBuffer = codeLines.pop();
            for (const line of codeLines) {
                if (line.trim()) addLog('info', line);
            }
            break;
    }
}

// ==================== UI 更新 ====================

function updateStatus(message) {
    document.getElementById('current-status').textContent = message;
}

function startIteration(iteration, total) {
    document.getElementById('current-iteration').textContent = iteration;
    updateProgressBar((iteration / total) * 100);
}

function updateIterationProgress(iteration, message) {
    // 进度信息写入日志即可
}

function completeIteration(iteration, success, accuracy, message) {
    if (accuracy != null) {
        document.getElementById('accuracy').textContent = `${(accuracy * 100).toFixed(2)}%`;
    }
}

function completeTraining(success, data) {
    stopTimer();
    updateProgressBar(100);

    if (success) {
        updateStatus('训练完成');
        addLog('success', '训练成功完成!');
        showResult(data);
    } else {
        updateStatus('训练失败');
        addLog('error', '训练失败');
        showError(data.message || '训练过程中发生错误');
    }

    document.getElementById('start-training-btn').disabled = false;
    document.getElementById('start-training-btn').textContent = '开始训练';

    // 刷新
    const tenantId = document.getElementById('tenant-input').value.trim();
    loadTenantList();
    loadTrainingHistory();
    if (tenantId) loadTenantScripts(tenantId);
}

function showResult(data) {
    const resultSection = document.getElementById('result-section');
    const resultCard = document.getElementById('result-card');
    const resultDownloads = document.getElementById('result-downloads');
    const tenantId = document.getElementById('tenant-input').value.trim();

    resultCard.className = 'result-card result-success';
    resultCard.innerHTML = `
        <div class="result-row">
            <div class="result-item">
                <div class="label">训练状态</div>
                <div class="value" style="color: #28a745; font-size: 20px;">成功</div>
            </div>
            <div class="result-item">
                <div class="label">最终准确率</div>
                <div class="value score">${data.final_accuracy ? (data.final_accuracy * 100).toFixed(2) + '%' : 'N/A'}</div>
            </div>
            <div class="result-item">
                <div class="label">训练轮次</div>
                <div class="value">${data.iterations || 'N/A'}</div>
            </div>
            <div class="result-item">
                <div class="label">训练时长</div>
                <div class="value">${document.getElementById('elapsed-time').textContent}</div>
            </div>
            <div class="result-item">
                <div class="label">脚本ID</div>
                <div class="value" style="font-size: 12px;">${data.script_id || 'N/A'}</div>
            </div>
        </div>
    `;

    resultDownloads.innerHTML = `
        <div class="download-row">
            <button class="btn btn-download" onclick="downloadScript()">下载脚本 (.py)</button>
            <button class="btn btn-download" onclick="downloadTrainingFiles('${tenantId}', 'output')">下载输出结果 (.xlsx)</button>
            <button class="btn btn-download" onclick="downloadTrainingFiles('${tenantId}', 'comparison')">下载差异对比 (.xlsx)</button>
            <button class="btn btn-adjust" onclick="openAdjustModal('${tenantId}')">调整逻辑</button>
            <button class="btn btn-secondary" onclick="resetTraining()">重新训练</button>
        </div>
    `;
}

function showError(message) {
    const resultSection = document.getElementById('result-section');
    const resultCard = document.getElementById('result-card');

    resultCard.className = 'result-card result-error';
    resultCard.innerHTML = `
        <div class="result-row">
            <div class="result-item">
                <div class="label">训练状态</div>
                <div class="value" style="color: #dc3545; font-size: 20px;">失败</div>
            </div>
            <div class="result-item" style="flex: 3;">
                <div class="label">错误信息</div>
                <div class="value" style="font-size: 13px; color: #dc3545;">${message}</div>
            </div>
        </div>
    `;
}

// ==================== 进度条和日志 ====================

function flushCodeStreamBuffer() {
    if (codeStreamBuffer.trim()) addLog('info', codeStreamBuffer);
    codeStreamBuffer = '';
}

function updateProgressBar(percent) {
    document.getElementById('progress-bar').style.width = percent + '%';
    document.getElementById('progress-text').textContent = Math.round(percent) + '%';
}

function addLog(level, message) {
    const logContent = document.getElementById('log-content');
    const timestamp = new Date().toLocaleTimeString();
    const entry = document.createElement('div');
    entry.className = `log-entry ${level}`;
    entry.innerHTML = `<span class="log-timestamp">[${timestamp}]</span>${message}`;
    logContent.appendChild(entry);
    logContent.scrollTop = logContent.scrollHeight;
}

function startTimer() {
    timerInterval = setInterval(() => {
        const elapsed = Date.now() - trainingStartTime;
        const minutes = Math.floor(elapsed / 60000);
        const seconds = Math.floor((elapsed % 60000) / 1000);
        document.getElementById('elapsed-time').textContent =
            `${String(minutes).padStart(2, '0')}:${String(seconds).padStart(2, '0')}`;
    }, 1000);
}

function stopTimer() {
    if (timerInterval) { clearInterval(timerInterval); timerInterval = null; }
}

function resetProgress() {
    codeStreamBuffer = '';
    document.getElementById('current-status').textContent = '等待训练';
    document.getElementById('current-iteration').textContent = '0';
    document.getElementById('elapsed-time').textContent = '--:--';
    document.getElementById('accuracy').textContent = '-';
    document.getElementById('log-content').innerHTML =
        '<div class="log-entry info"><span class="log-timestamp">[--:--:--]</span>等待训练开始...</div>';
    updateProgressBar(0);
}

function clearResultSection() {
    const resultCard = document.getElementById('result-card');
    const resultDownloads = document.getElementById('result-downloads');
    resultCard.innerHTML = '<p style="color: #999; text-align: center; padding: 20px;">等待训练完成...</p>';
    resultDownloads.innerHTML = '';
}

function resetTraining() {
    clearResultSection();
    document.getElementById('start-training-btn').disabled = false;
    document.getElementById('start-training-btn').textContent = '开始训练';
    resetProgress();
}

// ==================== 下载 ====================

function downloadScript() {
    const tenantId = document.getElementById('tenant-input').value.trim()
        || document.getElementById('tenant-id').textContent;
    if (tenantId && tenantId !== '-') {
        window.location.href = `/api/download-script/${tenantId}`;
    } else {
        alert('无法下载脚本：租户ID无效');
    }
}

async function downloadTrainingFiles(tenantId, fileType) {
    try {
        const response = await fetch(`/api/training-logs/${tenantId}`);
        if (!response.ok) { alert('获取训练日志失败'); return; }
        const data = await response.json();
        const logs = data.logs || [];

        let pattern;
        if (fileType === 'output') {
            pattern = /_output_.*\.xlsx$/i;
        } else if (fileType === 'comparison') {
            pattern = /(_comparison_|差异对比).*\.xlsx$/i;
        }

        const matched = logs.find(f => pattern && pattern.test(f.filename));
        if (matched) {
            window.location.href = `/api/download-log/${tenantId}/${matched.filename}`;
        } else {
            alert(`未找到${fileType === 'output' ? '输出结果' : '差异对比'}文件`);
        }
    } catch (e) {
        alert('下载失败: ' + e.message);
    }
}

// ==================== 调整逻辑弹窗 ====================

let adjustDiffData = {};

async function openAdjustModal(tenantId) {
    const modal = document.getElementById('adjust-modal');
    const loading = document.getElementById('adjust-loading');
    const columnsList = document.getElementById('adjust-columns-list');
    const rulesSummary = document.getElementById('adjust-rules-summary');

    modal.style.display = 'flex';
    loading.style.display = 'block';
    loading.innerHTML = '<div class="loading">正在实时验证脚本并分析差异列...</div>';
    columnsList.innerHTML = '';
    rulesSummary.innerHTML = '';
    adjustDiffData = {};

    // 调用增强版 training-detail（live=true，实时执行脚本对比）
    let detail = null;
    try {
        // 构建URL，附带薪资参数（避免脚本因参数为空报错）
        let detailUrl = `/api/training-detail/${tenantId}?live=true`;
        const salaryMonthEl = document.getElementById('salary-month');
        const standardHoursEl = document.getElementById('standard-hours');
        if (salaryMonthEl && salaryMonthEl.value.trim()) {
            const parts = salaryMonthEl.value.trim().split('-');
            if (parts.length === 2) {
                const year = parseInt(parts[0]);
                const month = parseInt(parts[1]);
                if (!isNaN(year) && !isNaN(month)) {
                    detailUrl += `&salary_year=${year}&salary_month=${month}`;
                }
            }
        }
        if (standardHoursEl && standardHoursEl.value.trim()) {
            detailUrl += `&monthly_standard_hours=${standardHoursEl.value.trim()}`;
        }
        const response = await fetch(detailUrl);
        if (response.ok) {
            detail = await response.json();
        }
    } catch (e) {
        console.error('training-detail failed:', e);
    }

    let fieldDiffs = (detail && detail.field_diff_samples) || {};
    adjustDiffData = fieldDiffs;

    // 显示规则摘要
    if (detail && detail.rules_content) {
        const rulesPreview = detail.rules_content.length > 500
            ? detail.rules_content.substring(0, 500) + '...'
            : detail.rules_content;
        rulesSummary.innerHTML = `
            <div class="adjust-rules-box">
                <div class="adjust-rules-header" onclick="this.parentElement.classList.toggle('expanded')">
                    当前使用的规则 <span class="toggle-icon">&#9654;</span>
                </div>
                <div class="adjust-rules-content"><pre>${escapeHtml(rulesPreview)}</pre></div>
            </div>
        `;
    }

    renderAdjustColumns(tenantId, fieldDiffs, detail);
}

function renderAdjustColumns(tenantId, fieldDiffs, detail) {
    const loading = document.getElementById('adjust-loading');
    const columnsList = document.getElementById('adjust-columns-list');
    loading.style.display = 'none';

    const diffColumns = Object.keys(fieldDiffs);
    const columnCodeMap = (detail && detail.column_code_map) || {};
    const liveValidated = detail && detail.live_validated;

    if (diffColumns.length === 0) {
        // 无差异列，提供手动输入（同时列出所有预期列名供参考）
        let expectedInfo = '';
        if (detail && detail.expected_columns && detail.expected_columns.length > 0) {
            expectedInfo = `<div class="adjust-column-info"><span>预期输出列: ${detail.expected_columns.join(', ')}</span></div>`;
        }
        columnsList.innerHTML = `
            <div class="adjust-column-card no-diff" style="border-left-color: #667eea;">
                <div class="adjust-column-header">
                    <span class="adjust-column-name">手动调整</span>
                    <span class="adjust-column-badge" style="background: #e7f3ff; color: #667eea;">自定义</span>
                </div>
                ${expectedInfo}
                <div class="form-group" style="margin-bottom: 8px;">
                    <input type="text" id="manual-target-columns" placeholder="输入目标列名，多列用逗号分隔" style="width:100%; padding:8px 10px; border:1px solid #ddd; border-radius:6px; font-size:13px;">
                </div>
                <textarea class="adjust-column-textarea" id="manual-adjust-request" placeholder="输入修改意见，如：加班费改为按1.5倍计算"></textarea>
            </div>
        `;
        return;
    }

    // 渲染差异列
    const currentScore = detail && detail.score != null ? (detail.score * 100).toFixed(2) + '%' : 'N/A';
    let html = `<div style="font-size:13px; color:#666; margin-bottom:12px;">
        当前准确率: <strong>${currentScore}</strong>${liveValidated ? ' (实时验证)' : ''} |
        共 <strong>${diffColumns.length}</strong> 个差异列，在需要修改的列中填写修改意见：
    </div>`;

    const sorted = diffColumns.sort((a, b) => (fieldDiffs[b].count || 0) - (fieldDiffs[a].count || 0));

    sorted.forEach(col => {
        const info = fieldDiffs[col];
        const count = info.count || 0;
        const formula = info.formula || '';
        const codeLogic = columnCodeMap[col] || '';

        let codeSection = '';
        if (codeLogic) {
            const codeId = 'code-' + col.replace(/[^a-zA-Z0-9\u4e00-\u9fa5]/g, '_');
            codeSection = `
                <div class="adjust-code-box" id="box-${codeId}">
                    <div class="adjust-code-header" onclick="this.parentElement.classList.toggle('expanded')">
                        当前代码逻辑 <span class="toggle-icon">&#9654;</span>
                    </div>
                    <div class="adjust-code-content"><pre>${escapeHtml(codeLogic)}</pre></div>
                </div>
            `;
        }

        html += `
            <div class="adjust-column-card has-diff">
                <div class="adjust-column-header">
                    <span class="adjust-column-name">${escapeHtml(col)}</span>
                    <span class="adjust-column-badge diff">${count} 处差异</span>
                </div>
                <div class="adjust-column-info">
                    <span>当前公式: <code>${escapeHtml(formula) || '无公式/纯值'}</code></span>
                </div>
                ${codeSection}
                <textarea class="adjust-column-textarea" data-column="${escapeHtml(col)}" placeholder="填写修改意见（留空则不调整此列）"></textarea>
            </div>
        `;
    });

    columnsList.innerHTML = html;
}

function escapeHtml(str) {
    if (!str) return '';
    return str.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
}

function closeAdjustModal() {
    document.getElementById('adjust-modal').style.display = 'none';
}

async function submitAdjustment() {
    const tenantId = document.getElementById('tenant-input').value.trim();
    if (!tenantId) { alert('租户ID不能为空'); return; }

    let targetColumns = [];
    let adjustmentParts = [];

    const manualInput = document.getElementById('manual-target-columns');
    const manualRequest = document.getElementById('manual-adjust-request');

    if (manualInput && manualInput.value.trim()) {
        targetColumns = manualInput.value.trim().split(/[,，]/).map(s => s.trim()).filter(Boolean);
        adjustmentParts.push(manualRequest ? manualRequest.value.trim() : '');
    } else {
        const textareas = document.querySelectorAll('#adjust-columns-list textarea[data-column]');
        textareas.forEach(ta => {
            const text = ta.value.trim();
            if (text) {
                const col = ta.getAttribute('data-column');
                targetColumns.push(col);
                adjustmentParts.push(`${col}: ${text}`);
            }
        });
    }

    if (targetColumns.length === 0 || adjustmentParts.join('').trim() === '') {
        alert('请至少填写一列的修改意见');
        return;
    }

    const adjustmentRequest = adjustmentParts.join('\n');
    const submitBtn = document.getElementById('adjust-submit-btn');
    submitBtn.disabled = true;
    submitBtn.textContent = '修正中...';

    addLog('info', `开始调整逻辑: 目标列=[${targetColumns.join(', ')}]`);
    addLog('info', `修改意见: ${adjustmentRequest}`);

    try {
        const formData = new FormData();
        formData.append('tenant_id', tenantId);
        formData.append('adjustment_request', adjustmentRequest);
        formData.append('target_columns', targetColumns.join(','));

        // 附带薪资参数（避免脚本因参数为空报错）
        const salaryMonthEl = document.getElementById('salary-month');
        const standardHoursEl = document.getElementById('standard-hours');
        if (salaryMonthEl && salaryMonthEl.value.trim()) {
            const parts = salaryMonthEl.value.trim().split('-');
            if (parts.length === 2) {
                const year = parseInt(parts[0]);
                const month = parseInt(parts[1]);
                if (!isNaN(year) && !isNaN(month)) {
                    formData.append('salary_year', year);
                    formData.append('salary_month', month);
                }
            }
        }
        if (standardHoursEl && standardHoursEl.value.trim()) {
            formData.append('monthly_standard_hours', standardHoursEl.value.trim());
        }

        const response = await fetch('/api/adjust-code', {
            method: 'POST',
            body: formData
        });

        const result = await response.json();

        if (!response.ok) {
            throw new Error(result.detail || '调整请求失败');
        }

        const columnsList = document.getElementById('adjust-columns-list');

        // 处理执行失败的情况（脚本出错才显示失败）
        if (result.status === 'execution_failed' || result.status === 'ai_generation_failed' || result.status === 'no_output') {
            columnsList.innerHTML += `
                <div class="adjust-result rejected">
                    <strong>脚本执行失败</strong> — ${escapeHtml(result.error || result.message || '未知错误')}
                    <br>执行耗时: ${result.execution_time || '-'}秒
                </div>
            `;
            addLog('error', `调整失败: ${result.error || result.message || '未知错误'}`);
        } else {
            // 正常完成：显示对比结果
            const adopted = result.adopted;
            const origScore = result.original_score != null ? result.original_score : '-';
            const newScore = result.new_score != null ? result.new_score : '-';
            const matchedCells = result.matched_cells != null ? result.matched_cells : '-';
            const totalCells = result.total_cells != null ? result.total_cells : '-';
            const totalDiff = result.total_differences != null ? result.total_differences : '-';

            columnsList.innerHTML += `
                <div class="adjust-result adopted">
                    <strong>已采纳</strong> — 新脚本ID: ${result.new_script_id || '-'}
                    <br>原始分数: ${origScore}% → 新分数: ${newScore}%
                    <br>匹配: ${matchedCells}/${totalCells} 个单元格，差异 ${totalDiff} 处
                    <br>执行耗时: ${result.execution_time || '-'}秒
                </div>
            `;

            addLog('success', `调整完成: 已采纳 原始=${origScore}% 新=${newScore}%`);

            // 刷新列表显示新脚本
            await Promise.all([loadTenantList(), loadTrainingHistory()]);
            loadTenantStatus();
            loadTenantScripts(tenantId);
        }
    } catch (e) {
        addLog('error', `调整失败: ${e.message}`);
        alert('调整失败: ' + e.message);
    } finally {
        submitBtn.disabled = false;
        submitBtn.textContent = '开始修正';
    }
}

// ==================== 文件列表显示 ====================

function updateFileList(inputId, listId) {
    const input = document.getElementById(inputId);
    const list = document.getElementById(listId);

    console.log('updateFileList called:', inputId, 'files:', input?.files?.length);

    if (!input || !list) {
        console.error('Element not found:', inputId, listId);
        return;
    }

    if (!input.files || input.files.length === 0) {
        list.innerHTML = '';
        return;
    }

    let html = '<div style="margin-top: 8px; font-size: 13px; color: #666;">';
    Array.from(input.files).forEach(file => {
        html += `<div style="padding: 4px 0;">📄 ${file.name}</div>`;
    });
    html += '</div>';
    list.innerHTML = html;
    console.log('File list updated:', list.innerHTML);
}
