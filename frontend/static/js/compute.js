// 全局变量
let tenantListData = [];
let currentTenantId = '';
let currentScriptId = '';

// 页面加载完成后初始化
document.addEventListener('DOMContentLoaded', function() {
    // 认证检查
    if (!AUTH.requireAuth()) return;
    AUTH.renderUserInfo(document.querySelector('header'));
    if (AUTH.isAdmin()) {
        const adminNav = document.getElementById('nav-admin');
        if (adminNav) adminNav.style.display = '';
    }

    loadTenantList();

    // 租户输入框事件
    const input = document.getElementById('tenant-input');
    if (input) {
        input.addEventListener('focus', showTenantDropdown);
        input.addEventListener('input', filterTenantDropdown);
        input.addEventListener('blur', function() {
            setTimeout(() => {
                document.getElementById('tenant-dropdown').style.display = 'none';
                if (input.value.trim()) {
                    onTenantSelected(input.value.trim());
                }
            }, 200);
        });
    }

    // 文件选择事件
    document.getElementById('source-files').addEventListener('change', () => {
        updateFileList();
        checkCanCompute();
    });

    // 计算按钮
    document.getElementById('compute-btn').addEventListener('click', startCompute);

    // 脚本选择器
    document.getElementById('script-selector').addEventListener('change', onScriptSelected);
});

// 点击外部关闭下拉框
document.addEventListener('click', function(e) {
    const combo = document.getElementById('tenant-combo');
    const dropdown = document.getElementById('tenant-dropdown');
    if (combo && dropdown && !combo.contains(e.target)) {
        dropdown.style.display = 'none';
    }
});

// ==================== 租户列表 ====================

async function loadTenantList() {
    try {
        const resp = await AUTH.authFetch('/api/training-history');
        const data = await resp.json();
        console.log('Training history data:', data);

        // API返回 {history: {tenant_id: {...}}}
        const historyData = data.history || {};
        tenantListData = Object.keys(historyData).map(tid => ({
            tenant_id: tid,
            best_score: historyData[tid].best_score || 0
        })).sort((a, b) => b.best_score - a.best_score);

        console.log('Tenant list loaded:', tenantListData.length, 'tenants');
    } catch (e) {
        console.error('加载租户列表失败:', e);
    }
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

function renderTenantDropdown(tenants) {
    const dropdown = document.getElementById('tenant-dropdown');
    if (tenants.length === 0) {
        dropdown.innerHTML = '<div class="combo-item">无匹配租户</div>';
        return;
    }
    dropdown.innerHTML = tenants.map(t => `
        <div class="combo-item" onclick="selectTenant('${t.tenant_id}')">
            <span class="combo-id">${t.tenant_id}</span>
            <span class="combo-score">${(t.best_score * 100).toFixed(2)}%</span>
        </div>
    `).join('');
}

function selectTenant(tenantId) {
    document.getElementById('tenant-input').value = tenantId;
    document.getElementById('tenant-dropdown').style.display = 'none';
    onTenantSelected(tenantId);
}

async function onTenantSelected(tenantId) {
    currentTenantId = tenantId;
    await loadTenantScripts(tenantId);
    checkCanCompute();
}

// ==================== 脚本列表 ====================

async function loadTenantScripts(tenantId) {
    const group = document.getElementById('script-selector-group');
    const selector = document.getElementById('script-selector');
    const info = document.getElementById('script-info');

    try {
        const resp = await AUTH.authFetch(`/api/tenant-scripts/${encodeURIComponent(tenantId)}`);
        if (!resp.ok) {
            group.style.display = 'none';
            return;
        }

        const data = await resp.json();
        if (!data.scripts || data.scripts.length === 0) {
            group.style.display = 'none';
            addLog('warning', '该租户没有可用的训练脚本');
            return;
        }

        // 按分数排序，最高分在前
        const scripts = data.scripts.sort((a, b) => b.score - a.score);

        selector.innerHTML = scripts.map(s => `
            <option value="${s.script_id}" ${s.is_active ? 'selected' : ''}>
                ${s.script_id} - ${(s.score * 100).toFixed(2)}% ${s.is_active ? '(当前)' : ''}
            </option>
        `).join('');

        group.style.display = 'block';

        // 自动选择最高分
        currentScriptId = scripts[0].script_id;
        updateScriptInfo(scripts[0]);

    } catch (e) {
        console.error('加载脚本列表失败:', e);
        group.style.display = 'none';
    }
}

function onScriptSelected() {
    const selector = document.getElementById('script-selector');
    currentScriptId = selector.value;

    // 更新脚本信息显示
    const option = selector.options[selector.selectedIndex];
    const text = option.textContent;
    const scoreMatch = text.match(/([\d.]+)%/);
    if (scoreMatch) {
        updateScriptInfo({
            script_id: currentScriptId,
            score: parseFloat(scoreMatch[1]) / 100
        });
    }
}

function updateScriptInfo(script) {
    const info = document.getElementById('script-info');
    info.innerHTML = `
        <div style="margin-top: 8px; padding: 8px; background: #f0f8ff; border-radius: 4px; font-size: 13px;">
            <div>脚本ID: <strong>${script.script_id}</strong></div>
            <div>准确率: <strong style="color: #28a745;">${(script.score * 100).toFixed(2)}%</strong></div>
        </div>
    `;
}

// ==================== 文件列表 ====================

function updateFileList() {
    const input = document.getElementById('source-files');
    const list = document.getElementById('source-file-list');

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
}

function checkCanCompute() {
    const btn = document.getElementById('compute-btn');
    const hasFiles = document.getElementById('source-files').files.length > 0;
    const hasTenant = currentTenantId && currentScriptId;
    btn.disabled = !(hasFiles && hasTenant);
}

// ==================== 计算执行 ====================

async function startCompute() {
    const btn = document.getElementById('compute-btn');
    const files = document.getElementById('source-files').files;

    if (!currentTenantId || !currentScriptId) {
        alert('请先选择租户和计算版本');
        return;
    }

    if (files.length === 0) {
        alert('请选择源文件');
        return;
    }

    btn.disabled = true;
    btn.textContent = '计算中...';
    clearResult();
    addLog('info', '准备开始计算...');
    updateStatus('计算中');
    updateProgress(10);

    const formData = new FormData();
    formData.append('tenant_id', currentTenantId);
    formData.append('script_id', currentScriptId);

    Array.from(files).forEach(file => {
        formData.append('source_files', file);
    });

    // 添加可选参数
    const salaryMonth = document.getElementById('salary-month').value.trim();
    const standardHours = document.getElementById('standard-hours').value.trim();

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

    if (standardHours) formData.append('standard_hours', standardHours);

    try {
        updateProgress(20);
        addLog('info', '正在连接服务器...');

        const resp = await AUTH.authFetch('/api/compute/stream', {
            method: 'POST',
            body: formData
        });

        if (!resp.ok) {
            throw new Error(`HTTP error! status: ${resp.status}`);
        }

        addLog('info', '连接成功，开始接收数据...');
        const reader = resp.body.getReader();
        const decoder = new TextDecoder();
        let buffer = '';

        updateProgress(30);

        while (true) {
            const { done, value } = await reader.read();
            if (done) break;

            buffer += decoder.decode(value, { stream: true });
            const lines = buffer.split('\n');
            buffer = lines.pop(); // 保留不完整的行

            for (const line of lines) {
                if (line.startsWith('data: ')) {
                    const data = line.slice(6);
                    if (data.trim()) {
                        try {
                            const event = JSON.parse(data);
                            handleComputeEvent(event);
                        } catch (e) {
                            console.error('解析事件失败:', e, data);
                        }
                    }
                }
            }
        }

    } catch (e) {
        console.error('计算失败:', e);
        addLog('error', `计算失败: ${e.message}`);
        updateStatus('计算失败');
        showError(e.message);
    } finally {
        btn.disabled = false;
        btn.textContent = '开始计算';
    }
}

function handleComputeEvent(event) {
    switch (event.type) {
        case 'status':
            addLog('info', event.message);
            // 根据状态更新进度
            if (event.message.includes('计算开始')) {
                updateProgress(15);
            }
            break;

        case 'log':
            addLog(event.level, event.message);
            // 根据日志内容递进式更新进度
            if (event.message.includes('保存源文件')) {
                updateProgress(25);
            } else if (event.message.includes('表头匹配')) {
                updateProgress(35);
            } else if (event.message.includes('匹配完成')) {
                updateProgress(45);
            } else if (event.message.includes('开始执行')) {
                updateProgress(55);
            } else if (event.message.includes('预加载源数据') || event.message.includes('性能优化')) {
                updateProgress(65);
            } else if (event.message.includes('生成输出文件')) {
                updateProgress(85);
            } else if (event.message.includes('结果已保存')) {
                updateProgress(95);
            }
            break;

        case 'complete':
            updateProgress(100);
            updateStatus('计算完成');
            addLog('success', '✓ 计算成功完成!');
            if (event.data) {
                showResult(event.data);
            }
            break;

        case 'error':
            addLog('error', `✗ 错误: ${event.message}`);
            updateStatus('计算失败');
            showError(event.message);
            break;

        default:
            console.log('未知事件类型:', event);
    }
}

// ==================== 结果显示 ====================

function showResult(data) {
    const resultCard = document.getElementById('result-card');
    const resultDownloads = document.getElementById('result-downloads');

    resultCard.className = 'result-card result-success';
    resultCard.innerHTML = `
        <div class="result-row">
            <div class="result-item">
                <div class="label">计算状态</div>
                <div class="value" style="color: #28a745; font-size: 20px;">成功</div>
            </div>
            <div class="result-item">
                <div class="label">输出文件</div>
                <div class="value">${data.output_file || 'N/A'}</div>
            </div>
            <div class="result-item">
                <div class="label">处理行数</div>
                <div class="value">${data.rows_processed || 'N/A'}</div>
            </div>
        </div>
    `;

    resultDownloads.innerHTML = `
        <div class="download-row">
            <button class="btn btn-download" onclick="downloadResult('${data.output_file}')">下载计算结果</button>
            <button class="btn btn-secondary" onclick="resetCompute()">重新计算</button>
        </div>
    `;
}

function showError(message) {
    const resultCard = document.getElementById('result-card');
    resultCard.className = 'result-card result-error';
    resultCard.innerHTML = `
        <div class="result-row">
            <div class="result-item">
                <div class="label">计算状态</div>
                <div class="value" style="color: #dc3545; font-size: 20px;">失败</div>
            </div>
            <div class="result-item" style="flex: 3;">
                <div class="label">错误信息</div>
                <div class="value" style="font-size: 13px; color: #dc3545;">${message}</div>
            </div>
        </div>
    `;
}

function clearResult() {
    const resultCard = document.getElementById('result-card');
    const resultDownloads = document.getElementById('result-downloads');
    resultCard.innerHTML = '<p style="color: #999; text-align: center; padding: 20px;">等待计算完成...</p>';
    resultDownloads.innerHTML = '';
    document.getElementById('log-content').innerHTML = '';
}

function resetCompute() {
    clearResult();
    updateStatus('等待计算');
    updateProgress(0);
    document.getElementById('source-files').value = '';
    updateFileList();
    checkCanCompute();
}

// ==================== 下载 ====================

async function downloadResult(filename) {
    if (!filename) {
        alert('没有可下载的文件');
        return;
    }
    try {
        window.location.href = `/api/download-compute-result/${encodeURIComponent(currentTenantId)}/${encodeURIComponent(filename)}`;
    } catch (e) {
        alert('下载失败: ' + e.message);
    }
}

// ==================== UI更新 ====================

function updateStatus(text) {
    document.getElementById('current-status').textContent = text;
}

function updateProgress(percent) {
    document.getElementById('progress-bar').style.width = percent + '%';
    document.getElementById('progress-text').textContent = Math.round(percent) + '%';
}

function addLog(level, message) {
    const logContent = document.getElementById('log-content');
    const time = new Date().toLocaleTimeString('zh-CN', { hour12: false });
    const entry = document.createElement('div');
    entry.className = `log-entry ${level}`;
    entry.innerHTML = `<span class="log-timestamp">[${time}]</span>${message}`;
    logContent.appendChild(entry);
    logContent.scrollTop = logContent.scrollHeight;
}
