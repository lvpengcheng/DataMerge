let _sessionSettings = null;
let _savingAssets = false;

async function refreshSessionAssets() {
    const id = _currentSessionId;
    if (!id || _savingAssets) return;
    try {
        const response = await AUTH.authFetch(`/api/training/chat/sessions/${id}/assets`);
        if (!response.ok) return;
        const data = await response.json();
        if (id !== _currentSessionId || _savingAssets) return;
        _renderSessionAssets(data);
        if (!_sessionSettings) _restoreSessionSettings(data.settings);
    } catch (error) { console.warn('刷新训练文件列表失败:', error); }
}

function _readSessionSettings() {
    const period = document.getElementById('salary-month').value.trim();
    if (period && !/^\d{4}[-/]\d{1,2}$/.test(period)) throw new Error('薪资年月格式应为 2026-09');
    const [year, month] = period ? period.split(/[-/]/).map(Number) : [null, null];
    const hours = document.getElementById('standard-hours').value;
    const headers = document.getElementById('manual-headers').value.trim();
    return {ai_provider: document.getElementById('ai-provider').value,
        mode: document.getElementById('mode').value, salary_year: year, salary_month: month,
        monthly_standard_hours: hours ? Number(hours) : null, manual_headers: headers ? JSON.parse(headers) : null,
        multi_sheet_source: document.getElementById('multi-sheet-source').checked,
        use_history: document.getElementById('use-history').checked};
}

function _restoreSessionSettings(settings) {
    const values = settings || {};
    const provider = document.getElementById('ai-provider');
    if (values.ai_provider && !Array.from(provider.options).some(o => o.value === values.ai_provider))
        provider.add(new Option(values.ai_provider, values.ai_provider));
    if (values.ai_provider) provider.value = values.ai_provider;
    document.getElementById('mode').value = values.mode || 'formula';
    document.getElementById('salary-month').value = values.salary_year && values.salary_month
        ? `${values.salary_year}-${String(values.salary_month).padStart(2, '0')}` : '';
    document.getElementById('standard-hours').value = values.monthly_standard_hours ?? '';
    document.getElementById('manual-headers').value = values.manual_headers ? JSON.stringify(values.manual_headers, null, 2) : '';
    document.getElementById('multi-sheet-source').checked = !!values.multi_sheet_source;
    document.getElementById('use-history').checked = !!values.use_history;
    _sessionSettings = _readSessionSettings();
    document.getElementById('session-asset-actions').style.display = '';
    onModeChange();
}

function _renderSessionAssets(data) {
    const panel = document.getElementById('session-assets');
    const sessionId = data.session?.id || _currentSessionId;
    const expanded = panel.dataset.sessionId === String(sessionId)
        && !!panel.querySelector('details')?.open;
    panel.replaceChildren(); panel.style.display = '';
    panel.dataset.sessionId = String(sessionId);
    const details = document.createElement('details'); details.open = expanded;
    const heading = document.createElement('summary');
    heading.textContent = '当前训练文件（点击展开 / 收起）'; heading.style.cursor = 'pointer';
    details.append(heading); panel.append(details);
    const edit = document.createElement('button');
    edit.type = 'button'; edit.className = 'btn btn-action'; edit.textContent = '添加 / 替换文件与配置';
    edit.onclick = toggleAttachPopover; details.append(edit);
    for (const [label, category, names] of [
        ['源文件', 'source', data.source_file_names || []],
        ['目标文件', 'expected', data.expected_file_name ? [data.expected_file_name] : []],
        ['规则附件', 'rule-file', data.rule_file_names || []],
    ]) {
        const row = document.createElement('div'); row.append(document.createTextNode(`${label}：`));
        if (!names.length) row.append(document.createTextNode('未找到，可重新上传'));
        for (const name of names) {
            const link = document.createElement('a');
            link.href = '#'; link.className = 'file-download-link'; link.textContent = name;
            link.onclick = event => { event.preventDefault(); _downloadOriginalFile(sessionId, category, name); };
            row.append(link, document.createTextNode('　'));
        }
        details.append(row);
    }
    if (data.validation_stale) {
        const note = document.createElement('div');
        note.textContent = '输入已更新，历史评分仅对应旧文件；请执行修正或重新生成以验证当前输入。';
        panel.append(note); _currentAccuracy = null;
    }
}

async function saveSessionAssets() {
    if (!_currentSessionId || _savingAssets) return false;
    const sessionId = _currentSessionId;
    const status = document.getElementById('session-asset-status');
    const button = document.getElementById('save-session-assets');
    try {
        const settings = _readSessionSettings();
        const inputs = ['source-files', 'target-file', 'rule-files'].map(id => document.getElementById(id));
        const changed = Object.fromEntries(Object.entries(settings).filter(([key, value]) =>
            JSON.stringify(value) !== JSON.stringify(_sessionSettings?.[key])));
        if (!inputs.some(input => input.files.length) && !Object.keys(changed).length) return true;
        const form = new FormData();
        form.append('settings', JSON.stringify(changed));
        form.append('source_mode', document.getElementById('source-update-mode').value);
        form.append('file_passwords', JSON.stringify(_filePasswordsMap || {}));
        for (const file of inputs[0].files) form.append('source_files', file);
        if (inputs[1].files[0]) form.append('expected_result', inputs[1].files[0]);
        for (const file of inputs[2].files) form.append('rule_files', file);
        _savingAssets = true; button.disabled = true; _setUIStreaming(_isStreaming);
        status.textContent = '正在排队验证并保存，原文件在成功前保持可用…';
        const response = await AUTH.authFetch(`/api/training/chat/sessions/${sessionId}/assets`, {method: 'POST', body: form});
        const data = await response.json();
        if (!response.ok) throw new Error(data.detail || '保存失败');
        if (_currentSessionId === sessionId) {
            _resetUploadState(); _restoreSessionSettings(data.settings); _renderSessionAssets(data);
            status.textContent = '已保存，下一次操作使用新配置。';
        }
        return true;
    } catch (error) {
        status.textContent = `未保存：${error.message}`;
        alert(`文件与配置未保存：${error.message}`); return false;
    } finally {
        _savingAssets = false; button.disabled = false; _setUIStreaming(_isStreaming);
    }
}
