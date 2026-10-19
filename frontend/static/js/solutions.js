/* Immutable solution lifecycle, deliberately separate from editable training. */
let solutionDialog;
async function solutionRequest(url, options) {
    const response = await AUTH.authFetch(url, options);
    if (!response.ok) {
        const data = await response.json().catch(() => ({}));
        throw new Error(typeof data.detail === 'string' ? data.detail : '操作失败，请检查输入或稍后重试');
    }
    return response;
}
function solutionElement(tag, text, parent) {
    const el = document.createElement(tag);
    if (text) el.textContent = text;
    if (parent) parent.appendChild(el);
    return el;
}
async function openSolutionTraining(row) {
    if ((typeof _isStreaming !== 'undefined' && _isStreaming) ||
        (typeof _savingAssets !== 'undefined' && _savingAssets)) {
        throw new Error('方案已保存。请等待当前操作结束，再点击该方案的“进入修正”。');
    }
    row = await (await solutionRequest('/api/solutions/' + row.id + '/training', {method: 'POST'})).json();
    if (!row.session_id) throw new Error('训练会话恢复失败，请稍后重试。');
    await selectSession(row.session_id);
    if (String(_currentSessionId) !== String(row.session_id)) throw new Error('训练会话加载失败，请重新点击“进入修正”。');
    document.getElementById('tenant-input').value = row.tenant_id;
    await loadSessions();
    solutionDialog.close();
}
async function openSolutions() {
    if (solutionDialog?.dataset.busy === 'true') {
        if (!solutionDialog.open) solutionDialog.showModal();
        return;
    }
    if (!solutionDialog) {
        solutionDialog = document.createElement('dialog');
        solutionDialog.className = 'solutions-dialog';
        solutionDialog.setAttribute('aria-labelledby', 'solutions-title');
        document.body.appendChild(solutionDialog);
    }
    solutionDialog.replaceChildren();
    const heading = solutionElement('div', '', solutionDialog);
    heading.className = 'solutions-heading';
    solutionElement('h2', '计算方案', heading).id = 'solutions-title';
    const close = solutionElement('button', '关闭', heading);
    close.dataset.solutionClose = 'true';
    close.onclick = () => solutionDialog.close();
    const body = solutionElement('div', '', solutionDialog);
    body.className = 'solutions-body';
    const intro = solutionElement('p', '选择已保存的脚本即可转方案，当前训练可以继续。方案固定本次转换时的代码和资料，后续训练不会改变该版本。', body);
    intro.className = 'solutions-intro';
    const tenantSection = solutionElement('section', '', body);
    tenantSection.className = 'solutions-card';
    solutionElement('h3', '选择租户', tenantSection);
    const tenantRow = solutionElement('div', '', tenantSection);
    tenantRow.className = 'solutions-tenant-row';
    const tenantLabel = solutionElement('label', '租户名称', tenantRow);
    tenantLabel.htmlFor = 'solutions-tenant-input';
    const tenant = solutionElement('input', '', tenantRow);
    tenant.id = 'solutions-tenant-input';
    tenant.value = (typeof _currentTenantId !== 'undefined' && _currentTenantId) || document.getElementById('tenant-input')?.value || '';
    tenant.placeholder = '选择现有租户或输入新租户名称';
    const listId = 'solution-tenant-list';
    tenant.setAttribute('list', listId);
    const tenants = solutionElement('datalist', '', solutionDialog);
    tenants.id = listId;
    solutionRequest('/api/tenants').then(r => r.json()).then(data => {
        for (const item of (Array.isArray(data) ? data : data.tenants || [])) {
            const option = solutionElement('option', '', tenants);
            option.value = typeof item === 'string' ? item : item.tenant_id || item.name || '';
        }
    }).catch(() => {});
    const status = solutionElement('p', '', body);
    status.className = 'solutions-status';
    status.setAttribute('role', 'status');
    status.setAttribute('aria-live', 'polite');
    const content = solutionElement('div', '', body);
    content.className = 'solutions-content';
    const importArea = solutionElement('section', '', body);
    importArea.className = 'solutions-card';
    solutionElement('h3', '导入完整方案包', importArea);
    solutionElement('p', '选择目标脚本导入一个方案版本，或输入新脚本名称创建。租户仅用于确定脚本归属。原方案版本仍保留，导入版本验收后启用。', importArea).className = 'solutions-muted';
    const targetRow = solutionElement('div', '', importArea);
    targetRow.className = 'solutions-form-row';
    const targetLabel = solutionElement('label', '目标脚本', targetRow);
    const targetScript = solutionElement('select', '', targetLabel);
    targetScript.id = 'solutions-import-script';
    solutionElement('option', '创建新脚本', targetScript).value = '';
    const nameLabel = solutionElement('label', '新脚本名称', targetRow);
    const scriptName = solutionElement('input', '', nameLabel);
    scriptName.id = 'solutions-import-name';
    scriptName.placeholder = '输入尚未创建的脚本名称';
    function updateImportTarget() { nameLabel.hidden = !!targetScript.value; }
    targetScript.onchange = updateImportTarget;
    const importRow = solutionElement('div', '', importArea);
    importRow.className = 'solutions-form-row';
    const fileLabel = solutionElement('label', '方案文件（ZIP）', importRow);
    const file = solutionElement('input', '', fileLabel);
    file.type = 'file'; file.accept = '.zip';
    const newLabel = solutionElement('label', '', importArea);
    newLabel.className = 'solutions-checkbox';
    const newTenant = solutionElement('input', '', newLabel); newTenant.type = 'checkbox';
    newLabel.append('创建上方填写的新租户');
    const trustLabel = solutionElement('label', '', importArea);
    trustLabel.className = 'solutions-checkbox';
    const trusted = solutionElement('input', '', trustLabel); trusted.type = 'checkbox';
    trustLabel.append('我已确认来源可信，允许导入其中的可执行代码');
    const importActions = solutionElement('div', '', importArea);
    importActions.className = 'solutions-actions';
    const importButton = solutionElement('button', '导入方案', importActions);
    const reload = solutionElement('button', '加载脚本和方案', tenantRow);
    let busy = false;
    let loadedTenant = null;
    async function action(message, fn) {
        if (busy) return;
        busy = true; status.textContent = message;
        solutionDialog.dataset.busy = 'true';
        const controls = Array.from(solutionDialog.querySelectorAll('button:not([data-solution-close]), input, select, textarea'));
        const disabled = controls.map(control => control.disabled);
        controls.forEach(control => control.disabled = true);
        try { await fn(); }
        catch (error) { status.textContent = error.message; }
        finally {
            busy = false; solutionDialog.dataset.busy = 'false';
            controls.forEach((control, index) => { if (control.isConnected) control.disabled = disabled[index]; });
        }
    }
    async function refresh() {
        const data = await (await solutionRequest('/api/solutions?tenant_id=' + encodeURIComponent(tenant.value.trim()))).json();
        const previous = loadedTenant === tenant.value.trim() ? targetScript.value : null;
        loadedTenant = tenant.value.trim();
        targetScript.replaceChildren();
        solutionElement('option', '创建新脚本', targetScript).value = '';
        for (const script of data.scripts) {
            const option = solutionElement('option', script.name + '（脚本 ' + script.id + '）', targetScript);
            option.value = script.id;
            if (previous === null && typeof _currentSessionId !== 'undefined' && String(script.session_id) === String(_currentSessionId)) option.selected = true;
        }
        if (previous !== null && Array.from(targetScript.options).some(o => o.value === previous)) targetScript.value = previous;
        updateImportTarget();
        content.replaceChildren();
        const convertCard = solutionElement('section', '', content);
        convertCard.className = 'solutions-card';
        solutionElement('h3', '将脚本转为方案', convertCard);
        const convertRow = solutionElement('div', '', convertCard);
        convertRow.className = 'solutions-form-row';
        const scriptLabel = solutionElement('label', '已保存的脚本', convertRow);
        const select = solutionElement('select', '', scriptLabel);
        for (const script of data.scripts) {
            const option = solutionElement('option', script.name + '（脚本 ' + script.id + '）', select);
            option.value = script.id;
            if (typeof _currentSessionId !== 'undefined' && String(script.session_id) === String(_currentSessionId)) option.selected = true;
        }
        const convert = solutionElement('button', '转为方案', convertRow);
        convert.className = 'solutions-primary';
        convert.disabled = !data.scripts.length;
        convert.onclick = () => action('正在保存代码和训练资料…', async () => {
            if (!select.value) throw new Error('请先在智训中将代码设为最佳脚本');
            const row = await (await solutionRequest('/api/solutions', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({script_id:select.value, tenant_id:tenant.value.trim()})})).json();
            await refresh(); status.textContent = '方案 ' + row.version + ' 已保存，请验收后启用。';
        });
        if (!data.scripts.length) solutionElement('p', '暂无已保存脚本，请先将需要的代码设为最佳，再转方案。', convertCard).className = 'solutions-muted';
        const versions = solutionElement('section', '', content);
        versions.className = 'solutions-versions';
        function renderVersions() {
        versions.replaceChildren();
        solutionElement('h3', '所选脚本的方案版本', versions);
        const rows = data.solutions.filter(row => String(row.script_id) === select.value);
        if (!rows.length) solutionElement('p', '当前脚本尚无方案，可从上方脚本转换或在下方导入。', versions).className = 'solutions-muted';
        for (const row of rows) {
            const box = solutionElement('article', '', versions);
            box.className = 'solutions-card';
            solutionElement('h4', row.name, box);
            solutionElement('p', '版本 ' + row.version, box).className = 'solutions-muted';
            solutionElement('p', row.is_active ? '已启用：智算可直接选择此版本' : '待验收 / 启用', box);
            if (row.validation) solutionElement('p', row.validation.error || (row.validation.passed ? '验收通过：结果与训练期望一致' : '验收未通过，请检查训练结果'), box);
            if (row.validation?.completed) {
                solutionElement('p', '实测准确率：' + (row.validation.accuracy * 100).toFixed(2) + '%；差异单元格：' + row.validation.total_differences, box);
                const details = solutionElement('details', '', box);
                solutionElement('summary', '查看差异证据', details);
                const pre = solutionElement('pre', JSON.stringify(row.validation.details || {}, null, 2), details);
                pre.style.cssText = 'white-space:pre-wrap;max-height:240px;overflow:auto';
            }
            const actions = solutionElement('div', '', box);
            actions.className = 'solutions-actions';
            function button(label, fn) {
                const b = solutionElement('button', label, actions);
                b.onclick = () => action('正在' + label + '，请稍候…', fn);
            }
            button('进入修正', () => openSolutionTraining(row));
            button('验收', async () => {
                await solutionRequest('/api/solutions/' + row.id + '/validate', {method:'POST'});
                await refresh(); status.textContent = '验收完成，请查看结果。';
            });
            if (!row.is_active && row.validation?.passed) button('启用此版本', async () => {
                await solutionRequest('/api/solutions/' + row.id + '/publish', {method:'POST'});
                await refresh(); status.textContent = '已启用，智算将使用此方案冻结的代码。';
            });
            if (!row.is_active && row.validation?.completed && !row.validation.passed && !row.validation.formula_warning && !(row.validation.missing_sheets || []).length) {
                const reason = solutionElement('textarea', '', box);
                box.insertBefore(reason, actions);
                reason.placeholder = '如果这些业务差异符合预期，请填写认可原因（保留实测分数）';
                button('确认接受差异并启用', async () => {
                    if (!reason.value.trim()) throw new Error('请先查看差异证据并填写认可原因');
                    await solutionRequest('/api/solutions/' + row.id + '/publish', {method:'POST', headers:{'Content-Type':'application/json'}, body:JSON.stringify({accept_differences:true, reason:reason.value.trim()})});
                    await refresh(); status.textContent = '已记录差异认可并启用，实测分数保持不变。';
                });
            }
            button('导出完整包', async () => {
                const blob = await (await solutionRequest('/api/solutions/' + row.id + '/export')).blob();
                const url = URL.createObjectURL(blob), a = document.createElement('a');
                a.href = url; a.download = 'solution_' + row.version + '.zip'; a.click();
                setTimeout(() => URL.revokeObjectURL(url), 30000);
                status.textContent = '完整方案包已导出。';
            });
        }
        }
        select.onchange = renderVersions;
        renderVersions();
    }
    reload.onclick = () => action('正在加载…', async () => { await refresh(); status.textContent = ''; });
    importButton.onclick = () => action('正在导入并恢复训练资料…', async () => {
        if (!tenant.value.trim() || !file.files.length || !trusted.checked) throw new Error('请填写租户、选择 ZIP，并确认代码来源可信');
        if (loadedTenant !== tenant.value.trim()) throw new Error('租户已变更，请先点击“加载脚本和方案”再选择目标脚本');
        if (!targetScript.value && !scriptName.value.trim()) throw new Error('请输入新脚本名称，或选择现有脚本');
        const form = new FormData();
        form.append('file', file.files[0]); form.append('tenant_id', tenant.value.trim());
        form.append('create_tenant', newTenant.checked); form.append('trust_code', trusted.checked);
        if (targetScript.value) form.append('script_id', targetScript.value);
        else form.append('script_name', scriptName.value.trim());
        const row = await (await solutionRequest('/api/solutions/import', {method:'POST', body:form})).json();
        await refresh();
        status.textContent = '导入完成。请在本环境验收后启用。';
        await openSolutionTraining(row);
    });
    if (!solutionDialog.open) solutionDialog.showModal();
    if (tenant.value.trim()) await action('正在加载…', async () => { await refresh(); status.textContent = ''; });
}
