/** Key-set selection; the matching UI and field formulas remain unchanged. */
(() => {
    Tools._remapDataScopeFiles = function(config, fileMap) {
        if (!config) return;
        if (fileMap[config.source_file]) config.source_file = fileMap[config.source_file];
        // Resolved scheme expressions use JSON-quoted table names. Replace each
        // token once, so swapped same-structure roles cannot cascade.
        if (config.expression) config.expression = config.expression.replace(/"(?:[^"\\]|\\.)*"/g, token => {
            try {
                const value = JSON.parse(token);
                return Object.prototype.hasOwnProperty.call(fileMap, value) ? JSON.stringify(fileMap[value]) : token;
            } catch (_) { return token; }
        });
    };
    Tools._dataScopeEnabled = function() {
        return !!document.getElementById('int-union-source-keys')?.checked ||
            (this._integrationKind === 'split' && document.getElementById('split-include-source-keys')?.checked !== false);
    };
    Tools._readDataScope = function() {
        if (!this._dataScopeEnabled()) return {mode: 'all'};
        return {
            mode: document.getElementById('int-data-scope-mode')?.value || 'all',
            source_file: document.getElementById('int-data-scope-file')?.value || '',
            filter_expression: document.getElementById('int-data-scope-filter')?.value.trim() || '',
            expression: document.getElementById('int-data-scope-expression')?.value.trim() || '',
        };
    };
    Tools._refreshDataScope = function() {
        const box = document.getElementById('int-data-scope');
        if (!box) return;
        box.style.display = this._dataScopeEnabled() ? '' : 'none';
        const mode = document.getElementById('int-data-scope-mode').value;
        document.getElementById('int-data-scope-source').style.display = mode === 'source' ? '' : 'none';
        document.getElementById('int-data-scope-formula').style.display = mode === 'expression' ? '' : 'none';
        document.getElementById('int-data-scope-note').textContent = mode === 'all'
            ? '保留模板已有记录，补齐各数据表主键，保持原方案行为。'
            : '仅输出所选范围的有效主键，每个主键一条；模板和其他表只提供样式及关联字段，不额外增加人员。筛选结果为空时会提示调整条件。';
    };
    Tools._renderDataScope = function(config = {}) {
        let box = document.getElementById('int-data-scope');
        if (!box) {
            box = document.createElement('div');
            box.id = 'int-data-scope';
            document.getElementById('int-union-source-keys')?.closest('.form-group')?.insertAdjacentElement('afterend', box);
        }
        const files = this._intFiles().map(file => file.name);
        const source = config.source_file || this._intNonMainFiles()[0]?.name || files[0] || '';
        box.style.cssText = 'margin:8px 0;padding:10px;border:1px solid #dce8f8;border-radius:6px;';
        box.innerHTML = `
            <label>数据范围 <select id="int-data-scope-mode">
                <option value="all">所有表主键并集（原规则）</option>
                <option value="source">以指定表 / Sheet 为准</option>
                <option value="expression">按表达式确定</option>
            </select></label>
            <div id="int-data-scope-source" style="margin-top:8px;">
                <label>主数据表 <select id="int-data-scope-file">${this._intFiles().map(file => `<option value="${_escape(file.name)}" ${file.name === source ? 'selected' : ''}>${_escape(file.name)}${file.sheet ? ' · ' + _escape(file.sheet) : ''}</option>`).join('')}</select></label>
                <div style="margin-top:8px;"><label>筛选条件（可不填）
                    <input id="int-data-scope-filter" style="width:320px;max-width:100%;" value="${_escape(config.filter_expression || '')}" placeholder="例如：[状态]='在职'"></label></div>
            </div>
            <div id="int-data-scope-formula" style="margin-top:8px;">
                <textarea id="int-data-scope-expression" rows="3" style="width:100%;box-sizing:border-box;" placeholder='INTERSECT("B.xlsx","C.xlsx")'>${_escape(config.expression || '')}</textarea>
                <div style="margin:6px 0;display:flex;gap:6px;flex-wrap:wrap;">
                    <button type="button" class="btn btn-sm" data-scope-example="UNION">并集示例</button>
                    <button type="button" class="btn btn-sm" data-scope-example="INTERSECT">交集示例</button>
                    <button type="button" class="btn btn-sm" data-scope-example="EXCEPT">差集示例</button>
                    <button type="button" class="btn btn-sm" data-scope-example="FILTER">筛选示例</button>
                </div>
                <div style="font-size:12px;color:#777;">UNION 合并、INTERSECT 交集、EXCEPT 前表减后表；可嵌套，例如 EXCEPT(UNION("B.xlsx","C.xlsx"),"D.xlsx")。<br>FILTER("B.xlsx", "[状态]='在职'") 按该表字段筛选。表名须与上传列表一致；多 Sheet 请用列表中的完整表名。<br>可用表：${files.map(_escape).join('；')}</div>
            </div>
            <p id="int-data-scope-note" style="font-size:12px;color:#777;margin-bottom:0;"></p>`;
        const mode = document.getElementById('int-data-scope-mode');
        mode.value = config.mode || 'all';
        mode.addEventListener('change', () => this._refreshDataScope());
        box.querySelectorAll('[data-scope-example]').forEach(button => button.addEventListener('click', () => {
            const sources = this._intNonMainFiles().map(file => file.name);
            const first = sources[0] || files[0] || '';
            const second = sources[1] || files[0] || first;
            const op = button.dataset.scopeExample;
            const input = document.getElementById('int-data-scope-expression');
            input.value = op === 'FILTER' ? `FILTER(${JSON.stringify(first)}, "[状态]='在职'")`
                : `${op}(${JSON.stringify(first)},${JSON.stringify(second)})`;
        }));
        this._refreshDataScope();
    };

    const render = Tools._renderIntegrateConfig;
    Tools._renderIntegrateConfig = function(...args) {
        render.apply(this, args);
        this._renderDataScope();
        document.getElementById('int-union-source-keys')?.addEventListener('change', () => this._refreshDataScope());
    };
    const fill = Tools._intFillFromScheme;
    Tools._intFillFromScheme = function(scheme) {
        fill.call(this, scheme);
        this._renderDataScope(scheme?.config?.data_scope || {});
    };
})();
