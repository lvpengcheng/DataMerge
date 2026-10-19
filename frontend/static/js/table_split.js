/** Shared integration wizard, with independent split routes, schemes and permissions. */
(() => {
    const activateTab = Tools._activateTab;
    Tools._activateTab = function(tab) {
        if (tab === 'table-split' || tab === 'data-integrate') {
            const kind = tab === 'table-split' ? 'split' : 'integrate';
            if (kind !== (this._integrationKind || 'integrate') && this._integrationPending) {
                this._setIntegrateStatus('当前操作正在进行，请完成后再切换整合或拆表', '');
                return;
            }
            const permission = kind === 'split' ? 'tools.table_split' : 'tools.data_integrate';
            if (!AUTH.hasPerm(permission)) return;
            const panel = document.querySelector('#integrate-scheme-list-view')?.closest('.tool-page');
            const host = document.getElementById('tab-' + tab);
            if (panel && host) host.appendChild(panel);
            if (kind !== this._integrationKind) {
                this._integrationKind = kind;
                this._intShowList();
                _integrateSchemeList = [];
                _intSchemePage = 1;
                _intSchemeFilterKw = '';
                const filter = document.getElementById('int-scheme-filter');
                if (filter) filter.value = '';
            }
            if (panel) {
                panel.querySelector('h2').textContent = kind === 'split' ? '拆表' : '多表整合对比';
                const desc = document.querySelector('#integrate-scheme-list-view .desc');
                desc.textContent = kind === 'split'
                    ? '上传一个或多个 Excel，选择模板，按关联键整合后按字段或条件拆表。方案按组织共享，支持新增、应用、修改、另存为、导入、导出和删除。'
                    : '同组织/子组织创建的整合对比方案。可新增、应用、修改和删除；非创建人可另存为新方案。';
                const intro = document.querySelector('#integrate-work-view > .desc');
                intro.textContent = kind === 'split'
                    ? '可只上传一张表直接拆分，也可选择模板并整合其他数据表后拆分。保留所选 Sheet 的表头、数据行后的汇总与备注、格式和公式，并保留工作簿的其他 Sheet。结果以 ZIP 下载。'
                    : '选择主表模板，按关联键把对照表的值回填到主表，可选生成差异 Sheet，保留其他 Sheet、格式和未覆盖的公式。';
                for (const id of ['btn-int-new-scheme', 'btn-int-import-scheme']) {
                    document.getElementById(id).dataset.perm = permission + '.create';
                }
                const uploadHint = document.querySelector('#integrate-upload-zone > div');
                uploadHint.textContent = kind === 'split' ? '点击或拖拽 Excel 文件到此处（支持单表直接拆分或多表整合后拆分）'
                    : '点击或拖拽 Excel 文件到此处（多文件模式需 ≥2 个；全部 Sheet 模式可只上传 1 个）';
                AUTH.applyPermFilter(panel);
            }
        }
        activateTab.call(this, tab);
        if (tab === 'table-split') this.loadIntegrateSchemes();
    };

    // Do not let a late response populate a wizard for the other workflow.
    for (const name of ['loadIntegrateSchemes', '_analyzeIntegrate', '_intReparseHeaders',
        '_intSaveScheme', '_doIntegrateExecute', '_intDeleteSchemeRow', '_intExportScheme',
        '_intImportScheme', '_intPrefillForApply', '_intPrefillForEdit']) {
        const original = Tools[name];
        if (!original) continue;
        Tools[name] = async function(...args) {
            this._integrationPending = (this._integrationPending || 0) + 1;
            try { return await original.apply(this, args); }
            finally { this._integrationPending--; }
        };
    }

    Tools._readSplitConfig = function() {
        const integer = id => {
            const value = document.getElementById(id)?.value.trim();
            return value ? Number(value) : null;
        };
        return {
            mode: document.getElementById('split-group-mode')?.value || 'value',
            column: document.getElementById('split-group-column')?.value || '',
            expression: document.getElementById('split-group-expression')?.value || '',
            data_row_start: integer('split-data-start'), data_row_end: integer('split-data-end'),
            include_source_keys: document.getElementById('split-include-source-keys')?.checked !== false,
        };
    };

    Tools._renderSplitSettings = function(config = {}) {
        if (this._integrationKind !== 'split') return;
        let box = document.getElementById('table-split-settings');
        if (!box) {
            box = document.createElement('div');
            box.id = 'table-split-settings';
            const target = document.getElementById('int-execute')?.closest('.actions') ||
                document.getElementById('int-save-row');
            const container = document.getElementById('integrate-config');
            container.insertBefore(box, target?.parentNode === container ? target : container.firstChild);
        }
        const main = this._intFileMeta(this._intMainFile()) || {};
        const columns = main.columns || [];
        const column = config.column || columns.find(c => c.includes('工作地')) || columns[0] || '';
        const detected = main.data_row_start && main.data_row_end ? `（当前自动识别第${main.data_row_start}至${main.data_row_end}行）` : '';
        box.style.cssText = 'margin:18px 0;padding:16px;background:#f3f8ff;border:1px solid #cbdcf2;border-radius:8px;';
        box.innerHTML = `
            <h3 style="margin-top:0;">⑦ 拆表规则</h3>
            <p>整合完成后按本规则分组，每组生成一个 Excel；只删除该组以外的数据行，保留上下方内容并更新相关公式引用。</p>
            <label style="display:block;margin-bottom:10px;"><input id="split-include-source-keys" type="checkbox" ${config.include_source_keys !== false ? 'checked' : ''} style="width:auto;"> 纳入数据表中模板未包含的记录（默认启用）</label>
            <p style="font-size:13px;color:#555;">例如模板只有苏州、数据表还有深圳：启用后两地都会生成文件；关闭后仅拆模板已有记录。按关联键补入新记录，字段值仍按上方映射回填，不会自动累加模板金额。结果命名为“分组值_主表名称.xlsx”。</p>
            <div style="display:flex;gap:16px;flex-wrap:wrap;margin-bottom:10px;">
                <label>数据起始行 <input id="split-data-start" type="number" min="1" step="1" value="${config.data_row_start || ''}" placeholder="自动识别" style="width:110px;"></label>
                <label>数据结束行 <input id="split-data-end" type="number" min="1" step="1" value="${config.data_row_end || ''}" placeholder="自动识别" style="width:110px;"></label>
            </div>
            <p style="font-size:13px;color:#555;">行号从1开始，只包含明细，不要包含表头、合计或备注。留空时每次自动识别${_escape(detected)}；新增整合行会自动纳入拆分。</p>
            <label>拆分方式 <select id="split-group-mode"><option value="value">按字段值：每个不同值一份</option><option value="expression">按表达式：条件结果相同的放一份</option></select></label>
            <div id="split-column-row" style="margin-top:10px;"><label>拆分字段 <select id="split-group-column">${this._optList(columns, column)}</select></label></div>
            <div id="split-expression-row" style="margin-top:10px;display:none;">
                <label for="split-group-expression">分组表达式</label>
                <textarea id="split-group-expression" rows="3" style="width:100%;box-sizing:border-box;" placeholder='IF(工作地="上海","上海","其他")'>${_escape(config.expression || '')}</textarea>
                <div style="display:flex;gap:8px;flex-wrap:wrap;margin:8px 0;">
                    <button class="btn btn-sm" type="button" id="split-insert-column">插入所选字段</button>
                    <button class="btn btn-sm" type="button" id="split-example-city">示例：上海 / 其他</button>
                </div>
                <p style="font-size:13px;color:#555;">字段名含空格或符号时写成 [列名]。支持 IF、AND、OR、NOT、CONTAINS、STARTSWITH、ENDSWITH，及 =、&lt;&gt;、&gt;、&gt;=、&lt;、&lt;= 和四则运算。文本用英文双引号。<br>
                例：IF([工作地]="上海","上海","其他")；IF(AND([工作地]="上海",[金额]&gt;=5000),"上海高金额","其他")。布尔结果自动分为“满足条件 / 不满足条件”，空结果归入“空白”，每次最多200组。</p>
            </div>`;
        const mode = document.getElementById('split-group-mode');
        mode.value = config.mode || 'value';
        const refresh = () => {
            document.getElementById('split-expression-row').style.display = mode.value === 'expression' ? '' : 'none';
            document.getElementById('split-column-row').style.display = '';
        };
        mode.addEventListener('change', refresh);
        refresh();
        document.getElementById('split-insert-column').addEventListener('click', () => {
            const input = document.getElementById('split-group-expression');
            const value = document.getElementById('split-group-column').value;
            input.setRangeText(`[${value}]`, input.selectionStart, input.selectionEnd, 'end');
            input.focus();
        });
        document.getElementById('split-example-city').addEventListener('click', () => {
            const value = document.getElementById('split-group-column').value;
            document.getElementById('split-group-expression').value = `IF([${value}]="上海","上海","其他")`;
        });
        const button = document.getElementById('int-execute');
        if (button) button.textContent = '整合并拆表下载';
    };

    const render = Tools._renderIntegrateConfig;
    Tools._renderIntegrateConfig = function(...args) {
        render.apply(this, args);
        this._renderSplitSettings();
        document.getElementById('int-main-file')?.addEventListener('change', () => this._renderSplitSettings());
    };
    const fill = Tools._intFillFromScheme;
    Tools._intFillFromScheme = function(scheme) {
        fill.call(this, scheme);
        this._renderSplitSettings(scheme?.config?.split_config || {});
    };
})();
