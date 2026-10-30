/** Each target owns its editor state; uploads and header discovery are shared. */
(() => {
    const clone = value => JSON.parse(JSON.stringify(value));
    const enabled = () => Tools._integrationKind !== 'split';
    const allFiles = () => (_integrateAnalysis?.files || []);
    const originalFiles = Tools._intFiles;
    Tools._intFiles = function() {
        const files = originalFiles.call(this);
        const task = enabled() && this._outputTasks?.[this._outputTaskIndex];
        return task ? files.filter(f => task.active_files.includes(f.name)) : files;
    };
    Tools._captureOutputTask = function() {
        const current = this._outputTasks?.[this._outputTaskIndex];
        if (!current || !document.getElementById('int-main-file')) return;
        const key_map = {};
        const active = new Set(current.active_files);
        document.querySelectorAll('.int-key').forEach(s => {
            if (active.has(s.dataset.file)) key_map[s.dataset.file] = s.value;
        });
        Object.assign(current, {
            main_file: this._intMainFile(), key_map,
            overwrite_pairs: this._readSectionPairs('ow'), compare_pairs: this._readSectionPairs('cmp'),
            output_mode: Number(document.querySelector('[name="int-output-mode"]:checked')?.value || 1),
            name_col: document.getElementById('int-name-col')?.value || null,
            id_col: document.getElementById('int-id-col')?.value || null,
            diff_order: document.getElementById('int-diff-order')?.value || 'id_name',
            date_key_mode: document.getElementById('int-date-mode')?.value || 'off', normalize_keys: true,
            union_source_keys: !!document.getElementById('int-union-source-keys')?.checked,
            data_scope: this._readDataScope(),
            output_file: document.getElementById('int-task-output-file')?.value.trim() || current.output_file,
            output_sheet: document.getElementById('int-task-output-sheet')?.value.trim() || current.output_sheet,
        });
    };
    const fill = Tools._intFillFromScheme;
    const render = Tools._renderIntegrateConfig;
    Tools._showOutputTask = function(index) {
        const schemeName = document.getElementById('int-scheme-name')?.value;
        this._outputTaskIndex = index;
        const task = this._outputTasks[index];
        const files = allFiles().filter(f => task.active_files.includes(f.name));
        files.sort((a, b) => Number(b.name === task.main_file) - Number(a.name === task.main_file));
        render.call(this, {..._integrateAnalysis, files});
        if (schemeName != null) document.getElementById('int-scheme-name').value = schemeName;
        const pairs = kind => (task[kind] || []).map(p => ({...p, source_role: task.active_files.indexOf(p.source_file)}));
        fill.call(this, {main_file: task.main_file, role_files: task.active_files, config: {
            ...task, key_map_by_role: Object.fromEntries(task.active_files.map((f,i) => [i, task.key_map[f]])),
            overwrite_pairs: pairs('overwrite_pairs'), compare_pairs: pairs('compare_pairs'),
        }});
        this._renderOutputTaskTabs();
        // Template selection is handled by the target picker to avoid erasing formulas.
        document.getElementById('int-main-file').disabled = true;
        if (_intMode === 'apply') document.getElementById('int-execute').disabled = !_intApplyReady;
    };
    Tools._renderOutputTaskTabs = function() {
        document.getElementById('int-output-tasks')?.remove();
        const task = this._outputTasks[this._outputTaskIndex];
        const box = document.createElement('section');
        box.id = 'int-output-tasks';
        box.style.cssText = 'padding:14px;margin-bottom:18px;border:1px solid #cbdcf2;border-radius:8px;background:#f4f8ff;';
        box.innerHTML = `<div style="display:flex;gap:8px;flex-wrap:wrap;" role="tablist" aria-label="目标Sheet配置">
            ${this._outputTasks.map((t,i) => `<button type="button" role="tab" aria-selected="${i === this._outputTaskIndex}" class="btn ${i === this._outputTaskIndex ? 'btn-primary' : ''}" data-output-task="${i}">${_escape(t.output_sheet)}</button>`).join('')}
            <button type="button" class="btn" id="int-add-targets">＋ 选择目标Sheet</button>
            <button type="button" class="btn" id="int-copy-target">复制当前配置到…</button>
            <button type="button" class="btn" id="int-remove-target" ${this._outputTasks.length < 2 ? 'disabled' : ''}>移除当前目标</button></div>
            <p>每个页签独立设置来源、关联键、数据范围和公式。切换自动保留；相同输出文件名合并为一个工作簿，不同文件名打包下载。</p>
            <label>输出文件名 <input id="int-task-output-file" value="${_escape(task.output_file)}" style="min-width:240px;"></label>
            <label style="margin-left:12px;">输出Sheet名 <input id="int-task-output-sheet" maxlength="31" value="${_escape(task.output_sheet)}"></label>
            <details style="margin-top:12px;" open><summary>本目标使用的来源Sheet</summary>
            ${allFiles().filter(f => f.name !== task.main_file).map(f => `<label style="display:block;margin:6px 0;"><input type="checkbox" class="int-task-source" value="${_escape(f.name)}" ${task.active_files.includes(f.name) ? 'checked' : ''} style="width:auto;"> ${_escape(f.physical_file || f.name)} · ${_escape(f.sheet || '')}</label>`).join('')}</details>
            <div style="font-size:12px;color:#666;margin-top:8px;">同一输出文件使用同一个模板工作簿，可保留其他Sheet及跨Sheet公式。不同模板请选择不同输出文件名。</div>`;
        document.getElementById('integrate-config').prepend(box);
        document.getElementById('int-main-file').onchange = () => {
            const meta = allFiles().find(f => f.name === this._intMainFile());
            task.main_file = meta.name;
            this._captureOutputTask();
            task.output_sheet = meta.sheet || '结果';
            task.output_file = '整合结果_' + (meta.physical_file || meta.name).replace(/\.(xlsx|xlsm|xls)$/i, '') + '.xlsx';
            this._renderOutputTaskTabs();
        };
        box.querySelectorAll('[data-output-task]').forEach(button => button.onclick = () => {
            this._captureOutputTask(); this._showOutputTask(Number(button.dataset.outputTask));
        });
        box.querySelectorAll('.int-task-source').forEach(input => input.onchange = () => {
            this._captureOutputTask();
            if (!input.checked && [...task.overwrite_pairs, ...task.compare_pairs].some(p =>
                p.source_file === input.value || (p.source_expr || p.source_col || '').replace(/"(?:[^"]|"")*"/g, '').includes(input.value + '.'))) {
                input.checked = true;
                this._setIntegrateStatus('该来源仍被公式引用，请先修改相关公式再取消选择', 'error');
                return;
            }
            task.active_files = [task.main_file, ...[...box.querySelectorAll('.int-task-source:checked')].map(el => el.value)];
            const removed = new Set(Object.keys(task.key_map).filter(f => !task.active_files.includes(f)));
            // Do not silently discard formulas; execution identifies invalid references.
            for (const f of removed) delete task.key_map[f];
            for (const f of task.active_files) if (!task.key_map[f]) {
                const meta = allFiles().find(x => x.name === f); task.key_map[f] = meta.suggested_key || meta.columns[0];
            }
            this._showOutputTask(this._outputTaskIndex);
        });
        document.getElementById('int-add-targets').onclick = () => this._chooseOutputTargets(false);
        document.getElementById('int-copy-target').onclick = () => this._chooseOutputTargets(true);
        document.getElementById('int-remove-target').onclick = () => {
            this._outputTasks.splice(this._outputTaskIndex, 1); this._showOutputTask(0);
        };
    };
    Tools._chooseOutputTargets = function(copyCurrent) {
        this._captureOutputTask();
        const current = this._outputTasks[this._outputTaskIndex];
        const files = allFiles();
        const books = [...new Set(files.map(f => f.physical_file || f.name))];
        this.openModal(copyCurrent ? '复制配置到目标Sheet' : '选择模板和目标Sheet', `<p>选择模板文件，再勾选需要配置的目标Sheet。解析出的其他Sheet可作为数据来源。</p>
            <select id="int-target-book">${books.map(b => `<option>${_escape(b)}</option>`).join('')}</select>
            <div id="int-target-sheet-choices" style="margin-top:12px;"></div>`, () => {
            const selected = [...document.querySelectorAll('.int-new-target:checked:not(:disabled)')].map(el => el.value);
            if (!selected.length) {this.closeModal(); return;}
            for (const name of selected) {
                const meta = files.find(f => f.name === name);
                const task = copyCurrent ? clone(current) : {overwrite_pairs: [], compare_pairs: [], output_mode: 1, data_scope: {mode:'all'}};
                const oldMain = task.main_file;
                task.main_file = name;
                task.output_sheet = meta.sheet || '结果';
                task.output_file = '整合结果_' + (meta.physical_file || meta.name).replace(/\.(xlsx|xlsm|xls)$/i, '') + '.xlsx';
                task.active_files = [name, ...files.filter(f => f.name !== name && !selected.includes(f.name) && !this._outputTasks.some(t => t.main_file === f.name)).map(f => f.name)];
                if (copyCurrent) {
                    task.active_files = [name, ...current.active_files.filter(f => f !== oldMain && f !== name)];
                    for (const kind of ['overwrite_pairs', 'compare_pairs']) {
                        task[kind] = task[kind].filter(p => meta.columns.includes(p.a_col)).map(p => {
                            const expr = p.source_expr || p.source_col || '';
                            const prefix = (oldMain + '.').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
                            const next = expr.replace(new RegExp('"(?:[^\"]|\"\")*"|' + prefix, 'g'), token => token.startsWith('"') ? token : name + '.');
                            return {...p, source_file: p.source_file === oldMain ? name : p.source_file, source_expr: next, source_col: next};
                        });
                    }
                    this._remapDataScopeFiles(task.data_scope, {[oldMain]: name});
                }
                task.key_map = Object.fromEntries(task.active_files.map(f => {
                    const info = files.find(x => x.name === f); return [f, current.key_map[f] || info.suggested_key || info.columns[0]];
                }));
                while (this._outputTasks.some(t => t.output_file === task.output_file && t.output_sheet === task.output_sheet)) {
                    task.output_file = '副本_' + task.output_file;
                }
                this._outputTasks.push(task);
            }
            this.closeModal(); this._showOutputTask(this._outputTasks.length - selected.length);
        }, {confirmText: copyCurrent ? '复制并配置' : '添加目标', wide: true});
        const select = document.getElementById('int-target-book');
        const populate = () => {
            document.getElementById('int-target-sheet-choices').innerHTML = files.filter(f => (f.physical_file || f.name) === select.value).map(f => {
                const exists = !copyCurrent && this._outputTasks.some(t => t.main_file === f.name);
                return `<label style="display:block;margin:8px;"><input class="int-new-target" type="checkbox" value="${_escape(f.name)}" ${exists ? 'checked disabled' : ''} style="width:auto;"> ${_escape(f.sheet || f.name)}${exists ? '（已添加）' : ''}</label>`;
            }).join('');
        };
        select.value = files.find(f => f.name === current.main_file)?.physical_file || current.main_file;
        select.onchange = populate; populate();
    };
    Tools._renderIntegrateConfig = function(data) {
        this._outputTasks = null;
        render.call(this, data);
        if (!enabled()) return;
        const main = this._intMainFile(), meta = allFiles().find(f => f.name === main);
        this._outputTaskIndex = 0;
        this._outputTasks = [{main_file: main, active_files: allFiles().map(f => f.name), key_map: {},
            output_file: '整合结果_' + (meta.physical_file || meta.name).replace(/\.(xlsx|xlsm|xls)$/i, '') + '.xlsx', output_sheet: meta.sheet || '结果'}];
        this._captureOutputTask(); this._renderOutputTaskTabs();
    };
    Tools._intFillFromScheme = function(scheme) {
        if (!enabled()) return fill.call(this, scheme);
        this._outputTasks = null;
        fill.call(this, scheme);
        if (scheme.config?.output_tasks?.length) {
            this._outputTasks = clone(scheme.config.output_tasks);
            this._showOutputTask(0);
        } else {
            const meta = allFiles().find(f => f.name === scheme.main_file);
            this._outputTaskIndex = 0;
            this._outputTasks = [{main_file: scheme.main_file, active_files: scheme.role_files || allFiles().map(f => f.name),
                output_file: '整合结果_' + (meta.physical_file || meta.name).replace(/\.(xlsx|xlsm|xls)$/i, '') + '.xlsx', output_sheet: meta.sheet || '结果'}];
            this._captureOutputTask();
            // Parsing all sheets can discover tables absent from an older scheme.
            // Render only its saved roles, matching the task's submitted key map.
            this._showOutputTask(0);
        }
    };
    const fetch = Tools._intFetch;
    Tools._intFetch = async function(url, options) {
        if (enabled() && this._outputTasks?.length && /\/(execute|scheme\/save)(\?|$)/.test(url)) {
            this._captureOutputTask();
            const payload = JSON.parse(options.body);
            for (const task of this._outputTasks) {
                if (!task.overwrite_pairs.length) throw new Error(`目标「${task.output_sheet}」尚未设置覆盖字段和公式`);
            }
            Object.assign(payload, clone(this._outputTasks[0]), {output_tasks: clone(this._outputTasks)});
            options = {...options, body: JSON.stringify(payload)};
        }
        return fetch.call(this, url, options);
    };
    const show = Tools._intShowWork;
    Tools._intShowWork = function(...args) {this._outputTasks = null; return show.apply(this,args);};
})();
