/** Select participating sheets and confirm saved-sheet roles before applying. */
Object.assign(Tools, {
    _intChooseSheets(analysis) {
        return new Promise(resolve => {
            const groups = new Map();
            (analysis.workbooks || []).forEach(book => groups.set(book.file, []));
            (analysis.files || []).forEach(file => {
                const name = file.physical_file || file.name;
                if (!groups.has(name)) groups.set(name, []);
                groups.get(name).push(file);
            });
            const minimum = this._integrationKind === 'split' ? 1 : 2;
            this.openModal('选择参与的 Sheet', `<p>勾选每个 Excel 要保留参与的 Sheet，至少选择 ${minimum} 张。未勾选的 Sheet 不参与后续设置和计算。</p>
                ${[...groups].map(([name, files], i) => `<fieldset class="int-sheet-group" style="margin:12px 0;padding:12px;">
                    <legend>${_escape(name)}</legend>
                    <label><input type="checkbox" class="int-sheet-select-all" data-group="${i}" ${files.length ? 'checked' : 'disabled'} style="width:auto;"> 全选 / 全不选</label>
                    ${files.map(file => `<label style="display:block;margin:7px 0;"><input type="checkbox" class="int-choose-sheet" data-group="${i}" value="${_escape(file.name)}" checked style="width:auto;"> ${_escape(file.sheet)} <span style="color:#888;font-size:12px;">（${file.columns.length}列）</span></label>`).join('')}
                    ${((analysis.workbooks || []).find(book => book.file === name)?.sheets || []).filter(sheet => !sheet.selectable).map(sheet => `<label style="display:block;margin:7px 0;color:#888;"><input type="checkbox" disabled style="width:auto;"> ${_escape(sheet.sheet)}（未识别到有效列头）</label>`).join('')}
                    </fieldset>`).join('')}
                <div id="int-sheet-choice-error" style="color:#c00;"></div>`, () => {
                _modalCancelCallback = () => resolve(null);
                const selected = [...document.querySelectorAll('.int-choose-sheet:checked')].map(input => input.value);
                if (selected.length < minimum) {
                    document.getElementById('int-sheet-choice-error').textContent = `请至少选择${minimum}张Sheet`;
                    return;
                }
                resolve(selected);
                this.closeModal();
            }, {wide: true, confirmText: '确认并继续', onCancel: () => resolve(null)});
            document.querySelectorAll('.int-sheet-select-all').forEach(input => input.addEventListener('change', () => {
                document.querySelectorAll(`.int-choose-sheet[data-group="${input.dataset.group}"]`).forEach(cb => {cb.checked = input.checked;});
            }));
            document.querySelectorAll('.int-choose-sheet').forEach(input => input.addEventListener('change', () => {
                const group = document.querySelector(`.int-sheet-select-all[data-group="${input.dataset.group}"]`);
                const boxes = [...document.querySelectorAll(`.int-choose-sheet[data-group="${input.dataset.group}"]`)];
                group.checked = boxes.every(cb => cb.checked);
                group.indeterminate = boxes.some(cb => cb.checked) && !group.checked;
            }));
        });
    },

    _intChooseSheetRelations(info, schemeId) {
        return new Promise(resolve => {
            const tables = new Map(info.tables.map(table => [table.name, table]));
            this.openModal('确认方案与当前 Sheet 的对应关系', `<p>Sheet 名称发生变化或存在多个可匹配的表。请确认每个方案角色对应的当前 Sheet，每张表只能选择一次。</p>
                ${info.roles.map(role => `<div style="margin:12px 0;">
                    <label>${role.role_index === 0 ? '主表模板' : '数据表' + role.role_index}：${_escape(role.saved_file)} · ${_escape(role.saved_sheet)}</label>
                    <select class="int-sheet-relation" data-role="${role.role_index}" style="display:block;min-width:300px;max-width:100%;margin-top:5px;">
                    ${role.candidates.map(name => { const table = tables.get(name);
                        return `<option value="${_escape(name)}" ${name === role.selected ? 'selected' : ''}>${_escape(table.physical_file)} · ${_escape(table.sheet)}</option>`;
                    }).join('')}</select></div>`).join('')}
                <button class="btn btn-sm" type="button" id="int-sheet-ai-match">AI 预匹配</button>
                <span id="int-sheet-ai-status" style="font-size:12px;margin-left:8px;">已提供初步匹配，可调整后确认。</span>
                <div id="int-sheet-relation-error" style="color:#c00;margin-top:8px;"></div>`, () => {
                _modalCancelCallback = () => resolve(null);
                const selected = [...document.querySelectorAll('.int-sheet-relation')].map(select => select.value);
                if (new Set(selected).size !== selected.length) {
                    document.getElementById('int-sheet-relation-error').textContent = '同一张Sheet不能对应多个方案角色，请调整选择';
                    return;
                }
                resolve(selected);
                this.closeModal();
            }, {wide: true, confirmText: '确认对应关系', onCancel: () => resolve(null)});
            const selects = [...document.querySelectorAll('.int-sheet-relation')];
            let revision = 0;
            selects.forEach(select => select.addEventListener('change', () => {revision++;}));
            const button = document.getElementById('int-sheet-ai-match');
            const status = document.getElementById('int-sheet-ai-status');
            button.addEventListener('click', async () => {
                const submitted = revision;
                button.disabled = true;
                status.textContent = 'AI预匹配中，可继续手工选择…';
                try {
                    const response = await this._intFetch('/api/tools/integrate/suggest-sheet-mapping', {
                        method: 'POST', headers: {'Content-Type': 'application/json'},
                        body: JSON.stringify({session_id: _integrateAnalysis.session_id, scheme_id: schemeId}),
                    });
                    if (!response.ok) throw new Error('AI预匹配暂不可用，请手工选择');
                    const result = await response.json();
                    if (!button.isConnected || document.getElementById('modal-overlay').style.display === 'none') return;
                    if (revision !== submitted) {status.textContent = '已保留你手工调整的选择。'; return;}
                    if (result.ai_suggested) {
                        selects.forEach((select, i) => {select.value = result.role_files[i];});
                        status.textContent = 'AI预匹配完成，请确认对应关系。';
                    } else status.textContent = result.message || '请手工确认对应关系。';
                } catch (error) {if (button.isConnected) status.textContent = error.message;}
                finally {button.disabled = false;}
            });
        });
    },

    async _intValidateChosenSheets(result, schemeId) {
        if (!result.sheet_mapping?.required) return result;
        const roleFiles = await this._intChooseSheetRelations(result.sheet_mapping, schemeId);
        if (!roleFiles) return null;
        const response = await this._intFetch('/api/tools/integrate/scheme/apply-validate', {
            method: 'POST', headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({session_id: _integrateAnalysis.session_id, scheme_id: schemeId, role_files: roleFiles}),
        });
        if (!response.ok) {await _alertErr(response, 'Sheet关系校验失败'); return null;}
        const validated = await response.json();
        if (!validated.ok) {this._setIntegrateStatus((validated.reasons || ['Sheet关系校验失败']).join('；'), 'error'); return null;}
        if (validated.files) {
            _integrateAnalysis.files = validated.files;
            // Refresh saved ranges/columns without clearing the current wizard.
            validated.files.forEach(file => {
                const row = [...document.querySelectorAll('.int-header-range')].find(el => el.dataset.file === file.name);
                if (row) {
                    row.querySelector('.int-header-start').value = file.header_start;
                    row.querySelector('.int-header-end').value = file.header_end;
                }
            });
            this._intRenderKeyMap();
            this._intRenderDiffCols();
            this._intRenderPicker('ow');
            this._intRenderPicker('cmp');
        }
        return validated;
    },
});
