/**
 * tools.js - 智能小工具页面逻辑
 * 包含: Sheet拆分 / 模版管理 / 训练历史 / 计算历史 / 数据对比
 */

let _modalCallback = null;
let _splitFiles = [];
let _mergeFiles = [];
let _mergeAnalysis = null;   // analyze 返回的结果
let _mergeGroups = [];       // 当前结果列（直连 or AI 合并后）
let _mergeTemplates = [];     // 已保存的命名模版（含 config）
let _mergeAppliedTplId = null; // 当前套用的方案 id（带模版则生成时按模版填充）
let _integrateFiles = [];
let _integrateAnalysis = null;   // integrate/analyze 返回的结果
let _integrateSchemes = [];      // analyze 返回的匹配方案列表
const _ALLOWED_EXT = new Set(['xlsx', 'xls', 'xlsm']);

async function _alertErr(resp, fallback) {
    let msg = fallback;
    try { const j = await resp.json(); msg = j.detail || j.message || fallback; } catch (_) {
        try { msg = await resp.text(); } catch (__) {}
    }
    alert(msg);
}

function _escape(s) {
    return String(s).replace(/[&<>"']/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[c]));
}

const Tools = {
    _tplTenants: [],
    _trainingPage: 1,
    _computePage: 1,
    _pageSize: 20,

    // ==================== 初始化 ====================
    async init() {
        if (!AUTH.requireAuth()) return;
        AUTH.renderUserInfo(document.querySelector('header'));
        // nav-admin / tab-btn 的显隐由 data-perm + applyPermFilter 处理(在 renderUserInfo 中触发)

        this.initTabs();
        this.initSplitSheet();
        this.initDataMerge();
        this.initDataIntegrate();
    },

    initTabs() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                btn.classList.add('active');
                const tab = btn.dataset.tab;
                document.getElementById('tab-' + tab).classList.add('active');
                if (tab === 'templates') this.loadTemplateTenants().then(() => this.loadTemplates());
                else if (tab === 'training-history') this.loadTrainingHistory();                else if (tab === 'compute-history') this.loadComputeHistory();
                else if (tab === 'data-compare') this.loadCompareHistory();
            });
        });
    },

    // ==================== 弹窗工具 ====================
    openModal(title, bodyHtml, onConfirm) {
        document.getElementById('modal-title').textContent = title;
        document.getElementById('modal-body').innerHTML = bodyHtml;
        document.getElementById('modal-overlay').style.display = 'flex';
        _modalCallback = onConfirm;
    },

    closeModal() {
        document.getElementById('modal-overlay').style.display = 'none';
        _modalCallback = null;
    },

    confirmModal() {
        if (_modalCallback) {
            const cb = _modalCallback;
            _modalCallback = null;
            cb();
        }
    },

    // ==================== Sheet 拆分 ====================
    initSplitSheet() {
        const zone = document.getElementById('upload-zone');
        const input = document.getElementById('file-input');
        if (!zone || !input) return;

        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault();
            zone.classList.remove('dragover');
            this._addSplitFiles(e.dataTransfer.files);
        });
        input.addEventListener('change', () => this._addSplitFiles(input.files));
        document.getElementById('btn-split').addEventListener('click', () => this._doSplit());
    },

    _addSplitFiles(fileList) {
        for (const f of Array.from(fileList || [])) {
            const ext = (f.name.split('.').pop() || '').toLowerCase();
            if (!_ALLOWED_EXT.has(ext)) continue;
            if (_splitFiles.some(x => x.name === f.name && x.size === f.size)) continue;
            _splitFiles.push(f);
        }
        this._renderSplitList();
    },

    _renderSplitList() {
        const box = document.getElementById('split-file-list');
        if (_splitFiles.length === 0) {
            box.innerHTML = '';
            document.getElementById('btn-split').disabled = true;
            return;
        }
        box.innerHTML = _splitFiles.map((f, i) => `
            <div class="file-row">
                <span>📄 ${_escape(f.name)} <span style="color:#999;">(${(f.size / 1024).toFixed(1)} KB)</span></span>
                <span class="rm" data-i="${i}">×</span>
            </div>
        `).join('');
        box.querySelectorAll('.rm').forEach(el => el.addEventListener('click', (e) => {
            const i = parseInt(e.target.dataset.i, 10);
            _splitFiles.splice(i, 1);
            this._renderSplitList();
        }));
        document.getElementById('btn-split').disabled = false;
    },

    _setSplitStatus(text, kind) {
        const el = document.getElementById('split-status');
        el.textContent = text || '';
        el.className = 'status' + (kind ? ' ' + kind : '');
    },

    async _doSplit() {
        if (_splitFiles.length === 0) return;
        const btn = document.getElementById('btn-split');
        btn.disabled = true;
        this._setSplitStatus('拆分中，可能需要一会儿...');

        try {
            const fd = new FormData();
            _splitFiles.forEach(f => fd.append('files', f));

            const resp = await AUTH.authFetch('/api/tools/split-by-banner', {
                method: 'POST',
                body: fd,
            });

            if (!resp.ok) {
                let msg = `HTTP ${resp.status}`;
                try {
                    const j = await resp.json();
                    if (j && j.detail) msg = j.detail;
                } catch (_) {}
                throw new Error(msg);
            }

            const errCount = parseInt(resp.headers.get('X-Split-Errors') || '0', 10);
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url;
            a.download = 'split_results.zip';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(url);

            if (errCount > 0) {
                this._setSplitStatus(`完成，${errCount} 个文件失败，详情见 zip 内 _errors.txt`, 'error');
            } else {
                this._setSplitStatus('完成，已下载', 'ok');
            }

            _splitFiles.length = 0;
            const fileInput = document.getElementById('file-input');
            if (fileInput) fileInput.value = '';
            this._renderSplitList();
        } catch (e) {
            this._setSplitStatus(`失败: ${e.message}`, 'error');
        } finally {
            btn.disabled = (_splitFiles.length === 0);
        }
    },

    // ==================== 多表数据合并 ====================
    // ==================== 多表整合对比 ====================
    initDataIntegrate() {
        const zone = document.getElementById('integrate-upload-zone');
        const input = document.getElementById('integrate-file-input');
        if (!zone || !input) return;
        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault(); zone.classList.remove('dragover');
            this._addIntegrateFiles(e.dataTransfer.files);
        });
        input.addEventListener('change', () => this._addIntegrateFiles(input.files));
        document.getElementById('btn-integrate-analyze').addEventListener('click', () => this._analyzeIntegrate());
    },

    _addIntegrateFiles(fileList) {
        for (const f of Array.from(fileList || [])) {
            const ext = (f.name.split('.').pop() || '').toLowerCase();
            if (!_ALLOWED_EXT.has(ext)) continue;
            if (_integrateFiles.some(x => x.name === f.name && x.size === f.size)) continue;
            _integrateFiles.push(f);
        }
        this._renderIntegrateList();
    },

    _renderIntegrateList() {
        const box = document.getElementById('integrate-file-list');
        if (_integrateFiles.length === 0) {
            box.innerHTML = '';
            document.getElementById('btn-integrate-analyze').disabled = true;
            return;
        }
        box.innerHTML = _integrateFiles.map((f, i) => `
            <div class="file-row">
                <span>📄 ${_escape(f.name)} <span style="color:#999;">(${(f.size / 1024).toFixed(1)} KB)</span></span>
                <span class="rm" data-i="${i}">×</span>
            </div>`).join('');
        box.querySelectorAll('.rm').forEach(el => el.addEventListener('click', (e) => {
            _integrateFiles.splice(parseInt(e.target.dataset.i, 10), 1);
            this._renderIntegrateList();
        }));
        document.getElementById('btn-integrate-analyze').disabled = (_integrateFiles.length < 2);
    },

    _setIntegrateStatus(text, kind) {
        const el = document.getElementById('integrate-status');
        el.textContent = text || '';
        el.className = 'status' + (kind ? ' ' + kind : '');
    },

    async _analyzeIntegrate() {
        if (_integrateFiles.length < 2) { this._setIntegrateStatus('请至少上传 2 个文件', 'error'); return; }
        const btn = document.getElementById('btn-integrate-analyze');
        btn.disabled = true;
        this._setIntegrateStatus('解析中...');
        try {
            const fd = new FormData();
            _integrateFiles.forEach(f => fd.append('files', f));
            fd.append('tenant_id', '__tools_integrate__');
            const resp = await AUTH.authFetch('/api/tools/integrate/analyze', { method: 'POST', body: fd });
            if (!resp.ok) { await _alertErr(resp, '解析失败'); this._setIntegrateStatus('解析失败', 'error'); return; }
            _integrateAnalysis = await resp.json();
            this._renderIntegrateConfig(_integrateAnalysis);
            this._setIntegrateStatus('解析完成', 'ok');
        } catch (e) {
            this._setIntegrateStatus(`失败: ${e.message}`, 'error');
        } finally {
            btn.disabled = false;
        }
    },

    // 整合对比配置向导（阶段1-3：主表/关联键/覆盖对/对比对/输出方式/生成）
    _intFiles() { return (_integrateAnalysis && _integrateAnalysis.files) || []; },
    _intFileMeta(name) { return this._intFiles().find(f => f.name === name) || null; },
    _intMainFile() { const s = document.getElementById('int-main-file'); return s ? s.value : ''; },
    _intColsOf(name) { const f = this._intFileMeta(name); return (f && f.columns) || []; },
    _intNonMainFiles() { const m = this._intMainFile(); return this._intFiles().filter(f => f.name !== m); },
    _optList(cols, sel) {
        return (cols || []).map(c => `<option value="${_escape(c)}" ${c === sel ? 'selected' : ''}>${_escape(c)}</option>`).join('');
    },

    _renderIntegrateConfig(data) {
        const files = data.files || [];
        const box = document.getElementById('integrate-config');
        box.style.display = 'block';
        const fileOpts = files.map((f, i) => `<option value="${_escape(f.name)}" ${i === 0 ? 'selected' : ''}>${_escape(f.name)}</option>`).join('');
        box.innerHTML = `
            <div style="display:flex;align-items:center;gap:8px;flex-wrap:wrap;margin-bottom:12px;padding:8px;background:#eef4fb;border-radius:6px;">
                <span style="font-weight:600;">已保存方案：</span>
                <select id="int-scheme-select" style="min-width:220px;"><option value="">（选择匹配当前上传表的方案）</option></select>
                <button type="button" class="btn btn-sm" id="int-apply-scheme">套用</button>
                <button type="button" class="btn btn-sm" id="int-delete-scheme">删除</button>
                <span class="status" id="int-scheme-status" style="margin-left:8px;"></span>
            </div>
            <h3>① 选择主表（模板）</h3>
            <div class="form-group">
                <select id="int-main-file" style="min-width:280px;">${fileOpts}</select>
                <span style="color:#888;font-size:12px;margin-left:8px;">主表将被原地更新（保留其余 sheet 与公式），对照表的值按关联键回填到主表。</span>
            </div>

            <h3 style="margin-top:16px;">② 关联键（每个文件）</h3>
            <div id="int-key-map"></div>

            <h3 style="margin-top:16px;">③ 覆盖字段（勾选基准列 → 匹配对照列）</h3>
            <div id="int-ow-picker" style="border:1px solid #e3e7ed;border-radius:6px;padding:8px;max-height:150px;overflow:auto;"></div>
            <div style="margin:6px 0;">
                <button type="button" class="btn btn-sm" id="int-ow-match">智能匹配</button>
                <label style="margin-left:8px;font-size:13px;"><input type="checkbox" id="int-ow-ai" style="width:auto;"> 用 AI 匹配差异命名</label>
                <span class="status" id="int-ow-status" style="margin-left:8px;"></span>
            </div>
            <table class="data-table"><thead><tr><th>基准字段（主表列）</th><th>匹配的对照列（来源）</th><th style="width:32px;"></th></tr></thead>
                <tbody id="int-ow-list-rows"></tbody></table>
            <div style="color:#888;font-size:12px;margin-top:4px;">勾选后在下方选/改对照列。只写值不写公式；同一主表列多行时靠上优先。</div>

            <h3 style="margin-top:16px;">④ 对比字段（可选，输出方式2用）</h3>
            <div id="int-cmp-picker" style="border:1px solid #e3e7ed;border-radius:6px;padding:8px;max-height:150px;overflow:auto;"></div>
            <div style="margin:6px 0;">
                <button type="button" class="btn btn-sm" id="int-cmp-match">智能匹配</button>
                <label style="margin-left:8px;font-size:13px;"><input type="checkbox" id="int-cmp-ai" style="width:auto;"> 用 AI 匹配差异命名</label>
                <span class="status" id="int-cmp-status" style="margin-left:8px;"></span>
            </div>
            <table class="data-table"><thead><tr><th>基准字段（主表列）</th><th>匹配的对照列（来源）</th><th style="width:32px;"></th></tr></thead>
                <tbody id="int-cmp-list-rows"></tbody></table>

            <h3 style="margin-top:16px;">⑤ 输出方式</h3>
            <div class="form-group">
                <label style="margin-right:16px;"><input type="radio" name="int-output-mode" value="1" checked style="width:auto;"> 方式1：只更新主表</label>
                <label><input type="radio" name="int-output-mode" value="2" style="width:auto;"> 方式2：主表 + 差异 sheet</label>
            </div>
            <div id="int-diff-settings" style="display:none;padding:8px;background:#f4f7fb;border-radius:6px;">
                <span style="font-size:13px;">差异 sheet 定位（主表列）：</span>
                姓名 <select id="int-name-col" style="min-width:120px;"></select>
                身份证 <select id="int-id-col" style="min-width:120px;"></select>
                顺序 <select id="int-diff-order"><option value="id_name">身份证,姓名</option><option value="name_id">姓名,身份证</option></select>
            </div>

            <div class="actions" style="margin-top:16px;">
                <button class="btn btn-primary" id="int-execute">生成并下载</button>
                <span class="status" id="int-exec-status"></span>
            </div>
            <div class="actions" style="margin-top:10px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                <span>保存为方案：</span>
                <input type="text" id="int-scheme-name" placeholder="方案名称（如 5月工资整合）" style="min-width:200px;">
                <button type="button" class="btn" id="int-save-scheme">保存为方案</button>
                <span class="status" id="int-save-status"></span>
            </div>`;

        this._intRenderKeyMap();
        this._intRenderDiffCols();
        this._intRenderPicker('ow');
        this._intRenderPicker('cmp');
        document.getElementById('int-main-file').addEventListener('change', () => {
            this._intRenderKeyMap();
            this._intRenderDiffCols();
            this._intRenderPicker('ow');   // 换主表 → 重列基准字段（清空已选）
            this._intRenderPicker('cmp');
        });
        document.getElementById('int-ow-match').addEventListener('click', () =>
            this._intMatchSection('ow', document.getElementById('int-ow-ai').checked));
        document.getElementById('int-cmp-match').addEventListener('click', () =>
            this._intMatchSection('cmp', document.getElementById('int-cmp-ai').checked));
        document.querySelectorAll('input[name="int-output-mode"]').forEach(r =>
            r.addEventListener('change', () => {
                document.getElementById('int-diff-settings').style.display =
                    (document.querySelector('input[name="int-output-mode"]:checked').value === '2') ? 'block' : 'none';
            }));
        document.getElementById('int-execute').addEventListener('click', () => this._doIntegrateExecute());

        // 方案栏：填充匹配到的方案 + 套用/删除/保存
        this._intPopulateSchemes(data.matched_schemes || []);
        document.getElementById('int-apply-scheme').addEventListener('click', () => this._intApplyScheme());
        document.getElementById('int-delete-scheme').addEventListener('click', () => this._intDeleteScheme());
        document.getElementById('int-save-scheme').addEventListener('click', () => this._intSaveScheme());
    },

    _intPopulateSchemes(schemes) {
        _integrateSchemes = schemes || [];
        const sel = document.getElementById('int-scheme-select');
        if (!sel) return;
        sel.innerHTML = '<option value="">（选择匹配当前上传表的方案）</option>' +
            _integrateSchemes.map((s, i) => `<option value="${i}">${_escape(s.name)}</option>`).join('');
        const st = document.getElementById('int-scheme-status');
        if (st) st.textContent = _integrateSchemes.length ? `${_integrateSchemes.length} 个可用方案` : '无匹配方案';
    },

    _intApplyScheme() {
        const sel = document.getElementById('int-scheme-select');
        const idx = sel.value;
        if (idx === '') { alert('请先选择一个方案'); return; }
        const scheme = _integrateSchemes[parseInt(idx, 10)];
        if (!scheme) return;
        const cfg = scheme.config || {};
        const f2f = scheme.fp_to_file || {};

        // 主表
        const mainSel = document.getElementById('int-main-file');
        mainSel.value = scheme.main_file;
        this._intRenderKeyMap();
        this._intRenderDiffCols();

        // 关联键（按指纹→文件）
        Object.entries(cfg.key_map_by_fp || {}).forEach(([fp, key]) => {
            const file = f2f[fp];
            const el = document.querySelector(`.int-key[data-file="${CSS.escape(file || '')}"]`);
            if (el) el.value = key;
        });

        // 覆盖/对比（指纹→文件）：先重列基准字段勾选区，再按方案勾选并预填对照列（按 a_col 归并多源）
        this._intRenderPicker('ow');
        this._intRenderPicker('cmp');
        const trans = (pairs) => (pairs || []).map(p => ({
            a_col: p.a_col, source_file: f2f[p.source_fp], source_col: p.source_col,
        })).filter(p => p.source_file);
        const applyPairs = (pairs, kind) => {
            const byA = {};
            trans(pairs).forEach(p => { (byA[p.a_col] = byA[p.a_col] || []).push({ file: p.source_file, col: p.source_col }); });
            Object.entries(byA).forEach(([a, srcs]) => {
                const cb = document.querySelector(`.int-${kind}-pick[data-col="${CSS.escape(a)}"]`);
                if (cb) cb.checked = true;
                this._intAddListRow(kind, a, srcs);
            });
        };
        applyPairs(cfg.overwrite_pairs, 'ow');
        applyPairs(cfg.compare_pairs, 'cmp');

        // 输出方式 + 差异定位
        const mode = String(cfg.output_mode || 1);
        const radio = document.querySelector(`input[name="int-output-mode"][value="${mode}"]`);
        if (radio) { radio.checked = true; radio.dispatchEvent(new Event('change')); }
        if (cfg.name_col) { const e = document.getElementById('int-name-col'); if (e) e.value = cfg.name_col; }
        if (cfg.id_col) { const e = document.getElementById('int-id-col'); if (e) e.value = cfg.id_col; }
        if (cfg.diff_order) { const e = document.getElementById('int-diff-order'); if (e) e.value = cfg.diff_order; }

        document.getElementById('int-scheme-status').textContent = `已套用方案「${scheme.name}」`;
    },

    async _intDeleteScheme() {
        const sel = document.getElementById('int-scheme-select');
        const idx = sel.value;
        if (idx === '') { alert('请先选择要删除的方案'); return; }
        const scheme = _integrateSchemes[parseInt(idx, 10)];
        if (!scheme || !confirm(`删除方案「${scheme.name}」？`)) return;
        try {
            const resp = await AUTH.authFetch('/api/tools/integrate/scheme/' + scheme.id, { method: 'DELETE' });
            if (!resp.ok) { await _alertErr(resp, '删除失败'); return; }
            _integrateSchemes.splice(parseInt(idx, 10), 1);
            this._intPopulateSchemes(_integrateSchemes);
        } catch (e) { alert('删除失败: ' + e.message); }
    },

    async _intSaveScheme() {
        const name = (document.getElementById('int-scheme-name').value || '').trim();
        if (!name) { alert('请填写方案名称'); return; }
        const key_map = {};
        document.querySelectorAll('.int-key').forEach(s => { key_map[s.dataset.file] = s.value; });
        const overwrite_pairs = this._readSectionPairs('ow');
        const compare_pairs = this._readSectionPairs('cmp');
        const payload = {
            session_id: _integrateAnalysis.session_id, name, main_file: this._intMainFile(), key_map,
            overwrite_pairs, compare_pairs,
            name_col: document.getElementById('int-name-col')?.value || null,
            id_col: document.getElementById('int-id-col')?.value || null,
            diff_order: document.getElementById('int-diff-order')?.value || 'id_name',
            output_mode: parseInt(document.querySelector('input[name="int-output-mode"]:checked').value, 10),
            normalize_keys: true,
        };
        const st = document.getElementById('int-save-status');
        st.textContent = '保存中...'; st.className = 'status';
        try {
            const resp = await AUTH.authFetch('/api/tools/integrate/scheme/save', {
                method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload),
            });
            if (!resp.ok) { await _alertErr(resp, '保存失败'); st.textContent = '保存失败'; st.className = 'status error'; return; }
            st.textContent = '方案已保存'; st.className = 'status ok';
        } catch (e) { st.textContent = '失败: ' + e.message; st.className = 'status error'; }
    },

    _intRenderKeyMap() {
        const box = document.getElementById('int-key-map');
        const main = this._intMainFile();
        box.innerHTML = this._intFiles().map(f => {
            const sk = f.suggested_key || (f.columns || [])[0] || '';
            const tag = f.name === main ? '<b style="color:#1565c0;">[主表]</b> ' : '';
            return `<div style="margin:4px 0;">
                <span style="display:inline-block;min-width:240px;">${tag}${_escape(f.name)}：</span>
                <select class="int-key" data-file="${_escape(f.name)}">${this._optList(f.columns, sk)}</select>
            </div>`;
        }).join('');
    },

    _intRenderDiffCols() {
        const main = this._intMainFile();
        const meta = this._intFileMeta(main) || {};
        const cols = meta.columns || [];
        const nameSel = document.getElementById('int-name-col');
        const idSel = document.getElementById('int-id-col');
        if (nameSel) nameSel.innerHTML = this._optList(cols, meta.suggested_name_col || '');
        if (idSel) idSel.innerHTML = this._optList(cols, meta.suggested_id_col || '');
    },

    // 按多表合并习惯：所有基准列打勾 → 勾中的进下方列表 → 每行点击弹出勾选对照列（支持多表多选，靠前优先）
    // 渲染某套(kind=ow|cmp)的基准列勾选区；勾选联动下方列表增删
    _intRenderPicker(kind) {
        const box = document.getElementById(`int-${kind}-picker`);
        const listBody = document.getElementById(`int-${kind}-list-rows`);
        if (!box) return;
        if (listBody) listBody.innerHTML = '';
        const cols = this._intColsOf(this._intMainFile());
        if (!cols.length) { box.innerHTML = '<span style="color:#888;">主表无可用字段</span>'; return; }
        box.innerHTML = cols.map(c =>
            `<label style="display:inline-flex;align-items:center;gap:4px;margin:2px 12px 2px 0;font-size:13px;">
                <input type="checkbox" class="int-${kind}-pick" data-col="${_escape(c)}" style="width:auto;"> ${_escape(c)}</label>`).join('');
        box.querySelectorAll(`.int-${kind}-pick`).forEach(cb =>
            cb.addEventListener('change', (e) => {
                const col = e.target.dataset.col;
                if (e.target.checked) this._intAddListRow(kind, col);
                else {
                    const tr = document.querySelector(`#int-${kind}-list-rows tr[data-a-col="${CSS.escape(col)}"]`);
                    if (tr) tr.remove();
                }
            }));
    },

    // 列表行：对照列单元格点击弹出勾选（多表多选）；选中的对照列存于 tr.dataset.src(JSON)
    _intAddListRow(kind, col, presetSources) {
        const tbody = document.getElementById(`int-${kind}-list-rows`);
        if (!tbody || tbody.querySelector(`tr[data-a-col="${CSS.escape(col)}"]`)) return;
        const tr = document.createElement('tr');
        tr.dataset.aCol = col;
        tr.dataset.src = JSON.stringify(presetSources || []);
        tr.innerHTML = `
            <td>${_escape(col)}</td>
            <td class="int-src-cell" style="cursor:pointer;min-width:240px;"></td>
            <td style="text-align:center;"><span class="rm" style="cursor:pointer;color:#c00;">×</span></td>`;
        tr.querySelector('.int-src-cell').addEventListener('click', () => this._intOpenSrcPicker(kind, tr));
        tr.querySelector('.rm').addEventListener('click', () => {
            const cb = document.querySelector(`.int-${kind}-pick[data-col="${CSS.escape(col)}"]`);
            if (cb) cb.checked = false;
            tr.remove();
        });
        tbody.appendChild(tr);
        this._intUpdateSrcCell(tr);
    },

    _intUpdateSrcCell(tr) {
        const cell = tr.querySelector('.int-src-cell');
        let picks = [];
        try { picks = JSON.parse(tr.dataset.src || '[]'); } catch (_) {}
        cell.innerHTML = picks.length
            ? picks.map(p => `<span style="background:#eef4fb;padding:1px 6px;border-radius:3px;margin:1px;display:inline-block;font-size:12px;">${_escape(p.file)} · ${_escape(p.col)}</span>`).join(' ')
            : '<span style="color:#999;">点击选择对照列…</span>';
    },

    // 弹出勾选对照列：按对照文件分组（支持多表），可多选（靠前优先）
    _intOpenSrcPicker(kind, tr) {
        let selected = [];
        try { selected = JSON.parse(tr.dataset.src || '[]'); } catch (_) {}
        const selSet = new Set(selected.map(s => `${s.file}||${s.col}`));
        const files = this._intNonMainFiles();
        const body = `
            <div style="font-size:12px;color:#888;margin-bottom:8px;">可跨多张对照表多选；勾多个时靠上的优先（取首个非空值）。</div>
            ${files.map(f => `
                <div style="margin-bottom:10px;">
                    <div style="font-weight:600;color:#2c3e50;margin-bottom:4px;">📄 ${_escape(f.name)}</div>
                    <div style="display:flex;flex-wrap:wrap;gap:4px 14px;">
                        ${(f.columns || []).map(c => {
                            const checked = selSet.has(`${f.name}||${c}`) ? 'checked' : '';
                            return `<label style="display:inline-flex;align-items:center;gap:4px;font-size:13px;">
                                <input type="checkbox" class="int-srcpick-cb" data-file="${_escape(f.name)}" data-col="${_escape(c)}" ${checked} style="width:auto;"> ${_escape(c)}</label>`;
                        }).join('')}
                    </div>
                </div>`).join('')}`;
        this.openModal(`为「${tr.dataset.aCol || ''}」选择对照列（可多选，支持多表）`, body, () => {
            const picks = [];
            document.querySelectorAll('#modal-body .int-srcpick-cb:checked').forEach(cb =>
                picks.push({ file: cb.dataset.file, col: cb.dataset.col }));
            tr.dataset.src = JSON.stringify(picks);
            this._intUpdateSrcCell(tr);
            this.closeModal();
        });
    },

    // 智能匹配：对已勾选的基准列匹配对照列——命中自动选（单个），没命中留空让人工选
    async _intMatchSection(kind, useAi) {
        const rows = [...document.querySelectorAll(`#int-${kind}-list-rows tr`)];
        const st = document.getElementById(`int-${kind}-status`);
        if (!rows.length) { st.textContent = '请先在上方勾选基准字段'; st.className = 'status error'; return; }
        const source_cols = [];
        this._intNonMainFiles().forEach(f => (f.columns || []).forEach(c => source_cols.push({ file: f.name, col: c })));
        st.textContent = '匹配中...'; st.className = 'status';
        try {
            const resp = await AUTH.authFetch('/api/tools/integrate/match', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ session_id: _integrateAnalysis.session_id, main_file: this._intMainFile(), source_cols, use_ai: !!useAi }),
            });
            if (!resp.ok) { await _alertErr(resp, '匹配失败'); st.textContent = '匹配失败'; st.className = 'status error'; return; }
            const data = await resp.json();
            const bestByA = {};
            (data.pairs || []).forEach(p => { if (!(p.a_col in bestByA)) bestByA[p.a_col] = p; });
            let hit = 0, kept = 0, miss = 0;
            rows.forEach(tr => {
                let cur = [];
                try { cur = JSON.parse(tr.dataset.src || '[]'); } catch (_) {}
                if (cur.length) { kept++; return; }   // 已手动选择 → 以手动为准，不被 AI 覆盖
                const m = bestByA[tr.dataset.aCol];    // 未选的行才用 AI/精确匹配填充
                if (m) { tr.dataset.src = JSON.stringify([{ file: m.source_file, col: m.source_col }]); hit++; }
                else { tr.dataset.src = '[]'; miss++; }   // 匹配不到 → 留空，人工点选
                this._intUpdateSrcCell(tr);
            });
            st.textContent = `AI匹配填充 ${hit} 个，保留已选 ${kept} 个${miss ? `，${miss} 个未匹配请手动选` : ''}`;
            st.className = 'status ok';
        } catch (e) { st.textContent = '失败: ' + e.message; st.className = 'status error'; }
    },

    _readSectionPairs(kind) {
        const out = [];
        document.querySelectorAll(`#int-${kind}-list-rows tr`).forEach(tr => {
            let picks = [];
            try { picks = JSON.parse(tr.dataset.src || '[]'); } catch (_) {}
            picks.forEach(p => out.push({ a_col: tr.dataset.aCol, source_file: p.file, source_col: p.col }));
        });
        return out;
    },

    async _doIntegrateExecute() {
        const main = this._intMainFile();
        const key_map = {};
        document.querySelectorAll('.int-key').forEach(s => { key_map[s.dataset.file] = s.value; });
        const { overwrite_pairs, compare_pairs } = { overwrite_pairs: this._readSectionPairs('ow'), compare_pairs: this._readSectionPairs('cmp') };
        const mode = parseInt(document.querySelector('input[name="int-output-mode"]:checked').value, 10);
        if (overwrite_pairs.length === 0) { alert('请至少在「覆盖字段」里选好一列'); return; }
        if (mode === 2 && compare_pairs.length === 0) { alert('输出方式2需在「对比字段」里选好一列'); return; }
        const payload = {
            session_id: _integrateAnalysis.session_id, main_file: main, key_map,
            overwrite_pairs, compare_pairs, output_mode: mode,
            name_col: document.getElementById('int-name-col')?.value || null,
            id_col: document.getElementById('int-id-col')?.value || null,
            diff_order: document.getElementById('int-diff-order')?.value || 'id_name',
            normalize_keys: true,
        };
        const st = document.getElementById('int-exec-status');
        const btn = document.getElementById('int-execute');
        btn.disabled = true; st.textContent = '生成中...'; st.className = 'status';
        try {
            const resp = await AUTH.authFetch('/api/tools/integrate/execute', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(payload),
            });
            if (!resp.ok) { await _alertErr(resp, '生成失败'); st.textContent = '生成失败'; st.className = 'status error'; return; }
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = '整合结果_' + main;
            document.body.appendChild(a); a.click(); a.remove();
            URL.revokeObjectURL(url);
            const info = `命中${resp.headers.get('X-Integrate-Matched') || 0}行 覆盖${resp.headers.get('X-Integrate-Cells') || 0}格 差异${resp.headers.get('X-Integrate-Diffs') || 0}行`;
            st.textContent = '已生成下载（' + info + '）'; st.className = 'status ok';
        } catch (e) { st.textContent = '失败: ' + e.message; st.className = 'status error'; }
        finally { btn.disabled = false; }
    },

    initDataMerge() {
        const zone = document.getElementById('merge-upload-zone');
        const input = document.getElementById('merge-file-input');
        if (!zone || !input) return;
        zone.addEventListener('click', () => input.click());
        zone.addEventListener('dragover', (e) => { e.preventDefault(); zone.classList.add('dragover'); });
        zone.addEventListener('dragleave', () => zone.classList.remove('dragover'));
        zone.addEventListener('drop', (e) => {
            e.preventDefault(); zone.classList.remove('dragover');
            this._addMergeFiles(e.dataTransfer.files);
        });
        input.addEventListener('change', () => this._addMergeFiles(input.files));
        document.getElementById('btn-merge-analyze').addEventListener('click', () => this._analyzeMerge());
        document.getElementById('btn-merge-execute').addEventListener('click', () => this._doMerge());
        document.getElementById('btn-merge-match').addEventListener('click', () => this._matchMerge());
        document.getElementById('merge-select-all').addEventListener('click', () => this._mergeToggleAll(true));
        document.getElementById('merge-select-none').addEventListener('click', () => this._mergeToggleAll(false));
        document.getElementById('btn-save-template').addEventListener('click', () => this._saveMergeTemplate());
        document.getElementById('btn-apply-template').addEventListener('click', () => this._applyMergeTemplate());
        document.getElementById('btn-delete-template').addEventListener('click', () => this._deleteMergeTemplate());
        document.getElementById('btn-merge-skeleton').addEventListener('click', () => this._downloadMergeSkeleton());
    },

    _addMergeFiles(fileList) {
        for (const f of Array.from(fileList || [])) {
            const ext = (f.name.split('.').pop() || '').toLowerCase();
            if (!_ALLOWED_EXT.has(ext)) continue;
            if (_mergeFiles.some(x => x.name === f.name && x.size === f.size)) continue;
            _mergeFiles.push(f);
        }
        this._renderMergeList();
    },

    _renderMergeList() {
        const box = document.getElementById('merge-file-list');
        if (_mergeFiles.length === 0) {
            box.innerHTML = '';
            document.getElementById('btn-merge-analyze').disabled = true;
            return;
        }
        box.innerHTML = _mergeFiles.map((f, i) => `
            <div class="file-row">
                <span>📄 ${_escape(f.name)} <span style="color:#999;">(${(f.size / 1024).toFixed(1)} KB)</span></span>
                <span class="rm" data-i="${i}">×</span>
            </div>`).join('');
        box.querySelectorAll('.rm').forEach(el => el.addEventListener('click', (e) => {
            _mergeFiles.splice(parseInt(e.target.dataset.i, 10), 1);
            this._renderMergeList();
        }));
        document.getElementById('btn-merge-analyze').disabled = (_mergeFiles.length < 2);
    },

    _setMergeStatus(text, kind) {
        const el = document.getElementById('merge-status');
        el.textContent = text || '';
        el.className = 'status' + (kind ? ' ' + kind : '');
    },

    async _analyzeMerge() {
        if (_mergeFiles.length < 2) { this._setMergeStatus('请至少上传 2 个文件', 'error'); return; }
        const btn = document.getElementById('btn-merge-analyze');
        btn.disabled = true;
        this._setMergeStatus('解析中...');
        try {
            const fd = new FormData();
            _mergeFiles.forEach(f => fd.append('files', f));
            fd.append('tenant_id', '__tools_merge__');
            const resp = await AUTH.authFetch('/api/tools/merge/analyze', { method: 'POST', body: fd });
            if (!resp.ok) { await _alertErr(resp, '解析失败'); this._setMergeStatus('解析失败', 'error'); return; }
            _mergeAnalysis = await resp.json();
            this._renderMergeConfig(_mergeAnalysis);
            const hit = (_mergeAnalysis.cache_hit_files || []).length;
            this._setMergeStatus(`解析完成${hit ? `（${hit} 个文件有历史匹配缓存）` : ''}`, 'ok');
        } catch (e) {
            this._setMergeStatus(`失败: ${e.message}`, 'error');
        } finally {
            btn.disabled = false;
        }
    },

    _renderMergeConfig(data) {
        const files = data.files || [];

        // 主键下拉（每文件）——默认选中 suggested_key（身份证/工号等唯一标识列）
        document.getElementById('merge-key-map').innerHTML = files.map(f => {
            const sk = f.suggested_key || (f.columns || [])[0] || '';
            return `
            <div style="margin:4px 0;">
                <span style="display:inline-block;min-width:160px;">${_escape(f.name)}：</span>
                <select class="mg-key" data-file="${_escape(f.name)}">
                    ${(f.columns || []).map(c => `<option value="${_escape(c)}" ${c === sk ? 'selected' : ''}>${_escape(c)}</option>`).join('')}
                </select>
            </div>`;
        }).join('');

        // 基准文件
        document.getElementById('merge-base-file').innerHTML =
            files.map(f => `<option value="${_escape(f.name)}">${_escape(f.name)}</option>`).join('');
        // 模版骨架基准表（以哪个上传表为模版底）
        const skb = document.getElementById('merge-skeleton-base');
        if (skb) skb.innerHTML = files.map(f => `<option value="${_escape(f.name)}">${_escape(f.name)}</option>`).join('');

        // 字段选择（按文件分块，标注来源 + 每文件独立全选/全不选）
        document.getElementById('merge-field-select').innerHTML = files.map((f, fi) => `
            <div style="margin-bottom:10px;">
                <div style="display:flex;align-items:center;gap:8px;margin-bottom:4px;">
                    <span style="font-weight:600;color:#2c3e50;">📄 ${_escape(f.name)}</span>
                    <button type="button" class="btn btn-sm mg-file-all" data-fi="${fi}">全选</button>
                    <button type="button" class="btn btn-sm mg-file-none" data-fi="${fi}">全不选</button>
                </div>
                <div style="display:flex;flex-wrap:wrap;gap:4px 16px;">
                    ${(f.columns || []).map(c => `
                        <label style="display:inline-flex;align-items:center;gap:4px;font-size:13px;">
                            <input type="checkbox" class="mg-field" data-file="${_escape(f.name)}" data-col="${_escape(c)}">
                            ${_escape(c)}
                        </label>`).join('')}
                </div>
            </div>`).join('');

        // 每文件全选/全不选（用 data-fi 索引定位该文件名，避免文件名含特殊字符）
        document.querySelectorAll('#merge-field-select .mg-file-all').forEach(btn =>
            btn.addEventListener('click', () => this._mergeToggleFile(files[parseInt(btn.dataset.fi, 10)].name, true)));
        document.querySelectorAll('#merge-field-select .mg-file-none').forEach(btn =>
            btn.addEventListener('click', () => this._mergeToggleFile(files[parseInt(btn.dataset.fi, 10)].name, false)));

        // 默认结果列 = 每个勾选字段直连一列；字段勾选变化 → 增量同步（保留已有顺序）
        document.getElementById('merge-field-select').addEventListener('change', () => this._syncFieldsToGroups());
        this._rebuildDirectGroups();
        document.getElementById('merge-config').style.display = 'block';
        _mergeAppliedTplId = null;   // 新一轮解析，清除上次套用的方案
        this._setFillHint();
        this._loadMergeTemplates();
    },

    async _loadMergeTemplates() {
        try {
            const resp = await AUTH.authFetch('/api/tools/merge/templates');
            if (!resp.ok) return;
            _mergeTemplates = await resp.json();
            const sel = document.getElementById('merge-template-select');
            sel.innerHTML = '<option value="">（选择已保存的模版）</option>' +
                _mergeTemplates.map((t, i) => `<option value="${i}">${_escape(t.name)}${t.has_template ? ' 🗎' : ''}</option>`).join('');
            this._setFillHint();
        } catch (_) {}
    },

    _setTplStatus(text, kind) {
        const el = document.getElementById('merge-template-status');
        if (el) { el.textContent = text || ''; el.className = 'status' + (kind ? ' ' + kind : ''); }
    },

    // 当前结果列（含可能改过名的输出列名）——骨架/保存/执行共用
    _currentGroups() {
        return (_mergeGroups || []).map((g, gi) => {
            const inp = document.querySelector(`.mg-name[data-gi="${gi}"]`);
            const nm = (inp && inp.value.trim()) || g.name;
            return { name: nm, sources: g.sources };
        });
    },

    async _downloadMergeSkeleton() {
        const groups = this._currentGroups();
        if (!groups.length) { this._setMergeExecStatus('请先选择结果列', 'error'); return; }
        const btn = document.getElementById('btn-merge-skeleton');
        const old = btn.textContent; btn.disabled = true; btn.textContent = '生成中...';
        try {
            const baseSel = document.getElementById('merge-skeleton-base');
            const resp = await AUTH.authFetch('/api/tools/merge/template-skeleton', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    result_columns: groups.map(g => ({ name: g.name })),
                    session_id: (_mergeAnalysis && _mergeAnalysis.session_id) || null,
                    base_file: (baseSel && baseSel.value) || null,
                }),
            });
            if (!resp.ok) { await _alertErr(resp, '生成骨架失败'); return; }
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = 'merge_template_skeleton.xlsx';
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            URL.revokeObjectURL(url);
            this._setMergeExecStatus('已下载模版骨架，编辑后可作为“模版文件”随方案上传', 'ok');
        } catch (e) { this._setMergeExecStatus('生成骨架失败: ' + e.message, 'error'); }
        finally { btn.disabled = false; btn.textContent = old; }
    },

    _setFillHint() {
        const el = document.getElementById('merge-tpl-fill-hint');
        if (!el) return;
        const tpl = (_mergeTemplates || []).find(t => t.id === _mergeAppliedTplId);
        el.textContent = (tpl && tpl.has_template) ? `将按方案「${tpl.name}」的模版填充（带格式/公式）` : '';
    },

    async _saveMergeTemplate() {
        const name = (document.getElementById('merge-template-name').value || '').trim();
        if (!name) { this._setSaveStatus('请输入方案名称', 'error'); return; }
        // 收集当前配置：主键列名（取首个文件的主键下拉值）、结果列(按列名)、模式
        const keyEls = document.querySelectorAll('.mg-key');
        const key_field = keyEls.length ? keyEls[0].value : '';
        const groups = this._currentGroups().map(g => ({
            name: g.name,
            source_cols: [...new Set(g.sources.map(s => s.col))],
        }));
        const config = {
            key_field,
            merge_mode: document.getElementById('merge-mode').value,
            normalize_keys: document.getElementById('merge-normalize-keys').checked,
            result_columns: groups,
        };
        const fileInput = document.getElementById('merge-template-file');
        const tplFile = fileInput && fileInput.files && fileInput.files[0];
        const fd = new FormData();
        fd.append('name', name);
        fd.append('config', JSON.stringify(config));
        if (tplFile) fd.append('template', tplFile);

        const btn = document.getElementById('btn-save-template');
        const oldText = btn.textContent;
        btn.disabled = true;
        btn.textContent = '保存中...';
        this._setSaveStatus('保存中...');
        try {
            const resp = await AUTH.authFetch('/api/tools/merge/template/save', {
                method: 'POST', body: fd,   // multipart，勿手动设 Content-Type
            });
            if (!resp.ok) { await _alertErr(resp, '保存失败'); this._setSaveStatus('保存失败', 'error'); return; }
            const j = await resp.json();
            this._setSaveStatus(`✓ 已保存方案「${name}」` + (j.has_template ? '（含模版）' : ''), 'ok');
            if (fileInput) fileInput.value = '';
            await this._loadMergeTemplates();
            // 保存后自动视为已套用该方案
            if (j.id != null) { _mergeAppliedTplId = j.id; this._setFillHint(); }
        } catch (e) { this._setSaveStatus('保存失败: ' + e.message, 'error'); }
        finally { btn.disabled = false; btn.textContent = oldText; }
    },

    _setSaveStatus(text, kind) {
        const el = document.getElementById('merge-save-status');
        if (el) { el.textContent = text || ''; el.className = 'status' + (kind ? ' ' + kind : ''); }
    },

    _applyMergeTemplate() {
        const idx = document.getElementById('merge-template-select').value;
        if (idx === '') { this._setTplStatus('请先选择模版', 'error'); return; }
        const tpl = (_mergeTemplates || [])[parseInt(idx, 10)];
        if (!tpl || !tpl.config) return;
        const cfg = tpl.config;
        _mergeAppliedTplId = tpl.id;   // 记录套用的方案；带模版则生成时按模版填充
        const files = (_mergeAnalysis && _mergeAnalysis.files) || [];

        // 1) 主键：各文件里选中与 key_field 同名的列
        if (cfg.key_field) {
            document.querySelectorAll('.mg-key').forEach(sel => {
                const opt = Array.from(sel.options).find(o => o.value === cfg.key_field);
                if (opt) sel.value = cfg.key_field;
            });
        }
        // 2) 模式 / 归一化
        if (cfg.merge_mode) document.getElementById('merge-mode').value = cfg.merge_mode;
        if (typeof cfg.normalize_keys === 'boolean') document.getElementById('merge-normalize-keys').checked = cfg.normalize_keys;

        // 3) 结果列：按 source_cols 列名在已上传文件里匹配，重建 _mergeGroups
        const missing = [];
        const groups = [];
        (cfg.result_columns || []).forEach(rc => {
            const sources = [];
            (rc.source_cols || []).forEach(colName => {
                files.forEach(f => {
                    if ((f.columns || []).includes(colName)) sources.push({ file: f.name, col: colName });
                });
            });
            if (sources.length) groups.push({ name: rc.name, sources });
            else missing.push(rc.name);
        });
        if (groups.length) {
            _mergeGroups = groups;
            this._renderGroupsPreview();
            // 同步字段勾选：只勾选被结果列用到的列
            const used = new Set(groups.flatMap(g => g.sources.map(s => s.file + '|||' + s.col)));
            document.querySelectorAll('#merge-field-select .mg-field').forEach(cb => {
                cb.checked = used.has(cb.dataset.file + '|||' + cb.dataset.col);
            });
        }
        this._setTplStatus(
            `已套用「${tpl.name}」` + (missing.length ? `（这些列在当前文件中没找到，已跳过：${missing.join('、')}）` : ''),
            missing.length ? 'error' : 'ok');
        this._setFillHint();
    },

    async _deleteMergeTemplate() {
        const idx = document.getElementById('merge-template-select').value;
        if (idx === '') { this._setTplStatus('请先选择要删除的模版', 'error'); return; }
        const tpl = (_mergeTemplates || [])[parseInt(idx, 10)];
        if (!tpl) return;
        if (!confirm(`确认删除模版「${tpl.name}」？`)) return;
        try {
            const resp = await AUTH.authFetch(`/api/tools/merge/template/${tpl.id}`, { method: 'DELETE' });
            if (!resp.ok) { this._setTplStatus('删除失败', 'error'); return; }
            this._setTplStatus(`已删除「${tpl.name}」`, 'ok');
            this._loadMergeTemplates();
        } catch (e) { this._setTplStatus('删除失败: ' + e.message, 'error'); }
    },

    _mergeToggleAll(checked) {
        document.querySelectorAll('#merge-field-select .mg-field').forEach(cb => { cb.checked = checked; });
        this._rebuildDirectGroups();
    },

    _mergeToggleFile(fileName, checked) {
        document.querySelectorAll('#merge-field-select .mg-field').forEach(cb => {
            if (cb.dataset.file === fileName) cb.checked = checked;
        });
        this._rebuildDirectGroups();
    },

    _selectedFields() {
        return Array.from(document.querySelectorAll('#merge-field-select .mg-field:checked'))
            .map(cb => ({ file: cb.dataset.file, col: cb.dataset.col }));
    },

    _rebuildDirectGroups() {
        // 直连：每个勾选字段各成一列；同名列加来源后缀区分
        const sel = this._selectedFields();
        const nameCount = {};
        sel.forEach(s => { nameCount[s.col] = (nameCount[s.col] || 0) + 1; });
        _mergeGroups = sel.map(s => ({
            name: nameCount[s.col] > 1 ? `${s.col}（${s.file}）` : s.col,
            sources: [{ file: s.file, col: s.col }],
        }));
        this._renderGroupsPreview();
        this._setMergeMatchStatus('');
    },

    // 增量同步：勾选变化时保留已有结果列顺序（含拖动顺序、AI 多源组），新勾选追加到末尾，取消的移除
    _syncFieldsToGroups() {
        this._syncGroupNames();   // 先保住已改过的列名
        const selSet = new Set(this._selectedFields().map(s => s.file + '|||' + s.col));
        // 1) 现有结果列：剔除已取消勾选的来源；来源全没了的组删除
        let groups = (_mergeGroups || []).map(g => ({
            name: g.name,
            sources: (g.sources || []).filter(s => selSet.has(s.file + '|||' + s.col)),
        })).filter(g => g.sources.length > 0);
        // 2) 已被现有结果列覆盖的字段
        const covered = new Set(groups.flatMap(g => g.sources.map(s => s.file + '|||' + s.col)));
        const usedNames = new Set(groups.map(g => g.name));
        // 3) 新勾选但未覆盖的字段 → 按 DOM 顺序追加到末尾
        this._selectedFields().forEach(s => {
            const key = s.file + '|||' + s.col;
            if (covered.has(key)) return;
            covered.add(key);
            let nm = s.col;
            if (usedNames.has(nm)) nm = `${s.col}（${s.file}）`;
            usedNames.add(nm);
            groups.push({ name: nm, sources: [{ file: s.file, col: s.col }] });
        });
        _mergeGroups = groups;
        this._renderGroupsPreview();
    },

    _renderGroupsPreview() {
        const tbody = document.querySelector('#merge-groups-table tbody');
        tbody.innerHTML = (_mergeGroups || []).map((g, gi) => {
            const multi = g.sources.length > 1;
            const srcText = g.sources.map(s => `${_escape(s.file)} · ${_escape(s.col)}`).join('<br>');
            return `<tr class="mg-row" draggable="true" data-gi="${gi}">
                <td class="mg-drag" style="cursor:grab;text-align:center;color:#bbb;user-select:none;font-weight:bold;" title="拖动调整列顺序">⋮⋮</td>
                <td><input class="mg-name" data-gi="${gi}" value="${_escape(g.name)}" style="width:180px;"></td>
                <td>${srcText}${multi ? ' <span style="color:#c0392b;font-size:12px;">[多源·值不一致将标红]</span>' : ''}</td>
            </tr>`;
        }).join('');
        this._bindGroupDrag();
    },

    // 把当前预览表里（可能改过的）列名回写到 _mergeGroups，避免重渲染丢失
    _syncGroupNames() {
        document.querySelectorAll('#merge-groups-table .mg-name').forEach(inp => {
            const gi = parseInt(inp.dataset.gi, 10);
            if (_mergeGroups[gi]) _mergeGroups[gi].name = (inp.value.trim() || _mergeGroups[gi].name);
        });
    },

    // 结果列拖动排序：拖动顺序即最终输出列顺序
    _bindGroupDrag() {
        const tbody = document.querySelector('#merge-groups-table tbody');
        if (!tbody) return;
        let dragGi = null;
        tbody.querySelectorAll('tr.mg-row').forEach(tr => {
            tr.addEventListener('dragstart', (e) => {
                dragGi = parseInt(tr.dataset.gi, 10);
                tr.style.opacity = '0.4';
                e.dataTransfer.effectAllowed = 'move';
            });
            tr.addEventListener('dragend', () => { tr.style.opacity = ''; });
            tr.addEventListener('dragover', (e) => {
                e.preventDefault();
                e.dataTransfer.dropEffect = 'move';
                tr.style.borderTop = '2px solid #2c7be5';
            });
            tr.addEventListener('dragleave', () => { tr.style.borderTop = ''; });
            tr.addEventListener('drop', (e) => {
                e.preventDefault();
                tr.style.borderTop = '';
                const targetGi = parseInt(tr.dataset.gi, 10);
                if (dragGi === null || isNaN(targetGi) || dragGi === targetGi) { dragGi = null; return; }
                this._syncGroupNames();
                const arr = _mergeGroups.slice();
                const [moved] = arr.splice(dragGi, 1);
                const insertAt = dragGi < targetGi ? targetGi - 1 : targetGi;
                arr.splice(insertAt, 0, moved);
                _mergeGroups = arr;
                dragGi = null;
                this._renderGroupsPreview();
            });
        });
    },

    async _matchMerge() {
        const sel = this._selectedFields();
        if (!sel.length) { this._setMergeMatchStatus('请先勾选字段', 'error'); return; }
        const useAi = document.getElementById('merge-use-ai').checked;
        const btn = document.getElementById('btn-merge-match');
        btn.disabled = true;
        this._setMergeMatchStatus(useAi ? 'AI 匹配中...' : '合并同名列中...');
        try {
            const resp = await AUTH.authFetch('/api/tools/merge/match', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    session_id: _mergeAnalysis.session_id,
                    tenant_id: '__tools_merge__',
                    selected: sel,
                    use_ai: useAi,
                    ai_provider: 'deepseek',
                }),
            });
            if (!resp.ok) { await _alertErr(resp, '匹配失败'); this._setMergeMatchStatus('匹配失败', 'error'); return; }
            const data = await resp.json();
            _mergeGroups = (data.groups || []).map(g => ({ name: g.name, sources: g.sources }));
            this._renderGroupsPreview();
            const merged = _mergeGroups.filter(g => g.sources.length > 1).length;
            this._setMergeMatchStatus(`完成：合并出 ${merged} 个多源列${useAi ? '（含 AI）' : '（仅同名）'}，可在下方改名`, 'ok');
        } catch (e) {
            this._setMergeMatchStatus(`失败: ${e.message}`, 'error');
        } finally {
            btn.disabled = false;
        }
    },

    _setMergeMatchStatus(text, kind) {
        const el = document.getElementById('merge-match-status');
        if (!el) return;
        el.textContent = text || '';
        el.className = 'status' + (kind ? ' ' + kind : '');
    },

    async _doMerge() {
        if (!_mergeAnalysis) return;
        // 用预览表里的（可能改过名的）结果列
        const groups = this._currentGroups();
        if (!groups.length) { this._setMergeExecStatus('请至少勾选一个字段', 'error'); return; }
        const key_map = {};
        document.querySelectorAll('.mg-key').forEach(sel => { key_map[sel.dataset.file] = sel.value; });

        const usingTpl = !!((_mergeTemplates || []).find(t => t.id === _mergeAppliedTplId && t.has_template));
        const btn = document.getElementById('btn-merge-execute');
        btn.disabled = true;
        this._setMergeExecStatus(usingTpl ? '按模版填充生成中...' : '生成中...');
        try {
            const body = {
                session_id: _mergeAnalysis.session_id,
                tenant_id: '__tools_merge__',
                key_map,
                result_columns: groups,
                merge_mode: document.getElementById('merge-mode').value,
                base_file: document.getElementById('merge-base-file').value,
                normalize_keys: document.getElementById('merge-normalize-keys').checked,
                template_id: _mergeAppliedTplId,
            };
            const resp = await AUTH.authFetch('/api/tools/merge/execute', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify(body),
            });
            if (!resp.ok) { await _alertErr(resp, '合并失败'); this._setMergeExecStatus('合并失败', 'error'); return; }
            const conflicts = resp.headers.get('X-Merge-Conflicts') || '0';
            const rows = resp.headers.get('X-Merge-Rows') || '0';
            const blob = await resp.blob();
            const url = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = url; a.download = 'merged_result.xlsx';
            document.body.appendChild(a); a.click(); document.body.removeChild(a);
            URL.revokeObjectURL(url);
            this._setMergeExecStatus(
                usingTpl ? `完成：已按模版填充 ${rows} 行并下载` :
                `完成：${rows} 行，${conflicts} 个冲突主键已标红，已下载`, 'ok');
        } catch (e) {
            this._setMergeExecStatus(`失败: ${e.message}`, 'error');
        } finally {
            btn.disabled = false;
        }
    },

    _setMergeExecStatus(text, kind) {
        const el = document.getElementById('merge-exec-status');
        el.textContent = text || '';
        el.className = 'status' + (kind ? ' ' + kind : '');
    },

    // ==================== 训练历史 ====================
    _renderPagination(containerId, currentPage, totalCount, pageSize, callbackName) {
        const container = document.getElementById(containerId);
        if (!container) return;
        if (!totalCount) { container.innerHTML = ''; return; }
        const totalPages = Math.max(1, Math.ceil(totalCount / pageSize));
        let html = `<span class="pg-info">共 ${totalCount} 条 · 第 ${currentPage}/${totalPages} 页</span>`;
        html += `<button class="pg-btn" ${currentPage <= 1 ? 'disabled' : ''} onclick="Tools.${callbackName}(${currentPage - 1})">上一页</button>`;
        const set = new Set([1, totalPages, currentPage, currentPage - 1, currentPage + 1, currentPage - 2, currentPage + 2]);
        const pages = [...set].filter(p => p >= 1 && p <= totalPages).sort((a, b) => a - b);
        let prev = 0;
        for (const p of pages) {
            if (p - prev > 1) html += '<span class="pg-ellipsis">…</span>';
            html += `<button class="pg-btn ${p === currentPage ? 'active' : ''}" onclick="Tools.${callbackName}(${p})">${p}</button>`;
            prev = p;
        }
        html += `<button class="pg-btn" ${currentPage >= totalPages ? 'disabled' : ''} onclick="Tools.${callbackName}(${currentPage + 1})">下一页</button>`;
        container.innerHTML = html;
    },

    async loadTrainingHistory(page = 1) {
        const tenantId = document.getElementById('training-tenant-filter')?.value || '';
        this._trainingPage = page;
        const offset = (page - 1) * this._pageSize;
        let url = `/api/training/sessions?limit=${this._pageSize}&offset=${offset}`;
        if (tenantId) url += `&tenant_id=${encodeURIComponent(tenantId)}`;
        const resp = await AUTH.authFetch(url);
        if (!resp.ok) return;
        const result = await resp.json();
        const tbody = document.querySelector('#training-history-table tbody');
        if (!result.items || !result.items.length) {
            tbody.innerHTML = '<tr><td colspan="9" class="empty-state">暂无训练记录</td></tr>';
        } else {
            tbody.innerHTML = result.items.map(s => `<tr>
                <td>${s.id}</td>
                <td>${s.tenant_id}</td>
                <td>${s.mode || '-'}</td>
                <td><span class="status-${s.status}">${s.status}</span></td>
                <td>${s.total_iterations || 0}</td>
                <td>${s.best_accuracy != null ? (s.best_accuracy * 100).toFixed(1) + '%' : '-'}</td>
                <td>${s.started_at ? new Date(s.started_at).toLocaleString() : '-'}</td>
                <td>${s.finished_at ? new Date(s.finished_at).toLocaleString() : '-'}</td>
                <td>
                    <button class="btn btn-sm" onclick="Tools.showTrainingDetail(${s.id})">详情</button>
                </td>
            </tr>`).join('');
        }
        this._renderPagination('training-history-pagination', page, result.total || 0, this._pageSize, 'loadTrainingHistory');
    },

    async showTrainingDetail(sessionId) {
        const resp = await AUTH.authFetch(`/api/training/sessions/${sessionId}/iterations`);
        if (!resp.ok) return alert('获取详情失败');
        const iterations = await resp.json();
        let html = `<div style="max-height:500px;overflow-y:auto;">`;
        if (!iterations.length) html += '<p>暂无迭代记录</p>';
        iterations.forEach(it => {
            html += `<div style="border:1px solid #eee;border-radius:8px;padding:12px;margin-bottom:10px;">
                <div style="display:flex;justify-content:space-between;margin-bottom:8px;">
                    <strong>第 ${it.iteration_num} 轮</strong>
                    <span>准确率: ${it.accuracy != null ? (it.accuracy * 100).toFixed(1) + '%' : '-'}</span>
                    <span class="status-${it.status}">${it.status}</span>
                </div>
                ${it.generated_code ? `<details><summary>查看代码 (${it.generated_code.length} 字符)</summary><pre style="font-size:11px;max-height:200px;overflow:auto;background:#f5f5f5;padding:8px;border-radius:4px;">${it.generated_code.substring(0, 3000)}</pre></details>` : ''}
                ${it.error_details ? `<div style="color:red;font-size:12px;">错误: ${JSON.stringify(it.error_details)}</div>` : ''}
            </div>`;
        });
        html += '</div>';
        this.openModal(`训练会话 #${sessionId} - 迭代详情`, html, null);
    },

    // ==================== 模版管理 ====================
    async loadTemplateTenants() {
        const resp = await AUTH.authFetch('/api/admin/tenant-auth/tenants');
        if (!resp.ok) return;
        this._tplTenants = await resp.json();
        const sel = document.getElementById('tpl-tenant-filter');
        if (sel) {
            sel.innerHTML = '<option value="">全部</option><option value="__global__">仅全局</option>' +
                this._tplTenants.map(t => `<option value="${t}">租户: ${t}</option>`).join('');
        }
    },

    async loadTemplates() {
        const tenantId = document.getElementById('tpl-tenant-filter')?.value || '';
        let url = '/api/admin/templates';
        if (tenantId) url += `?tenant_id=${encodeURIComponent(tenantId)}`;
        const resp = await AUTH.authFetch(url);
        if (!resp.ok) return;
        const list = await resp.json();
        this.renderTemplates(list);
    },

    renderTemplates(list) {
        const tbody = document.querySelector('#templates-table tbody');
        if (!list.length) {
            tbody.innerHTML = '<tr><td colspan="9" class="empty-state">暂无模版</td></tr>';
            return;
        }
        tbody.innerHTML = list.map(t => `<tr>
            <td>${t.id}</td>
            <td>${t.name}</td>
            <td>${t.tenant_id ? '<span class="tag">租户: ' + t.tenant_id + '</span>' : '<span class="tag" style="background:#e8f5e9;color:#2e7d32">全局</span>'}</td>
            <td>${t.file_name}</td>
            <td>${t.file_name_rule || '-'}</td>
            <td>${t.encrypt_password || '<span style="color:#999">不加密</span>'}</td>
            <td>${t.report_mode === 'block' ? '<span class="tag" style="background:#fff3e0;color:#e65100">block</span>' : t.report_mode === 'zip' ? '<span class="tag" style="background:#e8eaf6;color:#283593">zip</span>' : t.report_mode === 'sheet' ? '<span class="tag" style="background:#e8f5e9;color:#2e7d32">sheet</span>' : 'fill'}${t.group_by ? ' <small>(' + t.group_by + ')</small>' : ''}${t.split_by ? ' <small style="color:#1565c0;">[拆分:' + t.split_by + ']</small>' : ''}</td>
            <td>${t.is_active ? '<span style="color:green">启用</span>' : '<span style="color:#999">停用</span>'}</td>
            <td class="actions">
                <button class="btn btn-sm" onclick="Tools.downloadTemplate(${t.id}, '${t.file_name.replace(/'/g, "\\'")}')">下载</button>
                <button class="btn btn-sm" onclick="Tools.showEditTemplate(${t.id})">编辑</button>
                <button class="btn btn-sm btn-danger" onclick="Tools.deleteTemplate(${t.id}, '${t.name.replace(/'/g, "\\'")}')">停用</button>
            </td>
        </tr>`).join('');
    },

    showCreateTemplate() {
        const tenantOptions = this._tplTenants.map(t =>
            `<option value="${t}">租户: ${t}</option>`
        ).join('');
        this.openModal('新建模版', `
            <div style="display:flex;flex-direction:column;gap:12px;">
                <div class="form-group"><label>租户</label>
                    <select id="m-tpl-tenant" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                        <option value="">全局（所有租户可用）</option>
                        ${tenantOptions}
                    </select>
                </div>
                <div class="form-group"><label>模版名称</label>
                    <input id="m-tpl-name" required style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                </div>
                <div class="form-group"><label>描述</label>
                    <input id="m-tpl-desc" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                </div>
                <div class="form-group"><label>模版文件</label>
                    <input id="m-tpl-file" type="file" accept=".xlsx,.xls,.xlsm">
                </div>
                <div class="form-group"><label>文件名规则</label>
                    <input id="m-tpl-name-rule" placeholder="如: {year}{month}_薪资表_{姓名}" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">可用变量: {year} {month} {date} {tenant} {列名} {列名[:N]} {列名[-N:]}</small>
                </div>
                <div class="form-group"><label>加密规则</label>
                    <input id="m-tpl-encrypt-rule" placeholder="如: {身份证号码[:6]} 或 {姓名[:1]}{身份证号码[-6:]}" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">留空表示不加密。可用变量: {列名} {列名[:N]}前N位 {列名[-N:]}后N位</small>
                </div>
                <div class="form-group"><label>报表模式</label>
                    <select id="m-tpl-report-mode" onchange="Tools._toggleModeFields()" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                        <option value="fill">fill — 整表填充</option>
                        <option value="block">block — 分组合并（每组一块）</option>
                        <option value="zip">zip — 分组打包（每组一文件）</option>
                        <option value="sheet">sheet — 分组多Sheet（每组一个Sheet）</option>
                    </select>
                </div>
                <div class="form-group"><label>文件拆分字段</label>
                    <input id="m-tpl-split-by" placeholder="如：部门（留空则不拆分文件）" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">按此列值将数据拆分到不同文件中，拆分后自动打包为 zip</small>
                </div>
                <div class="form-group"><label>整表带出结果Sheet</label>
                    <input id="m-tpl-carry-sheets" placeholder="如：明细,#2（逗号分隔，可填sheet名或#序号）" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">把计算结果中这些 sheet 整表(值+格式)追加到报表末尾；#N 表示结果表第 N 个 sheet。留空不带出</small>
                </div>
                <div id="m-tpl-mode-fields" style="display:none;">
                    <div class="form-group"><label>分组字段</label>
                        <input id="m-tpl-group-by" placeholder="如: 工号" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                        <small style="color:#888;">按此列的值分组，每组独立填充模板</small>
                    </div>
                    <div class="form-group" id="m-tpl-skip-rows-group"><label>块间空行数</label>
                        <input id="m-tpl-skip-rows" type="number" value="1" min="0" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    </div>
                    <div class="form-group" id="m-tpl-name-field-group" style="display:none;"><label>文件命名字段</label>
                        <input id="m-tpl-name-field" placeholder="如: 姓名（用于 zip 内文件名）" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    </div>
                    <div class="form-group" id="m-tpl-show-empty-group"><label style="display:flex;align-items:center;gap:6px;">
                        <input id="m-tpl-show-empty" type="checkbox" checked> 显示空月份（多月合并时补齐无数据的月份）
                    </label></div>
                </div>
        `, async () => {
            const fileInput = document.getElementById('m-tpl-file');
            if (!fileInput.files.length) return alert('请选择模版文件');
            const name = document.getElementById('m-tpl-name').value.trim();
            if (!name) return alert('请输入模版名称');
            const encryptRule = document.getElementById('m-tpl-encrypt-rule').value.trim();
            const fd = new FormData();
            fd.append('file', fileInput.files[0]);
            fd.append('name', name);
            fd.append('tenant_id', document.getElementById('m-tpl-tenant').value);
            fd.append('description', document.getElementById('m-tpl-desc').value || '');
            fd.append('file_name_rule', document.getElementById('m-tpl-name-rule').value || '');
            fd.append('encrypt_type', encryptRule ? 'password' : 'none');
            fd.append('encrypt_password', encryptRule);
            fd.append('report_mode', document.getElementById('m-tpl-report-mode').value);
            fd.append('group_by', document.getElementById('m-tpl-group-by')?.value || '');
            fd.append('skip_rows', document.getElementById('m-tpl-skip-rows')?.value || '1');
            fd.append('name_field', document.getElementById('m-tpl-name-field')?.value || '');
            fd.append('split_by', document.getElementById('m-tpl-split-by')?.value || '');
            fd.append('show_empty_period', document.getElementById('m-tpl-show-empty')?.checked ? 'true' : 'false');
            fd.append('carry_over_sheets', document.getElementById('m-tpl-carry-sheets')?.value || '');
            const resp = await AUTH.authFetch('/api/admin/templates', { method: 'POST', body: fd });
            if (resp.ok) { this.closeModal(); this.loadTemplates(); }
            else { await _alertErr(resp, '创建失败'); }
        });
    },

    async showEditTemplate(id) {
        const resp = await AUTH.authFetch(`/api/admin/templates/${id}`);
        if (!resp.ok) return alert('获取模版失败');
        const t = await resp.json();
        const tenantOptions = this._tplTenants.map(tn =>
            `<option value="${tn}" ${t.tenant_id === tn ? 'selected' : ''}>租户: ${tn}</option>`
        ).join('');
        this.openModal('编辑模版', `
            <div style="display:flex;flex-direction:column;gap:12px;">
                <div class="form-group"><label>租户</label>
                    <select id="m-tpl-tenant" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                        <option value="" ${!t.tenant_id ? 'selected' : ''}>全局（所有租户可用）</option>
                        ${tenantOptions}
                    </select>
                </div>
                <div class="form-group"><label>模版名称</label>
                    <input id="m-tpl-name" value="${t.name}" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                </div>
                <div class="form-group"><label>描述</label>
                    <input id="m-tpl-desc" value="${t.description || ''}" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                </div>
                <div class="form-group"><label>替换文件（可选）</label>
                    <input id="m-tpl-file" type="file" accept=".xlsx,.xls,.xlsm">
                    <small style="color:#888;">当前文件: ${t.file_name}</small>
                </div>
                <div class="form-group"><label>文件名规则</label>
                    <input id="m-tpl-name-rule" value="${t.file_name_rule || ''}" placeholder="如: {year}{month}_薪资表_{姓名}" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">可用变量: {year} {month} {date} {tenant} {列名} {列名[:N]} {列名[-N:]}</small>
                </div>
                <div class="form-group"><label>加密规则</label>
                    <input id="m-tpl-encrypt-rule" value="${t.encrypt_password || ''}" placeholder="如: {身份证号码[:6]} 或 {姓名[:1]}{身份证号码[-6:]}" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">留空表示不加密。可用变量: {列名} {列名[:N]}前N位 {列名[-N:]}后N位</small>
                </div>
                <div class="form-group"><label>报表模式</label>
                    <select id="m-tpl-report-mode" onchange="Tools._toggleModeFields()" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                        <option value="fill" ${(t.report_mode||'fill')==='fill'?'selected':''}>fill — 整表填充</option>
                        <option value="block" ${t.report_mode==='block'?'selected':''}>block — 分组合并（每组一块）</option>
                        <option value="zip" ${t.report_mode==='zip'?'selected':''}>zip — 分组打包（每组一文件）</option>
                        <option value="sheet" ${t.report_mode==='sheet'?'selected':''}>sheet — 分组多Sheet（每组一个Sheet）</option>
                    </select>
                </div>
                <div class="form-group"><label>文件拆分字段</label>
                    <input id="m-tpl-split-by" value="${t.split_by || ''}" placeholder="如：部门（留空则不拆分文件）" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">按此列值将数据拆分到不同文件中，拆分后自动打包为 zip</small>
                </div>
                <div class="form-group"><label>整表带出结果Sheet</label>
                    <input id="m-tpl-carry-sheets" value="${(t.carry_over_sheets || '').replace(/"/g, '&quot;')}" placeholder="如：明细,#2（逗号分隔，可填sheet名或#序号）" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    <small style="color:#888;">把计算结果中这些 sheet 整表(值+格式)追加到报表末尾；#N 表示结果表第 N 个 sheet。留空不带出</small>
                </div>
                <div id="m-tpl-mode-fields" style="display:${(t.report_mode==='block'||t.report_mode==='zip'||t.report_mode==='sheet')?'block':'none'};">
                    <div class="form-group"><label>分组字段</label>
                        <input id="m-tpl-group-by" value="${t.group_by || ''}" placeholder="如: 工号" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                        <small style="color:#888;">按此列的值分组，每组独立填充模板</small>
                    </div>
                    <div class="form-group" id="m-tpl-skip-rows-group" style="display:${t.report_mode==='block'?'block':'none'};"><label>块间空行数</label>
                        <input id="m-tpl-skip-rows" type="number" value="${t.skip_rows ?? 1}" min="0" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    </div>
                    <div class="form-group" id="m-tpl-name-field-group" style="display:${t.report_mode==='zip'?'block':'none'};"><label>文件命名字段</label>
                        <input id="m-tpl-name-field" value="${t.name_field || ''}" placeholder="如: 姓名（用于 zip 内文件名）" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                    </div>
                    <div class="form-group" id="m-tpl-show-empty-group"><label style="display:flex;align-items:center;gap:6px;">
                        <input id="m-tpl-show-empty" type="checkbox" ${t.show_empty_period !== false ? 'checked' : ''}> 显示空月份（多月合并时补齐无数据的月份）
                    </label></div>
                </div>
        `, async () => {
            const fd = new FormData();
            const fileInput = document.getElementById('m-tpl-file');
            if (fileInput.files.length) fd.append('file', fileInput.files[0]);
            const encryptRule = document.getElementById('m-tpl-encrypt-rule').value.trim();
            fd.append('tenant_id', document.getElementById('m-tpl-tenant').value);
            fd.append('name', document.getElementById('m-tpl-name').value);
            fd.append('description', document.getElementById('m-tpl-desc').value || '');
            fd.append('file_name_rule', document.getElementById('m-tpl-name-rule').value || '');
            fd.append('encrypt_type', encryptRule ? 'password' : 'none');
            fd.append('encrypt_password', encryptRule);
            fd.append('report_mode', document.getElementById('m-tpl-report-mode').value);
            fd.append('group_by', document.getElementById('m-tpl-group-by')?.value || '');
            fd.append('skip_rows', document.getElementById('m-tpl-skip-rows')?.value || '1');
            fd.append('name_field', document.getElementById('m-tpl-name-field')?.value || '');
            fd.append('split_by', document.getElementById('m-tpl-split-by')?.value || '');
            fd.append('show_empty_period', document.getElementById('m-tpl-show-empty')?.checked ? 'true' : 'false');
            fd.append('carry_over_sheets', document.getElementById('m-tpl-carry-sheets')?.value || '');
            const resp = await AUTH.authFetch(`/api/admin/templates/${id}`, { method: 'PUT', body: fd });
            if (resp.ok) { this.closeModal(); this.loadTemplates(); }
            else { await _alertErr(resp, '更新失败'); }
        });
    },

    async deleteTemplate(id, name) {
        if (!confirm(`确定停用模版 ${name}？`)) return;
        const resp = await AUTH.authFetch(`/api/admin/templates/${id}`, { method: 'DELETE' });
        if (resp.ok) this.loadTemplates();
        else alert('操作失败');
    },

    _toggleModeFields() {
        const mode = document.getElementById('m-tpl-report-mode')?.value || 'fill';
        const fields = document.getElementById('m-tpl-mode-fields');
        const skipGroup = document.getElementById('m-tpl-skip-rows-group');
        const nameGroup = document.getElementById('m-tpl-name-field-group');
        if (fields) fields.style.display = (mode === 'block' || mode === 'zip' || mode === 'sheet') ? 'block' : 'none';
        if (skipGroup) skipGroup.style.display = mode === 'block' ? 'block' : 'none';
        if (nameGroup) nameGroup.style.display = mode === 'zip' ? 'block' : 'none';
    },

    downloadTemplate(id, fileName) {
        this._fetchAndDownload(`/api/admin/templates/${id}/download`, fileName);
    },

    // ==================== 下载工具（计算历史共用） ====================
    async downloadAsset(assetId, fileName, format) {
        try {
            let url = `/api/assets/${assetId}/download`;
            if (format) url += `?format=${format}`;
            const resp = await AUTH.authFetch(url);
            if (!resp.ok) return alert('下载失败: ' + resp.statusText);
            const blob = await resp.blob();
            const blobUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = blobUrl;
            let name = fileName || 'download.xlsx';
            if (format === 'pdf') name = name.replace(/\.(xlsx?|csv)$/i, '') + '.pdf';
            else if (format === 'encrypted') name = name.replace(/\.(xlsx?)$/i, '') + '_加密.xlsx';
            const contentType = resp.headers.get('content-type') || '';
            if (contentType.includes('zip') && !name.endsWith('.zip')) {
                name = name.replace(/\.(xlsx?)$/i, '') + '.zip';
            }
            a.download = name;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(blobUrl);
        } catch (e) {
            alert('下载失败: ' + e.message);
        }
    },

    downloadAssetEncrypted(assetId, fileName) {
        const password = prompt('请输入加密密码（默认123456）:', '123456');
        if (password === null) return;
        let url = `/api/assets/${assetId}/download?format=encrypted`;
        if (password) url += `&password=${encodeURIComponent(password)}`;
        this._fetchAndDownload(url, fileName.replace(/\.(xlsx?)$/i, '') + '_加密.xlsx');
    },

    async _fetchAndDownload(url, fileName) {
        try {
            const sep = url.includes('?') ? '&' : '?';
            const bustedUrl = `${url}${sep}_=${Date.now()}`;
            const resp = await AUTH.authFetch(bustedUrl, { cache: 'no-store' });
            if (!resp.ok) return alert('下载失败: ' + resp.statusText);
            const blob = await resp.blob();
            const blobUrl = URL.createObjectURL(blob);
            const a = document.createElement('a');
            a.href = blobUrl;
            a.download = fileName;
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
            URL.revokeObjectURL(blobUrl);
        } catch (e) {
            alert('下载失败: ' + e.message);
        }
    },

    _buildDownloadDropdown(assetId, fileName, btnStyle) {
        const id = `dl-${assetId}-${Date.now()}`;
        const style = btnStyle || '';
        return `<div style="position:relative;display:inline-block;">
            <button class="btn btn-sm" style="${style}" onclick="document.getElementById('${id}').style.display=document.getElementById('${id}').style.display==='block'?'none':'block'">
                下载 ▾
            </button>
            <div id="${id}" style="display:none;position:absolute;right:0;top:100%;background:#fff;border:1px solid #ddd;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.15);z-index:100;min-width:130px;">
                <div style="padding:6px 12px;cursor:pointer;font-size:12px;white-space:nowrap;" onmouseover="this.style.background='#f0f0f0'" onmouseout="this.style.background='#fff'" onclick="Tools.downloadAsset(${assetId},'${fileName.replace(/'/g, "\\'")}');this.parentElement.style.display='none'">
                    原始文件
                </div>
                <div style="padding:6px 12px;cursor:pointer;font-size:12px;white-space:nowrap;" onmouseover="this.style.background='#f0f0f0'" onmouseout="this.style.background='#fff'" onclick="Tools.downloadAsset(${assetId},'${fileName.replace(/'/g, "\\'")}','pdf');this.parentElement.style.display='none'">
                    下载 PDF
                </div>
                <div style="padding:6px 12px;cursor:pointer;font-size:12px;white-space:nowrap;" onmouseover="this.style.background='#f0f0f0'" onmouseout="this.style.background='#fff'" onclick="Tools.downloadAssetEncrypted(${assetId},'${fileName.replace(/'/g, "\\'")}');this.parentElement.style.display='none'">
                    加密 Excel
                </div>
            </div>
        </div>`;
    },

    // ==================== 计算历史 ====================
    async loadComputeHistory(page = 1) {
        const tenantId = document.getElementById('compute-tenant-filter')?.value || '';
        const status = document.getElementById('compute-status-filter')?.value || '';
        this._computePage = page;
        const offset = (page - 1) * this._pageSize;
        let url = `/api/compute2/tasks?limit=${this._pageSize}&offset=${offset}`;
        if (tenantId) url += `&tenant_id=${encodeURIComponent(tenantId)}`;
        if (status) url += `&status=${encodeURIComponent(status)}`;
        const resp = await AUTH.authFetch(url);
        if (!resp.ok) return;
        const result = await resp.json();
        const tbody = document.querySelector('#compute-history-table tbody');
        if (!result.items || !result.items.length) {
            tbody.innerHTML = '<tr><td colspan="10" class="empty-state">暂无计算记录</td></tr>';
        } else {
            tbody.innerHTML = result.items.map(t => `<tr>
                <td>${t.id}</td>
                <td>${t.tenant_id}</td>
                <td>${t.salary_year && t.salary_month ? t.salary_year + '-' + String(t.salary_month).padStart(2,'0') : '-'}</td>
                <td>${t.script_id || (t.analysis_report?.original_script_id || '-')}</td>
                <td><span class="status-${t.status}">${t.status}</span></td>
                <td>${t.inputs ? t.inputs.length : 0}</td>
                <td>${t.duration_seconds != null ? t.duration_seconds.toFixed(1) : '-'}</td>
                <td>${t.created_at ? new Date(t.created_at).toLocaleString() : '-'}</td>
                <td>${t.finished_at ? new Date(t.finished_at).toLocaleString() : '-'}</td>
                <td>
                    <button class="btn btn-sm" onclick="Tools.showComputeDetail(${t.id})">详情</button>
                    ${t.status === 'completed' ? `<button class="btn btn-sm btn-primary" style="margin-left:4px;" onclick="Tools.showGenerateReport(${t.id}, '${t.tenant_id}')">下载报表</button>` : ''}
                </td>
            </tr>`).join('');
        }
        this._renderPagination('compute-history-pagination', page, result.total || 0, this._pageSize, 'loadComputeHistory');
    },

    async showComputeDetail(taskId) {
        const resp = await AUTH.authFetch(`/api/compute2/tasks/${taskId}`);
        if (!resp.ok) return alert('获取详情失败');
        const task = await resp.json();

        let html = '<div style="max-height:500px;overflow-y:auto;">';

        html += `<div style="margin-bottom:16px;">
            <h4 style="margin:0 0 8px">基本信息</h4>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:13px;">
                <div>租户: <strong>${task.tenant_id}</strong></div>
                <div>状态: <span class="status-${task.status}">${task.status}</span></div>
                <div>耗时: ${task.duration_seconds != null ? task.duration_seconds.toFixed(1) + '秒' : '-'}</div>
                <div>脚本ID: ${task.script_id || '-'}</div>
            </div>
        </div>`;

        if (task.inputs && task.inputs.length) {
            html += `<div style="margin-bottom:16px;">
                <h4 style="margin:0 0 8px">输入文件 (${task.inputs.length})</h4>`;
            task.inputs.forEach(inp => {
                const downloadDropdown = inp.asset_id && inp.file_name
                    ? this._buildDownloadDropdown(inp.asset_id, inp.file_name || '')
                    : '';
                html += `<div style="padding:6px 10px;background:#f8f9fa;border-radius:4px;margin-bottom:4px;font-size:13px;display:flex;align-items:center;justify-content:space-between;">
                    <span>${inp.asset_name || inp.file_name || '未知'} <span style="color:#888;">(${inp.role})</span></span>
                    <span style="display:flex;align-items:center;gap:6px;">${inp.file_name ? '<span style="color:#999;font-size:11px;">' + inp.file_name + '</span>' : ''}${downloadDropdown}</span>
                </div>`;
            });
            html += '</div>';
        }

        if (task.output_assets && task.output_assets.length) {
            html += `<div style="margin-bottom:16px;">
                <h4 style="margin:0 0 8px">结果文件 (${task.output_assets.length})</h4>`;
            task.output_assets.forEach(asset => {
                const sizeKb = asset.file_size ? (asset.file_size / 1024).toFixed(1) + ' KB' : '';
                const downloadDropdown = this._buildDownloadDropdown(asset.id, asset.file_name || '', 'background:#2e7d32;color:#fff;');
                html += `<div style="padding:6px 10px;background:#e8f5e9;border-radius:4px;margin-bottom:4px;font-size:13px;display:flex;align-items:center;justify-content:space-between;">
                    <span>${asset.name} <span style="color:#999;font-size:11px;">${sizeKb}</span></span>
                    ${downloadDropdown}
                </div>`;
            });
            html += '</div>';
        }

        if (task.result_summary) {
            html += `<div style="margin-bottom:16px;">
                <h4 style="margin:0 0 8px">结果摘要</h4>
                <pre style="font-size:12px;background:#f5f5f5;padding:8px;border-radius:4px;">${JSON.stringify(task.result_summary, null, 2)}</pre>
            </div>`;
        }

        if (task.error_message) {
            html += `<div style="margin-bottom:16px;">
                <h4 style="margin:0 0 8px;color:red;">错误信息</h4>
                <div style="color:red;font-size:13px;background:#fff5f5;padding:8px;border-radius:4px;">${task.error_message}</div>
            </div>`;
        }

        html += '</div>';
        this.openModal(`计算任务 #${taskId} - 详情`, html, null);
    },

    async showGenerateReport(taskId, tenantId) {
        const resp = await AUTH.authFetch(`/api/admin/templates?tenant_id=${encodeURIComponent(tenantId)}&include_global=true`);
        if (!resp.ok) return alert('加载模版列表失败');
        const templates = await resp.json();
        if (!templates.length) return alert('暂无可用模版，请先在模版管理中创建');

        const tplOptions = templates.map(t =>
            `<option value="${t.id}">${t.name}${t.tenant_id ? '' : ' (全局)'}</option>`
        ).join('');

        this.openModal('下载报表', `
            <div style="display:flex;flex-direction:column;gap:14px;">
                <div class="form-group"><label>选择模版</label>
                    <select id="m-rpt-tpl" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;">
                        ${tplOptions}
                    </select>
                </div>
                <div class="form-group">
                    <label style="display:flex;align-items:center;gap:6px;cursor:pointer;">
                        启用历史数据（合并多个周期的计算结果）
                        <input type="checkbox" id="m-rpt-history" onchange="document.getElementById('m-rpt-period').style.display=this.checked?'flex':'none'">
                    </label>
                </div>
                <div id="m-rpt-period" style="display:none;gap:10px;align-items:center;">
                    <label>薪资周期从
                        <input type="month" id="m-rpt-from" style="padding:6px;border:1px solid #ddd;border-radius:4px;">
                    </label>
                    <label>至
                        <input type="month" id="m-rpt-to" style="padding:6px;border:1px solid #ddd;border-radius:4px;">
                    </label>
                </div>
            </div>
        `, async () => {
            const tplId = document.getElementById('m-rpt-tpl').value;
            const selectedTpl = templates.find(t => String(t.id) === String(tplId));
            const useHistory = document.getElementById('m-rpt-history').checked;
            const periodFrom = document.getElementById('m-rpt-from').value;
            const periodTo = document.getElementById('m-rpt-to').value;

            if (useHistory && (!periodFrom || !periodTo)) {
                return alert('启用历史时请选择薪资周期范围');
            }

            this.closeModal();
            const loadingEl = document.getElementById('loading-overlay');
            const loadingText = document.getElementById('loading-text');
            loadingEl.style.display = 'flex';
            loadingText.textContent = '报表生成中，请稍候...';

            try {
                const fd = new FormData();
                fd.append('task_id', taskId);
                fd.append('use_history', useHistory);
                if (useHistory) {
                    fd.append('period_from', periodFrom);
                    fd.append('period_to', periodTo);
                }

                const resp = await AUTH.authFetch(`/api/admin/templates/${tplId}/generate-report`, {
                    method: 'POST', body: fd
                });
                if (!resp.ok) {
                    let msg = '报表生成失败';
                    try { const err = await resp.json(); msg = err.detail || msg; } catch (_) {}
                    alert(msg);
                    return;
                }

                const blob = await resp.blob();
                if (!blob || blob.size === 0) {
                    alert('报表生成异常：文件为空（0 字节），请检查模版配置');
                    return;
                }

                const cd = resp.headers.get('content-disposition') || '';
                const tplMode = (selectedTpl && selectedTpl.report_mode) || 'fill';
                let fileName = tplMode === 'zip' ? '报表.zip' : '报表.xlsx';
                const fnMatch = cd.match(/filename\*?=(?:UTF-8''|")?([^";]+)/i);
                if (fnMatch) fileName = decodeURIComponent(fnMatch[1].replace(/"/g, ''));

                const blobUrl = URL.createObjectURL(blob);
                const a = document.createElement('a');
                a.href = blobUrl;
                a.download = fileName;
                document.body.appendChild(a);
                a.click();
                document.body.removeChild(a);
                URL.revokeObjectURL(blobUrl);
            } catch (e) {
                alert('报表生成失败: ' + e.message);
            } finally {
                loadingEl.style.display = 'none';
            }
        });
    },

    // ==================== 数据对比 ====================
    async startCompare() {
        const sourceFile = document.getElementById('compare-source-file').files[0];
        const targetFile = document.getElementById('compare-target-file').files[0];
        const primaryKeys = document.getElementById('compare-primary-keys').value.trim();

        if (!sourceFile || !targetFile) {
            alert('请选择基准文件和目标文件');
            return;
        }

        const loadingEl = document.getElementById('loading-overlay');
        const loadingText = document.getElementById('loading-text');
        loadingEl.style.display = 'flex';
        loadingText.textContent = '正在对比，请稍候...';

        try {
            const fd = new FormData();
            fd.append('source_file', sourceFile);
            fd.append('compare_file', targetFile);
            fd.append('primary_keys', primaryKeys || '工号,中文姓名');

            const resp = await AUTH.authFetch('/api/compare', { method: 'POST', body: fd });
            if (!resp.ok) {
                const err = await resp.json().catch(() => ({}));
                alert('对比失败: ' + (err.detail || resp.statusText));
                return;
            }
            const result = await resp.json();
            this.renderCompareResult(result);
            this.loadCompareHistory();
        } catch (e) {
            alert('对比失败: ' + e.message);
        } finally {
            loadingEl.style.display = 'none';
        }
    },

    renderCompareResult(result) {
        const container = document.getElementById('compare-result');
        container.style.display = 'block';

        const pct = (Math.min(1, result.match_rate) * 100).toFixed(1);
        const color = result.match_rate >= 0.95 ? '#4caf50' : result.match_rate >= 0.8 ? '#ff9800' : '#f44336';

        let summaryHtml = `
            <div style="display:flex;gap:20px;flex-wrap:wrap;">
                <div style="flex:1;min-width:200px;padding:16px;background:#f8f9fa;border-radius:8px;border-left:4px solid ${color};">
                    <div style="font-size:24px;font-weight:bold;color:${color};">${pct}%</div>
                    <div style="color:#666;font-size:13px;">总匹配率 (${result.matched_cells || 0}/${result.total_cells || 0})</div>
                </div>
                <div style="flex:1;min-width:200px;padding:16px;background:#f8f9fa;border-radius:8px;">
                    <div style="font-size:16px;font-weight:bold;">${result.different_cells || 0}</div>
                    <div style="color:#666;font-size:13px;">差异单元格</div>
                </div>
                ${result.download_url ? `<div style="display:flex;align-items:center;">
                    <button class="btn btn-primary" onclick="Tools._fetchAndDownload('${result.download_url}', '差异对比.xlsx')">下载差异报告</button>
                </div>` : ''}
            </div>`;
        document.getElementById('compare-summary').innerHTML = summaryHtml;

        let sheetHtml = '';
        const perSheet = result.per_sheet || {};
        if (Object.keys(perSheet).length > 0) {
            sheetHtml = '<h4 style="margin:16px 0 8px;">各Sheet匹配详情</h4>';
            sheetHtml += '<div style="display:grid;grid-template-columns:repeat(auto-fill,minmax(250px,1fr));gap:12px;">';
            for (const [name, info] of Object.entries(perSheet)) {
                const sPct = (Math.min(1, info.match_rate || 0) * 100).toFixed(1);
                const sColor = (info.match_rate || 0) >= 0.95 ? '#4caf50' : (info.match_rate || 0) >= 0.8 ? '#ff9800' : '#f44336';
                sheetHtml += `<div style="padding:12px;background:#fff;border:1px solid #e0e0e0;border-radius:8px;">
                    <div style="font-weight:bold;margin-bottom:4px;">${name} ${info.missing ? '<span style="color:red;">(缺失)</span>' : ''}</div>
                    <div style="font-size:20px;font-weight:bold;color:${sColor};">${sPct}%</div>
                    <div style="color:#888;font-size:12px;">匹配 ${info.matched_cells || 0}/${info.total_cells || 0} 单元格</div>
                </div>`;
            }
            sheetHtml += '</div>';
        }

        if (result.missing_sheets && result.missing_sheets.length) {
            sheetHtml += `<div style="margin-top:12px;padding:8px 12px;background:#fff3cd;border-radius:6px;color:#856404;">
                目标文件中缺失的Sheet: ${result.missing_sheets.join(', ')}
            </div>`;
        }
        if (result.warning) {
            sheetHtml += `<div style="margin-top:12px;padding:8px 12px;background:#f8d7da;border-radius:6px;color:#721c24;">
                注意: ${result.warning}对比结果可能不准确。
            </div>`;
        }
        document.getElementById('compare-sheet-details').innerHTML = sheetHtml;
    },

    async loadCompareHistory() {
        try {
            const resp = await AUTH.authFetch('/api/compare/history');
            if (!resp.ok) return;
            const items = await resp.json();
            const tbody = document.querySelector('#compare-history-table tbody');
            if (!items || !items.length) {
                tbody.innerHTML = '<tr><td colspan="6" class="empty-state">暂无对比记录</td></tr>';
                return;
            }
            tbody.innerHTML = items.map(item => `<tr>
                <td>${item.created_at ? new Date(item.created_at).toLocaleString() : '-'}</td>
                <td>${item.source_file || '-'}</td>
                <td>${item.compare_file || '-'}</td>
                <td><span style="color:${(item.match_rate || 0) >= 0.95 ? '#4caf50' : '#f44336'}">${(Math.min(1, item.match_rate || 0) * 100).toFixed(1)}%</span></td>
                <td>${item.sheet_count || 1}</td>
                <td>
                    ${item.download_url ? `<button class="btn btn-sm" onclick="Tools._fetchAndDownload('${item.download_url}', '差异对比.xlsx')">下载</button>` : ''}
                    <button class="btn btn-sm" style="margin-left:4px;" onclick="Tools.showCompareDetail('${item.session_id}')">详情</button>
                </td>
            </tr>`).join('');
        } catch (e) {
            console.error('加载对比历史失败:', e);
        }
    },

    async showCompareDetail(sessionId) {
        try {
            const resp = await AUTH.authFetch(`/api/compare/history/${sessionId}`);
            if (!resp.ok) return alert('获取详情失败');
            const result = await resp.json();
            this.renderCompareResult(result);
        } catch (e) {
            alert('获取详情失败: ' + e.message);
        }
    },
};

document.addEventListener('DOMContentLoaded', () => {
    Tools.init();
    document.addEventListener('click', (e) => {
        if (!e.target.closest('[id^="dl-"]') && !e.target.closest('.btn')) {
            document.querySelectorAll('[id^="dl-"]').forEach(el => el.style.display = 'none');
        }
    });
});
