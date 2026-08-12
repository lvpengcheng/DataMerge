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
let _integrateSchemeList = [];   // 方案列表页数据（GET /schemes）
let _intMode = 'create';         // create | apply | edit
let _intEditingSchemeId = null;  // edit 模式：正在修改的方案 id
let _intApplyTarget = null;      // apply 模式：目标方案 {id,name}
let _intSchemePage = 1;          // 方案列表分页：当前页
let _intSchemeFilterKw = '';     // 方案列表筛选：关键词
const _ALLOWED_EXT = new Set(['xlsx', 'xls', 'xlsm']);

async function _alertErr(resp, fallback) {
    let msg = fallback;
    try {
        const j = await resp.json();
        // detail 可能是对象/数组（FastAPI 422 校验错误），序列化后展示，避免弹出 [object Object]
        msg = typeof j.detail === 'object' ? JSON.stringify(j.detail) : (j.detail || j.message || fallback);
    } catch (_) {
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
        this.initSop();
        this.initSmartAssemble();
    },

    initTabs() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                this._activateTab(btn.dataset.tab);
            });
        });
        // 刷新后恢复上次所在页签（无记录/该页签无权限则用默认激活项）
        let saved = null;
        try { saved = localStorage.getItem('tools_active_tab'); } catch (e) {}
        const savedBtn = saved && document.querySelector('.tab-btn[data-tab="' + saved + '"]');
        if (savedBtn && savedBtn.offsetParent !== null) {   // offsetParent==null → 被 display:none 隐藏(无权限)
            this._activateTab(saved);
        }
    },

    _activateTab(tab) {
        const btn = document.querySelector('.tab-btn[data-tab="' + tab + '"]');
        const content = document.getElementById('tab-' + tab);
        if (!btn || !content) return;
        document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
        document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
        btn.classList.add('active');
        content.classList.add('active');
        try { localStorage.setItem('tools_active_tab', tab); } catch (e) {}
        if (tab === 'templates') this.loadTemplateTenants().then(() => this.loadTemplates());
        else if (tab === 'training-history') this.loadTrainingHistory();
        else if (tab === 'compute-history') this.loadComputeHistory();
        else if (tab === 'data-compare') this.loadCompareHistory();
        else if (tab === 'data-integrate') this.loadIntegrateSchemes();
        else if (tab === 'sop') this.loadSops();
        else if (tab === 'smart-assemble') this.initSmartAssemble();
    },

    // ==================== 弹窗工具 ====================
    openModal(title, bodyHtml, onConfirm, opts) {
        document.getElementById('modal-title').textContent = title;
        document.getElementById('modal-body').innerHTML = bodyHtml;
        document.getElementById('modal-overlay').style.display = 'flex';
        // 内容较宽的弹窗用 modal--wide 加宽
        const modalEl = document.getElementById('modal');
        if (modalEl) modalEl.classList.toggle('modal--wide', !!(opts && opts.wide));
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
        const btnNew = document.getElementById('btn-int-new-scheme');
        if (btnNew) btnNew.addEventListener('click', () => this._intShowWork('create'));
        const btnBack = document.getElementById('btn-int-back');
        if (btnBack) btnBack.addEventListener('click', () => this._intShowList());
    },

    // ==================== 整合对比：方案列表页 ====================
    _intSetListStatus(text, kind) {
        const el = document.getElementById('int-list-status');
        if (el) { el.textContent = text || ''; el.className = 'status' + (kind ? ' ' + kind : ''); }
    },

    async loadIntegrateSchemes() {
        this._intShowList();
        this._intSetListStatus('加载中...');
        try {
            const resp = await AUTH.authFetch('/api/tools/integrate/schemes');
            if (!resp.ok) { await _alertErr(resp, '加载方案失败'); this._intSetListStatus('加载失败', 'error'); return; }
            _integrateSchemeList = await resp.json();
            this._intRenderSchemeList();
            this._intSetListStatus(_integrateSchemeList.length ? `${_integrateSchemeList.length} 个方案` : '暂无方案', 'ok');
        } catch (e) {
            this._intSetListStatus('失败: ' + e.message, 'error');
        }
    },

    // 方案列表：筛选（按名称）+ 分页（每页 10 条）
    _intOnFilter(v) {
        _intSchemeFilterKw = (v || '').trim();
        _intSchemePage = 1;
        this._intRenderSchemeList();
    },

    _intGoPage(p) {
        _intSchemePage = p;
        this._intRenderSchemeList();
    },

    _intRenderPagination(total) {
        const box = document.getElementById('int-scheme-pagination');
        if (!box) return;
        const pageSize = 10;
        const totalPages = Math.max(1, Math.ceil(total / pageSize));
        if (totalPages <= 1) { box.innerHTML = ''; return; }
        let html = '';
        for (let p = 1; p <= totalPages; p++) {
            html += `<button class="btn btn-sm ${p === _intSchemePage ? 'btn-primary' : ''}" onclick="Tools._intGoPage(${p})">${p}</button>`;
        }
        html += `<span style="margin-left:8px;font-size:12px;color:#888;">共 ${total} 个方案</span>`;
        box.innerHTML = html;
    },

    _intRenderSchemeList() {
        const tbody = document.querySelector('#integrate-scheme-table tbody');
        if (!tbody) return;
        const kw = _intSchemeFilterKw.toLowerCase();
        const filtered = kw
            ? _integrateSchemeList.filter(s => (s.name || '').toLowerCase().includes(kw))
            : _integrateSchemeList;
        const pageSize = 10;
        const totalPages = Math.max(1, Math.ceil(filtered.length / pageSize));
        if (_intSchemePage > totalPages) _intSchemePage = totalPages;
        const pageItems = filtered.slice((_intSchemePage - 1) * pageSize, _intSchemePage * pageSize);

        if (!filtered.length) {
            tbody.innerHTML = '<tr><td colspan="5" style="text-align:center;color:#999;">' +
                (kw ? '没有匹配的方案' : '暂无方案，点击「新增方案」创建') + '</td></tr>';
            this._intRenderPagination(0);
            return;
        }
        const canApply = AUTH.hasPerm('tools.data_integrate.apply');
        const canEditPerm = AUTH.hasPerm('tools.data_integrate.edit');
        const canDelPerm = AUTH.hasPerm('tools.data_integrate.delete');
        tbody.innerHTML = pageItems.map(s => {
            // 最后修改时间：有 updated_at 用之，否则回退创建时间
            const t = ((s.updated_at || s.created_at) || '').replace('T', ' ').slice(0, 19);
            const btns = [];
            if (canApply) btns.push(`<button class="btn btn-sm btn-primary int-row-apply" data-id="${s.id}">应用</button>`);
            // 同组织可见即可修改配置（非创建人只能另存为，保存区会禁用「保存修改」）
            if (canEditPerm && s.can_modify) btns.push(`<button class="btn btn-sm int-row-edit" data-id="${s.id}">修改</button>`);
            if (canDelPerm && s.can_edit) btns.push(`<button class="btn btn-sm int-row-del" data-id="${s.id}">删除</button>`);
            return `<tr>
                <td>${_escape(s.name)}</td>
                <td style="text-align:center;">${s.file_count || ''}</td>
                <td>${_escape(s.creator_name || '')}</td>
                <td>${_escape(t)}</td>
                <td style="white-space:nowrap;">${btns.join(' ') || '<span style="color:#bbb;">无权限</span>'}</td>
            </tr>`;
        }).join('');
        const findById = (id) => _integrateSchemeList.find(x => x.id === id) || null;
        tbody.querySelectorAll('.int-row-apply').forEach(b =>
            b.addEventListener('click', () => this._intStartApply(findById(parseInt(b.dataset.id, 10)))));
        tbody.querySelectorAll('.int-row-edit').forEach(b =>
            b.addEventListener('click', () => this._intStartEdit(findById(parseInt(b.dataset.id, 10)))));
        tbody.querySelectorAll('.int-row-del').forEach(b =>
            b.addEventListener('click', () => this._intDeleteSchemeRow(findById(parseInt(b.dataset.id, 10)))));
        this._intRenderPagination(filtered.length);
    },

    _intResetWork() {
        _integrateFiles = [];
        _integrateAnalysis = null;
        _integrateSchemes = [];
        this._renderIntegrateList();
        const cfg = document.getElementById('integrate-config');
        if (cfg) { cfg.style.display = 'none'; cfg.innerHTML = ''; }
        this._setIntegrateStatus('');
        const rz = document.getElementById('int-apply-reasons');
        if (rz) { rz.style.display = 'none'; rz.innerHTML = ''; }
    },

    _intShowList() {
        _intMode = 'create'; _intEditingSchemeId = null; _intApplyTarget = null;
        const lv = document.getElementById('integrate-scheme-list-view');
        const wv = document.getElementById('integrate-work-view');
        if (lv) lv.style.display = '';
        if (wv) wv.style.display = 'none';
        this._intResetWork();
    },

    _intShowWork(mode, scheme) {
        _intMode = mode;
        _intEditingSchemeId = (mode === 'edit' && scheme) ? scheme.id : null;
        _intApplyTarget = (mode === 'apply' && scheme) ? { id: scheme.id, name: scheme.name } : null;
        const lv = document.getElementById('integrate-scheme-list-view');
        const wv = document.getElementById('integrate-work-view');
        if (lv) lv.style.display = 'none';
        if (wv) wv.style.display = '';
        this._intResetWork();
        const titleMap = {
            create: '新增方案',
            apply: '应用方案' + (scheme ? `：${scheme.name}` : ''),
            edit: '修改方案' + (scheme ? `：${scheme.name}` : ''),
        };
        const tt = document.getElementById('int-work-title');
        if (tt) tt.textContent = titleMap[mode] || '';
    },

    _intStartApply(scheme) {
        if (!scheme) return;
        this._intShowWork('apply', scheme);
        this._setIntegrateStatus('请上传本次要计算的数据文件后点「解析字段」', '');
    },

    _intStartEdit(scheme) {
        if (!scheme) return;
        this._intShowWork('edit', scheme);
        this._setIntegrateStatus('请上传与该方案结构一致的文件后点「解析字段」，回填后可调整并保存修改', '');
    },

    async _intDeleteSchemeRow(scheme) {
        if (!scheme || !confirm(`删除方案「${scheme.name}」？`)) return;
        try {
            const resp = await AUTH.authFetch('/api/tools/integrate/scheme/' + scheme.id, { method: 'DELETE' });
            if (!resp.ok) { await _alertErr(resp, '删除失败'); return; }
            this.loadIntegrateSchemes();
        } catch (e) { alert('删除失败: ' + e.message); }
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
            if (_intMode === 'apply') await this._intRunApply();
            else if (_intMode === 'edit') this._intPrefillForEdit();
        } catch (e) {
            this._setIntegrateStatus(`失败: ${e.message}`, 'error');
        } finally {
            btn.disabled = false;
        }
    },

    _intFindMatched(id) {
        return (_integrateAnalysis.matched_schemes || []).find(s => s.id === id) || null;
    },

    // 同结构表歧义：上传文件里存在表头结构相同的表（一个角色可对应多个文件），
    // 无法自动确定对应关系，弹出匹配框让操作人员手动指定"方案角色 ↔ 上传文件"。
    // 确认后更新 scheme.fp_to_file；取消返回 false。
    _intAskMatchMapping(scheme, ambiguous) {
        return new Promise(resolve => {
            const rows = ambiguous.map((a, i) => {
                const cur = (a.role_index != null && scheme.role_files && scheme.role_files[a.role_index])
                    ? scheme.role_files[a.role_index]
                    : ((scheme.fp_to_file || {})[a.fp] || (a.candidates || [])[0] || '');
                const saved = a.saved_file ? `（保存时：${_escape(a.saved_file)}）` : '';
                const opts = (a.candidates || []).map(c =>
                    `<option value="${_escape(c)}" ${c === cur ? 'selected' : ''}>${_escape(c)}</option>`).join('');
                return `<div style="margin-bottom:10px;display:flex;align-items:center;gap:8px;">
                    <label style="font-weight:500;min-width:130px;">${_escape(a.label)}${saved}：</label>
                    <select id="int-amb-${i}" style="min-width:280px;">${opts}</select>
                </div>`;
            }).join('');
            this.openModal('存在相同结构的表，请确认对应关系', `
                <div style="font-size:13px;color:#555;margin-bottom:12px;">
                    上传的文件中存在表头结构相同的表，无法自动确定与方案中角色的对应关系。
                    请为每个角色指定实际对应的上传文件（每个文件只能指定给一个角色）：
                </div>
                ${rows}
                <div style="margin-top:12px;display:flex;justify-content:flex-end;gap:8px;">
                    <button class="btn" onclick="Tools._intAmbResolve(false)">取消</button>
                    <button class="btn btn-primary" onclick="Tools._intAmbResolve(true)">确定匹配</button>
                </div>
            `, null);
            this._intAmbResolve = (ok) => {
                if (!ok) { this.closeModal(); resolve(false); return; }
                const mapping = {};
                const fileUsed = new Set();
                let dup = null;
                ambiguous.forEach((a, i) => {
                    const el = document.getElementById(`int-amb-${i}`);
                    if (!el) return;
                    const sel = el.value;
                    if (fileUsed.has(sel)) dup = sel;
                    fileUsed.add(sel);
                    // 以角色索引为准（同结构角色指纹相同，fp 键会互相覆盖）
                    if (a.role_index != null && scheme.role_files) scheme.role_files[a.role_index] = sel;
                    mapping[a.fp] = sel;
                });
                if (dup) { alert(`文件「${dup}」被指定给了多个角色，请重新确认`); return; }
                Object.assign(scheme.fp_to_file, mapping);
                this.closeModal();
                resolve(true);
            };
        });
    },

    // apply 模式：analyze 后调 apply-validate；不一致硬阻断并列出理由；一致则回填 + 直接生成下载。
    // 存在同结构表歧义时，先弹匹配框让操作人员手动确认对应关系，再回填执行。
    async _intRunApply() {
        if (!_intApplyTarget) return;
        const rz = document.getElementById('int-apply-reasons');
        try {
            const resp = await AUTH.authFetch('/api/tools/integrate/scheme/apply-validate', {
                method: 'POST', headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ session_id: _integrateAnalysis.session_id, scheme_id: _intApplyTarget.id }),
            });
            if (!resp.ok) { await _alertErr(resp, '校验失败'); return; }
            const res = await resp.json();
            if (!res.ok) {
                if (rz) {
                    rz.style.display = 'block';
                    rz.innerHTML = `<strong>上传文件与方案「${_escape(_intApplyTarget.name)}」不一致，无法应用：</strong>` +
                        '<ul style="margin:6px 0 0 18px;">' +
                        (res.reasons || []).map(r => `<li>${_escape(r)}</li>`).join('') + '</ul>';
                }
                this._setIntegrateStatus('校验未通过', 'error');
                return;
            }
            if (rz) { rz.style.display = 'none'; rz.innerHTML = ''; }
            const scheme = this._intFindMatched(_intApplyTarget.id);
            if (!scheme) { this._setIntegrateStatus('方案未能匹配到上传文件', 'error'); return; }
            const ambiguous = (scheme.ambiguous || []).filter(a => (a.candidates || []).length > 0);
            if (ambiguous.length) {
                const ok = await this._intAskMatchMapping(scheme, ambiguous);
                if (!ok) { this._setIntegrateStatus('已取消应用', ''); return; }
            }
            this._intFillFromScheme(scheme);
            this._setIntegrateStatus('校验通过，正在生成结果...', 'ok');
            await this._doIntegrateExecute();
        } catch (e) {
            this._setIntegrateStatus('应用失败: ' + e.message, 'error');
        }
    },

    // edit 模式：若上传文件结构与原方案匹配则回填配置供修改，否则从空白配置起
    _intPrefillForEdit() {
        if (!_intEditingSchemeId) return;
        const scheme = this._intFindMatched(_intEditingSchemeId);
        if (scheme) {
            this._intFillFromScheme(scheme);
            this._setIntegrateStatus('已回填原方案配置，可调整后保存修改', 'ok');
        } else {
            this._setIntegrateStatus('上传文件结构与原方案不同，将按当前配置另存为该方案', '');
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
            <h3>① 选择主表（模板）</h3>
            <div class="form-group">
                <select id="int-main-file" style="min-width:280px;">${fileOpts}</select>
                <span style="color:#888;font-size:12px;margin-left:8px;">主表将被原地更新（保留其余 sheet 与公式），对照表的值按关联键回填到主表。</span>
            </div>

            <h3 style="margin-top:16px;">② 关联键（每个文件）</h3>
            <div id="int-key-map"></div>
            <div class="form-group" style="margin-top:6px;">
                <label style="font-size:13px;">日期关联键归一：</label>
                <select id="int-date-mode">
                    <option value="off">关闭（纯文本原样比较）</option>
                    <option value="yearmonthday" selected>按年月日（2026-02-26，精确到日）</option>
                    <option value="yearmonth">按年月（2026-02，区分年度）</option>
                    <option value="month">按月（忽略年/日）</option>
                    <option value="day">按日（忽略年）</option>
                </select>
                <span style="color:#888;font-size:12px;margin-left:8px;">关联键是日期/月份时用：把 datetime、2026-02、2月26日 等归到同一粒度再匹配，两边同粒度才对得上。</span>
            </div>

            <h3 style="margin-top:16px;">③ 覆盖字段（勾选基准列 → 匹配对照列 / 公式）</h3>
            <div style="margin:6px 0;">
                <button type="button" class="btn btn-sm" id="int-ow-all">全选</button>
                <button type="button" class="btn btn-sm" id="int-ow-none">全不选</button>
                <button type="button" class="btn btn-sm" id="int-ow-match">智能匹配</button>
                <label style="margin-left:8px;font-size:13px;"><input type="checkbox" id="int-ow-ai" style="width:auto;"> 用 AI 匹配差异命名</label>
                <span class="status" id="int-ow-status" style="margin-left:8px;"></span>
            </div>
            <div id="int-ow-picker" style="border:1px solid #e3e7ed;border-radius:6px;padding:8px;max-height:150px;overflow:auto;"></div>
            <table class="data-table"><thead><tr><th>基准字段（主表列）</th><th>匹配的对照列（来源）</th><th style="width:32px;"></th></tr></thead>
                <tbody id="int-ow-list-rows"></tbody></table>
            <div style="color:#888;font-size:12px;margin-top:4px;">点单元格可选多列并用公式(基本工资+本月奖金，支持 +-*/、括号)组合；对照表同一主键多行时各列先跨行求和再代入。多张对照表都填时靠上优先取首个非空。</div>

            <h3 style="margin-top:16px;">④ 对比字段（可选，输出方式2用）</h3>
            <div style="margin:6px 0;">
                <button type="button" class="btn btn-sm" id="int-cmp-all">全选</button>
                <button type="button" class="btn btn-sm" id="int-cmp-none">全不选</button>
                <button type="button" class="btn btn-sm" id="int-cmp-match">智能匹配</button>
                <label style="margin-left:8px;font-size:13px;"><input type="checkbox" id="int-cmp-ai" style="width:auto;"> 用 AI 匹配差异命名</label>
                <span class="status" id="int-cmp-status" style="margin-left:8px;"></span>
            </div>
            <div id="int-cmp-picker" style="border:1px solid #e3e7ed;border-radius:6px;padding:8px;max-height:150px;overflow:auto;"></div>
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
            <div class="actions" id="int-save-row" style="margin-top:10px;display:flex;align-items:center;gap:8px;flex-wrap:wrap;">
                <span id="int-save-label">保存为方案：</span>
                <input type="text" id="int-scheme-name" placeholder="方案名称（如 5月工资整合）" style="min-width:200px;">
                <button type="button" class="btn" id="int-save-scheme">保存为方案</button>
                <button type="button" class="btn" id="int-save-as" style="display:none;">另存为</button>
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
        document.getElementById('int-ow-all').addEventListener('click', () => this._intToggleAllPick('ow', true));
        document.getElementById('int-ow-none').addEventListener('click', () => this._intToggleAllPick('ow', false));
        document.getElementById('int-cmp-all').addEventListener('click', () => this._intToggleAllPick('cmp', true));
        document.getElementById('int-cmp-none').addEventListener('click', () => this._intToggleAllPick('cmp', false));
        document.querySelectorAll('input[name="int-output-mode"]').forEach(r =>
            r.addEventListener('change', () => {
                document.getElementById('int-diff-settings').style.display =
                    (document.querySelector('input[name="int-output-mode"]:checked').value === '2') ? 'block' : 'none';
            }));
        document.getElementById('int-execute').addEventListener('click', () => this._doIntegrateExecute());

        // 保存区：按模式调整（create=保存为方案；edit=保存修改+另存为；apply=隐藏，不保存）
        const saveRow = document.getElementById('int-save-row');
        const saveBtn = document.getElementById('int-save-scheme');
        const saveAsBtn = document.getElementById('int-save-as');
        const nameInp = document.getElementById('int-scheme-name');
        if (_intMode === 'apply') {
            if (saveRow) saveRow.style.display = 'none';
        } else if (_intMode === 'edit') {
            const cur = _integrateSchemeList.find(s => s.id === _intEditingSchemeId);
            if (nameInp && cur) nameInp.value = cur.name;
            const canOverwrite = !!(cur && cur.can_edit);   // 仅创建人/管理员可保存修改覆盖原方案
            if (saveBtn) {
                saveBtn.textContent = '保存修改';
                saveBtn.disabled = !canOverwrite;
                saveBtn.title = canOverwrite ? '' : '仅创建人/管理员可保存修改到原方案，其他人请使用「另存为」';
            }
            if (saveAsBtn) saveAsBtn.style.display = 'inline-block';   // 另存为：原方案不动，按修改后的配置新建一个方案
            const lbl = document.getElementById('int-save-label');
            if (lbl) lbl.textContent = canOverwrite
                ? '保存修改到方案：'
                : '保存修改到方案（非创建人仅可另存为）：';
        }
        if (saveBtn) saveBtn.addEventListener('click', () => this._intSaveScheme(false));
        if (saveAsBtn) saveAsBtn.addEventListener('click', () => this._intSaveScheme(true));
    },

    _intFillFromScheme(scheme) {
        if (!scheme) return;
        const cfg = scheme.config || {};
        const f2f = scheme.fp_to_file || {};

        // 主表
        const mainSel = document.getElementById('int-main-file');
        mainSel.value = scheme.main_file;
        this._intRenderKeyMap();
        this._intRenderDiffCols();

        // 关联键：新方案按角色索引（同结构角色指纹相同，fp 不可作键）；旧方案回退按指纹
        const roleFiles = scheme.role_files || [];
        const keyByRole = cfg.key_map_by_role || {};
        if (Object.keys(keyByRole).length) {
            Object.entries(keyByRole).forEach(([idx, key]) => {
                const file = roleFiles[parseInt(idx, 10)];
                const el = file && document.querySelector(`.int-key[data-file="${CSS.escape(file)}"]`);
                if (el) el.value = key;
            });
        } else {
            Object.entries(cfg.key_map_by_fp || {}).forEach(([fp, key]) => {
                const file = f2f[fp];
                const el = document.querySelector(`.int-key[data-file="${CSS.escape(file || '')}"]`);
                if (el) el.value = key;
            });
        }

        // 覆盖/对比：新方案按 source_role 角色索引取文件（同结构角色各配各的），旧方案回退按指纹
        this._intRenderPicker('ow');
        this._intRenderPicker('cmp');
        const trans = (pairs) => (pairs || []).map(p => ({
            a_col: p.a_col,
            source_file: (p.source_role != null && roleFiles[p.source_role])
                ? roleFiles[p.source_role] : f2f[p.source_fp],
            expr: p.source_expr || p.source_col,
        })).filter(p => p.source_file);
        const applyPairs = (pairs, kind) => {
            const byA = {};
            trans(pairs).forEach(p => { (byA[p.a_col] = byA[p.a_col] || []).push({ file: p.source_file, expr: p.expr }); });
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
        if (cfg.date_key_mode) { const e = document.getElementById('int-date-mode'); if (e) e.value = cfg.date_key_mode; }
    },

    // 保存方案：asNew=true 为「另存为」（不带 scheme_id，按修改后的配置新建方案，原方案不动）
    async _intSaveScheme(asNew) {
        const name = (document.getElementById('int-scheme-name').value || '').trim();
        if (!name) { alert('请填写方案名称'); return; }
        if (asNew) {
            if (_intMode === 'edit' && _intEditingSchemeId) {
                const cur = _integrateSchemeList.find(s => s.id === _intEditingSchemeId);
                if (cur && name === cur.name) { alert('另存为需使用新的方案名称，请修改后再保存'); return; }
            }
            // 重名预检：与其它方案同名（后端也有同名校验兜底）
            const dup = _integrateSchemeList.find(s => s.name === name && s.id !== _intEditingSchemeId);
            if (dup) { alert(`已存在同名方案「${name}」，请换个名称`); return; }
        }
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
            date_key_mode: document.getElementById('int-date-mode')?.value || 'off',
        };
        // 另存为不带 scheme_id（走新建分支）；保存修改才带
        if (!asNew && _intMode === 'edit' && _intEditingSchemeId) payload.scheme_id = _intEditingSchemeId;
        const st = document.getElementById('int-save-status');
        st.textContent = '保存中...'; st.className = 'status';
        try {
            const resp = await AUTH.authFetch('/api/tools/integrate/scheme/save', {
                method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(payload),
            });
            if (!resp.ok) { await _alertErr(resp, '保存失败'); st.textContent = '保存失败'; st.className = 'status error'; return; }
            st.textContent = asNew ? '已另存为新方案' : (_intMode === 'edit' ? '修改已保存' : '方案已保存');
            st.className = 'status ok';
            setTimeout(() => this.loadIntegrateSchemes(), 600);   // 保存后返回方案列表
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

    // 全选/全不选基准列：改 checked 后必须触发 change，以复用上面的 _intAddListRow 联动（增删下方列表行）
    _intToggleAllPick(kind, checked) {
        document.querySelectorAll(`.int-${kind}-pick`).forEach(cb => {
            if (cb.checked !== checked) {
                cb.checked = checked;
                cb.dispatchEvent(new Event('change'));
            }
        });
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
            ? picks.map(p => `<span style="background:#eef4fb;padding:1px 6px;border-radius:3px;margin:1px;display:inline-block;font-size:12px;">${_escape(p.file)} · ${_escape(p.expr || p.col || '')}</span>`).join(' ')
            : '<span style="color:#999;">点击选择对照列…</span>';
    },

    // 公式校验：去掉本表所有列名后，只允许 +-*/、数字、括号、小数点、空白。
    // 列名被改动/写错 → 无法识别的残余会留在 rest 里，据此判定不合法（实现「列名不可编辑」）。
    _intCheckFx(file, expr) {
        const cols = (this._intFileMeta(file)?.columns || []).slice().sort((a, b) => b.length - a.length);
        let rest = expr || '';
        cols.forEach(c => { if (c) rest = rest.split(c).join(' '); });
        return { ok: /^[0-9eE.+\-*/()\s]*$/.test(rest), rest: rest.trim() };
    },

    // 公式分词：按已知列名（长名优先）把公式拆成 {t:'col'|'op'|'other', v} 序列，空白丢弃。
    // 供勾选联动做增量增删列，避免整体重写冲掉用户手动写的括号/常数/运算符。
    _intTokenize(expr, file) {
        const cols = (this._intFileMeta(file)?.columns || []).slice().filter(Boolean).sort((a, b) => b.length - a.length);
        const s = expr || '';
        const toks = [];
        const isOp = ch => '+-*/()'.indexOf(ch) >= 0;
        const colAt = pos => cols.find(c => s.startsWith(c, pos));
        let i = 0;
        while (i < s.length) {
            if (/\s/.test(s[i])) { i++; continue; }
            const c = colAt(i);
            if (c) { toks.push({ t: 'col', v: c }); i += c.length; continue; }
            if (isOp(s[i])) { toks.push({ t: 'op', v: s[i] }); i++; continue; }
            let j = i;
            while (j < s.length && !/\s/.test(s[j]) && !isOp(s[j]) && !colAt(j)) j++;
            toks.push({ t: 'other', v: s.slice(i, j) });
            i = j;
        }
        return toks;
    },

    // 增量·勾上：公式里没有该列才在末尾用 + 追加；已存在则原样返回（不动手动公式）。
    _intFxAddCol(expr, col, file) {
        const cur = (expr || '').trim();
        const toks = this._intTokenize(cur, file);
        if (toks.some(t => t.t === 'col' && t.v === col)) return cur;
        return cur ? `${cur}+${col}` : col;
    },

    // 增量·取消：删掉该列 token 及紧邻的一个运算符，再清理空括号/首尾悬空运算符；
    // 删完若已无任何列则清空（对应「取消所有勾选 → 空」）。公式里本就没有该列则原样返回。
    _intFxRemoveCol(expr, col, file) {
        let toks = this._intTokenize(expr || '', file);
        if (!toks.some(t => t.t === 'col' && t.v === col)) return (expr || '');
        for (let k = toks.length - 1; k >= 0; k--) {
            if (!(toks[k].t === 'col' && toks[k].v === col)) continue;
            const prev = toks[k - 1], next = toks[k + 1];
            if (prev && prev.t === 'op' && '+-*/'.indexOf(prev.v) >= 0) { toks.splice(k - 1, 2); k--; }
            else if (next && next.t === 'op' && '+-*/'.indexOf(next.v) >= 0) { toks.splice(k, 2); }
            else { toks.splice(k, 1); }
        }
        if (!toks.some(t => t.t === 'col')) return '';
        let out = toks.map(t => t.v).join(''), prev = null;
        while (out !== prev) {
            prev = out;
            out = out.replace(/\(\)/g, '')          // 空括号
                     .replace(/^[+\-*/]+/, '')       // 首部悬空运算符
                     .replace(/[+\-*/]+$/, '')       // 尾部悬空运算符
                     .replace(/\(\s*[+*/]/g, '(')     // 左括号后紧跟 +*/（保留一元 -）
                     .replace(/[+\-*/]\)/g, ')');     // 右括号前的悬空运算符
        }
        return out.trim();
    },

    // 公式里列名的集合签名（排序后拼接）：用于判定键盘编辑是否动了列名。
    _intColSig(expr, file) {
        return this._intTokenize(expr, file).filter(t => t.t === 'col').map(t => t.v).sort().join('');
    },

    // 弹出勾选对照列：按对照文件分组（支持多表）。每个对照表勾选若干列后，底部公式框自动
    // 用「+」连接已选列，可手动改成任意四则运算(基本工资+本月奖金 / (应发-扣款)*0.8)。
    // 每张对照表产出一条 {file, expr}；多表并存=优先级回退（靠上的先取，非空即用）。
    // 同一主键在对照表里有多行时，公式里每个列会先跨行求和再代入（后端 eval_source_expr）。
    _intOpenSrcPicker(kind, tr) {
        let selected = [];
        try { selected = JSON.parse(tr.dataset.src || '[]'); } catch (_) {}
        // 归一到 {file: expr}（兼容旧格式 {file,col}）
        const exprByFile = {};
        selected.forEach(s => { exprByFile[s.file] = (s.expr || s.col || exprByFile[s.file] || ''); });
        const files = this._intNonMainFiles();
        const body = `
            <div style="font-size:12px;color:#888;margin-bottom:8px;">
              每张对照表勾选列 → 下方公式框自动用「+」连接，可手动改（仅支持列名与 +-*/、数字、括号）。
              多张表都填时，靠上的优先（取首个非空）。同一主键多行会先把各列跨行求和再代入公式。
            </div>
            ${files.map((f, fi) => {
                const preset = exprByFile[f.name] || '';
                return `
                <div class="int-fgrp" data-file="${_escape(f.name)}" style="margin-bottom:12px;padding-bottom:8px;border-bottom:1px dashed #e0e0e0;">
                    <div style="font-weight:600;color:#2c3e50;margin-bottom:4px;">📄 ${_escape(f.name)}</div>
                    <div style="display:flex;flex-wrap:wrap;gap:4px 14px;margin-bottom:6px;">
                        ${(f.columns || []).map(c => {
                            const inExpr = preset && preset.indexOf(c) >= 0 ? 'checked' : '';
                            return `<label style="display:inline-flex;align-items:center;gap:4px;font-size:13px;">
                                <input type="checkbox" class="int-srcpick-cb" data-file="${_escape(f.name)}" data-col="${_escape(c)}" ${inExpr} style="width:auto;"> ${_escape(c)}</label>`;
                        }).join('')}
                    </div>
                    <div style="display:flex;align-items:flex-start;gap:6px;">
                        <span style="font-size:12px;color:#666;white-space:nowrap;padding-top:6px;">公式：</span>
                        <div style="flex:1;">
                            <textarea class="int-fx" data-file="${_escape(f.name)}" rows="5"
                                   placeholder="勾选上方列自动相加；列名不可改，只能编辑 + - * / 括号 数字（如 (基本工资+本月奖金)*0.8）"
                                   style="width:100%;box-sizing:border-box;font-size:13px;padding:4px 6px;line-height:1.5;resize:vertical;font-family:monospace;">${_escape(preset)}</textarea>
                            <div class="int-fx-err" data-file="${_escape(f.name)}" style="font-size:11px;color:#d32f2f;min-height:14px;margin-top:2px;"></div>
                        </div>
                    </div>
                </div>`;
            }).join('')}`;
        this.openModal(`为「${tr.dataset.aCol || ''}」选择对照列 / 公式（可多表）`, body, () => {
            const picks = [];
            for (const grp of document.querySelectorAll('#modal-body .int-fgrp')) {
                const file = grp.dataset.file;
                const fx = grp.querySelector('.int-fx');
                const expr = (fx?.value || '').trim();
                if (!expr) continue;
                // 校验：列名不可修改，去掉本表所有列名后只允许 +-*/、数字、括号、小数点、空白
                const chk = this._intCheckFx(file, expr);
                if (!chk.ok) {
                    alert(`「${file}」的公式含不支持的内容：列名不可修改，只能编辑 +、-、*、/、括号、数字。\n无法识别的部分：${chk.rest}`);
                    return; // 不关闭，让用户改
                }
                picks.push({ file, expr });
            }
            tr.dataset.src = JSON.stringify(picks);
            this._intUpdateSrcCell(tr);
            this.closeModal();
        });
        // 勾选联动（增量增删）+ 列名锁定：列名只能通过上方勾选增删；框内键盘编辑一旦改动列名集合
        // （删/改/手动加列）立即回退，只放行 + - * / 括号 数字 的编辑。
        document.querySelectorAll('#modal-body .int-fgrp').forEach(grp => {
            const fx = grp.querySelector('.int-fx');
            const err = grp.querySelector('.int-fx-err');
            const file = grp.dataset.file;
            let prev = fx.value;                          // 上一版已接受的公式
            let prevSig = this._intColSig(prev, file);    // 及其列名集合签名
            const validate = () => {
                const expr = (fx.value || '').trim();
                if (!expr) { err.textContent = ''; return; }
                const chk = this._intCheckFx(file, expr);
                err.textContent = chk.ok ? '' : `只能编辑 + - * / 括号 数字；无法识别：${chk.rest}`;
            };
            const accept = (val) => { fx.value = val; prev = val; prevSig = this._intColSig(val, file); validate(); };
            // 键盘编辑：列名集合变了就回退（列名不可删/改，只能靠勾选）；只改运算符/括号/数字才放行
            fx.addEventListener('input', () => {
                if (this._intColSig(fx.value, file) !== prevSig) {
                    fx.value = prev;
                    err.textContent = '列名不可删除/修改，请用上方勾选来增删列';
                    setTimeout(validate, 1500);
                    return;
                }
                prev = fx.value;
                validate();
            });
            grp.querySelectorAll('.int-srcpick-cb').forEach(cb => cb.addEventListener('change', () => {
                accept(cb.checked
                    ? this._intFxAddCol(fx.value, cb.dataset.col, file)
                    : this._intFxRemoveCol(fx.value, cb.dataset.col, file));
            }));
            validate();
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
                if (m) { tr.dataset.src = JSON.stringify([{ file: m.source_file, expr: m.source_col }]); hit++; }
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
            picks.forEach(p => {
                const expr = p.expr || p.col;   // 兼容旧格式 {file,col}
                if (!expr) return;
                // source_col 保留(=单列时的列名)，便于旧后端/方案兼容；source_expr 为求值公式
                out.push({ a_col: tr.dataset.aCol, source_file: p.file, source_expr: expr, source_col: expr });
            });
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
            date_key_mode: document.getElementById('int-date-mode')?.value || 'off',
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

        // 主键（每文件，可多选组成复合主键）——默认勾选 suggested_key（身份证/工号等唯一标识列）
        document.getElementById('merge-key-map').innerHTML = files.map(f => {
            const sk = f.suggested_key || (f.columns || [])[0] || '';
            return `
            <div style="margin:4px 0;display:flex;align-items:flex-start;gap:8px;">
                <span style="display:inline-block;min-width:160px;padding-top:2px;">${_escape(f.name)}：</span>
                <div style="display:flex;flex-wrap:wrap;gap:4px 14px;">
                    ${(f.columns || []).map(c => `
                        <label style="display:inline-flex;align-items:center;gap:4px;font-size:13px;">
                            <input type="checkbox" class="mg-key" data-file="${_escape(f.name)}" data-col="${_escape(c)}" ${c === sk ? 'checked' : ''}>
                            ${_escape(c)}
                        </label>`).join('')}
                </div>
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
        document.querySelectorAll('.mg-key:checked').forEach(cb => {
            (key_map[cb.dataset.file] = key_map[cb.dataset.file] || []).push(cb.dataset.col);
        });
        // 校验：每个文件都要至少选一个主键列
        const missing = (_mergeAnalysis.files || []).filter(f => !(key_map[f.name] || []).length);
        if (missing.length) {
            this._setMergeExecStatus(`请为每个文件至少选一个主键列（缺：${missing.map(f => f.name).join('、')}）`, 'error');
            return;
        }

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
                date_key_mode: document.getElementById('merge-date-mode').value,
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

    // ==================== 智能组表 ====================
    _saTenants: [],
    _saRules: [],
    _saEventSource: null,
    _saTaskId: null,
    _saLastEventId: 0,

    initSmartAssemble() {
        const srcInput = document.getElementById('sa-source-files');
        if (srcInput) srcInput.addEventListener('change', () => {
            document.getElementById('sa-source-file-list').textContent =
                Array.from(srcInput.files).map(f => f.name).join(', ');
        });
        const tplInput = document.getElementById('sa-template-file');
        if (tplInput) tplInput.addEventListener('change', () => {
            document.getElementById('sa-template-file-list').textContent =
                Array.from(tplInput.files).map(f => f.name).join(', ');
        });
        const tenantSel = document.getElementById('sa-tenant-select');
        if (tenantSel) tenantSel.addEventListener('change', () => this._saLoadRules());

        this._saLoadTenants();
        // 恢复进行中任务（1小时内）
        try {
            const saved = sessionStorage.getItem('sa_active_task');
            if (saved) {
                const s = JSON.parse(saved);
                if (s.taskId && Date.now() - (s.ts || 0) < 3600 * 1000) {
                    this._saResume(s.taskId, s.lastEventId || 0);
                } else {
                    sessionStorage.removeItem('sa_active_task');
                }
            }
        } catch (e) {}
    },

    async _saLoadTenants() {
        try {
            const resp = await AUTH.authFetch('/api/training-history');
            if (!resp.ok) return;
            const data = await resp.json();
            const historyData = data.history || {};
            this._saTenants = Object.keys(historyData);
            const sel = document.getElementById('sa-tenant-select');
            // 默认「空租户」（值为空，主要用全局规则）
            sel.innerHTML = '<option value="">（空租户）</option>' +
                this._saTenants.map(t => `<option value="${t}">${t}</option>`).join('');
            sel.value = '';
            this._saLoadRules();
        } catch (e) {}
    },

    async _saLoadRules() {
        const tenantId = document.getElementById('sa-tenant-select').value;
        try {
            const resp = await AUTH.authFetch('/api/assemble/rules?scope=available&tenant_id=' + encodeURIComponent(tenantId));
            if (!resp.ok) return;
            const data = await resp.json();
            this._saRules = data.items || [];
            const sel = document.getElementById('sa-rule-select');
            sel.innerHTML = '<option value="">（不选规则，AI 仅按结构分析）</option>' +
                this._saRules.map(r => `<option value="${r.id}">${r.name}（${r.scope === 'global' ? '全局' : r.tenant_id}）</option>`).join('');
            // 默认选中第一条全局规则（存在时）
            const globalRule = this._saRules.find(r => r.scope === 'global');
            if (globalRule) sel.value = String(globalRule.id);
        } catch (e) {}
    },

    // 对齐智算 addLog：时间戳 + 着色条目
    _saLog(level, message) {
        const logContent = document.getElementById('sa-log');
        if (!logContent) return;
        const time = new Date().toLocaleTimeString('zh-CN', { hour12: false });
        const entry = document.createElement('div');
        entry.className = 'log-entry ' + (level || 'info');
        entry.innerHTML = `<span class="log-timestamp">[${time}]</span>${message}`;
        logContent.appendChild(entry);
        logContent.scrollTop = logContent.scrollHeight;
    },

    _saSetStatus(text, pct) {
        document.getElementById('sa-status').textContent = text;
        if (pct != null) {
            document.getElementById('sa-progress').style.width = pct + '%';
            document.getElementById('sa-progress-text').textContent = Math.round(pct) + '%';
        }
    },

    // AI 生成的代码流累积到折叠代码区（对齐智训的代码区，不刷日志区）
    _saCodeBuf: '',
    _saThinkingBuf: '',

    _saAppendCode(chunk) {
        this._saCodeBuf += chunk;
        const el = document.getElementById('sa-code-content');
        if (el) {
            el.textContent = this._saCodeBuf;
            el.scrollTop = el.scrollHeight;
        }
        const cnt = document.getElementById('sa-code-count');
        if (cnt) cnt.textContent = `（${this._saCodeBuf.length} 字符）`;
    },

    _saAppendThinking(chunk) {
        this._saThinkingBuf += chunk;
        const logEl = document.getElementById('sa-log');
        const time = new Date().toLocaleTimeString('zh-CN', { hour12: false });
        let last = logEl && logEl.lastElementChild;
        if (last && last.classList.contains('thinking')) {
            last.innerHTML = `<span class="log-timestamp">[${time}]</span>${this._saThinkingBuf}`;
        } else {
            this._saLog('thinking', this._saThinkingBuf);
        }
    },

    async saStart() {
        const tenantId = document.getElementById('sa-tenant-select').value;
        const ruleId = document.getElementById('sa-rule-select').value;
        const force = document.getElementById('sa-force-rematch').checked;
        const srcFiles = document.getElementById('sa-source-files').files;
        const tplFile = document.getElementById('sa-template-file').files[0];
        // 租户可留空（空租户走通用全局规则，后端落到 __assemble__ 工具租户）
        if (!srcFiles.length) return alert('请选择源文件');
        if (!tplFile) return alert('请选择模板文件');

        const btn = document.getElementById('sa-start-btn');
        btn.disabled = true;
        btn.textContent = '组表中...';
        document.getElementById('sa-log').innerHTML = '';
        document.getElementById('sa-code-content').textContent = '';
        document.getElementById('sa-code-count').textContent = '';
        this._saCodeBuf = '';
        this._saThinkingBuf = '';
        document.getElementById('sa-result-card').style.display = 'none';
        this._saSetStatus('提交任务...', 5);

        const fd = new FormData();
        fd.append('tenant_id', tenantId);
        fd.append('rule_id', ruleId || '0');
        fd.append('force_rematch', force ? 'true' : 'false');
        for (const f of srcFiles) fd.append('source_files', f);
        fd.append('template_file', tplFile);

        try {
            const resp = await AUTH.authFetch('/api/assemble/submit', { method: 'POST', body: fd });
            if (!resp.ok) return _alertErr(resp, '提交失败');
            const data = await resp.json();
            this._saTaskId = data.task_id;
            this._saLastEventId = 0;
            sessionStorage.setItem('sa_active_task', JSON.stringify({ taskId: data.task_id, lastEventId: 0, ts: Date.now() }));
            this._saConnectStream(data.task_id, 0);
        } catch (e) {
            btn.disabled = false;
            btn.textContent = '开始组表';
            alert('提交失败: ' + e.message);
        }
    },

    _saResume(taskId, lastEventId) {
        this._saTaskId = taskId;
        this._saLastEventId = lastEventId || 0;
        const btn = document.getElementById('sa-start-btn');
        btn.disabled = true;
        btn.textContent = '组表中...';
        document.getElementById('sa-log').innerHTML = '';
        this._saLog('info', '（检测到进行中的组表任务 #' + taskId + '，恢复进度...）');
        this._saConnectStream(taskId, lastEventId || 0);
    },

    _saConnectStream(taskId, fromId) {
        if (this._saEventSource) {
            this._saEventSource.close();
            this._saEventSource = null;
        }
        const es = new EventSource('/api/assemble/tasks/' + taskId + '/stream?last_event_id=' + fromId);
        this._saEventSource = es;
        es.onmessage = (e) => {
            try {
                const event = JSON.parse(e.data);
                if (e.lastEventId) {
                    this._saLastEventId = parseInt(e.lastEventId);
                    sessionStorage.setItem('sa_active_task', JSON.stringify({
                        taskId: taskId, lastEventId: this._saLastEventId, ts: Date.now() }));
                }
                this._saHandleEvent(event);
                if (event.type === 'complete' || event.type === 'error') {
                    es.close();
                    this._saEventSource = null;
                    sessionStorage.removeItem('sa_active_task');
                }
            } catch (err) {
                console.error('解析组表SSE事件失败:', err);
            }
        };
        es.onerror = () => {
            es.close();
            this._saEventSource = null;
            this._saLog('warning', '⚠️ 连接中断，尝试重连...');
            setTimeout(() => this._saReconnect(taskId), 2000);
        };
    },

    async _saReconnect(taskId) {
        try {
            const resp = await AUTH.authFetch('/api/assemble/tasks/' + taskId + '/status');
            if (!resp.ok) return setTimeout(() => this._saReconnect(taskId), 5000);
            const st = await resp.json();
            if (st.status === 'completed') {
                this._saShowResult(st);
                return;
            }
            if (st.status === 'error') {
                this._saLog('error', '❌ ' + (st.error || '组表失败'));
                this._saResetBtn();
                return;
            }
            this._saConnectStream(taskId, this._saLastEventId);
        } catch (e) {
            setTimeout(() => this._saReconnect(taskId), 5000);
        }
    },

    _saHandleEvent(event) {
        switch (event.type) {
            case 'status': {
                const pct = event.status === 'pending' ? 5
                    : event.status === 'analyzing' ? 15
                    : event.status === 'generating' ? 40
                    : event.status === 'executing' ? 70
                    : event.status === 'complete' ? 100 : 5;
                this._saSetStatus(event.message || event.status, pct);
                break;
            }
            case 'log':
                if (!event.message) break;
                // AI 生成时的 [CODE] 代码流 → 累积到折叠代码区（对齐智训的代码区展示）
                const codeMatch = event.message.match(/\[CODE\]\s*([\s\S]*)/);
                if (codeMatch) {
                    this._saAppendCode(codeMatch[1]);
                    break;
                }
                let level = 'info';
                if (event.message.includes('✅')) level = 'success';
                else if (event.message.includes('⚠️')) level = 'warning';
                else if (event.message.includes('❌') || event.message.includes('失败')) level = 'error';
                this._saLog(level, event.message);
                break;
            case 'mapping':
                this._saLog('mapping', event.message || '');
                break;
            case 'thinking':
                // 思考流合并成一块显示（不逐 chunk 刷屏）
                this._saAppendThinking(event.content || '');
                break;
            case 'complete':
                this._saSetStatus('组表完成', 100);
                this._saLog('success', '✅ ' + (event.message || '组表完成'));
                this._saShowResult({ status: 'completed', output_files: event.output_files || [], matched_from_cache: event.matched_from_cache });
                break;
            case 'error':
                this._saSetStatus('组表失败', 0);
                this._saLog('error', '❌ ' + (event.message || '组表失败'));
                this._saResetBtn();
                break;
        }
    },

    _saShowResult(st) {
        const resultCard = document.getElementById('sa-result-card');
        const resultDownloads = document.getElementById('sa-result-downloads');
        const files = st.output_files || [];

        resultCard.className = 'result-card result-success';
        resultCard.innerHTML = `
            <div class="result-row">
                <div class="result-item">
                    <div class="label">组表状态</div>
                    <div class="value" style="color: #28a745; font-size: 20px;">成功${st.matched_from_cache ? '（命中存档）' : ''}</div>
                </div>
                <div class="result-item">
                    <div class="label">输出文件</div>
                    <div class="value">${files.length ? files.map(f => `${f.includes('_纯值') ? '纯值版' : '原版'}：${f}`).join('<br>') : 'N/A'}</div>
                </div>
            </div>
        `;
        resultDownloads.innerHTML = files.length
            ? files.map(f => `<button class="btn btn-download" onclick="Tools._saDownload(${this._saTaskId}, '${f.replace(/'/g, "\\'")}')">下载${f.includes('_纯值') ? '纯值版' : '原版'}</button>`).join('')
            : '';

        document.getElementById('sa-feedback-bar').style.display = st.status === 'completed' ? '' : 'none';
        document.getElementById('sa-feedback-msg').textContent = '';
        this._saResetBtn();
    },

    _saDownload(taskId, fileName) {
        const sep = fileName.includes('?') ? '&' : '?';
        const url = '/api/assemble/download/' + taskId + '/' + encodeURIComponent(fileName) + sep + '_=' + Date.now();
        const a = document.createElement('a');
        a.href = url;
        a.setAttribute('download', fileName);
        document.body.appendChild(a);
        a.click();
        a.remove();
    },

    async saFeedback(correct) {
        if (!this._saTaskId) return;
        const fd = new FormData();
        fd.append('correct', correct ? 'true' : 'false');
        try {
            const resp = await AUTH.authFetch('/api/assemble/tasks/' + this._saTaskId + '/feedback', { method: 'POST', body: fd });
            if (!resp.ok) return _alertErr(resp, '反馈提交失败');
            const data = await resp.json();
            document.getElementById('sa-feedback-msg').textContent = data.message || '';
            document.getElementById('sa-feedback-bar').style.opacity = '0.6';
            if (!correct) {
                this._saLog('error', '❌ 已反馈：命中映射标待复核，建议勾选"强制重新匹配"重新组表');
            }
        } catch (e) {
            alert('反馈失败: ' + e.message);
        }
    },

    _saResetBtn() {
        const btn = document.getElementById('sa-start-btn');
        btn.disabled = false;
        btn.textContent = '开始组表';
    },

    // ==================== SOP 维护 ====================
    _sopEntries: [],

    initSop() {
        // SOP 数据在 tab 激活时懒加载（_activateTab → loadSops），此处仅占位
    },

    async loadSops() {
        const keyword = document.getElementById('sop-keyword')?.value || '';
        const status = document.getElementById('sop-status-filter')?.value || '';
        const params = [];
        if (keyword) params.push('keyword=' + encodeURIComponent(keyword));
        if (status) params.push('status=' + encodeURIComponent(status));
        const url = '/api/tools/sop/entries' + (params.length ? '?' + params.join('&') : '');
        try {
            const resp = await AUTH.authFetch(url);
            if (!resp.ok) return _alertErr(resp, '加载SOP列表失败');
            const data = await resp.json();
            this._sopEntries = data.items || [];
            this._renderSopList();
        } catch (e) {
            alert('加载SOP列表失败: ' + e.message);
        }
    },

    _renderSopList() {
        const tbody = document.querySelector('#sop-table tbody');
        const items = this._sopEntries || [];
        if (!items.length) {
            tbody.innerHTML = '<tr><td colspan="8" class="empty-state">暂无SOP条目</td></tr>';
            return;
        }
        const canCreate = AUTH.hasPerm('tools.sop.create');
        const canReview = AUTH.hasPerm('tools.sop.review');
        const canManage = AUTH.hasPerm('tools.sop.manage');
        tbody.innerHTML = items.map(e => {
            const btns = [`<button class="btn btn-sm" onclick="Tools.showSopDetail(${e.id})">详情</button>`];
            if (canCreate) btns.push(`<button class="btn btn-sm" style="margin-left:4px;" onclick="Tools.showSopUpload(${e.id})">上传文件</button>`);
            if (canReview && e.status === 'ai_passed') btns.push(`<button class="btn btn-sm" style="margin-left:4px;background:#ffc107;" onclick="Tools.showSopReview(${e.id})">审核</button>`);
            if (canManage) btns.push(`<button class="btn btn-sm btn-danger" style="margin-left:4px;" onclick="Tools.deleteSop(${e.id})">删除</button>`);
            return `<tr>
                <td>${e.id}</td>
                <td>${_escape(e.customer_name)}</td>
                <td>${_escape(e.description || '-')}</td>
                <td>${this._sopStatusHtml(e.status)}</td>
                <td>${e.round_no || 0}</td>
                <td>${_escape(e.ai_comment || '-')}</td>
                <td>${e.updated_at ? new Date(e.updated_at).toLocaleString() : '-'}</td>
                <td class="actions">${btns.join('')}</td>
            </tr>`;
        }).join('');
    },

    _sopStatusHtml(status) {
        const map = {
            draft: ['草稿', '#9e9e9e'],
            ai_analyzing: ['AI分析中', '#2196f3'],
            ai_failed: ['有问题', '#f44336'],
            ai_passed: ['待人工审核', '#ff9800'],
            completed: ['已完成', '#4caf50'],
            rejected: ['已打回', '#e91e63'],
            failed: ['分析失败', '#f44336'],
        };
        const [label, color] = map[status] || [status, '#9e9e9e'];
        return `<span style="color:${color};font-weight:600;">${label}</span>`;
    },

    showCreateSop() {
        this.openModal('新建SOP', `
            <div style="display:flex;flex-direction:column;gap:10px;">
                <div><label style="font-weight:500;">客户名称 <span style="color:#f44336;">*</span></label>
                    <input type="text" id="m-sop-customer" placeholder="客户名称" style="width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:6px;margin-top:4px;"></div>
                <div><label style="font-weight:500;">SOP 描述</label>
                    <textarea id="m-sop-desc" placeholder="SOP 大体描述" rows="4" style="width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:6px;margin-top:4px;"></textarea></div>
            </div>`, async () => {
            const name = document.getElementById('m-sop-customer').value.trim();
            if (!name) return alert('请填写客户名称');
            const resp = await AUTH.authFetch('/api/tools/sop/entries', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ customer_name: name, description: document.getElementById('m-sop-desc').value }),
            });
            if (!resp.ok) return _alertErr(resp, '新建失败');
            this.closeModal();
            this.loadSops();
        });
    },

    showSopUpload(entryId) {
        const entry = (this._sopEntries || []).find(e => e.id === entryId) || {};
        this.openModal(`上传文件 - ${entry.customer_name || entryId}`, `
            <div style="display:flex;flex-direction:column;gap:10px;">
                <div><label style="font-weight:500;">源文件（可多选）<span style="color:#f44336;">*</span></label>
                    <input type="file" id="m-sop-source" accept=".xlsx,.xls,.xlsm" multiple style="margin-top:4px;"></div>
                <div><label style="font-weight:500;">结果文件 <span style="color:#f44336;">*</span></label>
                    <input type="file" id="m-sop-result" accept=".xlsx,.xls,.xlsm" style="margin-top:4px;"></div>
                <div><label style="font-weight:500;">规则文件（可选）</label>
                    <input type="file" id="m-sop-rule" accept=".xlsx,.xls,.xlsm,.docx,.doc,.pdf,.txt,.md" style="margin-top:4px;"></div>
                <div style="font-size:12px;color:#888;">上传后将自动开始 AI 后台分析，请稍候。</div>
            </div>`, () => {
            this._sopUpload(entryId);
        });
    },

    async _sopUpload(entryId) {
        const fd = new FormData();
        const srcFiles = document.getElementById('m-sop-source').files;
        const res = document.getElementById('m-sop-result').files[0];
        const rul = document.getElementById('m-sop-rule').files[0];
        if (!srcFiles || !srcFiles.length) return alert('请选择至少一个源文件');
        if (!res) return alert('请选择结果文件');
        for (const f of srcFiles) fd.append('source_files', f);
        fd.append('result_file', res);
        if (rul) fd.append('rule_file', rul);
        const resp = await AUTH.authFetch(`/api/tools/sop/entries/${entryId}/upload`, { method: 'POST', body: fd });
        if (!resp.ok) return _alertErr(resp, '上传失败');
        const data = await resp.json();
        this.closeModal();
        this._pollSopRound(data.round_id, entryId);
    },

    async _pollSopRound(roundId, entryId) {
        let tries = 0;
        const maxTries = 72;   // 每 2.5s 一次，最长约 3 分钟
        document.getElementById('loading-overlay').style.display = 'flex';
        document.getElementById('loading-text').textContent = 'AI 正在后台分析文件，请稍候...';
        const timer = setInterval(async () => {
            tries++;
            try {
                const resp = await AUTH.authFetch(`/api/tools/sop/rounds/${roundId}`);
                if (!resp.ok) throw new Error('查询分析状态失败');
                const r = await resp.json();
                const done = ['ai_passed', 'ai_failed', 'failed', 'completed', 'rejected'].includes(r.status);
                if (done || tries >= maxTries) {
                    clearInterval(timer);
                    document.getElementById('loading-overlay').style.display = 'none';
                    this.loadSops();
                    if (r.status === 'ai_passed') {
                        alert('AI 分析完成：符合大规则，等待人工审核。');
                    } else if (r.status === 'ai_failed') {
                        const a = r.ai_analysis || {};
                        alert(`AI 判定不符合大规则。\n\n${a.summary || ''}\n\n问题：\n${(a.issues || []).join('\n') || '（见详情）'}\n\n建议：\n${(a.suggestions || []).join('\n') || ''}`);
                    } else if (r.status === 'failed') {
                        alert('AI 分析失败：' + (r.error_message || '未知错误'));
                    } else if (tries >= maxTries) {
                        alert('分析超时，请手动刷新列表查看结果。');
                    }
                }
            } catch (e) {
                clearInterval(timer);
                document.getElementById('loading-overlay').style.display = 'none';
                this.loadSops();
                alert('查询分析状态失败: ' + e.message);
            }
        }, 2500);
    },

    async showSopDetail(entryId) {
        try {
            const resp = await AUTH.authFetch(`/api/tools/sop/entries/${entryId}`);
            if (!resp.ok) return _alertErr(resp, '获取详情失败');
            const entry = await resp.json();
            const canCreate = AUTH.hasPerm('tools.sop.create');
            const canReview = AUTH.hasPerm('tools.sop.review');
            const rounds = (entry.rounds || []).slice().reverse();   // 最新一轮在前
            const roundsHtml = rounds.map(r => {
                const srcNames = (r.source_file_names && r.source_file_names.length) ? r.source_file_names : [r.source_file_name];
                const files = [];
                srcNames.forEach((n, i) => {
                    const label = srcNames.length > 1 ? `源文件${i + 1}` : '源文件';
                    files.push(`<button class="btn btn-sm" ${i ? 'style="margin-left:4px;"' : ''} onclick="Tools.downloadSopFile(${r.id},'source','${String(n).replace(/'/g, "\\'")}',${i})">${label}</button>`);
                });
                files.push(`<button class="btn btn-sm" style="margin-left:4px;" onclick="Tools.downloadSopFile(${r.id},'result','${r.result_file_name.replace(/'/g, "\\'")}')">结果文件</button>`);
                if (r.rule_file_name) files.push(`<button class="btn btn-sm" style="margin-left:4px;" onclick="Tools.downloadSopFile(${r.id},'rule','${r.rule_file_name.replace(/'/g, "\\'")}')">规则文件</button>`);
                const a = r.ai_analysis || {};
                let aiBlock;
                if (r.status === 'ai_analyzing') {
                    aiBlock = '<div style="margin-top:6px;color:#2196f3;">AI 分析中...</div>';
                } else {
                    const verdictHtml = a.passed === undefined
                        ? (r.ai_comment || '-')
                        : `<span style="color:${a.passed ? '#4caf50' : '#f44336'};font-weight:600;">${a.passed ? '符合' : '不符合'}</span>　评分：${a.score ?? '-'}`;
                    aiBlock = `<div style="margin-top:6px;">
                        <div>AI评价：${verdictHtml}</div>
                        ${r.ai_comment ? `<div style="margin-top:4px;">${_escape(r.ai_comment)}</div>` : ''}
                        ${(a.issues && a.issues.length) ? `<div style="margin-top:4px;color:#f44336;">问题：${_escape(a.issues.join('；'))}</div>` : ''}
                        ${(a.suggestions && a.suggestions.length) ? `<div style="margin-top:4px;color:#ff9800;">建议：${_escape(a.suggestions.join('；'))}</div>` : ''}
                    </div>`;
                }
                const reviewBlock = r.review_status
                    ? `<div style="margin-top:6px;">最终评价：<span style="color:${r.review_status === 'completed' ? '#4caf50' : '#e91e63'};font-weight:600;">${r.review_status === 'completed' ? '已完成' : '打回重写'}</span>　${_escape(r.review_comment || '')}　（审核人：${_escape(r.reviewer_name || '-')}　${r.reviewed_at ? new Date(r.reviewed_at).toLocaleString() : ''}）</div>`
                    : '';
                const errBlock = r.error_message ? `<div style="margin-top:6px;color:#f44336;">错误：${_escape(r.error_message)}</div>` : '';
                let reviewBtns = '';
                if (canReview && entry.status === 'ai_passed' && r.id === entry.latest_round_id) {
                    reviewBtns = `<button class="btn btn-sm" style="background:#4caf50;color:#fff;margin-left:4px;" onclick="Tools.reviewSop(${entry.id}, ${r.id}, 'completed')">审核通过</button>
                        <button class="btn btn-sm btn-danger" style="margin-left:4px;" onclick="Tools.showSopRejectModal(${entry.id}, ${r.id}, ${r.round_no})">打回重写</button>`;
                }
                return `<div style="border:1px solid #e0e0e0;border-radius:8px;padding:12px;margin-bottom:12px;${r.id === entry.latest_round_id ? 'border-left:4px solid #2196f3;' : 'opacity:.85;'}">
                    <div style="display:flex;align-items:center;gap:10px;flex-wrap:wrap;">
                        <span style="font-weight:600;">第 ${r.round_no} 轮</span>
                        ${this._sopStatusHtml(r.status)}
                        ${files.join('')}
                        ${reviewBtns}
                    </div>
                    ${aiBlock}
                    ${reviewBlock}
                    ${errBlock}
                </div>`;
            }).join('');
            const body = `
                <div style="margin-bottom:12px;">
                    <div><strong>客户名称：</strong>${_escape(entry.customer_name)}</div>
                    <div style="margin-top:4px;"><strong>SOP 描述：</strong>${_escape(entry.description || '-')}</div>
                    <div style="margin-top:4px;"><strong>当前状态：</strong>${this._sopStatusHtml(entry.status)}</div>
                    <div style="margin-top:4px;"><strong>创建人：</strong>${_escape(entry.created_by_name || '-')}　<strong>创建时间：</strong>${entry.created_at ? new Date(entry.created_at).toLocaleString() : '-'}</div>
                </div>
                <hr style="border:none;border-top:1px solid #eee;margin:12px 0;">
                ${roundsHtml || '<div class="empty-state">暂无上传记录</div>'}
                ${canCreate ? `<div style="margin-top:12px;"><button class="btn btn-primary" onclick="Tools.showSopUpload(${entry.id})">+ 重新上传（开启新一轮）</button></div>` : ''}`;
            this.openModal('SOP 详情', body, null, { wide: true });
        } catch (e) {
            alert('获取详情失败: ' + e.message);
        }
    },

    showSopReview(entryId) {
        this.showSopDetail(entryId);
    },

    async reviewSop(entryId, roundId, verdict) {
        if (verdict === 'rejected') { this.showSopRejectModal(entryId, roundId); return; }
        if (!confirm('确认标记该 SOP 为已完成？')) return;
        await this._submitSopReview(entryId, roundId, verdict, '');
    },

    // 打回重写：专门弹窗填写打回原因/完善意见（不再用浏览器 prompt）
    showSopRejectModal(entryId, roundId, roundNo) {
        const entry = (this._sopEntries || []).find(e => e.id === entryId) || {};
        this.openModal(`打回重写 - ${entry.customer_name || entryId}`, `
            <div style="font-size:13px;color:#555;margin-bottom:10px;">
                第 ${roundNo || '-'} 轮未通过人工审核。请填写打回原因/完善意见，将反馈给上传人重新完善后再上传。
            </div>
            <textarea id="sop-reject-comment" rows="5" placeholder="请输入打回原因/完善意见（必填）"
                style="width:100%;box-sizing:border-box;min-height:110px;"></textarea>
            <div style="margin-top:12px;display:flex;justify-content:flex-end;gap:8px;">
                <button class="btn" onclick="Tools.closeModal()">取消</button>
                <button class="btn btn-danger" onclick="Tools.submitSopReject(${entryId}, ${roundId})">确认打回</button>
            </div>
        `, null);
        const ta = document.getElementById('sop-reject-comment');
        if (ta) setTimeout(() => ta.focus(), 50);
    },

    async submitSopReject(entryId, roundId) {
        const ta = document.getElementById('sop-reject-comment');
        const comment = (ta ? ta.value : '').trim();
        if (!comment) { alert('请填写打回原因/完善意见'); if (ta) ta.focus(); return; }
        await this._submitSopReview(entryId, roundId, 'rejected', comment);
    },

    async _submitSopReview(entryId, roundId, verdict, comment) {
        const resp = await AUTH.authFetch(`/api/tools/sop/entries/${entryId}/review`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ round_id: roundId, verdict, comment }),
        });
        if (!resp.ok) return _alertErr(resp, verdict === 'rejected' ? '打回失败' : '审核失败');
        this.closeModal();
        this.loadSops();
    },

    async deleteSop(entryId) {
        const entry = (this._sopEntries || []).find(e => e.id === entryId);
        if (!confirm(`确认删除 SOP「${entry ? entry.customer_name : entryId}」？该操作不可恢复。`)) return;
        const resp = await AUTH.authFetch(`/api/tools/sop/entries/${entryId}`, { method: 'DELETE' });
        if (!resp.ok) return _alertErr(resp, '删除失败');
        this.loadSops();
    },

    downloadSopFile(roundId, kind, fileName, idx = 0) {
        let url = `/api/tools/sop/rounds/${roundId}/download?kind=${kind}`;
        if (idx) url += `&idx=${idx}`;
        this._fetchAndDownload(url, fileName);
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
