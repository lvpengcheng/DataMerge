/** 独立 SOP 审核平台；所有数据范围由服务端再次校验。 */
const SOP = {
    entries: [],
    page: 1,
    pageSize: 10,
    escape(value) {
        return String(value ?? '').replace(/[&<>"']/g, ch => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[ch]));
    },
    error(message) { document.getElementById('sop-error').textContent = message || ''; },
    async request(url, options) {
        const resp = await AUTH.authFetch(url, options);
        if (!resp.ok) {
            const data = await resp.json().catch(() => ({}));
            throw new Error(data.detail || `请求失败 (${resp.status})`);
        }
        return resp;
    },
    filters() {
        const p = new URLSearchParams();
        for (const [id, key] of [['sop-keyword','keyword'],['sop-status','status'],['sop-month','batch_month'],['sop-salesperson','salesperson']]) {
            if (key === 'salesperson' && !this.canSearchSalesperson()) continue;
            const value = document.getElementById(id).value.trim();
            if (value) p.set(key, value);
        }
        return p;
    },
    canSearchSalesperson() { return AUTH.isAdmin() || AUTH.hasPerm('sop.review') || AUTH.hasPerm('sop.report'); },
    validateExcelFiles(form) {
        for (const input of form.querySelectorAll('input[type=file]')) {
            for (const file of input.files) {
                if (!/\.(xlsx|xls)$/i.test(file.name)) throw new Error('源文件、目标文件和规则文件仅支持 xlsx、xls 格式');
            }
        }
    },
    async load() {
        this.error('');
        try {
            const data = await (await this.request('/api/tools/sop/entries?' + this.filters())).json();
            this.entries = data.items || [];
            this.page = 1;
            const counts = data.counts || {};
            for (const [id, key] of [['sop-total','total'],['sop-pending-upload','pending_upload'],['sop-pending-review','pending_review'],['sop-rejected','rejected'],['sop-completed','completed']]) {
                document.getElementById(id).textContent = counts[key] || 0;
            }
            document.querySelectorAll('[data-sop-status]').forEach(card => {
                card.setAttribute('aria-pressed', String(card.dataset.sopStatus === document.getElementById('sop-status').value));
            });
            this.render();
        } catch (e) { this.error(e.message); }
    },
    status(value) {
        return ({pending_upload:'待上传',draft:'待上传',awaiting_files:'待上传',pending_review:'待审核',ai_queued:'待审核',ai_analyzing:'待审核',ai_failed:'驳回',
                 ai_passed:'待审核',rejected:'驳回',completed:'通过',failed:'驳回'})[value] || value;
    },
    render() {
        const body = document.querySelector('#sop-table tbody');
        const total = this.entries.length;
        const pages = Math.max(1, Math.ceil(total / this.pageSize));
        this.page = Math.max(1, Math.min(this.page, pages));
        const start = (this.page - 1) * this.pageSize;
        document.getElementById('sop-page-info').textContent = `共 ${total} 条，每页 10 条，第 ${this.page} / ${pages} 页`;
        document.getElementById('sop-page-prev').disabled = this.page <= 1;
        document.getElementById('sop-page-next').disabled = this.page >= pages;
        if (!this.entries.length) { body.innerHTML = '<tr><td colspan="9">暂无记录</td></tr>'; return; }
        body.innerHTML = this.entries.slice(start, start + this.pageSize).map(e => {
            const actions = [`<button class="btn" data-action="detail" data-id="${e.id}">详情</button>`];
            if (AUTH.hasPerm('sop.upload') && ['draft','awaiting_files','rejected','ai_failed','failed'].includes(e.status))
                actions.push(`<button class="btn" data-action="upload" data-id="${e.id}">上传</button>`);
            if (AUTH.hasPerm('sop.upload'))
                actions.push(`<button class="btn" data-action="supplementary" data-id="${e.id}">补充上传</button>`);
            if (AUTH.hasPerm('sop.review') && e.status === 'ai_passed')
                actions.push(`<button class="btn" data-action="detail" data-id="${e.id}">审核</button>`);
            if (AUTH.isAdmin()) {
                if (['ai_failed','rejected','ai_passed','completed'].includes(e.status))
                    actions.push(`<button class="btn" data-action="resend" data-id="${e.id}">补发通知</button>`);
                actions.push(`<button class="btn" data-action="delete" data-id="${e.id}">删除</button>`);
            }
            return `<tr><td>${this.escape(e.batch_month || '-')}</td><td>${this.escape(e.customer_name)}</td><td>${this.escape(e.project_number || '-')}</td>
                <td>${this.escape(e.salesperson_name || '-')}</td><td>${this.escape(this.status(e.business_status || e.status))}</td>
                <td>${e.round_no ? 'v' + e.round_no : '-'}</td><td><span class="sop-ai-summary" title="${this.escape(e.ai_comment || '-')}">${this.escape(e.ai_comment || '-')}</span></td>
                <td>${e.updated_at ? this.escape(new Date(e.updated_at).toLocaleString()) : '-'}</td>
                <td><div class="sop-row-actions">${actions.join('')}</div></td></tr>`;
        }).join('');
    },
    dialog(html, kind = 'default') {
        document.getElementById('sop-dialog-body').innerHTML = html;
        const dialog = document.getElementById('sop-dialog');
        dialog.classList.toggle('sop-detail-dialog', kind === 'detail');
        dialog.showModal();
    },
    close() { document.getElementById('sop-dialog').close(); },
    async create() {
        if (!AUTH.isAdmin()) { this.error('仅系统管理员可以新建 SOP。'); return; }
        let choices;
        try { choices = await (await this.request('/api/tools/sop/create-options')).json(); }
        catch (e) { this.error(e.message); return; }
        this.dialog(`<h2>新建 SOP</h2><form id="sop-create-form" class="sop-form">
            <label class="sop-full">客户（租户）<select id="sop-create-customer" required><option value="">请选择客户</option>${choices.customers.map(name => `<option value="${this.escape(name)}">${this.escape(name)}</option>`).join('')}</select></label>
            <label class="sop-full">项目编号 <span class="sop-required">*</span><input id="sop-create-project-number" maxlength="100" required placeholder="请输入该租户内唯一的项目编号"></label>
            <label>业务员<select id="sop-create-assignee" required><option value="">请选择业务员</option>${choices.users.map(user => `<option value="${user.id}">${this.escape(user.display_name || user.username)}（${this.escape(user.username)}）</option>`).join('')}</select></label>
            <label>业务员邮箱<input id="sop-create-email" type="email" readonly placeholder="选择业务员后自动带出"></label>
            <label class="sop-full">集约月份批次<input id="sop-create-month" type="month" required></label>
            <label class="sop-full">SOP 描述<textarea id="sop-create-description" rows="5" placeholder="请输入 SOP 描述（选填）"></textarea></label>
            ${!choices.customers.length ? '<p class="sop-muted sop-full">暂无可选客户，请先导入客户清单。</p>' : ''}
            ${!choices.users.length ? '<p class="sop-muted sop-full">暂无可选业务员，请在用户管理中启用用户、填写邮箱并授予 SOP 上传权限。</p>' : ''}
            <p class="sop-error sop-full" id="sop-create-error" role="alert"></p>
            <div class="sop-actions sop-review-actions sop-full"><button type="button" class="btn" id="sop-create-cancel">取消</button><button type="submit" class="btn btn-primary">确认新建</button></div>
        </form>`);
        document.getElementById('sop-create-cancel').onclick = () => this.close();
        document.getElementById('sop-create-assignee').onchange = event => {
            const user = choices.users.find(user => user.id === Number(event.target.value));
            document.getElementById('sop-create-email').value = user ? user.email : '';
        };
        document.getElementById('sop-create-form').onsubmit = async event => {
            event.preventDefault();
            const customer = document.getElementById('sop-create-customer');
            const name = customer.value.trim();
            const error = document.getElementById('sop-create-error');
            error.textContent = '';
            if (!name) { error.textContent = '请填写客户名称'; customer.focus(); return; }
            const button = event.currentTarget.querySelector('button[type=submit]');
            button.disabled = true;
            button.textContent = '创建中…';
            try {
                await this.request('/api/tools/sop/entries', {method:'POST',
                    headers:{'Content-Type':'application/json'},
                    body:JSON.stringify({customer_name:name,
                        project_number:document.getElementById('sop-create-project-number').value.trim(),
                        assignee_user_id:Number(document.getElementById('sop-create-assignee').value),
                        batch_month:document.getElementById('sop-create-month').value.replace('-', ''),
                        description:document.getElementById('sop-create-description').value.trim()})});
                this.close();
                // 清除筛选，避免新建的待上传记录被之前的状态或月份筛选隐藏。
                for (const id of ['sop-keyword', 'sop-month', 'sop-status', 'sop-salesperson']) document.getElementById(id).value = '';
                await this.load();
            } catch (e) { error.textContent = e.message; }
            finally { button.disabled = false; button.textContent = '确认新建'; }
        };
    },
    async upload(entryId) {
        try {
            const list = (await (await this.request('/api/tools/sop/entries')).json()).items || [];
            const options = list.filter(e => ['draft','awaiting_files','rejected','ai_failed','failed'].includes(e.status));
            if (!options.length) throw new Error('没有可上传的客户，请联系管理员分配清单');
            this.dialog(`<h2>上传 SOP 文件</h2><form id="sop-upload-form" class="sop-form">
                <div class="sop-full sop-customer-field">
                    <label for="sop-customer-search">选择客户</label>
                    <div class="sop-customer-combobox" id="sop-customer-combobox">
                        <input id="sop-customer-search" type="text" role="combobox" aria-autocomplete="list" aria-controls="sop-customer-options" aria-expanded="false" autocomplete="off" required placeholder="输入客户名称搜索，或直接选择">
                        <button type="button" id="sop-customer-toggle" aria-label="展开客户列表" aria-controls="sop-customer-options">▾</button>
                        <div id="sop-customer-options" class="sop-customer-options" role="listbox" aria-label="客户列表" hidden></div>
                    </div>
                    <input id="sop-customer-select" type="hidden">
                </div>
                <label><span>源文件（可多选） <span class="sop-required" aria-hidden="true">*</span></span><input type="file" name="source_files" multiple accept=".xlsx,.xls" required></label>
                <label><span>目标文件（结果表） <span class="sop-required" aria-hidden="true">*</span></span><input type="file" name="result_file" accept=".xlsx,.xls" required></label>
                <label><span>规则文件 <span class="sop-required" aria-hidden="true">*</span></span><input type="file" name="rule_file" accept=".xlsx,.xls" required></label>
                <p class="sop-muted sop-full">源文件和目标文件仅支持 xlsx、xls 格式；规则文件仅支持 xlsx、xls 格式。</p>
                <button class="btn btn-primary" type="submit">提交审核</button><p class="sop-error" id="sop-upload-error"></p>
            </form>`);
            const select = document.getElementById('sop-customer-select');
            const search = document.getElementById('sop-customer-search');
            const dropdown = document.getElementById('sop-customer-options');
            const field = document.getElementById('sop-customer-combobox');
            let visible = [], active = -1;
            const label = e => `${e.customer_name} (${e.batch_month || '-'})`;
            const closeList = () => {
                dropdown.hidden = true;
                search.setAttribute('aria-expanded', 'false');
                search.removeAttribute('aria-activedescendant');
            };
            const choose = e => {
                select.value = String(e.id);
                search.value = label(e);
                search.setCustomValidity('');
                closeList();
            };
            const openList = (keyword = '') => {
                visible = options.filter(e => label(e).toLowerCase().includes(keyword.trim().toLowerCase()));
                active = -1;
                dropdown.innerHTML = visible.map((e, index) => `<button type="button" role="option" tabindex="-1" id="sop-customer-option-${index}" data-index="${index}" aria-selected="${String(e.id) === select.value}">${this.escape(label(e))}</button>`).join('') || '<p class="sop-muted">没有匹配的客户</p>';
                dropdown.hidden = false;
                search.setAttribute('aria-expanded', 'true');
                search.removeAttribute('aria-activedescendant');
            };
            search.oninput = () => {
                select.value = '';
                search.setCustomValidity('请从下拉列表中选择客户');
                openList(search.value);
            };
            search.onclick = search.onfocus = () => openList(select.value ? '' : search.value);
            document.getElementById('sop-customer-toggle').onclick = () => {
                const wasOpen = !dropdown.hidden;
                search.focus();
                if (wasOpen) closeList(); else openList(select.value ? '' : search.value);
            };
            dropdown.onmousedown = event => event.preventDefault();
            dropdown.onclick = event => {
                const option = event.target.closest('button[data-index]');
                if (option) choose(visible[Number(option.dataset.index)]);
            };
            field.onfocusout = event => { if (!field.contains(event.relatedTarget)) closeList(); };
            search.onkeydown = event => {
                if (event.key === 'Escape' && !dropdown.hidden) {
                    event.preventDefault(); event.stopPropagation(); closeList();
                } else if (event.key === 'ArrowDown' || event.key === 'ArrowUp') {
                    event.preventDefault();
                    if (dropdown.hidden) openList(select.value ? '' : search.value);
                    if (!visible.length) return;
                    active = active < 0 ? (event.key === 'ArrowDown' ? 0 : visible.length - 1)
                        : (active + (event.key === 'ArrowDown' ? 1 : -1) + visible.length) % visible.length;
                    dropdown.querySelectorAll('[role="option"]').forEach((option, index) => option.classList.toggle('active', index === active));
                    const option = dropdown.children[active];
                    search.setAttribute('aria-activedescendant', option.id);
                    option.scrollIntoView({block:'nearest'});
                } else if (event.key === 'Enter' && !dropdown.hidden) {
                    event.preventDefault();
                    if (active >= 0 || visible.length === 1) choose(visible[active >= 0 ? active : 0]);
                }
            };
            const initial = options.find(e => e.id === Number(entryId));
            if (initial) choose(initial);
            document.getElementById('sop-upload-form').onsubmit = async event => {
                event.preventDefault();
                const id = Number(select.value);
                if (!options.some(e => e.id === id)) {
                    search.setCustomValidity('请从下拉列表中选择客户');
                    search.reportValidity();
                    return;
                }
                const form = event.currentTarget;
                const button = form.querySelector('button[type=submit]');
                button.disabled = true;
                const originalText = button.textContent;
                button.textContent = '正在提交并处理初审，请稍候…';
                try {
                    this.validateExcelFiles(form);
                    const data = await (await this.request(`/api/tools/sop/entries/${id}/upload`,
                        {method:'POST', body:new FormData(form)})).json();
                    this.close();
                    await this.load();
                    const feedback = data.status === 'ai_passed' ? '初审通过，等待复审。'
                        : data.status === 'ai_failed' ? '初审未通过。'
                        : data.status === 'failed' ? '初审执行失败。'
                        : data.status === 'ai_queued' ? '已进入定时初审队列。' : '初审处理中。';
                    alert(`版本 v${data.round_no}：${feedback}\n${data.error_message || data.ai_comment || ''}`);
                } catch (e) { document.getElementById('sop-upload-error').textContent = e.message; }
                finally { button.disabled = false; button.textContent = originalText; }
            };
        } catch (e) { this.error(e.message); }
    },
    async supplementary(id) {
        const entry = this.entries.find(entry => entry.id === Number(id));
        this.dialog(`<h2>补充上传：${this.escape(entry?.customer_name || '')}</h2>
            <p class="sop-muted">可多次上传，同名文件替换，不同名文件追加。</p>
            <form id="sop-supplementary-form" class="sop-form">
                <label class="sop-full">补充资料（可多选）<input type="file" name="files" multiple required accept=".xlsx,.xls,.txt,.docx,.doc,.zip,.rar,.7z"></label>
                <p class="sop-muted sop-full">支持格式：xlsx、xls、txt、docx、doc，以及 zip、rar、7z 压缩包。</p>
                <p class="sop-error sop-full" id="sop-supplementary-error" role="alert"></p>
                <button class="btn btn-primary" type="submit">提交补充资料</button>
            </form>`);
        document.getElementById('sop-supplementary-form').onsubmit = async event => {
            event.preventDefault();
            const form = event.currentTarget;
            const button = form.querySelector('button[type=submit]');
            button.disabled = true;
            try {
                for (const file of form.querySelector('input[type=file]').files) {
                    if (!/\.(xlsx|xls|txt|docx|doc|zip|rar|7z)$/i.test(file.name)) throw new Error('请选择支持的补充资料格式');
                }
                await this.request(`/api/tools/sop/entries/${id}/supplementary`, {method:'POST', body:new FormData(form)});
                this.close();
                await this.load();
                await this.detail(id);
            } catch (e) { document.getElementById('sop-supplementary-error').textContent = e.message; }
            finally { button.disabled = false; }
        };
    },
    async detail(id) {
        try {
            const e = await (await this.request(`/api/tools/sop/entries/${id}`)).json();
            const rounds = (e.rounds || []).map(r => {
                const sourceNames = r.source_file_names?.length ? r.source_file_names : (r.source_file_name ? [r.source_file_name] : []);
                const fileButton = (name, kind, index = 0) => `<button class="btn" data-file="${r.id}" data-kind="${kind}" data-index="${index}" data-name="${this.escape(name)}">${this.escape(name)}</button>`;
                const fileGroup = (label, buttons) => `<section class="sop-file-group"><h3>${label}</h3><div class="sop-actions sop-file-actions">${buttons || '<span class="sop-muted">暂无文件</span>'}</div></section>`;
                const files = fileGroup('源文件', sourceNames.map((name, index) => fileButton(name, 'source', index)).join(''))
                    + fileGroup('结果表', r.result_file_name ? fileButton(r.result_file_name, 'result') : '')
                    + fileGroup('规则文件', r.rule_file_name ? fileButton(r.rule_file_name, 'rule') : '');
                const packageButton = `<button class="btn" data-file="${r.id}" data-kind="package" data-name="sop_${e.id}_v${r.round_no}.zip">下载压缩包</button>`;
                const review = AUTH.hasPerm('sop.review') && e.status === 'ai_passed' && e.latest_round_id === r.id
                    ? `<div class="sop-actions sop-review-actions"><button class="btn btn-primary" data-review="completed" data-round="${r.id}" data-entry="${e.id}">通过</button><button class="btn" data-review="rejected" data-round="${r.id}" data-entry="${e.id}">驳回</button></div>` : '';
                return `<div class="sop-round"><strong>v${r.round_no} · ${this.escape(this.status(r.status))}</strong>
                    <p class="sop-comment">${this.escape(r.ai_comment || r.error_message || '')}</p><div class="sop-actions sop-file-actions">${packageButton}</div><div class="sop-file-groups">${files}</div>
                    ${r.review_comment ? `<p class="sop-comment">审核意见：${this.escape(r.review_comment)}</p>` : ''}${review}</div>`;
            }).join('');
            const supplementary = (e.supplementary_files || []).map((file, index) =>
                `<button class="btn" data-supplementary="${e.id}" data-index="${index}" data-name="${this.escape(file.name)}">${this.escape(file.name)}</button>`).join('');
            this.dialog(`<div class="sop-detail-heading"><h2>${this.escape(e.customer_name)}</h2>
                <p>业务员：${this.escape(e.salesperson_name || '-')}　集约月份批次：${this.escape(e.batch_month || '-')}</p>
                <p>项目编号：${this.escape(e.project_number || '-')}　状态：${this.escape(this.status(e.status))}</p></div>
                <div class="sop-detail-content">${supplementary ? `<section class="sop-file-group"><h3>补充资料</h3><div class="sop-actions">${supplementary}</div></section>` : ''}
                ${rounds || '<p>尚未上传</p>'}</div>`, 'detail');
        } catch (e) { this.error(e.message); }
    },
    async review(entryId, roundId, verdict) {
        if (verdict === 'rejected') {
            const dialog = document.getElementById('sop-review-dialog');
            const form = document.getElementById('sop-review-form');
            document.getElementById('sop-review-comment').value = '';
            form.onsubmit = async event => {
                event.preventDefault();
                const comment = document.getElementById('sop-review-comment').value.trim();
                if (!comment) return;
                dialog.close();
                await this.submitReview(entryId, roundId, verdict, comment);
            };
            dialog.showModal();
            return;
        }
        if (verdict === 'completed' && !confirm('确认审核通过？')) return;
        await this.submitReview(entryId, roundId, verdict, '');
    },
    async submitReview(entryId, roundId, verdict, comment) {
        try {
            await this.request(`/api/tools/sop/entries/${entryId}/review`, {method:'POST',
                headers:{'Content-Type':'application/json'}, body:JSON.stringify({round_id:roundId,verdict,comment:comment || ''})});
            this.close(); await this.load();
        } catch (e) { this.error(e.message); }
    },
    async download(url, filename, onError) {
        try {
            const blob = await (await this.request(url)).blob();
            const link = document.createElement('a');
            const objectUrl = URL.createObjectURL(blob);
            link.href = objectUrl; link.download = filename || 'download';
            link.style.display = 'none';
            document.body.appendChild(link);
            link.click();
            link.remove();
            setTimeout(() => URL.revokeObjectURL(objectUrl), 60000);
            return true;
        } catch (e) {
            if (onError) onError(e.message);
            else this.error(e.message);
            return false;
        }
    },
    async init() {
        if (!AUTH.requireAuth()) return;
        if (!AUTH.hasPerm('menu.sop')) { this.error('账号未开通 SOP 入口权限，请联系管理员。'); return; }
        AUTH.renderUserInfo(document.querySelector('header'));
        document.getElementById('sop-salesperson').hidden = !this.canSearchSalesperson();
        document.getElementById('sop-salesperson').onkeydown = event => { if (event.key === 'Enter') this.load(); };
        document.getElementById('sop-search').onclick = () => this.load();
        document.getElementById('sop-status').onchange = () => this.load();
        document.querySelectorAll('[data-sop-status]').forEach(card => {
            card.onclick = () => {
                document.getElementById('sop-status').value = card.dataset.sopStatus;
                this.load();
                document.getElementById('sop-table').scrollIntoView({block:'nearest', behavior:'smooth'});
            };
        });
        document.getElementById('sop-page-prev').onclick = () => { this.page -= 1; this.render(); };
        document.getElementById('sop-page-next').onclick = () => { this.page += 1; this.render(); };
        document.getElementById('sop-create-open').onclick = () => this.create();
        document.getElementById('sop-upload-open').onclick = () => this.upload();
        document.getElementById('sop-close').onclick = () => this.close();
        document.getElementById('sop-review-cancel').onclick = () => document.getElementById('sop-review-dialog').close();
        document.getElementById('sop-export').onclick = () => {
            const params = this.filters();
            this.download('/api/tools/sop/report?' + params, 'sop_report.xlsx');
        };
        document.getElementById('sop-table').onclick = event => {
            const button = event.target.closest('button[data-action]');
            if (!button) return;
            if (button.dataset.action === 'upload') this.upload(button.dataset.id);
            else if (button.dataset.action === 'supplementary') this.supplementary(button.dataset.id);
            else if (button.dataset.action === 'resend') {
                this.request(`/api/tools/sop/entries/${button.dataset.id}/resend-notice`, {method:'POST'})
                    .then(() => alert('通知已加入发送队列')).catch(e => this.error(e.message));
            } else if (button.dataset.action === 'delete') {
                if (!confirm('确认删除该 SOP 及其全部文件？')) return;
                this.request(`/api/tools/sop/entries/${button.dataset.id}`, {method:'DELETE'})
                    .then(() => this.load()).catch(e => this.error(e.message));
            } else this.detail(button.dataset.id);
        };
        document.getElementById('sop-dialog').onclick = event => {
            const supplementary = event.target.closest('button[data-supplementary]');
            if (supplementary) {
                this.download(`/api/tools/sop/entries/${supplementary.dataset.supplementary}/supplementary/${supplementary.dataset.index}`, supplementary.dataset.name);
                return;
            }
            const file = event.target.closest('button[data-file]');
            if (file) {
                const url = `/api/tools/sop/rounds/${file.dataset.file}/${file.dataset.kind === 'package' ? 'package' : 'download?kind=' + file.dataset.kind + '&idx=' + (file.dataset.index || 0)}`;
                this.download(url, file.dataset.name); return;
            }
            const review = event.target.closest('button[data-review]');
            if (review) this.review(review.dataset.entry, review.dataset.round, review.dataset.review);
        };
        this.load();
    },
};
document.addEventListener('DOMContentLoaded', () => {
    if (document.getElementById('sop-table')) SOP.init();
});
