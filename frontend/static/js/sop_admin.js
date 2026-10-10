/** SOP 系统管理；仅管理员加载管理数据。 */
const SOP_ADMIN = {
    escape: value => SOP.escape(value),
    error: message => SOP.error(message),
    request: (url, options) => SOP.request(url, options),
    download: (url, filename, onError) => SOP.download(url, filename, onError),
    async aiSchedule() {
        const error = document.getElementById('sop-ai-schedule-error');
        try {
            const data = await (await this.request('/api/tools/sop/ai-schedule')).json();
            document.getElementById('sop-ai-mode').value = data.mode;
            document.getElementById('sop-ai-hours').value = data.interval_hours;
            document.getElementById('sop-ai-hours').disabled = data.mode !== 'scheduled';
            const last = data.last_run_at ? new Date(data.last_run_at).toLocaleString() : '尚未扫描';
            const next = data.mode === 'scheduled' ? `下次扫描：${new Date(data.next_run_at).toLocaleString()}` : '待初审记录将尽快处理';
            document.getElementById('sop-ai-schedule-status').textContent = `已保存：${data.mode === 'scheduled' ? '每 ' + data.interval_hours + ' 小时审核一次' : '立即审核'}。待初审 ${data.pending_count} 条，审核中 ${data.running_count} 条。上次扫描：${last}；${next}。`;
            error.textContent = '';
        } catch (e) { error.textContent = e.message; }
    },
    currentRules(items) {
        // 兼容更新期间仍返回历史版本的接口，每个适用范围只展示最新版。
        const latest = new Map();
        for (const rule of items) {
            const key = rule.scope === 'global' ? 'global' : `customer:${rule.customer_name || ''}`;
            if (!latest.has(key) || Number(rule.id) > Number(latest.get(key).id)) latest.set(key, rule);
        }
        return [...latest.values()];
    },
    async rules() {
        try {
            const data = await (await this.request('/api/tools/sop/rules')).json();
            document.getElementById('sop-rules-list').innerHTML = this.currentRules(data.items || []).map(r =>
                `<div class="sop-rule-row"><div><strong>${this.escape(r.scope === 'global' ? '全局规则' : r.customer_name)}</strong><p>${this.escape(r.name || r.file_name)}</p><span class="sop-muted">${this.escape(r.file_name)} · ${r.is_active ? '已启用' : '已停用'}</span></div><div class="sop-actions">
                <button class="btn" data-rule-toggle="${r.id}">${r.is_active ? '停用' : '启用'}</button>
                <button class="btn" data-rule-download="${r.id}" data-name="${this.escape(r.file_name)}">下载</button>
                <button class="btn" data-rule-delete="${r.id}">删除</button><span class="sop-rule-download-status" role="status"></span></div></div>`).join('') || '<p class="sop-muted">暂无当前规则，请上传规则文件。</p>';
        } catch (e) { this.error(e.message); }
    },
    async reviewerConfig() {
        try {
            const data = await (await this.request('/api/tools/sop/reviewer-config')).json();
            const select = document.getElementById('sop-reviewer-user');
            const available = data.users.filter(user => user.selectable);
            const selected = new Set((data.user_ids || (data.user_id ? [data.user_id] : [])).map(Number));
            select.innerHTML = data.users.map(user =>
                `<option value="${user.id}"${user.selectable ? '' : ' disabled'}${selected.has(user.id) && user.selectable ? ' selected' : ''}>${this.escape(user.display_name || user.username)}（${this.escape(user.username)} · ${this.escape(user.email || '无邮箱')}）${user.selectable ? '' : ' — ' + this.escape(user.disabled_reason)}</option>`).join('');
            document.getElementById('sop-reviewer-save').disabled = !available.length;
            document.getElementById('sop-reviewer-status').textContent = data.reviewer_valid
                ? `当前复审通知接收人：${data.reviewer_name}`
                : data.reviewer_name ? `已配置的复审人中存在失效账号，请检查并重新保存：${data.reviewer_name}`
                : available.length ? '用户列表来自用户管理。可选择多位复审人，AI 通过后的通知将发送到所有有效复审人的邮箱。'
                : data.users.length ? '用户表中暂无符合条件的复审人。请根据下拉框提示，在用户管理中完善邮箱、启用状态和 SOP 权限。'
                : '用户表中暂无用户，请先在用户管理中创建用户。';
        } catch (e) { this.error(e.message); }
    },
    async setup() {
        try {
            const data = await (await this.request('/api/tools/sop/setup-status')).json();
            const missing = [];
            if (!data.public_url_configured) missing.push('SOP_PUBLIC_BASE_URL');
            if (!data.smtp_configured) missing.push('SMTP 发件账户');
            if (!data.reviewer_configured) missing.push('复审人选择');
            else {
                if (!data.reviewer_email_configured) missing.push('复审人的邮箱');
                if (!data.reviewer_can_review) missing.push('复审人的启用状态及 SOP 审核权限');
            }
            document.getElementById('sop-setup').textContent = missing.length
                ? `邮件通知配置待完善：${missing.join('、')}` : '邮件通知配置已就绪';
        } catch (e) { this.error(e.message); }
    },
    async init() {
        if (!AUTH.requireAuth()) return;
        AUTH.renderUserInfo(document.querySelector('header'));
        if (!AUTH.isAdmin()) { this.error('仅系统管理员可以访问 SOP 系统管理。'); return; }
        document.getElementById('sop-admin').style.display = '';
        document.getElementById('sop-ai-mode').onchange = event => {
            document.getElementById('sop-ai-hours').disabled = event.target.value !== 'scheduled';
        };
        document.getElementById('sop-ai-refresh').onclick = () => this.aiSchedule();
        document.getElementById('sop-ai-schedule-form').onsubmit = async event => {
            event.preventDefault();
            const button = document.getElementById('sop-ai-save');
            button.disabled = true;
            try {
                await this.request('/api/tools/sop/ai-schedule', {method:'PUT', headers:{'Content-Type':'application/json'},
                    body:JSON.stringify({mode:document.getElementById('sop-ai-mode').value,
                        interval_hours:Number(document.getElementById('sop-ai-hours').value)})});
                await this.aiSchedule();
            } catch (e) { document.getElementById('sop-ai-schedule-error').textContent = e.message; }
            finally { button.disabled = false; }
        };
        document.getElementById('sop-reviewer-form').onsubmit = async event => {
            event.preventDefault();
            try {
                await this.request('/api/tools/sop/reviewer-config', {method:'PUT',
                    headers:{'Content-Type':'application/json'},
                    body:JSON.stringify({user_ids:Array.from(document.getElementById('sop-reviewer-user').selectedOptions, option => Number(option.value))})});
                await this.reviewerConfig();
                await this.setup();
            } catch (e) { this.error(e.message); }
        };
        document.getElementById('sop-rule-form').onsubmit = async event => {
            event.preventDefault();
            const form = event.currentTarget;
            try { await this.request('/api/tools/sop/rules', {method:'POST',body:new FormData(form)});
                form.reset(); await this.rules(); } catch (e) { this.error(e.message); }
        };
        document.getElementById('sop-rules-list').onclick = async event => {
            const download = event.target.closest('button[data-rule-download]');
            if (download) {
                if (download.disabled) return;
                const status = download.closest('.sop-rule-row').querySelector('.sop-rule-download-status');
                download.disabled = true;
                download.textContent = '下载中…';
                status.textContent = '';
                status.classList.remove('sop-error');
                try {
                    const ok = await this.download(`/api/tools/sop/rules/${download.dataset.ruleDownload}/download`, download.dataset.name, message => {
                        status.classList.add('sop-error');
                        status.textContent = `下载失败：${message}`;
                    });
                    if (ok) status.textContent = '已交给浏览器下载';
                } finally {
                    download.disabled = false;
                    download.textContent = '下载';
                }
                return;
            }
            const remove = event.target.closest('button[data-rule-delete]');
            if (remove) {
                if (!confirm('确认删除该规则文件？')) return;
                try { await this.request(`/api/tools/sop/rules/${remove.dataset.ruleDelete}`, {method:'DELETE'}); this.rules(); }
                catch (e) { this.error(e.message); }
                return;
            }
            const button = event.target.closest('button[data-rule-toggle]');
            if (!button) return;
            try { await this.request(`/api/tools/sop/rules/${button.dataset.ruleToggle}/toggle`, {method:'POST'}); this.rules(); }
            catch (e) { this.error(e.message); }
        };
        await Promise.all([this.rules(), this.setup(), this.reviewerConfig(), this.aiSchedule()]);
    },
};
document.addEventListener('DOMContentLoaded', () => SOP_ADMIN.init());
