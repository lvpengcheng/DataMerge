/**
 * admin.js - 管理后台页面逻辑
 */

// 缓存数据
let _roles = [];
let _orgs = [];
let _orgsFlatMap = {};  // id -> org
let _modalCallback = null;

/** 安全解析错误响应，防止非 JSON 响应（如 nginx 502/504 纯文本）导致二次报错 */
async function _alertErr(resp, fallback) {
    let msg = fallback;
    try { const j = await resp.json(); msg = j.detail || j.message || fallback; } catch (_) {
        try { msg = await resp.text(); } catch (__) {}
    }
    alert(msg);
}

const Admin = {
    // ==================== 初始化 ====================
    async init() {
        if (!AUTH.requireAuth()) return;
        if (!AUTH.isAdmin()) {
            alert('需要管理员权限');
            window.location.href = '/training';
            return;
        }
        AUTH.renderUserInfo(document.querySelector('header'));
        this.initTabs();
        await this.loadRoles();
        await this.loadOrgs();
        await this.loadUsers();
    },

    initTabs() {
        document.querySelectorAll('.tab-btn').forEach(btn => {
            btn.addEventListener('click', () => {
                document.querySelectorAll('.tab-btn').forEach(b => b.classList.remove('active'));
                document.querySelectorAll('.tab-content').forEach(c => c.classList.remove('active'));
                btn.classList.add('active');
                const tabId = 'tab-' + btn.dataset.tab;
                document.getElementById(tabId).classList.add('active');

                // 切换 tab 时加载数据
                const tab = btn.dataset.tab;
                if (tab === 'users') this.loadUsers();
                else if (tab === 'roles') this.loadRoles().then(() => this.renderRoles());
                else if (tab === 'orgs') this.loadOrgs().then(() => this.renderOrgs());
                else if (tab === 'tenant-auth') this.loadTenantAuth();
                else if (tab === 'ref-data') this.loadRefCategories().then(() => this.loadRefData());
                else if (tab === 'scripts') this.loadScripts();
            });
        });
    },

    // ==================== 用户管理 ====================
    async loadUsers() {
        const resp = await AUTH.authFetch('/api/admin/users');
        if (!resp.ok) return;
        const users = await resp.json();
        this.renderUsers(users);
    },

    renderUsers(users) {
        const tbody = document.querySelector('#users-table tbody');
        if (!users.length) {
            tbody.innerHTML = '<tr><td colspan="7" class="empty-state">暂无用户</td></tr>';
            return;
        }
        tbody.innerHTML = users.map(u => `
            <tr>
                <td>${u.id}</td>
                <td>${u.username}</td>
                <td>${u.display_name || '-'}</td>
                <td>${u.org_name || '-'}</td>
                <td><span class="tag">${u.role_name || '-'}</span></td>
                <td><span class="${u.is_active ? 'status-active' : 'status-inactive'}">${u.is_active ? '启用' : '禁用'}</span></td>
                <td class="actions">
                    <button class="btn btn-sm" onclick="Admin.showEditUser(${u.id})">编辑</button>
                    <button class="btn btn-sm" onclick="Admin.resetPassword(${u.id}, '${u.username}')">重置密码</button>
                    ${u.is_active ? `<button class="btn btn-sm btn-danger" onclick="Admin.disableUser(${u.id}, '${u.username}')">禁用</button>` : ''}
                </td>
            </tr>
        `).join('');
    },

    showCreateUser() {
        this.openModal('新建用户', `
            <div class="form-group"><label>用户名</label><input id="m-username" required></div>
            <div class="form-group"><label>密码</label><input id="m-password" type="password" value="123456"></div>
            <div class="form-group"><label>显示名</label><input id="m-display-name"></div>
            <div class="form-group"><label>邮箱</label><input id="m-email"></div>
            <div class="form-group"><label>电话</label><input id="m-phone"></div>
            <div class="form-group"><label>组织</label><select id="m-org">${this.orgOptions()}</select></div>
            <div class="form-group"><label>角色</label><select id="m-role">${this.roleOptions()}</select></div>
        `, async () => {
            const resp = await AUTH.authFetch('/api/admin/users', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    username: document.getElementById('m-username').value,
                    password: document.getElementById('m-password').value || '123456',
                    display_name: document.getElementById('m-display-name').value,
                    email: document.getElementById('m-email').value,
                    phone: document.getElementById('m-phone').value,
                    org_id: parseInt(document.getElementById('m-org').value) || null,
                    role_id: parseInt(document.getElementById('m-role').value) || null,
                }),
            });
            if (resp.ok) { this.closeModal(); this.loadUsers(); }
            else { await _alertErr(resp, '创建失败'); }
        });
    },

    async showEditUser(id) {
        const resp = await AUTH.authFetch(`/api/admin/users/${id}`);
        if (!resp.ok) return;
        const u = await resp.json();

        this.openModal('编辑用户', `
            <div class="form-group"><label>用户名</label><input value="${u.username}" disabled></div>
            <div class="form-group"><label>显示名</label><input id="m-display-name" value="${u.display_name || ''}"></div>
            <div class="form-group"><label>邮箱</label><input id="m-email" value="${u.email || ''}"></div>
            <div class="form-group"><label>电话</label><input id="m-phone" value="${u.phone || ''}"></div>
            <div class="form-group"><label>组织</label><select id="m-org">${this.orgOptions(u.org_id)}</select></div>
            <div class="form-group"><label>角色</label><select id="m-role">${this.roleOptions(u.role_id)}</select></div>
            <div class="form-group"><label>状态</label><select id="m-active">
                <option value="true" ${u.is_active ? 'selected' : ''}>启用</option>
                <option value="false" ${!u.is_active ? 'selected' : ''}>禁用</option>
            </select></div>
        `, async () => {
            const resp = await AUTH.authFetch(`/api/admin/users/${id}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    display_name: document.getElementById('m-display-name').value,
                    email: document.getElementById('m-email').value,
                    phone: document.getElementById('m-phone').value,
                    org_id: parseInt(document.getElementById('m-org').value) || null,
                    role_id: parseInt(document.getElementById('m-role').value) || null,
                    is_active: document.getElementById('m-active').value === 'true',
                }),
            });
            if (resp.ok) { this.closeModal(); this.loadUsers(); }
            else { await _alertErr(resp, '更新失败'); }
        });
    },

    async resetPassword(id, username) {
        if (!confirm(`确定重置 ${username} 的密码为 123456？`)) return;
        const resp = await AUTH.authFetch(`/api/admin/users/${id}/reset-password`, { method: 'POST' });
        if (resp.ok) alert('密码已重置为 123456');
        else alert('重置失败');
    },

    async disableUser(id, username) {
        if (!confirm(`确定禁用用户 ${username}？`)) return;
        const resp = await AUTH.authFetch(`/api/admin/users/${id}`, { method: 'DELETE' });
        if (resp.ok) this.loadUsers();
        else alert('操作失败');
    },

    // ==================== 角色管理 ====================
    async loadRoles() {
        const resp = await AUTH.authFetch('/api/admin/roles');
        if (!resp.ok) return;
        _roles = await resp.json();
    },

    renderRoles() {
        const tbody = document.querySelector('#roles-table tbody');
        if (!_roles.length) {
            tbody.innerHTML = '<tr><td colspan="6" class="empty-state">暂无角色</td></tr>';
            return;
        }
        tbody.innerHTML = _roles.map(r => `
            <tr>
                <td>${r.id}</td>
                <td>${r.name}</td>
                <td>${r.description || '-'}</td>
                <td>${Object.keys(r.permissions || {}).map(k => `<span class="tag">${k}</span>`).join(' ')}</td>
                <td>${r.is_system ? '是' : '否'}</td>
                <td class="actions">
                    <button class="btn btn-sm" onclick="Admin.showEditRole(${r.id})">编辑</button>
                    ${!r.is_system ? `<button class="btn btn-sm btn-danger" onclick="Admin.deleteRole(${r.id}, '${r.name}')">删除</button>` : ''}
                </td>
            </tr>
        `).join('');
    },

    showCreateRole() {
        this.openModal('新建角色', `
            <div class="form-group"><label>角色名</label><input id="m-role-name" required></div>
            <div class="form-group"><label>描述</label><input id="m-role-desc"></div>
            <div class="form-group"><label>权限（JSON）</label><textarea id="m-role-perms" rows="3">{"can_train": true, "can_compute": true}</textarea></div>
        `, async () => {
            let perms = {};
            try { perms = JSON.parse(document.getElementById('m-role-perms').value); } catch(e) { alert('权限JSON格式错误'); return; }
            const resp = await AUTH.authFetch('/api/admin/roles', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    name: document.getElementById('m-role-name').value,
                    description: document.getElementById('m-role-desc').value,
                    permissions: perms,
                }),
            });
            if (resp.ok) { this.closeModal(); await this.loadRoles(); this.renderRoles(); }
            else { await _alertErr(resp, '创建失败'); }
        });
    },

    async showEditRole(id) {
        const role = _roles.find(r => r.id === id);
        if (!role) return;

        this.openModal('编辑角色', `
            <div class="form-group"><label>角色名</label><input id="m-role-name" value="${role.name}" ${role.is_system ? 'disabled' : ''}></div>
            <div class="form-group"><label>描述</label><input id="m-role-desc" value="${role.description || ''}"></div>
            <div class="form-group"><label>权限（JSON）</label><textarea id="m-role-perms" rows="3">${JSON.stringify(role.permissions || {}, null, 2)}</textarea></div>
        `, async () => {
            let perms = {};
            try { perms = JSON.parse(document.getElementById('m-role-perms').value); } catch(e) { alert('权限JSON格式错误'); return; }
            const resp = await AUTH.authFetch(`/api/admin/roles/${id}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    description: document.getElementById('m-role-desc').value,
                    permissions: perms,
                    ...(role.is_system ? {} : { name: document.getElementById('m-role-name').value }),
                }),
            });
            if (resp.ok) { this.closeModal(); await this.loadRoles(); this.renderRoles(); }
            else { await _alertErr(resp, '更新失败'); }
        });
    },

    async deleteRole(id, name) {
        if (!confirm(`确定删除角色 ${name}？`)) return;
        const resp = await AUTH.authFetch(`/api/admin/roles/${id}`, { method: 'DELETE' });
        if (resp.ok) { await this.loadRoles(); this.renderRoles(); }
        else { await _alertErr(resp, '删除失败'); }
    },

    // ==================== 组织管理 ====================
    async loadOrgs() {
        const resp = await AUTH.authFetch('/api/admin/organizations');
        if (!resp.ok) return;
        _orgs = await resp.json();
        // 构建扁平映射
        _orgsFlatMap = {};
        const flatten = (list) => {
            list.forEach(o => { _orgsFlatMap[o.id] = o; if (o.children) flatten(o.children); });
        };
        flatten(_orgs);
    },

    renderOrgs() {
        const tbody = document.querySelector('#orgs-table tbody');
        const rows = [];
        const renderTree = (list, level = 0) => {
            list.forEach(o => {
                const indent = '&nbsp;'.repeat(level * 4) + (level > 0 ? '└─ ' : '');
                rows.push(`
                    <tr>
                        <td>${o.id}</td>
                        <td>${indent}${o.name}</td>
                        <td>${o.parent_id ? (_orgsFlatMap[o.parent_id]?.name || o.parent_id) : '-'}</td>
                        <td>${o.description || '-'}</td>
                        <td><span class="${o.is_active ? 'status-active' : 'status-inactive'}">${o.is_active ? '启用' : '禁用'}</span></td>
                        <td class="actions">
                            <button class="btn btn-sm" onclick="Admin.showEditOrg(${o.id})">编辑</button>
                            <button class="btn btn-sm btn-danger" onclick="Admin.deleteOrg(${o.id}, '${o.name}')">删除</button>
                        </td>
                    </tr>
                `);
                if (o.children && o.children.length) renderTree(o.children, level + 1);
            });
        };
        renderTree(_orgs);
        tbody.innerHTML = rows.length ? rows.join('') : '<tr><td colspan="6" class="empty-state">暂无组织</td></tr>';
    },

    showCreateOrg() {
        this.openModal('新建组织', `
            <div class="form-group"><label>组织名称</label><input id="m-org-name" required></div>
            <div class="form-group"><label>上级组织</label><select id="m-org-parent">${this.orgOptions(null, true)}</select></div>
            <div class="form-group"><label>描述</label><input id="m-org-desc"></div>
        `, async () => {
            const resp = await AUTH.authFetch('/api/admin/organizations', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    name: document.getElementById('m-org-name').value,
                    parent_id: parseInt(document.getElementById('m-org-parent').value) || null,
                    description: document.getElementById('m-org-desc').value,
                }),
            });
            if (resp.ok) { this.closeModal(); await this.loadOrgs(); this.renderOrgs(); }
            else { await _alertErr(resp, '创建失败'); }
        });
    },

    async showEditOrg(id) {
        const org = _orgsFlatMap[id];
        if (!org) return;

        this.openModal('编辑组织', `
            <div class="form-group"><label>组织名称</label><input id="m-org-name" value="${org.name}"></div>
            <div class="form-group"><label>上级组织</label><select id="m-org-parent">${this.orgOptions(org.parent_id, true)}</select></div>
            <div class="form-group"><label>描述</label><input id="m-org-desc" value="${org.description || ''}"></div>
        `, async () => {
            const resp = await AUTH.authFetch(`/api/admin/organizations/${id}`, {
                method: 'PUT',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    name: document.getElementById('m-org-name').value,
                    parent_id: parseInt(document.getElementById('m-org-parent').value) || null,
                    description: document.getElementById('m-org-desc').value,
                }),
            });
            if (resp.ok) { this.closeModal(); await this.loadOrgs(); this.renderOrgs(); }
            else { await _alertErr(resp, '更新失败'); }
        });
    },

    async deleteOrg(id, name) {
        if (!confirm(`确定删除组织 ${name}？`)) return;
        const resp = await AUTH.authFetch(`/api/admin/organizations/${id}`, { method: 'DELETE' });
        if (resp.ok) { await this.loadOrgs(); this.renderOrgs(); }
        else { await _alertErr(resp, '删除失败'); }
    },

    // ==================== 租户授权 ====================
    async loadTenantAuth() {
        const resp = await AUTH.authFetch('/api/admin/tenant-auth');
        if (!resp.ok) return;
        const auths = await resp.json();
        this.renderTenantAuth(auths);
    },

    renderTenantAuth(auths) {
        const tbody = document.querySelector('#auth-table tbody');
        if (!auths.length) {
            tbody.innerHTML = '<tr><td colspan="6" class="empty-state">暂无授权记录</td></tr>';
            return;
        }
        tbody.innerHTML = auths.map(a => `
            <tr>
                <td>${a.id}</td>
                <td>${a.tenant_id}</td>
                <td>${a.org_name || a.org_id}</td>
                <td><span class="tag ${a.auth_type === 'owner' ? 'tag-owner' : 'tag-shared'}">${a.auth_type === 'owner' ? '所有者' : '共享'}</span></td>
                <td>${a.granted_at || '-'}</td>
                <td class="actions">
                    <button class="btn btn-sm btn-danger" onclick="Admin.revokeAuth(${a.id})">撤销</button>
                </td>
            </tr>
        `).join('');
    },

    async showGrantAuth() {
        // 加载可选的租户列表
        const resp = await AUTH.authFetch('/api/admin/tenant-auth/tenants');
        const tenants = resp.ok ? await resp.json() : [];
        const tenantOpts = tenants.map(t => `<option value="${t}">${t}</option>`).join('');

        this.openModal('新增租户授权', `
            <div class="form-group"><label>租户</label><select id="m-auth-tenant">${tenantOpts}</select></div>
            <div class="form-group"><label>组织</label><select id="m-auth-org">${this.orgOptions()}</select></div>
            <div class="form-group"><label>授权类型</label><select id="m-auth-type">
                <option value="shared">共享</option>
                <option value="owner">所有者</option>
            </select></div>
        `, async () => {
            const resp = await AUTH.authFetch('/api/admin/tenant-auth', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({
                    tenant_id: document.getElementById('m-auth-tenant').value,
                    org_id: parseInt(document.getElementById('m-auth-org').value),
                    auth_type: document.getElementById('m-auth-type').value,
                }),
            });
            if (resp.ok) { this.closeModal(); this.loadTenantAuth(); }
            else { await _alertErr(resp, '授权失败'); }
        });
    },

    async revokeAuth(id) {
        if (!confirm('确定撤销此授权？')) return;
        const resp = await AUTH.authFetch(`/api/admin/tenant-auth/${id}`, { method: 'DELETE' });
        if (resp.ok) this.loadTenantAuth();
        else alert('撤销失败');
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
            _modalCallback = null;  // 立即清空，防止重复触发
            cb();
        }
    },

    // ==================== 选项生成 ====================
    orgOptions(selectedId = null, includeNone = false) {
        let html = includeNone ? '<option value="">无</option>' : '<option value="">请选择</option>';
        Object.values(_orgsFlatMap).forEach(o => {
            html += `<option value="${o.id}" ${o.id === selectedId ? 'selected' : ''}>${o.name}</option>`;
        });
        return html;
    },

    roleOptions(selectedId = null) {
        let html = '<option value="">请选择</option>';
        _roles.forEach(r => {
            html += `<option value="${r.id}" ${r.id === selectedId ? 'selected' : ''}>${r.name} - ${r.description || ''}</option>`;
        });
        return html;
    },

    // ==================== 基础数据管理 ====================
    _refCategories: [],
    _refTenants: [],

    async loadRefCategories() {
        // 并行加载分类和租户列表
        const [catResp, tenantResp] = await Promise.all([
            AUTH.authFetch('/api/assets/reference-categories'),
            AUTH.authFetch('/api/assets/tenants'),
        ]);
        if (catResp.ok) {
            this._refCategories = await catResp.json();
            const sel = document.getElementById('ref-category-filter');
            if (sel) {
                sel.innerHTML = '<option value="">全部分类</option>' +
                    this._refCategories.map(c => `<option value="${c.id}">${c.name}</option>`).join('');
            }
        }
        if (tenantResp.ok) {
            this._refTenants = await tenantResp.json();
            // 填充作用域筛选中的租户选项
            const scopeSel = document.getElementById('ref-scope-filter');
            if (scopeSel) {
                scopeSel.innerHTML = '<option value="">全部</option>' +
                    '<option value="global">仅全局</option>' +
                    this._refTenants.map(t => `<option value="tenant:${t.tenant_id}">租户: ${t.tenant_id}</option>`).join('');
            }
        }
    },

    async loadRefData() {
        const categoryId = document.getElementById('ref-category-filter')?.value || '';
        const scopeVal = document.getElementById('ref-scope-filter')?.value || '';
        let url = '/api/assets?asset_type=reference';
        if (categoryId) url += `&category_id=${categoryId}`;
        // 解析作用域筛选
        if (scopeVal === 'global') {
            url += '&scope=global';
        } else if (scopeVal.startsWith('tenant:')) {
            url += '&scope=tenant&tenant_id=' + encodeURIComponent(scopeVal.replace('tenant:', ''));
        }
        const resp = await AUTH.authFetch(url);
        if (!resp.ok) return;
        const assets = await resp.json();
        const tbody = document.querySelector('#ref-data-table tbody');
        if (!assets.length) {
            tbody.innerHTML = '<tr><td colspan="9" class="empty-state">暂无基础数据</td></tr>';
            return;
        }
        tbody.innerHTML = assets.map(a => `<tr>
            <td>${a.id}</td>
            <td>${a.name}</td>
            <td>${a.category_name || '-'}</td>
            <td>${a.tenant_id ? '<span class="tag">租户: ' + a.tenant_id + '</span>' : '<span class="tag" style="background:#e8f5e9;color:#2e7d32">全局</span>'}</td>
            <td>${a.file_name}</td>
            <td>v${a.version}</td>
            <td>${a.effective_from || '-'}</td>
            <td>${a.is_active ? '<span style="color:green">启用</span>' : '<span style="color:#999">停用</span>'}</td>
            <td>
                <button class="btn btn-sm" onclick="Admin.previewAsset(${a.id})">预览</button>
                <button class="btn btn-sm btn-danger" onclick="Admin.deleteAsset(${a.id})">停用</button>
            </td>
        </tr>`).join('');
    },

    showUploadRefData() {
        const catOptions = this._refCategories.map(c =>
            `<option value="${c.id}">${c.name}</option>`
        ).join('');
        const tenantOptions = this._refTenants.map(t =>
            `<option value="${t.tenant_id}">租户: ${t.tenant_id}</option>`
        ).join('');
        this.openModal('上传基础数据', `
            <div style="display:flex;flex-direction:column;gap:12px;">
                <label>分类：<select id="ref-upload-category" style="padding:6px;border:1px solid #ddd;border-radius:4px;">${catOptions}</select></label>
                <label>名称：<input id="ref-upload-name" type="text" placeholder="如：2025年最低工资标准" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;"></label>
                <label>作用域：
                    <select id="ref-upload-scope" style="padding:6px;border:1px solid #ddd;border-radius:4px;">
                        <option value="">全局（所有租户可用）</option>
                        ${tenantOptions}
                    </select>
                </label>
                <label>生效日期：<input id="ref-upload-from" type="date" style="padding:6px;border:1px solid #ddd;border-radius:4px;"></label>
                <label>失效日期：<input id="ref-upload-to" type="date" style="padding:6px;border:1px solid #ddd;border-radius:4px;"></label>
                <label>文件：<input id="ref-upload-file" type="file" accept=".xlsx,.xls"></label>
            </div>
        `, async () => {
            const fileInput = document.getElementById('ref-upload-file');
            if (!fileInput.files.length) return alert('请选择文件');
            const fd = new FormData();
            fd.append('file', fileInput.files[0]);
            fd.append('asset_type', 'reference');
            fd.append('category_id', document.getElementById('ref-upload-category').value);
            fd.append('name', document.getElementById('ref-upload-name').value || fileInput.files[0].name);
            // 作用域 → tenant_id
            const scopeVal = document.getElementById('ref-upload-scope').value;
            if (scopeVal) fd.append('tenant_id', scopeVal);
            // 日期
            const ef = document.getElementById('ref-upload-from').value;
            if (ef) fd.append('effective_from', ef);
            const et = document.getElementById('ref-upload-to').value;
            if (et) fd.append('effective_to', et);
            const resp = await AUTH.authFetch('/api/assets/upload', {method: 'POST', body: fd});
            if (resp.ok) {
                this.closeModal();
                this.loadRefData();
            } else {
                alert('上传失败: ' + (await resp.text()));
            }
        });
    },

    async previewAsset(assetId) {
        const resp = await AUTH.authFetch(`/api/assets/${assetId}/preview?rows=10`);
        if (!resp.ok) return alert('预览失败');
        const data = await resp.json();
        let html = '';
        for (const [sheet, info] of Object.entries(data)) {
            html += `<h4>${sheet}</h4><div style="overflow-x:auto;"><table class="data-table" style="font-size:12px;">`;
            html += '<thead><tr>' + info.headers.map(h => `<th>${h}</th>`).join('') + '</tr></thead>';
            html += '<tbody>' + info.data.map(row =>
                '<tr>' + row.map(v => `<td>${v}</td>`).join('') + '</tr>'
            ).join('') + '</tbody></table></div>';
        }
        this.openModal('数据预览', html, null);
    },

    async deleteAsset(assetId) {
        if (!confirm('确定停用此数据？')) return;
        await AUTH.authFetch(`/api/assets/${assetId}`, {method: 'DELETE'});
        this.loadRefData();
    },

    // ==================== 脚本管理 ====================
    async loadScripts() {
        const tenant = (document.getElementById('scripts-tenant-filter')?.value || '').trim();
        const includeInactive = document.getElementById('scripts-include-inactive')?.checked || false;
        const params = new URLSearchParams();
        if (tenant) params.set('tenant_id', tenant);
        if (includeInactive) params.set('include_inactive', 'true');
        const resp = await AUTH.authFetch(`/api/admin/scripts?${params.toString()}`);
        if (!resp.ok) {
            alert('加载脚本失败');
            return;
        }
        const data = await resp.json();
        this.renderScripts(data.items || []);
    },

    renderScripts(items) {
        const tbody = document.querySelector('#scripts-table tbody');
        if (!items.length) {
            tbody.innerHTML = '<tr><td colspan="10" class="empty-state">暂无脚本</td></tr>';
            return;
        }
        tbody.innerHTML = items.map(s => {
            const acc = (s.accuracy != null) ? `${(s.accuracy * 100).toFixed(1)}%` : '-';
            const status = s.is_active
                ? '<span class="status-active">启用中</span>'
                : '<span class="status-inactive">已停用</span>';
            const updated = s.updated_at ? s.updated_at.replace('T', ' ').slice(0, 19) : '-';
            const sourceLink = s.source_session_id
                ? `<a href="/training#session=${s.source_session_id}" target="_blank">#${s.source_session_id}</a>`
                : '-';
            const action = s.is_active
                ? `<button class="btn btn-sm btn-danger" onclick="Admin.disableScript(${s.id}, '${(s.name || '').replace(/'/g, "\\'")}')">停用</button>`
                : `<button class="btn btn-sm" onclick="Admin.enableScript(${s.id}, '${(s.name || '').replace(/'/g, "\\'")}')">恢复</button>`;
            return `
                <tr>
                    <td>${s.id}</td>
                    <td>${s.tenant_id}</td>
                    <td>${s.name || '-'}</td>
                    <td>${s.mode || '-'}</td>
                    <td>v${s.version}</td>
                    <td>${acc}</td>
                    <td>${status}</td>
                    <td>${sourceLink}</td>
                    <td>${updated}</td>
                    <td class="actions">${action}</td>
                </tr>
            `;
        }).join('');
    },

    async disableScript(scriptId, name) {
        if (!confirm(`确定停用脚本「${name}」？\n停用后智训和智算将无法选择此脚本。`)) return;
        const resp = await AUTH.authFetch(`/api/admin/scripts/${scriptId}/disable`, { method: 'POST' });
        const data = await resp.json().catch(() => ({}));
        if (!resp.ok) {
            alert(data.detail || '停用失败');
            return;
        }
        alert(data.message || '已停用');
        this.loadScripts();
    },

    async enableScript(scriptId, name) {
        if (!confirm(`确定恢复脚本「${name}」？`)) return;
        const resp = await AUTH.authFetch(`/api/admin/scripts/${scriptId}/enable`, { method: 'POST' });
        const data = await resp.json().catch(() => ({}));
        if (!resp.ok) {
            alert(data.detail || '恢复失败');
            return;
        }
        alert(data.message || '已恢复');
        this.loadScripts();
    },
};

document.addEventListener('DOMContentLoaded', () => {
    Admin.init();
});
