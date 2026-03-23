/**
 * admin.js - 管理后台页面逻辑
 */

// 缓存数据
let _roles = [];
let _orgs = [];
let _orgsFlatMap = {};  // id -> org
let _modalCallback = null;

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
                else if (tab === 'templates') this.loadTemplateTenants().then(() => this.loadTemplates());
                else if (tab === 'training-history') this.loadTrainingHistory();
                else if (tab === 'compute-history') this.loadComputeHistory();
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
            else { const e = await resp.json(); alert(e.detail || '创建失败'); }
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
            else { const e = await resp.json(); alert(e.detail || '更新失败'); }
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
            else { const e = await resp.json(); alert(e.detail || '创建失败'); }
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
            else { const e = await resp.json(); alert(e.detail || '更新失败'); }
        });
    },

    async deleteRole(id, name) {
        if (!confirm(`确定删除角色 ${name}？`)) return;
        const resp = await AUTH.authFetch(`/api/admin/roles/${id}`, { method: 'DELETE' });
        if (resp.ok) { await this.loadRoles(); this.renderRoles(); }
        else { const e = await resp.json(); alert(e.detail || '删除失败'); }
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
            else { const e = await resp.json(); alert(e.detail || '创建失败'); }
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
            else { const e = await resp.json(); alert(e.detail || '更新失败'); }
        });
    },

    async deleteOrg(id, name) {
        if (!confirm(`确定删除组织 ${name}？`)) return;
        const resp = await AUTH.authFetch(`/api/admin/organizations/${id}`, { method: 'DELETE' });
        if (resp.ok) { await this.loadOrgs(); this.renderOrgs(); }
        else { const e = await resp.json(); alert(e.detail || '删除失败'); }
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
            else { const e = await resp.json(); alert(e.detail || '授权失败'); }
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
        if (_modalCallback) _modalCallback();
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

    // ==================== 训练历史 ====================
    async loadTrainingHistory() {
        const tenantId = document.getElementById('training-tenant-filter')?.value || '';
        let url = '/api/training/sessions?limit=50';
        if (tenantId) url += `&tenant_id=${tenantId}`;
        const resp = await AUTH.authFetch(url);
        if (!resp.ok) return;
        const result = await resp.json();
        const tbody = document.querySelector('#training-history-table tbody');
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
                <button class="btn btn-sm" onclick="Admin.showTrainingDetail(${s.id})">详情</button>
            </td>
        </tr>`).join('');
    },

    async showTrainingDetail(sessionId) {
        const resp = await AUTH.authFetch(`/api/training/sessions/${sessionId}/iterations`);
        if (!resp.ok) return alert('获取详情失败');
        const iterations = await resp.json();
        let html = `<div style="max-height:500px;overflow-y:auto;">`;
        if (!iterations.length) {
            html += '<p>暂无迭代记录</p>';
        }
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
    _tplTenants: [],

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
            tbody.innerHTML = '<tr><td colspan="8" class="empty-state">暂无模版</td></tr>';
            return;
        }
        tbody.innerHTML = list.map(t => `<tr>
            <td>${t.id}</td>
            <td>${t.name}</td>
            <td>${t.tenant_id ? '<span class="tag">租户: ' + t.tenant_id + '</span>' : '<span class="tag" style="background:#e8f5e9;color:#2e7d32">全局</span>'}</td>
            <td>${t.file_name}</td>
            <td>${t.file_name_rule || '-'}</td>
            <td>${t.encrypt_password || '<span style="color:#999">不加密</span>'}</td>
            <td>${t.is_active ? '<span style="color:green">启用</span>' : '<span style="color:#999">停用</span>'}</td>
            <td class="actions">
                <button class="btn btn-sm" onclick="Admin.downloadTemplate(${t.id}, '${t.file_name.replace(/'/g, "\\'")}')">下载</button>
                <button class="btn btn-sm" onclick="Admin.showEditTemplate(${t.id})">编辑</button>
                <button class="btn btn-sm btn-danger" onclick="Admin.deleteTemplate(${t.id}, '${t.name.replace(/'/g, "\\'")}')">停用</button>
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
            const resp = await AUTH.authFetch('/api/admin/templates', { method: 'POST', body: fd });
            if (resp.ok) { this.closeModal(); this.loadTemplates(); }
            else { const e = await resp.json(); alert(e.detail || '创建失败'); }
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
            const resp = await AUTH.authFetch(`/api/admin/templates/${id}`, { method: 'PUT', body: fd });
            if (resp.ok) { this.closeModal(); this.loadTemplates(); }
            else { const e = await resp.json(); alert(e.detail || '更新失败'); }
        });
    },

    async deleteTemplate(id, name) {
        if (!confirm(`确定停用模版 ${name}？`)) return;
        const resp = await AUTH.authFetch(`/api/admin/templates/${id}`, { method: 'DELETE' });
        if (resp.ok) this.loadTemplates();
        else alert('操作失败');
    },

    downloadTemplate(id, fileName) {
        this._fetchAndDownload(`/api/admin/templates/${id}/download`, fileName);
    },

    // ==================== 计算历史 ====================
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
            // 根据格式调整文件名
            let name = fileName || 'download.xlsx';
            if (format === 'pdf') name = name.replace(/\.(xlsx?|csv)$/i, '') + '.pdf';
            else if (format === 'encrypted') name = name.replace(/\.(xlsx?)$/i, '') + '_加密.xlsx';
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
        if (password === null) return;  // 取消
        let url = `/api/assets/${assetId}/download?format=encrypted`;
        if (password) url += `&password=${encodeURIComponent(password)}`;
        this._fetchAndDownload(url, fileName.replace(/\.(xlsx?)$/i, '') + '_加密.xlsx');
    },

    async _fetchAndDownload(url, fileName) {
        try {
            const resp = await AUTH.authFetch(url);
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
                <div style="padding:6px 12px;cursor:pointer;font-size:12px;white-space:nowrap;" onmouseover="this.style.background='#f0f0f0'" onmouseout="this.style.background='#fff'" onclick="Admin.downloadAsset(${assetId},'${fileName.replace(/'/g, "\\'")}');this.parentElement.style.display='none'">
                    原始文件
                </div>
                <div style="padding:6px 12px;cursor:pointer;font-size:12px;white-space:nowrap;" onmouseover="this.style.background='#f0f0f0'" onmouseout="this.style.background='#fff'" onclick="Admin.downloadAsset(${assetId},'${fileName.replace(/'/g, "\\'")}','pdf');this.parentElement.style.display='none'">
                    下载 PDF
                </div>
                <div style="padding:6px 12px;cursor:pointer;font-size:12px;white-space:nowrap;" onmouseover="this.style.background='#f0f0f0'" onmouseout="this.style.background='#fff'" onclick="Admin.downloadAssetEncrypted(${assetId},'${fileName.replace(/'/g, "\\'")}');this.parentElement.style.display='none'">
                    加密 Excel
                </div>
            </div>
        </div>`;
    },

    async loadComputeHistory() {
        const tenantId = document.getElementById('compute-tenant-filter')?.value || '';
        const status = document.getElementById('compute-status-filter')?.value || '';
        let url = '/api/compute2/tasks?limit=50';
        if (tenantId) url += `&tenant_id=${tenantId}`;
        if (status) url += `&status=${status}`;
        const resp = await AUTH.authFetch(url);
        if (!resp.ok) return;
        const result = await resp.json();
        const tbody = document.querySelector('#compute-history-table tbody');
        if (!result.items || !result.items.length) {
            tbody.innerHTML = '<tr><td colspan="9" class="empty-state">暂无计算记录</td></tr>';
            return;
        }
        tbody.innerHTML = result.items.map(t => `<tr>
            <td>${t.id}</td>
            <td>${t.tenant_id}</td>
            <td>${t.script_id || (t.analysis_report?.original_script_id || '-')}</td>
            <td><span class="status-${t.status}">${t.status}</span></td>
            <td>${t.inputs ? t.inputs.length : 0}</td>
            <td>${t.duration_seconds != null ? t.duration_seconds.toFixed(1) : '-'}</td>
            <td>${t.created_at ? new Date(t.created_at).toLocaleString() : '-'}</td>
            <td>${t.finished_at ? new Date(t.finished_at).toLocaleString() : '-'}</td>
            <td>
                <button class="btn btn-sm" onclick="Admin.showComputeDetail(${t.id})">详情</button>
                ${t.status === 'completed' ? `<button class="btn btn-sm btn-primary" style="margin-left:4px;" onclick="Admin.showGenerateReport(${t.id}, '${t.tenant_id}')">下载报表</button>` : ''}
            </td>
        </tr>`).join('');
    },

    async showComputeDetail(taskId) {
        const resp = await AUTH.authFetch(`/api/compute2/tasks/${taskId}`);
        if (!resp.ok) return alert('获取详情失败');
        const task = await resp.json();

        let html = '<div style="max-height:500px;overflow-y:auto;">';

        // 基本信息
        html += `<div style="margin-bottom:16px;">
            <h4 style="margin:0 0 8px">基本信息</h4>
            <div style="display:grid;grid-template-columns:1fr 1fr;gap:6px;font-size:13px;">
                <div>租户: <strong>${task.tenant_id}</strong></div>
                <div>状态: <span class="status-${task.status}">${task.status}</span></div>
                <div>耗时: ${task.duration_seconds != null ? task.duration_seconds.toFixed(1) + '秒' : '-'}</div>
                <div>脚本ID: ${task.script_id || '-'}</div>
            </div>
        </div>`;

        // 输入文件（含下载下拉）
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

        // 输出文件（含下载下拉）
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

        // 结果摘要
        if (task.result_summary) {
            html += `<div style="margin-bottom:16px;">
                <h4 style="margin:0 0 8px">结果摘要</h4>
                <pre style="font-size:12px;background:#f5f5f5;padding:8px;border-radius:4px;">${JSON.stringify(task.result_summary, null, 2)}</pre>
            </div>`;
        }

        // 错误信息
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
        // 加载全局+该租户的模版
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
            const useHistory = document.getElementById('m-rpt-history').checked;
            const periodFrom = document.getElementById('m-rpt-from').value;
            const periodTo = document.getElementById('m-rpt-to').value;

            if (useHistory && (!periodFrom || !periodTo)) {
                return alert('启用历史时请选择薪资周期范围');
            }

            // 显示加载状态
            const confirmBtn = document.getElementById('modal-confirm');
            const origText = confirmBtn.textContent;
            confirmBtn.textContent = '生成中...';
            confirmBtn.disabled = true;

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
                    const err = await resp.json();
                    alert(err.detail || '报表生成失败');
                    return;
                }

                // 下载生成的文件
                const blob = await resp.blob();
                const cd = resp.headers.get('content-disposition') || '';
                let fileName = '报表.xlsx';
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

                this.closeModal();
            } catch (e) {
                alert('报表生成失败: ' + e.message);
            } finally {
                confirmBtn.textContent = origText;
                confirmBtn.disabled = false;
            }
        });
    },
};

document.addEventListener('DOMContentLoaded', () => {
    Admin.init();
    // 点击空白处关闭下载下拉菜单
    document.addEventListener('click', (e) => {
        if (!e.target.closest('[id^="dl-"]') && !e.target.closest('.btn')) {
            document.querySelectorAll('[id^="dl-"]').forEach(el => el.style.display = 'none');
        }
    });
});
