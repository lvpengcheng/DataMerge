/**
 * admin.js - 管理后台页面逻辑
 */

// 缓存数据
let _roles = [];
let _orgs = [];
let _orgsFlatMap = {};  // id -> org
let _modalCallback = null;
let _migItems = [];     // 测试迁移：拉取到的脚本列表
let _migUseSourceTenant = false;   // 是否处于“沿用来源租户”模式

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
        if (!AUTH.hasPerm('can_manage_users')) {
            alert('无权访问管理后台');
            window.location.href = '/dashboard';
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
                else if (tab === 'test-migration') this.initMigration();
                else if (tab === 'sop-rules') this.loadSopRules();
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
                    <button class="btn btn-sm" onclick="Admin.setPassword(${u.id}, '${u.username}')">修改密码</button>
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

    async setPassword(id, username) {
        const pwd = prompt(`为用户 ${username} 设定新密码（至少 6 位）：`);
        if (pwd === null) return;                    // 取消
        if (!pwd || pwd.trim().length < 6) { alert('密码至少 6 位'); return; }
        const resp = await AUTH.authFetch(`/api/admin/users/${id}/set-password`, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json' },
            body: JSON.stringify({ new_password: pwd.trim() }),
        });
        if (resp.ok) {
            alert(`用户 ${username} 密码已修改`);
        } else {
            let detail = '修改失败';
            try { const d = await resp.json(); if (d && d.detail) detail = d.detail; } catch (_) {}
            alert(detail);
        }
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
            <div class="form-group"><label>权限</label>
                ${this._renderPermissionTree({can_train: true, can_compute: true})}
            </div>
        `, async () => {
            const perms = this._collectPermissions();
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
        }, { wide: true });
        this._bindPermissionTree();
    },

    async showEditRole(id) {
        const role = _roles.find(r => r.id === id);
        if (!role) return;

        this.openModal('编辑角色', `
            <div class="form-group"><label>角色名</label><input id="m-role-name" value="${role.name}" ${role.is_system ? 'disabled' : ''}></div>
            <div class="form-group"><label>描述</label><input id="m-role-desc" value="${role.description || ''}"></div>
            <div class="form-group"><label>权限</label>
                ${this._renderPermissionTree(role.permissions || {})}
            </div>
        `, async () => {
            const perms = this._collectPermissions();
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
        }, { wide: true });
        this._bindPermissionTree();
    },

    // ============ 权限树定义与渲染（按真实菜单结构组织） ============
    _PERMISSION_TREE: [
        {
            label: '首页', children: [
                { key: 'can_view', label: '查看首页（仪表盘）' },
            ]
        },
        {
            label: '规则', children: [
                { key: 'menu.rules', label: '查看规则页' },
            ]
        },
        {
            label: '智训', children: [
                { key: 'can_train', label: '进入智训页 / 训练脚本' },
            ]
        },
        {
            label: '智算', children: [
                { key: 'can_compute', label: '进入智算页 / 执行脚本' },
            ]
        },
        {
            label: '智能小工具', children: [
                { key: 'menu.tools', label: '进入智能小工具页' },
                { key: 'tools.split_sheet', label: '└ Sheet 拆分' },
                { key: 'tools.data_merge', label: '└ 多表数据合并' },
                { key: 'tools.data_integrate', label: '└ 多表整合对比' },
                { key: 'tools.data_integrate.create', label: '　　└ 新增方案' },
                { key: 'tools.data_integrate.apply', label: '　　└ 应用方案' },
                { key: 'tools.data_integrate.edit', label: '　　└ 修改方案' },
                { key: 'tools.data_integrate.delete', label: '　　└ 删除方案' },
                { key: 'tools.templates', label: '└ 模版管理（管理员）' },
                { key: 'tools.training_history', label: '└ 训练历史' },
                { key: 'tools.compute_history', label: '└ 计算历史' },
                { key: 'tools.data_compare', label: '└ 数据对比（管理员）' },
                { key: 'tools.sop', label: '└ SOP维护' },
                { key: 'tools.sop.create', label: '　　└ 新建SOP / 上传文件' },
                { key: 'tools.sop.review', label: '　　└ 人工审核' },
                { key: 'tools.sop.manage', label: '　　└ 规则文件管理 / 删除SOP' },
            ]
        },
        {
            label: '管理后台', children: [
                { key: 'admin', label: '超级管理员（隐式授予所有权限）' },
                { key: 'can_manage_users', label: '进入管理后台' },
                { key: 'admin.users', label: '└ 用户管理' },
                { key: 'admin.roles', label: '└ 角色管理' },
                { key: 'admin.orgs', label: '└ 组织管理' },
                { key: 'admin.tenant_auth', label: '└ 租户授权' },
                { key: 'admin.ref_data', label: '└ 基础数据' },
                { key: 'admin.scripts', label: '└ 脚本管理' },
            ]
        },
    ],

    _renderPermissionTree(currentPerms) {
        const known = new Set();
        const groups = this._PERMISSION_TREE.map((group, gi) => {
            const items = group.children.map(item => {
                known.add(item.key);
                const checked = currentPerms[item.key] === true ? 'checked' : '';
                return `
                    <label class="perm-leaf" title="${item.key}">
                        <input type="checkbox" class="perm-leaf-cb" data-perm-key="${item.key}" data-group="${gi}" ${checked}>
                        <span class="perm-leaf-label">${item.label}</span>
                    </label>`;
            }).join('');
            const total = group.children.length;
            return `
                <div class="perm-group" data-group="${gi}">
                    <div class="perm-group-header" data-toggle-group="${gi}">
                        <span class="perm-group-toggle">▾</span>
                        <input type="checkbox" class="perm-group-cb" data-group="${gi}" onclick="event.stopPropagation();">
                        <span class="perm-group-label">${group.label}</span>
                        <span class="perm-group-count" data-count="${gi}">0 / ${total}</span>
                    </div>
                    <div class="perm-group-items">${items}</div>
                </div>`;
        }).join('');

        // 兼容已存在但未在树中定义的 key
        const extras = Object.keys(currentPerms).filter(k => !known.has(k));
        let extraBlock = '';
        if (extras.length) {
            const extraItems = extras.map(k => `
                <label class="perm-leaf" title="自定义权限">
                    <input type="checkbox" class="perm-leaf-cb" data-perm-key="${k}" data-group="extra" ${currentPerms[k] ? 'checked' : ''}>
                    <span class="perm-leaf-label">${k}</span>
                </label>`).join('');
            extraBlock = `
                <div class="perm-group perm-group--extra" data-group="extra">
                    <div class="perm-group-header">
                        <span class="perm-group-toggle">▾</span>
                        <span class="perm-group-label">自定义权限（树外保留）</span>
                        <span class="perm-group-count" data-count="extra">0 / ${extras.length}</span>
                    </div>
                    <div class="perm-group-items">${extraItems}</div>
                </div>`;
        }

        return `
            <div class="perm-tree-toolbar">
                <button type="button" class="btn-mini" data-perm-action="expand-all">展开全部</button>
                <button type="button" class="btn-mini" data-perm-action="collapse-all">收起全部</button>
                <button type="button" class="btn-mini" data-perm-action="check-all">全部勾选</button>
                <button type="button" class="btn-mini" data-perm-action="uncheck-all">全部清空</button>
                <span class="perm-tree-summary" id="perm-tree-summary"></span>
            </div>
            <div id="perm-tree" class="perm-tree">${groups}${extraBlock}</div>`;
    },

    _bindPermissionTree() {
        const tree = document.getElementById('perm-tree');
        if (!tree) return;

        const updateSummary = () => {
            const all = tree.querySelectorAll('.perm-leaf-cb');
            const checked = tree.querySelectorAll('.perm-leaf-cb:checked');
            const summary = document.getElementById('perm-tree-summary');
            if (summary) summary.textContent = `已勾选 ${checked.length} / ${all.length}`;
        };

        const syncGroup = (gi) => {
            const group = tree.querySelector(`.perm-group[data-group="${gi}"]`);
            const groupCb = group ? group.querySelector('.perm-group-cb') : null;
            const leaves = tree.querySelectorAll(`.perm-leaf-cb[data-group="${gi}"]`);
            const countEl = tree.querySelector(`[data-count="${gi}"]`);
            if (countEl) {
                const total = leaves.length;
                const cnt = Array.from(leaves).filter(cb => cb.checked).length;
                countEl.textContent = `${cnt} / ${total}`;
            }
            if (!groupCb || !leaves.length) return;
            const all = Array.from(leaves).every(cb => cb.checked);
            const some = Array.from(leaves).some(cb => cb.checked);
            groupCb.checked = all;
            groupCb.indeterminate = !all && some;
        };

        // 折叠/展开（点 header 区域，但点 checkbox/leaf 不触发）
        tree.querySelectorAll('.perm-group-header[data-toggle-group]').forEach(h => {
            h.addEventListener('click', (e) => {
                if (e.target.closest('.perm-group-cb')) return;
                const group = h.parentElement;
                if (group) group.classList.toggle('collapsed');
            });
        });

        // group 复选 → 同步所有 leaf
        tree.querySelectorAll('.perm-group-cb').forEach(gcb => {
            const gi = gcb.dataset.group;
            syncGroup(gi);
            gcb.addEventListener('change', () => {
                tree.querySelectorAll(`.perm-leaf-cb[data-group="${gi}"]`).forEach(lcb => { lcb.checked = gcb.checked; });
                gcb.indeterminate = false;
                syncGroup(gi);
                updateSummary();
            });
        });

        // leaf 改 → 同步所属 group
        tree.querySelectorAll('.perm-leaf-cb').forEach(lcb => {
            lcb.addEventListener('change', () => {
                syncGroup(lcb.dataset.group);
                updateSummary();
            });
        });

        // 工具栏
        const toolbar = tree.previousElementSibling;
        if (toolbar) {
            toolbar.querySelectorAll('[data-perm-action]').forEach(btn => {
                btn.addEventListener('click', () => {
                    const act = btn.dataset.permAction;
                    if (act === 'expand-all') {
                        tree.querySelectorAll('.perm-group').forEach(g => g.classList.remove('collapsed'));
                    } else if (act === 'collapse-all') {
                        tree.querySelectorAll('.perm-group').forEach(g => g.classList.add('collapsed'));
                    } else if (act === 'check-all') {
                        tree.querySelectorAll('.perm-leaf-cb').forEach(cb => { cb.checked = true; });
                        tree.querySelectorAll('.perm-group-cb').forEach(gcb => { gcb.checked = true; gcb.indeterminate = false; });
                        tree.querySelectorAll('.perm-group').forEach(g => syncGroup(g.dataset.group));
                        updateSummary();
                    } else if (act === 'uncheck-all') {
                        tree.querySelectorAll('.perm-leaf-cb').forEach(cb => { cb.checked = false; });
                        tree.querySelectorAll('.perm-group-cb').forEach(gcb => { gcb.checked = false; gcb.indeterminate = false; });
                        tree.querySelectorAll('.perm-group').forEach(g => syncGroup(g.dataset.group));
                        updateSummary();
                    }
                });
            });
        }

        updateSummary();
    },

    _collectPermissions() {
        const tree = document.getElementById('perm-tree');
        if (!tree) return {};
        const perms = {};
        tree.querySelectorAll('.perm-leaf-cb').forEach(cb => {
            if (cb.checked) perms[cb.dataset.permKey] = true;
        });
        return perms;
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
    openModal(title, bodyHtml, onConfirm, opts) {
        document.getElementById('modal-title').textContent = title;
        document.getElementById('modal-body').innerHTML = bodyHtml;
        document.getElementById('modal-overlay').style.display = 'flex';
        // 角色/权限等内容较宽的弹窗用 modal--wide 加宽，避免权限树被挤在 480px 里
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
        const showInactive = document.getElementById('ref-show-inactive')?.checked || false;
        let url = '/api/assets?asset_type=reference';
        if (categoryId) url += `&category_id=${categoryId}`;
        if (showInactive) url += '&is_active=false';   // 展示已停用，便于启用/删除
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
            tbody.innerHTML = `<tr><td colspan="9" class="empty-state">${showInactive ? '暂无已停用的基础数据' : '暂无基础数据'}</td></tr>`;
            return;
        }
        tbody.innerHTML = assets.map(a => {
            const actions = a.is_active
                ? `<button class="btn btn-sm" onclick="Admin.previewAsset(${a.id})">预览</button>
                   <button class="btn btn-sm" onclick="Admin.updateAssetVersion(${a.id})">更新</button>
                   <button class="btn btn-sm btn-danger" onclick="Admin.deleteAsset(${a.id})">停用</button>
                   <button class="btn btn-sm btn-danger" onclick="Admin.hardDeleteAsset(${a.id})">删除</button>`
                : `<button class="btn btn-sm" onclick="Admin.previewAsset(${a.id})">预览</button>
                   <button class="btn btn-sm btn-primary" onclick="Admin.enableAsset(${a.id})">启用</button>
                   <button class="btn btn-sm btn-danger" onclick="Admin.hardDeleteAsset(${a.id})">删除</button>`;
            return `<tr>
            <td>${a.id}</td>
            <td>${a.name}</td>
            <td>${a.category_name || '-'}</td>
            <td>${a.tenant_id ? '<span class="tag">租户: ' + a.tenant_id + '</span>' : '<span class="tag" style="background:#e8f5e9;color:#2e7d32">全局</span>'}</td>
            <td>${a.file_name}</td>
            <td>v${a.version}</td>
            <td>${a.effective_from || '-'}</td>
            <td>${a.is_active ? '<span style="color:green">启用</span>' : '<span style="color:#999">停用</span>'}</td>
            <td>${actions}</td>
        </tr>`;
        }).join('');
    },

    showUploadRefData() {
        const catOptions = this._refCategories.map(c =>
            `<option value="${c.id}">${c.name}</option>`
        ).join('');
        const tenantOptions = this._refTenants.map(t =>
            `<option value="${t.tenant_id}">租户: ${t.tenant_id}</option>`
        ).join('');
        // 全局作用域仅管理员可管理（非管理员上传全局会被后端 403）
        const _isAdmin = (AUTH.getUser() || {}).role_name === 'admin';
        const globalOption = _isAdmin ? '<option value="">全局（所有租户可用）</option>' : '';
        this.openModal('上传基础数据', `
            <div style="display:flex;flex-direction:column;gap:12px;">
                <label>分类：<select id="ref-upload-category" style="padding:6px;border:1px solid #ddd;border-radius:4px;">${catOptions}</select></label>
                <label>名称：<input id="ref-upload-name" type="text" placeholder="留空用文件名；多选时按各自文件名（此框忽略）" style="width:100%;padding:6px;border:1px solid #ddd;border-radius:4px;"></label>
                <label>作用域：
                    <select id="ref-upload-scope" style="padding:6px;border:1px solid #ddd;border-radius:4px;">
                        ${globalOption}
                        ${tenantOptions}
                    </select>
                </label>
                <label>生效日期：<input id="ref-upload-from" type="date" style="padding:6px;border:1px solid #ddd;border-radius:4px;"></label>
                <label>失效日期：<input id="ref-upload-to" type="date" style="padding:6px;border:1px solid #ddd;border-radius:4px;"></label>
                <label>文件：<input id="ref-upload-file" type="file" accept=".xlsx,.xls" multiple></label>
                <div style="font-size:12px;color:#888;">可一次选多个文件批量上传，每个文件以其文件名作为基础数据名称。</div>
            </div>
        `, async () => {
            const fileInput = document.getElementById('ref-upload-file');
            if (!fileInput.files.length) return alert('请选择文件');
            const catId = document.getElementById('ref-upload-category').value;
            const scopeVal = document.getElementById('ref-upload-scope').value;
            const ef = document.getElementById('ref-upload-from').value;
            const et = document.getElementById('ref-upload-to').value;

            if (fileInput.files.length > 1) {
                // 批量上传：每个文件用自己的文件名，共享分类/作用域/日期
                const fd = new FormData();
                for (const f of fileInput.files) fd.append('files', f);
                fd.append('asset_type', 'reference');
                if (catId) fd.append('category_id', catId);
                if (scopeVal) fd.append('tenant_id', scopeVal);
                if (ef) fd.append('effective_from', ef);
                if (et) fd.append('effective_to', et);
                const resp = await AUTH.authFetch('/api/assets/upload-batch', {method: 'POST', body: fd});
                if (resp.ok) {
                    const r = await resp.json();
                    this.closeModal();
                    this.loadRefData();
                    const failMsg = (r.failed && r.failed.length)
                        ? '\n失败 ' + r.failed.length + ' 个：\n' + r.failed.map(x => `${x.filename}: ${x.error}`).join('\n')
                        : '';
                    alert(`成功上传 ${r.created ? r.created.length : 0} 个${failMsg}`);
                } else {
                    alert('批量上传失败: ' + (await resp.text()));
                }
                return;
            }

            // 单文件（保持原路径）
            const fd = new FormData();
            fd.append('file', fileInput.files[0]);
            fd.append('asset_type', 'reference');
            if (catId) fd.append('category_id', catId);
            fd.append('name', document.getElementById('ref-upload-name').value || fileInput.files[0].name);
            if (scopeVal) fd.append('tenant_id', scopeVal);
            if (ef) fd.append('effective_from', ef);
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

    async hardDeleteAsset(assetId) {
        if (!confirm('物理删除此数据？将同时删除文件与记录，不可恢复！')) return;
        const resp = await AUTH.authFetch(`/api/assets/${assetId}?hard=true`, {method: 'DELETE'});
        if (!resp.ok) return alert('删除失败: ' + (await resp.text()));
        this.loadRefData();
    },

    async enableAsset(assetId) {
        const resp = await AUTH.authFetch(`/api/assets/${assetId}`, {
            method: 'PUT',
            headers: {'Content-Type': 'application/json'},
            body: JSON.stringify({is_active: true}),
        });
        if (!resp.ok) return alert('启用失败: ' + (await resp.text()));
        this.loadRefData();
    },

    updateAssetVersion(assetId) {
        // 行内选新文件 → 以新版本更新（保留历史、停用旧版）
        const picker = document.createElement('input');
        picker.type = 'file';
        picker.accept = '.xlsx,.xls';
        picker.onchange = async () => {
            if (!picker.files.length) return;
            const fd = new FormData();
            fd.append('file', picker.files[0]);
            const resp = await AUTH.authFetch(`/api/assets/${assetId}/new-version`, {method: 'POST', body: fd});
            if (!resp.ok) return alert('更新失败: ' + (await resp.text()));
            alert('已更新为新版本');
            this.loadRefData();
        };
        picker.click();
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
            const delBtn = `<button class="btn btn-sm btn-danger" onclick="Admin.deleteScript(${s.id}, '${(s.name || '').replace(/'/g, "\\'")}')">删除</button>`;
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
                    <td class="actions">${action} ${delBtn}</td>
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

    async deleteScript(scriptId, name) {
        if (!confirm(`确定物理删除脚本「${name}」？\n此操作不可恢复：将删除脚本文件，并断开历史计算任务与该脚本的关联（计算历史与结果文件保留）。`)) return;
        const resp = await AUTH.authFetch(`/api/admin/scripts/${scriptId}`, { method: 'DELETE' });
        const data = await resp.json().catch(() => ({}));
        if (!resp.ok) {
            alert(data.detail || '删除失败');
            return;
        }
        alert(data.message || '已删除');
        this.loadScripts();
    },

    // ==================== 测试迁移 ====================
    async initMigration() {
        // 载入目标租户下拉
        const sel = document.getElementById('mig-target-tenant');
        if (sel && !sel.dataset.loaded) {
            try {
                const resp = await AUTH.authFetch('/api/admin/tenant-auth/tenants');
                if (resp.ok) {
                    const tenants = await resp.json();
                    // 首项“空租户”：迁移时沿用各脚本来源租户并自动创建
                    const opts = ['<option value="">（空）沿用来源租户（自动创建）</option>']
                        .concat((tenants || []).map(t => `<option value="${t}">${t}</option>`));
                    sel.innerHTML = opts.join('');
                    sel.dataset.loaded = '1';
                }
            } catch (_) {}
        }
    },

    _setMigStatus(text) {
        const el = document.getElementById('mig-status');
        if (el) el.textContent = text || '';
    },

    async loadMigrationScripts() {
        // 空值 = 沿用各脚本来源租户，属于合法选择，不再拦截
        const tenant = (document.getElementById('mig-target-tenant')?.value || '').trim();
        this._setMigStatus('正在连接测试环境并拉取...');
        try {
            const resp = await AUTH.authFetch(`/api/admin/migration/remote-scripts?target_tenant=${encodeURIComponent(tenant)}`);
            if (!resp.ok) { await _alertErr(resp, '拉取失败'); this._setMigStatus('拉取失败'); return; }
            const data = await resp.json();
            _migItems = data.items || [];
            _migUseSourceTenant = !!data.use_source_tenant;
            this.renderMigrationList();
            const modeTip = data.use_source_tenant ? '（沿用来源租户）' : '';
            this._setMigStatus(`来源：${data.source_url || ''}${modeTip}，共 ${_migItems.length} 个脚本`);
        } catch (e) {
            this._setMigStatus('拉取失败: ' + e.message);
        }
    },

    renderMigrationList() {
        const tbody = document.querySelector('#migration-table tbody');
        const _esc = (s) => String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c]));
        if (!_migItems.length) {
            tbody.innerHTML = '<tr><td colspan="8" style="text-align:center;color:#999;">无数据，请先拉取</td></tr>';
            document.getElementById('mig-select-all').checked = false;
            return;
        }
        const kw = (document.getElementById('mig-filter')?.value || '').trim().toLowerCase();
        // 保留原始下标（data-i 指向 _migItems），过滤仅隐藏不匹配行
        const rows = _migItems.map((it, i) => {
            if (kw) {
                const hay = [it.name, it.tenant_id, it.hash, it.mode]
                    .map(v => String(v == null ? '' : v).toLowerCase()).join(' ');
                if (!hay.includes(kw)) return '';
            }
            const acc = (it.accuracy != null) ? (it.accuracy * 100).toFixed(1) + '%' : '-';
            let stat = '';
            if (it.already_migrated) stat = '<span style="color:#2e7d32;">已迁移</span>';
            else if (it.exists_by_hash) stat = '<span style="color:#e65100;">已存在(将覆盖)</span>';
            else stat = '<span style="color:#888;">未迁移</span>';
            // 撞名预警：沿用来源租户模式下，目标租户在本环境已存在 → 将合并进已有租户
            if (_migUseSourceTenant && it.dest_tenant_exists) {
                stat += ' <span style="color:#c62828;" title="目标租户已存在，脚本将合并进该租户">⚠合并</span>';
            }
            return `<tr>
                <td><input type="checkbox" class="mig-cb" data-i="${i}"></td>
                <td>${_esc(it.name)}</td>
                <td><code style="font-size:12px;">${_esc(it.hash)}</code></td>
                <td>${_esc(it.tenant_id)}</td>
                <td>${_esc(it.mode)}</td>
                <td>${acc}</td>
                <td>${stat}</td>
                <td><input type="text" class="mig-newname" data-i="${i}" placeholder="沿用原名"
                       value="${_esc(it.new_name || '')}" style="width:140px;"></td>
            </tr>`;
        }).filter(Boolean);
        tbody.innerHTML = rows.length ? rows.join('')
            : '<tr><td colspan="8" style="text-align:center;color:#999;">无匹配脚本</td></tr>';
        const selAll = document.getElementById('mig-select-all');
        if (selAll) selAll.checked = false;
    },

    migToggleAll(checked) {
        document.querySelectorAll('#migration-table .mig-cb').forEach(cb => { cb.checked = checked; });
    },

    _selectedMigItems() {
        return Array.from(document.querySelectorAll('#migration-table .mig-cb:checked'))
            .map(cb => {
                const i = parseInt(cb.dataset.i, 10);
                const it = _migItems[i];
                if (!it) return null;
                const nameInput = document.querySelector(`#migration-table .mig-newname[data-i="${i}"]`);
                const new_name = nameInput ? nameInput.value.trim() : '';
                return { db_id: it.db_id, hash: it.hash, name: it.name, tenant_id: it.tenant_id, new_name };
            })
            .filter(Boolean);
    },

    async doMigrate() {
        // 空值 = 沿用来源租户，属于合法选择
        const tenant = (document.getElementById('mig-target-tenant')?.value || '').trim();
        const items = this._selectedMigItems();
        if (!items.length) { alert('请勾选要迁移的脚本'); return; }
        // 撞名预警：沿用来源租户模式下，若来源租户在本环境已存在 → 二次确认合并
        if (_migUseSourceTenant) {
            const dupTenants = [...new Set(
                Array.from(document.querySelectorAll('#migration-table .mig-cb:checked'))
                    .map(cb => _migItems[parseInt(cb.dataset.i, 10)])
                    .filter(it => it && it.dest_tenant_exists)
                    .map(it => it.dest_tenant))];
            if (dupTenants.length) {
                if (!confirm(`以下来源租户在本环境已存在，迁移将把脚本合并进这些已有租户（不会新建/改名）：\n${dupTenants.join('、')}\n是否继续？`)) return;
            }
        }
        await this._runMigrate(tenant, items, false);
    },

    async _runMigrate(tenant, items, overwrite) {
        this._setMigStatus(overwrite ? '覆盖迁移中...' : '迁移中...');
        try {
            const resp = await AUTH.authFetch('/api/admin/migration/import', {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ target_tenant_id: tenant, items, overwrite }),
            });
            if (!resp.ok) { await _alertErr(resp, '迁移失败'); this._setMigStatus('迁移失败'); return; }
            const r = await resp.json();
            const imported = r.imported || [], conflicts = r.conflicts || [], skipped = r.skipped || [];
            // 有冲突且本次未覆盖 → 询问是否覆盖
            if (conflicts.length && !overwrite) {
                const names = conflicts.map(c => c.name || c.hash).join('、');
                if (confirm(`以下脚本在目标租户已存在相同代码哈希：\n${names}\n是否覆盖迁移？`)) {
                    // 覆盖：把冲突项 + 已成功项合并重发（仅冲突项需覆盖，已导入的无需重发）
                    // 保留用户在原始 items 里填的 new_name（conflicts 来自后端响应，不含 new_name）
                    await this._runMigrate(tenant, conflicts.map(c => {
                        const orig = items.find(it => it.hash === c.hash && it.db_id === c.db_id);
                        return { db_id: c.db_id, hash: c.hash, name: c.name, tenant_id: c.tenant_id,
                                 new_name: orig ? (orig.new_name || '') : '' };
                    }), true);
                    return;
                }
            }
            let msg = `迁移完成：成功 ${imported.length}`;
            if (conflicts.length && !overwrite) msg += `，跳过冲突 ${conflicts.length}`;
            if (skipped.length) msg += `，失败 ${skipped.length}`;
            this._setMigStatus(msg);
            if (skipped.length) {
                alert('以下未迁移：\n' + skipped.map(s => `${s.name || ''}: ${s.reason}`).join('\n'));
            }
            this.loadMigrationScripts();  // 刷新已迁移状态
        } catch (e) {
            this._setMigStatus('迁移失败: ' + e.message);
        }
    },

    // ==================== SOP 规则文件管理 ====================
    _sopRules: [],

    async loadSopRules() {
        try {
            const resp = await AUTH.authFetch('/api/tools/sop/rules');
            if (!resp.ok) return _alertErr(resp, '加载规则列表失败');
            const data = await resp.json();
            this._sopRules = data.items || [];
            this._renderSopRules();
        } catch (e) {
            alert('加载规则列表失败: ' + e.message);
        }
    },

    _renderSopRules() {
        const tbody = document.querySelector('#sop-rules-table tbody');
        const items = this._sopRules || [];
        if (!items.length) {
            tbody.innerHTML = '<tr><td colspan="8" class="empty-state">暂无规则文件</td></tr>';
            return;
        }
        tbody.innerHTML = items.map(r => `<tr>
            <td>${r.id}</td>
            <td>${r.scope === 'global' ? '<span class="tag">全局</span>' : '<span class="tag" style="background:#e3f2fd;">按客户</span>'}</td>
            <td>${r.customer_name ? this._escapeHtml(r.customer_name) : '-'}</td>
            <td>${this._escapeHtml(r.name || '-')}</td>
            <td>${this._escapeHtml(r.file_name)}</td>
            <td><span class="${r.is_active ? 'status-active' : 'status-inactive'}">${r.is_active ? '启用' : '停用'}</span></td>
            <td>${r.updated_at ? new Date(r.updated_at).toLocaleString() : '-'}</td>
            <td class="actions">
                <button class="btn btn-sm" onclick="Admin.downloadSopRule(${r.id}, '${r.file_name.replace(/'/g, "\\'")}')">下载</button>
                <button class="btn btn-sm" style="margin-left:4px;" onclick="Admin.toggleSopRule(${r.id})">${r.is_active ? '停用' : '启用'}</button>
                <button class="btn btn-sm btn-danger" style="margin-left:4px;" onclick="Admin.deleteSopRule(${r.id})">删除</button>
            </td>
        </tr>`).join('');
    },

    _escapeHtml(s) {
        return String(s == null ? '' : s).replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
    },

    showCreateSopRule() {
        this.openModal('上传规则文件', `
            <div style="display:flex;flex-direction:column;gap:10px;">
                <div><label style="font-weight:500;">作用域 <span style="color:#f44336;">*</span></label>
                    <select id="m-soprule-scope" style="width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:6px;margin-top:4px;" onchange="Admin._onSopRuleScopeChange()">
                        <option value="global">全局大规则</option>
                        <option value="customer">按客户专属规则</option>
                    </select></div>
                <div id="m-soprule-cust-wrap" style="display:none;"><label style="font-weight:500;">客户名称 <span style="color:#f44336;">*</span></label>
                    <input type="text" id="m-soprule-customer" placeholder="客户名称（需与 SOP 条目中的客户名称一致）" style="width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:6px;margin-top:4px;"></div>
                <div><label style="font-weight:500;">规则文件 <span style="color:#f44336;">*</span></label>
                    <input type="file" id="m-soprule-file" accept=".txt,.md,.docx,.doc,.pdf,.xlsx,.xls" style="margin-top:4px;"></div>
                <div><label style="font-weight:500;">规则名称</label>
                    <input type="text" id="m-soprule-name" placeholder="规则名称/说明" style="width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:6px;margin-top:4px;"></div>
                <div><label style="font-weight:500;">描述</label>
                    <textarea id="m-soprule-desc" rows="2" placeholder="描述（可选）" style="width:100%;padding:8px 10px;border:1px solid #ddd;border-radius:6px;margin-top:4px;"></textarea></div>
            </div>`, () => {
            this._uploadSopRule();
        });
    },

    _onSopRuleScopeChange() {
        const scope = document.getElementById('m-soprule-scope').value;
        document.getElementById('m-soprule-cust-wrap').style.display = scope === 'customer' ? '' : 'none';
    },

    async _uploadSopRule() {
        const scope = document.getElementById('m-soprule-scope').value;
        const cust = document.getElementById('m-soprule-customer').value.trim();
        const file = document.getElementById('m-soprule-file').files[0];
        if (!file) return alert('请选择规则文件');
        if (scope === 'customer' && !cust) return alert('按客户规则必须填写客户名称');
        const fd = new FormData();
        fd.append('file', file);
        fd.append('scope', scope);
        fd.append('customer_name', cust);
        fd.append('name', document.getElementById('m-soprule-name').value);
        fd.append('description', document.getElementById('m-soprule-desc').value);
        const resp = await AUTH.authFetch('/api/tools/sop/rules', { method: 'POST', body: fd });
        if (!resp.ok) return _alertErr(resp, '上传失败');
        this.closeModal();
        this.loadSopRules();
    },

    async toggleSopRule(id) {
        const resp = await AUTH.authFetch(`/api/tools/sop/rules/${id}/toggle`, { method: 'POST' });
        if (!resp.ok) return _alertErr(resp, '切换失败');
        this.loadSopRules();
    },

    async deleteSopRule(id) {
        if (!confirm('确认删除该规则文件？删除后不再参与 AI 分析。')) return;
        const resp = await AUTH.authFetch(`/api/tools/sop/rules/${id}`, { method: 'DELETE' });
        if (!resp.ok) return _alertErr(resp, '删除失败');
        this.loadSopRules();
    },

    downloadSopRule(id, fileName) {
        this._downloadViaAuth(`/api/tools/sop/rules/${id}/download`, fileName);
    },

    async _downloadViaAuth(url, fileName) {
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
};

document.addEventListener('DOMContentLoaded', () => {
    Admin.init();
});
