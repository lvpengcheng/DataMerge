/**
 * auth.js - 共享认证工具
 * 在 training.html、compute.html、admin.html 中引入
 */

const AUTH = {
    TOKEN_KEY: 'datamerge_token',
    USER_KEY: 'datamerge_user',

    getToken() {
        return localStorage.getItem(this.TOKEN_KEY);
    },

    setToken(token) {
        localStorage.setItem(this.TOKEN_KEY, token);
    },

    getUser() {
        const data = localStorage.getItem(this.USER_KEY);
        return data ? JSON.parse(data) : null;
    },

    setUser(user) {
        localStorage.setItem(this.USER_KEY, JSON.stringify(user));
    },

    logout() {
        localStorage.removeItem(this.TOKEN_KEY);
        localStorage.removeItem(this.USER_KEY);
        window.location.href = '/login';
    },

    isLoggedIn() {
        return !!this.getToken();
    },

    isAdmin() {
        const user = this.getUser();
        if (!user) return false;
        if (user.role_name === 'admin') return true;
        const p = user.permissions || {};
        return p.admin === true;
    },

    /** 取当前用户的权限字典 */
    getPermissions() {
        const user = this.getUser();
        return (user && user.permissions) || {};
    },

    /** 是否拥有某个权限键。admin 视为拥有所有权限 */
    hasPerm(key) {
        if (!key) return true;
        if (this.isAdmin()) return true;
        const p = this.getPermissions();
        return p && p[key] === true;
    },

    /**
     * 按 data-perm 控制元素显隐(支持空格/逗号分隔多 key,任一通过即显示)。
     * 用法: <a href="..." data-perm="menu.tools">...</a>
     * 对有权限的元素强制清空 inline display(覆盖模板中的 display:none)。
     */
    applyPermFilter(root) {
        const scope = root || document;
        scope.querySelectorAll('[data-perm]').forEach(el => {
            const raw = el.getAttribute('data-perm') || '';
            const keys = raw.split(/[\s,]+/).filter(Boolean);
            if (keys.length === 0) return;
            const ok = keys.some(k => this.hasPerm(k));
            if (ok) {
                el.style.display = '';
                el.removeAttribute('data-perm-hidden');
            } else {
                el.style.display = 'none';
                el.setAttribute('data-perm-hidden', '1');
            }
        });
    },

    /** 返回 Authorization 头 */
    getAuthHeaders() {
        const token = this.getToken();
        return token ? { 'Authorization': `Bearer ${token}` } : {};
    },

    /**
     * fetch 包装器：自动带 token，401 时跳转登录
     * 用法: await AUTH.authFetch('/api/xxx', { method: 'POST', body: formData })
     */
    async authFetch(url, options = {}) {
        const headers = { ...this.getAuthHeaders(), ...(options.headers || {}) };
        const resp = await fetch(url, { ...options, headers });
        if (resp.status === 401) {
            this.logout();
            return resp;
        }
        return resp;
    },

    /** 页面加载时检查登录状态，未登录跳转 /login */
    requireAuth() {
        if (!this.isLoggedIn()) {
            window.location.href = '/login';
            return false;
        }
        return true;
    },

    /** 在 header 中渲染用户信息 + 退出按钮 */
    renderUserInfo(headerElement) {
        const user = this.getUser();
        if (!user || !headerElement) return;

        // 优先填充已有的 #user-info 容器
        let userDiv = headerElement.querySelector('#user-info') || headerElement.querySelector('.user-info');
        if (userDiv && userDiv.children.length > 0) {
            // 已渲染过，但仍需确保权限过滤已应用
            try { this.applyPermFilter(document); } catch (_) {}
            return;
        }

        if (!userDiv) {
            userDiv = document.createElement('div');
            userDiv.className = 'user-info';
            headerElement.appendChild(userDiv);
        }

        userDiv.innerHTML = `
            <span class="user-name">${user.display_name || user.username}</span>
            <button class="btn-logout" onclick="AUTH.showChangePassword()">修改密码</button>
            <button class="btn-logout" onclick="AUTH.logout()">退出</button>
        `;

        // 渲染完成后立即应用权限过滤(导航/Tab 等)
        try { this.applyPermFilter(document); } catch (_) {}
    },

    /** 弹出「修改密码」对话框（自助改自己的密码，需校验原密码）。自带样式，全站可用。 */
    showChangePassword() {
        // 已存在则不重复注入
        if (document.getElementById('auth-cp-overlay')) return;

        const ov = document.createElement('div');
        ov.id = 'auth-cp-overlay';
        ov.style.cssText = 'position:fixed;inset:0;background:rgba(0,0,0,.45);z-index:99999;display:flex;align-items:center;justify-content:center;';
        ov.innerHTML = `
            <div style="background:#fff;border-radius:10px;width:360px;max-width:92vw;box-shadow:0 8px 30px rgba(0,0,0,.2);overflow:hidden;">
                <div style="padding:14px 18px;border-bottom:1px solid #eee;font-size:16px;font-weight:600;color:#333;">修改密码</div>
                <div style="padding:18px;display:flex;flex-direction:column;gap:12px;">
                    <input id="auth-cp-old" type="password" placeholder="原密码" autocomplete="current-password"
                           style="padding:9px 12px;border:1px solid #ddd;border-radius:6px;font-size:14px;">
                    <input id="auth-cp-new" type="password" placeholder="新密码（至少 6 位）" autocomplete="new-password"
                           style="padding:9px 12px;border:1px solid #ddd;border-radius:6px;font-size:14px;">
                    <input id="auth-cp-new2" type="password" placeholder="确认新密码" autocomplete="new-password"
                           style="padding:9px 12px;border:1px solid #ddd;border-radius:6px;font-size:14px;">
                    <div id="auth-cp-msg" style="min-height:18px;font-size:13px;color:#d32f2f;"></div>
                </div>
                <div style="padding:12px 18px;border-top:1px solid #eee;display:flex;justify-content:flex-end;gap:10px;">
                    <button id="auth-cp-cancel" class="btn-logout">取消</button>
                    <button id="auth-cp-ok" class="btn-logout" style="border-color:#1976d2;color:#fff;background:#1976d2;">确定</button>
                </div>
            </div>`;
        document.body.appendChild(ov);

        const close = () => { const el = document.getElementById('auth-cp-overlay'); if (el) el.remove(); };
        const msg = (t) => { const m = document.getElementById('auth-cp-msg'); if (m) m.textContent = t || ''; };

        ov.addEventListener('click', (e) => { if (e.target === ov) close(); });
        document.getElementById('auth-cp-cancel').onclick = close;

        const submit = async () => {
            const oldp = document.getElementById('auth-cp-old').value;
            const newp = document.getElementById('auth-cp-new').value;
            const newp2 = document.getElementById('auth-cp-new2').value;
            if (!oldp || !newp || !newp2) { msg('请填写完整'); return; }
            if (newp.length < 6) { msg('新密码至少 6 位'); return; }
            if (newp !== newp2) { msg('两次输入的新密码不一致'); return; }
            if (newp === oldp) { msg('新密码不能与原密码相同'); return; }
            const okBtn = document.getElementById('auth-cp-ok');
            okBtn.disabled = true; msg('');
            try {
                const resp = await this.authFetch('/api/auth/change-password', {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json' },
                    body: JSON.stringify({ old_password: oldp, new_password: newp }),
                });
                if (resp.ok) {
                    close();
                    alert('密码修改成功，请重新登录');
                    this.logout();
                } else {
                    let detail = '修改失败';
                    try { const d = await resp.json(); if (d && d.detail) detail = d.detail; } catch (_) {}
                    msg(detail);
                    okBtn.disabled = false;
                }
            } catch (err) {
                msg('网络错误，请重试');
                okBtn.disabled = false;
            }
        };

        document.getElementById('auth-cp-ok').onclick = submit;
        document.getElementById('auth-cp-new2').addEventListener('keydown', (e) => { if (e.key === 'Enter') submit(); });
        document.getElementById('auth-cp-old').focus();
    }
};
