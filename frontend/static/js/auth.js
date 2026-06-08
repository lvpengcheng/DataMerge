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
            <button class="btn-logout" onclick="AUTH.logout()">退出</button>
        `;

        // 渲染完成后立即应用权限过滤(导航/Tab 等)
        try { this.applyPermFilter(document); } catch (_) {}
    }
};
