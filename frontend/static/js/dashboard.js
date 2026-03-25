/**
 * dashboard.js - 首页租户计算状态总览
 */

document.addEventListener('DOMContentLoaded', function () {
    if (!AUTH.requireAuth()) return;
    AUTH.renderUserInfo(document.querySelector('header'));
    if (AUTH.isAdmin()) {
        const adminNav = document.getElementById('nav-admin');
        if (adminNav) adminNav.style.display = '';
    }
    loadDashboard();
});

async function loadDashboard() {
    const grid = document.getElementById('tenant-grid');
    const loading = document.getElementById('loading-indicator');
    const empty = document.getElementById('empty-state');
    const summary = document.getElementById('dashboard-summary');

    try {
        const resp = await AUTH.authFetch('/api/dashboard/tenants');
        if (!resp.ok) throw new Error('请求失败');
        const tenants = await resp.json();

        loading.style.display = 'none';

        if (!tenants || tenants.length === 0) {
            empty.style.display = 'block';
            return;
        }

        // 统计摘要
        const computed = tenants.filter(t => t.current_month_computed).length;
        summary.innerHTML = `
            <span class="summary-item">共 <strong>${tenants.length}</strong> 个租户</span>
            <span class="summary-item computed">已计算 <strong>${computed}</strong></span>
            <span class="summary-item pending">待计算 <strong>${tenants.length - computed}</strong></span>
        `;

        // 排序：未计算的排前面
        const sorted = [...tenants].sort((a, b) => {
            if (a.current_month_computed === b.current_month_computed) {
                return a.tenant_id.localeCompare(b.tenant_id);
            }
            return a.current_month_computed ? 1 : -1;
        });

        grid.innerHTML = sorted.map(renderTenantCard).join('');

    } catch (err) {
        loading.style.display = 'none';
        grid.innerHTML = '<div class="error-state">加载失败，请刷新重试</div>';
        console.error('Dashboard load error:', err);
    }
}

function renderTenantCard(t) {
    const statusClass = t.current_month_computed ? 'status-done' : 'status-pending';
    const statusText = t.current_month_computed ? '本月已计算' : '本月未计算';

    const lastMonthInfo = t.last_month_compute_time
        ? `<div class="card-meta">
               <span class="meta-label">上月计算时间</span>
               <span class="meta-value">${formatDateTime(t.last_month_compute_time)}</span>
           </div>
           <div class="card-meta">
               <span class="meta-label">耗时</span>
               <span class="meta-value">${formatDuration(t.last_month_duration_seconds)}</span>
           </div>`
        : '<div class="card-meta"><span class="meta-label">上月</span><span class="meta-value text-muted">无计算记录</span></div>';

    const actionBtn = t.current_month_computed
        ? ''
        : `<a href="/compute?tenant=${encodeURIComponent(t.tenant_id)}" class="btn-compute">去计算</a>`;

    return `
        <div class="tenant-card ${statusClass}">
            <div class="card-header">
                <h3 class="card-title">${escapeHtml(t.tenant_name || t.tenant_id)}</h3>
                <span class="status-badge ${statusClass}">${statusText}</span>
            </div>
            <div class="card-body">
                ${lastMonthInfo}
            </div>
            <div class="card-footer">
                ${actionBtn}
            </div>
        </div>
    `;
}

function formatDateTime(isoStr) {
    if (!isoStr) return '-';
    const d = new Date(isoStr);
    const pad = n => String(n).padStart(2, '0');
    return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
}

function formatDuration(seconds) {
    if (!seconds && seconds !== 0) return '-';
    if (seconds < 60) return `${seconds.toFixed(1)}秒`;
    const m = Math.floor(seconds / 60);
    const s = Math.round(seconds % 60);
    return `${m}分${s}秒`;
}

function escapeHtml(text) {
    const div = document.createElement('div');
    div.textContent = text;
    return div.innerHTML;
}
