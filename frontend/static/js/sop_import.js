/** 独立客户清单导入页。 */
const SOP_IMPORT = {
    request: (url, options) => SOP.request(url, options),
    error: message => SOP.error(message),
    async init() {
        if (!AUTH.requireAuth()) return;
        AUTH.renderUserInfo(document.querySelector('header'));
        if (!AUTH.isAdmin()) { this.error('仅系统管理员可以导入客户清单。'); return; }
        document.getElementById('sop-import-panel').style.display = '';
        document.getElementById('sop-import-form').onsubmit = async event => {
            event.preventDefault();
            try {
                const data = await (await this.request('/api/tools/sop/import', {method:'POST',body:new FormData(event.currentTarget)})).json();
                document.getElementById('sop-import-result').textContent = `新增 ${data.created}，更新 ${data.updated}${data.errors.length ? '\n' + data.errors.join('\n') : ''}`;
            } catch (e) { this.error(e.message); }
        };
    },
};
document.addEventListener('DOMContentLoaded', () => SOP_IMPORT.init());
