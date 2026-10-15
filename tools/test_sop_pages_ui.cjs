// Offline checks for separate SOP pages and user selection. No server or mail calls.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

const operationHtml = fs.readFileSync('frontend/templates/sop.html', 'utf8');
const adminHtml = fs.readFileSync('frontend/templates/sop_admin.html', 'utf8');
const importHtml = fs.readFileSync('frontend/templates/sop_import.html', 'utf8');
assert.ok(!operationHtml.includes('id="sop-import-form"'));
assert.ok(!operationHtml.includes('id="sop-reviewer-form"'));
assert.ok(!adminHtml.includes('id="sop-import-form"'));
assert.ok(importHtml.includes('id="sop-import-form"'));
assert.ok(!adminHtml.includes('id="sop-table"'));
assert.match(operationHtml, /href="\/sop\/admin"[^>]*data-perm="admin"/);

function environment(html, admin) {
    const elements = new Map([...html.matchAll(/id="([^"]+)"/g)].map(match => [match[1], {
        style: {}, value: '', innerHTML: '', textContent: '', disabled: false,
    }]));
    const callbacks = [];
    const requests = [];
    let config = {user_id: 42, reviewer_valid: true, reviewer_name: '复审人', users: [
        {id: 42, username: 'reviewer', display_name: '复审人', email: 'reviewer@example.com', selectable: true},
        {id: 43, username: '<inactive>', email: null, selectable: false, disabled_reason: '账号已停用、未填写邮箱'},
    ]};
    const context = vm.createContext({
        console, URLSearchParams,
        FormData: function () {},
        document: {
            getElementById: id => elements.get(id) || null,
            querySelector: () => ({}),
            addEventListener: (name, callback) => callbacks.push(callback),
        },
        AUTH: {
            requireAuth: () => true, hasPerm: () => true,
            renderUserInfo() {}, isAdmin: () => admin,
            async authFetch(url, options) {
                requests.push({url, options});
                const data = url.endsWith('/reviewer-config') ? config
                    : url.endsWith('/setup-status') ? {public_url_configured: true, smtp_configured: true,
                        reviewer_configured: true, reviewer_email_configured: true, reviewer_can_review: true}
                    : url.endsWith('/import') ? {created: 1, updated: 0, errors: []}
                    : {items: []};
                return {ok: true, async json() { return data; }};
            },
        },
    });
    vm.runInContext(fs.readFileSync('frontend/static/js/sop.js', 'utf8'), context);
    if (html === adminHtml) vm.runInContext(fs.readFileSync('frontend/static/js/sop_admin.js', 'utf8'), context);
    if (html === importHtml) vm.runInContext(fs.readFileSync('frontend/static/js/sop_import.js', 'utf8'), context);
    return {context, elements, requests, callbacks, setConfig: value => { config = value; }};
}

(async () => {
    const managed = environment(adminHtml, true);
    const validate = inputs => vm.runInContext('SOP.validateUploadFiles', managed.context)({querySelectorAll: () => inputs});
    for (const extension of ['xlsx', 'xls', 'txt', 'md', 'markdown', 'pdf', 'doc', 'docx']) {
        validate([{name: 'source_files', files: []}, {name: 'result_file', files: []},
            {name: 'rule_file', files: [{name: 'rules.' + extension}]}]);
    }
    assert.throws(() => validate([{name: 'source_files', files: [{name: 'source.pdf'}]}]), /xlsx/);
    assert.throws(() => validate([{name: 'rule_file', files: [{name: 'rules.zip'}]}]), /规则文件/);
    // Loading the shared business script on the admin page must make no business requests.
    managed.callbacks[0]();
    assert.equal(managed.requests.length, 0);
    await vm.runInContext('SOP_ADMIN.init()', managed.context);
    const rules = vm.runInContext(`SOP_ADMIN.currentRules([
        {id:1,scope:'global',is_active:false},
        {id:2,scope:'global',is_active:true},
        {id:3,scope:'customer',customer_name:'A',is_active:false},
        {id:4,scope:'customer',customer_name:'A',is_active:true}
    ])`, managed.context);
    assert.deepEqual(Array.from(rules, rule => rule.id), [2,4]);
    const select = managed.elements.get('sop-reviewer-user');
    assert.match(select.innerHTML, /value="42" selected/);
    assert.match(select.innerHTML, /value="43" disabled/);
    assert.ok(select.innerHTML.includes('&lt;inactive&gt;'));
    assert.ok(select.innerHTML.includes('账号已停用、未填写邮箱'));
    assert.equal(managed.elements.get('sop-reviewer-save').disabled, false);
    select.selectedOptions = [{value: '42'}, {value: '44'}];
    await managed.elements.get('sop-reviewer-form').onsubmit({preventDefault() {}});
    const saved = managed.requests.find(request => request.options?.method === 'PUT');
    assert.deepEqual(JSON.parse(saved.options.body).user_ids, [42,44]);
    const importing = environment(importHtml, true);
    await vm.runInContext('SOP_IMPORT.init()', importing.context);
    await importing.elements.get('sop-import-form').onsubmit({preventDefault() {}, currentTarget: {}});
    assert.match(importing.elements.get('sop-import-result').textContent, /新增 1/);
    assert.ok(!managed.requests.some(request => request.url.includes('/entries')));

    managed.setConfig({user_id: null, reviewer_valid: false, users: [
        {id: 43, username: 'inactive', selectable: false, disabled_reason: '账号已停用'},
    ]});
    await vm.runInContext('SOP_ADMIN.reviewerConfig()', managed.context);
    assert.equal(managed.elements.get('sop-reviewer-save').disabled, true);
    assert.match(managed.elements.get('sop-reviewer-status').textContent, /暂无符合条件/);

    const forbidden = environment(adminHtml, false);
    await vm.runInContext('SOP_ADMIN.init()', forbidden.context);
    assert.equal(forbidden.requests.length, 0);
    assert.match(forbidden.elements.get('sop-error').textContent, /仅系统管理员/);
    const forbiddenImport = environment(importHtml, false);
    await vm.runInContext('SOP_IMPORT.init()', forbiddenImport.context);
    assert.equal(forbiddenImport.requests.length, 0);
    assert.match(forbiddenImport.elements.get('sop-error').textContent, /仅系统管理员/);
    console.log('SOP page separation, user selection, saving, import and admin access checks passed.');
})().catch(error => { console.error(error); process.exitCode = 1; });
