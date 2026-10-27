const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('frontend/static/js/compute.js', 'utf8');
const elements = new Map();
function element() {
    return {style: {}, disabled: false, textContent: '', value: '', children: [],
        appendChild(child) {this.children.push(child);},
        insertBefore(child) {this.children.unshift(child);},
        set innerHTML(value) {this.children = []; this.html = value;},
        get innerHTML() {return this.html;}};
}
const logs = [];
let payload;
let fetcher = async () => ({ok: true, json: async () => payload});
const ctx = vm.createContext({console, encodeURIComponent,
    document: {getElementById(id) {if (!elements.has(id)) elements.set(id, element()); return elements.get(id);},
        createElement: element},
    AUTH: {authFetch: (...args) => fetcher(...args)},
    toggleTemplateInput() {}, checkCanCompute() {}, updateScriptInfo() {},
    addLog(level, text) {logs.push(text);},
});
vm.runInContext("let currentScriptId = 'old'; let scriptsCache = []; let _tenantScriptsLoadId = 0;\n" +
    source.slice(source.indexOf('async function loadTenantScripts('), source.indexOf('function onScriptSelected()')), ctx);
(async () => {
    payload = {scripts: [], pending_solutions: [{name: '汇总'}]};
    await ctx.loadTenantScripts('山东太保');
    assert.equal(vm.runInContext('currentScriptId', ctx), '');
    assert.equal(elements.get('script-selector').disabled, true);
    assert.match(elements.get('script-info').textContent, /汇总.*验收并启用/);
    assert.equal(elements.get('script-info').children[0].href, '/training');
    assert.ok(!logs.some(text => text.includes('没有可用')));
    payload = {scripts: [{script_id: 'solution_new', name: '汇总', is_active: true}], pending_solutions: []};
    await ctx.loadTenantScripts('山东太保');
    assert.equal(vm.runInContext('currentScriptId', ctx), 'solution_new');
    assert.equal(elements.get('script-selector').disabled, false);
    payload = {scripts: [], unavailable_solutions: [{name: '缺失方案'}]};
    await ctx.loadTenantScripts('山东太保');
    assert.equal(elements.get('script-selector-group').style.display, 'none');
    assert.equal(elements.get('script-selector').disabled, true);
    assert.equal(elements.get('script-info').textContent, '');
    assert.ok(!logs.some(text => /方案文件缺失|恢复完整方案|重新导入原方案/.test(text)));
    payload = {scripts: [{script_id: 'legacy', name: '旧脚本', is_active: true}],
        unavailable_solutions: [{name: '缺失方案'}], requires_script_selection: true};
    await ctx.loadTenantScripts('山东太保');
    assert.equal(vm.runInContext('currentScriptId', ctx), 'legacy', 'unavailable packages must not block a listed script');
    assert.equal(elements.get('script-selector').value, 'legacy');
    assert.equal(elements.get('script-selector').disabled, false);
    assert.ok(!elements.get('script-selector').innerHTML.includes('原方案不可用'));
    let resolveSlow;
    fetcher = () => new Promise(resolve => {resolveSlow = resolve;});
    const slow = ctx.loadTenantScripts('old-tenant');
    fetcher = async () => ({ok: true, json: async () => ({scripts: [], pending_solutions: []})});
    await ctx.loadTenantScripts('empty-tenant');
    resolveSlow({ok: true, json: async () => payload});
    await slow;
    assert.equal(vm.runInContext('currentScriptId', ctx), '', 'stale responses must not restore a previous tenant script');
    console.log('Imported scheme status, activation visibility and tenant switching checks passed');
})().catch(error => {console.error(error); process.exitCode = 1;});
