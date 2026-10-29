const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('frontend/static/js/training.js', 'utf8');
const start = source.indexOf('async function downloadCode()');
const end = source.indexOf('// ==================== 输入处理', start);
let contentType, requested, downloaded;
const context = vm.createContext({
    _currentSessionId: 7, _currentTenantId: '工资',
    AUTH: {authFetch: async url => {
        requested = url;
        return {ok: true, blob: async () => Buffer.from('content'), headers: {get: () => contentType}};
    }},
    URL: {createObjectURL: () => 'blob:test', revokeObjectURL() {}},
    document: {createElement: () => ({click() {downloaded = this.download;}}),
        body: {appendChild() {}, removeChild() {}}},
    alert: message => {throw Error(message);},
});
vm.runInContext(source.slice(start, end), context);
(async () => {
    contentType = 'text/x-python; charset=utf-8';
    await context.downloadCode();
    assert.equal(requested, '/api/training/chat/sessions/7/code?download=true');
    assert.equal(downloaded, 'script_工资.py');
    contentType = 'application/zip';
    await context.downloadCode();
    assert.equal(downloaded, 'script_工资.py');
    const compute = fs.readFileSync('frontend/static/js/compute.js', 'utf8');
    vm.runInContext(compute.slice(compute.indexOf('function resultDisplayName('), compute.indexOf('async function downloadResult(')), context);
    assert.equal(context.resultDisplayName('工资 · 方案 abc012345678_202610_20261028102030_012abc34_原版.zip'),
        '工资_202610_20261028102030_原版.zip');
    const solutions = fs.readFileSync('frontend/static/js/solutions.js', 'utf8');
    assert.ok(!solutions.includes("button('验收'"));
    assert.ok(!solutions.includes('确认接受差异并启用'));
    console.log('Single/multi script downloads and hidden solution labels passed');
})().catch(error => {console.error(error); process.exitCode = 1;});
