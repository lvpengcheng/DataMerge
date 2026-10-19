const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const {File} = require('node:buffer');
const source = fs.readFileSync('frontend/static/js/compute.js', 'utf8');
const templates = [new File(['detail'], '明细.xlsx'), new File(['summary'], '汇总.xlsx')];
const elements = {
    'compute-btn': {}, 'source-files': {files: [new File(['source'], '源.xlsx')]},
    'template-file': {files: templates}, 'salary-month': {value: ''}, 'standard-hours': {value: ''},
};
const submitted = [];
const ctx = vm.createContext({
    console, FormData, Date, Math, alert: assert.fail,
    document: {getElementById: id => elements[id]},
    currentTenantId: 'tenant', currentScriptId: 'script_1',
    scriptsCache: [{script_id: 'script_1', name: '双目标', mode: 'template'}],
    _encryptionCheckInProgress: false,
    _filePasswordsMap: {'template::明细.xlsx': 'one', 'template::汇总.xlsx': 'two'},
    _autoCheckEncryption: async () => true, clearResult() {}, addLog() {},
    updateStatus() {}, updateProgress() {}, _closePrecheckDialog() {},
    _readComputeSubmitJson: async () => ({task_id: '7'}), _saveActiveTask() {}, _connectComputeStream() {},
    AUTH: {async authFetch(url, options) { submitted.push(options.body); return {ok: true, status: 200}; }},
});
const start = source.indexOf('async function startCompute()');
const end = source.indexOf('\nfunction ', start + 1);
vm.runInContext(source.slice(start, end < 0 ? undefined : end), ctx);
const validationStart = source.indexOf('function _targetMappingIssue(');
const validationEnd = source.indexOf('\n}', validationStart) + 2;
vm.runInContext(source.slice(validationStart, validationEnd), ctx);
(async () => {
    const html = fs.readFileSync('frontend/templates/compute.html', 'utf8');
    assert.match(html, /<input[^>]*id="template-file"[^>]*multiple/);
    await ctx.startCompute();
    assert.deepEqual(submitted[0].getAll('template_files').map(f => f.name), ['明细.xlsx', '汇总.xlsx']);
    assert.equal(submitted[0].get('template_file'), null);
    assert.equal(submitted[0].getAll('source_files').length, 1);
    assert.deepEqual(JSON.parse(submitted[0].get('file_passwords')), {'template::明细.xlsx': 'one', 'template::汇总.xlsx': 'two'});
    elements['template-file'].files = [];
    await ctx.startCompute();
    assert.equal(submitted[1].getAll('template_files').length, 0, '不上传时复用智训模板');
    const selects = [
        {dataset: {targetWorkbook: '明细.xlsx'}, value: '结果'},
        {dataset: {targetWorkbook: '汇总.xlsx'}, value: '结果'},
    ];
    const overlay = {querySelectorAll: () => selects};
    assert.equal(ctx._targetMappingIssue(overlay), '', '不同文件允许同名 Sheet');
    selects[1].dataset.targetWorkbook = '明细.xlsx';
    assert.match(ctx._targetMappingIssue(overlay), /同一文件/);
    selects[1].value = '';
    assert.equal(ctx._targetMappingIssue(overlay), '', '明确跳过不占用 Sheet');
    console.log('PASS: all templates and separate passwords submitted, no-template fallback, workbook-scoped same-sheet confirmation and per-file duplicate protection');
})().catch(error => { console.error(error); process.exitCode = 1; });
