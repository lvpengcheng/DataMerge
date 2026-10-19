const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('frontend/static/js/compute.js', 'utf8');
const start = source.indexOf('async function startCompute()');
const end = source.indexOf('\nfunction ', start + 1);
const elements = {
    'compute-btn': {}, 'source-files': {files: [new Blob(['source'])]},
    'template-file': {files: []}, 'salary-month': {value: '2026-09'},
    'standard-hours': {value: ''},
};
let submitted;
const context = vm.createContext({
    console, FormData, Date, Math, alert: assert.fail,
    document: {getElementById: id => elements[id]},
    currentTenantId: 'tenant', currentScriptId: 'script_c7d1ccba20f7',
    scriptsCache: [{script_id: 'script_c7d1ccba20f7', name: '工资计算'}],
    _encryptionCheckInProgress: false, _filePasswordsMap: null,
    _autoCheckEncryption: async () => true, clearResult() {}, addLog() {},
    updateStatus() {}, updateProgress() {}, _closePrecheckDialog() {},
    _readComputeSubmitJson: async () => ({task_id: '7'}), _saveActiveTask() {},
    _connectComputeStream() {},
    AUTH: {async authFetch(url, options) {
        submitted = options.body;
        return {ok: true, status: 200};
    }},
});
vm.runInContext(source.slice(start, end < 0 ? undefined : end), context);
(async () => {
    await context.startCompute();
    assert.equal(submitted.get('script_display_name'), '工资计算');
    assert.equal(submitted.get('script_id'), 'script_c7d1ccba20f7');
    assert.equal(submitted.get('salary_month'), '9');
    console.log('Selected script name is submitted separately from its internal ID');
})().catch(error => {console.error(error); process.exitCode = 1;});
