const fs = require('fs');
const vm = require('vm');
const assert = require('node:assert/strict');
const path = require('path');

const nodes = {
    'manual-headers': {value: ''}, 'multi-sheet-source': {checked: false},
    'use-history': {checked: false}, 'attach-popover': {style: {}},
    'source-files': {files: [{name: 'source.xlsx'}]},
    'target-file': {files: [{name: '明细.xlsx'}, {name: '汇总.xlsx'}]},
    'rule-files': {files: []}, 'ai-provider': {value: 'deepseek'},
    'mode': {value: 'template'}, 'salary-month': {value: '2026-10'},
    'standard-hours': {value: '174'},
};
class Form {
    constructor() { this.entries = []; }
    append(key, value) { this.entries.push([key, value]); }
}
const context = {document: {addEventListener() {}, getElementById: id => nodes[id]},
    window: {}, console, setTimeout, clearTimeout, FormData: Form,
    alert: message => { throw new Error(message); }};
vm.createContext(context);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../frontend/static/js/training.js'), 'utf8'), context);
vm.runInContext(`_currentTenantId = 'GWS';
    _addMessage = () => {}; _setUIStreaming = () => {}; _resetUploadState = () => {};
    _fetchTrainingSSE = (url, options) => { globalThis.sent = options.body; };
    _submitStartTraining('按两个文件分别输出', [], [{name:'明细.xlsx'}, {name:'汇总.xlsx'}], [],
        'deepseek', 'template', '2026-10', '174', {'明细.xlsx':['明细'], '汇总.xlsx':['汇总']});`, context);
assert.equal(context.sent.entries.filter(([key]) => key === 'target_files').length, 2);
assert.equal(context.sent.entries.find(([key]) => key === 'message')[1], '按两个文件分别输出');
assert.deepEqual(JSON.parse(context.sent.entries.find(([key]) => key === 'target_sheets')[1]),
    {'明细.xlsx': ['明细'], '汇总.xlsx': ['汇总']});

(async () => {
    vm.runInContext(`globalThis.picked = [];
        _pickTargetSheets = async file => { picked.push(file.name); return ['目标']; };
        _submitStartTraining = (...args) => { globalThis.submitted = args; };
        _startTraining('rules');`, context);
    await new Promise(resolve => setImmediate(resolve));
    assert.deepEqual(Array.from(context.picked), ['明细.xlsx', '汇总.xlsx']);
    assert.equal(context.submitted[2].length, 2);
    assert.deepEqual(JSON.parse(JSON.stringify(context.submitted[8])), {'明细.xlsx':['目标'], '汇总.xlsx':['目标']});
    const elements = {'result-card': {}, 'result-downloads': {}};
    const compute = {document: {addEventListener() {}, getElementById: id => elements[id]}, setTimeout, clearTimeout};
    vm.createContext(compute);
    vm.runInContext(fs.readFileSync(path.join(__dirname, '../frontend/static/js/compute.js'), 'utf8'), compute);
    compute.showResult({output_files: ['结果.zip'], rows_processed: 4});
    assert.ok(elements['result-downloads'].innerHTML.includes('下载ZIP交付包'));
    assert.ok(!elements['result-downloads'].innerHTML.includes('下载全部'));
    console.log('PASS: multiple target upload, per-workbook sheet picker, ZIP download display');
})().catch(error => { console.error(error); process.exitCode = 1; });
