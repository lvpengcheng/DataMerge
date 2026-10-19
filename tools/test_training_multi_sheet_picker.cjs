const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

class FormDataStub {
    constructor() { this.items = []; }
    append(key, value) { this.items.push([key, value]); }
    get(key) { return this.items.find(item => item[0] === key)?.[1]; }
    getAll(key) { return this.items.filter(item => item[0] === key).map(item => item[1]); }
}
const escape = value => String(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
const decode = value => value.replace(/&quot;/g, '"').replace(/&#39;/g, "'").replace(/&lt;/g, '<').replace(/&gt;/g, '>').replace(/&amp;/g, '&');
const nodes = new Map();
let checkboxes = [];
const list = {
    set innerHTML(html) {
        this.html = html;
        checkboxes = [...html.matchAll(/<input[^>]*data-file-index="(\d+)" data-name="([^"]*)"[^>]*>/g)].map(match => ({
            checked: true,
            getAttribute(name) { return name === 'data-file-index' ? match[1] : decode(match[2]); },
        }));
    },
};
nodes.set('sheet-picker-list', list);
nodes.set('sheet-picker-modal', {style: {}, querySelector() { return {style: {}}; }});
nodes.set('sheet-picker-title', {});
nodes.set('sheet-picker-confirm', {});
const document = {
    addEventListener() {},
    getElementById(id) { return nodes.get(id); },
    querySelectorAll(selector) { return selector.endsWith(':checked') ? checkboxes.filter(c => c.checked) : checkboxes; },
    createElement() { return {set textContent(value) { this.innerHTML = escape(value); }}; },
};
const alerts = [], requests = [], submissions = [];
let responses;
const ctx = vm.createContext({
    document, window: {}, console: {log() {}, warn() {}}, FormData: FormDataStub,
    getComputedStyle: () => ({}), alert: text => alerts.push(text),
    AUTH: {authFetch: async (url, options) => {
        requests.push(options.body);
        return responses(options.body.get('target_file').name);
    }},
});
vm.runInContext(fs.readFileSync('frontend/static/js/training.js', 'utf8'), ctx);
const flush = () => new Promise(resolve => setImmediate(resolve));
const response = sheets => ({ok: true, json: async () => ({sheets})});
const files = [{name: '员工明细.xls'}, {name: '部门汇总.xlsx'}];
const simple = value => JSON.parse(JSON.stringify(value));

(async () => {
    responses = async () => response([{name: '结果', preview: [['<img src=x>', 0]]}, {name: '说明&"附表', preview: []}]);
    vm.runInContext('_filePasswordsMap = {"员工明细.xls": "first", "部门汇总.xlsx": "second"}', ctx);
    const pending = ctx._pickTargetSheets(files);
    assert.equal(nodes.get('sheet-picker-confirm').disabled, true);
    ctx.closeSheetPicker(false); // Loading must never submit an empty/all-sheets choice.
    await flush();
    assert.equal(checkboxes.length, 4);
    assert.match(list.html, /员工明细.xls/);
    assert.match(list.html, /部门汇总.xlsx/);
    assert.match(list.html, /&lt;img src=x&gt; \| 0/);
    assert.doesNotMatch(list.html, /<img/);
    assert.equal(requests[0].get('file_password'), 'first');
    assert.equal(requests[1].get('file_password'), 'second');
    ctx.sheetPickerToggleAll(false, 1);
    assert.deepEqual(checkboxes.map(c => c.checked), [true, true, false, false]);
    ctx.closeSheetPicker(false);
    assert.match(alerts.at(-1), /部门汇总.xlsx/);
    assert.equal(nodes.get('sheet-picker-modal').style.display, 'flex');
    checkboxes[1].checked = false;
    checkboxes[2].checked = true;
    ctx.closeSheetPicker(false);
    assert.deepEqual(simple(await pending), {'员工明细.xlsx': ['结果'], '部门汇总.xlsx': ['结果']});

    const single = ctx._pickTargetSheets(files[1]);
    await flush();
    ctx.sheetPickerToggleAll(false);
    checkboxes[1].checked = true;
    ctx.closeSheetPicker(false);
    assert.deepEqual(simple(await single), ['说明&"附表']);

    const normalized = ctx._pickTargetSheets([{name: '明细.XLSX'}, {name: '汇总.xlsm'}]);
    await flush();
    ctx.closeSheetPicker(false);
    assert.deepEqual(simple(await normalized), {
        '明细.xlsx': ['结果', '说明&"附表'], '汇总.xlsx': ['结果', '说明&"附表'],
    });

    let completeOld;
    responses = () => new Promise(resolve => { completeOld = resolve; });
    const cancelled = ctx._pickTargetSheets(files[0]);
    ctx.closeSheetPicker(true);
    assert.equal(await cancelled, null);
    responses = async () => response([{name: '新选择', preview: []}]);
    const newer = ctx._pickTargetSheets(files[1]);
    await flush();
    completeOld(response([{name: '旧请求', preview: []}]));
    await flush();
    assert.doesNotMatch(list.html, /旧请求/);
    assert.match(list.html, /新选择/);
    ctx.closeSheetPicker(false);
    assert.deepEqual(simple(await newer), ['新选择']);

    responses = async name => name === files[1].name
        ? {ok: false, status: 400, json: async () => ({detail: '密码错误'})}
        : response([{name: '结果'}]);
    await assert.rejects(ctx._pickTargetSheets(files), /部门汇总.xlsx：密码错误/);
    assert.equal(nodes.get('sheet-picker-modal').style.display, 'none');
    await assert.rejects(ctx._pickTargetSheets([{name: '明细.xls'}, {name: '明细.XLSX'}]), /重名/);

    // Exercise the real start + multipart submission path, rather than only the picker helpers.
    for (const [id, value] of Object.entries({
        'source-files': {files: [{name: '源.xlsx'}]}, 'target-file': {files}, 'rule-files': {files: []},
        'ai-provider': {value: 'openai'}, mode: {value: 'template'}, 'salary-month': {value: ''},
        'standard-hours': {value: ''}, 'manual-headers': {value: ''},
        'multi-sheet-source': {checked: false}, 'attach-popover': {style: {}},
    })) nodes.set(id, value);
    ctx.submissions = submissions;
    vm.runInContext(`_addMessage = () => {}; _setUIStreaming = () => {}; _resetUploadState = () => {};
        _fetchTrainingSSE = (url, options) => submissions.push({url, form: options.body});`, ctx);
    responses = async () => response([{name: '结果'}, {name: '说明'}]);
    const requestCount = requests.length;
    ctx._startTraining('生成双目标');
    ctx._startTraining('重复点击');
    await flush();
    assert.equal(requests.length - requestCount, 2, '重复点击不得启动第二组解析');
    assert.equal(submissions.length, 0);
    checkboxes[1].checked = false;
    checkboxes[2].checked = false;
    ctx.closeSheetPicker(false);
    await flush();
    assert.equal(submissions.length, 1);
    const {form} = submissions[0];
    assert.deepEqual(form.getAll('target_files'), files);
    assert.deepEqual(JSON.parse(form.get('target_sheets')), {'员工明细.xlsx': ['结果'], '部门汇总.xlsx': ['说明']});
    responses = async () => { throw new Error('解析失败'); };
    ctx._startTraining('失败不应训练');
    await flush();
    assert.equal(submissions.length, 1, '解析失败不能默认处理全部 Sheet');
    console.log('PASS: grouped multi-file selection, independent same-name sheets, passwords, per-file validation, escaping, legacy single-file payload, loading/cancellation, stale results, duplicate guard and training multipart submission');
})().catch(error => { console.error(error); process.exitCode = 1; });
