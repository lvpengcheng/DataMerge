// Offline UI contract checks; no server, account, or AI requests.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
class Element {
    constructor(tagName) { this.tagName = tagName; this.dataset = {}; this.value = ''; this.files = []; this.checked = false; this.style = {}; this.children = []; this.options = []; }
    querySelector(tagName) { return this.children.find(child => child.tagName === tagName); }
    append(...children) { this.children.push(...children); }
    replaceChildren(...children) { this.children = children; }
    add(option) { this.options.push(option); }
}
const fields = new Map();
const field = id => { if (!fields.has(id)) fields.set(id, new Element()); return fields.get(id); };
const settings = {ai_provider: 'deepseek', mode: 'formula', salary_year: 2026, salary_month: 9,
    monthly_standard_hours: 174, manual_headers: null, multi_sheet_source: false, use_history: false};
let request, calls = 0, fail = false, resetCount = 0;
const context = vm.createContext({console, FormData, File, Option: function(label, value) { this.label = label; this.value = value; },
    _currentSessionId: 7, _currentAccuracy: 1, _isStreaming: false, _filePasswordsMap: {},
    document: {getElementById: field, createElement: tag => new Element(tag), createTextNode: text => ({textContent: text})},
    onModeChange() {}, toggleAttachPopover() {}, _setUIStreaming() {}, alert() {},
    _downloadOriginalFile(...args) { context.download = args; },
    _resetUploadState() { resetCount++; for (const id of ['source-files', 'target-file', 'rule-files']) field(id).files = []; },
    AUTH: {async authFetch(url, options) {
        calls++; request = {url, options};
        return {ok: !fail, async json() { return fail ? {detail: 'bad workbook'} : {
            settings: {...settings, ai_provider: 'claude'}, source_file_names: ['old.xlsx', 'new.xlsx'],
            expected_file_name: 'target.xlsx', rule_file_names: ['SOP.docx'], validation_stale: true}; }};
    }},
});
vm.runInContext(fs.readFileSync('frontend/static/js/training_assets.js', 'utf8'), context);
(async () => {
    context._restoreSessionSettings(settings);
    assert.equal(field('ai-provider').value, 'deepseek');
    assert.equal(field('salary-month').value, '2026-09');
    field('ai-provider').value = 'claude';
    field('source-update-mode').value = 'merge';
    field('source-files').files = [new File(['excel'], 'new.xlsx')];
    assert.equal(await context.saveSessionAssets(), true);
    assert.equal(request.url, '/api/training/chat/sessions/7/assets');
    assert.deepEqual(JSON.parse(request.options.body.get('settings')), {ai_provider: 'claude'});
    assert.equal(request.options.body.get('source_mode'), 'merge');
    assert.equal(request.options.body.getAll('source_files')[0].name, 'new.xlsx');
    assert.equal(resetCount, 1);
    assert.equal(context._currentAccuracy, null);
    assert.equal(await context.saveSessionAssets(), true);
    assert.equal(calls, 1, 'unchanged settings must not be submitted repeatedly');
    const panel = field('session-assets');
    assert.equal(panel.querySelector('details').open, false);
    panel.querySelector('details').open = true;
    context._renderSessionAssets({source_file_names: ['old.xlsx']});
    assert.equal(panel.querySelector('details').open, true);
    const links = panel.querySelector('details').children.flatMap(row => row.children || []).filter(child => child.onclick);
    context._currentSessionId = 8;
    links.find(link => link.textContent === 'old.xlsx').onclick({preventDefault() {}});
    assert.deepEqual(context.download, [7, 'source', 'old.xlsx']);
    context._renderSessionAssets({source_file_names: ['other.xlsx']});
    assert.equal(panel.querySelector('details').open, false);
    fail = true;
    field('source-files').files = [new File(['bad'], 'bad.xlsx')];
    assert.equal(await context.saveSessionAssets(), false);
    assert.equal(field('source-files').files[0].name, 'bad.xlsx', 'failed updates keep selected files for correction');
    assert.equal(resetCount, 1);
    console.log('Training assets UI: restore, merge upload, provider switch, scoped download, retry preservation passed.');
})().catch(error => { console.error(error); process.exitCode = 1; });
