// Offline regression checks: loading a scheme must precede explicit execution.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');

function environment() {
    const elements = new Map();
    const element = id => {
        if (!elements.has(id)) elements.set(id, {disabled: false, checked: false, style: {}, textContent: '', innerHTML: ''});
        return elements.get(id);
    };
    const calls = [], filled = [];
    const analysis = {session_id: 'test', files: [
        {name: 'main.xlsx', fingerprint: 'auto-main', header_start: 1, header_end: 1},
        {name: 'source.xlsx', fingerprint: 'source', header_start: 1, header_end: 1},
    ], matched_schemes: []};
    let validation = {ok: true, resolved_scheme: {id: 12, config: {}}};
    const context = vm.createContext({
        console, FormData: class {append() {}},
        document: {addEventListener() {}, getElementById: element, querySelectorAll: () => []},
        AUTH: {async authFetch(url, options) {
            calls.push({url, body: options?.body});
            const data = url.endsWith('/apply-validate') ? validation : structuredClone(analysis);
            if (url.endsWith('/reparse-headers')) {
                const ranges = JSON.parse(options.body).header_ranges;
                data.files.forEach(f => Object.assign(f, {header_start: ranges[f.name].start_row, header_end: ranges[f.name].end_row}));
            }
            return {ok: true, async json() {return data;}};
        }},
    });
    vm.runInContext(fs.readFileSync('frontend/static/js/tools.js', 'utf8') + '\nthis.tools = Tools;', context);
    const tools = context.tools;
    tools._renderIntegrateConfig = () => {vm.runInContext('_intApplyReady = false', context);};
    tools._intFillFromScheme = scheme => filled.push(scheme);
    tools._intConfirmColumnMappings = async () => true;
    tools._intAskMatchMapping = async () => true;
    vm.runInContext(`
        _intMode = 'apply'; _intApplyTarget = {id: 12, name: 'test'};
        _integrateFiles = [{name: 'main.xlsx'}, {name: 'source.xlsx'}];
        _integrateSchemeList = [{id: 12, config: {
            roles: [{file: 'main.xlsx', fp: 'main'}, {file: 'source.xlsx', fp: 'source'}],
            header_ranges_by_role: {'0': [1, 3], '1': [1, 1]}
        }}];
    `, context);
    return {tools, context, calls, filled, element, setValidation: v => {validation = v;}};
}

(async () => {
    const env = environment();
    await env.tools._analyzeIntegrate();
    assert.deepEqual(env.calls.map(c => c.url.split('/').pop()), ['analyze', 'reparse-headers', 'apply-validate']);
    assert.equal(JSON.parse(env.calls[1].body).header_ranges['main.xlsx'].end_row, 3);
    assert.equal(env.filled.length, 1, 'analysis must load the selected scheme');
    assert.equal(env.element('int-execute').disabled, false);
    assert.equal(env.calls.some(c => c.url.endsWith('/execute')), false, 'loading must never generate a download');

    env.context.document.querySelectorAll = () => env.context.headerRows;
    env.context.headerRows = ['main.xlsx', 'source.xlsx'].map(name => ({
        dataset: {file: name}, querySelector: selector => ({value: selector.includes('start') ? '1' : '2'}),
    }));
    env.calls.length = 0;
    await env.tools._intReparseHeaders();
    assert.deepEqual(env.calls.map(c => c.url.split('/').pop()), ['reparse-headers', 'apply-validate']);
    assert.equal(JSON.parse(env.calls[0].body).header_ranges['main.xlsx'].end_row, 2, 'manual range must be preserved');
    assert.equal(env.filled.length, 2, 'manual reparse must reload scheme mappings');

    env.setValidation({ok: false, reasons: ['missing required column']});
    await env.tools._intPrefillForApply();
    assert.equal(env.element('int-execute').disabled, true);
    assert.match(env.element('int-apply-reasons').innerHTML, /missing required column/);
    env.calls.length = 0;
    await env.tools._doIntegrateExecute();
    assert.equal(env.calls.length, 0, 'failed validation must block execution');

    env.setValidation({ok: true, requires_confirmation: true, mapping_suggestions: [], resolved_scheme: {id: 12}});
    env.tools._intConfirmColumnMappings = async () => false;
    await env.tools._intPrefillForApply();
    assert.equal(env.element('int-execute').disabled, true);
    assert.equal(env.filled.length, 2, 'cancelled column mapping must not be applied');

    env.setValidation({ok: true, resolved_scheme: {id: 12}});
    vm.runInContext('_integrateAnalysis.matched_schemes = [{id: 12, ambiguous: [{candidates: ["a", "b"]}]}]', env.context);
    env.tools._intAskMatchMapping = async () => false;
    await env.tools._intPrefillForApply();
    assert.equal(env.element('int-execute').disabled, true);
    assert.equal(env.filled.length, 2, 'cancelled role assignment must not be applied');
    console.log('Integration scheme prefill, header restoration, reparse and validation checks passed');
})().catch(error => {console.error(error); process.exitCode = 1;});
