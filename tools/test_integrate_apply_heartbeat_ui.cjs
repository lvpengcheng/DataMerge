const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const source = fs.readFileSync(path.join(__dirname, '../frontend/static/js/tools.js'), 'utf8');
const method = source.slice(source.indexOf('    async _intFetch('), source.indexOf('    _intPerm('));
let requested;
let payload = {ok: true};
const ctx = vm.createContext({Response, AUTH: {authFetch: async (url) => {
    requested = url;
    return new Response('   \n' + JSON.stringify(payload), {headers: {'Content-Type': 'application/json'}});
}}});
vm.runInContext(`var tools = {${method}};`, ctx);
(async () => {
    for (const kind of ['integrate', 'split']) {
        ctx.tools._integrationKind = kind;
        for (const suffix of ['scheme/apply-validate', 'reparse-headers', 'execute']) {
            payload = {ok: true};
            const response = await ctx.tools._intFetch('/api/tools/integrate/' + suffix);
            assert.equal(response.ok, true);
            assert.deepEqual(await response.json(), {ok: true});
            assert.equal(requested, `/api/tools/${kind === 'split' ? 'table-split' : 'integrate'}/${suffix}?stream=true`);
            payload = {error: '无权应用方案', status_code: 403};
            const failed = await ctx.tools._intFetch('/api/tools/integrate/' + suffix);
            assert.equal(failed.status, 403);
            assert.deepEqual(await failed.json(), {detail: '无权应用方案'});
        }
        await ctx.tools._intFetch('/api/tools/integrate/result?session_id=test&token=test');
        assert.ok(!requested.includes('stream=true'));
    }
    const executeMethod = source.slice(source.indexOf('    async _doIntegrateExecute()'), source.indexOf('    initDataMerge()'));
    vm.runInContext(`Object.assign(tools, {${executeMethod}}); var _intMode = 'create'; var _integrateAnalysis = {session_id: 'test'};`, ctx);
    const elements = {'int-exec-status': {}, 'int-execute': {}};
    let downloaded, clicked;
    ctx.document = {
        querySelectorAll: () => [], querySelector: () => ({value: '1'}),
        getElementById: id => elements[id] || {},
        createElement: () => ({click() {clicked = true;}, remove() {}}),
        body: {appendChild() {}},
    };
    ctx.URL = {createObjectURL: blob => {downloaded = blob; return 'blob:test';}, revokeObjectURL() {}};
    Object.assign(ctx.tools, {_intMainFile: () => 'main.xlsx', _readSectionPairs: () => [{a_col: '金额'}],
        _setIntegrateStatus() {}, _readSplitConfig: () => ({})});
    for (const kind of ['integrate', 'split']) {
        ctx.tools._integrationKind = kind;
        ctx.AUTH.authFetch = async url => url.includes('/execute?stream=true')
            ? new Response('  ' + JSON.stringify({download_url: '/api/tools/integrate/result?token=ready'}),
                {headers: {'Content-Type': 'application/json'}})
            : new Response('PK-binary', {headers: {'Content-Type': 'application/zip', 'X-Integrate-Matched': '12'}});
        clicked = false;
        await ctx.tools._doIntegrateExecute();
        assert.equal(await downloaded.text(), 'PK-binary');
        assert.ok(clicked);
        assert.ok(elements['int-exec-status'].textContent.includes('命中12行'));
        assert.equal(elements['int-execute'].disabled, false);
    }
    console.log('PASS: both workflows request heartbeats, consume JSON, preserve errors and binary downloads');
})().catch(error => {console.error(error); process.exitCode = 1;});
