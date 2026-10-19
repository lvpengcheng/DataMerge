const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const nodes = new Map();
const doc = {
    addEventListener() {},
    getElementById(id) { return nodes.get(id) || null; },
    createElement() {
        return {
            style: {}, targets: [],
            set innerHTML(html) {
                this.html = html;
                for (const id of ['_pre_confirm', '_pre_cancel', '_pre_validation_error']) {
                    nodes.set(id, {style: {}, disabled: false, textContent: ''});
                }
                this.targets = [...html.matchAll(/<select data-target-key="([^"]+)"[^>]*>([\s\S]*?)<\/select>/g)].map(match => {
                    const options = [...match[2].matchAll(/<option value="([^"]*)"([^>]*)>/g)];
                    return {dataset: {targetKey: match[1]}, value: options.find(o => /\bselected\b/.test(o[2]))?.[1] ?? options[0][1],
                        matches(selector) { return selector === 'select[data-target-key]'; }};
                });
            },
            querySelectorAll(selector) { return selector === 'select[data-target-key]' ? this.targets : []; },
            remove() { nodes.delete(this.id); },
        };
    },
    body: {
        appendChild(el) { nodes.set(el.id, el); el.parentNode = this; },
        removeChild(el) { el.remove(); },
    },
};
const ctx = vm.createContext({document: doc, setTimeout, clearTimeout, console});
vm.runInContext(fs.readFileSync('frontend/static/js/compute.js', 'utf8'), ctx);
(async () => {
    const pending = ctx._showPrecheckDialog({has_review_items: true, target_confirmation_required: true,
        target_candidates: [
            {key: '当月1', selected_sheet: '当月1', all_sheets: ['当月1', '当月2'], status: 'matched'},
            {key: '当月2', selected_sheet: '当月2', all_sheets: ['当月1', '当月2'], status: 'matched'},
            {key: '当月3', selected_sheet: null, all_sheets: ['当月1', '当月2'], status: 'missing'},
        ], source_mapping_overview: [{file: '上传.xlsx', sheet: '9月', expected_file: '智训.xlsx', expected_sheet: '上月1'}],
    });
    const overlay = doc.getElementById('_compute_precheck_overlay');
    assert.equal(overlay.targets.length, 3, '所有关键 Sheet 都显示，包括同名匹配和完全缺失的表');
    assert.match(overlay.html, /上传.xlsx.*9月.*智训.xlsx.*上月1/);
    assert.equal(overlay.targets[2].value, '__pending_target__', '未找到不能默认等同于跳过');
    const confirm = doc.getElementById('_pre_confirm');
    assert.equal(confirm.disabled, true);
    assert.throws(() => ctx._collectConfirmedTargetMap(overlay), /全部关键 Sheet/);
    overlay.targets[2].value = '当月1';
    overlay.onchange({target: overlay.targets[2]});
    assert.equal(confirm.disabled, true, '同一个模板 Sheet 不能同时分给两个目标');
    overlay.targets[2].value = '';
    overlay.onchange({target: overlay.targets[2]});
    assert.equal(confirm.disabled, false, '明确跳过后才可以确认计算');
    await confirm.onclick();
    const result = await pending;
    assert.equal(result.mapping_finalized, true);
    assert.deepEqual(JSON.parse(JSON.stringify(result.confirmed_target_map)), {'当月1': '当月1', '当月2': '当月2', '当月3': ''});
    console.log('PASS: complete target review, source mapping visibility, pending/skip distinction, duplicate guard and final payload');
    const logs = [];
    nodes.set('log-content', {appendChild(el) { logs.push(el.html); }});
    nodes.set('result-card', {});
    ctx.addLog('error', "expected <class 'float'>");
    ctx.showError("expected <class 'float'>");
    assert.match(logs[0], /expected &lt;class &#39;float&#39;&gt;/);
    assert.match(nodes.get('result-card').innerHTML, /expected &lt;class &#39;float&#39;&gt;/);
    console.log('PASS: exception type text remains visible in logs and error result');
})().catch(error => {console.error(error); process.exitCode = 1;});
