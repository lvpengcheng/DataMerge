const fs = require('fs');
const vm = require('vm');
const assert = require('node:assert/strict');
const path = require('path');
const ctx = {document: {addEventListener() {}}, setTimeout, clearTimeout};
vm.createContext(ctx);
vm.runInContext(fs.readFileSync(path.join(__dirname, '../frontend/static/js/compute.js'), 'utf8'), ctx);
const plain = x => JSON.parse(JSON.stringify(x));
function mapping(file, sheet, target, columns) {
    return {[file]: {expected_file: 'trained.xlsx', sheet_mapping: {[sheet]: target}, header_mapping_by_sheet: {[sheet]: columns}}};
}

const first = mapping('a.xlsx', '工资', '工资表', {'编号': '工号', '金额': '工资'});
{
    const rows = [
        {expected_path:'上月.xlsx > 第一批 > 金额', suggested_path:'上传.xlsx > 第一批 > 金额'},
        {expected_path:'上月.xlsx > 第二批 > 金额', suggested_path:'上传.xlsx > 第一批 > 金额'},
    ];
    assert.ok(ctx._constrainSourceSheetSuggestions(rows, {}).every(row => row.suggested_path === null),
        '两张训练表争用同一源表时不能默认选中任何冲突建议');
    const fixed = {'上传.xlsx':{expected_file:'上月.xlsx', sheet_mapping:{第一批:'第一批'}}};
    const guarded = ctx._constrainSourceSheetSuggestions(rows, fixed);
    assert.equal(guarded[0].suggested_path, rows[0].suggested_path);
    assert.equal(guarded[1].suggested_path, null, '缺失第二批不能再占用已确定的第一批');
    const updates = ctx._sourceSheetFieldUpdates(guarded, new Map(), JSON.stringify(['上月.xlsx','第二批']),
        JSON.stringify(['上传.xlsx','第二批']), ['上传.xlsx > 第二批 > 金额']);
    assert.equal(updates[0].value,'上传.xlsx > 第二批 > 金额');
    console.log('PASS: global sheet ownership removes conflicting defaults and manual selection fills the chosen sheet');
}
{
    const rows = [
        {expected_path:'trained.xlsx > 上月1 > 编号'},
        {expected_path:'trained.xlsx > 上月1 > 金额'},
        {expected_path:'trained.xlsx > 上月2 > 编号'},
    ];
    const paths = ['七月.xlsx > 第一表 > 编号', '七月.xlsx > 第一表 > 金额',
                   '七月.xlsx > 第二表 > 编号', '七月.xlsx > 第二表 > 金额'];
    const mixed = new Map([[0,paths[0]],[1,paths[3]],[2,paths[2]]]);
    const updates = ctx._sourceSheetFieldUpdates(rows, mixed, JSON.stringify(['trained.xlsx','上月1']),
        JSON.stringify(['七月.xlsx','第一表']), paths);
    assert.deepEqual(plain(updates.map(row => row.value)), paths.slice(0,2));
    assert.ok(updates.every(row => row.options.every(path => path.includes('第一表'))));
    assert.deepEqual(plain(updates.map(row => row.index)), [0,1], '只更新所属训练表');
    const overlay = {querySelectorAll: () => updates.map(row => ({dataset:{aiIdx:String(row.index)},value:row.value}))};
    const submitted = ctx._buildFileMappingFromAiSelections(overlay, rows, {});
    assert.deepEqual(plain(submitted['七月.xlsx'].sheet_mapping), {'第一表':'上月1'});
    const skipped = ctx._sourceSheetFieldUpdates(rows, mixed, JSON.stringify(['trained.xlsx','上月1']), '', paths);
    assert.ok(skipped.every(row => row.value === '' && row.options.length === 0));
    console.log('PASS: source Sheet selection immediately replaces all mixed fields and constrains their options');
}
const second = mapping('a.xlsx', '补贴', '补贴表', {'编号': '工号', '金额': '补贴'});
const merged = ctx._mergeConfirmations({confirmed_mapping: {file_mapping: first}}, {confirmed_mapping: {file_mapping: second}});
assert.deepEqual(plain(merged.confirmed_mapping.file_mapping['a.xlsx'].header_mapping_by_sheet), {
    工资: {编号:'工号', 金额:'工资'}, 补贴: {编号:'工号', 金额:'补贴'}
});
const noRows = ctx._mergeConfirmations(merged, {confirmed_mapping: {file_mapping: {}}});
assert.deepEqual(plain(noRows), plain(merged));
const remapped = ctx._mergeFileMappings(first, mapping('b.xlsx', '本月', '工资表', {'编号':'工号', '新金额':'工资'}));
assert.deepEqual(Object.keys(remapped), ['b.xlsx']);
assert.equal(first['a.xlsx'].header_mapping_by_sheet.工资.金额, '工资');
const sheetOnlyRemap = ctx._mergeFileMappings(first, mapping('b.xlsx', '本月', '工资表', {'新金额':'工资'}));
assert.equal(sheetOnlyRemap['a.xlsx'], undefined, '更换源 Sheet 时旧来源必须整体移除');
const rows = [{expected_path:'trained.xlsx > 工资表 > 工号'}, {expected_path:'trained.xlsx > 工资表 > 工资'}];
const overlay = {querySelectorAll: selector => selector.includes('data-ai-idx') ? [
    {dataset:{aiIdx:'0'}, value:'b.xlsx > 本月 > 编号'},
    {dataset:{aiIdx:'1'}, value:'b.xlsx > 本月 > 新金额'}
] : []};
const picked = ctx._buildFileMappingFromAiSelections(overlay, rows, first);
assert.equal(picked['b.xlsx'].header_mapping_by_sheet.本月.新金额, '工资');
assert.equal(picked['a.xlsx'], undefined);
const unchanged = {expected_path:'a.xlsx > Sheet > 编号', suggested_path:'a.xlsx > Sheet > 编号', confidence:1};
assert.equal(ctx._isUnchangedMapping(unchanged), true);
assert.equal(ctx._isUnchangedMapping({...unchanged, suggested_path:'b.xlsx > Sheet > 编号'}), false);
assert.equal(ctx._isUnchangedMapping({...unchanged, suggested_path:'a.xlsx > 新Sheet > 编号'}), false);
assert.equal(ctx._isUnchangedMapping({...unchanged, suggested_path:'a.xlsx > Sheet > 工号'}), false);
assert.equal(ctx._isUnchangedMapping({...unchanged, confidence:0.99}), false);
console.log('PASS: multi-round selections, sheet-scoped columns, remap replacement, conflict rejection, changed/stable grouping');


(async () => {
    const nodes = new Map();
    let appends = 0;
    const doc = {
        addEventListener() {},
        getElementById(id) { return nodes.get(id) || null; },
        createElement() {
            return {
                style: {}, parentNode: null,
                set innerHTML(html) {
                    this.html = html;
                    for (const id of ['_pre_confirm', '_pre_cancel', '_pre_validation_error', '_pre_skip_history']) {
                        nodes.set(id, {style: {}, disabled: false, checked: false, textContent: ''});
                    }
                },
                querySelectorAll() { return []; },
                remove() { nodes.delete(this.id); this.parentNode = null; }
            };
        },
        body: {
            appendChild(el) { appends++; nodes.set(el.id, el); el.parentNode = this; },
            removeChild(el) { el.remove(); }
        }
    };
    ctx.document = doc;
    const emptyDecision = await ctx._showPrecheckDialog({has_review_items:false});
    assert.equal(emptyDecision.mapping_finalized, true);
    assert.equal(appends, 0, '后端明确无审核项时不得创建空弹窗');
    for (const has_review_items of [undefined, true]) {
        const decision = await ctx._showPrecheckDialog({
            has_review_items, mapping_requires_confirmation:true,
            source_sheet_reviews:[{expected_file:'missing.xlsx', expected_sheet:'S'}],
            actual_sources:[], actual_paths:[], review_file_mapping:{},
        });
        assert.equal(decision.mapping_finalized, true);
        assert.equal(appends, 0, '训练侧缺口非空但没有可选上传来源时，不创建弹窗');
    }
    console.log('PASS: backend no-review decision advances without rendering an empty dialog');
    const pending = {target_candidates: [], history_warnings:['请确认历史数据缺失']};
    const firstDialog = ctx._showPrecheckDialog(pending);
    const overlay = doc.getElementById('_compute_precheck_overlay');
    doc.getElementById('_pre_confirm').onclick();
    const decision = await firstDialog;
    assert.equal(doc.getElementById('_compute_precheck_overlay'), overlay, '提交校验时保持原窗口');
    assert.equal(doc.getElementById('_pre_confirm').disabled, true);
    const previous = ctx._mergeConfirmations(null, decision);
    const secondDialog = ctx._showPrecheckDialog(pending, previous);
    assert.equal(appends, 1, '服务端仍有待处理项时复用窗口，不再次弹出');
    doc.getElementById('_pre_confirm').onclick();
    assert.match(doc.getElementById('_pre_validation_error').textContent, /本次选择与上次相同/);
    assert.equal(doc.getElementById('_pre_confirm').disabled, false, '相同参数不重复提交');
    doc.getElementById('_pre_skip_history').checked = true;
    doc.getElementById('_pre_confirm').onclick();
    assert.equal((await secondDialog).skip_history_check, true);
    ctx._closePrecheckDialog();
    assert.equal(doc.getElementById('_compute_precheck_overlay'), null);
    console.log('PASS: one persistent confirmation window, unchanged selection blocked locally, corrected selection submitted');
    const refreshDialog = ctx._showPrecheckDialog({session_id:'session-1',
        target_candidates:[{key:'当月2', all_sheets:['202608(1)','202608(2)'], candidates:[]}]});
    const refreshOverlay = doc.getElementById('_compute_precheck_overlay');
    const target = {value:'202608(2)', dataset:{targetKey:'当月2'},
                    matches: selector => selector.includes('data-target-key')};
    const fileChoice = {value:'new.xlsx', dataset:{renameUploaded:'old.xlsx'}, style:{},
                    matches: selector => selector.includes('data-rename-uploaded')};
    const oldDuplicates = [0,1].map(i => ({value:'old.xlsx > old > amount', dataset:{aiIdx:String(i)}, style:{}}));
    refreshOverlay.querySelectorAll = selector => selector === 'select[data-target-key]' ? [target]
        : selector === 'select[data-ai-idx]' ? oldDuplicates
        : selector === 'select[data-rename-uploaded]' ? [fileChoice] : [];
    let sent;
    ctx.AUTH = {authFetch: async (url, options) => {
        sent = JSON.parse(options.body);
        return {ok:true};
    }};
    ctx._readComputeSubmitJson = async () => ({session_id:'session-1', mapping_refreshed:true,
        target_candidates:[{key:'当月2', all_sheets:['202608(1)','202608(2)'], candidates:[]}]});
    refreshOverlay.onchange({target});
    assert.equal(doc.getElementById('_pre_confirm').textContent, '确认匹配并开始计算', '只改目标不重新推断源字段');
    assert.equal(sent, undefined);
    refreshOverlay.onchange({target:fileChoice});
    assert.equal(doc.getElementById('_pre_confirm').textContent, '应用表关系并刷新字段');
    await doc.getElementById('_pre_confirm').onclick();
    assert.equal(sent.refresh_only, true);
    assert.equal(sent.confirmed_target_map.当月2, '202608(2)');
    assert.equal(sent.confirmed_mapping, undefined, '旧自动重复建议不能在刷新时被固化为人工决定');
    assert.equal(doc.getElementById('_compute_precheck_overlay'), refreshOverlay);
    refreshOverlay.querySelectorAll = () => [];
    doc.getElementById('_pre_confirm').onclick();
    assert.ok(await refreshDialog, '刷新之后允许进入下一步，不被相同确认参数拦住');
    console.log('PASS: target edits do not rematch source columns; source-file edits can refresh explicitly');
    ctx._closePrecheckDialog();
    const choices = {};
    ctx._showPrecheckDialog({target_candidates:[{key:'当月2', candidates:[{sheet:'202608(1)',score:1}],
        all_sheets:['202608(1)','202608(2)']}]}, null, choices);
    ctx._showPrecheckDialog({mapping_refreshed:true}, {confirmed_target_map:{当月2:'202608(2)'}}, choices);
    const html = doc.getElementById('_compute_precheck_overlay').html;
    assert.ok(html.includes('data-target-key="当月2"'), '已解决的 Sheet 仍可修改');
    assert.ok(html.includes('value="202608(2)" selected'), '回显手工选择而非第一推荐');
    console.log('PASS: resolved sheet choices remain editable with confirmed value selected');
    ctx._showPrecheckDialog({
        actual_sources:[{file:'source.xlsx', sheet:'原始工资表', original_file:'原始上传.xls', original_sheet:'原始工资表'}],
        actual_paths:['source.xlsx > 原始工资表 > 工号'],
        source_sheet_reviews:[{expected_file:'训练名称.xlsx', expected_sheet:'训练Sheet',
            suggested_file:'source.xlsx', suggested_sheet:'原始工资表', recommendation_source:'ai'}],
        file_mapping:{'source.xlsx':{expected_file:'训练名称.xlsx', sheet_mapping:{'原始工资表':'训练Sheet'}, header_mapping:{'工号':'工号'}}}
    });
    const originalHtml = doc.getElementById('_compute_precheck_overlay').html;
    assert.ok(originalHtml.includes('原始上传.xls') && originalHtml.includes('原始工资表'), '匹配项必须展示原始文件和真实 Sheet');
    assert.ok(originalHtml.includes('data-upload-source-key="[&quot;source.xlsx&quot;,&quot;原始工资表&quot;]"'), '提交值保留稳定身份，不受显示标签影响');
    console.log('PASS: original upload labels and full source details preserve stable mapping identities');
    // 实际提交循环：旧服务返回空审核（仍带 true 标志），无需任何点击即继续到任务流。
    ctx._closePrecheckDialog();
    const appendCount = appends;
    nodes.set('compute-btn', {disabled:false, textContent:''});
    nodes.set('source-files', {files:['upload.xlsx']});
    nodes.set('salary-month', {value:''});
    nodes.set('standard-hours', {value:''});
    ctx.FormData = class { append() {} set() {} };
    ctx.console = console;
    ctx._autoCheckEncryption = async () => true;
    for (const name of ['clearResult','addLog','updateStatus','updateProgress','_saveActiveTask']) {
        ctx[name] = () => {};
    }
    ctx.showError = error => { throw new Error(error); };
    let streamedTask;
    ctx._connectComputeStream = id => { streamedTask = id; };
    const requests = [];
    ctx.AUTH = {authFetch: async (url, options) => {
        requests.push({url, options});
        return {ok:true, data:requests.length === 1 ? {
            error_type:'precheck_failed', session_id:'empty-review',
            has_review_items:true, mapping_requires_confirmation:true,
            source_sheet_reviews:[{expected_file:'base.xlsx', expected_sheet:'S'}],
            actual_sources:[], actual_paths:[], review_file_mapping:{},
        } : {task_id:'computed-without-click'}};
    }};
    ctx._readComputeSubmitJson = async response => response.data;
    vm.runInContext("currentTenantId = 'test'; currentScriptId = 'test-script';", ctx);
    await ctx.startCompute();
    assert.equal(appends, appendCount, '整个提交循环不能生成空审核框');
    assert.equal(requests.length, 2);
    assert.equal(requests[1].url, '/api/compute/session/empty-review/confirm');
    assert.equal(JSON.parse(requests[1].options.body).mapping_finalized, true);
    assert.equal(streamedTask, 'computed-without-click');
    console.log('PASS: empty review proceeds through submit, session confirmation and task stream without a click');
})().catch(error => { console.error(error); process.exitCode = 1; });


{
    const rows = [{expected_path:'trained.xlsx > 工资表 > 工号'}, {expected_path:'trained.xlsx > 工资表 > 工资'}];
    const overlay = {querySelectorAll: () => [
        {dataset:{aiIdx:'0'}, value:'a.xlsx > 工资 > 编号'},
        {dataset:{aiIdx:'1'}, value:''}
    ]};
    const picked = ctx._buildFileMappingFromAiSelections(overlay, rows, first);
    assert.deepEqual(plain(picked['a.xlsx'].header_mapping_by_sheet.工资), {编号:'工号'});
    const unmatched = ctx._collectUnmatchedColumns(overlay, rows);
    assert.deepEqual(plain(unmatched), [['trained.xlsx','工资表','工资']]);
    const decision = ctx._mergeConfirmations({confirmed_mapping:{file_mapping:first}},
        {confirmed_mapping:{file_mapping:picked, unmatched_columns:unmatched}});
    assert.deepEqual(plain(decision.confirmed_mapping.file_mapping['a.xlsx'].header_mapping_by_sheet.工资), {编号:'工号'});
    assert.equal(decision.confirmed_mapping.unmatched_columns.length,1);
    const restored = ctx._mergeConfirmations(decision, {confirmed_mapping:{file_mapping:first}});
    assert.equal(restored.confirmed_mapping.unmatched_columns.length,0);
    const summary = ctx._precheckSummary({missing_columns:[{error:'缺少列'.repeat(2000), expected_columns:['A','B']}]});
    assert.ok(summary.length < 180);
    assert.doesNotMatch(summary,/待确认.*列|无匹配/);
    console.log('PASS: empty selection continues, old mapping removed, skip persists and can be undone, missing summary bounded');
}

{
    const rows = [{expected_path:'七月薪酬.xlsx > 上月1 > 工号'},
                  {expected_path:'七月薪酬.xlsx > 上月1 > 金额'},
                  {expected_path:'其他.xlsx > 数据 > 备注'}];
    const overlay = {_pendingFileTargets: new Set(['七月薪酬.xlsx']), querySelectorAll: () => [
        {dataset:{aiIdx:'0'}, value:''},
        {dataset:{aiIdx:'1', mappingEdited:'1'}, value:''},
        {dataset:{aiIdx:'2'}, value:''}
    ]};
    assert.deepEqual(plain(ctx._collectUnmatchedColumns(overlay, rows)),
        [['七月薪酬.xlsx','上月1','金额'], ['其他.xlsx','数据','备注']]);
    console.log('PASS: newly confirmed file is resolved before default blank columns are skipped; explicit skips remain');
}
