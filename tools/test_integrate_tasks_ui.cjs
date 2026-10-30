const assert = require('node:assert/strict');
const path = require('node:path');
const {chromium} = require('playwright');
const root = path.resolve(__dirname, '..');
(async () => {
    const browser = await chromium.launch({channel: 'msedge', headless: true});
    try {
        const page = await browser.newPage();
        const errors = [], requests = [];
        page.on('pageerror', e => errors.push(e.message));
        page.on('dialog', d => {errors.push(d.message()); d.dismiss();});
        await page.addInitScript(() => {
            localStorage.setItem('datamerge_token', 'test');
            localStorage.setItem('datamerge_user', JSON.stringify({username:'test',role_name:'admin'}));
        });
        await page.route('**/*', route => {
            const url = new URL(route.request().url());
            if (url.pathname.startsWith('/static/')) return route.fulfill({path:path.join(root,'frontend',url.pathname)});
            if (url.pathname === '/tools') return route.fulfill({path:path.join(root,'frontend/templates/tools.html'),contentType:'text/html'});
            requests.push({url: url.pathname, body: route.request().postData()});
            if (url.pathname.endsWith('/apply-validate')) return route.fulfill({json:{ok:true,
                resolved_scheme:{main_file:'test__工资.xlsx',role_files:['test__工资.xlsx','test__工资源.xlsx'],
                    config:{key_map_by_role:{'0':'工号','1':'工号'},
                        overwrite_pairs:[{a_col:'金额',source_role:1,source_expr:'金额*2'}],compare_pairs:[]}}}});
            return route.fulfill({json: []});
        });
        await page.goto('http://datamerge.test/tools');
        await page.locator('[data-tab="data-integrate"]').click();
        await page.evaluate(() => {
            Tools._integrationKind = 'integrate'; Tools._intShowWork('create');
            Tools.loadIntegrateSchemes = async () => {};
            _integrateAnalysis = {session_id:'a'.repeat(32), files:['工资','奖金','工资源','奖金源'].map((sheet,i) => ({
                name: 'test__'+sheet+'.xlsx', physical_file:i<2?'模板.xlsx':'来源.xlsx', sheet,
                columns:['工号','金额'], suggested_key:'工号',header_start:1,header_end:1,
            }))};
            Tools._renderIntegrateConfig(_integrateAnalysis);
            Tools._intAddListRow('ow','金额',[{file:'test__工资源.xlsx',expr:'金额*2'}]);
        });
        await page.locator('#int-task-output-file').fill('结果.xlsx');
        await page.locator('#int-add-targets').click();
        await page.locator('.int-new-target[value="test__奖金.xlsx"]').check();
        await page.locator('#modal-confirm').click();
        assert.equal(await page.locator('#int-main-file').inputValue(),'test__奖金.xlsx');
        assert.equal(await page.locator('#int-ow-list-rows tr').count(),0);
        await page.evaluate(() => Tools._intAddListRow('ow','金额',[{file:'test__奖金源.xlsx',expr:'金额*3'}]));
        await page.locator('#int-task-output-file').fill('结果.xlsx');
        await page.locator('[data-output-task="0"]').click();
        assert.equal(await page.locator('#int-task-output-file').inputValue(),'结果.xlsx');
        assert.match(await page.locator('#int-ow-list-rows tr').getAttribute('data-src'),/金额\*2/);
        await page.locator('.int-task-source[value="test__奖金源.xlsx"]').uncheck();
        assert.equal(await page.locator('.int-key[data-file="test__奖金源.xlsx"]').count(),0);
        await page.locator('[data-output-task="1"]').click();
        assert.match(await page.locator('#int-ow-list-rows tr').getAttribute('data-src'),/金额\*3/);
        assert.equal(await page.locator('.int-key[data-file="test__奖金源.xlsx"]').count(),1);
        await page.locator('#int-scheme-name').fill('多目标方案');
        await page.locator('#int-save-scheme').click();
        await page.waitForFunction(() => document.getElementById('int-save-status').textContent.includes('已保存'));
        const saved = JSON.parse(requests.find(r => r.url.endsWith('/scheme/save')).body);
        assert.equal(saved.output_tasks.length,2);
        assert.equal(saved.main_file,'test__工资.xlsx');
        assert.match(saved.output_tasks[0].overwrite_pairs[0].source_expr,/金额\*2$/);
        assert.match(saved.output_tasks[1].overwrite_pairs[0].source_expr,/金额\*3$/);
        await page.evaluate(tasks => {
            Tools._intFillFromScheme({main_file:tasks[0].main_file, role_files:tasks[0].active_files,
                config:{output_tasks:tasks,key_map_by_role:{},overwrite_pairs:[]}});
        }, saved.output_tasks);
        assert.equal(await page.locator('[role="tab"][data-output-task]').count(),2);
        assert.match(await page.locator('#int-ow-list-rows tr').getAttribute('data-src'),/金额\*2/);
        await page.locator('#int-copy-target').click();
        await page.locator('.int-new-target[value="test__奖金.xlsx"]').check();
        await page.locator('#modal-confirm').click();
        assert.equal(await page.locator('[role="tab"][data-output-task]').count(),3);
        assert.match(await page.locator('#int-ow-list-rows tr').getAttribute('data-src'),/金额\*2/);
        await page.locator('#int-remove-target').click();
        assert.equal(await page.locator('[role="tab"][data-output-task]').count(),2);
        // A saved single-target scheme may use only some of the parsed sheets.
        await page.evaluate(async () => {
            const analysis = _integrateAnalysis;
            Tools._intShowWork('apply',{id:77,name:'山东太保-拆表V2'});
            _integrateAnalysis = analysis;
            Tools._renderIntegrateConfig(_integrateAnalysis);
            await Tools._intPrefillForApply();
            Tools._captureOutputTask();
        });
        const legacy = await page.evaluate(() => Tools._outputTasks[0]);
        assert.deepEqual(Object.keys(legacy.key_map).sort(),[...legacy.active_files].sort());
        assert.equal(await page.locator('.int-key').count(),2);
        assert.equal(await page.locator('#int-execute').isEnabled(),true);
        await page.evaluate(async () => Tools._intFetch('/api/tools/integrate/execute',{
            method:'POST',headers:{'Content-Type':'application/json'},body:JSON.stringify({session_id:_integrateAnalysis.session_id})}));
        const executed=JSON.parse(requests.findLast(r => r.url.endsWith('/execute')).body);
        assert.deepEqual(Object.keys(executed.output_tasks[0].key_map).sort(),[...executed.output_tasks[0].active_files].sort());
        assert.deepEqual(errors,[]);
        console.log('Target tabs: independent formulas/sources, switching, copy/remove, save and restore passed');
    } finally {await browser.close();}
})().catch(e => {console.error(e);process.exitCode=1;});
