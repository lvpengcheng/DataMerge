const assert = require('node:assert/strict');
const path = require('node:path');
const {chromium} = require('playwright');
const root = path.resolve(__dirname, '..');
const files = ['A.xlsx', 'B.xlsx', 'C.xlsx'].map((name, i) => ({name, physical_file: name,
    sheet: i ? '在职' : '模板', fingerprint: `fp${i}`, columns: ['工号', '工作地', '金额', '状态'],
    suggested_key: '工号', header_start: 1, header_end: 1, data_row_start: 2, data_row_end: 21}));
(async () => {
    const browser = await chromium.launch({channel: 'msedge', headless: true});
    try {
        for (const tab of ['data-integrate', 'table-split']) {
            const page = await browser.newPage({viewport: {width: 1440, height: 1000}, acceptDownloads: true});
            const errors = [];
            let saved, executed;
            page.on('pageerror', error => errors.push(error.message));
            page.on('dialog', dialog => {errors.push(dialog.message()); dialog.dismiss();});
            await page.addInitScript(() => {
                localStorage.setItem('datamerge_token', 'test');
                localStorage.setItem('datamerge_user', JSON.stringify({username: 'test', role_name: 'admin'}));
            });
            await page.route('**/*', async route => {
                const pathname = new URL(route.request().url()).pathname;
                if (pathname.startsWith('/static/')) return route.fulfill({path: path.join(root, 'frontend', pathname)});
                if (pathname === '/tools') return route.fulfill({path: path.join(root, 'frontend/templates/tools.html'), contentType: 'text/html'});
                if (pathname.endsWith('/analyze')) return route.fulfill({json: {session_id: 'a'.repeat(32), files}});
                if (pathname.endsWith('/scheme/save')) {
                    saved = route.request().postDataJSON();
                    return route.fulfill({json: {ok: true, id: 1, name: saved.name}});
                }
                if (pathname.endsWith('/execute')) {
                    executed = route.request().postDataJSON();
                    return route.fulfill({body: Buffer.from('test Excel'), contentType: 'application/octet-stream'});
                }
                return route.fulfill({json: []});
            });
            await page.goto('http://datamerge.test/tools');
            await page.locator(`[data-tab="${tab}"]`).click();
            await page.locator('#btn-int-new-scheme').click();
            await page.locator('#integrate-file-input').setInputFiles(files.map(f => ({name: f.name,
                mimeType: 'application/octet-stream', buffer: Buffer.from('test')})));
            await page.locator('#btn-integrate-analyze').click();
            await page.locator('#int-data-scope-mode').waitFor({state: 'attached'});
            assert.equal(await page.locator('#int-data-scope').isVisible(), tab === 'table-split');
            await page.locator('#int-union-source-keys').check();
            assert.equal(await page.locator('#int-data-scope').isVisible(), true);
            await page.locator('#int-data-scope-mode').selectOption('source');
            await page.locator('#int-data-scope-file').selectOption('B.xlsx');
            await page.locator('#int-data-scope-filter').fill('[状态]="在职"');
            await page.evaluate(() => {
                Tools._intAddListRow('ow', '金额', [{file: 'B.xlsx', expr: '金额'}]);
            });
            const download = page.waitForEvent('download');
            await page.locator('#int-execute').click();
            await download;
            assert.equal(executed.main_file, 'A.xlsx');
            assert.equal(executed.data_scope.source_file, 'B.xlsx');
            assert.equal(executed.data_scope.filter_expression, '[状态]="在职"');
            assert.equal(executed.union_source_keys, true);

            await page.locator('#int-data-scope-mode').selectOption('expression');
            await page.locator('[data-scope-example="INTERSECT"]').click();
            assert.equal(await page.locator('#int-data-scope-expression').inputValue(), 'INTERSECT("B.xlsx","C.xlsx")');
            await page.locator('#int-scheme-name').fill('范围方案');
            await page.locator('#int-save-scheme').click();
            await page.waitForFunction(() => document.getElementById('int-save-status').textContent.includes('已保存'));
            assert.equal(saved.data_scope.expression, 'INTERSECT("B.xlsx","C.xlsx")');
            assert.equal(saved.data_scope.mode, 'expression');
            await page.evaluate(() => Tools._intFillFromScheme({main_file: 'A.xlsx', role_files: ['A.xlsx', 'B.xlsx', 'C.xlsx'],
                config: {union_source_keys: true, data_scope: {mode: 'source', source_file: 'C.xlsx', filter_expression: '[状态]="离职"'},
                    key_map_by_role: {'0': '工号', '1': '工号', '2': '工号'},
                    split_config: {mode: 'value', column: '工作地', include_source_keys: true}}}));
            assert.equal(await page.locator('#int-data-scope-mode').inputValue(), 'source');
            assert.equal(await page.locator('#int-data-scope-file').inputValue(), 'C.xlsx');
            assert.equal(await page.locator('#int-data-scope-filter').inputValue(), '[状态]="离职"');
            await page.locator('#int-union-source-keys').uncheck();
            if (tab === 'table-split') await page.locator('#split-include-source-keys').uncheck();
            assert.equal(await page.locator('#int-data-scope').isVisible(), false);
            assert.deepEqual(await page.evaluate(() => Tools._readDataScope()), {mode: 'all'});
            const remapped = await page.evaluate(() => {
                const scope = {mode: 'expression', expression: 'EXCEPT("B.xlsx","C.xlsx")'};
                Tools._remapDataScopeFiles(scope, {'B.xlsx': 'C.xlsx', 'C.xlsx': 'B.xlsx'});
                return scope;
            });
            assert.equal(remapped.expression, 'EXCEPT("C.xlsx","B.xlsx")');
            await page.locator('#int-union-source-keys').check();
            await page.screenshot({path: path.join(root, `.codex-build/data-scope-${tab}.png`), fullPage: true});
            assert.deepEqual(errors, []);
            await page.close();
        }
        console.log('Both tools: scope visibility, source/filter, set expressions, execute/save payloads and scheme prefill passed');
    } finally { await browser.close(); }
})().catch(error => {console.error(error); process.exitCode = 1;});
