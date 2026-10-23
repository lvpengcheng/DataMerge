const assert = require('node:assert/strict');
const path = require('node:path');
const {chromium} = require('playwright');
const root = path.resolve(__dirname, '..');
const files = ['新模板', '新工资', '门店列表'].map((sheet, i) => ({
    name: `本月__${sheet}.xlsx`, physical_file: '本月.xlsx', sheet, sheet_explicit: true,
    fingerprint: `fp${i}`, columns: ['工号', '金额', '工作地'], suggested_key: '工号', header_start: 1, header_end: 1,
}));
(async () => {
    const browser = await chromium.launch({channel: 'msedge', headless: true});
    try {
        for (const tab of ['data-integrate', 'table-split']) {
            const page = await browser.newPage({viewport: {width: 1440, height: 1000}});
            const errors = [], requests = [];
            let selected = [];
            page.on('pageerror', e => errors.push(e.message));
            page.on('dialog', d => {errors.push(d.message()); d.dismiss();});
            await page.addInitScript(() => {
                localStorage.setItem('datamerge_token', 'test');
                localStorage.setItem('datamerge_user', JSON.stringify({username: 'test', role_name: 'admin'}));
            });
            await page.route('**/*', async route => {
                const url = new URL(route.request().url());
                if (url.pathname.startsWith('/static/')) return route.fulfill({path: path.join(root, 'frontend', url.pathname)});
                if (url.pathname === '/tools') return route.fulfill({path: path.join(root, 'frontend/templates/tools.html'), contentType: 'text/html'});
                const body = route.request().method() === 'POST' && route.request().headers()['content-type']?.includes('application/json')
                    ? route.request().postDataJSON() : null;
                requests.push({url: url.pathname, body});
                if (url.pathname.endsWith('/analyze')) return route.fulfill({json: {
                    session_id: 'a'.repeat(32), files, parse_all_sheets: true, sheet_selection_pending: true, matched_schemes: []}});
                if (url.pathname.endsWith('/select-sheets')) {
                    selected = files.filter(f => body.selected_tables.includes(f.name));
                    return route.fulfill({json: {session_id: 'a'.repeat(32), files: selected, parse_all_sheets: true, matched_schemes: []}});
                }
                if (url.pathname.endsWith('/suggest-sheet-mapping')) return route.fulfill({json: {
                    role_files: [selected[1].name, selected[0].name], ai_suggested: true}});
                if (url.pathname.endsWith('/apply-validate')) {
                    const roleFiles = body.role_files || selected.map(f => f.name);
                    return route.fulfill({json: {ok: true, files: selected, sheet_mapping: body.role_files ? null : {
                        required: true, roles: selected.map((f, i) => ({role_index: i, saved_file: '上月.xlsx',
                            saved_sheet: i ? '工资' : '模板', selected: f.name, candidates: selected.map(t => t.name)})), tables: selected},
                        resolved_scheme: {id: 1, main_file: roleFiles[0], role_files: roleFiles,
                            config: {key_map_by_role: {'0': '工号', '1': '工号'},
                                overwrite_pairs: [{a_col: '金额', source_role: 1, source_col: '金额'}],
                                split_config: {mode: 'value', column: '工作地'}}}}});
                }
                return route.fulfill({json: []});
            });
            await page.goto('http://datamerge.test/tools');
            await page.locator(`[data-tab="${tab}"]`).click();
            await page.locator('#btn-int-new-scheme').click();
            await page.locator('#int-parse-all-sheets').check();
            await page.locator('#integrate-file-input').setInputFiles({name: '本月.xlsx',
                mimeType: 'application/octet-stream', buffer: Buffer.from('test')});
            await page.locator('#btn-integrate-analyze').click();
            await page.locator('.int-choose-sheet').first().waitFor();
            assert.equal(await page.locator('#integrate-config').isVisible(), false);
            assert.equal(await page.locator('.int-sheet-group').count(), 1);
            await page.locator('.int-choose-sheet').nth(2).uncheck();
            await page.locator('#modal-confirm').click();
            await page.locator('#int-main-file').waitFor();
            assert.equal(await page.locator('#int-main-file option').count(), 2);
            assert.equal(await page.locator('.int-key').count(), 2);
            assert.equal(requests.find(r => r.url.endsWith('/select-sheets')).body.selected_tables.includes(files[2].name), false);
            await page.evaluate(() => {
                _intMode = 'apply'; _intApplyTarget = {id: 1, name: '方案'};
                void Tools._intPrefillForApply();
            });
            await page.locator('.int-sheet-relation').first().waitFor();
            assert.equal(await page.locator('#int-execute').isDisabled(), true);
            await page.locator('#int-sheet-ai-match').click();
            await page.waitForFunction(() => document.getElementById('int-sheet-ai-status').textContent.includes('完成'));
            assert.equal(await page.locator('.int-sheet-relation').first().inputValue(), selected[1].name);
            await page.screenshot({path: path.join(root, `.codex-build/sheet-relations-${tab}.png`), fullPage: true});
            await page.locator('#modal-confirm').click();
            await page.waitForFunction(() => _intApplyReady === true);
            assert.equal(await page.locator('#int-main-file').inputValue(), selected[1].name);
            const confirmed = requests.find(r => r.url.endsWith('/apply-validate') && r.body?.role_files);
            assert.deepEqual(confirmed.body.role_files, [selected[1].name, selected[0].name]);
            assert.equal(requests.some(r => r.url.endsWith('/execute')), false);
            // Cancellation must leave the application blocked, including after
            // a validation error (the global modal supports retry).
            await page.evaluate(() => {void Tools._intPrefillForApply();});
            await page.locator('.int-sheet-relation').first().waitFor();
            await page.locator('.int-sheet-relation').nth(1).selectOption(selected[0].name);
            await page.locator('.int-sheet-relation').first().selectOption(selected[0].name);
            await page.locator('#modal-confirm').click();
            assert.match(await page.locator('#int-sheet-relation-error').innerText(), /不能对应多个/);
            await page.locator('.modal-footer button').first().click();
            await page.waitForFunction(() => !Tools._integrationPending);
            assert.equal(await page.locator('#int-execute').isDisabled(), true);
            assert.deepEqual(errors, []);
            await page.close();
        }
        console.log('Both tools: per-workbook sheet selection, AI pre-match, confirmed role validation and cancellation passed');
    } finally {await browser.close();}
})().catch(e => {console.error(e); process.exitCode = 1;});
