const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {chromium} = require('playwright');
const root = path.resolve(__dirname, '..');
const cols = ['金额', '姓名', '计算字段', '固定字段', '缺失字段', '人工字段'];
const files = [
    {name: '主表.xlsx', columns: cols, sheet: '模板', available_sheets: ['模板'], header_start: 1, header_end: 1},
    {name: 'A.xlsx', columns: ['金额A', '姓名', 'ID'], sheet: '门店', available_sheets: ['门店', '人员'], header_start: 1, header_end: 1},
    {name: 'B.xlsx', columns: ['金额B', '姓名', 'ID'], sheet: '人员', header_start: 1, header_end: 1},
];
const source = (file, col) => ({source_file: file, source_col: col, sheet: '人员', label: col});
const results = [
    {a_col: '金额', label: '金额', status: 'matched', sources: [source('A.xlsx', '金额A')], reason: '确定映射'},
    {a_col: '姓名', label: '姓名', status: 'needs_review', sources: [source('A.xlsx', '姓名'), source('B.xlsx', '姓名')], reason: '多个候选来源'},
    {a_col: '计算字段', label: '计算字段', status: 'formula', sources: [source('A.xlsx', '金额A'), source('B.xlsx', '金额B')], reason: '需计算', suggestion: '请确认应加还是应减'},
    {a_col: '固定字段', label: '固定字段', status: 'constant', sources: [], reason: '固定业务值', suggestion: '请确认固定值'},
    {a_col: '缺失字段', label: '缺失字段', status: 'unmatched', sources: [], reason: '未找到对应'},
];
(async () => {
    const browser = await chromium.launch({channel: 'msedge', headless: true});
    try {
        for (const tab of ['data-integrate', 'table-split']) {
            const page = await browser.newPage({viewport: {width: 1440, height: 1000}});
            const errors = [], requests = [];
            page.on('pageerror', error => errors.push(error.message));
            page.on('dialog', dialog => {errors.push(dialog.message()); dialog.dismiss();});
            await page.addInitScript(() => {
                localStorage.setItem('datamerge_token', 'test');
                localStorage.setItem('datamerge_user', JSON.stringify({username: 'test', role_name: 'admin'}));
            });
            await page.route('**/*', async route => {
                const url = new URL(route.request().url());
                if (url.pathname.startsWith('/static/')) return route.fulfill({path: path.join(root, 'frontend', url.pathname)});
                if (url.pathname === '/tools') return route.fulfill({path: path.join(root, 'frontend/templates/tools.html'), contentType: 'text/html'});
                if (url.pathname.endsWith('/analyze')) return route.fulfill({json: {session_id: 'a'.repeat(32), files}});
                if (url.pathname.endsWith('/reparse-headers')) {
                    const body = route.request().postDataJSON();
                    requests.push(body);
                    return route.fulfill({json: {session_id: 'a'.repeat(32), files: files.map(f => ({...f,
                        sheet: body.sheet_selections[f.name] || f.sheet}))}});
                }
                if (url.pathname.endsWith('/match')) {
                    const body = route.request().postDataJSON();
                    requests.push(body);
                    const selected = results.filter(r => body.target_cols.includes(r.a_col));
                    return route.fulfill({json: {results: selected,
                        warnings: ['测试提示：请检查人员Sheet'],
                        pairs: selected.filter(r => r.status === 'matched').flatMap(r =>
                            r.sources.map(source => ({a_col: r.a_col, ...source})))}});
                }
                return route.fulfill({json: []});
            });
            await page.goto('http://datamerge.test/tools');
            await page.locator(`[data-tab="${tab}"]`).click();
            await page.locator('#btn-int-new-scheme').click();
            await page.locator('#integrate-file-input').setInputFiles([
                {name: '主表.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('test')},
                {name: 'A.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('test')},
                {name: 'B.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('test')},
            ]);
            await page.locator('#btn-integrate-analyze').click();
            await page.locator('#int-main-file').waitFor();
            assert.equal(await page.locator('.int-sheet-select, .int-table-enabled').count(), 0);
            await page.locator('#int-ow-all').click();
            await page.evaluate(() => {
                const row = [...document.querySelectorAll('#int-ow-list-rows tr')].find(r => r.dataset.aCol === '人工字段');
                row.dataset.src = JSON.stringify([{file: 'A.xlsx', expr: '金额A'}]);
                Tools._intUpdateSrcCell(row);
            });
            await page.locator('#int-ow-ai').check();
            await page.locator('#int-ow-match').click();
            await page.waitForFunction(() => document.getElementById('int-ow-status').textContent.includes('智能匹配填充'));
            const request = requests[requests.length - 1];
            assert.equal(request.target_cols.includes('人工字段'), false);
            assert.equal(request.use_ai, true);
            const state = await page.evaluate(() => Object.fromEntries([...document.querySelectorAll('#int-ow-list-rows tr')]
                .map(r => [r.dataset.aCol, JSON.parse(r.dataset.src)])));
            assert.equal(state['金额'][0].expr, '金额A');
            assert.deepEqual(state['姓名'], []);
            assert.deepEqual(state['计算字段'], []);
            assert.deepEqual(state['固定字段'], []);
            assert.equal(state['人工字段'][0].expr, '金额A');
            assert.equal(await page.locator('.field-match-report, [data-match-row], [data-match-edit]').count(), 0);
            assert.equal(await page.locator('body').innerText().then(t =>
                /理由与操作|来源 \/ 建议|测试提示：请检查人员Sheet/.test(t)), false);
            assert.match(await page.locator('#int-ow-status').innerText(), /4 个不匹配/);
            await page.locator('#int-ow-match').click();
            await page.waitForFunction(() => !Tools._integrationPending);
            assert.equal(requests[requests.length - 1].target_cols.includes('金额'), false);
            assert.equal(requests[requests.length - 1].target_cols.includes('人工字段'), false);
            const fixedRow = page.locator('#int-ow-list-rows tr').filter({has: page.locator('td', {hasText: /^人工字段$/})});
            await fixedRow.locator('.int-src-cell').click();
            await page.locator('#int-unified-fx').fill('"aaaa"');
            await page.evaluate(() => Tools.confirmModal());
            assert.equal(await fixedRow.getAttribute('data-src').then(s => JSON.parse(s)[0].expr), '"aaaa"');
            await page.locator('#int-union-source-keys').check();
            assert.equal(await fixedRow.getAttribute('data-src').then(s => JSON.parse(s)[0].expr), '"aaaa"');
            await page.locator('#int-union-source-keys').uncheck();
            assert.equal(await fixedRow.getAttribute('data-src').then(s => JSON.parse(s)[0].expr), '"aaaa"');
            await page.screenshot({path: path.join(root, `.codex-build/structured-matching-${tab}.png`), fullPage: true});
            assert.equal(await page.locator('.int-key').count(), 3);
            assert.deepEqual(errors, []);
            await page.close();
        }
        console.log('Both tools: original layout and matching controls, direct fill, unmatched fields and existing mappings preserved');
    } finally { await browser.close(); }
})().catch(error => {console.error(error); process.exitCode = 1;});
