const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const {chromium} = require('playwright');
const root = path.resolve(__dirname, '..');
const permissions = {'menu.tools': true, 'tools.table_split': true,
    ...Object.fromEntries(['create', 'apply', 'edit', 'delete'].map(a => [`tools.table_split.${a}`, true]))};
const files = [{name: '工资.xlsx', physical_file: '工资.xlsx', sheet: 'Sheet1', columns: ['工号', '工作地', '金额'],
    suggested_key: '工号', fingerprint: 'main', header_start: 1, header_end: 2, data_row_start: 3, data_row_end: 8}];
let saved, executed;
const errors = [];
(async () => {
    const browser = await chromium.launch({channel: 'msedge', headless: true});
    try {
        const page = await browser.newPage({viewport: {width: 1440, height: 1000}, acceptDownloads: true});
        page.on('pageerror', error => errors.push(error.message));
        page.on('dialog', dialog => {errors.push(dialog.message()); dialog.dismiss();});
        await page.addInitScript(perms => {
            localStorage.setItem('datamerge_token', 'test');
            localStorage.setItem('datamerge_user', JSON.stringify({username: 'test', permissions: perms}));
        }, permissions);
        await page.route('**/*', async route => {
            const url = new URL(route.request().url());
            const name = url.pathname;
            if (name.startsWith('/static/')) {
                return route.fulfill({path: path.join(root, 'frontend', name)});
            }
            if (name === '/tools') return route.fulfill({path: path.join(root, 'frontend/templates/tools.html'), contentType: 'text/html'});
            if (name === '/api/tools/table-split/schemes') return route.fulfill({json: saved ? [{id: 1, name: saved.name, config: saved, can_edit: true, can_modify: true, file_count: 1}] : []});
            if (name === '/api/tools/table-split/analyze') return route.fulfill({json: {session_id: 'a'.repeat(32), files, matched_schemes: []}});
            if (name === '/api/tools/table-split/scheme/save') {
                saved = route.request().postDataJSON();
                return route.fulfill({json: {ok: true, id: 1, name: saved.name}});
            }
            if (name === '/api/tools/table-split/execute') {
                executed = route.request().postDataJSON();
                return route.fulfill({body: Buffer.from('mock ZIP'), contentType: 'application/zip', headers: {'X-Split-Groups': '2', 'X-Split-Rows': '6'}});
            }
            return route.fulfill({json: []});
        });
        await page.goto('http://datamerge.test/tools');
        await page.locator('[data-tab="table-split"]').click();
        assert.equal(await page.locator('#tab-table-split h2').textContent(), '拆表');
        assert.equal(await page.locator('[data-tab="data-integrate"]').isVisible(), false);
        await page.locator('#btn-int-new-scheme').click();
        await page.locator('#integrate-file-input').setInputFiles({name: '工资.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('test')});
        assert.equal(await page.locator('#btn-integrate-analyze').isEnabled(), true);
        await page.locator('#btn-integrate-analyze').click();
        await page.locator('#split-group-mode').waitFor();
        await page.locator('#split-group-mode').selectOption('expression');
        await page.locator('#split-example-city').click();
        assert.equal(await page.locator('#split-group-expression').inputValue(), 'IF([工作地]="上海","上海","其他")');
        await page.locator('#split-data-start').fill('3');
        await page.locator('#split-data-end').fill('8');
        const download = page.waitForEvent('download');
        await page.locator('#int-execute').click();
        assert.equal((await download).suggestedFilename(), '拆表结果.zip');
        assert.equal(executed.split_config.expression, 'IF([工作地]="上海","上海","其他")');
        assert.equal(executed.overwrite_pairs.length, 0);
        assert.equal(executed.split_config.data_row_start, 3);
        assert.equal(executed.split_config.include_source_keys, true);
        await page.locator('#int-scheme-name').fill('上海和其他');
        await page.screenshot({path: path.join(root, '.codex-build/table-split-desktop.png'), fullPage: true});
        await page.locator('#int-save-scheme').click();
        await page.waitForFunction(() => document.getElementById('int-list-status').textContent.includes('1 个方案'));
        assert.equal(saved.split_config.data_row_end, 8);
        assert.equal(saved.name, '上海和其他');
        assert.equal(await page.locator('.int-row-edit').count(), 1);
        await page.setViewportSize({width: 390, height: 844});
        await page.screenshot({path: path.join(root, '.codex-build/table-split-mobile.png'), fullPage: true});
        assert.deepEqual(errors, []);
        console.log('Split-only permissions, single upload, expression grouping, ZIP download and scheme save passed');
    } finally { await browser.close(); }
})().catch(error => {console.error(error); process.exitCode = 1;});
