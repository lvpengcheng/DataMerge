const assert = require('node:assert/strict');
const fs = require('node:fs/promises');
const path = require('node:path');
const {chromium} = require('playwright');
const root = path.resolve(__dirname, '..');
const document = '<!-- DM_RULES BUNDLE START -->\n# 数据处理规则\n## 基本信息\n输出：工资.xlsx、银行.xlsx\n' +
    '<!-- DM_RULES MANIFEST {"version":1,"targets":[]} -->\n## 公共规则\n保留文本主键\n' +
    '## 目标文件：工资.xlsx\n明细：工号、金额\n## 目标文件：银行.xlsx\n明细：银行卡、金额\n' +
    '## 跨目标核对规则\n先核对人员范围\n<!-- DM_RULES BUNDLE END -->';
const event = data => `data: ${JSON.stringify(data)}\n\n`;
(async () => {
    const browser = await chromium.launch({channel: 'msedge', headless: true});
    try {
        const page = await browser.newPage({viewport: {width: 1440, height: 1000}, acceptDownloads: true});
        const errors = [], uploads = [];
        let failChat = false;
        page.on('pageerror', error => errors.push(error.message));
        await page.addInitScript(() => {
            localStorage.setItem('datamerge_token', 'test');
            localStorage.setItem('datamerge_user', JSON.stringify({username: 'test', role_name: 'admin'}));
        });
        await page.route('**/*', async route => {
            const name = new URL(route.request().url()).pathname;
            if (name.startsWith('/static/')) return route.fulfill({path: path.join(root, 'frontend', name)});
            if (name === '/rules') return route.fulfill({path: path.join(root, 'frontend/templates/rules.html'), contentType: 'text/html'});
            if (name === '/api/rules/sessions') return route.fulfill({json: {sessions: []}});
            if (name === '/api/admin/ai-providers') return route.fulfill({json: {items: []}});
            if (name === '/api/files/check-encrypted') return route.fulfill({json: {encrypted_files: []}});
            if (name === '/api/rules/organize/stream') {
                uploads.push(route.request().postDataBuffer().toString('utf8'));
                return route.fulfill({contentType: 'text/event-stream', body:
                    event({type: 'thinking', content: '正在分别整理两个目标…'}) +
                    event({type: 'chunk', content: document}) +
                    event({type: 'complete', content: document, session_id: 1, is_rules_document: true})});
            }
            if (name === '/api/rules/chat') return route.fulfill({contentType: 'text/event-stream', body: failChat
                ? event({type: 'chunk', content: '## 基本信息\n这是不完整的内容'}) + event({type: 'error', message: '本次整理失败'})
                : event({type: 'complete', content: '请说明要修改哪个目标。', is_rules_document: false, session_id: 1})});
            return route.fulfill({json: {}});
        });
        await page.goto('http://datamerge.test/rules');
        // Exercise the real nested scroll areas while thoughts grow and formal output arrives.
        await page.evaluate(() => {
            const message = _addMessage('assistant', '', true);
            let thoughts = '';
            for (let index = 0; index < 3; index++) {
                thoughts += ('思考过程第' + index + '阶段\n').repeat(80);
                _updateStreamingMessage(message, index === 2 ? '规则正文\n'.repeat(100) : undefined, thoughts, true);
                const block = message.querySelector('.thinking-block');
                if (block.scrollHeight <= block.clientHeight ||
                    Math.abs(block.scrollHeight - block.clientHeight - block.scrollTop) > 1) {
                    throw new Error('思考框没有保持在底部');
                }
                const chat = document.getElementById('chat-messages');
                if (Math.abs(chat.scrollHeight - chat.clientHeight - chat.scrollTop) > 1) {
                    throw new Error('对话框没有保持在底部');
                }
                block.scrollTop = 0;
            }
            _clearChatUI();
        });
        assert.equal(await page.locator('#target-file').getAttribute('multiple'), '');
        await page.locator('#source-files').setInputFiles({name: 'source.xlsx', mimeType: 'application/octet-stream', buffer: Buffer.from('source')});
        await page.locator('#target-file').setInputFiles(['工资.xlsx', '银行.xlsx'].map(name => ({name,
            mimeType: 'application/octet-stream', buffer: Buffer.from('target')})));
        await page.locator('#chat-input').fill('请分别整理两个交付文件的规则');
        await page.locator('#send-btn').click();
        await page.waitForFunction(() => !_isStreaming && _lastAIContent.includes('银行.xlsx'));
        assert.equal((uploads[0].match(/name="target_files"/g) || []).length, 2);
        const visible = await page.locator('#chat-messages').innerText();
        assert.ok(visible.includes('正在分别整理两个目标…'), 'thinking must remain visible alongside the rules');
        assert.ok(visible.includes('目标文件：工资.xlsx') && visible.includes('目标文件：银行.xlsx'));
        assert.equal(visible.includes('DM_RULES'), false, 'internal boundaries must not appear in the UI');
        const first = page.waitForEvent('download');
        await page.locator('#download-btn').click();
        const file = await first;
        assert.equal(file.suggestedFilename(), 'rules.md');
        assert.equal(await fs.readFile(await file.path(), 'utf8'), document);
        await page.locator('#chat-input').fill('需要修改哪个文件？');
        await page.locator('#send-btn').click();
        await page.waitForFunction(() => !_isStreaming && document.getElementById('chat-messages').textContent.includes('请说明'));
        assert.equal(await page.evaluate(() => _lastAIContent), document);
        failChat = true;
        await page.locator('#chat-input').fill('请修改工资.xlsx的规则');
        await page.locator('#send-btn').click();
        await page.waitForFunction(() => !_isStreaming && document.getElementById('chat-messages').textContent.includes('请求失败'));
        assert.equal(await page.evaluate(() => _lastAIContent), document, 'failed/partial content cannot replace the last complete rules');
        const again = page.waitForEvent('download');
        await page.locator('#download-btn').click();
        assert.equal(await fs.readFile(await (await again).path(), 'utf8'), document);
        await page.screenshot({path: path.join(root, '.codex-build/multi-target-rules-ui.png'), fullPage: true});
        assert.deepEqual(errors, []);
        console.log('Multi-target upload, readable rules, complete download and preservation after clarification/failure passed');
    } finally {await browser.close();}
})().catch(error => {console.error(error); process.exitCode = 1;});
