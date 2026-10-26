// Offline browser regression: long multi-turn streams, lazy full text and interrupted output.
const assert = require('node:assert/strict');
const path = require('node:path');
const {chromium} = require('playwright');

(async () => {
    const browser = await chromium.launch({channel: 'msedge', headless: true});
    try {
        const page = await browser.newPage({viewport: {width: 1280, height: 900}});
        const errors = [];
        page.on('pageerror', error => errors.push(error.message));
        await page.setContent(`<div class="container"><div class="chat-panel">
            <div id="chat-messages" class="chat-messages"></div>
            <div class="chat-input-area"><textarea id="chat-input"></textarea></div>
            <button id="send-btn"></button><button id="generate-btn"></button>
            <button id="regenerate-btn"></button></div></div>`);
        await page.addStyleTag({path: path.resolve('frontend/static/css/training.css')});
        await page.addScriptTag({path: path.resolve('frontend/static/js/training.js')});
        await page.evaluate(() => {
            window._savingAssets = false;
            window.markdownCalls = 0;
            const original = _renderMarkdown;
            _renderMarkdown = text => { window.markdownCalls++; return original(text); };
        });
        for (let round = 0; round < 3; round++) {
            const result = await page.evaluate(() => {
                let body = '', thinking = '';
                const start = performance.now();
                for (let i = 0; i < 500; i++) {
                    const chunk = `第${i}段规则：金额与人员范围核对。\n`.repeat(12);
                    body += chunk;
                    thinking += chunk;
                    _handleSSEEvent({type: 'chat_chunk', content: chunk});
                    _handleSSEEvent({type: 'thinking', content: chunk});
                }
                window.expectedBody = body;
                window.expectedThinking = thinking;
                return {length: body.length, elapsed: performance.now() - start};
            });
            assert.ok(result.length > 100000);
            await page.waitForTimeout(130);
            assert.equal(await page.evaluate(() => markdownCalls), 0, 'no full markdown parsing during stream');
            assert.ok(await page.evaluate(() =>
                _chatStreamEl.querySelector('.msg-content-body').textContent.length <= 16000 &&
                _chatStreamEl.querySelector('.thinking-block').textContent.length <= 16000));
            await page.locator('#chat-input').fill(`第${round + 1}轮继续输入中文`);
            await page.evaluate(() => _handleSSEEvent({type: 'chat_done', content: expectedBody}));
            const latest = page.locator('.message').last();
            const details = latest.locator('.msg-content-body details');
            assert.equal(await details.locator('div').textContent(), '', 'collapsed full reply has no body DOM');
            await details.locator('summary').click();
            await page.waitForFunction(() => document.querySelector('#chat-messages > .message:last-child .msg-content-body details > div').textContent.length > 100000);
            const visible = await details.locator('div').textContent();
            assert.equal(visible.replace(/\n/g, ''), (await page.evaluate(() => expectedBody)).replace(/\n/g, ''));
            await details.locator('summary').click();
            await page.waitForFunction(() => !document.querySelector('#chat-messages > .message:last-child .msg-content-body details > div').childNodes.length);
            await latest.locator('.thinking-block summary').click();
            await page.waitForFunction(() => !!document.querySelector('#chat-messages > .message:last-child .thinking-block code'));
            assert.equal(await latest.locator('.thinking-block code').textContent(), await page.evaluate(() => expectedThinking));
            await latest.locator('.thinking-block summary').click();
            await page.waitForTimeout(30);
            await page.evaluate(() => { window.markdownCalls = 0; });
            console.log(`Round ${round + 1}: ${result.length} characters; queued in ${Math.round(result.elapsed)}ms; input and full text verified`);
        }
        await page.evaluate(() => {
            const textElement = document.createElement('div');
            let accumulated = '';
            for (let i = 0; i < 1000; i++) {
                accumulated += `第${i}段增量文字。`.repeat(8);
                _appendStreamText(textElement, accumulated);
                if (textElement.textContent !== accumulated.slice(-16000)) throw new Error('Incremental tail mismatch');
            }
            _handleSSEEvent({type: 'chat_chunk', content: '**短回复**'});
            _handleSSEEvent({type: 'chat_done', content: '**短回复**'});
            if (document.querySelector('#chat-messages > .message:last-child strong')?.textContent !== '短回复') {
                throw new Error('Final markdown formatting missing');
            }
            _appendCodeStream('旧代码\n'.repeat(20000));
            _renderCodeStreamOnce();
            if (_codeStreamEl.querySelector('code').textContent.length > 16000) throw new Error('Unbounded code stream');
            _replaceCodeStream('新代码完整替换');
            if (_codeStreamEl.querySelector('code').textContent !== '新代码完整替换') throw new Error('Replacement failed');
            _finishCodeStream();
            _renderIterationCodeSummary({iteration_num: 1, generated_code: '历史完整代码\n'.repeat(20000)});
            _handleSSEEvent({type: 'thinking', content: '断流前思考'});
            _handleSSEEvent({type: 'chat_chunk', content: '断流前正文'});
            _setUIStreaming(false);
            if (_chatThinkingTimer !== null || _chatStreamEl !== null) throw new Error('Pending stream after disconnect');
        });
        await page.waitForTimeout(150);
        assert.ok((await page.locator('#chat-messages').textContent()).includes('断流前正文'));
        const history = page.locator('.iteration-code-summary');
        assert.equal(await history.locator('code').count(), 0, 'historical code loads lazily');
        await history.locator('summary').click();
        await page.waitForFunction(() => !!document.querySelector('.iteration-code-summary code'));
        assert.equal(await history.locator('code').textContent(), '历史完整代码\n'.repeat(20000));
        assert.deepEqual(errors, []);
        console.log('Lazy history, code replacement and interrupted stream cleanup passed');
    } finally { await browser.close(); }
})().catch(error => { console.error(error); process.exitCode = 1; });
