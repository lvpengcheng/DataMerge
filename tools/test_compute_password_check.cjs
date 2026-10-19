const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('frontend/static/js/compute.js', 'utf8');
const btn = {};
let prompts = 0;
let cancelled = false;
const requested = [];
const ctx = vm.createContext({Uint8Array, console,
    document: {getElementById: () => btn},
    alert() {}, checkCanCompute() {},
    _promptFilePasswords: async names => { prompts++; requested.push([...names]); return cancelled ? null : Object.fromEntries(names.map(n => [n, 'test'])); },
});
vm.runInContext('let _filePasswordsMap = null; let _encryptionCheckInProgress = false; let _encryptionChecksPending = 0; let _encryptionCheckQueue = Promise.resolve();\n' +
    source.slice(source.indexOf('async function _probeModernExcelEncryption'), source.indexOf('// 页面加载完成后初始化')), ctx);
const file = (name, bytes) => ({name, slice: () => ({arrayBuffer: async () => new Uint8Array(bytes).buffer})});
(async () => {
    const encrypted = file('template.xlsx', [0xd0, 0xcf, 0x11, 0xe0]);
    assert.equal(await ctx._autoCheckEncryption([encrypted]), true);
    assert.equal(prompts, 1);
    assert.equal(await ctx._autoCheckEncryption([encrypted]), true);
    assert.equal(prompts, 1, 'do not repeat password prompt before submit');
    cancelled = true;
    assert.equal(await ctx._autoCheckEncryption([file('new.xlsx', [0xd0, 0xcf, 0x11, 0xe0])]), false);
    assert.equal(await ctx._autoCheckEncryption([file('plain.xlsx', [0x50, 0x4b])]), true);
    cancelled = false;
    const second = file('summary.xlsx', [0xd0, 0xcf, 0x11, 0xe0]);
    assert.equal(await ctx._autoCheckEncryption([encrypted, second], 'template'), true);
    assert.deepEqual(requested.at(-1), ['template::template.xlsx', 'template::summary.xlsx']);
    const before = prompts;
    assert.equal(await ctx._autoCheckEncryption([encrypted, second], 'template'), true);
    assert.equal(prompts, before, 'each template password is retained separately');
    console.log('Password detection and cancellation checks passed');
})().catch(e => { console.error(e); process.exitCode = 1; });
