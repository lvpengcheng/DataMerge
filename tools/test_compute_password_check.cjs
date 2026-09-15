const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('frontend/static/js/compute.js', 'utf8');
const btn = {};
let prompts = 0;
let cancelled = false;
const ctx = vm.createContext({Uint8Array, console,
    document: {getElementById: () => btn},
    alert() {},
    _promptFilePasswords: async names => { prompts++; return cancelled ? null : Object.fromEntries(names.map(n => [n, 'test'])); },
});
vm.runInContext('let _filePasswordsMap = null; let _encryptionCheckInProgress = false;\n' +
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
    console.log('Password detection and cancellation checks passed');
})().catch(e => { console.error(e); process.exitCode = 1; });
