// Offline scroll behavior regression: no server or AI calls.
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const source = fs.readFileSync('frontend/static/js/training.js', 'utf8');
const frames = [];
const thinking = {scrollTop: 0, scrollHeight: 850};
const code = {scrollTop: 0, scrollHeight: 1200};
const content = {querySelectorAll: () => [thinking, code]};
const container = {scrollTop: 0, scrollHeight: 2300, contains: el => el === content};
const context = vm.createContext({
    document: {getElementById: () => container},
    requestAnimationFrame: callback => { frames.push(callback); return frames.length; },
});
vm.runInContext(source.slice(source.indexOf('const _streamScrollTargets'),
    source.indexOf('function _updateStreamingMessage')), context);
context._scrollStreamingToBottom(content);
context._scrollStreamingToBottom(content);
assert.equal(frames.length, 1, 'coalesce token updates into one frame');
assert.equal(thinking.scrollTop, 0, 'wait for render');
frames.shift()();
assert.equal(thinking.scrollTop, 850);
assert.equal(code.scrollTop, 1200);
assert.equal(container.scrollTop, 2300);
thinking.scrollHeight = 1000;
context._scrollStreamingToBottom(content);
frames.shift()();
assert.equal(thinking.scrollTop, 1000, 'follow later chunks');
context._scrollStreamingToBottom(content);
container.contains = () => false;
container.scrollTop = 12;
frames.shift()();
assert.equal(container.scrollTop, 12, 'ignore previous session messages');
assert.match(source.slice(source.indexOf('function _updateStreamingMessage'),
    source.indexOf('function _finishStreamingMessage')), /_scrollStreamingToBottom\(contentDiv\)/);
assert.match(source.slice(source.indexOf('function _renderCodeStreamOnce'),
    source.indexOf('function _finishCodeStream')), /_scrollStreamingToBottom\(_codeStreamEl\)/);
console.log('Training stream scroll checks passed');
