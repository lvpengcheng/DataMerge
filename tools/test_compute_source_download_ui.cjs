const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const alerts = [];
const context = vm.createContext({
    document:{addEventListener() {}},
    alert: message => alerts.push(message),
    AUTH:{authFetch:async () => ({ok:false,status:404,statusText:'Not Found',
        json:async () => ({detail:'源文件未保留或已清理，无法下载'})})},
    Date, console,
});
vm.runInContext(fs.readFileSync(path.join(__dirname,'../frontend/static/js/tools.js'),'utf8'),context);
(async () => {
    await vm.runInContext('Tools.downloadAsset(10,"工资.xlsx")',context);
    await vm.runInContext('Tools._fetchAndDownload("/api/assets/10/download?format=encrypted","工资.xlsx")',context);
    assert.deepEqual(alerts,Array(2).fill('下载失败: 源文件未保留或已清理，无法下载'));
    console.log('History downloads display the server explanation for missing sources');
})().catch(error => {console.error(error);process.exitCode=1;});
