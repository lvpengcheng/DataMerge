const assert=require('node:assert/strict');
const path=require('node:path');
const {chromium}=require('playwright');
const root=path.resolve(__dirname,'..');
const formula='IF((其他&"")="转移",(银行账号&""),"")';
(async()=>{
 const browser=await chromium.launch({channel:'msedge',headless:true});
 try {
  for(const kind of ['integrate','split']) {
   const page=await browser.newPage();
   const errors=[],requests=[];let mode='success',release,arrived;
   page.on('pageerror',e=>errors.push(e.message));
   await page.addInitScript(()=>{
    localStorage.setItem('datamerge_token','test');
    localStorage.setItem('datamerge_user',JSON.stringify({username:'test',role_name:'admin'}));
   });
   await page.route('**/*',async route=>{
    const url=new URL(route.request().url());
    if(url.pathname.startsWith('/static/'))return route.fulfill({path:path.join(root,'frontend',url.pathname)});
    if(url.pathname==='/tools')return route.fulfill({path:path.join(root,'frontend/templates/tools.html'),contentType:'text/html'});
    if(url.pathname.endsWith('/generate-formula')) {
     requests.push({url:url.pathname,body:route.request().postDataJSON()});
     if(mode==='held')await new Promise(resolve=>{release=resolve;arrived();});
     return route.fulfill({json:mode==='clarify'?{formula:null,needs_clarification:true,message:'请说明没有匹配员工时如何处理'}:
      {formula,explanation:'按转移标记回填银行卡号'}});
    }
    return route.fulfill({json:[]});
   });
   await page.goto('http://datamerge.test/tools');
   await page.locator(`[data-tab="${kind==='split'?'table-split':'data-integrate'}"]`).click();
   await page.evaluate(()=>{
    Tools._intShowWork('create');
    _integrateAnalysis={session_id:'f'.repeat(32),files:[
     {name:'template.xlsx',columns:['员工号','银行卡号'],sheet:'Sheet1',suggested_key:'员工号'},
     {name:'source.xlsx',columns:['员工号','其他','金额','银行账号'],sheet:'Sheet1',suggested_key:'员工号'}]};
    Tools._renderIntegrateConfig(_integrateAnalysis);
    document.getElementById('int-union-source-keys').checked=true;
    Tools._intAddListRow('ow','银行卡号');
   });
   await page.locator('#int-ow-list-rows .int-src-cell').click();
   assert.equal(await page.locator('#int-formula-description').isVisible(),false);
   await page.locator('#int-formula-ai-panel > summary').click();
   assert.equal(await page.locator('#int-formula-description').getAttribute('rows'),'10');
   assert.equal(await page.locator('#int-formula-description').getAttribute('maxlength'),'20000');
   await page.locator('#int-generate-formula').click();
   assert.match(await page.locator('#int-formula-generation-status').innerText(),/请先输入/);
   await page.locator('#int-formula-description').fill('有转移标记返回银行卡号，否则不覆盖');
   await page.locator('#int-formula-expected-result').fill('有转移标记时应返回银行卡，不能返回0');
   await page.locator('#int-generate-formula').click();
   await page.waitForFunction(()=>document.getElementById('int-unified-fx').value.startsWith('IF('));
   assert.equal(await page.locator('#int-unified-fx').inputValue(),formula);
   assert.equal(await page.locator('.int-srcpick-cb[data-token="银行账号"]').isChecked(),true);
   await page.locator('#modal-confirm').click();
   const pair=await page.evaluate(()=>Tools._readSectionPairs('ow')[0]);
   assert.equal(pair.source_expr,formula,'AI business logic must not acquire a template prefix');
   assert.equal(pair.formula_description,'有转移标记返回银行卡号，否则不覆盖');
   assert.equal(pair.formula_exact,true);
   assert.equal(pair.formula_expected_result,'有转移标记时应返回银行卡，不能返回0');
   assert.equal(requests[0].url,`/api/tools/${kind==='split'?'table-split':'integrate'}/generate-formula`);
   assert.equal(requests[0].body.target_col,'银行卡号');
   assert.equal(requests[0].body.expected_result,pair.formula_expected_result);
   await page.locator('#int-ow-list-rows .int-src-cell').click();
   assert.equal(await page.locator('#int-formula-description').isVisible(),false);
   await page.locator('#int-formula-ai-panel > summary').click();
   assert.equal(await page.locator('#int-formula-description').inputValue(),pair.formula_description);
   assert.equal(await page.locator('#int-formula-expected-result').inputValue(),pair.formula_expected_result);
   mode='clarify';
   await page.locator('#int-generate-formula').click();
   await page.waitForFunction(()=>document.getElementById('int-formula-generation-status').textContent.includes('请说明'));
   assert.equal(await page.locator('#int-unified-fx').inputValue(),formula);
   mode='held';
   const pending=new Promise(resolve=>{arrived=resolve;});
   await page.locator('#int-generate-formula').click();await pending;
   await page.locator('#int-unified-fx').fill('"手工值"');release();
   await page.waitForFunction(()=>document.getElementById('int-formula-generation-status').textContent.includes('已保留'));
   assert.equal(await page.locator('#int-unified-fx').inputValue(),'"手工值"');
   await page.locator('.modal-close').click();
   assert.equal((await page.evaluate(()=>Tools._readSectionPairs('ow')[0])).source_expr,formula);
   assert.deepEqual(errors,[]);
   await page.close();
  }
  console.log('Both workflows: AI description/generation, field sync, exact formulas, reopen, clarification, cancellation and stale-response protection passed');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
