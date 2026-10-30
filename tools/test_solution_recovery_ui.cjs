const assert=require('node:assert/strict'),path=require('node:path');
const {chromium}=require('playwright');
(async()=>{
 const browser=await chromium.launch({channel:'msedge',headless:true});
 try{
  const page=await browser.newPage();const requests=[];
  await page.setContent('<input id="tenant-input" value="满记迁移测试">');
  await page.evaluate(()=>{window.AUTH={authFetch:async(url,options)=>{
   window.calls=(window.calls||[]);window.calls.push({url,body:options?.body});
   if(url==='/api/solutions/import'){
    window.importTargets=window.importTargets||[];window.importTargets.push(options.body.get('script_id'));
   }
   if(url==='/api/tenants')return new Response(JSON.stringify({tenants:[]}));
   if(options?.method==='POST')return new Response(JSON.stringify({id:'new'}));
   return new Response(JSON.stringify({scripts:[{id:'solution_old',name:'满记迁移1028',needs_restore:true,bound_script_id:964},
     {id:972,name:'现有脚本'}],
    solutions:[{id:'old',name:'满记迁移1028',script_id:964,is_active:true}]}));
  }};});
  await page.addScriptTag({path:path.resolve('frontend/static/js/solutions.js')});
  await page.evaluate(()=>openSolutions());
  assert.equal(await page.getByRole('button',{name:'转为方案',exact:true}).isEnabled(),true);
  assert.equal(await page.getByRole('button',{name:'进入修正',exact:true}).count(),1);
  assert.match(await page.locator('.solutions-content select').innerText(),/从迁移方案恢复/);
  await page.getByRole('button',{name:'转为方案',exact:true}).click();
  await page.waitForFunction(()=>window.calls.some(c=>c.url==='/api/solutions'&&c.body));
  const call=await page.evaluate(()=>window.calls.find(c=>c.url==='/api/solutions'&&c.body));
  assert.equal(JSON.parse(call.body).script_id,'solution_old');
  await page.waitForFunction(()=>document.querySelector('dialog').dataset.busy==='false');
  const target=page.locator('#solutions-import-script');
  assert.equal(await target.isEnabled(),true);
  assert.match(await target.innerText(),/满记迁移1028（覆盖并恢复脚本）/);
  assert.match(await target.innerText(),/现有脚本/);
  await target.selectOption('solution_old');
  await page.getByRole('button',{name:'加载脚本和方案',exact:true}).click();
  await page.waitForFunction(()=>document.querySelector('dialog').dataset.busy==='false');
  assert.equal(await target.inputValue(),'solution_old');
  await page.locator('input[type=file]').setInputFiles({name:'方案.zip',mimeType:'application/zip',buffer:Buffer.from('fixture')});
  await page.locator('.solutions-checkbox input').nth(1).check();
  await page.evaluate(()=>{openSolutionTraining=async()=>{};});
  for(const id of ['solution_old','972']){
   await target.selectOption(id);
   await page.getByRole('button',{name:'导入方案',exact:true}).click();
   await page.waitForFunction(()=>document.querySelector('dialog').dataset.busy==='false');
  }
  assert.deepEqual(await page.evaluate(()=>window.importTargets),['solution_old','972']);
  console.log('Migrated snapshot remains selectable and convertible when its editable script is missing');
  console.log('Import target accepts both existing scripts and archived solutions, retaining selection after reload');
 }finally{await browser.close();}
})().catch(e=>{console.error(e);process.exitCode=1;});
