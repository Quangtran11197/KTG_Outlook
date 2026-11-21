// Isolated fixture tests: no Outlook, no real mailbox changes.
const fs=require('node:fs'),path=require('node:path'),os=require('node:os'),cp=require('node:child_process'),assert=require('node:assert/strict');
const {chromium}=require('playwright');
(async()=>{
 const project=path.resolve(__dirname,'..'),dir=process.env.KTG_TEST_DIR||fs.mkdtempSync(path.join(os.tmpdir(),'ktg-bubble-'));
 fs.mkdirSync(dir,{recursive:true});const ready=path.join(dir,'ready.json'),native=process.env.KTG_SERVER_EXE;
 const server=cp.spawn(native||path.join(process.env.SystemRoot,'System32/WindowsPowerShell/v1.0/powershell.exe'),
 native?['--headless','--data-dir',path.join(dir,'data')]:['-NoProfile','-ExecutionPolicy','Bypass','-File',path.join(project,'scripts/server.ps1'),'-NoBrowser','-ReadyFile',ready,'-DataDirectory',path.join(dir,'data')],{windowsHide:true,stdio:'pipe'});
 let startup='',browser;server.stdout.on('data',b=>startup+=b.toString());
 try{
  for(let i=0;i<100&&(native?!startup.includes('\n'):!fs.existsSync(ready));i++)await new Promise(r=>setTimeout(r,100));
  const url=native?startup.split(/\r?\n/)[0]:JSON.parse(fs.readFileSync(ready,'utf8')).url;
  browser=await chromium.launch({headless:true,executablePath:process.env.KTG_CHROME||'C:/Program Files/Google/Chrome/Application/chrome.exe'});
  const page=await browser.newPage({viewport:{width:1656,height:1000},acceptDownloads:true}),errors=[];page.on('pageerror',e=>errors.push(e.message));
  let mode='idle',captured=[],deleteRequests=0,scanRequest=null,autoMode=false,autoRequests=0;const dialogMessages=[];
  const mails=Array.from({length:205},(_,i)=>({id:String(i%103),storeId:i<103?'STORE-A':'STORE-B',sourceFolderId:i<103?'FA':'FB',subject:i===0?'<img src=x onerror=alert(1)>':i===1?'=SUM(1,2)':'Test '+i,sender:'sender@example.invalid',folder:i<103?'Box A / Inbox':'Box B / Sent Items',folderSegments:[i<103?'Box A':'Box B',i<103?'Inbox':'Sent Items'],size:1024,date:'2023-01-01 12:00',internetMessageId:''}));
  await page.route('**/api/**',async route=>{
   const req=route.request(),u=new URL(req.url());let body={};
   if(u.pathname==='/api/start'){
    const data=req.postDataJSON();
    if(data.operation==='scan'){mode='scan';scanRequest=data;}
    else{captured=data.messages;mode=data.operation;if(mode==='export'){assert.equal(data.bulkCopy,true,'Native batching enabled by default');autoMode=data.autoDeleteAfterExport===true;if(autoMode){autoRequests++;assert.equal(data.autoDeleteConfirmed,true);assert.equal(data.autoDeleteCount,captured.length);}}if(['delete','delete_archived'].includes(mode)){deleteRequests++;assert.equal(data.confirmPermanentDelete,true);assert.equal(data.confirmationText,'XOA '+captured.length);if(mode==='delete_archived')assert.equal(data.archiveJobId,'job-100-1');}}
    body={started:true};
   }else if(u.pathname==='/api/state')body={jobId:'job-100-1',status:mode==='idle'?'idle':'completed',processed:mode==='idle'?0:205,total:mode==='idle'?0:205,copied:mode==='export'?captured.length:0,deleted:['delete','delete_archived'].includes(mode)||(mode==='export'&&autoMode)?captured.length:0,errors:0,elapsedSeconds:mode==='idle'?0:1,batches:mode==='scan'?1:0,resultBatches:['export','delete','delete_archived'].includes(mode)?1:0,note:mode==='idle'?'':'Fixture only: no real mail'};
   else if(u.pathname==='/api/batch')body=u.searchParams.get('kind')==='messages'?mails:captured.flatMap(m=>[{key:m.storeId+'|'+m.id,status:['delete','delete_archived'].includes(mode)?'delete_requested':'verified',destinationId:'fake',error:'',archiveProof:{fixtureOnly:'Must stay out of UI result cache'}},...(mode==='export'&&autoMode?[{key:m.storeId+'|'+m.id,status:'delete_requested',error:''}]:[])]);
   else if(u.pathname==='/api/pickpst')body='D:\\MailArchive\\Fixture.pst';
   await route.fulfill({status:200,contentType:'application/json',body:JSON.stringify(body)});
  });
  await page.goto(url);assert.equal(await page.locator('section').count(),5);
  assert.equal(await page.locator('#cancel').textContent(),'Dừng tác vụ');
  assert.equal(await page.locator('#scan').textContent(),'Quét cache Outlook');
  assert.equal(await page.locator('#autoDeleteAfterExport').isChecked(),true,'Auto cleanup enabled by default');
  await page.screenshot({path:path.join(dir,'bubble-empty.png'),fullPage:true});
  await page.locator('#scan').click();await page.waitForFunction(()=>document.querySelector('#count').textContent==='205 email');
  assert.equal(scanRequest.allMail,false);assert.equal(await page.locator('#rows tr').count(),100);assert.equal(await page.locator('#rows img').count(),0);
  await page.getByRole('button',{name:'Box A / Inbox',exact:true}).click();
  assert.equal(await page.locator('#count').textContent(),'103 email');
  await page.locator('#selectAll').click();assert.match(await page.locator('#selection').textContent(),/103 email/);
  await page.locator('#clear').click();await page.getByRole('button',{name:'Tất cả thư',exact:true}).click();
  await page.locator('#search').fill('Test 204');assert.equal(await page.locator('#count').textContent(),'1 email');await page.locator('#search').fill('');
  const downloadPromise=page.waitForEvent('download');await page.locator('#csv').click();const download=await downloadPromise;const csv=fs.readFileSync(await download.path(),'utf8');
  assert.ok(csv.includes('=SUM'));assert.ok(csv.includes('"'+"'=SUM"),'CSV formulas escaped');
  await page.locator('input[name=action][value=delete]').check();await page.locator('#selectAll').click();
  const dialogsBeforeDelete=dialogMessages.length;
  await page.locator('#applyAction').click();
  assert.equal(await page.locator('#deleteDialog').isVisible(),false,'Direct delete skips typed confirmation');
  assert.equal(dialogMessages.length,dialogsBeforeDelete,'Direct delete has no confirmation popup');
  await page.waitForFunction(()=>document.querySelector('#selection').textContent.startsWith('Đã chọn 0 email'));
  assert.equal(deleteRequests,1);assert.equal(captured.length,205);
  // Reset with a new synthetic scan, then check both PST choices and source retention.
  await page.locator('#scan').click();await page.waitForFunction(()=>document.querySelector('#selectAll').disabled===false);
  await page.locator('input[name=action][value=existing]').check();await page.locator('#pickPst').click();await page.waitForFunction(()=>document.querySelector('#pst').value.endsWith('Fixture.pst'));assert.equal(await page.locator('#pst').inputValue(),'D:\\MailArchive\\Fixture.pst');
  await page.locator('#autoDeleteAfterExport').uncheck();
  await page.locator('input[name=action][value=new]').check();await page.locator('#selectAll').click();
  page.on('dialog',d=>{dialogMessages.push(d.message());d.accept();});await page.locator('#applyAction').click();
  await page.waitForFunction(()=>document.querySelector('#selection').textContent.startsWith('Đã chọn 0 email'));
  assert.equal(captured.length,205);assert.equal(deleteRequests,1,'PST never triggers delete');
  assert.equal(await page.evaluate(()=>[...model.results.values()].some(r=>'archiveProof' in r)),false,'UI does not retain audit receipts');
  await page.locator('input[name=action][value=delete]').check();assert.equal(await page.locator('#selectAll').isEnabled(),true,'Exported originals remain selectable for a separate delete');
  assert.equal(await page.locator('#confirmArchiveDelete').isVisible(),false,'New PST does not enable the existing-PST cleanup flow');
  await page.locator('#scan').click();await page.waitForFunction(()=>document.querySelector('#selectAll').disabled===false);
  await page.locator('input[name=action][value=existing]').check();await page.locator('#selectAll').click();await page.locator('#applyAction').click();
  await page.waitForFunction(()=>!document.querySelector('#confirmArchiveDelete').hidden);
  assert.equal(deleteRequests,1,'Existing PST never deletes without post-copy confirmation');
  await page.locator('#confirmArchiveDelete').click();assert.match(await page.locator('#deleteSummary').textContent(),/PST/);
  await page.locator('#deleteDismiss').click();assert.equal(deleteRequests,1);
  await page.locator('#confirmArchiveDelete').click();await page.locator('#deleteAcknowledge').check();await page.locator('#deletePhrase').fill('XOA 205');await page.locator('#deleteConfirm').click();
  await page.waitForFunction(()=>document.querySelector('#confirmArchiveDelete').hidden);
  assert.equal(deleteRequests,2,'Archived cleanup is a distinct confirmed request');
  for(const pstMode of ['existing','new']){
   await page.locator('#scan').click();await page.waitForFunction(()=>!document.querySelector('#selectAll').disabled);
   await page.locator('input[name=action][value='+pstMode+']').check();await page.locator('#autoDeleteAfterExport').check();await page.locator('#selectAll').click();
   const countBefore=dialogMessages.length;await page.locator('#applyAction').click();
   await page.waitForFunction(()=>document.querySelector('#selection').textContent.startsWith('Đã chọn 0 email'));
   assert.equal(dialogMessages.length,countBefore+1,'Only the initial confirmation dialog');
   assert.match(dialogMessages.at(-1),/TỰ ĐỘNG XÓA/);
   assert.equal(await page.locator('#deleteDialog').isVisible(),false);
   assert.equal(await page.locator('#confirmArchiveDelete').isVisible(),false);
   assert.equal(deleteRequests,2,'No second cleanup request');
  }
  assert.equal(autoRequests,2,'Both PST modes support automatic cleanup');
  await page.screenshot({path:path.join(dir,'bubble-desktop.png'),fullPage:true});
  await page.setViewportSize({width:390,height:844});await page.screenshot({path:path.join(dir,'bubble-mobile.png'),fullPage:true});assert.deepEqual(errors,[]);
  console.log(JSON.stringify({passed:true,native:!!native,fixtureMessages:205,deleteRequests,autoRequests,javaScriptErrors:errors,screenshots:dir},null,2));
 }finally{if(browser)await browser.close();server.kill();}
})().catch(e=>{console.error(e);process.exitCode=1;});
