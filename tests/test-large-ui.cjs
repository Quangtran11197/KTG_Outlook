// Scan-only synthetic workload; all API calls are intercepted, no Outlook access.
const {chromium}=require('playwright'),cp=require('node:child_process'),path=require('node:path'),fs=require('node:fs'),os=require('node:os'),assert=require('node:assert/strict');
(async()=>{
 const exe=process.env.KTG_SERVER_EXE;if(!exe)throw Error('Set KTG_SERVER_EXE');
 const dir=fs.mkdtempSync(path.join(os.tmpdir(),'ktg-large-ui-')),server=cp.spawn(exe,['--headless','--data-dir',dir],{windowsHide:true,stdio:'pipe'});
 let startup='',browser;server.stdout.on('data',b=>startup+=b);
 try{
  for(let i=0;i<100&&!startup.includes('\n');i++)await new Promise(r=>setTimeout(r,100));
  browser=await chromium.launch({headless:true,executablePath:process.env.KTG_CHROME||'C:/Program Files/Google/Chrome/Application/chrome.exe'});
  const page=await browser.newPage(),errors=[];page.on('pageerror',e=>errors.push(e.message));
  const mails=Array.from({length:70000},(_,i)=>({id:String(i),storeId:'STORE',sourceFolderId:'f'+i%100,subject:i===69999?'unique-last-mail':'Mail '+i,sender:'sender@example.invalid',size:1024,date:'2026-01-01 12:00',folder:'Mailbox / Folder '+i%100,folderSegments:['Mailbox','Folder '+i%100],pstFolderSegments:['Folder '+i%100],internetMessageId:'<'+i+'>'}));
  let scanned=false,requests=0;
  await page.route('**/api/**',async route=>{
   const u=new URL(route.request().url());let body={};
   if(u.pathname==='/api/start'){assert.equal(route.request().postDataJSON().operation,'scan');scanned=true;body={started:true};}
   else if(u.pathname==='/api/state')body={status:scanned?'completed':'idle',processed:scanned?70000:0,total:scanned?70000:0,batches:scanned?35:0,errors:0};
   else if(u.pathname==='/api/batch'){requests++;const n=Number(u.searchParams.get('index'));body=mails.slice(n*2000,(n+1)*2000);}
   await route.fulfill({status:200,contentType:'application/json',body:JSON.stringify(body)});
  });
  await page.goto(startup.split(/\r?\n/)[0]);const start=Date.now();await page.locator('#scan').click();
  await page.waitForFunction(()=>document.querySelector('#count').textContent.replace(/\D/g,'')==='70000',{},{timeout:30000});const loadMs=Date.now()-start;
  assert.equal(requests,35);assert.equal(await page.locator('#rows tr').count(),100);assert.equal(await page.locator('#folders button').count(),101);
  await page.getByRole('button',{name:'Mailbox / Folder 99',exact:true}).click();assert.equal((await page.locator('#count').textContent()).replace(/\D/g,''),'700');
  await page.getByRole('button',{name:'Tất cả thư',exact:true}).click();await page.locator('#search').fill('unique-last-mail');assert.equal(await page.locator('#rows tr').count(),1);
  await page.locator('#search').fill('');await page.locator('#selectAll').click();assert.match((await page.locator('#selection').textContent()).replace(/[,\.]/g,''),/70000 email/);
  await page.locator('#clear').click();assert.deepEqual(errors,[]);
  console.log(JSON.stringify({passed:true,synthetic:true,messages:70000,folders:100,batchRequests:requests,loadMs,renderedRows:100,errors},null,2));
 }finally{if(browser)await browser.close();server.kill();}
})().catch(e=>{console.error(e);process.exitCode=1;});
