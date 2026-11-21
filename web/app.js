'use strict';
const $=id=>document.getElementById(id);
const token=new URLSearchParams(location.search).get('token')||sessionStorage.getItem('ktg-token')||'';
sessionStorage.setItem('ktg-token',token);history.replaceState(null,'',location.pathname);
const model={messages:[],byKey:new Map(),selected:new Set(),results:new Map(),folder:'',page:0,busy:false,mode:'',batches:0,resultBatches:0,polling:false,logged:'',request:0,switching:false,closing:false,picking:false,pendingDelete:[],archive:null,exportPstMode:'',pendingArchiveJob:null,autoCleanup:false,bulkMode:false,folderCounts:new Map(),filterCache:null};
const identity=m=>m.storeId+'|'+m.id;
const action=()=>document.querySelector('input[name=action]:checked').value;
const leaf=m=>m.folderSegments?.at(-1)||m.folder?.split(' / ').at(-1)||'Mailbox';
const bytes=n=>n>=1073741824?(n/1073741824).toFixed(2)+' GB':n>=1048576?(n/1048576).toFixed(2)+' MB':(n/1024).toFixed(1)+' KB';
const selectable=m=>!m.deleted&&!m.needsRescan&&(action()==='delete'||!m.exported);
async function api(path,body){
 const res=await fetch(path,{method:body===undefined?'GET':'POST',headers:{'x-ktg-token':token,...(body===undefined?{}:{'Content-Type':'application/json'})},body:body===undefined?undefined:JSON.stringify(body)});
 const value=await res.json();if(!res.ok||value?.error)throw new Error(value?.error||'Không kết nối được ứng dụng.');return value;
}
function log(note){$('log').textContent+=(($('log').textContent?'\n':'')+new Date().toLocaleTimeString()+' · '+note);$('log').scrollTop=$('log').scrollHeight;}
function filtered(){
 const q=$('search').value.toLocaleLowerCase(),cache=model.filterCache;
 if(cache&&cache.q===q&&cache.folder===model.folder&&cache.length===model.messages.length)return cache.rows;
 const rows=!q&&!model.folder?model.messages:model.messages.filter(m=>(!model.folder||m.folder===model.folder)&&(!q||(m.subject+' '+m.sender+' '+m.folder).toLocaleLowerCase().includes(q)));
 model.filterCache={q,folder:model.folder,length:model.messages.length,rows};return rows;
}
function cell(tr,text){const td=document.createElement('td');td.textContent=text;tr.append(td);return td;}
function setBusy(busy){
 model.busy=busy;for(const id of ['scan','applyAction','quit','pickPst'])$(id).disabled=busy||model.closing||model.picking;
 $('confirmArchiveDelete').disabled=busy||model.closing;
 $('autoDeleteAfterExport').disabled=busy||model.closing;
 $('cancel').disabled=!busy||model.closing;for(const radio of document.querySelectorAll('input[name=action]'))radio.disabled=busy||model.closing;
 $('clear').disabled=busy||model.closing;
 if(busy){for(const id of ['selectPage','selectAll','clear'])$(id).disabled=true;for(const input of $('rows').querySelectorAll('input'))input.disabled=true;}
 if(model.closing)for(const b of document.querySelectorAll('button'))b.disabled=true;
}
window.ktgRequestClose=()=>{model.closing=true;setBusy(true);log('Đang dừng tác vụ để đóng ứng dụng. Chờ Outlook hoàn thành thao tác hiện tại.');$('status').textContent='Đang đóng';};
function updateCount(){
 let size=0;for(const key of model.selected)size+=Number(model.byKey.get(key)?.size||0);
 $('selection').textContent='Đã chọn '+model.selected.size.toLocaleString()+' email · '+bytes(size)+' (ước tính)';$('clear').hidden=!model.selected.size;
}
function render(){
 const all=filtered(),pages=Math.max(1,Math.ceil(all.length/100));model.page=Math.min(model.page,pages-1);$('rows').replaceChildren();
 const page=all.slice(model.page*100,model.page*100+100);
 for(const m of page){
  const key=identity(m),result=model.results.get(key),tr=document.createElement('tr'),td=cell(tr,'');const check=document.createElement('input');
  check.type='checkbox';check.checked=model.selected.has(key);check.disabled=!selectable(m)||model.busy;check.setAttribute('aria-label','Chọn '+m.subject);
  check.onchange=()=>{check.checked?model.selected.add(key):model.selected.delete(key);updateCount();};td.append(check);
  cell(tr,m.date);const folder=cell(tr,leaf(m));folder.title=m.folder;cell(tr,m.sender);cell(tr,m.subject);cell(tr,bytes(Number(m.size||0)));
  cell(tr,m.deleted?'Đã yêu cầu xóa':result?.error|| (m.exported?'Đã kiểm tra PST':'Chưa xuất'));$('rows').append(tr);
 }
 $('count').textContent=all.length.toLocaleString()+' email';
 $('pageInfo').textContent=all.length?(all.length.toLocaleString()+' email · Trang '+(model.page+1)+' / '+pages):'Không có email phù hợp';
 $('prev').disabled=model.page===0;$('next').disabled=model.page===pages-1;
 $('selectPage').textContent='Chọn '+page.filter(selectable).length+' email trên trang';
 $('selectPage').disabled=!page.some(selectable)||model.busy;$('selectAll').disabled=!all.some(selectable)||model.busy;
 $('selectAll').textContent='Chọn tất cả email trong '+(model.folder?leaf({folder:model.folder}):'tất cả thư');
 $('csv').disabled=!model.messages.length;updateCount();
}
function updateFolders(){
 const counts=model.folderCounts;
 if(model.folder&&!counts.has(model.folder))model.folder='';$('folders').replaceChildren();
 function add(label,value,count){
  const b=document.createElement('button');b.type='button';b.title=value||'Tất cả thư';b.setAttribute('aria-pressed',String(model.folder===value));b.setAttribute('aria-label',value||'Tất cả thư');
  b.append(document.createTextNode(label));const badge=document.createElement('span');badge.className='folder-count';badge.textContent=count.toLocaleString();b.append(badge);
  b.onclick=()=>{model.folder=value;model.page=0;updateFolders();render();};$('folders').append(b);
 }
 add('Tất cả thư','',model.messages.length);
 for(const [folder,count]of [...counts].sort((a,b)=>a[0].localeCompare(b[0])))add(leaf({folder}),folder,count);
}
function updateAction(){
 const deleting=action()==='delete';$('pstFields').hidden=deleting;
 $('actionHint').textContent=deleting?'Bấm Thực hiện lựa chọn để xóa ngay các email đã chọn và xóa tiếp trong Deleted Items. Không hỏi xác nhận; không dọn những email khác.':$('autoDeleteAfterExport').checked?'Đang bật tự xóa: xác minh bản PST rồi hard-delete đúng thư nguồn Exchange, không chuyển qua Deleted Items. Thư sao chép lỗi hoặc đối chiếu không đạt sẽ được giữ lại.':'Đang tắt tự xóa: lưu PST và giữ thư nguồn. Với PST có sẵn, bạn có thể xác nhận xóa riêng sau.';
 for(const key of model.selected){const m=model.byKey.get(key);if(!m||!selectable(m))model.selected.delete(key);}render();
}
async function start(request){
 if(model.busy||model.closing)return;model.request++;model.switching=true;setBusy(true);
 try{
  await api('/api/start',request);model.mode=request.operation;model.batches=0;model.resultBatches=0;model.logged='';
  if(request.operation==='export'){model.bulkMode=request.bulkCopy===true;model.exportPstMode=request.pstMode;model.archive=null;model.autoCleanup=request.autoDeleteAfterExport===true;}
  if(request.operation==='scan'){model.messages=[];model.folderCounts.clear();model.filterCache=null;model.byKey.clear();model.selected.clear();model.results.clear();model.folder='';model.page=0;model.archive=null;updateFolders();render();}
  $('confirmArchiveDelete').hidden=true;
  $('report').hidden=true;log(request.operation==='scan'?'Bắt đầu quét cache Exchange.':(['delete','delete_archived'].includes(request.operation)?'Bắt đầu xóa các email đã chọn: ':'Bắt đầu sao chép PST: ')+request.messages.length+' email.');$('status').textContent='Đang chạy';
 }catch(e){setBusy(false);log(e.message);throw e;}finally{model.request++;model.switching=false;}
}
async function poll(){
 if(model.polling||model.switching||model.picking||model.closing)return;model.polling=true;const epoch=model.request;
 try{
  const s=await api('/api/state');if(epoch!==model.request)return;let changed=false;
  while(model.batches<(s.batches||0)){
   const end=Math.min(model.batches+4,s.batches),indices=Array.from({length:end-model.batches},(_,i)=>model.batches+i);
   const chunks=await Promise.all(indices.map(index=>api('/api/batch?kind=messages&index='+index)));if(epoch!==model.request)return;
   for(const chunk of chunks)for(const m of chunk){const key=identity(m);if(!model.byKey.has(key)){model.messages.push(m);model.byKey.set(key,m);model.folderCounts.set(m.folder,(model.folderCounts.get(m.folder)||0)+1);}}
   model.batches=end;changed=true;
  }
  while(model.resultBatches<(s.resultBatches||0)){
   const chunk=await api('/api/batch?kind=results&index='+model.resultBatches);if(epoch!==model.request)return;
   for(const result of chunk){const {archiveProof,...summary}=result;model.results.set(result.key,summary);const m=model.byKey.get(result.key);
    if(m&&result.status==='verified'){m.exported=true;if(model.mode==='export'&&model.exportPstMode==='existing'&&!model.autoCleanup){if(!model.archive)model.archive={jobId:s.jobId,keys:new Set()};model.archive.keys.add(result.key);}}
    if(m&&result.status==='delete_requested')m.deleted=true;if(m&&result.needsRescan)m.needsRescan=true;
    if(result.status==='verified'||result.status==='delete_requested'||result.needsRescan)model.selected.delete(result.key);
   }model.resultBatches++;changed=true;
  }
  const wasBusy=model.busy,running=s.status==='running';setBusy(running);if(changed){updateFolders();render();}else if(wasBusy!==running)render();
  const names={idle:'IDLE',running:'Đang chạy',completed:'Hoàn tất',completed_with_errors:'Xong, có lỗi',cancelled:'Đã dừng',error:'Lỗi'};
  $('status').textContent=names[s.status]||s.status;const secs=Number(s.elapsedSeconds||0),done=Number(s.processed||0);
  $('progress').textContent=done.toLocaleString()+'/'+(s.total?s.total.toLocaleString():'?')+' · Lỗi/cảnh báo: '+(s.errors||0)+(secs?' · '+secs.toFixed(1)+' giây':'')+(model.mode==='export'?' · PST đã kiểm tra: '+(s.copied||0):'')+(['delete','delete_archived'].includes(model.mode)?' · Đã yêu cầu xóa: '+(s.deleted||0):'');
  if(s.bulkBatches)$('progress').textContent+=' · Lô MAPI: '+s.bulkBatches;
  if(model.mode==='export'&&model.autoCleanup)$('progress').textContent+=' · Tự xóa nguồn: '+(s.deleted||0)+(s.phase==='cleanup'?' · Đang xử lý '+(s.cleanupProcessed||0)+'/'+(s.cleanupTotal||0):'');
  $('bar').hidden=!running;if(running&&!s.total)$('bar').removeAttribute('value');else{$('bar').max=Math.max(1,s.total||done);$('bar').value=done;}
  if(s.note&&s.note!==model.logged){log(s.note);model.logged=s.note;}
  if(s.logPath)$('logPath').textContent='Nhật ký: '+s.logPath;
  $('report').hidden=running||!['export','delete','delete_archived'].includes(model.mode);$('report').disabled=running;
  const archived=archiveCandidates();$('confirmArchiveDelete').hidden=running||!archived.length;
  $('confirmArchiveDelete').textContent='Xác nhận đã lưu PST và xóa '+archived.length+' email nguồn';
 }catch(e){$('progress').textContent=e.message;}finally{model.polling=false;}
}
function selected(){return model.messages.filter(m=>model.selected.has(identity(m))&&selectable(m));}
function archiveCandidates(){return model.archive?model.messages.filter(m=>model.archive.keys.has(identity(m))&&m.exported&&!m.deleted&&!m.needsRescan):[];}
function showDeleteDialog(messages,archiveJobId=null){
 model.pendingDelete=messages;model.pendingArchiveJob=archiveJobId;
 $('deleteSummary').textContent=archiveJobId?'Đã kiểm tra '+messages.length+' bản trong PST. Xác nhận xóa các email nguồn tương ứng. Ứng dụng sẽ mở lại bản PST và kiểm tra thêm một lần trước khi xóa.':'Bạn đang chọn '+messages.length.toLocaleString()+' email để xóa riêng.';
 $('deletePhraseLabel').textContent='Nhập chính xác: XOA '+messages.length;$('deletePhrase').value='';$('deleteAcknowledge').checked=false;$('deleteConfirm').disabled=true;$('deleteDialog').showModal();
}
$('scan').onclick=()=>start({operation:'scan',allMail:$('allMail').checked,from:$('from').value,to:$('to').value}).catch(()=>{});
$('allMail').onchange=()=>{$('from').disabled=$('to').disabled=$('allMail').checked;};
$('cancel').onclick=async()=>{try{await api('/api/cancel',{});log('Đã yêu cầu dừng tác vụ; chờ Outlook hoàn thành thao tác hiện tại.');}catch(e){log(e.message);}};
for(const radio of document.querySelectorAll('input[name=action]'))radio.onchange=updateAction;
$('autoDeleteAfterExport').onchange=updateAction;
$('applyAction').onclick=async()=>{
 const messages=selected();if(!messages.length){alert('Hãy chọn email phù hợp với thao tác.');return;}
 if(action()==='delete'){
  try{await start({operation:'delete',messages,confirmPermanentDelete:true,confirmationText:'XOA '+messages.length});}catch(e){alert(e.message);}return;
 }
 if(!$('pst').value.trim()){alert('Hãy nhập đường dẫn PST đầy đủ.');return;}
 const autoDeleteAfterExport=$('autoDeleteAfterExport').checked;
 const question=autoDeleteAfterExport?'Lưu '+messages.length.toLocaleString()+' email vào PST, kiểm tra bản sao rồi HARD-DELETE vĩnh viễn thư nguồn Exchange (bỏ qua Deleted Items)?\nKhông hỏi lại sau khi lưu. Chỉ các bản sao kiểm tra đạt mới được xóa nguồn.':'Sao chép '+messages.length.toLocaleString()+' email vào PST và GIỮ thư nguồn?';
 if(!confirm(question))return;
 try{await start({operation:'export',pst:$('pst').value.trim(),pstMode:action(),messages,bulkCopy:$('bulkCopy').checked,autoDeleteAfterExport,...(autoDeleteAfterExport?{autoDeleteConfirmed:true,autoDeleteCount:messages.length}:{})});}catch(e){alert(e.message);}
};
function checkDelete(){$('deleteConfirm').disabled=!$('deleteAcknowledge').checked||$('deletePhrase').value!=='XOA '+model.pendingDelete.length;}
$('deleteAcknowledge').onchange=checkDelete;$('deletePhrase').oninput=checkDelete;
$('deleteDismiss').onclick=()=>{$('deleteDialog').close();model.pendingDelete=[];};
$('deleteDialog').onclose=()=>{model.pendingDelete=[];};
$('confirmArchiveDelete').onclick=()=>{const messages=archiveCandidates();if(messages.length&&model.archive?.jobId)showDeleteDialog(messages,model.archive.jobId);};
$('deleteForm').onsubmit=async e=>{
 e.preventDefault();if($('deleteConfirm').disabled)return;
 const messages=[...model.pendingDelete],confirmationText=$('deletePhrase').value,archiveJobId=model.pendingArchiveJob;$('deleteDialog').close();
 try{await start({operation:archiveJobId?'delete_archived':'delete',messages,confirmPermanentDelete:true,confirmationText,...(archiveJobId?{archiveJobId}:{})});}catch(err){alert(err.message);}
};
$('search').oninput=()=>{model.page=0;render();};
$('selectPage').onclick=()=>{const rows=filtered().slice(model.page*100,model.page*100+100).filter(selectable),all=rows.every(m=>model.selected.has(identity(m)));for(const m of rows)all?model.selected.delete(identity(m)):model.selected.add(identity(m));render();};
$('selectAll').onclick=()=>{const rows=filtered().filter(selectable),all=rows.every(m=>model.selected.has(identity(m)));for(const m of rows)all?model.selected.delete(identity(m)):model.selected.add(identity(m));render();};
$('clear').onclick=()=>{model.selected.clear();render();};
$('prev').onclick=()=>{model.page--;render();};$('next').onclick=()=>{model.page++;render();};
function download(blob,name){const url=URL.createObjectURL(blob),a=document.createElement('a');a.href=url;a.download=name;a.click();setTimeout(()=>URL.revokeObjectURL(url),5000);}
function csvCell(value){let s=String(value??'');if(/^[=+@\-\t\r\n]/.test(s))s="'"+s;return '"'+s.replaceAll('"','""')+'"';}
$('csv').onclick=()=>{
 const rows=[['Ngày','Thư mục','Người gửi','Tiêu đề','Dung lượng (byte)','StoreID','EntryID','Trạng thái']];
 for(const m of model.messages)rows.push([m.date,m.folder,m.sender,m.subject,m.size,m.storeId,m.id,m.deleted?'delete_requested':m.exported?'verified':'scanned']);
 download(new Blob(['\ufeff'+rows.map(r=>r.map(csvCell).join(',')).join('\r\n')],{type:'text/csv;charset=utf-8'}),'KTG-email-scan.csv');
};
$('report').onclick=async()=>{try{const r=await fetch('/api/report',{headers:{'x-ktg-token':token}});if(!r.ok)throw new Error('Không tải được báo cáo');download(await r.blob(),'KTG-operation-report.csv');}catch(e){log(e.message);}};
$('pickPst').onclick=async()=>{
 model.picking=true;setBusy(model.busy);
 try{const path=await api('/api/pickpst?mode='+action(),{});if(typeof path==='string')$('pst').value=path;}
 catch(e){log('Không chọn được file; bạn có thể nhập đường dẫn trực tiếp. '+e.message);}
 finally{model.picking=false;setBusy(model.busy);}
};
$('quit').onclick=async()=>{try{await api('/api/quit',{});clearInterval(timer);$('status').textContent='Đã đóng';for(const b of document.querySelectorAll('button'))b.disabled=true;}catch(e){log(e.message);}};
const timer=setInterval(poll,1000);updateFolders();updateAction();poll();
