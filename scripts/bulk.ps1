# Copy-only native MAPI capability. Original deletion remains in guarded core.ps1.
function Initialize-BulkMapi([string]$Dll) {
 if('KtgNativeBulk.Api' -as [type]){return}
 if(-not [Environment]::Is64BitProcess){throw 'Chế độ MAPI theo lô cần Outlook/PowerShell 64-bit trên bản này. Tắt tùy chọn theo lô để dùng chế độ cũ.'}
 if(-not [IO.File]::Exists($Dll)){throw 'Thiếu bộ xử lý MAPI. Hãy giải nén đầy đủ bản phát hành mới.'}
 $safe=$Dll.Replace('"','""')
 Add-Type -TypeDefinition @"
using System;using System.Runtime.InteropServices;
namespace KtgNativeBulk {public static class Api {
[DllImport(@"$safe",CharSet=CharSet.Unicode)] public static extern int KtgOpen(string profile,out IntPtr handle);
[DllImport(@"$safe")] public static extern void KtgClose(IntPtr handle);
[DllImport(@"$safe",CharSet=CharSet.Unicode)] public static extern int KtgPrepareCleanup(IntPtr h,string s,string f,string d,string target,string root,string paths,string rows,[MarshalAs(UnmanagedType.BStr)]out string result);
[DllImport(@"$safe",CharSet=CharSet.Unicode)] public static extern int KtgCleanupStep(IntPtr h,string token,int step,[MarshalAs(UnmanagedType.BStr)]out string result);
[DllImport(@"$safe",CharSet=CharSet.Unicode)] public static extern int KtgCopy(IntPtr h,string s,string f,string d,string t,string ids,[MarshalAs(UnmanagedType.BStr)]out string result);
}}
"@
}
function Wait-BulkGate($Context) {
 $pause=Join-Path $Context.Directory 'pause';$note=$Context.Note
 while([IO.File]::Exists($pause)){
  if(Test-Cancelled $Context){$Context.Status='cancelled';return $false}
  $Context.Paused=$true;$Context.Note='Đã tạm dừng giữa các bước an toàn. Có thể tiếp tục sử dụng Outlook.';Save-JobState $Context
  [Threading.Thread]::Sleep(200)
 }
 if($Context.Paused){$Context.Paused=$false;$Context.Note=$note;Save-JobState $Context -Force}
 if(Test-Cancelled $Context){$Context.Status='cancelled';return $false}
 return $true
}
function Merge-NativeCopyManifests($Partials,$Failures) {
 $copies=New-Object 'Collections.Generic.List[object]';$skipped=New-Object 'Collections.Generic.List[object]';$errors=New-Object 'Collections.Generic.List[string]'
 foreach($failure in $Failures){if($failure){$errors.Add([string]$failure)}}
 foreach($partial in $Partials){
  foreach($copy in @($partial.copies)){$copies.Add($copy)}
  foreach($item in @($partial.skipped)){$skipped.Add($item)}
  if($partial.error){$errors.Add([string]$partial.error)}
  if($partial.PSObject.Properties['hresult'] -and [int]$partial.hresult -lt 0){$errors.Add(('CopyMessages HRESULT 0x'+([int]$partial.hresult).ToString('X8')))}
 }
 return [pscustomobject]@{attempted=$true;copies=@($copies.ToArray());skipped=@($skipped.ToArray());error=($errors -join '; ')}
}
function Invoke-NativeCopyResults($Root,[string]$DestinationStoreId,[string]$DestinationFolderId,$Messages,$Eligible,$Manifest,$Context) {
 $writer=$null;$eligibleIds=@{};$copies=@{};$duplicates=@{};$skipped=@{}
 foreach($m in $Eligible){$eligibleIds[[string]$m.id]=$true}
 if($Manifest -and $Manifest.PSObject.Properties['skipped']){foreach($entry in @($Manifest.skipped)){if($entry.id -and $entry.error){$skipped[[string]$entry.id]=[string]$entry.error}}}
 foreach($copy in @($Manifest.copies)){
  $id=[string]$copy.id
  if(-not $eligibleIds.ContainsKey($id) -or $duplicates.ContainsKey($id)){continue}
  if($copies.ContainsKey($id)){$copies.Remove($id);$duplicates[$id]=$true;continue}
  $copies[$id]=$copy
 }
 try{
  $writer=New-Object IO.StreamWriter((Join-Path $Context.Directory 'report.csv'),[bool]$Context.AppendReport,(New-Object Text.UTF8Encoding($true)))
  if(-not $Context.AppendReport){Write-CsvRecord $writer @('sourceStoreId','sourceEntryId','destinationStoreId','destinationEntryId','status','error','temporaryCopyEntryId','sourceCurrentEntryId')}
  $rootId=[string]$Root.EntryID
  foreach($m in $Messages){
   $identity=Get-Identity $m;$id=[string]$m.id;$destinationId='';$errorText='';$proof=$null;$status='failed'
   if($copies.ContainsKey($id)){
    $copy=$copies[$id]
    try{
     if([string]::IsNullOrWhiteSpace([string]$copy.destinationId) -or [string]::IsNullOrWhiteSpace([string]$copy.sourceSearchKey) -or [string]::IsNullOrWhiteSpace([string]$copy.sourceModifiedFileTime)){throw 'Biên nhận MAPI thiếu ID, Search Key hoặc thời điểm nguồn.'}
     $segments=@(Get-PstFolderSegments $m);$sourceTicks=[datetime]::FromFileTimeUtc([long]$copy.sourceModifiedFileTime).ToLocalTime().Ticks
     $subject=[string]$copy.subject;$internet=[string]$copy.internetMessageId;$size=[long]$copy.sourceSize;$attachments=[int]$copy.sourceAttachments
     if($size -le 0 -or $attachments -lt 0){throw 'Snapshot nguồn MAPI không hợp lệ.'}
     $destinationId=[string]$copy.destinationId
     $sourceSpec=@{id=$id;storeId=[string]$m.storeId;sourceFolderId=[string]$m.sourceFolderId;subject=$subject;internetMessageId=$internet}
     $proof=@{sourceMessage=$sourceSpec;sourceSearchKey=[string]$copy.sourceSearchKey;sourceModifiedTicks=$sourceTicks;sourceModifiedFileTime=[string]$copy.sourceModifiedFileTime;destinationStoreId=$DestinationStoreId;destinationId=$destinationId;destinationFolderId=$DestinationFolderId;destinationRootId=$rootId;destinationSegments=$segments;subject=$subject;size=$size;attachmentCount=$attachments;internetMessageId=$internet;pstMode=$(if($Context.ContainsKey('ExportMode')){$Context.ExportMode}else{'existing'})}
     $status='verified';$Context.Copied++
     if($Context.AutoCleanup){$Context.VerifiedReceipts[$identity]=[pscustomobject]$proof}
    }catch{$errorText=$_.Exception.Message;$Context.Errors++}
   }else{
    if(-not $eligibleIds.ContainsKey($id)){$errorText='Thư chưa tải đầy đủ vào cache; giữ thư nguồn.'}
    elseif($skipped.ContainsKey($id)){$errorText='Không thể đọc thư từ cache MAPI: '+$skipped[$id]+'; giữ thư nguồn.'}
    elseif($Manifest -and $Manifest.PSObject.Properties['error'] -and $Manifest.error){$errorText=[string]$Manifest.error}else{$errorText='MAPI chưa xác minh được bản sao trong PST; giữ thư nguồn.'}
    $Context.Errors++
   }
   $Context.Processed++
   $Context.Results.Add(@{key=$identity;status=$status;destinationId=$destinationId;error=$errorText;archiveProof=$proof})
   Write-CsvRecord $writer @($m.storeId,$m.id,$DestinationStoreId,$destinationId,$status,$errorText,'','')
   if($Context.Results.Count -ge 100){Flush-ResultBatch $Context $writer}
  }
  Flush-ResultBatch $Context $writer;$Context.AppendReport=$true
  $Context.Note='Đã sao chép và đối chiếu MAPI '+$Context.Copied+'/'+$Context.Total+' email trong các lô. Lỗi: '+$Context.Errors+'. '+$(if($Context.AutoCleanup){'Chuẩn bị xử lý tự xóa các bản đã kiểm tra.'}else{'Thư gốc được giữ nguyên.'})
  Save-JobState $Context -Force
 }finally{if($writer){$writer.Dispose()}}
}
function Invoke-BulkArchiveExport($Namespace,$Root,[string]$DestinationStoreId,$Request,$Context,[scriptblock]$CopyProvider=$null) {
 $Context.AutoCleanup=Get-AutoCleanupEnabled $Request
 $selected=@(Select-UniqueMessages @($Request.messages));$Context.Total=$selected.Count
 $Context.DeferCompletion=$true;$Context.NativeCopyMode=$true
 $cache=@{};$groups=New-Object Collections.Specialized.OrderedDictionary;$handle=[IntPtr]::Zero
 try{
  foreach($m in $selected){
   if(-not $m.PSObject.Properties['sourceFolderId'] -or [string]::IsNullOrWhiteSpace([string]$m.sourceFolderId)){throw 'Hãy quét lại trước khi chuyển theo lô.'}
   $segments=@(Get-PstFolderSegments $m)
   $key=([string]$m.storeId)+'|'+([string]$m.sourceFolderId)+'|'+($segments -join [char]0)
   if(-not $groups.Contains($key)){$groups.Add($key,(New-Object 'Collections.Generic.List[object]'))}
   $groups[$key].Add($m)
  }
  if(-not $CopyProvider){
   $dll=Join-Path $PSScriptRoot 'ktg-mapi-x64.dll';if(-not [IO.File]::Exists($dll)){$dll=Join-Path $PSScriptRoot '..\native\bin\ktg-mapi-x64.dll'};Initialize-BulkMapi $dll
   $hr=[KtgNativeBulk.Api]::KtgOpen([string]$Namespace.CurrentProfileName,[ref]$handle)
   if($hr -lt 0){throw ('Không mở được phiên MAPI riêng (0x'+$hr.ToString('X8')+'). Chưa sao chép hoặc xóa nguồn. Có thể tắt chế độ theo lô.')}
  }
  foreach($group in $groups.Values){
   for($offset=0;$offset -lt $group.Count;$offset+=1000){
    if(-not (Wait-BulkGate $Context)){break}
    $end=[math]::Min($offset+1000,$group.Count);$batch=@(for($i=$offset;$i -lt $end;$i++){$group[$i]})
    $Context.Phase='copy';$Context.Note='Đang sao chép lô MAPI '+($Context.BulkBatches+1)+' ('+$batch.Count+' thư).';Save-JobState $Context -Force
    $destination=Get-DestinationFolder $Root @(Get-PstFolderSegments $batch[0]) $cache $Context
    $eligibleList=New-Object 'Collections.Generic.List[object]'
    if($CopyProvider){
     for($checkOffset=0;$checkOffset -lt $batch.Count;$checkOffset+=100){
      if(-not (Wait-BulkGate $Context)){break}
      $checkEnd=[math]::Min($checkOffset+100,$batch.Count);$Context.Note='Kiểm tra cache lô '+($Context.BulkBatches+1)+': '+$checkEnd+'/'+$batch.Count+' thư.';Save-JobState $Context -Force
      for($i=$checkOffset;$i -lt $checkEnd;$i++){$m=$batch[$i];$item=$null;try{$item=$Namespace.GetItemFromID([string]$m.id,[string]$m.storeId);if([int]$item.Class -eq 43 -and [int]$item.DownloadState -eq 1 -and (Test-SameOutlookId $Namespace ([string]$item.EntryID) ([string]$m.id))){$eligibleList.Add($m)}}catch{}finally{Release-Com $item}}
     }
    }else{foreach($m in $batch){$eligibleList.Add($m)}}
    $eligible=@($eligibleList.ToArray())
    if($Context.Status -eq 'cancelled'){break}
    $Context.NativeCopies=@{};$nativeError='';$manifest=$null
    $messageById=@{};foreach($message in $batch){$messageById[[string]$message.id]=$message}
    try{
     if($eligible.Count -eq 0){$manifest=[pscustomobject]@{copies=@();attempted=$false;error='Không có thư đầy đủ cache trong lô.'}}
     elseif($CopyProvider){$manifest=& $CopyProvider $eligible $destination}
     else{
      $partials=New-Object 'Collections.Generic.List[object]';$failures=New-Object 'Collections.Generic.List[string]'
      for($nativeOffset=0;$nativeOffset -lt $eligible.Count;$nativeOffset+=1000){
       if(-not (Wait-BulkGate $Context)){break}
       $nativeEnd=[math]::Min($nativeOffset+1000,$eligible.Count);$nativeBatch=@(for($i=$nativeOffset;$i -lt $nativeEnd;$i++){$eligible[$i]})
       $Context.Note='Đang gửi lô MAPI '+($Context.BulkBatches+1)+' ('+$nativeBatch.Count+' thư); chờ sao chép và xác minh PST hoàn tất.';Save-JobState $Context -Force
       $json='';$ids=@($nativeBatch|ForEach-Object{[string]$_.id}) -join "`n"
       $hr=[KtgNativeBulk.Api]::KtgCopy($handle,[string]$nativeBatch[0].storeId,[string]$nativeBatch[0].sourceFolderId,$DestinationStoreId,[string]$destination.EntryId,$ids,[ref]$json)
       $Context.Note='MAPI đã trả kết quả lô '+($Context.BulkBatches+1)+'; đang xử lý biên nhận xác minh.';Save-JobState $Context -Force
       if($hr -lt 0){$detail='';try{$failureManifest=$json|ConvertFrom-Json;if($failureManifest.error){$detail=': '+[string]$failureManifest.error}}catch{};$failures.Add(('MAPI lô con lỗi 0x'+$hr.ToString('X8')+$detail+'; giữ nguồn của lô con.'));continue}
       try{$partial=$json|ConvertFrom-Json;$partials.Add($partial)}catch{$failures.Add('Không đọc được biên nhận MAPI lô con; giữ nguồn.')}
      }
      $manifest=Merge-NativeCopyManifests $partials.ToArray() $failures.ToArray()
     }
     Write-AtomicJson (Join-Path $Context.Directory ('native-batch-'+$Context.BulkBatches+'.json')) $manifest
     # Only mappings returned by our copy-only native worker are accepted.
     foreach($entry in @($manifest.copies)){
      $entryId=[string]$entry.id
      if(-not $messageById.ContainsKey($entryId)){continue}
      $match=$messageById[$entryId]
      if(-not $entry.PSObject.Properties['error']){$entry|Add-Member error ''}
      $Context.NativeCopies[(Get-Identity $match)]=$entry
     }
    }catch{$nativeError=$_.Exception.Message}
    $Context.BulkBatches++
    if($nativeError){$Context.Note=$nativeError;Save-JobState $Context -Force}
    # Every mapped copy still passes all existing source/cache/PST checks.
    if($CopyProvider){Invoke-CopyExport $Namespace $Root $DestinationStoreId $batch $Context}else{
     Invoke-NativeCopyResults $Root $DestinationStoreId ([string]$destination.EntryId) $batch $eligible $manifest $Context
    }
    if(-not (Wait-BulkGate $Context)){break}
    if($Context.AutoCleanup){
     $pending=@($batch|Where-Object{$Context.VerifiedReceipts.ContainsKey((Get-Identity $_))})
     if($pending.Count){
      $Context.Phase='cleanup';Save-JobState $Context -Force
      $cleanup=[pscustomobject]@{operation='delete_archived';messages=$pending;confirmPermanentDelete=$true;confirmationText=('XOA '+$pending.Count)}
      if($CopyProvider){Invoke-ConfirmedDelete $Namespace $cleanup $Context $Context.VerifiedReceipts -Automatic}else{Invoke-NativeCleanup $Namespace $handle $pending $Context $Context.VerifiedReceipts}
     }
    }
    $Context.VerifiedReceipts.Clear();$Context.NativeCopies.Clear()
    if($Context.Status -eq 'cancelled'){break}
    [Threading.Thread]::Sleep(150)
   }
   if($Context.Status -eq 'cancelled'){break}
  }
  if($Context.Status -eq 'running'){$Context.Status=if($Context.Errors){'completed_with_errors'}else{'completed'}}
  $Context.Note='Lô MAPI: '+$Context.BulkBatches+'. PST đã kiểm tra: '+$Context.Copied+'/'+$Context.Total+'. Đã xử lý xóa nguồn: '+$Context.Deleted+'. Lỗi: '+$Context.Errors+'. Đối chiếu CSV; không tự thử lại các mục lỗi.'
 }finally{
  if($handle -ne [IntPtr]::Zero){[KtgNativeBulk.Api]::KtgClose($handle)}
  foreach($entry in $cache.Values){Release-Com $entry.Folder}
  $Context.DeferCompletion=$false;$Context.NativeCopyMode=$false;Save-JobState $Context -Force
 }
}

function Write-DurableCleanupJson([string]$Path,$Value) {
 $bytes=[Text.Encoding]::UTF8.GetBytes((ConvertTo-Json -InputObject $Value -Depth 16 -Compress))
 $file=New-Object IO.FileStream(($Path+'.tmp'),[IO.FileMode]::Create,[IO.FileAccess]::Write,[IO.FileShare]::None)
 try{$file.Write($bytes,0,$bytes.Length);$file.Flush($true)}finally{$file.Dispose()}
 if([IO.File]::Exists($Path)){[IO.File]::Replace(($Path+'.tmp'),$Path,[NullString]::Value)}else{[IO.File]::Move(($Path+'.tmp'),$Path)}
}
function Convert-CleanupHex([string]$Text){return [BitConverter]::ToString([Text.Encoding]::UTF8.GetBytes($Text)).Replace('-','')}
function Invoke-NativeCleanupBatch($Namespace,[IntPtr]$Handle,$Messages,$Context,$Receipts) {
 if(-not (Wait-BulkGate $Context)){return}
 $rows=New-Object 'Collections.Generic.List[string]';$byId=@{};$first=$null
 foreach($m in $Messages){
  $key=Get-Identity $m
  if(-not $Receipts.ContainsKey($key)){throw 'Thiếu biên nhận PST; không xóa nguồn.'}
  $proof=$Receipts[$key]
  if(-not $proof.PSObject.Properties['sourceModifiedFileTime'] -or -not $proof.PSObject.Properties['sourceSearchKey']){throw 'Biên nhận cũ thiếu MAPI timestamp/Search Key; giữ thư nguồn.'}
  if(-not $first){$first=$proof}else{
   if([string]$proof.sourceMessage.storeId -ine [string]$first.sourceMessage.storeId -or [string]$proof.sourceMessage.sourceFolderId -ine [string]$first.sourceMessage.sourceFolderId -or [string]$proof.destinationFolderId -ine [string]$first.destinationFolderId){throw 'Không trộn thư mục trong cùng lô xóa.'}
  }
  $id=[string]$proof.sourceMessage.id
  if($byId.ContainsKey($id)){throw 'Trùng ID nguồn trong lô xóa.'};$byId[$id]=@{message=$m;proof=$proof}
  $rows.Add((@($id,[string]$proof.destinationId,[string]$proof.sourceModifiedFileTime,[string]$proof.size,[string]$proof.attachmentCount,(Convert-CleanupHex ([string]$proof.subject)),(Convert-CleanupHex ([string]$proof.internetMessageId)),[string]$proof.sourceSearchKey) -join "`t"))
 }
 if(-not $first){return}
 $store=$null
 try{$store=$Namespace.GetStoreFromID([string]$first.sourceMessage.storeId);if(-not (Test-CachedExchangeStore $store)){throw 'Chỉ xóa nguồn Exchange đã cache.'}}finally{Release-Com $store}
 $paths=@($first.destinationSegments|ForEach-Object{Convert-CleanupHex ([string]$_)}) -join "`n"
 if(-not $Context.ContainsKey('CleanupBatchSequence')){$Context.CleanupBatchSequence=0};$Context.CleanupBatchSequence++
 $token='';$json='';$manifest=$null;$writer=$null;$base=Join-Path $Context.Directory ('cleanup-batch-'+$Context.BulkBatches+'-'+$Context.CleanupBatchSequence)
 try{
  $hr=[KtgNativeBulk.Api]::KtgPrepareCleanup($Handle,[string]$first.sourceMessage.storeId,[string]$first.sourceMessage.sourceFolderId,[string]$first.destinationStoreId,[string]$first.destinationFolderId,[string]$first.destinationRootId,$paths,($rows.ToArray() -join "`n"),[ref]$json)
  if($hr -lt 0){throw ('Không chuẩn bị được lô xóa MAPI: 0x'+$hr.ToString('X8')+'. Chưa gọi chuyển/xóa nguồn.')}
  $manifest=$json|ConvertFrom-Json;$token=[string]$manifest.planToken
  # Fail closed if the durable intent cannot be saved before any mutation.
  Write-DurableCleanupJson ($base+'-intent.json') @{plan=$manifest;receipts=$Receipts}
  if(-not (Wait-BulkGate $Context)){return}
  $json='';$hr=[KtgNativeBulk.Api]::KtgCleanupStep($Handle,$token,1,[ref]$json)
  if($hr -lt 0){throw ('Lệnh hard-delete nguồn trả lỗi 0x'+$hr.ToString('X8')+'. Không tự thử lại; đối chiếu intent.')}
  $manifest=$json|ConvertFrom-Json
  # The durable intent already names the exact selected source IDs and carries
  # the verified PST receipts. Step 1 hard-deletes those source IDs in one MAPI
  # DeleteMessages call; never retry if the outcome is partial/unknown.
  Write-DurableCleanupJson ($base+'-result.json') $manifest
  $writer=New-Object IO.StreamWriter((Join-Path $Context.Directory 'report.csv'),$true,(New-Object Text.UTF8Encoding($true)))
  $Context.CleanupTotal=$Context.Total
  foreach($row in @($manifest.rows)){
   if(-not $byId.ContainsKey([string]$row.id)){throw 'Native trả về ID ngoài lô được kiểm tra.'}
   $m=$byId[[string]$row.id].message;$key=Get-Identity $m
   if($row.status -eq 'delete_requested'){$Context.Deleted++}else{$Context.Errors++}
   $Context.CleanupProcessed++
   Write-CsvRecord $writer @($m.storeId,$m.id,'','',$row.status,$row.error,'',$row.currentId)
   $Context.Results.Add(@{key=$key;status=[string]$row.status;error=[string]$row.error;currentId=[string]$row.currentId;needsRescan=($row.status -ne 'delete_requested')})
  }
  if(@($manifest.rows).Count -ne $Messages.Count){throw 'Thiếu kết quả từng thư sau lô xóa; không tự thử lại.'}
  Flush-ResultBatch $Context $writer;Save-JobState $Context -Force
 }finally{
  if($writer){$writer.Dispose()}
  if($token){$ignored='';[void][KtgNativeBulk.Api]::KtgCleanupStep($Handle,$token,3,[ref]$ignored)}
 }
}
function Invoke-NativeCleanup($Namespace,[IntPtr]$Handle,$Messages,$Context,$Receipts) {
 for($offset=0;$offset -lt $Messages.Count;$offset+=100){
  if(-not (Wait-BulkGate $Context)){return}
  $end=[math]::Min($offset+100,$Messages.Count);$slice=@(for($i=$offset;$i -lt $end;$i++){$Messages[$i]});$sliceReceipts=@{}
  foreach($m in $slice){$key=Get-Identity $m;if(-not $Receipts.ContainsKey($key)){throw 'Thiếu biên nhận PST; không xóa nguồn.'};$sliceReceipts[$key]=$Receipts[$key]}
  $Context.Note='Đang dọn nguồn cho lô '+$Context.BulkBatches+': '+$end+'/'+$Messages.Count+' thư.';Save-JobState $Context -Force
  Invoke-NativeCleanupBatch $Namespace $Handle $slice $Context $sliceReceipts
 }
}
