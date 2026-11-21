param([string]$TestDirectory)
$ErrorActionPreference='Stop'
if(-not $TestDirectory){$TestDirectory=Join-Path ([IO.Path]::GetTempPath()) ('KTG-bulk-'+[guid]::NewGuid().ToString('N'))}
& "$PSScriptRoot\test-delete.ps1" -TestDirectory (Join-Path $TestDirectory 'guards')|Out-Null
. "$PSScriptRoot\..\scripts\core.ps1"
. "$PSScriptRoot\..\scripts\bulk.ps1"
$script:checks=0
function Assert($ok,$message){$script:checks++;if(-not $ok){throw ('FAIL: '+$message)}}
foreach($scenario in @('auto','copy-only','partial','changed','header-only','bad-copy','empty-subject','cancel-after-copy')){
 $script:testNs=New-Object KtgDeleteTest.Namespace;$script:calls=0
 $ctx=New-JobContext (Join-Path $TestDirectory $scenario);$ctx.ExportMode='new'
 $messages=@(for($i=0;$i -lt 123;$i++){
  $id='id'+$i;$script:testNs.Items['SOURCE|'+$id]=New-Object KtgDeleteTest.Mail -Property @{Owner=$script:testNs;EntryID=$id}
  [pscustomobject]@{id=$id;storeId='SOURCE';sourceFolderId='inbox';subject='Mail';internetMessageId='<id>';pstFolderSegments=@('Inbox','SPL')}
 })
 if($scenario -eq 'header-only'){$script:testNs.Items['SOURCE|id1'].DownloadState=0}
 if($scenario -eq 'empty-subject'){$script:testNs.Items['SOURCE|id1'].Subject='';$messages[1].subject=''}
 $script:scenario=$scenario;$script:cancelPath=Join-Path $ctx.Directory 'cancel'
 $provider={param($batch,$destination)
  $script:calls++;$copies=@()
  foreach($m in $batch){
   if($script:scenario -eq 'partial' -and $m.id -eq 'id1'){continue}
   $source=$script:testNs.Items['SOURCE|'+$m.id]
   if($source.DownloadState -ne 1){throw 'Uncached source reached native provider'}
   $copy=$source.Copy().Move($destination.Folder)
   if($script:scenario -eq 'bad-copy' -and $m.id -eq 'id1'){$copy.Subject='Mismatch'}
   $copies+=,[pscustomobject]@{id=$m.id;destinationId=$copy.EntryID;sourceModifiedFileTime=$source.LastModificationTime.ToUniversalTime().ToFileTimeUtc().ToString();sourceSize=$source.Size;sourceAttachments=$source.Attachments.Count;error=''}
   if($script:scenario -eq 'changed' -and $m.id -eq 'id1'){$source.Size=9999}
  }
  if($script:scenario -eq 'cancel-after-copy'){[IO.File]::WriteAllText($script:cancelPath,'cancel')}
  [pscustomobject]@{copies=$copies;attempted=$true;error=''}
 }
 $auto=$scenario -ne 'copy-only'
 $request=[pscustomobject]@{operation='export';pstMode='new';messages=$messages;autoDeleteAfterExport=$auto;autoDeleteConfirmed=$auto;autoDeleteCount=123;bulkCopy=$true}
 Invoke-BulkArchiveExport $script:testNs $script:testNs.DestinationStore.Root 'PST' $request $ctx $provider
 if($scenario -eq 'cancel-after-copy'){
  Assert ($ctx.Status -eq 'cancelled' -and $ctx.Deleted -eq 0 -and $script:calls -eq 1) 'Stop after native copy never starts source cleanup'
  continue
 }
 $expected=if($scenario -in @('partial','changed','header-only','bad-copy')){122}else{123}
 Assert ($ctx.Copied -eq $expected -and $ctx.Processed -eq 123 -and $ctx.Total -eq 123) ($scenario+' preserves global totals across batches')
 Assert ($script:calls -eq 1 -and $ctx.BulkBatches -eq 1) '123 messages use one 1,000-message outer batch'
 $wantedDeletes=if($auto){$expected}else{0}
 Assert ($ctx.Deleted -eq $wantedDeletes) 'Only individually verified PST copies permit source deletion'
 $rows=@(Import-Csv (Join-Path $ctx.Directory 'report.csv'))
 Assert (@($rows|Where-Object status -eq 'verified').Count -eq $expected -and @($rows|Where-Object status -eq 'delete_requested').Count -eq $wantedDeletes) 'Copy audit is appended across batches and cleanup'
 $inbox=$script:testNs.DestinationStore.Root.Folders.Item('Inbox');$spl=$inbox.Folders.Item('SPL')
 Assert ($spl.StoreID -eq 'PST' -and $ctx.Status -in @('completed','completed_with_errors')) 'Visible Inbox/SPL and terminal status preserved'
 if($scenario -in @('partial','changed','header-only','bad-copy')){Assert ($script:testNs.Items.ContainsKey('SOURCE|id1')) 'Failed or changed original remains'}
}
$ctx=New-JobContext (Join-Path $TestDirectory 'pause-cancel')
[IO.File]::WriteAllText((Join-Path $ctx.Directory 'pause'),'pause');[IO.File]::WriteAllText((Join-Path $ctx.Directory 'cancel'),'cancel')
Assert (-not (Wait-BulkGate $ctx)) 'Cancel works while paused'
Assert ($ctx.Status -eq 'cancelled') 'Paused cancel is reported'
$ctx=New-JobContext (Join-Path $TestDirectory 'pause-resume')
$marker=Join-Path $ctx.Directory 'pause';[IO.File]::WriteAllText($marker,'pause')
$job=Start-Job -ArgumentList $marker -ScriptBlock {param($p) Start-Sleep -Milliseconds 200;[IO.File]::Delete($p)}
try{Assert (Wait-BulkGate $ctx) 'Resume returns to processing at safe boundary';Assert (-not $ctx.Paused -and $ctx.Status -eq 'running') 'Resume preserves running state'}finally{Wait-Job $job -Timeout 10|Out-Null;Remove-Job $job -Force}
@{passed=$true;checks=$script:checks;synthetic=$true;test='bulk-copy-and-guarded-cleanup';testDirectory=$TestDirectory}|ConvertTo-Json
