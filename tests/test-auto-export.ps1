param([string]$TestDirectory)
$ErrorActionPreference='Stop'
if(-not $TestDirectory){$TestDirectory=Join-Path ([IO.Path]::GetTempPath()) ('KTG-auto-'+[guid]::NewGuid().ToString('N'))}
# Load and run the existing delete/receipt guard tests in this process first.
& "$PSScriptRoot\test-delete.ps1" -TestDirectory (Join-Path $TestDirectory 'delete-guards') | Out-Null
. "$PSScriptRoot\..\scripts\core.ps1"
$script:checks=0
function Assert($value,[string]$message){$script:checks++;if(-not $value){throw ('FAIL: '+$message)}}
function MailMetadata([string]$id){
 [pscustomobject]@{id=$id;storeId='SOURCE';sourceFolderId='inbox';subject='Mail';internetMessageId='<id>';pstFolderSegments=@('Inbox','SPL')}
}
function AutoRequest([string]$mode,$messages,[bool]$enabled=$true){
 [pscustomobject]@{operation='export';pstMode=$mode;messages=@($messages);autoDeleteAfterExport=$enabled;autoDeleteConfirmed=$enabled;autoDeleteCount=@($messages).Count}
}
foreach($scenario in @('existing','new','copy-only','legacy-copy-only','copy-failed','missing-pst','source-changed','cancel-before-cleanup','cancel-during-cleanup','delete-failed','mixed')){
 [KtgDeleteTest.Stats]::Reset()
 $ns=New-Object KtgDeleteTest.Namespace
 $context=New-JobContext (Join-Path $TestDirectory $scenario)
 $source=New-Object KtgDeleteTest.Mail -Property @{Owner=$ns}
 $ns.Items['SOURCE|id1']=$source
 $messages=@((MailMetadata 'id1'))
 $mode=if($scenario -eq 'new'){'new'}else{'existing'};$context.ExportMode=$mode
 switch($scenario){
  'copy-failed'{$source.FailCopy=$true}
  'missing-pst'{$ns.MissingPst=$true}
  'source-changed'{$source.ChangeSourceAfterCopy=$true}
  'cancel-before-cleanup'{$source.CancelPath=Join-Path $context.Directory 'cancel'}
  'cancel-during-cleanup'{$source.CancelAfterSourceMove=Join-Path $context.Directory 'cancel'}
  'delete-failed'{$source.FailDelete=$true}
  'mixed'{$ns.Items['SOURCE|id2']=New-Object KtgDeleteTest.Mail -Property @{Owner=$ns;EntryID='id2';FailCopy=$true};$messages+=MailMetadata 'id2'}
 }
 $request=AutoRequest $mode $messages ($scenario -ne 'copy-only')
 if($scenario -eq 'legacy-copy-only'){$request=[pscustomobject]@{operation='export';pstMode=$mode;messages=$messages}}
 Invoke-ArchiveExport $ns $ns.DestinationStore.Root 'PST' $request $context
 if($scenario -in @('existing','new')){
  Assert ($context.Copied -eq 1 -and $context.Deleted -eq 1 -and $context.Errors -eq 0) ($scenario+' copies then auto-deletes without a second request')
  Assert ([KtgDeleteTest.Stats]::Copies -eq 1 -and [KtgDeleteTest.Stats]::Moves -eq 1 -and [KtgDeleteTest.Stats]::Deletes -eq 1) 'Exactly one original moved/deleted'
  $report=@(Import-Csv (Join-Path $context.Directory 'report.csv'))
  Assert ($report[0].status -eq 'verified' -and $report[-1].status -eq 'delete_requested') 'Copy and deletion audit retained in one CSV'
  Assert ($context.Processed -eq 1 -and $context.CleanupProcessed -eq 1 -and $context.CleanupTotal -eq 1) 'Separate copy/cleanup progress'
 }elseif($scenario -in @('copy-only','legacy-copy-only')){
  Assert ($context.Copied -eq 1 -and $context.Deleted -eq 0 -and [KtgDeleteTest.Stats]::Deletes -eq 0) 'Opt-out preserves originals'
 }elseif($scenario -eq 'mixed'){
  Assert ($context.Copied -eq 1 -and $context.Deleted -eq 1 -and $context.Errors -eq 1) 'Only verified subset auto-deleted'
  Assert ([KtgDeleteTest.Stats]::Deletes -eq 1 -and $ns.Items['SOURCE|id2'].Parent.EntryID -eq 'inbox') 'Failed copy stays in its source folder'
 }elseif($scenario -eq 'cancel-during-cleanup'){
  Assert ($context.Status -eq 'cancelled' -and $context.Copied -eq 1 -and [KtgDeleteTest.Stats]::Moves -eq 1 -and [KtgDeleteTest.Stats]::Deletes -eq 0) 'Stop after source move prevents permanent deletion'
 }elseif($scenario -eq 'delete-failed'){
  Assert ($context.Copied -eq 1 -and $context.Deleted -eq 0 -and $context.Errors -eq 1 -and [KtgDeleteTest.Stats]::Deletes -eq 1) 'Delete failure is reported and never automatically retried'
 }else{
  Assert ($context.Deleted -eq 0 -and [KtgDeleteTest.Stats]::Moves -eq 0 -and [KtgDeleteTest.Stats]::Deletes -eq 0) ($scenario+' never moves/deletes original')
  if($scenario -eq 'cancel-before-cleanup'){Assert ($context.Status -eq 'cancelled') 'Cancellation stops auto-cleanup'}
 }
}
# A valid message in an orphan folder must never authorize source deletion.
$dir=Join-Path $TestDirectory 'existing'
$receipts=Read-VerifiedPstReceipts $dir
$proof=$receipts['SOURCE|id1']
$ns=New-Object KtgDeleteTest.Namespace
$ns.Items['PST|'+$proof.destinationId]=New-Object KtgDeleteTest.Mail -Property @{EntryID=$proof.destinationId;Parent=(New-Object KtgDeleteTest.Folder -Property @{StoreID='PST';EntryID=$proof.destinationFolderId})}
$ns.DestinationStore.Root.EntryID=$proof.destinationRootId
[KtgDeleteTest.Stats]::Reset();$threw=$false
try{Assert-ArchivedCopy $ns $proof}catch{$threw=$true}
Assert ($threw -and [KtgDeleteTest.Stats]::Deletes -eq 0) 'Orphan PST folder rejected even when message can be opened by EntryID'
# Cached topology is scoped to this cleanup, with bounded reuse and expiry.
$ns=New-Object KtgDeleteTest.Namespace
$root=$ns.DestinationStore.Root;$inbox=$root.Folders.Add('Inbox',6);$spl=$inbox.Folders.Add('SPL',6)
$proof=[pscustomobject]@{pstMode='existing';destinationStoreId='PST';destinationId='copy';destinationFolderId=$spl.EntryID;destinationRootId=$root.EntryID;destinationSegments=@('Inbox','SPL');subject='Mail';size=4096;attachmentCount=2;internetMessageId='<id>'}
$ns.Items['PST|copy']=New-Object KtgDeleteTest.Mail -Property @{EntryID='copy';Parent=$spl}
$pathCache=@{};[KtgDeleteTest.Stats]::Reset()
Assert-ArchivedCopy $ns $proof -PathCache $pathCache
$reads=[KtgDeleteTest.Stats]::PathReads
Assert-ArchivedCopy $ns $proof -PathCache $pathCache
Assert ([KtgDeleteTest.Stats]::PathReads -eq $reads -and [KtgDeleteTest.Stats]::Lookups -eq 2) 'Reuse topology but reopen/validate each PST message'
foreach($entry in $pathCache.Values){$entry.uses=100}
Assert-ArchivedCopy $ns $proof -PathCache $pathCache
Assert ([KtgDeleteTest.Stats]::PathReads -gt $reads) 'Topology rechecked after 100 uses'
$inbox.Folders.Remove('SPL');foreach($entry in $pathCache.Values){$entry.tick=[Environment]::TickCount-501}
$threw=$false;try{Assert-ArchivedCopy $ns $proof -PathCache $pathCache}catch{$threw=$true}
Assert $threw 'Expired topology rejects removed/moved visible folder'
$pathCache=@{};$threw=$false;try{Assert-ArchivedCopy $ns $proof -PathCache $pathCache}catch{$threw=$true}
Assert $threw 'Fresh cleanup never trusts another job cache'
foreach($kind in @('missing-confirmation','wrong-count','string-flag','missing-folder')){
 $request=AutoRequest 'existing' @((MailMetadata 'id1'))
 switch($kind){
  'missing-confirmation'{$request.autoDeleteConfirmed=$false}
  'wrong-count'{$request.autoDeleteCount=2}
  'string-flag'{$request.autoDeleteAfterExport='true'}
  'missing-folder'{$request.messages[0].sourceFolderId=''}
 }
 [KtgDeleteTest.Stats]::Reset();$ns=New-Object KtgDeleteTest.Namespace
 $threw=$false
 try{Invoke-ArchiveExport $ns (New-Object KtgDeleteTest.Folder) 'PST' $request (New-JobContext (Join-Path $TestDirectory $kind))}catch{$threw=$true}
 Assert ($threw -and [KtgDeleteTest.Stats]::Lookups -eq 0 -and [KtgDeleteTest.Stats]::Deletes -eq 0) ($kind+' rejected before mail access')
}
@{passed=$true;checks=$script:checks;test='automatic-export-synthetic-only';testDirectory=$TestDirectory}|ConvertTo-Json
