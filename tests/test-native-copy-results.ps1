param()
$ErrorActionPreference='Stop'
. "$PSScriptRoot\..\scripts\core.ps1"
. "$PSScriptRoot\..\scripts\bulk.ps1"
$dir=Join-Path ([IO.Path]::GetTempPath()) ('KTG-native-copy-results-'+[guid]::NewGuid().ToString('N'))
function Check($condition,[string]$message){if(-not $condition){throw $message}}
try{
 $context=New-JobContext $dir 2;$context.AutoCleanup=$true;$context.ExportMode='existing'
 $filetime=[datetime]::UtcNow.ToFileTimeUtc().ToString()
 $messages=@(
  [pscustomobject]@{id='AA01';storeId='SRC';sourceFolderId='SF';pstFolderSegments=@('Inbox','SPL');subject='old';internetMessageId='old'}
  [pscustomobject]@{id='AA02';storeId='SRC';sourceFolderId='SF';pstFolderSegments=@('Inbox','SPL');subject='skip';internetMessageId=''}
 )
 $eligible=@($messages[0]);$manifest=[pscustomobject]@{error='one cached item skipped';copies=@([pscustomobject]@{id='AA01';destinationId='BB01';sourceSearchKey='AABBCC';sourceModifiedFileTime=$filetime;sourceSize=321;sourceAttachments=2;subject='verified subject';internetMessageId='<verified@example.test>'},[pscustomobject]@{id='UNRELATED';destinationId='BAD';sourceSearchKey='DDEEFF';sourceModifiedFileTime=$filetime;sourceSize=1;sourceAttachments=0;subject='extra';internetMessageId=''})}
 $root=[pscustomobject]@{EntryID='ROOT'}
 Invoke-NativeCopyResults $root 'DST' 'DF' $messages $eligible $manifest $context
 Check ($context.Copied -eq 1 -and $context.Errors -eq 1 -and $context.Processed -eq 2) 'Native result counters mismatch'
 Check ($context.VerifiedReceipts.Count -eq 1 -and $context.VerifiedReceipts.ContainsKey('SRC|AA01')) 'Verified receipt missing or extra receipt accepted'
 $proof=$context.VerifiedReceipts['SRC|AA01'];Check ($proof.destinationId -eq 'BB01' -and $proof.destinationSegments.Count -eq 2 -and $proof.sourceModifiedFileTime -eq $filetime -and $proof.sourceSearchKey -eq 'AABBCC') 'Native proof fields mismatch'
 Check ($context.ResultBatches -eq 1 -and $context.Results.Count -eq 0) 'Result batch was not durably flushed'

 $context2=New-JobContext (Join-Path $dir 'duplicate') 1;$context2.AutoCleanup=$true
 $one=@($messages[0]);$duplicates=[pscustomobject]@{copies=@([pscustomobject]@{id='AA01';destinationId='BB01';sourceSearchKey='AABBCC';sourceModifiedFileTime=$filetime;sourceSize=321;sourceAttachments=2;subject='s';internetMessageId=''},[pscustomobject]@{id='AA01';destinationId='BB02';sourceSearchKey='AABBCC';sourceModifiedFileTime=$filetime;sourceSize=321;sourceAttachments=2;subject='s';internetMessageId=''})}
 Invoke-NativeCopyResults $root 'DST' 'DF' $one $one $duplicates $context2
 Check ($context2.Copied -eq 0 -and $context2.Errors -eq 1 -and $context2.VerifiedReceipts.Count -eq 0) 'Duplicate native IDs were accepted'
 $script:cleanupSizes=New-Object 'Collections.Generic.List[int]'
 function Invoke-NativeCleanupBatch($Namespace,[IntPtr]$Handle,$Messages,$Context,$Receipts){$script:cleanupSizes.Add($Messages.Count);Check ($Receipts.Count -eq $Messages.Count) 'Cleanup slice has an incomplete receipt set'}
 $cleanupContext=New-JobContext (Join-Path $dir 'cleanup-slices') 123;$cleanupContext.BulkBatches=1;$cleanupMessages=New-Object 'Collections.Generic.List[object]';$cleanupReceipts=@{}
 for($i=0;$i -lt 123;$i++){$m=[pscustomobject]@{id=('S'+$i);storeId='SRC'};$cleanupMessages.Add($m);$cleanupReceipts[(Get-Identity $m)]=[pscustomobject]@{sourceModifiedFileTime='1'}}
 Invoke-NativeCleanup $null ([IntPtr]::Zero) $cleanupMessages.ToArray() $cleanupContext $cleanupReceipts
 Check (($script:cleanupSizes -join ',') -eq '100,23') '1,000-message export cleanup is split into at-most-100 receipt-verified chunks'
 $skipContext=New-JobContext (Join-Path $dir 'cache-skip') 1;$skipMessage=@($messages[0]);$skipManifest=[pscustomobject]@{attempted=$true;copies=@();skipped=@([pscustomobject]@{id='AA01';error='cached item unavailable'})}
 Invoke-NativeCopyResults $root 'DST' 'DF' $skipMessage $skipMessage $skipManifest $skipContext
 $skipReport=[IO.File]::ReadAllText((Join-Path $skipContext.Directory 'report.csv'));Check ($skipContext.Copied -eq 0 -and $skipContext.Errors -eq 1 -and $skipReport.Contains('cached item unavailable')) 'Cache-only MAPI rejection is reported per email without accepting a copy'
 $merged=Merge-NativeCopyManifests @([pscustomobject]@{copies=@([pscustomobject]@{id='OK'});skipped=@([pscustomobject]@{id='S1';error='stage miss'});error='';hresult=0},[pscustomobject]@{copies=@();skipped=@([pscustomobject]@{id='S2';error='verify miss'});error='finalize warning';hresult=-2147024809}) @('subbatch failed')
 Check (@($merged.copies).Count -eq 1 -and @($merged.skipped).Count -eq 2 -and $merged.skipped[1].error -eq 'verify miss' -and $merged.error.Contains('finalize warning') -and $merged.error.Contains('80070057')) 'Per-email native reasons and subbatch HRESULTs survive manifest merge'
 @{passed=$true;scenarios=5;cleanupChunkSizes=$script:cleanupSizes.ToArray();syntheticOnly=$true;directory=$dir}|ConvertTo-Json -Compress
}finally{}
