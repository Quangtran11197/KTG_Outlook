param([string]$Core,[string]$Fixture,[string]$TestDirectory,[int]$Count=2000)
$ErrorActionPreference='Stop'
if(-not $Core){$Core=Join-Path $PSScriptRoot '..\scripts\core.ps1'}
if(-not $Fixture){$Fixture=Join-Path $PSScriptRoot 'test-delete.ps1'}
if(-not $TestDirectory){$TestDirectory=Join-Path ([IO.Path]::GetTempPath()) ('KTG-export-bench-'+[guid]::NewGuid().ToString('N'))}
& $Fixture -TestDirectory (Join-Path $TestDirectory 'guards')|Out-Null
. $Core
$script:count=$Count
$ns=New-Object KtgDeleteTest.Namespace;$messages=New-Object 'Collections.Generic.List[object]'
for($i=0;$i -lt $Count;$i++){
 $id='id'+$i;$ns.Items['SOURCE|'+$id]=New-Object KtgDeleteTest.Mail -Property @{Owner=$ns;EntryID=$id}
 $messages.Add([pscustomobject]@{id=$id;storeId='SOURCE';sourceFolderId='inbox';subject='Mail';internetMessageId='<id>';pstFolderSegments=@('Inbox',('Folder'+($i%10)))})
}
[KtgDeleteTest.Stats]::Reset()
$ctx=New-JobContext (Join-Path $TestDirectory 'export');$ctx.ExportMode='existing'
$request=[pscustomobject]@{operation='export';pstMode='existing';messages=$messages.ToArray();autoDeleteAfterExport=$true;autoDeleteConfirmed=$true;autoDeleteCount=$Count}
Invoke-ArchiveExport $ns $ns.DestinationStore.Root 'PST' $request $ctx
$elapsed=$ctx.Timer.Elapsed.TotalSeconds;$pathReads=[KtgDeleteTest.Stats]::PathReads
if($ctx.Copied -ne $Count -or $ctx.Deleted -ne $Count -or $ctx.Errors -ne 0){throw ('Incorrect counts: '+$ctx.Copied+'/'+$ctx.Deleted+'/'+$ctx.Errors+' '+$ctx.Note)}
$rows=@(Import-Csv (Join-Path $ctx.Directory 'report.csv'))
if(@($rows|Where-Object status -eq 'verified').Count -ne $Count -or @($rows|Where-Object status -eq 'delete_requested').Count -ne $Count -or $rows.Count -ne $Count*4){throw 'Incomplete audit'}
foreach($m in $messages){if($ns.Items.ContainsKey('SOURCE|'+$m.id) -or $ns.Items.ContainsKey('SOURCE|moved-'+$m.id)){throw 'Source remains'}}
$proofs=Read-VerifiedPstReceipts $ctx.Directory
foreach($proof in $proofs.Values){Assert-ArchivedCopy $ns $proof}
$result=@{passed=$true;synthetic=$true;count=$Count;copied=$ctx.Copied;deleted=$ctx.Deleted;seconds=[math]::Round($elapsed,3);reportRows=$rows.Count;folders=10;reopenedCopies=$proofs.Count;sourceLookups=[KtgDeleteTest.Stats]::Lookups;topologyReads=$pathReads;privateDirectory=$TestDirectory}
Write-AtomicJson (Join-Path $TestDirectory 'result.json') $result;$result|ConvertTo-Json
