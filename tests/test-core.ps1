param([int]$BenchmarkCount=40000,[string]$TestDirectory)
$ErrorActionPreference='Stop'
. "$PSScriptRoot\..\scripts\core.ps1"
if(-not $TestDirectory){$TestDirectory=Join-Path ([IO.Path]::GetTempPath()) ('KTG-test-'+[guid]::NewGuid().ToString('N'))}
[IO.Directory]::CreateDirectory($TestDirectory)|Out-Null
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
namespace KtgTest {
 public static class Stats { public static int Copies, OriginalMoves, Deletes, Lookups; public static void Reset(){Copies=OriginalMoves=Deletes=Lookups=0;} }
 public class Attachments {public int Count=2;}
 public class Accessor {public string GetProperty(string key){return "<test@example>";}}
 public class Folder {
   public string EntryID=Guid.NewGuid().ToString();
   public string StoreID="PST";
   public Folders Folders {get;set;}
   public Folder(){Folders=new Folders();}
 }
 public class Folders {
   Dictionary<string,Folder> entries=new Dictionary<string,Folder>(StringComparer.OrdinalIgnoreCase);
   public Folder Item(string name){return entries[name];}
   public Folder Add(string name,int kind){var f=new Folder();entries.Add(name,f);return f;}
 }
 public class Mail {
   public int Class=43; public int DownloadState=1;
   public string EntryID; public string Subject="Test";
   public long Size=4096; public Attachments Attachments=new Attachments();
   public Accessor PropertyAccessor=new Accessor();
   public Folder Parent=new Folder{StoreID="SOURCE"};
   public bool IsCopy, FailCopy, FailMove, BadSubject, BadDestination;
   public Mail Copy(){
     if(FailCopy)throw new Exception("Injected Copy failure");
     Stats.Copies++;
     return new Mail{EntryID="TEMP-"+EntryID,IsCopy=true,FailMove=FailMove,BadSubject=BadSubject,BadDestination=BadDestination};
   }
   public Mail Move(Folder destination){
     if(!IsCopy){Stats.OriginalMoves++;throw new Exception("Attempt to move original");}
     if(FailMove)throw new Exception("Injected Move failure");
     Parent=BadDestination?new Folder{StoreID="WRONG"}:destination;
     EntryID="PST-"+EntryID;if(BadSubject)Subject="MISMATCH";return this;
   }
   public void Delete(){Stats.Deletes++;throw new Exception("Delete must never be used");}
 }
 public class Namespace {
   public Dictionary<string,Mail> Items=new Dictionary<string,Mail>();
   public Mail GetItemFromID(string id,string store){Stats.Lookups++;return Items[store+"|"+id];}
   public bool CompareEntryIDs(string a,string b){return a=="LONG-"+b || b=="LONG-"+a || String.Equals(a,b,StringComparison.OrdinalIgnoreCase);}
 }
}
'@
function Assert($Condition,[string]$Message){if(-not $Condition){throw ('FAIL: '+$Message)}}
function Message([string]$Id,[string]$Store='SOURCE',[string[]]$Segments=@('Mailbox','Inbox')){
    [pscustomobject]@{id=$Id;storeId=$Store;folderSegments=$Segments;internetMessageId='<test@example>'}
}
function Run-Case([string]$Name,$Messages,$Namespace){
    $ctx=New-JobContext (Join-Path $TestDirectory $Name)
    Invoke-CopyExport $Namespace (New-Object KtgTest.Folder) 'PST' $Messages $ctx
    return $ctx
}
[KtgTest.Stats]::Reset()
$ns=New-Object KtgTest.Namespace
$one=Message 'same-id' 'STORE-A';$two=Message 'same-id' 'STORE-B'
$ns.Items['STORE-A|same-id']=New-Object KtgTest.Mail -Property @{EntryID='same-id';Parent=(New-Object KtgTest.Folder -Property @{StoreID='STORE-A'})}
$ns.Items['STORE-B|same-id']=New-Object KtgTest.Mail -Property @{EntryID='same-id';Parent=(New-Object KtgTest.Folder -Property @{StoreID='STORE-B'})}
$ctx=Run-Case 'identity' @($one,$one,$two) $ns
Assert ($ctx.Copied -eq 2 -and $ctx.Total -eq 2) 'Deduplicate by StoreID plus EntryID'
Assert ($ctx.FolderResolutions -eq 1 -and $ctx.FolderCacheHits -eq 1) 'Destination folder reused without mailbox wrapper'
Assert ([KtgTest.Stats]::OriginalMoves -eq 0 -and [KtgTest.Stats]::Deletes -eq 0) 'Originals untouched'
$ns=New-Object KtgTest.Namespace
$ns.Items['SOURCE|short']=New-Object KtgTest.Mail -Property @{EntryID='LONG-short'}
$ctx=Run-Case 'equivalent-entry-id' @((Message 'short' 'SOURCE' @('Mailbox','Hộp thư','SPL'))) $ns
Assert ($ctx.Copied -eq 1 -and $ctx.Errors -eq 0) 'Equivalent table/OOM entry IDs accepted by Outlook comparison'
Assert ($ctx.FolderResolutions -eq 2) 'Inbox/SPL created directly under PST root'
Assert ((@(Get-PstFolderSegments (Message 'x' 'SOURCE' @('Mailbox','Hộp thư','SPL'))) -join '/') -eq 'Inbox/SPL') 'Localized legacy Inbox maps to Inbox/SPL'
$ns=New-Object KtgTest.Namespace
$cases=New-Object 'Collections.Generic.List[object]'
foreach($kind in @('header','copy-failure','move-failure','bad-subject','bad-destination','bad-id')){
    $m=Message $kind;$cases.Add($m)
    $item=New-Object KtgTest.Mail -Property @{EntryID=$kind}
    switch($kind){
        'header'{$item.DownloadState=0}
        'copy-failure'{$item.FailCopy=$true}
        'move-failure'{$item.FailMove=$true}
        'bad-subject'{$item.BadSubject=$true}
        'bad-destination'{$item.BadDestination=$true}
        'bad-id'{$item.EntryID='changed'}
    }
    $ns.Items['SOURCE|'+$kind]=$item
}
$ctx=Run-Case 'failures' $cases.ToArray() $ns
Assert ($ctx.Errors -eq 6 -and $ctx.Copied -eq 0) 'All injected failures are recorded'
$invalidNs=New-Object KtgTest.Namespace;$invalidNs.Items['SOURCE|changed']=New-Object KtgTest.Mail -Property @{EntryID='different'}
$invalidCtx=Run-Case 'no-empty-folder' @((Message 'changed')) $invalidNs
Assert ($invalidCtx.Errors -eq 1 -and $invalidCtx.FolderResolutions -eq 0) 'Invalid source does not create empty folders'
$report=@(Import-Csv (Join-Path $ctx.Directory 'report.csv'))
Assert ($report.Count -eq 6) 'Report includes failures'
Assert ($report[2].temporaryCopyEntryId -ne '') 'Temporary copy tracked after Move failure'
Assert ([KtgTest.Stats]::OriginalMoves -eq 0 -and [KtgTest.Stats]::Deletes -eq 0) 'No original moved/deleted on failure'
$cancel=New-JobContext (Join-Path $TestDirectory 'cancel')
[IO.File]::WriteAllText((Join-Path $cancel.Directory 'cancel'),'cancel')
Invoke-CopyExport $ns (New-Object KtgTest.Folder) 'PST' @($cases[0]) $cancel
Assert ($cancel.Status -eq 'cancelled' -and $cancel.Processed -eq 0) 'Cooperative cancellation'
$encoded=New-Object IO.StringWriter
Write-CsvRecord $encoded @('=HYPERLINK("bad")','hello, "world"')
Assert ($encoded.ToString().StartsWith('"''=')) 'CSV formula injection protected'
# Regression: a provider default Inbox can exist outside the visible PST tree.
$ns=New-Object KtgTest.Namespace
$ns.Items['SOURCE|visible']=New-Object KtgTest.Mail -Property @{EntryID='visible'}
$visibleRoot=New-Object KtgTest.Folder
$orphan=New-Object KtgTest.Folder
$ctx=New-JobContext (Join-Path $TestDirectory 'ignore-orphan-default')
Invoke-CopyExport $ns $visibleRoot 'PST' @((Message 'visible' 'SOURCE' @('Mailbox','Inbox','SPL'))) $ctx @{Inbox=@{Folder=$orphan;EntryId=$orphan.EntryID}}
$visibleInbox=$visibleRoot.Folders.Item('Inbox');$visibleSpl=$visibleInbox.Folders.Item('SPL')
$batch=Get-Content (Join-Path $ctx.Directory 'results-0.json') -Raw -Encoding UTF8|ConvertFrom-Json
Assert ($ctx.Copied -eq 1 -and $batch[0].archiveProof.destinationFolderId -eq $visibleSpl.EntryID) 'Provider orphan Inbox ignored; copy reachable through visible root'
Assert ([KtgExportIo.Fast]::PathKey([string[]]@('ab','c')) -cne [KtgExportIo.Fast]::PathKey([string[]]@('a','bc'))) 'Path cache keys preserve segment boundaries'
Assert ([KtgExportIo.Fast]::PathKey([string[]]@('a/b')) -cne [KtgExportIo.Fast]::PathKey([string[]]@('a','b'))) 'Slash in folder name does not merge paths'
$csv=New-Object IO.StringWriter
Write-CsvRecord $csv @('Tiếng Việt',"quote`" newline`n",'=SUM(1,2)',-2,'+x','@x')
$decoded=@($csv.ToString()|ConvertFrom-Csv -Header a,b,c,d,e,f)
Assert ($decoded[0].a -ceq 'Tiếng Việt' -and $decoded[0].b -ceq "quote`" newline`n" -and $decoded[0].c -ceq "'=SUM(1,2)" -and $decoded[0].d -ceq "'-2") 'Fast CSV preserves Unicode, quoting, newlines and formula protection'
# Synthetic benchmark: in-memory Outlook objects, not real COM/PST throughput.
[KtgTest.Stats]::Reset()
$ns=New-Object KtgTest.Namespace
$messages=New-Object 'Collections.Generic.List[object]'
for($i=0;$i -lt $BenchmarkCount;$i++){
    $id=[string]$i
    $segments=if(($i%2)-eq 0){@('Mailbox','Inbox')}else{@('Mailbox','Sent Items')}
    $messages.Add((Message $id 'SOURCE' $segments))
    $ns.Items['SOURCE|'+$id]=New-Object KtgTest.Mail -Property @{EntryID=$id}
}
$ctx=Run-Case 'benchmark' $messages.ToArray() $ns
Assert ($ctx.Copied -eq $BenchmarkCount -and $ctx.Errors -eq 0) 'Benchmark verified count'
Assert ($ctx.FolderResolutions -eq 2) 'Only Inbox and Sent Items resolved, no mailbox wrapper'
Assert ($ctx.FolderCacheHits -eq ($BenchmarkCount-2)) 'Cache hit count'
Assert ([KtgTest.Stats]::Lookups -eq $BenchmarkCount -and [KtgTest.Stats]::Copies -eq $BenchmarkCount) 'One source lookup and copy per mail'
Assert ([KtgTest.Stats]::OriginalMoves -eq 0 -and [KtgTest.Stats]::Deletes -eq 0) 'Benchmark preserves every original'
$rows=@(Import-Csv (Join-Path $ctx.Directory 'report.csv'))
Assert ($rows.Count -eq $BenchmarkCount) 'Complete benchmark report'
$result=@{
    test='synthetic-mock-only';count=$BenchmarkCount;seconds=$ctx.Timer.Elapsed.TotalSeconds
    folderResolutions=$ctx.FolderResolutions;folderCacheHits=$ctx.FolderCacheHits
    resultBatches=$ctx.ResultBatches;sourceLookups=[KtgTest.Stats]::Lookups
    originalMoves=[KtgTest.Stats]::OriginalMoves;deletes=[KtgTest.Stats]::Deletes
    testDirectory=$TestDirectory;passed=$true
}
Write-AtomicJson (Join-Path $TestDirectory 'test-results.json') $result
$result|ConvertTo-Json
