param([string]$TestDirectory)
$ErrorActionPreference='Stop'
. "$PSScriptRoot\..\scripts\core.ps1"
if(-not $TestDirectory){$TestDirectory=Join-Path ([IO.Path]::GetTempPath()) ('KTG-scan-'+[guid]::NewGuid().ToString('N'))}
# Import only the folder function, so the production entry point cannot attach to Outlook.
$tokens=$null;$errors=$null
$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $PSScriptRoot '..\scripts\bridge.ps1'),[ref]$tokens,[ref]$errors)
if($errors.Count -gt 0){throw ($errors|Out-String)}
$function=$ast.Find({param($node) $node -is [Management.Automation.Language.FunctionDefinitionAst] -and $node.Name -eq 'Read-MailFolder'},$true)
Invoke-Expression $function.Extent.Text
Add-Type -TypeDefinition @'
using System;
namespace KtgScanTest {
 public class Columns {public void RemoveAll(){} public object Add(string name){return null;}}
 public class Table {
   public Columns Columns=new Columns();
   public bool EndOfTable=false;
   public string CancelPath;
   public object[,] GetArray(int count) {
     EndOfTable=true;
     if(!String.IsNullOrEmpty(CancelPath))System.IO.File.WriteAllText(CancelPath,"cancel");
     return new object[,] {
       {"id1","Tiêu đề thử","IPM.Note",new DateTime(2023,2,1),1024,"raw","smtp@example.invalid","<one>"},
       {"id2","Outside range","IPM.Note",new DateTime(2024,2,1),2048,"raw","smtp@example.invalid","<two>"},
       {"id3","Calendar","IPM.Appointment",new DateTime(2023,2,1),300,"raw","smtp@example.invalid","<three>"}
     };
   }
 }
 public class Children {public int Count=0;public Folder Item(int index){throw new Exception("No children");}}
 public class Folder {
   public int DefaultItemType=0;public string EntryID="folder1";public string Name="Inbox";
   public Children Folders=new Children();public string Filter="";
   public string CancelPath;
   public bool Empty; public Table GetTable(string filter,int flags){Filter=filter;return new Table{CancelPath=CancelPath,EndOfTable=Empty};}
 }
}
'@
function Assert($Condition,[string]$Message){if(-not $Condition){throw ('FAIL: '+$Message)}}
Assert (Test-CachedExchangeStore ([pscustomobject]@{ExchangeStoreType=0;IsCachedExchange=$true})) 'Primary cached Exchange mailbox included'
Assert (Test-CachedExchangeStore ([pscustomobject]@{ExchangeStoreType=1;IsCachedExchange=$true})) 'Secondary cached Exchange mailbox included'
Assert (-not (Test-CachedExchangeStore ([pscustomobject]@{ExchangeStoreType=3;IsCachedExchange=$true}))) 'PST excluded even if cache flag is true'
Assert (-not (Test-CachedExchangeStore ([pscustomobject]@{ExchangeStoreType=0;IsCachedExchange=$false}))) 'Online-only Exchange excluded'
Assert (-not (Test-CachedExchangeStore ([pscustomobject]@{ExchangeStoreType=0}))) 'Unknown cache status excluded'
Assert (-not (Test-CachedExchangeStore ([pscustomobject]@{ExchangeStoreType=2;IsCachedExchange=$false}))) 'Uncached public folders excluded'
$script:defaultFolderMap=@{'folder1'='Inbox'}
$script:sentId='sent';$script:fromDate=[datetime]'2023-01-01';$script:toDate=[datetime]'2024-01-01'
$folder=New-Object KtgScanTest.Folder
$ctx=New-JobContext (Join-Path $TestDirectory 'all')
Read-MailFolder $folder @('Mailbox','Inbox') $false 'SOURCE' ([pscustomobject]@{allMail=$true}) $ctx
Flush-ScanBatch $ctx
Assert ($ctx.Processed -eq 2 -and $ctx.Errors -eq 0) 'Scan reads 2D table and skips non-mail'
$rows=Get-Content -LiteralPath (Join-Path $ctx.Directory 'messages-0.json') -Raw -Encoding UTF8|ConvertFrom-Json
Assert ($rows[0].id -eq 'id1' -and $rows[0].sender -eq 'smtp@example.invalid') 'Table column mapping'
Assert ($rows[0].subject -eq 'Tiêu đề thử') 'Unicode preserved'
Assert ($rows[0].folderSegments.Count -eq 2 -and $rows[0].storeId -eq 'SOURCE') 'Folder segments and store identity'
Assert (($rows[0].pstFolderSegments -join '/') -eq 'Inbox') 'PST path excludes mailbox and uses canonical Inbox'
$ctx=New-JobContext (Join-Path $TestDirectory 'range')
Read-MailFolder $folder @('Mailbox','Inbox') $false 'SOURCE' ([pscustomobject]@{allMail=$false}) $ctx
Assert ($ctx.Processed -eq 1 -and $ctx.Errors -eq 0) 'Date range checked'
Assert ($folder.Filter.StartsWith('[ReceivedTime]')) 'Inbox filter'
$ctx=New-JobContext (Join-Path $TestDirectory 'sent')
Read-MailFolder $folder @('Mailbox','Sent') $true 'SOURCE' ([pscustomobject]@{allMail=$false}) $ctx
Assert ($folder.Filter.StartsWith('[SentOn]')) 'Sent filter'
$ctx=New-JobContext (Join-Path $TestDirectory 'cancel')
[IO.File]::WriteAllText((Join-Path $ctx.Directory 'cancel'),'cancel')
Read-MailFolder $folder @('Mailbox','Inbox') $false 'SOURCE' ([pscustomobject]@{allMail=$true}) $ctx
Assert ($ctx.Processed -eq 0 -and $ctx.Status -eq 'cancelled') 'Scan cancellation'
$ctx=New-JobContext (Join-Path $TestDirectory 'cancel-after-table-read')
$folder.CancelPath=Join-Path $ctx.Directory 'cancel'
Read-MailFolder $folder @('Mailbox','Inbox') $false 'SOURCE' ([pscustomobject]@{allMail=$true}) $ctx
Assert ($ctx.Processed -eq 0 -and $ctx.Status -eq 'cancelled') 'Cancellation checked inside table batch'
$empty=New-Object KtgScanTest.Folder;$empty.Empty=$true
$ctx=New-JobContext (Join-Path $TestDirectory 'empty')
Read-MailFolder $empty @('Mailbox','Inbox') $false 'SOURCE' ([pscustomobject]@{allMail=$true}) $ctx
Assert ($ctx.Processed -eq 0 -and $ctx.Errors -eq 0 -and $ctx.Batches -eq 0) 'Empty folders have no batches'
$matrix=(New-Object KtgScanTest.Table).GetArray(2000)
$subject="Unicode Việt `"quote`" \ slash `n `r `t `0"
$matrix[0,1]=$subject;$matrix[0,3]=([datetime]'2023-02-01').ToOADate();$matrix[1,3]=[DBNull]::Value
$out=Join-Path $TestDirectory 'encoder-edge.json'
$result=[KtgFastScan.Encoder]::Write($matrix,[int[]]@(0,1,2,3,-1,5,-1,-1),'SOURCE','folder1',[string[]]@('Mailbox','Inbox'),[string[]]@('Inbox'),$true,[datetime]::MinValue,[datetime]::MaxValue,$out,(Join-Path $TestDirectory 'no-cancel'))
$decoded=Get-Content $out -Raw -Encoding UTF8|ConvertFrom-Json
Assert ($result.Count -eq 1 -and $decoded[0].subject -ceq $subject) 'JSON escapes Unicode, quotes, slashes, controls without data loss'
Assert ($decoded[0].sender -eq 'raw' -and $decoded[0].size -eq 0 -and $decoded[0].internetMessageId -eq '') 'Missing optional columns use safe defaults'
Assert ($decoded[0].date -eq '2023-02-01 00:00') 'OLE date and null dates handled'
@{passed=$true;checks=20;testDirectory=$TestDirectory}|ConvertTo-Json