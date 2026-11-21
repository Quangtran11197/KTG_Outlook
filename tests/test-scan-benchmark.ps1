param([string]$ScriptsDirectory,[string]$TestDirectory,[int]$Count=70000)
$ErrorActionPreference='Stop'
if(-not $ScriptsDirectory){$ScriptsDirectory=Join-Path $PSScriptRoot '..\scripts'}
if(-not $TestDirectory){$TestDirectory=Join-Path ([IO.Path]::GetTempPath()) ('KTG-scan-bench-'+[guid]::NewGuid().ToString('N'))}
. (Join-Path $ScriptsDirectory 'core.ps1')
$t=$null;$e=$null;$ast=[Management.Automation.Language.Parser]::ParseFile((Join-Path $ScriptsDirectory 'bridge.ps1'),[ref]$t,[ref]$e)
$f=$ast.Find({param($n)$n -is [Management.Automation.Language.FunctionDefinitionAst] -and $n.Name -eq 'Read-MailFolder'},$true);Invoke-Expression $f.Extent.Text
Add-Type @'
using System;
namespace LargeScan {
public class Columns {public void RemoveAll(){}public object Add(string name){return null;}}
public class Table {public Columns Columns=new Columns();public int Position,Total,Calls;public bool EndOfTable {get{return Position>=Total;}}public object[,] GetArray(int count){Calls++;int n=Math.Min(count,Total-Position);var a=new object[n,8];for(int i=0;i<n;i++){a[i,0]="id"+(Position+i);a[i,1]="Tiếng Việt \"test\"\n"+(Position+i);a[i,2]="IPM.Note";a[i,3]=new DateTime(2026,1,1);a[i,4]=1024L;a[i,5]="raw";a[i,6]="smtp@example.invalid";a[i,7]="<"+(Position+i)+">";}Position+=n;return a;}}
public class Children {public int Count=0;}
public class Folder {public int DefaultItemType=0;public string EntryID="folder",Name="Inbox";public Children Folders=new Children();public Table Table=new Table();public Table GetTable(string f,int flags){return Table;}}
}
'@
$script:defaultFolderMap=@{'folder'='Inbox'};$script:sentId='sent';$script:fromDate=[datetime]'2026-01-01';$script:toDate=[datetime]'2027-01-01'
$folder=New-Object LargeScan.Folder;$folder.Table.Total=$Count;$ctx=New-JobContext $TestDirectory
Read-MailFolder $folder @('Mailbox','Inbox') $false 'SOURCE' ([pscustomobject]@{allMail=$true}) $ctx
Flush-ScanBatch $ctx;$elapsed=$ctx.Timer.Elapsed.TotalSeconds
if($ctx.Processed -ne $Count -or $ctx.Errors -ne 0){throw ('Wrong result '+$ctx.Processed+' '+$ctx.Note)}
$files=@(Get-ChildItem $TestDirectory -Filter messages-*.json);$decoded=0
foreach($file in $files){$batch=Get-Content $file.FullName -Raw -Encoding UTF8|ConvertFrom-Json;$decoded+=@($batch).Count;foreach($m in $batch){if($m.sender -ne 'smtp@example.invalid' -or $m.subject -notlike 'Tiếng Việt*' -or ($m.pstFolderSegments -join '/') -ne 'Inbox'){throw 'Metadata mismatch'}}}
if($decoded -ne $Count){throw 'Lost rows'}
@{passed=$true;synthetic=$true;count=$Count;seconds=[math]::Round($elapsed,3);batches=$ctx.Batches;tableCalls=$folder.Table.Calls;decoded=$decoded}|ConvertTo-Json
