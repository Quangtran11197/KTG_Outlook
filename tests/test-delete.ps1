param([string]$TestDirectory)
$ErrorActionPreference='Stop'
. "$PSScriptRoot\..\scripts\core.ps1"
if(-not $TestDirectory){$TestDirectory=Join-Path ([IO.Path]::GetTempPath()) ('KTG-delete-'+[guid]::NewGuid().ToString('N'))}
Add-Type -TypeDefinition @'
using System;
using System.Collections.Generic;
namespace KtgDeleteTest {
 public static class Stats {public static int Moves,Deletes,Lookups,Copies,PstMoves,PathReads; public static void Reset(){Moves=Deletes=Lookups=Copies=PstMoves=PathReads=0;}}
 public class Folder {public string EntryID="inbox",StoreID="SOURCE";public Folders Folders;public Folder(){Folders=new Folders(this);}}
 public class Folders {
  Folder parent;Dictionary<string,Folder> items=new Dictionary<string,Folder>();
  public Folders(Folder owner){parent=owner;}
  public Folder Item(string name){Stats.PathReads++;return items[name];} public void Remove(string name){items.Remove(name);}
  public Folder Add(string name,int kind){var folder=new Folder{EntryID=Guid.NewGuid().ToString(),StoreID=parent.StoreID};items[name]=folder;return folder;}
 }
 public class Store {
  public int ExchangeStoreType=0;public bool IsCachedExchange=true;
  public Folder Trash=new Folder{EntryID="trash"};
  public Folder Root=new Folder{StoreID="PST"};public Folder GetRootFolder(){Stats.PathReads++;return Root;}
  public Folder GetDefaultFolder(int kind){if(kind!=3)throw new Exception("Wrong default");return Trash;}
 }
 public class Accessor {public string GetProperty(string name){return "<id>";}}
 public class Attachments {public int Count=2;}
 public class Mail {
  public int Class=43;public string EntryID="id1",Subject="Mail"; public Folder Parent=new Folder();
  public long Size=4096;public Attachments Attachments=new Attachments();
  public DateTime LastModificationTime=new DateTime(2026,1,1);
  public int DownloadState=1;public bool IsCopy,FailCopy,ChangeSourceAfterCopy;public Namespace Owner;
  public Accessor PropertyAccessor=new Accessor();
  public bool FailMove,FailDelete,BadDestination,Fresh,SilentDelete,FailReopen;public string CancelPath,CancelAfterSourceMove;
  public Mail Copy(){
   Stats.Copies++;if(FailCopy)throw new Exception("Copy failed");
   return new Mail{EntryID="temp-"+EntryID,Subject=Subject,IsCopy=true,Owner=Owner,CancelPath=CancelPath,ChangeSourceAfterCopy=ChangeSourceAfterCopy};
  }
  public Mail Move(Folder destination){
   if(IsCopy){
    Stats.PstMoves++;var result=new Mail{EntryID="pst-"+EntryID,Subject=Subject,Parent=destination,Owner=Owner,IsCopy=true};
    if(Owner!=null){Owner.Items[destination.StoreID+"|"+result.EntryID]=result;if(ChangeSourceAfterCopy)Owner.Items["SOURCE|"+EntryID.Substring(5)].Size=9999;}
    if(!String.IsNullOrEmpty(CancelPath))System.IO.File.WriteAllText(CancelPath,"cancel");
    return result;
   }
   Stats.Moves++;if(FailMove)throw new Exception("Move failed");
   if(!String.IsNullOrEmpty(CancelAfterSourceMove))System.IO.File.WriteAllText(CancelAfterSourceMove,"cancel");
   if(!String.IsNullOrEmpty(CancelPath))System.IO.File.WriteAllText(CancelPath,"cancel");
    var moved=new Mail{EntryID="moved-"+EntryID,Subject=Subject,Parent=BadDestination?new Folder{EntryID="wrong"}:destination,FailDelete=FailDelete,SilentDelete=SilentDelete,FailReopen=FailReopen,Owner=Owner};
    if(Owner!=null){Owner.Items.Remove("SOURCE|"+EntryID);Owner.Items["SOURCE|"+moved.EntryID]=moved;}
    return moved;
  }
  public void Delete(){if(Parent.EntryID!="trash")throw new Exception("Delete outside trash");Stats.Deletes++;if(FailDelete)throw new Exception("Delete failed");if(Fresh&&!SilentDelete&&Owner!=null)Owner.Items.Remove("SOURCE|"+EntryID);}
 }
 public class Namespace {
  public Store Store=new Store(),DestinationStore=new Store();public Dictionary<string,Mail> Items=new Dictionary<string,Mail>();public bool MissingPst;
  public List<string> Opened=new List<string>();
  public Store GetStoreFromID(string id){return id=="PST"?DestinationStore:Store;}
  public Mail GetItemFromID(string id,string store){Stats.Lookups++;Opened.Add(store+"|"+id);if(store=="PST"&&MissingPst)throw new Exception("PST copy missing");if(!Items.ContainsKey(store+"|"+id))throw new System.Runtime.InteropServices.COMException("MAPI_E_NOT_FOUND",unchecked((int)0x8004010F));var item=Items[store+"|"+id];if(item.FailReopen&&id.StartsWith("moved-"))throw new Exception("Reopen failed");item.Owner=this;item.Fresh=true;return item;}
 }
}
'@
function Assert($Value,[string]$Message){if(-not $Value){throw ('FAIL: '+$Message)}}
function Message([string]$Id='id1',[string]$Folder='inbox'){
 return [pscustomobject]@{id=$Id;storeId='SOURCE';sourceFolderId=$Folder;subject='Mail';internetMessageId='<id>'}
}
function Request($Messages){
 $unique=@(Select-UniqueMessages $Messages)
 return [pscustomobject]@{operation='delete';messages=@($Messages);confirmPermanentDelete=$true;confirmationText=('XOA '+$unique.Count)}
}
function RunCase([string]$Name,$Request,$Namespace){
 $ctx=New-JobContext (Join-Path $TestDirectory $Name)
 Invoke-ConfirmedDelete $Namespace $Request $ctx
 return $ctx
}
[KtgDeleteTest.Stats]::Reset()
foreach($bad in @(
 [pscustomobject]@{operation='delete'},
 [pscustomobject]@{messages=@((Message));confirmPermanentDelete=$false;confirmationText='XOA 1'},
 [pscustomobject]@{messages=@((Message));confirmPermanentDelete='true';confirmationText='XOA 1'},
 [pscustomobject]@{messages=@((Message));confirmPermanentDelete=$true;confirmationText='XOA 2'},
 [pscustomobject]@{messages=@((Message));confirmPermanentDelete=$true;confirmationText='xoa 1'},
 [pscustomobject]@{messages=@([pscustomobject]@{id='id';storeId='SOURCE'});confirmPermanentDelete=$true;confirmationText='XOA 1'}
)){
 $threw=$false
 try{Invoke-ConfirmedDelete (New-Object KtgDeleteTest.Namespace) $bad (New-JobContext (Join-Path $TestDirectory ('invalid-'+[guid]::NewGuid().ToString('N'))))}catch{$threw=$true}
 Assert $threw 'Missing/wrong confirmation or missing source folder refused'
}
Assert ([KtgDeleteTest.Stats]::Moves -eq 0 -and [KtgDeleteTest.Stats]::Deletes -eq 0 -and [KtgDeleteTest.Stats]::Lookups -eq 0) 'No mail access for invalid confirmation'
$ns=New-Object KtgDeleteTest.Namespace;$ns.Items['SOURCE|id1']=New-Object KtgDeleteTest.Mail
$m=Message
$ctx=RunCase 'confirmed' (Request @($m,$m)) $ns
Assert ($ctx.Deleted -eq 1 -and $ctx.Total -eq 1 -and [KtgDeleteTest.Stats]::Moves -eq 1 -and [KtgDeleteTest.Stats]::Deletes -eq 1) 'Confirmed delete deduplicates, moves then requests delete'
$rows=@(Import-Csv (Join-Path $ctx.Directory 'report.csv'))
Assert ($rows[0].status -eq 'delete_intent' -and $rows[1].status -eq 'in_deleted_items' -and $rows[2].status -eq 'delete_requested') 'Intent and stage recorded'
Assert ([IO.File]::Exists($ctx.LogPath)) 'Operations log exists'
[KtgDeleteTest.Stats]::Reset();$ns=New-Object KtgDeleteTest.Namespace
$ns.Items['SOURCE|id1']=New-Object KtgDeleteTest.Mail -Property @{Parent=(New-Object KtgDeleteTest.Folder -Property @{EntryID='trash'})}
$ctx=RunCase 'already-trash' (Request @((Message 'id1' 'trash'))) $ns
Assert ($ctx.Deleted -eq 1 -and [KtgDeleteTest.Stats]::Moves -eq 0 -and [KtgDeleteTest.Stats]::Deletes -eq 1) 'Already in Deleted Items: no second move'
foreach($failure in @('wrong-folder','wrong-subject','wrong-id','not-cached','move-failed','wrong-destination','delete-failed','silent-delete','reopen-failed')){
 [KtgDeleteTest.Stats]::Reset();$ns=New-Object KtgDeleteTest.Namespace;$item=New-Object KtgDeleteTest.Mail
 switch($failure){
  'wrong-folder'{$item.Parent.EntryID='other'}
  'wrong-subject'{$item.Subject='other'}
  'wrong-id'{$item.EntryID='other'}
  'not-cached'{$ns.Store.IsCachedExchange=$false}
  'move-failed'{$item.FailMove=$true}
  'wrong-destination'{$item.BadDestination=$true}
  'delete-failed'{$item.FailDelete=$true}
  'silent-delete'{$item.SilentDelete=$true}
  'reopen-failed'{$item.FailReopen=$true}
 }
 $ns.Items['SOURCE|id1']=$item;$ctx=RunCase $failure (Request @((Message))) $ns
 Assert ($ctx.Deleted -eq 0 -and $ctx.Errors -eq 1) ($failure+' is recorded, not reported as successful')
 if($failure -notin @('delete-failed','silent-delete')){Assert ([KtgDeleteTest.Stats]::Deletes -eq 0) ($failure+' never reaches Delete')}
 $results=Get-Content -LiteralPath (Join-Path $ctx.Directory 'results-0.json') -Raw -Encoding UTF8|ConvertFrom-Json
 Assert $results[0].needsRescan 'Failure requires rescan, not blind retry'
}
[KtgDeleteTest.Stats]::Reset();$ns=New-Object KtgDeleteTest.Namespace
$ctx=New-JobContext (Join-Path $TestDirectory 'cancel-after-move')
$ns.Items['SOURCE|id1']=New-Object KtgDeleteTest.Mail -Property @{CancelPath=(Join-Path $ctx.Directory 'cancel')}
Invoke-ConfirmedDelete $ns (Request @((Message))) $ctx
Assert ($ctx.Status -eq 'cancelled' -and [KtgDeleteTest.Stats]::Moves -eq 1 -and [KtgDeleteTest.Stats]::Deletes -eq 0) 'Cancel after move leaves item in trash, no permanent delete'
$session=Join-Path $TestDirectory 'archive-session'
$exportDirectory=Join-Path $session 'job-123-1'
[IO.Directory]::CreateDirectory($exportDirectory)|Out-Null
$proof=@{pstMode='existing';sourceMessage=(Message);sourceModifiedTicks=([datetime]'2026-01-01').Ticks;destinationStoreId='PST';destinationId='pst-id';destinationFolderId='pst-inbox-spl';subject='Mail';size=4096;attachmentCount=2;internetMessageId='<id>'}
Write-AtomicJson (Join-Path $exportDirectory 'results-0.json') @(@{key='SOURCE|id1';status='verified';archiveProof=$proof})
$archiveScenarios=@('verified','missing-copy','wrong-pst-folder','wrong-pst-subject','unverified-source','wrong-pst-mode','source-size-changed','source-attachments-changed','source-edited')
foreach($scenario in $archiveScenarios){
 [KtgDeleteTest.Stats]::Reset();$ns=New-Object KtgDeleteTest.Namespace
 $ns.Items['SOURCE|id1']=New-Object KtgDeleteTest.Mail
 if($scenario -eq 'source-size-changed'){$ns.Items['SOURCE|id1'].Size=9000}
 if($scenario -eq 'source-attachments-changed'){$ns.Items['SOURCE|id1'].Attachments.Count=3}
 if($scenario -eq 'source-edited'){$ns.Items['SOURCE|id1'].LastModificationTime=[datetime]'2026-01-02'}
 $copy=New-Object KtgDeleteTest.Mail -Property @{EntryID='pst-id';Parent=(New-Object KtgDeleteTest.Folder -Property @{StoreID='PST';EntryID='pst-inbox-spl'})}
 if($scenario -ne 'missing-copy'){$ns.Items['PST|pst-id']=$copy}
 if($scenario -eq 'wrong-pst-folder'){$copy.Parent.EntryID='different-folder'}
 if($scenario -eq 'wrong-pst-subject'){$copy.Subject='different-subject'}
 $request=Request @((Message));$request.operation='delete_archived';$request|Add-Member archiveJobId 'job-123-1'
 $ctx=New-JobContext (Join-Path $session ('job-456-'+([array]::IndexOf($archiveScenarios,$scenario)+1)))
 $receipts=Get-ArchiveReceipts $request $ctx.Directory
 if($scenario -eq 'unverified-source'){$receipts.Clear()}
 if($scenario -eq 'wrong-pst-mode'){$receipts['SOURCE|id1'].pstMode='new'}
 Invoke-ConfirmedDelete $ns $request $ctx $receipts
 if($scenario -eq 'verified'){
  Assert ($ctx.Deleted -eq 1 -and [KtgDeleteTest.Stats]::Deletes -eq 1) 'Verified PST copy permits confirmed source delete'
  Assert ($ns.Opened[0] -eq 'PST|pst-id') 'PST is reopened before accessing source'
 }else{
  Assert ($ctx.Deleted -eq 0 -and $ctx.Errors -eq 1 -and [KtgDeleteTest.Stats]::Moves -eq 0 -and [KtgDeleteTest.Stats]::Deletes -eq 0) ($scenario+' preserves original')
 }
}
$badRequest=Request @((Message));$badRequest|Add-Member archiveJobId '..\other'
$threw=$false;try{[void](Get-ArchiveReceipts $badRequest (Join-Path $session 'job-999-1'))}catch{$threw=$true}
Assert $threw 'Archive job path traversal rejected'
@{passed=$true;test='delete-synthetic-only';scenarios=28;testDirectory=$TestDirectory}|ConvertTo-Json
