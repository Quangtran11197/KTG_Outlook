param([Parameter(Mandatory=$true)][string]$InputFile)
$ErrorActionPreference='Stop'
. "$PSScriptRoot\core.ps1"
. "$PSScriptRoot\bulk.ps1"
$directory=[IO.Path]::GetDirectoryName([IO.Path]::GetFullPath($InputFile))
$context=New-JobContext $directory
$outlook=$null; $namespace=$null; $root=$null; $destinationStore=$null
function Read-MailFolder($Folder, [string[]]$Segments, [bool]$Sent, [string]$StoreId, $Request, $Context, [string[]]$PstSegments=@()) {
    if (Test-Cancelled $Context) { $Context.Status='cancelled'; return }
    $table=$null; $columns=$null; $children=$null
    $relativePath=@($PstSegments)
    if($script:defaultFolderMap.ContainsKey([string]$Folder.EntryID)){$relativePath=@($script:defaultFolderMap[[string]$Folder.EntryID])}
    try {
        if ([int]$Folder.DefaultItemType -eq 0) {
            $dateColumn=if($Sent){'SentOn'}else{'ReceivedTime'}
            $filter=''
            if (-not [bool]$Request.allMail) {
                $culture=[Globalization.CultureInfo]::CurrentCulture
                $filter="[$dateColumn] >= '"+$script:fromDate.ToString('g',$culture)+"' AND [$dateColumn] < '"+$script:toDate.ToString('g',$culture)+"'"
            }
            $table=$Folder.GetTable($filter,0)
            if(-not $table.EndOfTable){
            $sourceFolderId=[string]$Folder.EntryID
            $columns=$table.Columns; $columns.RemoveAll()
            $map=@{}
            foreach($name in @('EntryID','Subject','MessageClass',$dateColumn,'Size','SenderEmailAddress','http://schemas.microsoft.com/mapi/proptag/0x5D01001F','http://schemas.microsoft.com/mapi/proptag/0x1035001F')) {
                try { $column=$columns.Add($name); Release-Com $column; $map[$name]=$map.Count }
                catch { if($name -in @('EntryID','Subject','MessageClass',$dateColumn)){throw} }
            }
            Initialize-FastScan
            $indices=New-Object 'int[]' 8
            $names=@('EntryID','Subject','MessageClass',$dateColumn,'Size','SenderEmailAddress','http://schemas.microsoft.com/mapi/proptag/0x5D01001F','http://schemas.microsoft.com/mapi/proptag/0x1035001F')
            for($c=0;$c -lt 8;$c++){$indices[$c]=if($map.ContainsKey($names[$c])){[int]$map[$names[$c]]}else{-1}}
            while (-not $table.EndOfTable) {
                if(Test-Cancelled $Context){$Context.Status='cancelled';return}
                $rows=$table.GetArray(1000)
                if($null -eq $rows -or $rows.Length -eq 0){break}
                $path=Join-Path $Context.Directory ('messages-'+$Context.Batches+'.json')
                $result=[KtgFastScan.Encoder]::Write($rows,$indices,$StoreId,$sourceFolderId,[string[]]$Segments,[string[]]$relativePath,[bool]$Request.allMail,$script:fromDate,$script:toDate,$path,(Join-Path $Context.Directory 'cancel'))
                if($result.Count -gt 0){$Context.Processed+=$result.Count;$Context.Batches++}
                Save-JobState $Context
                if($result.Cancelled){$Context.Status='cancelled';return}
            }
        }
            }
    } catch { $Context.Errors++; $Context.Note='Không đọc đầy đủ thư mục '+($Segments -join ' / ')+': '+$_.Exception.Message }
    finally {Release-Com $columns; Release-Com $table}
    try {
        $children=$Folder.Folders
        $count=[int]$children.Count
        for($i=1;$i -le $count;$i++) {
            if(Test-Cancelled $Context){$Context.Status='cancelled';break}
            $child=$null
            try {
                $child=$children.Item($i)
                $isSent=$Sent -or ([string]$child.EntryID -eq $script:sentId)
                Read-MailFolder $child @($Segments+@([string]$child.Name)) $isSent $StoreId $Request $Context @($relativePath+@([string]$child.Name))
            } finally {Release-Com $child}
        }
    } catch {$Context.Errors++;$Context.Note='Không đọc được thư mục con: '+$_.Exception.Message}
    finally {Release-Com $children}
}
try {
    Save-JobState $context -Force
    $request=Get-Content -LiteralPath $InputFile -Raw -Encoding UTF8 | ConvertFrom-Json
    if ($request.operation -notin @('scan','export','delete','delete_archived')) { throw 'Thao tác không được hỗ trợ.' }
    $archiveReceipts=$null
    if ($request.operation -in @('delete','delete_archived')) {[void]@(Get-ConfirmedDeleteMessages $request)}
    if ($request.operation -eq 'delete_archived'){$archiveReceipts=Get-ArchiveReceipts $request $directory}
    if ($request.operation -eq 'export' -and @($request.messages).Count -eq 0) {throw 'Chưa chọn thư.'}
    if ($request.operation -eq 'export'){$context.AutoCleanup=Get-AutoCleanupEnabled $request}
    $context.Note=switch($request.operation){'scan'{'Bắt đầu quét cache Exchange cục bộ.'} 'export'{if($context.AutoCleanup){'Bắt đầu lưu PST và tự xóa nguồn sau khi bản sao được kiểm tra.'}else{'Bắt đầu sao chép PST; giữ nguyên thư gốc.'}} 'delete'{'Bắt đầu thao tác xóa riêng đã được xác nhận.'} 'delete_archived'{'Mở lại bản PST để kiểm tra trước khi xóa nguồn đã xác nhận.'}}
    Save-JobState $context -Force
    # Attach only: do not start Outlook or select a profile without the user's action.
    try{$outlook=[Runtime.InteropServices.Marshal]::GetActiveObject('Outlook.Application')}
    catch{throw 'Chưa kết nối được Outlook Classic. Hãy mở Outlook Classic, đăng nhập profile và chọn Work Offline; chạy Outlook và công cụ bằng cùng tài khoản/quyền.'}
    $namespace=$outlook.GetNamespace('MAPI')
    $bulkRequested=$request.operation -eq 'export' -and $request.PSObject.Properties['bulkCopy'] -and $request.bulkCopy -is [bool] -and $request.bulkCopy
    if (-not [bool]$namespace.Offline -and -not $bulkRequested -and $request.operation -ne 'scan') {throw 'Chế độ hiện tại cần Work Offline. Chuyển sang MAPI theo lô để xuất trong khi Outlook vẫn trực tuyến.'}
    if($request.operation -eq 'scan') {
        $script:fromDate=[datetime]::MinValue;$script:toDate=[datetime]::MaxValue
        if(-not [bool]$request.allMail) {
            $script:fromDate=[datetime]::ParseExact([string]$request.from,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture)
            $script:toDate=[datetime]::ParseExact([string]$request.to,'yyyy-MM-dd',[Globalization.CultureInfo]::InvariantCulture).AddDays(1)
            if($script:fromDate -ge $script:toDate){throw 'Khoảng ngày không hợp lệ.'}
        }
        $stores=$null; $exchangeCount=0
        try {
            $stores=$namespace.Stores
            for($s=1;$s -le $stores.Count;$s++) {
                $store=$null; $storeRoot=$null; $sent=$null
                try {
                    $store=$stores.Item($s)
                    if(-not (Test-CachedExchangeStore $store)){continue}
                    $exchangeCount++;$script:sentId=''
                    $script:defaultFolderMap=@{}
                    foreach($definition in @(@{Id=6;Name='Inbox'},@{Id=5;Name='Sent Items'},@{Id=3;Name='Deleted Items'},@{Id=16;Name='Drafts'},@{Id=23;Name='Junk Email'},@{Id=4;Name='Outbox'})){
                        $defaultFolder=$null
                        try{$defaultFolder=$store.GetDefaultFolder([int]$definition.Id);$script:defaultFolderMap[[string]$defaultFolder.EntryID]=$definition.Name}catch{}finally{Release-Com $defaultFolder}
                    }
                    try{$sent=$store.GetDefaultFolder(5);$script:sentId=[string]$sent.EntryID}finally{Release-Com $sent}
                    $storeRoot=$store.GetRootFolder()
                    # Keep mailbox identity for selection; PST paths are stored separately.
                    $sourceStoreId=[string]$store.StoreID
                    $hasher=[Security.Cryptography.SHA256]::Create()
                    try{$hash=[BitConverter]::ToString($hasher.ComputeHash($script:Utf8.GetBytes($sourceStoreId))).Replace('-','').Substring(0,8)}finally{$hasher.Dispose()}
                    $sourceLabel=[string]$store.DisplayName+' ['+$hash+']'
                    Read-MailFolder $storeRoot @($sourceLabel) $false $sourceStoreId $request $context
                } finally {Release-Com $storeRoot; Release-Com $store}
                if(Test-Cancelled $context){$context.Status='cancelled';break}
            }
        } finally {Release-Com $stores}
        if($exchangeCount -eq 0){throw 'Không tìm thấy mailbox Exchange có cache cục bộ trong profile Outlook đang mở. Hãy kiểm tra Cached Exchange Mode.'}
        Flush-ScanBatch $context
        $context.Total=$context.Processed
        if($context.Status -eq 'running'){$context.Status=if($context.Errors -gt 0){'completed_with_errors'}else{'completed'}}
        if($context.Errors -eq 0){
            if($context.Status -eq 'cancelled'){$context.Note='Đã dừng tác vụ quét. Danh sách thư đã đọc được giữ lại.'}
            else{$context.Note='Đã quét cache cục bộ của '+$exchangeCount+' mailbox Exchange trong profile Outlook đang mở; bỏ qua PST và hộp thư chỉ online.'}
        }
    } elseif($request.operation -in @('delete','delete_archived')){
        Invoke-ConfirmedDelete $namespace $request $context $archiveReceipts
    } else {
        $pst=[IO.Path]::GetFullPath([string]$request.pst)
        if($pst -notmatch '^[A-Za-z]:\\' -or [IO.Path]::GetExtension($pst) -ine '.pst'){throw 'Chọn đường dẫn .pst trên ổ đĩa cục bộ.'}
        $drive=New-Object IO.DriveInfo([IO.Path]::GetPathRoot($pst))
        if($drive.DriveType -eq [IO.DriveType]::Network){throw 'Không xuất PST trên ổ mạng.'}
        foreach($cloud in @($env:OneDrive,$env:OneDriveCommercial,$env:OneDriveConsumer)) {
            if($cloud -and $pst.StartsWith($cloud.TrimEnd('\')+'\',[StringComparison]::OrdinalIgnoreCase)){throw 'Hãy chọn PST ngoài thư mục OneDrive.'}
        }
        if(-not [IO.Directory]::Exists([IO.Path]::GetDirectoryName($pst))){throw 'Thư mục chứa PST chưa tồn tại.'}
        $exists=[IO.File]::Exists($pst)
        if($request.pstMode -eq 'new' -and $exists){throw 'Tên PST mới đã tồn tại; chọn tên khác.'}
        if($request.pstMode -eq 'existing' -and -not $exists){throw 'PST được chọn chưa tồn tại.'}
        if($request.pstMode -notin @('new','existing')){throw 'Chọn PST mới hoặc có sẵn.'}
        $stores=$namespace.Stores
        try {
            for($i=1;$i -le $stores.Count;$i++) {
                $candidate=$stores.Item($i);$keep=$false
                try {if([string]::Equals([string]$candidate.FilePath,$pst,[StringComparison]::OrdinalIgnoreCase)){$destinationStore=$candidate;$keep=$true;break}}finally{if(-not $keep){Release-Com $candidate}}
            }
        } finally {Release-Com $stores}
        if($null -eq $destinationStore) {
            $namespace.AddStoreEx($pst,2)
            $stores=$namespace.Stores
            try {for($i=1;$i -le $stores.Count;$i++){$candidate=$stores.Item($i);$keep=$false;try{if([string]::Equals([string]$candidate.FilePath,$pst,[StringComparison]::OrdinalIgnoreCase)){$destinationStore=$candidate;$keep=$true;break}}finally{if(-not $keep){Release-Com $candidate}}}}finally{Release-Com $stores}
        }
        if($null -eq $destinationStore -or [int]$destinationStore.ExchangeStoreType -ne 3){throw 'Không tìm thấy PST đích hợp lệ.'}
        $root=$destinationStore.GetRootFolder()
        if($request.pstMode -eq 'new'){$root.Name=[IO.Path]::GetFileNameWithoutExtension($pst)}
        # Resolve every path from the visible PST root. GetDefaultFolder(6)
        # on a fresh PST can return an Inbox outside its visible folder tree.
        $context.ExportMode=[string]$request.pstMode
        if($bulkRequested){Invoke-BulkArchiveExport $namespace $root ([string]$destinationStore.StoreID) $request $context}
        else{Invoke-ArchiveExport $namespace $root ([string]$destinationStore.StoreID) $request $context}
    }
} catch { $context.Status='error';$context.Errors++;$context.Note=$_.Exception.Message }
finally {
    Flush-ScanBatch $context
    Release-Com $root; Release-Com $destinationStore; Release-Com $namespace; Release-Com $outlook
    Save-JobState $context -Force
}
if($context.Status -eq 'error'){exit 1}
