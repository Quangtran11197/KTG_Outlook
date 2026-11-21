# Shared implementation. No Outlook is instantiated by loading this file.
Set-StrictMode -Version 2
$script:Utf8 = New-Object Text.UTF8Encoding($false)
function Release-Com($value) {
    if ($null -ne $value -and [Runtime.InteropServices.Marshal]::IsComObject($value)) {
        [void][Runtime.InteropServices.Marshal]::ReleaseComObject($value)
    }
}
function Test-CachedExchangeStore($Store) {
    # Fail closed: PST, online-only stores and unknown cache state are not scanned.
    try {
        return ([int]$Store.ExchangeStoreType -in @(0,1,2) -and [bool]$Store.IsCachedExchange)
    } catch { return $false }
}
function Write-AtomicJson([string]$Path, $Value) {
    $temp = $Path + '.tmp'
    [IO.File]::WriteAllText($temp, (ConvertTo-Json -InputObject $Value -Depth 14 -Compress), $script:Utf8)
    if ([IO.File]::Exists($Path)) { [IO.File]::Replace($temp, $Path, [NullString]::Value) }
    else { [IO.File]::Move($temp, $Path) }
}
function New-JobContext([string]$Directory, [int]$Total = 0) {
    [IO.Directory]::CreateDirectory($Directory) | Out-Null
    return @{
        Directory=$Directory; Timer=[Diagnostics.Stopwatch]::StartNew(); LastUpdate=0L
        Status='running'; Note=''; Processed=0; Total=$Total; Copied=0; Deleted=0; Errors=0
        AutoCleanup=$false; Phase=''; CleanupProcessed=0; CleanupTotal=0; VerifiedReceipts=@{}; DeferCompletion=$false; AppendReport=$false; NativeCopyMode=$false; NativeCopies=@{}; BulkBatches=0; Paused=$false
        LogPath=$(if($env:KTG_LOG_FILE){$env:KTG_LOG_FILE}else{Join-Path $Directory 'operations.log'}); LastLog=''
        Batches=0; ResultBatches=0; FolderCacheHits=0; FolderResolutions=0
        Batch=(New-Object 'Collections.Generic.List[object]')
        Results=(New-Object 'Collections.Generic.List[object]')
    }
}
function Save-JobState($Context, [switch]$Force) {
    $elapsed=$Context.Timer.ElapsedMilliseconds
    if (-not $Force -and $elapsed-$Context.LastUpdate -lt 500) { return }
    $Context.LastUpdate=$elapsed
    Write-AtomicJson (Join-Path $Context.Directory 'state.json') @{
        status=$Context.Status; note=$Context.Note; processed=$Context.Processed
        total=$Context.Total; copied=$Context.Copied; deleted=$Context.Deleted; errors=$Context.Errors
        elapsedSeconds=[math]::Round($elapsed/1000,2); batches=$Context.Batches
        resultBatches=$Context.ResultBatches; folderCacheHits=$Context.FolderCacheHits
        folderResolutions=$Context.FolderResolutions
        report=(Join-Path $Context.Directory 'report.csv')
        logPath=$Context.LogPath
        paused=$Context.Paused;bulkBatches=$Context.BulkBatches;autoDeleteAfterExport=$Context.AutoCleanup;phase=$Context.Phase;cleanupProcessed=$Context.CleanupProcessed;cleanupTotal=$Context.CleanupTotal
        jobId=[IO.Path]::GetFileName($Context.Directory)
    }
    $signature=$Context.Status+'|'+$Context.Note
    if($Context.LastLog -ne $signature){
        $line=([datetime]::Now.ToString('yyyy-MM-dd HH:mm:ss')+' | '+$Context.Status+' | '+$Context.Note+[Environment]::NewLine)
        try{[IO.File]::AppendAllText($Context.LogPath,$line,$script:Utf8)}
        catch{$Context.LogPath=Join-Path $Context.Directory 'operations.log';[IO.File]::AppendAllText($Context.LogPath,$line,$script:Utf8)}
        $Context.LastLog=$signature
    }
}
function Flush-ScanBatch($Context) {
    if ($Context.Batch.Count -eq 0) { return }
    $index=$Context.Batches
    Write-AtomicJson (Join-Path $Context.Directory ('messages-'+$index+'.json')) @($Context.Batch.ToArray())
    $Context.Batches++; $Context.Batch.Clear()
    Save-JobState $Context -Force
}
function Flush-ResultBatch($Context, $Writer) {
    if ($Context.Results.Count -eq 0) { return }
    $index=$Context.ResultBatches
    Write-AtomicJson (Join-Path $Context.Directory ('results-'+$index+'.json')) @($Context.Results.ToArray())
    $Context.ResultBatches++; $Context.Results.Clear()
    if ($null -ne $Writer) { $Writer.Flush() }
    Save-JobState $Context -Force
}
function Test-Cancelled($Context) {
    return [IO.File]::Exists((Join-Path $Context.Directory 'cancel'))
}
function Get-Identity($Message) { return ([string]$Message.storeId+'|'+[string]$Message.id) }
function Get-AutoCleanupEnabled($Request) {
    $flag=$Request.PSObject.Properties['autoDeleteAfterExport']
    if(-not $flag){return $false}
    if($flag.Value -isnot [bool]){throw 'Chế độ tự xóa phải là giá trị bật/tắt hợp lệ.'}
    if(-not $flag.Value){return $false}
    $confirmed=$Request.PSObject.Properties['autoDeleteConfirmed']
    $count=$Request.PSObject.Properties['autoDeleteCount']
    if($Request.operation -ne 'export' -or -not $confirmed -or $confirmed.Value -isnot [bool] -or -not $confirmed.Value -or
       -not $count -or ($count.Value -isnot [int] -and $count.Value -isnot [long])){throw 'Thiếu xác nhận tự xóa trước khi bắt đầu lưu PST.'}
    $selected=@(Select-UniqueMessages @($Request.messages))
    if($selected.Count -eq 0 -or [long]$count.Value -ne $selected.Count){throw 'Số email cho chế độ tự xóa không khớp.'}
    foreach($m in $selected){if(-not $m.PSObject.Properties['sourceFolderId'] -or [string]::IsNullOrWhiteSpace([string]$m.sourceFolderId)){throw 'Hãy quét lại trước khi bật tự xóa.'}}
    return $true
}
function Test-SameOutlookId($Namespace,[string]$First,[string]$Second) {
    if([string]::IsNullOrWhiteSpace($First) -or [string]::IsNullOrWhiteSpace($Second)){return $false}
    if([string]::Equals($First,$Second,[StringComparison]::OrdinalIgnoreCase)){return $true}
    try{return [bool]$Namespace.CompareEntryIDs($First,$Second)}catch{return $false}
}
function Get-PstFolderSegments($Message) {
    if($Message.PSObject.Properties['pstFolderSegments']){
        $segments=@($Message.pstFolderSegments)
    }else{
        # Compatibility with older scans, whose first segment is the mailbox label.
        $segments=@($Message.folderSegments | Select-Object -Skip 1)
        $aliases=@{'Hộp thư'='Inbox';'Hộp thư đến'='Inbox';'Các mục đã gửi'='Sent Items';'Các mục đã Xóa'='Deleted Items';'Bản thảo'='Drafts';'Email rác'='Junk Email';'Hộp thư bên ngoài'='Outbox'}
        if($segments.Count -gt 0 -and $aliases.ContainsKey([string]$segments[0])){$segments[0]=$aliases[[string]$segments[0]]}
    }
    if($segments.Count -eq 0){return @('Inbox')}
    return $segments
}
function Get-ConfirmedDeleteMessages($Request) {
    $confirmation=$Request.PSObject.Properties['confirmPermanentDelete']
    $phrase=$Request.PSObject.Properties['confirmationText']
    $messages=$Request.PSObject.Properties['messages']
    if($null -eq $confirmation -or $confirmation.Value -isnot [bool] -or -not $confirmation.Value -or
       $null -eq $phrase -or $null -eq $messages){throw 'Xóa cần xác nhận riêng: nhập XOA và số email đã chọn.'}
    $selected=@(Select-UniqueMessages @($messages.Value))
    if($selected.Count -eq 0 -or [string]$phrase.Value -cne ('XOA '+$selected.Count)){
        throw 'Xác nhận xóa không khớp số email đã chọn.'
    }
    foreach($m in $selected){
        if(-not $m.PSObject.Properties['sourceFolderId'] -or [string]::IsNullOrWhiteSpace([string]$m.sourceFolderId)){
            throw 'Thiếu định danh thư mục nguồn; hãy quét lại trước khi xóa.'
        }
    }
    return $selected
}
function Select-UniqueMessages($Messages) {
    $seen=New-Object 'Collections.Generic.HashSet[string]' ([StringComparer]::OrdinalIgnoreCase)
    foreach($message in $Messages) {
        if ([string]::IsNullOrWhiteSpace([string]$message.id) -or [string]::IsNullOrWhiteSpace([string]$message.storeId)) {
            throw 'Mail thiếu EntryID hoặc StoreID; hãy quét lại.'
        }
        if ($seen.Add((Get-Identity $message))) { $message }
    }
}
function Get-DestinationFolder($Root, [string[]]$Segments, $Cache, $Context) {
    # JSON encoding preserves segment boundaries, even when folder names contain '/'.
    if(-not ('KtgExportIo.Fast' -as [type])){Initialize-ExportIo}
    $key=[KtgExportIo.Fast]::PathKey($Segments)
    if ($Cache.ContainsKey($key)) { $Context.FolderCacheHits++; return $Cache[$key] }
    $current=$Root; $path=New-Object 'Collections.Generic.List[string]'
    foreach($segment in $Segments) {
        if ([string]::IsNullOrWhiteSpace($segment) -or $segment.Length -gt 255) { throw 'Tên thư mục PST không hợp lệ.' }
        $path.Add($segment)
        $prefix=[KtgExportIo.Fast]::PathKey($path.ToArray())
        if ($Cache.ContainsKey($prefix)) { $current=$Cache[$prefix].Folder; continue }
        $folders=$null; $next=$null
        try {
            $folders=$current.Folders
            try { $next=$folders.Item($segment) } catch { $next=$folders.Add($segment,6) }
            if ($null -eq $next) { throw 'Không mở được thư mục PST.' }
            $Cache[$prefix]=@{Folder=$next; EntryId=[string]$next.EntryID}
            $current=$next; $Context.FolderResolutions++
        } finally { Release-Com $folders }
    }
    if ($Segments.Count -eq 0) { throw 'Thiếu đường dẫn thư mục đích.' }
    return $Cache[$key]
}
function Write-CsvRecord($Writer, $Values) {
    if(-not ('KtgExportIo.Fast' -as [type])){Initialize-ExportIo}
    [KtgExportIo.Fast]::Csv($Writer,[object[]]$Values)
}
function Invoke-CopyExport($Namespace, $Root, [string]$DestinationStoreId, $Messages, $Context, $DefaultFolders=@{}) {
    $cache=@{}; $writer=$null
    # Never seed this cache with Store.GetDefaultFolder: a new PST may expose
    # a default folder outside the visible tree. Resolve from Root.Folders.
    try {
        $rootId=[string]$Root.EntryID
        $selected=@(Select-UniqueMessages $Messages)
        if(-not $Context.DeferCompletion){$Context.Total=$selected.Count}
        $writer=New-Object IO.StreamWriter((Join-Path $Context.Directory 'report.csv'),[bool]$Context.AppendReport,(New-Object Text.UTF8Encoding($true)))
        if(-not $Context.AppendReport){Write-CsvRecord $writer @('sourceStoreId','sourceEntryId','destinationStoreId','destinationEntryId','status','error','temporaryCopyEntryId','sourceCurrentEntryId')}
        Save-JobState $Context -Force
        foreach($m in $selected) {
            if (Test-Cancelled $Context) { $Context.Status='cancelled'; break }
            $source=$null; $copy=$null; $moved=$null; $parent=$null; $attachments=$null
            $destinationId=''; $temporaryId=''; $errorText=''; $resultStatus='failed'
            try {
                if ([string]::Equals([string]$m.storeId,$DestinationStoreId,[StringComparison]::OrdinalIgnoreCase)) { throw 'Nguồn trùng PST đích.' }
                $source=$Namespace.GetItemFromID([string]$m.id,[string]$m.storeId)
                if ($null -eq $source -or [int]$source.Class -ne 43) { throw 'Thư không còn tồn tại hoặc không phải MailItem.' }
                $sourceId=[string]$source.EntryID
                if (-not (Test-SameOutlookId $Namespace $sourceId ([string]$m.id))) { throw 'Outlook trả về thư khác; hãy quét lại.' }
                if ([int]$source.DownloadState -ne 1) { throw 'Thư chưa tải đầy đủ vào cache; bỏ qua.' }
                $subject=[string]$source.Subject
                $size=[long]$source.Size
                $sourceModifiedTicks=0L;try{$sourceModifiedTicks=([datetime]$source.LastModificationTime).Ticks}catch{}
                $attachments=$source.Attachments; $attachmentCount=[int]$attachments.Count
                Release-Com $attachments; $attachments=$null
                $sourceParent=$null
                try{
                    $sourceParent=$source.Parent
                    if([string]$sourceParent.StoreID -ine [string]$m.storeId){throw 'Mailbox nguồn không khớp; hãy quét lại.'}
                    if($m.PSObject.Properties['sourceFolderId'] -and -not (Test-SameOutlookId $Namespace ([string]$sourceParent.EntryID) ([string]$m.sourceFolderId))){throw 'Thư đã đổi thư mục nguồn; hãy quét lại.'}
                    $sourceSpec=@{id=$sourceId;storeId=[string]$m.storeId;sourceFolderId=[string]$sourceParent.EntryID;subject=$subject;internetMessageId=[string]$m.internetMessageId}
                }finally{Release-Com $sourceParent}
                # Invalid sources never cause empty destination folders to be created.
                $pstSegments=@(Get-PstFolderSegments $m)
                $destination=Get-DestinationFolder $Root $pstSegments $cache $Context
                # Only the new temporary copy is moved. The selected source is never moved/deleted.
                if($Context.NativeCopyMode){
                    $key=Get-Identity $m
                    if(-not $Context.NativeCopies.ContainsKey($key)){throw 'Lô MAPI không có bản sao kiểm tra được; giữ nguồn. Không tự thử lại.'}
                    $native=$Context.NativeCopies[$key]
                    if($native.error){throw [string]$native.error}
                    $nativeTicks=[datetime]::FromFileTimeUtc([long]$native.sourceModifiedFileTime).ToLocalTime().Ticks
                    if([math]::Abs($nativeTicks - $sourceModifiedTicks) -gt 10000 -or [long]$native.sourceSize -ne $size -or [int]$native.sourceAttachments -ne $attachmentCount){throw 'Nguồn thay đổi trong lúc sao chép lô; giữ nguồn.'}
                    $moved=$Namespace.GetItemFromID([string]$native.destinationId,$DestinationStoreId)
                }else{
                    $copy=$source.Copy()
                    if($null -eq $copy){throw 'Outlook không trả về bản sao.'}
                    $temporaryId=[string]$copy.EntryID
                    $moved=$copy.Move($destination.Folder)
                }
                if ($null -eq $moved) { throw 'Không nhận được bản sao trong PST.' }
                $destinationId=[string]$moved.EntryID
                $parent=$moved.Parent
                if ([int]$moved.Class -ne 43 -or
                    -not [string]::Equals([string]$parent.StoreID,$DestinationStoreId,[StringComparison]::OrdinalIgnoreCase) -or
                    -not (Test-SameOutlookId $Namespace ([string]$parent.EntryID) $destination.EntryId)) {
                    throw 'Bản sao không nằm trong thư mục PST dự kiến.'
                }
                if ([string]$moved.Subject -cne $subject) { throw 'Tiêu đề bản sao không khớp.' }
                $attachments=$moved.Attachments
                if ([int]$attachments.Count -ne $attachmentCount) { throw 'Số tệp đính kèm không khớp.' }
                if ($size -gt 0 -and [long]$moved.Size -le 0) { throw 'Dung lượng bản sao không hợp lệ.' }
                if (-not [string]::IsNullOrWhiteSpace([string]$m.internetMessageId)) {
                    $accessor=$null
                    try {
                        $accessor=$moved.PropertyAccessor
                        $internetId=[string]$accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F')
                        if ($internetId -cne [string]$m.internetMessageId) { throw 'Internet Message ID không khớp.' }
                    } finally { Release-Com $accessor }
                }
                $resultStatus='verified'; $Context.Copied++
            } catch { $Context.Errors++; $errorText=$_.Exception.Message }
            finally {
                Release-Com $attachments; Release-Com $parent
                Release-Com $moved; Release-Com $copy; Release-Com $source
            }
            $Context.Processed++
            $proof=$null
            if($resultStatus -eq 'verified'){
                $proof=@{sourceMessage=$sourceSpec;sourceModifiedTicks=$sourceModifiedTicks;destinationStoreId=$DestinationStoreId;destinationId=$destinationId;destinationFolderId=$destination.EntryId;destinationRootId=$rootId;destinationSegments=$pstSegments;subject=$subject;size=$size;attachmentCount=$attachmentCount;internetMessageId=[string]$m.internetMessageId;pstMode=$(if($Context.ContainsKey('ExportMode')){$Context.ExportMode}else{'new'})}
            }
            if($proof -and $Context.NativeCopyMode){$proof['sourceModifiedFileTime']=[string]$native.sourceModifiedFileTime}
            if($proof -and $Context.AutoCleanup){$Context.VerifiedReceipts[(Get-Identity $m)]=[pscustomobject]$proof}
            $Context.Results.Add(@{key=(Get-Identity $m);status=$resultStatus;destinationId=$destinationId;error=$errorText;archiveProof=$proof})
            Write-CsvRecord $writer @($m.storeId,$m.id,$DestinationStoreId,$destinationId,$resultStatus,$errorText,$temporaryId,'')
            if ($Context.Results.Count -ge 100 -or $Context.Timer.ElapsedMilliseconds-$Context.LastUpdate -ge 500) {
                Flush-ResultBatch $Context $writer
            }
            Save-JobState $Context
        }
        if ($Context.Status -eq 'running' -and -not $Context.AutoCleanup -and -not $Context.DeferCompletion) { $Context.Status=if($Context.Errors -gt 0){'completed_with_errors'}else{'completed'} }
        $Context.Note='Đã sao chép và kiểm tra '+$Context.Copied+'/'+$Context.Total+' email vào PST. Lỗi: '+$Context.Errors+'. '+$(if($Context.AutoCleanup){'Chuẩn bị xử lý tự xóa các bản đã kiểm tra.'}else{'Thư gốc được giữ nguyên.'})
    } finally {
        Flush-ResultBatch $Context $writer
        if ($null -ne $writer) { $writer.Dispose() }
        $Context.AppendReport=$true
        foreach($entry in $cache.Values) { Release-Com $entry.Folder }
        Save-JobState $Context -Force
    }
}
function Get-ArchiveReceipts($Request,[string]$CurrentDirectory) {
    $property=$Request.PSObject.Properties['archiveJobId']
    if(-not $property -or [string]$property.Value -notmatch '^job-[0-9]+-[0-9]+$'){throw 'Thiếu phiên xuất PST hợp lệ.'}
    $directory=Join-Path ([IO.Path]::GetDirectoryName($CurrentDirectory)) ([string]$property.Value)
    if(-not [IO.Directory]::Exists($directory) -or $directory -eq $CurrentDirectory){throw 'Không tìm thấy phiên xuất PST để đối chiếu.'}
    return (Read-VerifiedPstReceipts $directory)
}
function Read-VerifiedPstReceipts([string]$directory) {
    $receipts=@{}
    foreach($file in Get-ChildItem -LiteralPath $directory -Filter 'results-*.json'){
        $batch=Get-Content -LiteralPath $file.FullName -Raw -Encoding UTF8|ConvertFrom-Json
        foreach($result in $batch){
            if($result.status -eq 'verified' -and $result.PSObject.Properties['archiveProof'] -and $result.archiveProof){$receipts[[string]$result.key]=$result.archiveProof}
        }
    }
    return $receipts
}
function Assert-ArchivedCopy($Namespace,$Proof,[switch]$AllowNewPst,[hashtable]$PathCache=$null) {
    if($Proof.pstMode -ne 'existing' -and -not ($AllowNewPst -and $Proof.pstMode -eq 'new')){throw 'Phiên lưu PST không hợp lệ cho thao tác này.'}
    $item=$null;$parent=$null;$attachments=$null;$accessor=$null
    try{
        $pathKey='';$checkPath=$true
        if($null -ne $PathCache -and $Proof.PSObject.Properties['destinationRootId']){
            $pathKey=([string]$Proof.destinationStoreId)+'|'+([string]$Proof.destinationFolderId)+'|'+([string]$Proof.destinationRootId)+'|'+(@($Proof.destinationSegments) -join [char]0)
            if($PathCache.ContainsKey($pathKey)){
                $entry=$PathCache[$pathKey]
                if(([Environment]::TickCount - [long]$entry.tick) -ge 0 -and ([Environment]::TickCount - [long]$entry.tick) -lt 500 -and $entry.uses -lt 100){$checkPath=$false;$entry.uses++}
            }
        }
        if($checkPath -and $Proof.PSObject.Properties['destinationRootId']){
            # A valid EntryID alone can point at a hidden/orphan PST folder.
            # Resolve the recorded path from the actual visible store root too.
            $store=$null;$visible=$null;$folders=$null
            try{
                $store=$Namespace.GetStoreFromID([string]$Proof.destinationStoreId)
                $visible=$store.GetRootFolder()
                if([string]$visible.StoreID -ine [string]$Proof.destinationStoreId -or -not (Test-SameOutlookId $Namespace ([string]$visible.EntryID) ([string]$Proof.destinationRootId))){throw 'Gốc PST đã thay đổi; giữ thư nguồn.'}
                foreach($segment in @($Proof.destinationSegments)){
                    $folders=$visible.Folders
                    $next=$folders.Item([string]$segment)
                    Release-Com $folders;$folders=$null
                    Release-Com $visible;$visible=$next
                }
                if(-not (Test-SameOutlookId $Namespace ([string]$visible.EntryID) ([string]$Proof.destinationFolderId))){throw 'Bản PST không nằm trong đường dẫn hiển thị; giữ thư nguồn.'}
                if($null -ne $PathCache){$PathCache[$pathKey]=@{tick=[Environment]::TickCount;uses=1}}
            }finally{Release-Com $folders;Release-Com $visible;Release-Com $store}
        }
        $item=$Namespace.GetItemFromID([string]$Proof.destinationId,[string]$Proof.destinationStoreId)
        if($null -eq $item -or [int]$item.Class -ne 43 -or -not (Test-SameOutlookId $Namespace ([string]$item.EntryID) ([string]$Proof.destinationId))){throw 'Không mở lại được bản sao PST; giữ thư nguồn.'}
        $parent=$item.Parent
        if([string]$parent.StoreID -ine [string]$Proof.destinationStoreId -or -not (Test-SameOutlookId $Namespace ([string]$parent.EntryID) ([string]$Proof.destinationFolderId))){throw 'Bản sao PST đã đổi thư mục; giữ thư nguồn.'}
        $attachments=$item.Attachments
        if([string]$item.Subject -cne [string]$Proof.subject -or [int]$attachments.Count -ne [int]$Proof.attachmentCount -or ([long]$Proof.size -gt 0 -and [long]$item.Size -le 0)){throw 'Bản sao PST không còn khớp; giữ thư nguồn.'}
        if(-not [string]::IsNullOrWhiteSpace([string]$Proof.internetMessageId)){
            $accessor=$item.PropertyAccessor
            if([string]$accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F') -cne [string]$Proof.internetMessageId){throw 'Internet Message ID trong PST không khớp; giữ thư nguồn.'}
        }
    }finally{Release-Com $accessor;Release-Com $attachments;Release-Com $parent;Release-Com $item}
}
function Assert-SourceMatchesArchive($Source,$Proof) {
    $attachments=$null
    try{
        $attachments=$Source.Attachments
        if([long]$Source.Size -ne [long]$Proof.size -or [int]$attachments.Count -ne [int]$Proof.attachmentCount){throw 'Nội dung hoặc tệp đính kèm nguồn đã đổi sau khi lưu PST; giữ thư nguồn.'}
        if($Proof.PSObject.Properties['sourceModifiedTicks'] -and [long]$Proof.sourceModifiedTicks -gt 0 -and
           ([datetime]$Source.LastModificationTime).Ticks -ne [long]$Proof.sourceModifiedTicks){throw 'Thư nguồn đã được sửa sau khi lưu PST; giữ thư nguồn.'}
    }finally{Release-Com $attachments}
}
function Invoke-ConfirmedDelete($Namespace, $Request, $Context, $ArchiveReceipts=$null, [switch]$Automatic) {
    # This capability is separate from export and re-validates confirmation at its entry point.
    $selected=@(Get-ConfirmedDeleteMessages $Request)
    if($Automatic){$Context.CleanupTotal=if($Context.DeferCompletion){$Context.Total}else{$selected.Count}}else{if(-not $Context.DeferCompletion){$Context.Total=$selected.Count}}
    $cache=@{};$pathCache=@{};$writer=$null
    try {
        $writer=New-Object IO.StreamWriter((Join-Path $Context.Directory 'report.csv'),[bool]$Automatic,(New-Object Text.UTF8Encoding($true)))
        if(-not $Automatic){Write-CsvRecord $writer @('sourceStoreId','sourceEntryId','currentEntryId','status','error')}
        $writeDeleteRecord={
            param($storeId,$sourceId,$currentId,$status,$errorText)
            if($Automatic){Write-CsvRecord $writer @($storeId,$sourceId,'','',$status,$errorText,'',$currentId)}
            else{Write-CsvRecord $writer @($storeId,$sourceId,$currentId,$status,$errorText)}
            # Intent and post-Move ID must reach the audit before mutation.
            # Final success records can share the normal result-batch flush.
            if($status -in @('delete_intent','in_deleted_items','failed')){$writer.Flush()}
        }
        Save-JobState $Context -Force
        foreach($m in $selected){
            if($Context.NativeCopyMode -and -not (Wait-BulkGate $Context)){break}
            if(Test-Cancelled $Context){$Context.Status='cancelled';break}
            $source=$null;$moved=$null;$reopened=$null;$probe=$null;$parent=$null;$accessor=$null;$currentId=[string]$m.id
            $status='failed';$errorText=''
            $resultKey=Get-Identity $m
            try{
                if($null -ne $ArchiveReceipts){
                    if(-not $ArchiveReceipts.ContainsKey($resultKey)){throw 'Email chưa có bản PST được kiểm tra; giữ thư nguồn.'}
                    $proof=$ArchiveReceipts[$resultKey]
                    Assert-ArchivedCopy $Namespace $proof -AllowNewPst:$Automatic -PathCache $pathCache
                    $m=[pscustomobject]$proof.sourceMessage
                    $currentId=[string]$m.id
                }
                $storeId=[string]$m.storeId
                if(-not $cache.ContainsKey($storeId)){
                    $store=$null;$trash=$null;$keep=$false
                    try{
                        $store=$Namespace.GetStoreFromID($storeId)
                        if(-not (Test-CachedExchangeStore $store)){throw 'Chỉ cho phép xóa trong mailbox Exchange có cache cục bộ.'}
                        $trash=$store.GetDefaultFolder(3)
                        $cache[$storeId]=@{Store=$store;Folder=$trash;EntryId=[string]$trash.EntryID};$keep=$true
                    }finally{if(-not $keep){Release-Com $trash;Release-Com $store}}
                }
                $destination=$cache[$storeId]
                $source=$Namespace.GetItemFromID([string]$m.id,$storeId)
                if($null -eq $source -or [int]$source.Class -ne 43 -or
                   -not (Test-SameOutlookId $Namespace ([string]$source.EntryID) ([string]$m.id))){
                    throw 'Thư nguồn đã thay đổi; hãy quét lại.'
                }
                $parent=$source.Parent
                if(-not [string]::Equals([string]$parent.StoreID,$storeId,[StringComparison]::OrdinalIgnoreCase) -or
                   -not (Test-SameOutlookId $Namespace ([string]$parent.EntryID) ([string]$m.sourceFolderId)) -or
                   [string]$source.Subject -cne [string]$m.subject){throw 'Mailbox, thư mục hoặc tiêu đề đã thay đổi; từ chối xóa.'}
                if(-not [string]::IsNullOrWhiteSpace([string]$m.internetMessageId)){
                    $accessor=$source.PropertyAccessor
                    if([string]$accessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F') -cne [string]$m.internetMessageId){throw 'Internet Message ID không khớp; từ chối xóa.'}
                }
                if($null -ne $ArchiveReceipts){Assert-SourceMatchesArchive $source $proof}
                if(Test-Cancelled $Context){$Context.Status='cancelled';break}
                # Persist intent before any original is moved or deleted.
                & $writeDeleteRecord $storeId $m.id $currentId 'delete_intent' ''
                $target=$source
                if(-not (Test-SameOutlookId $Namespace ([string]$parent.EntryID) $destination.EntryId)){
                    Release-Com $parent;$parent=$null
                    $moved=$source.Move($destination.Folder)
                    if($null -eq $moved -or [int]$moved.Class -ne 43){throw 'Không nhận được thư sau khi chuyển vào Deleted Items.'}
                    $target=$moved;$currentId=[string]$target.EntryID
                    $parent=$target.Parent
                }
                if(-not (Test-SameOutlookId $Namespace ([string]$parent.EntryID) $destination.EntryId) -or
                   -not [string]::Equals([string]$parent.StoreID,$storeId,[StringComparison]::OrdinalIgnoreCase)){
                    throw 'Thư không ở Deleted Items dự kiến; từ chối yêu cầu xóa vĩnh viễn.'
                }
                & $writeDeleteRecord $storeId $m.id $currentId 'in_deleted_items' ''
                if(Test-Cancelled $Context){$Context.Status='cancelled';throw 'Đã dừng; thư hiện ở Deleted Items, chưa yêu cầu xóa vĩnh viễn. Hãy quét lại.'}
                # Do not call Delete on Move's returned wrapper: Exchange may retain
                # its original folder context and only soft-delete it a second time.
                Release-Com $parent;$parent=$null
                Release-Com $moved;$moved=$null
                Release-Com $source;$source=$null;$target=$null
                $reopened=$Namespace.GetItemFromID($currentId,$storeId)
                $parent=$reopened.Parent
                if([int]$reopened.Class -ne 43 -or
                   -not (Test-SameOutlookId $Namespace ([string]$reopened.EntryID) $currentId) -or
                   -not (Test-SameOutlookId $Namespace ([string]$parent.EntryID) $destination.EntryId) -or
                   [string]$parent.StoreID -ine $storeId -or [string]$reopened.Subject -cne [string]$m.subject){throw 'Không mở lại được đúng thư trong Deleted Items; chưa xóa vĩnh viễn.'}
                if($null -ne $ArchiveReceipts){
                    $attachments=$null;$reopenedAccessor=$null
                    try{
                        # Moving a message may change its MAPI size/modified time.
                        # Exact source snapshot was checked before Move above.
                        $attachments=$reopened.Attachments
                        if([int]$attachments.Count -ne [int]$proof.attachmentCount -or [long]$reopened.Size -le 0){throw 'Thư mở lại không khớp bản PST; chưa xóa vĩnh viễn.'}
                        if(-not [string]::IsNullOrWhiteSpace([string]$proof.internetMessageId)){
                            $reopenedAccessor=$reopened.PropertyAccessor
                            if([string]$reopenedAccessor.GetProperty('http://schemas.microsoft.com/mapi/proptag/0x1035001F') -cne [string]$proof.internetMessageId){throw 'Internet ID sau khi chuyển không khớp; chưa xóa vĩnh viễn.'}
                        }
                    }finally{Release-Com $reopenedAccessor;Release-Com $attachments}
                }
                if(Test-Cancelled $Context){$Context.Status='cancelled';throw 'Đã dừng trước khi xóa vĩnh viễn.'}
                $reopened.Delete()
                Release-Com $parent;$parent=$null
                Release-Com $reopened;$reopened=$null
                # A successful COM return alone is not proof of removal. Unknown
                # lookup errors must remain failures, never be treated as not-found.
                $absent=$false
                try{$probe=$Namespace.GetItemFromID($currentId,$storeId)}catch{
                    $e=$_.Exception
                    while($e.InnerException){$e=$e.InnerException}
                    if($e.HResult -eq -2147221233 -or $e.HResult -eq -2147221238){$absent=$true}else{throw}
                }
                if(-not $absent){throw 'Outlook nhận lệnh nhưng thư vẫn truy cập được trong mailbox; chưa xác nhận xóa khỏi Deleted Items. Không tự thử lại.'}
                $status='delete_requested';$Context.Deleted++
            }catch{$Context.Errors++;$errorText=$_.Exception.Message}
            finally{Release-Com $accessor;Release-Com $parent;Release-Com $probe;Release-Com $reopened;Release-Com $moved;Release-Com $source}
            if($Automatic){$Context.CleanupProcessed++}else{$Context.Processed++}
            & $writeDeleteRecord $m.storeId $m.id $currentId $status $errorText
            $Context.Results.Add(@{key=$resultKey;status=$status;error=$errorText;currentId=$currentId;needsRescan=($status -ne 'delete_requested')})
            if($Context.Results.Count -ge 100 -or $Context.Timer.ElapsedMilliseconds-$Context.LastUpdate -ge 500){Flush-ResultBatch $Context $writer}
            Save-JobState $Context
        }
        if($Context.Status -eq 'running' -and -not $Context.DeferCompletion){$Context.Status=if($Context.Errors -gt 0){'completed_with_errors'}else{'completed'}}
        $Context.Note=$(if($Automatic){'PST đã kiểm tra: '+$Context.Copied+'/'+$Context.Total+'. Đã tự động gửi '+$Context.Deleted+' yêu cầu xóa nguồn và Deleted Items.'}else{'Đã gửi '+$Context.Deleted+' yêu cầu xóa vĩnh viễn sau xác nhận riêng.'})+' Lỗi: '+$Context.Errors+'. Không tự thử lại các mục lỗi; đối chiếu báo cáo. Chính sách lưu giữ Exchange có thể vẫn áp dụng.'
    }finally{
        Flush-ResultBatch $Context $writer
        if($writer){$writer.Dispose()}
        foreach($item in $cache.Values){Release-Com $item.Folder;Release-Com $item.Store}
        Save-JobState $Context -Force
    }
}
function Invoke-ArchiveExport($Namespace,$Root,[string]$DestinationStoreId,$Request,$Context,$DefaultFolders=@{}) {
    $Context.AutoCleanup=Get-AutoCleanupEnabled $Request
    $Context.Phase='copy'
    Invoke-CopyExport $Namespace $Root $DestinationStoreId @($Request.messages) $Context $DefaultFolders
    if(-not $Context.AutoCleanup){return}
    if($Context.Status -eq 'cancelled' -or (Test-Cancelled $Context)){
        $Context.Status='cancelled';$Context.Note='Đã dừng tác vụ; các bản PST đã tạo được giữ lại. Không bắt đầu tự xóa nguồn.';Save-JobState $Context -Force;return
    }
    # These receipts are created only by this worker and flushed to disk before
    # cleanup. Avoid parsing every result JSON again for a large automatic job.
    $receipts=$Context.VerifiedReceipts
    $pending=@(Select-UniqueMessages @($Request.messages)|Where-Object{$receipts.ContainsKey((Get-Identity $_))})
    if($pending.Count -eq 0){
        $Context.Status=if($Context.Errors -gt 0){'completed_with_errors'}else{'completed'}
        $Context.Note='Không có bản PST kiểm tra thành công; không xóa thư nguồn.';Save-JobState $Context -Force;return
    }
    $Context.Phase='cleanup';$Context.Note='Đang mở lại bản PST và tự xóa nguồn đã lưu thành công; không hỏi lại.';Save-JobState $Context -Force
    $internalRequest=[pscustomobject]@{operation='delete_archived';messages=$pending;confirmPermanentDelete=$true;confirmationText=('XOA '+$pending.Count)}
    Invoke-ConfirmedDelete $Namespace $internalRequest $Context $receipts -Automatic
}

function Initialize-FastScan {
    if('KtgFastScan.Encoder' -as [type]){return}
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.Globalization;
namespace KtgFastScan {
 public sealed class Result {public int Count;public bool Cancelled;}
 public static class Encoder {
  static void Quote(StringBuilder b,string s){
   b.Append('"');foreach(char c in s??""){switch(c){case '"':b.Append("\\\"");break;case '\\':b.Append("\\\\");break;case '\n':b.Append("\\n");break;case '\r':b.Append("\\r");break;case '\t':b.Append("\\t");break;default:if(c<32)b.Append("\\u"+((int)c).ToString("x4"));else b.Append(c);break;}}b.Append('"');
  }
  static string Text(object[,] rows,int r,int c){return c<0||rows[r,c]==null||rows[r,c] is DBNull?"":Convert.ToString(rows[r,c],CultureInfo.InvariantCulture);}
  static void Field(StringBuilder b,string name,string value){Quote(b,name);b.Append(':');Quote(b,value);}
  static string ArrayJson(string[] a){var b=new StringBuilder("[");for(int i=0;i<a.Length;i++){if(i>0)b.Append(',');Quote(b,a[i]);}return b.Append(']').ToString();}
  public static Result Write(object[,] rows,int[] columns,string store,string folderId,string[] segments,string[] pstSegments,bool all,DateTime from,DateTime to,string path,string cancel){
   var result=new Result();var b=new StringBuilder(Math.Max(1024,rows.GetLength(0)*800));b.Append('[');
   var common=new StringBuilder();common.Append(',');Field(common,"storeId",store);common.Append(',');Field(common,"sourceFolderId",folderId);common.Append(',');Field(common,"folder",String.Join(" / ",segments));common.Append(",\"folderSegments\":").Append(ArrayJson(segments)).Append(",\"pstFolderSegments\":").Append(ArrayJson(pstSegments));string suffix=common.ToString();
   for(int r=0;r<rows.GetLength(0);r++){
    if(r%100==0&&File.Exists(cancel)){result.Cancelled=true;break;}
    if(!Text(rows,r,columns[2]).StartsWith("IPM.Note",StringComparison.OrdinalIgnoreCase))continue;
    object d=rows[r,columns[3]];if(d==null||d is DBNull)continue;
    DateTime date=d is DateTime?(DateTime)d:DateTime.FromOADate(Convert.ToDouble(d,CultureInfo.InvariantCulture));if(!all&&(date<from||date>=to))continue;
    string sender=Text(rows,r,columns[6]);if(String.IsNullOrWhiteSpace(sender))sender=Text(rows,r,columns[5]);
    if(result.Count>0)b.Append(',');b.Append('{');Field(b,"id",Text(rows,r,columns[0]));b.Append(',');Field(b,"subject",Text(rows,r,columns[1]));b.Append(',');Field(b,"sender",sender);b.Append(',');Field(b,"internetMessageId",Text(rows,r,columns[7]));b.Append(',');Field(b,"date",date.ToString("yyyy-MM-dd HH:mm",CultureInfo.InvariantCulture));
    long size=columns[4]<0||rows[r,columns[4]]==null||rows[r,columns[4]] is DBNull?0:Convert.ToInt64(rows[r,columns[4]],CultureInfo.InvariantCulture);
    b.Append(",\"size\":").Append(size.ToString(CultureInfo.InvariantCulture)).Append(suffix).Append('}');result.Count++;
   }
   if(result.Count>0){b.Append(']');File.WriteAllText(path+".tmp",b.ToString(),new UTF8Encoding(false));File.Move(path+".tmp",path);}
   return result;
  }
 }
}
'@
}

function Initialize-ExportIo {
    if('KtgExportIo.Fast' -as [type]){return}
    Add-Type -TypeDefinition @'
using System;
using System.IO;
using System.Text;
using System.Globalization;
namespace KtgExportIo {
 public static class Fast {
  public static string PathKey(string[] segments){
   var b=new StringBuilder();foreach(string s in segments){if(s==null)b.Append("-1:");else b.Append(s.Length).Append(':').Append(s);}return b.ToString();
  }
  public static void Csv(TextWriter writer,object[] values){
   var b=new StringBuilder();for(int i=0;i<values.Length;i++){
    if(i>0)b.Append(',');string s=Convert.ToString(values[i],CultureInfo.CurrentCulture)??"";
    b.Append('"');if(s.Length>0&&("=+@-\t\r\n".IndexOf(s[0])>=0))b.Append('\'');
    foreach(char c in s){if(c=='"')b.Append('"');b.Append(c);}b.Append('"');
   }writer.WriteLine(b.ToString());
  }
 }
}
'@
}
