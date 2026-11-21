param([string]$TestDirectory,[string]$ServerExecutable)
$ErrorActionPreference='Stop'
if(-not $TestDirectory){$TestDirectory=Join-Path ([IO.Path]::GetTempPath()) ('KTG-server-'+[guid]::NewGuid().ToString('N'))}
[IO.Directory]::CreateDirectory($TestDirectory)|Out-Null
$root=Split-Path $PSScriptRoot -Parent
$ready=Join-Path $TestDirectory 'ready.json'
$start=New-Object Diagnostics.ProcessStartInfo
if($ServerExecutable){
    $start.FileName=[IO.Path]::GetFullPath($ServerExecutable)
    $start.Arguments='--headless --data-dir "'+$TestDirectory+'\data"'
}else{
    $start.FileName=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
    $start.Arguments='-NoProfile -ExecutionPolicy Bypass -File "'+$root+'\scripts\server.ps1" -NoBrowser -ReadyFile "'+$ready+'" -DataDirectory "'+$TestDirectory+'\data"'
}
$start.CreateNoWindow=$true;$start.UseShellExecute=$false;$start.RedirectStandardError=$true;$start.RedirectStandardOutput=$true
$process=New-Object Diagnostics.Process;$process.StartInfo=$start
[void]$process.Start()
$stderr=$process.StandardError.ReadToEndAsync()
if($ServerExecutable){
    $firstLine=$process.StandardOutput.ReadLineAsync()
    if(-not $firstLine.Wait(10000)){throw 'Native server startup timeout'}
    $nativeUrl=[uri]$firstLine.Result
    $nativeToken=$nativeUrl.Query.Substring('?token='.Length)
    $script:connection=[pscustomobject]@{origin=$nativeUrl.GetLeftPart([UriPartial]::Authority);token=$nativeToken}
}
$stdout=$process.StandardOutput.ReadToEndAsync()
function Assert($Condition,[string]$Message){if(-not $Condition){throw ('FAIL: '+$Message)}}
function Http([string]$Path,[string]$Method='GET',[string]$Body,[switch]$NoToken,[string]$Origin){
    $request=[Net.HttpWebRequest]::Create($script:connection.origin+$Path)
    $request.Proxy=$null;$request.Method=$Method;$request.Timeout=10000
    if(-not $NoToken){$request.Headers.Add('x-ktg-token',$script:connection.token)}
    if($Origin){$request.Headers.Add('Origin',$Origin)}
    if($Method -eq 'POST'){
        $bytes=[Text.Encoding]::UTF8.GetBytes($Body)
        $request.ContentType='application/json';$request.ContentLength=$bytes.Length
        $s=$request.GetRequestStream();$s.Write($bytes,0,$bytes.Length);$s.Dispose()
    }
    try{$response=$request.GetResponse()}catch [Net.WebException]{if($_.Exception.Response){$response=$_.Exception.Response}else{throw}}
    try{$reader=New-Object IO.StreamReader($response.GetResponseStream());$text=$reader.ReadToEnd();$reader.Dispose();return @{Status=[int]$response.StatusCode;Body=$text;Mime=$response.ContentType}}finally{$response.Dispose()}
}
try{
    if(-not $ServerExecutable){
        for($i=0;$i -lt 100 -and -not [IO.File]::Exists($ready);$i++){Start-Sleep -Milliseconds 100;if($process.HasExited){throw $stderr.Result}}
        Assert ([IO.File]::Exists($ready)) 'Server starts'
        $script:connection=Get-Content -LiteralPath $ready -Raw|ConvertFrom-Json
    }
    $r=Http '/';Assert ($r.Status -eq 200 -and $r.Body.Contains('KTG Outlook')) 'HTML served'
    $r=Http '/app.js';Assert ($r.Status -eq 200 -and $r.Mime.Contains('javascript')) 'JavaScript MIME'
    $r=Http '/api/state' -NoToken;Assert ($r.Status -eq 403) 'Missing token rejected'
    $r=Http '/api/state' -Origin 'https://example.invalid';Assert ($r.Status -eq 403) 'Foreign origin rejected'
    $r=Http '/api/state';Assert (($r.Body|ConvertFrom-Json).status -eq 'idle') 'Initial state'
    $r=Http '/api/batch?kind=../../request&index=0';Assert ($r.Status -eq 400) 'Traversal rejected'
    $r=Http '/api/start' 'POST' '{"operation":"delete"}'
    Assert ($r.Status -eq 200 -and ($r.Body|ConvertFrom-Json).started) 'Worker launched'
    $paused=Http '/api/pause' 'POST' '{}'
    Assert ($paused.Status -eq 200 -and ($paused.Body|ConvertFrom-Json).pauseToggled) 'Pause endpoint marks active job'
    $resumed=Http '/api/pause' 'POST' '{}'
    Assert ($resumed.Status -eq 200 -and ($resumed.Body|ConvertFrom-Json).pauseToggled) 'Resume endpoint toggles active job'
    $state=$null
    for($i=0;$i -lt 100;$i++){
        Start-Sleep -Milliseconds 100
        $state=(Http '/api/state').Body|ConvertFrom-Json
        if($state.status -eq 'error'){break}
    }
    Assert ($state.status -eq 'error' -and $state.note.Contains('Xóa cần xác nhận riêng')) 'Unconfirmed delete rejected before Outlook access'
    Start-Sleep -Milliseconds 200
    $r=Http '/api/start' 'POST' '{"operation":"delete_archived","confirmPermanentDelete":true,"confirmationText":"XOA 1","archiveJobId":"../other","messages":[{"id":"fake","storeId":"fake","sourceFolderId":"fake","subject":"fixture","internetMessageId":""}]}'
    Assert ($r.Status -eq 200 -and ($r.Body|ConvertFrom-Json).started) 'Archive cleanup validation worker starts'
    for($i=0;$i -lt 100;$i++){Start-Sleep -Milliseconds 100;$state=(Http '/api/state').Body|ConvertFrom-Json;if($state.status -eq 'error'){break}}
    Assert ($state.status -eq 'error' -and $state.note.Contains('Thiếu phiên xuất PST')) 'Invalid archive receipt refused before Outlook access'
    Start-Sleep -Milliseconds 200
    $r=Http '/api/start' 'POST' '{"operation":"export","autoDeleteAfterExport":true,"messages":[{"id":"fake","storeId":"fake","sourceFolderId":"fake","subject":"fixture"}]}'
    Assert ($r.Status -eq 200 -and ($r.Body|ConvertFrom-Json).started) 'Automatic cleanup validation worker starts'
    for($i=0;$i -lt 100;$i++){Start-Sleep -Milliseconds 100;$state=(Http '/api/state').Body|ConvertFrom-Json;if($state.status -eq 'error'){break}}
    Assert ($state.status -eq 'error' -and $state.note.Contains('Thiếu xác nhận tự xóa')) 'Automatic cleanup without upfront consent refused before Outlook access'
    Start-Sleep -Milliseconds 200
    $r=Http '/api/quit' 'POST' '{}'
    Assert ($r.Status -eq 200) 'Server quits cleanly'
    Assert ($process.WaitForExit(10000)) 'Server process exits'
    [IO.File]::WriteAllText((Join-Path $TestDirectory 'server-output.log'),$stdout.Result+[Environment]::NewLine+$stderr.Result)
    @{passed=$true;checks=15;native=[bool]$ServerExecutable;testDirectory=$TestDirectory}|ConvertTo-Json
}finally{
    if(-not $process.HasExited){$process.Kill()}
    $process.Dispose()
}
