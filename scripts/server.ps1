param([switch]$NoBrowser,[string]$ReadyFile,[string]$DataDirectory)
$ErrorActionPreference='Stop'
. "$PSScriptRoot\core.ps1"
$project=Split-Path $PSScriptRoot -Parent
if(-not $DataDirectory){$DataDirectory=Join-Path $project 'data'}
$random=New-Object byte[] 32
$rng=[Security.Cryptography.RandomNumberGenerator]::Create()
try{$rng.GetBytes($random)}finally{$rng.Dispose()}
$token=([BitConverter]::ToString($random)).Replace('-','').ToLowerInvariant()
$session=Join-Path $DataDirectory ('session-'+$token.Substring(0,12))
[IO.Directory]::CreateDirectory($session)|Out-Null
$listener=New-Object Net.Sockets.TcpListener([Net.IPAddress]::Loopback,0)
$listener.Start()
$port=$listener.LocalEndpoint.Port
$authority='127.0.0.1:'+$port
$origin='http://'+$authority
$url=$origin+'/?token='+$token
$job=$null;$worker=$null;$sequence=0;$script:errorTask=$null
if($ReadyFile){Write-AtomicJson $ReadyFile @{url=$url;token=$token;origin=$origin;session=$session}}
Write-Host 'KTG Outlook Mail Manager - PowerShell launcher'
Write-Host ('Local address: '+$url)
Write-Host ('Reports: '+$session)
if(-not $NoBrowser){Start-Process $url -WindowStyle Normal}
function Read-LocalRequest($Stream) {
    $header=New-Object 'Collections.Generic.List[byte]'
    while($true){
        $b=$Stream.ReadByte()
        if($b -lt 0){throw 'Incomplete request'}
        $header.Add([byte]$b);$n=$header.Count
        if($n -ge 4 -and $header[$n-4] -eq 13 -and $header[$n-3] -eq 10 -and $header[$n-2] -eq 13 -and $header[$n-1] -eq 10){break}
        if($n -gt 16384){throw 'Header too large'}
    }
    $lines=([Text.Encoding]::ASCII.GetString($header.ToArray())) -split "\r\n"
    $first=$lines[0] -split ' '
    if($first.Count -ne 3 -or $first[2] -ne 'HTTP/1.1'){throw 'Invalid HTTP version'}
    $headers=@{}
    foreach($line in $lines[1..($lines.Count-1)]){
        if(-not $line){continue}
        $pair=$line -split ':',2
        if($pair.Count -ne 2){throw 'Invalid header'}
        $name=$pair[0].Trim().ToLowerInvariant()
        if($headers.ContainsKey($name)){throw 'Duplicate header'}
        $headers[$name]=$pair[1].Trim()
    }
    if($headers.ContainsKey('transfer-encoding')){throw 'Transfer encoding not supported'}
    $length=0
    if($headers.ContainsKey('content-length')){
        if(-not [int]::TryParse($headers['content-length'],[ref]$length) -or $length -lt 0 -or $length -gt 67108864){throw 'Invalid content length'}
    }
    $body=New-Object byte[] $length;$offset=0
    while($offset -lt $length){$read=$Stream.Read($body,$offset,$length-$offset);if($read -eq 0){throw 'Incomplete body'};$offset+=$read}
    return @{Method=$first[0];Target=$first[1];Headers=$headers;Body=$body}
}
function Send-LocalResponse($Stream,[string]$Code,[string]$Mime,[byte[]]$Body){
    $nl=[string][char]13+[char]10
    $headers="HTTP/1.1 $Code"+$nl+"Content-Type: $Mime"+$nl+"Content-Length: "+$Body.Length+$nl+
        "Cache-Control: no-store"+$nl+"X-Content-Type-Options: nosniff"+$nl+"Referrer-Policy: no-referrer"+$nl+
        "Content-Security-Policy: default-src 'self'; script-src 'self'; style-src 'self'; connect-src 'self'; frame-ancestors 'none'; base-uri 'none'; form-action 'none'"+$nl+"Connection: close"+$nl+$nl
    $bytes=[Text.Encoding]::ASCII.GetBytes($headers)
    $Stream.Write($bytes,0,$bytes.Length);$Stream.Write($Body,0,$Body.Length)
}
try {
    $quit=$false
    while(-not $quit){
        $client=$listener.AcceptTcpClient();$stream=$null
        try {
            $client.ReceiveTimeout=10000;$client.SendTimeout=10000;$stream=$client.GetStream()
            try{$r=Read-LocalRequest $stream}catch{Send-LocalResponse $stream '400 Bad Request' 'text/plain' ($script:Utf8.GetBytes('Invalid or oversized request'));continue}
            $code='200 OK';$mime='application/json; charset=utf-8';$body='{}';$raw=$null
            $api=$r.Target.StartsWith('/api/')
            if($r.Headers['host'] -cne $authority -or
                ($r.Headers.ContainsKey('origin') -and $r.Headers['origin'] -cne $origin) -or
                ($api -and $r.Headers['x-ktg-token'] -cne $token)){
                Send-LocalResponse $stream '403 Forbidden' 'text/plain' ($script:Utf8.GetBytes('Forbidden'));continue
            }
            $active=$null -ne $worker -and -not $worker.HasExited
            $path=($r.Target -split '\?',2)[0]
            if($r.Method -eq 'GET' -and $path -in @('/','/app.js','/style.css')){
                $file=switch($path){'/'{'index.html';$mime='text/html; charset=utf-8'} '/app.js'{'app.js';$mime='text/javascript; charset=utf-8'} '/style.css'{'style.css';$mime='text/css; charset=utf-8'}}
                $raw=[IO.File]::ReadAllBytes((Join-Path (Join-Path $project 'web') $file))
            }elseif($r.Method -eq 'GET' -and $path -eq '/api/state'){
                if($job -and [IO.File]::Exists((Join-Path $job 'state.json'))){
                    $body=[IO.File]::ReadAllText((Join-Path $job 'state.json'),$script:Utf8)
                    if(-not $active -and $body.Contains('"status":"running"')){$body='{"status":"error","note":"Worker exited unexpectedly; check worker-error.log and report.csv before retrying.","errors":1,"batches":0,"resultBatches":0}'}
                }elseif($active){$body='{"status":"running","processed":0,"total":0,"copied":0,"errors":0,"batches":0,"resultBatches":0,"elapsedSeconds":0}'}
                elseif($job){$body='{"status":"error","note":"Worker failed to start; check worker-error.log.","errors":1,"batches":0,"resultBatches":0}'}
                else{$body='{"status":"idle","processed":0,"total":0,"copied":0,"errors":0,"batches":0,"resultBatches":0,"elapsedSeconds":0}'}
            }elseif($r.Method -eq 'GET' -and $r.Target -match '^/api/batch\?kind=(messages|results)&index=(\d{1,9})$'){
                $name=$Matches[1]+'-'+([int]$Matches[2])+'.json';$body='[]'
                if($job -and [IO.File]::Exists((Join-Path $job $name))){$raw=[IO.File]::ReadAllBytes((Join-Path $job $name))}
            }elseif($r.Method -eq 'GET' -and $path -eq '/api/report'){
                $mime='text/csv; charset=utf-8';$body=''
                if($job -and [IO.File]::Exists((Join-Path $job 'report.csv'))){$raw=[IO.File]::ReadAllBytes((Join-Path $job 'report.csv'))}
            }elseif($r.Method -eq 'POST' -and $path -eq '/api/start'){
                if($active){$code='409 Conflict';$body='{"error":"A job is already running"}'}
                else {
                    if($worker -and $worker.HasExited -and $script:errorTask.IsCompleted){[IO.File]::WriteAllText((Join-Path $job 'worker-error.log'),$script:errorTask.Result,$script:Utf8);$worker.Dispose()}
                    $sequence++;$job=Join-Path $session ('job-'+[DateTime]::UtcNow.ToString('yyyyMMddHHmmssfff')+'-'+$sequence)
                    [IO.Directory]::CreateDirectory($job)|Out-Null
                    $inputPath=Join-Path $job 'request.json';[IO.File]::WriteAllBytes($inputPath,$r.Body)
                    $start=New-Object Diagnostics.ProcessStartInfo
                    $start.FileName=Join-Path $env:SystemRoot 'System32\WindowsPowerShell\v1.0\powershell.exe'
                    $start.Arguments='-NoProfile -NonInteractive -Sta -ExecutionPolicy Bypass -File "'+(Join-Path $PSScriptRoot 'bridge.ps1')+'" -InputFile "'+$inputPath+'"'
                    $start.UseShellExecute=$false;$start.CreateNoWindow=$true;$start.RedirectStandardError=$true
                    $start.EnvironmentVariables['KTG_LOG_FILE']=Join-Path $DataDirectory 'operations.log'
                    $worker=New-Object Diagnostics.Process;$worker.StartInfo=$start
                    [void]$worker.Start()
                    $script:errorTask=$worker.StandardError.ReadToEndAsync()
                    $body='{"started":true}'
                }
            }elseif($r.Method -eq 'POST' -and $path -eq '/api/pause' -and $active){
                $marker=Join-Path $job 'pause'
                if([IO.File]::Exists($marker)){[IO.File]::Delete($marker)}else{[IO.File]::WriteAllText($marker,'pause')}
                $body='{"pauseToggled":true}'
            }elseif($r.Method -eq 'POST' -and $path -eq '/api/cancel'){
                if($job){[IO.File]::WriteAllText((Join-Path $job 'cancel'),'cancel')}
                $body='{"cancelRequested":true}'
            }elseif($r.Method -eq 'POST' -and $path -eq '/api/quit' -and -not $active){
                $body='{"closed":true}';$quit=$true
            }else{$code='400 Bad Request';$body='{"error":"Unsupported request"}'}
            if($worker -and $worker.HasExited -and $script:errorTask.IsCompleted){[IO.File]::WriteAllText((Join-Path $job 'worker-error.log'),$script:errorTask.Result,$script:Utf8)}
            if($null -eq $raw){$raw=$script:Utf8.GetBytes($body)}
            Send-LocalResponse $stream $code $mime $raw
        }catch{
            Write-Warning $_.Exception.Message
            if($null -ne $stream){try{Send-LocalResponse $stream '500 Internal Server Error' 'application/json' ($script:Utf8.GetBytes('{"error":"Server error; see the launcher window"}'))}catch{}}
        }finally{if($null -ne $stream){$stream.Dispose()};$client.Dispose()}
    }
}finally{$listener.Stop();if($worker){$worker.Dispose()}}
