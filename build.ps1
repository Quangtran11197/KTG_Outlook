$ErrorActionPreference='Stop'
Set-Location $PSScriptRoot
$cargoCommand=Get-Command cargo -ErrorAction SilentlyContinue
$cargo=if($cargoCommand){$cargoCommand.Source}else{Join-Path ([Environment]::GetFolderPath('UserProfile')) '.cargo\bin\cargo.exe'}
if(-not (Test-Path -LiteralPath $cargo)){
    throw 'Chưa có Cargo. Cài Rust stable và Microsoft C++ Build Tools theo README, sau đó chạy lại.'
}
$cargoTarget=if($env:CARGO_TARGET_DIR){$env:CARGO_TARGET_DIR}else{Join-Path $PSScriptRoot 'target'}
& "$PSScriptRoot\native\build.cmd"
if($LASTEXITCODE -ne 0){throw 'Native MAPI build failed'}
& $cargo test --locked --target-dir $cargoTarget
if($LASTEXITCODE -ne 0){throw 'Rust tests failed'}
& "$PSScriptRoot\tests\test-core.ps1"
& "$PSScriptRoot\tests\test-scan.ps1"
& "$PSScriptRoot\tests\test-scan-benchmark.ps1" -Count 70000
& "$PSScriptRoot\tests\test-delete.ps1"
& "$PSScriptRoot\tests\test-auto-export.ps1"
& "$PSScriptRoot\tests\test-bulk.ps1"
& "$PSScriptRoot\tests\test-native-copy-results.ps1"
& "$PSScriptRoot\tests\test-server.ps1"
& $cargo build --release --locked --target-dir $cargoTarget
if($LASTEXITCODE -ne 0){throw 'Rust build failed'}
Copy-Item -LiteralPath (Join-Path $cargoTarget 'release\ktg-outlook-mail-manager.exe') -Destination "$PSScriptRoot\KTG-Outlook-Mail-Manager-Desktop-0.5.23.exe"
Write-Host 'Built KTG-Outlook-Mail-Manager-Desktop-0.5.23.exe'
& "$PSScriptRoot\tests\test-server.ps1" -ServerExecutable "$PSScriptRoot\KTG-Outlook-Mail-Manager-Desktop-0.5.23.exe"




