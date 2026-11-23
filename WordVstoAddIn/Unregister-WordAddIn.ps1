$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$dllPath = Join-Path $scriptDir "bin\Release\net48\WordVstoAddIn.dll"
if (-not (Test-Path $dllPath)) {
    $dllPath = Join-Path $scriptDir "WordVstoAddIn.dll"
}

$regAsm64 = "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
$regAsm32 = "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"

if (Test-Path $dllPath) {
    if (Test-Path $regAsm64) { & $regAsm64 $dllPath /unregister /nologo | Out-Host }
    if (Test-Path $regAsm32) { & $regAsm32 $dllPath /unregister /nologo | Out-Host }
}

$addInRegPath = "HKCU:\Software\Microsoft\Office\Word\Addins\WordVstoAddIn.Connect"
if (Test-Path $addInRegPath) {
    Remove-Item -Path $addInRegPath -Recurse -Force
}

Write-Host "WordVstoAddIn unregistered cleanly." -ForegroundColor Green
