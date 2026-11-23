param(
    [ValidateSet("x86", "x64", "Auto")]
    [string]$OfficeArch = "Auto"
)

$ErrorActionPreference = "Stop"

$scriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
$csproj = Join-Path $scriptDir "WordVstoAddIn.csproj"

# Detect installed Microsoft Office architecture when set to Auto
if ($OfficeArch -eq "Auto") {
    $c2r = Get-ItemProperty -Path "HKLM:\SOFTWARE\Microsoft\Office\ClickToRun\Configuration" -ErrorAction SilentlyContinue
    if ($c2r -and $c2r.Platform -eq "x64") {
        $OfficeArch = "x64"
    } elseif (Test-Path "C:\Program Files (x86)\Microsoft Office\root\Office16\WINWORD.EXE") {
        $OfficeArch = "x86"
    } elseif (Test-Path "C:\Program Files\Microsoft Office\root\Office16\WINWORD.EXE") {
        $OfficeArch = "x64"
    } else {
        $OfficeArch = "x86"
    }
    Write-Host "Detected Microsoft Word architecture: $OfficeArch" -ForegroundColor Yellow
}

if (Test-Path $csproj) {
    Write-Host "Building WordVstoAddIn ($OfficeArch, Release)..." -ForegroundColor Cyan
    dotnet build $csproj -c Release -p:Platform=$OfficeArch | Out-Host
    $candidateArchPath = Join-Path $scriptDir "bin\$OfficeArch\Release\net48\WordVstoAddIn.dll"
    $candidateDefaultPath = Join-Path $scriptDir "bin\Release\net48\WordVstoAddIn.dll"
    $dllPath = if (Test-Path $candidateArchPath) { $candidateArchPath } else { $candidateDefaultPath }
} else {
    $dllPath = Join-Path $scriptDir "WordVstoAddIn.dll"
}

if (-not (Test-Path $dllPath)) {
    throw "WordVstoAddIn.dll not found at: $dllPath"
}

$regAsm = if ($OfficeArch -eq "x64") {
    "$env:WINDIR\Microsoft.NET\Framework64\v4.0.30319\RegAsm.exe"
} else {
    "$env:WINDIR\Microsoft.NET\Framework\v4.0.30319\RegAsm.exe"
}

if (-not (Test-Path $regAsm)) {
    throw "RegAsm.exe not found at: $regAsm"
}

Write-Host "Registering COM assembly ($OfficeArch) via $regAsm..." -ForegroundColor Cyan
& $regAsm $dllPath /codebase /nologo | Out-Host

$addInRegPath = "HKCU:\Software\Microsoft\Office\Word\Addins\WordVstoAddIn.Connect"
New-Item -Path $addInRegPath -Force | Out-Null
New-ItemProperty -Path $addInRegPath -Name "FriendlyName" -Value "Internship Word Tools Add-In" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addInRegPath -Name "Description" -Value "ASP.NET Core 8 & Word COM Interop Document Tools" -PropertyType String -Force | Out-Null
New-ItemProperty -Path $addInRegPath -Name "LoadBehavior" -Value 3 -PropertyType DWord -Force | Out-Null

Write-Host "WordVstoAddIn ($OfficeArch) registered successfully! Restart Microsoft Word to see the 'Internship Word Tools' tab." -ForegroundColor Green
