$ErrorActionPreference = "Stop"

$dllPath = "$PSScriptRoot\bin\CoolWord.dll"
if (Test-Path $dllPath) {
    Write-Host "==> Đang gỡ cài đặt Add-in CoolWord khỏi Microsoft Word..." -ForegroundColor Yellow
    [Reflection.Assembly]::LoadFrom($dllPath) | Out-Null
    $msg = ""
    [CoolWord.AddIn.AddInInstaller]::UninstallForCurrentUser([ref]$msg) | Out-Null
    Write-Host $msg -ForegroundColor Green
} else {
    Remove-Item -Path "HKCU:\Software\Microsoft\Office\Word\Addins\CoolWord.AddIn" -Recurse -Force -ErrorAction SilentlyContinue
    Write-Host "Đã xóa khóa Registry Add-in CoolWord." -ForegroundColor Green
}
