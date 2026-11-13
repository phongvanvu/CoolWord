$ErrorActionPreference = "Stop"

$dllPath = "$PSScriptRoot\bin\CoolWord.dll"
if (-not (Test-Path $dllPath)) {
    & "$PSScriptRoot\build.ps1"
}

Write-Host "==> Đang đăng ký Add-in CoolWord vào Microsoft Word (Current User - không cần quyền Admin)..." -ForegroundColor Cyan
[Reflection.Assembly]::LoadFrom($dllPath) | Out-Null
$msg = ""
$ok = [CoolWord.AddIn.AddInInstaller]::InstallForCurrentUser([ref]$msg)
if ($ok) {
    Write-Host $msg -ForegroundColor Green
} else {
    throw $msg
}
