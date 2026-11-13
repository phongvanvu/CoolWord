$ErrorActionPreference = "Stop"

$msbuild = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
if (-not (Test-Path $msbuild)) {
    $msbuild = "C:\Windows\Microsoft.NET\Framework\v4.0.30319\MSBuild.exe"
}

Write-Host "==> Đang biên dịch giải pháp CoolWord (Release)..." -ForegroundColor Cyan
& $msbuild "$PSScriptRoot\CoolWord.sln" /p:Configuration=Release /v:minimal
if ($LASTEXITCODE -ne 0) {
    throw "Biên dịch giải pháp thất bại!"
}

Write-Host "`n==> Đang chạy bộ kiểm thử tự động (CoolWord.Tests)..." -ForegroundColor Cyan
& "$PSScriptRoot\bin\CoolWord.Tests.exe"
if ($LASTEXITCODE -ne 0) {
    throw "Kiểm thử tự động thất bại!"
}

Write-Host "`n==> Hoàn tất biên dịch và kiểm thử tại thư mục: $PSScriptRoot\bin" -ForegroundColor Green
