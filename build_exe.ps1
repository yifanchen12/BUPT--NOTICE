$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path $Python)) {
    py -3.14 -m venv (Join-Path $Root ".venv")
}

& $Python -m pip install --upgrade pip
& $Python -m pip install -r (Join-Path $Root "requirements.txt")

Push-Location $Root
try {
    & $Python -m PyInstaller `
        --noconfirm `
        --clean `
        --onefile `
        --name "bupt_notice_crawler" `
        --collect-all playwright `
        --hidden-import bs4 `
        "bupt_notice_crawler.py"
}
finally {
    Pop-Location
}

Write-Host ""
Write-Host "打包完成：$Root\dist\bupt_notice_crawler.exe"
Write-Host "运行示例：.\dist\bupt_notice_crawler.exe --pages 5"
