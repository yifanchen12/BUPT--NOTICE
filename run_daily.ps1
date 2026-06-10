param(
    [int]$Pages = 3,
    [int]$MaxItems = 120,
    [switch]$IncludeOffCampus,
    [switch]$Visible
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Exe = Join-Path $Root "dist\bupt_notice_crawler.exe"

if (-not (Test-Path $Exe)) {
    throw "找不到已打包程序：$Exe。请先运行 .\build_exe.ps1"
}

$OutputDir = Join-Path $Root "outputs\daily"
$LogDir = Join-Path $Root "runtime\logs"
New-Item -ItemType Directory -Force -Path $OutputDir, $LogDir | Out-Null

$DateText = Get-Date -Format "yyyy-MM-dd"
$Output = Join-Path $OutputDir "校内通知活动_$DateText.xlsx"
$LogFile = Join-Path $LogDir "daily_$DateText.log"

$ArgsList = @(
    "--today-only",
    "--pages", $Pages,
    "--max-items", $MaxItems,
    "--output", $Output
)

if (-not $Visible) {
    $ArgsList += "--headless"
}
if ($IncludeOffCampus) {
    $ArgsList += "--include-off-campus"
}

"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Start daily crawl" | Out-File -FilePath $LogFile -Append -Encoding utf8
& $Exe @ArgsList *>> $LogFile
$Code = $LASTEXITCODE
"[$(Get-Date -Format 'yyyy-MM-dd HH:mm:ss')] Exit code: $Code" | Out-File -FilePath $LogFile -Append -Encoding utf8

if ($Code -ne 0) {
    throw "每日爬取失败，详见日志：$LogFile"
}

Write-Host "完成：$Output"
