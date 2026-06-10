param(
    [string]$At = "08:30",
    [int]$Pages = 3,
    [int]$MaxItems = 120,
    [string]$TaskName = "BUPTNoticeDailyCrawler",
    [switch]$IncludeOffCampus
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$DailyScript = Join-Path $Root "run_daily.ps1"

if (-not (Test-Path $DailyScript)) {
    throw "找不到每日运行脚本：$DailyScript"
}

$PowerShell = Join-Path $env:SystemRoot "System32\WindowsPowerShell\v1.0\powershell.exe"
$ArgumentList = @(
    "-NoProfile",
    "-ExecutionPolicy", "Bypass",
    "-File", "`"$DailyScript`"",
    "-Pages", $Pages,
    "-MaxItems", $MaxItems
)

if ($IncludeOffCampus) {
    $ArgumentList += "-IncludeOffCampus"
}

$Action = New-ScheduledTaskAction -Execute $PowerShell -Argument ($ArgumentList -join " ") -WorkingDirectory $Root
$Trigger = New-ScheduledTaskTrigger -Daily -At ([datetime]::Parse($At))
$Principal = New-ScheduledTaskPrincipal -UserId ([System.Security.Principal.WindowsIdentity]::GetCurrent().Name) -LogonType Interactive -RunLevel Limited
$Settings = New-ScheduledTaskSettingsSet -StartWhenAvailable -ExecutionTimeLimit (New-TimeSpan -Hours 2)

Register-ScheduledTask `
    -TaskName $TaskName `
    -Action $Action `
    -Trigger $Trigger `
    -Principal $Principal `
    -Settings $Settings `
    -Description "Daily BUPT portal notice crawler. Output: outputs\daily\校内通知活动_YYYY-MM-DD.xlsx" `
    -Force | Out-Null

Write-Host "计划任务已安装：$TaskName"
Write-Host "运行时间：每天 $At"
Write-Host "输出目录：$Root\outputs\daily"
Write-Host "日志目录：$Root\runtime\logs"
