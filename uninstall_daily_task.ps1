param(
    [string]$TaskName = "BUPTNoticeDailyCrawler"
)

$ErrorActionPreference = "Stop"
Unregister-ScheduledTask -TaskName $TaskName -Confirm:$false
Write-Host "计划任务已删除：$TaskName"
