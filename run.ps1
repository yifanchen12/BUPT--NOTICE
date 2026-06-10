param(
    [int]$Pages = 3,
    [int]$MaxItems = 120,
    [string]$Output = "",
    [switch]$TodayOnly,
    [string]$Date = "",
    [switch]$IncludeAll,
    [switch]$IncludeOffCampus,
    [switch]$Headless
)

$ErrorActionPreference = "Stop"
$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
$Python = Join-Path $Root ".venv\Scripts\python.exe"

if (-not (Test-Path $Python)) {
    py -3.14 -m venv (Join-Path $Root ".venv")
    & $Python -m pip install --upgrade pip
    & $Python -m pip install -r (Join-Path $Root "requirements.txt")
}

$ArgsList = @(
    (Join-Path $Root "bupt_notice_crawler.py"),
    "--pages", $Pages,
    "--max-items", $MaxItems
)

if ($Output -ne "") {
    $ArgsList += @("--output", $Output)
}
if ($TodayOnly) {
    $ArgsList += "--today-only"
}
if ($Date -ne "") {
    $ArgsList += @("--date", $Date)
}
if ($IncludeAll) {
    $ArgsList += "--include-all"
}
if ($IncludeOffCampus) {
    $ArgsList += "--include-off-campus"
}
if ($Headless) {
    $ArgsList += "--headless"
}

& $Python @ArgsList
