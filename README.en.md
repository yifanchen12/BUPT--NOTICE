# TongZhi

A BUPT portal notice crawler that extracts campus activity information and exports it to Excel.

TongZhi opens the local Chrome or Edge browser, visits the “Campus Notices” section of the Beijing University of Posts and Telecommunications information portal, reads notice articles, extracts activity location, time, department, and related metadata, then exports the result to an Excel file. By default, it only keeps activities that happen on the BUPT campus. Online meetings, off-campus venues, other university venues, and notices with unrecognized locations are excluded unless you explicitly loosen the filter.

Chinese documentation: [README.md](README.md)

## Features

- Crawl campus notice list pages and article bodies.
- Extract location, time, department, activity name, publish date, notice title, and source URL.
- Export a sorted Excel table by location, time, and department.
- Filter notices by today or by a specific publish date.
- Optionally include off-campus, online, or uncertain-location notices for manual review.
- Build a Windows single-file executable.
- Install a Windows scheduled task for daily runs.

## Security Notes

- The tool does not store usernames or passwords.
- On the first run, if the CAS login page appears, log in manually in the browser.
- Browser login state is stored under `runtime/chrome-profile` and reused later.
- Local artifacts such as `runtime/`, `outputs/`, `dist/`, and `.venv/` are ignored by Git and should not be committed.

## Requirements

- Windows
- Python 3.14 or a compatible version
- Chrome or Edge
- PowerShell

Dependencies are listed in [requirements.txt](requirements.txt):

```text
beautifulsoup4
openpyxl
playwright
pyinstaller
```

## Run from Source

First run:

```powershell
.\run.ps1 -Pages 5
```

Common examples:

```powershell
.\run.ps1 -Pages 8 -MaxItems 200
.\run.ps1 -Pages 5 -Output ".\outputs\activities.xlsx"
.\run.ps1 -Pages 5 -TodayOnly
.\run.ps1 -Pages 5 -Date "2026-05-14"
.\run.ps1 -Pages 3 -IncludeAll
```

To temporarily loosen location filtering and include off-campus, online, or uncertain-location notices for review:

```powershell
.\run.ps1 -Pages 3 -IncludeOffCampus
```

## CLI Usage

Common direct Python commands:

```powershell
python .\bupt_notice_crawler.py --pages 5 --max-items 120
python .\bupt_notice_crawler.py --pages 5 --today-only
python .\bupt_notice_crawler.py --pages 5 --date 2026-05-14
python .\bupt_notice_crawler.py --pages 5 --include-off-campus
python .\bupt_notice_crawler.py --pages 5 --headless
```

Show all options:

```powershell
python .\bupt_notice_crawler.py --help
```

## Build the Executable

```powershell
.\build_exe.ps1
```

Run the packaged executable:

```powershell
.\dist\bupt_notice_crawler.exe --pages 5
```

Only process notices published today:

```powershell
.\dist\bupt_notice_crawler.exe --pages 5 --today-only
```

Only process notices published on a specific date:

```powershell
.\dist\bupt_notice_crawler.exe --pages 5 --date 2026-05-14
```

Loosen location filtering:

```powershell
.\dist\bupt_notice_crawler.exe --pages 5 --include-off-campus
```

## Excel Output

The main sheet is named `活动整理` and contains these columns:

1. Activity Location
2. Activity Time
3. Department
4. Activity Name
5. Publish Date
6. Notice Title
7. Source URL
8. Notes

If time or department cannot be identified from a notice, the notes column will flag it for manual review.

## Daily Scheduled Run

Before installing the task, run the packaged executable manually once and finish CAS login:

```powershell
.\dist\bupt_notice_crawler.exe --pages 1
```

Install the Windows scheduled task. By default, it runs every day at 08:30 and only processes notices published that day:

```powershell
.\install_daily_task.ps1
```

Set a custom run time:

```powershell
.\install_daily_task.ps1 -At "20:30" -Pages 5
```

The scheduled task calls:

```powershell
.\run_daily.ps1
```

Daily output path:

```text
outputs\daily\校内通知活动_YYYY-MM-DD.xlsx
```

Daily log path:

```text
runtime\logs\daily_YYYY-MM-DD.log
```

Remove the scheduled task:

```powershell
.\uninstall_daily_task.ps1
```

## Campus Location Matching

The crawler prioritizes BUPT campus keywords such as BUPT, campus names, teaching buildings, main building, research building, economics and management building, student development center, student activity center, science hall, gymnasium, library, lecture hall, meeting room, classroom, auditorium, playground, and cafeteria.

If new common venue names appear in school notices, add them to `CAMPUS_LOCATION_KEYWORDS` in [bupt_notice_crawler.py](bupt_notice_crawler.py).

## Development

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Run:

```powershell
.\run.ps1 -Pages 3
```
