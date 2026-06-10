# TongZhi

[中文](README.md)

TongZhi is a structured activity-information extractor for the “Campus Notices” section of the Beijing University of Posts and Telecommunications information portal. It opens the portal through a local browser, reads notice list pages and article bodies, extracts activity location, time, department, activity name, publish date, and source URL, then exports the results to an Excel workbook.

## Purpose

Campus notice articles are not fully standardized. Time, venue, department, and activity names may appear in the title, article body, or publish metadata. TongZhi is designed to automate the repetitive part of this workflow while keeping ambiguous records visible for manual review.

The default filtering strategy is conservative. It only exports activities identified as happening on the BUPT campus. Online meetings, off-campus venues, other university venues, and records with unclear locations are excluded unless the filter is explicitly loosened.

## Features

- Crawl BUPT campus notice list pages and article bodies.
- Extract activity location, time, department, activity name, publish date, notice title, and source URL.
- Export a sorted Excel workbook by location, time, and department.
- Filter notices by today or by a specific publish date.
- Include off-campus, online, or uncertain-location records when manual review is needed.
- Build a Windows single-file executable with PyInstaller.
- Install a Windows scheduled task for daily automated runs.

## Repository Layout

```text
.
├── bupt_notice_crawler.py      # Main crawler
├── run.ps1                     # Source-mode runner
├── build_exe.ps1               # PyInstaller build script
├── run_daily.ps1               # Daily run script
├── install_daily_task.ps1      # Windows scheduled-task installer
├── uninstall_daily_task.ps1    # Windows scheduled-task remover
├── requirements.txt            # Python dependencies
├── README.md                   # Chinese documentation
└── README.en.md                # English documentation
```

The following local runtime artifacts are ignored by Git:

```text
.venv/      # Python virtual environment
dist/       # Packaged executable
build/      # PyInstaller build cache
runtime/    # Browser profile, login state, and logs
outputs/    # Excel output files
```

## Requirements

- Windows
- PowerShell
- Python 3.14 or a compatible version
- Chrome or Edge

Python dependencies:

```text
beautifulsoup4
openpyxl
playwright
pyinstaller
```

## Quick Start

Run the crawler from source:

```powershell
.\run.ps1 -Pages 5
```

On the first run, the browser may open a CAS login page. Log in manually in that browser window. The session state is stored locally under:

```text
runtime\chrome-profile
```

Later runs reuse this browser profile.

## Common Usage

Crawl 8 list pages and read up to 200 notices:

```powershell
.\run.ps1 -Pages 8 -MaxItems 200
```

Write to a custom output path:

```powershell
.\run.ps1 -Pages 5 -Output ".\outputs\activities.xlsx"
```

Only process notices published today:

```powershell
.\run.ps1 -Pages 5 -TodayOnly
```

Only process notices published on a specific date:

```powershell
.\run.ps1 -Pages 5 -Date "2026-05-14"
```

Export all extracted activity records:

```powershell
.\run.ps1 -Pages 3 -IncludeAll
```

Include off-campus, online, or uncertain-location records for review:

```powershell
.\run.ps1 -Pages 3 -IncludeOffCampus
```

## CLI Options

The Python entry point can also be called directly:

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

Common options:

| Option | Description |
| --- | --- |
| `--pages` | Number of notice list pages to crawl |
| `--max-items` | Maximum number of notices to read |
| `--output` | Excel output path |
| `--today-only` | Only process notices published today |
| `--date` | Only process notices published on `YYYY-MM-DD` |
| `--include-all` | Export all extracted activity records |
| `--include-off-campus` | Include off-campus, online, and uncertain-location records |
| `--headless` | Run without a visible browser window after login state exists |
| `--profile-dir` | Custom browser profile directory |
| `--chrome-path` | Custom Chrome or Edge executable path |

## Excel Output

The default output path uses this pattern:

```text
outputs\校内通知活动_YYYYMMDD_HHMMSS.xlsx
```

The main sheet is named `活动整理` and contains:

| Column | Description |
| --- | --- |
| 活动地点 | Extracted activity venue |
| 活动时间 | Extracted activity time |
| 部门 | Publishing or organizing department |
| 活动名称 | Activity name inferred from title and body |
| 发布时间 | Notice publish date |
| 通知标题 | Original notice title |
| 原文链接 | Source article URL |
| 备注 | Review notes or filtering hints |

## Build the Executable

Run:

```powershell
.\build_exe.ps1
```

The executable will be created at:

```text
dist\bupt_notice_crawler.exe
```

Examples:

```powershell
.\dist\bupt_notice_crawler.exe --pages 5
.\dist\bupt_notice_crawler.exe --pages 5 --today-only
.\dist\bupt_notice_crawler.exe --pages 5 --date 2026-05-14
.\dist\bupt_notice_crawler.exe --pages 5 --include-off-campus
```

## Daily Scheduled Run

Before installing the scheduled task, run the packaged executable manually once and complete CAS login:

```powershell
.\dist\bupt_notice_crawler.exe --pages 1
```

Install the Windows scheduled task. By default, it runs every day at 08:30:

```powershell
.\install_daily_task.ps1
```

Set a custom run time and page count:

```powershell
.\install_daily_task.ps1 -At "20:30" -Pages 5
```

The scheduled task calls:

```powershell
.\run_daily.ps1
```

Daily output:

```text
outputs\daily\校内通知活动_YYYY-MM-DD.xlsx
```

Daily logs:

```text
runtime\logs\daily_YYYY-MM-DD.log
```

Remove the scheduled task:

```powershell
.\uninstall_daily_task.ps1
```

## Campus Location Matching

TongZhi prioritizes BUPT campus-related location keywords, including campus names, teaching buildings, the main building, research buildings, student activity spaces, science hall, gymnasium, library, lecture halls, meeting rooms, classrooms, auditorium, playground, and cafeterias.

If new common venue names appear in portal notices, add them to `CAMPUS_LOCATION_KEYWORDS` in [bupt_notice_crawler.py](bupt_notice_crawler.py).

## Security and Privacy

- TongZhi does not store usernames or passwords.
- CAS login is completed manually by the user in the browser.
- Browser session state is stored only on the local machine under `runtime/chrome-profile`.
- Do not commit `runtime/`, `outputs/`, `dist/`, `.venv/`, or other local artifacts to a public repository.

## Development

Create a virtual environment and install dependencies:

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

Run:

```powershell
.\run.ps1 -Pages 3
```

## Limitations

TongZhi depends on the portal page structure and keyword-based extraction rules. If the portal markup changes, or if a notice does not include clear time, venue, or department information, the generated result may require manual review.

## License

This project is licensed under the MIT License. See [LICENSE](LICENSE) for the copyright notice.
