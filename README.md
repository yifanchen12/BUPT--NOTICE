# TongZhi

北京邮电大学信息门户“校内通知”活动整理工具。

TongZhi 会打开本机 Chrome 或 Edge，访问北邮信息门户的校内通知列表，读取通知正文，提取活动地点、活动时间、发布部门等信息，并导出为 Excel。默认只保留发生在北邮校内的活动；线上会议、校外地点、其他高校会场以及无法识别地点的通知默认不会进入结果表。

English documentation: [README.en.md](README.en.md)

## 功能

- 自动抓取校内通知列表和正文。
- 提取活动地点、活动时间、部门、活动名称、发布日期和原文链接。
- 按“活动地点 -> 活动时间 -> 部门”排序导出 Excel。
- 支持只整理当天发布的通知或指定日期发布的通知。
- 支持放宽地点筛选，便于人工复核校外、线上或地点不明确的通知。
- 支持打包为 Windows 单文件 exe。
- 支持安装 Windows 计划任务，每天自动运行。

## 安全说明

- 程序不会保存账号和密码。
- 首次运行如果出现 CAS 登录页，请在浏览器中正常登录。
- 登录状态会保存在 `runtime/chrome-profile`，后续运行会自动复用。
- `runtime/`、`outputs/`、`dist/`、`.venv/` 等本地产物已被 `.gitignore` 排除，不应提交到仓库。

## 环境要求

- Windows
- Python 3.14 或兼容版本
- Chrome 或 Edge
- PowerShell

依赖见 [requirements.txt](requirements.txt)：

```text
beautifulsoup4
openpyxl
playwright
pyinstaller
```

## 直接运行

首次运行：

```powershell
.\run.ps1 -Pages 5
```

常用参数：

```powershell
.\run.ps1 -Pages 8 -MaxItems 200
.\run.ps1 -Pages 5 -Output ".\outputs\活动整理.xlsx"
.\run.ps1 -Pages 5 -TodayOnly
.\run.ps1 -Pages 5 -Date "2026-05-14"
.\run.ps1 -Pages 3 -IncludeAll
```

如果需要临时放宽地点筛选，把校外、线上或地点不明确的结果也输出用于复核：

```powershell
.\run.ps1 -Pages 3 -IncludeOffCampus
```

## 命令行参数

主程序支持以下常用参数：

```powershell
python .\bupt_notice_crawler.py --pages 5 --max-items 120
python .\bupt_notice_crawler.py --pages 5 --today-only
python .\bupt_notice_crawler.py --pages 5 --date 2026-05-14
python .\bupt_notice_crawler.py --pages 5 --include-off-campus
python .\bupt_notice_crawler.py --pages 5 --headless
```

更多参数可运行：

```powershell
python .\bupt_notice_crawler.py --help
```

## 打包 exe

```powershell
.\build_exe.ps1
```

打包完成后运行：

```powershell
.\dist\bupt_notice_crawler.exe --pages 5
```

只整理当天发布的通知：

```powershell
.\dist\bupt_notice_crawler.exe --pages 5 --today-only
```

只整理指定日期发布的通知：

```powershell
.\dist\bupt_notice_crawler.exe --pages 5 --date 2026-05-14
```

放宽地点筛选：

```powershell
.\dist\bupt_notice_crawler.exe --pages 5 --include-off-campus
```

## Excel 输出字段

主表名称为 `活动整理`，字段顺序为：

1. 活动地点
2. 活动时间
3. 部门
4. 活动名称
5. 发布时间
6. 通知标题
7. 原文链接
8. 备注

如果某条通知没有识别出时间或部门，备注列会提示，便于人工复核。

## 每天自动运行

首次请先手动运行一次，完成 CAS 登录并保存登录状态：

```powershell
.\dist\bupt_notice_crawler.exe --pages 1
```

然后安装 Windows 计划任务。默认每天 08:30 运行，只整理当天发布的校内事项：

```powershell
.\install_daily_task.ps1
```

指定运行时间：

```powershell
.\install_daily_task.ps1 -At "20:30" -Pages 5
```

计划任务会调用：

```powershell
.\run_daily.ps1
```

每日输出位置：

```text
outputs\daily\校内通知活动_YYYY-MM-DD.xlsx
```

每日日志位置：

```text
runtime\logs\daily_YYYY-MM-DD.log
```

删除计划任务：

```powershell
.\uninstall_daily_task.ps1
```

## 校内地点判定

程序会优先识别北邮、校区、教学楼、主楼、科研楼、经管楼、学生发展中心、学生活动中心、科学会堂、体育馆、图书馆、报告厅、会议室、教室、礼堂、操场、食堂等校内地点特征。

如果学校通知中出现新的常用场地名，可以在 [bupt_notice_crawler.py](bupt_notice_crawler.py) 的 `CAMPUS_LOCATION_KEYWORDS` 中补充。

## 开发说明

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

运行脚本：

```powershell
.\run.ps1 -Pages 3
```
