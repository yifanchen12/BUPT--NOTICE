# BUPT-NOTICE 

[English](README.en.md)

BUPT-NOTICE 是一个面向北京邮电大学信息门户“校内通知”栏目的活动信息整理工具。它通过本机浏览器访问通知列表与正文，自动提取活动地点、时间、发布部门、活动名称、发布时间和原文链接，并导出结构化 Excel 表格，适合用于日常活动汇总、部门信息整理和人工复核前的数据预处理。

## 项目定位

学校通知的正文格式并不完全统一，活动时间、地点和部门信息可能分布在标题、正文段落或发布信息中。BUPT-NOTICE 的目标不是替代人工判断，而是把高频、重复、格式相近的整理工作自动化，并把无法可靠识别的内容标记出来，方便后续复核。

默认策略偏保守：只输出识别为北邮校内地点的活动；线上会议、校外场地、其他高校会场以及地点不明确的通知默认会被过滤。

## 核心功能

- 抓取北邮信息门户校内通知列表与通知正文。
- 从正文中提取活动地点、活动时间、发布部门、活动名称等字段。
- 按“活动地点 -> 活动时间 -> 部门”排序生成 Excel。
- 支持按发布日期筛选，包括当天通知和指定日期通知。
- 支持放宽地点筛选，输出校外、线上或地点不明确的记录供人工复核。
- 支持使用 PyInstaller 打包为 Windows 单文件可执行程序。
- 支持通过 Windows 计划任务进行每日自动整理。

## 目录结构

```text
.
├── bupt_notice_crawler.py      # 主程序
├── run.ps1                     # 源码运行入口
├── build_exe.ps1               # exe 打包脚本
├── run_daily.ps1               # 每日自动运行脚本
├── install_daily_task.ps1      # 安装 Windows 计划任务
├── uninstall_daily_task.ps1    # 删除 Windows 计划任务
├── requirements.txt            # Python 依赖
├── README.md                   # 中文文档
└── README.en.md                # 英文文档
```

运行过程中会生成以下本地目录，它们已被 `.gitignore` 排除：

```text
.venv/      # Python 虚拟环境
dist/       # 打包产物
build/      # PyInstaller 构建缓存
runtime/    # 浏览器登录状态、日志等运行数据
outputs/    # Excel 输出文件
```

## 环境要求

- Windows
- PowerShell
- Python 3.14 或兼容版本
- Chrome 或 Edge

Python 依赖：

```text
beautifulsoup4
openpyxl
playwright
pyinstaller
```

## 快速开始

首次运行建议使用可见浏览器模式，以便完成 CAS 登录：

```powershell
.\run.ps1 -Pages 5
```

如果出现统一身份认证页面，请在打开的浏览器中手动登录。登录状态会保存在：

```text
runtime\chrome-profile
```

后续运行会复用该登录状态。

## 常用运行方式

抓取 8 页、最多读取 200 条通知：

```powershell
.\run.ps1 -Pages 8 -MaxItems 200
```

指定输出文件：

```powershell
.\run.ps1 -Pages 5 -Output ".\outputs\活动整理.xlsx"
```

只整理当天发布的通知：

```powershell
.\run.ps1 -Pages 5 -TodayOnly
```

只整理指定日期发布的通知：

```powershell
.\run.ps1 -Pages 5 -Date "2026-05-14"
```

输出全部可提取记录：

```powershell
.\run.ps1 -Pages 3 -IncludeAll
```

放宽地点筛选，包含校外、线上或地点不明确的记录：

```powershell
.\run.ps1 -Pages 3 -IncludeOffCampus
```

## 命令行参数

也可以直接调用 Python 主程序：

```powershell
python .\bupt_notice_crawler.py --pages 5 --max-items 120
python .\bupt_notice_crawler.py --pages 5 --today-only
python .\bupt_notice_crawler.py --pages 5 --date 2026-05-14
python .\bupt_notice_crawler.py --pages 5 --include-off-campus
python .\bupt_notice_crawler.py --pages 5 --headless
```

查看完整参数：

```powershell
python .\bupt_notice_crawler.py --help
```

常用参数说明：

| 参数 | 说明 |
| --- | --- |
| `--pages` | 抓取通知列表页数 |
| `--max-items` | 最多读取通知数量 |
| `--output` | Excel 输出路径 |
| `--today-only` | 只整理当天发布的通知 |
| `--date` | 只整理指定日期发布的通知，格式为 `YYYY-MM-DD` |
| `--include-all` | 输出所有提取到的活动记录 |
| `--include-off-campus` | 包含校外、线上或地点不明确记录 |
| `--headless` | 无头模式运行，适合已有登录状态的自动任务 |
| `--profile-dir` | 指定浏览器用户数据目录 |
| `--chrome-path` | 指定 Chrome 或 Edge 可执行文件路径 |

## Excel 输出

默认输出路径形如：

```text
outputs\校内通知活动_YYYYMMDD_HHMMSS.xlsx
```

主表名称为 `活动整理`，字段如下：

| 字段 | 说明 |
| --- | --- |
| 活动地点 | 从标题或正文中识别出的活动地点 |
| 活动时间 | 从标题或正文中识别出的活动时间 |
| 部门 | 通知发布部门或正文中识别出的组织部门 |
| 活动名称 | 根据标题和正文整理的活动名称 |
| 发布时间 | 通知发布日期 |
| 通知标题 | 原始通知标题 |
| 原文链接 | 信息门户原文地址 |
| 备注 | 无法识别、需复核或筛选原因 |

## 打包为 exe

运行打包脚本：

```powershell
.\build_exe.ps1
```

打包完成后，可执行文件位于：

```text
dist\bupt_notice_crawler.exe
```

示例：

```powershell
.\dist\bupt_notice_crawler.exe --pages 5
.\dist\bupt_notice_crawler.exe --pages 5 --today-only
.\dist\bupt_notice_crawler.exe --pages 5 --date 2026-05-14
.\dist\bupt_notice_crawler.exe --pages 5 --include-off-campus
```

## 每日自动运行

首次安装计划任务前，建议先手动运行一次打包后的程序，完成 CAS 登录：

```powershell
.\dist\bupt_notice_crawler.exe --pages 1
```

安装 Windows 计划任务，默认每天 08:30 运行：

```powershell
.\install_daily_task.ps1
```

指定运行时间和抓取页数：

```powershell
.\install_daily_task.ps1 -At "20:30" -Pages 5
```

每日任务会调用：

```powershell
.\run_daily.ps1
```

每日输出文件：

```text
outputs\daily\校内通知活动_YYYY-MM-DD.xlsx
```

每日日志文件：

```text
runtime\logs\daily_YYYY-MM-DD.log
```

删除计划任务：

```powershell
.\uninstall_daily_task.ps1
```

## 校内地点识别

TongZhi 会优先识别与北邮校内相关的地点特征，包括校区、教学楼、主楼、科研楼、经管楼、学生发展中心、学生活动中心、科学会堂、体育馆、图书馆、报告厅、会议室、教室、礼堂、操场、食堂等。

如果通知中出现新的常用场地名，可以在 [bupt_notice_crawler.py](bupt_notice_crawler.py) 的 `CAMPUS_LOCATION_KEYWORDS` 中补充关键词。

## 安全与隐私

- 程序不会保存账号或密码。
- 登录过程由用户在浏览器中手动完成。
- 浏览器登录状态只保存在本机 `runtime/chrome-profile`。
- 不要将 `runtime/`、`outputs/`、`dist/`、`.venv/` 等本地产物提交到公开仓库。

## 开发

创建虚拟环境并安装依赖：

```powershell
py -3.14 -m venv .venv
.\.venv\Scripts\python.exe -m pip install --upgrade pip
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
```

运行：

```powershell
.\run.ps1 -Pages 3
```

## 适用边界

TongZhi 依赖通知页面结构、正文格式和关键词规则。若门户页面结构变化，或通知正文缺少明确时间、地点、部门信息，提取结果可能需要人工复核。

## 许可证

本项目采用 MIT License。版权声明见 [LICENSE](LICENSE)。
