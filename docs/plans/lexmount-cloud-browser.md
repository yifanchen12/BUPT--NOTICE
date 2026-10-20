# Lexmount Cloud Browser integration plan

## Requirement

Add an opt-in Lexmount CloudBrowser backend to BUPT-NOTICE. Keep the current local Chrome/Edge backend and reuse the existing notice extraction, activity filtering, and Excel export.

## Reuse survey

- `bupt_notice_crawler.py`: already owns browser navigation, CAS login detection, pagination, HTML parsing, filtering, and workbook output. Reuse all parsing/export functions.
- `run.ps1`: existing source launcher can forward a browser-provider option.
- `requirements.txt` and `build_exe.ps1`: current dependency/install and PyInstaller packaging paths.
- No automated test suite or test command is documented in `README.md`; the documented manual entry point is `run.ps1`.

## Files

- `bupt_notice_crawler.py`: add `local`/`lexmount` provider selection; connect Playwright to Lexmount's CDP URL; show the remote inspector URL for manual CAS login; disable local cookie forwarding in cloud mode; reuse the current parser and Excel writer; close the browser session on success and failure.
- `requirements.txt`: add the official Lexmount Python SDK and its documented compatible Playwright version.
- `run.ps1`: forward the selected provider from PowerShell.
- `build_exe.ps1`: collect Lexmount SDK resources for the packaged executable.
- `README.md`, `README.en.md`: document credentials, cloud-mode use, manual login, and the existing local-mode fallback.
- `docs/plans/lexmount-cloud-browser.md`: retain this approved implementation plan.

## Interface

- CLI `--browser-provider {local,lexmount}`, default `local`.
- Lexmount credentials come only from `LEXMOUNT_API_KEY` and `LEXMOUNT_PROJECT_ID`; `LEXMOUNT_BASE_URL` remains optional and follows the SDK default. Never persist or print credentials.
- In Lexmount mode, navigate through a cloud Chromium session, let the user complete CAS login at the session inspector URL, then process the same pages as local mode.
- Each Lexmount invocation uses a fresh session; persistent contexts and unattended scheduled Lexmount runs are out of scope for this change.

## Build order

1. Add provider/session handling and reuse the existing page-processing pipeline.
2. Update launcher and packaging dependencies.
3. Document configuration and limitations.

## Validation

- `python -m py_compile bupt_notice_crawler.py` checks syntax.
- `python bupt_notice_crawler.py --help` confirms CLI compatibility.
- Cloud end-to-end login requires the user's Lexmount API Key, Project ID, and manual CAS login; no credentials will be requested or committed.

## Verified facts and assumptions

- Verified: the current application uses a local persistent Playwright profile and keeps parsing/export in the same Python script.
- Verified: Lexmount's current Python quickstart uses `lexmount==0.6.1` with `playwright==1.63.0` and connects Playwright over CDP.
- Assumption: the user's Lexmount project has CloudBrowser API access and permits the BUPT portal from its network region.
