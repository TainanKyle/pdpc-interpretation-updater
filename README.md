# PDPC Administrative Interpretation Updater

A desktop tool that scrapes the **administrative interpretations (行政函釋)** for
each article of Taiwan's Personal Data Protection Act from the Personal Data
Protection Commission (Preparatory Office) website
(<https://www.pdpc.gov.tw/News_Html/100/>), compares them against your master
Excel workbook using the **document number (發文字號)** as the unique key, and
**automatically appends any interpretations that are on the website but missing
from your workbook** to the matching article. The original file is backed up
before anything is written.

The user interface is in Traditional Chinese; button labels below are quoted as
they appear in the app.

## For users: how to run

### Option A — Windows self-build package (nothing to install, recommended)

1. Download `PDPC-tool-selfbuild.zip` from the
   [latest release](https://github.com/TainanKyle/pdpc-interpretation-updater/releases/latest).
2. Extract it and follow `說明.txt` inside: double-click `1-打包成exe.bat` to
   build `PDPC函釋更新工具.exe` on your own PC (no installation, admin rights or
   network needed). Building locally avoids the "downloaded executable" blocks
   common on corporate PCs; if the `.exe` is still blocked, use
   `2-直接執行.bat` to run the tool without an `.exe`.

On macOS, use Option B, or build an `.app` yourself with `build_macos.sh`
(see [Building executables](#building-executables)).

### Using the tool

Steps:

1. Click **「選擇 Excel 檔案…」** (Choose Excel file) and pick your master
   workbook (an `.xlsx` with 「條號」 and 「最新函釋」 columns).
2. The window shows **「✓ 已辨識欄位」** (columns detected). Check that it looks
   right.
3. Click **「試跑（不更新檔案）」** (Dry run) first to see which interpretations
   would be added. **The file is not touched.**
4. When satisfied, click **「開始更新」** (Start update). The program will:
   - create a **「備份」** (backup) folder next to your workbook and save the
     original as `<your file name>_backup_<date-time>.xlsx`
   - append the new interpretations to the 「最新函釋」 column of the matching
     articles
5. When finished, an **「更新完成」** (Update complete) dialog shows how many
   entries were added and the **new total number of interpretations**. The log
   area at the bottom lists every added interpretation with a link to its
   article page on the website. Log text can be **selected and copied** (for
   example, to paste a link into your browser): select it and press **Cmd+C**
   (macOS) / **Ctrl+C** (Windows), or use the menu **「編輯 → 複製」**
   (Edit → Copy).

> First-launch notes (only for an executable built on another computer and then
> copied or downloaded; one you build yourself is not affected):
> - **macOS**: if you see "cannot be opened because it is from an unidentified
>   developer", go to **System Settings → Privacy & Security** and click
>   **Open Anyway**.
> - **Windows**: if SmartScreen blocks it, click **More info → Run anyway**.
>   (This is normal for executables without a code-signing certificate; it is
>   not a virus.)

### Option B — Run with Python

```bash
pip install -r requirements.txt
python app.py          # open the GUI
# or a quick command-line check of the scraper:
python core.py
```

## Workbook format

The tool detects columns from the first (header) row. It only needs a
**「條號」** (article number) column and a **「最新函釋」** (latest
interpretations) column. The **「條文內容」** (article text) column in between
is never modified. See `template.xlsx` for an example.

| 條號 | 條文內容 | 最新函釋(共147條) |
|------|----------|-------------------|
| 第2條 | 本法用詞，定義如下：… | `115.1.23個人資料保護委員會籌備處個資籌法字第1150000083號` |

- **One row per article**; all interpretations for that article go in the
  「最新函釋」 cell, **one per line**.
- Each line is **issue date + document number**, matching the website's format
  (e.g. `115.1.23個人資料保護委員會籌備處個資籌法字第1150000083號`; dates use the
  ROC/Minguo calendar).
- Matching uses **only the document number** (date formatting and full-width /
  half-width differences are ignored), so existing entries are never duplicated.
- The 「條文內容」 column is **never modified**; it is left empty on newly added
  article rows.
- Articles that exist on the website but not in the workbook are appended as a
  new row at the end (「條號」 is written as `第X條`).
- The **「(共N條)」** suffix in the 「最新函釋」 header is **recalculated on every
  update**: it counts distinct document numbers across the whole sheet (a
  document number cited under several articles counts once), so you never need
  to maintain it by hand.

## For developers

### Files

| File | Purpose |
|------|---------|
| `core.py` | Shared core: scraping, parsing, Excel comparison and write-back (no GUI dependencies) |
| `app.py` | Desktop GUI shell (tkinter); only handles file selection and triggering |
| `requirements.txt` | Runtime dependencies: `requests`, `openpyxl`, `truststore` (validates TLS against the OS certificate store so corporate TLS-inspecting proxies work) |
| `build_macos.sh` / `build_windows.bat` | Per-platform packaging scripts |
| `make_windows_bundle.sh` / `bundle/` | Offline Windows bundle that lets users on locked-down PCs build the `.exe` themselves |
| `template.xlsx` | Example master workbook |

### Building executables

PyInstaller **cannot cross-compile**; build once on each target OS.

- **macOS**: `bash build_macos.sh` → `dist/PDPC函釋更新工具.app`
- **Windows**: run `build_windows.bat` on Windows → `dist\PDPC函釋更新工具.exe`
- **Windows self-build package**: `bash make_windows_bundle.sh` (on macOS or
  Linux) → `dist-windows/PDPC函釋工具-自建包.zip`, which builds the `.exe` on
  the target Windows PC (published on Releases as `PDPC-tool-selfbuild.zip`)

### How it works

- The list page is a SPA, but its HTML embeds a `DefaultData` JSON blob
  containing every article. The program discovers **dynamically** all articles
  that have an 【行政函釋】 link (nothing is hard-coded, so new articles are
  picked up automatically).
- Interpretation pages are server-side rendered but break HTML tree parsers, so
  the fixed `發文字號／發文日期／要旨` markers are parsed directly with regular
  expressions.
- The site's JSON API requires a CSRF token and frequently returns 500, so it is
  **not used**.

## License

Copyright (C) 2026 Kyle Huang

This program is free software: you can redistribute it and/or modify it under
the terms of the GNU General Public License as published by the Free Software
Foundation, either version 3 of the License, or (at your option) any later
version. See [LICENSE](LICENSE) for the full text.
