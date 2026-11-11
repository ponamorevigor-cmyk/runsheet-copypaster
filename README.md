# CopyPaster | Excel Runsheet Automation

A small Python utility that eliminates repetitive manual editing in **audio recording and localization runsheets**.

## What it does

For each `.xlsx` file in a folder, CopyPaster:

1. Opens the first worksheet.
2. Copies the values in **SOURCE TEXT (column G)** into **TRANSLATION (column I)**, starting at **row 10**.
3. Saves the result as a new `<original_name>_1.xlsx` file, leaving the original unchanged.
4. Repeats this for the remaining files in the folder.

This is useful when preparing **source-language runsheets** where the translation field intentionally needs to contain the same text as the source. The runsheet also feeds an audio QA workflow.

> **Important:** Do not use this tool on translated runsheets where column I already contains target-language translations that must be retained. The generated copy will replace those values with source text.

## Why I built it

A typical localization project involved **approximately 220 runsheet files**, received in batches of roughly **20 files**. The manual process required opening each file, copying text between columns, saving a separate copy and moving to the next file.

| Workflow | Approximate time per 20-file batch |
| --- | ---: |
| Manual editing | 20 minutes |
| Automated processing | 2 minutes |
| **Estimated reduction** | **90%** |

These are practical estimates, not formally benchmarked timings. The original tool was adopted by **two additional project managers** and included in the updated project workflow instructions.

## Requirements

- Windows with desktop **Microsoft Excel** installed
- Python 3.10+ and [`xlwings`](https://docs.xlwings.org/) when running the Python source
- No Python installation needed when running a separately packaged Windows executable

Install the dependency:

```bash
python -m pip install -r requirements.txt
```

## Run

1. Put `CopyPaster.py` in a folder containing the runsheets (`.xlsx`).
2. Run `python CopyPaster.py` from that folder, or place and launch a packaged `CopyPaster.exe` alongside those files.
3. Review the generated files with suffix `_1.xlsx`.

```text
runsheets/
├── CopyPaster.py
├── lesson_01.xlsx
├── lesson_02.xlsx
├── lesson_01_1.xlsx   # generated
└── lesson_02_1.xlsx   # generated
```

## Safeguards in this public-source version

- Does not overwrite the original workbooks or pre-existing output files.
- Skips Excel temporary (`~$...`) files and already-generated `_1.xlsx` files.
- Copies values in one block rather than cell by cell.
- Handles errors per file and closes Excel when processing finishes.

The script is intentionally tailored to a specific runsheet layout: **first worksheet, G → I, starting at row 10**. Use it on copies or test files before processing other workbook templates.

## Build an optional Windows executable

```bash
python -m pip install -r requirements.txt
python -m pip install pyinstaller
python -m PyInstaller --onefile --console --name CopyPaster CopyPaster.py
```

The executable will be under `dist/CopyPaster.exe`. Test it with non-confidential sample files before distribution.

## Notes

This repository contains **no real client runsheets**. The original utility was used in a production localization workflow; the revised public-source version contains reliability improvements and has **not** been end-to-end tested with Microsoft Excel in this environment.
