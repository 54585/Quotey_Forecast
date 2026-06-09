# Hybrid Excel + Python Quotey Forecast Processor

This version keeps Excel as the final review/reporting format, but moves the fragile row-by-row calculations out of the workbook and into Python.

## One-click usage

1. Unzip the package.
2. Open the `Quotey Forecast` folder.
3. Put your latest raw Quotey export into the `Inputs` folder.
   - If the folder does not exist yet, double-click the `.bat` once and it will create it.
   - The workbook must contain a tab named `Export`.
4. Double-click:

```text
Run_Quotey_Forecast.bat
```

The batch file will:

- find a usable local Python runtime,
- create a local `.venv` folder on the first run,
- install required packages from `requirements.txt`,
- auto-detect whether the export headers start on row 1 or row 3,
- normalize common export header variants such as `OpptyMonth`, `CustGrp1 Desc`, `Sls Region`, and `Line Sum USD`,
- apply the background image `Grey 44.PNG` to every generated worksheet,
- include `CustGrp 2 Desc` in `Forecast_Detail` immediately to the right of `CustGrp 1 Desc`,
- find the newest `.xlsx` file in `Inputs`,
- run `forecast_processor.py`,
- save the clean workbook into `Outputs`,
- open the finished workbook automatically.

## Recommended folder structure

```text
Quotey Forecast/
  Run_Quotey_Forecast.bat
  forecast_processor.py
  requirements.txt
  Inputs/
    Your latest Quotey export.xlsx
  Outputs/
    Quotey_Forecast_Clean_YYYY-MM-DD.xlsx
```

## Manual run option

You can still run the processor manually:

```bash
python forecast_processor.py "Inputs\Quotey Forecast - Copy.xlsx" --output "Outputs\Quotey_Forecast_Clean.xlsx"
```

By default, the script now calculates `In Quarter Revenue` using service-date overlap against the LSG commission quarter windows. The current implementation uses Sunday-to-Saturday fiscal quarters, with `Q1` starting on the Sunday on or before January 1. For 2026, `Q4` is explicitly extended through `2026-12-31`.

For the 2026 commission calendar from the PDF, that means:

- `Q1 2026`: `2025-12-28` through `2026-03-28`
- `Q2 2026`: `2026-03-29` through `2026-06-27`
- `Q3 2026`: `2026-06-28` through `2026-09-26`
- `Q4 2026`: `2026-09-27` through `2026-12-31`

Rows are anchored to their service dates first (`Line Start`, then `Hdr Start`, then `Oppty Month`) so late-2025 service dates can still roll into `Q1 2026` correctly.

You can also provide the quarter end directly:

```bash
python forecast_processor.py "Inputs\Quotey Forecast - Copy.xlsx" --quarter-end 2026-06-27 --output "Outputs\Quotey_Forecast_Clean.xlsx"
```

## Output workbook tabs

- `Dashboard`
- `Forecast_Detail`
- `Summary_By_Region`
- `Summary_By_Stage`
- `Summary_By_Rep`
- `Summary_By_Customer`

## Troubleshooting

If the batch file fails, open `run_log.txt` in the same folder. Common causes:

- No usable Python runtime was available on the PC.
- The raw workbook does not have a tab named `Export`.
- The export layout changed and no recognizable header row was found.
- The raw export is open in Excel and locked.
- Corporate network settings blocked package installation.

If package installation is blocked, ask IT to allow these Python packages:

```text
pandas
openpyxl
xlsxwriter
```
