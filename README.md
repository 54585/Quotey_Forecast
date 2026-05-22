# Hybrid Excel + Python Quotey Forecast Processor

This version keeps Excel as the final review/reporting format, but moves the fragile row-by-row calculations out of the workbook and into Python.

## One-click usage

1. Unzip the package.
2. Open the `hybrid_forecast` folder.
3. Put your latest raw Quotey export into the `Raw_Exports` folder.
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
- find the newest `.xlsx` file in `Raw_Exports`,
- run `forecast_processor.py`,
- save the clean workbook into `Outputs`,
- open the finished workbook automatically.

## Recommended folder structure

```text
hybrid_forecast/
  Run_Quotey_Forecast.bat
  forecast_processor.py
  requirements.txt
  Raw_Exports/
    Your latest Quotey export.xlsx
  Outputs/
    Quotey_Forecast_Clean_YYYY-MM-DD.xlsx
```

## Manual run option

You can still run the processor manually:

```bash
python forecast_processor.py "Raw_Exports\Quotey Forecast - Copy.xlsx" --output "Outputs\Quotey_Forecast_Clean.xlsx"
```

By default, the script reads the quarter-end date from `Export!B2`.

If `Export!B2` is blank or contains a non-date token such as `Q03`, the script infers a quarter end from the workbook date columns so the run can still finish.

You can also provide the quarter end directly:

```bash
python forecast_processor.py "Raw_Exports\Quotey Forecast - Copy.xlsx" --quarter-end 2026-06-27 --output "Outputs\Quotey_Forecast_Clean.xlsx"
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
