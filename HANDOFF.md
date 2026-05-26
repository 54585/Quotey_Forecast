# Hybrid Forecast Handoff

Use this file as the durable project map for Codex sessions in this folder. Keep it current enough that a new session can safely understand how to run, modify, test, commit, and hand off the project without relying on chat history.

## Project Snapshot

- Project name: `hybrid_forecast`
- Local path: `C:\Users\olufemi.amurawaiye\OneDrive - Thermo Fisher Scientific\Desktop\Codex Programs\hybrid_forecast`
- Primary entry point: `Run_Quotey_Forecast.bat`
- Main run command: Double-click `Run_Quotey_Forecast.bat`
- Main test or self-check command: `.\.venv\Scripts\python.exe .\forecast_processor.py ".\Inputs\Test Exports.xlsx" --output ".\Outputs\test_output.xlsx"`
- Runtime and key dependencies: Local `.venv` plus `pandas`, `openpyxl`, `xlsxwriter`
- UI expectation: This project is a batch-driven desktop utility, not a Tkinter app.
- Local Git repository: Initialized on 2026-05-20.
- Git remote: Not configured yet. Ask the user before adding one or pushing.

## Maintenance Rule

Whenever Codex makes a program update in this project, it must update all three maintenance artifacts in the same work session:

- `README.md` with user-facing behavior or runbook changes.
- `HANDOFF.md` with implementation details, migrations, safety rules, and next-session notes.
- Local Git tracking through the project-root `.git/` repository.
- Git history with a commit that captures the source and documentation changes, unless the user explicitly asks not to commit yet.

After every three committed program updates, Codex must ask the user to confirm the Git remote for this project, then push only the tracked program and documentation files to that confirmed remote. Do not include local runtime state, secrets, backup files, generated exports, logs, or machine-specific cache files in that push.

## Session Start Checklist

1. Read `README.md` and `HANDOFF.md` before changing code.
2. Confirm the current folder is `hybrid_forecast`.
3. Check `git status`, current branch, and configured remote.
4. Keep `.venv/`, `Outputs/`, `Inputs/*.xlsx`, and `run_log.txt` out of commits unless the user explicitly asks for a tracked sample file.
5. Identify the safest validation path before editing. Prefer the local `Test Exports.xlsx` workbook if it is available and not open in Excel.
6. Preserve unrelated user changes. Do not reset, delete, or overwrite files unless the user explicitly asks.

## Program Update Checklist

1. Inspect the relevant code and current runtime state.
2. Make the smallest change that fully handles the request.
3. Run focused validation with the local `.venv`.
4. Update `README.md` for behavior, setup, or runbook changes.
5. Update `HANDOFF.md` for implementation details, safety rules, unresolved risks, and next-session notes.
6. Stage only source and documentation files that belong in Git.
7. Commit source and documentation changes unless the user says not to commit yet.
8. Update the committed-update counter below.
9. When the counter reaches 3, confirm the Git remote with the user, push tracked source and docs, then reset the counter to 0.

## Committed-Update Counter

- Committed program updates since last push: 1
- Last push date: 2026-05-22
- Last confirmed Git remote: `https://github.com/54585/hybrid_forecast.git`

## Local State And Exclusions

- `.venv/` is local runtime state and should not be committed.
- `Outputs/` contains generated workbooks and should not be committed, except for the placeholder text file.
- `Inputs/*.xlsx` contains user data exports and should not be committed.
- `run_log.txt` is a local troubleshooting file and should not be committed.

## Safety Rules

- Ask before pushing to any remote.
- Ask before deleting or overwriting raw export files.
- Keep customer exports, generated workbooks, and logs out of Git.
- If the workbook is open in Excel and locked, close it first rather than working around the lock.

## Implementation Notes

- `Run_Quotey_Forecast.bat` now prefers a real local Python runtime, skips the Windows Store `python.exe` stub, creates a local `.venv`, installs requirements, and uses PowerShell `Get-Date -Format yyyy-MM-dd` for a stable output filename.
- The input folder was renamed from `Raw_Exports` to `Inputs` on 2026-05-26. Launcher, docs, Git ignore rules, and sample commands were updated to use `Inputs`.
- `forecast_processor.py` now auto-detects whether the `Export` sheet headers start on row 1 or row 3. The sample workbook `Inputs\Test Exports.xlsx` uses row 1 headers.
- `forecast_processor.py` no longer crashes when expected columns are missing; it uses safe default Series values instead of scalar fallbacks.
- `forecast_processor.py` now normalizes common alternate export headers into the canonical internal names. Verified examples from the 2026-05-22 export: `OpptyMonth`, `CustGrp1 Desc`, `Sls Region`, `Sls Territory`, `#`, and `Line Sum USD`.
- Every generated worksheet now applies the background image `C:\Users\olufemi.amurawaiye\OneDrive - Thermo Fisher Scientific\Documents\Shades of Grey\Shades of Grey\Grey 44.PNG`.
- If `Export!B2` is blank or contains a non-date token such as `Q03`, the processor infers a quarter-end date from workbook date columns so the run can still complete.
- Verified on 2026-05-20 with `Test Exports.xlsx`: the processor created `Outputs\test_output.xlsx` successfully and processed 760 rows.
- Verified on 2026-05-22 with `data - 2026-05-22T172921.787.xlsx`: the processor created `Outputs\repro_2026-05-22.xlsx` successfully and processed 387 rows.
- Verified on 2026-05-22 with `Outputs\background_check_2026-05-22.xlsx`: the workbook package contains `xl/media/image1.png`, and the worksheet XML files contain `<picture>` references, confirming the background image is embedded.
- Verified on 2026-05-26 that `Run_Quotey_Forecast.bat` now reads from `Inputs\...`; the test run stopped only because `Outputs\Quotey_Forecast_Clean_2026-05-26.xlsx` was already locked and could not be overwritten.

## Important Files

- `Run_Quotey_Forecast.bat`: one-click launcher
- `forecast_processor.py`: workbook parsing and output generation
- `requirements.txt`: Python dependencies
- `README.md`: end-user run instructions
- `.gitignore`: local-runtime and export exclusions

## Next-Session Notes

- Re-check the inferred quarter-end behavior with the user if revenue totals look off. The current fallback uses workbook date columns when `Export!B2` is blank or contains a non-date token like `Q03`.
- If the user wants this published, ask: `What Git address should this project use?`
- Git remote confirmed on 2026-05-22 as `https://github.com/54585/hybrid_forecast.git`.
