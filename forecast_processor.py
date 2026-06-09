#!/usr/bin/env python3
"""
Hybrid Excel + Python forecast processor for Quotey forecast exports.

Purpose
-------
Reads the raw Export tab from the existing Quotey forecast workbook, moves the
row-by-row forecast calculations into Python, and produces a simplified Excel
workbook for review.

Default input assumptions
-------------------------
- Input workbook has an `Export` worksheet.
- The processor auto-detects whether headers start on row 1 or row 3.
- The workbook may have a quarter-end date in Export!B2 when --quarter-end is omitted.

Example
-------
python forecast_processor.py "Quotey Forecast - Copy.xlsx" --output "Quotey_Forecast_Clean.xlsx"
python forecast_processor.py input.xlsx --quarter-end 2026-06-27 --output output.xlsx
"""

from __future__ import annotations

import argparse
import re
from dataclasses import dataclass
from pathlib import Path
from typing import Dict, Iterable, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook

BACKGROUND_IMAGE_PATH = Path(
    r"C:\Users\olufemi.amurawaiye\OneDrive - Thermo Fisher Scientific\Documents\Shades of Grey\Shades of Grey\Grey 44.PNG"
)

INTERNAL_HELP_COLUMNS = {
    "# of Days in Quarter",
    "In Quarter Revenue",
}

DATE_COLUMNS = [
    "Oppty Month",
    "Hdr Start",
    "Hdr End",
    "Line Start",
    "Line End",
    "Created On",
    "Changed On",
    "Last Contact Date",
    "ValidTo End",
    "closed or declined date",
    "Prior Cvr Line Start",
    "Prior Cvr Line End",
]

MONEY_COLUMNS = [
    "1YR Value",
    "line amt usd",
    "Quote Total USD",
    "In Quarter Revenue",
    "Expected 1YR Value",
    "Expected In Quarter Revenue",
]

HEADER_MARKERS = {
    "Oppty Month",
    "Quote#",
    "Line#",
    "SoldTo Name",
    "Sales Rep Name",
    "In Forecast",
}

COLUMN_ALIASES = {
    "Oppty Month": ["OpptyMonth"],
    "CustGrp 1 Desc": ["CustGrp1 Desc"],
    "Sales Region": ["Sls Region"],
    "Sales Territory": ["Sls Territory"],
    "Line#": ["#"],
    "Line Start": ["Hdr Start", "Header Start Period"],
    "Line End": ["Hdr End"],
    "1YR Value": ["Line Sum USD", "line amt usd", "Quote Total USD"],
}

OUTPUT_DETAIL_COLUMNS = [
    "Sales Region",
    "Sales Territory",
    "Sales Rep Name",
    "SoldTo Name",
    "Quote#",
    "Line#",
    "In Forecast",
    "Big Deal",
    "CustGrp 1 Desc",
    "CustGrp 2 Desc",
    "Win Probability",
    "1YR Value",
    "In Quarter Revenue",
    "Expected 1YR Value",
    "Expected In Quarter Revenue",
    "Quote Total USD",
    "Oppty Month",
    "Line Start",
    "Line End",
    "ValidTo End",
    "Aged Days",
    "Status",
    "model",
    "Serial#",
    "CP Email",
]


@dataclass
class ForecastConfig:
    input_file: Path
    output_file: Path
    export_sheet: str = "Export"
    header_row_1_based: int = 0
    data_start_row_1_based: int = 0
    quarter_end: Optional[pd.Timestamp] = None
    quarter_start: Optional[pd.Timestamp] = None


def parse_args() -> ForecastConfig:
    parser = argparse.ArgumentParser(description="Create a simplified forecast workbook from a Quotey export.")
    parser.add_argument("input_file", help="Source workbook with an Export tab")
    parser.add_argument("--output", "-o", default="Quotey_Forecast_Clean.xlsx", help="Output workbook path")
    parser.add_argument("--export-sheet", default="Export", help="Name of source worksheet")
    parser.add_argument("--header-row", type=int, default=0, help="1-based row containing column headers. Use 0 to auto-detect.")
    parser.add_argument("--data-start-row", type=int, default=0, help="1-based first data row. Use 0 to infer from header row.")
    parser.add_argument("--quarter-end", default=None, help="Quarter-end date, YYYY-MM-DD. Defaults to Export!B2")
    parser.add_argument("--quarter-start", default=None, help="Optional quarter-start date, YYYY-MM-DD. Used for overlap mode")
    args = parser.parse_args()

    quarter_end = pd.to_datetime(args.quarter_end) if args.quarter_end else None
    quarter_start = pd.to_datetime(args.quarter_start) if args.quarter_start else None

    return ForecastConfig(
        input_file=Path(args.input_file),
        output_file=Path(args.output),
        export_sheet=args.export_sheet,
        header_row_1_based=args.header_row,
        data_start_row_1_based=args.data_start_row,
        quarter_end=quarter_end,
        quarter_start=quarter_start,
    )


def normalize_header(value: object, index: int) -> str:
    if value is None or str(value).strip() == "":
        return f"Unused_{index}"
    return str(value).strip()


def detect_header_row(raw: pd.DataFrame) -> int:
    best_row_index = 0
    best_score = -1
    scan_limit = min(len(raw), 10)

    for row_index in range(scan_limit):
        values = {normalize_header(value, idx + 1) for idx, value in enumerate(raw.iloc[row_index].tolist())}
        score = len(values & HEADER_MARKERS)
        if score > best_score:
            best_score = score
            best_row_index = row_index

    return best_row_index


def safe_series(df: pd.DataFrame, column: str, default: object = 0) -> pd.Series:
    if column in df.columns:
        return df[column]
    return pd.Series([default] * len(df), index=df.index)


def apply_column_aliases(df: pd.DataFrame) -> pd.DataFrame:
    renamed = df.copy()
    for canonical, aliases in COLUMN_ALIASES.items():
        if canonical in renamed.columns:
            continue
        for alias in aliases:
            if alias in renamed.columns:
                renamed = renamed.rename(columns={alias: canonical})
                break
    return renamed


def infer_quarter_end_from_dates(df: pd.DataFrame) -> pd.Timestamp:
    for column_name in ("Line Start", "Oppty Month", "Hdr Start"):
        if column_name not in df.columns:
            continue
        dates = pd.to_datetime(df[column_name], errors="coerce").dropna()
        if dates.empty:
            continue
        anchor = dates.max().normalize()
        quarter_month = ((anchor.month - 1) // 3 + 1) * 3
        quarter_end = pd.Timestamp(year=anchor.year, month=quarter_month, day=1) + pd.offsets.MonthEnd(0)
        return quarter_end
    raise ValueError(
        "Quarter-end date was not provided, Export!B2 is blank, and the workbook did not contain usable date columns for inference."
    )


def build_commission_quarter_windows(year: int) -> list[Tuple[pd.Timestamp, pd.Timestamp, str]]:
    jan_first = pd.Timestamp(year=year, month=1, day=1)
    q1_start = jan_first - pd.Timedelta(days=(jan_first.dayofweek + 1) % 7)
    q2_start = q1_start + pd.Timedelta(weeks=13)
    q3_start = q2_start + pd.Timedelta(weeks=13)
    q4_start = q3_start + pd.Timedelta(weeks=13)

    starts = [q1_start, q2_start, q3_start, q4_start]
    ends = [
        q2_start - pd.Timedelta(days=1),
        q3_start - pd.Timedelta(days=1),
        q4_start - pd.Timedelta(days=1),
        pd.Timestamp(year=year, month=12, day=31),
    ]
    return [(start, end, f"Q{index} {year}") for index, (start, end) in enumerate(zip(starts, ends), start=1)]


def find_commission_quarter_window(anchor_date: pd.Timestamp) -> Tuple[pd.Timestamp, pd.Timestamp, str]:
    if pd.isna(anchor_date):
        raise ValueError("Commission quarter could not be determined because the anchor date was blank.")

    normalized = pd.Timestamp(anchor_date).normalize()
    candidate_years = [normalized.year - 1, normalized.year, normalized.year + 1]
    for year in candidate_years:
        for start, end, label in build_commission_quarter_windows(year):
            if start <= normalized <= end:
                return start, end, label

    raise ValueError(
        f"No commission-quarter window matched anchor date {normalized.date()}. "
        "Update the commission-calendar logic before processing this workbook."
    )


def determine_row_quarter_anchor(row: pd.Series) -> pd.Timestamp:
    for column_name in ("Line Start", "Hdr Start", "Oppty Month"):
        value = row.get(column_name)
        if pd.notna(value):
            return pd.Timestamp(value)
    return pd.NaT


def read_quarter_end(config: ForecastConfig) -> pd.Timestamp:
    if config.quarter_end is not None:
        return config.quarter_end
    workbook = load_workbook(config.input_file, data_only=True, read_only=True)
    worksheet = workbook[config.export_sheet]
    value = worksheet["B2"].value
    workbook.close()
    quarter_end = pd.to_datetime(value, errors="coerce")
    if pd.notna(quarter_end):
        return quarter_end

    export_df = read_export(config)
    return infer_quarter_end_from_dates(export_df)


def read_export(config: ForecastConfig) -> pd.DataFrame:
    raw = pd.read_excel(
        config.input_file,
        sheet_name=config.export_sheet,
        header=None,
        engine="openpyxl",
    )
    if config.header_row_1_based and config.header_row_1_based > 0:
        header_index = config.header_row_1_based - 1
    else:
        header_index = detect_header_row(raw)

    if config.data_start_row_1_based and config.data_start_row_1_based > 0:
        data_index = config.data_start_row_1_based - 1
    else:
        data_index = header_index + 1

    headers = [normalize_header(value, idx + 1) for idx, value in enumerate(raw.iloc[header_index].tolist())]
    df = raw.iloc[data_index:].copy()
    df.columns = headers
    df = apply_column_aliases(df)

    # Drop completely empty rows and unneeded blank export columns.
    df = df.dropna(how="all")
    df = df[[column for column in df.columns if not column.startswith("Unused_")]]

    for column in DATE_COLUMNS:
        if column in df.columns:
            df[column] = pd.to_datetime(df[column], errors="coerce")

    for column in ["1YR Value", "line amt usd", "Quote Total USD", "Aged Days", "Days"]:
        if column in df.columns:
            df[column] = pd.to_numeric(df[column], errors="coerce")

    for column in ["Quote#", "Line#", "ShipTo#", "SoldTo#", "Serial#"]:
        if column in df.columns:
            df[column] = df[column].astype("string").fillna("").str.replace(r"\.0$", "", regex=True)

    return df


def probability_from_stage(value: object) -> float:
    """Extract probability from stage text such as `Working - 65%`. Defaults conservatively."""
    if value is None or pd.isna(value):
        return 0.0
    text = str(value)
    match = re.search(r"(\d{1,3})\s*%", text)
    if match:
        pct = min(max(int(match.group(1)), 0), 100)
        return pct / 100.0
    lowered = text.lower()
    if "reviewed" in lowered or "closed" in lowered:
        return 1.0
    if "working" in lowered:
        return 0.65
    return 0.0


def calculate_days_remaining(line_start: pd.Timestamp, quarter_end: pd.Timestamp) -> float:
    """Matches the workbook's simple DATEDIF(Line Start, quarter_end, 'D') behavior."""
    if pd.isna(line_start) or pd.isna(quarter_end):
        return 0.0
    days = (quarter_end.normalize() - line_start.normalize()).days
    return float(max(days, 0))


def calculate_days_overlap(line_start: pd.Timestamp, line_end: pd.Timestamp, quarter_start: pd.Timestamp, quarter_end: pd.Timestamp) -> float:
    """Optional more precise mode: count service days that overlap the selected quarter window."""
    if pd.isna(line_start) or pd.isna(line_end) or pd.isna(quarter_start) or pd.isna(quarter_end):
        return 0.0
    start = max(line_start.normalize(), quarter_start.normalize())
    end = min(line_end.normalize(), quarter_end.normalize())
    if end < start:
        return 0.0
    return float((end - start).days + 1)


def add_forecast_fields(df: pd.DataFrame, quarter_end: pd.Timestamp, quarter_start: Optional[pd.Timestamp]) -> pd.DataFrame:
    out = df.copy()
    one_year_value = pd.to_numeric(safe_series(out, "1YR Value", 0), errors="coerce").fillna(0)

    if quarter_start is not None:
        effective_quarter_start = pd.Timestamp(quarter_start).normalize()
        effective_quarter_end = pd.Timestamp(quarter_end).normalize()
        out["Quarter Start"] = effective_quarter_start
        out["Quarter End"] = effective_quarter_end
        out["Commission Quarter"] = f"Manual {effective_quarter_start.date()} to {effective_quarter_end.date()}"
        out["# of Days in Quarter"] = [
            calculate_days_overlap(start, end, effective_quarter_start, effective_quarter_end)
            for start, end in zip(safe_series(out, "Line Start", pd.NaT), safe_series(out, "Line End", pd.NaT))
        ]
        out["Calculation Mode"] = "Manual quarter window overlap"
    else:
        quarter_starts: list[pd.Timestamp] = []
        quarter_ends: list[pd.Timestamp] = []
        quarter_labels: list[str] = []
        quarter_days: list[float] = []

        for _, row in out.iterrows():
            line_start = row.get("Line Start", pd.NaT)
            line_end = row.get("Line End", pd.NaT)
            anchor_date = determine_row_quarter_anchor(row)
            if pd.isna(anchor_date):
                quarter_starts.append(pd.NaT)
                quarter_ends.append(pd.NaT)
                quarter_labels.append("Unmapped")
                quarter_days.append(0.0)
                continue
            q_start, q_end, q_label = find_commission_quarter_window(anchor_date)
            quarter_starts.append(q_start)
            quarter_ends.append(q_end)
            quarter_labels.append(q_label)
            quarter_days.append(calculate_days_overlap(line_start, line_end, q_start, q_end))

        out["Quarter Start"] = quarter_starts
        out["Quarter End"] = quarter_ends
        out["Commission Quarter"] = quarter_labels
        out["# of Days in Quarter"] = quarter_days
        out["Calculation Mode"] = "2026 LSG commission-quarter overlap from PDF calendar"

    out["In Quarter Revenue"] = one_year_value / 365.0 * pd.to_numeric(out["# of Days in Quarter"], errors="coerce").fillna(0)
    out["Win Probability"] = safe_series(out, "CustGrp 1 Desc", "").apply(probability_from_stage)
    out["Expected 1YR Value"] = one_year_value * out["Win Probability"]
    out["Expected In Quarter Revenue"] = out["In Quarter Revenue"] * out["Win Probability"]

    if "In Forecast" in out.columns:
        out["Forecast Flag"] = out["In Forecast"].fillna("No").astype(str).str.strip().str.title()
    else:
        out["Forecast Flag"] = "Unknown"

    if "Big Deal" in out.columns:
        out["Big Deal Flag"] = out["Big Deal"].fillna("No").astype(str).str.strip().str.title()
    else:
        out["Big Deal Flag"] = "Unknown"

    return out


def summarize(df: pd.DataFrame, group_cols: Iterable[str]) -> pd.DataFrame:
    cols = [col for col in group_cols if col in df.columns]
    if not cols:
        cols = ["Forecast Flag"] if "Forecast Flag" in df.columns else []
    summary_frame = df.copy()
    summary_frame["Quote#"] = safe_series(summary_frame, "Quote#", "")
    summary_frame["1YR Value"] = pd.to_numeric(safe_series(summary_frame, "1YR Value", 0), errors="coerce").fillna(0)
    summary_frame["In Quarter Revenue"] = pd.to_numeric(safe_series(summary_frame, "In Quarter Revenue", 0), errors="coerce").fillna(0)
    summary_frame["Expected 1YR Value"] = pd.to_numeric(safe_series(summary_frame, "Expected 1YR Value", 0), errors="coerce").fillna(0)
    summary_frame["Expected In Quarter Revenue"] = pd.to_numeric(safe_series(summary_frame, "Expected In Quarter Revenue", 0), errors="coerce").fillna(0)
    summary_frame["Win Probability"] = pd.to_numeric(safe_series(summary_frame, "Win Probability", 0), errors="coerce").fillna(0)
    result = (
        summary_frame.groupby(cols, dropna=False)
        .agg(
            Lines=("Quote#", "count"),
            Quotes=("Quote#", pd.Series.nunique),
            One_Year_Value=("1YR Value", "sum"),
            In_Quarter_Revenue=("In Quarter Revenue", "sum"),
            Expected_One_Year_Value=("Expected 1YR Value", "sum"),
            Expected_In_Quarter_Revenue=("Expected In Quarter Revenue", "sum"),
            Average_Win_Probability=("Win Probability", "mean"),
        )
        .reset_index()
        .sort_values("In_Quarter_Revenue", ascending=False)
    )
    return result


def apply_sheet_background(worksheet) -> None:
    if BACKGROUND_IMAGE_PATH.exists():
        worksheet.set_background(str(BACKGROUND_IMAGE_PATH))


def make_dashboard(writer: pd.ExcelWriter, forecast: pd.DataFrame, quarter_end: pd.Timestamp, quarter_start: Optional[pd.Timestamp]) -> None:
    workbook = writer.book
    worksheet = workbook.add_worksheet("Dashboard")
    writer.sheets["Dashboard"] = worksheet
    apply_sheet_background(worksheet)

    title_fmt = workbook.add_format({"bold": True, "font_size": 18, "font_color": "#1F4E78"})
    section_fmt = workbook.add_format({"bold": True, "font_size": 12, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "border": 1})
    label_fmt = workbook.add_format({"bold": True, "border": 1, "bg_color": "#D9EAF7"})
    value_fmt = workbook.add_format({"border": 1, "num_format": "$#,##0"})
    count_fmt = workbook.add_format({"border": 1, "num_format": "#,##0"})
    pct_fmt = workbook.add_format({"border": 1, "num_format": "0.0%"})
    note_fmt = workbook.add_format({"italic": True, "font_color": "#666666"})

    total_lines = len(forecast)
    total_quotes = forecast["Quote#"].nunique() if "Quote#" in forecast.columns else 0
    in_forecast = forecast[forecast["Forecast Flag"].eq("Yes")]
    not_forecast = forecast[forecast["Forecast Flag"].ne("Yes")]

    worksheet.write("A1", "Quotey Forecast Dashboard", title_fmt)
    if quarter_start is not None:
        worksheet.write("A2", f"Quarter end: {quarter_end.date()}")
        worksheet.write("B2", f"Quarter start: {quarter_start.date()}")
    else:
        worksheet.write("A2", "Commission quarter: row-specific 2026 LSG PDF calendar")
    worksheet.write("A3", "Generated by forecast_processor.py", note_fmt)

    kpis = [
        ("Total lines", total_lines, count_fmt),
        ("Unique quotes", total_quotes, count_fmt),
        ("In Forecast Revenue", in_forecast["In Quarter Revenue"].sum(), value_fmt),
        ("Not in Forecast Revenue", not_forecast["In Quarter Revenue"].sum(), value_fmt),
        ("Expected In-Qtr Revenue", forecast["Expected In Quarter Revenue"].sum(), value_fmt),
        ("Avg Win Probability", forecast["Win Probability"].mean(), pct_fmt),
    ]
    worksheet.write("A5", "KPI", section_fmt)
    worksheet.write("B5", "Value", section_fmt)
    for row, (label, value, fmt) in enumerate(kpis, start=5):
        worksheet.write(row, 0, label, label_fmt)
        worksheet.write(row, 1, value, fmt)

    status = summarize(forecast, ["Forecast Flag"])
    status_start = 5
    status_col = 4
    worksheet.write(status_start - 1, status_col, "Forecast Status", section_fmt)
    for col_idx, col_name in enumerate(status.columns):
        worksheet.write(status_start, status_col + col_idx, col_name, section_fmt)
    money_names = {"One_Year_Value", "In_Quarter_Revenue", "Expected_One_Year_Value", "Expected_In_Quarter_Revenue"}
    for r, row in enumerate(status.itertuples(index=False), start=status_start + 1):
        for c, value in enumerate(row):
            col_name = status.columns[c]
            fmt = value_fmt if col_name in money_names else pct_fmt if col_name == "Average_Win_Probability" else count_fmt if col_name in {"Lines", "Quotes"} else None
            worksheet.write(r, status_col + c, value, fmt)

    top_customers = (
        forecast.groupby("SoldTo Name", dropna=False)
        .agg(In_Quarter_Revenue=("In Quarter Revenue", "sum"), Quotes=("Quote#", pd.Series.nunique))
        .reset_index()
        .sort_values("In_Quarter_Revenue", ascending=False)
        .head(10)
    ) if "SoldTo Name" in forecast.columns else pd.DataFrame()

    start_row = 14
    worksheet.write(start_row, 0, "Top Customers by In-Quarter Revenue", section_fmt)
    if not top_customers.empty:
        for c, col_name in enumerate(top_customers.columns):
            worksheet.write(start_row + 1, c, col_name, section_fmt)
        for r, row in enumerate(top_customers.itertuples(index=False), start=start_row + 2):
            for c, value in enumerate(row):
                fmt = value_fmt if c == 1 else count_fmt if c == 2 else None
                worksheet.write(r, c, value, fmt)

        chart = workbook.add_chart({"type": "bar"})
        end_row = start_row + 1 + len(top_customers)
        chart.add_series({
            "name": "In-Quarter Revenue",
            "categories": ["Dashboard", start_row + 2, 0, end_row, 0],
            "values": ["Dashboard", start_row + 2, 1, end_row, 1],
        })
        chart.set_title({"name": "Top Customers"})
        chart.set_x_axis({"num_format": "$#,##0"})
        chart.set_legend({"none": True})
        worksheet.insert_chart("E15", chart, {"x_scale": 1.35, "y_scale": 1.2})

    worksheet.set_column("A:A", 28)
    worksheet.set_column("B:B", 18)
    worksheet.set_column("C:C", 18)
    worksheet.set_column("E:K", 18)
    worksheet.freeze_panes(5, 0)


def write_output(forecast: pd.DataFrame, config: ForecastConfig, quarter_end: pd.Timestamp) -> None:
    output = config.output_file
    output.parent.mkdir(parents=True, exist_ok=True)

    visible_columns = [col for col in OUTPUT_DETAIL_COLUMNS if col in forecast.columns]
    detail = forecast[visible_columns].copy()

    by_region = summarize(forecast, ["Sales Region", "Sales Territory", "Forecast Flag"])
    by_stage = summarize(forecast, ["CustGrp 1 Desc", "Forecast Flag"])
    by_rep = summarize(forecast, ["Sales Rep Name", "Forecast Flag"])
    by_customer = summarize(forecast, ["SoldTo Name", "Forecast Flag"])

    with pd.ExcelWriter(output, engine="xlsxwriter", datetime_format="yyyy-mm-dd", date_format="yyyy-mm-dd") as writer:
        make_dashboard(writer, forecast, quarter_end, config.quarter_start)
        detail.to_excel(writer, sheet_name="Forecast_Detail", index=False)
        by_region.to_excel(writer, sheet_name="Summary_By_Region", index=False)
        by_stage.to_excel(writer, sheet_name="Summary_By_Stage", index=False)
        by_rep.to_excel(writer, sheet_name="Summary_By_Rep", index=False)
        by_customer.to_excel(writer, sheet_name="Summary_By_Customer", index=False)

        workbook = writer.book
        header_fmt = workbook.add_format({"bold": True, "font_color": "#FFFFFF", "bg_color": "#1F4E78", "border": 1})
        money_fmt = workbook.add_format({"num_format": "$#,##0", "border": 1})
        pct_fmt = workbook.add_format({"num_format": "0.0%", "border": 1})
        date_fmt = workbook.add_format({"num_format": "yyyy-mm-dd", "border": 1})
        default_fmt = workbook.add_format({"border": 1})

        for sheet_name, frame in {
            "Forecast_Detail": detail,
            "Summary_By_Region": by_region,
            "Summary_By_Stage": by_stage,
            "Summary_By_Rep": by_rep,
            "Summary_By_Customer": by_customer,
        }.items():
            worksheet = writer.sheets[sheet_name]
            apply_sheet_background(worksheet)
            worksheet.freeze_panes(1, 0)
            for col_idx, col_name in enumerate(frame.columns):
                worksheet.write(0, col_idx, col_name, header_fmt)
                width = min(max(len(str(col_name)) + 2, 12), 32)
                if col_name in {"SoldTo Name", "Sales Rep Name", "CP Email", "CustGrp 1 Desc"}:
                    width = 28
                worksheet.set_column(col_idx, col_idx, width, default_fmt)
                if col_name in MONEY_COLUMNS or col_name in {"One_Year_Value", "In_Quarter_Revenue", "Expected_One_Year_Value", "Expected_In_Quarter_Revenue"}:
                    worksheet.set_column(col_idx, col_idx, 18, money_fmt)
                elif col_name in {"Win Probability", "Average_Win_Probability"}:
                    worksheet.set_column(col_idx, col_idx, 16, pct_fmt)
                elif col_name in DATE_COLUMNS:
                    worksheet.set_column(col_idx, col_idx, 14, date_fmt)

            # Add Excel table for better filtering and readability.
            if len(frame.columns) > 0:
                table_name = re.sub(r"[^A-Za-z0-9_]", "_", sheet_name)[:25]
                worksheet.add_table(0, 0, max(len(frame), 1), len(frame.columns) - 1, {
                    "name": table_name,
                    "columns": [{"header": str(col)} for col in frame.columns],
                    "style": "Table Style Medium 2",
                })


def main() -> None:
    config = parse_args()
    try:
        quarter_end = read_quarter_end(config)
        source = read_export(config)
        forecast = add_forecast_fields(source, quarter_end, config.quarter_start)
        write_output(forecast, config, quarter_end)
    except PermissionError as exc:
        raise SystemExit(
            "Could not open the workbook. Close the raw export and the output file in Excel, "
            f"then run the tool again.\nFile: {exc.filename}"
        ) from exc

    print(f"Created {config.output_file}")
    print(f"Rows processed: {len(forecast):,}")
    if config.quarter_start is not None:
        print(f"Quarter end: {quarter_end.date()}")
    else:
        print("Quarter end: row-specific 2026 LSG commission calendar")


if __name__ == "__main__":
    main()
