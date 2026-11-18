"""
Automated Excel Data Cleaning & Standardization Demo
----------------------------------------------------
Demonstrates automated spreadsheet processing using pandas and openpyxl:
- Standardizes column headers to snake_case
- Normalizes dates, emails, whitespace, and text casing
- Removes duplicate rows
- Handles missing values with clear defaults
- Exports cleaned records and an executive Summary tab to Excel
"""

import re
import sys
from datetime import datetime
from pathlib import Path

import pandas as pd
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter


def standardize_header(col: str) -> str:
    cleaned = re.sub(r"[^\w\s]", "", str(col).strip().lower())
    return re.sub(r"\s+", "_", cleaned) or "column"


def parse_date(val):
    if pd.isna(val):
        return pd.NA
    if isinstance(val, (datetime, pd.Timestamp)):
        return val.strftime("%Y-%m-%d")
    text = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m-%d-%Y", "%Y/%m/%d", "%m/%d/%Y"):
        try:
            return datetime.strptime(text, fmt).strftime("%Y-%m-%d")
        except ValueError:
            continue
    parsed = pd.to_datetime(text, errors="coerce")
    return parsed.strftime("%Y-%m-%d") if pd.notna(parsed) else pd.NA


def clean_excel_file(input_path: Path, output_path: Path) -> None:
    excel_file = pd.ExcelFile(input_path, engine="openpyxl")
    first_sheet = excel_file.sheet_names[0]
    df = pd.read_excel(excel_file, sheet_name=first_sheet).dropna(how="all")
    total_read = len(df)

    # 1. Standardize column names
    df.columns = [standardize_header(c) for c in df.columns]

    # 2. Clean whitespace, casing, dates, and emails
    for col in df.columns:
        if pd.api.types.is_object_dtype(df[col]) or pd.api.types.is_string_dtype(df[col]):
            df[col] = df[col].apply(lambda x: " ".join(x.strip().split()) if isinstance(x, str) else x)
            df[col] = df[col].replace({"": pd.NA, "nan": pd.NA, "None": pd.NA, "N/A": pd.NA})

        if "date" in col:
            df[col] = df[col].apply(parse_date)
        elif "email" in col:
            df[col] = df[col].apply(lambda x: x.lower() if isinstance(x, str) else x)
        elif col.endswith("_id"):
            df[col] = df[col].apply(lambda x: x.upper() if isinstance(x, str) else x)
        elif any(k in col for k in ("name", "category", "status", "department")):
            df[col] = df[col].apply(lambda x: x.title() if isinstance(x, str) else x)

    # 3. Remove duplicate rows
    duplicates_removed = int(df.duplicated().sum())
    df = df.drop_duplicates(keep="first").reset_index(drop=True)

    # 4. Handle missing values
    missing_handled = int(df.isna().sum().sum())
    for col in df.columns:
        if pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].fillna(0)
        elif "date" in col or "email" in col:
            df[col] = df[col].fillna("Not Provided")
        else:
            df[col] = df[col].fillna("Unknown")

    if "quantity" in df.columns:
        df["quantity"] = df["quantity"].astype(int)

    # 5. Build processing summary
    summary_df = pd.DataFrame(
        {
            "Metric": [
                "Source File",
                "Worksheet Processed",
                "Total Records Read",
                "Duplicates Removed",
                "Missing Values Handled",
                "Final Clean Records",
            ],
            "Value": [
                input_path.name,
                first_sheet,
                total_read,
                duplicates_removed,
                missing_handled,
                len(df),
            ],
        }
    )

    # 6. Export formatted Excel workbook
    with pd.ExcelWriter(output_path, engine="openpyxl") as writer:
        df.to_excel(writer, sheet_name="Cleaned_Data", index=False)
        summary_df.to_excel(writer, sheet_name="Summary", index=False)

        header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
        header_font = Font(color="FFFFFF", bold=True)
        thin_border = Border(
            left=Side(style="thin", color="D9D9D9"),
            right=Side(style="thin", color="D9D9D9"),
            top=Side(style="thin", color="D9D9D9"),
            bottom=Side(style="thin", color="D9D9D9"),
        )

        for sheet in ("Cleaned_Data", "Summary"):
            ws = writer.sheets[sheet]
            for cell in ws[1]:
                cell.fill = header_fill
                cell.font = header_font
                cell.alignment = Alignment(horizontal="center")
            for row in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=ws.max_column):
                for cell in row:
                    cell.border = thin_border
            for idx, col_cells in enumerate(ws.columns, start=1):
                width = max((len(str(c.value or "")) for c in col_cells), default=10)
                ws.column_dimensions[get_column_letter(idx)].width = min(width + 4, 40)

    print(f"Done! Cleaned {len(df)} records ({duplicates_removed} duplicates removed, {missing_handled} missing cells filled).")


if __name__ == "__main__":
    in_path = Path(sys.argv[1]) if len(sys.argv) > 1 else Path("demo_input.xlsx")
    out_path = Path(sys.argv[2]) if len(sys.argv) > 2 else Path("demo_output.xlsx")
    clean_excel_file(in_path, out_path)
