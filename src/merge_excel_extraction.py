#!/usr/bin/env python3
"""Merge all Excel files and all sheets in a folder into one dataset."""

from __future__ import annotations

import argparse
import re
from pathlib import Path

import pandas as pd


def find_excel_files(input_dir: Path, excluded_names: set[str]) -> list[Path]:
    files = sorted(list(input_dir.glob("*.xlsx")) + list(input_dir.glob("*.xls")))
    return [p for p in files if p.name not in excluded_names]


def normalize_columns(columns: pd.Index) -> list[str]:
    normalized: list[str] = []
    seen: dict[str, int] = {}

    for col in columns:
        base = re.sub(r"\s+", " ", str(col).strip()).lower()
        if not base:
            base = "unnamed"

        count = seen.get(base, 0) + 1
        seen[base] = count
        normalized.append(base if count == 1 else f"{base}_{count}")

    return normalized


def merge_workbooks(excel_files: list[Path]) -> pd.DataFrame:
    frames: list[pd.DataFrame] = []

    for file_path in excel_files:
        workbook = pd.ExcelFile(file_path)
        for sheet_name in workbook.sheet_names:
            df = pd.read_excel(file_path, sheet_name=sheet_name)
            df.columns = normalize_columns(df.columns)
            df.insert(0, "source_sheet", sheet_name)
            df.insert(0, "source_workbook", file_path.name)
            df.insert(2, "source_row_number", range(1, len(df) + 1))
            frames.append(df)

    if not frames:
        raise ValueError("No sheet data found in the provided Excel files.")

    merged = pd.concat(frames, ignore_index=True, sort=False)
    if "study number" not in merged.columns:
        raise KeyError("Required column not found after merge: study number")

    # Keep only rows where study number is present and not just whitespace.
    keep_mask = merged["study number"].astype("string").str.strip().ne("")
    keep_mask = keep_mask.fillna(False)
    return merged.loc[keep_mask].reset_index(drop=True)


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Merge all Excel files and all their sheets into one output dataset.",
    )
    parser.add_argument(
        "--input-dir",
        default="data/data_extraction",
        help="Folder containing .xlsx/.xls files (default: data/data_extraction)",
    )
    parser.add_argument(
        "--output-base",
        default="merged_all_sheets",
        help="Output file base name without extension (default: merged_all_sheets)",
    )
    args = parser.parse_args()

    input_dir = Path(args.input_dir)
    if not input_dir.exists() or not input_dir.is_dir():
        raise FileNotFoundError(f"Input directory not found: {input_dir}")

    output_xlsx = input_dir / f"{args.output_base}.xlsx"
    output_csv = input_dir / f"{args.output_base}.csv"

    excluded_names = {output_xlsx.name, output_csv.name}
    excel_files = find_excel_files(input_dir, excluded_names)
    if not excel_files:
        raise FileNotFoundError(
            f"No .xlsx/.xls files found in {input_dir} after exclusions: {sorted(excluded_names)}"
        )

    merged = merge_workbooks(excel_files)
    merged.to_excel(output_xlsx, sheet_name="merged_data", index=False)
    merged.to_csv(output_csv, index=False)

    print(f"Input folder: {input_dir}")
    print(f"Workbook files merged: {len(excel_files)}")
    print(f"Rows: {len(merged)}")
    print(f"Columns: {len(merged.columns)}")
    print(f"Created: {output_xlsx}")
    print(f"Created: {output_csv}")


if __name__ == "__main__":
    main()
