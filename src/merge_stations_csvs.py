"""
TempO_UA — merge station CSV exports

What it does:
- reads all CSV files exported from the TempO app for separate stations;
- normalizes column names;
- merges them into one table;
- saves combined results as CSV and Excel.

Expected input folder:
  data/raw/station_exports/

Expected output folder:
  data/processed/

Run from the project root:
  python src/merge_station_csvs.py

Demo/test run without real testers:
  python src/merge_station_csvs.py --demo
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, List, Tuple

import pandas as pd


DEFAULT_INPUT_DIR = Path("data/raw/station_exports")
DEFAULT_OUTPUT_DIR = Path("data/processed")


COLUMN_ALIASES = {
    "station": "Station",
    "станція": "Station",
    "stanciya": "Station",
    "stationnumber": "Station",
    "station_number": "Station",
    "спортсмен": "Athlete",
    "athlete": "Athlete",
    "athletename": "Athlete",
    "athlete_name": "Athlete",
    "імя": "Athlete",
    "ім'я": "Athlete",
    "ім’я": "Athlete",
    "номер": "Athlete",
    "time_sec": "Time_sec",
    "час_сек": "Time_sec",
    "час": "Time_sec",
    "time": "Time_sec",
    "stationtime": "Time_sec",
    "station_time": "Time_sec",
    "corrected": "Corrected",
    "виправлено": "Corrected",
}


def normalize_key(value: str) -> str:
    """Make column names comparable."""
    value = str(value).strip().lower()
    value = value.replace(" ", "_").replace("-", "_")
    value = value.replace("’", "'")
    value = re.sub(r"[^a-zа-яіїєґ0-9_']+", "", value)
    return value


def detect_answer_column(column_name: str) -> Tuple[bool, int | None]:
    """Detect task answer columns like Завдання_1, Task_1, Answer_1."""
    raw = str(column_name).strip()
    key = normalize_key(raw)

    patterns = [
        r"^завдання_?(\d+)$",
        r"^task_?(\d+)$",
        r"^answer_?(\d+)$",
        r"^q_?(\d+)$",
    ]

    for pattern in patterns:
        match = re.match(pattern, key)
        if match:
            return True, int(match.group(1))

    return False, None


def read_csv_safely(path: Path) -> pd.DataFrame:
    """Read CSV exported from browser/Excel with a few common encodings/delimiters."""
    last_error = None

    for encoding in ["utf-8-sig", "utf-8", "cp1251"]:
        try:
            return pd.read_csv(path, sep=None, engine="python", encoding=encoding)
        except Exception as exc:  # keep trying other encodings
            last_error = exc

    raise RuntimeError(f"Could not read {path}. Last error: {last_error}")


def standardize_columns(df: pd.DataFrame, source_file: str) -> pd.DataFrame:
    """Rename known columns to a stable format and normalize answer columns."""
    rename_map: Dict[str, str] = {}

    for col in df.columns:
        key = normalize_key(col)

        if key in COLUMN_ALIASES:
            rename_map[col] = COLUMN_ALIASES[key]
            continue

        is_answer, task_num = detect_answer_column(col)
        if is_answer and task_num is not None:
            rename_map[col] = f"Task_{task_num}"

    df = df.rename(columns=rename_map).copy()
    df["Source_file"] = source_file

    required = ["Station", "Athlete", "Time_sec"]
    missing = [col for col in required if col not in df.columns]
    if missing:
        raise ValueError(
            f"File {source_file} is missing required columns: {missing}. "
            f"Found columns: {list(df.columns)}"
        )

    if "Corrected" not in df.columns:
        df["Corrected"] = "no"

    df["Station"] = pd.to_numeric(df["Station"], errors="coerce").astype("Int64")
    df["Time_sec"] = pd.to_numeric(df["Time_sec"], errors="coerce")
    df["Athlete"] = df["Athlete"].astype(str).str.strip()
    df["Corrected"] = df["Corrected"].fillna("no").astype(str).str.lower().str.strip()

    # Clean answer columns to uppercase single-letter values where possible.
    for col in df.columns:
        if re.match(r"^Task_\d+$", col):
            df[col] = df[col].fillna("").astype(str).str.strip().str.upper()

    return df


def collect_task_columns(frames: List[pd.DataFrame]) -> List[str]:
    task_numbers = set()
    for df in frames:
        for col in df.columns:
            match = re.match(r"^Task_(\d+)$", col)
            if match:
                task_numbers.add(int(match.group(1)))
    return [f"Task_{n}" for n in sorted(task_numbers)]


def merge_station_exports(input_dir: Path, output_dir: Path) -> pd.DataFrame:
    files = sorted(input_dir.glob("*.csv"))

    if not files:
        raise FileNotFoundError(
            f"No CSV files found in {input_dir}. "
            "Put station exports there, for example station_1.csv, station_2.csv."
        )

    frames: List[pd.DataFrame] = []
    for path in files:
        raw = read_csv_safely(path)
        standardized = standardize_columns(raw, source_file=path.name)
        frames.append(standardized)

    task_cols = collect_task_columns(frames)

    base_cols = ["Station", "Athlete", "Time_sec"]
    end_cols = ["Corrected", "Source_file"]
    final_cols = base_cols + task_cols + end_cols

    aligned_frames = []
    for df in frames:
        for col in task_cols:
            if col not in df.columns:
                df[col] = ""
        aligned_frames.append(df[final_cols])

    merged = pd.concat(aligned_frames, ignore_index=True)
    merged = merged.sort_values(by=["Station", "Athlete"], kind="stable")

    output_dir.mkdir(parents=True, exist_ok=True)
    csv_path = output_dir / "merged_station_results.csv"
    xlsx_path = output_dir / "merged_station_results.xlsx"

    merged.to_csv(csv_path, index=False, encoding="utf-8-sig")

    try:
        with pd.ExcelWriter(xlsx_path) as writer:
            merged.to_excel(writer, index=False, sheet_name="Merged results")
    except Exception as exc:
        print(f"Excel export failed: {exc}")
        print(f"CSV was still saved to: {csv_path}")
    else:
        print(f"Saved Excel file to: {xlsx_path}")

    print(f"Saved CSV file to: {csv_path}")
    print(f"Merged files: {len(files)}")
    print(f"Merged rows: {len(merged)}")

    return merged


def create_demo_files(input_dir: Path) -> None:
    input_dir.mkdir(parents=True, exist_ok=True)

    demo_1 = pd.DataFrame(
        [
            {"Станція": 1, "Спортсмен": "101", "Час_сек": 35, "Завдання_1": "A", "Завдання_2": "C", "Завдання_3": "Z", "Corrected": "no"},
            {"Станція": 1, "Спортсмен": "102", "Час_сек": 42, "Завдання_1": "B", "Завдання_2": "C", "Завдання_3": "A", "Corrected": "yes"},
        ]
    )
    demo_2 = pd.DataFrame(
        [
            {"Станція": 2, "Спортсмен": "101", "Час_сек": 28, "Завдання_1": "D", "Завдання_2": "F", "Завдання_3": "Z", "Corrected": "no"},
            {"Станція": 2, "Спортсмен": "102", "Час_сек": 31, "Завдання_1": "D", "Завдання_2": "E", "Завдання_3": "Z", "Corrected": "no"},
        ]
    )

    demo_1.to_csv(input_dir / "station_1_demo.csv", index=False, encoding="utf-8-sig")
    demo_2.to_csv(input_dir / "station_2_demo.csv", index=False, encoding="utf-8-sig")

    print(f"Demo files created in: {input_dir}")


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Merge TempO station CSV exports into one table.")
    parser.add_argument("--input", type=Path, default=DEFAULT_INPUT_DIR, help="Folder with station CSV exports.")
    parser.add_argument("--output", type=Path, default=DEFAULT_OUTPUT_DIR, help="Folder for merged output files.")
    parser.add_argument("--demo", action="store_true", help="Create demo station CSV files before merging.")
    return parser.parse_args()


def main() -> None:
    args = parse_args()

    if args.demo:
        create_demo_files(args.input)

    merged = merge_station_exports(args.input, args.output)
    print("\nPreview:")
    print(merged.head(10).to_string(index=False))


if __name__ == "__main__":
    main()
