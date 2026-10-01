#!/usr/bin/env python3
"""
Fill official TempO protocol template with station CSV exports.

Inputs:
  data/raw/protocol_template.xlsx
  data/raw/station_exports/*.csv
  optional: data/raw/solutions.csv
  optional: data/raw/athletes.xlsx OR data/raw/athletes.csv

Output:
  data/processed/filled_protocol.xlsx

Run from project root:
  python src/fill_protocol_template.py

What it does:
  - opens official Excel template
  - optionally fills athlete list into the left part of the template
  - reads station CSV exports from the TempO app
  - fills answers into columns like 1-1, 1-2, ...
  - fills station time into the nearest "Час" column after each station block
  - optionally fills correct answers from solutions.csv into the row under 1-1, 1-2, ...
  - keeps existing formulas/styles in the template
"""

from __future__ import annotations

import argparse
import re
from pathlib import Path
from typing import Dict, List, Optional, Tuple

import pandas as pd
from openpyxl import load_workbook


RAW_DIR = Path("data/raw")
STATION_EXPORTS_DIR = RAW_DIR / "station_exports"
TEMPLATE_PATH = RAW_DIR / "protocol_template.xlsx"
SOLUTIONS_PATH = RAW_DIR / "solutions.csv"
ATHLETES_XLSX = RAW_DIR / "athletes.xlsx"
ATHLETES_CSV = RAW_DIR / "athletes.csv"

PROCESSED_DIR = Path("data/processed")
OUT_PATH = PROCESSED_DIR / "filled_protocol.xlsx"


# ----------------------------
# Basic helpers
# ----------------------------

def clean_text(value) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip()


def normalize_key(value) -> str:
    """
    Normalize athlete key for matching.
    Examples:
      1 -> "1"
      1.0 -> "1"
      " 001 " -> "001"
    """
    text = clean_text(value)
    if text == "":
        return ""
    try:
        number = float(text.replace(",", "."))
        if number.is_integer():
            return str(int(number))
    except ValueError:
        pass
    return text


def clean_answer(value) -> str:
    if value is None:
        return ""
    if pd.isna(value):
        return ""
    return str(value).strip().upper()


def read_csv_safely(path: Path) -> pd.DataFrame:
    try:
        return pd.read_csv(path, sep=None, engine="python", encoding="utf-8-sig")
    except UnicodeDecodeError:
        return pd.read_csv(path, sep=None, engine="python", encoding="cp1251")


def find_column(columns: List[str], candidates: List[str]) -> Optional[str]:
    norm = {str(c).lower().strip(): c for c in columns}
    for candidate in candidates:
        key = candidate.lower().strip()
        if key in norm:
            return norm[key]
    return None


# ----------------------------
# Template detection
# ----------------------------

def find_header_row(ws) -> int:
    """
    Find row with athlete-table headers.
    Looks for "№" and something like "Прізвище Ім’я".
    """
    for row in range(1, min(ws.max_row, 30) + 1):
        values = [clean_text(ws.cell(row=row, column=col).value).lower() for col in range(1, ws.max_column + 1)]
        has_number = any(v in ["№", "#", "no", "номер"] for v in values)
        has_name = any(("прізвище" in v or "ім" in v or "name" in v or "спортсмен" in v) for v in values)
        if has_number and has_name:
            return row
    raise ValueError("Не знайшла рядок заголовків спортсменів. Очікую колонки типу '№' і 'Прізвище Ім’я'.")


def find_col_by_header(ws, header_row: int, candidates: List[str]) -> Optional[int]:
    for col in range(1, ws.max_column + 1):
        value = clean_text(ws.cell(row=header_row, column=col).value).lower()
        for candidate in candidates:
            if value == candidate.lower():
                return col
    return None


def find_task_header_row(ws) -> int:
    """
    Find row with max count of task headers like 1-1, 1-2...
    """
    best_row = 1
    best_count = 0
    for row in range(1, min(ws.max_row, 20) + 1):
        count = 0
        for col in range(1, ws.max_column + 1):
            value = clean_text(ws.cell(row=row, column=col).value)
            if re.match(r"^\d+-\d+$", value):
                count += 1
        if count > best_count:
            best_count = count
            best_row = row

    if best_count == 0:
        raise ValueError("Не знайшла заголовки завдань типу 1-1, 1-2, 2-1...")
    return best_row


def map_task_columns(ws, task_header_row: int) -> Dict[Tuple[int, int], int]:
    mapping = {}
    for col in range(1, ws.max_column + 1):
        value = clean_text(ws.cell(row=task_header_row, column=col).value)
        if re.match(r"^\d+-\d+$", value):
            station, task = map(int, value.split("-"))
            mapping[(station, task)] = col
    return mapping


def map_time_columns(ws, task_header_row: int, task_cols: Dict[Tuple[int, int], int]) -> Dict[int, int]:
    """
    Find 'Час' column for each station.
    In the template it is usually after 1-1..1-5, 2-1..2-5 etc.
    We search near the task header row and assign each 'Час' to the last station block on the left.
    """
    station_by_last_task_col = {}
    for (station, task), col in task_cols.items():
        station_by_last_task_col[col] = station

    time_cols = {}
    search_rows = list(range(max(1, task_header_row), min(ws.max_row, task_header_row + 3) + 1))

    for row in search_rows:
        for col in range(1, ws.max_column + 1):
            value = clean_text(ws.cell(row=row, column=col).value).lower()
            if value == "час":
                # assign to nearest station task block on the left
                left_task_cols = [task_col for task_col in station_by_last_task_col.keys() if task_col < col]
                if not left_task_cols:
                    continue
                nearest_task_col = max(left_task_cols)
                station = station_by_last_task_col[nearest_task_col]
                if station not in time_cols:
                    time_cols[station] = col

    return time_cols


def find_solution_row(task_header_row: int) -> int:
    """
    Usually correct answers are placed directly under 1-1, 1-2...
    """
    return task_header_row + 1


# ----------------------------
# Input data
# ----------------------------

def load_station_exports() -> pd.DataFrame:
    files = sorted(STATION_EXPORTS_DIR.glob("*.csv"))
    if not files:
        raise FileNotFoundError(f"У папці {STATION_EXPORTS_DIR} немає CSV-файлів зі станцій.")

    frames = []

    for path in files:
        df = read_csv_safely(path)
        df.columns = [str(c).strip() for c in df.columns]

        station_col = find_column(df.columns.tolist(), ["Станція", "Station", "stationNumber", "station"])
        athlete_col = find_column(df.columns.tolist(), ["Спортсмен", "Athlete", "athleteName", "participant", "name"])
        time_col = find_column(df.columns.tolist(), ["Час_сек", "Time_sec", "stationTime", "time_sec", "Time"])

        if station_col is None:
            raise ValueError(f"{path.name}: не знайшла колонку станції.")
        if athlete_col is None:
            raise ValueError(f"{path.name}: не знайшла колонку спортсмена.")
        if time_col is None:
            raise ValueError(f"{path.name}: не знайшла колонку часу.")

        task_cols = []
        for col in df.columns:
            match = re.match(r"^(?:Завдання|Task|task)[_\s-]*(\d+)$", str(col), flags=re.IGNORECASE)
            if match:
                task_cols.append((int(match.group(1)), col))
        task_cols = sorted(task_cols)

        if not task_cols:
            raise ValueError(f"{path.name}: не знайшла колонки завдань типу Завдання_1, Завдання_2...")

        out = pd.DataFrame()
        out["source_file"] = path.name
        out["station"] = pd.to_numeric(df[station_col], errors="coerce").astype("Int64")
        out["athlete_key"] = df[athlete_col].apply(normalize_key)
        out["time_sec"] = pd.to_numeric(df[time_col], errors="coerce").fillna(0).astype(int)

        for task_num, col in task_cols:
            out[f"task_{task_num}"] = df[col].apply(clean_answer)

        out = out[out["athlete_key"] != ""]
        frames.append(out)

    return pd.concat(frames, ignore_index=True)


def load_solutions() -> Dict[Tuple[int, int], str]:
    if not SOLUTIONS_PATH.exists():
        return {}

    sol = read_csv_safely(SOLUTIONS_PATH)
    sol.columns = [str(c).strip() for c in sol.columns]

    station_col = find_column(sol.columns.tolist(), ["Station", "Станція"])
    task_col = find_column(sol.columns.tolist(), ["Task", "Завдання"])
    answer_col = find_column(sol.columns.tolist(), ["CorrectAnswer", "Correct_Answer", "Правильна_відповідь", "Правильна відповідь"])

    if station_col is None or task_col is None or answer_col is None:
        raise ValueError("solutions.csv має мати колонки Station, Task, CorrectAnswer.")

    mapping = {}
    for _, row in sol.iterrows():
        station = int(float(row[station_col]))
        task = int(float(row[task_col]))
        mapping[(station, task)] = clean_answer(row[answer_col])

    return mapping


def load_athletes_if_exists() -> Optional[pd.DataFrame]:
    if ATHLETES_XLSX.exists():
        return pd.read_excel(ATHLETES_XLSX)
    if ATHLETES_CSV.exists():
        return read_csv_safely(ATHLETES_CSV)
    return None


# ----------------------------
# Filling template
# ----------------------------

def build_existing_athlete_row_map(ws, header_row: int, number_col: int, name_col: Optional[int]) -> Dict[str, int]:
    """
    Build mapping from athlete number/name to worksheet row.
    Prefer number column. Also add name as fallback if present.
    """
    mapping: Dict[str, int] = {}

    for row in range(header_row + 1, ws.max_row + 1):
        number_key = normalize_key(ws.cell(row=row, column=number_col).value)
        name_key = normalize_key(ws.cell(row=row, column=name_col).value) if name_col else ""

        if number_key:
            mapping[number_key] = row
        if name_key:
            mapping[name_key] = row

    return mapping


def fill_athletes(ws, athletes: pd.DataFrame, header_row: int) -> None:
    """
    Fill athlete list into the left part of template.
    Matches columns by header names.
    If athletes file does not have №, it assigns 1..N.
    """
    if athletes is None:
        return

    athletes = athletes.copy()
    athletes.columns = [str(c).strip() for c in athletes.columns]

    # Template headers -> col index
    template_cols = {}
    for col in range(1, ws.max_column + 1):
        header = clean_text(ws.cell(row=header_row, column=col).value)
        if header:
            template_cols[header.lower()] = col

    # Source headers normalized
    source_cols = {str(c).lower().strip(): c for c in athletes.columns}

    def source_for_template(template_header: str) -> Optional[str]:
        key = template_header.lower().strip()
        if key in source_cols:
            return source_cols[key]

        aliases = {
            "№": ["№", "no", "номер", "start", "стартовий"],
            "прізвище ім’я": ["прізвище ім’я", "прізвище ім'я", "спортсмен", "athlete", "name", "піб"],
            "дата народж.": ["дата народж.", "дата народження", "birth", "dob"],
            "група": ["група", "group"],
            "квал.": ["квал.", "кваліфікація", "qual"],
            "регіон": ["регіон", "region"],
            "клас": ["клас", "class"],
            "клуб": ["клуб", "club"],
            "тренер": ["тренер", "coach"],
            "дюсш": ["дюсш", "school"],
        }

        for alias in aliases.get(key, []):
            alias_key = alias.lower().strip()
            if alias_key in source_cols:
                return source_cols[alias_key]

        return None

    start_row = header_row + 1

    # Fill only as many rows as athletes provided
    for i, (_, athlete) in enumerate(athletes.iterrows(), start=0):
        row = start_row + i

        for template_header_lower, col in template_cols.items():
            template_header = clean_text(ws.cell(row=header_row, column=col).value)
            source_col = source_for_template(template_header)

            if source_col is not None:
                ws.cell(row=row, column=col).value = athlete[source_col]
            elif template_header_lower == "№":
                ws.cell(row=row, column=col).value = i + 1


def fill_solutions(ws, solutions: Dict[Tuple[int, int], str], solution_row: int, task_cols: Dict[Tuple[int, int], int]) -> None:
    if not solutions:
        return

    for key, col in task_cols.items():
        if key in solutions:
            ws.cell(row=solution_row, column=col).value = solutions[key]


def fill_station_results(ws, records: pd.DataFrame, athlete_rows: Dict[str, int], task_cols: Dict[Tuple[int, int], int], time_cols: Dict[int, int]) -> None:
    missing_athletes = []
    missing_task_cols = []
    missing_time_cols = []

    for _, rec in records.iterrows():
        athlete_key = normalize_key(rec["athlete_key"])
        station = int(rec["station"])
        row = athlete_rows.get(athlete_key)

        if row is None:
            missing_athletes.append(athlete_key)
            continue

        # Fill task answers
        task_columns_in_record = [c for c in records.columns if re.match(r"^task_\d+$", c)]
        for task_col_name in task_columns_in_record:
            task = int(task_col_name.split("_")[1])
            template_col = task_cols.get((station, task))
            if template_col is None:
                missing_task_cols.append(f"{station}-{task}")
                continue
            ws.cell(row=row, column=template_col).value = clean_answer(rec[task_col_name])

        # Fill station time
        time_col = time_cols.get(station)
        if time_col is None:
            missing_time_cols.append(str(station))
        else:
            ws.cell(row=row, column=time_col).value = int(rec["time_sec"])

    if missing_athletes:
        unique = sorted(set(missing_athletes))
        print("WARNING: не знайшла рядки спортсменів для цих ключів:", ", ".join(unique[:20]))
        if len(unique) > 20:
            print(f"  ... і ще {len(unique) - 20}")

    if missing_task_cols:
        unique = sorted(set(missing_task_cols))
        print("WARNING: не знайшла колонки завдань у шаблоні:", ", ".join(unique[:20]))
        if len(unique) > 20:
            print(f"  ... і ще {len(unique) - 20}")

    if missing_time_cols:
        unique = sorted(set(missing_time_cols))
        print("WARNING: не знайшла колонки часу для станцій:", ", ".join(unique))


def force_formula_recalculation(wb) -> None:
    """
    Ask Excel to recalculate formulas when the workbook is opened.
    """
    try:
        wb.calculation.fullCalcOnLoad = True
        wb.calculation.forceFullCalc = True
    except Exception:
        pass


def main() -> None:
    parser = argparse.ArgumentParser(description="Fill official TempO protocol template.")
    parser.add_argument("--template", default=str(TEMPLATE_PATH), help="Path to protocol template xlsx.")
    parser.add_argument("--out", default=str(OUT_PATH), help="Output xlsx path.")
    args = parser.parse_args()

    template_path = Path(args.template)
    out_path = Path(args.out)

    if not template_path.exists():
        raise FileNotFoundError(f"Не знайшла шаблон: {template_path}")

    records = load_station_exports()
    solutions = load_solutions()

    wb = load_workbook(template_path)
    ws = wb.active

    athlete_header_row = find_header_row(ws)
    task_header_row = find_task_header_row(ws)
    solution_row = find_solution_row(task_header_row)

    number_col = find_col_by_header(ws, athlete_header_row, ["№", "#", "No", "номер"])
    name_col = find_col_by_header(ws, athlete_header_row, ["Прізвище Ім’я", "Прізвище Ім'я", "Спортсмен", "Athlete", "Name"])

    if number_col is None:
        raise ValueError("Не знайшла колонку '№' у шаблоні.")

    # Optional: fill athlete list from data/raw/athletes.xlsx or athletes.csv
    athletes = load_athletes_if_exists()
    if athletes is not None:
        fill_athletes(ws, athletes, athlete_header_row)

    # Rebuild mapping after optional fill
    athlete_rows = build_existing_athlete_row_map(ws, athlete_header_row, number_col, name_col)

    task_cols = map_task_columns(ws, task_header_row)
    time_cols = map_time_columns(ws, task_header_row, task_cols)

    fill_solutions(ws, solutions, solution_row, task_cols)
    fill_station_results(ws, records, athlete_rows, task_cols, time_cols)

    force_formula_recalculation(wb)

    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)
    wb.save(out_path)

    print("Done.")
    print(f"Template: {template_path}")
    print(f"Station records loaded: {len(records)}")
    print(f"Solutions loaded: {len(solutions)}")
    print(f"Output: {out_path}")


if __name__ == "__main__":
    main()
