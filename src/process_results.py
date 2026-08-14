import pandas as pd
from pathlib import Path


RAW_PATH = Path("")
PROCESSED_DIR = Path("")
SUMMARY_PATH = PROCESSED_DIR / ".csv"


def load_results(path: Path) -> pd.DataFrame:
    """Load raw TempO results from CSV"""
    if not path.exists():
        raise FileNotFoundError(f"File not found: {path}")

    return pd.read_csv(path)


def prepare_results(df: pd.DataFrame) -> pd.DataFrame:
    """Add correctness and penalty columns"""
    df = df.copy()

    df["is_correct"] = df["answer"] == df["correct_answer"]

    # неправильна відповідь = час + штраф
    # ставимо штраф = 30 секунд
    df["penalty_sec"] = df["is_correct"].apply(lambda x: 0 if x else 30)
    df["total_time_sec"] = df["time_sec"] + df["penalty_sec"]

    return df


def summarize_results(df: pd.DataFrame) -> pd.DataFrame:
    """Create participant-level summary"""

    return summary


def main():
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    print(summary)


if __name__ == "__main__":
    main()
