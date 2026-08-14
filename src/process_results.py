import pandas as pd
from pathlib import Path


RAW_PATH = Path("")
PROCESSED_DIR = Path("")
SUMMARY_PATH = PROCESSED_DIR / ".csv"


def load_results(path: Path) -> pd.DataFrame:
    """Load raw TempO results from CSV"""

    return pd.read_csv(path)


def prepare_results(df: pd.DataFrame) -> pd.DataFrame:
    """Add correctness and penalty columns"""

    # неправильна відповідь = час + штраф
    # Поки ставимо умовний штраф 30 секунд

    return df


def summarize_results(df: pd.DataFrame) -> pd.DataFrame:
    """Create participant-level summary"""

    return summary


def main():
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    print(summary)


if __name__ == "__main__":
    main()
