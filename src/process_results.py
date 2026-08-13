import pandas as pd
from pathlib import Path


RAW_PATH = Path("")
PROCESSED_DIR = Path("")
SUMMARY_PATH = PROCESSED_DIR / ".csv"


def load_results(path: Path) -> pd.DataFrame:

    return pd.read_csv(path)


def prepare_results(df: pd.DataFrame) -> pd.DataFrame:

    return df


def summarize_results(df: pd.DataFrame) -> pd.DataFrame:

    return sammary


def main():
    PROCESSED_DIR.mkdir(parents=True, exist_ok=True)

    print(sammary)


if __name__ == "__main__":
    main()
