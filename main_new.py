import pandas as pd
from pathlib import Path
from dotenv import load_dotenv, env
from core.themes import extract_gift_themes_from_df, save_list_to_json

load_dotenv()


class Config:
    input_file_name: str = env("FILE_NAME")
    input_file_path: Path = Path(__file__).parent / input_file_name

    load_gift_themes: bool = env("LOAD_GIFT_THEMES", "true").lower() == "true"


def load_df(path: str) -> pd.DataFrame:
    """Loads a file into a DataFrame."""
    if not Path(path).exists():
        raise FileNotFoundError(f"File not found: {path}")

    if str(path).lower().endswith(".csv"):
        return pd.read_csv(path)

    if str(path).lower().endswith(".xlsx"):
        return pd.read_excel(path)

    raise ValueError(f"Unsupported file format: {path}")


def main() -> None:
    config = Config()
    df = load_df(config.input_file_path)

    if config.load_gift_themes:
        themes = extract_gift_themes_from_df(df)
        save_list_to_json(themes, out_path="gift_themes.json")


if __name__ == "__main__":
    main()
