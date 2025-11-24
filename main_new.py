from pathlib import Path
import pandas as pd
import json

from core.themes import extract_gift_themes_from_df, save_list_to_json
from core.draw import Person, assign_draws
from core.email import send_email_via_outlook


class Config:
    def __init__(self) -> None:
        from dotenv import load_dotenv
        import os

        load_dotenv()

        self.input_file_name: str = os.getenv("FILE_NAME")
        self.input_file_path: Path = Path(__file__).parent / self.input_file_name

        self.load_gift_themes: bool = (
            os.getenv("LOAD_GIFT_THEMES", "true").lower() == "true"
        )
        self.column_map: dict = json.loads(os.getenv("COLUMN_MAP"))

        self.send_backup_emails: str | None = (
            os.getenv("SEND_BACKUP_EMAILS", None)
        ).strip()

    def is_valid(self) -> bool:
        """Checks if the input file exists."""
        return self.input_file_path.exists()


def load_df(path: str) -> pd.DataFrame:
    """Loads a file into a DataFrame."""
    if not Path(path).exists():
        raise FileNotFoundError(f"File not found: {path}")

    if str(path).lower().endswith(".csv"):
        return pd.read_csv(path)

    if str(path).lower().endswith(".xlsx"):
        return pd.read_excel(path)

    raise ValueError(f"Unsupported file format: {path}")


def get_data_from_df(df: pd.DataFrame, column_map: dict[str, str]) -> list[Person]:
    persons = []

    for _, row in df.iterrows():
        kwargs = {}

        for col_name, attr_name in column_map.items():
            if col_name in df.columns:
                kwargs[attr_name] = row[col_name]

        persons.append(Person(**kwargs))

    return persons


def main() -> None:
    """Main function to execute the program logic."""
    print("https://github.com/Niewiaro/christmas-draw")

    config = Config()
    if not config.is_valid():
        raise FileNotFoundError(f"Input file not found: {config.input_file_path}")

    df = load_df(config.input_file_path)

    if config.load_gift_themes:
        themes = extract_gift_themes_from_df(df)
        save_list_to_json(themes, out_path="gift_themes.json")

    persons = get_data_from_df(df, config.column_map)

    if not persons:
        raise ValueError("No persons found in the input data.")

    print(f"Loaded {len(persons)} persons from {config.input_file_name}")

    assign_draws(persons)

    if config.send_backup_emails:
        backup_data = [
            {"person": str(person), "draw": str(person.draw)} for person in persons
        ]
        backup_body = json.dumps(backup_data, indent=4, ensure_ascii=False)
        # send_email_via_outlook(
        #     to=config.send_backup_emails,
        #     subject="Backup of Secret Santa Draws",
        #     body=backup_body,
        # )
        print(f"Backup email sent to {config.send_backup_emails}")
        # print(backup_body)


if __name__ == "__main__":
    main()
