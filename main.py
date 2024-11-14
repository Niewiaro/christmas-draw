from pathlib import Path
from typing import List


class Config:
    """Class to store configuration details."""

    def __init__(self) -> None:
        self.input_name = "input.xlsx"
        self.input_path = Path.cwd() / self.input_name

    def is_valid(self) -> bool:
        """Checks if the input file exists."""
        return self.input_path.exists()


class Person:
    """Represents a person with name, email and a draw assignment."""

    def __init__(self, *, name: str, email: str) -> None:
        self.name = name
        self.email = email
        self.draw = None

    def __repr__(self) -> str:
        return f"name: {self.name}\nemail: {self.email}\ndraw: {self.draw}"


def get_data_from_excel(path: Path) -> List[Person]:
    """Reads person data from an Excel file."""
    import pandas as pd

    try:
        file = pd.ExcelFile(path)
        with file as xlsx:
            df = pd.read_excel(xlsx, "Sheet1")
            result = [
                Person(name=name, email=email)
                for name, email in df[["Nazwa", "Adres e-mail"]].to_numpy()
            ]
        return result
    except Exception as e:
        print(f"Error reading file: {e}")
        return []


def perform_draw(persons: List[Person]) -> None:
    """Assigns a random draw to each person."""
    import random

    names = [person.name for person in persons]
    random.shuffle(names)

    for i, person in enumerate(persons):
        person.draw = names[i]


def send_email(
    *, to: str, subject: str, body: str = None, html_body: str = None
) -> None:
    """Sends an email via Outlook."""
    import win32com.client as win32

    try:
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject

        if body:
            mail.Body = body

        if html_body:
            mail.HTMLBody = html_body

        mail.Send()
    except Exception as e:
        print(f"Error sending email: {e}")


def main() -> None:
    """Main function to execute the program logic."""
    print("https://github.com/Niewiaro/christmas-draw")

    config = Config()
    if not config.is_valid():
        print(f"Error: {config.input_name} does not exist.")
        return

    persons = get_data_from_excel(config.input_path)
    if not persons:
        print("No data found.")
        return

    perform_draw(persons)
    for person in persons:
        # Uncomment the line below to send emails
        # send_email(to=person.email, subject="Test", body=f"Hello {person.name}")
        print(f"{person}\n")


if __name__ == "__main__":
    main()
