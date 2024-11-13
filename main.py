import win32com.client as win32
import pandas as pd
from pathlib import Path


class Config:
    def __init__(self) -> None:
        self.input_name = "input.xlsx"
        self.input_path = Path.cwd() / self.input_name


class Person:
    def __init__(self, *, name: str, email: str) -> None:
        self.name = name
        self.email = email

    def __repr__(self) -> str:
        return f"{self.name}\t{self.email}"


def get_data_from_excel(path) -> list:
    file = pd.ExcelFile(path)
    with file as xlsx:
        df = pd.read_excel(xlsx, "Sheet1")
        # print(df.to_string())
        result = [
            Person(name=name, email=email)
            for name, email in df[["Nazwa", "Adres e-mail"]].to_numpy()
        ]
        return result


def send_email(
    *, to: str, subject: str, body: str = None, html_body: str = None
) -> None:
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject

    if body is not None:
        mail.Body = body

    if html_body is not None:
        mail.HTMLBody = html_body

    mail.Send()


def main() -> None:
    print("https://github.com/Niewiaro/christmas-draw")

    config = Config()
    persons = get_data_from_excel(config.input_path)
    for person in persons:
        send_email(to=person.email, subject="Test", body=f"Hello {person.name}")


if __name__ == "__main__":
    main()
