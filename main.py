from pathlib import Path
from typing import List


class Config:
    """Class to store configuration details."""

    def __init__(self) -> None:
        self.input_name = "input.xlsx"
        self.input_path = Path.cwd() / self.input_name

        self.email_subject = "🎄 List do Świętego Mikołaja! 🎅"
        self.email_html_body = """
<!DOCTYPE html>
<html lang="pl">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Wesołych Świąt!</title>
    <style>
        body {{
            font-family: Arial, sans-serif;
            background: #002244;
            color: #333;
            margin: 0;
            padding: 0;
        }}
        .container {{
            max-width: 600px;
            margin: 100px auto 100px auto;
            background-color: #fdf6e3;
            border-radius: 15px;
            overflow: hidden;
            box-shadow: 0 4px 8px rgba(0, 0, 0, 0.2);
            border: 5px double #d32f2f;
            padding: 20px;
        }}
        .header {{
            background-color: #d32f2f;
            color: #fff;
            text-align: center;
            padding: 20px;
            border-bottom: 3px dashed #fff;
        }}
        .header h1 {{
            margin: 0;
            font-size: 28px;
        }}
        .content {{
            padding: 20px;
            line-height: 1.8;
        }}
        .content h2 {{
            color: #d32f2f;
            text-align: center;
            margin-bottom: 20px;
        }}
        .st00pka {{
            text-align: center;
            background-color: #d32f2f;
            color: #fff;
            padding: 15px;
            font-size: 14px;
            border-top: 3px dashed #fff;
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🎄 Cześć, Mikołaju {name}! 🎅</h1>
        </div>
        <div class="content">
            <h2>📜 List od {draw} 🎁</h2>
            <p>Drogi Mikołaju {name},</p>
            <p>W tym roku zostałeś wybrany, aby spełnić moje świąteczne marzenie! 🎅</p>
            <p>Oto kilka informacji, które mogą Ci pomóc w przygotowaniu prezentu:</p>
            <ul>
                <li>🎁 <strong>Co chcę dostać:</strong> {gift}</li>
                <li>🎨 <strong>Hobby:</strong> {hobby}</li>
                <li>📚 <strong>Co robię w wolnym czasie:</strong> {free_time}</li>
                <li>🎨 <strong>Ulubiony kolor:</strong> {favorite_color}</li>
                <li>❌ <strong>Czego nie chcę dostać:</strong> {not_want}</li>
            </ul>
            <p>Nie mogę się doczekać niespodzianki, którą dla mnie przygotujesz. Mam nadzieję, że znajdziesz coś, co wywoła uśmiech na mojej twarzy! 🎄✨</p>
            <p>Wesołych Świąt i dużo świątecznej magii! 🌟</p>
            <p>Z pozdrowieniami,<br>Twoje Ukochanie Świąteczne Dziecko, {draw}</p>
        </div>
        <div class="st00pka">
            <p>Dziękuję, że jesteś częścią tych magicznych Świąt! ❄️</p>
            <p>🎄 Wesołych Świąt! 🎁</p>
        </div>
    </div>
</body>
</html>
"""

        self.map = [
            "Imię i nazwisko",
            "Adres e-mail",
            "Bardzo chcę dostać...",
            "Moje hobby to...",
            "W wolnym czasie uwielbiam...",
            "Mój ulubiony kolor to...",
            "Bardzo nie chcę dostać...",
        ]

    def is_valid(self) -> bool:
        """Checks if the input file exists."""
        return self.input_path.exists()


class Person:
    """Represents a person with name, email and a draw assignment."""

    def __init__(
        self,
        *,
        name: str,
        email: str,
        gift: str = None,
        hobby: str = None,
        free_time: str = None,
        favorite_color: str = None,
        not_want: str = None,
    ) -> None:
        self.name = name
        self.email = email
        self.draw = None
        self.gift = gift
        self.hobby = hobby
        self.free_time = free_time
        self.favorite_color = favorite_color
        self.not_want = not_want

    def __repr__(self) -> str:
        return f"name: {self.name}\nemail: {self.email}\ndraw: {self.draw}"

    def __str__(self) -> str:
        return self.name


def get_data_from_excel(path: Path, map: List[str]) -> List[Person]:
    """Reads person data from an Excel file."""
    import pandas as pd

    try:
        file = pd.ExcelFile(path)
        with file as xlsx:
            df = pd.read_excel(xlsx, "Sheet1")
            result = [
                Person(
                    name=name,
                    email=email,
                    gift=gift,
                    hobby=hobby,
                    free_time=free_time,
                    favorite_color=favorite_color,
                    not_want=not_want,
                )
                for name, email, gift, hobby, free_time, favorite_color, not_want in df[
                    map
                ].to_numpy()
            ]
        return result
    except Exception as e:
        print(f"Error reading file: {e}")
        return []


def perform_draw(persons: List[Person]) -> None:
    """Assigns a random draw to each person."""
    import random, copy

    bad_draw = True

    while bad_draw:
        bad_draw = False
        draws = copy.deepcopy(persons)
        random.shuffle(draws)

        for i, person in enumerate(persons):
            if person.name == draws[i].name:
                bad_draw = True
                break
            person.draw = draws[i]


def send_email(
    *,
    person: Person = None,
    to: str = None,
    subject: str = "",
    body: str = None,
    html_body: str = None,
    attribute_not_found: str = "Brak informacji",
) -> None:
    """Sends an email via Outlook."""
    import win32com.client as win32

    try:
        if person and person.email:
            to = person.email
        elif to is None:
            raise ValueError("to is None")

        # Format the email body dynamically by replacing placeholders
        html_body_formatted = html_body.format(
            name=person.name,
            draw=person.draw.name,
            gift=person.draw.gift or attribute_not_found,
            hobby=person.draw.hobby or attribute_not_found,
            free_time=person.draw.free_time or attribute_not_found,
            favorite_color=person.draw.favorite_color or attribute_not_found,
            not_want=person.draw.not_want or attribute_not_found,
        )

        # Setting up Outlook email
        outlook = win32.Dispatch("outlook.application")
        mail = outlook.CreateItem(0)
        mail.To = to
        mail.Subject = subject

        if body:
            mail.Body = body

        if html_body_formatted:
            mail.HTMLBody = html_body_formatted

        # print(f"mail.To:\t{mail.To}")
        # print(f"mail.Subject:\t{mail.Subject}")
        # print(f"mail.Body:\t{mail.Body}")
        # print(f"mail.HTMLBody:\t{mail.HTMLBody}")
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

    persons = get_data_from_excel(config.input_path, config.map)
    if not persons:
        print("No data found.")
        return

    perform_draw(persons)
    for person in persons:
        print(f"Sending email to {person}...")
        # print(f"Got {person.draw.name}...")
        send_email(
            person=person,
            subject=config.email_subject,
            html_body=config.email_html_body,
        )
        print("Done\n")


if __name__ == "__main__":
    main()
