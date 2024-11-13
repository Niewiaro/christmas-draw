import win32com.client as win32


class Person:
    def __init__(self, /, name: str, surname: str, email: str) -> None:
        self.name = name
        self.surname = surname
        self.email = email

    def __repr__(self) -> str:
        return f"{self.name}\t{self.surname}\t{self.email}"


def send_email(to: str, subject: str, body: str = None, html_body: str = None):
    outlook = win32.Dispatch("outlook.application")
    mail = outlook.CreateItem(0)
    mail.To = to
    mail.Subject = subject

    if body is not None:
        mail.Body = "Message body"

    if html_body is not None:
        mail.HTMLBody = "<h2>HTML Message body</h2>"

    mail.Send()


def main() -> None:
    print("https://github.com/Niewiaro/christmas-draw")

    persons = [
        Person(name="Jan", surname="Mak", email="jan.mak@example.com"),
        Person(name="Anna", surname="Nowak", email="anna.nowak@example.com"),
        Person(name="Piotr", surname="Ćwir", email="piotr.cwir@example.com"),
    ]

    for person in persons:
        print(person)


if __name__ == "__main__":
    main()
