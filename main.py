class Person:
    def __init__(self, /, name: str, surname: str, email: str) -> None:
        self.name = name
        self.surname = surname
        self.email = email

    def __repr__(self) -> str:
        return f"{self.name}\t{self.surname}\t{self.email}"


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
