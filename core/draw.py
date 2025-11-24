from dataclasses import dataclass, field
import re

EMAIL_REGEX = re.compile(r"^[^@\s]+@[^@\s]+\.[^@\s]+$")


@dataclass
class Person:
    """Represents a person participating in Secret Santa."""

    name: str
    email: str
    gift: str | None = None
    hobby: str | None = None
    free_time: str | None = None
    favorite_color: str | None = None
    not_want: str | None = None

    # Not shown in repr, but stored normally
    draw: Person | None = field(default=None, repr=False)

    def __post_init__(self):
        """Runs automatically after __init__."""

        # Email validation
        if not EMAIL_REGEX.match(self.email):
            raise ValueError(f"Invalid email: {self.email!r}")

    def __str__(self) -> str:
        return self.name

    def __str__(self) -> str:
        # return f"{self.name} <{self.email}>"
        return f"{self.name}"


def assign_draws(persons: list[Person]) -> None:
    import random

    if len(persons) < 2:
        raise ValueError("At least two persons are required for the draw.")

    # List of persons to be randomly permuted
    targets = persons.copy()

    # Shuffle until no one draws themselves
    # (for n > 3 it usually succeeds in 1–2 attempts)
    while True:
        random.shuffle(targets)

        if all(p is not t for p, t in zip(persons, targets)):
            break

    # Assign the results to the objects
    for p, t in zip(persons, targets):
        p.draw = t
