import pandas as pd


def extract_gift_themes_from_df(df: pd.DataFrame, cols: list = None) -> list:
    if cols is None:
        cols = [
            "Zaproponuj pierwszy motyw prezentowy",
            "Zaproponuj drugi motyw prezentowy",
            "Zaproponuj trzeci motyw prezentowy",
        ]

    seen = set()
    result = []
    for col in cols:
        if col not in df.columns:
            continue
        for val in df[col].dropna().astype(str):
            s = val.strip()
            if s == "":
                continue
            if s not in seen:
                seen.add(s)
                result.append(s)
            else:
                print(f"Duplicate theme found and skipped: {s}")

    return result


def save_list_to_json(list: list, out_path: str, sort: bool = True) -> None:
    from pathlib import Path
    import json

    out = Path(out_path)
    out.parent.mkdir(parents=True, exist_ok=True)
    if sort:
        list = sorted(list, key=lambda s: s.lower())
    with out.open("w", encoding="utf-8") as fh:
        json.dump(list, fh, ensure_ascii=False, indent=2)


def assign_gift_themes(persons, themes):
    """
    Assigns each person exactly one unique gift theme.
    `persons` — list of Person objects
    `themes` — list of strings (themes)
    """
    import random

    if len(themes) < len(persons):
        raise ValueError("Not enough themes to assign to each person.")

    selected_themes = random.sample(themes, len(persons))

    for person, theme in zip(persons, selected_themes):
        person.gift_theme = theme
