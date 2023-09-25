import re

START_PAGE = 43
NUM_PAGES_TO_SKIP = 4  # Last 4 pages are useless

REPLACEMENTS = (
            ("\t", " "),
            ("\n-", ""),
            ("5th World Psoriasis & Psoriatic Arthritis Conference 2018", ""),
            ("Acta Derm Venereol 2018", ""),
            ("Poster abstracts", ""),
            ("www.medicaljournals.se/acta", ""),
            (".P", ".\nP"),
            ("Background", "Introduction"),
            ("\xad", " "),
            ("Introduction", "\nIntroduction"),
            ("1,2", "\n"),
        )


def delete_numbers_from_text(text: str) -> str:
    return re.sub("[0-9]+", "", text)


def clean_text_for_names(text: str):
    text = re.sub("-", " ", text)
    return delete_numbers_from_text(text)
