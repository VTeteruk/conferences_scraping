import re


def delete_numbers_from_text(text: str) -> str:
    return re.sub("[0-9]+", "", text)


def clean_text_for_names(text: str):
    text = re.sub("-", " ", text)
    return delete_numbers_from_text(text)
