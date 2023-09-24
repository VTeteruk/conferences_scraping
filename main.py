import logging
import re
import PyPDF2
from openpyxl.reader.excel import load_workbook


logging.basicConfig(
    level=logging.INFO
)


def extract_text_from_pdf(file_path: str, start_page: int = 0, end_page: int = None) -> str:
    reader = PyPDF2.PdfReader(file_path)

    if end_page is None:
        end_page = len(reader.pages) - 4  # Last 4 pages are useless

    # Extract text from specified pages and join it into a single string
    return " ".join(
        [
            reader.pages[index].extract_text()
            for index
            in range(start_page, end_page)
        ]
    )


def clean_text(text: str) -> str:
    return (
        text
        .replace("\t", " ")
        .replace("\n-", "")
        .replace("5th World Psoriasis & Psoriatic Arthritis Conference 2018", "")
        .replace("Acta Derm Venereol 2018", "")
        .replace("Poster abstracts", "")
        .replace("www.medicaljournals.se/acta", "")
        .replace(".P", ".\nP")
        .replace("Background", "Introduction")
        .replace("\xad", " ")
        .replace("Introduction", "\nIntroduction")
    )


def divide_text_by_session_name(text: str) -> list[str]:
    session_names = re.findall("P[0-9]{3}", text)
    sections = []
    start_index = 0

    for session_name in session_names:
        end_index = text.find(session_name, start_index)

        if end_index != -1:
            section_text = text[start_index:end_index].strip()
            sections.append(section_text)
            start_index = end_index

    last_section = text[start_index:].strip()
    sections.append(last_section)

    # Remove sections with less than 3 characters (page numbers)
    return [section for section in sections if len(section) > 3]


def delete_numbers(text: str) -> str:
    return re.sub("[0-9]+", "", text)


def clean_name_text(text: str):
    text = re.sub("-", " ", text)
    return delete_numbers(text)


def find_topic(data: list) -> tuple[str, int, str | None]:
    topic = ""
    for index, text in enumerate(data):
        if text.isupper():
            topic += text
        else:
            # Check if it wasn't separated correctly
            only_upper_text = re.sub("[a-z]", "+", text).split("+")
            length_of_upper_text = len(only_upper_text[0])

            # Check if it is only upper case text
            if len(only_upper_text[0].replace("-", " ").split()[0]) > 1:
                topic += only_upper_text[0][:-1]
                return topic, index + 1, text[length_of_upper_text - 1:]
            else:
                return topic, index + 1, None


def find_name(data: list) -> tuple[str, int]:
    name = []
    for index in range(0, len(data) - 1):
        text = clean_name_text(data[index])
        future_text = clean_name_text(data[index + 1])

        name.append(data[index])
        if not (text.endswith(" ") or future_text.startswith(",")):
            return "".join(name), index + 1


def split_location_and_presentation(data: list) -> list[str, str]:
    if "Introduction" in "".join(data):
        return "".join(data).split("Introduction")
    else:
        return [data[0], "".join(data[1:])]


def get_affiliation_from_full_location(full_location: str) -> tuple[str, str]:
    split_location = full_location.split(", ")
    if len(split_location) == 1:
        affiliation = split_location
        location = ""
    else:
        *affiliation, location = split_location

    return " ".join(affiliation), location


def parse_section(section: str) -> dict:
    split_data = section.split("\n")
    session = split_data[0].strip()
    topic, start_index, possible_name = find_topic(split_data[1:])

    if possible_name:
        split_data[start_index] = possible_name

    name_and_index = find_name(split_data[start_index:])
    name = name_and_index[0]

    start_index += name_and_index[1]

    full_location, presentation = split_location_and_presentation(split_data[start_index:])

    affiliation, location = get_affiliation_from_full_location(full_location)

    presentation = (
                   "Introduction" if presentation.startswith(":") else "Introduction "
               ) + presentation

    logging.info(f"Session {session} was parsed")

    return {
        "name": delete_numbers(name),
        "affiliation": delete_numbers(affiliation),
        "location": delete_numbers(location),
        "session": session,
        "topic": topic,
        "presentation": presentation,
    }


def add_data_to_excel(file: str, articles: list, start_row: int):
    wb = load_workbook(file)

    # Select the active sheet
    ws = wb.active

    for i, article in enumerate(articles, start=start_row):
        ws.cell(row=i, column=1, value=article.get("name"))
        ws.cell(row=i, column=2, value=article.get("affiliation"))
        ws.cell(row=i, column=3, value=article.get("location"))
        ws.cell(row=i, column=4, value=article.get("session"))
        ws.cell(row=i, column=5, value=article.get("topic"))
        ws.cell(row=i, column=6, value=article.get("presentation"))

    wb.save(file_name)

    logging.info("Data was added to excel file successfully")


if __name__ == "__main__":
    # Extract text from the PDF
    pdf_text = extract_text_from_pdf("data_to_parse/Book.pdf", start_page=43)

    # Clean the extracted text
    cleaned_text = clean_text(pdf_text)

    # Divide the text into sections based on session names
    all_parts = divide_text_by_session_name(cleaned_text)

    # Parse each section to extract information
    parsed_articles = [parse_section(part) for part in all_parts]

    # Add parsed data to an Excel file
    file_name = (
        "results/"
        "Data Entry - 5th World Psoriasis & Psoriatic "
        "Arthritis Conference 2018 - Case format (2).xlsx"
    )

    add_data_to_excel(
        file=file_name,
        articles=parsed_articles,
        start_row=7
    )
