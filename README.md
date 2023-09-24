# Conferences Parser

This Python script is designed to parse data from a PDF document containing information from the 5th World Psoriasis & Psoriatic Arthritis Conference 2018 and store it in an Excel file. It cleans, extracts, and organizes the data for further analysis.
___
## Requirements

- Python 3.x
- Install libraries using the following command:
```bash
pip install -r requrements.txt
```
___
# Usage
* Place the PDF file you want to parse in the data_to_parse directory.
* Make sure you have the necessary libraries installed.
* Run the script:
```bash
python main.py
```

The parsed data will be added to an Excel file named:
`Data Entry - 5th World Psoriasis & Psoriatic Arthritis Conference 2018 - Case format (2).xlsx`
in the `results` directory.
___
# Configuration
You can customize the script's behavior by modifying the following variables in the script:
1. `start_page`: The starting page for PDF text extraction.
2. `end_page`: The ending page for PDF text extraction.
3. `file_name`: The name of the Excel file where the parsed data will be saved.
