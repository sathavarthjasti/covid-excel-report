# COVID-19 Excel Report Generator

This project reads a small, public COVID-19 dataset and generates a styled Excel report
showing the top 20 countries by confirmed cases on the most recent date.

## Features
- Pulls real data from Johns Hopkins COVID-19 dataset
- Filters the latest available snapshot
- Generates Excel file with:
  - Styled table
  - Bar chart of top 20 countries

## Tech Stack
- Python 3
- Pandas
- OpenPyXL
- Requests

## How to Run
1. Install dependencies:
   ```
   pip install pandas openpyxl requests
   ```

2. Run the script:
   ```
   python covid_cases_excel_report.py
   ```

3. Output:
   A file named `covid_country_summary.xlsx` will be created in your directory.

## Dataset Source
[Johns Hopkins CSSE COVID-19 Dataset](https://github.com/datasets/covid-19)

## License
MIT
