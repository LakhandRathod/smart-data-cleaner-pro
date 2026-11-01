# Smart Data Cleaner Pro

A desktop-based Excel and CSV data cleaning application built with Python, CustomTkinter and Pandas.

## Features

- Import Excel (`.xlsx`, `.xls`) and CSV (`.csv`) files
- Remove duplicate rows and completely blank rows
- Trim leading and trailing spaces from text columns
- Convert text columns to Title Case
- Optionally remove rows containing missing values
- Preview source data and compare before/after results
- Export cleaned data to Excel and CSV
- Save a local cleaning history
- Dark and light appearance modes
- Optional drag-and-drop file upload

## Requirements

- Python 3.10 or newer recommended
- Windows, macOS or Linux desktop environment

## Installation

Clone the repository:

```bash
git clone https://github.com/YOUR-USERNAME/smart-data-cleaner-pro.git
cd smart-data-cleaner-pro
```

Create and activate a virtual environment (recommended):

### Windows PowerShell

```powershell
py -m venv .venv
.\.venv\Scripts\Activate.ps1
```

Install dependencies:

```powershell
python -m pip install -r requirements.txt
```

Run the application:

```powershell
python app/main.py
```

## How to use

1. Open the application.
2. Browse for an Excel or CSV file.
3. Choose an output folder.
4. Select the cleaning operations.
5. Click **Clean Data**.
6. Review the before/after comparison and cleaning history.
7. Find the exported `_cleaned.xlsx` and `_cleaned.csv` files in the selected output folder.

## Project structure

```text
smart-data-cleaner-pro/
├── app/
│   └── main.py
├── docs/
├── screenshots/
├── .gitignore
├── README.md
└── requirements.txt
```

## Notes

- This is a desktop application, not a web application. GitHub stores and shares the source code; GitHub Pages cannot run this Python desktop GUI.
- Drag-and-drop support is optional and requires `tkinterdnd2`.
- Your input spreadsheets are processed locally. Do not commit private or sensitive datasets to GitHub.
- Cleaning history is stored locally in `cleaning_history.json` when the application runs.

## Future improvements

- Unit tests and automated checks
- More data validation and cleaning rules
- Build a Windows executable
- Add configurable cleaning rules and downloadable reports

## License

Choose and add a license before accepting contributions or allowing reuse. If you do not add one, normal copyright restrictions apply.
