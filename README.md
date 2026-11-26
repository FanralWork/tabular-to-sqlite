# Tabular Files to SQLite

A Tkinter application that imports Excel, CSV, and delimited text files into SQLite for local inspection and querying.

## Features

- Imports `.xls`, `.xlsx`, `.csv`, and `.txt` into a SQLite table.
- Supports column renaming, optional first-row headers, and text encoding and delimiter settings.
- Processes CSV and TXT in configurable chunks and displays conversion progress.

## Stack and architecture

The Tkinter interface in `gui_app_optimized.py` calls `DataProcessorOptimized` in a background thread. The processor uses pandas with openpyxl or xlrd for Excel reading and writes to SQLite. `helper_windows.py` implements the settings, column mapping, and progress dialogs. `gui_app.py` and `data_processor.py` contain the earlier implementation.

## Setup

Use Python with Tkinter installed. This setup was checked with Python 3.14 and the pinned packages in `requirements.txt`.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe gui_app_optimized.py
```

## Usage

Select a format and input file. Set the destination to a new SQLite file. Adjust the header, delimiter, or column mapping if necessary, then start the conversion.

You can also call the processor from Python.

```python
from data_processor_optimized import DataProcessorOptimized

processor = DataProcessorOptimized({"chunk_size": 10_000})
processor.convert_file_to_db(
    input_file_path="sample.csv",
    db_file_path="sample.db",
    use_first_row_as_header=True,
)
```

Run the CSV and TXT header regression tests with `python -m unittest -q test_csv_txt_headers`.

## Status and limitations

Excel input is read from the first sheet into memory. Use a new output path to avoid replacing a destination table. The earlier README identified MIT as the license. A standalone license file and rights confirmation for distribution remain pending.

## Кратко для портфолио

Настольная утилита переноса Excel, CSV и TXT в SQLite. Реализованы настройки импорта, порционная обработка текстовых файлов и отображение прогресса.
