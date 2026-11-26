# Импорт Excel, CSV и TXT в SQLite

Настольное приложение переносит табличные файлы в SQLite для последующего просмотра и поиска.

## Возможности

- Импортирует `.xls`, `.xlsx`, `.csv` и `.txt`.
- Позволяет настроить имена столбцов, заголовки, кодировку и разделитель.
- Читает CSV и TXT порциями заданного размера и показывает ход обработки.

## Устройство

Интерфейс Tkinter из `gui_app_optimized.py` запускает `DataProcessorOptimized` в фоновом потоке. Процессор читает таблицы через pandas, openpyxl или xlrd и записывает данные в SQLite. В `helper_windows.py` находятся окна настроек, сопоставления столбцов и прогресса. В `gui_app.py` и `data_processor.py` сохранена ранняя реализация.

## Установка и запуск

Установите Python с Tkinter. Локальные проверки выполнялись с Python 3.14 и версиями пакетов из `requirements.txt`.

```powershell
python -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe gui_app_optimized.py
```

Выберите входной файл и укажите новый файл назначения SQLite. При необходимости настройте заголовки, разделитель и столбцы, затем запустите импорт.

Процессор можно вызвать и из Python:

```python
from data_processor_optimized import DataProcessorOptimized

processor = DataProcessorOptimized({"chunk_size": 10_000})
processor.convert_file_to_db(
    input_file_path="sample.csv",
    db_file_path="sample.db",
    use_first_row_as_header=True,
)
```

Для проверки заголовков CSV и TXT выполните `python -m unittest -q test_csv_txt_headers`.

## Ограничения

Первый лист Excel загружается в память целиком. Чтобы сохранить существующую таблицу, указывайте новый файл назначения. В ранней документации упомянута лицензия MIT. Отдельный файл лицензии требует подтверждения условий распространения.
