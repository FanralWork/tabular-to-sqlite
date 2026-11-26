import pandas as pd
import sqlite3
import os
from tqdm import tqdm
import math


class DataProcessorOptimized:
    def __init__(self, settings=None):
        self.log_messages = []
        self.settings = settings or {
            'remove_empty_cols': True,
            'overwrite_existing': True,
            'chunk_size': 10000  # Размер чанка по умолчанию
        }

    def log(self, message):
        """Добавляет сообщение в лог и выводит его"""
        self.log_messages.append(message)
        print(message)

    def validate_file_path(self, file_path, allowed_extensions=None):
        """Проверяет существование файла и его расширение"""
        if allowed_extensions is None:
            allowed_extensions = ['.xlsx', '.xls', '.csv', '.txt']

        if not file_path:
            return False, "Не указан путь к файлу."

        if not os.path.exists(file_path):
            return False, f"Файл не существует: {file_path}"

        if not file_path.lower().endswith(tuple(allowed_extensions)):
            ext_list = ', '.join(allowed_extensions)
            return False, f"Файл не имеет разрешенного расширения ({ext_list})."

        return True, ""

    def get_file_line_count(self, file_path, encoding='utf-8', separator=','):
        """Подсчитывает количество строк в файле для отображения прогресса"""
        try:
            if file_path.lower().endswith(('.xlsx', '.xls')):
                # Для Excel файлов используем pandas для получения количества строк
                df_temp = pd.read_excel(file_path, sheet_name=0, nrows=0)
                # Это не даст точного количества строк, но мы можем использовать другой подход
                # для Excel файлов мы будем использовать общее количество строк из листа
                df_full = pd.read_excel(file_path, sheet_name=0, header=None)
                return len(df_full)
            else:
                # Для CSV и TXT файлов подсчитываем строки напрямую
                # Сначала пробуем с основной кодировкой
                try:
                    with open(file_path, 'r', encoding=encoding, errors='replace') as f:
                        count = sum(1 for line in f)
                    return count - 1  # Вычитаем 1, потому что первая строка может быть заголовком
                except UnicodeDecodeError:
                    # Если основная кодировка не работает, пробуем fallback кодировки
                    fallback_encodings = ['utf-8-sig', 'cp1251', 'iso-8859-1']
                    for enc in fallback_encodings:
                        try:
                            with open(file_path, 'r', encoding=enc, errors='replace') as f:
                                count = sum(1 for line in f)
                            return count - 1  # Вычитаем 1, потому что первая строка может быть заголовком
                        except UnicodeDecodeError:
                            continue
                        except UnicodeError:
                            continue
                        except Exception:
                            continue
                except UnicodeError:
                    # Если основная кодировка не работает, пробуем fallback кодировки
                    fallback_encodings = ['utf-8-sig', 'cp1251', 'iso-8859-1']
                    for enc in fallback_encodings:
                        try:
                            with open(file_path, 'r', encoding=enc, errors='replace') as f:
                                count = sum(1 for line in f)
                            return count - 1  # Вычитаем 1, потому что первая строка может быть заголовком
                        except UnicodeDecodeError:
                            continue
                        except UnicodeError:
                            continue
                        except Exception:
                            continue
                except Exception:
                    # Если основная кодировка не работает, пробуем fallback кодировки
                    fallback_encodings = ['utf-8-sig', 'cp1251', 'iso-8859-1']
                    for enc in fallback_encodings:
                        try:
                            with open(file_path, 'r', encoding=enc, errors='replace') as f:
                                count = sum(1 for line in f)
                            return count - 1  # Вычитаем 1, потому что первая строка может быть заголовком
                        except UnicodeDecodeError:
                            continue
                        except UnicodeError:
                            continue
                        except Exception:
                            continue

                # Если ни одна кодировка не сработала, возвращаем None
                return None
        except:
            # Если не удалось определить количество строк, возвращаем None
            return None

    def clean_table_name(self, name):
        """Очищает имя таблицы для использования в SQL"""
        cleaned = "".join(c for c in name if c.isalnum() or c == '_')
        if not cleaned or cleaned[0].isdigit():
            cleaned = f"table_{cleaned}" if cleaned else "table_data"
        return cleaned

    def remove_empty_columns(self, df):
        """Удаляет пустые столбцы из DataFrame"""
        # Проверяем, нужно ли удалять пустые столбцы согласно настройкам
        if not self.settings.get('remove_empty_cols', True):
            self.log("Удаление пустых столбцов отключено в настройках.")
            return df, []

        initial_column_count = len(df.columns)
        empty_columns = []

        for col in df.columns:
            # Проверяем, все ли значения в столбце являются null или NaN
            if df[col].isna().all():
                empty_columns.append(col)
            else:
                # Для не-null столбцов проверяем, являются ли они строковыми и если да, то проверяем на пробелы
                if df[col].dtype == "object":  # Объектные столбцы обычно содержат строки
                    try:
                        # Проверяем, все ли ненулевые значения содержат только пробелы
                        non_null_values = df[col].dropna()
                        if len(non_null_values) > 0 and non_null_values.astype(str).str.strip().eq('').all():
                            empty_columns.append(col)
                    except:
                        # Если возникла ошибка при обработке строк, пропускаем эту проверку
                        # и полагаемся только на проверку isna().all()
                        pass

        if empty_columns:
            self.log(f"Найдено {len(empty_columns)} пустых столбцов: {empty_columns}")
            df = df.drop(columns=empty_columns)
            self.log(f"Столбцов уменьшено с {initial_column_count} до {len(df.columns)}")
        else:
            self.log("Пустые столбцы не найдены.")

        return df, empty_columns

    def _detect_csv_params(self, file_path):
        """
        Определяет наиболее вероятные параметры CSV файла (разделитель)
        Кодировка будет передана отдельно
        """
        import csv

        # Определяем возможные разделители
        separators = [',', ';', '\t', '|', ':', ' ']

        best_params = (None, None)  # (encoding, separator)
        max_columns = 0

        # Пробуем разные разделители с UTF-8 (основная кодировка)
        for sep in separators:
            try:
                # Читаем только заголовки (первую строку)
                df_sample = pd.read_csv(file_path, encoding='utf-8', sep=sep, nrows=0)

                # Если получили больше одного столбца, это хороший признак
                if len(df_sample.columns) > max_columns:
                    max_columns = len(df_sample.columns)
                    best_params = ('utf-8', sep)

                    # Если уже получили хотя бы 2 столбца, считаем это хорошим результатом
                    if max_columns >= 2:
                        break
            except UnicodeDecodeError:
                continue
            except Exception:
                continue  # любая другая ошибка - пробуем следующую комбинацию

        # Если не нашли подходящую комбинацию, возвращаем UTF-8 с запятой
        if best_params[0] is None:
            return ('utf-8', ',')

        return best_params[0], best_params[1]

    def convert_file_to_db(self, input_file_path, db_file_path, column_mapping=None, use_first_row_as_header=False, txt_delimiter=None, encoding='utf-8', progress_callback=None):
        """Основной метод конвертации Excel, CSV или TXT файла в SQLite базу данных с поддержкой больших файлов"""
        self.log_messages = []  # Очистка предыдущих сообщений

        # Определение типа файла и проверка
        if input_file_path.lower().endswith(('.xlsx', '.xls')):
            # Это Excel файл
            is_valid, error_msg = self.validate_file_path(input_file_path, ['.xlsx', '.xls'])
            file_type = "Excel"
        elif input_file_path.lower().endswith('.csv'):
            # Это CSV файл
            is_valid, error_msg = self.validate_file_path(input_file_path, ['.csv'])
            file_type = "CSV"
        elif input_file_path.lower().endswith('.txt'):
            # Это TXT файл
            is_valid, error_msg = self.validate_file_path(input_file_path, ['.txt'])
            file_type = "TXT"
        else:
            raise ValueError(f"Неподдерживаемый тип файла: {input_file_path}")

        if not is_valid:
            raise ValueError(error_msg)

        # Создание имени таблицы из имени файла
        table_name = os.path.splitext(os.path.basename(input_file_path))[0]
        table_name = self.clean_table_name(table_name)

        # Проверка пути к базе данных
        db_dir = os.path.dirname(db_file_path)
        if db_dir and not os.path.exists(db_dir):
            raise FileNotFoundError(f"Директория не существует: {db_dir}")

        # Проверка, существует ли файл базы данных и нужно ли перезаписывать
        if os.path.exists(db_file_path) and not self.settings.get('overwrite_existing', True):
            raise FileExistsError(f"Файл базы данных уже существует: {db_file_path}. Перезапись отключена в настройках.")

        # Определение размера чанка
        chunk_size = self.settings.get('chunk_size', 10000)

        self.log(f"Начинаем конвертацию {file_type} файла: {input_file_path}")
        self.log(f"Размер чанка: {chunk_size}")
        self.log(f"Кодировка: {encoding}")

        # Подсчет количества строк для отображения прогресса (если возможно)
        total_rows = self.get_file_line_count(input_file_path, encoding=encoding)
        if total_rows:
            self.log(f"Обнаружено примерно {total_rows} строк в файле")
            total_chunks = math.ceil(total_rows / chunk_size) if total_rows > 0 else 1
        else:
            total_chunks = None
            self.log("Не удалось определить общее количество строк, прогресс будет приблизительным")

        try:
            if file_type == "Excel":
                # Для Excel файлов обработка происходит по-другому, так как pandas не поддерживает chunksize для Excel
                self.log("Обработка Excel файла...")

                # Сначала получаем оригинальные названия столбцов (из первой строки), как если бы читали с заголовками
                df_with_headers = pd.read_excel(input_file_path, sheet_name=0, nrows=0)
                original_columns = list(df_with_headers.columns)

                # Затем читаем весь файл без заголовков, чтобы контролировать, как обрабатывается первая строка
                df_full = pd.read_excel(input_file_path, sheet_name=0, header=None)

                if not use_first_row_as_header:
                    # Если не использовать первую строку как заголовки, используем стандартные имена
                    df = df_full.copy()
                    df.columns = [f"Column_{i}" for i in range(len(df.columns))]
                else:
                    # Если использовать первую строку как заголовки, используем оригинальные имена для переименования
                    df = df_full.copy()
                    df.columns = original_columns  # Используем оригинальные имена (значения из первой строки)

                # Применение переименования столбцов, если предоставлены соответствия
                if column_mapping:
                    self.log("Применение переименования столбцов...")
                    df = self.rename_columns(df, column_mapping)

                # Использовать первую строку как заголовки, если указано
                if use_first_row_as_header:
                    self.log("Применение первой строки как заголовков столбцов...")
                    df = self.use_first_row_as_header(df)

                # Удаление пустых столбцов
                df, removed_columns = self.remove_empty_columns(df)

                # Сохранение в базу данных
                self.log(f"Сохранение в базу данных: {db_file_path}")

                # Обновляем прогресс, если предоставлен callback
                if progress_callback:
                    progress_callback(50, "Сохранение данных в базу данных...")

                conn = sqlite3.connect(db_file_path)
                df.to_sql(table_name, conn, index=False, if_exists='replace')
                conn.commit()
                conn.close()

                # Обновляем прогресс завершения
                if progress_callback:
                    progress_callback(100, "Конвертация завершена успешно!")

                self.log(f"Данные сохранены в таблицу: {table_name}")
                self.log(f"Всего записей: {len(df)}")
                self.log(f"Всего столбцов: {len(df.columns)}")

            elif file_type in ["CSV", "TXT"]:
                # Для CSV и TXT файлов используем chunksize для потоковой обработки
                self.log(f"Обработка {file_type} файла с использованием чанков...")

                # Определяем параметры для чтения файла (только разделитель, кодировка передана явно)
                if file_type == "CSV":
                    _, detected_separator = self._detect_csv_params(input_file_path)
                else:  # TXT
                    _, detected_separator = self._detect_csv_params(input_file_path)
                    # Если пользователь указал разделитель для TXT, используем его
                    if txt_delimiter:
                        detected_separator = txt_delimiter

                self.log(f"Используем кодировку: {encoding}, разделитель: {detected_separator}")

                # Создаем соединение с базой данных
                conn = sqlite3.connect(db_file_path)

                # Определяем имена столбцов и применяем правила переименования перед чтением чанков.
                try:
                    if use_first_row_as_header:
                        # Сначала определим структуру файла с заголовками, чтобы знать, как должны называться столбцы
                        try:
                            # Попробуем сначала с on_bad_lines='skip' (новые версии pandas)
                            temp_df = pd.read_csv(
                                input_file_path,
                                encoding=encoding,
                                sep=detected_separator,
                                nrows=0,
                                on_bad_lines='skip',
                                encoding_errors='replace'  # Заменяем проблемные символы
                            )
                        except TypeError:
                            try:
                                # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                                temp_df = pd.read_csv(
                                    input_file_path,
                                    encoding=encoding,
                                    sep=detected_separator,
                                    nrows=0,
                                    error_bad_lines=False,
                                    warn_bad_lines=True,
                                    encoding_errors='replace'  # Заменяем проблемные символы
                                )
                            except TypeError:
                                # Если параметр encoding_errors не поддерживается, читаем без него
                                temp_df = pd.read_csv(
                                    input_file_path,
                                    encoding=encoding,
                                    sep=detected_separator,
                                    nrows=0,
                                    error_bad_lines=False,
                                    warn_bad_lines=True
                                )
                        expected_columns = list(temp_df.columns)

                        # Также нужно обработать переименование столбцов, если они заданы
                        if column_mapping:
                            for old_name, new_name in column_mapping.items():
                                if str(old_name) in expected_columns:
                                    idx = expected_columns.index(str(old_name))
                                    expected_columns[idx] = str(new_name)
                    else:
                        # Определяем количество столбцов из первого чанка
                        try:
                            # Попробуем сначала с on_bad_lines='skip' (новые версии pandas)
                            temp_df = pd.read_csv(
                                input_file_path,
                                encoding=encoding,
                                sep=detected_separator,
                                nrows=0,
                                header=None,
                                on_bad_lines='skip',
                                encoding_errors='replace'  # Заменяем проблемные символы
                            )
                        except TypeError:
                            try:
                                # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                                temp_df = pd.read_csv(
                                    input_file_path,
                                    encoding=encoding,
                                    sep=detected_separator,
                                    nrows=0,
                                    header=None,
                                    error_bad_lines=False,
                                    warn_bad_lines=True,
                                    encoding_errors='replace'  # Заменяем проблемные символы
                                )
                            except TypeError:
                                # Если параметр encoding_errors не поддерживается, читаем без него
                                temp_df = pd.read_csv(
                                    input_file_path,
                                    encoding=encoding,
                                    sep=detected_separator,
                                    nrows=0,
                                    header=None,
                                    error_bad_lines=False,
                                    warn_bad_lines=True
                                )
                        expected_columns = [f"Column_{j}" for j in range(len(temp_df.columns))]

                        # Применяем переименование, если задано
                        if column_mapping:
                            for old_idx in range(len(expected_columns)):
                                old_col_name = expected_columns[old_idx]
                                if old_col_name in column_mapping:
                                    expected_columns[old_idx] = str(column_mapping[old_col_name])
                except pd.errors.ParserError as e:
                    self.log(f"Обнаружена ошибка парсинга при определении структуры {file_type} файла: {e}")
                    # Если возникает ошибка парсинга, пробуем снова с пропуском проблемных строк
                    try:
                        # Попробуем снова с on_bad_lines='skip' (новые версии pandas)
                        temp_df = pd.read_csv(
                            input_file_path,
                            encoding=encoding,
                            sep=detected_separator,
                            nrows=0,
                            header=None,
                            on_bad_lines='skip',
                            encoding_errors='replace'  # Заменяем проблемные символы
                        )
                    except TypeError:
                        try:
                            # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                            temp_df = pd.read_csv(
                                input_file_path,
                                encoding=encoding,
                                sep=detected_separator,
                                nrows=0,
                                header=None,
                                error_bad_lines=False,
                                warn_bad_lines=True,
                                encoding_errors='replace'  # Заменяем проблемные символы
                            )
                        except TypeError:
                            # Если параметр encoding_errors не поддерживается, читаем без него
                            temp_df = pd.read_csv(
                                input_file_path,
                                encoding=encoding,
                                sep=detected_separator,
                                nrows=0,
                                header=None,
                                error_bad_lines=False,
                                warn_bad_lines=True
                            )
                    expected_columns = [f"Column_{j}" for j in range(len(temp_df.columns))]
                    self.log(f"Продолжаем обработку {file_type} файла с пропуском проблемных строк")

                # При наличии заголовка pandas использует первую строку только для имён столбцов.
                header = 0 if use_first_row_as_header else None
                try:
                    # Попробуем сначала с on_bad_lines='skip' (новые версии pandas)
                    chunk_reader = pd.read_csv(
                        input_file_path,
                        encoding=encoding,
                        sep=detected_separator,
                        chunksize=chunk_size,
                        header=header,
                        on_bad_lines='skip',
                        encoding_errors='replace'  # Заменяем проблемные символы
                    )
                except TypeError:
                    try:
                        # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                        chunk_reader = pd.read_csv(
                            input_file_path,
                            encoding=encoding,
                            sep=detected_separator,
                            chunksize=chunk_size,
                            header=header,
                            error_bad_lines=False,
                            warn_bad_lines=True,
                            encoding_errors='replace'  # Заменяем проблемные символы
                        )
                    except TypeError:
                        # Если параметр encoding_errors не поддерживается, читаем без него
                        chunk_reader = pd.read_csv(
                            input_file_path,
                            encoding=encoding,
                            sep=detected_separator,
                            chunksize=chunk_size,
                            header=header,
                            on_bad_lines='skip'
                        )
                except TypeError:
                    # Если параметр encoding_errors не поддерживается, читаем без него
                    chunk_reader = pd.read_csv(
                        input_file_path,
                        encoding=encoding,
                        sep=detected_separator,
                        chunksize=chunk_size,
                        header=header,
                        error_bad_lines=False,
                        warn_bad_lines=True
                    )

                # Инициализируем прогресс, если предоставлен callback
                if progress_callback:
                    progress_callback(0, f"Начало обработки {file_type} файла...")

                # Создаем tqdm прогресс-бар (только для консоли, не для GUI)
                if total_chunks:
                    pbar = tqdm(total=total_chunks, desc="Обработка чанков", unit="чанк")
                else:
                    pbar = tqdm(desc="Обработка чанков", unit="чанк")

                try:
                    first_chunk = True
                    processed_chunks = 0

                    for i, chunk in enumerate(chunk_reader):
                        # Устанавливаем правильные имена столбцов
                        chunk.columns = expected_columns[:len(chunk.columns)]  # На случай, если в последнем чанке меньше столбцов

                        # Заголовок задаёт схему таблицы, даже если первый чанк пуст
                        # или значения некоторых столбцов появятся только в следующих чанках.
                        if i > 0 or not use_first_row_as_header:
                            chunk, _ = self.remove_empty_columns(chunk)

                        # Сохраняем чанк в базу данных
                        if i == 0:
                            # Для первого чанка создаем таблицу
                            chunk.to_sql(table_name, conn, index=False, if_exists='replace')
                        else:
                            # Для последующих чанков добавляем к существующей таблице
                            chunk.to_sql(table_name, conn, index=False, if_exists='append')

                        # Обновляем прогресс-бар
                        pbar.update(1)

                        # Обновляем прогресс для GUI, если предоставлен callback
                        if progress_callback and total_chunks:
                            progress_percent = min(int((i + 1) / total_chunks * 100), 100)
                            progress_callback(progress_percent, f"Обработка чанка {i + 1} из {total_chunks}...")

                        # Логируем прогресс каждые 10 чанков
                        if (i + 1) % 10 == 0:
                            self.log(f"Обработано чанков: {i + 1}")

                        processed_chunks = i + 1

                finally:
                    pbar.close()
                    conn.close()

                    # Обновляем прогресс завершения
                    if progress_callback:
                        progress_callback(100, f"Обработка {file_type} файла завершена. Обработано {processed_chunks} чанков.")

                self.log(f"Обработка {file_type} файла завершена")
                self.log(f"Данные сохранены в таблицу: {table_name}")

            self.log("Конвертация завершена успешно!")

            # Формирование результата
            result = "\n".join(self.log_messages)
            return result

        except UnicodeDecodeError as e:
            raise Exception(f"Ошибка декодирования при обработке {file_type} файла: {e}. Попробуйте изменить кодировку файла.")
        except FileNotFoundError:
            raise FileNotFoundError(f"Файл не найден: {input_file_path}")
        except PermissionError:
            raise PermissionError(f"Нет доступа к файлу: {input_file_path}")
        except pd.errors.EmptyDataError:
            raise ValueError(f"{file_type} файл пустой или не содержит данных.")
        except Exception as e:
            raise Exception(f"Ошибка при чтении {file_type} файла: {e}")

    def convert_xlsx_to_db(self, excel_file_path, db_file_path, column_mapping=None):
        """Метод конвертации Excel файла в SQLite базу данных (сохранен для обратной совместимости)"""
        return self.convert_file_to_db(excel_file_path, db_file_path, column_mapping)

    def rename_columns(self, df, column_mapping):
        """Переименование столбцов в DataFrame"""
        self.log("Начинаем процесс переименования столбцов...")
        self.log(f"Типы столбцов в DataFrame: {[(col, type(col)) for col in df.columns]}")
        self.log(f"Все столбцы в DataFrame: {list(df.columns)}")
        self.log(f"Запрошено переименование: {column_mapping}")

        # Проверяем, какие столбцы можно переименовать
        existing_columns = set(str(col) for col in df.columns)  # Преобразуем все в строки для сравнения
        renameable_columns = {}

        for old_name, new_name in column_mapping.items():
            # Преобразуем old_name в строку для поиска в списке столбцов
            old_name_str = str(old_name)

            # Если имя содержит образец (например, "Column_0 (образец: значение)"),
            # извлекаем только основное имя столбца
            if " (образец:" in old_name_str:
                actual_old_name = old_name_str.split(" (образец:")[0]
            else:
                actual_old_name = old_name_str

            # Проверяем, есть ли столбец с таким именем в DataFrame
            matched_col = None
            for col in df.columns:
                col_str = str(col)
                # Проверяем как полное совпадение, так и совпадение без образца
                if col_str == actual_old_name or (col_str.startswith(actual_old_name + " (образец:") if " (образец:" in col_str else col_str == actual_old_name):
                    matched_col = col
                    break

            if matched_col is not None:
                # Проверяем, не пустое ли новое имя
                if new_name.strip():
                    renameable_columns[matched_col] = new_name
                    self.log(f"Столбец '{matched_col}' будет переименован в '{new_name}'")
                else:
                    self.log(f"Пропущено переименование столбца '{old_name}': новое имя пустое")
            else:
                self.log(f"Столбец '{old_name}' не найден в данных для переименования (поиск как '{actual_old_name}')")

        if renameable_columns:
            df = df.rename(columns=renameable_columns)
            self.log(f"Переименовано {len(renameable_columns)} столбцов: {renameable_columns}")
        else:
            self.log("Не найдено столбцов для переименования.")

        self.log(f"Финальные столбцы в DataFrame: {list(df.columns)}")
        return df

    def use_first_row_as_header(self, df):
        """Использовать первую строку как заголовки столбцов"""
        self.log("Используем первую строку как заголовки столбцов...")

        if df.empty:
            self.log("DataFrame пустой, нечего обрабатывать")
            return df

        # Сохраняем текущие названия столбцов
        old_columns = df.columns.tolist()

        # Используем первую строку как новые названия столбцов
        new_header = df.iloc[0]  # первая строка

        # Создаем маппинг старых имен в новые значения из первой строки
        column_mapping = {}
        for old_col in old_columns:
            # Получаем значение из первой строки для текущего столбца
            new_name_from_first_row = new_header[old_col]
            column_mapping[old_col] = new_name_from_first_row

        df = df.rename(columns=column_mapping)

        # Удаляем первую строку из данных (теперь она стала заголовком)
        df = df.iloc[1:].reset_index(drop=True)

        self.log(f"Старые названия столбцов: {old_columns}")
        self.log(f"Новые названия столбцов (из первой строки): {list(df.columns)}")
        self.log(f"Количество строк после обработки: {len(df)}")

        return df
