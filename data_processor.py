import pandas as pd
import sqlite3
import os


class DataProcessor:
    def __init__(self, settings=None):
        self.log_messages = []
        self.settings = settings or {
            'remove_empty_cols': True,
            'overwrite_existing': True
        }
        
    def log(self, message):
        """Добавляет сообщение в лог и выводит его"""
        self.log_messages.append(message)
        print(message)
        
    def validate_file_path(self, file_path, allowed_extensions=None):
        """Проверяет существование файла и его расширение"""
        if allowed_extensions is None:
            allowed_extensions = ['.xlsx', '.xls']

        if not file_path:
            return False, "Не указан путь к файлу."

        if not os.path.exists(file_path):
            return False, f"Файл не существует: {file_path}"

        if not file_path.lower().endswith(tuple(allowed_extensions)):
            ext_list = ', '.join(allowed_extensions)
            return False, f"Файл не имеет разрешенного расширения ({ext_list})."

        return True, ""

    def validate_csv_path(self, file_path):
        """Проверяет существование CSV файла"""
        return self.validate_file_path(file_path, ['.csv'])

    def validate_txt_path(self, file_path):
        """Проверяет существование TXT файла"""
        return self.validate_file_path(file_path, ['.txt'])
        
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
        Определяет наиболее вероятные параметры CSV файла (кодировку и разделитель)
        """
        import csv
        try:
            import chardet
            chardet_available = True
        except ImportError:
            chardet_available = False

        if chardet_available:
            # Определяем кодировку файла
            with open(file_path, 'rb') as f:
                raw_data = f.read(10000)  # читаем первые 10KB для определения кодировки
                encoding_result = chardet.detect(raw_data)
                detected_encoding = encoding_result['encoding']

                # Если chardet не смог определить кодировку или результат ненадежный, используем наши варианты
                if detected_encoding is None or encoding_result['confidence'] < 0.7:
                    encodings = ['utf-8', 'utf-8-sig', 'cp1251', 'iso-8859-1']
                else:
                    encodings = [detected_encoding, 'utf-8', 'utf-8-sig', 'cp1251', 'iso-8859-1']
        else:
            # Если chardet недоступен, используем стандартные кодировки
            encodings = ['utf-8', 'utf-8-sig', 'cp1251', 'iso-8859-1']

        # Определяем возможные разделители
        separators = [',', ';', '\t', '|', ':', ' ']

        best_params = (None, None)  # (encoding, separator)
        max_columns = 0

        # Пробуем разные комбинации
        for encoding in encodings:
            for sep in separators:
                try:
                    # Читаем только заголовки (первую строку)
                    df_sample = pd.read_csv(file_path, encoding=encoding, sep=sep, nrows=0)

                    # Если получили больше одного столбца, это хороший признак
                    if len(df_sample.columns) > max_columns:
                        max_columns = len(df_sample.columns)
                        best_params = (encoding, sep)

                        # Если уже получили хотя бы 2 столбца, считаем это хорошим результатом
                        if max_columns >= 2:
                            break
                except UnicodeDecodeError:
                    continue
                except Exception:
                    continue  # любая другая ошибка - пробуем следующую комбинацию

            # Если уже нашли хороший результат, можно прекратить
            if max_columns >= 2:
                break

        # Если не нашли подходящую комбинацию, используем pandas по умолчанию
        if best_params[0] is None:
            # Попробуем pandas автоопределение
            try:
                df_sample = pd.read_csv(file_path, nrows=0)
                # pandas обычно использует запятую как разделитель по умолчанию
                return ('utf-8', ',')  # возвращаем стандартные параметры
            except:
                # Если ничего не работает, возвращаем стандартные значения
                return ('utf-8', ',')

        return best_params[0], best_params[1]
        
    def convert_file_to_db(self, input_file_path, db_file_path, column_mapping=None, use_first_row_as_header=False, txt_delimiter=None):
        """Основной метод конвертации Excel, CSV или TXT файла в SQLite базу данных"""
        self.log_messages = []  # Очистка предыдущих сообщений

        # Определение типа файла и проверка
        if input_file_path.lower().endswith(('.xlsx', '.xls')):
            # Это Excel файл
            is_valid, error_msg = self.validate_file_path(input_file_path)
            file_type = "Excel"
        elif input_file_path.lower().endswith('.csv'):
            # Это CSV файл
            is_valid, error_msg = self.validate_csv_path(input_file_path)
            file_type = "CSV"
        elif input_file_path.lower().endswith('.txt'):
            # Это TXT файл
            is_valid, error_msg = self.validate_txt_path(input_file_path)
            file_type = "TXT"
        else:
            raise ValueError(f"Неподдерживаемый тип файла: {input_file_path}")

        if not is_valid:
            raise ValueError(error_msg)

        # Чтение файла в зависимости от типа
        self.log(f"Загрузка {file_type} файла: {input_file_path}")
        try:
            if file_type == "Excel":
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
                    # Теперь переименование будет работать с этими именами, и потом use_first_row_as_header
                    # возьмет эти значения как новые заголовки
            elif file_type == "CSV":
                # Определяем наиболее вероятный разделитель и кодировку
                detected_encoding, detected_separator = self._detect_csv_params(input_file_path)

                # Сначала получаем оригинальные названия столбцов (из первой строки), как если бы читали с заголовками
                df_with_headers = pd.read_csv(input_file_path, encoding=detected_encoding, sep=detected_separator, nrows=0)
                original_columns = list(df_with_headers.columns)

                # Затем читаем весь файл без заголовков, чтобы контролировать, как обрабатывается первая строка
                df_full = pd.read_csv(input_file_path, encoding=detected_encoding, sep=detected_separator, header=None)

                if not use_first_row_as_header:
                    # Если не использовать первую строку как заголовки, используем стандартные имена
                    df = df_full.copy()
                    df.columns = [f"Column_{i}" for i in range(len(df.columns))]
                else:
                    # Если использовать первую строку как заголовки, используем оригинальные имена для переименования
                    df = df_full.copy()
                    df.columns = original_columns  # Используем оригинальные имена (значения из первой строки)
                    # Теперь переименование будет работать с этими именами, и потом use_first_row_as_header
                    # возьмет эти значения как новые заголовки

                # Логируем использование нестандартного разделителя
                if detected_separator != ',':
                    self.log(f"Использован разделитель '{detected_separator}' для чтения CSV файла")
                elif detected_encoding != 'utf-8':
                    self.log(f"Использована кодировка {detected_encoding} для чтения CSV файла")
            elif file_type == "TXT":
                # Для TXT файлов используем переданный разделитель или определяем автоматически
                detected_encoding, detected_separator = self._detect_csv_params(input_file_path)

                # Если пользователь указал разделитель для TXT, используем его
                if txt_delimiter:
                    separator_to_use = txt_delimiter
                    encoding_to_use = detected_encoding
                else:
                    separator_to_use = detected_separator
                    encoding_to_use = detected_encoding

                try:
                    # Сначала получаем оригинальные названия столбцов (из первой строки), как если бы читали с заголовками
                    try:
                        # Попробуем сначала с on_bad_lines='skip' (новые версии pandas)
                        df_with_headers = pd.read_csv(input_file_path, encoding=encoding_to_use, sep=separator_to_use, nrows=0, on_bad_lines='skip')
                    except TypeError:
                        # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                        df_with_headers = pd.read_csv(input_file_path, encoding=encoding_to_use, sep=separator_to_use, nrows=0, error_bad_lines=False, warn_bad_lines=True)
                    original_columns = list(df_with_headers.columns)

                    # Затем читаем весь файл без заголовков, чтобы контролировать, как обрабатывается первая строка
                    # Используем параметр error_bad_lines=False (для старых версий pandas) или on_bad_lines='skip' (для новых)
                    try:
                        # Попробуем сначала с on_bad_lines='skip' (новые версии pandas)
                        df_full = pd.read_csv(input_file_path, encoding=encoding_to_use, sep=separator_to_use, header=None, on_bad_lines='skip')
                    except TypeError:
                        # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                        df_full = pd.read_csv(input_file_path, encoding=encoding_to_use, sep=separator_to_use, header=None, error_bad_lines=False, warn_bad_lines=True)

                    if not use_first_row_as_header:
                        # Если не использовать первую строку как заголовки, используем стандартные имена
                        df = df_full.copy()
                        df.columns = [f"Column_{i}" for i in range(len(df.columns))]
                    else:
                        # Если использовать первую строку как заголовки, используем оригинальные имена для переименования
                        df = df_full.copy()
                        df.columns = original_columns  # Используем оригинальные имена (значения из первой строки)
                        # Теперь переименование будет работать с этими именами, и потом use_first_row_as_header
                        # возьмет эти значения как новые заголовки

                except pd.errors.ParserError as e:
                    # Если возникает ошибка парсинга, пробуем снова с пропуском проблемных строк
                    try:
                        # Попробуем снова с on_bad_lines='skip' (новые версии pandas)
                        df_full = pd.read_csv(input_file_path, encoding=encoding_to_use, sep=separator_to_use, header=None, on_bad_lines='skip')
                    except TypeError:
                        # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                        df_full = pd.read_csv(input_file_path, encoding=encoding_to_use, sep=separator_to_use, header=None, error_bad_lines=False, warn_bad_lines=True)

                    if not use_first_row_as_header:
                        # Если не использовать первую строку как заголовки, используем стандартные имена
                        df = df_full.copy()
                        df.columns = [f"Column_{i}" for i in range(len(df.columns))]
                    else:
                        # Если использовать первую строку как заголовки, используем оригинальные имена для переименования
                        # Но в случае ошибки парсинга, используем стандартные имена
                        df = df_full.copy()
                        df.columns = [f"Column_{i}" for i in range(len(df.columns))]

                    self.log(f"Обнаружены и пропущены строки с ошибками формата в TXT файле")

                # Логируем использование разделителя
                self.log(f"Использован разделитель '{separator_to_use}' для чтения TXT файла")
                if encoding_to_use != 'utf-8':
                    self.log(f"Использована кодировка {encoding_to_use} для чтения TXT файла")

            self.log(f"Загружено строк: {len(df)}, столбцов: {len(df.columns)}")

            if df.empty:
                raise ValueError(f"{file_type} файл пустой.")

        except FileNotFoundError:
            raise FileNotFoundError(f"Файл не найден: {input_file_path}")
        except PermissionError:
            raise PermissionError(f"Нет доступа к файлу: {input_file_path}")
        except pd.errors.EmptyDataError:
            raise ValueError(f"{file_type} файл пустой или не содержит данных.")
        except Exception as e:
            raise Exception(f"Ошибка при чтении {file_type} файла: {e}")

        # Использовать первую строку как заголовки, если указано
        if use_first_row_as_header:
            self.log("Применение первой строки как заголовков столбцов...")
            df = self.use_first_row_as_header(df)

        # Переименование столбцов, если предоставлены соответствия
        if column_mapping:
            self.log("Применение переименования столбцов...")
            df = self.rename_columns(df, column_mapping)

        # Удаление пустых столбцов
        df, removed_columns = self.remove_empty_columns(df)

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

        # Сохранение в базу данных
        self.log(f"Сохранение в базу данных: {db_file_path}")
        try:
            conn = sqlite3.connect(db_file_path)
            df.to_sql(table_name, conn, index=False, if_exists='replace')
            conn.commit()
            conn.close()

            self.log(f"Данные сохранены в таблицу: {table_name}")
            self.log(f"Всего записей: {len(df)}")
            self.log(f"Всего столбцов: {len(df.columns)}")

            # Формирование результата
            result = "\n".join(self.log_messages)
            return result

        except sqlite3.Error as e:
            raise Exception(f"Ошибка базы данных: {e}")
        except PermissionError:
            raise PermissionError(f"Нет доступа к файлу базы данных: {db_file_path}")
        except Exception as e:
            raise Exception(f"Ошибка при сохранении в базу данных: {e}")

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