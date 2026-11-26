import tkinter as tk
from tkinter import ttk, messagebox, filedialog
from data_processor import DataProcessor
from helper_windows import SettingsWindow, RenameColumnsWindow
import threading
import os
import pandas as pd


class ExcelToDbApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Конвертер (xls/csv/txt в db)")
        self.root.geometry("800x600")

        # Настройки по умолчанию
        self.settings = {
            'remove_empty_cols': True,
            'overwrite_existing': True
        }

        # Инициализация процессора данных
        self.processor = DataProcessor(self.settings)

        # Переменные для хранения путей
        self.input_file_path = tk.StringVar()
        self.output_file_path = tk.StringVar()

        # Переменная для хранения соответствий переименования столбцов
        self.column_mapping = {}

        # Тип конвертации (только Excel/CSV to DB)
        self.conversion_type = tk.StringVar(value="file_to_db")  # file_to_db

        # Переменная для хранения разделителя для TXT файлов
        self.txt_delimiter = tk.StringVar(value=',')  # По умолчанию запятая

        # Создание интерфейса
        self.create_widgets()
        
    def create_widgets(self):
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Заголовок
        title_label = ttk.Label(main_frame, text="Конвертер (xls/csv/txt в db)", font=("Arial", 16))
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Выбор входного файла
        ttk.Label(main_frame, text="XLS/CSV/TXT файл:").grid(row=1, column=0, sticky=tk.W, pady=5)
        input_entry = ttk.Entry(main_frame, textvariable=self.input_file_path, width=50)
        input_entry.grid(row=1, column=1, padx=(5, 5), pady=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Обзор", command=self.browse_input_file).grid(row=1, column=2, pady=5)

        # Выбор выходного файла
        ttk.Label(main_frame, text="Файл БД:").grid(row=2, column=0, sticky=tk.W, pady=5)
        output_entry = ttk.Entry(main_frame, textvariable=self.output_file_path, width=50)
        output_entry.grid(row=2, column=1, padx=(5, 5), pady=5, sticky=(tk.W, tk.E))
        ttk.Button(main_frame, text="Обзор", command=self.browse_output_file).grid(row=2, column=2, pady=5)

        # Кнопки действий (кроме конвертировать)
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=3, column=0, columnspan=3, pady=10)

        ttk.Button(button_frame, text="Просмотр/переименование столбцов", command=self.view_and_rename_columns).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Выбор разделителя TXT", command=self.select_txt_delimiter_conditionally).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Настройки", command=self.open_settings).pack(side=tk.LEFT)

        # Кнопка конвертировать над полем с логами
        self.convert_button = ttk.Button(main_frame, text="Конвертировать", command=self.start_conversion)
        self.convert_button.grid(row=4, column=0, columnspan=3, pady=10)

        # Текстовое поле для вывода информации
        self.info_text = tk.Text(main_frame, height=15, width=80)
        self.info_text.grid(row=5, column=0, columnspan=3, pady=10)

        # Настройка растягивания элементов
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(5, weight=1)

        # Инициализация интерфейса - установка правильных меток
        self.on_conversion_type_change()
        
    def on_conversion_type_change(self):
        """Обновляет интерфейс - устанавливает правильные метки"""
        # Обновляем текст меток для Excel/CSV в DB
        for child in self.root.winfo_children():
            if isinstance(child, ttk.Frame):
                for widget in child.winfo_children():
                    if isinstance(widget, ttk.Label) and widget.cget("text") == "Входной файл:":
                        widget.config(text="Excel/CSV файл:")
                    elif isinstance(widget, ttk.Label) and widget.cget("text") == "Выходной файл:":
                        widget.config(text="Файл БД:")

    def browse_input_file(self):
        """Выбор входного файла для конвертации в DB"""
        file_path = filedialog.askopenfilename(
            title="Выберите XLS, CSV или TXT файл",
            filetypes=[("Все поддерживаемые файлы", "*.xlsx *.xls *.csv *.txt"), ("Excel файлы", "*.xlsx *.xls"), ("CSV файлы", "*.csv"), ("TXT файлы", "*.txt"), ("Все файлы", "*.*")]
        )

        if file_path:
            self.input_file_path.set(file_path)

            # Если выбран TXT файл, автоматически открыть окно выбора разделителя
            if file_path.lower().endswith('.txt'):
                self.select_txt_delimiter()

            # Автоматически открыть вкладку просмотра/переименования столбцов
            self.view_and_rename_columns()

    def select_txt_delimiter_conditionally(self):
        """Открывает диалог выбора разделителя только если выбран TXT файл"""
        input_file = self.input_file_path.get()
        if not input_file or not input_file.lower().endswith('.txt'):
            messagebox.showwarning("Предупреждение", "Эта функция доступна только для TXT файлов")
            return

        self.select_txt_delimiter()

    def browse_output_file(self):
        """Выбор выходного файла (только для DB)"""
        file_path = filedialog.asksaveasfilename(
            title="Сохранить базу данных как",
            defaultextension=".db",
            filetypes=[("SQLite файлы", "*.db"), ("Все файлы", "*.*")]
        )

        if file_path:
            self.output_file_path.set(file_path)
            
    def start_conversion(self):
        # Проверка заполнения полей в зависимости от типа конвертации
        if not self.input_file_path.get():
            messagebox.showerror("Ошибка", "Пожалуйста, выберите входной файл")
            return

        if not self.output_file_path.get():
            messagebox.showerror("Ошибка", "Пожалуйста, укажите выходной файл")
            return

        # Запуск процесса конвертации в отдельном потоке
        self.convert_button.config(state='disabled')

        thread = threading.Thread(target=self.run_conversion)
        thread.daemon = True
        thread.start()
        
    def run_conversion(self):
        try:
            # Очистка текстового поля
            self.info_text.delete(1.0, tk.END)

            # Проверяем, установлен ли флаг использования первой строки как заголовков
            use_header = getattr(self, 'use_first_row_as_header', False)

            # Определяем, является ли файл TXT
            is_txt_file = self.input_file_path.get().lower().endswith('.txt')

            # Вызов метода конвертации из процессора данных (только file_to_db)
            if is_txt_file:
                result = self.processor.convert_file_to_db(
                    self.input_file_path.get(),
                    self.output_file_path.get(),
                    self.column_mapping,
                    use_header,
                    txt_delimiter=self.txt_delimiter.get()
                )
            else:
                result = self.processor.convert_file_to_db(
                    self.input_file_path.get(),
                    self.output_file_path.get(),
                    self.column_mapping,
                    use_header
                )

            # Обновление интерфейса в основном потоке
            self.root.after(0, self.on_conversion_complete, result)

        except Exception as e:
            self.root.after(0, self.on_conversion_error, str(e))
    
    def on_conversion_complete(self, result):
        self.convert_button.config(state='enabled')

        # Отображение результата
        self.info_text.insert(tk.END, result)
        messagebox.showinfo("Успех", "Конвертация завершена успешно!")

    def on_conversion_error(self, error_message):
        self.convert_button.config(state='enabled')

        # Отображение ошибки
        self.info_text.insert(tk.END, f"Ошибка: {error_message}\n")
        messagebox.showerror("Ошибка", f"Произошла ошибка: {error_message}")

    def open_settings(self):
        """Открывает окно настроек"""
        def update_settings(settings):
            self.settings.update(settings)
            self.update_processor_settings()

        SettingsWindow(self.root, update_settings)

    def view_and_rename_columns(self):
        """Открывает окно для просмотра и переименования столбцов"""
        input_file = self.input_file_path.get()
        if not input_file:
            messagebox.showwarning("Предупреждение", "Сначала выберите входной файл")
            return

        try:
            # Загружаем только заголовки столбцов для отображения в зависимости от типа файла
            if input_file.lower().endswith('.csv'):
                # Для CSV файлов используем ту же логику, что и в _get_csv_columns
                try:
                    # Используем ту же логику определения параметров, что и в DataProcessor
                    # Для интерфейса переименования столбцов, мы всегда хотим получить текущие имена столбцов
                    # как если бы файл был прочитан с заголовками (по умолчанию)
                    columns = self._get_csv_columns(input_file)
                    if columns is None:
                        messagebox.showerror("Ошибка", "Не удалось определить структуру CSV файла")
                        return

                    # Загрузим первую строку данных для отображения в интерфейсе
                    detected_encoding, detected_separator = self._detect_encoding_and_separator(input_file)
                    try:
                        df_first_row = pd.read_csv(input_file, encoding=detected_encoding, sep=detected_separator, nrows=1, header=None)
                        first_row_values = df_first_row.iloc[0] if not df_first_row.empty else []
                    except:
                        first_row_values = []

                    # Для правильного сопоставления имен столбцов, нужно понимать,
                    # как будут именоваться столбцы в реальном DataFrame
                    # Если use_first_row_as_header не установлен (по умолчанию False),
                    # то столбцы будут иметь индексные имена
                    if not getattr(self, 'use_first_row_as_header', False):
                        # Показываем пользователю стандартные имена с образцами из первой строки
                        # когда header=None в процессе конвертации
                        try:
                            # Определим параметры и загрузим один ряд с header=None чтобы увидеть реальные имена столбцов
                            df_no_header = pd.read_csv(input_file, encoding=detected_encoding, sep=detected_separator, header=None, nrows=0)
                            columns = []
                            for i in range(len(df_no_header.columns)):
                                sample_val = str(first_row_values[i])[:20] + "..." if i < len(first_row_values) and len(str(first_row_values[i])) > 20 else str(first_row_values[i]) if i < len(first_row_values) else ""
                                columns.append(f"Column_{i} (образец: {sample_val})")
                        except:
                            # Если не удалось, используем стандартные имена на основе количества столбцов
                            columns = []
                            for i in range(len(columns)):
                                sample_val = str(first_row_values[i])[:20] + "..." if i < len(first_row_values) and len(str(first_row_values[i])) > 20 else str(first_row_values[i]) if i < len(first_row_values) else ""
                                columns.append(f"Column_{i} (образец: {sample_val})")
                    else:
                        # Если используется первая строка как заголовки, показываем оригинальные имена
                        # которые были получены ранее, но также добавим значения первой строки для контекста
                        columns_with_samples = []
                        for i, col_name in enumerate(columns):
                            sample_val = str(first_row_values[i])[:20] + "..." if i < len(first_row_values) and len(str(first_row_values[i])) > 20 else str(first_row_values[i]) if i < len(first_row_values) else ""
                            columns_with_samples.append(f"{col_name} (образец: {sample_val})")
                        columns = columns_with_samples

                except Exception as e:
                    messagebox.showerror("Ошибка", f"Ошибка при определении структуры CSV файла: {e}")
                    return
            elif input_file.lower().endswith('.txt'):
                # Для TXT файлов используем выбранный пользователем разделитель
                try:
                    # Для интерфейса переименования столбцов, мы хотим получить текущие имена столбцов
                    # используя выбранный пользователем разделитель
                    detected_encoding, _ = self._detect_encoding_and_separator(input_file)

                    # Используем выбранный пользователем разделитель
                    try:
                        # Попробуем сначала с on_bad_lines='skip' (новые версии pandas)
                        df_sample = pd.read_csv(input_file, encoding=detected_encoding, sep=self.txt_delimiter.get(), nrows=0, on_bad_lines='skip')
                    except TypeError:
                        # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                        df_sample = pd.read_csv(input_file, encoding=detected_encoding, sep=self.txt_delimiter.get(), nrows=0, error_bad_lines=False, warn_bad_lines=True)
                    columns = list(df_sample.columns)

                    # Загрузим первую строку данных для отображения в интерфейсе
                    try:
                        df_first_row = pd.read_csv(input_file, encoding=detected_encoding, sep=self.txt_delimiter.get(), nrows=1, header=None)
                        first_row_values = df_first_row.iloc[0] if not df_first_row.empty else []
                    except:
                        first_row_values = []

                    # Для правильного сопоставления имен столбцов, нужно понимать,
                    # как будут именоваться столбцы в реальном DataFrame
                    # Если use_first_row_as_header не установлен (по умолчанию False),
                    # то столбцы будут иметь имена Column_0, Column_1 и т.д.
                    # Если установлен, то будут использоваться значения из первой строки
                    if not getattr(self, 'use_first_row_as_header', False):
                        # Показываем пользователю стандартные имена с образцами из первой строки
                        columns = []
                        for i in range(len(df_sample.columns)):
                            sample_val = str(first_row_values[i])[:20] + "..." if i < len(first_row_values) and len(str(first_row_values[i])) > 20 else str(first_row_values[i]) if i < len(first_row_values) else ""
                            columns.append(f"Column_{i} (образец: {sample_val})")
                    else:
                        # Если используется первая строка как заголовки, показываем оригинальные имена
                        # с образцами из первой строки
                        columns_with_samples = []
                        for i, col_name in enumerate(columns):
                            sample_val = str(first_row_values[i])[:20] + "..." if i < len(first_row_values) and len(str(first_row_values[i])) > 20 else str(first_row_values[i]) if i < len(first_row_values) else ""
                            columns_with_samples.append(f"{col_name} (образец: {sample_val})")
                        columns = columns_with_samples

                except Exception as e:
                    messagebox.showerror("Ошибка", f"Ошибка при чтении TXT файла с разделителем '{self.txt_delimiter.get()}': {e}")
                    return
            elif input_file.lower().endswith(('.xlsx', '.xls')):
                try:
                    # Читаем Excel файл как с заголовками (по умолчанию)
                    df_with_headers = pd.read_excel(input_file, sheet_name=0, nrows=0)
                    columns = list(df_with_headers.columns)

                    # Загрузим первую строку данных для отображения в интерфейсе
                    try:
                        df_first_row = pd.read_excel(input_file, sheet_name=0, nrows=1, header=None)
                        first_row_values = df_first_row.iloc[0] if not df_first_row.empty else []
                    except:
                        first_row_values = []

                    # Для правильного сопоставления имен столбцов, нужно понимать,
                    # как будут именоваться столбцы в реальном DataFrame
                    # Если use_first_row_as_header не установлен (по умолчанию False),
                    # то столбцы будут иметь индексные имена
                    if not getattr(self, 'use_first_row_as_header', False):
                        # Загружаем Excel файл без заголовков, чтобы получить реальные имена столбцов
                        df_no_header = pd.read_excel(input_file, sheet_name=0, header=None, nrows=0)
                        columns = []
                        for i in range(len(df_no_header.columns)):
                            sample_val = str(first_row_values[i])[:20] + "..." if i < len(first_row_values) and len(str(first_row_values[i])) > 20 else str(first_row_values[i]) if i < len(first_row_values) else ""
                            columns.append(f"Column_{i} (образец: {sample_val})")
                    else:
                        # Если используется первая строка как заголовки, показываем оригинальные имена
                        # с образцами из первой строки
                        columns_with_samples = []
                        for i, col_name in enumerate(columns):
                            sample_val = str(first_row_values[i])[:20] + "..." if i < len(first_row_values) and len(str(first_row_values[i])) > 20 else str(first_row_values[i]) if i < len(first_row_values) else ""
                            columns_with_samples.append(f"{col_name} (образец: {sample_val})")
                        columns = columns_with_samples
                except Exception as e:
                    messagebox.showerror("Ошибка", f"Ошибка при чтении Excel файла: {e}")
                    return
            else:
                messagebox.showerror("Ошибка", "Неподдерживаемый формат файла")
                return

            # Сбрасываем флаг использования первой строки как заголовков
            self.use_first_row_as_header = False

            # Открываем окно переименования столбцов
            def update_column_mapping(mapping):
                self.column_mapping = mapping
                if mapping:
                    messagebox.showinfo("Информация", f"Выбрано {len(mapping)} столбцов для переименования")
                else:
                    messagebox.showinfo("Информация", "Переименование столбцов не применяется")

            RenameColumnsWindow(self.root, columns, update_column_mapping, self.set_use_first_row_as_header)

        except Exception as e:
            messagebox.showerror("Ошибка", f"Не удалось загрузить информацию о столбцах: {e}")

    def _get_csv_columns(self, file_path):
        """
        Получает названия столбцов из CSV файла, определяя подходящие параметры
        """
        try:
            import chardet
            chardet_available = True
        except ImportError:
            chardet_available = False

        try:
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
                    return list(df_sample.columns)
                except:
                    return None  # Не удалось определить

            # Загружаем с найденными параметрами
            try:
                try:
                    # Попробуем сначала с on_bad_lines='skip' (новые версии pandas)
                    df = pd.read_csv(file_path, encoding=best_params[0], sep=best_params[1], nrows=0, on_bad_lines='skip')
                except TypeError:
                    # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                    df = pd.read_csv(file_path, encoding=best_params[0], sep=best_params[1], nrows=0, error_bad_lines=False, warn_bad_lines=True)
                return list(df.columns)
            except UnicodeDecodeError:
                # Если основная попытка не удалась из-за ошибки кодировки, пробуем CP1251 как резервный вариант
                try:
                    try:
                        # Попробуем сначала с on_bad_lines='skip' (новые версии pandas)
                        df = pd.read_csv(file_path, encoding='cp1251', sep=best_params[1], nrows=0, on_bad_lines='skip')
                    except TypeError:
                        # Если не поддерживается on_bad_lines, используем error_bad_lines (старые версии pandas)
                        df = pd.read_csv(file_path, encoding='cp1251', sep=best_params[1], nrows=0, error_bad_lines=False, warn_bad_lines=True)
                    return list(df.columns)
                except:
                    return None  # Не удалось загрузить с найденными параметрами
            except:
                return None  # Не удалось загрузить с найденными параметрами
        except Exception as e:
            # Если произошла ошибка на уровне файла или chardet, возвращаем None
            print(f"Ошибка в _get_csv_columns: {e}")
            return None

    def _detect_encoding_and_separator(self, file_path):
        """
        Определяет кодировку и разделитель для файла
        """
        try:
            import chardet
            chardet_available = True
        except ImportError:
            chardet_available = False

        try:
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

            # Если не нашли подходящую комбинацию, используем первую доступную кодировку
            if best_params[0] is None:
                return encodings[0], ','  # возвращаем первую кодировку и запятую по умолчанию

            return best_params[0], best_params[1]
        except Exception as e:
            # Если произошла ошибка, возвращаем utf-8 и запятую по умолчанию
            print(f"Ошибка в _detect_encoding_and_separator: {e}")
            return 'utf-8', ','

    def set_use_first_row_as_header(self, value):
        """Устанавливает флаг использования первой строки как заголовков"""
        self.use_first_row_as_header = value

    def select_txt_delimiter(self):
        """Открывает диалог для выбора разделителя для TXT файлов"""
        input_file = self.input_file_path.get()
        if not input_file or not input_file.lower().endswith('.txt'):
            messagebox.showwarning("Предупреждение", "Сначала выберите TXT файл")
            return

        # Создаем новое окно для выбора разделителя
        delimiter_window = tk.Toplevel(self.root)
        delimiter_window.title("Выбор разделителя для TXT файла")
        delimiter_window.geometry("700x500")
        delimiter_window.resizable(True, True)

        # Центрируем окно
        delimiter_window.transient(self.root)
        delimiter_window.grab_set()

        main_frame = ttk.Frame(delimiter_window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        ttk.Label(main_frame, text="Выберите разделитель для TXT файла:").pack(anchor=tk.W, pady=(0, 10))

        # Создаем фрейм для разделителя и превью
        content_frame = ttk.Frame(main_frame)
        content_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Левая часть - выбор разделителя
        left_frame = ttk.Frame(content_frame)
        left_frame.pack(side=tk.LEFT, fill=tk.Y, padx=(0, 10))

        # Переменная для хранения выбранного разделителя
        selected_delimiter = tk.StringVar(value=self.txt_delimiter.get())

        # Стандартные разделители
        delimiters = [
            (',', 'Запятая (,)'),
            (';', 'Точка с запятой (;)'),
            (':', 'Двоеточие (:)'),
            ('\t', 'Табуляция (\\t)'),
            ('|', 'Вертикальная черта (|)'),
            (' ', 'Пробел'),
            ('  ', 'Два пробела'),
            ('@@@', 'Другой...')
        ]

        # Создаем радиокнопки для стандартных разделителей
        for delimiter, label in delimiters[:-1]:  # Все кроме последнего
            ttk.Radiobutton(left_frame, text=label, variable=selected_delimiter, value=delimiter).pack(anchor=tk.W, pady=2)

        # Для "Другой..." создаем отдельную секцию
        other_frame = ttk.Frame(left_frame)
        other_frame.pack(anchor=tk.W, pady=(10, 0))

        ttk.Radiobutton(other_frame, text=delimiters[-1][1], variable=selected_delimiter, value='custom').pack(side=tk.LEFT)

        # Поле ввода для пользовательского разделителя
        custom_delimiter_var = tk.StringVar()
        custom_entry = ttk.Entry(other_frame, textvariable=custom_delimiter_var, width=10)
        custom_entry.pack(side=tk.LEFT, padx=(10, 0))

        # Если текущий разделитель не стандартный, устанавливаем его как пользовательский
        if self.txt_delimiter.get() not in [d[0] for d in delimiters[:-1]]:
            selected_delimiter.set('custom')
            custom_delimiter_var.set(self.txt_delimiter.get())

        # Правая часть - превью файла
        right_frame = ttk.Frame(content_frame)
        right_frame.pack(side=tk.RIGHT, fill=tk.BOTH, expand=True)

        ttk.Label(right_frame, text="Превью первых 20 строк файла:").pack(anchor=tk.W)

        # Создаем текстовое поле с прокруткой для превью
        preview_frame = ttk.Frame(right_frame)
        preview_frame.pack(fill=tk.BOTH, expand=True, pady=(5, 0))

        self.preview_text = tk.Text(preview_frame, wrap=tk.NONE, height=20, width=60)
        preview_scroll_y = ttk.Scrollbar(preview_frame, orient=tk.VERTICAL, command=self.preview_text.yview)
        preview_scroll_x = ttk.Scrollbar(preview_frame, orient=tk.HORIZONTAL, command=self.preview_text.xview)

        self.preview_text.configure(yscrollcommand=preview_scroll_y.set, xscrollcommand=preview_scroll_x.set)

        self.preview_text.grid(row=0, column=0, sticky="nsew")
        preview_scroll_y.grid(row=0, column=1, sticky="ns")
        preview_scroll_x.grid(row=1, column=0, sticky="ew")

        preview_frame.grid_rowconfigure(0, weight=1)
        preview_frame.grid_columnconfigure(0, weight=1)

        # Определяем функцию обновления превью до её использования
        def update_preview(delimiter):
            """Обновляет превью с учетом выбранного разделителя"""
            # Очищаем текущее содержимое
            self.preview_text.config(state=tk.NORMAL)
            self.preview_text.delete(1.0, tk.END)

            input_file = self.input_file_path.get()
            if not input_file:
                return

            try:
                # Определяем кодировку файла
                detected_encoding, _ = self._detect_encoding_and_separator(input_file)

                # Если разделитель - это специальный символ, преобразуем его
                actual_delimiter = delimiter
                if delimiter == '\\t':
                    actual_delimiter = '\t'

                # Читаем файл с учетом выбранного разделителя для демонстрации
                with open(input_file, 'r', encoding=detected_encoding) as f:
                    lines = []
                    for i, line in enumerate(f):
                        if i >= 20:  # Показываем только первые 20 строк
                            break
                        # Показываем, как будет выглядеть строка с выбранным разделителем
                        lines.append(line.rstrip('\n'))

                preview_content = '\n'.join(lines)
                self.preview_text.insert(tk.END, preview_content)
                self.preview_text.config(state=tk.DISABLED)  # Только для чтения
            except Exception as e:
                self.preview_text.insert(tk.END, f"Ошибка загрузки файла: {str(e)}")
                self.preview_text.config(state=tk.DISABLED)

        # Загружаем превью файла
        def load_file_preview():
            try:
                with open(input_file, 'r', encoding='utf-8') as f:
                    lines = []
                    for i, line in enumerate(f):
                        if i >= 20:  # Показываем только первые 20 строк
                            break
                        lines.append(line.rstrip('\n'))

                # Если не удалось открыть с UTF-8, пробуем другие кодировки
                if not lines:
                    encodings = ['cp1251', 'iso-8859-1', 'utf-8-sig']
                    for enc in encodings:
                        try:
                            with open(input_file, 'r', encoding=enc) as f:
                                lines = []
                                for i, line in enumerate(f):
                                    if i >= 20:  # Показываем только первые 20 строк
                                        break
                                    lines.append(line.rstrip('\n'))
                            break
                        except:
                            continue

                preview_content = '\n'.join(lines)
                self.preview_text.insert(tk.END, preview_content)
                self.preview_text.config(state=tk.DISABLED)  # Только для чтения
            except Exception as e:
                self.preview_text.insert(tk.END, f"Ошибка загрузки файла: {str(e)}")
                self.preview_text.config(state=tk.DISABLED)

        load_file_preview()

        # Функция для обработки выбора "другой"
        def on_select_change(*args):
            if selected_delimiter.get() != 'custom':
                custom_entry.config(state='disabled')
            else:
                custom_entry.config(state='enabled')
            # Обновляем превью при изменении разделителя
            update_preview(selected_delimiter.get())

        selected_delimiter.trace('w', on_select_change)
        on_select_change()  # Вызываем сразу для начального состояния

        # Кнопки OK и Отмена
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(pady=(0, 10))

        def continue_clicked():
            if selected_delimiter.get() == 'custom':
                delimiter = custom_delimiter_var.get()
                if not delimiter:
                    messagebox.showerror("Ошибка", "Введите разделитель")
                    return
            else:
                delimiter = selected_delimiter.get()

            # Если пользователь ввел экранированный символ, обрабатываем его
            if delimiter == '\\t':
                delimiter = '\t'
            elif delimiter == '\\n':
                delimiter = '\n'
            elif delimiter == '\\r':
                delimiter = '\r'

            self.txt_delimiter.set(delimiter)
            delimiter_window.destroy()
            # После выбора разделителя открываем окно переименования столбцов
            self.view_and_rename_columns()

        ttk.Button(button_frame, text="Продолжить", command=continue_clicked).pack(side=tk.LEFT, padx=(0, 10))
        ttk.Button(button_frame, text="Отмена", command=delimiter_window.destroy).pack(side=tk.LEFT)

    def update_processor_settings(self):
        """Обновляет настройки процессора данных"""
        self.processor.settings = self.settings


def main():
    root = tk.Tk()
    app = ExcelToDbApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()