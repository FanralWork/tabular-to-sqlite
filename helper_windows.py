
import tkinter as tk
from tkinter import ttk, scrolledtext


class RenameColumnsWindow:
    def __init__(self, parent, columns, rename_callback, use_first_row_as_header_callback=None):
        self.window = tk.Toplevel(parent)
        self.window.title("Переименовать столбцы")
        self.window.geometry("700x500")
        self.window.resizable(True, True)

        self.columns = columns
        self.rename_callback = rename_callback
        self.use_first_row_as_header_callback = use_first_row_as_header_callback

        # Создание виджетов
        self.create_widgets()

    def create_widgets(self):
        # Основной фрейм
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        title_label = ttk.Label(main_frame, text="Переименование столбцов", font=("Arial", 14))
        title_label.pack(pady=(0, 10))

        # Чекбокс для использования первой строки как заголовков
        if self.use_first_row_as_header_callback is not None:
            self.use_first_row_var = tk.BooleanVar(value=False)
            header_checkbox = ttk.Checkbutton(
                main_frame,
                text="Использовать первую строку как заголовки столбцов",
                variable=self.use_first_row_var,
                command=self.on_header_checkbox_change
            )
            header_checkbox.pack(anchor=tk.W, pady=(0, 10))

        # Создание фрейма для таблицы
        self.table_frame = ttk.Frame(main_frame)
        self.table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

        # Создание Treeview для отображения и редактирования столбцов
        columns_headers = ('№', 'Текущее имя', 'Новое имя')
        self.tree = ttk.Treeview(self.table_frame, columns=columns_headers, show='headings', height=15)

        # Настройка заголовков
        self.tree.heading('№', text='№')
        self.tree.heading('Текущее имя', text='Текущее имя')
        self.tree.heading('Новое имя', text='Новое имя')

        # Настройка ширины колонок
        self.tree.column('№', width=50, anchor=tk.CENTER)
        self.tree.column('Текущее имя', width=250, anchor=tk.W)
        self.tree.column('Новое имя', width=250, anchor=tk.W)

        # Добавление данных
        for i, col in enumerate(self.columns):
            self.tree.insert('', tk.END, values=(i+1, col, col))

        # Добавление скроллбара
        scrollbar = ttk.Scrollbar(self.table_frame, orient=tk.VERTICAL, command=self.tree.yview)
        self.tree.configure(yscrollcommand=scrollbar.set)

        # Упаковка дерева и скроллбара
        self.tree.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        # Кнопки управления
        button_frame = ttk.Frame(main_frame)
        button_frame.pack(fill=tk.X)

        reset_button = ttk.Button(button_frame, text="Сбросить все", command=self.reset_all)
        reset_button.pack(side=tk.LEFT, padx=(0, 10))

        apply_button = ttk.Button(button_frame, text="Применить", command=self.apply_changes)
        apply_button.pack(side=tk.RIGHT)

        cancel_button = ttk.Button(button_frame, text="Отмена", command=self.window.destroy)
        cancel_button.pack(side=tk.RIGHT, padx=(0, 10))

        # Добавляем возможность редактирования полей "Новое имя"
        self.tree.bind('<Double-1>', self.on_double_click)

    def on_header_checkbox_change(self):
        """Обработка изменения состояния чекбокса использования первой строки как заголовков"""
        if self.use_first_row_var.get():
            # Если чекбокс отмечен, скрываем таблицу переименования
            self.table_frame.pack_forget()
        else:
            # Если чекбокс снят, показываем таблицу переименования
            self.table_frame.pack(fill=tk.BOTH, expand=True, pady=(0, 10))

    def apply_changes(self):
        """Применение изменений и вызов callback"""
        if self.use_first_row_as_header_callback is not None and self.use_first_row_var.get():
            # Если отмечен чекбокс использования первой строки как заголовков
            self.use_first_row_as_header_callback(True)
            self.window.destroy()
        else:
            # Иначе собираем переименования из таблицы
            renamed_columns = {}
            print("Начинаем анализ изменений столбцов...")

            for i, item in enumerate(self.tree.get_children()):
                values = self.tree.item(item, "values")
                old_name = values[1]
                new_name = values[2]

                # Если old_name содержит образец (например, "Column_0 (образец: значение)"),
                # извлекаем только основное имя столбца
                if " (образец:" in old_name:
                    actual_old_name = old_name.split(" (образец:")[0]
                else:
                    actual_old_name = old_name

                print(f"Столбец {i+1}: '{old_name}' -> '{new_name}', изменен: {old_name != new_name}, actual: '{actual_old_name}'")

                if old_name != new_name:
                    renamed_columns[actual_old_name] = new_name

            # Отладочный вывод
            print(f"Переименованные столбцы: {renamed_columns}")

            self.rename_callback(renamed_columns)
            self.window.destroy()

    def on_double_click(self, event):
        """Обработка двойного клика для редактирования поля 'Новое имя'"""
        region = self.tree.identify("region", event.x, event.y)
        column = self.tree.identify_column(event.x)

        # Проверяем, что клик был по колонке "Новое имя" (index 3, т.к. начинается с 1)
        if region == "cell" and column == "#3":
            item = self.tree.selection()[0]
            current_value = self.tree.item(item, "values")[2]  # Текущее значение нового имени

            # Создаем временное текстовое поле для редактирования
            bbox = self.tree.bbox(item, column)
            if bbox:
                x, y, width, height = bbox
                entry = ttk.Entry(self.tree)
                entry.place(x=x, y=y, width=width, height=height)
                entry.insert(0, current_value)
                entry.select_range(0, tk.END)  # Выделяем весь текст для удобства
                entry.focus()

                # Функция для сохранения изменений
                def save_edit(event=None):
                    # Обновляем значение в дереве
                    new_value = entry.get()
                    if new_value.strip():  # Проверяем, что имя не пустое
                        # Получаем старые значения
                        old_values = self.tree.item(item, "values")
                        # Обновляем только новое имя
                        new_values = (old_values[0], old_values[1], new_value)
                        self.tree.item(item, values=new_values)

                    # Уничтожаем Entry после небольшой задержки, чтобы избежать проблем с событиями
                    self.tree.after(1, entry.destroy)

                # Привязываем события для сохранения
                entry.bind("<Return>", save_edit)
                entry.bind("<FocusOut>", save_edit)
                entry.bind("<Escape>", lambda e: entry.destroy())

    def reset_all(self):
        """Сброс всех имен к исходным"""
        for i, item in enumerate(self.tree.get_children()):
            original_name = self.tree.item(item, "values")[1]  # Берем текущее имя
            new_values = (i+1, original_name, original_name)
            self.tree.item(item, values=new_values)


class SettingsWindow:
    def __init__(self, parent, settings_callback, is_optimized=False):
        self.window = tk.Toplevel(parent)
        self.window.title("Настройки")
        self.window.geometry("400x300")
        self.window.resizable(False, False)

        self.settings_callback = settings_callback
        self.is_optimized = is_optimized

        # Переменные для настроек
        self.remove_empty_cols_var = tk.BooleanVar(value=True)
        self.overwrite_existing_var = tk.BooleanVar(value=True)
        self.chunk_size_var = tk.IntVar(value=10000)  # Размер чанка по умолчанию

        # Создание виджетов
        self.create_widgets()

    def create_widgets(self):
        # Основной фрейм
        main_frame = ttk.Frame(self.window, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Заголовок
        title_label = ttk.Label(main_frame, text="Настройки конвертации", font=("Arial", 14))
        title_label.grid(row=0, column=0, columnspan=2, pady=(0, 20))

        # Чекбокс для удаления пустых столбцов
        remove_empty_check = ttk.Checkbutton(
            main_frame,
            text="Удалять пустые столбцы",
            variable=self.remove_empty_cols_var
        )
        remove_empty_check.grid(row=1, column=0, sticky=tk.W, pady=5)

        # Чекбокс для перезаписи существующих файлов
        overwrite_check = ttk.Checkbutton(
            main_frame,
            text="Перезаписывать существующие файлы",
            variable=self.overwrite_existing_var
        )
        overwrite_check.grid(row=2, column=0, sticky=tk.W, pady=5)

        # Поле для настройки размера чанка (только для оптимизированной версии)
        if self.is_optimized:
            chunk_size_frame = ttk.Frame(main_frame)
            chunk_size_frame.grid(row=3, column=0, columnspan=2, sticky=tk.W, pady=5)

            ttk.Label(chunk_size_frame, text="Размер чанка:").pack(side=tk.LEFT)
            chunk_size_spinbox = ttk.Spinbox(
                chunk_size_frame,
                from_=100,
                to=1000000,
                increment=100,
                textvariable=self.chunk_size_var,
                width=10
            )
            chunk_size_spinbox.pack(side=tk.LEFT, padx=(10, 0))

            ttk.Label(chunk_size_frame, text="(для файлов > 100 МБ)").pack(side=tk.LEFT, padx=(10, 0))

        # Выбор кодировки файла (только для оптимизированной версии)
        if self.is_optimized:
            encoding_frame = ttk.Frame(main_frame)
            encoding_frame.grid(row=4 if self.is_optimized else 3, column=0, columnspan=2, sticky=tk.W, pady=5)

            ttk.Label(encoding_frame, text="Кодировка файла:").pack(side=tk.LEFT)

            # Доступные кодировки
            encodings = ['utf-8', 'utf-8-sig', 'cp1251', 'iso-8859-1', 'latin-1']
            self.encoding_var = tk.StringVar(value=self.settings.get('file_encoding', 'utf-8'))

            encoding_combo = ttk.Combobox(
                encoding_frame,
                textvariable=self.encoding_var,
                values=encodings,
                state="readonly",
                width=15
            )
            encoding_combo.pack(side=tk.LEFT, padx=(10, 0))

        # Выбор режима обработки ошибок декодирования (только для оптимизированной версии)
        if self.is_optimized:
            error_handling_frame = ttk.Frame(main_frame)
            error_handling_frame.grid(row=5 if self.is_optimized else 4, column=0, columnspan=2, sticky=tk.W, pady=5)

            ttk.Label(error_handling_frame, text="Обработка ошибок декодирования:").pack(side=tk.LEFT)

            # Доступные режимы обработки ошибок
            error_modes = ['strict', 'ignore', 'replace']
            self.error_mode_var = tk.StringVar(value=self.settings.get('error_handling_mode', 'replace'))

            error_mode_combo = ttk.Combobox(
                error_handling_frame,
                textvariable=self.error_mode_var,
                values=error_modes,
                state="readonly",
                width=15
            )
            error_mode_combo.pack(side=tk.LEFT, padx=(10, 0))

            # Добавляем пояснение
            tooltip = tk.Label(error_handling_frame, text="(strict: ошибка при проблемных символах, ignore: пропустить, replace: заменить на ?)", font=("Arial", 8))
            tooltip.pack(side=tk.LEFT, padx=(5, 0))

        # Кнопки OK и Cancel
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=5 if self.is_optimized else 3, column=0, columnspan=2, pady=20)

        ok_button = ttk.Button(button_frame, text="OK", command=self.save_settings)
        ok_button.pack(side=tk.LEFT, padx=(0, 10))

        cancel_button = ttk.Button(button_frame, text="Отмена", command=self.window.destroy)
        cancel_button.pack(side=tk.LEFT)

    def save_settings(self):
        # Сохранение настроек через callback
        settings = {
            'remove_empty_cols': self.remove_empty_cols_var.get(),
            'overwrite_existing': self.overwrite_existing_var.get()
        }

        # Добавляем размер чанка, если используется оптимизированная версия
        if self.is_optimized:
            settings['chunk_size'] = self.chunk_size_var.get()
            # Добавляем выбранную кодировку
            settings['file_encoding'] = self.encoding_var.get()
            # Добавляем режим обработки ошибок
            settings['error_handling_mode'] = self.error_mode_var.get()

        self.settings_callback(settings)
        self.window.destroy()


class ProgressDialog:
    def __init__(self, parent, title="Выполняется обработка", max_value=100):
        self.window = tk.Toplevel(parent)
        self.window.title(title)
        self.window.geometry("400x100")
        self.window.resizable(False, False)

        # Центрируем окно относительно родительского
        self.window.transient(parent)
        self.window.grab_set()

        # Создаем виджеты
        self.create_widgets(max_value)

        # Закрытие окна по крестику не должно прерывать процесс
        self.window.protocol("WM_DELETE_WINDOW", self.on_close)

    def create_widgets(self, max_value):
        main_frame = ttk.Frame(self.window, padding="20")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Метка с информацией
        self.label = ttk.Label(main_frame, text="Обработка данных...", font=("Arial", 10))
        self.label.pack(pady=(0, 10))

        # Прогресс-бар
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            main_frame,
            variable=self.progress_var,
            maximum=max_value,
            length=300
        )
        self.progress_bar.pack(pady=(0, 10))

        # Кнопка отмены (временно отключена, так как отмена требует дополнительной логики)
        self.cancel_button = ttk.Button(
            main_frame,
            text="Отмена",
            command=self.on_cancel,
            state=tk.DISABLED  # Пока отключаем кнопку отмены
        )
        self.cancel_button.pack()

        self.current_value = 0
        self.max_value = max_value

    def update_progress(self, value, text=None):
        """Обновляет прогресс-бар и текст"""
        self.current_value = value
        self.progress_var.set(value)

        if text:
            self.label.config(text=text)

        # Обновляем окно
        self.window.update()

    def increment(self, step=1, text=None):
        """Увеличивает прогресс на заданное значение"""
        new_value = min(self.current_value + step, self.max_value)
        self.update_progress(new_value, text)

    def on_cancel(self):
        """Обработка нажатия кнопки отмены"""
        # Пока просто игнорируем, так как отмена требует дополнительной логики в основном процессе
        pass

    def on_close(self):
        """Обработка закрытия окна"""
        # Игнорируем закрытие окна, чтобы не прерывать процесс
        pass

    def close(self):
        """Закрывает окно прогресса"""
        try:
            self.window.destroy()
        except tk.TclError:
            # Окно уже закрыто
            pass