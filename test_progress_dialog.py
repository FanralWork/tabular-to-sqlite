"""
Тестирование нового окна прогресса в GUI приложении
"""
import tkinter as tk
from tkinter import ttk
import tempfile
import os
import pandas as pd
from gui_app_optimized import ExcelToDbApp

def create_test_file():
    """Создает тестовый CSV файл для проверки"""
    # Создаем временный файл
    temp_dir = tempfile.gettempdir()
    csv_path = os.path.join(temp_dir, "test_large_file.csv")
    db_path = os.path.join(temp_dir, "test_output.db")
    
    # Создаем тестовые данные
    data = {
        'Name': [f'Name_{i}' for i in range(1000)],
        'Age': [20 + (i % 50) for i in range(1000)],
        'City': [f'City_{i % 10}' for i in range(1000)]
    }
    
    df = pd.DataFrame(data)
    df.to_csv(csv_path, index=False)
    
    return csv_path, db_path

def test_progress_dialog():
    """Тестирует отображение прогресса в новом окне"""
    print("Создание тестового файла...")
    csv_path, db_path = create_test_file()
    
    print(f"Тестовый файл: {csv_path}")
    print(f"Выходной файл: {db_path}")
    
    # Создаем приложение
    root = tk.Tk()
    app = ExcelToDbApp(root)
    
    # Устанавливаем пути к файлам
    app.input_file_path.set(csv_path)
    app.output_file_path.set(db_path)
    
    # Устанавливаем небольшой размер чанка для демонстрации прогресса
    app.processor.settings['chunk_size'] = 100
    
    print("Запуск приложения...")
    print("Проверьте, появляется ли окно прогресса при нажатии кнопки 'Конвертировать'")
    
    # Запускаем приложение
    root.mainloop()
    
    # Удаляем временные файлы
    try:
        os.remove(csv_path)
        if os.path.exists(db_path):
            os.remove(db_path)
    except:
        pass

if __name__ == "__main__":
    test_progress_dialog()