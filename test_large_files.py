"""
Тестирование обработки больших файлов с использованием оптимизированного процессора
"""
import os
import sys
import pandas as pd
from data_processor_optimized import DataProcessorOptimized

def create_large_test_file(filename, num_rows=100000, num_cols=10):
    """
    Создает тестовый CSV файл большого размера для тестирования
    """
    print(f"Создание тестового файла {filename} с {num_rows} строками и {num_cols} столбцами...")
    
    # Создаем данные для теста
    data = {}
    for i in range(num_cols):
        col_name = f"Column_{i}"
        # Создаем данные с разными типами для лучшего тестирования
        if i % 3 == 0:
            data[col_name] = [f"value_{j}_{i}" for j in range(num_rows)]
        elif i % 3 == 1:
            data[col_name] = [float(j * i) for j in range(num_rows)]
        else:
            data[col_name] = [j + i for j in range(num_rows)]
    
    df = pd.DataFrame(data)
    df.to_csv(filename, index=False)
    print(f"Файл {filename} создан. Размер: {os.path.getsize(filename) / (1024*1024):.2f} MB")


def test_large_file_processing():
    """
    Тестирует обработку большого файла с использованием оптимизированного процессора
    """
    # Создаем тестовые файлы
    test_csv_file = "large_test.csv"
    test_db_file = "large_test.db"
    
    # Создаем большой тестовый файл
    create_large_test_file(test_csv_file, num_rows=50000, num_cols=20)
    
    # Создаем процессор с настройками для обработки больших файлов
    processor = DataProcessorOptimized({
        'remove_empty_cols': True,
        'overwrite_existing': True,
        'chunk_size': 5000  # Небольшой размер чанка для тестирования
    })
    
    print("\nНачинаем конвертацию большого файла...")
    
    try:
        # Конвертируем файл в базу данных
        result = processor.convert_file_to_db(
            input_file_path=test_csv_file,
            db_file_path=test_db_file
        )
        
        print("\nКонвертация завершена успешно!")
        print(result)
        
        # Проверяем, что файл базы данных создан
        if os.path.exists(test_db_file):
            print(f"\nФайл базы данных создан: {test_db_file}")
            print(f"Размер базы данных: {os.path.getsize(test_db_file) / (1024*1024):.2f} MB")
        else:
            print("Файл базы данных не был создан")
        
    except Exception as e:
        print(f"Ошибка при конвертации: {e}")
        import traceback
        traceback.print_exc()
    
    finally:
        # Удаляем тестовые файлы
        if os.path.exists(test_csv_file):
            os.remove(test_csv_file)
            print(f"\nТестовый файл {test_csv_file} удален")
        
        if os.path.exists(test_db_file):
            os.remove(test_db_file)
            print(f"Тестовый файл базы данных {test_db_file} удален")


def test_different_chunk_sizes():
    """
    Тестирует влияние разных размеров чанков на производительность
    """
    print("\nТестирование разных размеров чанков...")
    
    # Создаем тестовый файл
    test_csv_file = "chunk_test.csv"
    test_db_file = "chunk_test.db"
    
    create_large_test_file(test_csv_file, num_rows=20000, num_cols=10)
    
    chunk_sizes = [1000, 5000, 10000, 20000]
    
    for chunk_size in chunk_sizes:
        print(f"\nТестирование с размером чанка: {chunk_size}")
        
        # Создаем процессор с конкретным размером чанка
        processor = DataProcessorOptimized({
            'remove_empty_cols': True,
            'overwrite_existing': True,
            'chunk_size': chunk_size
        })
        
        try:
            result = processor.convert_file_to_db(
                input_file_path=test_csv_file,
                db_file_path=f"{test_db_file.replace('.db', f'_{chunk_size}.db')}"
            )
            
            print(f"Чанк размером {chunk_size}: Успешно")
            
        except Exception as e:
            print(f"Чанк размером {chunk_size}: Ошибка - {e}")
    
    # Удаляем тестовые файлы
    if os.path.exists(test_csv_file):
        os.remove(test_csv_file)
        print(f"\nТестовый файл {test_csv_file} удален")
    
    for chunk_size in chunk_sizes:
        db_file = f"{test_db_file.replace('.db', f'_{chunk_size}.db')}"
        if os.path.exists(db_file):
            os.remove(db_file)
            print(f"Тестовый файл базы данных {db_file} удален")


if __name__ == "__main__":
    print("Тестирование обработки больших файлов...")
    
    # Проверяем, установлен ли tqdm
    try:
        import tqdm
        print("Библиотека tqdm установлена")
    except ImportError:
        print("Предупреждение: библиотека tqdm не установлена. Установите её командой: pip install tqdm")
        sys.exit(1)
    
    # Запускаем тесты
    test_large_file_processing()
    test_different_chunk_sizes()
    
    print("\nТестирование завершено!")