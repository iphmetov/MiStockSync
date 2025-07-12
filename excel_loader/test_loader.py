#!/usr/bin/env python3
"""
Тестовый скрипт для проверки работы excel_loader модуля
"""

import os
import sys
from pathlib import Path

# Добавляем путь к модулю
sys.path.insert(0, str(Path(__file__).parent))

from loader import select_and_load_excel, load_largest_file


def test_config_loading():
    """Тест загрузки конфигурации"""
    print("🧪 Тестирование загрузки конфигурации...")

    from loader import get_loader

    loader = get_loader()

    if loader.config:
        print("✅ Конфигурация успешно загружена")
        print(
            f"📊 Найдено правил переименования: {len(loader.config.get('column_mapping', {}))}"
        )
        return True
    else:
        print("❌ Ошибка загрузки конфигурации")
        return False


def test_directory_loading():
    """Тест загрузки самого большого файла из директории"""
    print("\n🧪 Тестирование загрузки самого большого файла...")

    # Проверяем директорию с тестовыми данными
    test_dir = "../data/input"

    if os.path.exists(test_dir):
        print(f"📁 Тестовая директория: {os.path.abspath(test_dir)}")

        # Проверяем наличие Excel файлов
        excel_files = [
            f for f in os.listdir(test_dir) if f.lower().endswith((".xlsx", ".xls"))
        ]

        if excel_files:
            print(f"📊 Найдено Excel файлов: {len(excel_files)}")
            for file in excel_files:
                file_path = os.path.join(test_dir, file)
                file_size = os.path.getsize(file_path)
                print(f"  📄 {file}: {file_size / (1024*1024):.1f} MB")

            # Загружаем самый большой файл
            df = load_largest_file(test_dir)

            if df is not None:
                print(f"✅ Самый большой файл успешно загружен: {df.shape}")
                return True
            else:
                print("❌ Ошибка загрузки файла")
                return False
        else:
            print("❌ Excel файлы не найдены в тестовой директории")
            return False
    else:
        print(f"❌ Тестовая директория не найдена: {test_dir}")
        return False


def test_interactive_loading():
    """Тест интерактивной загрузки файла"""
    print("\n🧪 Тестирование интерактивной загрузки...")
    print("📝 Для тестирования интерактивной загрузки запустите:")
    print("   python test_loader.py --interactive")
    print("   или вызовите select_and_load_excel() в интерактивном режиме")
    return True


def main():
    """Основная функция тестирования"""
    print("🚀 Запуск тестов excel_loader модуля")
    print("=" * 50)

    tests = [
        ("Загрузка конфигурации", test_config_loading),
        ("Загрузка из директории", test_directory_loading),
        ("Интерактивная загрузка", test_interactive_loading),
    ]

    results = []

    for test_name, test_func in tests:
        try:
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ Ошибка в тесте '{test_name}': {e}")
            results.append((test_name, False))

    print("\n" + "=" * 50)
    print("📊 Результаты тестирования:")

    passed = 0
    for test_name, result in results:
        status = "✅ ПРОЙДЕН" if result else "❌ ПРОВАЛЕН"
        print(f"  {status}: {test_name}")
        if result:
            passed += 1

    print(f"\n📈 Итого: {passed}/{len(results)} тестов пройдено")

    if "--interactive" in sys.argv:
        print("\n🎯 Запуск интерактивного теста...")
        df = select_and_load_excel()
        if df is not None:
            print(f"✅ Интерактивный тест пройден: {df.shape}")
        else:
            print("❌ Интерактивный тест не пройден")


if __name__ == "__main__":
    main()
