#!/usr/bin/env python3
"""
Тестовый скрипт для проверки системы множественных конфигураций
"""

import os
import sys
from pathlib import Path

# Добавляем путь к модулю
sys.path.insert(0, str(Path(__file__).parent))

from loader import (
    get_available_configs,
    select_and_load_excel,
    load_largest_file,
    load_with_config,
    get_loader,
)


def test_available_configs():
    """Тест получения списка доступных конфигов"""
    print("🧪 Тестирование списка доступных конфигов...")

    configs = get_available_configs()
    print(f"📊 Найдено конфигураций: {len(configs)}")

    for config in configs:
        loader = get_loader(config)
        supplier_name = loader.config.get("supplier_name", config)
        description = loader.config.get("description", "Нет описания")
        column_count = len(loader.config.get("column_mapping", {}))

        print(f"  📋 {config}: {supplier_name}")
        print(f"     📝 {description}")
        print(f"     🏷️ Столбцов в маппинге: {column_count}")
        print()

    return len(configs) > 0


def test_base_config():
    """Тест загрузки самого большого файла с base конфигом"""
    print("🧪 Тестирование base конфига для основной базы...")

    test_dir = "../data/input"

    if os.path.exists(test_dir):
        print(f"📁 Тестовая директория: {os.path.abspath(test_dir)}")

        # Загружаем с base конфигом
        print("🔄 Загружаем самый большой файл с base конфигом...")
        base_df = load_largest_file(test_dir, config_name="base")

        if base_df is not None:
            print(f"✅ База данных загружена: {base_df.shape}")
            print(f"🏷️ Столбцы после обработки:")
            for i, col in enumerate(base_df.columns):
                print(f"  {i+1:2d}. {col}")
            return True
        else:
            print("❌ Ошибка загрузки base_df")
            return False
    else:
        print(f"❌ Тестовая директория не найдена: {test_dir}")
        return False


def test_vitya_config():
    """Тест конфига для прайса Витя"""
    print("\n🧪 Тестирование vitya конфига...")

    # Проверяем что конфиг загружается
    loader = get_loader("vitya")

    print(f"📋 Конфиг поставщика: {loader.config.get('supplier_name')}")
    print(
        f"💱 Валюта: {loader.config.get('settings', {}).get('currency', 'Не указана')}"
    )

    # Показываем маппинг столбцов
    mapping = loader.config.get("column_mapping", {})
    print(f"🔄 Маппинг столбцов ({len(mapping)} правил):")
    for old_col, new_col in mapping.items():
        print(f"  '{old_col}' → '{new_col}'")

    # Показываем игнорируемые столбцы
    ignored = loader.config.get("ignore_columns", [])
    if ignored:
        print(f"🚫 Игнорируемые столбцы ({len(ignored)}):")
        for col in ignored:
            print(f"  - {col}")

    return True


def test_config_validation():
    """Тест валидации конфигов"""
    print("\n🧪 Тестирование валидации конфигов...")

    configs_to_test = ["base", "vitya", "dima"]

    for config_name in configs_to_test:
        try:
            loader = get_loader(config_name)
            validation = loader.config.get("validation", {})

            required_cols = validation.get("required_columns", [])
            price_range = (
                f"{validation.get('price_min', 0)} - {validation.get('price_max', '∞')}"
            )

            print(f"📋 {config_name} ({loader.config.get('supplier_name')}):")
            print(f"  ✅ Обязательные столбцы: {required_cols}")
            print(f"  💰 Диапазон цен: {price_range}")

        except Exception as e:
            print(f"❌ Ошибка в конфиге {config_name}: {e}")

    return True


def test_direct_file_loading():
    """Тест прямой загрузки файла с разными конфигами"""
    print("\n🧪 Тестирование прямой загрузки файла...")

    test_file = "../data/input/price_2.xlsx"

    if os.path.exists(test_file):
        print(f"📄 Тестовый файл: {os.path.basename(test_file)}")

        configs_to_test = ["default", "vitya"]

        for config_name in configs_to_test:
            print(f"\n🔄 Загружаем с конфигом '{config_name}'...")

            df = load_with_config(test_file, config_name)

            if df is not None:
                print(f"  ✅ Успешно: {df.shape}")
                print(f"  🏷️ Первые 5 столбцов: {list(df.columns[:5])}")
            else:
                print(f"  ❌ Ошибка загрузки с конфигом {config_name}")

        return True
    else:
        print(f"❌ Тестовый файл не найден: {test_file}")
        return False


def main():
    """Основная функция тестирования"""
    print("🚀 Тестирование системы множественных конфигураций")
    print("=" * 60)

    tests = [
        ("Список доступных конфигов", test_available_configs),
        ("Base конфиг (основная база)", test_base_config),
        ("Vitya конфиг", test_vitya_config),
        ("Валидация конфигов", test_config_validation),
        ("Прямая загрузка файлов", test_direct_file_loading),
    ]

    results = []

    for test_name, test_func in tests:
        try:
            print(f"\n{'='*20} {test_name} {'='*20}")
            result = test_func()
            results.append((test_name, result))
        except Exception as e:
            print(f"❌ Ошибка в тесте '{test_name}': {e}")
            results.append((test_name, False))

    print("\n" + "=" * 60)
    print("📊 Результаты тестирования:")

    passed = 0
    for test_name, result in results:
        status = "✅ ПРОЙДЕН" if result else "❌ ПРОВАЛЕН"
        print(f"  {status}: {test_name}")
        if result:
            passed += 1

    print(f"\n📈 Итого: {passed}/{len(results)} тестов пройдено")

    if passed == len(results):
        print("\n🎉 Все тесты пройдены! Система конфигов готова к использованию.")
        print("\n📋 Доступные функции:")
        print("- select_and_load_excel(config_name='vitya')  # Диалог с конфигом")
        print("- load_largest_file('./data', config_name='base')  # Самый большой файл")
        print("- load_with_config('file.xlsx', 'vitya')  # Прямая загрузка")
        print("- get_available_configs()  # Список конфигов")
    else:
        print(f"\n❌ {len(results) - passed} тестов не пройдены.")


if __name__ == "__main__":
    main()
