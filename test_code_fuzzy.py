#!/usr/bin/env python3
# -*- coding: utf-8 -*-

"""
Тест функции нечеткого поиска по строкам с реальными данными

ВНИМАНИЕ: Для ускорения тестирования количество обрабатываемых строк ограничено до 20
вместо полного файла. Это позволяет быстро протестировать функциональность
без ожидания обработки сотен строк.
"""

import sys
import os
import pandas as pd
import difflib

# Добавляем путь к основному модулю
sys.path.append(os.path.dirname(os.path.abspath(__file__)))

from main import MiStockSyncApp, TRSH

# Создаем экземпляр класса (без GUI)
app = MiStockSyncApp(None)

# Устанавливаем конфигурацию для Вити (нужно для правильного определения колонок)
app.current_config = "vitya"

# Ограничиваем количество строк для тестирования
# Для ускорения тестирования установлено ограничение в 20 строк
# Если нужно протестировать больше строк, измените это значение
MAX_TEST_ROWS = 20  # Максимальное количество строк для тестирования

print("🔍 Тестирование функций нечеткого поиска с реальными данными")
print("=" * 80)
print(f"📊 Текущий порог схожести (TRSH): {TRSH:.1%}")
print(f"📊 Ограничение тестирования: максимум {MAX_TEST_ROWS} строк")
print("=" * 80)

# Загружаем реальные данные через excel_loader для правильного переименования колонок
try:
    print("📁 Загружаем реальные данные через excel_loader...")

    # База данных
    base_file = "data/base_ОСНОВА 30.07.2025.xlsx"
    if os.path.exists(base_file):
        print(f"✅ Загружаем базу: {base_file}")
        try:
            # Используем excel_loader для правильного переименования колонок
            from excel_loader.loader import load_with_config

            base_df = load_with_config(base_file, "base")
            print(f"   📊 База содержит {len(base_df)} строк")
            print(f"   📋 Колонки: {list(base_df.columns)}")
        except Exception as e:
            print(f"   ⚠️ Ошибка загрузки через excel_loader: {e}")
            print("   🔄 Пробуем загрузить напрямую...")
            base_df = pd.read_excel(base_file)
            print(f"   📊 База содержит {len(base_df)} строк")
            print(f"   📋 Колонки: {list(base_df.columns)}")
    else:
        print(f"❌ Файл базы не найден: {base_file}")
        base_df = None

    # Данные поставщика Вити
    vitya_file = "data/JHT_Xiaomi_аксессуары31.07xlsx.xlsx"
    if os.path.exists(vitya_file):
        print(f"✅ Загружаем данные Вити: {vitya_file}")
        try:
            # Используем excel_loader для правильного переименования колонок
            vitya_df = load_with_config(vitya_file, "vitya")
            print(f"   📊 Витя содержит {len(vitya_df)} строк")
            print(f"   📋 Колонки: {list(vitya_df.columns)}")

            # Ограничиваем количество строк для тестирования
            if len(vitya_df) > MAX_TEST_ROWS:
                print(f"   🔄 Ограничиваем до {MAX_TEST_ROWS} строк для тестирования")
                vitya_df = vitya_df.head(MAX_TEST_ROWS)
                print(f"   📊 После ограничения: {len(vitya_df)} строк")
        except Exception as e:
            print(f"   ⚠️ Ошибка загрузки через excel_loader: {e}")
            print("   🔄 Пробуем загрузить напрямую...")
            vitya_df = pd.read_excel(vitya_file)
            print(f"   📊 Витя содержит {len(vitya_df)} строк")
            print(f"   📋 Колонки: {list(vitya_df.columns)}")

            # Ограничиваем количество строк для тестирования
            if len(vitya_df) > MAX_TEST_ROWS:
                print(f"   🔄 Ограничиваем до {MAX_TEST_ROWS} строк для тестирования")
                vitya_df = vitya_df.head(MAX_TEST_ROWS)
                print(f"   📊 После ограничения: {len(vitya_df)} строк")
    else:
        print(f"❌ Файл Вити не найден: {vitya_file}")
        vitya_df = None

    # Данные поставщика Димы
    dimi_file = "data/DiMi_Opt_Price.xlsx_31-07.xlsx"
    if os.path.exists(dimi_file):
        print(f"✅ Загружаем данные Димы: {dimi_file}")
        try:
            # Используем excel_loader для правильного переименования колонок
            dimi_df = load_with_config(dimi_file, "dimi")
            print(f"   📊 Дима содержит {len(dimi_df)} строк")
            print(f"   📋 Колонки: {list(dimi_df.columns)}")

            # Ограничиваем количество строк для тестирования
            if len(dimi_df) > MAX_TEST_ROWS:
                print(f"   🔄 Ограничиваем до {MAX_TEST_ROWS} строк для тестирования")
                dimi_df = dimi_df.head(MAX_TEST_ROWS)
                print(f"   📊 После ограничения: {len(dimi_df)} строк")
        except Exception as e:
            print(f"   ⚠️ Ошибка загрузки через excel_loader: {e}")
            print("   🔄 Пробуем загрузить напрямую...")
            dimi_df = pd.read_excel(dimi_file)
            print(f"   📊 Дима содержит {len(dimi_df)} строк")
            print(f"   📋 Колонки: {list(dimi_df.columns)}")

            # Ограничиваем количество строк для тестирования
            if len(dimi_df) > MAX_TEST_ROWS:
                print(f"   🔄 Ограничиваем до {MAX_TEST_ROWS} строк для тестирования")
                dimi_df = dimi_df.head(MAX_TEST_ROWS)
                print(f"   📊 После ограничения: {len(dimi_df)} строк")
    else:
        print(f"❌ Файл Дими не найден: {dimi_file}")
        dimi_df = None

except Exception as e:
    print(f"❌ Ошибка загрузки данных: {e}")
    base_df = None
    vitya_df = None
    dimi_df = None

print("\n" + "=" * 80)

# Тестируем функцию compare_by_fuzzy_string_matching
if base_df is not None and vitya_df is not None:
    print("🔍 Тестирование функции compare_by_fuzzy_string_matching")
    print("-" * 50)

    # Предобработка данных Вити (если нужно)
    try:
        # Применяем предобработку как в основном приложении
        processed_vitya = app.preprocess_vitya_fixed_v3(vitya_df.copy())
        print(f"✅ Предобработка Вити завершена, осталось {len(processed_vitya)} строк")

        # Вызываем функцию нечеткого поиска
        result = app.compare_by_fuzzy_string_matching(processed_vitya, base_df, "vitya")

        # Выводим результаты
        print(f"✅ Функция вернула список с {len(result)} элементами")

        if len(result) > 0:
            print(f"\n📋 Примеры найденных совпадений:")
            for i, match in enumerate(result[:5]):
                print(
                    f"  {i+1}. '{match['supplier_name'][:50]}...' -> '{match['base_name'][:50]}...' "
                    f"(схожесть: {match['similarity_ratio']:.2%})"
                )
        else:
            print("⚠️ Функция не нашла совпадений")

    except Exception as e:
        print(f"❌ Ошибка при тестировании compare_by_fuzzy_string_matching: {e}")

print("\n" + "=" * 80)

# Тестируем новую функцию find_item_by_fuzzy_matching
if base_df is not None:
    print("🔍 Тестирование новой функции find_item_by_fuzzy_matching")
    print("-" * 50)

    # Устанавливаем базу данных в приложении
    app.base_df = base_df

    # Тестовые названия товаров
    test_names = [
        'Монитор Xiaomi Redmi Display 27" G PRO 27Q  180Hz (P27QDA-RGP)',
        "Смарт-часы Xiaomi Redmi Watch 5  (M2462W1) EU",
        "USB Flash накопитель Xiaomi U-Disk Thumb Drive 64 Гб (XMUP21YM)",
        "Микроволновая печь Xiaomi Mijia Microwave Oven (MWB020) 20L",
        "Пароварка Xiaomi Mijia Multifunctional Electric Steamer S1 (MES03) 13L",
        "Автоматическая машина для приготовления соевого молока Mijia (MJDJJ01DEM) 1L",
        "Массажный пистолет Mijia Fascia Gun 3 Mini   (MJJMQ07YM)",
        "Мультистиллер SenCiciMen X9 Pro EU   New!!!(С плёнка!!)",
        "Умный блендер с функцией нагреваXiaomi Miiia Smart Sound Blender S2 (MJPBJ02DEM) 1.5L",
    ]

    print(f"🧪 Тестируем {len(test_names)} названий товаров...")

    for i, test_name in enumerate(test_names, 1):
        try:
            found_name, row_number, color, price = app.find_item_by_fuzzy_matching(
                test_name
            )

            # Вычисляем процент схожести
            if found_name != "Не найдено":
                similarity = difflib.SequenceMatcher(
                    None, test_name.lower(), found_name.lower()
                ).ratio()
                similarity_percent = f"{similarity:.2%}"
            else:
                similarity_percent = "N/A"

            print(
                f"  {i}. '{test_name[:40]}...' -> '{found_name[:40]}...' "
                f"(строка: {row_number}, цвет: {color}, цена: {price}, схожесть: {similarity_percent})"
            )
        except Exception as e:
            print(f"  {i}. ❌ Ошибка: {e}")


print("✅ Тест завершен")
