"""
MiStockSync - Приложение для синхронизации прайсов
Версия: 1.0.0
"""

import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
import sys
import os
import pandas as pd
from datetime import datetime
import logging

# Добавляем путь к модулю excel_loader
sys.path.insert(0, os.path.join(os.path.dirname(__file__), "excel_loader"))

try:
    from excel_loader.loader import (
        select_and_load_excel,
        get_available_configs,
        load_largest_file,
    )
except ImportError as e:
    print(f"Ошибка импорта excel_loader: {e}")
    sys.exit(1)


# КОНСТАНТЫ ДЛЯ ФИЛЬТРАЦИИ ДАННЫХ
# ================================

# Для фильтрации баланса Вити
VITYA_BALANCE_AVAILABLE = "Имеются в нал."

# Для фильтрации баланса Димы
DIMI_BALANCE_EXPECTED = "Ожидается"

# Минимальная цена для фильтрации (исключаем 0 и NaN)
MIN_PRICE_THRESHOLD = 0.01


class MiStockSyncApp:
    def __init__(self, root):
        self.root = root
        self.root.title("MiStockSync - Управление прайсами")
        self.root.geometry("800x600")

        # Настройка логирования
        self.setup_logging()

        # Данные
        self.current_df = None
        self.current_config = None
        self.base_df = None
        self.auto_load_base = tk.BooleanVar(value=True)  # Чекбокс автозагрузки базы
        self.comparison_result = None  # Результаты сравнения

        # Создаем интерфейс
        self.create_widgets()

        # Загружаем доступные конфиги
        self.load_available_configs()

    def setup_logging(self):
        """Настройка системы логирования"""
        # Создаем логгер для приложения
        self.logger = logging.getLogger("MiStockSync")
        self.logger.setLevel(logging.INFO)

        # Удаляем существующие обработчики
        for handler in self.logger.handlers[:]:
            self.logger.removeHandler(handler)

        # Консольный обработчик
        console_handler = logging.StreamHandler()
        console_handler.setLevel(logging.INFO)

        # Формат сообщений
        formatter = logging.Formatter(
            "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        )
        console_handler.setFormatter(formatter)

        # Добавляем обработчик
        self.logger.addHandler(console_handler)

        # Настройка файлового логирования
        logs_dir = "logs"
        if not os.path.exists(logs_dir):
            os.makedirs(logs_dir)

        log_file = os.path.join(
            logs_dir, f"mistocksync_{datetime.now().strftime('%Y%m%d')}.log"
        )
        file_handler = logging.FileHandler(log_file, encoding="utf-8")
        file_handler.setLevel(logging.INFO)
        file_handler.setFormatter(formatter)
        self.logger.addHandler(file_handler)

        self.logger.info("🚀 MiStockSync запущен")
        self.logger.info("📋 Система логирования настроена")

    def create_widgets(self):
        """Создание элементов интерфейса"""

        # Главный фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Настройка растяжения
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(2, weight=1)

        # Заголовок
        title_label = ttk.Label(
            main_frame,
            text="MiStockSync - Синхронизация прайсов",
            font=("Arial", 16, "bold"),
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 20))

        # Выбор конфигурации
        config_frame = ttk.LabelFrame(main_frame, text="Выбор поставщика", padding="10")
        config_frame.grid(
            row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10)
        )
        config_frame.columnconfigure(1, weight=1)

        ttk.Label(config_frame, text="Конфигурация:").grid(
            row=0, column=0, sticky=tk.W, padx=(0, 10)
        )

        self.config_var = tk.StringVar()
        self.config_combo = ttk.Combobox(
            config_frame, textvariable=self.config_var, state="readonly"
        )
        self.config_combo.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        # Кнопки загрузки
        buttons_frame = ttk.Frame(config_frame)
        buttons_frame.grid(row=0, column=2, sticky=tk.E)

        ttk.Button(
            buttons_frame, text="📁 Выбрать файл", command=self.select_file
        ).grid(row=0, column=0, padx=(0, 5))

        # НОВЫЙ ЧЕКБОКС ВМЕСТО КНОПКИ
        ttk.Checkbutton(
            buttons_frame, text="📊 Загрузка базы авто", variable=self.auto_load_base
        ).grid(row=0, column=1, sticky=tk.W)

        # Область вывода информации
        info_frame = ttk.LabelFrame(main_frame, text="Информация о файле", padding="10")
        info_frame.grid(
            row=2, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10)
        )
        info_frame.columnconfigure(0, weight=1)
        info_frame.rowconfigure(0, weight=1)

        self.info_text = scrolledtext.ScrolledText(info_frame, width=80, height=15)
        self.info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Кнопки действий
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(
            row=3, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0)
        )

        ttk.Button(action_frame, text="🔄 Очистить", command=self.clear_info).grid(
            row=0, column=0, sticky=tk.W
        )
        ttk.Button(
            action_frame, text="📋 Показать данные", command=self.show_data_sample
        ).grid(row=0, column=1, padx=(10, 0))
        ttk.Button(
            action_frame, text="💾 Сохранить обработанный", command=self.save_data
        ).grid(row=0, column=2, padx=(10, 0))
        ttk.Button(
            action_frame, text="🔍 Сравнить с базой", command=self.compare_with_base
        ).grid(row=0, column=3, padx=(10, 0))

        # Новые кнопки после сравнения
        self.report_button = ttk.Button(
            action_frame,
            text="📊 Сохранить отчет",
            command=self.save_report,
            state="disabled",
        )
        self.report_button.grid(row=0, column=4, padx=(10, 0))

        self.add_to_base_button = ttk.Button(
            action_frame, text="📥 Добавить в базу", command=self.add_to_base
        )
        self.add_to_base_button.grid(row=0, column=5, padx=(10, 0))

        # Статус бар
        self.status_var = tk.StringVar(value="Готов к работе")
        status_bar = ttk.Label(
            main_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W
        )
        status_bar.grid(
            row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0)
        )

    def load_available_configs(self):
        """Загрузка списка доступных конфигураций"""
        self.log_info("📋 Загрузка доступных конфигураций...")
        try:
            configs = get_available_configs()
            self.config_combo["values"] = configs

            # НОВОЕ: Устанавливаем "auto" по умолчанию
            if "auto" in configs:
                self.config_combo.set("auto")
            elif configs:
                self.config_combo.set(configs[0])

            self.log_info(f"✅ Найдено конфигураций: {len(configs)}")
            self.log_info(f"📋 Доступные конфиги: {', '.join(configs)}")
        except Exception as e:
            self.log_error(f"Ошибка загрузки конфигураций: {e}")

    def select_file(self):
        """Выбор и загрузка файла"""
        self.log_info("📁 Выбор файла для загрузки...")

        # Создаем папку data/input если её нет
        input_dir = "data/input"
        if not os.path.exists(input_dir):
            os.makedirs(input_dir)
            self.log_info(f"📁 Создана папка: {input_dir}")

        # Сначала показываем диалог выбора файла
        from tkinter import filedialog

        file_path = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            initialdir=input_dir,
        )

        if not file_path:
            self.log_info("ℹ️ Выбор файла отменен")
            return

        self.log_info(f"📁 Выбран файл: {os.path.basename(file_path)}")

        # НОВОЕ: Автоматически определяем конфиг
        if self.config_var.get() == "auto":
            detected_config = self.auto_select_config(file_path)
            config_name = detected_config
        else:
            config_name = self.config_var.get()

        if not config_name:
            messagebox.showwarning(
                "Предупреждение", "Не удалось определить конфигурацию"
            )
            return

        try:
            self.status_var.set("Загрузка файла...")
            self.root.update()

            # Загружаем файл с определенным конфигом
            from excel_loader.loader import load_with_config

            df = load_with_config(file_path, config_name)

            if df is not None:
                self.current_df = df
                self.current_config = config_name
                self.show_file_info(df, config_name)
                self.status_var.set("Файл загружен успешно")

                # Сбрасываем конфигурацию на "auto" для следующей загрузки
                if "auto" in self.config_combo["values"]:
                    self.config_combo.set("auto")
                    self.log_info(
                        "🔄 Конфигурация сброшена на 'auto' для следующей загрузки"
                    )
            else:
                self.status_var.set("Файл не загружен")

        except Exception as e:
            self.log_error(f"Ошибка загрузки файла: {e}")
            self.status_var.set("Ошибка загрузки")

    def load_largest(self):
        """Загрузка самого большого файла"""

        # Директория с данными
        data_dir = "data/input"

        try:
            self.status_var.set("Поиск самого большого файла...")
            self.root.update()

            # Находим самый большой файл
            excel_files = []
            for file in os.listdir(data_dir):
                if file.endswith((".xlsx", ".xls")):
                    file_path = os.path.join(data_dir, file)
                    file_size = os.path.getsize(file_path)
                    excel_files.append((file_path, file_size))

            if not excel_files:
                self.status_var.set("Файлы не найдены")
                return

            largest_file_path = max(excel_files, key=lambda x: x[1])[0]

            # НОВОЕ: Автоматически определяем конфиг для самого большого файла
            if self.config_var.get() == "auto":
                detected_config = self.auto_select_config(largest_file_path)
                config_name = detected_config
            else:
                config_name = self.config_var.get()

            from excel_loader.loader import load_with_config

            df = load_with_config(largest_file_path, config_name)

            if df is not None:
                self.current_df = df
                self.current_config = config_name
                self.show_file_info(df, config_name)
                self.status_var.set("Самый большой файл загружен")
            else:
                self.status_var.set("Файл не загружен")

        except Exception as e:
            self.log_error(f"Ошибка загрузки самого большого файла: {e}")
            self.status_var.set("Ошибка загрузки")

    def show_file_info(self, df, config_name):
        """Показ информации о загруженном файле"""
        self.log_info(f"📊 Отображение информации о файле (конфиг: {config_name})")
        self.clear_info()

        # Основная информация
        info = f"📊 ИНФОРМАЦИЯ О ФАЙЛЕ\n"
        info += f"{'='*50}\n"
        info += f"Конфигурация: {config_name}\n"
        info += f"Дата загрузки: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n"
        info += f"Строк: {len(df):,}\n"
        info += f"Столбцов: {len(df.columns):,}\n"
        info += f"Размер в памяти: {df.memory_usage(deep=True).sum() / 1024 / 1024:.2f} MB\n\n"

        # Информация о столбцах
        info += f"📋 СТОЛБЦЫ ({len(df.columns)}):\n"
        info += f"{'-'*30}\n"
        for i, col in enumerate(df.columns, 1):
            info += f"{i:2d}. {col}\n"

        # Типы данных
        info += f"\n📊 ТИПЫ ДАННЫХ:\n"
        info += f"{'-'*30}\n"
        for col in df.columns:
            non_null = df[col].notna().sum()
            info += f"{col}: {str(df[col].dtype)} ({non_null:,} не пустых)\n"

        # Статистика по пустым значениям
        info += f"\n❌ ПУСТЫЕ ЗНАЧЕНИЯ:\n"
        info += f"{'-'*30}\n"
        null_counts = df.isnull().sum()
        for col in df.columns:
            if null_counts[col] > 0:
                info += f"{col}: {null_counts[col]:,} пустых\n"

        if null_counts.sum() == 0:
            info += "Пустых значений нет! ✅\n"

        self.info_text.insert(tk.END, info)
        self.log_info(
            f"✅ Файл загружен: {len(df)} строк, {len(df.columns)} столбцов, {df.memory_usage(deep=True).sum() / 1024 / 1024:.2f} MB"
        )

    def show_data_sample(self):
        """Показ образца данных"""
        if self.current_df is None:
            messagebox.showwarning("Предупреждение", "Сначала загрузите файл")
            return

        self.clear_info()

        df = self.current_df

        info = f"📋 ОБРАЗЕЦ ДАННЫХ (первые 10 строк)\n"
        info += f"{'='*80}\n\n"

        # Показываем первые 10 строк
        sample_df = df.head(10)
        info += sample_df.to_string(max_cols=10, max_colwidth=20) + "\n\n"

        # Уникальные значения для важных столбцов
        important_cols = [
            "article",
            "name",
            "price",
            "article_vitya",
            "price_usd",
            "price_rub",
        ]
        existing_cols = [col for col in important_cols if col in df.columns]

        if existing_cols:
            info += f"📊 УНИКАЛЬНЫЕ ЗНАЧЕНИЯ (топ-10):\n"
            info += f"{'-'*50}\n"
            for col in existing_cols[:3]:  # Показываем только первые 3
                unique_vals = df[col].value_counts().head(10)
                info += f"\n{col.upper()}:\n"
                for val, count in unique_vals.items():
                    info += f"  {val}: {count:,}\n"

        self.info_text.insert(tk.END, info)

    def save_data(self):
        """Сохранение обработанных данных"""
        self.log_info("💾 Начало сохранения обработанных данных...")

        if self.current_df is None:
            self.log_error("Файл не загружен")
            messagebox.showwarning("Предупреждение", "Сначала загрузите файл")
            return

        from tkinter import filedialog

        # Создаем папку data/output если её нет
        output_dir = "data/output"
        if not os.path.exists(output_dir):
            os.makedirs(output_dir)
            self.log_info(f"📁 Создана папка: {output_dir}")

        self.log_info(f"📁 Открываем диалог сохранения в папке: {output_dir}")

        file_path = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            filetypes=[
                ("Excel files", "*.xlsx"),
                ("CSV files", "*.csv"),
                ("All files", "*.*"),
            ],
            initialdir=output_dir,
        )

        if file_path:
            try:
                self.status_var.set("Предобработка данных...")
                self.root.update()

                # Предобрабатываем данные перед сохранением
                processed_df = self.preprocess_supplier_data(
                    self.current_df, self.current_config
                )

                if file_path.endswith(".xlsx"):
                    processed_df.to_excel(file_path, index=False)
                elif file_path.endswith(".csv"):
                    processed_df.to_csv(file_path, index=False, encoding="utf-8")

                self.log_info(f"Обработанные данные сохранены: {file_path}")
                self.log_info(
                    f"Исходно: {len(self.current_df)} строк → Обработано: {len(processed_df)} строк"
                )
                messagebox.showinfo(
                    "Успех", f"Обработанные данные сохранены в {file_path}"
                )
                self.status_var.set("Готов к работе")

            except Exception as e:
                self.log_error(f"Ошибка сохранения: {e}")
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")
                self.status_var.set("Ошибка сохранения")

    def compare_with_base(self):
        """Сравнение текущего файла с базой данных"""
        self.log_info("🔍 Начало сравнения с базой данных...")

        if self.current_df is None:
            self.log_error("Файл поставщика не загружен")
            messagebox.showwarning(
                "Предупреждение", "Сначала загрузите файл поставщика"
            )
            return

        # Автозагрузка базы (как описано выше)
        # НОВАЯ ЛОГИКА: проверяем чекбокс
        if self.auto_load_base.get():
            self.status_var.set("Автозагрузка базы данных...")
            self.root.update()

            # Загружаем базу данных если еще не загружена
            if self.base_df is None:
                data_dir = "data/input"

                self.base_df = load_largest_file(data_dir, "base")
                if self.base_df is None:
                    messagebox.showerror("Ошибка", "Не удалось загрузить базу данных")
                    return

                self.log_info("База данных автоматически загружена")
        else:
            # Если чекбокс выключен, база должна быть загружена вручную
            if self.base_df is None:
                messagebox.showwarning(
                    "Предупреждение", "Сначала загрузите базу данных"
                )
                return

        # НОВОЕ: Предобработка данных поставщика
        self.status_var.set("Предобработка данных поставщика...")
        self.root.update()

        processed_supplier_df = self.preprocess_supplier_data(
            self.current_df, self.current_config
        )

        # Выполняем сравнение с предобработанными данными
        self.status_var.set("Сравнение с базой...")
        self.root.update()

        comparison_result = self.perform_comparison(processed_supplier_df, self.base_df)
        self.show_comparison_result(comparison_result)

        # Сохраняем результат сравнения и активируем кнопку отчета
        self.comparison_result = comparison_result

        # Проверяем нет ли ошибки в результате сравнения
        if "error" in comparison_result:
            self.log_error(
                f"Ошибка в результате сравнения: {comparison_result['error']}"
            )
            self.log_info("❌ Кнопка 'Сохранить отчет' НЕ активирована из-за ошибки")
        else:
            self.log_info("🔘 Активируем кнопку 'Сохранить отчет'...")
            self.report_button.config(state="normal")
            self.log_info("✅ Кнопка 'Сохранить отчет' активирована")

        self.status_var.set("Сравнение завершено")

    def perform_comparison(self, supplier_df, base_df):
        """Выполняет сравнение файла поставщика с базой данных"""

        # Определяем ключевые столбцы для сравнения
        if self.current_config == "vitya":
            supplier_article_col = "article_vitya"
            base_article_col = "article_vitya"
            supplier_price_col = "price_usd"
            base_price_col = "price_vitya_usd"
        elif self.current_config == "dimi":
            supplier_article_col = "article_dimi"
            base_article_col = "article_dimi"
            supplier_price_col = "price_usd"
            base_price_col = "price_dimi_usd"
        else:
            # Для других конфигов используем общие столбцы
            supplier_article_col = "article"
            base_article_col = "article"
            supplier_price_col = "price"
            base_price_col = "price"

        # Проверяем наличие нужных столбцов
        if supplier_article_col not in supplier_df.columns:
            return {
                "error": f"Столбец {supplier_article_col} не найден в файле поставщика"
            }

        if base_article_col not in base_df.columns:
            return {"error": f"Столбец {base_article_col} не найден в базе данных"}

        # Очищаем данные от NaN и пустых значений
        supplier_clean = supplier_df.dropna(
            subset=[supplier_article_col, supplier_price_col]
        )
        base_clean = base_df.dropna(subset=[base_article_col])

        # Создаем словари для быстрого поиска
        supplier_dict = {}
        for _, row in supplier_clean.iterrows():
            article_value = row[supplier_article_col]
            # Для article_vitya используем int значение напрямую, для других - строку
            if self.current_config == "vitya" and isinstance(article_value, int):
                article = str(article_value)
            else:
                article = str(article_value).strip()

            if article and article != "nan" and article != "None":
                supplier_dict[article] = {
                    "price": (
                        row[supplier_price_col]
                        if pd.notna(row[supplier_price_col])
                        else 0
                    ),
                    "name": row.get("name", ""),
                    "index": row.name,
                }

        base_dict = {}
        for _, row in base_clean.iterrows():
            article_value = row[base_article_col]
            # Для article_vitya используем int значение напрямую, для других - строку
            if self.current_config == "vitya" and isinstance(article_value, int):
                article = str(article_value)
            else:
                article = str(article_value).strip()

            if article and article != "nan" and article != "None":
                base_dict[article] = {
                    "price": (
                        row[base_price_col] if pd.notna(row[base_price_col]) else 0
                    ),
                    "name": row.get("name", ""),
                    "index": row.name,
                }

        # Анализируем совпадения
        matches = []
        price_changes = []
        new_items = []

        for article, supplier_data in supplier_dict.items():
            if article in base_dict:
                base_data = base_dict[article]
                match_info = {
                    "article": article,
                    "supplier_price": supplier_data["price"],
                    "base_price": base_data["price"],
                    "name": supplier_data["name"] or base_data["name"],
                    "price_diff": supplier_data["price"] - base_data["price"],
                    "price_change_percent": 0,
                }

                if base_data["price"] > 0:
                    match_info["price_change_percent"] = (
                        (supplier_data["price"] - base_data["price"])
                        / base_data["price"]
                        * 100
                    )

                matches.append(match_info)

                # Значительные изменения цены (больше 5%)
                if abs(match_info["price_change_percent"]) > 5:
                    price_changes.append(match_info)
            else:
                new_items.append(
                    {
                        "article": article,
                        "price": supplier_data["price"],
                        "name": supplier_data["name"],
                    }
                )

        # НОВОЕ: Поиск по кодам в наименованиях, если мало совпадений по артикулам
        code_matches = []
        if len(matches) < len(supplier_dict) * 0.3:  # Если меньше 30% совпадений
            self.log_info(
                "🔍 Мало совпадений по артикулам, запускаем поиск по кодам..."
            )
            code_matches = self.compare_by_product_code(
                supplier_df, base_df, self.current_config
            )

        return {
            "supplier_total": len(supplier_dict),
            "base_total": len(base_dict),
            "matches": matches,
            "price_changes": price_changes,
            "new_items": new_items,
            "code_matches": code_matches,  # Новое поле
            "match_rate": (
                len(matches) / len(supplier_dict) * 100 if supplier_dict else 0
            ),
        }

    def show_comparison_result(self, result):
        """Показ результатов сравнения"""
        if "error" in result:
            messagebox.showerror("Ошибка", result["error"])
            return

        self.clear_info()

        info = f"🔍 РЕЗУЛЬТАТЫ СРАВНЕНИЯ С БАЗОЙ ДАННЫХ\n"
        info += f"{'='*60}\n"
        info += f"Конфигурация: {self.current_config}\n"
        info += f"Дата сравнения: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n"

        # Общая статистика
        info += f"📊 ОБЩАЯ СТАТИСТИКА:\n"
        info += f"{'-'*40}\n"
        info += f"Товаров у поставщика: {result['supplier_total']:,}\n"
        info += f"Товаров в базе: {result['base_total']:,}\n"
        info += f"Совпадений найдено: {len(result['matches']):,}\n"
        info += f"Процент совпадений: {result['match_rate']:.1f}%\n\n"

        # Изменения цен
        if result["price_changes"]:
            info += f"💰 ЗНАЧИТЕЛЬНЫЕ ИЗМЕНЕНИЯ ЦЕН (>5%):\n"
            info += f"{'-'*50}\n"
            for i, item in enumerate(result["price_changes"][:10], 1):
                change_sign = "📈" if item["price_diff"] > 0 else "📉"
                info += f"{i:2d}. {item['article']}: {item['base_price']:.2f} → {item['supplier_price']:.2f} "
                info += f"({item['price_change_percent']:+.1f}%) {change_sign}\n"
            if len(result["price_changes"]) > 10:
                info += f"... и еще {len(result['price_changes']) - 10} изменений\n"
        else:
            info += f"💰 ИЗМЕНЕНИЯ ЦЕН: Значительных изменений не найдено ✅\n"

        info += "\n"

        # Новые товары
        if result["new_items"]:
            info += f"🆕 НОВЫЕ ТОВАРЫ У ПОСТАВЩИКА:\n"
            info += f"{'-'*40}\n"
            for i, item in enumerate(result["new_items"][:10], 1):
                info += f"{i:2d}. {item['article']}: {item['price']:.2f} - {item['name'][:30]}\n"
            if len(result["new_items"]) > 10:
                info += f"... и еще {len(result['new_items']) - 10} новых товаров\n"
        else:
            info += f"🆕 НОВЫЕ ТОВАРЫ: Не найдено\n"

        # Результаты поиска по кодам
        if result.get("code_matches"):
            info += f"\n🔍 СОВПАДЕНИЯ ПО КОДАМ В НАИМЕНОВАНИЯХ:\n"
            info += f"{'-'*50}\n"
            for i, match in enumerate(result["code_matches"][:10], 1):
                info += f"{i:2d}. Код: {match['code']}\n"
                info += f"    Поставщик: {match['supplier_name'][:40]}...\n"
                info += f"    База: {match['base_name'][:40]}...\n"
                info += f"    Цены: {match['supplier_price']:.2f} ↔ {match['base_price']:.2f}\n\n"
            if len(result["code_matches"]) > 10:
                info += f"... и еще {len(result['code_matches']) - 10} совпадений по кодам\n"

        info += f"\n🎉 СРАВНЕНИЕ ЗАВЕРШЕНО!"

        self.info_text.insert(tk.END, info)
        self.log_info(
            f"Сравнение завершено: {len(result['matches'])} совпадений из {result['supplier_total']} товаров"
        )

    def clean_invisible_chars(self, text):
        """Убирает невидимые и непечатаемые символы из текста"""
        if pd.isna(text):
            return None

        # Преобразуем в строку
        text_str = str(text).strip()

        # ЗАКОММЕНТИРОВАНО: Убираем невидимые и непечатаемые символы
        # import unicodedata
        # cleaned = "".join(
        #     char
        #     for char in text_str
        #     if unicodedata.category(char)[0]
        #     in ("L", "N", "P", "S", "M")  # Буквы, цифры, знаки, символы, диакритики
        #     or char in (" ", "\t", "\n")  # Обычные пробелы
        #     or unicodedata.category(char) == "Zs"  # Обычные пробелы (Space separators)
        # )
        # cleaned = " ".join(cleaned.split())

        # УПРОЩЕНО: Просто убираем лишние пробелы
        cleaned = " ".join(text_str.split())

        return cleaned if cleaned else None

    def clean_article_vitya_simple(self, article):
        """Простая очистка артикула Вити - убираем ТОЛЬКО апострофы и префикс '000, результат ВСЕГДА int"""
        # Проверяем на NaN или None
        if pd.isna(article) or article is None:
            return None

        # Преобразуем в строку
        cleaned = str(article).strip()

        # Проверяем, что получилась не пустая строка и не 'nan'
        if not cleaned or cleaned.lower() == "nan":
            return None

        # 1. Убираем ТОЛЬКО апострофы
        cleaned = cleaned.replace("'", "")

        # 2. Убираем префикс '000 если есть
        if cleaned.startswith("000"):
            cleaned = cleaned[3:]

        # 3. ВСЕГДА преобразуем в int
        if cleaned.isdigit():
            return int(cleaned) if cleaned else 0
        elif cleaned == "":
            return 0
        else:
            # Если есть нецифровые символы, извлекаем только цифры
            import re

            digits = re.findall(r"\d+", cleaned)
            if digits:
                return int("".join(digits))
            else:
                return 0

    def filter_by_price(self, df, price_column="price_usd"):
        """
        Фильтрация данных по цене - убирает строки где price_usd является NaN, пустой или <= 0

        Args:
            df: DataFrame для фильтрации
            price_column: название столбца с ценой (по умолчанию 'price_usd')

        Returns:
            DataFrame с отфильтрованными данными
        """
        if price_column not in df.columns:
            self.log_info(
                f"⚠️ Столбец '{price_column}' не найден, фильтрация по цене пропущена"
            )
            return df

        initial_count = len(df)

        # Фильтруем: убираем NaN, пустые значения и цены <= MIN_PRICE_THRESHOLD
        filtered_df = df[
            (df[price_column].notna()) & (df[price_column] > MIN_PRICE_THRESHOLD)
        ].copy()

        final_count = len(filtered_df)
        removed_count = initial_count - final_count

        if removed_count > 0:
            self.log_info(f"💰 Фильтрация по цене ({price_column}):")
            self.log_info(f"   Удалено строк: {removed_count}")
            self.log_info(f"   Осталось строк: {final_count}")

            # Показываем статистику удаленных строк
            nan_count = df[price_column].isna().sum()
            zero_count = (df[price_column] == 0).sum()
            low_price_count = (
                (df[price_column] > 0) & (df[price_column] <= MIN_PRICE_THRESHOLD)
            ).sum()

            self.log_info(f"   📊 Причины удаления:")
            if nan_count > 0:
                self.log_info(f"      NaN/пустые: {nan_count}")
            if zero_count > 0:
                self.log_info(f"      Нулевые цены: {zero_count}")
            if low_price_count > 0:
                self.log_info(
                    f"      Слишком низкие (<={MIN_PRICE_THRESHOLD}): {low_price_count}"
                )
        else:
            self.log_info(
                f"✅ Фильтрация по цене: все {final_count} строк прошли фильтр"
            )

        return filtered_df

    def preprocess_vitya_fixed_v3(self, df):
        """ИСПРАВЛЕННАЯ предобработка для Вити с фильтрацией по цене и балансу"""
        self.log_info("🔧 Запуск предобработки для Витя...")

        # Копируем данные для безопасности
        processed_df = df.copy()
        initial_count = len(processed_df)

        # 1. Фильтрация по цене - убираем строки с NaN, пустыми или нулевыми ценами
        self.log_info("💰 Фильтруем по цене...")
        processed_df = self.filter_by_price(processed_df, "price_usd")

        # 2. Фильтрация по балансу - оставляем только товары в наличии
        if "balance" in processed_df.columns:
            self.log_info(
                f"📦 Фильтруем по балансу (оставляем только '{VITYA_BALANCE_AVAILABLE}')..."
            )

            balance_before = len(processed_df)
            processed_df = processed_df[
                processed_df["balance"] == VITYA_BALANCE_AVAILABLE
            ].copy()
            balance_after = len(processed_df)

            removed_balance = balance_before - balance_after
            if removed_balance > 0:
                self.log_info(f"   📦 Удалено строк без наличия: {removed_balance}")
                self.log_info(f"   📦 Осталось строк в наличии: {balance_after}")
            else:
                self.log_info(f"   📦 Все {balance_after} строк имеют товары в наличии")
        else:
            self.log_info(
                "⚠️ Столбец 'balance' не найден, фильтрация по наличию пропущена"
            )

        # 3. Очистка артикулов - активируем очистку
        if "article_vitya" in processed_df.columns:
            self.log_info("🧹 Очистка артикулов Витя...")

            processed_df["article_vitya"] = processed_df["article_vitya"].apply(
                self.clean_article_vitya_simple
            )

        # 4. Добавляем метку поставщика
        processed_df["supplier_name"] = "Витя"

        # 5. Отладочная информация
        if "article_vitya" in processed_df.columns:
            valid_articles = processed_df["article_vitya"].notna().sum()
            self.log_info(f"🔢 Валидных артикулов Витя: {valid_articles}")
            if valid_articles > 0:
                sample_articles = (
                    processed_df["article_vitya"].dropna().head(5).tolist()
                )
                self.log_info(f"📝 Примеры артикулов: {sample_articles}")

        # 6. Финальная статистика
        final_count = len(processed_df)
        total_removed = initial_count - final_count

        self.log_info(f"✅ Предобработка Витя завершена:")
        self.log_info(f"   📊 Исходно: {initial_count} строк")
        self.log_info(f"   📊 Итого: {final_count} строк")
        self.log_info(f"   📊 Удалено: {total_removed} строк")

        return processed_df

    def preprocess_dimi_fixed(self, df):
        """УПРОЩЕННАЯ предобработка данных для поставщика Дима с фильтрацией"""
        self.log_info("🔧 Запуск предобработки для Дима...")

        # Копируем данные
        processed_df = df.copy()
        initial_count = len(processed_df)

        # 1. Фильтрация по цене - убираем строки с NaN, пустыми или нулевыми ценами
        self.log_info("💰 Фильтруем по цене...")
        processed_df = self.filter_by_price(processed_df, "price_usd")

        # 2. Фильтрация по балансу - убираем строки где balance или balance1 = "Ожидается"
        balance_columns = ["balance", "balance1"]
        found_balance_columns = [
            col for col in balance_columns if col in processed_df.columns
        ]

        if found_balance_columns:
            self.log_info(
                f"📦 Фильтруем по балансу (убираем '{DIMI_BALANCE_EXPECTED}')..."
            )

            balance_before = len(processed_df)

            # Создаем условие: НИ balance, НИ balance1 не должны быть "Ожидается"
            for col in found_balance_columns:
                processed_df = processed_df[processed_df[col] != DIMI_BALANCE_EXPECTED]

            processed_df = processed_df.copy()
            balance_after = len(processed_df)

            removed_balance = balance_before - balance_after
            if removed_balance > 0:
                self.log_info(
                    f"   📦 Удалено строк с '{DIMI_BALANCE_EXPECTED}': {removed_balance}"
                )
                self.log_info(f"   📦 Осталось строк в наличии: {balance_after}")

                # Показываем детализацию по столбцам
                for col in found_balance_columns:
                    expected_count = (df[col] == DIMI_BALANCE_EXPECTED).sum()
                    if expected_count > 0:
                        self.log_info(
                            f"      {col}: {expected_count} строк с '{DIMI_BALANCE_EXPECTED}'"
                        )
            else:
                self.log_info(
                    f"   📦 Все {balance_after} строк прошли фильтр по балансу"
                )
        else:
            self.log_info(
                "⚠️ Столбцы balance/balance1 не найдены, фильтрация по наличию пропущена"
            )

        # 3. Очистка артикулов - активируем очистку
        if "article_dimi" in processed_df.columns:
            self.log_info("🧹 Очистка артикулов Дима...")

            def clean_article_dimi_simple(article):
                """Упрощенная очистка артикула Димы - ТОЛЬКО апострофы и префикс '000"""
                if pd.isna(article):
                    return None

                # Преобразуем в строку и убираем лишние пробелы
                cleaned = str(article).strip()

                if not cleaned or cleaned.lower() == "nan":
                    return None

                # Убираем апострофы
                cleaned = cleaned.replace("'", "")

                # ДЛЯ ДИМЫ: Убираем префикс '000 если есть
                if cleaned.startswith("000"):
                    cleaned = cleaned[3:]

                return cleaned if cleaned else None

            processed_df["article_dimi"] = processed_df["article_dimi"].apply(
                clean_article_dimi_simple
            )

        # 4. Добавляем метку поставщика
        processed_df["supplier_name"] = "Дима"

        # 5. Отладочная информация
        if "article_dimi" in processed_df.columns:
            valid_articles = processed_df["article_dimi"].notna().sum()
            self.log_info(f"🔢 Валидных артикулов Дима: {valid_articles}")
            if valid_articles > 0:
                sample_articles = processed_df["article_dimi"].dropna().head(5).tolist()
                self.log_info(f"📝 Примеры артикулов: {sample_articles}")

        # 6. Финальная статистика
        final_count = len(processed_df)
        total_removed = initial_count - final_count

        self.log_info(f"✅ Предобработка Дима завершена:")
        self.log_info(f"   📊 Исходно: {initial_count} строк")
        self.log_info(f"   📊 Итого: {final_count} строк")
        self.log_info(f"   📊 Удалено: {total_removed} строк")

        return processed_df

    def preprocess_supplier_data(self, df, config_name):
        """Универсальная предобработка в зависимости от конфига"""

        if config_name == "vitya":
            return self.preprocess_vitya_fixed_v3(df)
        elif config_name == "dimi":
            return self.preprocess_dimi_fixed(df)
        else:
            self.log_info(f"📋 Предобработка для {config_name} не требуется")
            return df

    def detect_config_by_filename(self, file_path):
        """Автоматическое определение конфига по имени файла"""

        filename = os.path.basename(file_path).upper()  # Имя файла в верхнем регистре

        self.log_info(f"🔍 Определение конфига для файла: {filename}")

        # Правила определения конфига
        if "JHT" in filename:
            detected_config = "vitya"
            self.log_info("✅ Обнаружен прайс Вити (содержит JHT)")

        elif "DIMI" in filename or "DIMA" in filename:
            detected_config = "dimi"
            self.log_info("✅ Обнаружен прайс Димы (содержит DiMi/DiMa)")

        elif "BASE" in filename or "БАЗА" in filename:
            detected_config = "base"
            self.log_info("✅ Обнаружена база данных (содержит BASE/БАЗА)")

        else:
            detected_config = "auto"  # По умолчанию
            self.log_info("ℹ️ Конфиг не определен, используется AUTO")

        return detected_config

    def find_product_code_in_name(self, product_name):
        """Извлечение кода товара из наименования"""
        # БОЛВАНКА ДЛЯ БУДУЩЕЙ РЕАЛИЗАЦИИ
        # TODO: Добавить логику поиска кодов в наименовании

        if pd.isna(product_name) or not isinstance(product_name, str):
            return None

        import re

        # Примеры паттернов для поиска кодов:
        patterns = [
            r"\b\d{6,}\b",  # 6+ цифр подряд
            r"[A-Z]{2,}\d{3,}",  # Буквы + цифры (XM123)
            r"\d{3,}[A-Z]{1,2}",  # Цифры + буквы (123XM)
            # TODO: Добавить специфичные паттерны для каждого поставщика
        ]

        for pattern in patterns:
            matches = re.findall(pattern, product_name.upper())
            if matches:
                # Возвращаем первое найденное совпадение
                return matches[0]

        return None

    def compare_by_product_code(self, supplier_df, base_df, supplier_config):
        """Поиск совпадений по кодам товаров в наименованиях"""
        # БОЛВАНКА ДЛЯ БУДУЩЕЙ РЕАЛИЗАЦИИ
        # TODO: Реализовать логику сравнения по кодам в наименованиях

        self.log_info("🔍 Поиск совпадений по кодам в наименованиях...")

        code_matches = []

        # Извлекаем коды из наименований поставщика
        supplier_codes = {}
        for idx, row in supplier_df.iterrows():
            if "name" in row and pd.notna(row["name"]):
                code = self.find_product_code_in_name(row["name"])
                if code:
                    supplier_codes[code] = {
                        "index": idx,
                        "name": row["name"],
                        "price": (
                            row.get("price_usd", 0)
                            if supplier_config == "vitya"
                            else row.get("price", 0)
                        ),
                    }

        # Извлекаем коды из наименований базы
        base_codes = {}
        for idx, row in base_df.iterrows():
            if "name" in row and pd.notna(row["name"]):
                code = self.find_product_code_in_name(row["name"])
                if code:
                    base_codes[code] = {
                        "index": idx,
                        "name": row["name"],
                        "price": row.get("price", 0),
                    }

        # Ищем совпадения
        for code, supplier_data in supplier_codes.items():
            if code in base_codes:
                base_data = base_codes[code]
                code_matches.append(
                    {
                        "code": code,
                        "supplier_name": supplier_data["name"],
                        "base_name": base_data["name"],
                        "supplier_price": supplier_data["price"],
                        "base_price": base_data["price"],
                    }
                )

        self.log_info(f"✅ Найдено совпадений по кодам: {len(code_matches)}")
        return code_matches

    def auto_select_config(self, file_path):
        """Автоматический выбор и установка конфига"""

        detected_config = self.detect_config_by_filename(file_path)

        # Проверяем, есть ли такой конфиг в списке доступных
        available_configs = self.config_combo["values"]

        if detected_config in available_configs:
            self.config_combo.set(detected_config)
            self.log_info(f"🎯 Конфиг автоматически изменен на: {detected_config}")
            return detected_config
        else:
            self.log_info(f"⚠️ Конфиг {detected_config} не найден, оставляем текущий")
            return self.config_var.get()

    def clear_info(self):
        """Очистка области информации"""
        self.info_text.delete(1.0, tk.END)

    def log_info(self, message):
        """Логирование информации"""
        # Логируем в консоль и файл
        self.logger.info(message)

        # Также выводим в GUI
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] {message}\n"
        self.info_text.insert(tk.END, log_message)
        self.info_text.see(tk.END)

    def log_error(self, message):
        """Логирование ошибок"""
        # Логируем в консоль и файл
        self.logger.error(f"❌ ОШИБКА: {message}")

        # Также выводим в GUI
        timestamp = datetime.now().strftime("%H:%M:%S")
        log_message = f"[{timestamp}] ❌ ОШИБКА: {message}\n"
        self.info_text.insert(tk.END, log_message)
        self.info_text.see(tk.END)

    def save_report(self):
        """Сохранение отчета о сравнении в Excel"""
        self.log_info("🔘 Нажата кнопка 'Сохранить отчет'")

        if self.comparison_result is None:
            self.log_info("❌ Результат сравнения отсутствует")
            messagebox.showwarning(
                "Предупреждение", "Сначала выполните сравнение с базой"
            )
            return

        self.log_info("✅ Результат сравнения найден, открываем диалог сохранения...")

        try:
            from tkinter import filedialog
            import os

            # Создаем папку data/output если её нет
            output_dir = "data/output"
            if not os.path.exists(output_dir):
                os.makedirs(output_dir)
                self.log_info(f"📁 Создана папка: {output_dir}")

            # Создаем имя файла с временной меткой
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_filename = f"summary_report_{timestamp}.xlsx"
            self.log_info(f"📁 Предлагаемое имя файла: {default_filename}")

            file_path = filedialog.asksaveasfilename(
                title="Сохранить отчет о сравнении",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=default_filename,
                initialdir=output_dir,
            )

            self.log_info(
                f"📁 Выбранный путь: {file_path if file_path else 'Отменено пользователем'}"
            )
        except Exception as e:
            self.log_error(f"Ошибка при открытии диалога: {e}")
            return

        if file_path:
            try:
                self.log_info("💾 Начинаем сохранение отчета...")
                self.status_var.set("Сохранение отчета...")
                self.root.update()

                # Создаем сводную таблицу
                self.log_info("📊 Создаем сводную таблицу...")
                summary_data = [
                    {
                        "Поставщик": self.current_config.upper(),
                        "Товаров": self.comparison_result["supplier_total"],
                        "Совпадений": len(self.comparison_result["matches"]),
                        "Процент совпадений": f"{self.comparison_result['match_rate']:.1f}%",
                        "Изменений цен": len(self.comparison_result["price_changes"]),
                        "Новых товаров": len(self.comparison_result["new_items"]),
                        "Совпадений по кодам": len(
                            self.comparison_result.get("code_matches", [])
                        ),
                    }
                ]
                self.log_info(f"✅ Сводная таблица создана: {summary_data[0]}")

                # Сохраняем в Excel с несколькими листами
                self.log_info("📝 Создаем Excel файл...")
                with pd.ExcelWriter(file_path, engine="openpyxl") as writer:
                    # Лист с общей сводкой
                    self.log_info("📄 Создаем лист 'Сводка'...")
                    summary_df = pd.DataFrame(summary_data)
                    summary_df.to_excel(writer, sheet_name="Сводка", index=False)

                    # Лист с совпадениями
                    if self.comparison_result["matches"]:
                        self.log_info(
                            f"📄 Создаем лист 'Совпадения' ({len(self.comparison_result['matches'])} записей)..."
                        )
                        matches_df = pd.DataFrame(self.comparison_result["matches"])
                        matches_df.to_excel(
                            writer, sheet_name="Совпадения", index=False
                        )

                    # Лист с изменениями цен
                    if self.comparison_result["price_changes"]:
                        self.log_info(
                            f"📄 Создаем лист 'Изменения цен' ({len(self.comparison_result['price_changes'])} записей)..."
                        )
                        price_changes_df = pd.DataFrame(
                            self.comparison_result["price_changes"]
                        )
                        price_changes_df.to_excel(
                            writer, sheet_name="Изменения цен", index=False
                        )

                    # Лист с новыми товарами
                    if self.comparison_result["new_items"]:
                        self.log_info(
                            f"📄 Создаем лист 'Новые товары' ({len(self.comparison_result['new_items'])} записей)..."
                        )
                        new_items_df = pd.DataFrame(self.comparison_result["new_items"])
                        new_items_df.to_excel(
                            writer, sheet_name="Новые товары", index=False
                        )

                    # Лист с совпадениями по кодам
                    if self.comparison_result.get("code_matches"):
                        self.log_info(
                            f"📄 Создаем лист 'Совпадения по кодам' ({len(self.comparison_result['code_matches'])} записей)..."
                        )
                        code_matches_df = pd.DataFrame(
                            self.comparison_result["code_matches"]
                        )
                        code_matches_df.to_excel(
                            writer, sheet_name="Совпадения по кодам", index=False
                        )

                self.log_info("✅ Excel файл создан успешно")

                self.log_info(f"📊 Отчет сохранен: {file_path}")
                self.log_info(f"   Листов создано: {len(summary_data)} + детализация")
                messagebox.showinfo("Успех", f"Отчет сохранен в {file_path}")
                self.status_var.set("Отчет сохранен")

            except Exception as e:
                self.log_error(f"Ошибка сохранения отчета: {e}")
                messagebox.showerror("Ошибка", f"Не удалось сохранить отчет: {e}")
                self.status_var.set("Ошибка сохранения отчета")
        else:
            self.log_info("ℹ️ Сохранение отчета отменено пользователем")

    def add_to_base(self):
        """Добавление товаров в базу данных (заглушка)"""
        self.log_info("🔄 Добавление товаров в базу данных...")
        self.log_info("📋 Скоро будет реализован функционал добавления товаров в базу!")
        self.log_info("🚀 Планируется:")
        self.log_info("   - Добавление новых товаров из прайса поставщика")
        self.log_info("   - Обновление цен для существующих товаров")
        self.log_info("   - Выбор товаров для добавления")
        self.log_info("   - Резервное копирование базы перед изменениями")

        messagebox.showinfo(
            "Функция в разработке",
            "Скоро будет реализован функционал добавления товаров в базу!\n\n"
            "Планируется:\n"
            "• Добавление новых товаров\n"
            "• Обновление цен\n"
            "• Выбор товаров для добавления\n"
            "• Резервное копирование",
        )


def main():
    """Главная функция приложения"""
    # Базовое логирование для main функции
    print("🚀 Запуск MiStockSync GUI...")
    print("📋 Инициализация интерфейса...")

    root = tk.Tk()
    app = MiStockSyncApp(root)

    # Центрируем окно
    root.update_idletasks()
    width = root.winfo_width()
    height = root.winfo_height()
    x = (root.winfo_screenwidth() // 2) - (width // 2)
    y = (root.winfo_screenheight() // 2) - (height // 2)
    root.geometry(f"{width}x{height}+{x}+{y}")

    app.logger.info("🖥️ GUI интерфейс готов к работе")
    print("✅ Приложение готово к работе!")

    root.mainloop()


if __name__ == "__main__":
    main()
