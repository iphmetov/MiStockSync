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
import json
import shutil

# Опциональные импорты для точечного обновления Excel
try:
    from openpyxl import load_workbook

    OPENPYXL_AVAILABLE = True
except ImportError:
    OPENPYXL_AVAILABLE = False

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

# Для фильтрации баланса Вити - список допустимых статусов
VITYA_BALANCE_AVAILABLE = ["Имеются в нал.", "Распродажа"]

# Для фильтрации баланса Димы
DIMI_BALANCE_EXPECTED = "Ожидается"

# Минимальная цена для фильтрации (исключаем 0 и NaN)
MIN_PRICE_THRESHOLD = 0.01

# Константы для обновления цен (из notebook)
MIN_PRICE_CHANGE_PERCENT = 0.1  # Минимальное изменение для обновления
MAX_PRICE_CHANGE_PERCENT = 100.0  # Максимальное разрешенное изменение
SIGNIFICANT_CHANGE_PERCENT = 20.0  # Порог "значительного" изменения


class MiStockSyncApp:
    def __init__(self, root):
        self.root = root
        # Заголовок устанавливается в main()
        self.root.geometry("1000x800")

        # Настройка логирования
        self.setup_logging()

        # Загружаем настройки из файла
        self.settings = self.load_settings()

        # Данные
        self.current_df = None
        self.current_config = None
        self.base_df = None
        self.auto_load_base = tk.BooleanVar(value=True)  # Чекбокс автозагрузки базы
        self.comparison_result = None  # Результаты сравнения

        # Настройки интерфейса (применяем загруженные настройки)
        self.current_font_size = self.settings.get("font_size", "normal")
        self.auto_load_base_enabled = self.settings.get("auto_load_base", True)

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

    def load_settings(self):
        """Загрузка настроек из файла settings.json"""
        settings_file = "settings.json"
        default_settings = {"auto_load_base": True, "font_size": "normal"}

        try:
            if os.path.exists(settings_file):
                with open(settings_file, "r", encoding="utf-8") as f:
                    settings = json.load(f)

                # Проверяем наличие всех нужных ключей
                for key, default_value in default_settings.items():
                    if key not in settings:
                        settings[key] = default_value

                self.logger.info(f"⚙️ Настройки загружены из {settings_file}")
                return settings
            else:
                self.logger.info(
                    "⚙️ Файл настроек не найден, используются значения по умолчанию"
                )
                return default_settings

        except Exception as e:
            self.logger.error(f"❌ Ошибка загрузки настроек: {e}")
            return default_settings

    def save_settings(self, settings):
        """Сохранение настроек в файл settings.json"""
        settings_file = "settings.json"

        try:
            with open(settings_file, "w", encoding="utf-8") as f:
                json.dump(settings, f, indent=2, ensure_ascii=False)

            self.logger.info(f"💾 Настройки сохранены в {settings_file}")
            return True

        except Exception as e:
            self.logger.error(f"❌ Ошибка сохранения настроек: {e}")
            return False

    def create_widgets(self):
        """Создание элементов интерфейса"""

        # Создаем главное меню
        self.create_menu()

        # Главный фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Настройка растяжения
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)
        main_frame.rowconfigure(3, weight=1)

        # Заголовок
        title_label = ttk.Label(
            main_frame,
            text="MiStockSync - Синхронизация прайсов",
            font=("Arial", 16, "bold"),
        )
        title_label.grid(row=0, column=0, columnspan=3, pady=(0, 10))

        # Мини-панель инструментов
        toolbar_frame = ttk.Frame(main_frame)
        toolbar_frame.grid(
            row=1, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10)
        )

        # Контейнер для инструментов (прижатый к левому краю)
        tools_container = ttk.Frame(toolbar_frame)
        tools_container.grid(row=0, column=0, sticky=tk.W)

        # Оставляем toolbar для будущих быстрых действий (пока пустой)
        # TODO: Добавить кнопки быстрого доступа к основным функциям

        # Выбор конфигурации
        config_frame = ttk.LabelFrame(main_frame, text="Выбор поставщика", padding="10")
        config_frame.grid(
            row=2, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10)
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

        # Область вывода информации
        info_frame = ttk.LabelFrame(main_frame, text="Информация о файле", padding="10")
        info_frame.grid(
            row=3, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(0, 10)
        )
        info_frame.columnconfigure(0, weight=1)
        info_frame.rowconfigure(0, weight=1)

        self.info_text = scrolledtext.ScrolledText(info_frame, width=80, height=15)
        self.info_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Применяем загруженный размер шрифта
        self.apply_font_size(self.current_font_size)

        # Кнопки действий
        action_frame = ttk.Frame(main_frame)
        action_frame.grid(
            row=4, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(10, 0)
        )

        self.show_data_button = ttk.Button(
            action_frame,
            text="📋 Показать данные",
            command=self.show_data_sample,
            state="disabled",
        )
        self.show_data_button.grid(row=0, column=0, sticky=tk.W)

        self.save_data_button = ttk.Button(
            action_frame,
            text="💾 Сохранить обработанный",
            command=self.save_data,
            state="disabled",
        )
        self.save_data_button.grid(row=0, column=1, padx=(10, 0))

        self.compare_button = ttk.Button(
            action_frame,
            text="🔍 Сравнить с базой",
            command=self.compare_with_base,
            state="disabled",
        )
        self.compare_button.grid(row=0, column=2, padx=(10, 0))

        # Новые кнопки после сравнения
        self.update_prices_button = ttk.Button(
            action_frame,
            text="🏷️ Обновить цены",
            command=self.update_prices,
            state="disabled",
        )
        self.update_prices_button.grid(row=0, column=3, padx=(10, 0))

        self.report_button = ttk.Button(
            action_frame,
            text="📊 Сохранить отчет",
            command=self.save_report,
            state="disabled",
        )
        self.report_button.grid(row=0, column=4, padx=(10, 0))

        self.add_to_base_button = ttk.Button(
            action_frame,
            text="📥 Добавить новый товар в базу",
            command=self.add_to_base,
            state="disabled",
        )
        self.add_to_base_button.grid(row=0, column=5, padx=(10, 0))

        # Продвинутый статус-бар
        self.create_advanced_status_bar(main_frame)

    def create_menu(self):
        """Создание главного меню приложения"""

        # Создаем главное меню
        self.menubar = tk.Menu(self.root)

        # === МЕНЮ "ФАЙЛ" ===
        file_menu = tk.Menu(self.menubar, tearoff=0)
        file_menu.add_command(
            label="📁 Открыть файл", command=self.select_file, accelerator="Ctrl+O"
        )
        file_menu.add_separator()
        file_menu.add_command(
            label="⚙️ Настройки", command=self.show_settings, accelerator="Ctrl+,"
        )
        file_menu.add_separator()
        file_menu.add_command(
            label="🚪 Выход", command=self.quit_application, accelerator="Ctrl+Q"
        )
        self.menubar.add_cascade(label="📁 Файл", menu=file_menu)

        # === МЕНЮ "ПРАВКА" ===
        edit_menu = tk.Menu(self.menubar, tearoff=0)
        edit_menu.add_command(
            label="✂️ Вырезать", command=self.cut_text, accelerator="Ctrl+X"
        )
        edit_menu.add_command(
            label="📋 Копировать", command=self.copy_text, accelerator="Ctrl+C"
        )
        edit_menu.add_separator()
        edit_menu.add_command(
            label="🔘 Выделить все", command=self.select_all_text, accelerator="Ctrl+A"
        )
        edit_menu.add_command(
            label="🔄 Инвертировать выделенное",
            command=self.invert_selection,
            accelerator="Ctrl+I",
        )
        self.menubar.add_cascade(label="✏️ Правка", menu=edit_menu)

        # === МЕНЮ "ВИД" ===
        view_menu = tk.Menu(self.menubar, tearoff=0)
        view_menu.add_command(
            label="🧹 Очистить", command=self.clear_info, accelerator="Ctrl+L"
        )
        view_menu.add_command(
            label="🔄 Обновить", command=self.refresh_interface, accelerator="F5"
        )
        view_menu.add_separator()

        # Подменю размеров шрифта
        font_menu = tk.Menu(view_menu, tearoff=0)
        font_menu.add_command(
            label="📝 Обычный шрифт", command=lambda: self.change_font_size("normal")
        )
        font_menu.add_command(
            label="📄 Средний шрифт", command=lambda: self.change_font_size("medium")
        )
        font_menu.add_command(
            label="📊 Крупный шрифт", command=lambda: self.change_font_size("large")
        )
        view_menu.add_cascade(label="🔤 Размер шрифта", menu=font_menu)

        self.menubar.add_cascade(label="👁️ Вид", menu=view_menu)

        # === МЕНЮ "СПРАВКА" ===
        help_menu = tk.Menu(self.menubar, tearoff=0)
        help_menu.add_command(
            label="📖 Помощь", command=self.show_help, accelerator="F1"
        )
        help_menu.add_separator()
        help_menu.add_command(
            label="ℹ️ О программе", command=self.show_about, accelerator="Ctrl+F1"
        )
        self.menubar.add_cascade(label="❓ Справка", menu=help_menu)

        # Привязываем меню к окну
        self.root.config(menu=self.menubar)

        # Горячие клавиши
        self.setup_hotkeys()

    def setup_hotkeys(self):
        """Настройка горячих клавиш"""
        # Файл
        self.root.bind("<Control-o>", lambda e: self.select_file())
        self.root.bind("<Control-comma>", lambda e: self.show_settings())
        self.root.bind("<Control-q>", lambda e: self.quit_application())

        # Правка
        self.root.bind("<Control-x>", lambda e: self.cut_text())
        self.root.bind("<Control-c>", lambda e: self.copy_text())
        self.root.bind("<Control-a>", lambda e: self.select_all_text())
        self.root.bind("<Control-i>", lambda e: self.invert_selection())

        # Вид
        self.root.bind("<Control-l>", lambda e: self.clear_info())
        self.root.bind("<F5>", lambda e: self.refresh_interface())

        # Справка
        self.root.bind("<F1>", lambda e: self.show_help())
        self.root.bind("<Control-F1>", lambda e: self.show_about())

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
            # Запускаем красивый прогресс-бар для загрузки
            self.start_progress("Загрузка файла", 5, "file")

            # Шаг 1: Подготовка
            self.update_progress(1, "Подготовка к загрузке")
            from excel_loader.loader import load_with_config

            # Шаг 2: Загрузка Excel файла
            self.update_progress(2, "Чтение Excel файла")
            df = load_with_config(file_path, config_name)

            if df is not None:
                # Шаг 3: Обработка данных
                self.update_progress(3, "Обработка данных")
                self.current_df = df
                self.current_config = config_name

                # Шаг 4: Отображение информации
                self.update_progress(4, "Подготовка отображения")
                self.show_file_info(df, config_name)

                # Шаг 5: Финализация
                self.update_progress(5, "Завершение загрузки")

                # Обновляем состояние кнопок
                self.update_buttons_state()

                # Завершаем с красивым сообщением
                rows = len(df)
                cols = len(df.columns)
                size_mb = df.memory_usage(deep=True).sum() / 1024 / 1024
                self.finish_progress(
                    f"✅ Загружено: {rows:,} строк, {cols} столбцов ({size_mb:.1f} МБ)"
                )

                # Сбрасываем конфигурацию на "auto" для следующей загрузки
                if "auto" in self.config_combo["values"]:
                    self.config_combo.set("auto")
                    self.log_info(
                        "🔄 Конфигурация сброшена на 'auto' для следующей загрузки"
                    )
            else:
                self.finish_progress("Файл не был загружен", auto_reset=False)
                self.set_status("Файл не загружен", "error")

        except Exception as e:
            self.log_error(f"Ошибка загрузки файла: {e}")
            self.finish_progress("Ошибка загрузки файла", auto_reset=False)
            self.set_status(f"Ошибка: {str(e)}", "error")

    def load_largest(self):
        """Загрузка самого большого файла"""

        # Директория с данными
        data_dir = "data/input"

        try:
            self.set_status("Поиск самого большого файла...", "loading")
            self.root.update()

            # Находим самый большой файл
            excel_files = []
            for file in os.listdir(data_dir):
                if file.endswith((".xlsx", ".xls")):
                    file_path = os.path.join(data_dir, file)
                    file_size = os.path.getsize(file_path)
                    excel_files.append((file_path, file_size))

            if not excel_files:
                self.log_error("Excel файлы не найдены в data/input")
                self.set_status("Файлы не найдены", "warning")
                return

            # Сортируем по размеру и берем самый большой
            excel_files.sort(key=lambda x: x[1], reverse=True)
            largest_file_path, largest_size = excel_files[0]

            self.log_info(
                f"Найден самый большой файл: {os.path.basename(largest_file_path)} ({largest_size} bytes)"
            )

            # Автоматически определяем конфиг
            config_name = self.auto_select_config(largest_file_path)

            from excel_loader.loader import load_with_config

            df = load_with_config(largest_file_path, config_name)

            if df is not None:
                self.current_df = df
                self.current_config = config_name
                self.show_file_info(df, config_name)
                self.set_status("Самый большой файл загружен", "success")

                # Обновляем состояние кнопок
                self.update_buttons_state()
            else:
                self.set_status("Файл не загружен", "error")

        except Exception as e:
            self.log_error(f"Ошибка загрузки самого большого файла: {e}")
            self.set_status("Ошибка загрузки", "error")

    def show_file_info(self, df, config_name):
        """Показ информации о загруженном файле"""
        self.log_info(f"📊 Отображение информации о файле (конфиг: {config_name})")
        # Очищаем только текстовое поле, НЕ сбрасывая данные
        self.info_text.delete(1.0, tk.END)

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

        # Очищаем только текстовое поле, НЕ сбрасывая данные
        self.info_text.delete(1.0, tk.END)

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
                self.set_status("Предобработка данных...", "save")

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
                self.set_status("Готов к работе", "info")

            except Exception as e:
                self.log_error(f"Ошибка сохранения: {e}")
                messagebox.showerror("Ошибка", f"Не удалось сохранить файл: {e}")
                self.set_status("Ошибка сохранения", "error")

    def compare_with_base(self):
        """Сравнение текущего файла с базой данных"""
        self.log_info("🔍 Начало сравнения с базой данных...")

        if self.current_df is None:
            self.log_error("Файл поставщика не загружен")
            messagebox.showwarning(
                "Предупреждение", "Сначала загрузите файл поставщика"
            )
            return

        # НОВАЯ ЛОГИКА: проверяем настройку автозагрузки
        if self.auto_load_base_enabled:
            self.set_status("Автозагрузка базы данных...", "loading")
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
            # Если автозагрузка выключена, база должна быть загружена вручную
            if self.base_df is None:
                messagebox.showwarning(
                    "Предупреждение",
                    "Сначала загрузите базу данных или включите автозагрузку в настройках",
                )
                return

        # НОВОЕ: Предобработка данных поставщика
        self.set_status("Предобработка данных поставщика...", "loading")
        self.root.update()

        processed_supplier_df = self.preprocess_supplier_data(
            self.current_df, self.current_config
        )

        # Выполняем сравнение с предобработанными данными
        self.set_status("Сравнение с базой...", "compare")
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
            self.log_info("🔘 Активируем кнопки после успешного сравнения...")
            self.update_buttons_state()

        self.set_status("Сравнение завершено", "success")

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

        # Очищаем только текстовое поле, НЕ сбрасывая данные
        self.info_text.delete(1.0, tk.END)

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

        # 2. Фильтрация по балансу - оставляем только товары в наличии И на распродаже
        if "balance" in processed_df.columns:
            self.log_info(
                f"📦 Фильтруем по балансу (оставляем только {VITYA_BALANCE_AVAILABLE})..."
            )

            balance_before = len(processed_df)
            # Новая логика: фильтруем по списку значений
            processed_df = processed_df[
                processed_df["balance"].isin(VITYA_BALANCE_AVAILABLE)
            ].copy()
            balance_after = len(processed_df)

            removed_balance = balance_before - balance_after
            if removed_balance > 0:
                self.log_info(f"   📦 Удалено строк без наличия: {removed_balance}")
                self.log_info(f"   📦 Осталось строк в наличии: {balance_after}")

                # Показываем статистику по каждому типу баланса
                for status in VITYA_BALANCE_AVAILABLE:
                    status_count = (processed_df["balance"] == status).sum()
                    if status_count > 0:
                        self.log_info(f"      '{status}': {status_count} товаров")
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
        """Очистка области информации и сброс состояния"""
        self.info_text.delete(1.0, tk.END)

        # Сбрасываем данные
        self.current_df = None
        self.base_df = None
        self.comparison_result = None
        self.current_config = None

        # Обновляем состояние кнопок
        self.update_buttons_state()

        # Сбрасываем статус
        self.set_status("Готов к работе", "info")

        self.log_info("🧹 Интерфейс очищен, все данные сброшены")

    def update_buttons_state(self, log_changes=True):
        """Обновление состояния кнопок в зависимости от загруженных данных"""
        # Кнопки, которые зависят от загруженного файла поставщика
        file_loaded = self.current_df is not None
        file_state = "normal" if file_loaded else "disabled"

        self.show_data_button.config(state=file_state)
        self.save_data_button.config(state=file_state)
        self.compare_button.config(state=file_state)
        self.update_prices_button.config(state=file_state)

        # Кнопки, которые зависят от выполненного сравнения
        comparison_done = self.comparison_result is not None
        comparison_state = "normal" if comparison_done else "disabled"

        self.report_button.config(state=comparison_state)

        # Кнопка "Добавить новый товар в базу" активна только если есть новые товары
        has_new_items = False
        new_items_count = 0
        if self.comparison_result is not None:
            new_items = self.comparison_result.get("new_items", [])
            new_items_count = len(new_items)
            has_new_items = new_items_count > 0

        add_to_base_state = "normal" if has_new_items else "disabled"
        self.add_to_base_button.config(state=add_to_base_state)

        # Логирование изменений (опционально)
        if log_changes:
            if file_loaded:
                self.log_info("✅ Файл загружен - основные кнопки активны")
            if comparison_done:
                self.log_info("✅ Сравнение выполнено - кнопки отчетов активны")
            if has_new_items:
                self.log_info(
                    f"📥 Обнаружено новых товаров: {new_items_count} - кнопка добавления активна"
                )
            elif comparison_done and not has_new_items:
                self.log_info(
                    "ℹ️ Новых товаров не найдено - кнопка добавления неактивна"
                )
            if not file_loaded and not comparison_done:
                self.log_info("⚪ Данные отсутствуют - кнопки деактивированы")

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
                self.set_status("Сохранение отчета...", "save")
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

                    # Настраиваем ширину столбцов для Сводки
                    worksheet = writer.sheets["Сводка"]
                    worksheet.column_dimensions["A"].width = 20  # Поставщик
                    worksheet.column_dimensions["B"].width = 12  # Товаров
                    worksheet.column_dimensions["C"].width = 15  # Совпадений
                    worksheet.column_dimensions["D"].width = 18  # Процент совпадений
                    worksheet.column_dimensions["E"].width = 15  # Изменений цен
                    worksheet.column_dimensions["F"].width = 15  # Новых товаров
                    worksheet.column_dimensions["G"].width = 20  # Совпадений по кодам

                    # Лист с совпадениями
                    if self.comparison_result["matches"]:
                        self.log_info(
                            f"📄 Создаем лист 'Совпадения' ({len(self.comparison_result['matches'])} записей)..."
                        )
                        matches_df = pd.DataFrame(self.comparison_result["matches"])
                        matches_df.to_excel(
                            writer, sheet_name="Совпадения", index=False
                        )

                        # Настраиваем ширину столбцов для Совпадений
                        worksheet = writer.sheets["Совпадения"]
                        # Ищем столбец с name и устанавливаем ширину 110
                        if "name" in matches_df.columns:
                            name_col_index = matches_df.columns.get_loc("name")
                            name_col_letter = chr(
                                65 + name_col_index
                            )  # A=65, B=66, C=67...
                            worksheet.column_dimensions[name_col_letter].width = 110

                        # Устанавливаем стандартную ширину для остальных столбцов
                        for i, col in enumerate(matches_df.columns):
                            col_letter = chr(65 + i)
                            if col != "name":  # name уже настроен выше
                                if "article" in col.lower():
                                    worksheet.column_dimensions[col_letter].width = 15
                                elif "price" in col.lower() or "diff" in col.lower():
                                    worksheet.column_dimensions[col_letter].width = 15
                                elif "color" in col.lower():
                                    worksheet.column_dimensions[col_letter].width = 20
                                else:
                                    worksheet.column_dimensions[col_letter].width = 18

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

                        # Настраиваем ширину столбцов для Изменений цен
                        worksheet = writer.sheets["Изменения цен"]
                        # Ищем столбец с name и устанавливаем ширину 110
                        if "name" in price_changes_df.columns:
                            name_col_index = price_changes_df.columns.get_loc("name")
                            name_col_letter = chr(65 + name_col_index)
                            worksheet.column_dimensions[name_col_letter].width = 110

                        # Устанавливаем стандартную ширину для остальных столбцов
                        for i, col in enumerate(price_changes_df.columns):
                            col_letter = chr(65 + i)
                            if col != "name":
                                if "article" in col.lower():
                                    worksheet.column_dimensions[col_letter].width = 15
                                elif (
                                    "price" in col.lower()
                                    or "diff" in col.lower()
                                    or "percent" in col.lower()
                                ):
                                    worksheet.column_dimensions[col_letter].width = 15
                                else:
                                    worksheet.column_dimensions[col_letter].width = 18

                    # Лист с новыми товарами
                    if self.comparison_result["new_items"]:
                        self.log_info(
                            f"📄 Создаем лист 'Новые товары' ({len(self.comparison_result['new_items'])} записей)..."
                        )
                        new_items_df = pd.DataFrame(self.comparison_result["new_items"])
                        new_items_df.to_excel(
                            writer, sheet_name="Новые товары", index=False
                        )

                        # Настраиваем ширину столбцов для Новых товаров
                        worksheet = writer.sheets["Новые товары"]
                        # Ищем столбец с name и устанавливаем ширину 110
                        if "name" in new_items_df.columns:
                            name_col_index = new_items_df.columns.get_loc("name")
                            name_col_letter = chr(65 + name_col_index)
                            worksheet.column_dimensions[name_col_letter].width = 110

                        # Устанавливаем стандартную ширину для остальных столбцов
                        for i, col in enumerate(new_items_df.columns):
                            col_letter = chr(65 + i)
                            if col != "name":
                                if "article" in col.lower():
                                    worksheet.column_dimensions[col_letter].width = 15
                                elif "price" in col.lower():
                                    worksheet.column_dimensions[col_letter].width = 15
                                elif "color" in col.lower() or "balance" in col.lower():
                                    worksheet.column_dimensions[col_letter].width = 20
                                else:
                                    worksheet.column_dimensions[col_letter].width = 18

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

                        # Настраиваем ширину столбцов для Совпадений по кодам
                        worksheet = writer.sheets["Совпадения по кодам"]
                        # Ищем столбцы с name и устанавливаем ширину 110
                        for col_name in ["name", "supplier_name", "base_name"]:
                            if col_name in code_matches_df.columns:
                                name_col_index = code_matches_df.columns.get_loc(
                                    col_name
                                )
                                name_col_letter = chr(65 + name_col_index)
                                worksheet.column_dimensions[name_col_letter].width = 110

                        # Устанавливаем стандартную ширину для остальных столбцов
                        for i, col in enumerate(code_matches_df.columns):
                            col_letter = chr(65 + i)
                            if col not in ["name", "supplier_name", "base_name"]:
                                if "article" in col.lower() or "code" in col.lower():
                                    worksheet.column_dimensions[col_letter].width = 15
                                elif "confidence" in col.lower():
                                    worksheet.column_dimensions[col_letter].width = 15
                                else:
                                    worksheet.column_dimensions[col_letter].width = 18

                    # Лист с предупреждениями (значительные изменения цен)
                    warnings_data = []

                    # Добавляем значительные изменения цен как предупреждения
                    for change in self.comparison_result.get("price_changes", []):
                        if (
                            abs(change.get("price_change_percent", 0))
                            > SIGNIFICANT_CHANGE_PERCENT
                        ):
                            warnings_data.append(
                                {
                                    "Тип предупреждения": "Значительное изменение цены",
                                    "Артикул": change.get("article", ""),
                                    "Наименование": change.get("name", ""),
                                    "Цена базы": change.get("base_price", 0),
                                    "Цена поставщика": change.get("supplier_price", 0),
                                    "Изменение %": f"{change.get('price_change_percent', 0):+.1f}%",
                                    "Разница": change.get("price_diff", 0),
                                    "Описание": f"Изменение цены превышает {SIGNIFICANT_CHANGE_PERCENT}%",
                                }
                            )

                    # Добавляем предупреждения о товарах без цены в базе
                    for match in self.comparison_result.get("matches", []):
                        if (
                            match.get("base_price", 0) <= 0
                            and match.get("supplier_price", 0) > 0
                        ):
                            warnings_data.append(
                                {
                                    "Тип предупреждения": "Отсутствует цена в базе",
                                    "Артикул": match.get("article", ""),
                                    "Наименование": match.get("name", ""),
                                    "Цена базы": match.get("base_price", 0),
                                    "Цена поставщика": match.get("supplier_price", 0),
                                    "Изменение %": "Новая цена",
                                    "Разница": match.get("supplier_price", 0),
                                    "Описание": "В базе нет цены, но есть у поставщика",
                                }
                            )

                    # Создаем лист Предупреждения если есть данные
                    if warnings_data:
                        self.log_info(
                            f"📄 Создаем лист 'Предупреждения' ({len(warnings_data)} записей)..."
                        )
                        warnings_df = pd.DataFrame(warnings_data)
                        warnings_df.to_excel(
                            writer, sheet_name="Предупреждения", index=False
                        )

                        # Настраиваем ширину столбцов для Предупреждений
                        worksheet = writer.sheets["Предупреждения"]
                        worksheet.column_dimensions["A"].width = (
                            25  # Тип предупреждения
                        )
                        worksheet.column_dimensions["B"].width = 15  # Артикул
                        worksheet.column_dimensions["C"].width = (
                            110  # Наименование (широкий)
                        )
                        worksheet.column_dimensions["D"].width = 15  # Цена базы
                        worksheet.column_dimensions["E"].width = 18  # Цена поставщика
                        worksheet.column_dimensions["F"].width = 15  # Изменение %
                        worksheet.column_dimensions["G"].width = 12  # Разница
                        worksheet.column_dimensions["H"].width = 40  # Описание
                    else:
                        self.log_info("ℹ️ Предупреждений для отчета не найдено")

                self.log_info("✅ Excel файл создан успешно")

                self.log_info(f"📊 Отчет сохранен: {file_path}")
                self.log_info(f"   Листов создано: {len(summary_data)} + детализация")
                messagebox.showinfo("Успех", f"Отчет сохранен в {file_path}")
                self.set_status("Отчет сохранен", "success")

            except Exception as e:
                self.log_error(f"Ошибка сохранения отчета: {e}")
                messagebox.showerror("Ошибка", f"Не удалось сохранить отчет: {e}")
                self.set_status("Ошибка сохранения отчета", "error")
        else:
            self.log_info("ℹ️ Сохранение отчета отменено пользователем")

    def update_prices(self):
        """Обновление цен в базе данных"""
        self.log_info("🔄 Начало обновления цен в базе данных...")

        # Проверяем, что данные загружены
        if self.current_df is None:
            self.log_error("❌ Файл поставщика не загружен")
            messagebox.showwarning(
                "Предупреждение", "Сначала загрузите файл поставщика"
            )
            return

        if self.base_df is None:
            self.log_info("📁 База данных не загружена, выполняем автозагрузку...")

            # Автозагрузка базы данных (такая же логика, как в compare_with_base)
            self.set_status("Автозагрузка базы данных...", "loading")
            self.root.update()

            data_dir = "data/input"
            self.base_df = load_largest_file(data_dir, "base")

            if self.base_df is None:
                self.log_error("❌ Не удалось загрузить базу данных")
                messagebox.showerror(
                    "Ошибка", "Не удалось загрузить базу данных из data/input"
                )
                return

            self.log_info("✅ База данных автоматически загружена для обновления цен")

        if self.comparison_result is None:
            self.log_info("📊 Результат сравнения отсутствует, выполняем сравнение...")

            # Автоматически выполняем сравнение
            self.set_status("Выполнение сравнения для обновления цен...", "compare")
            self.root.update()

            # Предобработка данных поставщика
            processed_supplier_df = self.preprocess_supplier_data(
                self.current_df, self.current_config
            )

            # Выполняем сравнение
            comparison_result = self.perform_comparison(
                processed_supplier_df, self.base_df
            )

            if "error" in comparison_result:
                self.log_error(
                    f"❌ Ошибка при автосравнении: {comparison_result['error']}"
                )
                messagebox.showerror(
                    "Ошибка",
                    f"Не удалось выполнить сравнение: {comparison_result['error']}",
                )
                return

            # Сохраняем результат и показываем его
            self.comparison_result = comparison_result
            self.show_comparison_result(comparison_result)
            self.update_buttons_state()

            self.log_info("✅ Сравнение автоматически выполнено для обновления цен")

        # Диалог выбора резервной копии
        backup_choice = messagebox.askyesnocancel(
            "Резервная копия",
            "Создать резервную копию базы данных перед обновлением цен?\n\n"
            "💡 Рекомендуется для безопасности данных\n\n"
            "Да - выбрать папку для backup\n"
            "Нет - обновить без backup\n"
            "Отмена - прервать операцию",
        )

        if backup_choice is None:  # Отмена
            self.log_info("❌ Обновление цен отменено пользователем")
            return

        backup_path = None
        if backup_choice:  # Пользователь выбрал "Да"
            from tkinter import filedialog

            # Предзаполненное имя файла backup
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            default_name = f"BACKUP_base_{self.current_config}_{timestamp}.xlsx"

            backup_path = filedialog.asksaveasfilename(
                title="Выберите место для сохранения резервной копии",
                defaultextension=".xlsx",
                filetypes=[("Excel files", "*.xlsx"), ("All files", "*.*")],
                initialfile=default_name,
                initialdir="data/output",
            )

            if not backup_path:  # Пользователь отменил выбор папки
                self.log_info(
                    "❌ Обновление цен отменено - не выбрана папка для backup"
                )
                return

        # Запускаем progress bar
        self.start_progress("Анализ изменений цен", 7, "update")

        # Проверяем наличие совпадений для обновления
        self.update_progress(1, "Проверка совпадений")
        matches = self.comparison_result.get("matches", [])
        if not matches:
            self.log_info("ℹ️ Нет совпадений для обновления цен")
            self.finish_progress("Нет изменений для обновления", auto_reset=True)
            messagebox.showinfo(
                "Информация", "Нет совпадений товаров для обновления цен"
            )
            return

        # Фильтруем товары, которые имеют изменения цен больше MIN_PRICE_CHANGE_PERCENT
        price_updates = []
        for match in matches:
            supplier_price = match.get("supplier_price")
            base_price = match.get("base_price")

            # Проверяем что есть цена поставщика и она отличается от базовой
            if supplier_price is not None and supplier_price > 0:
                base_price = base_price if base_price is not None else 0

                # Если цены отличаются или в базе нет цены (0)
                if supplier_price != base_price:
                    # Вычисляем процент изменения
                    if base_price > 0:
                        price_change_percent = abs(
                            (supplier_price - base_price) / base_price * 100
                        )
                    else:
                        price_change_percent = (
                            100  # Новая цена вместо 0 - всегда обновляем!
                        )

                    # Проверяем что изменение больше минимального порога
                    # Для товаров с base_price = 0 всегда обновляем
                    if (
                        price_change_percent >= MIN_PRICE_CHANGE_PERCENT
                        or base_price == 0
                    ):
                        price_updates.append(match)
                        if base_price == 0:
                            self.log_info(
                                f"📌 Добавлен {match.get('article')}: новая цена в базе (было 0 → {supplier_price})"
                            )
                    else:
                        self.log_info(
                            f"⏭️ Пропущен {match.get('article')}: изменение слишком мало ({price_change_percent:.1f}%)"
                        )

        self.update_progress(2, f"Найдено {len(price_updates)} товаров для обновления")

        if not price_updates:
            self.log_info("ℹ️ Нет изменений цен для обновления")
            self.finish_progress("Все цены актуальны", auto_reset=True)
            messagebox.showinfo(
                "Информация", "Все цены уже актуальны, обновление не требуется"
            )
            return

        self.log_info(f"📊 Найдено {len(price_updates)} товаров с изменениями цен")

        # Показываем диалог подтверждения
        backup_message = (
            "Резервная копия будет создана.\n\n"
            if backup_path
            else "Резервная копия НЕ будет создана.\n\n"
        )
        result = messagebox.askyesno(
            "Подтверждение обновления",
            f"Будет обновлено {len(price_updates)} товаров.\n\n"
            f"{backup_message}"
            "Продолжить обновление цен?",
            icon="question",
        )

        if result:
            self.log_info("✅ Пользователь подтвердил обновление цен")

            # Запускаем процесс обновления цен
            try:
                self.update_progress(3, "Создание резервной копии")
                self.root.update()

                # 1. Создаем резервную копию базы (если выбрана)
                if backup_path:
                    self.log_info("💾 Создание резервной копии базы...")
                    try:
                        import shutil

                        # Определяем путь к оригинальному файлу базы
                        base_file_path = "data/input"
                        original_path = None

                        if os.path.exists(base_file_path):
                            base_files = []
                            for file in os.listdir(base_file_path):
                                if file.endswith(
                                    (".xlsx", ".xls")
                                ) and not file.startswith("~"):
                                    full_path = os.path.join(base_file_path, file)
                                    file_size = os.path.getsize(full_path)
                                    base_files.append((full_path, file_size, file))

                            if base_files:
                                base_files.sort(key=lambda x: x[1], reverse=True)
                                original_path = base_files[0][0]

                        if original_path:
                            # Создаем папку если не существует
                            os.makedirs(os.path.dirname(backup_path), exist_ok=True)
                            shutil.copy(original_path, backup_path)
                            self.log_info(f"💾 Backup создан: {backup_path}")
                        else:
                            self.log_error("❌ Не найден файл базы для backup")

                    except Exception as backup_error:
                        self.log_error(f"❌ Ошибка создания backup: {backup_error}")
                        messagebox.showerror(
                            "Ошибка",
                            f"Не удалось создать резервную копию: {backup_error}",
                        )
                        self.finish_progress("Ошибка создания backup", auto_reset=True)
                        return
                else:
                    self.log_info(
                        "ℹ️ Резервная копия не создается (выбрано пользователем)"
                    )

                # 2. Применяем обновления цен с проверками
                self.update_progress(4, "Применение обновлений в памяти")
                self.log_info("🔄 Применение обновлений цен...")
                updates_applied = 0
                updates_skipped = 0
                warnings = []

                # Определяем столбцы для обновления
                if self.current_config == "vitya":
                    base_price_col = "price_vitya_usd"
                    article_col = "article_vitya"
                elif self.current_config == "dimi":
                    base_price_col = "price_dimi_usd"
                    article_col = "article_dimi"
                else:
                    base_price_col = "price"
                    article_col = "article"

                # Проверяем что столбец существует в базе
                if base_price_col not in self.base_df.columns:
                    self.log_error(
                        f"❌ Столбец {base_price_col} не найден в базе данных"
                    )
                    messagebox.showerror(
                        "Ошибка", f"Столбец {base_price_col} не найден в базе данных"
                    )
                    self.set_status("Ошибка обновления", "error")
                    return

                # Обрабатываем каждое обновление
                for update in price_updates:
                    article = update.get("article")
                    supplier_price = update.get("supplier_price", 0)
                    base_price = update.get("base_price", 0)

                    if not article or supplier_price <= 0:
                        continue

                    # Вычисляем процент изменения
                    if base_price > 0:
                        price_change_percent = abs(
                            (supplier_price - base_price) / base_price * 100
                        )
                    else:
                        price_change_percent = 100  # Новая цена вместо 0

                    # Проверяем пороги безопасности
                    if price_change_percent < MIN_PRICE_CHANGE_PERCENT:
                        updates_skipped += 1
                        self.log_info(
                            f"⏭️ Пропущено {article}: изменение слишком мало ({price_change_percent:.1f}%)"
                        )
                        continue

                    if price_change_percent > MAX_PRICE_CHANGE_PERCENT:
                        warnings.append(
                            {
                                "article": article,
                                "old_price": base_price,
                                "new_price": supplier_price,
                                "change_percent": price_change_percent,
                                "reason": f"Большое изменение ({price_change_percent:.1f}%)",
                            }
                        )
                        updates_skipped += 1
                        self.log_info(
                            f"⚠️ Пропущено {article}: изменение слишком большое ({price_change_percent:.1f}%)"
                        )
                        continue

                    # Находим строку в базе для обновления
                    try:
                        if self.current_config == "vitya":
                            # Для Вити ищем по int значению
                            base_matches = self.base_df[
                                self.base_df[article_col] == int(article)
                            ]
                        else:
                            # Для остальных ищем по строке
                            base_matches = self.base_df[
                                self.base_df[article_col].astype(str).str.strip()
                                == str(article).strip()
                            ]

                        if len(base_matches) > 0:
                            # Обновляем цену в первой найденной строке
                            base_idx = base_matches.index[0]
                            old_price = self.base_df.loc[base_idx, base_price_col]
                            self.base_df.loc[base_idx, base_price_col] = supplier_price
                            updates_applied += 1

                            self.log_info(
                                f"💰 Обновлено {article}: {old_price} → {supplier_price} ({price_change_percent:+.1f}%)"
                            )
                        else:
                            self.log_info(
                                f"❓ Артикул {article} не найден в базе для обновления"
                            )
                            updates_skipped += 1

                    except Exception as e:
                        self.log_error(f"❌ Ошибка обновления {article}: {e}")
                        updates_skipped += 1

                # 3. Показываем результаты
                self.update_progress(5, "Подготовка отчета результатов")
                self.log_info("✅ Обновление цен завершено")
                self.log_info(f"   💰 Цен обновлено: {updates_applied}")
                self.log_info(f"   ⏭️ Пропущено: {updates_skipped}")
                self.log_info(f"   ⚠️ Предупреждений: {len(warnings)}")

                # 4. Показываем диалог результатов
                result_message = f"Обновление цен завершено!\n\n"
                result_message += f"✅ Обновлено цен: {updates_applied}\n"
                result_message += f"⏭️ Пропущено: {updates_skipped}\n"
                result_message += f"⚠️ Предупреждений: {len(warnings)}\n\n"
                if backup_path:
                    result_message += (
                        f"💾 Резервная копия создана: {os.path.basename(backup_path)}\n"
                    )
                else:
                    result_message += f"ℹ️ Резервная копия не создавалась\n"
                result_message += f"🔄 База данных обновлена"

                if warnings:
                    result_message += (
                        f"\n\n⚠️ Внимание: {len(warnings)} товаров требуют проверки"
                    )

                messagebox.showinfo("Обновление завершено", result_message)

                # 5. СОХРАНЯЕМ ИЗМЕНЕНИЯ В EXCEL ФАЙЛ с сохранением форматирования
                if updates_applied > 0:
                    self.update_progress(6, "Сохранение в Excel файл")
                    self.log_info("💾 Сохранение изменений в Excel файл...")

                    # Определяем путь к оригинальному файлу базы
                    base_file_path = "data/input"
                    original_path = None

                    # Ищем файл базы (самый большой .xlsx файл)
                    if os.path.exists(base_file_path):
                        base_files = []
                        for file in os.listdir(base_file_path):
                            if file.endswith((".xlsx", ".xls")) and not file.startswith(
                                "~"
                            ):
                                full_path = os.path.join(base_file_path, file)
                                file_size = os.path.getsize(full_path)
                                base_files.append((full_path, file_size, file))

                        if base_files:
                            # Берем самый большой файл (это должна быть база)
                            base_files.sort(key=lambda x: x[1], reverse=True)
                            original_path = base_files[0][0]

                    if original_path:
                        # Создаем отдельный backup для Excel функции (всегда)
                        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                        excel_backup_filename = (
                            f"EXCEL_backup_{self.current_config}_{timestamp}.xlsx"
                        )
                        excel_backup_path = os.path.join(
                            "data/output", excel_backup_filename
                        )

                        # Применяем точечное обновление Excel (всегда!)
                        success = self.update_excel_prices_preserve_formatting(
                            original_path,
                            excel_backup_path,
                            price_updates,
                            self.current_config,
                        )

                        if success:
                            self.log_info(
                                "✅ Excel файл успешно обновлен с сохранением форматирования"
                            )
                        else:
                            self.log_error("❌ Ошибка обновления Excel файла")
                    else:
                        self.log_error("❌ Не найден файл базы для обновления")

                # 6. Завершаем операцию
                self.update_progress(7, "Завершение операции")
                self.finish_progress("Цены успешно обновлены!", auto_reset=True)
                self.update_buttons_state()

            except Exception as e:
                self.log_error(f"❌ Ошибка при обновлении цен: {e}")
                self.finish_progress("Ошибка обновления цен", auto_reset=True)
                messagebox.showerror("Ошибка", f"Ошибка при обновлении цен: {e}")
        else:
            self.log_info("❌ Пользователь отменил обновление цен")
            self.finish_progress("Обновление отменено", auto_reset=True)

    def refresh_interface(self):
        """Обновление интерфейса"""
        self.log_info("🔄 Обновление интерфейса...")

        # Обновляем список доступных конфигураций
        self.load_available_configs()

        # Обновляем статус
        self.set_status("Интерфейс обновлён", "success")
        self.root.update()

        self.log_info("✅ Интерфейс обновлён")

    def show_help(self):
        """Показать справку по использованию"""
        help_text = """🚀 MiStockSync - Справка

📁 Загрузка файлов:
• Файлы Вити должны содержать 'JHT' в названии
• Файлы Димы должны содержать 'DiMi' в названии  
• База данных должна содержать 'BASE' в названии

🔍 Процесс работы:
1. Выберите файл поставщика (или поставьте галочку 'авто')
2. Нажмите 'Сравнить с базой' для анализа
3. Используйте 'Сохранить отчет' для Excel отчёта
4. 'Обновить цены' для применения изменений

⚙️ Фильтрация:
• Витя: только товары "Имеются в нал."
• Дима: исключает товары "Ожидается"
• Цены: исключает NaN, пустые и нулевые

📊 Папки:
• data/input - исходные файлы
• data/output - результаты работы
• logs/ - файлы логов"""

        messagebox.showinfo("Справка", help_text)
        self.log_info("❓ Показана справка пользователю")

    def create_advanced_status_bar(self, main_frame):
        """Создание продвинутого многосекционного статус-бара"""
        # Основной фрейм статус-бара
        self.status_frame = ttk.Frame(main_frame, relief=tk.SUNKEN, borderwidth=1)
        self.status_frame.grid(
            row=5, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(5, 0)
        )
        self.status_frame.columnconfigure(0, weight=1)

        # Внутренний фрейм для компонентов
        inner_frame = ttk.Frame(self.status_frame)
        inner_frame.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=5, pady=2)
        inner_frame.columnconfigure(0, weight=1)

        # Левая часть - основной статус с иконкой
        self.status_main = tk.StringVar(value="🚀 Готов к работе")
        self.status_label = ttk.Label(
            inner_frame, textvariable=self.status_main, anchor=tk.W
        )
        self.status_label.grid(row=0, column=0, sticky=(tk.W, tk.E), padx=(0, 10))

        # Центр - прогресс-бар (скрыт по умолчанию)
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            inner_frame, variable=self.progress_var, length=200, mode="determinate"
        )
        self.progress_bar.grid(row=0, column=1, padx=(10, 10))
        self.progress_bar.grid_remove()  # Изначально скрыт

        # Правая часть - дополнительная информация
        self.status_info = tk.StringVar(value="")
        self.info_label = ttk.Label(
            inner_frame, textvariable=self.status_info, anchor=tk.E
        )
        self.info_label.grid(row=0, column=2, padx=(10, 0))

        # Инициализация переменных для прогресса ⓘ
        self.is_progress_visible = False
        self.current_operation = None

    def set_status(self, message, status_type="info", show_time=True):
        """Установка красивого статуса с иконками и цветами"""
        icons = {
            "loading": "⏳",
            "success": "✅",
            "error": "❌",
            "warning": "⚠️",
            "info": "🚀",
            "file": "📁",
            "save": "💾",
            "compare": "🔍",
            "update": "🏷️",
            "report": "📊",
            "backup": "🛡️",
        }

        # Цвета для разных типов статусов
        colors = {
            "loading": "#9932CC",  # Orange
            "success": "#228B22",  # ForestGreen
            "error": "#DC143C",  # Crimson
            "warning": "#FFD700",  # Gold
            "info": "#4169E1",  # RoyalBlue
            "file": "#9932CC",  # DarkOrchid
            "save": "#9932CC",  # LimeGreen #32CD32
            "compare": "#9932CC",  # DodgerBlue #1E90FF
            "update": "#9932CC",  # BlueViolet
            "report": "#20B2AA",  # LightSeaGreen
            "backup": "#CD853F",  # Peru
        }

        icon = icons.get(status_type, "🚀")
        color = colors.get(status_type, "#000000")

        formatted_message = f"{icon} {message}"
        self.status_main.set(formatted_message)
        self.status_label.config(foreground=color)

        # Добавляем время если нужно
        if show_time:
            current_time = datetime.now().strftime("%H:%M:%S")
            self.status_info.set(f"🕐 {current_time}")

        # Принудительное обновление GUI
        self.root.update_idletasks()

    def start_progress(self, message, total_steps, operation_type="loading"):
        """Запуск прогресс-бара для длительной операции"""
        self.current_operation = {
            "message": message,
            "total": total_steps,
            "current": 0,
            "type": operation_type,
        }

        # Настраиваем прогресс-бар
        self.progress_var.set(0)
        self.progress_bar.config(maximum=total_steps)

        # Показываем прогресс-бар
        self.progress_bar.grid()
        self.is_progress_visible = True

        # Устанавливаем статус
        self.set_status(f"{message} (0/{total_steps})", operation_type, show_time=True)

        self.root.update_idletasks()

    def update_progress(self, step, message=None):
        """Обновление прогресс-бара"""
        if not self.is_progress_visible or not self.current_operation:
            return

        self.current_operation["current"] = step
        self.progress_var.set(step)

        # Обновляем сообщение
        if message:
            display_message = message
        else:
            display_message = self.current_operation["message"]

        total = self.current_operation["total"]
        operation_type = self.current_operation["type"]

        # Вычисляем процент
        percent = int((step / total) * 100) if total > 0 else 0

        self.set_status(
            f"{display_message} ({step}/{total}) - {percent}%",
            operation_type,
            show_time=True,
        )

        self.root.update_idletasks()

    def finish_progress(self, success_message="Операция завершена", auto_reset=True):
        """Завершение прогресс-бара"""
        if not self.is_progress_visible:
            return

        # Скрываем прогресс-бар
        self.progress_bar.grid_remove()
        self.is_progress_visible = False

        # Показываем финальное сообщение
        self.set_status(success_message, "success", show_time=True)

        # Автосброс через 3 секунды
        if auto_reset:
            self.root.after(3000, lambda: self.set_status("Готов к работе", "info"))

        self.current_operation = None
        self.root.update_idletasks()

    def set_temp_status(self, message, status_type="info", duration=2000):
        """Временный статус с автосбросом"""
        old_status = self.status_main.get()
        old_color = self.status_label.cget("foreground")

        self.set_status(message, status_type)

        # Автоматический сброс
        def reset_status():
            self.status_main.set(old_status)
            self.status_label.config(foreground=old_color)

        self.root.after(duration, reset_status)

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

    def show_settings(self):
        """Показать окно настроек с автозагрузкой базы"""
        self.log_info("⚙️ Открытие окна настроек...")

        # Создаем окно настроек
        settings_window = tk.Toplevel(self.root)
        settings_window.title("Настройки MiStockSync")
        settings_window.resizable(False, False)

        # Устанавливаем иконку для окна
        self.set_window_icon(settings_window)

        # Центрируем окно относительно главного окна
        window_width = 450
        window_height = 420  # Увеличено с 350 из-за добавления настроек шрифта
        self.center_window(settings_window, window_width, window_height)

        # Делаем окно модальным
        settings_window.transient(self.root)
        settings_window.grab_set()

        # Заголовок
        ttk.Label(
            settings_window, text="⚙️ Настройки приложения", font=("Arial", 14, "bold")
        ).pack(pady=10)

        # Рамка с настройками
        settings_frame = ttk.LabelFrame(
            settings_window, text="Основные настройки", padding="10"
        )
        settings_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # Настройка автозагрузки базы
        auto_load_frame = ttk.Frame(settings_frame)
        auto_load_frame.pack(fill="x", pady=10)

        ttk.Label(
            auto_load_frame,
            text="📊 Автозагрузка базы данных:",
            font=("Arial", 10, "bold"),
        ).pack(anchor="w")

        auto_load_var = tk.BooleanVar(value=self.auto_load_base_enabled)
        auto_load_check = ttk.Checkbutton(
            auto_load_frame,
            text="Автоматически загружать базу при сравнении",
            variable=auto_load_var,
        )
        auto_load_check.pack(anchor="w", padx=20, pady=5)

        ttk.Label(
            auto_load_frame,
            text="При включении база будет загружаться автоматически\nпри первом сравнении с поставщиком.",
            font=("Arial", 8),
            foreground="gray",
        ).pack(anchor="w", padx=20)

        # Разделитель
        ttk.Separator(settings_frame, orient="horizontal").pack(fill="x", pady=15)

        # Настройка размера шрифта
        font_frame = ttk.Frame(settings_frame)
        font_frame.pack(fill="x", pady=10)

        ttk.Label(
            font_frame,
            text="🔤 Размер шрифта по умолчанию:",
            font=("Arial", 10, "bold"),
        ).pack(anchor="w")

        font_size_var = tk.StringVar(value=self.current_font_size)

        font_options = [
            ("📝 Обычный", "normal"),
            ("📄 Средний", "medium"),
            ("📊 Крупный", "large"),
        ]

        for text, value in font_options:
            ttk.Radiobutton(
                font_frame, text=text, variable=font_size_var, value=value
            ).pack(anchor="w", padx=20, pady=2)

        ttk.Label(
            font_frame,
            text="Изменения применятся к главному окну информации.",
            font=("Arial", 8),
            foreground="gray",
        ).pack(anchor="w", padx=20, pady=(5, 0))

        # Разделитель
        ttk.Separator(settings_frame, orient="horizontal").pack(fill="x", pady=15)

        # Функции в разработке
        ttk.Label(
            settings_frame, text="🚧 В разработке:", font=("Arial", 10, "bold")
        ).pack(anchor="w")

        planned_features = [
            "• Настройка путей к файлам по умолчанию",
            "• Пороги для фильтрации цен",
            "• Параметры автосохранения",
            "• Настройки логирования",
            "• Цветовые схемы интерфейса",
        ]

        for feature in planned_features:
            ttk.Label(settings_frame, text=feature, font=("Arial", 9)).pack(
                anchor="w", padx=10
            )

        # Кнопки
        button_frame = ttk.Frame(settings_window)
        button_frame.pack(pady=10)

        def save_settings():
            """Сохранить настройки"""
            # Сохраняем автозагрузку базы
            self.auto_load_base_enabled = auto_load_var.get()
            self.settings["auto_load_base"] = auto_load_var.get()

            # Сохраняем размер шрифта
            new_font_size = font_size_var.get()
            if new_font_size != self.current_font_size:
                self.current_font_size = new_font_size
                self.settings["font_size"] = new_font_size
                # Применяем новый размер шрифта сразу
                self.apply_font_size(new_font_size)

            # Сохраняем настройки в файл
            if self.save_settings(self.settings):
                self.log_info(
                    f"💾 Настройки сохранены: автозагрузка={auto_load_var.get()}, шрифт={new_font_size}"
                )
                messagebox.showinfo("Настройки", "Настройки успешно сохранены!")
            else:
                messagebox.showerror("Ошибка", "Не удалось сохранить настройки")

            settings_window.destroy()

        def cancel_settings():
            """Отменить изменения"""
            self.log_info("↩️ Изменения настроек отменены")
            settings_window.destroy()

        ttk.Button(button_frame, text="💾 Сохранить", command=save_settings).pack(
            side="left", padx=5
        )
        ttk.Button(button_frame, text="❌ Отмена", command=cancel_settings).pack(
            side="left", padx=5
        )

    def quit_application(self):
        """Выход из приложения с подтверждением"""
        self.log_info("🚪 Запрос на выход из приложения...")

        result = messagebox.askyesno(
            "Подтверждение выхода",
            "Вы действительно хотите выйти из MiStockSync?\n\n"
            "Все несохраненные данные будут потеряны.",
            icon="question",
        )

        if result:
            self.log_info("👋 Завершение работы приложения...")
            self.logger.info("📋 Приложение закрыто пользователем")
            self.root.quit()
        else:
            self.log_info("↩️ Выход отменен пользователем")

    def show_about(self):
        """Показать информацию о программе"""
        self.log_info("ℹ️ Показ информации о программе...")

        # Создаем отдельное окно вместо простого messagebox
        about_window = tk.Toplevel(self.root)
        about_window.title("О программе")
        about_window.resizable(False, False)

        # Устанавливаем иконку для окна
        self.set_window_icon(about_window)

        # Центрируем окно относительно главного окна
        window_width = 320  # Увеличено с 300 из-за длинного текста
        window_height = 350  # Увеличено с 240 из-за дополнительного текста
        self.center_window(about_window, window_width, window_height)

        # Делаем окно модальным
        about_window.transient(self.root)
        about_window.grab_set()

        # Главный фрейм
        main_frame = ttk.Frame(about_window, padding="20")
        main_frame.pack(fill="both", expand=True)

        # Большая иконка приложения (эмодзи)
        ttk.Label(main_frame, text="🚀", font=("Arial", 48)).pack()

        # Название и версия
        ttk.Label(
            main_frame, text="MiStockSync v0.0.9", font=("Arial", 14, "bold")
        ).pack(pady=5)

        # Дата
        ttk.Label(
            main_frame,
            text=f"📅 {datetime.now().strftime('%Y-%m-%d')}",
            font=("Arial", 9),
        ).pack()

        # Краткое описание
        ttk.Label(
            main_frame,
            text="Синхронизация прайс-листов\nс базой данных товаров\n\n• Автозагрузка базы данных\n• Настройка размера шрифта\n• Сохранение пользовательских настроек",
            font=("Arial", 9),
            justify="center",
        ).pack(pady=10)

        # Кнопка закрытия
        ttk.Button(
            main_frame, text="✅ ОК", command=about_window.destroy, width=10
        ).pack(pady=10)

        self.log_info("ℹ️ Информация о программе показана")

    # === ФУНКЦИИ МЕНЮ "ПРАВКА" ===
    def cut_text(self):
        """Вырезать выделенный текст"""
        try:
            focused_widget = self.root.focus_get()
            if hasattr(focused_widget, "selection_get"):
                text = focused_widget.selection_get()
                self.root.clipboard_clear()
                self.root.clipboard_append(text)
                focused_widget.delete("sel.first", "sel.last")
                self.log_info("✂️ Текст вырезан в буфер обмена")
        except tk.TclError:
            self.log_info("⚠️ Нет выделенного текста для вырезания")

    def copy_text(self):
        """Копировать выделенный текст"""
        try:
            focused_widget = self.root.focus_get()
            if hasattr(focused_widget, "selection_get"):
                text = focused_widget.selection_get()
                self.root.clipboard_clear()
                self.root.clipboard_append(text)
                self.log_info("📋 Текст скопирован в буфер обмена")
        except tk.TclError:
            self.log_info("⚠️ Нет выделенного текста для копирования")

    def select_all_text(self):
        """Выделить весь текст в активном поле"""
        try:
            focused_widget = self.root.focus_get()
            if focused_widget == self.info_text:
                # Для ScrolledText
                focused_widget.tag_add("sel", "1.0", "end")
                self.log_info("🔘 Весь текст выделен")
            elif hasattr(focused_widget, "select_range"):
                # Для Entry
                focused_widget.select_range(0, tk.END)
                self.log_info("🔘 Весь текст выделен")
        except:
            self.log_info("⚠️ Нет активного текстового поля")

    def invert_selection(self):
        """Инвертировать выделение (заглушка)"""
        self.log_info("🔄 Инвертирование выделения")
        messagebox.showinfo(
            "Функция в разработке",
            "Инвертирование выделения будет добавлено в следующих версиях",
        )

    # === ФУНКЦИИ РАЗМЕРА ШРИФТА ===
    def change_font_size(self, size_type):
        """Изменить размер шрифта в интерфейсе"""
        if size_type in ["normal", "medium", "large"]:
            # Применяем новый размер шрифта
            self.apply_font_size(size_type)

            # Сохраняем в настройки
            self.current_font_size = size_type
            self.settings["font_size"] = size_type
            self.save_settings(self.settings)

            size_names = {"normal": "обычный", "medium": "средний", "large": "крупный"}

            self.log_info(f"🔤 Установлен {size_names[size_type]} размер шрифта")
        else:
            self.log_info("⚠️ Неизвестный размер шрифта")

    def apply_font_size(self, size_type):
        """Применение размера шрифта к текстовому полю"""
        sizes = {
            "normal": ("Arial", 9),
            "medium": ("Arial", 11),
            "large": ("Arial", 13),
        }

        if size_type in sizes and hasattr(self, "info_text"):
            font_family, font_size = sizes[size_type]
            self.info_text.configure(font=(font_family, font_size))

    def center_window(self, window, width, height, parent=None):
        """Центрирование окна относительно родительского окна или экрана"""
        if parent is None:
            parent = self.root

        # Получаем размеры и позицию родительского окна
        parent.update_idletasks()
        parent_x = parent.winfo_x()
        parent_y = parent.winfo_y()
        parent_width = parent.winfo_width()
        parent_height = parent.winfo_height()

        # Вычисляем позицию для центрирования относительно родительского окна
        x = parent_x + (parent_width - width) // 2
        y = parent_y + (parent_height - height) // 2

        # Устанавливаем размер и позицию
        window.geometry(f"{width}x{height}+{x}+{y}")

    def set_window_icon(self, window):
        """Установка иконки для дочернего окна"""
        try:
            from PIL import Image, ImageTk

            icon = ImageTk.PhotoImage(Image.open("assets/icon.png"))
            window.iconphoto(False, icon)
        except Exception:
            # Если не удалось загрузить иконку, пропускаем
            pass

    def create_backup_base(self):
        """Создание резервной копии базы перед обновлением цен"""

        if self.base_df is None:
            self.log_error("❌ База данных не загружена для создания backup")
            return False

        try:
            # Создаем папку для backup если не существует
            backup_dir = "data/output"
            os.makedirs(backup_dir, exist_ok=True)

            # Создаем имя файла backup
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_filename = f"BACKUP_base_{self.current_config}_{timestamp}.xlsx"
            backup_path = os.path.join(backup_dir, backup_filename)

            # Сохраняем backup
            self.base_df.to_excel(backup_path, index=False)

            self.log_info(f"💾 Резервная копия создана: {backup_filename}")
            self.log_info(f"📁 Путь: {backup_path}")

            return True

        except Exception as e:
            self.log_error(f"❌ Ошибка создания резервной копии: {e}")
            return False

    def update_excel_prices_preserve_formatting(
        self, original_path, backup_path, price_updates, supplier_config
    ):
        """
        Точечное обновление цен в Excel файле с сохранением всего форматирования
        Изменяются ТОЛЬКО значения ценовых ячеек, всё остальное остается как было
        """

        self.log_info("🔧 Точечное обновление цен с сохранением форматирования...")

        try:
            # Проверяем наличие openpyxl
            if not OPENPYXL_AVAILABLE:
                self.log_error(
                    "❌ Библиотека openpyxl не установлена. Используйте: pip install openpyxl"
                )
                return False

            # 1. Создаем backup
            os.makedirs("data/output", exist_ok=True)
            shutil.copy(original_path, backup_path)
            self.log_info(f"💾 Backup создан: {os.path.basename(backup_path)}")

            # 2. Открываем Excel файл через openpyxl (сохраняет форматирование)
            workbook = load_workbook(original_path)
            worksheet = workbook.active  # Берем первый лист

            # 3. Определяем столбец для обновления цен (реальные названия в базе)
            if supplier_config == "vitya":
                price_column_name = "Цена Витя в $"
                article_column_name = "Артикул Витя"
            elif supplier_config == "dimi":
                price_column_name = "Цена Дима в $"
                article_column_name = "Артикул Дима"
            else:
                self.log_error(f"❌ Неподдерживаемая конфигурация: {supplier_config}")
                return False

            # 4. Находим индексы столбцов в Excel файле
            header_row = 1  # Предполагаем что заголовки в первой строке
            price_col_idx = None
            article_col_idx = None

            for col_idx in range(1, worksheet.max_column + 1):
                cell_value = worksheet.cell(row=header_row, column=col_idx).value
                if cell_value == price_column_name:
                    price_col_idx = col_idx
                elif cell_value == article_column_name:
                    article_col_idx = col_idx

            if not price_col_idx or not article_col_idx:
                self.log_error(
                    f"❌ Не найдены столбцы в Excel: {price_column_name}, {article_column_name}"
                )
                return False

            self.log_info(
                f"📍 Найдены столбцы: {article_column_name} (col {article_col_idx}), {price_column_name} (col {price_col_idx})"
            )

            # 5. Применяем только изменения цен
            updates_applied = 0

            for update in price_updates:
                article_to_find = str(update.get("article", "")).strip()
                new_price = update.get("supplier_price", 0)

                if not article_to_find or new_price <= 0:
                    continue

                # Ищем строку с нужным артикулом
                for row_idx in range(2, worksheet.max_row + 1):  # Начинаем с 2-й строки
                    cell_value = worksheet.cell(
                        row=row_idx, column=article_col_idx
                    ).value

                    if cell_value is not None:
                        if supplier_config == "vitya":
                            # Для Вити сравниваем как int
                            try:
                                if isinstance(cell_value, (int, float)) and int(
                                    cell_value
                                ) == int(float(article_to_find)):
                                    found_match = True
                                else:
                                    found_match = False
                            except (ValueError, TypeError):
                                found_match = False
                        else:
                            # Для Димы сравниваем как строки
                            found_match = str(cell_value).strip() == article_to_find

                        if found_match:
                            # ОБНОВЛЯЕМ ТОЛЬКО ЗНАЧЕНИЕ ЯЧЕЙКИ (форматирование сохраняется!)
                            old_value = worksheet.cell(
                                row=row_idx, column=price_col_idx
                            ).value
                            worksheet.cell(
                                row=row_idx, column=price_col_idx, value=new_price
                            )
                            updates_applied += 1

                            self.log_info(
                                f"   ✅ {article_to_find}: {old_value} → {new_price}"
                            )
                            break

            # 6. Сохраняем файл (форматирование полностью сохраняется)
            workbook.save(original_path)
            workbook.close()

            self.log_info(f"✅ Применено {updates_applied} обновлений цен")
            self.log_info(
                f"🎨 Сохранено ВСЁ форматирование: размеры ячеек, цвета, картинки и т.д."
            )

            return True

        except Exception as e:
            self.log_error(f"❌ Ошибка обновления Excel файла: {e}")
            return False


def main():
    """Главная функция приложения"""
    # Базовое логирование для main функции
    print("🚀 Запуск MiStockSync GUI...")
    print("📋 Инициализация интерфейса...")

    root = tk.Tk()

    # Настройка иконки приложения
    try:
        from PIL import Image, ImageTk

        icon = ImageTk.PhotoImage(Image.open("assets/icon.png"))
        root.iconphoto(False, icon)
        print("✅ Иконка приложения загружена")
    except Exception as e:
        print(f"⚠️ Не удалось загрузить иконку: {e}")

    # Устанавливаем заголовок
    root.title("🚀 MiStockSync - Управление прайсами")

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
