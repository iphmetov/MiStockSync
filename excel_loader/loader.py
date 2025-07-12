"""
Excel Loader Module Enhanced
Функции для загрузки и обработки Excel файлов с поддержкой множественных конфигураций
"""

import os
import json
import logging
import glob
from datetime import datetime
from pathlib import Path
from typing import Optional, Tuple, Dict, Any, List
import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd


class ExcelLoaderEnhanced:
    """Расширенный класс для загрузки и обработки Excel файлов с множественными конфигами"""

    def __init__(self, config_name: str = "default"):
        """
        Инициализация загрузчика

        Args:
            config_name: Имя конфигурации (без .json)
        """
        self.config_name = config_name
        self.config = self._load_config(config_name)
        self._setup_logging()

    def get_available_configs(self) -> List[str]:
        """Получение списка доступных конфигураций"""
        configs_dir = os.path.join(os.path.dirname(__file__), "configs")
        if not os.path.exists(configs_dir):
            return ["default"]

        config_files = glob.glob(os.path.join(configs_dir, "*_config.json"))
        config_names = [
            os.path.basename(f).replace("_config.json", "") for f in config_files
        ]
        return sorted(config_names)

    def _load_config(self, config_name: str = "default") -> dict:
        """Загрузка конфигурации по имени"""
        configs_dir = os.path.join(os.path.dirname(__file__), "configs")
        config_path = os.path.join(configs_dir, f"{config_name}_config.json")

        # Если конфиг не найден, попробуем default
        if not os.path.exists(config_path):
            config_path = os.path.join(configs_dir, "default_config.json")
            print(f"Конфиг {config_name} не найден, используем default")

        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = json.load(f)
                print(f"✅ Загружен конфиг: {config.get('supplier_name', config_name)}")
                return config
        except FileNotFoundError:
            print(f"❌ Конфигурационный файл не найден: {config_path}")
            return self._get_fallback_config()
        except json.JSONDecodeError as e:
            print(f"❌ Ошибка парсинга конфигурации: {e}")
            return self._get_fallback_config()

    def _get_fallback_config(self) -> dict:
        """Резервная конфигурация если основная не загрузилась"""
        return {
            "supplier_name": "Резервная",
            "column_mapping": {},
            "ignore_columns": [],
            "settings": {"skip_empty_rows": True},
            "data_types": {},
            "validation": {"required_columns": []},
        }

    def _setup_logging(self):
        """Настройка логирования"""
        log_dir = Path("logs")
        log_dir.mkdir(exist_ok=True)

        log_file = log_dir / f"excel_loader_{datetime.now().strftime('%Y%m%d')}.log"

        self.logger = logging.getLogger(f"excel_loader_{self.config_name}")
        self.logger.setLevel(logging.INFO)

        if not self.logger.handlers:
            file_handler = logging.FileHandler(log_file, encoding="utf-8")
            file_handler.setLevel(logging.INFO)

            console_handler = logging.StreamHandler()
            console_handler.setLevel(logging.INFO)

            formatter = logging.Formatter(
                "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
            )
            file_handler.setFormatter(formatter)
            console_handler.setFormatter(formatter)

            self.logger.addHandler(file_handler)
            self.logger.addHandler(console_handler)

    def _get_file_info(self, file_path: str) -> dict:
        """Получение информации о файле"""
        try:
            stat = os.stat(file_path)
            return {
                "size": stat.st_size,
                "size_mb": round(stat.st_size / (1024 * 1024), 2),
                "created": datetime.fromtimestamp(stat.st_ctime).strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
                "modified": datetime.fromtimestamp(stat.st_mtime).strftime(
                    "%Y-%m-%d %H:%M:%S"
                ),
                "owner": stat.st_uid if hasattr(stat, "st_uid") else "Unknown",
            }
        except Exception as e:
            self.logger.error(f"Ошибка получения информации о файле: {e}")
            return {}

    def _apply_column_mapping(self, df: pd.DataFrame) -> pd.DataFrame:
        """Применение маппинга столбцов из конфига"""
        if not self.config.get("column_mapping"):
            return df

        mapping = {}
        for old_col in df.columns:
            if not isinstance(old_col, str):
                old_col = str(old_col)

            for config_key, config_value in self.config["column_mapping"].items():
                if old_col.lower().strip() == config_key.lower().strip():
                    mapping[old_col] = config_value
                    break

        if mapping:
            df = df.rename(columns=mapping)
            self.logger.info(f"Применено переименование столбцов: {mapping}")

        return df

    def _remove_ignored_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Удаление игнорируемых столбцов"""
        ignore_columns = self.config.get("ignore_columns", [])
        if not ignore_columns:
            return df

        columns_to_drop = []
        for col in df.columns:
            col_str = str(col) if col is not None else ""
            for ignore_pattern in ignore_columns:
                if ignore_pattern.lower() in col_str.lower():
                    columns_to_drop.append(col)
                    break

        if columns_to_drop:
            df = df.drop(columns=columns_to_drop, errors="ignore")
            self.logger.info(f"Удалены игнорируемые столбцы: {columns_to_drop}")

        return df

    def _fix_unnamed_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Исправление Unnamed столбцов"""
        new_columns = []
        for col in df.columns:
            col_str = str(col) if col is not None else "Column"

            if "Unnamed" in col_str:
                first_val = df[col].iloc[0] if len(df) > 0 else None
                if pd.notna(first_val) and str(first_val).strip():
                    new_columns.append(str(first_val).strip())
                else:
                    new_columns.append(f"Column_{len(new_columns) + 1}")
            else:
                new_columns.append(col_str)

        df.columns = new_columns
        return df

    def _apply_data_types(self, df: pd.DataFrame) -> pd.DataFrame:
        """Применение типов данных из конфига"""
        data_types = self.config.get("data_types", {})
        if not data_types:
            return df

        for column, dtype in data_types.items():
            if column in df.columns:
                try:
                    if dtype == "float":
                        df[column] = pd.to_numeric(df[column], errors="coerce")
                    elif dtype == "int":
                        df[column] = pd.to_numeric(df[column], errors="coerce").astype(
                            "Int64"
                        )
                    elif dtype == "string":
                        df[column] = df[column].astype(str)

                    self.logger.info(f"Применен тип {dtype} для столбца {column}")
                except Exception as e:
                    self.logger.warning(
                        f"Не удалось применить тип {dtype} для {column}: {e}"
                    )

        return df

    def _validate_data(self, df: pd.DataFrame) -> bool:
        """Валидация данных согласно конфигу"""
        validation = self.config.get("validation", {})

        # Проверка обязательных столбцов
        required_columns = validation.get("required_columns", [])
        missing_columns = [col for col in required_columns if col not in df.columns]

        if missing_columns:
            error_msg = f"Отсутствуют обязательные столбцы: {missing_columns}"
            self.logger.error(error_msg)
            messagebox.showerror("Ошибка валидации", error_msg)
            return False

        # Проверка диапазона цен
        price_columns = [col for col in df.columns if "price" in col.lower()]
        price_min = validation.get("price_min", 0)
        price_max = validation.get("price_max", float("inf"))

        for price_col in price_columns:
            if price_col in df.columns:
                invalid_prices = df[
                    (df[price_col] < price_min) | (df[price_col] > price_max)
                ]
                if not invalid_prices.empty:
                    self.logger.warning(
                        f"Найдены цены вне диапазона в столбце {price_col}"
                    )

        return True

    def _show_file_info(self, df: pd.DataFrame, file_path: str):
        """Вывод информации о загруженном файле"""
        file_info = self._get_file_info(file_path)

        info_text = f"""
📁 Файл: {os.path.basename(file_path)}
⚙️ Конфиг: {self.config.get('supplier_name', self.config_name)}
📊 Размер: {file_info.get('size_mb', 0)} MB
📋 Строк: {len(df)}
📋 Столбцов: {len(df.columns)}
📅 Создан: {file_info.get('created', 'Unknown')}
📅 Изменен: {file_info.get('modified', 'Unknown')}

🏷️ Названия столбцов:
{', '.join(df.columns.tolist())}
        """

        print(info_text)
        self.logger.info(
            f"Загружен файл: {file_path}, строк: {len(df)}, столбцов: {len(df.columns)}"
        )

    def select_and_load_excel(self, config_name: str = None) -> Optional[pd.DataFrame]:
        """
        Диалог выбора и загрузки Excel файла с выбором конфига

        Args:
            config_name: Имя конфига для применения (если None, используется текущий)

        Returns:
            pandas.DataFrame или None при ошибке
        """
        if config_name and config_name != self.config_name:
            # Переключаемся на другой конфиг
            self.config_name = config_name
            self.config = self._load_config(config_name)

        try:
            root = tk.Tk()
            root.withdraw()

            file_path = filedialog.askopenfilename(
                title=f"Выберите Excel файл (конфиг: {self.config.get('supplier_name', self.config_name)})",
                filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")],
            )

            if not file_path:
                self.logger.info("Файл не выбран")
                return None

            file_path_str = str(file_path) if file_path else ""
            if not file_path_str.lower().endswith((".xlsx", ".xls")):
                error_msg = "Выбран неподдерживаемый формат файла. Поддерживаются только .xlsx и .xls файлы."
                messagebox.showerror("Ошибка", error_msg)
                self.logger.error(error_msg)
                return None

            df = self._load_excel_file(file_path)

            if df is not None:
                self._show_file_info(df, file_path)
                messagebox.showinfo(
                    "Успех",
                    f"Файл успешно загружен с конфигом '{self.config.get('supplier_name')}'!\nСтрок: {len(df)}\nСтолбцов: {len(df.columns)}",
                )

            return df

        except Exception as e:
            error_msg = f"Неожиданная ошибка: {str(e)}"
            messagebox.showerror("Ошибка", error_msg)
            self.logger.error(error_msg)
            return None
        finally:
            try:
                root.destroy()
            except:
                pass

    def _load_excel_file(self, file_path: str) -> Optional[pd.DataFrame]:
        """Загрузка Excel файла с применением конфигурации"""
        try:
            df = pd.read_excel(file_path, sheet_name=0)

            # ОТКЛЮЧЕНО: Исправляем Unnamed столбцы
            # df = self._fix_unnamed_columns(df)

            # Удаляем игнорируемые столбцы
            df = self._remove_ignored_columns(df)

            # Применяем маппинг столбцов
            df = self._apply_column_mapping(df)

            # Применяем типы данных
            df = self._apply_data_types(df)

            # Убираем пустые строки если настроено
            if self.config.get("settings", {}).get("skip_empty_rows", True):
                df = df.dropna(how="all")

            # Валидация данных
            if not self._validate_data(df):
                return None

            return df

        except FileNotFoundError:
            error_msg = f"Файл не найден: {file_path}"
            messagebox.showerror("Ошибка", error_msg)
            self.logger.error(error_msg)
            return None
        except PermissionError:
            error_msg = f"Нет прав доступа к файлу: {file_path}"
            messagebox.showerror("Ошибка", error_msg)
            self.logger.error(error_msg)
            return None
        except Exception as e:
            error_msg = f"Ошибка загрузки файла {file_path}: {str(e)}"
            messagebox.showerror("Ошибка", error_msg)
            self.logger.error(error_msg)
            return None

    def load_largest_file(
        self, directory_path: str, config_name: str = None
    ) -> Optional[pd.DataFrame]:
        """
        Загрузка самого большого Excel файла из директории

        Args:
            directory_path: Путь к директории
            config_name: Имя конфига для применения

        Returns:
            pandas.DataFrame или None при ошибке
        """
        if config_name and config_name != self.config_name:
            self.config_name = config_name
            self.config = self._load_config(config_name)

        try:
            if not os.path.exists(directory_path):
                error_msg = f"Директория не найдена: {directory_path}"
                self.logger.error(error_msg)
                return None

            excel_files = []
            for file in os.listdir(directory_path):
                file_str = str(file) if file is not None else ""
                if file_str.lower().endswith((".xlsx", ".xls")):
                    file_path = os.path.join(directory_path, file)
                    file_size = os.path.getsize(file_path)
                    excel_files.append((file_path, file_size))

            if not excel_files:
                error_msg = f"Excel файлы не найдены в директории: {directory_path}"
                self.logger.warning(error_msg)
                return None

            largest_file = max(excel_files, key=lambda x: x[1])
            file_path, file_size = largest_file

            self.logger.info(
                f"Найден самый большой файл: {file_path} ({file_size} bytes)"
            )

            df = self._load_excel_file(file_path)

            if df is not None:
                self._show_file_info(df, file_path)

            return df

        except Exception as e:
            error_msg = f"Ошибка загрузки самого большого файла: {str(e)}"
            self.logger.error(error_msg)
            return None


# Глобальные переменные для экземпляров загрузчиков
_loaders = {}


def get_loader(config_name: str = "default") -> ExcelLoaderEnhanced:
    """Получение экземпляра загрузчика (singleton для каждого конфига)"""
    global _loaders
    if config_name not in _loaders:
        _loaders[config_name] = ExcelLoaderEnhanced(config_name)
    return _loaders[config_name]


def select_and_load_excel(config_name: str = "default") -> Optional[pd.DataFrame]:
    """
    Диалог выбора и загрузки Excel файла с указанным конфигом

    Args:
        config_name: Имя конфигурации (default, base, vitya, dima, ...)

    Returns:
        pandas.DataFrame или None при ошибке
    """
    return get_loader(config_name).select_and_load_excel()


def load_largest_file(
    directory_path: str, config_name: str = "base"
) -> Optional[pd.DataFrame]:
    """
    Загрузка самого большого Excel файла из директории с указанным конфигом

    Args:
        directory_path: Путь к директории
        config_name: Имя конфигурации (default, base, vitya, dima, ...)

    Returns:
        pandas.DataFrame или None при ошибке
    """
    return get_loader(config_name).load_largest_file(directory_path)


def get_available_configs() -> List[str]:
    """Получение списка доступных конфигураций"""
    return get_loader().get_available_configs()


def load_with_config(file_path: str, config_name: str) -> Optional[pd.DataFrame]:
    """
    Прямая загрузка файла с указанным конфигом

    Args:
        file_path: Путь к файлу
        config_name: Имя конфигурации

    Returns:
        pandas.DataFrame или None при ошибке
    """
    loader = get_loader(config_name)
    return loader._load_excel_file(file_path)
