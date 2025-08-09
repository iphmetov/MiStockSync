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
        if not config_name or not isinstance(config_name, str):
            config_name = "default"

        self.config_name = config_name
        self.config = self._load_config(config_name)
        self._setup_logging()

    def get_available_configs(self) -> List[str]:
        """Получение списка доступных конфигураций"""
        try:
            configs_dir = os.path.join(os.path.dirname(__file__), "configs")
            if not os.path.exists(configs_dir):
                return ["default"]

            config_files = glob.glob(os.path.join(configs_dir, "*_config.json"))
            config_names = []
            for f in config_files:
                try:
                    config_name = os.path.basename(f).replace("_config.json", "")
                    if config_name:
                        config_names.append(config_name)
                except Exception as e:
                    self.logger.warning(f"Ошибка при обработке конфига {f}: {e}")
                    continue

            return sorted(config_names) if config_names else ["default"]
        except Exception as e:
            self.logger.error(f"Ошибка при получении списка конфигов: {e}")
            return ["default"]

    def _load_config(self, config_name: str = "default") -> dict:
        """Загрузка конфигурации по имени"""
        if not config_name or not isinstance(config_name, str):
            config_name = "default"

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
        try:
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
        except Exception as e:
            # Если логирование не настроилось, создаем простой логгер
            self.logger = logging.getLogger(f"excel_loader_{self.config_name}")
            self.logger.setLevel(logging.WARNING)
            if not self.logger.handlers:
                console_handler = logging.StreamHandler()
                console_handler.setLevel(logging.WARNING)
                formatter = logging.Formatter("%(levelname)s - %(message)s")
                console_handler.setFormatter(formatter)
                self.logger.addHandler(console_handler)

    def _get_file_info(self, file_path: str) -> dict:
        """Получение информации о файле"""
        if not file_path or not isinstance(file_path, str):
            self.logger.error("Некорректный путь к файлу")
            return {}

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
        # Проверяем, что DataFrame не пустой
        if df is None or df.empty:
            return df

        if not self.config.get("column_mapping"):
            return df

        # Создаем список названий столбцов в DataFrame в виде строк
        df_column_names = []
        for col in df.columns:
            col_str = str(col) if col is not None else ""
            df_column_names.append(col_str)

        mapping = {}
        for old_col in df.columns:
            # Безопасное преобразование в строку
            old_col_str = str(old_col) if old_col is not None else ""

            for config_key, config_value in self.config["column_mapping"].items():
                # Безопасное преобразование ключа конфигурации в строку
                config_key_str = str(config_key) if config_key is not None else ""

                if old_col_str.lower().strip() == config_key_str.lower().strip():
                    mapping[old_col] = config_value
                    break

        if mapping:
            df = df.rename(columns=mapping)
            self.logger.info(f"Применено переименование столбцов: {mapping}")

        return df

    def _remove_ignored_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Удаление игнорируемых столбцов"""
        # Проверяем, что DataFrame не пустой
        if df is None or df.empty:
            return df

        ignore_columns = self.config.get("ignore_columns", [])
        if not ignore_columns:
            return df

        # Создаем список названий столбцов в DataFrame в виде строк
        df_column_names = []
        for col in df.columns:
            col_str = str(col) if col is not None else ""
            df_column_names.append(col_str)

        columns_to_drop = []
        for col in df.columns:
            # Безопасное преобразование в строку
            col_str = str(col) if col is not None else ""

            for ignore_pattern in ignore_columns:
                # Безопасное преобразование паттерна в строку
                ignore_str = str(ignore_pattern) if ignore_pattern is not None else ""

                if ignore_str.lower() in col_str.lower():
                    columns_to_drop.append(col)
                    break

        if columns_to_drop:
            df = df.drop(columns=columns_to_drop, errors="ignore")
            self.logger.info(f"Удалены игнорируемые столбцы: {columns_to_drop}")

        return df

    def _fix_unnamed_columns(self, df: pd.DataFrame) -> pd.DataFrame:
        """Исправление Unnamed столбцов"""
        if df is None or df.empty or len(df.columns) == 0:
            return df

        new_columns = []
        for col in df.columns:
            # Безопасное преобразование в строку
            col_str = str(col) if col is not None else "Column"

            if "Unnamed" in col_str:
                try:
                    first_val = df[col].iloc[0] if len(df) > 0 else None
                    if pd.notna(first_val) and str(first_val).strip():
                        new_columns.append(str(first_val).strip())
                    else:
                        new_columns.append(f"Column_{len(new_columns) + 1}")
                except Exception as e:
                    self.logger.warning(f"Ошибка при обработке столбца {col_str}: {e}")
                    new_columns.append(f"Column_{len(new_columns) + 1}")
            else:
                new_columns.append(col_str)

        df.columns = new_columns
        return df

    def _apply_data_types(self, df: pd.DataFrame) -> pd.DataFrame:
        """Применение типов данных из конфига"""
        # Проверяем, что DataFrame не пустой
        if df is None or df.empty:
            return df

        data_types = self.config.get("data_types", {})
        if not data_types:
            return df

        # Создаем список названий столбцов в DataFrame в виде строк
        df_column_names = []
        for col in df.columns:
            col_str = str(col) if col is not None else ""
            df_column_names.append(col_str)

        for column, dtype in data_types.items():
            # Безопасное преобразование названия столбца в строку
            column_str = str(column) if column is not None else ""

            if column_str in df_column_names:
                try:
                    if dtype == "float":
                        df[column_str] = pd.to_numeric(df[column_str], errors="coerce")
                    elif dtype == "int":
                        df[column_str] = pd.to_numeric(
                            df[column_str], errors="coerce"
                        ).astype("Int64")
                    elif dtype == "string":
                        df[column_str] = df[column_str].astype(str)

                    self.logger.info(f"Применен тип {dtype} для столбца {column_str}")
                except Exception as e:
                    self.logger.warning(
                        f"Не удалось применить тип {dtype} для {column_str}: {e}"
                    )

        return df

    def _validate_data(self, df: pd.DataFrame) -> bool:
        """Валидация данных согласно конфигу"""
        # Проверяем, что DataFrame не пустой
        if df is None or df.empty:
            error_msg = "DataFrame пустой или не содержит данных"
            self.logger.error(error_msg)
            messagebox.showerror("Ошибка валидации", error_msg)
            return False

        validation = self.config.get("validation", {})

        # Проверка обязательных столбцов
        required_columns = validation.get("required_columns", [])
        missing_columns = []

        # Создаем список названий столбцов в DataFrame в виде строк
        df_column_names = []
        for col in df.columns:
            col_str = str(col) if col is not None else ""
            df_column_names.append(col_str)

        for required_col in required_columns:
            # Безопасное преобразование в строку
            required_col_str = str(required_col) if required_col is not None else ""
            if required_col_str not in df_column_names:
                missing_columns.append(required_col_str)

        if missing_columns:
            error_msg = f"Отсутствуют обязательные столбцы: {missing_columns}"
            self.logger.error(error_msg)
            messagebox.showerror("Ошибка валидации", error_msg)
            return False

        # Проверка диапазона цен
        price_columns = []
        for col in df.columns:
            # Безопасное преобразование в строку
            col_str = str(col) if col is not None else ""
            if "price" in col_str.lower():
                price_columns.append(col)

        price_min = validation.get("price_min", 0)
        price_max = validation.get("price_max", float("inf"))

        for price_col in price_columns:
            # Безопасное преобразование названия столбца в строку
            price_col_str = str(price_col) if price_col is not None else ""
            if price_col_str in df_column_names:
                invalid_prices = df[
                    (df[price_col_str] < price_min) | (df[price_col_str] > price_max)
                ]
                if not invalid_prices.empty:
                    self.logger.warning(
                        f"Найдены цены вне диапазона в столбце {price_col_str}"
                    )

        return True

    def _show_file_info(self, df: pd.DataFrame, file_path: str):
        """Вывод информации о загруженном файле"""
        # Проверяем, что DataFrame не пустой
        if df is None or df.empty:
            self.logger.warning("DataFrame пустой, нечего показывать")
            return

        file_info = self._get_file_info(file_path)

        # Безопасное преобразование названий столбцов в строки
        column_names = []
        for col in df.columns:
            col_str = str(col) if col is not None else "Unknown"
            column_names.append(col_str)

        info_text = f"""
📁 Файл: {os.path.basename(file_path)}
⚙️ Конфиг: {self.config.get('supplier_name', self.config_name)}
📊 Размер: {file_info.get('size_mb', 0)} MB
📋 Строк: {len(df)}
📋 Столбцов: {len(df.columns)}
📅 Создан: {file_info.get('created', 'Unknown')}
📅 Изменен: {file_info.get('modified', 'Unknown')}

🏷️ Названия столбцов:
{', '.join(column_names)}
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

            # Безопасное преобразование в строку
            file_path_str = str(file_path) if file_path is not None else ""
            if not file_path_str or not file_path_str.lower().endswith(
                (".xlsx", ".xls")
            ):
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

            # Проверяем, что DataFrame не пустой
            if df is None or df.empty:
                error_msg = f"Файл {file_path} пустой или не содержит данных"
                self.logger.error(error_msg)
                messagebox.showerror("Ошибка", error_msg)
                return None

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
    ) -> Optional[Tuple[pd.DataFrame, str]]:
        """
        Загрузка самого большого Excel файла из директории

        Args:
            directory_path: Путь к директории
            config_name: Имя конфига для применения

        Returns:
            Tuple[pandas.DataFrame, str] или None при ошибке
            (DataFrame, путь_к_файлу)
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
                # Безопасное преобразование в строку
                file_str = str(file) if file is not None else ""
                if file_str and file_str.lower().endswith((".xlsx", ".xls")):
                    file_path = os.path.join(directory_path, file)
                    try:
                        file_size = os.path.getsize(file_path)
                        excel_files.append((file_path, file_size))
                    except (OSError, IOError) as e:
                        self.logger.warning(
                            f"Не удалось получить размер файла {file_path}: {e}"
                        )
                        continue

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
                return df, file_path

            return None

        except Exception as e:
            error_msg = f"Ошибка загрузки самого большого файла: {str(e)}"
            self.logger.error(error_msg)
            return None


# Глобальные переменные для экземпляров загрузчиков
_loaders = {}


def get_loader(config_name: str = "default") -> ExcelLoaderEnhanced:
    """Получение экземпляра загрузчика (singleton для каждого конфига)"""
    global _loaders
    if not config_name or not isinstance(config_name, str):
        config_name = "default"

    if config_name not in _loaders:
        try:
            _loaders[config_name] = ExcelLoaderEnhanced(config_name)
        except Exception as e:
            print(f"Ошибка создания загрузчика для конфига {config_name}: {e}")
            # Пробуем создать загрузчик с default конфигом
            config_name = "default"
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
    try:
        if not config_name or not isinstance(config_name, str):
            config_name = "default"
        return get_loader(config_name).select_and_load_excel()
    except Exception as e:
        print(f"Ошибка в select_and_load_excel: {e}")
        return None


def load_largest_file(
    directory_path: str, config_name: str = "base"
) -> Optional[Tuple[pd.DataFrame, str]]:
    """
    Загрузка самого большого Excel файла из директории с указанным конфигом

    Args:
        directory_path: Путь к директории
        config_name: Имя конфигурации (default, base, vitya, dima, ...)

    Returns:
        Tuple[pandas.DataFrame, str] или None при ошибке
        (DataFrame, путь_к_файлу)
    """
    try:
        if not directory_path or not isinstance(directory_path, str):
            print("Некорректный путь к директории")
            return None

        if not config_name or not isinstance(config_name, str):
            config_name = "base"

        return get_loader(config_name).load_largest_file(directory_path)
    except Exception as e:
        print(f"Ошибка в load_largest_file: {e}")
        return None


def get_available_configs() -> List[str]:
    """Получение списка доступных конфигураций"""
    try:
        return get_loader().get_available_configs()
    except Exception as e:
        print(f"Ошибка при получении списка конфигов: {e}")
        return ["default"]


def load_with_config(file_path: str, config_name: str) -> Optional[pd.DataFrame]:
    """
    Прямая загрузка файла с указанным конфигом

    Args:
        file_path: Путь к файлу
        config_name: Имя конфигурации

    Returns:
        pandas.DataFrame или None при ошибке
    """
    try:
        if not file_path or not isinstance(file_path, str):
            print("Некорректный путь к файлу")
            return None

        if not config_name or not isinstance(config_name, str):
            config_name = "default"

        loader = get_loader(config_name)
        return loader._load_excel_file(file_path)
    except Exception as e:
        print(f"Ошибка в load_with_config: {e}")
        return None
