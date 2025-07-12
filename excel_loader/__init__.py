"""
Excel Loader Module
Модуль для загрузки и обработки Excel файлов в pandas DataFrame
"""

from .loader import select_and_load_excel, load_largest_file

__version__ = "1.0.0"
__all__ = ["select_and_load_excel", "load_largest_file"] 