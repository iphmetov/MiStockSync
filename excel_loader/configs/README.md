# Конфигурации Excel Loader

Этот каталог содержит конфигурационные файлы для различных форматов Excel файлов.

## Структура файлов

```
configs/
├── base_config.json      # Основная база данных
├── dimi_config.json      # Прайс поставщика Дима
├── vitya_config.json     # Прайс поставщика Витя
├── auto_config.json      # Автоматическая конфигурация
└── README.md             # Этот файл
```

## Описание конфигураций

### 1. **base** - Основная база данных

Конфигурация для загрузки основной базы данных товаров.

- **Файл**: `base_config.json`
- **Описание**: Содержит все товары с артикулами разных поставщиков
- **Основные столбцы**: article, name, price, article_vitya, article_dimi, price_vitya_usd, price_dimi_usd
- **Тип данных**: Смешанный (int для витя, string для дима)

### 2. **vitya** - Прайс Витя

Конфигурация для прайс-листов поставщика Витя.

- **Файл**: `vitya_config.json`
- **Описание**: Прайс-лист с артикулами в числовом формате
- **Основные столбцы**: article_vitya, name, price_usd, balance
- **Тип данных**: article_vitya = int

### 3. **dimi** - Прайс Дима

Конфигурация для прайс-листов поставщика Дима.

- **Файл**: `dimi_config.json`
- **Описание**: Прайс-лист с артикулами в строковом формате
- **Основные столбцы**: article_dimi, name, price_usd, balance
- **Тип данных**: article_dimi = string
- **Обязательные поля**: article_dimi, name, price_usd

### 4. **auto** - Автоматическая конфигурация

Универсальная конфигурация с автоматическим определением столбцов.

- **Файл**: `auto_config.json`
- **Описание**: Попытка автоматически определить структуру файла
- **Использование**: Когда неизвестен формат файла

## Использование

```python
from excel_loader import load_with_config

# Загрузка с конкретной конфигурацией
df = load_with_config('file.xlsx', 'vitya')

# Загрузка с автоматическим определением
df = load_with_config('file.xlsx', 'auto')

# Получение списка доступных конфигураций
from excel_loader import get_available_configs
configs = get_available_configs()
print(configs)  # ['base', 'default', 'dimi', 'vitya']
```

## Настройка конфигураций

Каждая конфигурация содержит:

- **column_mapping**: Сопоставление столбцов в Excel файле с внутренними именами
- **ignore_columns**: Столбцы, которые нужно игнорировать
- **data_types**: Типы данных для каждого столбца
- **validation**: Правила валидации данных
- **settings**: Дополнительные настройки загрузки

## Добавление новой конфигурации

1. Создайте новый JSON файл с именем `{name}_config.json`
2. Следуйте структуре существующих конфигураций
3. Добавьте маппинг столбцов в `column_mapping`
4. Укажите типы данных в `data_types`
5. Настройте валидацию в `validation`

Пример:

```json
{
    "supplier_name": "Новый поставщик",
    "description": "Описание конфигурации",
    "column_mapping": {
        "Артикул": "article",
        "Название": "name",
        "Цена": "price"
    },
    "data_types": {
        "article": "string",
        "price": "float"
    },
    "validation": {
        "required_columns": ["article", "name"],
        "price_min": 0,
        "price_max": 100000
    }
}
```
