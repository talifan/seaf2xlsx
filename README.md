# Конвертер YAML ↔ XLSX (SEAF1 и SEAF2)

Этот набор скриптов обеспечивает полную совместимость между форматами SEAF1, SEAF2 и XLSX, позволяя конвертировать данные в обоих направлениях.

## Основные файлы

### Конвертеры для SEAF2 (Множественное число)
- `_seaf2_yaml_to_xlsx_ta.py` - **[Рекомендуемый]** Конвертация YAML (SEAF2) → XLSX. Включает строгую типизацию объектов через `CLASS_NAME_MAP`.
- `_seaf2_xlsx_to_yaml.py` - Конвертация XLSX → YAML (SEAF2).
- `_seaf2_yaml_to_xlsx.py` - Базовая версия экспорта SEAF2.

### Оригинальные скрипты (для SEAF1 - Единственное число)
- `yaml_to_xlsx.py` - Конвертация YAML (SEAF1) → XLSX
- `xlsx_to_yaml.py` - Конвертация XLSX → YAML (SEAF1)

### Тестирование
- `tests/run_tests.py` - Тестирование конвертации SEAF1 (Roundtrip)
- `tests/run_tests_seaf2.py` - Тестирование конвертации SEAF2 (Roundtrip)
- `tests/run_cross_tests.py` - Тестирование совместимости (SEAF1 ↔ SEAF2)

## Ключевые различия между SEAF1 и SEAF2

| Сущность | SEAF1 | SEAF2 |
|----------|-------|-------|
| Пространство имен | `seaf.ta.*` | `seaf.company.ta.*` |
| Имена сущностей | единственное число | множественное число |
| Файлы | `office.yaml` | `dc_office.yaml` |
| Файлы | `components_network.yaml` | `network_component.yaml` |

## Использование

### 1. Конвертация SEAF2 → XLSX

```bash
# Рекомендуемый способ для папки ta
python _seaf2_yaml_to_xlsx_ta.py --config ta_to_xlsx_config.yaml
```

Пример конфигурации (`ta_to_xlsx_config.yaml`):
```yaml
yaml_dir: ../ta
out_xlsx_dir: ta_xlsx
xlsx_files:
  - ta_regions.xlsx
  - ta_segments.xlsx
  - ta_kb.xlsx
  - ta_services.xlsx
  - ta_components.xlsx
  - ta_links.xlsx
```

### 2. Конвертация XLSX → SEAF2

```bash
python _seaf2_xlsx_to_yaml.py --config ta_xlsx_to_yaml_config.yaml
```

### 3. Конвертация SEAF1 ↔ XLSX

```bash
# YAML -> XLSX
python yaml_to_xlsx.py --config seaf1_roundtrip_export.yaml

# XLSX -> YAML
python xlsx_to_yaml.py --config seaf1_roundtrip_import.yaml
```

## Особенности и улучшения

### 1. Строгая классификация объектов
В скрипте `_seaf2_yaml_to_xlsx_ta.py` реализован маппинг `CLASS_NAME_MAP`, который гарантирует корректное определение класса (например, `K8s Cluster` вместо дефолтного `Server`) независимо от регистра или специфики множественного числа в YAML.

### 2. Авто-локация (Auto Location)
- **XLSX → YAML:** Если колонка "ЦОД" пуста, локация автоматически вычисляется по именам подключенных сетей (например, `*.dc01.*` → `*.dc.01`).
- **YAML → XLSX:** Если в YAML поле `location` отсутствует, колонка "ЦОД" в Excel заполняется на основе анализа сетевых связей.

### 3. Нормализация и защита данных
- **Homoglyph Protection:** Скрипт защищен от опечаток в типах сервисов (авто-замена латинских букв на кириллические в визуально похожих символах: C, A, E, O, P, X, y).
- **Deterministic Lists:** Списки (сети, локации) в Excel всегда **отсортированы и разделены запятыми**, что обеспечивает консистентность файлов.

## Структура данных SEAF2

### Основные пространства имен:
- `seaf.company.ta.services.dc_regions` - Регионы ЦОД
- `seaf.company.ta.services.dc_azs` - Зоны доступности
- `seaf.company.ta.services.dcs` - Центры обработки данных
- `seaf.company.ta.services.dc_offices` - Офисы
- `seaf.company.ta.services.network_segments` - Сетевые сегменты
- `seaf.company.ta.services.networks` - Сети
- `seaf.company.ta.services.kbs` - База знаний
- `seaf.company.ta.components.networks` - Сетевые компоненты

## Зависимости
```bash
pip install pyyaml pandas openpyxl
```
