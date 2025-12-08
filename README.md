# Конвертер YAML ↔ XLSX (SEAF1 и SEAF2)

Этот набор скриптов обеспечивает полную совместимость между форматами SEAF1, SEAF2 и XLSX, позволяя конвертировать данные в обоих направлениях.

## Основные файлы

### Конвертеры для SEAF2
- `_seaf2_yaml_to_xlsx.py` - Конвертация YAML (SEAF2) → XLSX
- `_seaf2_xlsx_to_yaml.py` - Конвертация XLSX → YAML (SEAF2)

### Оригинальные скрипты (для SEAF1)
- `yaml_to_xlsx.py` - Конвертация YAML (SEAF1) → XLSX
- `xlsx_to_yaml.py` - Конвертация XLSX → YAML (SEAF1)
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
python _seaf2_yaml_to_xlsx.py --config config_seaf2_y2x.yaml
```

Пример конфигурации:
```yaml
yaml_dir: example/seaf2
out_xlsx_dir: output_xlsx
xlsx_files:
  - regions_az_dc_offices.xlsx
  - segments_nets_netdevices.xlsx
  - kb_services.xlsx
```

### 2. Конвертация XLSX → SEAF2

```bash
python _seaf2_xlsx_to_yaml.py --config config_seaf2_x2y.yaml
```

Пример конфигурации:
```yaml
xlsx_files:
  - output_xlsx/regions_az_dc_offices.xlsx
  - output_xlsx/segments_nets_netdevices.xlsx
  - output_xlsx/kb_services.xlsx
out_yaml_dir: output_seaf2
```

### 3. Конвертация SEAF1 → XLSX

```bash
python yaml_to_xlsx.py --config config_y2x.yaml
```

### 4. Конвертация XLSX → SEAF1

```bash
python xlsx_to_yaml.py --config config_x2y.yaml
```

### 5. Тестирование конвертации SEAF1 (Roundtrip)

```bash
python tests/run_tests.py
```

### 6. Тестирование конвертации SEAF2 (Roundtrip)

```bash
python tests/run_tests_seaf2.py
```

### 7. Тестирование конвертации между форматами (Cross-check)

Для тестирования полной совместимости между форматами (SEAF1 → XLSX → SEAF2 и обратно) используйте:

```bash
python tests/run_cross_tests.py
```

Этот скрипт автоматически:
1. Конвертирует SEAF1 → XLSX → SEAF2 → XLSX
2. Конвертирует SEAF2 → XLSX → SEAF1 → XLSX
3. Сравнивает результаты на каждом этапе.

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

## Особенности обработки сетевых устройств

### Парсинг расположений
Колонка `Расположение` может содержать одно значение или список площадок. После парсинга в объект попадает либо строка, либо список идентификаторов (`sbs.dc.01`, `sbs.office.msk-hq`, ...).

### Автоопределение сегмента
- Если в XLSX заполнено поле `Сетевой сегмент/зона (ID)`, оно переносится без изменений.
- В противном случае скрипт собирает все сети из `Подключённые сети`, находит их сегменты и выбирает те, у которых `sber.location` совпадает с текущей площадкой устройства.

### Разделение по площадкам
Когда устройство привязано к нескольким площадкам, создаются индивидуальные копии (`<исходный ID>-03`, `<исходный ID>-05`, …). Для каждой копии вычисляется собственный сегмент.

### Контроль отсутствующих сетей
Если устройство ссылается на сеть, которой нет в выгруженных данных, идентификатор фиксируется и выводится предупреждение `WARN: Networks referenced by devices but missing in network data: ...`.

## Валидация

### Для SEAF1
- Ссылки `AZ → Region`, `DC → AZ`, `Network → Segment`, `Device → Network` и др. проверяются в `validate_refs`.
- `validate_enums` контролирует допустимые значения перечислений.
- Любые несоответствия подсвечиваются через `WARN`/`ERROR` в отчёте, но выполнение не прерывается.

### Для SEAF2
- Аналогичная валидация с учётом новой структуры пространств имён
- Дополнительная проверка новых атрибутов SEAF2

## Выходные данные и отчёт

После конвертации печатаются сводки:
- `--- Source XLSX Analysis ---`
- `--- Destination YAML Analysis ---`
- `--- Conversion Summary ---`

С числом сущностей до/после и пометкой `OK/FAIL`.

## Зависимости

```bash
pip install pyyaml pandas openpyxl
```

## Совместимость

Скрипты полностью совместимы с существующими XLSX файлами и сохраняют ту же структуру полей и свойств. Обеспечивается полная обратная совместимость с форматом SEAF1.

## Результаты тестирования

Все тесты пройдены успешно:
- ✅ Конвертация SEAF1 → XLSX → SEAF1 работает корректно
- ✅ Конвертация SEAF2 → XLSX → SEAF2 работает корректно
- ✅ Конвертация между форматами через XLSX работает корректно
- ✅ Обратная совместимость с SEAF1 обеспечена
- ✅ Структура XLSX файлов сохраняется неизменной

Эти процедуры позволяют обеспечить воспроизводимость данных и быстро находить расхождения между табличными шаблонами и YAML‑представлением.
