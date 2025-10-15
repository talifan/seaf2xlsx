## Конвертация XLSX ↔ YAML

### `xlsx_to_yaml.py`
Скрипт для конвертации данных из XLSX-файлов в YAML-файлы.
- **Вход:** `--config config_x2y.yaml` с полями `xlsx_files` (список XLSX-файлов для обработки) и `out_yaml_dir` (директория для сохранения YAML-файлов).
- **Нормализация данных:** Выполняет очистку и нормализацию текстовых полей, идентификаторов и списков (CSV/`;`), а также плоско заменяет переносы строк, чтобы не ломать итоговый YAML.
- **Поддерживаемые сущности:** Помимо инфраструктурных сущностей, скрипт собирает лист `Сервисы КБ` и формирует `kb.yaml` в соответствии со схемой `seaf.ta.services.kb`.
- **Разбиение сетей:** Генерирует отдельные YAML-файлы для сетей (`networks_<location>.yaml`) и включает их в `root.yaml`.
- **Валидация:** Проверяет целостность ссылок (`AZ→Region`, `DC→AZ`, `Office→Region`, `KB service→Network` и т.д.) и соответствие значений перечислениям (`enum`).
- **Отчетность:** После завершения конвертации выводит подробный отчет о количестве сущностей, найденных в исходных XLSX-файлах и созданных в YAML-файлах. Пример отчета:
```
--- Source XLSX Analysis ---
  - Found 52 entities in sheet for 'components.network'
  - Found 2 entities in sheet for 'dc'
  - Found 1 entities in sheet for 'dc_az'
  - Found 1 entities in sheet for 'dc_region'
  - Found 53 entities in sheet for 'network'
  - Found 22 entities in sheet for 'network_segment'
  - Found 1 entities in sheet for 'office'

--- Destination YAML Analysis ---
  - Created 52 entities for 'components.network'
  - Created 2 entities for 'dc'
  - Created 1 entities for 'dc_az'
  - Created 1 entities for 'dc_region'
  - Created 53 entities for 'network'
  - Created 22 entities for 'network_segment'
  - Created 1 entities for 'office'

--- Conversion Summary ---
  - components.network        | Source: 52    | Dest: 52    | OK
  - dc                        | Source: 2     | Dest: 2     | OK
  - dc_az                     | Source: 1     | Dest: 1     | OK
  - dc_region                 | Source: 1     | Dest: 1     | OK
  - network                   | Source: 53    | Dest: 53    | OK
  - network_segment           | Source: 22    | Dest: 22    | OK
  - office                    | Source: 1     | Dest: 1     | OK

YAML written to: out_yaml_homecinema (validation OK)
```

### `yaml_to_xlsx.py`
Скрипт для конвертации данных из YAML-файлов в XLSX-файлы.
- **Вход:** `--config config_y2x.yaml` с полями `yaml_dir` (директория с YAML-файлами) и `out_xlsx_dir` (директория для сохранения XLSX-файлов).
- **Сборка данных:** Собирает данные из всех релевантных YAML-файлов (включая `networks_*.yaml`, `components_network*.yaml` и `kb.yaml`) для экспорта. Для KB-сервисов формирует отдельный файл `kb_services.xlsx` с сохранением переносов строк в колонке сетей.
- **Отчетность:** После завершения конвертации выводит подробный отчет о количестве сущностей, найденных в исходных YAML-файлах и созданных в XLSX-файлах. Пример отчета:
```
--- Source YAML Analysis ---
  - Found 52 entities for 'components.network'
  - Found 2 entities for 'dc'
  - Found 1 entities for 'dc_az'
  - Found 1 entities for 'dc_region'
  - Found 53 entities for 'network'
  - Found 22 entities for 'network_segment'
  - Found 1 entities for 'office'

--- Destination XLSX Analysis ---
  - Created 52 rows for 'components.network'
  - Created 2 rows for 'dc'
  - Created 1 rows for 'dc_az'
  - Created 1 rows for 'dc_region'
  - Created 53 rows for 'network'
  - Created 22 rows for 'network_segment'
  - Created 1 rows for 'office'

--- Conversion Summary ---
  - components.network        | Source: 52    | Dest: 52    | OK
  - dc                        | Source: 2     | Dest: 2     | OK
  - dc_az                     | Source: 1     | Dest: 1     | OK
  - dc_region                 | Source: 1     | Dest: 1     | OK
  - network                   | Source: 53    | Dest: 53    | OK
  - network_segment           | Source: 22    | Dest: 22    | OK
  - office                    | Source: 1     | Dest: 1     | OK

Conversion complete. XLSX files written to: out_xlsx_homecinema
```

### Тестирование полной конвертации (Round-Trip)
Для проверки корректности двусторонней конвертации используется отдельный скрипт `run_tests.py`.
- **Назначение:** Проверяет, что данные, которые обрабатываются скриптами конвертации, могут быть без потерь преобразованы из YAML в XLSX, а затем обратно в YAML, и снова в XLSX.
- **Принцип работы:**
    1.  Конвертирует исходные YAML-файлы (включая `kb.yaml`) в набор XLSX-файлов (`regions_az_dc_offices.xlsx`, `segments_nets_netdevices.xlsx`, `kb_services.xlsx`).
    2.  Конвертирует эти промежуточные XLSX-файлы обратно в YAML.
    3.  Конвертирует полученные YAML-файлы в финальные XLSX-файлы.
    4.  Сравнивает промежуточные и финальные XLSX-файлы. Если они идентичны, тест считается пройденным.
- **Запуск:** `python run_tests.py`
- **Результат:** Выводит `STATUS: PASSED` или `STATUS: FAILED` с указанием расхождений.
