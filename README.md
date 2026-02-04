# Конвертер YAML ↔ XLSX (SEAF1 и SEAF2)

Этот набор скриптов обеспечивает полную совместимость между форматами SEAF1 (сингулярные ключи), SEAF2 (множественные ключи) и XLSX, позволяя конвертировать данные в обоих направлениях.

## Основные файлы

### Скрипты SEAF1 (сингулярные ключи)
- `yaml_to_xlsx.py` - Конвертация YAML → XLSX
- `xlsx_to_yaml.py` - Конвертация XLSX → YAML
- Ключи в YAML: `seaf.ta.services.compute_service`, `seaf.ta.services.dc_region` и т.д.

### Скрипты SEAF2 (множественные ключи)
- `_seaf2_yaml_to_xlsx.py` - Конвертация YAML (SEAF2) → XLSX
- `_seaf2_xlsx_to_yaml.py` - Конвертация XLSX → YAML (SEAF2)
- Ключи в YAML: `seaf.company.ta.services.compute_services`, `seaf.company.ta.services.dc_regions` и т.д.

## Ключевые различия

| Параметр | SEAF1 (Standard) | SEAF2 |
|----------|----------|-------|
| Пространство имен | `seaf.ta.*` | `seaf.company.ta.*` |
| Имена сущностей | единственное число (singular) | множественное число (plural) |
| Тех. сервисы (root) | `compute_service` | `compute_services` |
| Файл офисов | `office.yaml` | `dc_office.yaml` |
| Файл компонентов | `components_network.yaml` | `network_component.yaml` |

## Поддерживаемые сущности (Entity Support)

### Вычислительные ресурсы (Лист: Тех. сервисы)
- Compute Service, Cluster, Monitoring, Backup
- Cluster Virtualization, Software, Storage
- Kubernetes (Кластеры и Deployments)

### Инфраструктурные компоненты (Лист: Компоненты)
- Servers (Физические и Виртуальные)
- HW Storage (Аппаратные СХД)
- User Devices (Рабочие места, IoT)
- K8s Nodes, Namespaces, HPA
- Network Devices (Маршрутизаторы, МСЭ и др.)

### Топология и Связи
- Regions, AZ, DC, Offices (Лист: Регионы)
- Network Segments, Networks (Лист: Сети)
- Logical Links, Network Links (Лист: Связи)
- Stands, Environments (Лист: Стенды и окружения)

## Особенности

### 1. Авто-локация (Auto Location)
- **XLSX → YAML:** Если колонка "ЦОД" пуста, локация вычисляется по именам подключенных сетей (например, `*.dc01.*` → `*.dc.01`).
- **YAML → XLSX:** Если в YAML поле `location` пустое, колонка "ЦОД" в Excel будет заполнена на основе анализа сетевых связей.

### 2. Нормализация данных
- **Homoglyph Protection:** Скрипт защищен от опечаток в типах сервисов (авто-замена латинских букв на кириллические в похожих символах: C, A, E, O и др.).
- **Deterministic Lists:** Все списки (сети, локации) в Excel теперь всегда **отсортированы и разделены запятыми**, что гарантирует идентичность файлов при повторной конвертации.

## Обработка ошибок
- **Проверка конфига:** Валидация путей и наличия обязательных параметров.
- **I/O Safety:** Обработка ошибок доступа (например, если файл открыт в Excel).
- **Graceful Skip:** Если один из файлов в списке отсутствует, скрипт выведет ошибку в `stderr`, но продолжит обработку остальных данных.

## Использование

```bash
# Конвертация в Excel
python yaml_to_xlsx.py --config config.yaml

# Конвертация в YAML
python xlsx_to_yaml.py --config config.yaml
```

## Зависимости
```bash
pip install pyyaml pandas openpyxl
```
