# Конвертер YAML ↔ XLSX (SEAF2)

Скрипты в этой папке конвертируют данные модели SEAF2 (`seaf.company.ta.*`) в обе стороны:
- YAML -> XLSX
- XLSX -> YAML

Поддерживается roundtrip-проверка на примере `../data/example`.

## Основные файлы

- `yaml_to_xlsx.py` - экспорт SEAF2 YAML в набор XLSX.
- `xlsx_to_yaml.py` - импорт XLSX обратно в SEAF2 YAML.
- `seaf2_config_yaml_to_xlsx.yaml` - конфиг экспорта.
- `seaf2_config_xlsx_to_yaml.yaml` - конфиг импорта.

## Зависимости

```bash
python -m pip install -U pip
python -m pip install pandas openpyxl pyyaml
```

## Структура данных SEAF2

Ключевые пространства имен:
- `seaf.company.ta.services.dc_regions`
- `seaf.company.ta.services.dc_azs`
- `seaf.company.ta.services.dcs`
- `seaf.company.ta.services.dc_offices`
- `seaf.company.ta.services.network_segments`
- `seaf.company.ta.services.networks`
- `seaf.company.ta.services.kbs`
- `seaf.company.ta.services.compute_services`
- `seaf.company.ta.services.k8s`
- `seaf.company.ta.services.cluster_virtualizations`
- `seaf.company.ta.services.monitorings`
- `seaf.company.ta.services.backups`
- `seaf.company.ta.components.networks`

## Команды запуска

Все команды запускать из папки `seaf2xlsx/`.

### 1) YAML -> XLSX

```bash
python -X utf8 yaml_to_xlsx.py --config seaf2_config_yaml_to_xlsx.yaml
```

Результат пишется в `output_xlsx/`:
- `regions_az_dc_offices.xlsx`
- `segments_nets_netdevices.xlsx`
- `kb_services.xlsx`
- `tech_services.xlsx`
- `components.xlsx`
- `links.xlsx`

### 2) XLSX -> YAML

```bash
python -X utf8 xlsx_to_yaml.py --config seaf2_config_xlsx_to_yaml.yaml --force
```

Результат пишется в `output_seaf2/`:
- доменные YAML-файлы (`dc.yaml`, `network_segment.yaml`, `compute_service.yaml`, `k8s.yaml`, ...)
- `root.yaml` (список импортов)
- `seaf_full.yaml` (объединенный файл для последующей генерации drawio)

## Roundtrip (пример)

Полный прогон:

```bash
python -X utf8 yaml_to_xlsx.py --config seaf2_config_yaml_to_xlsx.yaml
python -X utf8 xlsx_to_yaml.py --config seaf2_config_xlsx_to_yaml.yaml --force
```

Пример ожидаемого фрагмента итогового отчета `xlsx_to_yaml.py`:

```text
--- Conversion Summary ---
  - backups                   | Source: 1     | Dest: 1     | OK
  - cluster_virtualizations   | Source: 3     | Dest: 3     | OK
  - components.networks       | Source: 53    | Dest: 53    | OK
  - compute_services          | Source: 32    | Dest: 32    | OK
  - dc_azs                    | Source: 2     | Dest: 2     | OK
  - dc_offices                | Source: 1     | Dest: 1     | OK
  - dc_regions                | Source: 1     | Dest: 1     | OK
  - dcs                       | Source: 2     | Dest: 2     | OK
  - k8s                       | Source: 3     | Dest: 3     | OK
  - kbs                       | Source: 26    | Dest: 26    | OK
  - monitorings               | Source: 2     | Dest: 2     | OK
  - network_segments          | Source: 24    | Dest: 24    | OK
  - networks                  | Source: 53    | Dest: 53    | OK
  - softwares                 | Source: 6     | Dest: 6     | OK
  - storages                  | Source: 3     | Dest: 3     | OK
```

## Ожидаемое поведение

- `Deployment` (k8s deployments) намеренно не участвует в roundtrip-конвертации.
- `K8s Cluster`, `Cluster Virtualization`, `Monitoring`, `Backup` сохраняют свои типы и не деградируют в `compute_services`.
- В обработке сегментов и сетей используется верхний уровень полей (`location`, `zone`, `provider`) без `sber.*`.

## Проверка генерации drawio после roundtrip

Из корня репозитория:

```bash
python -X utf8 seaf2drawio.py -s seaf2xlsx/output_seaf2/seaf_full.yaml -d result/final_fixed.drawio
```

Ожидаемый результат в логе:

```text
Result: GENERATION MATCHES YAML (by schema)
```
