## Конвертация XLSX ↔ YAML
- `scripts/xlsx_to_yaml.py`:
  - Вход: `--config scripts/config_xlsx_to_yaml.example.yaml` с полями `xlsx_files`, `out_yaml_dir`.
  - Нормализация: удаление неразрывных/управляющих пробелов, тримминг, схлопывание; для ID/ссылок — удаление внутренних пробелов; списки — CSV/`;` с нормализацией.
  - Поля `Расположение` и `Сетевой сегмент/зона(ID)` на листе 'Сети' теперь поддерживают множественные значения, указанные через запятую.
  - Разбиение сетей: файлы `networks_<location>.yaml` + `networks_misc.yaml` (если нет локации); `root.yaml` включает все `networks_*.yaml`.
  - Валидация связей: `AZ→Region`, `DC→AZ`, `Office→Region`, `Segment→(DC|Office)`, `Network→(Segments,Locations)`, `Device→(Segment,Networks)`.
  - Валидация enum: `Network.type ∈ {LAN,WAN}`, при `LAN` — `lan_type ∈ {Проводная, Беспроводная}`, устройства (`realization_type`, `type`).
  - Зона сегмента (`sber.zone`) поддерживается из XLSX, но валидация зон по enum из DZO выключена по требованию.
  - Отчёты: ошибки и предупреждения печатаются в консоль и сохраняются в `scripts/out_yaml/validation_report.txt`; при ошибках exit code = 2.

- `scripts/yaml_to_xlsx.py`:
  - Вход: `--config scripts/config_yaml_to_xlsx.example.yaml` с полями `yaml_dir`, `out_xlsx_dir`.
  - Собирает сети из всех `networks_*.yaml` для обратного экспорта.

Пример конфигов: `scripts/config_xlsx_to_yaml.example.yaml`, `scripts/config_yaml_to_xlsx.example.yaml`