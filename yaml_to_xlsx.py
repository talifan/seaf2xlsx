import sys
import argparse
from pathlib import Path
from typing import Dict, Any


def ensure_deps():
    try:
        import pandas  # noqa: F401
        import yaml  # noqa: F401
        import openpyxl  # noqa: F401
    except Exception:
        import subprocess
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', 'openpyxl', 'pyyaml'])


def read_yaml(path: Path) -> Dict[str, Any]:
    import yaml
    with path.open('r', encoding='utf-8') as f:
        return yaml.safe_load(f) or {}


def to_series_rows(d: Dict[str, Any]):
    """Flatten mapping into DataFrame-ready rows."""
    rows = []
    for _id, payload in d.items():
        row = {'ID': _id}
        if isinstance(payload, dict):
            for k, v in payload.items():
                row[k] = v
        else:
            row['value'] = payload
        rows.append(row)
    return rows


def save_regions_az_dc_offices(yaml_dir: Path, out_xlsx: Path):
    import pandas as pd
    # Regions
    regions_yaml = yaml_dir / 'dc_region.yaml'
    az_yaml = yaml_dir / 'dc_az.yaml'
    dc_yaml = yaml_dir / 'dc.yaml'
    office_yaml = yaml_dir / 'office.yaml'

    with pd.ExcelWriter(out_xlsx, engine='openpyxl') as writer:
        # Regions
        d = read_yaml(regions_yaml).get('seaf.ta.services.dc_region', {})
        rows = []
        for _id, p in d.items():
            rows.append({
                'ID Региона': _id,
                'Наименование': p.get('title'),
                'Описание': p.get('description'),
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name='Регионы', index=False)

        # AZ
        d = read_yaml(az_yaml).get('seaf.ta.services.dc_az', {})
        rows = []
        for _id, p in d.items():
            rows.append({
                'ID AZ': _id,
                'Наименование': p.get('title'),
                'Описание': p.get('description'),
                'Поставщик': p.get('vendor'),
                'Регион': p.get('region'),
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name='AZ', index=False)

        # DC
        d = read_yaml(dc_yaml).get('seaf.ta.services.dc', {})
        rows = []
        for _id, p in d.items():
            rows.append({
                'ID DC': _id,
                'Наименование': p.get('title'),
                'Описание': p.get('description'),
                'Поставщик': p.get('vendor'),
                'Tier': p.get('tier'),
                'Тип': p.get('type'),
                'Кол-во стоек': p.get('rack_qty'),
                'Адрес': p.get('address'),
                'Форма владения': p.get('ownership'),
                'AZ': p.get('availabilityzone'),
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name='DC', index=False)

        # Offices
        d = read_yaml(office_yaml).get('seaf.ta.services.office', {})
        rows = []
        for _id, p in d.items():
            rows.append({
                'ID Офиса': _id,
                'Наименование': p.get('title'),
                'Описание': p.get('description'),
                'Адрес': p.get('address'),
                'Регион': p.get('region'),
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name='Офисы', index=False)


def save_segments_nets_devices(yaml_dir: Path, out_xlsx: Path):
    import pandas as pd
    seg_yaml = yaml_dir / 'network_segment.yaml'
    # collect networks from all networks_*.yaml

    with pd.ExcelWriter(out_xlsx, engine='openpyxl') as writer:
        # Segments
        d = read_yaml(seg_yaml).get('seaf.ta.services.network_segment', {})
        rows = []
        for _id, p in d.items():
            rows.append({
                'ID сетевые сегмента/зоны': _id,
                'Наименование': p.get('title'),
                'Описание': p.get('description'),
                'Расположение': (p.get('sber') or {}).get('location'),
                'Зона': (p.get('sber') or {}).get('zone'),
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name='Сегменты', index=False)

        # Networks
        d = {}
        for p in sorted(yaml_dir.glob('networks_*.yaml')):
            d.update(read_yaml(p).get('seaf.ta.services.network', {}) or {})
        rows = []
        for _id, p in d.items():
            sber = p.get('sber') or {}
            # Preserve provider/bandwidth from top-level if present, fallback to vendor block
            provider = p.get('provider') if p.get('provider') is not None else sber.get('provider')
            speed = p.get('bandwidth') if p.get('bandwidth') is not None else sber.get('bandwidth')
            vrf = p.get('VRF')
            seg_val = p.get('segment')
            if isinstance(seg_val, list):
                seg_out = ', '.join(seg_val)
            else:
                seg_out = seg_val or ''
            loc_val = p.get('location')
            if isinstance(loc_val, list):
                loc_out = ', '.join(loc_val)
            else:
                loc_out = loc_val or ''
            rows.append({
                'ID Network': _id,
                'Наименование': p.get('title'),
                'Описание': p.get('description'),
                'Тип сети': p.get('type'),
                'VLAN': p.get('vlan'),
                'VRF  ': vrf,
                'Провайдер': provider,
                'Скорость': speed,
                'Резервирование': sber.get('reservation'),
                'Тип сети (проводная, беспроводная)': p.get('lan_type'),
                'Адрес сети': p.get('ipnetwork'),
                'WAN Адрес': p.get('wan_ip'),
                'Расположение': loc_out,
                'Сетевой сегмент/зона(ID)': seg_out,
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name='Сети', index=False)

        # Devices (merge from any components_network*.yaml in the dir)
        d = {}
        for pth in sorted(yaml_dir.glob('components_network*.yaml')):
            d.update(read_yaml(pth).get('seaf.ta.components.network', {}) or {})
        rows = []
        for _id, p in d.items():
            rows.append({
                'ID Устройства': _id,
                'Наименование': p.get('title'),
                'Тип реализации': p.get('realization_type'),
                'Тип': p.get('type'),
                'Модель': p.get('model'),
                'Назначение': p.get('purpose'),
                'IP адрес': p.get('address'),
                'Описание': p.get('description'),
                'Расположение (ID сегмента/зоны)': p.get('segment'),
                'Подключенные сети (список)': ', '.join(p.get('network_connection') or []),
            })
        pd.DataFrame(rows).to_excel(writer, sheet_name='Сетевые устройства', index=False)

        # Legend placeholder sheet if desired
        pd.DataFrame({'EXTERNAL-NET': ['INTERNET', 'TRANSPORT-WAN', 'INET-EDGE'],
                      '- Внешние сети (партнеры, подрядчики и т.д.)': ['- Интернет', '- Транспортная сеть', '- Сегмент подключения к интернет']
                      }).to_excel(writer, sheet_name='-----', index=False)


def main():
    ensure_deps()
    parser = argparse.ArgumentParser(description='Convert YAML to XLSX (roundtrip helper)')
    parser.add_argument('--config', type=str, help='Path to YAML config with yaml inputs and xlsx outputs')
    args = parser.parse_args()

    yaml_dir = Path('scripts') / 'out_yaml'
    out_dir = Path('scripts') / 'out_xlsx'

    if args.config:
        import yaml
        with open(args.config, 'r', encoding='utf-8') as f:
            cfg = yaml.safe_load(f) or {}
        # allow overriding dirs
        if cfg.get('yaml_dir'):
            yaml_dir = Path(cfg['yaml_dir'])
        if cfg.get('out_xlsx_dir') or cfg.get('out_dir'):
            out_dir = Path(cfg.get('out_xlsx_dir') or cfg.get('out_dir'))

    out_dir.mkdir(parents=True, exist_ok=True)
    save_regions_az_dc_offices(yaml_dir, out_dir / 'regions_az_dc_offices_roundtrip.xlsx')
    save_segments_nets_devices(yaml_dir, out_dir / 'segments_nets_netdevices_roundtrip.xlsx')
    print(f'Roundtrip XLSX written to: {out_dir}')


if __name__ == '__main__':
    main()
