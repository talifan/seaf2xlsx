import sys
import argparse
from pathlib import Path
from typing import Dict, Any, List
import pandas as pd
import yaml

# --- Helper Functions: Start ---

def ensure_deps():
    try:
        import pandas, yaml, openpyxl
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', 'openpyxl', 'pyyaml'])

def read_yaml(path: Path) -> Dict[str, Any]:
    with path.open('r', encoding='utf-8') as f:
        return yaml.safe_load(f) or {}

# --- Helper Functions: End ---

# --- Reporting Functions: Start ---

def count_entities_in_yaml_dir(yaml_dir: Path) -> Dict[str, int]:
    """Counts entities in the source YAML directory."""
    counts = {}
    if not yaml_dir.exists():
        return counts
    
    # Define which entity types are handled by this script
    handled_entities = {
        'seaf.ta.services.dc_region',
        'seaf.ta.services.dc_az',
        'seaf.ta.services.dc',
        'seaf.ta.services.office',
        'seaf.ta.services.network_segment',
        'seaf.ta.services.network',
        'seaf.ta.components.network'
    }

    for p in sorted(yaml_dir.glob('**/*.yaml')):
        data = read_yaml(p)
        for key, value in data.items():
            if key in handled_entities and isinstance(value, dict):
                entity_name = key.replace('seaf.ta.services.', '').replace('seaf.ta.components.', 'components.')
                counts[entity_name] = counts.get(entity_name, 0) + len(value)
    return counts

def count_entities_in_xlsx(xlsx_files: List[Path]) -> Dict[str, int]:
    """Counts rows in the output XLSX files."""
    counts = {}
    sheet_map = {
        'Регионы': 'dc_region',
        'AZ': 'dc_az',
        'DC': 'dc',
        'Офисы': 'office',
        'Сегменты': 'network_segment',
        'Сети': 'network',
        'Сетевые устройства': 'components.network'
    }
    for file_path in xlsx_files:
        if not file_path.exists():
            continue
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                if sheet_name in sheet_map:
                    df = xls.parse(sheet_name).dropna(how='all')
                    entity_name = sheet_map[sheet_name]
                    counts[entity_name] = counts.get(entity_name, 0) + len(df)
        except Exception as e:
            print(f"WARN: Could not read back Excel file {file_path.name} for counting: {e}", file=sys.stderr)
    return counts

# --- Reporting Functions: End ---

# --- Conversion Functions: Start ---

def save_regions_az_dc_offices(yaml_dir: Path, out_xlsx: Path):
    with pd.ExcelWriter(out_xlsx, engine='openpyxl') as writer:
        d = read_yaml(yaml_dir / 'dc_region.yaml').get('seaf.ta.services.dc_region', {})
        pd.DataFrame([{'ID Региона': _id, 'Наименование': p.get('title'), 'Описание': p.get('description')} for _id, p in d.items()]).to_excel(writer, sheet_name='Регионы', index=False)

        d = read_yaml(yaml_dir / 'dc_az.yaml').get('seaf.ta.services.dc_az', {})
        pd.DataFrame([{'ID AZ': _id, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Поставщик': p.get('vendor'), 'Регион': p.get('region')} for _id, p in d.items()]).to_excel(writer, sheet_name='AZ', index=False)

        d = read_yaml(yaml_dir / 'dc.yaml').get('seaf.ta.services.dc', {})
        pd.DataFrame([{'ID DC': _id, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Поставщик': p.get('vendor'), 'Tier': p.get('tier'), 'Тип': p.get('type'), 'Кол-во стоек': p.get('rack_qty'), 'Адрес': p.get('address'), 'Форма владения': p.get('ownership'), 'AZ': p.get('availabilityzone')} for _id, p in d.items()]).to_excel(writer, sheet_name='DC', index=False)

        d = read_yaml(yaml_dir / 'office.yaml').get('seaf.ta.services.office', {})
        pd.DataFrame([{'ID Офиса': _id, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Адрес': p.get('address'), 'Регион': p.get('region')} for _id, p in d.items()]).to_excel(writer, sheet_name='Офисы', index=False)

def save_segments_nets_devices(yaml_dir: Path, out_xlsx: Path):
    with pd.ExcelWriter(out_xlsx, engine='openpyxl') as writer:
        d = read_yaml(yaml_dir / 'network_segment.yaml').get('seaf.ta.services.network_segment', {})
        pd.DataFrame([{'ID сетевые сегмента/зоны': _id, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Расположение': (p.get('sber') or {}).get('location'), 'Зона': (p.get('sber') or {}).get('zone')} for _id, p in d.items()]).to_excel(writer, sheet_name='Сегменты', index=False)

        nets_data = {}
        for p in sorted(yaml_dir.glob('networks_*.yaml')):
            nets_data.update(read_yaml(p).get('seaf.ta.services.network', {}))
        rows = []
        for _id, p in nets_data.items():
            sber = p.get('sber') or {}
            provider = p.get('provider') or sber.get('provider')
            vrf = p.get('VRF')
            seg_out = ', '.join(p.get('segment', []))
            loc_out = ', '.join(p.get('location', []))
            rows.append({'ID Network': _id, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Тип сети': p.get('type'), 'VLAN': p.get('vlan'), 'VRF  ': vrf, 'Провайдер': provider, 'Резервирование': sber.get('reservation'), 'Тип сети (проводная, беспроводная)': p.get('lan_type'), 'Адрес сети': p.get('ipnetwork'), 'WAN Адрес': p.get('wan_ip'), 'Расположение': loc_out, 'Сетевой сегмент/зона(ID)': seg_out})
        pd.DataFrame(rows).to_excel(writer, sheet_name='Сети', index=False)

        dev_data = {}
        for pth in sorted(yaml_dir.glob('components_network*.yaml')):
            dev_data.update(read_yaml(pth).get('seaf.ta.components.network', {}))
        pd.DataFrame([{'ID Устройства': _id, 'Наименование': p.get('title'), 'Тип реализации': p.get('realization_type'), 'Тип': p.get('type'), 'Модель': p.get('model'), 'Назначение': p.get('purpose'), 'IP адрес': p.get('address'), 'Описание': p.get('description'), 'Расположение (ID сегмента/зоны)': p.get('segment'), 'Подключенные сети (список)': ', '.join(p.get('network_connection') or [])} for _id, p in dev_data.items()]).to_excel(writer, sheet_name='Сетевые устройства', index=False)

# --- Conversion Functions: End ---

def main():
    ensure_deps()
    parser = argparse.ArgumentParser(description='Convert YAML to XLSX with reporting.')
    parser.add_argument('--config', type=str, required=True, help='Path to YAML config')
    args = parser.parse_args()

    config_path = Path(args.config)
    with config_path.open('r', encoding='utf-8') as f:
        cfg = yaml.safe_load(f) or {}

    config_dir = config_path.parent
    yaml_dir = Path(cfg.get('yaml_dir'))
    if not yaml_dir.is_absolute():
        yaml_dir = config_dir / yaml_dir
    
    out_dir = Path(cfg.get('out_xlsx_dir'))
    if not out_dir.is_absolute():
        out_dir = config_dir / out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    print("--- Source YAML Analysis ---")
    source_counts = count_entities_in_yaml_dir(yaml_dir)
    if not source_counts:
        print("No source entities found in specified YAML directory.")
    else:
        for entity, count in sorted(source_counts.items()):
            print(f"  - Found {count} entities for '{entity}'")

    out_files = [out_dir / f for f in cfg.get('xlsx_files', [])]
    reg_file = next((f for f in out_files if 'regions' in f.name), None)
    seg_file = next((f for f in out_files if 'segments' in f.name), None)

    if reg_file:
        save_regions_az_dc_offices(yaml_dir, reg_file)
        print(f"Saved regions, AZs, DCs, and offices to {reg_file.name}")
    if seg_file:
        save_segments_nets_devices(yaml_dir, seg_file)
        print(f"Saved segments, networks, and devices to {seg_file.name}")

    if not reg_file and not seg_file:
        print("\nERROR: No output files specified in config under 'xlsx_files'. Nothing to do.", file=sys.stderr)
        sys.exit(1)

    print("\n--- Destination XLSX Analysis ---")
    dest_counts = count_entities_in_xlsx(out_files)
    if not dest_counts:
        print("No destination entities were created.")
    else:
        for entity, count in sorted(dest_counts.items()):
            print(f"  - Created {count} rows for '{entity}'")

    print("\n--- Conversion Summary ---")
    all_keys = sorted(list(set(source_counts.keys()) | set(dest_counts.keys())))
    for key in all_keys:
        s_count = source_counts.get(key, 0)
        d_count = dest_counts.get(key, 0)
        status = "OK" if s_count == d_count else "FAIL"
        print(f"  - {key:<25} | Source: {s_count:<5} | Dest: {d_count:<5} | {status}")

    print(f'\nConversion complete. XLSX files written to: {out_dir}')

if __name__ == '__main__':
    main()
