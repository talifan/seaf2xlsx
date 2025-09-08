import sys
import re
import argparse
from pathlib import Path
from typing import List, Dict, Any
import math
import io
import yaml
import pandas as pd

# --- Helper Functions: Start ---

def ensure_deps():
    try:
        import pandas, yaml, openpyxl
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', 'openpyxl', 'pyyaml'])

def read_excel(path: Path):
    return pd.ExcelFile(path)


def non_empty_rows(df):
    return df.dropna(how='all')

def ws_clean(s: Any) -> Any:
    if s is None: return None
    if isinstance(s, float) and math.isnan(s): return None
    s = str(s)
    if not s: return None
    s = s.replace('\u00A0', ' ').replace('\xa0', ' ').replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
    s = re.sub(r'\s+', ' ', s).strip()
    if s.lower() in {"nan", "none", "null", "n/a", "na", ""}: return None
    return s or None

def id_clean(s: Any) -> str | None:
    s = ws_clean(s)
    if s is None: return None
    return re.sub(r'\s+', '', s) or None

def to_list(val) -> List[str]:
    if val is None: return []
    if isinstance(val, list):
        return [t for x in val if (t := id_clean(x))]
    s = ws_clean(str(val)) or ''
    return [t for p in re.split(r'[;,]', s) if (t := id_clean(p))]

def safe_num(v):
    try:
        if v is None or (isinstance(v, float) and str(v) == 'nan'): return None
        if isinstance(v, (int,)):
            return int(v)
        if isinstance(v, float):
            return int(v) if v.is_integer() else v
        s = str(v).strip()
        if not s: return None
        return int(s) if s.isdigit() else float(s)
    except (ValueError, TypeError):
        return None

class IndentedDumper(yaml.SafeDumper):
    def increase_indent(self, flow=False, indentless=False):
        return super(IndentedDumper, self).increase_indent(flow, False)

def write_yaml(path: Path, data: Dict[str, Any]):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8') as f:
        yaml.dump(data, f, Dumper=IndentedDumper, allow_unicode=True, sort_keys=False)

# --- Helper Functions: End ---

# --- Reporting Functions: Start ---

def count_entities_in_xlsx(xlsx_files: List[Path]) -> Dict[str, int]:
    """Counts rows in relevant sheets of input XLSX files."""
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
                    df = non_empty_rows(xls.parse(sheet_name))
                    entity_name = sheet_map[sheet_name]
                    counts[entity_name] = counts.get(entity_name, 0) + len(df)
        except Exception as e:
            print(f"WARN: Could not process Excel file {file_path.name}: {e}", file=sys.stderr)
    return counts

def count_entities_in_yaml_dir(yaml_dir: Path) -> Dict[str, int]:
    """Counts entities in the output YAML directory."""
    counts = {}
    if not yaml_dir.exists():
        return counts
    for p in sorted(yaml_dir.glob('**/*.yaml')):
        with p.open('r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
            if not isinstance(data, dict):
                continue
            for key, value in data.items():
                if isinstance(value, dict):
                    entity_name = key.replace('seaf.ta.services.', '').replace('seaf.ta.components.', 'components.')
                    counts[entity_name] = counts.get(entity_name, 0) + len(value)
    return counts

# --- Reporting Functions: End ---

# --- Conversion Functions: Start ---

def convert_regions_az_dc_offices(xlsx_path: Path, out_dir: Path):
    xls = read_excel(xlsx_path)
    regions = {}
    if 'Регионы' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Регионы'))
        for _, row in df.iterrows():
            if not (rid := id_clean(row.get('ID Региона', ''))): continue
            regions[rid] = {'description': ws_clean(row.get('Описание')), 'external_id': rid.split('.')[-1], 'title': ws_clean(row.get('Наименование'))}
    write_yaml(out_dir / 'dc_region.yaml', {'seaf.ta.services.dc_region': regions})
    azs = {}
    if 'AZ' in xls.sheet_names:
        df = non_empty_rows(xls.parse('AZ'))
        for _, row in df.iterrows():
            if not (aid := id_clean(row.get('ID AZ', ''))): continue
            azs[aid] = {'description': ws_clean(row.get('Описание')), 'external_id': aid.split('.')[-1], 'region': id_clean(row.get('Регион')), 'title': ws_clean(row.get('Наименование')), 'vendor': ws_clean(row.get('Поставщик'))}
    write_yaml(out_dir / 'dc_az.yaml', {'seaf.ta.services.dc_az': azs})
    dcs = {}
    if 'DC' in xls.sheet_names:
        df = non_empty_rows(xls.parse('DC'))
        for _, row in df.iterrows():
            if not (did := id_clean(row.get('ID DC', ''))): continue
            dcs[did] = {'address': ws_clean(row.get('Адрес')), 'availabilityzone': id_clean(row.get('AZ')), 'description': ws_clean(row.get('Описание')), 'external_id': did.split('.')[-1], 'ownership': ws_clean(row.get('Форма владения')), 'rack_qty': safe_num(row.get('Кол-во стоек')), 'tier': ws_clean(row.get('Tier')), 'title': ws_clean(row.get('Наименование')), 'type': ws_clean(row.get('Тип')), 'vendor': ws_clean(row.get('Поставщик'))}
    write_yaml(out_dir / 'dc.yaml', {'seaf.ta.services.dc': dcs})
    offices = {}
    if 'Офисы' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Офисы'))
        for _, row in df.iterrows():
            if not (oid := id_clean(row.get('ID Офиса', ''))): continue
            offices[oid] = {'address': ws_clean(row.get('Адрес')), 'description': ws_clean(row.get('Описание')), 'external_id': oid.split('.')[-1], 'region': id_clean(row.get('Регион')), 'title': ws_clean(row.get('Наименование'))}
    write_yaml(out_dir / 'office.yaml', {'seaf.ta.services.office': offices})

def write_networks_per_location(nets: Dict[str, Any], out_dir: Path) -> List[str]:
    per_loc: Dict[str, Dict[str, Any]] = {}
    misc: Dict[str, Any] = {}
    for nid, entry in nets.items():
        if not (locs := entry.get('location')):
            misc[nid] = entry
            continue
        for loc in locs:
            token = re.sub(r'[^A-Za-z0-9]+', '_', str(loc)).strip('_') or 'loc'
            if m := re.match(r'^flix.dc.(\d+)$', str(loc)): token = f'dc{m.group(1)}'
            if m := re.match(r'^flix.office.(.+)$', str(loc)): token = f'office_{m.group(1)}'
            per_loc.setdefault(token, {})[nid] = entry
    written: List[str] = []
    for token, subset in per_loc.items():
        fname = f'networks_{token}.yaml'
        write_yaml(out_dir / fname, {'seaf.ta.services.network': subset})
        written.append(fname)
    if misc:
        fname = 'networks_misc.yaml'
        write_yaml(out_dir / fname, {'seaf.ta.services.network': misc})
        written.append(fname)
    return written

def convert_segments_nets_devices(xlsx_path: Path, out_dir: Path):
    xls = read_excel(xlsx_path)
    segments = {}
    if 'Сегменты' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Сегменты'))
        for _, row in df.iterrows():
            if not (sid := id_clean(row.get('ID сетевые сегмента/зоны', ''))): continue
            seg: Dict[str, Any] = {'title': ws_clean(row.get('Наименование')), 'description': ws_clean(row.get('Описание'))}
            if loc := id_clean(row.get('Расположение')): seg.setdefault('sber', {})['location'] = loc
            if zone := ws_clean(row.get('Зона')): seg.setdefault('sber', {})['zone'] = zone
            segments[sid] = seg
    write_yaml(out_dir / 'network_segment.yaml', {'seaf.ta.services.network_segment': segments})
    nets = {}
    if 'Сети' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Сети'))
        for _, row in df.iterrows():
            if not (nid := id_clean(row.get('ID Network', ''))): continue
            entry: Dict[str, Any] = {'title': ws_clean(row.get('Наименование')), 'description': ws_clean(row.get('Описание'))}
            if ntype := ws_clean(row.get('Тип сети')): entry['type'] = ntype
            if ntype == 'LAN':
                if (vlan := safe_num(row.get('VLAN'))) is not None: entry['vlan'] = vlan
                if ipn := ws_clean(row.get('Адрес сети')): entry['ipnetwork'] = ipn
                if lan_type := ws_clean(row.get('Тип сети (проводная, беспроводная)')): entry['lan_type'] = lan_type
            elif ntype == 'WAN':
                if wan := ws_clean(row.get('WAN Адрес')): entry['wan_ip'] = wan
            if prov := ws_clean(row.get('Провайдер')): entry['provider'] = prov
            if speed := safe_num(row.get('Скорость')): entry['bandwidth'] = speed
            if seg := id_clean(row.get('Сетевой сегмент/зона(ID)')): entry['segment'] = [seg]
            if location := id_clean(row.get('Расположение')): entry['location'] = [location]
            if vrf := (ws_clean(row.get('VRF  ')) or ws_clean(row.get('VRF'))): entry['VRF'] = vrf
            nets[nid] = entry
    write_networks_per_location(nets, out_dir)
    devices = {}
    if 'Сетевые устройства' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Сетевые устройства'))
        for _, row in df.iterrows():
            if not (did := id_clean(row.get('ID Устройства', ''))): continue
            entry = {'title': ws_clean(row.get('Наименование')), 'realization_type': ws_clean(row.get('Тип реализации')), 'type': ws_clean(row.get('Тип')), 'model': ws_clean(row.get('Модель')), 'purpose': ws_clean(row.get('Назначение')), 'address': id_clean(row.get('IP адрес')), 'description': ws_clean(row.get('Описание'))}
            if seg := id_clean(row.get('Расположение (ID сегмента/зоны)')): entry['segment'] = seg
            if nets_list := to_list(row.get('Подключенные сети (список)')): entry['network_connection'] = nets_list
            devices[did] = entry
    write_yaml(out_dir / 'components_network.yaml', {'seaf.ta.components.network': devices})

def write_root(out_dir: Path):
    imports = [p.name for p in sorted(out_dir.glob('*.yaml')) if p.name != 'root.yaml']
    write_yaml(out_dir / 'root.yaml', {'imports': imports})

# --- Conversion Functions: End ---

# --- Validation Functions: Start ---

def validate_refs(out_dir: Path) -> Dict[str, Any]:
    report: Dict[str, Any] = {'errors': [], 'warnings': []}
    def load(name: str): return yaml.safe_load((out_dir / name).read_text(encoding='utf-8')) or {}
    try:
        regions = load('dc_region.yaml').get('seaf.ta.services.dc_region', {})
        azs = load('dc_az.yaml').get('seaf.ta.services.dc_az', {})
        dcs = load('dc.yaml').get('seaf.ta.services.dc', {})
        offices = load('office.yaml').get('seaf.ta.services.office', {})
        segments = load('network_segment.yaml').get('seaf.ta.services.network_segment', {})
        devices = load('components_network.yaml').get('seaf.ta.components.network', {})
        networks: Dict[str, Any] = {}
        for p in sorted(out_dir.glob('networks_*.yaml')):
            networks.update(load(p.name).get('seaf.ta.services.network', {}))
        region_ids, az_ids, dc_ids, office_ids, seg_ids, net_ids = set(regions.keys()), set(azs.keys()), set(dcs.keys()), set(offices.keys()), set(segments.keys()), set(networks.keys())
        for i, d in azs.items():
            if (r := d.get('region')) and r not in region_ids: report['errors'].append(f'AZ {i} refs missing Region {r}')
        for i, d in dcs.items():
            if (az := d.get('availabilityzone')) and az not in az_ids: report['errors'].append(f'DC {i} refs missing AZ {az}')
        for i, d in offices.items():
            if (r := d.get('region')) and r not in region_ids: report['errors'].append(f'Office {i} refs missing Region {r}')
        for i, s in segments.items():
            if (loc := (s.get('sber') or {}).get('location')) and loc not in dc_ids and loc not in office_ids: report['errors'].append(f'Segment {i} has unknown location {loc}')
        for i, n in networks.items():
            for s in n.get('segment') or []:
                if s not in seg_ids: report['errors'].append(f'Network {i} refs missing Segment {s}')
            for l in n.get('location') or []:
                if l not in dc_ids and l not in office_ids: report['errors'].append(f'Network {i} has unknown location {l}')
        for i, d in devices.items():
            if (s := d.get('segment')) and s not in seg_ids: report['errors'].append(f'Device {i} refs missing Segment {s}')
            for n in d.get('network_connection') or []:
                if n not in net_ids: report['errors'].append(f'Device {i} refs missing Network {n}')
    except FileNotFoundError as e:
        report['errors'].append(f"Validation failed: file not found - {e.filename}. Conversion might have been incomplete.")
    return report

def validate_enums(out_dir: Path, report: Dict[str, Any]) -> None:
    def load(name: str): return yaml.safe_load((out_dir / name).read_text(encoding='utf-8')) or {}
    try:
        devices = load('components_network.yaml').get('seaf.ta.components.network', {})
        networks: Dict[str, Any] = {}
        for p in sorted(out_dir.glob('networks_*.yaml')):
            networks.update(load(p.name).get('seaf.ta.services.network', {}))
        for i, n in networks.items():
            if (t := n.get('type')) and t not in ('LAN', 'WAN'): report['errors'].append(f'Network {i} has invalid type: {t}')
            if n.get('type') == 'LAN':
                if not n.get('lan_type'): report['errors'].append(f'Network {i} missing lan_type for LAN')
                if not n.get('ipnetwork'): report['errors'].append(f'Network {i} missing ipnetwork for LAN')
        dev_type_allowed = {'Маршрутизатор', 'МСЭ', 'Контроллер WiFi', 'Криптошлюз', 'VPN', 'NAT', 'Коммутатор'}
        realization_allowed = {'Виртуальный', 'Физический'}
        for i, d in devices.items():
            if (rt := d.get('realization_type')) and rt not in realization_allowed: report['errors'].append(f'Device {i} has invalid realization_type: {rt}')
            if (dt := d.get('type')) and dt not in dev_type_allowed: report['errors'].append(f'Device {i} has invalid type: {dt}')
    except FileNotFoundError as e:
        pass # Errors are already handled by validate_refs

# --- Validation Functions: End ---

def main():
    ensure_deps()
    parser = argparse.ArgumentParser(description='Convert XLSX to YAML with referential validation and reporting.')
    parser.add_argument('--config', type=str, required=True, help='Path to YAML config with inputs and output dir')
    args = parser.parse_args()

    config_path = Path(args.config)
    with config_path.open('r', encoding='utf-8') as f:
        cfg = yaml.safe_load(f) or {}
    
    config_dir = config_path.parent
    inputs = [config_dir / p for p in (cfg.get('xlsx_files') or [])]
    out_dir = Path(cfg.get('out_yaml_dir'))
    if not out_dir.is_absolute():
        out_dir = config_dir / out_dir
    out_dir.mkdir(parents=True, exist_ok=True)

    print("---", "Source XLSX Analysis", "---")
    source_counts = count_entities_in_xlsx(inputs)
    if not source_counts:
        print("No source entities found in specified XLSX files.")
    else:
        for entity, count in sorted(source_counts.items()):
            print(f"  - Found {count} entities in sheet for '{entity}'")

    processed_something = False
    for xlsx_path in inputs:
        if not xlsx_path.exists():
            print(f"WARN: Input XLSX file not found: {xlsx_path.name}. Skipping.", file=sys.stderr)
            continue
        try:
            xls = pd.ExcelFile(xlsx_path)
            if any(sheet in xls.sheet_names for sheet in ['Регионы', 'AZ', 'DC', 'Офисы']):
                convert_regions_az_dc_offices(xlsx_path, out_dir)
                processed_something = True
            if any(sheet in xls.sheet_names for sheet in ['Сегменты', 'Сети', 'Сетевые устройства']):
                convert_segments_nets_devices(xlsx_path, out_dir)
                processed_something = True
        except Exception as e:
            print(f"ERROR: Failed to process XLSX file {xlsx_path.name}: {e}", file=sys.stderr)

    if not processed_something:
        print("\nERROR: No valid input XLSX files found to process. Aborting.", file=sys.stderr)
        sys.exit(1)

    write_root(out_dir)
    report = validate_refs(out_dir)
    validate_enums(out_dir, report)

    print("\n---", "Destination YAML Analysis", "---")
    dest_counts = count_entities_in_yaml_dir(out_dir)
    if not dest_counts:
        print("No destination entities were created.")
    else:
        for entity, count in sorted(dest_counts.items()):
            print(f"  - Created {count} entities for '{entity}'")

    print("\n---", "Conversion Summary", "---")
    all_keys = sorted(list(set(source_counts.keys()) | set(dest_counts.keys())))
    for key in all_keys:
        s_count = source_counts.get(key, 0)
        d_count = dest_counts.get(key, 0)
        status = "OK" if s_count == d_count else "FAIL"
        print(f"  - {key:<25} | Source: {s_count:<5} | Dest: {d_count:<5} | {status}")

    if report['errors']:
        print("\nVALIDATION ERRORS:", file=sys.stderr)
        for e in report['errors']:
            print('-', e, file=sys.stderr)
        print(f'\nYAML written to: {out_dir} (with validation errors)', file=sys.stderr)
        sys.exit(2)
    else:
        print(f'\nYAML written to: {out_dir} (validation OK)')

if __name__ == '__main__':
    main()