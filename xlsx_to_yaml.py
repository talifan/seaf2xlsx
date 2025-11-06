import sys
import re
import argparse
from pathlib import Path
from typing import List, Dict, Any, Tuple
import math
import io
import yaml
import pandas as pd

DEBUG_LOG_FILE = Path('debug_script.log')

def log_debug(message):
    with DEBUG_LOG_FILE.open('a', encoding='utf-8') as f:
        f.write(f"{message}\n")

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

def parse_multiline_ids(val) -> List[str]:
    if val is None: return []
    if isinstance(val, list):
        return [t for x in val if (t := id_clean(x))]
    text = str(val).replace('\r', '\n')
    tokens: List[str] = []
    for line in text.split('\n'):
        segment = line.strip()
        if not segment: continue
        if segment.startswith('-'):
            segment = segment[1:].strip()
        for piece in re.split(r'[;,]', segment):
            if cleaned := id_clean(piece):
                tokens.append(cleaned)
    return tokens

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

def parse_locations(val: Any) -> List[str]:
    """Splits location string into individual IDs and normalizes tokens."""
    if val is None: return []
    s = ws_clean(str(val))
    if not s: return []
    return [t for p in re.split(r'[;,\s]+', s) if (t := id_clean(p))]

class IndentedDumper(yaml.SafeDumper):
    def increase_indent(self, flow=False, indentless=False):
        return super(IndentedDumper, self).increase_indent(flow, False)

def write_yaml(path: Path, data: Dict[str, Any]):
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8') as f:
        yaml.dump(sanitize_for_yaml(data), f, Dumper=IndentedDumper, allow_unicode=True, sort_keys=False)

def sanitize_for_yaml(value: Any) -> Any:
    if isinstance(value, dict):
        return {k: sanitize_for_yaml(v) for k, v in value.items() if not k.startswith('_')}
    if isinstance(value, list):
        return [sanitize_for_yaml(v) for v in value]
    if isinstance(value, str):
        return re.sub(r'[\r\n]+', ' ', value)
    return value

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
        'Сетевые устройства': 'components.network',
        'Сервисы КБ': 'kb'
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
                    row_count = len(df)
                    if entity_name == 'network':
                        id_count = 0
                        for col in ('ID Network',):
                            if col in df.columns:
                                id_count = df[col].apply(id_clean).dropna().shape[0]
                                break
                        if id_count:
                            row_count = id_count
                    counts[entity_name] = counts.get(entity_name, 0) + row_count
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
    
    # --- Regions ---
    regions, processed_ids = {}, set()
    if 'Регионы' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Регионы'))
        for index, row in df.iterrows():
            rid = id_clean(row.get('ID Региона'))
            if not rid:
                print(f"WARN: Skipping row {index + 2} in sheet 'Регионы' of '{xlsx_path.name}' due to missing ID.", file=sys.stderr)
                continue
            if rid in processed_ids:
                print(f"WARN: Skipping duplicate ID '{rid}' in sheet 'Регионы', row {index + 2} of '{xlsx_path.name}'.", file=sys.stderr)
                continue
            processed_ids.add(rid)
            regions[rid] = {'description': ws_clean(row.get('Описание')), 'external_id': rid.split('.')[-1], 'title': ws_clean(row.get('Наименование'))}
    if regions:
        write_yaml(out_dir / 'dc_region.yaml', {'seaf.ta.services.dc_region': regions})
    
    # --- AZs ---
    azs, processed_ids = {}, set()
    if 'AZ' in xls.sheet_names:
        df = non_empty_rows(xls.parse('AZ'))
        for index, row in df.iterrows():
            aid = id_clean(row.get('ID AZ'))
            if not aid:
                print(f"WARN: Skipping row {index + 2} in sheet 'AZ' of '{xlsx_path.name}' due to missing ID.", file=sys.stderr)
                continue
            if aid in processed_ids:
                print(f"WARN: Skipping duplicate ID '{aid}' in sheet 'AZ', row {index + 2} of '{xlsx_path.name}'.", file=sys.stderr)
                continue
            processed_ids.add(aid)
            azs[aid] = {'description': ws_clean(row.get('Описание')), 'external_id': aid.split('.')[-1], 'region': id_clean(row.get('Регион')), 'title': ws_clean(row.get('Наименование')), 'vendor': ws_clean(row.get('Поставщик'))}
    if azs:
        write_yaml(out_dir / 'dc_az.yaml', {'seaf.ta.services.dc_az': azs})

    # --- DCs ---
    dcs, processed_ids = {}, set()
    if 'DC' in xls.sheet_names:
        df = non_empty_rows(xls.parse('DC'))
        for index, row in df.iterrows():
            did = id_clean(row.get('ID DC'))
            if not did:
                print(f"WARN: Skipping row {index + 2} in sheet 'DC' of '{xlsx_path.name}' due to missing ID.", file=sys.stderr)
                continue
            if did in processed_ids:
                print(f"WARN: Skipping duplicate ID '{did}' in sheet 'DC', row {index + 2} of '{xlsx_path.name}'.", file=sys.stderr)
                continue
            processed_ids.add(did)
            dcs[did] = {'address': ws_clean(row.get('Адрес')), 'availabilityzone': id_clean(row.get('AZ')), 'description': ws_clean(row.get('Описание')), 'external_id': did.split('.')[-1], 'ownership': ws_clean(row.get('Форма владения')), 'rack_qty': safe_num(row.get('Кол-во стоек')), 'tier': ws_clean(row.get('Tier')), 'title': ws_clean(row.get('Наименование')), 'type': ws_clean(row.get('Тип')), 'vendor': ws_clean(row.get('Поставщик'))}
    if dcs:
        write_yaml(out_dir / 'dc.yaml', {'seaf.ta.services.dc': dcs})

    # --- Offices ---
    offices, processed_ids = {}, set()
    if 'Офисы' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Офисы'))
        for index, row in df.iterrows():
            oid = id_clean(row.get('ID Офиса'))
            if not oid:
                print(f"WARN: Skipping row {index + 2} in sheet 'Офисы' of '{xlsx_path.name}' due to missing ID.", file=sys.stderr)
                continue
            if oid in processed_ids:
                print(f"WARN: Skipping duplicate ID '{oid}' in sheet 'Офисы', row {index + 2} of '{xlsx_path.name}'.", file=sys.stderr)
                continue
            processed_ids.add(oid)
            offices[oid] = {'address': ws_clean(row.get('Адрес')), 'description': ws_clean(row.get('Описание')), 'external_id': oid.split('.')[-1], 'region': id_clean(row.get('Регион')), 'title': ws_clean(row.get('Наименование'))}
    if offices:
        write_yaml(out_dir / 'office.yaml', {'seaf.ta.services.office': offices})

def write_networks_per_location(nets: Dict[str, Any], out_dir: Path, company_prefix: str):
    per_loc: Dict[str, Dict[str, Any]] = {}
    misc: Dict[str, Any] = {}
    
    if not company_prefix:
        print("WARN: No company prefix found, cannot generate location-specific network files correctly.", file=sys.stderr)

    dc_pattern = re.compile(rf'{re.escape(company_prefix)}[.]dc[.](\d+)$')
    office_pattern = re.compile(rf'{re.escape(company_prefix)}[.]office[.](.+)$')

    for nid, entry in nets.items():
        if not (locs := entry.get('location')):
            misc[nid] = entry
            continue
        for loc in locs:
            token = re.sub(r'[^A-Za-z0-9]+', '_', str(loc)).strip('_') or 'loc'
            if company_prefix:
                if m := dc_pattern.match(str(loc)): token = f'dc{m.group(1)}'
                if m := office_pattern.match(str(loc)): token = f'office_{m.group(1)}'
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

def convert_segments_nets_devices(xlsx_path: Path, out_dir: Path) -> int:
    xls = read_excel(xlsx_path)
    auto_created_segments_count = 0

    # --- Segments ---
    segments, processed_ids = {}, set()
    if 'Сегменты' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Сегменты'))
        for index, row in df.iterrows():
            sid = id_clean(row.get('ID сетевые сегмента/зоны'))
            if not sid:
                print(f"WARN: Skipping row {index + 2} in sheet 'Сегменты' of '{xlsx_path.name}' due to missing ID.", file=sys.stderr)
                continue
            if sid in processed_ids:
                print(f"WARN: Skipping duplicate ID '{sid}' in sheet 'Сегменты', row {index + 2} of '{xlsx_path.name}'.", file=sys.stderr)
                continue
            processed_ids.add(sid)

            base_title = ws_clean(row.get('Наименование'))
            zone = ws_clean(row.get('Зона'))
            locs = parse_locations(row.get('Расположение'))

            if not locs:
                seg: Dict[str, Any] = {'title': base_title, 'description': ws_clean(row.get('Описание'))}
                if zone: seg.setdefault('sber', {})['zone'] = zone
                segments[sid] = seg
                continue

            primary_loc = locs[0]
            seg: Dict[str, Any] = {'title': base_title, 'description': ws_clean(row.get('Описание'))}
            seg.setdefault('sber', {})['location'] = primary_loc
            if zone: seg['sber']['zone'] = zone
            segments[sid] = seg

            if len(locs) > 1 and base_title and zone:
                company_prefix = sid.split('.')[0] if '.' in sid else ''
                for extra_loc_id in locs[1:]:
                    loc_postfix = extra_loc_id.split('.')[-1]
                    if extra_loc_id.startswith(f'{company_prefix}.dc.'):
                        loc_postfix = f"dc{loc_postfix}"
                    
                    zone_slug = zone.lower().replace(' ', '_')
                    new_seg_id = f"{company_prefix}.network_segment.{zone_slug}.{loc_postfix}"
                    new_seg_title = f"{base_title}_{loc_postfix}"

                    if new_seg_id in segments or new_seg_id in processed_ids:
                        continue
                    
                    new_segment = {
                        'title': new_seg_title,
                        'description': f'Автоматически созданный сегмент по шаблону из строки {index + 2} листа \'Сегменты\'',
                        'sber': {
                            'location': extra_loc_id,
                            'zone': zone
                        }
                    }
                    segments[new_seg_id] = new_segment
                    processed_ids.add(new_seg_id)
                    auto_created_segments_count += 1

    # --- Networks ---
    nets, processed_ids = {}, set()
    if 'Сети' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Сети'))
        for index, row in df.iterrows():
            nid = id_clean(row.get('ID Network'))
            if not nid:
                print(f"WARN: Skipping row {index + 2} in sheet 'Сети' of '{xlsx_path.name}' due to missing ID.", file=sys.stderr)
                continue
            if nid in processed_ids:
                print(f"WARN: Skipping duplicate ID '{nid}' in sheet 'Сети', row {index + 2} of '{xlsx_path.name}'.", file=sys.stderr)
                continue
            processed_ids.add(nid)
            entry: Dict[str, Any] = {
                '_source_info': f"Sheet 'Сети', row {index + 2}",
                'title': ws_clean(row.get('Наименование')), 
                'description': ws_clean(row.get('Описание'))
            }
            if ntype := ws_clean(row.get('Тип сети')): entry['type'] = ntype
            if ntype == 'LAN':
                if (vlan := safe_num(row.get('VLAN'))) is not None: entry['vlan'] = vlan
                if ipn := ws_clean(row.get('Адрес сети')): entry['ipnetwork'] = ipn
                for lan_col in ('Тип сети (проводная, беспроводная)', 'Тип LAN'):
                    if lan_col in row and (lan_type := ws_clean(row.get(lan_col))):
                        entry['lan_type'] = lan_type
                        break
            elif ntype == 'WAN':
                if wan := ws_clean(row.get('WAN Адрес')): entry['wan_ip'] = wan
            if prov := ws_clean(row.get('Провайдер')): entry['provider'] = prov
            if speed := safe_num(row.get('Скорость')): entry['bandwidth'] = speed
            segment_refs: List[str] = []
            for seg_col in ('Сетевой сегмент/зона(ID)', 'Сетевой сегмент/зона'):
                if seg_col in row:
                    segment_refs = parse_multiline_ids(row.get(seg_col))
                else:
                    segment_refs = []
                if segment_refs:
                    break
            if segment_refs:
                entry['segment'] = segment_refs
            if locations := parse_locations(row.get('Расположение')): entry['location'] = locations
            if vrf := (ws_clean(row.get('VRF  ')) or ws_clean(row.get('VRF'))): entry['VRF'] = vrf
            nets[nid] = entry

    # --- Auto-creation from Networks ---
    auto_created_from_nets = 0
    for nid, ndata in nets.items():
        net_locations = ndata.get('location', [])
        net_segment_ids = ndata.get('segment', [])

        if net_locations and len(net_segment_ids) == 1:
            primary_segment_id = net_segment_ids[0]
            if primary_segment_id not in segments:
                continue 

            primary_segment = segments[primary_segment_id]
            primary_segment_title = primary_segment.get('title')
            primary_segment_zone = (primary_segment.get('sber') or {}).get('zone')

            if not primary_segment_title or not primary_segment_zone:
                continue

            existing_locations_for_segment = {
                (s.get('sber') or {}).get('location')
                for s in segments.values()
                if s.get('title') == primary_segment_title and (s.get('sber') or {}).get('location')
            }

            company_prefix = nid.split('.')[0] if '.' in nid else ''
            for loc_id in net_locations:
                if loc_id not in existing_locations_for_segment:
                    loc_postfix = loc_id.split('.')[-1]
                    if loc_id.startswith(f'{company_prefix}.dc.'):
                        loc_postfix = f"dc{loc_postfix}"
                    
                    zone_slug = primary_segment_zone.lower().replace(' ', '_')
                    new_seg_id = f"{company_prefix}.network_segment.{zone_slug}.{loc_postfix}"

                    new_seg_title = f"{primary_segment_title}_{loc_postfix}"

                    target_list = ndata.setdefault('segment', [])

                    if new_seg_id in segments or new_seg_id in processed_ids:
                        if new_seg_id not in target_list:
                            target_list.append(new_seg_id)
                        existing_locations_for_segment.add(loc_id)
                        continue

                    new_segment = {
                        'title': new_seg_title,
                        'description': f'Автоматически созданный сегмент для расположения {loc_id}',
                        'sber': {
                            'location': loc_id,
                            'zone': primary_segment_zone
                        }
                    }
                    segments[new_seg_id] = new_segment
                    target_list.append(new_seg_id)
                    existing_locations_for_segment.add(loc_id)
                    processed_ids.add(new_seg_id)
                    auto_created_from_nets += 1
    
    auto_created_segments_count += auto_created_from_nets

    company_prefix_for_files = ''
    if nets:
        first_net_id = next(iter(nets.keys()))
        if '.' in first_net_id:
            company_prefix_for_files = first_net_id.split('.')[0]

    if segments:
        write_yaml(out_dir / 'network_segment.yaml', {'seaf.ta.services.network_segment': segments})
    if nets:
        write_networks_per_location(nets, out_dir, company_prefix_for_files)

    # --- Devices ---
    devices, processed_ids = {}, set()
    if 'Сетевые устройства' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Сетевые устройства'))
        for index, row in df.iterrows():
            did = id_clean(row.get('ID Устройства'))
            if not did:
                print(f"WARN: Skipping row {index + 2} in sheet 'Сетевые устройства' of '{xlsx_path.name}' due to missing ID.", file=sys.stderr)
                continue
            if did in processed_ids:
                print(f"WARN: Skipping duplicate ID '{did}' in sheet 'Сетевые устройства', row {index + 2} of '{xlsx_path.name}'.", file=sys.stderr)
                continue
            processed_ids.add(did)
            entry = {'title': ws_clean(row.get('Наименование')), 'realization_type': ws_clean(row.get('Тип реализации')), 'type': ws_clean(row.get('Тип')), 'model': ws_clean(row.get('Модель')), 'purpose': ws_clean(row.get('Назначение')), 'address': id_clean(row.get('IP адрес')), 'description': ws_clean(row.get('Описание'))}
            if seg := id_clean(row.get('Расположение (ID сегмента/зоны)')): entry['segment'] = seg
            if nets_list := to_list(row.get('Подключенные сети (список)')): entry['network_connection'] = nets_list
            devices[did] = entry
    if devices:
        write_yaml(out_dir / 'components_network.yaml', {'seaf.ta.components.network': devices})
    
    return auto_created_segments_count

def convert_kb_services(xlsx_path: Path, out_dir: Path):
    xls = read_excel(xlsx_path)
    if 'Сервисы КБ' not in xls.sheet_names:
        return
        
    kb_services, processed_ids = {}, set()
    df = non_empty_rows(xls.parse('Сервисы КБ'))
    for index, row in df.iterrows():
        svc_id = id_clean(row.get('ID КБ сервиса'))
        if not svc_id:
            print(f"WARN: Skipping row {index + 2} in sheet 'Сервисы КБ' of '{xlsx_path.name}' due to missing ID.", file=sys.stderr)
            continue
        svc_id = svc_id.rstrip(':;,')
        if svc_id in processed_ids:
            print(f"WARN: Skipping duplicate ID '{svc_id}' in sheet 'Сервисы КБ', row {index + 2} of '{xlsx_path.name}'.", file=sys.stderr)
            continue
        processed_ids.add(svc_id)
        
        service: Dict[str, Any] = {}
        if desc := ws_clean(row.get('Описание')): service['description'] = desc
        if tag := ws_clean(row.get('Tag')): service['tag'] = tag
        title = ws_clean(row.get('Название сервиса')) or ws_clean(row.get('Название')) or ws_clean(row.get('Технология'))
        if title: service['title'] = title
        if tech := ws_clean(row.get('Технология')): service['technology'] = tech
        if sname := ws_clean(row.get('Название ПО')): service['software_name'] = sname
        if status := ws_clean(row.get('Статус')): service['status'] = status
        networks = parse_multiline_ids(row.get('Подключенные сети'))
        if networks: service['network_connection'] = networks
        kb_services[svc_id] = service
    if kb_services:
        write_yaml(out_dir / 'kb.yaml', {'seaf.ta.services.kb': kb_services})

def write_root(out_dir: Path):
    imports = [p.name for p in sorted(out_dir.glob('*.yaml')) if p.name != 'root.yaml']
    if imports:
        write_yaml(out_dir / 'root.yaml', {'imports': imports})

# --- Conversion Functions: End ---

# --- Validation Functions: Start ---

def validate_refs(out_dir: Path) -> Dict[str, Any]:
    report: Dict[str, Any] = {'errors': [], 'warnings': []}
    
    def load_if_exists(name: str):
        path = out_dir / name
        if path.exists():
            return yaml.safe_load(path.read_text(encoding='utf-8')) or {}
        return {}

    try:
        regions = load_if_exists('dc_region.yaml').get('seaf.ta.services.dc_region', {})
        azs = load_if_exists('dc_az.yaml').get('seaf.ta.services.dc_az', {})
        dcs = load_if_exists('dc.yaml').get('seaf.ta.services.dc', {})
        offices = load_if_exists('office.yaml').get('seaf.ta.services.office', {})
        segments = load_if_exists('network_segment.yaml').get('seaf.ta.services.network_segment', {})
        devices = load_if_exists('components_network.yaml').get('seaf.ta.components.network', {})
        kb_services = load_if_exists('kb.yaml').get('seaf.ta.services.kb', {})
        
        networks: Dict[str, Any] = {}
        for p in sorted(out_dir.glob('networks_*.yaml')):
            networks.update(load_if_exists(p.name).get('seaf.ta.services.network', {}))

        (region_ids, az_ids, dc_ids, office_ids, seg_ids, net_ids) = (
            set(regions.keys()), set(azs.keys()), set(dcs.keys()), set(offices.keys()), set(segments.keys()), set(networks.keys())
        )
        
        # Existing checks
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
        for i, svc in kb_services.items():
            for net in svc.get('network_connection') or []:
                if net not in net_ids:
                    report['warnings'].append(f'KB service {i} refs missing Network {net}')
        for i, d in devices.items():
            if (s := d.get('segment')) and s not in seg_ids: report['errors'].append(f'Device {i} refs missing Segment {s}')
            for n in d.get('network_connection') or []:
                if n not in net_ids: report['errors'].append(f'Device {i} refs missing Network {n}')

        # New validation check: Network location vs. Segment locations
        for net_id, net_data in networks.items():
            if not (net_locs := net_data.get('location')):
                continue

            segment_locs = set()
            for seg_id in net_data.get('segment', []):
                if seg_id in segments:
                    segment = segments[seg_id]
                    if seg_loc := (segment.get('sber') or {}).get('location'):
                        segment_locs.add(seg_loc)

            for net_loc in net_locs:
                if net_loc not in segment_locs:
                    report['errors'].append(f"Network '{net_id}' is in location '{net_loc}', but none of its associated segments exist in that location.")

    except Exception as e:
        report['errors'].append(f"An unexpected error occurred during validation: {e}")
    return report

def validate_enums(out_dir: Path, report: Dict[str, Any]) -> None:
    def load_if_exists(name: str):
        path = out_dir / name
        if path.exists():
            return yaml.safe_load(path.read_text(encoding='utf-8')) or {}
        return {}
    try:
        devices = load_if_exists('components_network.yaml').get('seaf.ta.components.network', {})
        networks: Dict[str, Any] = {}
        for p in sorted(out_dir.glob('networks_*.yaml')):
            networks.update(load_if_exists(p.name).get('seaf.ta.services.network', {}))
        kb_services = load_if_exists('kb.yaml').get('seaf.ta.services.kb', {})
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
        status_allowed = {'Используется', 'Создается', 'Не используется', 'Выводится'}
        for i, svc in kb_services.items():
            if (status := svc.get('status')) and status not in status_allowed:
                report['warnings'].append(f'KB service {i} has unexpected status: {status}')
    except Exception as e:
        report['errors'].append(f"An unexpected error occurred during enum validation: {e}")

# --- Validation Functions: End ---

def main():
    try:
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
        
        # Clean output directory
        if out_dir.exists():
            for item in out_dir.glob('*'):
                if item.is_file():
                    item.unlink()
        out_dir.mkdir(parents=True, exist_ok=True)

        print("---", "Source XLSX Analysis", "---")
        source_counts = count_entities_in_xlsx(inputs)
        if not source_counts:
            print("No source entities found in specified XLSX files.")
        else:
            for entity, count in sorted(source_counts.items()):
                print(f"  - Found {count} entities in sheet for '{entity}'")

        processed_something = False
        total_auto_created_segments = 0

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
                    auto_segments = convert_segments_nets_devices(xlsx_path, out_dir)
                    total_auto_created_segments += auto_segments
                    processed_something = True

                if 'Сервисы КБ' in xls.sheet_names:
                    convert_kb_services(xlsx_path, out_dir)
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

        auto_created_segments_actual = max(dest_counts.get('network_segment', 0) - source_counts.get('network_segment', 0), 0)
        total_auto_created_segments = auto_created_segments_actual
        if auto_created_segments_actual > 0:
            print(f"\nINFO: Auto-created {auto_created_segments_actual} new network segments.")

        if not dest_counts:
            print("No destination entities were created.")
        else:
            for entity, count in sorted(dest_counts.items()):
                print(f"  - Created {count} entities for '{entity}'")

        print("\n---", "Conversion Summary", "---")
        all_keys = sorted(list(set(source_counts.keys()) | set(dest_counts.keys())))

        def format_dest_counts(src: int, dest: int) -> str:
            delta = dest - src
            if delta > 0:
                if src > 0:
                    return f"{src} + {delta} (= {dest})"
                return f"{dest} (+{delta})"
            if delta < 0 and src > 0:
                return f"{src} - {abs(delta)} (= {dest})"
            return str(dest)

        for key in all_keys:
            s_count = source_counts.get(key, 0)
            d_count = dest_counts.get(key, 0)
            delta = d_count - s_count
            status = "OK"

            if key in {'network_segment', 'network'}:
                if delta < 0:
                    status = "FAIL"
            elif s_count != d_count:
                status = "FAIL"

            dest_display = format_dest_counts(s_count, d_count)
            print(f"  - {key:<25} | Source: {s_count:<5} | Dest: {dest_display:<25} | {status}")

        if report['errors']:
            print("\nVALIDATION ERRORS:", file=sys.stderr)
            for e in report['errors']:
                print('-', e, file=sys.stderr)
            print(f'\nYAML written to: {out_dir} (with validation errors)', file=sys.stderr)
            sys.exit(0) # Exit with 0 even if there are validation errors, but report them
        else:
            print(f'\nYAML written to: {out_dir} (validation OK)')
            sys.exit(0)
    finally:
        print("Script finished.", file=sys.stderr)

if __name__ == '__main__':
    main()
