import sys
import re
import argparse
from pathlib import Path
from typing import List, Dict, Any
import math
import io


def ensure_deps():
    try:
        import pandas  # noqa: F401
        import yaml  # noqa: F401
    except Exception:
        # Try to install missing deps
        import subprocess
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', 'openpyxl', 'pyyaml'])


def read_excel(path: Path):
    import pandas as pd
    return pd.ExcelFile(path)


def non_empty_rows(df):
    # Drop fully empty rows
    return df.dropna(how='all')


def ws_clean(s: Any) -> Any:
    """Normalize whitespace in general text fields: strip and collapse internal spaces.
    Keeps spaces inside titles but removes leading/trailing and excessive spacing.
    """
    if s is None:
        return None
    # handle pandas NaN
    if isinstance(s, float) and math.isnan(s):
        return None
    s = str(s)
    if not s:
        return None
    # replace non-breaking space and control whitespace with normal space
    s = s.replace('\u00A0', ' ').replace('\xa0', ' ').replace('\t', ' ').replace('\n', ' ').replace('\r', ' ')
    # collapse multiple spaces
    s = re.sub(r'\s+', ' ', s)
    s = s.strip()
    # treat textual sentinels as empty
    if s.lower() in {"nan", "none", "null", "n/a", "na", ""}:
        return None
    return s or None


def id_clean(s: Any) -> str | None:
    """Normalize identifiers and references: remove all internal whitespace and trim.
    Do not alter punctuation except spaces.
    """
    s = ws_clean(s)
    if s is None:
        return None
    # remove all whitespace inside
    s = re.sub(r'\s+', '', s)
    return s or None


def to_list(val) -> List[str]:
    if val is None:
        return []
    # allow both list and CSV/semicolon-separated strings
    if isinstance(val, list):
        parts = []
        for x in val:
            t = id_clean(x)
            if t:
                parts.append(t)
        return parts
    s = str(val)
    # normalize whitespace first
    s = ws_clean(s) or ''
    # split by comma or semicolon
    raw = re.split(r'[;,]', s)
    out = []
    for p in raw:
        t = id_clean(p)
        if t:
            out.append(t)
    return out


def safe_num(v):
    try:
        if v is None or (isinstance(v, float) and str(v) == 'nan'):
            return None
        if isinstance(v, (int,)):
            return int(v)
        if isinstance(v, float):
            if v.is_integer():
                return int(v)
            return v
        # Try parse int
        s = str(v).strip()
        if not s:
            return None
        if s.isdigit():
            return int(s)
        return float(s)
    except Exception:
        return None


def write_yaml(path: Path, data: Dict[str, Any]):
    import yaml
    path.parent.mkdir(parents=True, exist_ok=True)
    with path.open('w', encoding='utf-8') as f:
        yaml.safe_dump(data, f, allow_unicode=True, sort_keys=False)


def convert_regions_az_dc_offices(xlsx_path: Path, out_dir: Path):
    import pandas as pd
    xls = read_excel(xlsx_path)

    # Регионы
    regions = {}
    if 'Регионы' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Регионы'))
        for _, row in df.iterrows():
            rid = id_clean(row.get('ID Региона', ''))
            if not rid:
                continue
            regions[rid] = {
                'description': ws_clean(row.get('Описание', None)),
                'external_id': rid.split('.')[-1] if rid else None,
                'title': ws_clean(row.get('Наименование', None)),
            }
    write_yaml(out_dir / 'dc_region.yaml', {'seaf.ta.services.dc_region': regions})

    # AZ
    azs = {}
    if 'AZ' in xls.sheet_names:
        df = non_empty_rows(xls.parse('AZ'))
        for _, row in df.iterrows():
            aid = id_clean(row.get('ID AZ', ''))
            if not aid:
                continue
            azs[aid] = {
                'description': ws_clean(row.get('Описание', None)),
                'external_id': aid.split('.')[-1] if aid else None,
                'region': id_clean(row.get('Регион', '')),
                'title': ws_clean(row.get('Наименование', None)),
                'vendor': ws_clean(row.get('Поставщик', None)),
            }
    write_yaml(out_dir / 'dc_az.yaml', {'seaf.ta.services.dc_az': azs})

    # DC
    dcs = {}
    if 'DC' in xls.sheet_names:
        df = non_empty_rows(xls.parse('DC'))
        for _, row in df.iterrows():
            did = id_clean(row.get('ID DC', ''))
            if not did:
                continue
            dcs[did] = {
                'address': ws_clean(row.get('Адрес', None)),
                'availabilityzone': id_clean(row.get('AZ', '')),
                'description': ws_clean(row.get('Описание', None)),
                'external_id': did.split('.')[-1] if did else None,
                'ownership': ws_clean(row.get('Форма владения', None)),
                'rack_qty': safe_num(row.get('Кол-во стоек', None)),
                'tier': ws_clean(row.get('Tier', '')),
                'title': ws_clean(row.get('Наименование', None)),
                'type': ws_clean(row.get('Тип', None)),
                'vendor': ws_clean(row.get('Поставщик', None)),
            }
    write_yaml(out_dir / 'dc.yaml', {'seaf.ta.services.dc': dcs})

    # Офисы
    offices = {}
    if 'Офисы' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Офисы'))
        for _, row in df.iterrows():
            oid = id_clean(row.get('ID Офиса', ''))
            if not oid:
                continue
            offices[oid] = {
                'address': ws_clean(row.get('Адрес', None)),
                'description': ws_clean(row.get('Описание', None)),
                'external_id': oid.split('.')[-1] if oid else None,
                'region': id_clean(row.get('Регион', '')),
                'title': ws_clean(row.get('Наименование', None)),
            }
    write_yaml(out_dir / 'office.yaml', {'seaf.ta.services.office': offices})


def simplify_location_token(loc: str) -> str:
    """Return a concise token for filenames based on location id."""
    # flix.dc.01 -> dc01
    m = re.match(r'^flix\.dc\.(\d+)$', loc)
    if m:
        return f'dc{m.group(1)}'
    # flix.office.hq -> office_hq
    m = re.match(r'^flix\.office\.(.+)$', loc)
    if m:
        return f'office_{m.group(1)}'
    # fallback: sanitize
    return re.sub(r'[^A-Za-z0-9]+', '_', loc).strip('_') or 'loc'


def write_networks_per_location(nets: Dict[str, Any], out_dir: Path) -> List[str]:
    """Split networks into files by each location listed; duplicate entries across files if multi-located.

    Returns list of filenames written relative to out_dir.
    """
    per_loc: Dict[str, Dict[str, Any]] = {}
    misc: Dict[str, Any] = {}
    for nid, entry in nets.items():
        locs = entry.get('location') or []
        if not locs:
            misc[nid] = entry
            continue
        for loc in locs:
            token = simplify_location_token(str(loc))
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
    import pandas as pd
    xls = read_excel(xlsx_path)

    # Сегменты
    segments = {}
    if 'Сегменты' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Сегменты'))
        for _, row in df.iterrows():
            sid = id_clean(row.get('ID сетевые сегмента/зоны', ''))
            if not sid:
                continue
            seg: Dict[str, Any] = {
                'title': ws_clean(row.get('Наименование', None)),
                'description': ws_clean(row.get('Описание', None)),
            }
            loc = id_clean(row.get('Расположение', ''))
            zone = ws_clean(row.get('Зона', None))
            if loc:
                seg['sber'] = {'location': loc}
                if zone:
                    seg['sber']['zone'] = zone
            elif zone:
                seg['sber'] = {'zone': zone}
            segments[sid] = seg
    write_yaml(out_dir / 'network_segment.yaml', {'seaf.ta.services.network_segment': segments})

    # Сети
    nets = {}
    if 'Сети' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Сети'))
        for _, row in df.iterrows():
            nid = id_clean(row.get('ID Network', ''))
            if not nid:
                continue
            entry: Dict[str, Any] = {}
            entry['title'] = ws_clean(row.get('Наименование', None))
            entry['description'] = ws_clean(row.get('Описание', None))
            ntype = ws_clean(row.get('Тип сети', None))
            if ntype:
                entry['type'] = ntype
            # Fields by type
            if ntype == 'LAN':
                vlan = safe_num(row.get('VLAN', None))
                if vlan is not None:
                    entry['vlan'] = vlan
                ipn = ws_clean(row.get('Адрес сети', None))
                if ipn:
                    entry['ipnetwork'] = ipn
                lan_type = ws_clean(row.get('Тип сети (проводная, беспроводная)', None))
                if lan_type:
                    entry['lan_type'] = lan_type
            elif ntype == 'WAN':
                wan = ws_clean(row.get('WAN Адрес', None))
                if wan:
                    entry['wan_ip'] = wan
            # Optional provider (preserve across roundtrip)
            prov = ws_clean(row.get('Провайдер', None))
            if prov:
                entry['provider'] = prov
            speed = safe_num(row.get('Скорость', None))
            if speed is not None:
                entry['bandwidth'] = speed

            segs = to_list(row.get('Сетевой сегмент/зона(ID)', ''))
            if segs:
                entry['segment'] = segs
            locations = to_list(row.get('Расположение', ''))
            if locations:
                entry['location'] = locations
            # Map WAN speed to schema 'bandwidth'; VRF to schema 'VRF'
            if ntype == 'WAN':
                speed_raw = row.get('Скорость', None)
                speed = safe_num(speed_raw if speed_raw is not None else None)
                if speed is not None and not (isinstance(speed, float) and math.isnan(speed)):
                    entry['bandwidth'] = speed
            vrf = ws_clean(row.get('VRF  ', None)) or ws_clean(row.get('VRF', None))
            if vrf:
                entry['VRF'] = vrf
            nets[nid] = entry
    # Split networks by location to multiple files
    network_files = write_networks_per_location(nets, out_dir)

    # Сетевые устройства
    devices = {}
    if 'Сетевые устройства' in xls.sheet_names:
        df = non_empty_rows(xls.parse('Сетевые устройства'))
        for _, row in df.iterrows():
            did = id_clean(row.get('ID Устройства', ''))
            if not did:
                continue
            entry: Dict[str, Any] = {
                'title': ws_clean(row.get('Наименование', None)),
                'realization_type': ws_clean(row.get('Тип реализации', None)),
                'type': ws_clean(row.get('Тип', None)),
                'model': ws_clean(row.get('Модель', None)),
                'purpose': ws_clean(row.get('Назначение', None)),
                'address': id_clean(row.get('IP адрес', '')),
                'description': ws_clean(row.get('Описание', None)),
            }
            seg = id_clean(row.get('Расположение (ID сегмента/зоны)', ''))
            if seg:
                entry['segment'] = seg
            nets_list = to_list(row.get('Подключенные сети (список)', None))
            if nets_list:
                entry['network_connection'] = nets_list
            devices[did] = entry
    write_yaml(out_dir / 'components_network.yaml', {'seaf.ta.components.network': devices})
    return {
        'network_files': network_files,
    }


def validate_refs(out_dir: Path) -> Dict[str, Any]:
    """Validate referential integrity across generated YAML dictionaries.

    Returns a report dict with lists of errors and warnings.
    """
    import yaml
    report: Dict[str, Any] = {'errors': [], 'warnings': []}

    def load(name: str):
        with (out_dir / name).open('r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}

    regions = load('dc_region.yaml').get('seaf.ta.services.dc_region', {})
    azs = load('dc_az.yaml').get('seaf.ta.services.dc_az', {})
    dcs = load('dc.yaml').get('seaf.ta.services.dc', {})
    offices = load('office.yaml').get('seaf.ta.services.office', {})
    segments = load('network_segment.yaml').get('seaf.ta.services.network_segment', {})
    # Load all networks_*.yaml and merge
    networks: Dict[str, Any] = {}
    for p in sorted(out_dir.glob('networks_*.yaml')):
        nd = load(p.name).get('seaf.ta.services.network', {})
        networks.update(nd)
    devices = load('components_network.yaml').get('seaf.ta.components.network', {})

    region_ids = set(regions.keys())
    az_ids = set(azs.keys())
    dc_ids = set(dcs.keys())
    office_ids = set(offices.keys())
    seg_ids = set(segments.keys())
    net_ids = set(networks.keys())

    # AZ → Region
    for aid, a in azs.items():
        r = (a or {}).get('region')
        if r and r not in region_ids:
            report['errors'].append(f'AZ {aid} references missing Region {r}')

    # DC → AZ
    for did, d in dcs.items():
        az = (d or {}).get('availabilityzone')
        if az and az not in az_ids:
            report['errors'].append(f'DC {did} references missing AZ {az}')

    # Office → Region
    for oid, o in offices.items():
        r = (o or {}).get('region')
        if r and r not in region_ids:
            report['errors'].append(f'Office {oid} references missing Region {r}')

    # Segment → location (DC|Office) in sber.location
    for sid, s in segments.items():
        loc = ((s or {}).get('sber') or {}).get('location')
        if loc and (loc not in dc_ids and loc not in office_ids):
            report['errors'].append(f'Segment {sid} has unknown location {loc} (expect DC or Office ID)')

    # Network → segments[] and locations[] (each must exist)
    for nid, n in networks.items():
        for seg in (n or {}).get('segment') or []:
            if seg not in seg_ids:
                report['errors'].append(f'Network {nid} references missing Segment {seg}')
        for loc in (n or {}).get('location') or []:
            if loc not in dc_ids and loc not in office_ids:
                report['errors'].append(f'Network {nid} has unknown location {loc} (expect DC or Office ID)')

    # Device → segment and network_connection[]
    for did, d in devices.items():
        seg = (d or {}).get('segment')
        if seg and seg not in seg_ids:
            report['errors'].append(f'Device {did} references missing Segment {seg}')
        for net in (d or {}).get('network_connection') or []:
            if net not in net_ids:
                report['errors'].append(f'Device {did} references missing Network {net}')

    # Save report
    out_report = out_dir / 'validation_report.txt'
    with out_report.open('w', encoding='utf-8') as f:
        if report['errors']:
            f.write('ERRORS:\n')
            for e in report['errors']:
                f.write(f'- {e}\n')
        else:
            f.write('No referential errors found.\n')
        if report['warnings']:
            f.write('\nWARNINGS:\n')
            for w in report['warnings']:
                f.write(f'- {w}\n')
    return report


def validate_enums(out_dir: Path, report: Dict[str, Any]) -> None:
    """Validate enum-constrained fields based on available schemas (subset relevant to current XLSX).

    Mutates report by appending to errors/warnings.
    """
    import yaml
    def load(name: str):
        with (out_dir / name).open('r', encoding='utf-8') as f:
            return yaml.safe_load(f) or {}

    # Collect networks from all files
    networks: Dict[str, Any] = {}
    for p in sorted(out_dir.glob('networks_*.yaml')):
        d = load(p.name).get('seaf.ta.services.network', {}) or {}
        networks.update(d)
    devices = load('components_network.yaml').get('seaf.ta.components.network', {}) or {}

    # Network enums and type-specific requireds (per schema oneOf)
    for nid, n in networks.items():
        ntype = n.get('type')
        if ntype is not None and ntype not in ('LAN', 'WAN'):
            report['errors'].append(f'Network {nid} has invalid type enum: {ntype} (expected LAN|WAN)')
        if ntype == 'LAN':
            ltype = n.get('lan_type')
            if ltype is not None and ltype not in ('Проводная', 'Беспроводная'):
                report['errors'].append(f'Network {nid} has invalid lan_type enum: {ltype} (expected Проводная|Беспроводная)')
            # Required by schema: lan_type, ipnetwork
            if not n.get('lan_type'):
                report['errors'].append(f'Network {nid} missing required field lan_type for LAN')
            if not n.get('ipnetwork'):
                report['errors'].append(f'Network {nid} missing required field ipnetwork for LAN')
        if ntype == 'WAN':
            # Required by schema: wan_ip, bandwidth, autonomus_system
            if not n.get('wan_ip'):
                report['errors'].append(f'Network {nid} missing required field wan_ip for WAN')
            if n.get('bandwidth') is None:
                report['errors'].append(f'Network {nid} missing required field bandwidth for WAN')
            # autonomus_system is not required per current policy

    # Network device enums
    dev_type_allowed = {'Маршрутизатор', 'МСЭ', 'Контроллер WiFi', 'Криптошлюз', 'VPN', 'NAT', 'Коммутатор'}
    realization_allowed = {'Виртуальный', 'Физический'}
    for did, d in devices.items():
        rtype = d.get('realization_type')
        if rtype is not None and rtype not in realization_allowed:
            report['errors'].append(f'Device {did} has invalid realization_type enum: {rtype} (expected one of {sorted(realization_allowed)})')
        dtype = d.get('type')
        if dtype is not None and dtype not in dev_type_allowed:
            report['errors'].append(f'Device {did} has invalid type enum: {dtype} (expected one of {sorted(dev_type_allowed)})')

    # Validate segment zone enum (if present) using DZO extension schema
    segments = load('network_segment.yaml').get('seaf.ta.services.network_segment', {}) or {}
    def allowed_segment_zones() -> List[str]:
        p = Path('seaf-core/dzo_entities.yaml')
        if not p.exists():
            return []
        text = p.read_text(encoding='utf-8', errors='ignore')
        idx = text.find('seaf.ta.services.network_segment:')
        if idx == -1:
            return []
        sub = text[idx:]
        # find the first occurrence of 'zone:' followed by 'enum:'
        zidx = sub.find('\n')
        # narrow search window
        window = sub[:1000]
        # more direct scanning
        lines = window.splitlines()
        values: List[str] = []
        capture = False
        for i, line in enumerate(lines):
            if ' zone:' in line:
                # look ahead for enum listing
                for j in range(i+1, min(i+30, len(lines))):
                    l2 = lines[j]
                    if ' enum:' in l2:
                        capture = True
                        continue
                    if capture:
                        s = l2.strip()
                        if s.startswith('- '):
                            values.append(s[2:].strip())
                        else:
                            capture = False
                            break
        return values

    zones_allowed = set(allowed_segment_zones())
    if zones_allowed:
        for sid, s in segments.items():
            sber = (s or {}).get('sber') or {}
            zone = sber.get('zone')
            if zone and zone not in zones_allowed:
                report['errors'].append(f'Segment {sid} has invalid zone: {zone} (allowed: {sorted(zones_allowed)})')
            if not zone:
                report['warnings'].append(f'Segment {sid} has no zone set; allowed: {sorted(zones_allowed)}')


def write_root(out_dir: Path):
    imports = [
        'dc_region.yaml',
        'dc_az.yaml',
        'dc.yaml',
        'office.yaml',
        'network_segment.yaml',
        'components_network.yaml',
    ]
    # add network files (sorted for stability)
    imports.extend(sorted([p.name for p in out_dir.glob('networks_*.yaml')]))
    write_yaml(out_dir / 'root.yaml', {'imports': imports})


def main():
    ensure_deps()
    parser = argparse.ArgumentParser(description='Convert XLSX to YAML with referential validation')
    parser.add_argument('--config', type=str, help='Path to YAML config with inputs and output dir')
    args = parser.parse_args()

    # Defaults
    inputs = [
        'regions_az_dc_offices_v1.0.xlsx',
        'segments_nets_netdevices_v1.0.xlsx',
    ]
    out_dir = Path('scripts') / 'out_yaml'

    # Load config if provided
    if args.config:
        import yaml
        with open(args.config, 'r', encoding='utf-8') as f:
            cfg = yaml.safe_load(f) or {}
        # expect list of xlsx files
        inputs = cfg.get('xlsx_files') or cfg.get('inputs') or inputs
        out_dir = Path(cfg.get('out_yaml_dir') or cfg.get('out_dir') or out_dir)

    out_dir.mkdir(parents=True, exist_ok=True)

    # Run conversions depending on presence of expected files
    # regions + az + dc + offices
    reg_file = next((Path(p) for p in inputs if Path(p).name.lower().startswith('regions_az_dc_offices')), None)
    if reg_file and reg_file.exists():
        convert_regions_az_dc_offices(reg_file, out_dir)
    else:
        print('WARN: regions_az_dc_offices*.xlsx not provided or missing; skipping')

    seg_file = next((Path(p) for p in inputs if Path(p).name.lower().startswith('segments_nets_netdevices')), None)
    if seg_file and seg_file.exists():
        convert_segments_nets_devices(seg_file, out_dir)
    else:
        print('WARN: segments_nets_netdevices*.xlsx not provided or missing; skipping')

    write_root(out_dir)
    report = validate_refs(out_dir)
    validate_enums(out_dir, report)
    # Console report: both errors and warnings
    if report['errors']:
        print('VALIDATION ERRORS:', file=sys.stderr)
        for e in report['errors']:
            print('-', e, file=sys.stderr)
    if report['warnings']:
        print('VALIDATION WARNINGS:')
        for w in report['warnings']:
            print('-', w)
    if report['errors']:
        print(f'YAML written to: {out_dir} (with validation errors)', file=sys.stderr)
        print(f'Details: {out_dir / "validation_report.txt"}', file=sys.stderr)
        sys.stderr.flush()
        sys.exit(2)
    else:
        print(f'YAML written to: {out_dir} (validation OK)')


if __name__ == '__main__':
    main()
