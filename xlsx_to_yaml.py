import sys
import re
import argparse
from pathlib import Path
from typing import List, Dict, Any, Tuple, Set
from copy import deepcopy
import math
import yaml
import pandas as pd

DEBUG_LOG_FILE = Path('debug_script.log')

# --- Schema Definitions ---
# Mandatory: Script cannot function/identify objects without these.
# Optional: Script runs, but attributes will be missing.
SCHEMA_DEF = {
    'Регионы': {
        'mandatory': ['ID Региона'],
        'optional': ['Наименование', 'Описание']
    },
    'AZ': {
        'mandatory': ['ID AZ'],
        'optional': ['Наименование', 'Описание', 'Регион', 'Поставщик']
    },
    'DC': {
        'mandatory': ['ID DC'],
        'optional': ['Наименование', 'Описание', 'Адрес', 'AZ', 'Кол-во стоек', 'Tier']
    },
    'Офисы': {
        'mandatory': ['ID Офиса'],
        'optional': ['Наименование', 'Описание', 'Адрес', 'Регион']
    },
    'Сегменты': {
        'mandatory': ['ID сетевые сегмента/зоны'],
        'optional': ['Наименование', 'Описание', 'Расположение', 'Зона']
    },
    'Сети': {
        'mandatory': ['ID Network'],
        'optional': ['Наименование', 'Описание', 'Тип сети', 'Расположение', 'Сетевой сегмент/зона', 'VLAN', 'Адрес сети']
    },
    'Сетевые устройства': {
        'mandatory': ['ID Устройства'],
        'optional': ['Наименование', 'Тип устройства', 'Расположение', 'Подключенные сети', 'Описание', 'IP адрес']
    },
    'Сервисы КБ': {
        'mandatory': ['ID КБ сервиса'],
        'optional': ['Название сервиса', 'Описание', 'Статус', 'Технология', 'Название ПО', 'Подключенные сети']
    },
    'Тех. сервисы': {
        'mandatory': ['Идентификатор', 'Класс'],
        'optional': ['Наименование', 'Описание', 'Тип сервиса', 'Тип резервирования', 'Подключен к сети', 'ЦОД']
    },
    'Компоненты': {
        'mandatory': ['Идентификатор', 'Класс'],
        'optional': ['Наименование', 'Описание', 'Тип', 'Локация', 'Сети', 'Сегмент']
    }
}

SHEET_ALIASES = {
    'Tech Services': 'Тех. сервисы',
    'Сетевые устройства': 'Сетевые устройства', 
    '??????? ??????????': 'Сетевые устройства' # Encoding artifacts handling
}

def log_debug(message):
    try:
        with DEBUG_LOG_FILE.open('a', encoding='utf-8') as f:
            f.write(f"{message}\n")
    except Exception: pass

SPECIAL_ENTITY_MAP = {
    'Мониторинг': 'monitoring',
    'Логгирование': 'monitoring',
    'Резервное копирование': 'backup',
    'Бекапирование и восстановление данных': 'backup'
}

SVC_TYPE_MAP = {
    'E-mail': 'Коммуникации (АТС, Почта, мессенджеры, СМС шлюзы и т.д.)',
    'Реляционные СУБД': 'СУБД',
    'Системы управления сетевым адресным пространством (DHCP, DNS т.д.)': 'Управление сетевым адресным пространством (DHCP, DNS и т.д.)',
    'Удаленный доступ': 'Инфраструктура удаленного доступа',
    'Файловое хранилище': 'Файловый ресурс (FTP, NFS, SMB, S3 и т.д.)',
    'Управление конфигурациями': 'Управление и автоматизацией (Ansible, Terraform, Jenkins и т.д.)',
    'Виртуализация': 'Управление ИТ-службой, ИТ-инфраструктурой и ИТ-активами (CMDB, ITSM и т.д.)',
    'Управление облачной инфраструктурой': 'Управление ИТ-службой, ИТ-инфраструктурой и ИТ-активами (CMDB, ITSM и т.д.)',
    'Иное': 'Серверы приложений и т.д.',
    'Мониторинг': 'Серверы приложений и т.д.',
    'Логгирование': 'Серверы приложений и т.д.'
}

# --- Validation Logic ---

def normalize_sheet_name(name: str) -> str:
    return SHEET_ALIASES.get(name, name)

def validate_structure(xlsx_path: Path, force: bool) -> bool:
    print(f"\n[CHECK] Analyzing structure: {xlsx_path.name}")
    try:
        xls = pd.ExcelFile(xlsx_path)
    except Exception as e:
        print(f"[FATAL] Cannot open file: {e}")
        return False

    issues_found = False
    critical_missing = False
    
    found_sheets = set(normalize_sheet_name(s) for s in xls.sheet_names)
    
    # Simple check: does it have ANY relevant sheets?
    relevant_sheets = set(SCHEMA_DEF.keys())
    if not found_sheets.intersection(relevant_sheets):
        print(f"  [!!] No recognized SEAF sheets found. (Expected one of: {', '.join(relevant_sheets)})")
        return False if not force else True

    for sheet_original in xls.sheet_names:
        sheet_norm = normalize_sheet_name(sheet_original)
        if sheet_norm not in SCHEMA_DEF:
            continue
            
        print(f"  Sheet '{sheet_original}' (mapped to '{sheet_norm}'):")
        df = xls.parse(sheet_original).dropna(how='all')
        cols = set(df.columns)
        
        required = SCHEMA_DEF[sheet_norm]['mandatory']
        optional = SCHEMA_DEF[sheet_norm]['optional']
        
        # Check Mandatory
        missing_req = [c for c in required if c not in cols]
        if missing_req:
            print(f"    [CRITICAL] Missing MANDATORY columns: {missing_req}")
            critical_missing = True
            issues_found = True
        
        # Check Optional
        missing_opt = [c for c in optional if c not in cols]
        if missing_opt:
            print(f"    [WARN] Missing optional columns: {missing_opt}")
            issues_found = True
            
        if not missing_req and not missing_opt:
            print("    [OK] All expected columns found.")

    if critical_missing and not force:
        print("\n[STOP] Critical columns are missing. Unable to proceed reliably.")
        return False
        
    if issues_found and not force:
        choice = input("\n[?] Structural issues found. Continue anyway? [y/N]: ").strip().lower()
        if choice != 'y':
            return False
            
    return True

class DataValidator:
    def __init__(self):
        self.seen_ids: Dict[str, str] = {} # id -> sheet/type provenance
        self.known_networks: Set[str] = set()
        self.known_locations: Set[str] = set()
        self.errors: List[str] = []
        self.warnings: List[str] = []

    def register_id(self, entity_id: str, context: str):
        if not entity_id: return
        if entity_id in self.seen_ids:
            self.errors.append(f"Duplicate ID '{entity_id}' found in {context}. Previous occurence in {self.seen_ids[entity_id]}")
        else:
            self.seen_ids[entity_id] = context

    def register_network(self, net_id: str):
        if net_id: self.known_networks.add(net_id)

    def register_location(self, loc_id: str):
        if loc_id: self.known_locations.add(loc_id)

    def check_ref_network(self, net_refs: List[str], owner_id: str):
        for net in net_refs:
            if net and net not in self.known_networks:
                # We warn, not error, because networks might be in a file not yet processed if doing multi-file
                self.warnings.append(f"Object '{owner_id}' references unknown network '{net}'")

    def report(self):
        if not self.errors and not self.warnings:
            return
        print("\n--- Data Validation Report ---")
        for err in self.errors:
            print(f"  [DATA ERROR] {err}")
        for warn in self.warnings[:10]: # Limit warnings to avoid spam
            print(f"  [DATA WARN]  {warn}")
        if len(self.warnings) > 10:
            print(f"  ... and {len(self.warnings) - 10} more warnings.")

# Global Validator Instance
VALIDATOR = DataValidator()

def normalize_svc_type(val):
    if not val: return 'Серверы приложений и т.д.'
    trans = str.maketrans('CAEOPXy', 'САЕОРХу')
    def norm(s): return str(s).translate(trans).strip()
    val_norm = norm(val)
    allowed = ["Управление ИТ-службой, ИТ-инфраструктурой и ИТ-активами (CMDB, ITSM и т.д.)", "Управление и автоматизацией (Ansible, Terraform, Jenkins и т.д.)", "Управление разработкой и хранения кода (Gitlab, Jira и т.д.)", "Управление сетевым адресным пространством (DHCP, DNS и т.д.)", "Виртуализация рабочих мест (ВАРМ и VDI)", "Шлюз, Балансировщик, прокси", "СУБД", "Распределенный кэш", "Интеграционная шина  (MQ, ETL, API)", "Файловый ресурс (FTP, NFS, SMB, S3 и т.д.)", "Инфраструктура удаленного доступа", "Коммуникации (АТС, Почта, мессенджеры, СМС шлюзы и т.д.)", "Серверы приложений и т.д."]
    for item in allowed:
        if norm(item) == val_norm: return item
    return SVC_TYPE_MAP.get(val, 'Серверы приложений и т.д.')

def derive_location_from_network(net_id):
    if not net_id: return None
    prefix = net_id.split('.')[0] if '.' in net_id else 'seaf'
    m_dc = re.search(r'\.dc(\d+)', net_id)
    if m_dc: return f"{prefix}.dc.{m_dc.group(1)}"
    m_off = re.search(r'\.office\.([a-zA-Z0-9-]+)', net_id)
    if m_off:
        parts = net_id.split('.')
        try:
            idx = parts.index('office')
            if idx+1 < len(parts) and parts[idx+1] not in ['lan','wan']:
                return f"{prefix}.office.{parts[idx+1]}"
            return f"{prefix}.office"
        except ValueError: pass
    return None

def ensure_deps():
    try: import pandas, yaml, openpyxl
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', 'openpyxl', 'pyyaml'])

def read_excel(path: Path):
    if not path.exists(): raise FileNotFoundError(f"Excel file not found: {path}")
    try: return pd.ExcelFile(path)
    except Exception as e: raise RuntimeError(f"Failed to open Excel {path.name}: {e}")

def non_empty_rows(df): return df.dropna(how='all')

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
    return re.sub(r'\s+', '', s) if s else None

def parse_multiline_ids(val) -> List[str]:
    if val is None: return []
    if isinstance(val, list): return [t for x in val if (t := id_clean(x))]
    text = str(val).replace('\r', '\n')
    tokens = []
    for line in text.split('\n'):
        line = line.strip()
        if not line: continue
        if line.startswith('-'): line = line[1:].strip()
        for piece in re.split(r'[;,]', line):
            if cleaned := id_clean(piece): tokens.append(cleaned)
    return tokens

def parse_locations(val: Any) -> List[str]:
    if val is None: return []
    s = ws_clean(str(val))
    return [t for p in re.split(r'[;,\s]+', s) if (t := id_clean(p))] if s else []

class IndentedDumper(yaml.SafeDumper):
    def increase_indent(self, flow=False, indentless=False): return super(IndentedDumper, self).increase_indent(flow, False)

def write_yaml(path: Path, data: Dict[str, Any]):
    try:
        path.parent.mkdir(parents=True, exist_ok=True)
        with path.open('w', encoding='utf-8') as f:
            yaml.dump(sanitize_for_yaml(data), f, Dumper=IndentedDumper, allow_unicode=True, sort_keys=False)
    except Exception as e: print(f"ERROR: Failed to write YAML to {path}: {e}", file=sys.stderr)

def sanitize_for_yaml(value: Any) -> Any:
    if isinstance(value, dict): return {k: sanitize_for_yaml(v) for k, v in value.items() if not k.startswith('_')}
    if isinstance(value, list): return [sanitize_for_yaml(v) for v in value]
    if isinstance(value, str): return re.sub(r'[\r\n]+', ' ', value)
    return value

def count_entities_in_xlsx(xlsx_files: List[Path]) -> Dict[str, int]:
    counts = {}
    sheet_map = {'Регионы': 'dc_region', 'AZ': 'dc_az', 'DC': 'dc', 'Офисы': 'office', 'Сегменты': 'network_segment', 'Сети': 'network', 'Сетевые устройства': 'components.network', 'Сервисы КБ': 'kb', 'Тех. сервисы': 'tech_services', 'Tech Services': 'tech_services'}
    for file_path in xlsx_files:
        if not file_path.exists(): continue
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                if sheet_name in sheet_map:
                    df = non_empty_rows(xls.parse(sheet_name))
                    ename = sheet_map[sheet_name]
                    if ename == 'tech_services':
                        typed_ids = {}
                        for _, row in df.iterrows():
                            oid = id_clean(row.get('Идентификатор'))
                            if not oid: continue
                            svc_raw, res_val, cls_val = ws_clean(row.get('Тип сервиса')) or ws_clean(row.get('Класс')), ws_clean(row.get('Тип резервирования')), ws_clean(row.get('Класс'))
                            etype = 'compute_service'
                            if svc_raw in SPECIAL_ENTITY_MAP: etype = SPECIAL_ENTITY_MAP[svc_raw]
                            elif cls_val == 'Cluster' or (res_val and res_val.lower() in ['active-active', 'active-passive', 'n+1', 'да']): etype = 'cluster'
                            elif cls_val == 'Software': etype = 'software'
                            elif cls_val == 'Storage': etype = 'storage'
                            elif cls_val == 'Monitoring': etype = 'monitoring'
                            elif cls_val == 'Backup': etype = 'backup'
                            elif cls_val == 'Compute Service': etype = 'compute_service'
                            typed_ids.setdefault(etype, set()).add(oid)
                        for etype, ids in typed_ids.items():
                            counts[etype] = counts.get(etype, 0) + len(ids)
                    else: counts[ename] = counts.get(ename, 0) + len(df)
        except Exception: pass
    return counts

def count_entities_in_yaml_dir(yaml_dir: Path) -> Dict[str, int]:
    counts = {}
    if not yaml_dir.exists(): return counts
    for p in sorted(yaml_dir.glob('**/*.yaml')):
        try:
            with p.open('r', encoding='utf-8') as f:
                data = yaml.safe_load(f)
                if not isinstance(data, dict): continue
                for key, val in data.items():
                    if isinstance(val, dict):
                        ename = key.replace('seaf.ta.services.', '').replace('seaf.ta.components.', 'components.').replace('seaf.company.ta.services.', '').replace('seaf.company.ta.components.', 'components.')
                        counts[ename] = counts.get(ename, 0) + len(val)
        except Exception: continue
    return counts

def convert_regions_az_dc_offices(xls, out_dir: Path):
    reg, azs, dcs, off = {}, {}, {}, {}
    if 'Регионы' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('Регионы')).iterrows():
            rid = id_clean(row.get('ID Региона'))
            if rid: 
                reg[rid] = {'description': ws_clean(row.get('Описание')), 'external_id': rid.split('.')[-1], 'title': ws_clean(row.get('Наименование'))}
                VALIDATOR.register_id(rid, "Регионы")
                VALIDATOR.register_location(rid)
    if reg: write_yaml(out_dir / 'dc_region.yaml', {'seaf.ta.services.dc_region': reg})
    
    if 'AZ' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('AZ')).iterrows():
            aid = id_clean(row.get('ID AZ'))
            if aid: 
                azs[aid] = {'description': ws_clean(row.get('Описание')), 'external_id': aid.split('.')[-1], 'region': id_clean(row.get('Регион')), 'title': ws_clean(row.get('Наименование')), 'vendor': ws_clean(row.get('Поставщик'))}
                VALIDATOR.register_id(aid, "AZ")
                VALIDATOR.register_location(aid)
    if azs: write_yaml(out_dir / 'dc_az.yaml', {'seaf.ta.services.dc_az': azs})
    
    if 'DC' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('DC')).iterrows():
            did = id_clean(row.get('ID DC'))
            if did: 
                dcs[did] = {'address': ws_clean(row.get('Адрес')), 'availabilityzone': id_clean(row.get('AZ')), 'description': ws_clean(row.get('Описание')), 'external_id': did.split('.')[-1], 'ownership': ws_clean(row.get('Форма владения')), 'rack_qty': ws_clean(row.get('Кол-во стоек')), 'tier': ws_clean(row.get('Tier')), 'title': ws_clean(row.get('Наименование')), 'type': ws_clean(row.get('Тип')), 'vendor': ws_clean(row.get('Поставщик'))}
                VALIDATOR.register_id(did, "DC")
                VALIDATOR.register_location(did)
    if dcs: write_yaml(out_dir / 'dc.yaml', {'seaf.ta.services.dc': dcs})
    
    if 'Офисы' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('Офисы')).iterrows():
            oid = id_clean(row.get('ID Офиса'))
            if oid: 
                off[oid] = {'address': ws_clean(row.get('Адрес')), 'description': ws_clean(row.get('Описание')), 'external_id': oid.split('.')[-1], 'region': id_clean(row.get('Регион')), 'title': ws_clean(row.get('Наименование'))}
                VALIDATOR.register_id(oid, "Офисы")
                VALIDATOR.register_location(oid)
    if off: write_yaml(out_dir / 'office.yaml', {'seaf.ta.services.office': off})
    return (len(reg)+len(azs)+len(dcs)+len(off)) > 0

def convert_segments_nets_devices(xls, out_dir: Path):
    res = False
    if 'Сегменты' in xls.sheet_names:
        segments = {}
        for _, r in non_empty_rows(xls.parse('Сегменты')).iterrows():
            sid = id_clean(r.get('ID сетевые сегмента/зоны'))
            if sid:
                segments[sid] = {'title': ws_clean(r.get('Наименование')), 'description': ws_clean(r.get('Описание')), 'sber': {'location': parse_locations(r.get('Расположение'))[0] if parse_locations(r.get('Расположение')) else None, 'zone': ws_clean(r.get('Зона'))}}
                VALIDATOR.register_id(sid, "Сегменты")
        if segments: write_yaml(out_dir / 'network_segment.yaml', {'seaf.ta.services.network_segment': segments}); res = True
    
    if 'Сети' in xls.sheet_names:
        nets = {}
        for _, r in non_empty_rows(xls.parse('Сети')).iterrows():
            nid = id_clean(r.get('ID Network'))
            if not nid: continue
            VALIDATOR.register_id(nid, "Сети")
            VALIDATOR.register_network(nid)
            
            ntype = ws_clean(r.get('Тип сети'))
            entry = {'title': ws_clean(r.get('Наименование')), 'description': ws_clean(r.get('Описание')), 'type': ntype, 'location': parse_locations(r.get('Расположение')), 'segment': parse_multiline_ids(r.get('Сетевой сегмент/зона(ID)') or r.get('Сетевой сегмент/зона'))}
            if ntype == 'LAN':
                if vlan := ws_clean(r.get('VLAN')): entry['vlan'] = int(float(vlan))
                entry['ipnetwork'] = ws_clean(r.get('Адрес сети'))
                entry['lan_type'] = ws_clean(r.get('Тип сети (проводная, беспроводная)') or r.get('Тип LAN'))
            elif ntype == 'WAN': entry['wan_ip'] = ws_clean(r.get('WAN Адрес'))
            if prov := ws_clean(r.get('Провайдер')): entry['provider'] = prov
            if vrf := ws_clean(r.get('VRF  ') or ws_clean(r.get('VRF'))): entry['VRF'] = vrf
            nets[nid] = entry
        if nets:
            res = True
            prefix = next(iter(nets.keys())).split('.')[0] if '.' in next(iter(nets.keys())) else ''
            per_loc, misc = {}, {}
            for nid, entry in nets.items():
                locs = entry.get('location')
                if not locs: misc[nid] = entry; continue
                for loc in locs:
                    token = re.sub(r'[^A-Za-z0-9]+', '_', str(loc)).strip('_') or 'loc'
                    if prefix:
                        if m := re.search(rf'{re.escape(prefix)}\.dc\.(\d+)', loc): token = f'dc{m.group(1)}'
                        elif m := re.search(rf'{re.escape(prefix)}\.office\.(.+)', loc): token = f'office_{m.group(1)}'
                    per_loc.setdefault(token, {})[nid] = entry
            for t, s in per_loc.items(): write_yaml(out_dir / f'networks_{t}.yaml', {'seaf.ta.services.network': s})
            if misc: write_yaml(out_dir / 'networks_misc.yaml', {'seaf.ta.services.network': misc})
    
    sheet = next((s for s in xls.sheet_names if s in ['Сетевые устройства', '??????? ??????????']), None)
    devs = {}
    if sheet:
        for idx, row in non_empty_rows(xls.parse(sheet)).iterrows():
            did = id_clean(row.get('ID Устройства') or row.get('ID ??????????'))
            if not did: continue
            VALIDATOR.register_id(did, "Сетевые устройства")
            
            locs = parse_locations(row.get('Расположение'))
            conn_nets = parse_multiline_ids(row.get('Подключенные сети (список)') or row.get('Подключенные сети'))
            VALIDATOR.check_ref_network(conn_nets, did)
            
            obj = {'title': ws_clean(row.get('Наименование')) or did, 'realization_type': ws_clean(row.get('Тип реализации')), 'type': ws_clean(row.get('Тип устройства') or row.get('Тип')), 'network_connection': conn_nets, 'segment': id_clean(row.get('Расположение (ID сегмента/зоны)') or row.get('Сетевой сегмент/зона (ID)'))}
            for k, cs in [('model', ['Модель']), ('purpose', ['Назначение']), ('address', ['IP адрес']), ('description', ['Описание'])]:
                for c in cs:
                    if val := ws_clean(row.get(c)): obj[k] = val; break
            if len(locs) > 1:
                for l in locs: devs[f"{did}-{l.split('.')[-1]}"] = {**obj, 'location': l}
            else: obj['location'] = locs[0] if locs else None; devs[did] = obj
    
    sheet_comp = next((s for s in xls.sheet_names if s == 'Компоненты'), None)
    if sheet_comp:
        comp_config = {
            'Network Device': ('components_network.yaml', 'seaf.ta.components.network'),
            'K8s Namespace': ('k8s_namespace.yaml', 'seaf.ta.components.k8s_namespace'),
            'K8s HPA': ('k8s_hpa.yaml', 'seaf.ta.components.k8s_hpa')
        }
        collected = {k: {} for k in comp_config}

        for idx, row in non_empty_rows(xls.parse(sheet_comp)).iterrows():
            cls = ws_clean(row.get('Класс'))
            if cls not in comp_config: continue
            
            did = id_clean(row.get('Идентификатор'))
            if not did: continue
            VALIDATOR.register_id(did, "Компоненты")
            
            locs = parse_locations(row.get('Локация'))
            conn_nets = parse_multiline_ids(row.get('Сети'))
            VALIDATOR.check_ref_network(conn_nets, did)
            
            obj = {
                'title': ws_clean(row.get('Наименование')),
                'description': ws_clean(row.get('Описание')),
                'type': ws_clean(row.get('Тип')),
                'network_connection': conn_nets,
                'segment': id_clean(row.get('Сегмент'))
            }
            if len(locs) > 1:
                for l in locs: collected[cls][f"{did}-{l.split('.')[-1]}"] = {**obj, 'location': l}
            else:
                obj['location'] = locs[0] if locs else None
                collected[cls][did] = obj
        
        if collected['Network Device']:
             devs.update(collected['Network Device'])

        for cls, data in collected.items():
            if cls == 'Network Device': continue 
            if data:
                fn, ns = comp_config[cls]
                write_yaml(out_dir / fn, {ns: data})
                res = True

    if devs: write_yaml(out_dir / 'components_network.yaml', {'seaf.ta.components.network': devs}); res = True
    return res

def convert_kb_services(xls, out_dir: Path):
    if 'Сервисы КБ' not in xls.sheet_names: return False
    kb = {}
    for _, r in non_empty_rows(xls.parse('Сервисы КБ')).iterrows():
        sid = id_clean(r.get('ID КБ сервиса'))
        if not sid: continue
        VALIDATOR.register_id(sid, "Сервисы КБ")
        
        conn_nets = parse_multiline_ids(r.get('Подключенные сети'))
        VALIDATOR.check_ref_network(conn_nets, sid)
        
        title = ws_clean(r.get('Название сервиса')) or ws_clean(r.get('Название')) or ws_clean(r.get('Технология')) or sid
        kb[sid] = {'title': title, 'description': ws_clean(r.get('Описание')), 'status': ws_clean(r.get('Статус')), 'technology': ws_clean(r.get('Технология')), 'software_name': ws_clean(r.get('Название ПО')), 'tag': ws_clean(r.get('Tag')), 'network_connection': conn_nets}
    if kb: write_yaml(out_dir / 'kb.yaml', {'seaf.ta.services.kb': kb}); return True
    return False

def convert_tech_services(xls, out_dir: Path):
    sheet = next((s for s in xls.sheet_names if s in ['Тех. сервисы', 'Tech Services']), None)
    if not sheet: return False
    out_data = {'compute_service': {}, 'cluster': {}, 'monitoring': {}, 'backup': {}, 'software': {}, 'storage': {}}
    
    # Track unique IDs to avoid duplication if the same ID appears multiple times in Excel (e.g. multi-location)
    for _, row in non_empty_rows(xls.parse(sheet)).iterrows():
        oid = id_clean(row.get('Идентификатор'))
        if not oid: continue
        
        # Note: We don't register ID here immediately because rows might be split by location/network.
        # Ideally, we should check uniqueness after aggregating or flag it differently.
        # For now, let's register it to catch if other sheets use the same ID (unlikely but possible).
        # We accept multiple rows for same ID in Tech Services as "partial definitions" to be merged.
        # So we skip duplicate ID check within this loop or handle it gracefully.
        
        svc_raw, res_val, cls_val = ws_clean(row.get('Тип сервиса')) or ws_clean(row.get('Класс')), ws_clean(row.get('Тип резервирования')), ws_clean(row.get('Класс'))
        nets = parse_multiline_ids(row.get('Подключен к сети') or row.get('Подключен к  сети'))
        locs = parse_locations(row.get('ЦОД'))
        if not locs:
            for n in nets:
                if l := derive_location_from_network(n): locs.append(l)
        locs = sorted(list(set(locs)))
        
        VALIDATOR.check_ref_network(nets, oid)
        
        etype = 'compute_service'
        if svc_raw in SPECIAL_ENTITY_MAP: etype = SPECIAL_ENTITY_MAP[svc_raw]
        elif cls_val == 'Cluster' or (res_val and res_val.lower() in ['active-active','active-passive','n+1','да']): etype = 'cluster'
        elif cls_val == 'Software': etype = 'software'
        elif cls_val == 'Storage': etype = 'storage'
        elif cls_val == 'Monitoring': etype = 'monitoring'
        elif cls_val == 'Backup': etype = 'backup'
        elif cls_val == 'Compute Service': etype = 'compute_service'
        
        if oid in out_data[etype]:
            # Merge locations and networks for duplicate IDs (multi-page/multi-location export)
            existing = out_data[etype][oid]
            existing['location'] = sorted(list(set(existing['location'] + locs)))
            existing['network_connection'] = sorted(list(set(existing['network_connection'] + nets)))
            continue
        
        # Only register new IDs
        VALIDATOR.register_id(oid, "Тех. сервисы")

        obj = {'title': ws_clean(row.get('Наименование')), 'description': ws_clean(row.get('Описание')), 'location': locs, 'network_connection': nets, 'availabilityzone': []}
        if etype in ['compute_service', 'cluster']: obj['service_type'] = normalize_svc_type(svc_raw)
        if etype == 'cluster': obj['reservation_type'] = res_val
        elif etype == 'monitoring': obj.update({'role':['Monitoring'], 'ha': res_val is not None, 'monitored_services':[]})
        elif etype == 'backup': obj.update({'path':'/', 'backed_up_services':[]})
        elif etype == 'software': pass
        elif etype == 'storage': pass
        out_data[etype][oid] = obj
    
    emap = {
        'compute_service': ('compute_service.yaml', 'seaf.ta.services.compute_service'),
        'cluster': ('cluster.yaml', 'seaf.ta.services.cluster'),
        'monitoring': ('monitoring.yaml', 'seaf.ta.services.monitoring'),
        'backup': ('backup.yaml', 'seaf.ta.services.backup'),
        'software': ('software.yaml', 'seaf.ta.services.software'),
        'storage': ('storage.yaml', 'seaf.ta.services.storage')
    }
    found = False
    for k, (fn, root) in emap.items():
        if out_data[k]: 
            write_yaml(out_dir / fn, {root: out_data[k]})
            found = True
    return found

def main():
    try:
        ensure_deps()
        parser = argparse.ArgumentParser()
        parser.add_argument('--config', required=True)
        parser.add_argument('--force', '-y', action='store_true', help='Skip validation prompts and force processing')
        args = parser.parse_args()
        
        cpath = Path(args.config)
        if not cpath.exists(): sys.exit(1)
        with cpath.open('r', encoding='utf-8') as f: cfg = yaml.safe_load(f) or {}
        inputs, out_dir = [cpath.parent / p for p in (cfg.get('xlsx_files') or [])], cpath.parent / cfg.get('out_yaml_dir', 'out_yaml')
        
        if out_dir.exists():
            for i in out_dir.glob('*'):
                if i.is_file(): i.unlink()
        out_dir.mkdir(parents=True, exist_ok=True)
        
        # Pre-flight check
        valid_files = []
        for p in inputs:
            if not p.exists():
                print(f"ERROR: {p.name} not found.", file=sys.stderr)
                continue
            if validate_structure(p, args.force):
                valid_files.append(p)
        
        if not valid_files:
            print("ERROR: No valid files to process.", file=sys.stderr)
            sys.exit(1)

        src_counts = count_entities_in_xlsx(valid_files)
        processed = False
        
        # Process files
        for p in valid_files:
            try:
                xls = pd.ExcelFile(p)
                # Order matters for reference checking: regions/nets first usually
                if any(s in xls.sheet_names for s in ['Регионы','AZ','DC','Офисы']): 
                    if convert_regions_az_dc_offices(xls, out_dir): processed = True
                
                if any(s in xls.sheet_names for s in ['Сегменты','Сети','Сетевые устройства']): 
                    if convert_segments_nets_devices(xls, out_dir): processed = True
                
                if 'Сервисы КБ' in xls.sheet_names: 
                    if convert_kb_services(xls, out_dir): processed = True
                
                if any(s in xls.sheet_names for s in ['Тех. сервисы','Tech Services']): 
                    if convert_tech_services(xls, out_dir): processed = True
            except Exception as e: print(f"ERROR: {p.name}: {e}", file=sys.stderr)
            
        if not processed: print("ERROR: No valid data processed.", file=sys.stderr); sys.exit(1)
        
        # Post-processing reporting
        VALIDATOR.report()
        
        imports = [p.name for p in sorted(out_dir.glob('*.yaml')) if p.name != 'root.yaml']
        if imports: write_yaml(out_dir / 'root.yaml', {'imports': imports})
        dst_counts = count_entities_in_yaml_dir(out_dir)
        print("\n--- Conversion Summary ---")
        for k in sorted(list(set(src_counts.keys()) | set(dst_counts.keys()))):
            s, d = src_counts.get(k, 0), dst_counts.get(k, 0)
            print(f"  - {k:<25} | Source: {s:<5} | Dest: {d:<5} | {'OK' if s==d else 'FAIL'}")
    except Exception as e: print(f"FATAL: {e}", file=sys.stderr); sys.exit(1)

if __name__ == '__main__': main()