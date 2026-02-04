import sys
import re
import argparse
from pathlib import Path
from typing import List, Dict, Any, Tuple
from copy import deepcopy
import math
import yaml
import pandas as pd

DEBUG_LOG_FILE = Path('debug_script.log')

def log_debug(message):
    try:
        with DEBUG_LOG_FILE.open('a', encoding='utf-8') as f:
            f.write(f"{message}\n")
    except Exception: pass

SPECIAL_ENTITY_MAP = {
    'Мониторинг': 'monitorings',
    'Логгирование': 'monitorings',
    'Резервное копирование': 'backups',
    'Бекапирование и восстановление данных': 'backups'
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
            if idx+1 < len(parts) and parts[idx+1] not in ['lan','wan']: return f"{prefix}.office.{parts[idx+1]}"
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
    sheet_map = {'Регионы': 'dc_regions', 'AZ': 'dc_azs', 'DC': 'dcs', 'Офисы': 'dc_offices', 'Сегменты': 'network_segments', 'Сети': 'networks', 'Сетевые устройства': 'components.networks', 'Сервисы КБ': 'kbs', 'Тех. сервисы': 'tech_services', 'Tech Services': 'tech_services'}
    for file_path in xlsx_files:
        if not file_path.exists(): continue
        try:
            xls = pd.ExcelFile(file_path)
            for sheet_name in xls.sheet_names:
                if sheet_name in sheet_map:
                    df = non_empty_rows(xls.parse(sheet_name))
                    ename = sheet_map[sheet_name]
                    if ename == 'tech_services':
                        for _, row in df.iterrows():
                            svc_raw = ws_clean(row.get('Тип сервиса')) or ws_clean(row.get('Класс'))
                            res_val = ws_clean(row.get('Тип резервирования'))
                            etype = 'compute_services'
                            if svc_raw in SPECIAL_ENTITY_MAP: etype = SPECIAL_ENTITY_MAP[svc_raw]
                            elif svc_raw == 'Cluster' or (res_val and res_val.lower() in ['active-active', 'active-passive', 'n+1', 'да']): etype = 'clusters'
                            counts[etype] = counts.get(etype, 0) + 1
                    else: counts[ename] = counts.get(ename, 0) + len(df)
        except Exception as e: print(f"WARN: Count failed for {file_path.name}: {e}", file=sys.stderr)
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
                        ename = key.replace('seaf.company.ta.services.', '').replace('seaf.company.ta.components.', 'components.')
                        counts[ename] = counts.get(ename, 0) + len(val)
        except Exception: continue
    return counts

def convert_regions_az_dc_offices(xlsx_path: Path, out_dir: Path):
    try: xls = read_excel(xlsx_path)
    except Exception as e: print(f"ERROR: {e}", file=sys.stderr); return
    reg, azs, dcs, off, proc = {}, {}, {}, {}, set()
    if 'Регионы' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('Регионы')).iterrows():
            rid = id_clean(row.get('ID Региона'))
            if rid and rid not in proc:
                proc.add(rid)
                reg[rid] = {'description': ws_clean(row.get('Описание')), 'external_id': rid.split('.')[-1], 'title': ws_clean(row.get('Наименование'))}
    if reg: write_yaml(out_dir / 'dc_region.yaml', {'seaf.company.ta.services.dc_regions': reg})
    proc = set()
    if 'AZ' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('AZ')).iterrows():
            aid = id_clean(row.get('ID AZ'))
            if aid and aid not in proc:
                proc.add(aid)
                azs[aid] = {'description': ws_clean(row.get('Описание')), 'external_id': aid.split('.')[-1], 'region': id_clean(row.get('Регион')), 'title': ws_clean(row.get('Наименование')), 'vendor': ws_clean(row.get('Поставщик'))}
    if azs: write_yaml(out_dir / 'dc_az.yaml', {'seaf.company.ta.services.dc_azs': azs})
    proc = set()
    if 'DC' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('DC')).iterrows():
            did = id_clean(row.get('ID DC'))
            if did and did not in proc:
                proc.add(did)
                dcs[did] = {'address': ws_clean(row.get('Адрес')), 'availabilityzone': id_clean(row.get('AZ')), 'description': ws_clean(row.get('Описание')), 'external_id': did.split('.')[-1], 'ownership': ws_clean(row.get('Форма владения')), 'rack_qty': ws_clean(row.get('Кол-во стоек')), 'tier': ws_clean(row.get('Tier')), 'title': ws_clean(row.get('Наименование')), 'type': ws_clean(row.get('Тип')), 'vendor': ws_clean(row.get('Поставщик'))}
    if dcs: write_yaml(out_dir / 'dc.yaml', {'seaf.company.ta.services.dcs': dcs})
    proc = set()
    if 'Офисы' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('Офисы')).iterrows():
            oid = id_clean(row.get('ID Офиса'))
            if oid and oid not in proc:
                proc.add(oid)
                off[oid] = {'address': ws_clean(row.get('Адрес')), 'description': ws_clean(row.get('Описание')), 'external_id': oid.split('.')[-1], 'region': id_clean(row.get('Регион')), 'title': ws_clean(row.get('Наименование'))}
    if off: write_yaml(out_dir / 'dc_office.yaml', {'seaf.company.ta.services.dc_offices': off})

def convert_segments_nets_devices(xlsx_path: Path, out_dir: Path) -> int:
    try: xls = read_excel(xlsx_path)
    except Exception as e: print(f"ERROR: {e}", file=sys.stderr); return 0
    segments, proc_seg = {}, set()
    if 'Сегменты' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('Сегменты')).iterrows():
            sid = id_clean(row.get('ID сетевые сегмента/зоны'))
            if sid and sid not in proc_seg:
                proc_seg.add(sid)
                locs = parse_locations(row.get('Расположение'))
                segments[sid] = {'title': ws_clean(row.get('Наименование')), 'description': ws_clean(row.get('Описание')), 'sber': {'location': locs[0] if locs else None, 'zone': ws_clean(row.get('Зона'))}}
    nets, proc_net = {}, set()
    if 'Сети' in xls.sheet_names:
        for idx, row in non_empty_rows(xls.parse('Сети')).iterrows():
            nid = id_clean(row.get('ID Network'))
            if nid and nid not in proc_net:
                proc_net.add(nid)
                ntype = ws_clean(row.get('Тип сети'))
                entry = {'title': ws_clean(row.get('Наименование')), 'description': ws_clean(row.get('Описание')), 'type': ntype, 'location': parse_locations(row.get('Расположение')), 'segment': parse_multiline_ids(row.get('Сетевой сегмент/зона(ID)') or row.get('Сетевой сегмент/зона'))}
                if ntype == 'LAN':
                    if vlan := ws_clean(row.get('VLAN')):
                        try: entry['vlan'] = int(float(vlan))
                        except ValueError: pass
                    entry['ipnetwork'] = ws_clean(row.get('Адрес сети'))
                    entry['lan_type'] = ws_clean(row.get('Тип сети (проводная, беспроводная)') or row.get('Тип LAN'))
                elif ntype == 'WAN': entry['wan_ip'] = ws_clean(row.get('WAN Адрес'))
                if prov := ws_clean(row.get('Провайдер')): entry['provider'] = prov
                if vrf := ws_clean(row.get('VRF  ') or ws_clean(row.get('VRF'))): entry['VRF'] = vrf
                nets[nid] = entry
    if segments: write_yaml(out_dir / 'network_segment.yaml', {'seaf.company.ta.services.network_segments': segments})
    if nets: 
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
        for t, s in per_loc.items(): write_yaml(out_dir / f'networks_{t}.yaml', {'seaf.company.ta.services.networks': s})
        if misc: write_yaml(out_dir / 'networks_misc.yaml', {'seaf.company.ta.services.networks': misc})
    devs, proc_dev = {}, set()
    sheet = next((s for s in xls.sheet_names if s in ['Сетевые устройства', '??????? ??????????']), None)
    if sheet:
        for idx, row in non_empty_rows(xls.parse(sheet)).iterrows():
            did = id_clean(row.get('ID Устройства') or row.get('ID ??????????'))
            if did and did not in proc_dev:
                proc_dev.add(did)
                locs = parse_locations(row.get('Расположение'))
                obj = {'title': ws_clean(row.get('Наименование')) or did, 'realization_type': ws_clean(row.get('Тип реализации')), 'type': ws_clean(row.get('Тип устройства') or row.get('Тип')), 'network_connection': parse_multiline_ids(row.get('Подключенные сети (список)') or row.get('Подключенные сети')), 'segment': id_clean(row.get('Расположение (ID сегмента/зоны)') or row.get('Сетевой сегмент/зона (ID)'))}
                for k, cols in [('model', ['Модель']), ('purpose', ['Назначение']), ('address', ['IP адрес']), ('description', ['Описание'])]:
                    for c in cols:
                        if val := ws_clean(row.get(c)): obj[k] = val; break
                if len(locs) > 1:
                    for l in locs: devs[f"{did}-{l.split('.')[-1]}"] = {**obj, 'location': l}
                else: obj['location'] = locs[0] if locs else None; devs[did] = obj
    if devs: write_yaml(out_dir / 'network_component.yaml', {'seaf.company.ta.components.networks': devs})
    return 0

def convert_kb_services(xlsx_path: Path, out_dir: Path):
    try:
        xls = read_excel(xlsx_path)
        if 'Сервисы КБ' not in xls.sheet_names: return
        kb, proc = {}, set()
        for idx, row in non_empty_rows(xls.parse('Сервисы КБ')).iterrows():
            sid = id_clean(row.get('ID КБ сервиса'))
            if sid and sid not in proc:
                proc.add(sid)
                kb[sid] = {'title': ws_clean(row.get('Название сервиса')) or ws_clean(row.get('Название')), 'description': ws_clean(row.get('Описание')), 'status': ws_clean(row.get('Статус')), 'technology': ws_clean(row.get('Технология')), 'software_name': ws_clean(row.get('Название ПО')), 'tag': ws_clean(row.get('Tag')), 'network_connection': parse_multiline_ids(row.get('Подключенные сети'))}
        if kb: write_yaml(out_dir / 'kb.yaml', {'seaf.company.ta.services.kbs': kb})
    except Exception as e: print(f"WARN: KB failed for {xlsx_path.name}: {e}", file=sys.stderr)

def convert_tech_services(xlsx_path: Path, out_dir: Path):
    try:
        xls = read_excel(xlsx_path)
        sheet = next((s for s in xls.sheet_names if s in ['Тех. сервисы', 'Tech Services']), None)
        if not sheet: return
        out_data = {'compute_services': {}, 'clusters': {}, 'monitorings': {}, 'backups': {}}
        proc = set()
        df = non_empty_rows(xls.parse(sheet))
        for _, row in df.iterrows():
            oid = id_clean(row.get('Идентификатор'))
            if not oid or oid in proc: continue
            proc.add(oid)
            svc_raw = ws_clean(row.get('Тип сервиса')) or ws_clean(row.get('Класс'))
            res_val = ws_clean(row.get('Тип резервирования'))
            cls_val = ws_clean(row.get('Класс'))
            nets = parse_multiline_ids(row.get('Подключен к сети') or row.get('Подключен к  сети'))
            locs = parse_locations(row.get('ЦОД'))
            if not locs:
                for n in nets:
                    if l := derive_location_from_network(n): locs.append(l)
            locs = sorted(list(set(locs)))
            etype = 'compute_services'
            if svc_raw in SPECIAL_ENTITY_MAP: etype = SPECIAL_ENTITY_MAP[svc_raw]
            elif cls_val == 'Cluster' or (res_val and res_val.lower() in ['active-active','active-passive','n+1','да']): etype = 'clusters'
            elif cls_val == 'Compute Service': etype = 'compute_services'
            obj = {'title': ws_clean(row.get('Наименование')), 'description': ws_clean(row.get('Описание')), 'location': locs, 'network_connection': nets, 'availabilityzone': []}
            if etype in ['compute_services', 'clusters']: obj['service_type'] = normalize_svc_type(svc_raw)
            if etype == 'clusters': obj['reservation_type'] = res_val
            elif etype == 'monitorings': obj.update({'role':['Monitoring'], 'ha': res_val is not None, 'monitored_services':[]})
            elif etype == 'backups': obj.update({'path':'/', 'backed_up_services':[]})
            out_data[etype][oid] = obj
        emap = {'compute_services': ('compute_service.yaml', 'seaf.company.ta.services.compute_services'), 'clusters': ('cluster.yaml', 'seaf.company.ta.services.clusters'), 'monitorings': ('monitoring.yaml', 'seaf.company.ta.services.monitorings'), 'backups': ('backup.yaml', 'seaf.company.ta.services.backups')}
        for k, (fn, root) in emap.items():
            if out_data[k]: write_yaml(out_dir / fn, {root: out_data[k]})
    except Exception as e: print(f"ERROR: Tech failed for {xlsx_path.name}: {e}", file=sys.stderr)

def write_root(out_dir: Path):
    imports = [p.name for p in sorted(out_dir.glob('*.yaml')) if not p.name.startswith('_')]
    if imports: write_yaml(out_dir / '_root.yaml', {'imports': imports})

def main():
    try:
        ensure_deps()
        parser = argparse.ArgumentParser()
        parser.add_argument('--config', required=True)
        args = parser.parse_args()
        cpath = Path(args.config)
        if not cpath.exists(): sys.exit(1)
        with cpath.open('r', encoding='utf-8') as f: cfg = yaml.safe_load(f) or {}
        inputs = [cpath.parent / p for p in (cfg.get('xlsx_files') or [])]
        out_dir = cpath.parent / cfg.get('out_yaml_dir', 'out_yaml')
        if out_dir.exists():
            for i in out_dir.glob('*'):
                if i.is_file(): i.unlink()
        out_dir.mkdir(parents=True, exist_ok=True)
        src_counts = count_entities_in_xlsx(inputs)
        processed = False
        for p in inputs:
            if not p.exists(): print(f"ERROR: {p.name} not found.", file=sys.stderr); continue
            try:
                xls = pd.ExcelFile(p)
                if any(s in xls.sheet_names for s in ['Регионы','AZ','DC','Офисы']): convert_regions_az_dc_offices(p, out_dir); processed = True
                if any(s in xls.sheet_names for s in ['Сегменты','Сети','Сетевые устройства']): convert_segments_nets_devices(p, out_dir); processed = True
                if 'Сервисы КБ' in xls.sheet_names: convert_kb_services(p, out_dir); processed = True
                if any(s in xls.sheet_names for s in ['Тех. сервисы','Tech Services']): convert_tech_services(p, out_dir); processed = True
            except Exception as e: print(f"ERROR: {p.name}: {e}", file=sys.stderr)
        if not processed: sys.exit(1)
        write_root(out_dir)
        dst_counts = count_entities_in_yaml_dir(out_dir)
        print("\n--- Conversion Summary ---")
        for k in sorted(list(set(src_counts.keys()) | set(dst_counts.keys()))):
            s, d = src_counts.get(k, 0), dst_counts.get(k, 0)
            print(f"  - {k:<25} | Source: {s:<5} | Dest: {d:<5} | {'OK' if s==d else 'FAIL'}")
    except Exception as e: print(f"FATAL: {e}", file=sys.stderr); sys.exit(1)

if __name__ == '__main__': main()
