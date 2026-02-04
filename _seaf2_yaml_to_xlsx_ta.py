import sys
import argparse
from pathlib import Path
from typing import Dict, Any, List
import pandas as pd
import yaml
import re

def ensure_deps():
    try: import pandas, yaml, openpyxl
    except ImportError:
        import subprocess
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pandas', 'openpyxl', 'pyyaml'])

def read_yaml(path: Path) -> Dict[str, Any]:
    try:
        if not path.exists(): return {}
        with path.open('r', encoding='utf-8') as f: return yaml.safe_load(f) or {}
    except Exception: return {}

def normalize_val(val, default=None):
    if not val or not isinstance(val, str): return default if default else val
    trans = str.maketrans('CAEOPXy', 'САЕОРХу')
    return val.translate(trans).strip()

def format_list(items: Any) -> str:
    if not items: return ''
    if isinstance(items, str): return items
    if isinstance(items, list): return ', '.join(sorted([str(x) for x in items if x]))
    return str(items)

def derive_location_from_network(net_id):
    if not net_id: return None
    prefix = net_id.split('.')[0] if '.' in net_id else 'seaf'
    m = re.search(r'\.dc(\d+)', net_id)
    if m: return f"{prefix}.dc.{m.group(1)}"
    m = re.search(r'\.office\.([a-zA-Z0-9-]+)', net_id)
    if m:
        parts = net_id.split('.')
        try:
            idx = parts.index('office')
            if idx+1 < len(parts) and parts[idx+1] not in ['lan','wan']: return f"{prefix}.office.{parts[idx+1]}"
            return f"{prefix}.office"
        except ValueError: pass
    return None

ENTITY_MAP = {
    'dc_region': ['seaf.company.ta.services.dc_regions', 'seaf.ta.services.dc_regions'],
    'dc_az': ['seaf.company.ta.services.dc_azs', 'seaf.ta.services.dc_azs'],
    'dc': ['seaf.company.ta.services.dcs', 'seaf.ta.services.dcs'],
    'office': ['seaf.company.ta.services.dc_offices', 'seaf.ta.services.dc_offices'],
    'network_segment': ['seaf.company.ta.services.network_segments', 'seaf.ta.services.network_segment'],
    'network': ['seaf.company.ta.services.networks', 'seaf.ta.services.network'],
    'kb': ['seaf.company.ta.services.kbs', 'seaf.ta.services.kb'],
    'components.network': ['seaf.company.ta.components.networks', 'seaf.ta.components.networks'],
    'compute_service': ['seaf.company.ta.services.compute_services', 'seaf.ta.services.compute_service'],
    'cluster': ['seaf.company.ta.services.clusters', 'seaf.ta.services.cluster'],
    'monitoring': ['seaf.company.ta.services.monitorings', 'seaf.ta.services.monitoring'],
    'backup': ['seaf.company.ta.services.backups', 'seaf.ta.services.backup'],
    'software': ['seaf.company.ta.services.softwares', 'seaf.ta.services.software'],
    'storage': ['seaf.company.ta.services.storages', 'seaf.ta.services.storage'],
    'cluster_virtualization': ['seaf.company.ta.services.cluster_virtualizations', 'seaf.ta.services.cluster_virtualization'],
    'k8s': ['seaf.company.ta.services.k8s', 'seaf.ta.services.k8s'],
    'k8s_deployment': ['seaf.company.ta.services.k8s_deployments', 'seaf.ta.services.k8s_deployment'],
    'logical_link': ['seaf.company.ta.services.logical_links', 'seaf.ta.services.logical_link'],
    'network_link': ['seaf.company.ta.services.network_links', 'seaf.ta.services.network_link'],
    'stand': ['seaf.company.ta.services.stands', 'seaf.ta.services.stand'],
    'environment': ['seaf.company.ta.services.environments', 'seaf.ta.services.environment'],
    'server': ['seaf.company.ta.components.servers', 'seaf.ta.components.server'],
    'hw_storage': ['seaf.company.ta.components.hw_storages', 'seaf.ta.components.hw_storage'],
    'user_device': ['seaf.company.ta.components.user_devices', 'seaf.ta.components.user_device'],
    'k8s_node': ['seaf.company.ta.components.k8s_nodes', 'seaf.ta.components.k8s_node'],
    'k8s_namespace': ['seaf.company.ta.components.k8s_namespaces', 'seaf.ta.components.k8s_namespace'],
    'k8s_hpa': ['seaf.company.ta.components.k8s_hpas', 'seaf.ta.components.k8s_hpa']
}

# Mapping internal names to display Class names
CLASS_NAME_MAP = {
    'compute_service': 'Compute Service',
    'cluster': 'Cluster',
    'monitoring': 'Monitoring',
    'backup': 'Backup',
    'software': 'Software',
    'storage': 'Storage',
    'cluster_virtualization': 'Cluster Virtualization',
    'k8s': 'K8s Cluster',
    'k8s_deployment': 'Deployment',
    'server': 'Server',
    'hw_storage': 'HW Storage',
    'user_device': 'User Device',
    'k8s_node': 'K8s Node',
    'k8s_namespace': 'K8s Namespace',
    'k8s_hpa': 'K8s HPA',
    'logical_link': 'Logical Link',
    'network_link': 'Network Link',
    'stand': 'Stand',
    'environment': 'Environment'
}

def count_entities_in_yaml_dir(yaml_dir: Path) -> Dict[str, int]:
    counts = {}
    if not yaml_dir.exists(): return counts
    rev_lookup = {}
    for etype, keys in ENTITY_MAP.items():
        for k in keys: rev_lookup[k] = etype
    for p in sorted(yaml_dir.glob('**/*.yaml')):
        data = read_yaml(p)
        for key, val in data.items():
            if key in rev_lookup and isinstance(val, dict):
                etype = rev_lookup[key]
                counts[etype] = counts.get(etype, 0) + len(val)
            elif key.startswith('seaf.ta.reverse.'):
                counts['reverse'] = counts.get('reverse', 0) + len(val)
    return counts

def count_entities_in_xlsx(xlsx_files: List[Path]) -> Dict[str, int]:
    counts = {}
    # Use internal keys from ENTITY_MAP for sheet mapping consistency
    sheet_map = {
        'Регионы': 'dc_region', 'AZ': 'dc_az', 'DC': 'dc', 'Офисы': 'office', 
        'Сегменты': 'network_segment', 'Сети': 'network', 'Сервисы КБ': 'kb', 
        'Тех. сервисы': 'tech_services', 'Компоненты': 'components', 
        'Сетевые устройства': 'components.network', 'Связи': 'links', 
        'Стенды и окружения': 'stands', 'Reverse': 'reverse'
    }
    # Reverse map for display names back to internal keys
    rev_class_map = {v: k for k, v in CLASS_NAME_MAP.items()}

    for fp in xlsx_files:
        if not fp.exists(): continue
        try:
            xls = pd.ExcelFile(fp)
            for sn in xls.sheet_names:
                if sn in sheet_map:
                    df = xls.parse(sn).dropna(how='all')
                    ename = sheet_map[sn]
                    if ename == 'tech_services':
                        for _, row in df.iterrows():
                            cls_raw = row.get('Класс')
                            etype = rev_class_map.get(cls_raw, 'compute_service')
                            counts[etype] = counts.get(etype, 0) + 1
                    elif ename == 'components':
                        for _, row in df.iterrows():
                            cls_raw = row.get('Класс')
                            etype = rev_class_map.get(cls_raw, 'server')
                            counts[etype] = counts.get(etype, 0) + 1
                    elif ename == 'links':
                        for _, row in df.iterrows():
                            cls_raw = row.get('Класс')
                            etype = rev_class_map.get(cls_raw, 'logical_link')
                            counts[etype] = counts.get(etype, 0) + 1
                    elif ename == 'stands':
                        for _, row in df.iterrows():
                            cls_raw = row.get('Класс')
                            etype = rev_class_map.get(cls_raw, 'stand')
                            counts[etype] = counts.get(etype, 0) + 1
                    else: counts[ename] = counts.get(ename, 0) + len(df)
        except Exception: continue
    return counts

def save_regions_az_dc_offices(ydir: Path, writer: pd.ExcelWriter):
    rows_reg, rows_az, rows_dc, rows_off = [], [], [], []
    for p in sorted(ydir.glob('**/*.yaml')):
        data = read_yaml(p)
        for k, d in data.items():
            if k in ENTITY_MAP['dc_region']:
                for i, v in d.items(): rows_reg.append({'ID Региона': i, 'Наименование': v.get('title'), 'Описание': v.get('description')})
            elif k in ENTITY_MAP['dc_az']:
                for i, v in d.items(): rows_az.append({'ID AZ': i, 'Наименование': v.get('title'), 'Описание': v.get('description'), 'Поставщик': v.get('vendor'), 'Регион': v.get('region')})
            elif k in ENTITY_MAP['dc']:
                for i, v in d.items(): rows_dc.append({'ID DC': i, 'Наименование': v.get('title'), 'Описание': v.get('description'), 'Поставщик': v.get('vendor'), 'Tier': v.get('tier'), 'Тип': v.get('type'), 'Кол-во стоек': v.get('rack_qty'), 'Адрес': v.get('address'), 'Форма владения': v.get('ownership'), 'AZ': v.get('availabilityzone')})
            elif k in ENTITY_MAP['office']:
                for i, v in d.items(): rows_off.append({'ID Офиса': i, 'Наименование': v.get('title'), 'Описание': v.get('description'), 'Адрес': v.get('address'), 'Регион': v.get('region')})
    if rows_reg: pd.DataFrame(rows_reg).to_excel(writer, sheet_name='Регионы', index=False)
    if rows_az: pd.DataFrame(rows_az).to_excel(writer, sheet_name='AZ', index=False)
    if rows_dc: pd.DataFrame(rows_dc).to_excel(writer, sheet_name='DC', index=False)
    if rows_off: pd.DataFrame(rows_off).to_excel(writer, sheet_name='Офисы', index=False)

def save_segments_nets_devices(ydir: Path, writer: pd.ExcelWriter):
    rows_seg, rows_net, rows_dev = [], [], []
    for p in sorted(ydir.glob('**/*.yaml')):
        data = read_yaml(p)
        for k, d in data.items():
            if k in ENTITY_MAP['network_segment']:
                for i, v in d.items(): rows_seg.append({'ID сетевые сегмента/зоны': i, 'Наименование': v.get('title'), 'Описание': v.get('description'), 'Расположение': (v.get('sber') or {}).get('location'), 'Зона': (v.get('sber') or {}).get('zone')})
            elif k in ENTITY_MAP['network']:
                for i, v in d.items(): rows_net.append({'ID Network': i, 'Наименование': v.get('title'), 'Описание': v.get('description'), 'Тип сети': v.get('type'), 'VLAN': v.get('vlan'), 'VRF  ': v.get('VRF'), 'Провайдер': v.get('provider') or (v.get('sber') or {}).get('provider'), 'Тип сети (проводная, беспроводная)': v.get('lan_type'), 'Адрес сети': v.get('ipnetwork'), 'WAN Адрес': v.get('wan_ip'), 'Расположение': format_list(v.get('location')), 'Сетевой сегмент/зона(ID)': format_list(v.get('segment'))})
            elif k in ENTITY_MAP['components.network']:
                for i, v in d.items(): rows_dev.append({'ID Устройства': i, 'Наименование': v.get('title'), 'Тип реализации': v.get('realization_type'), 'Тип': v.get('type'), 'Модель': v.get('model'), 'Назначение': v.get('purpose'), 'IP адрес': v.get('address'), 'Описание': v.get('description'), 'Расположение (ID сегмента/зоны)': v.get('segment'), 'Подключенные сети (список)': format_list(v.get('network_connection'))})
    if rows_seg: pd.DataFrame(rows_seg).to_excel(writer, sheet_name='Сегменты', index=False)
    if rows_net: pd.DataFrame(rows_net).to_excel(writer, sheet_name='Сети', index=False)
    if rows_dev: pd.DataFrame(rows_dev).to_excel(writer, sheet_name='Сетевые устройства', index=False)

def save_kb_services(ydir: Path, writer: pd.ExcelWriter):
    rows = []
    for p in sorted(ydir.glob('**/*.yaml')):
        data = read_yaml(p)
        for k, d in data.items():
            if k in ENTITY_MAP['kb']:
                for i, v in d.items(): rows.append({'ID КБ сервиса': i, 'Tag': v.get('tag'), 'Описание': v.get('description'), 'Технология': v.get('technology'), 'Название ПО': v.get('software_name'), 'Статус': v.get('status'), 'Подключенные сети': format_list(v.get('network_connection'))})
    if rows: pd.DataFrame(rows).to_excel(writer, sheet_name='Сервисы КБ', index=False)

def save_tech_services(ydir: Path, writer: pd.ExcelWriter):
    rows = []
    kmap = {}
    for etype, keys in ENTITY_MAP.items():
        if etype in ['compute_service', 'cluster', 'monitoring', 'backup', 'software', 'storage', 'cluster_virtualization', 'k8s', 'k8s_deployment']:
            for k in keys: kmap[k] = CLASS_NAME_MAP[etype]
    for p in sorted(ydir.glob('**/*.yaml')):
        data = read_yaml(p)
        for rkey, ent in data.items():
            if rkey in kmap:
                cls = kmap[rkey]
                for i, d in ent.items():
                    ls = list(d.get('location') or [])
                    if not ls:
                        for n in (d.get('network_connection') or []):
                            if l := derive_location_from_network(n): ls.append(l)
                    obj = {'Идентификатор': i, 'Наименование': d.get('title'), 'Описание': d.get('description'), 'Подключен к сети': format_list(d.get('network_connection')), 'ЦОД': format_list(list(set(ls))), 'Класс': cls}
                    if 'Compute' in cls or 'Cluster' in cls: obj['Тип сервиса'] = normalize_val(d.get('service_type'), 'Серверы приложений и т.д.')
                    if 'Cluster' in cls and 'reservation_type' in d: obj['Тип резервирования'] = d['reservation_type']
                    rows.append(obj)
    if rows: pd.DataFrame(rows).to_excel(writer, sheet_name='Тех. сервисы', index=False)

def save_components(ydir: Path, writer: pd.ExcelWriter):
    rows = []
    cmap = {}
    for etype, keys in ENTITY_MAP.items():
        if etype in ['server', 'hw_storage', 'user_device', 'k8s_node', 'k8s_namespace', 'k8s_hpa']:
            for k in keys: cmap[k] = CLASS_NAME_MAP[etype]
    for p in sorted(ydir.glob('**/*.yaml')):
        data = read_yaml(p)
        for rkey, ent in data.items():
            if rkey in cmap:
                cls = cmap[rkey]
                for i, d in ent.items():
                    rows.append({'Идентификатор': i, 'Наименование': d.get('title'), 'Описание': d.get('description'), 'Класс': cls, 'Тип': d.get('type') or d.get('device_type'), 'Локация': format_list(d.get('location')), 'Сети': format_list(d.get('network_connection') or d.get('subnets')), 'Сегмент': d.get('segment')})
    if rows: pd.DataFrame(rows).to_excel(writer, sheet_name='Компоненты', index=False)

def save_links(ydir: Path, writer: pd.ExcelWriter):
    rows = []
    cmap = {}
    for etype, keys in ENTITY_MAP.items():
        if etype in ['logical_link', 'network_link']:
            for k in keys: cmap[k] = CLASS_NAME_MAP[etype]
    for p in sorted(ydir.glob('**/*.yaml')):
        data = read_yaml(p)
        for rkey, ent in data.items():
            if rkey in cmap:
                cls = cmap[rkey]
                for i, d in ent.items():
                    rows.append({'Идентификатор': i, 'Описание': d.get('description'), 'Класс': cls, 'Источник': d.get('source'), 'Приемник': format_list(d.get('target')), 'Направление': d.get('direction'), 'Сети': format_list(d.get('network_connection'))})
    if rows: pd.DataFrame(rows).to_excel(writer, sheet_name='Связи', index=False)

def save_stands(ydir: Path, writer: pd.ExcelWriter):
    rows = []
    cmap = {}
    for etype, keys in ENTITY_MAP.items():
        if etype in ['stand', 'environment']:
            for k in keys: cmap[k] = CLASS_NAME_MAP[etype]
    for p in sorted(ydir.glob('**/*.yaml')):
        data = read_yaml(p)
        for rkey, ent in data.items():
            if rkey in cmap:
                cls = cmap[rkey]
                for i, d in ent.items():
                    rows.append({'Идентификатор': i, 'Наименование': d.get('title'), 'Описание': d.get('description'), 'Класс': cls})
    if rows: pd.DataFrame(rows).to_excel(writer, sheet_name='Стенды и окружения', index=False)

def save_reverse(ydir: Path, writer: pd.ExcelWriter):
    rows = []
    for p in sorted(ydir.glob('**/*.yaml')):
        data = read_yaml(p)
        for key, val in data.items():
            if key.startswith('seaf.ta.reverse.'):
                ns = key.replace('seaf.ta.reverse.', '')
                for i, v in val.items():
                    rows.append({'Namespace': ns, 'ID': i, 'Name': v.get('name'), 'Type': v.get('type'), 'VPC': v.get('vpc_id'), 'AZ': v.get('az'), 'Subnets': format_list(v.get('subnets')), 'Description': v.get('description')})
    if rows: pd.DataFrame(rows).to_excel(writer, sheet_name='Reverse', index=False)

def main():
    try:
        ensure_deps()
        parser = argparse.ArgumentParser(); parser.add_argument('--config', required=True); args = parser.parse_args()
        cpath = Path(args.config)
        if not cpath.exists(): print(f"ERROR: Config not found: {cpath}", file=sys.stderr); sys.exit(1)
        with cpath.open('r', encoding='utf-8') as f: cfg = yaml.safe_load(f) or {}
        ydir, odir = cpath.parent / cfg.get('yaml_dir', '.'), cpath.parent / cfg.get('out_xlsx_dir', '.')
        odir.mkdir(parents=True, exist_ok=True)
        src_counts = count_entities_in_yaml_dir(ydir)
        for p in [odir / f for f in cfg.get('xlsx_files', [])]:
            try:
                with pd.ExcelWriter(p, engine='openpyxl') as writer:
                    name = p.name.lower()
                    if 'reg' in name: save_regions_az_dc_offices(ydir, writer)
                    if 'seg' in name: save_segments_nets_devices(ydir, writer)
                    if 'kb' in name: save_kb_services(ydir, writer)
                    if 'tech' in name or name == 'ta_services.xlsx' or name == 'services.xlsx': save_tech_services(ydir, writer)
                    if 'comp' in name: save_components(ydir, writer)
                    if 'link' in name: save_links(ydir, writer)
                    if 'stand' in name: save_stands(ydir, writer)
                    if 'reverse' in name: save_reverse(ydir, writer)
                    if not writer.book.sheetnames: pd.DataFrame([{'Info': 'No data'}]).to_excel(writer, sheet_name='Empty')
                print(f"Written data to {p.name}")
            except Exception as e: print(f"ERROR: Failed to write {p.name}: {e}", file=sys.stderr)
        dst_counts = count_entities_in_xlsx([odir / f for f in cfg.get('xlsx_files', [])])
        print("\n--- Conversion Summary ---")
        for k in sorted(list(set(src_counts.keys()) | set(dst_counts.keys()))):
            s, d = src_counts.get(k, 0), dst_counts.get(k, 0)
            print(f"  - {k:<25} | Source: {s:<5} | Dest: {d:<5} | {'OK' if s==d else 'FAIL'}")
    except Exception as e: print(f"FATAL: {e}", file=sys.stderr); sys.exit(1)

if __name__ == '__main__': main()
