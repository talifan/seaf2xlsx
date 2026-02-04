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
    return ', '.join(sorted([str(x) for x in items if x]))

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

# SEAF2 handling plural keys
ENTITY_MAP = {
    'dc_region': ['seaf.company.ta.services.dc_regions'],
    'dc_az': ['seaf.company.ta.services.dc_azs'],
    'dc': ['seaf.company.ta.services.dcs'],
    'office': ['seaf.company.ta.services.dc_offices'],
    'network_segment': ['seaf.company.ta.services.network_segments'],
    'network': ['seaf.company.ta.services.networks'],
    'kb': ['seaf.company.ta.services.kbs'],
    'components.network': ['seaf.company.ta.components.networks'],
    'compute_service': ['seaf.company.ta.services.compute_services'],
    'cluster': ['seaf.company.ta.services.clusters'],
    'monitoring': ['seaf.company.ta.services.monitorings'],
    'backup': ['seaf.company.ta.services.backups'],
    'software': ['seaf.company.ta.services.softwares'],
    'storage': ['seaf.company.ta.services.storages'],
    'cluster_virtualization': ['seaf.company.ta.services.cluster_virtualizations'],
    'k8s': ['seaf.company.ta.services.k8s'],
    'k8s_deployment': ['seaf.company.ta.services.k8s_deployments'],
    'logical_link': ['seaf.company.ta.services.logical_links'],
    'network_link': ['seaf.company.ta.services.network_links'],
    'stand': ['seaf.company.ta.services.stands'],
    'environment': ['seaf.company.ta.services.environments'],
    'server': ['seaf.company.ta.components.servers'],
    'hw_storage': ['seaf.company.ta.components.hw_storages'],
    'user_device': ['seaf.company.ta.components.user_devices'],
    'k8s_node': ['seaf.company.ta.components.k8s_nodes'],
    'k8s_namespace': ['seaf.company.ta.components.k8s_namespaces'],
    'k8s_hpa': ['seaf.company.ta.components.k8s_hpa']
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
                ename = rev_lookup[key]
                counts[ename] = counts.get(ename, 0) + len(val)
    return counts

def count_entities_in_xlsx(xlsx_files: List[Path]) -> Dict[str, int]:
    counts = {}
    sheet_map = {'Регионы': 'regions', 'AZ': 'dc_az', 'DC': 'dc', 'Офисы': 'office', 'Сегменты': 'network_segment', 'Сети': 'network', 'Сервисы КБ': 'kb', 'Тех. сервисы': 'tech_services', 'Компоненты': 'components', 'Сетевые устройства': 'components', 'Связи': 'links', 'Стенды и окружения': 'stands'}
    for fp in xlsx_files:
        if not fp.exists(): continue
        try:
            xls = pd.ExcelFile(fp)
            for sn in xls.sheet_names:
                if sn in sheet_map:
                    df = xls.parse(sn).dropna(how='all')
                    ename = sheet_map[sn]
                    if ename == 'tech_services':
                        cmap = {'Compute Service': 'compute_service', 'Cluster': 'cluster', 'Monitoring': 'monitoring', 'Backup': 'backup', 'Software': 'software', 'Storage': 'storage', 'Cluster Virtualization': 'cluster_virtualization', 'K8s Cluster': 'k8s', 'Deployment': 'k8s_deployment'}
                        for _, row in df.iterrows():
                            etype = cmap.get(row.get('Класс'), 'compute_service')
                            counts[etype] = counts.get(etype, 0) + 1
                    elif ename == 'components':
                        cmap = {'Server': 'server', 'HW Storage': 'hw_storage', 'User Device': 'user_device', 'K8s Node': 'k8s_node', 'K8s Namespace': 'k8s_namespace', 'K8s HPA': 'k8s_hpa', 'Network Device': 'components.network'}
                        for _, row in df.iterrows():
                            etype = cmap.get(row.get('Класс'), 'components.network')
                            counts[etype] = counts.get(etype, 0) + 1
                    else: counts[ename] = counts.get(ename, 0) + len(df)
        except Exception: continue
    return counts

def save_regions_az_dc_offices(ydir: Path, writer: pd.ExcelWriter):
    for fn, key, sn in [('dc_region.yaml', 'seaf.company.ta.services.dc_regions', 'Регионы'), ('dc_az.yaml', 'seaf.company.ta.services.dc_azs', 'AZ'), ('dc.yaml', 'seaf.company.ta.services.dcs', 'DC'), ('dc_office.yaml', 'seaf.company.ta.services.dc_offices', 'Офисы')]:
        d = read_yaml(ydir / fn).get(key, {})
        if d:
            rows = []
            for i, p in d.items():
                if sn == 'Регионы': rows.append({'ID Региона': i, 'Наименование': p.get('title'), 'Описание': p.get('description')})
                elif sn == 'AZ': rows.append({'ID AZ': i, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Поставщик': p.get('vendor'), 'Регион': p.get('region')})
                elif sn == 'DC': rows.append({'ID DC': i, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Поставщик': p.get('vendor'), 'Tier': p.get('tier'), 'Тип': p.get('type'), 'Кол-во стоек': p.get('rack_qty'), 'Адрес': p.get('address'), 'Форма владения': p.get('ownership'), 'AZ': p.get('availabilityzone')})
                elif sn == 'Офисы': rows.append({'ID Офиса': i, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Адрес': p.get('address'), 'Регион': p.get('region')})
            if rows: pd.DataFrame(rows).to_excel(writer, sheet_name=sn, index=False)

def save_segments_nets_devices(ydir: Path, writer: pd.ExcelWriter):
    d = read_yaml(ydir / 'network_segment.yaml').get('seaf.company.ta.services.network_segments', {})
    if d: pd.DataFrame([{'ID сетевые сегмента/зоны': i, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Расположение': (p.get('sber') or {}).get('location'), 'Зона': (p.get('sber') or {}).get('zone')} for i, p in d.items()]).to_excel(writer, sheet_name='Сегменты', index=False)
    nets = {}
    for p in sorted(ydir.glob('network*.yaml')):
        if not p.name.startswith('_') and p.name not in ['network_segment.yaml', 'network_component.yaml']:
            nets.update(read_yaml(p).get('seaf.company.ta.services.networks', {}))
    if nets:
        rows = [{'ID Network': i, 'Наименование': p.get('title'), 'Описание': p.get('description'), 'Тип сети': p.get('type'), 'VLAN': p.get('vlan'), 'VRF  ': p.get('VRF'), 'Провайдер': p.get('provider') or (p.get('sber') or {}).get('provider'), 'Тип сети (проводная, беспроводная)': p.get('lan_type'), 'Адрес сети': p.get('ipnetwork'), 'WAN Адрес': p.get('wan_ip'), 'Расположение': format_list(p.get('location')), 'Сетевой сегмент/зона(ID)': format_list(p.get('segment'))} for i, p in nets.items()]
        pd.DataFrame(rows).to_excel(writer, sheet_name='Сети', index=False)
    devs = {}
    for p in sorted(ydir.glob('network_component*.yaml')): devs.update(read_yaml(p).get('seaf.company.ta.components.networks', {}))
    if devs: pd.DataFrame([{'ID Устройства': i, 'Наименование': p.get('title'), 'Тип реализации': p.get('realization_type'), 'Тип': p.get('type'), 'Модель': p.get('model'), 'Назначение': p.get('purpose'), 'IP адрес': p.get('address'), 'Описание': p.get('description'), 'Расположение (ID сегмента/зоны)': p.get('segment'), 'Подключенные сети (список)': format_list(p.get('network_connection'))} for i, p in devs.items()]).to_excel(writer, sheet_name='Сетевые устройства', index=False)

def save_kb_services(ydir: Path, writer: pd.ExcelWriter):
    d = read_yaml(ydir / 'kb.yaml').get('seaf.company.ta.services.kbs', {})
    if not d: return
    rows = [{'ID КБ сервиса': i, 'Tag': p.get('tag'), 'Описание': p.get('description'), 'Технология': p.get('technology'), 'Название ПО': p.get('software_name'), 'Статус': p.get('status'), 'Подключенные сети': format_list(p.get('network_connection'))} for i, p in d.items()]
    pd.DataFrame(rows).to_excel(writer, sheet_name='Сервисы КБ', index=False)

def save_tech_services(ydir: Path, writer: pd.ExcelWriter):
    rows = []
    kmap = {
        'seaf.company.ta.services.compute_services': 'Compute Service',
        'seaf.company.ta.services.clusters': 'Cluster',
        'seaf.company.ta.services.monitorings': 'Monitoring',
        'seaf.company.ta.services.backups': 'Backup',
        'seaf.company.ta.services.softwares': 'Software',
        'seaf.company.ta.services.storages': 'Storage',
        'seaf.company.ta.services.cluster_virtualizations': 'Cluster Virtualization',
        'seaf.company.ta.services.k8s': 'K8s Cluster',
        'seaf.company.ta.services.k8s_deployments': 'Deployment'
    }
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
                    if 'compute' in rkey or 'cluster' in rkey: obj['Тип сервиса'] = normalize_val(d.get('service_type'), 'Серверы приложений и т.д.')
                    if 'cluster' in rkey and 'reservation_type' in d: obj['Тип резервирования'] = d['reservation_type']
                    rows.append(obj)
    if rows: pd.DataFrame(rows).to_excel(writer, sheet_name='Тех. сервисы', index=False)

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
                    if 'tech' in name or name == 'services.xlsx': save_tech_services(ydir, writer)
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