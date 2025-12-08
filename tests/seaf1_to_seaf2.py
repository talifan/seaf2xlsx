#!/usr/bin/env python3
import sys
import argparse
from pathlib import Path
from typing import Dict, Any
import yaml

def convert_seaf1_to_seaf2(input_dir: Path, output_dir: Path):
    """Конвертирует файлы из формата SEAF1 в SEAF2"""
    print(f"Converting SEAF1 to SEAF2...")
    print(f"Input directory: {input_dir}")
    print(f"Output directory: {output_dir}")
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Mapping of entity name changes
    entity_mappings = {
        'dc_region': 'dc_regions',
        'dc_az': 'dc_azs',
        'dc': 'dcs',
        'office': 'dc_offices',
        'network_segment': 'network_segments',
        'network': 'networks',
        'kb': 'kbs'
    }
    
    # Mapping of namespace changes
    namespace_mappings = {
        'seaf.ta.services.dc_region': 'seaf.company.ta.services.dc_regions',
        'seaf.ta.services.dc_az': 'seaf.company.ta.services.dc_azs',
        'seaf.ta.services.dc': 'seaf.company.ta.services.dcs',
        'seaf.ta.services.office': 'seaf.company.ta.services.dc_offices',
        'seaf.ta.services.network_segment': 'seaf.company.ta.services.network_segments',
        'seaf.ta.services.network': 'seaf.company.ta.services.networks',
        'seaf.ta.services.kb': 'seaf.company.ta.services.kbs',
        'seaf.ta.components.network': 'seaf.company.ta.components.networks'
    }
    
    # File name mappings
    file_mappings = {
        'components_network.yaml': 'network_component.yaml',
        'office.yaml': 'dc_office.yaml',
        'root.yaml': '_root.yaml'
    }
    
    converted_files = 0
    
    for yaml_file in input_dir.glob('*.yaml'):
        if yaml_file.name.startswith('_'):  # Skip _root.yaml
            continue
            
        with yaml_file.open('r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
            
        if not isinstance(data, dict):
            continue
            
        new_data = {}
        
        for key, value in data.items():
            # Update namespace
            new_key = namespace_mappings.get(key, key)
            
            if isinstance(value, dict):
                # Update entity names within the data
                new_value = {}
                for entity_id, entity_data in value.items():
                    new_value[entity_id] = entity_data
                new_data[new_key] = new_value
            else:
                new_data[new_key] = value
        
        # Determine output filename
        output_filename = file_mappings.get(yaml_file.name, yaml_file.name)
        output_path = output_dir / output_filename
        
        # Write converted file
        with output_path.open('w', encoding='utf-8') as f:
            yaml.dump(new_data, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
        
        print(f"  Converted: {yaml_file.name} → {output_filename}")
        converted_files += 1
    
    print(f"Conversion complete. {converted_files} files converted.")

def convert_seaf2_to_seaf1(input_dir: Path, output_dir: Path):
    """Конвертирует файлы из формата SEAF2 в SEAF1"""
    print(f"Converting SEAF2 to SEAF1...")
    print(f"Input directory: {input_dir}")
    print(f"Output directory: {output_dir}")
    
    output_dir.mkdir(parents=True, exist_ok=True)
    
    # Reverse mappings
    namespace_mappings = {
        'seaf.company.ta.services.dc_regions': 'seaf.ta.services.dc_region',
        'seaf.company.ta.services.dc_azs': 'seaf.ta.services.dc_az',
        'seaf.company.ta.services.dcs': 'seaf.ta.services.dc',
        'seaf.company.ta.services.dc_offices': 'seaf.ta.services.office',
        'seaf.company.ta.services.network_segments': 'seaf.ta.services.network_segment',
        'seaf.company.ta.services.networks': 'seaf.ta.services.network',
        'seaf.company.ta.services.kbs': 'seaf.ta.services.kb',
        'seaf.company.ta.components.networks': 'seaf.ta.components.network'
    }
    
    file_mappings = {
        'network_component.yaml': 'components_network.yaml',
        'dc_office.yaml': 'office.yaml',
        '_root.yaml': 'root.yaml'
    }
    
    converted_files = 0
    
    for yaml_file in input_dir.glob('*.yaml'):
        if yaml_file.name.startswith('_'):  # Skip _root.yaml
            continue
            
        with yaml_file.open('r', encoding='utf-8') as f:
            data = yaml.safe_load(f)
            
        if not isinstance(data, dict):
            continue
            
        new_data = {}
        
        for key, value in data.items():
            # Update namespace
            new_key = namespace_mappings.get(key, key)
            new_data[new_key] = value
        
        # Determine output filename
        output_filename = file_mappings.get(yaml_file.name, yaml_file.name)
        output_path = output_dir / output_filename
        
        # Write converted file
        with output_path.open('w', encoding='utf-8') as f:
            yaml.dump(new_data, f, default_flow_style=False, allow_unicode=True, sort_keys=False)
        
        print(f"  Converted: {yaml_file.name} → {output_filename}")
        converted_files += 1
    
    print(f"Conversion complete. {converted_files} files converted.")

def main():
    parser = argparse.ArgumentParser(description='Convert between SEAF1 and SEAF2 formats')
    parser.add_argument('direction', choices=['seaf1-to-seaf2', 'seaf2-to-seaf1'], 
                       help='Conversion direction')
    parser.add_argument('input_dir', type=str, help='Input directory path')
    parser.add_argument('output_dir', type=str, help='Output directory path')
    
    args = parser.parse_args()
    
    input_dir = Path(args.input_dir)
    output_dir = Path(args.output_dir)
    
    if not input_dir.exists():
        print(f"Error: Input directory {input_dir} does not exist.", file=sys.stderr)
        sys.exit(1)
    
    if args.direction == 'seaf1-to-seaf2':
        convert_seaf1_to_seaf2(input_dir, output_dir)
    elif args.direction == 'seaf2-to-seaf1':
        convert_seaf2_to_seaf1(input_dir, output_dir)

if __name__ == '__main__':
    main()