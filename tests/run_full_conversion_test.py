#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import subprocess
from pathlib import Path

def run_command(cmd, description):
    """Выполняет команду и выводит результат"""
    print(f"\n--- {description} ---")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    if result.returncode != 0:
        print(f"ERROR: {description} failed")
        print("STDOUT:", result.stdout)
        print("STDERR:", result.stderr)
        return False
    print(result.stdout.strip())
    if result.stderr.strip():
        print("STDERR:", result.stderr.strip())
    return True

def main():
    print("=== Full conversion test SEAF2 -> XLSX -> SEAF1 -> XLSX -> SEAF2 ===")
    
    # Step 1: SEAF2 -> XLSX
    cmd1 = [sys.executable, '_seaf2_yaml_to_xlsx.py', '--config', 'config_seaf2_y2x_1.yaml']
    if not run_command(cmd1, "Step 1: SEAF2 YAML -> XLSX"):
        sys.exit(1)
    
    # Step 2: XLSX -> SEAF1
    cmd2 = [sys.executable, 'xlsx_to_yaml.py', '--config', 'config_x2y_seaf1.yaml']
    if not run_command(cmd2, "Step 2: XLSX -> SEAF1 YAML"):
        sys.exit(1)
    
    # Step 3: SEAF1 -> XLSX
    cmd3 = [sys.executable, 'yaml_to_xlsx.py', '--config', 'config_y2x_seaf1.yaml']
    if not run_command(cmd3, "Step 3: SEAF1 YAML -> XLSX"):
        sys.exit(1)
    
    # Step 4: XLSX -> SEAF2
    cmd4 = [sys.executable, '_seaf2_xlsx_to_yaml.py', '--config', 'config_seaf2_from_xlsx.yaml']
    if not run_command(cmd4, "Step 4: XLSX -> SEAF2 YAML"):
        sys.exit(1)
    
    print("\n=== Full conversion chain completed successfully! ===")

if __name__ == '__main__':
    main()