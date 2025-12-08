import sys
import os
import shutil
import subprocess
import yaml
from pathlib import Path

# Paths
BASE_DIR = Path.cwd()
TEMP_DIR = BASE_DIR / "temp_repro"

# Scripts
SCRIPT_X2Y_SEAF1 = BASE_DIR / "xlsx_to_yaml.py"
SCRIPT_Y2X_SEAF1 = BASE_DIR / "yaml_to_xlsx.py"
SCRIPT_X2Y_SEAF2 = BASE_DIR / "_seaf2_xlsx_to_yaml.py"
SCRIPT_Y2X_SEAF2 = BASE_DIR / "_seaf2_yaml_to_xlsx.py"

# Source Data
SOURCE_SEAF1 = BASE_DIR / "example/seaf1"
SOURCE_SEAF2 = BASE_DIR / "example/seaf2"

def run_script(script, config_path):
    print(f"Running {script.name} with {config_path.name}...")
    res = subprocess.run([sys.executable, str(script), "--config", str(config_path)], capture_output=True, text=True, encoding='utf-8')
    if res.returncode != 0:
        print(f"ERROR running {script.name}")
        print(res.stderr)
    
    # Check for validation errors in stderr but don't fail the script run if exit code is 0
    if "VALIDATION ERRORS:" in res.stderr:
        print("INFO: Validation errors reported (expected for invalid test data or partial conversion scenarios).")
        print("STDOUT:", res.stdout)
        print("STDERR:", res.stderr)
    else:
        print(res.stdout)
        if res.stderr.strip():
            print("STDERR:", res.stderr.strip())

    if "FAIL" in res.stdout:
        print("!!! FAIL detected in output !!!")

def create_config(path, content):
    with path.open('w', encoding='utf-8') as f:
        yaml.dump(content, f)

def main():
    if TEMP_DIR.exists():
        shutil.rmtree(TEMP_DIR)
    TEMP_DIR.mkdir()

    # Chain 1: Seaf1 -> XLSX -> Seaf2 -> XLSX
    print("\n=== Chain 1: Seaf1 -> XLSX -> Seaf2 -> XLSX ===")
    
    # Step 1: Seaf1 -> XLSX
    c1_s1_config = TEMP_DIR / "c1_s1_config.yaml"
    c1_s1_xlsx_dir = TEMP_DIR / "c1_s1_xlsx"
    c1_s1_xlsx_dir.mkdir()
    
    # Filenames must match heuristics in scripts
    xlsx_files = ['regions.xlsx', 'segments_networks.xlsx', 'kb_services.xlsx']

    create_config(c1_s1_config, {
        'yaml_dir': str(SOURCE_SEAF1),
        'out_xlsx_dir': str(c1_s1_xlsx_dir),
        'xlsx_files': xlsx_files
    })
    run_script(SCRIPT_Y2X_SEAF1, c1_s1_config)

    # Step 2: XLSX -> Seaf2
    c1_s2_config = TEMP_DIR / "c1_s2_config.yaml"
    c1_s2_yaml_dir = TEMP_DIR / "c1_s2_yaml" # Seaf2 output
    
    create_config(c1_s2_config, {
        'xlsx_files': [str(c1_s1_xlsx_dir / f) for f in xlsx_files],
        'out_yaml_dir': str(c1_s2_yaml_dir)
    })
    run_script(SCRIPT_X2Y_SEAF2, c1_s2_config)

    # Step 3: Seaf2 -> XLSX
    c1_s3_config = TEMP_DIR / "c1_s3_config.yaml"
    c1_s3_xlsx_dir = TEMP_DIR / "c1_s3_xlsx"
    
    create_config(c1_s3_config, {
        'yaml_dir': str(c1_s2_yaml_dir),
        'out_xlsx_dir': str(c1_s3_xlsx_dir),
        'xlsx_files': xlsx_files
    })
    run_script(SCRIPT_Y2X_SEAF2, c1_s3_config)


    # Chain 2: Seaf2 -> XLSX -> Seaf1 -> XLSX
    print("\n=== Chain 2: Seaf2 -> XLSX -> Seaf1 -> XLSX ===")
    
    # Step 1: Seaf2 -> XLSX
    c2_s1_config = TEMP_DIR / "c2_s1_config.yaml"
    c2_s1_xlsx_dir = TEMP_DIR / "c2_s1_xlsx"
    c2_s1_xlsx_dir.mkdir()
    
    create_config(c2_s1_config, {
        'yaml_dir': str(SOURCE_SEAF2),
        'out_xlsx_dir': str(c2_s1_xlsx_dir),
        'xlsx_files': xlsx_files
    })
    run_script(SCRIPT_Y2X_SEAF2, c2_s1_config)

    # Step 2: XLSX -> Seaf1
    c2_s2_config = TEMP_DIR / "c2_s2_config.yaml"
    c2_s2_yaml_dir = TEMP_DIR / "c2_s2_yaml" # Seaf1 output
    
    create_config(c2_s2_config, {
        'xlsx_files': [str(c2_s1_xlsx_dir / f) for f in xlsx_files],
        'out_yaml_dir': str(c2_s2_yaml_dir)
    })
    run_script(SCRIPT_X2Y_SEAF1, c2_s2_config)

    # Step 3: Seaf1 -> XLSX
    c2_s3_config = TEMP_DIR / "c2_s3_config.yaml"
    c2_s3_xlsx_dir = TEMP_DIR / "c2_s3_xlsx"
    
    create_config(c2_s3_config, {
        'yaml_dir': str(c2_s2_yaml_dir),
        'out_xlsx_dir': str(c2_s3_xlsx_dir),
        'xlsx_files': xlsx_files
    })
    run_script(SCRIPT_Y2X_SEAF1, c2_s3_config)

if __name__ == "__main__":
    main()
