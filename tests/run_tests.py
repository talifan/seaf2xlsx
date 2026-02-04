import sys
from pathlib import Path
import subprocess
import yaml
import pandas as pd

# --- Test Configuration ---
CWD = Path.cwd()

# Directories
YAML_SOURCE_DIR = CWD / 'example/seaf1'
XLSX_INITIAL_DIR = CWD / 'out_xlsx_initial'
YAML_ROUNDTRIP_DIR = CWD / 'out_yaml_roundtrip'
XLSX_FINAL_DIR = CWD / 'out_xlsx_final'

# Config file names
YAML_TO_XLSX_CONFIG_1 = CWD / 'config_y2x_1.yaml'
XLSX_TO_YAML_CONFIG = CWD / 'config_x2y.yaml'
YAML_TO_XLSX_CONFIG_2 = CWD / 'config_y2x_2.yaml'

# --- Helper Functions ---

def run_script(script_name: str, config_path: Path) -> bool:
    """Runs a python script with a given config file."""
    cmd = [sys.executable, script_name, '--config', str(config_path)]
    print(f"--- Running: {' '.join(str(c) for c in cmd)} ---")
    result = subprocess.run(cmd, capture_output=True, text=True, encoding='utf-8')
    if result.returncode != 0:
        print(f"ERROR: Script {script_name} failed.")
        print("Stdout:", result.stdout)
        print("Stderr:", result.stderr)
        return False
    print(result.stdout.strip())
    if result.stderr.strip():
        print("Stderr:", result.stderr.strip())
    return True

def compare_xlsx_files(initial_dir: Path, final_dir: Path) -> dict:
    """Comparisons all XLSX files between two directories sheet by sheet."""
    diffs = {}
    initial_files = sorted(list(initial_dir.glob('*.xlsx')))
    final_files = sorted(list(final_dir.glob('*.xlsx')))

    if len(initial_files) != len(final_files):
        return {"file_count_mismatch": f"Initial: {len(initial_files)}, Final: {len(final_files)}"}

    for initial_path, final_path in zip(initial_files, final_files):
        if initial_path.name != final_path.name:
            diffs[f"{initial_path.name} vs {final_path.name}"] = "Filename mismatch"
            continue

        try:
            initial_xls = pd.ExcelFile(initial_path)
            final_xls = pd.ExcelFile(final_path)

            if initial_xls.sheet_names != final_xls.sheet_names:
                diffs[initial_path.name] = f"Sheet names differ. Initial: {initial_xls.sheet_names}, Final: {final_xls.sheet_names}"
                continue

            for sheet_name in initial_xls.sheet_names:
                if sheet_name == '-----':
                    continue
                df_initial = initial_xls.parse(sheet_name).fillna('')
                df_final = final_xls.parse(sheet_name).fillna('')
                
                first_col = df_initial.columns[0]
                df_initial = df_initial.sort_values(by=first_col).reset_index(drop=True)
                df_final = df_final.sort_values(by=first_col).reset_index(drop=True)

                if not df_initial.equals(df_final):
                    diffs.setdefault(initial_path.name, {})[sheet_name] = "Sheet content differs"

        except Exception as e:
            diffs[initial_path.name] = f"Error comparing file: {e}"
            
    return diffs

def main():
    """Main test execution function."""
    # --- Preparation ---
    print("--- Preparing test environment ---")
    for d in [XLSX_INITIAL_DIR, YAML_ROUNDTRIP_DIR, XLSX_FINAL_DIR]:
        d.mkdir(exist_ok=True)
        for f in d.glob('*'): f.unlink()

    # --- Create Configs ---
    with open(YAML_TO_XLSX_CONFIG_1, 'w', encoding='utf-8') as f:
        yaml.dump({'yaml_dir': str(YAML_SOURCE_DIR), 'out_xlsx_dir': str(XLSX_INITIAL_DIR), 'xlsx_files': ['regions_az_dc_offices.xlsx', 'segments_nets_netdevices.xlsx', 'kb_services.xlsx']}, f)
    
    with open(XLSX_TO_YAML_CONFIG, 'w', encoding='utf-8') as f:
        yaml.dump({'xlsx_files': [str(XLSX_INITIAL_DIR / 'regions_az_dc_offices.xlsx'), str(XLSX_INITIAL_DIR / 'segments_nets_netdevices.xlsx'), str(XLSX_INITIAL_DIR / 'kb_services.xlsx')], 'out_yaml_dir': str(YAML_ROUNDTRIP_DIR)}, f)

    with open(YAML_TO_XLSX_CONFIG_2, 'w', encoding='utf-8') as f:
        yaml.dump({'yaml_dir': str(YAML_ROUNDTRIP_DIR), 'out_xlsx_dir': str(XLSX_FINAL_DIR), 'xlsx_files': ['regions_az_dc_offices.xlsx', 'segments_nets_netdevices.xlsx', 'kb_services.xlsx']}, f)

    # --- STEP 1: YAML -> XLSX (Initial) ---
    print("\n[Step 1/4] Converting source YAML to INITIAL XLSX...")
    if not run_script('yaml_to_xlsx.py', YAML_TO_XLSX_CONFIG_1):
        sys.exit(1)

    # --- STEP 2: XLSX -> YAML (Roundtrip) ---
    print("\n[Step 2/4] Converting INITIAL XLSX to ROUNDTRIP YAML...")
    if not run_script('xlsx_to_yaml.py', XLSX_TO_YAML_CONFIG):
        sys.exit(1)

    # --- STEP 3: YAML -> XLSX (Final) ---
    print("\n[Step 3/4] Converting ROUNDTRIP YAML to FINAL XLSX...")
    if not run_script('yaml_to_xlsx.py', YAML_TO_XLSX_CONFIG_2):
        sys.exit(1)

    # --- STEP 4: Verification ---
    print("\n[Step 4/4] Verifying conversion results by comparing XLSX files...")
    differences = compare_xlsx_files(XLSX_INITIAL_DIR, XLSX_FINAL_DIR)

    # --- Final Report ---
    print("\n--- TEST RESULT ---")
    if not differences:
        print("STATUS: PASSED")
        print("XLSX files are identical after a full round-trip conversion.")
    else:
        print("STATUS: FAILED")
        print("Reason: XLSX files differ after round-trip.")
        print("Differences found:")
        for file, diff in differences.items():
            print(f"  - In file {file}: {diff}")

    # --- Cleanup ---
    for f in [YAML_TO_XLSX_CONFIG_1, XLSX_TO_YAML_CONFIG, YAML_TO_XLSX_CONFIG_2]:
        f.unlink()

if __name__ == '__main__':
    main()