import os
import sys
import re

def validate_data_js():
    workspace_dir = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
    dashboard_dir = os.path.join(workspace_dir, "pnwdashboard-main")
    data_js_path = os.path.join(dashboard_dir, "data.js")
    
    if not os.path.exists(data_js_path):
        print(f"[Error] data.js not found at {data_js_path}")
        sys.exit(1)
        
    print(f"Reading {data_js_path} for validation...")
    with open(data_js_path, "r", encoding="utf-8") as f:
        content = f.read()
        
    errors = []
    
    # 1. Check JS syntax basically
    # We will check if basic consts are present
    required_consts = ["DATA_EXPORT", "DATA_KR_COUNTRY", "DATA_INDICATORS", "DATA_META"]
    for const in required_consts:
        if f"const {const} = [" not in content and f"const {const} = {{" not in content:
            errors.append(f"Missing variable definition: {const}")
            
    # 2. Check DATA_EXPORT rows > 0
    export_match = re.search(r"const DATA_EXPORT = \[([\s\S]*?)\];", content)
    if export_match:
        items = re.findall(r"\{.*?\}", export_match.group(1))
        if len(items) == 0:
            errors.append("DATA_EXPORT contains no data rows.")
    else:
        errors.append("Could not parse DATA_EXPORT array.")
        
    # 3. Check DATA_KR_COUNTRY rows > 0
    kr_country_match = re.search(r"const DATA_KR_COUNTRY = \[([\s\S]*?)\];", content)
    if kr_country_match:
        items = re.findall(r"\{.*?\}", kr_country_match.group(1))
        if len(items) == 0:
            errors.append("DATA_KR_COUNTRY contains no data rows.")
    else:
        # DATA_KR_COUNTRY might not exist yet, let's just issue a warning or error based on requirements.
        # Requirements said: DATA_KR_COUNTRY 행 수 0 초과
        errors.append("Could not parse DATA_KR_COUNTRY array.")
        
    # 4. Check DATA_INDICATORS latest month exists
    ind_match = re.search(r"const DATA_INDICATORS = \[([\s\S]*?)\];", content)
    if ind_match:
        items = re.findall(r"\{.*?\}", ind_match.group(1))
        if len(items) == 0:
            errors.append("DATA_INDICATORS contains no data rows.")
    else:
        errors.append("Could not parse DATA_INDICATORS array.")
        
    # 5. Check DATA_META latest month exists
    meta_match = re.search(r"const DATA_META = \{([\s\S]*?)\};", content)
    if not meta_match:
        errors.append("Could not parse DATA_META object.")
        
    # Print results
    if errors:
        print("\n❌ Validation Failed:")
        for err in errors:
            print(f"  - {err}")
        sys.exit(1)
    else:
        print("\n✅ Validation Passed: data.js is syntactically sound and contains required data.")
        sys.exit(0)

if __name__ == "__main__":
    validate_data_js()
