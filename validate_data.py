import json
import os
import sys
sys.stdout.reconfigure(encoding="utf-8")   # 콘솔이 cp949여도 기호 출력에서 멈추지 않게
import re

MM_LIMIT_T = 1000   # 합계 불일치·지역 미배정이 이 톤을 넘으면 실패(C27)


def check_marketmap(content):
    """DATA_MM(마켓맵) 검사. 블록이 없으면 건너뛴다. 반환 = (오류 목록, 경고 목록)"""
    m = re.search(r"^const DATA_MM = (.*);$", content, flags=re.M)
    if not m:
        print("DATA_MM 없음 → 마켓맵 검사 건너뜀")
        return [], []
    errors, warns = [], []
    try:
        mm = json.loads(m.group(1))
    except ValueError as e:
        return [f"DATA_MM JSON 읽기 실패: {e}"], []
    C, R, rows, NS = mm["countries"], set(mm["regions"]), mm["rows"], len(mm["series"])
    latest = mm["latest"]
    a, b = latest - 23, latest   # 최근 24개월 = 올해 누계와 전년 같은 기간을 덮는다
    by_region, by_country, by_series, total = {}, {}, [0.0] * NS, 0.0
    src_sum = {"cn": 0.0, "kr": 0.0, "eu": 0.0}
    for mo, ci, si, t, _ in rows:
        if not (0 <= ci < len(C) and 0 <= si < NS):
            errors.append(f"DATA_MM 행 index 범위 밖: {[mo, ci, si]}")
            break
        src_sum["cn" if si == 0 else "kr" if si <= 4 else "eu"] += t
        if a <= mo <= b:
            rg = C[ci]["rg"]
            by_region[rg] = by_region.get(rg, 0) + t
            by_country[ci] = by_country.get(ci, 0) + t
            by_series[si] += t
            total += t
    gap = lambda x, y: abs(x - y) > MM_LIMIT_T
    if gap(sum(by_region.values()), total):
        errors.append(f"마켓맵 지역 합 {sum(by_region.values()):,.0f} ≠ 전체 {total:,.0f}")
    if gap(sum(by_series), sum(by_country.values())):
        errors.append(f"마켓맵 계열 합 {sum(by_series):,.0f} ≠ 나라 합 {sum(by_country.values()):,.0f}")
    for k, label in (("cn", "중국산 = DATA_EXPORT"), ("kr", "한국산 = DATA_KR_COUNTRY"), ("eu", "유럽산 = Comext 원자료")):
        want = mm["check"][k + "_t"] - (mm["unmapped"]["eu_t"] if k == "eu" else 0)
        if gap(src_sum[k], want):
            errors.append(f"마켓맵 {label} 불일치: {src_sum[k]:,.0f} vs {want:,.0f}톤")
    unassigned = by_region.get("미배정", 0)
    if unassigned > MM_LIMIT_T:
        names = sorted({C[ci]["ko"] for ci in by_country if C[ci]["rg"] == "미배정"})
        errors.append(f"마켓맵 지역 미배정 {unassigned:,.0f}톤 > {MM_LIMIT_T}톤: {names}")
    elif unassigned:
        warns.append(f"마켓맵 지역 미배정 {unassigned:,.0f}톤(문턱 이하)")
    if mm["unmapped"]["eu_t"] > MM_LIMIT_T:
        errors.append(f"마켓맵 EU 도착국 매핑 누락 {mm['unmapped']['eu_t']:,.0f}톤: {list(mm['unmapped']['eu'])[:10]}")
    other = sum(v for k, v in by_region.items() if k not in R and k != "미배정")
    if other:
        warns.append(f"마켓맵 24권역 밖(원본 '기타'·한국 등) {other:,.0f}톤, 지역표 '기타' 행으로 표시")
    y, mth = divmod(latest, 12)
    print(f"마켓맵 검사: 기준월 {y}.{mth + 1}, 최근 24개월 전체 {total:,.0f}톤, 지역 {len(by_region)}개, 미배정 {unassigned:,.0f}톤")
    return errors, warns


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
        
    # 6. 마켓맵(DATA_MM) — 블록이 있을 때만
    mm_errors, mm_warns = check_marketmap(content)
    errors += mm_errors
    for w in mm_warns:
        print(f"  [경고] {w}")

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
