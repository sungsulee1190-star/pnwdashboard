"""
인쇄용지 경쟁 분석 대시보드 — 월별 데이터 업데이트 스크립트
=================================================================
사용법:
  1. Excel 파일을 DRM 없이 저장 (Excel에서 열고 → 다른 이름으로 저장 → 일반 .xlsx)
  2. 이 스크립트를 실행:  python update_data.py

업데이트 대상:
  - 한국수출실적_*.xlsx  → DATA_KR_EXPORT, DATA_KR_REGION, DATA_KR_YEARLY, DATA_KR_TOTAL
  - 중국수출실적_*.xlsx  → DATA_EXPORT (중국산 국가별 수출량)
  - 26년_중국 아트지 가격.xlsx → DATA_PRICES (중국 내수가)

설치 필요 패키지:
  pip install pandas openpyxl
"""

import pandas as pd
import json
import re
import sys
from pathlib import Path
from datetime import datetime

# ─────────────────────────────────────────────
# 경로 설정
# ─────────────────────────────────────────────
BASE_DIR   = Path(__file__).parent
EXCEL_DIR  = BASE_DIR / "참조_백데이터_엑셀"
DATA_JS    = BASE_DIR / "data.js"

# ─────────────────────────────────────────────
# 엑셀 파일 자동 검색 (날짜 suffix 무시)
# ─────────────────────────────────────────────
def find_excel(pattern: str) -> Path | None:
    """패턴에 맞는 최신 파일 반환 (예: '한국수출실적'"""
    files = sorted(EXCEL_DIR.glob(f"*{pattern}*.xlsx"), key=lambda p: p.stat().st_mtime, reverse=True)
    if not files:
        print(f"  ⚠️  '{pattern}' 파일을 찾을 수 없습니다. 건너뜁니다.")
        return None
    print(f"  ✅ 발견: {files[0].name}")
    return files[0]

# ─────────────────────────────────────────────
# 공통 유틸
# ─────────────────────────────────────────────
def q_label(m: int) -> str:
    return f"Q{(m-1)//3+1}"

def h_label(m: int) -> str:
    return "H1" if m <= 6 else "H2"

def safe_float(v, default=0.0) -> float:
    try:
        f = float(v)
        return round(f, 1) if not pd.isna(f) else default
    except (TypeError, ValueError):
        return default

def safe_int(v, default=0) -> int:
    try:
        f = float(v)
        return int(f) if not pd.isna(f) else default
    except (TypeError, ValueError):
        return default

def js_array(name: str, rows: list[dict]) -> str:
    """Python dict 리스트 → JS const 선언 문자열"""
    lines = [f"const {name}=["]
    for r in rows:
        parts = []
        for k, v in r.items():
            if isinstance(v, str):
                parts.append(f'{k}:"{v}"')
            elif isinstance(v, float):
                parts.append(f"{k}:{v:.1f}" if v != int(v) else f"{k}:{int(v)}")
            else:
                parts.append(f"{k}:{v}")
        lines.append("  {" + ",".join(parts) + "},")
    lines.append("];")
    return "\n".join(lines)

def inject_js(content: str, name: str, new_block: str) -> str:
    """data.js 안의 특정 const 블록을 새 블록으로 교체"""
    # const NAME=[...]; 또는 const NAME = [...]; 패턴 (다중행)
    pattern = rf"const\s+{re.escape(name)}\s*=\s*\[.*?\];"
    if not re.search(pattern, content, re.DOTALL):
        print(f"  ⚠️  {name} 블록을 data.js에서 찾지 못했습니다. 파일 끝에 추가합니다.")
        return content + "\n\n" + new_block + "\n"
    return re.sub(pattern, new_block, content, flags=re.DOTALL)

# ═══════════════════════════════════════════════════════════════
# 1. 한국산 수출실적
# ═══════════════════════════════════════════════════════════════
def process_kr_export(path: Path) -> dict:
    """
    예상 시트 구조 (실제 파일 열어서 확인 후 아래 컬럼명 조정):

    [시트: "실적" 또는 첫번째 시트]
    columns: 년도, 월, 경쟁사, 지역(대분류), 국가, 수출량(MT), 수출금액(USD)
    경쟁사 값: 한솔 | 무림 | 한국 | 무림+한국

    ※ 실제 컬럼명이 다르면 아래 COL_MAP 딕셔너리를 수정하세요
    """
    COL_MAP = {
        "year":    ["년도", "연도", "year", "Year", "YR"],
        "month":   ["월", "month", "Month", "MO"],
        "company": ["경쟁사", "회사", "company", "Company", "COMP", "구분"],
        "region":  ["지역", "대분류", "region", "Region", "RG"],
        "country": ["국가", "country", "Country"],
        "volume":  ["수출량", "물량", "MT", "volume", "Volume", "수량"],
        "amount":  ["수출금액", "금액", "USD", "amount", "Amount", "AMT"],
    }

    print(f"\n[1] 한국산 수출실적 처리 중: {path.name}")

    # 시트 목록 확인
    xl = pd.ExcelFile(path)
    print(f"    시트: {xl.sheet_names}")
    # 첫 번째 시트 (또는 '실적' 시트)
    sheet = "실적" if "실적" in xl.sheet_names else xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet, header=0)
    print(f"    컬럼: {list(df.columns)}")
    print(f"    행수: {len(df)}")

    # 컬럼 자동 매핑
    col = {}
    for key, candidates in COL_MAP.items():
        for c in candidates:
            matches = [x for x in df.columns if str(x).strip() == c]
            if matches:
                col[key] = matches[0]
                break
    missing = [k for k in ["year","month","company","volume"] if k not in col]
    if missing:
        print(f"  ❌ 컬럼 매핑 실패: {missing}")
        print(f"     파일 컬럼: {list(df.columns)}")
        print(f"     update_data.py 상단 COL_MAP을 수정하세요.")
        return {}

    has_country = "country" in col
    has_region  = "region"  in col
    has_amount  = "amount"  in col

    # 데이터 정리
    df = df.dropna(subset=[col["year"], col["month"], col["company"], col["volume"]])
    df[col["year"]]   = df[col["year"]].apply(safe_int)
    df[col["month"]]  = df[col["month"]].apply(safe_int)
    df[col["volume"]] = df[col["volume"]].apply(safe_float)
    if has_amount:
        df[col["amount"]] = df[col["amount"]].apply(safe_float)
    df = df[(df[col["year"]] >= 2018) & (df[col["month"]].between(1,12)) & (df[col["volume"]] >= 0)]

    rows_export  = []  # DATA_KR_EXPORT (회사별 월합계)
    rows_region  = []  # DATA_KR_REGION (회사+지역별)
    rows_country = []  # DATA_KR_COUNTRY (회사+국가별) — 신규

    # ── 회사별 월합계 (DATA_KR_EXPORT)
    grp_cols = [col["year"], col["month"], col["company"]]
    agg = {col["volume"]: "sum"}
    if has_amount: agg[col["amount"]] = "sum"
    g = df.groupby(grp_cols).agg(agg).reset_index()

    for _, r in g.iterrows():
        y = safe_int(r[col["year"]])
        m = safe_int(r[col["month"]])
        v = safe_float(r[col["volume"]])
        a = safe_float(r[col["amount"]]) if has_amount else 0.0
        fob = round(a / v, 1) if v > 0 else 0.0
        rows_export.append({"y":y,"m":m,"q":q_label(m),"h":h_label(m),
                             "comp":str(r[col["company"]]).strip(),"v":safe_int(v),"amt":safe_int(a),"fob":fob})

    # ── 지역별 (DATA_KR_REGION)
    if has_region:
        grp_cols_r = [col["year"], col["month"], col["company"], col["region"]]
        agg_r = {col["volume"]: "sum"}
        if has_amount: agg_r[col["amount"]] = "sum"
        gr = df.groupby(grp_cols_r).agg(agg_r).reset_index()
        for _, r in gr.iterrows():
            y = safe_int(r[col["year"]]); m = safe_int(r[col["month"]])
            v = safe_float(r[col["volume"]]); a = safe_float(r[col["amount"]]) if has_amount else 0.0
            if v == 0: continue
            fob = round(a / v, 1) if v > 0 else 0.0
            rows_region.append({"y":y,"m":m,"q":q_label(m),"h":h_label(m),
                                 "comp":str(r[col["company"]]).strip(),
                                 "rg":str(r[col["region"]]).strip(),
                                 "v":safe_int(v),"amt":safe_int(a),"fob":fob})

    # ── 국가별 (DATA_KR_COUNTRY) — 신규
    if has_country:
        grp_cols_c = [col["year"], col["month"], col["company"], col["country"]]
        if has_region: grp_cols_c.append(col["region"])
        agg_c = {col["volume"]: "sum"}
        if has_amount: agg_c[col["amount"]] = "sum"
        gc = df.groupby(grp_cols_c).agg(agg_c).reset_index()
        for _, r in gc.iterrows():
            y = safe_int(r[col["year"]]); m = safe_int(r[col["month"]])
            v = safe_float(r[col["volume"]]); a = safe_float(r[col["amount"]]) if has_amount else 0.0
            if v == 0: continue
            fob = round(a / v, 1) if v > 0 else 0.0
            row = {"y":y,"m":m,"q":q_label(m),"h":h_label(m),
                   "comp":str(r[col["company"]]).strip(),
                   "c":str(r[col["country"]]).strip(),
                   "v":safe_int(v),"amt":safe_int(a),"fob":fob}
            if has_region: row["rg"] = str(r[col["region"]]).strip()
            rows_country.append(row)

    # ── 연간 집계 (DATA_KR_YEARLY)
    rows_yearly = []
    gy = g.groupby([col["year"], col["company"]]).agg({col["volume"]:"sum", **({col["amount"]:"sum"} if has_amount else {})}).reset_index()
    for _, r in gy.iterrows():
        v = safe_float(r[col["volume"]]); a = safe_float(r[col["amount"]]) if has_amount else 0.0
        fob = round(a / v, 1) if v > 0 else 0.0
        rows_yearly.append({"y":safe_int(r[col["year"]]),"comp":str(r[col["company"]]).strip(),"v":safe_int(v),"amt":safe_int(a),"fob":fob})

    # ── 전체 연간 합계 (DATA_KR_TOTAL)
    rows_total = []
    gt = g.groupby(col["year"]).agg({col["volume"]:"sum", **({col["amount"]:"sum"} if has_amount else {})}).reset_index()
    for _, r in gt.iterrows():
        v = safe_float(r[col["volume"]]); a = safe_float(r[col["amount"]]) if has_amount else 0.0
        fob = round(a / v, 1) if v > 0 else 0.0
        rows_total.append({"y":safe_int(r[col["year"]]),"v":safe_int(v),"amt":safe_int(a),"fob":fob})

    # 정렬
    rows_export.sort(key=lambda x: (x["y"], x["m"], x["comp"]))
    rows_region.sort(key=lambda x: (x["y"], x["m"], x["comp"], x.get("rg","")))
    rows_country.sort(key=lambda x: (x["y"], x["m"], x["comp"], x.get("c","")))
    rows_yearly.sort(key=lambda x: (x["y"], x["comp"]))
    rows_total.sort(key=lambda x: x["y"])

    print(f"    → 월합계 {len(rows_export)}행 / 지역 {len(rows_region)}행 / 국가 {len(rows_country)}행")
    return {
        "DATA_KR_EXPORT":  rows_export,
        "DATA_KR_REGION":  rows_region,
        "DATA_KR_COUNTRY": rows_country,
        "DATA_KR_YEARLY":  rows_yearly,
        "DATA_KR_TOTAL":   rows_total,
    }

# ═══════════════════════════════════════════════════════════════
# 2. 중국산 수출실적
# ═══════════════════════════════════════════════════════════════
def process_cn_export(path: Path) -> dict:
    """
    예상 컬럼: 년도, 월, 국가, 지역, 수출량(MT)
    ※ 실제 컬럼명이 다르면 아래 COL_MAP을 수정하세요
    """
    COL_MAP = {
        "year":    ["년도", "연도", "year", "Year"],
        "month":   ["월", "month", "Month"],
        "country": ["국가", "수출국", "country", "Country"],
        "region":  ["지역", "대분류", "region", "Region", "_rg", "지역분류"],
        "volume":  ["수출량", "물량", "MT", "volume", "수량(MT)", "수출량(MT)"],
    }

    print(f"\n[2] 중국산 수출실적 처리 중: {path.name}")
    xl = pd.ExcelFile(path)
    print(f"    시트: {xl.sheet_names}")
    sheet = xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet, header=0)
    print(f"    컬럼: {list(df.columns)}")

    col = {}
    for key, candidates in COL_MAP.items():
        for c in candidates:
            matches = [x for x in df.columns if str(x).strip() == c]
            if matches: col[key] = matches[0]; break

    missing = [k for k in ["year","month","country","volume"] if k not in col]
    if missing:
        print(f"  ❌ 컬럼 매핑 실패: {missing}. COL_MAP을 수정하세요.")
        return {}

    has_region = "region" in col
    df = df.dropna(subset=[col["year"], col["month"], col["country"], col["volume"]])
    df[col["year"]]   = df[col["year"]].apply(safe_int)
    df[col["month"]]  = df[col["month"]].apply(safe_int)
    df[col["volume"]] = df[col["volume"]].apply(safe_float)
    df = df[(df[col["year"]] >= 2022) & (df[col["month"]].between(1,12)) & (df[col["volume"]] >= 0)]

    rows = []
    for _, r in df.iterrows():
        row = {"y":safe_int(r[col["year"]]), "m":safe_int(r[col["month"]]),
               "c":str(r[col["country"]]).strip(),
               "_rg": str(r[col["region"]]).strip() if has_region else "기타",
               "v":safe_int(r[col["volume"]])}
        rows.append(row)

    rows.sort(key=lambda x: (x["y"], x["m"], x["c"]))
    print(f"    → {len(rows)}행")
    return {"DATA_EXPORT": rows}

# ═══════════════════════════════════════════════════════════════
# 3. 중국 아트지 가격
# ═══════════════════════════════════════════════════════════════
def process_cn_price(path: Path) -> dict:
    """
    예상 컬럼: 년도, 월, 지역(광동/장절호/북경 등), 종이종류, 평량, 가격(위안/톤)
    ※ 실제 구조가 다를 경우 아래를 수정하세요
    """
    COL_MAP = {
        "year":   ["년도", "연도", "year", "Year"],
        "month":  ["월", "month"],
        "region": ["지역", "region", "Region", "지역명"],
        "type":   ["종류", "종이", "type", "품목", "코트지/옵셋지"],
        "grade":  ["평량", "grade", "규격", "gsm"],
        "price":  ["가격", "price", "단가", "위안", "CNY", "Price"],
    }

    print(f"\n[3] 중국 아트지 가격 처리 중: {path.name}")
    xl = pd.ExcelFile(path)
    print(f"    시트: {xl.sheet_names}")
    sheet = xl.sheet_names[0]
    df = pd.read_excel(path, sheet_name=sheet, header=0)
    print(f"    컬럼: {list(df.columns)}")

    col = {}
    for key, candidates in COL_MAP.items():
        for c in candidates:
            matches = [x for x in df.columns if str(x).strip() == c]
            if matches: col[key] = matches[0]; break

    missing = [k for k in ["year","month","price"] if k not in col]
    if missing:
        print(f"  ❌ 컬럼 매핑 실패: {missing}. COL_MAP을 수정하세요.")
        return {}

    df = df.dropna(subset=[col["year"], col["month"], col["price"]])
    df[col["year"]]  = df[col["year"]].apply(safe_int)
    df[col["month"]] = df[col["month"]].apply(safe_int)
    df[col["price"]] = df[col["price"]].apply(safe_float)
    df = df[(df[col["year"]] >= 2024) & (df[col["month"]].between(1,12)) & (df[col["price"]] > 0)]

    rows = []
    for _, r in df.iterrows():
        row = {
            "t":  str(r[col["type"]]).strip() if "type" in col else "코트지",
            "g":  str(r[col["grade"]]).strip() if "grade" in col else "157gsm",
            "r":  str(r[col["region"]]).strip() if "region" in col else "",
            "y":  safe_int(r[col["year"]]),
            "m":  safe_int(r[col["month"]]),
            "p":  safe_float(r[col["price"]]),
        }
        rows.append(row)

    rows.sort(key=lambda x: (x["t"], x["g"], x["r"], x["y"], x["m"]))
    print(f"    → {len(rows)}행")

    # 기존 usd 환율 데이터는 별도 관리 (가격 파일에 없으면 유지)
    return {"DATA_CN_PRICES": rows}

# ═══════════════════════════════════════════════════════════════
# data.js 업데이트
# ═══════════════════════════════════════════════════════════════
def update_datajs(updates: dict):
    """여러 const 블록을 data.js에서 교체"""
    print(f"\n[4] data.js 업데이트 중...")
    content = DATA_JS.read_text(encoding="utf-8")

    for name, rows in updates.items():
        if not rows:
            print(f"  ⏭️  {name}: 데이터 없음, 건너뜀")
            continue
        block = js_array(name, rows)
        content = inject_js(content, name, block)
        print(f"  ✅ {name} 업데이트 완료 ({len(rows)}행)")

    # 업데이트 날짜 반영
    today = datetime.now().strftime("%Y.%m.%d")
    content = re.sub(r'최종 업데이트: [\d.]+', f'최종 업데이트: {today}', content)

    # 백업 후 저장
    backup = DATA_JS.with_suffix(f".backup_{datetime.now().strftime('%y%m%d_%H%M')}.js")
    DATA_JS.rename(backup)
    print(f"  💾 백업: {backup.name}")
    DATA_JS.write_text(content, encoding="utf-8")
    print(f"  ✅ data.js 저장 완료")

# ═══════════════════════════════════════════════════════════════
# 메인
# ═══════════════════════════════════════════════════════════════
def main():
    print("=" * 60)
    print("인쇄용지 대시보드 데이터 업데이트")
    print(f"실행 시각: {datetime.now().strftime('%Y-%m-%d %H:%M')}")
    print("=" * 60)

    all_updates = {}
    errors = []

    # 1. 한국산 수출실적
    kr_path = find_excel("한국수출실적")
    if kr_path:
        try:
            all_updates.update(process_kr_export(kr_path))
        except Exception as e:
            errors.append(f"한국수출실적: {e}")
            print(f"  ❌ 오류: {e}")

    # 2. 중국산 수출실적
    cn_path = find_excel("중국수출실적")
    if cn_path:
        try:
            all_updates.update(process_cn_export(cn_path))
        except Exception as e:
            errors.append(f"중국수출실적: {e}")
            print(f"  ❌ 오류: {e}")

    # 3. 중국 아트지 가격
    price_path = find_excel("중국 아트지 가격")
    if not price_path:
        price_path = find_excel("아트지 가격")
    if price_path:
        try:
            all_updates.update(process_cn_price(price_path))
        except Exception as e:
            errors.append(f"아트지 가격: {e}")
            print(f"  ❌ 오류: {e}")

    if all_updates:
        update_datajs(all_updates)

    print("\n" + "=" * 60)
    if errors:
        print(f"⚠️  일부 오류 발생:\n" + "\n".join(f"  - {e}" for e in errors))
    else:
        print("✅ 완료! index.html을 브라우저에서 새로고침하세요.")
    print("=" * 60)

if __name__ == "__main__":
    main()
