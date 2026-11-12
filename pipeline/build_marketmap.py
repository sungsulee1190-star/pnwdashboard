"""마켓맵(2페이지 마켓맵 모드) 데이터 빌드: data.js에 `const DATA_MM = {...};` 한 블록을 넣거나 바꾼다.

사용: python pipeline\\build_marketmap.py   (저장소 루트에서, data.js가 있어야 함)
입력 폴더 = pipeline/paths.local.json 의 "marketmap_dir" (마켓맵 프로젝트 폴더, git에 올리지 않음)
  - data.js 의 DATA_EXPORT(중국 월보) · DATA_KR_COUNTRY(한국 3사)
  - scripts/hansol_region_map_261112.json (없으면 hansol_region_map.json) : 나라 → 한솔 24권역(251031 L열)
  - scripts/marketmap_country_base.json : 나라 → 영문 경계 이름(en)·점 좌표(pt)·확인 화면의 L열 권역(l)·중국 원본 권역(rg)
  지역 순서: 지역표 → base의 l → base의 rg를 24권역 이름으로 변환. 셋 다 없으면 "미배정"(validate_data.py가 톤을 센다)
  - 데이터_Comext/eu27_by_country_month.json · eu27_intra_by_country_month.json : EU27 역외·역내 월별 톤
  - 데이터_공급국/asia_export_by_country.json : 일본·인도 → 도착국 연간 톤(참고값)
  - 데이터_공급국/deck_input.json : 국가 상세 장표 값
data.js의 다른 블록은 건드리지 않는다. data.js는 평문이라 git에 올리지 않는다(암호화본 data.enc.js만 올림).
"""
import datetime
import json
import os
import re
import sys

sys.stdout.reconfigure(encoding="utf-8")
ROOT = os.path.dirname(os.path.dirname(os.path.abspath(__file__)))
DATA_JS = os.path.join(ROOT, "data.js")
PATHS = os.path.join(ROOT, "pipeline", "paths.local.json")

SERIES = ["중국산", "한솔", "무림", "한국제지", "무림·한국 미분리", "유럽산 역외", "유럽산 EU역내"]
COMP = {"한솔": 1, "무림": 2, "한국": 3, "무림 울산+한국": 4}
# 한솔 24권역(목표 엑셀 행 순서)
REGIONS = ["01_미국", "02_오세아", "03_서유럽_영국", "04_서유럽_기타", "05_동유럽", "06_아프리카_남아공", "07_아프리카_기타",
           "08_남미_멕시코", "09_남미_기타", "10_중국", "11_몽골", "12_대만", "13_동남아_말싱", "14_동남아_인니", "15_동남아_베트남",
           "16_동남아_태국", "17_동남아_필리핀", "18_동남아_기타", "19_인도", "20_서남아_기타", "21_방글라", "22_중동_기타",
           "23_중동_이란", "24_일본"]
L_FIX = {"23_일본": "24_일본", "21_중동_기타": "22_중동_기타", "17_동남아_기타": "18_동남아_기타"}   # L열 옛 번호 → 24권역
CN2H = {"01_북미": "01_미국", "04_서유럽": "04_서유럽_기타", "05_유럽_기타": "05_동유럽", "16_동남아-태국": "16_동남아_태국",
        "21_서남아_2": "20_서남아_기타", "22_중동(기타)": "22_중동_기타", "23_중동(이란)": "23_중동_이란", "25_대한민국": "기타"}


def fail(msg):
    print(f"[오류] {msg}")
    sys.exit(1)


def load_json(path):
    with open(path, encoding="utf-8") as f:
        return json.load(f)


def js_rows(text, name):
    """data.js의 `const NAME = [ {y:..}, ... ];` 배열을 한 줄씩 읽는다(run_updates.py의 js_array 형식)."""
    s = text.find(f"const {name} = [")
    if s < 0:
        fail(f"data.js에 {name} 없음")
    e = text.index("];", s)
    out = []
    for ln in text[s:e].splitlines():
        ln = ln.strip().rstrip(",")
        if ln.startswith("{"):
            out.append(json.loads(re.sub(r"([{,]\s*)(\w+):", r'\1"\2":', ln)))
    return out


def month(y, m):
    return y * 12 + m - 1


def main():
    if not os.path.exists(DATA_JS):
        fail("data.js 없음. 먼저 pipeline\\update.bat(또는 run_updates.py)로 data.js를 만든다.")
    if not os.path.exists(PATHS):
        fail("pipeline/paths.local.json 없음. paths.example.json을 복사해 marketmap_dir를 적는다.")
    mm_dir = load_json(PATHS).get("marketmap_dir")
    if not mm_dir or not os.path.isdir(mm_dir):
        fail(f"paths.local.json의 marketmap_dir 폴더가 없다: {mm_dir}")
    p = lambda *a: os.path.join(mm_dir, *a)

    text = open(DATA_JS, encoding="utf-8").read()
    cn, kr = js_rows(text, "DATA_EXPORT"), js_rows(text, "DATA_KR_COUNTRY")

    rmap_name = "hansol_region_map_261112.json"
    if not os.path.exists(p("scripts", rmap_name)):
        rmap_name = "hansol_region_map.json"
    # 24권역이 아닌 값(예: "국가 추가 필요")은 버린다
    rmap = {k: v for k, v in load_json(p("scripts", rmap_name))["map"].items() if L_FIX.get(v, v) in REGIONS}
    base = load_json(p("scripts", "marketmap_country_base.json"))["countries"]
    eux = load_json(p("데이터_Comext", "eu27_by_country_month.json"))
    eui = load_json(p("데이터_Comext", "eu27_intra_by_country_month.json"))
    asia = load_json(p("데이터_공급국", "asia_export_by_country.json"))
    deck = {k: v for k, v in load_json(p("데이터_공급국", "deck_input.json")).items() if not k.startswith("_")}

    # 월 × 나라 × 계열 → 톤·금액
    vol, amt = {}, {}
    for r in cn:
        if r.get("v"):
            k = (month(r["y"], r["m"]), r["c"], 0)
            vol[k] = vol.get(k, 0) + r["v"]
    for r in kr:
        if r["comp"] not in COMP:
            fail(f"DATA_KR_COUNTRY에 모르는 회사: {r['comp']}")
        k = (month(r["y"], r["m"]), r["c"], COMP[r["comp"]])
        vol[k] = vol.get(k, 0) + (r.get("v") or 0)
        amt[k] = amt.get(k, 0) + (r.get("amt") or 0)

    names = sorted({k[1] for k in vol})
    ci = {c: i for i, c in enumerate(names)}
    countries = []
    for c in names:
        b = base.get(c, {})
        h = rmap.get(c) or b.get("l")
        if h:
            rg, rs = L_FIX.get(h, h), ("L번호정리" if h in L_FIX else "L열")
        elif b.get("rg") and b["rg"] != "미배정":
            rg, rs = CN2H.get(b["rg"], b["rg"]), "중국코드변환"
        else:
            rg, rs = "미배정", "미배정"
        countries.append({"ko": c, "en": b.get("en"), "pt": b.get("pt"), "rg": rg, "rs": rs})

    # EU 역외(5)·역내(6): 대시보드 나라 목록에 있는 도착국만 행으로. 없는 나라 톤은 unmapped로 기록
    eu_unmapped = {}
    for src, si in ((eux, 5), (eui, 6)):
        for mk, o in src.items():
            for ko, t in o.items():
                if ko in ci:
                    vol[(int(mk), ko, si)] = vol.get((int(mk), ko, si), 0) + t
                else:
                    eu_unmapped[ko] = eu_unmapped.get(ko, 0) + t

    rows = [[k[0], ci[k[1]], k[2], round(v, 1), round(amt.get(k, 0))] for k, v in sorted(vol.items()) if v or amt.get(k)]
    cn_m = [month(r["y"], r["m"]) for r in cn]
    kr_m = [month(r["y"], r["m"]) for r in kr]
    eu_m = [int(k) for k in eux]
    rng = {"cn": [min(cn_m), max(cn_m)], "kr": [min(kr_m), max(kr_m)], "eu": [min(eu_m), max(eu_m)]}
    latest = min(rng["cn"][1], rng["kr"][1], rng["eu"][1])   # 세 자료 공통 최신월(C31)

    asia_unmapped = {code: {y: round(sum(t for ko, t in o.items() if ko not in ci), 1) for y, o in ys.items()} for code, ys in asia.items()}
    out = {
        "version": 1,
        "built": datetime.date.today().isoformat(),
        "latest": latest,
        "range": rng,
        "series": SERIES,
        "regions": REGIONS,
        "countries": countries,
        "rows": rows,          # [월 번호(연*12+월-1), 나라 index, 계열 index, 톤, 금액 USD]
        "asia": asia,          # {JP|IN: {연도: {나라: 톤}}} 연간 참고값, 합계에 넣지 않음(C1)
        "deck": deck,
        "unmapped": {"eu_t": round(sum(eu_unmapped.values()), 1), "eu": {k: round(v, 1) for k, v in sorted(eu_unmapped.items(), key=lambda x: -x[1])},
                     "asia_t": asia_unmapped},
        "check": {"cn_t": round(sum(r.get("v") or 0 for r in cn), 1), "kr_t": round(sum(r.get("v") or 0 for r in kr), 1),
                  "eu_t": round(sum(t for src in (eux, eui) for o in src.values() for t in o.values()), 1)},
        "source": "중국 월보(DATA_EXPORT) + 한국 3사(DATA_KR_COUNTRY) + Eurostat Comext EU27(역외·역내) · 일본·인도 = UN Comtrade 연간",
    }

    block = "const DATA_MM = " + json.dumps(out, ensure_ascii=False, separators=(",", ":")) + ";"
    pat = re.compile(r"^const DATA_MM = .*?;$", re.M)
    if pat.search(text):
        text = pat.sub(lambda _: block, text, count=1)
        print("  - DATA_MM 블록 교체")
    else:
        text = text.rstrip("\n") + "\n\n" + block + "\n"
        print("  - DATA_MM 블록 추가")
    with open(DATA_JS, "w", encoding="utf-8") as f:
        f.write(text)

    y, m = divmod(latest, 12)
    un = [f"{c['ko']}({c['rg']})" for c in countries if c["rg"] not in REGIONS]
    print(f"기준월 {y}.{m + 1} · 나라 {len(countries)} · 행 {len(rows)} · 지역표 {rmap_name} · {len(block) / 1e6:.2f}MB")
    print(f"24권역 밖 나라: {un}")
    print(f"EU 도착국 중 나라 목록에 없음: {out['unmapped']['eu_t']}톤 {list(out['unmapped']['eu'])[:8]}")


if __name__ == "__main__":
    main()
