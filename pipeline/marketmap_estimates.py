"""Private evidence and market-map estimates. Source amounts stay separate."""
import csv
import hashlib
import json
import os
from collections import defaultdict

HASH = '0BF79DBF800EB6515A143117F3CEB631618A118663CE54F498DF454E55316073'.lower()
ALIASES = {'인도(인디아)': '인도', '터키': '튀르키예',
           '남아프리카': '남아공', '남아프리카 공화국': '남아공', '남아프리카공화국': '남아공',
           '뉴 칼레도니아': '뉴칼레도니아', '도미니카 공화국': '도미니카공화국',
           '러시아연방': '러시아', '싱가폴': '싱가포르', '아랍에미리트 연합': '아랍에미리트',
           '트리니다드 토바고': '트리니다드토바고', '파푸아 뉴기니': '파푸아뉴기니', '레유니온': '레위니옹'}

def canonical(name):
    return ALIASES.get(name, name)

def month(y, m):
    return y * 12 + m - 1

def load_basis(path):
    with open(path, 'rb') as f:
        digest = hashlib.sha256(f.read()).hexdigest()
    if digest != HASH:
        raise ValueError('한국제지 근거 CSV 해시 불일치')
    sums = defaultdict(lambda: [0.0, 0.0, False])
    annual = defaultdict(lambda: defaultdict(lambda: [0.0, 0.0, False]))
    with open(path, encoding='utf-8-sig', newline='') as f:
        for row in csv.DictReader(f):
            y = int(row['연도'].replace('년', '')) + 2000
            m = int(row['월'].replace('월', ''))
            if not month(2021, 1) <= month(y, m) <= month(2023, 7):
                continue
            company = row['회사']
            if company not in ('한솔', '무림', '한국', '무림 울산+한국'):
                continue
            v = float(row['수량(MT)'].replace(',', '') or 0)
            for cell in (sums[canonical(row['국가명'])], annual[canonical(row['국가명'])][str(y)]):
                cell[1] += v
                if company == '한국':
                    cell[0] += v
                    cell[2] = True
    return sums, annual

def export_basis(private_dir, basis, annual):
    """Write full-precision audit evidence only to the ignored private directory."""
    os.makedirs(private_dir, exist_ok=True)
    payload = {'period': '2021.01~2023.07', 'sha256': HASH,
               'source_status': '저장된 CSV 기준; 최신 원본 대사 미완료',
               'formula': '한국제지 수출량 / 한국산 3사 수출량',
               'combined_by_country': dict(basis),
               'annual_by_country': {country: dict(years) for country, years in annual.items()}}
    with open(os.path.join(private_dir, 'kr_split_basis.json'), 'w', encoding='utf-8') as f:
        json.dump(payload, f, ensure_ascii=False, indent=2)
    with open(os.path.join(private_dir, 'kr_split_basis.csv'), 'w', encoding='utf-8-sig', newline='') as f:
        writer = csv.writer(f)
        writer.writerow(['국가', '기간', '한국제지_톤', '한국산3사_톤', '한국제지행_존재', '비중'])
        for country, combined in sorted(basis.items()):
            for period, values in sorted(annual[country].items()):
                numerator, denominator, present = values
                writer.writerow([country, period if period != '2023' else '2023.01~07', numerator, denominator,
                                 present, numerator / denominator if present and denominator > 0 else ''])
            numerator, denominator, present = combined
            writer.writerow([country, '2021.01~2023.07', numerator, denominator, present,
                             numerator / denominator if present and denominator > 0 else ''])

def split_kr(vol, amt, ci, basis):
    records, original = [], []
    keys = sorted({(mo, country) for mo, country, si in vol if si in (1, 2, 3, 4) and mo >= month(2023, 8)})
    for mo, country in keys:
        get = lambda si: vol.get((mo, country, si), 0.0)
        pool, actual = get(4), get(3)
        total = sum(get(i) for i in (1, 2, 3, 4))
        b = basis.get(country)
        ratio = b[0] / b[1] if b and b[1] > 0 and b[2] else None
        target = total * ratio if ratio is not None else None
        add = target - actual if target is not None else None
        reason = None
        if not b or not b[2]: reason = '과거 한국제지 행 없음'
        elif b[1] <= 0: reason = '과거 한국산 분모 0 또는 음수'
        elif min(get(1), get(2), pool, actual, total, target) < 0: reason = '음수 입력 또는 추정치'
        elif add < 0: reason = '실제 한국제지 물량이 추정 목표 초과'
        elif add > pool: reason = '추가 추정량이 미분리 물량 초과'
        if reason is None:
            moorim = get(2)
            vol[(mo, country, 3)] = actual + add
            vol[(mo, country, 2)] = moorim + pool - add
            vol[(mo, country, 4)] = 0.0
            status = 'confirmed_zero' if ratio == 0 else 'estimated'
        else:
            status = 'missing' if ratio is None else 'conflict'
        affected = [2, 3, 4] if reason is None and pool else []
        original.append({'month': mo, 'country': ci[country], 'volume_t': total,
                         'amount_usd': sum(amt.get((mo, country, i), 0) for i in (1, 2, 3, 4)),
                         'series_amount_usd': [amt.get((mo, country, i), 0) for i in (1, 2, 3, 4)]})
        records.append({'month': mo, 'country': ci[country], 'status': status, 'reason': reason,
                        'original_pool_t': pool, 'kr_total_t': total, 'ratio': ratio,
                        'hankuk_t': vol.get((mo, country, 3), 0), 'moorim_t': vol.get((mo, country, 2), 0),
                        'affected_series': affected})
    return records, original

def build_local(capa, countries, cn, kr, mm_dir, actuals=None):
    caps = defaultdict(float)
    for row in capa:
        if str(row.get('product', '')).startswith('WFC'):
            caps[canonical(row.get('location_desc', ''))] += float(row.get('capa_mty') or 0)
    def available_capacity(country, mo):
        matching = [r for r in capa if canonical(r.get('location_desc', '')) == country
                    and str(r.get('product', '')).startswith('WFC')]
        if not matching: return None
        return sum(float(r.get('capa_mty') or 0) for r in matching
                   if (r.get('start') is None or r['start'] <= mo)
                   and (r.get('end') is None or mo <= r['end']))
    entries, exports = [], defaultdict(float)
    for country, data in (('중국', cn), ('한국', kr)):
        for r in data:
            exports[(country, month(r['y'], r['m']))] += r.get('v') or 0
    for country, data in (('중국', cn), ('한국', kr)):
        if country not in countries: continue
        for mo in sorted({month(r['y'], r['m']) for r in data}):
            cap, exp = available_capacity(country, mo), exports[(country, mo)]
            value = cap * .9 / 12 - exp if cap is not None else None
            reason = '보유 CAPA 기준 참고 계산. CAPA 시점·전체성 및 수출 범위 미확인; 재고·재수출 미반영.'
            if country == '한국': reason += ' 한국 수출은 3사만 포함하며 CAPA 범위와 다를 수 있음.'
            if value is not None and value < 0: reason += ' 계산 결과 음수: 자료 범위 확인 필요.'
            entries.append({'country': countries[country], 'start': mo, 'end': mo, 'value_t': value,
                'capacity_t': cap, 'exports_t': exp, 'status': 'conflict' if value is not None and value < 0 else 'reference' if value is not None else 'missing',
                'reason': reason, 'source': 'DATA_CAPA 제품별 capa_mty; DATA_EXPORT 또는 DATA_KR_COUNTRY', 'capacity_year': None})
    for code, country in [('JP', '일본'), ('IN', '인도')]:
        if country not in countries: continue
        for year in (2023, 2024, 2025):
            path = os.path.join(mm_dir, '데이터_공급국', f'comtrade_{code}_X_{year}.json')
            with open(path, encoding='utf8') as f: data = json.load(f)['data']
            found = [r for r in data if r.get('partnerCode') == 0 and str(r.get('cmdCode')) in ('481013', '481014', '481019')]
            exp = sum(r['netWgt'] for r in found) / 1000 if len(found) == 3 and all(r.get('netWgt') is not None for r in found) else None
            monthly_caps = [available_capacity(country, month(year, m)) for m in range(1, 13)]
            cap = caps.get(country)
            production = sum(c * .9 / 12 for c in monthly_caps) if all(c is not None for c in monthly_caps) else None
            value = production - exp if production is not None and exp is not None else None
            reason = '연간 참고값. CAPA 시점·전체성 미확인, 수출과 CAPA 출처가 다름; 월로 나누지 않음.'
            if code == 'IN': reason += ' 다른 최신 자료의 CAPA와 차이가 있어 원 DATA_CAPA를 유지.'
            if value is not None and value < 0: reason += ' 계산 결과 음수.'
            entries.append({'country': countries[country], 'start': month(year, 1), 'end': month(year, 12), 'value_t': value,
                'capacity_t': cap, 'exports_t': exp, 'status': 'conflict' if value is not None and value < 0 else 'reference' if value is not None else 'missing',
                'reason': reason, 'source': f'DATA_CAPA; UN Comtrade {code} {year} World', 'capacity_year': None})
    for name, idx in countries.items():
        if name not in ('중국', '한국', '일본', '인도'):
            entries.append({'country': idx, 'start': None, 'end': None, 'value_t': None, 'capacity_t': caps.get(name), 'exports_t': None,
                'status': 'missing', 'reason': '해당 국가 전체 CWF 수출 자료 없음. EU27 합계는 개별 국가에 사용하지 않음.',
                'source': 'DATA_CAPA', 'capacity_year': None})
    seen_actual = set()
    for actual in actuals or []:
        name = canonical(actual.get('country', ''))
        start, end = actual.get('start'), actual.get('end')
        if name not in countries or actual.get('product') != 'CWF':
            continue
        if not isinstance(start, int) or not isinstance(end, int) or start > end:
            raise ValueError('실제 내수 기간 오류')
        if actual.get('value_t') is None or actual['value_t'] < 0 or not actual.get('source'):
            raise ValueError('실제 내수 값 또는 출처 오류')
        key = (countries[name], start, end)
        if key in seen_actual: raise ValueError('실제 내수 중복 기간')
        seen_actual.add(key)
        replacement = {'country': key[0], 'start': start, 'end': end, 'value_t': actual['value_t'],
                       'capacity_t': None, 'exports_t': None, 'status': 'actual',
                       'reason': '같은 기간·품목의 실제 내수 출하', 'source': actual['source'], 'capacity_year': None}
        entries = [e for e in entries if (e['country'], e['start'], e['end']) != key]
        entries.append(replacement)
    return {'assumed_utilization': .9, 'entries': entries,
            'limitations': ['CAPA 전체성 및 유효기간 미확인', '재고 변동과 재수출 미반영', '참고값은 확정 수요 합계에 포함하지 않음']}
