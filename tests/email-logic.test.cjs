const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root = path.join(__dirname, '..');
function load(data = {}) {
  const context = vm.createContext({ window: {}, ...data });
  vm.runInContext(fs.readFileSync(path.join(root, 'dashboard_logic.js'), 'utf8'), context);
  return context.window.DashboardLogic;
}
const options = { year: 2026, monthFrom: 1, monthTo: 1, targetRegion: 'all', includeCN: true };

test('empty latest period has no invented date and skips invalid records', () => {
  const logic = load();
  assert.equal(logic.getLatestPeriod([]), null);
  assert.deepEqual(JSON.parse(JSON.stringify(logic.getLatestPeriod([{ y: 2026, m: 13 }, { y: 2025, m: 12 }]))), { year: 2025, month: 12 });
});
test('period selection rejects reversed months and invalid year/month values', () => {
  const logic = load();
  for (const period of [[2026, 3, 1], [2026, 0, 2], [NaN, 1, 2], [2026, 1, 13]]) {
    assert.throws(() => logic.filterDataByPeriod([], ...period), /기간|period/i);
    assert.match(logic.buildEmailDraft({ ...options, year: period[0], monthFrom: period[1], monthTo: period[2] }).text, /기간/);
  }
});
test('zero baseline is uncomputable growth and missing observations stay missing', () => {
  const logic = load();
  assert.equal(logic.calculateYoY([{ v: 12 }], [{ v: 0 }]).rate, null);
  assert.equal(logic.calculateYoY([{ v: 12 }], []).previous, null);
  assert.equal(logic.calculateYoY([], []).current, null);
  assert.equal(logic.calculateMoM([{ v: 0 }], [{ v: 2 }]).rate, -100);
});
test('regions match dashboard grouping including KR codes and canonical names', () => {
  const logic = load();
  const regions = logic.aggregateByRegion([{ _rg: '03_서유럽_영국', v: 3 }, { rg: '04_동유럽', v: 4 }, { rg: '09_동남아', v: 2 }]);
  assert.equal(regions.find(r => r.region === '유럽').volume, 7);
  assert.equal(logic.aggregateByCountry([{ c: '태국', rg: '09_동남아', v: 2 }], '동남아')[0].volume, 2);
  assert.equal(logic.cleanRegion('동남아'), '동남아');
  assert.equal(logic.cleanRegion('#N/A'), '기타');
});
test('contributions preserve region union without inventing zero observations', () => {
  const logic = load();
  const values = logic.calculateContribution([{ region: '동남아', volume: 10, count: 1 }], [{ region: '유럽', volume: 10, count: 1 }], 0);
  assert.equal(values.length, 2);
  assert.equal(values.find(r => r.region === '유럽').volume, null);
  assert.equal(values.find(r => r.region === '동남아').prevVolume, null);
  assert.ok(values.every(r => r.diff === null && r.contribution === null));
  const recordedZero = logic.calculateContribution([{ region: '유럽', volume: 0 }], [{ region: '유럽', volume: 10 }], -10)[0];
  assert.equal(recordedZero.diff, -10);
  assert.equal(recordedZero.contribution, -100);
  const unchanged = logic.calculateContribution([{ region: '유럽', volume: 10 }], [{ region: '유럽', volume: 10 }], 0)[0];
  assert.equal(unchanged.diff, 0);
  assert.equal(unchanged.contribution, 0);
});
test('missing regional observations never become zero exports or growth narratives', () => {
  const logic = load({ DATA_EXPORT: [{ y: 2026, m: 1, _rg: '유럽', v: 10 }, { y: 2025, m: 1, _rg: '동남아', v: 10 }] });
  const draft = logic.buildEmailDraft(options);
  assert.match(draft.text, /동남아\t데이터 없음\t비교 데이터 없음\t산출 불가/);
  assert.match(draft.text, /유럽\t10\t비교 데이터 없음\t산출 불가/);
  assert.doesNotMatch(draft.text, /-100\.0%|기준 0|지역의 감소량|지역의 증가량/);
});
test('regional comparison requires each selected month while recorded zero remains comparable', () => {
  const data = [
    { y: 2026, m: 1, _rg: '유럽', v: 100 },
    { y: 2026, m: 2, _rg: '동남아', v: 20 },
    { y: 2025, m: 1, _rg: '유럽', v: 50 },
    { y: 2025, m: 2, _rg: '유럽', v: 50 },
    { y: 2025, m: 1, _rg: '동남아', v: 10 },
    { y: 2025, m: 2, _rg: '동남아', v: 10 }
  ];
  const missingDraft = load({ DATA_EXPORT: data }).buildEmailDraft({ ...options, monthTo: 2 });
  assert.match(missingDraft.text, /유럽\t100\t비교 데이터 없음\t산출 불가/);
  assert.match(missingDraft.text, /유럽.*2026년 2월.*없/);
  assert.doesNotMatch(missingDraft.text, /지역의 감소량|지역의 증가량/);
  const zeroDraft = load({ DATA_EXPORT: [...data, { y: 2026, m: 2, _rg: '유럽', v: 0 }, { y: 2026, m: 1, _rg: '동남아', v: 0 }] }).buildEmailDraft({ ...options, monthTo: 2 });
  assert.match(zeroDraft.text, /유럽\t100\t0\.0%\t0/);
  assert.match(zeroDraft.text, /동남아\t20\t0\.0%\t0/);
  const declineDraft = load({ DATA_EXPORT: [{ y: 2026, m: 1, _rg: '유럽', v: 0 }, { y: 2025, m: 1, _rg: '유럽', v: 10 }] }).buildEmailDraft(options);
  assert.match(declineDraft.text, /유럽\t0\t-100\.0%\t-10/);
});
test('narrative ranks changes itself without relying on input order', () => {
  const logic = load();
  const narrative = logic.buildNarrative([{ region: '작은 증가', diff: 1 }, { region: '최대 증가', diff: 20 }, { region: '작은 감소', diff: -1 }, { region: '최대 감소', diff: -20 }]);
  assert.ok(narrative.indexOf('최대 증가') < narrative.indexOf('작은 증가'));
  assert.ok(narrative.indexOf('최대 감소') < narrative.indexOf('작은 감소'));
});
test('email target region scopes current and previous exports and reports top countries', () => {
  const logic = load({ DATA_EXPORT: [{ y: 2026, m: 1, _rg: '04_서유럽', c: '독일', v: 100 }, { y: 2025, m: 1, _rg: '04_서유럽', c: '독일', v: 50 }, { y: 2026, m: 1, _rg: '15_동남아_베트남', c: '베트남', v: 9 }] });
  const draft = logic.buildEmailDraft({ ...options, targetRegion: '동남아' });
  assert.match(draft.text, /9 MT/);
  assert.doesNotMatch(draft.text, /100 MT|독일|0\.0%/);
  assert.match(draft.text, /베트남/);
  assert.match(draft.text, /전년.*없/);
});
test('unchanged export totals render without crashing or claiming decline', () => {
  const logic = load({ DATA_EXPORT: [{ y: 2026, m: 1, _rg: '유럽', v: 10 }, { y: 2025, m: 1, _rg: '유럽', v: 10 }] });
  assert.match(logic.buildEmailDraft(options).text, /변동 없음/);
});
test('absent and incomplete requested months are explicitly missing, not a zero decline', () => {
  const logic = load({ DATA_EXPORT: [{ y: 2026, m: 1, v: 10 }, { y: 2025, m: 1, v: 20 }] });
  const partial = logic.buildEmailDraft({ ...options, monthTo: 2 });
  assert.match(partial.text, /2월.*없/);
  assert.doesNotMatch(partial.text, /-50\.0%/);
  const absent = logic.buildEmailDraft({ ...options, monthFrom: 3, monthTo: 3 });
  assert.match(absent.text, /없/);
  assert.doesNotMatch(absent.text, /0 MT.*감소/);
});
test('news uses real object schema, dates its source and escapes all supplied markup', () => {
  const logic = load({ DATA_NEWS: { last_updated: '2026-01-03', items: [{ date: '2026-01-03', title: '<img src=x onerror=alert(1)>', source: '<svg>', summary: '수요 & 공급\n<script>alert(1)</script>' }] }, DATA_EXPORT: [{ y: 2026, m: 1, _rg: '유럽', c: '<img src=x>', v: 2 }] });
  const draft = logic.buildEmailDraft({ ...options, includeNews: true });
  assert.doesNotMatch(draft.html, /<script>|<img|<svg>/);
  assert.match(draft.html, /&lt;script&gt;/);
  assert.match(draft.text, /수요 & 공급\n<script>/);
  assert.match(draft.text, /2026-01-03/);
  assert.match(draft.text, /\n.*중국산/);
});
test('KR email uses selected period/region quantity, amount and weighted FOB', () => {
  const logic = load({ DATA_KR_COUNTRY: [{ y: 2026, m: 1, c: '태국', comp: '무림', rg: '09_동남아', v: 10, amt: 7000, fob: 1 }, { y: 2026, m: 1, c: '베트남', comp: '한솔', rg: '09_동남아', v: 30, amt: 27000, fob: 1 }, { y: 2025, m: 1, c: '태국', comp: '무림', rg: '09_동남아', v: 20, amt: 18000 }, { y: 2026, m: 1, c: '독일', comp: '무림', rg: '03_서유럽', v: 999, amt: 999000 }] });
  const draft = logic.buildEmailDraft({ ...options, includeCN: false, includeKR: true, targetRegion: '동남아' });
  assert.match(draft.text, /40 MT/);
  assert.match(draft.text, /34,000/);
  assert.match(draft.text, /850.*USD\/MT/);
  assert.match(draft.text, /100\.0%/);
  assert.match(draft.text, /태국|베트남/);
  assert.doesNotMatch(draft.text, /999|참조 바랍니다/);
});
test('KR absent amount is explicitly missing and never silently converted into zero FOB', () => {
  const logic = load({ DATA_KR_COUNTRY: [{ y: 2026, m: 1, c: '<script>', comp: '<img>', rg: '유럽', v: 10 }] });
  const draft = logic.buildEmailDraft({ ...options, includeCN: false, includeKR: true });
  assert.match(draft.text, /금액.*없|FOB.*없/);
  assert.doesNotMatch(draft.html, /<script>|<img>/);
});
test('price email uses exact previous calendar month and selected prices, never generic claims', () => {
  const logic = load({ DATA_PRICE: { prices: [{ y: 2025, m: 12, t: '코트지', g: '157gsm', r: '광동', p: 5000 }, { y: 2026, m: 1, t: '코트지', g: '157gsm', r: '광동', p: 5500 }, { y: 2026, m: 2, t: '코트지', g: '157gsm', r: '광동', p: 9999 }, { y: 2025, m: 11, t: '<img>', r: '<script>', p: 50 }, { y: 2026, m: 1, t: '<img>', r: '<script>', p: 60 }] } });
  const draft = logic.buildEmailDraft({ ...options, includeCN: false, includePrice: true });
  assert.match(draft.text, /5,500.*CNY\/MT/);
  assert.match(draft.text, /10\.0%/);
  assert.match(draft.text, /전월.*없/);
  assert.doesNotMatch(draft.text, /9,999|변동성이/);
  assert.doesNotMatch(draft.html, /<script>|<img>/);
});
test('zero price remains zero with unavailable percentage and missing selection is explicit', () => {
  const logic = load({ DATA_PRICE: { prices: [{ y: 2025, m: 12, t: '코트지', r: '광동', p: 0 }, { y: 2026, m: 1, t: '코트지', r: '광동', p: 0 }] } });
  assert.match(logic.buildEmailDraft({ ...options, includePrice: true }).text, /0 CNY\/MT/);
  assert.match(logic.buildEmailDraft({ ...options, includePrice: true }).text, /기준.*0/);
  assert.match(logic.buildEmailDraft({ ...options, includePrice: true, year: 2030 }).text, /가격.*없/);
});

test('real repaired dataset renders bounded drafts with exact recorded volumes and live news schema', () => {
  const context = vm.createContext({ window: {} });
  vm.runInContext(fs.readFileSync(path.join(root, 'data.js'), 'utf8'), context);
  vm.runInContext(fs.readFileSync(path.join(root, 'dashboard_logic.js'), 'utf8'), context);
  const dataset = vm.runInContext('({cn: DATA_EXPORT, kr: DATA_KR_COUNTRY, prices: DATA_PRICE.prices, news: DATA_NEWS})', context);
  const logic = context.window.DashboardLogic;
  assert.ok(Array.isArray(dataset.news.items) && dataset.news.items.length > 0);
  for (const [kind, data] of [['CN', dataset.cn], ['KR', dataset.kr]]) {
    const latest = logic.getLatestPeriod(data);
    assert.ok(latest);
    const rows = data.filter(d => d.y === latest.year && d.m === latest.month);
    const expected = rows.reduce((sum, row) => sum + row.v, 0).toLocaleString('en-US', { maximumFractionDigits: 2 });
    for (const region of ['all', '유럽', '동남아']) {
      const draft = logic.buildEmailDraft({ ...options, year: latest.year, monthFrom: latest.month, monthTo: latest.month, targetRegion: region, includeCN: kind === 'CN', includeKR: kind === 'KR', includeNews: true, includePrice: true });
      if (region === 'all') assert.ok(draft.text.includes(expected + ' MT'));
      assert.ok(draft.text.includes(dataset.news.items[0].summary));
      assert.ok(draft.html.length < 100000);
      assert.doesNotMatch(draft.text, /undefined|NaN|Infinity/);
    }
  }
});
