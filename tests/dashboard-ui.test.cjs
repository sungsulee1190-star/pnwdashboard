const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const vm = require('node:vm');
const path = require('node:path');
const html = fs.readFileSync(path.join(__dirname, '../index.html'), 'utf8');

function harness(data = {}) {
  const elements = new Map();
  const nodes = [];
  function element(id = '') {
    const classes = new Set();
    const e = { id, value: '', innerHTML: '', textContent: '', style: {}, dataset: {}, children: [],
      classList: { add: (...x) => x.forEach(v => classes.add(v)), remove: (...x) => x.forEach(v => classes.delete(v)), contains: x => classes.has(x), toggle: (x, force) => { const yes = force ?? !classes.has(x); yes ? classes.add(x) : classes.delete(x); return yes; } },
      appendChild(x) { this.children.push(x); }, getAttribute(n) { return this[n] ?? null; }, parentElement: { appendChild() {} } };
    return e;
  }
  for (const match of html.matchAll(/<(\w+)\b([^>]+)>/g)) {
    const attrs = match[2];
    const id = /\bid="([^"]+)"/.exec(attrs)?.[1];
    const e = element(id);
    for (const c of /\bclass="([^"]+)"/.exec(attrs)?.[1]?.split(/\s+/) || []) e.classList.add(c);
    for (const d of attrs.matchAll(/data-([\w-]+)="([^"]+)"/g)) e.dataset[d[1]] = d[2];
    e.onclick = /onclick="([^"]+)"/.exec(attrs)?.[1] || '';
    if(id && !elements.has(id)) elements.set(id, e);
    nodes.push(e);
  }
  const select = s => nodes.filter(e => {
    if(s.startsWith('.')) return e.classList.contains(s.slice(1));
    const m = /^\[data-([\w-]+)="([^"]+)"\](\.on)?$/.exec(s);
    return m && e.dataset[m[1]] === m[2] && (!m[3] || e.classList.contains('on'));
  });
  const charts = new Map();
  let clipboard;
  const context = vm.createContext({ console, URL, Blob, ClipboardItem: class { constructor(x) { this.items=x; } }, setTimeout: f => f(),
    navigator: { clipboard: { write: x => { clipboard=x; return Promise.resolve(); }, writeText: () => Promise.resolve() } },
    document: { getElementById: id => elements.get(id) || null, querySelectorAll: select, querySelector: s => select(s)[0], addEventListener() {}, createElement: () => element() },
    Chart: class { constructor(e, config) { this.id=e.id;this.config=config; charts.set(e.id, config); } destroy() { if(charts.get(this.id)===this.config)charts.delete(this.id); } },
    DATA_EXPORT: [], DATA_KR_EXPORT: [], DATA_KR_COUNTRY: [], DATA_KR_REGION: [], DATA_KR_TOTAL: [], DATA_KR_YEARLY: [], DATA_CAPA: [], DATA_COMPANIES: [], DATA_NEWS: {items:[]}, DATA_PRICE: {prices:[],usd:[]}, DATA_INDICATORS: [], KR_REGIONS: [], KR_COMPETITORS: [], ...data });
  context.window = context;
  context.addEventListener = () => {};
  for(const script of html.matchAll(/<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g)) if(script[1].trim()) vm.runInContext(script[1],context);
  const defaults = {cnYear:'2026',cnMonthFrom:'1',cnMonthTo:'4',cnRegion:'all',krYear:'all',krMonthFrom:'1',krMonthTo:'12',krRegionFilter:'all',krPeriod:'monthly',krCountrySort:'v_desc',newsCoSel:'all'};
  for(const [id,value] of Object.entries(defaults)) elements.get(id).value=value;
  return {context, elements, charts, run: s => vm.runInContext(s,context), clipboard: () => clipboard};
}

test('every HTML id is unique so country autocomplete and table search are independent', () => {
  const ids = [...html.matchAll(/\bid="([^"]+)"/g)].map(m=>m[1]);
  assert.equal(ids.length, new Set(ids).size);
});

test('legacy company sources disable HTTP or executable URLs and escape source notes', () => {
 const h=harness({DATA_COMPANIES:[{id:'x',name:'회사',country:'한국',capa_kt:1,mills:[],products:[],links:{homepage:'http://example.com/',ir:'javascript:alert(1)',quarterly:'https://example.com/results',note:'<img src=x onerror=x>'},financials:[]}]});
 h.run("selectCo('x')");
 const links=h.elements.get('coLinks').innerHTML;
 assert.doesNotMatch(links,/href="http:|href="javascript:|<img/);
 assert.match(links,/href="https:\/\/example.com\/results"/);assert.match(links,/링크 확인 필요/);assert.match(links,/&lt;img/);
});

test('programmatic navigation opens the email tab and generates its draft', () => {
  const h=harness(); let generated=0; h.context.generateEmailDraft=()=>generated++;
  assert.doesNotThrow(()=>h.run("go('email')"));
  assert.equal(generated,1);
  assert.ok(h.elements.get('sec-email').classList.contains('active'));
});

test('navigation initializes map once and invalidates its size on each visit', () => {
  const h=harness();let initialized=0,resized=0;
  h.context.initMarket=()=>{initialized++;h.run('_M.mill = __map');};h.context.__map={invalidateSize:()=>resized++};
  h.run("go('map'); go('email'); go('map')");
  assert.equal(initialized,1);assert.equal(resized,2);
});

const krRows=[
  {y:2026,m:1,q:'Q1',h:'H1',comp:'한솔',c:'베트남',rg:'09_동남아',v:10,amt:100},
  {y:2026,m:2,q:'Q1',h:'H1',comp:'한솔',c:'베트남',rg:'09_동남아',v:20,amt:200},
  {y:2026,m:2,q:'Q1',h:'H1',comp:'한솔',c:'미국',rg:'01_미국',v:30,amt:300},
  {y:2026,m:2,q:'Q1',h:'H1',comp:'무림',c:'베트남',rg:'09_동남아',v:40,amt:400}
];
test('KR month, company, normalized region and country filters affect overview and country data', () => {
  const h=harness({DATA_KR_COUNTRY:krRows,DATA_KR_EXPORT:krRows.map(({c,rg,...x})=>x)});
  h.elements.get('krMonthFrom').value='2';h.elements.get('krMonthTo').value='2';h.elements.get('krRegionFilter').value='동남아';
  h.run("KR.year='2026'; KR.comp='한솔'; _krSelectedCountries=['베트남']");
  assert.equal(h.run('krGetF().reduce((s,x)=>s+x.v,0)'),20);
  assert.equal(h.run('krCountryGetData().reduce((s,x)=>s+x.v,0)'),20);
});
test('KR filter refresh redraws the currently visible country view', () => {
  const h=harness({DATA_KR_COUNTRY:krRows,DATA_KR_EXPORT:krRows});
  h.run("KR.view='country'; KR.comp='한솔'; krRefresh()");
  assert.ok(h.charts.get('krCCBar'));
  assert.equal(h.charts.get('krCCBar').data.datasets.length,1);
});

test('CN comparison retains normalized region and selected country filters', () => {
  const rows=[{y:2026,m:4,c:'독일',_rg:'04_서유럽',v:10},{y:2026,m:3,c:'독일',_rg:'04_서유럽',v:5},{y:2026,m:4,c:'프랑스',_rg:'04_서유럽',v:90}];
  const h=harness({DATA_EXPORT:rows});h.elements.get('cnRegion').value='유럽';h.run("_cnSelectedCountries=['독일']; setCnMode('2m')");
  assert.equal(h.charts.get('cnCVol').data.datasets[1].data[0],10);
  assert.equal(h.charts.get('cnCVol').data.labels.length,1);
  assert.notEqual(h.elements.get('cnParams').style.display,'none');
});
test('CN year comparisons match only months present in both years', () => {
  const rows=[{y:2026,m:1,c:'독일',_rg:'04_서유럽',v:10},{y:2026,m:2,c:'독일',_rg:'04_서유럽',v:20},{y:2025,m:1,c:'독일',_rg:'04_서유럽',v:5},{y:2025,m:2,c:'독일',_rg:'04_서유럽',v:10},{y:2025,m:12,c:'독일',_rg:'04_서유럽',v:1000}];
  const h=harness({DATA_EXPORT:rows});h.run("setCnMode('2y')");
  assert.equal(h.charts.get('cnCVol').data.datasets[0].data[0],15);
  assert.match(h.elements.get('cnVolTitle').textContent,/1[,~·\s-]*2월/);
});
test('CN quarter comparisons match partial-quarter month coverage', () => {
  const rows=[{y:2026,m:4,c:'독일',_rg:'04_서유럽',v:10},{y:2025,m:4,c:'독일',_rg:'04_서유럽',v:5},{y:2025,m:5,c:'독일',_rg:'04_서유럽',v:500}];
  const h=harness({DATA_EXPORT:rows});h.run("setCnMode('yoy_q')");
  assert.equal(h.charts.get('cnCVol').data.datasets[0].data[0],5);
  assert.match(h.elements.get('cnVolTitle').textContent,/4월/);
});
test('CN MoM compares latest selected month with previous calendar month across January rollover', () => {
  const rows=[{y:2026,m:1,c:'독일',_rg:'04_서유럽',v:30},{y:2025,m:12,c:'독일',_rg:'04_서유럽',v:20},{y:2025,m:12,c:'미국',_rg:'01_북미',v:1000}];
  const h=harness({DATA_EXPORT:rows});h.elements.get('cnMonthTo').value='1';h.elements.get('cnRegion').value='유럽';h.run("_cnSelectedCountries=['독일'];cnRefresh()");
  assert.match(h.elements.get('cnKpi').innerHTML,/50\.0%/);
  assert.match(h.elements.get('cnKpi').innerHTML,/2025년 12월/);
});
test('CN maximum export region sums all its countries', () => {
  const rows=[{y:2026,m:1,c:'미국',_rg:'01_북미',v:100},{y:2026,m:1,c:'독일',_rg:'04_서유럽',v:60},{y:2026,m:1,c:'프랑스',_rg:'04_서유럽',v:60}];
  const h=harness({DATA_EXPORT:rows});h.run('cnRefresh()');
  assert.match(h.elements.get('cnKpi').innerHTML,/최대 수출 지역<\/div>\s*<div class="kpi-v [^"]*">유럽/);
});
test('CN regional share uses all countries even when table stops at fifty', () => {
  const rows=Array.from({length:60},(_,i)=>({y:2026,m:1,c:'Country'+i,_rg:'04_서유럽',v:100}));
  const h=harness({DATA_EXPORT:rows});h.run('cnRefresh()');
  assert.match(h.elements.get('cnTbl').innerHTML,/1\.7%/);
});
test('mill modal shows unavailable machine dimensions and exact grade shares', () => {
  const h=harness();h.context.mill={mill:'A',company:'C',loc:'X',region:'China',capa:100,pulp:0,prods:[{p:'Minor',c:1},{p:'Major',c:99}]};h.run('openMillModal(mill)');
  const content=h.elements.get('millModalContent').innerHTML;
  assert.match(content,/1 MT \(1%\)/);assert.match(content,/확인되지 않음/);assert.doesNotMatch(content,/1200 mm|3\.5 m|자급율 0%/);
});
test('country capacity sidebar deduplicates pulp and aggregates grades per mill', () => {
  const rows=[{company:'C',mill:'A',region:'China',capa_mty:20,pulp_mty:70},{company:'C',mill:'A',region:'China',capa_mty:30,pulp_mty:70}];
  const h=harness({DATA_CAPA:rows});h.run("showCountryAnalysis('China')");const content=h.elements.get('sidebarContent').innerHTML;
  assert.match(content,/>70 <small/);assert.doesNotMatch(content,/>140 <small/);assert.match(content,/>50<\/td>/);
});
test('local file news sources are marked unavailable instead of browser links', () => {
  const h=harness({DATA_NEWS:{items:[{date:'2026-01-01',title:'Report',source:'Local',summary:'Text',url:'file:///c:/report.md'}]}});h.run('renderNews()');
  assert.doesNotMatch(h.elements.get('newsList').innerHTML,/href="file:/);assert.match(h.elements.get('newsList').innerHTML,/로컬/);
});
test('collapse updates arrow state along with content visibility', () => {
  const h=harness();h.run("toggleCollapse('priceReportContent')");
  assert.equal(h.elements.get('arrow-priceReportContent').textContent,'▼');
});
test('email summary aggregates Korea countries across companies', async () => {
  const h=harness({DATA_EXPORT:[{y:2026,m:2,c:'독일',v:10}],DATA_KR_COUNTRY:krRows});h.run('copyEmailSummary()');
  const text=await h.clipboard()[0].items['text/plain'].text();
  assert.match(text,/베트남: 60 MT/);assert.equal((text.split('■ 한국산')[1].match(/베트남/g)||[]).length,1);
});

test('combined Korean competitor filter recognizes the actual dataset company name', () => {
  const row={y:2026,m:1,q:'Q1',h:'H1',comp:'무림 울산+한국',c:'베트남',rg:'09_동남아',v:10,amt:100};
  const h=harness({DATA_KR_EXPORT:[row],DATA_KR_COUNTRY:[row]});h.run("KR.comp='무림+한국';krRefresh()");
  assert.equal(h.run('krGetF().reduce((s,x)=>s+x.v,0)'),10);
  assert.match(h.elements.get('krKpi').innerHTML,/무림\+한국/);
});
test('Korean KPI compares matching months instead of a partial year with a full year', () => {
  const rows=[{y:2026,m:1,comp:'한솔',v:20,amt:200},{y:2025,m:1,comp:'한솔',v:10,amt:100},{y:2025,m:12,comp:'한솔',v:1000,amt:10000}];
  const h=harness({DATA_KR_EXPORT:rows});h.run('krRefresh()');
  assert.match(h.elements.get('krKpi').innerHTML,/▲100\.0%/);assert.match(h.elements.get('krKpi').innerHTML,/2026\/2025년 1월/);
});
test('KR trend redraws selected company and month data rather than static annual totals', () => {
  const h=harness({DATA_KR_EXPORT:krRows});h.elements.get('krMonthFrom').value='2';h.elements.get('krMonthTo').value='2';h.run("KR.view='trend';KR.comp='한솔';krRefresh()");
  assert.equal(h.charts.get('krCTV').data.datasets[0].data[0],50);
  assert.match(h.charts.get('krCTV').data.labels[0],/2월/);
});
test('email copying falls back to plain text when HTML clipboard is unsupported', async () => {
  const h=harness();h.elements.get('emailPreviewHTML').innerHTML='<p>Draft</p>';h.elements.get('emailPreviewText').value='Draft';
  let copied='';h.context.ClipboardItem=undefined;h.context.navigator.clipboard.writeText=async t=>{copied=t;};
  await h.run('copyDraftHTML()');assert.equal(copied,'Draft');
});
test('email summary handles an empty export dataset without throwing', () => {
  const h=harness();assert.doesNotThrow(()=>h.run('copyEmailSummary()'));
  assert.match(h.elements.get('toastMsg').textContent,/데이터/);
});
test('price KPI does not show an unsupported fixed exchange rate as an actual USD price', () => {
  const rows=[1,2,3,4].map(m=>({y:2026,m,t:'코트지',r:'광동',p:5000}));
  const h=harness({DATA_PRICE:{prices:rows,usd:[]}});h.run('dPr()');
  assert.doesNotMatch(h.elements.get('priceKpi').innerHTML,/\$700/);
  assert.match(h.elements.get('priceKpi').innerHTML,/환율 데이터 없음/);
});
test('real datasets initialize export, price, company and news panels', () => {
  const h=harness();h.run(fs.readFileSync(path.join(__dirname,'../data.js'),'utf8'));h.run(fs.readFileSync(path.join(__dirname,'../dashboard_logic.js'),'utf8'));
  h.run("cnInit();go('krExport');go('price');go('company');go('news')");
  assert.ok(h.charts.get('cnCVol').data.datasets[0].data.length>0);
  assert.ok(h.charts.get('krCVol').data.datasets.length>0);
  assert.ok(h.elements.get('coGrid').innerHTML.includes('APP'));
  assert.ok(h.elements.get('newsList').innerHTML.includes('2026'));
});
test('price KPI renders one available month and marks missing calendar comparisons unavailable', () => {
  const h=harness({DATA_PRICE:{prices:[{y:2026,m:4,t:'코트지',r:'광동',p:5000}],usd:[]}});h.run('dPr()');
  assert.match(h.elements.get('priceKpi').innerHTML,/5,000/);assert.match(h.elements.get('priceKpi').innerHTML,/비교 데이터 없음/);
});
test('price MoM does not substitute an older nonconsecutive month', () => {
  const rows=[{y:2026,m:1,p:4000},{y:2026,m:2,p:4500},{y:2026,m:4,p:5000},{y:2025,m:12,p:4000}].map(x=>({...x,t:'코트지',r:'광동'}));
  const h=harness({DATA_PRICE:{prices:rows,usd:[]}});h.run('dPr()');
  assert.match(h.elements.get('priceKpi').innerHTML,/전월비[\s\S]*?비교 데이터 없음/);assert.doesNotMatch(h.elements.get('priceKpi').innerHTML,/11\.1%/);
});
test('nonfinite numeric results render as unavailable', () => {
  const h=harness();assert.equal(h.run('fmtF(Infinity)'),'-');assert.equal(h.run('fmt(Infinity)'),'-');
});
test('Korean insights recompute from selected company data', () => {
  const h=harness({DATA_KR_EXPORT:krRows});h.run("KR.view='insight';KR.comp='한솔';krRefresh()");
  assert.match(h.elements.get('krIns').innerHTML,/60/);assert.doesNotMatch(h.elements.get('krIns').innerHTML,/한솔-40\.3%/);
});
test('CN empty selection shows unavailable values and clears previously rendered charts', () => {
  const h=harness({DATA_EXPORT:[{y:2026,m:1,c:'독일',_rg:'04_서유럽',v:20}]});h.run('cnRefresh()');
  h.elements.get('cnRegion').value='유럽';h.run("_cnSelectedCountries=['태국'];cnRefresh()");
  assert.match(h.elements.get('cnKpi').innerHTML,/데이터 없음/);assert.doesNotMatch(h.elements.get('cnKpi').innerHTML,/>0<\/div>/);
  assert.match(h.elements.get('cnTbl').innerHTML,/데이터 없음/);assert.equal(h.charts.has('cnCVol'),false);assert.equal(h.charts.has('cnCReg'),false);
});
test('CN reversed month range shows validation instead of a zero export total', () => {
  const h=harness({DATA_EXPORT:[{y:2026,m:1,c:'독일',_rg:'04_서유럽',v:20}]});h.elements.get('cnMonthFrom').value='5';h.elements.get('cnMonthTo').value='2';h.run('cnRefresh()');
  assert.match(h.elements.get('cnKpi').innerHTML,/시작 월/);assert.equal(h.charts.size,0);
});
test('CN explicit observed zero remains a zero export volume with undefined shares', () => {
  const h=harness({DATA_EXPORT:[{y:2026,m:1,c:'독일',_rg:'04_서유럽',v:0}]});h.run('cnRefresh()');
  assert.match(h.elements.get('cnKpi').innerHTML,/총 수출량<\/div>\s*<div class="kpi-v [^"]*">0<\/div>/);
  assert.doesNotMatch(h.elements.get('cnTbl').innerHTML,/NaN|Infinity/);assert.equal(h.charts.get('cnCVol').data.datasets[0].data[0],0);
});
test('KR empty country and region intersection clears all overview and detail views', () => {
  const h=harness({DATA_KR_EXPORT:krRows,DATA_KR_COUNTRY:krRows,DATA_KR_REGION:krRows});h.run("KR.view='country';krRefresh();krTrend();krRegion()");
  h.elements.get('krRegionFilter').value='유럽';h.run("_krSelectedCountries=['베트남'];krRefresh()");
  assert.match(h.elements.get('krKpi').innerHTML,/데이터 없음/);assert.doesNotMatch(h.elements.get('krKpi').innerHTML,/\$0(?:M|\.0)/);
  for(const id of ['krTbl','krRTbl','krCTbl'])assert.match(h.elements.get(id).innerHTML,/데이터 없음/);
  assert.equal(h.charts.size,0);
});
test('KR reversed range remains invalid when country or region views refresh directly', () => {
  const h=harness({DATA_KR_EXPORT:krRows,DATA_KR_COUNTRY:krRows,DATA_KR_REGION:krRows});h.elements.get('krMonthFrom').value='5';h.elements.get('krMonthTo').value='2';h.run('krRefresh();krRegion();krCountryFilter()');
  for(const id of ['krKpi','krRTbl','krCTbl'])assert.match(h.elements.get(id).innerHTML,/시작 월/);assert.equal(h.charts.size,0);
});
test('KR explicit observed zero retains volume and amount without inventing an FOB', () => {
  const row={y:2026,m:1,comp:'한솔',c:'베트남',rg:'09_동남아',v:0,amt:0,q:'Q1',h:'H1'};
  const h=harness({DATA_KR_EXPORT:[row],DATA_KR_COUNTRY:[row],DATA_KR_REGION:[row]});h.run('krRefresh();krRegion();krCountryFilter()');
  assert.match(h.elements.get('krKpi').innerHTML,/총 수출량<\/div><div class="kpi-v">0<\/div>/);
  assert.match(h.elements.get('krKpi').innerHTML,/총 수출금액<\/div><div class="kpi-v">\$0M/);assert.doesNotMatch(h.elements.get('krKpi').innerHTML,/\$0\.0/);
  for(const id of ['krTbl','krRTbl','krCTbl'])assert.doesNotMatch(h.elements.get(id).innerHTML,/NaN|Infinity/);
});
test('CN comparison renders missing country observations unavailable instead of a fabricated decline', () => {
  const rows=[{y:2026,m:4,c:'독일',_rg:'04_서유럽',v:20},{y:2026,m:3,c:'프랑스',_rg:'04_서유럽',v:10}];
  const h=harness({DATA_EXPORT:rows});h.run("setCnMode('2m')");
  assert.equal(h.charts.get('cnCVol').data.datasets[0].data[0],null);
  assert.equal(h.charts.get('cnCVol').data.datasets[1].data[1],null);
  assert.doesNotMatch(h.elements.get('cnTbl').innerHTML,/100\.0%/);assert.match(h.elements.get('cnTbl').innerHTML,/데이터 없음/);
  assert.match(h.elements.get('cnKpi').innerHTML,/비교 불가/);
});
test('CN comparison with an absent current period clears charts rather than fabricating its date', () => {
  const h=harness({DATA_EXPORT:[{y:2025,m:12,c:'독일',_rg:'04_서유럽',v:20}]});h.run("setCnMode('2m')");
  assert.match(h.elements.get('cnKpi').innerHTML,/데이터 없음/);assert.equal(h.charts.size,0);
});
test('CN comparison retains a recorded zero when both country-month observations exist', () => {
  const rows=[{y:2026,m:4,c:'독일',_rg:'04_서유럽',v:0},{y:2026,m:3,c:'독일',_rg:'04_서유럽',v:10}];
  const h=harness({DATA_EXPORT:rows});h.run("setCnMode('2m')");
  assert.equal(h.charts.get('cnCVol').data.datasets[1].data[0],0);assert.match(h.elements.get('cnTbl').innerHTML,/100\.0%/);
});
