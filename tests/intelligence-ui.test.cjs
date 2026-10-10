const test = require('node:test');
const assert = require('node:assert/strict');
const intel = require('../competitor_intelligence.js');
const vm = require('node:vm');
const fs = require('node:fs');
const path = require('node:path');

test('unknown financials remain unknown and a disclosed zero remains zero', () => {
  for (const value of [null, undefined, '', NaN, Infinity, '12']) assert.equal(intel.formatMetric(value), '미공개/확인 필요');
  assert.equal(intel.formatMetric(0), '0');
  assert.equal(intel.formatMetric(-1250.5), '-1,250.5');
});

test('source links allow HTTPS only without embedded credentials', () => {
  for (const value of ['javascript:alert(1)', 'data:text/html,hello', '//evil.test', 'http://example.com', 'https://user:pass@example.com', 'https://example.com/\nfoo']) assert.equal(intel.safeURL(value), '');
  assert.equal(intel.safeURL('https://example.com/report?year=2025'), 'https://example.com/report?year=2025');
  assert.equal(intel.escapeHTML('<img src=x onerror="x">&'), '&lt;img src=x onerror=&quot;x&quot;&gt;&amp;');
});

test('filing mirrors stay distinct from primary official sources', () => {
  assert.equal(intel.sourceLabel({kind:'filing',url:'https://money.finance.sina.com.cn/corp/filing',title:'issuer filing mirrored by Sina'}),'공시 재게시본');
  assert.equal(intel.sourceLabel({kind:'filing',url:'https://www.hkexnews.hk/report.pdf'}),'공시');
  assert.equal(intel.normalizeNews({source_type:'official',date:'2026-08-01',url:'https://www.upm.com/news/verified'}).evidence_type,'official_release');
});

test('pipeline news single company converts to IDs and URLs deduplicate across sources', () => {
  const pipeline = {company:'upm', date:'2026-08-01', title:'실적 공시', url:'https://example.com/release#top', category:'official', collected_at:'2026-10-09T01:00:00Z'};
  const entries = intel.mergeNews([pipeline], [{...pipeline, company_ids:['upm'], url:'https://example.com/release', title:'중복'}], [{date:'2025-01-01',title:'과거',url:'https://example.com/older',keywords:['UPM']}]);
  assert.equal(entries.length, 2);
  assert.deepEqual(entries[0].company_ids, ['upm']);
  assert.equal(entries[0].title, '실적 공시');
  assert.equal(entries[0].evidence_type, 'official_release');
});

test('invalid dates and missing provenance cannot appear fresh or official', () => {
  assert.equal(intel.freshness(null, '2026-10-09').label, '확인일 없음');
  assert.equal(intel.freshness('not-a-date', '2026-10-09').label, '확인일 없음');
  assert.equal(intel.freshness('2026-10-10', '2026-10-09').label, '확인일 점검 필요');
  assert.equal(intel.freshness('2026-10-08', '2026-10-09').state, 'fresh');
  assert.equal(intel.freshness('2026-01-01', '2026-10-09').state, 'stale');
  assert.equal(intel.normalizeNews({title:'미확인',url:'#'}).evidence_type, 'unverified');
});

test('comparison labels retain reporting period currency unit and scope', () => {
  assert.equal(intel.financialContext({period:'2025 FY',currency:'JPY',unit:'million',scope:'연결 전체'}), '2025 FY · JPY 백만 · 연결 전체');
  assert.match(intel.financialContext({}), /미공개\/확인 필요/);
  assert.match(intel.financialContext({period:'FY2025',currency:'JPY',unit:'million',scope:'연결',period_start:'2025-04-01',period_end:'2026-03-31'}),/2025-04-01.*2026-03-31/);
});

test('failed/manual source updates retain partial coverage and last success', () => {
  const result = intel.updateState({overall:{status:'partial',last_successful_fetch:'2026-10-08'},sources:[{status:'success'},{status:'error'},{status:'manual'}]}, true);
  assert.equal(result.state, 'partial');
  assert.equal(result.success, 1);
  assert.equal(result.errors, 1);
  assert.equal(result.manual, 1);
  assert.equal(result.lastSuccess, '2026-10-08');
  assert.equal(intel.updateState(null, false).state, 'fallback');
});

test('collector top-level errors and missing status are never labelled healthy', () => {
  assert.equal(intel.updateState({status:'error',sources:[{status:'error'}]},true).state, 'error');
  assert.equal(intel.updateState({status:'partial',sources:[{status:'success'},{status:'error'}]},true).state, 'partial');
  assert.equal(intel.updateState({status:'unknown',sources:[{status:'unknown'}]},true).state, 'partial');
  assert.equal(intel.updateState(null,true).state, 'partial');
  assert.equal(intel.updateState({status:'success',sources:[{status:'success'}]},false).state, 'fallback');
});

test('metric definitions preserve exact reported profit and debt definitions', () => {
  assert.equal(intel.metricDefinition({metric_definitions:{operating_profit:'비교가능 EBIT'}},{},'operating_profit'), '비교가능 EBIT');
  assert.equal(intel.metricDefinition({metric_definitions:{operating_profit:'EBIT'}},{metric_definitions:{operating_profit:'조정 EBIT'}},'operating_profit'), '조정 EBIT');
  assert.equal(intel.metricDefinition({},{},'operating_profit'), '보고 정의 확인 필요');
  assert.equal(intel.metricDefinition({metric_definitions:{growth:'동일 범위 전년 매출 대비 계산'}},{},'revenue_growth_pct'), '동일 범위 전년 매출 대비 계산');
  assert.match(intel.metricDefinition({},{},'operating_margin_pct'), /영업이익.*매출/);
});

test('a later fetch failure preserves news while reporting the current attempt failed', async () => {
  let failed = false;
  const context = vm.createContext({URL, console, setTimeout, clearTimeout, document:{readyState:'loading',addEventListener(){},getElementById(){return null;}},location:{protocol:'http:'},
    fetch:async url => {
      if (failed) throw new Error('offline');
      return {ok:true,json:async () => url.includes('update-log') ? {status:'success',sources:[{status:'success'}]} : {schema_version:1,entries:[{title:'공식 발표',date:'2026-08-01',company:'upm',url:'https://example.com/official',category:'official'}]}};
    }});
  context.window = context;
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../competitor_intelligence.js'),'utf8'),context);
  await context.CompetitorIntelligence.refreshNews();
  assert.equal(context.CompetitorIntelligence.getNewsSnapshot().status.state, 'success');
  failed = true;
  await context.CompetitorIntelligence.refreshNews();
  const snapshot = context.CompetitorIntelligence.getNewsSnapshot();
  assert.equal(snapshot.entries.length, 1);
  assert.equal(snapshot.status.state, 'fallback');
  assert.equal(snapshot.fetchPhase, 'unavailable');
});

function browserHarness() {
  const elements = new Map();
  function element(tag = 'div') {
    let markup = '';
    const node = {tagName:tag,children:[],dataset:{},style:{},value:'',listeners:{},
      setAttribute(key,value) {this[key]=value;},
      addEventListener(key,handler) {this.listeners[key]=handler;},
      appendChild(child) {this.children.push(child); if(child.id) elements.set(child.id,child);},
      prepend(child) {this.children.unshift(child); if(child.id) elements.set(child.id,child);},
      before(child) {if(child.id) elements.set(child.id,child);}, after(child) {if(child.id) elements.set(child.id,child);}, focus(){this.focused=true;},
      querySelectorAll(selector) {const match=/^\[data-([\w-]+)\]$/.exec(selector); return match ? this.children.filter(child => child.dataset[match[1].replace(/-([a-z])/g,(_,c)=>c.toUpperCase())] !== undefined) : [];},
      set innerHTML(value) {
        markup = value; this.children=[];
        for (const match of value.matchAll(/<(\w+)\b([^>]*)>/g)) {
          const child=element(match[1]); const attrs=match[2]; child.id=/\bid="([^"]*)"/.exec(attrs)?.[1];
          for (const attr of attrs.matchAll(/data-([\w-]+)="([^"]*)"/g)) child.dataset[attr[1].replace(/-([a-z])/g,(_,c)=>c.toUpperCase())]=attr[2];
          child.checked=/\bchecked\b/.test(attrs); this.children.push(child); if(child.id) elements.set(child.id,child);
        }
      }, get innerHTML(){return markup;}
    };
    return node;
  }
  for(const id of ['sec-company','coDetail','coProfile','coGrid','sec-news','newsList','newsUpdate','newsCoSel']) {const node=element();node.id=id;elements.set(id,node);}
  elements.get('newsCoSel').value='all';
  const calls=[];
  const context=vm.createContext({URL,console,setTimeout,clearTimeout,location:{protocol:'file:'},
    document:{readyState:'complete',getElementById:id=>elements.get(id)||null,createElement:element},
    DATA_COMPANIES:[{id:'upm',name:'UPM <img onerror=x>',country:'핀란드',listed:true},{id:'sappi',name:'SAPPI',country:'남아공/유럽',listed:true}],
    DATA_NEWS:{items:[]}, _S:{cf:'all',nk:'all'}, selectCo:id=>calls.push(id),
    COMPETITOR_RESEARCH:{as_of:'2026-10-09',companies:[{id:'upm',business_model:'<script>alert(1)</script>',segments:['펄프'],sources:[{id:'src',url:'https://example.com/ir',title:'공식 자료'}],financials:[{period:'2026 H1',currency:'EUR',unit:'million',scope:'연결 전체',revenue:100,operating_profit:null,source_ids:['src'],metric_definitions:{operating_profit:'비교가능 EBIT'}}]}],news:[],market:[]}});
  context.window=context;
  vm.runInContext(fs.readFileSync(path.join(__dirname,'../competitor_intelligence.js'),'utf8'),context);
  return {context,elements,calls};
}

test('browser integration preserves original company selection and renders safe research and exact comparison criteria', () => {
  const {context,elements,calls}=browserHarness();
  assert.deepEqual(calls,['upm']);
  assert.match(elements.get('ci-legacy-note').textContent,/기준일.*출처.*확인/);
  const grid=elements.get('coGrid');
  assert.ok(grid.innerHTML.includes('&lt;img onerror=x&gt;'));
  assert.ok(!grid.innerHTML.includes('<script>'));
  const input=grid.querySelectorAll('[data-ci-compare]')[0];input.checked=true;input.listeners.change();
  const comparison=elements.get('ci-comparison');
  assert.equal(comparison.hidden,false);
  assert.match(comparison.innerHTML,/2026 H1 · EUR 백만 · 연결 전체/);
  assert.match(comparison.innerHTML,/비교가능 EBIT/);
  assert.match(comparison.innerHTML,/미공개\/확인 필요/);
  const panel=elements.get('ci-company-research');
  const financial=panel.querySelectorAll('[data-ci-tab]').find(node=>node.dataset.ciTab==='financial');
  financial.listeners.click();
  assert.match(panel.innerHTML,/지표 정의/);
  assert.match(panel.innerHTML,/비교가능 EBIT/);
  assert.match(elements.get('ci-news-status').innerHTML,/저장된 조사·뉴스 표시/);
  context.CompetitorIntelligence.init();
  assert.equal(calls.length,1);
});
