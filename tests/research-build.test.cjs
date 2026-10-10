const test = require('node:test');
const assert = require('node:assert/strict');
const builder = require('../research/build.cjs');
const record = () => ({as_of:'2026-10-09',source_policy:'공식 출처 확인',companies:[{id:'upm',business_model:'제지 및 소재',segments:['제지'],products:[{name:'인쇄용지',applications:'인쇄',competitive_position:'사업부',source_ids:['ir']}],geographies:['유럽'],financials:[{period:'2026H1',currency:'EUR',unit:'million',scope:'consolidated',revenue:4000,total_debt:null,source_ids:['ir']}],guidance:[],risks:[{text:'수요 변동',source_ids:['ir']}],sources:[{id:'ir',title:'공식 실적',url:'https://www.upm.com/results',published_at:'2026-07-23',accessed_at:'2026-10-09',kind:'results'}],coverage_note:'차입금 미확인',verified_at:'2026-10-09'}],news:[{id:'n1',date:'2026-07-23',title:'실적 발표',summary:'보고 실적',company_ids:['upm'],url:'https://www.upm.com/news/1?utm_source=x',source:'UPM',source_type:'official',topic:'실적'}],market:[]});
test('research validation retains absent metrics and exact scope and currency',()=>{
 const data=record(); builder.validateBundle(data); const out=builder.mergeBundles([data]);
 assert.equal(out.companies[0].financials[0].total_debt,null); assert.equal(out.companies[0].financials[0].currency,'EUR'); assert.equal(out.companies[0].financials[0].scope,'consolidated');
});
test('unknown source references, unsafe source URLs and malformed dates are rejected',()=>{
 for(const mutate of [d=>d.companies[0].products[0].source_ids=['fake'],d=>d.companies[0].sources[0].url='javascript:alert(1)',d=>d.companies[0].sources[0].published_at='2026-02-30']){
  const data=record();mutate(data);assert.throws(()=>builder.validateBundle(data));
 }
});
test('future publication dates and nonfinite financial values are rejected',()=>{
 const data=record();data.news[0].date='2026-10-10';assert.throws(()=>builder.validateBundle(data),/date|future/i);
 const num=record();num.companies[0].financials[0].revenue=Infinity;assert.throws(()=>builder.validateBundle(num),/revenue|number/i);
});
test('bundle merging rejects duplicate companies and unions news company mappings',()=>{
 const a=record();assert.throws(()=>builder.mergeBundles([a,a]),/duplicate/i);
 const b=record();b.companies[0].id='sappi';b.news[0].id='n2';b.news[0].company_ids=['sappi'];b.news[0].url='https://www.upm.com/news/1';
 const merged=builder.mergeBundles([a,b]);assert.equal(merged.news.length,1);assert.deepEqual(merged.news[0].company_ids,['sappi','upm']);
});
test('seed news distinguishes reviewed source data from a successful automated collection',()=>{
 const news=builder.seedNews(builder.mergeBundles([record()]));assert.equal(news.last_successful_fetch,null);assert.equal(news.entries[0].review_status,'reviewed_source');assert.equal(news.entries[0].evidence_type,'official_release');assert.equal(news.entries[0].collection_method,'manual_review');assert.equal(news.entries[0].company,'upm');assert.equal(news.reviewed_at,'2026-10-09');
});
