'use strict';
const fs = require('node:fs');
const path = require('node:path');
const COMPANY_IDS = ['app','chenming','huatai','sunpaper','sappi','upm','moorim','hankuk','oji','nippon','hansol','hokuetsu','arctic','burgo'];
const METRICS = ['revenue','operating_profit','net_profit','total_debt','cash','net_debt','equity','sales_volume','revenue_growth_pct','operating_margin_pct'];

function date(value, label, ceiling) {
 if (!/^\d{4}-\d{2}-\d{2}$/.test(value || '') || !Number.isFinite(Date.parse(value)) || new Date(value).toISOString().slice(0,10) !== value) throw new Error(`${label}: invalid date ${value}`);
 if (ceiling && value > ceiling) throw new Error(`${label}: future date ${value}`);
}
function https(value, label) {
 let url; try { url = new URL(value); } catch { throw new Error(`${label}: invalid URL`); }
 if (url.protocol !== 'https:' || url.username || url.password) throw new Error(`${label}: unsafe URL`);
 return url;
}
function canonicalUrl(value) {
 const url=https(value,'news');url.hash='';
 for(const key of [...url.searchParams.keys()]) if(/^utm_|^(fbclid|gclid)$/i.test(key))url.searchParams.delete(key);
 url.searchParams.sort();return url.href;
}
function validateBundle(bundle) {
 date(bundle.as_of,'as_of');
 if(!Array.isArray(bundle.companies)||!Array.isArray(bundle.news)||!Array.isArray(bundle.market))throw new Error('bundle arrays required');
 for(const company of bundle.companies){
  if(!COMPANY_IDS.includes(company.id))throw new Error(`unknown company ${company.id}`);
  date(company.verified_at,`${company.id} verified_at`,bundle.as_of);
  if(!company.business_model || !company.coverage_note)throw new Error(`${company.id}: business model and coverage note required`);
  const ids=new Set();
  for(const source of company.sources||[]){
   if(!source.id||ids.has(source.id)||!source.title)throw new Error(`${company.id}: invalid/duplicate source`);
   ids.add(source.id);https(source.url,`${company.id} source ${source.id}`);
   if(source.published_at)date(source.published_at,'source published_at',bundle.as_of);
   date(source.accessed_at,'source accessed_at',bundle.as_of);
  }
  if(!ids.size)throw new Error(`${company.id}: source required`);
  for(const entry of [...company.products||[],...company.financials||[],...company.guidance||[],...company.risks||[]]){
   if(!entry.source_ids?.length||entry.source_ids.some(id=>!ids.has(id)))throw new Error(`${company.id}: unresolved source reference`);
  }
  for(const financial of company.financials||[]){
   if(!financial.period||!financial.scope||financial.unit!=='million'||! /^[A-Z]{3}$/.test(financial.currency||''))throw new Error(`${company.id}: financial period/currency/unit/scope required`);
   for(const metric of METRICS)if(financial[metric]!=null&&(!Number.isFinite(financial[metric])||typeof financial[metric]!=='number'))throw new Error(`${company.id}: ${metric} must be finite number or null`);
  }
 }
 for(const entry of [...bundle.news,...bundle.market]){
  date(entry.date,'publication date',bundle.as_of);https(entry.url,'article');
  if(!entry.title||!entry.source||!entry.summary)throw new Error('article title, source, summary required');
  if((entry.company_ids||[]).some(id=>!COMPANY_IDS.includes(id)))throw new Error('unknown article company');
 }
 return bundle;
}
function mergeArticles(entries){
 const map=new Map();
 for(const entry of entries){
  const key=canonicalUrl(entry.url),prior=map.get(key);
  if(prior){prior.company_ids=[...new Set([...(prior.company_ids||[]),...(entry.company_ids||[])])].sort();}
  else map.set(key,{...entry,url:key,company_ids:[...(entry.company_ids||[])]});
 }
 return [...map.values()].sort((a,b)=>b.date.localeCompare(a.date)||a.title.localeCompare(b.title));
}
function mergeBundles(bundles){
 const seen=new Set(),companies=[];
 for(const bundle of bundles){validateBundle(bundle);for(const company of bundle.companies){if(seen.has(company.id))throw new Error(`duplicate company ${company.id}`);seen.add(company.id);companies.push(company);}}
 companies.sort((a,b)=>COMPANY_IDS.indexOf(a.id)-COMPANY_IDS.indexOf(b.id));
 return {schema_version:1,as_of:bundles.map(x=>x.as_of).sort().at(-1),source_policy:[...new Set(bundles.map(x=>x.source_policy).filter(Boolean))].join(' / '),companies,news:mergeArticles(bundles.flatMap(x=>x.news)),market:mergeArticles(bundles.flatMap(x=>x.market))};
}
function seedNews(bundle){
 return {schema_version:1,reviewed_at:bundle.as_of,last_successful_fetch:null,entries:bundle.news.map(n=>({id:n.id,company:n.company_ids[0]||'market',company_ids:n.company_ids,date:n.date,title:n.title,summary:n.summary,url:n.url,source:n.source,source_id:`research:${n.id}`,category:n.source_type||'official',evidence_type:n.source_type==='official'?'official_release':'unverified',review_status:'reviewed_source',collection_method:'manual_review',collected_at:`${bundle.as_of}T00:00:00Z`,topic:n.topic,significance:n.significance}))};
}
function build({seed=false}={}){
 const root=path.resolve(__dirname,'..');
 const bundle=mergeBundles(['asia.json','global.json'].map(file=>JSON.parse(fs.readFileSync(path.join(__dirname,file),'utf8'))));
 const missing=COMPANY_IDS.filter(id=>!bundle.companies.some(x=>x.id===id));if(missing.length)throw new Error(`missing companies: ${missing.join(', ')}`);
 const text=JSON.stringify(bundle,null,2).replace(/</g,'\\u003c').replace(/\u2028/g,'\\u2028').replace(/\u2029/g,'\\u2029');
 fs.writeFileSync(path.join(root,'intelligence_data.js'),`// Generated by node research/build.cjs; edit research/*.json.\nwindow.COMPETITOR_RESEARCH = ${text};\n`);
 if(seed){fs.mkdirSync(path.join(root,'intelligence'),{recursive:true});const target=path.join(root,'intelligence/news.json');if(fs.existsSync(target))throw new Error('seed refuses to overwrite existing collected news');fs.writeFileSync(target,JSON.stringify(seedNews(bundle),null,2)+'\n');}
 console.log(`Verified ${bundle.companies.length} companies, ${bundle.news.length} news, ${bundle.market.length} market records`);
 return bundle;
}
module.exports={COMPANY_IDS,validateBundle,canonicalUrl,mergeBundles,seedNews,build};
if(require.main===module)build({seed:process.argv.includes('--seed-news')});
