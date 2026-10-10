const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const root=path.resolve(__dirname,'..');
const {mergeBundles,COMPANY_IDS}=require('../research/build.cjs');
const read=file=>JSON.parse(fs.readFileSync(path.join(root,file),'utf8'));
test('shipped research has all monitored companies and valid fact-level evidence',()=>{
 const research=mergeBundles([read('research/asia.json'),read('research/global.json')]);
 assert.deepEqual(research.companies.map(c=>c.id),COMPANY_IDS);
 assert.ok(research.news.length>=14);
 assert.ok(research.market.length>=4);
 for(const record of research.companies){assert.ok(record.products.length>0,record.id);assert.ok(record.sources.length>=2,record.id);}
});
test('generated browser data matches audited source JSON without executable text injection',()=>{
 const expected=mergeBundles([read('research/asia.json'),read('research/global.json')]);
 const context=vm.createContext({window:{}});
 vm.runInContext(fs.readFileSync(path.join(root,'intelligence_data.js'),'utf8'),context);
 assert.deepEqual(JSON.parse(JSON.stringify(context.window.COMPETITOR_RESEARCH)),expected);
});
test('all local page scripts and styles exist and intelligence loads after shared dashboard code',()=>{
 const html=fs.readFileSync(path.join(root,'index.html'),'utf8');
 for(const match of html.matchAll(/(?:src|href)="([^"#]+\.(?:js|css))"/g))if(!match[1].startsWith('http'))assert.ok(fs.existsSync(path.join(root,match[1])),match[1]);
 assert.ok(html.indexOf('src="intelligence_data.js"')>html.indexOf('src="dashboard_logic.js"'));
 assert.ok(html.indexOf('src="competitor_intelligence.js"')>html.indexOf('src="intelligence_data.js"'));
});
