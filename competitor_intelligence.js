/* 출처를 확인할 수 있는 경쟁사 조사. 기존 대시보드에 독립적으로 추가됩니다. */
(function (root, factory) {
  'use strict';
  const api = factory(root);
  if (typeof module === 'object' && module.exports) module.exports = api;
  if (root && root.document) {
    root.CompetitorIntelligence = api;
    if (root.document.readyState === 'loading') root.document.addEventListener('DOMContentLoaded', api.init, { once: true });
    else api.init();
  }
})(typeof window !== 'undefined' ? window : globalThis, function (root) {
  'use strict';
  const UNKNOWN = '미공개/확인 필요';
  const array = value => Array.isArray(value) ? value : [];
  const escapeHTML = value => String(value ?? '').replace(/[&<>"']/g, char => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;',"'":'&#39;'}[char]));
  const E = escapeHTML;
  function safeURL(value) {
    if (typeof value !== 'string' || /[\u0000-\u0020\u007f]/.test(value) || !/^https:\/\//i.test(value)) return '';
    try {
      const url = new URL(value);
      return url.protocol === 'https:' && !url.username && !url.password ? url.href : '';
    } catch (_) { return ''; }
  }
  function formatMetric(value, suffix = '') {
    return typeof value === 'number' && Number.isFinite(value)
      ? new Intl.NumberFormat('ko-KR', { maximumFractionDigits: 2 }).format(value) + suffix : UNKNOWN;
  }
  function validDate(value) {
    if (typeof value !== 'string' || !/^\d{4}-\d{2}-\d{2}(?:T.*)?$/.test(value)) return null;
    const date = new Date(value);
    return Number.isNaN(date.getTime()) || date.toISOString().slice(0,10) !== value.slice(0,10) ? null : date;
  }
  function dateLabel(value) { return validDate(value) ? value.slice(0,10) : '확인일 없음'; }
  function freshness(value, now = new Date().toISOString()) {
    const date = validDate(value); const current = validDate(now);
    if (!date || !current) return { state:'unknown', label:'확인일 없음', days:null };
    const days = Math.floor((current - date) / 86400000);
    if (days < 0) return { state:'unknown', label:'확인일 점검 필요', days };
    return { state:days > 120 ? 'stale' : 'fresh', label:days > 120 ? '자료 재확인 필요' : '최근 확인', days };
  }
  function financialContext(record = {}) {
    const dates = validDate(record.period_start) && validDate(record.period_end) ? ` (${record.period_start} ~ ${record.period_end})` : validDate(record.period_end) ? ` (기준 ${record.period_end})` : '';
    return `${record.period || UNKNOWN}${dates} · ${record.currency || UNKNOWN} ${record.unit === 'million' ? '백만' : (record.unit || UNKNOWN)} · ${record.scope || UNKNOWN}`;
  }
  function metricDefinition(company = {}, row = {}, key) {
    const reported = row.metric_definitions?.[key] || company.metric_definitions?.[key];
    if (reported) return reported;
    if (key === 'revenue_growth_pct') return row.metric_definitions?.growth || company.metric_definitions?.growth || '공시 성장률 또는 동일 범위 비교기간 매출에서 계산: (당기 / 비교기간 매출 - 1) × 100. 비교기간은 원문 확인.';
    if (key === 'operating_margin_pct') return '같은 보고기간·범위의 영업이익 / 매출 × 100. 영업이익 정의를 함께 확인.';
    return '보고 정의 확인 필요';
  }
  function sourceLabel(source = {}) {
    if (source.provenance === 'filing_mirror' || /(?:^|\.)sina\.com\.cn\//i.test(source.url || '') || /mirror/i.test(source.title || '')) return '공시 재게시본';
    return ({annual_report:'연차보고서',results:'실적 발표',financial_report:'실적 보고서',filing:'공시',exchange_filing:'거래소 공시',news:'공식 뉴스',official_release:'공식 발표',company:'회사 자료'})[source.kind] || '출처 자료';
  }
  function notesText(notes) {
    if (typeof notes === 'string') return notes;
    if (Array.isArray(notes)) return notes.filter(value => typeof value === 'string').join(' · ');
    return '';
  }
  function normalizeNews(record = {}) {
    const companyIds = array(record.company_ids).length ? array(record.company_ids) : (typeof record.company === 'string' && record.company ? [record.company] : []);
    const url = safeURL(record.url);
    return {
      ...record, title:String(record.title || '제목 확인 필요'), summary:String(record.summary || ''),
      date:validDate(record.date) ? record.date.slice(0,10) : '', url,
      company_ids:companyIds.filter(id => typeof id === 'string'), keywords:array(record.keywords),
      evidence_type:url ? record.evidence_type || (record.category === 'official' || record.source_type === 'official' ? 'official_release' : 'unverified') : 'unverified',
      source:String(record.source || '출처 확인 필요')
    };
  }
  function mergeNews(...groups) {
    const seen = new Set(); const result = [];
    groups.forEach(group => array(group).forEach(raw => {
      if (!raw || typeof raw !== 'object') return;
      const entry = normalizeNews(raw);
      const key = entry.url ? entry.url.split('#')[0] : `${entry.date}|${entry.title}|${entry.company_ids.join(',')}`;
      if (seen.has(key)) return;
      seen.add(key); result.push(entry);
    }));
    return result.sort((a,b) => b.date.localeCompare(a.date));
  }
  function updateState(log, fetched) {
    const overall = log && log.overall || {};
    const rawSources = log && log.sources;
    const sources = Array.isArray(rawSources) ? rawSources : (rawSources && typeof rawSources === 'object' ? Object.values(rawSources) : []);
    const success = sources.filter(source => source.status === 'success').length;
    const errors = sources.filter(source => source.status === 'error').length;
    const manual = sources.filter(source => source.status === 'manual').length;
    const unknown = sources.filter(source => !['success','error','manual'].includes(source.status)).length;
    const overallStatus = overall.status || log && log.status;
    let state = fetched ? 'success' : 'fallback';
    if (fetched && (!log || errors || manual || unknown || !sources.length || ['partial','error','unknown','never_success'].includes(overallStatus))) state = 'partial';
    if (fetched && overallStatus === 'error' && !success) state = 'error';
    return { state, success, errors, manual, unknown, lastSuccess:overall.last_successful_fetch || log && log.last_successful_fetch || null, sources };
  }
  const state = { mounted:false, search:'', comparisons:new Set(), activeId:'', tab:'business', news:[], newsFetched:false, log:null, fetchPhase:'loading', newsSearch:'', onlyOfficial:false, originalSelect:null };
  const document = root && root.document;
  const byId = id => document && document.getElementById(id);
  function companyData() { return typeof DATA_COMPANIES !== 'undefined' ? array(DATA_COMPANIES) : array(root.DATA_COMPANIES); }
  function researchData() { return typeof COMPETITOR_RESEARCH !== 'undefined' ? COMPETITOR_RESEARCH : root.COMPETITOR_RESEARCH || {companies:[],news:[],market:[]}; }
  function fallbackNews() { return typeof DATA_NEWS !== 'undefined' ? DATA_NEWS : root.DATA_NEWS || {items:[]}; }
  function legacyState() { return typeof _S !== 'undefined' ? _S : root._S || {}; }
  function companies() {
    const base = companyData(); const researched = array(researchData().companies);
    const ids = new Set(base.map(company => company.id));
    return [...base, ...researched.filter(company => !ids.has(company.id))];
  }
  function companyName(id) { return companies().find(company => company.id === id)?.name || id || '업계 전체'; }
  function research(id) { return array(researchData().companies).find(company => company.id === id) || {id, coverage_note:'이 회사의 상세 조사는 아직 등록되지 않았습니다. 공식 IR 자료 확인이 필요합니다.'}; }
  function sourceLink(source, text) {
    const url = safeURL(source && source.url);
    const title = E(text || source && source.title || '원문 확인');
    return url ? `<a href="${E(url)}" target="_blank" rel="noopener noreferrer" class="ci-source-link">${title}<span aria-hidden="true"> ↗</span></a>` : `<span class="ci-muted">${title} · 링크 확인 필요</span>`;
  }
  function references(ids, company) {
    const sources = array(company.sources);
    return array(ids).map(id => {
      const source = sources.find(item => item.id === id);
      return source ? sourceLink(source, source.title || '공식 출처') : '<span class="ci-muted">출처 확인 필요</span>';
    }).join(' · ') || '<span class="ci-muted">출처 확인 필요</span>';
  }
  const empty = message => `<div class="ci-empty">${E(message || UNKNOWN)}</div>`;
  const pill = text => `<span class="ci-pill">${E(text)}</span>`;
  function regionMatch(company, filter) {
    const country = company.country || '';
    if (!filter || filter === 'all') return true;
    if (filter === 'listed') return !!company.listed;
    if (filter === 'china') return country.includes('중국');
    if (filter === 'japan') return country.includes('일본');
    if (filter === 'korea') return country.includes('한국');
    if (filter === 'europe') return /유럽|핀란드|스웨덴|폴란드|이탈리아|벨기에|남아공/.test(country);
    return true;
  }
  function renderHeader() {
    const host = byId('ci-overview'); if (!host) return;
    const data = researchData(); const all = companies(); const records = array(data.companies);
    const withFinancials = records.filter(company => array(company.financials).some(row => typeof row.revenue === 'number' && Number.isFinite(row.revenue))).length;
    const withDebt = records.filter(company => array(company.financials).some(row => typeof row.total_debt === 'number' && Number.isFinite(row.total_debt))).length;
    const verified = records.filter(company => array(company.sources).some(source => safeURL(source.url))).length;
    host.innerHTML = `<div class="ci-overview-copy"><span class="ci-eyebrow">공시와 원문으로 읽는 경쟁 구도</span><h2>경쟁사 인텔리전스</h2><p>사업 구조부터 실적, 차입금, 경영진 전망까지.<br>공개된 사실과 확인이 필요한 항목을 함께 확인합니다.</p><div class="ci-overview-foot">조사 기준 ${E(dateLabel(data.as_of))} <span>·</span> 재무 수치는 각 보고기간 기준</div></div><div class="ci-kpis"><div><strong>${all.length}</strong><span>모니터링 기업</span></div><div><strong>${verified}<small> / ${all.length}</small></strong><span>원문 출처 연결</span></div><div><strong>${withFinancials}<small> / ${all.length}</small></strong><span>매출 확인 기업</span></div><div><strong>${withDebt}<small> / ${all.length}</small></strong><span>차입금 확인 기업</span></div></div>`;
  }
  function renderGrid() {
    const grid = byId('coGrid'); if (!grid) return;
    const query = state.search.trim().toLocaleLowerCase();
    const list = companies().filter(company => regionMatch(company, legacyState().cf)).filter(company => {
      const record = research(company.id);
      return [company.name,company.id,company.country,record.business_model,...array(record.segments),...array(record.products).map(product => product.name)].join(' ').toLocaleLowerCase().includes(query);
    });
    grid.innerHTML = list.map(company => {
      const record = research(company.id); const sources = array(record.sources).filter(source => safeURL(source.url));
      const fresh = freshness(record.verified_at);
      return `<article class="ci-company-card ${state.activeId === company.id ? 'ci-selected' : ''}"><button type="button" class="ci-company-open" data-ci-company="${E(company.id)}" aria-pressed="${state.activeId === company.id}"><span class="ci-company-country">${E(company.country || '지역 확인 필요')} <span>${company.listed ? '상장' : '비상장'}</span></span><strong>${E(company.name || company.id)}</strong><span class="ci-company-model">${E(record.business_model || '사업 모델 조사 대기')}</span><span class="ci-company-meta">${sources.length ? `${sources.length}개 출처 연결` : '출처 확인 필요'}<span class="ci-dot ci-${fresh.state}"></span>${E(fresh.label)}</span></button><label class="ci-compare-check"><input type="checkbox" data-ci-compare="${E(company.id)}" ${state.comparisons.has(company.id) ? 'checked' : ''}> 비교에 추가</label></article>`;
    }).join('') || empty('검색 조건에 맞는 회사가 없습니다. 회사명, 제품명 또는 지역을 바꿔보세요.');
    const count = byId('ci-company-count'); if (count) count.textContent = `${list.length}개 기업`;
    grid.querySelectorAll('[data-ci-company]').forEach(button => button.addEventListener('click', () => selectCompany(button.dataset.ciCompany)));
    grid.querySelectorAll('[data-ci-compare]').forEach(input => input.addEventListener('change', () => {
      if (input.checked && state.comparisons.size >= 4) {
        input.checked = false; byId('ci-compare-message').textContent = '비교 기업은 최대 4개까지 선택할 수 있습니다.'; return;
      }
      input.checked ? state.comparisons.add(input.dataset.ciCompare) : state.comparisons.delete(input.dataset.ciCompare);
      renderComparison();
    }));
  }
  function renderComparison() {
    const host = byId('ci-comparison'); if (!host) return;
    const ids = [...state.comparisons];
    byId('ci-compare-message').textContent = ids.length ? `${ids.length} / 4개 선택 · 기간, 통화, 사업 범위를 먼저 확인하세요.` : '기업 카드에서 최대 4개를 선택해 공시 수치를 나란히 확인하세요.';
    host.hidden = ids.length === 0;
    host.innerHTML = `<div class="ci-panel-heading"><div><span class="ci-eyebrow">선택 기업 비교</span><h3>공시 기준을 함께 보는 재무 비교</h3></div><button type="button" id="ci-clear-comparison" class="ci-button ci-button-quiet">선택 초기화</button></div><p class="ci-note">서로 다른 기간·통화·사업 범위와 이익 정의를 직접 순위로 비교하지 않습니다. 금액 단위는 각 열의 표기 기준입니다.</p><div class="ci-table-scroll"><table class="ci-table"><caption class="ci-sr-only">선택 기업의 최근 등록 재무 수치 및 공시 기준</caption><thead><tr><th scope="col">항목</th>${ids.map(id => `<th scope="col">${E(companyName(id))}</th>`).join('')}</tr></thead><tbody>${[
      ['사업 구조', record => record.business_model || UNKNOWN],
      ['주요 사업 부문', record => array(record.segments).join(' · ') || UNKNOWN],
      ['주요 제품', record => array(record.products).map(product => product.name).join(' · ') || UNKNOWN],
      ['보고 기준', record => financialContext(array(record.financials)[0])],
      ['매출', record => formatMetric(array(record.financials)[0]?.revenue)],
      ['영업이익 (보고 정의)', record => formatMetric(array(record.financials)[0]?.operating_profit)],
      ['순이익 (보고 정의)', record => formatMetric(array(record.financials)[0]?.net_profit)],
      ['총차입금 (총부채와 구분)', record => formatMetric(array(record.financials)[0]?.total_debt)],
      ['현금', record => formatMetric(array(record.financials)[0]?.cash)],
      ['순차입금 (공시값)', record => formatMetric(array(record.financials)[0]?.net_debt)],
      ['자본', record => formatMetric(array(record.financials)[0]?.equity)],
      ['판매량', record => { const row = array(record.financials)[0]; return formatMetric(row?.sales_volume, row?.sales_volume_unit ? ` ${row.sales_volume_unit}` : ''); }],
      ['매출 성장률 (공시/계산값)', record => formatMetric(array(record.financials)[0]?.revenue_growth_pct, '%')],
      ['영업이익률 (공시/계산값)', record => formatMetric(array(record.financials)[0]?.operating_margin_pct, '%')],
      ['성장률 기준', record => metricDefinition(record,array(record.financials)[0],'revenue_growth_pct')],
      ['이익률 기준', record => metricDefinition(record,array(record.financials)[0],'operating_margin_pct')],
      ['영업이익 정의', record => metricDefinition(record,array(record.financials)[0],'operating_profit')],
      ['순이익 정의', record => metricDefinition(record,array(record.financials)[0],'net_profit')],
      ['차입금 정의', record => metricDefinition(record,array(record.financials)[0],'total_debt')],
      ['공시 범위 참고', record => notesText(array(record.financials)[0]?.notes) || record.coverage_note || UNKNOWN],
    ].map(([title,value]) => `<tr><th scope="row">${E(title)}</th>${ids.map(id => `<td>${E(value(research(id)))}</td>`).join('')}</tr>`).join('')}<tr><th scope="row">출처</th>${ids.map(id => `<td>${references(array(research(id).financials)[0]?.source_ids,research(id))}</td>`).join('')}</tr></tbody></table></div>`;
    byId('ci-clear-comparison').addEventListener('click', () => { state.comparisons.clear(); renderComparison(); renderGrid(); });
  }
  function businessView(record) {
    const products = array(record.products);
    return `<div class="ci-business-layout"><div><h4>사업 모델과 수익 구조</h4><p class="ci-body-copy">${E(record.business_model || UNKNOWN)}</p><h4>사업 부문</h4><div class="ci-pills">${array(record.segments).map(pill).join('') || E(UNKNOWN)}</div><h4>주요 사업 지역</h4><div class="ci-pills">${array(record.geographies).map(pill).join('') || E(UNKNOWN)}</div><p class="ci-note">사업별 매출 비중은 공식 공시가 있을 때만 기재합니다. 제품 구성만으로 매출 비중을 추정하지 않습니다.</p></div><div class="ci-product-list"><h4>제품 · 적용 시장 · 경쟁 포지션</h4>${products.map(product => `<article class="ci-product"><h5>${E(product.name || UNKNOWN)}</h5><p><span>적용 시장</span>${E(Array.isArray(product.applications) ? product.applications.join(', ') : product.applications || UNKNOWN)}</p><p><span>포지션</span>${E(product.competitive_position || UNKNOWN)}</p><div class="ci-citations">${references(product.source_ids,record)}</div></article>`).join('') || empty('제품별 상세 정보는 공식 자료 확인이 필요합니다.')}</div></div>`;
  }
  function financialView(record) {
    const rows = array(record.financials); const latest = rows[0];
    return `<div class="ci-financial-intro"><div><h4>최근 등록 보고기간</h4><p>${E(latest ? financialContext(latest) : UNKNOWN)}</p></div><span class="ci-status-pill">공시·계산값 기준</span></div><p class="ci-note">매출·이익은 해당 보고기간 실적, 차입금·현금·자본은 보고서의 재무상태 기준입니다. 총차입금은 총부채가 아니며 순차입금은 직접 계산하지 않습니다. 판매량과 생산량을 구분하고, 연결 전체 실적을 인쇄용지 부문 실적으로 해석하지 않습니다.</p>${rows.length ? `<div class="ci-table-scroll"><table class="ci-table ci-financial-table"><caption>보고기간별 실적·차입금 이력 · 금액은 각 행의 통화와 단위 기준</caption><thead><tr>${['기간 / 범위 / 단위','매출','영업이익','순이익','총차입금','현금','순차입금','자본','판매량','매출 성장률','영업이익률','출처'].map(title => `<th scope="col">${E(title)}</th>`).join('')}</tr></thead><tbody>${rows.map(row => `<tr><th scope="row">${E(financialContext(row))}</th>${['revenue','operating_profit','net_profit','total_debt','cash','net_debt','equity'].map(key => `<td class="ci-number">${E(formatMetric(row[key]))}</td>`).join('')}<td>${E(formatMetric(row.sales_volume,row.sales_volume_unit ? ` ${row.sales_volume_unit}` : ''))}</td><td>${E(formatMetric(row.revenue_growth_pct,'%'))}</td><td>${E(formatMetric(row.operating_margin_pct,'%'))}</td><td>${references(row.source_ids,record)}</td></tr>`).join('')}</tbody></table></div>` : empty('원문으로 확인된 재무 수치가 없습니다. 공식 IR·공시에서 기간, 통화, 연결 범위 및 차입금 정의를 확인한 후 등록하세요.')}<p class="ci-note">성장률은 공시값 또는 동일 범위 매출에서 계산한 값이며, 영업이익률은 영업이익 / 매출 × 100입니다. 상세 산식과 비교기간은 지표 정의와 원문에서 확인하세요. 조정 영업이익·EBITDA는 영업이익과 구분합니다. 공시 원문에서 확인하지 못한 항목은 모두 “미공개/확인 필요”로 표시합니다.</p>`;
  }
  function outlookView(record) {
    return `<div class="ci-business-layout"><div><h4>경영진 가이던스</h4><p class="ci-note">경영진이 발표한 전망이며 확정 실적이나 대시보드 예측이 아닙니다. 발표 후 수정 여부는 원문에서 확인하세요.</p>${array(record.guidance).map(item => `<article class="ci-product"><div class="ci-guidance-label">경영진 전망 · ${E(item.period || '기간 확인 필요')}</div><p>${E(item.text || UNKNOWN)}</p><div class="ci-citations">${references(item.source_ids,record)}</div></article>`).join('') || empty('등록된 공식 가이던스가 없습니다. 미공개 여부 또는 후속 공시 확인이 필요합니다.')}</div><div><h4>공시에서 확인한 리스크와 점검 항목</h4>${array(record.risks).map(item => `<article class="ci-product"><p>${E(item.text || UNKNOWN)}</p><div class="ci-citations">${references(item.source_ids,record)}</div></article>`).join('') || empty('확인된 리스크 자료가 아직 등록되지 않았습니다.')}</div></div>`;
  }
  function financialDefinitions(record) {
    const labels = {revenue:'매출',operating_profit:'영업이익',net_profit:'순이익',total_debt:'총차입금',cash:'현금',net_debt:'순차입금',equity:'자본',sales_volume:'판매량',revenue_growth_pct:'매출 성장률 기준',operating_margin_pct:'영업이익률 산식'};
    return `<details class="ci-metric-definitions"><summary>지표 정의 · 보고기간별 공시 범위 참고</summary>${array(record.financials).map(row => `<div class="ci-definition-period"><h4>${E(financialContext(row))}</h4><dl>${Object.entries(labels).map(([key,label]) => `<div><dt>${E(label)}</dt><dd>${E(metricDefinition(record,row,key))}</dd></div>`).join('')}</dl>${notesText(row.notes) ? `<p class="ci-note">${E(notesText(row.notes))}</p>` : ''}</div>`).join('') || empty('등록된 지표 정의가 없습니다. 공시 원문 확인이 필요합니다.')}</details>`;
  }
  function sourcesView(record) {
    return `<p class="ci-note">${E(researchData().source_policy || '회사 IR·거래소 공시·공식 보도자료를 우선 확인합니다. 확인일은 원문을 조회한 날짜이며 공시 발표일과 다릅니다.')}</p><div class="ci-source-list">${array(record.sources).map(source => `<article class="ci-source-row"><div>${sourceLink(source)}<p>발표일 ${E(source.published_at ? dateLabel(source.published_at) : '미기재')} · 조회일 ${E(dateLabel(source.accessed_at))}</p></div><span class="ci-pill">${E(sourceLabel(source))}</span></article>`).join('') || empty('등록된 공식 출처가 없습니다.')}</div><p class="ci-note">자동 뉴스 수집은 재무 수치나 가이던스를 자동 검증하지 않습니다. 원문 검토 후 재무 조사 자료를 수동 갱신해야 합니다.</p>`;
  }
  function renderDetail() {
    const host = byId('ci-company-research'); if (!host || !state.activeId) return;
    const record = research(state.activeId); const fresh = freshness(record.verified_at);
    const tabs = [['business','사업 · 제품'],['financial','실적 · 차입금'],['outlook','가이던스 · 리스크'],['sources','출처 · 확인일']];
    host.innerHTML = `<div class="ci-panel-heading"><div><span class="ci-eyebrow">기업 조사 브리프</span><h3>${E(companyName(state.activeId))}</h3><div class="ci-detail-meta"><span class="ci-dot ci-${fresh.state}"></span>${E(fresh.label)} · 최종 확인 ${E(dateLabel(record.verified_at))}</div></div><button type="button" id="ci-download-company" class="ci-button">조사 자료 내려받기 ↓</button></div>${record.coverage_note ? `<div class="ci-coverage-note"><strong>자료 범위</strong><span>${E(record.coverage_note)}</span></div>` : ''}<div class="ci-tabs" role="tablist" aria-label="기업 조사 항목">${tabs.map(([id,title]) => `<button type="button" role="tab" id="ci-tab-${id}" aria-selected="${state.tab === id}" aria-controls="ci-company-tabpanel" tabindex="${state.tab === id ? '0' : '-1'}" data-ci-tab="${id}">${title}</button>`).join('')}</div><div id="ci-company-tabpanel" role="tabpanel" aria-labelledby="ci-tab-${E(state.tab)}" tabindex="0">${state.tab === 'financial' ? financialView(record) + financialDefinitions(record) : state.tab === 'outlook' ? outlookView(record) : state.tab === 'sources' ? sourcesView(record) : businessView(record)}</div>`;
    const buttons = [...host.querySelectorAll('[data-ci-tab]')];
    buttons.forEach((button,index) => {
      button.addEventListener('click', () => { state.tab = button.dataset.ciTab; renderDetail(); byId(`ci-tab-${state.tab}`).focus(); });
      button.addEventListener('keydown', event => {
        let target = null;
        if (event.key === 'ArrowRight') target = (index + 1) % tabs.length;
        if (event.key === 'ArrowLeft') target = (index - 1 + tabs.length) % tabs.length;
        if (event.key === 'Home') target = 0;
        if (event.key === 'End') target = tabs.length - 1;
        if (target !== null) { event.preventDefault(); state.tab = tabs[target][0]; renderDetail(); byId(`ci-tab-${state.tab}`).focus(); }
      });
    });
    byId('ci-download-company').addEventListener('click', () => downloadJSON({as_of:researchData().as_of,company:companyName(state.activeId),...record}, `${state.activeId}-research.json`));
  }
  function selectCompany(id) {
    if (!companies().some(company => company.id === id)) return;
    state.activeId = id;
    if (state.originalSelect && companyData().some(company => company.id === id)) state.originalSelect(id);
    const profile = byId('coProfile');
    if (profile && !byId('ci-legacy-note')) {
      const notice = document.createElement('p'); notice.id = 'ci-legacy-note'; notice.className = 'ci-note';
      notice.textContent = '아래 설비·생산능력·종목 정보는 기존 참고 자료입니다. 기준일과 출처 확인이 필요하며, 최신 공시 분석은 위 조사 패널에서 확인하세요.';
      profile.prepend(notice);
    }
    const detail = byId('coDetail'); if (detail) detail.style.display = '';
    renderGrid(); renderDetail();
  }
  function downloadJSON(data, filename) {
    if (!document || !root.URL || !root.URL.createObjectURL) return;
    const url = root.URL.createObjectURL(new Blob([JSON.stringify(data,null,2)], {type:'application/json;charset=utf-8'}));
    const anchor = document.createElement('a'); anchor.href = url; anchor.download = filename;
    document.body.appendChild(anchor); anchor.click(); anchor.remove(); root.setTimeout(() => root.URL.revokeObjectURL(url), 1000);
  }
  function allNews() { return mergeNews(state.news,array(researchData().news),array(fallbackNews().items)); }
  function getNewsSnapshot() { return {entries:allNews(), status:updateState(state.log,state.newsFetched), fetchPhase:state.fetchPhase}; }
  function renderNews() {
    const host = byId('newsList'); if (!host) return;
    const filter = byId('newsCoSel')?.value || 'all'; const keyword = legacyState().nk || 'all'; const query = state.newsSearch.toLocaleLowerCase().trim();
    const items = allNews().filter(item => {
      const text = [item.title,item.summary,...item.keywords].join(' ').toLocaleLowerCase();
      const co = companies().find(company => company.id === filter);
      const legacyMatch = co && item.company_ids.length === 0 && (text.includes(filter.toLocaleLowerCase()) || array(item.keywords).some(tag => co.name.toLocaleLowerCase().includes(String(tag).toLocaleLowerCase())));
      return (filter === 'all' || item.company_ids.includes(filter) || legacyMatch) && (keyword === 'all' || text.includes(keyword)) && (!query || text.includes(query)) && (!state.onlyOfficial || item.evidence_type === 'official_release');
    });
    host.innerHTML = items.map(item => `<article class="ci-news-item"><div class="ci-news-meta"><span class="ci-news-date">${E(item.date || '발표일 확인 필요')}</span><span>${E(item.source)}</span><span class="ci-evidence ${item.evidence_type === 'official_release' ? 'ci-evidence-official' : ''}">${item.evidence_type === 'official_release' ? '공식 발표' : item.evidence_type === 'regulatory_filing' ? '공시' : '출처 검토 필요'}</span></div><h3>${item.url ? sourceLink(item,item.title) : E(item.title)}</h3>${item.summary ? `<p>${E(item.summary)}</p>` : ''}<div class="ci-pills">${item.company_ids.map(id => pill(companyName(id))).join('')}${item.keywords.map(pill).join('')}</div>${item.collected_at ? `<div class="ci-news-collected">수집 ${E(dateLabel(item.collected_at))} · 발표일과 구분</div>` : ''}</article>`).join('') || empty('현재 필터에 해당하는 뉴스가 없습니다. 선택 조건을 바꿔 확인하세요.');
    const update = byId('newsUpdate'); if (update) update.textContent = `${items.length}건 · 조사 기준 ${dateLabel(researchData().as_of)}`;
    renderUpdateStatus();
  }
  function renderUpdateStatus() {
    const host = byId('ci-news-status'); if (!host) return;
    const status = updateState(state.log,state.newsFetched);
    const titles = {success:'공식 뉴스 파일 반영',partial:'일부 정보원 확인 필요',error:'뉴스 갱신 오류 · 기존 자료 표시',fallback:'저장된 조사·뉴스 표시'};
    const current = state.fetchPhase === 'loading' ? '갱신 파일 확인 중' : titles[status.state];
    const lastSuccess = status.lastSuccess || state.newsFileSuccess;
    host.innerHTML = `<div class="ci-update-summary"><span class="ci-dot ci-${status.state === 'success' ? 'fresh' : 'stale'}"></span><strong>${E(current)}</strong><span>최근 수집 성공 ${E(dateLabel(lastSuccess))}</span></div><p>${state.fetchPhase === 'loading' ? '저장된 출처 자료를 먼저 표시하고 최신 수집 파일을 확인합니다.' : state.fetchPhase === 'unavailable' ? '수집 파일을 불러오지 못했습니다. 저장된 자료를 표시합니다. 로컬 파일로 열었다면 HTTP 서버에서 수집 상태를 확인하세요.' : `${status.success}개 정보원 성공 · ${status.errors}개 오류 · ${status.manual}개 수동 확인 · ${status.unknown}개 미확인`}</p><p class="ci-note">재무 실적·차입금·가이던스는 공식 IR 원문을 검토한 뒤 수동 갱신해야 합니다. 수집 성공일이 최신 공시일까지의 전체 수집을 보장하지는 않습니다.</p>${status.sources.length ? `<details><summary>정보원별 수집 상태 확인</summary><div class="ci-table-scroll"><table class="ci-table"><thead><tr><th>정보원</th><th>상태</th><th>시도일</th><th>성공일</th><th>최근 발표일</th><th>확인 내용</th></tr></thead><tbody>${status.sources.map(source => `<tr><th scope="row">${E(source.name || source.title || source.id || source.source_id || '정보원')}</th><td>${E(source.status === 'success' ? '성공' : source.status === 'error' ? '오류' : source.status === 'manual' ? '수동 확인' : source.status === 'never_success' ? '수집 이력 없음' : '미확인')}</td><td>${E(dateLabel(source.last_attempt_at))}</td><td>${E(dateLabel(source.last_success_at))}</td><td>${E(dateLabel(source.latest_publication_date))}</td><td>${E(source.error || source.reason || source.note || '—')}</td></tr>`).join('')}</tbody></table></div></details>` : ''}`;
  }
  async function refreshNews() {
    state.fetchPhase = 'loading'; renderUpdateStatus();
    if (!root.fetch || root.location && root.location.protocol === 'file:') { state.fetchPhase = 'unavailable'; renderNews(); return; }
    const fetchFile = async url => {
      const response = await root.fetch(url, {cache:'no-store'});
      if (!response.ok) throw new Error('수집 파일 읽기 실패');
      return response.json();
    };
    const results = await Promise.allSettled([fetchFile('./intelligence/news.json'),fetchFile('./intelligence/update-log.json')]);
    const newsResult = results[0]; const logResult = results[1];
    state.newsFetched = false;
    if (newsResult.status === 'fulfilled' && newsResult.value && newsResult.value.schema_version === 1 && Array.isArray(newsResult.value.entries)) {
      state.news = newsResult.value.entries; state.newsFetched = true; state.newsFileSuccess = newsResult.value.last_successful_fetch;
    }
    if (logResult.status === 'fulfilled' && logResult.value && typeof logResult.value === 'object') state.log = logResult.value;
    state.fetchPhase = state.newsFetched ? 'loaded' : 'unavailable';
    // A missing status file must not make an otherwise loaded feed appear fully healthy.
    if (state.newsFetched && logResult.status === 'rejected') state.log = {...state.log,overall:{...state.log?.overall,status:'partial'}};
    renderNews();
  }
  function renderMarket() {
    const host = byId('ci-market'); if (!host) return;
    const items = array(researchData().market);
    host.innerHTML = `<div class="ci-panel-heading"><div><span class="ci-eyebrow">시장을 움직이는 근거</span><h3>시장 조사 · 확인된 이슈</h3></div><button type="button" id="ci-download-research" class="ci-button ci-button-quiet">전체 조사 내려받기 ↓</button></div><p class="ci-note">${E(researchData().source_policy || '공식 공시와 원문을 기준으로 조사합니다. 발표일과 자료 확인일을 구분해 확인하세요.')}</p><div class="ci-market-grid">${items.map(item => {
      const sources = array(item.sources); const direct = safeURL(item.url) ? [{title:item.source || '공식 원문',url:item.url,published_at:item.published_at || item.date,accessed_at:item.accessed_at || researchData().as_of}] : [];
      const linked = sources.length ? sources : direct;
      return `<article class="ci-market-item"><span class="ci-market-category">${E(item.category || item.topic || '시장 조사')}</span><h4>${E(item.title || item.topic || '시장 조사')}</h4><p>${E(item.summary || item.text || item.finding || UNKNOWN)}</p>${item.implication ? `<div class="ci-market-implication"><strong>검토 포인트</strong><p>${E(item.implication)}</p></div>` : ''}<div class="ci-citations">${linked.map(source => `${sourceLink(source)}<span> · 발표 ${E(source.published_at ? dateLabel(source.published_at) : '미기재')} · 조회 ${E(dateLabel(source.accessed_at))}</span>`).join('<br>') || '원문 출처 확인 필요'}</div></article>`;
    }).join('') || empty('출처가 확인된 시장 조사 자료가 아직 없습니다. 조사 완료 후 이곳에 등록합니다.')}</div>`;
    byId('ci-download-research').addEventListener('click', () => downloadJSON(researchData(),'competitor-research.json'));
  }
  function init() {
    if (!document || state.mounted) return;
    const section = byId('sec-company'); const detail = byId('coDetail'); const grid = byId('coGrid');
    if (!section || !detail || !grid) return;
    state.mounted = true;
    const overview = document.createElement('div'); overview.id = 'ci-overview'; overview.className = 'ci-overview'; section.prepend(overview);
    const toolbar = document.createElement('div'); toolbar.className = 'ci-toolbar'; toolbar.innerHTML = '<label class="ci-search-label" for="ci-company-search">기업·제품 검색<input type="search" id="ci-company-search" placeholder="회사명, 제품명, 지역을 입력하세요"></label><span id="ci-company-count" class="ci-muted"></span><div id="ci-compare-message" class="ci-compare-message" role="status" aria-live="polite"></div>';
    grid.before(toolbar);
    const comparison = document.createElement('div'); comparison.id = 'ci-comparison'; comparison.className = 'ci-panel'; comparison.hidden = true; detail.before(comparison);
    const panel = document.createElement('div'); panel.id = 'ci-company-research'; panel.className = 'ci-panel'; detail.prepend(panel);
    const market = document.createElement('div'); market.id = 'ci-market'; market.className = 'ci-panel ci-market'; section.appendChild(market);
    byId('ci-company-search').addEventListener('input', event => { state.search = event.target.value; renderGrid(); });
    state.originalSelect = typeof root.selectCo === 'function' ? root.selectCo : null;
    root.selectCo = selectCompany; root.renderCoGrid = renderGrid;
    root.renderNews = renderNews;
    // Existing initCompany/initNews retain their navigation and region/keyword event bindings.
    const newsSection = byId('sec-news');
    if (newsSection) {
      const status = document.createElement('div'); status.id = 'ci-news-status'; status.className = 'ci-panel ci-news-status'; status.setAttribute('role','status'); status.setAttribute('aria-live','polite'); newsSection.prepend(status);
      const tools = document.createElement('div'); tools.className = 'ci-toolbar ci-news-tools'; tools.innerHTML = '<label class="ci-search-label" for="ci-news-search">뉴스 검색<input type="search" id="ci-news-search" placeholder="공식 발표 제목·내용 검색"></label><label class="ci-official-check"><input id="ci-news-official" type="checkbox"> 공식 발표만 보기</label><button type="button" id="ci-refresh-news" class="ci-button ci-button-quiet">수집 파일 다시 확인 ↻</button>';
      status.after(tools);
      byId('ci-news-search').addEventListener('input', event => { state.newsSearch = event.target.value; renderNews(); });
      byId('ci-news-official').addEventListener('change', event => { state.onlyOfficial = event.target.checked; renderNews(); });
      byId('ci-refresh-news').addEventListener('click', refreshNews);
      const selector = byId('newsCoSel');
      if (selector) { selector.setAttribute('aria-label','회사별 뉴스 필터'); selector.addEventListener('change',renderNews); }
    }
    renderHeader(); renderGrid(); renderComparison(); renderMarket(); renderNews();
    const active = legacyState().selCo || companies()[0]?.id;
    if (active) selectCompany(active);
    refreshNews();
  }
  return { init, selectCompany, refreshNews, renderNews, renderGrid, getNewsSnapshot, formatMetric, escapeHTML, safeURL, normalizeNews, mergeNews, freshness, financialContext, metricDefinition, sourceLabel, updateState };
});
