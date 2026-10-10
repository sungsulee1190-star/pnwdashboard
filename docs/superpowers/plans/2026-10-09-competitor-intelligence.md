# Competitor Intelligence Implementation Plan

> **For agentic workers:** Use superpowers:subagent-driven-development. User explicitly selected subagents; independent files permit parallel execution.

**Goal:** Correct the existing dashboard and provide sourced competitor analysis and repeatable news updates.

**Architecture:** Preserve static HTML/JS; separate researched facts from periodic news. Layer the intelligence module over the existing navigation and company view.

**Tech Stack:** Vanilla JS, Node built-in tests, Python stdlib, GitHub Actions.

**Spec:** ../specs/2026-10-09-competitor-intelligence-design.md

## Global Constraints
- Do not invent financial figures or forecasts; show period, currency, unit, scope and sources.
- Future articles excluded, failed fetch preserves old data, no unverified HTTPS links.
- Work only on assigned files; root coordinates interfaces and remote draft PR. No changes or deployment to main.

## Review Focus
- News collection failure must not erase the last successful dataset.
- Absent period/baseline must be distinguished from real zero export.
- Country/region filters must apply throughout each selected view.
- Unsafe source text/URLs must not execute script.
- Profit definitions/periods/scopes must remain attached to metrics.

### Task 1: Dashboard corrections
Files: index.html, data.js, tests/dashboard-ui.test.cjs, tests/data-syntax.test.cjs.
Interface: retain public navigation and render functions; cleanRg delegates to DashboardLogic.cleanRegion. Intelligence enhancement loads after existing inline scripts.
- [x] Reproduce data parse error and filtered period/navigation/map correctness failures.
- [x] Fix dataset syntax without changing original statistics and implement consistent filtering/comparison/mill aggregation.
- [x] Run direct Node tests and syntax validation; review all changed paths.

### Task 2: Email and shared analysis
Files: dashboard_logic.js, tests/email-logic.test.cjs.
Interface: window.DashboardLogic retains methods, adds cleanRegion and robust selected-period exports/price summaries.
- [x] Test missing/zero/reversed periods, DATA_NEWS.items, region/country aggregation and HTML escaping.
- [x] Correct actual summary generation and missing baselines; remove unsupported claims.
- [x] Run complete email test file.

### Task 3: Official company research
Files: research/asia.json and research/global.json, independently owned.
Interface: {as_of,companies,news,market,source_policy}; company records include exact metric periods/units/scopes, source IDs, guidance/risks/products.
- [x] Verify official company/result/filing sources for each of 14 companies.
- [x] Write Korean analysis and read financial values with exact unit/scope; missing fields null.
- [x] Check schema, sourced values, dates and news claims.

### Task 4: Repeatable updates
Files: scripts/*, config/intelligence-sources.json, tests/update-pipeline.test.py, .github/workflows/intelligence-update.yml, docs/UPDATES.md.
Interface: intelligence/news.json entries use company IDs; update-log.json retains per-source attempt/success/failure.
- [x] Test RSS/Atom or verified structured official newsroom parser, dedup/future dates/unsafe URLs/error preservation.
- [x] Implement bounded official fetching and atomic writes; scheduled workflow proposes PR.
- [x] Run Python tests and workflow configuration inspection.

### Task 5: Research and news interface
Files: competitor_intelligence.js/css, tests/intelligence-ui.test.cjs.
Interface: window.COMPETITOR_RESEARCH from intelligence_data.js; external module hooks existing company/news views. Dataset financials latest first.
- [x] Test formatting/missing metrics/source safety/schema conversion/freshness/dedup.
- [x] Implement business/products/financial/debt/history/outlook/risk/source views and transparent comparison; preserve legacy company profiles.
- [x] Add source-backed market research panel and news freshness/filtering.
- [x] Run UI helper tests and syntax checks.

### Task 6: Integration and review
Files: intelligence_data.js, intelligence/*.json, README.md, root QA support.
- [x] Merge research deterministically, validate every source ID and financial unit/scope.
- [x] Add script/style tags and default seed data/news update log.
- [x] Run all Node/Python tests and syntax checks; Node VM DOM flows pass. Actual browser flow unavailable due environment socket restrictions.
- [x] Independent review of code/data/workflow; address material findings.
- [x] Prepare complete local delivery archive; create remote feature branch and draft PR preserving the base tree and original operational data following user instruction to proceed.
