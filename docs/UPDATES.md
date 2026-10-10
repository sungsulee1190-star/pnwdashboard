# Official competitor intelligence updates

The dashboard reads `intelligence/news.json` and `intelligence/update-log.json`. The existing `DATA_NEWS` remains the offline fallback. The collector uses Python 3.10+ standard-library modules only; no dependency installation or API credential is required.

```sh
python3 tests/update-pipeline.test.py
python3 scripts/update_intelligence.py
```

The source registry in `config/intelligence-sources.json` maps all 14 dashboard company IDs. UPM, Oji Holdings and Hokuetsu currently have registered public official HTML newsroom adapters. Other companies are explicitly marked for manual IR review. Sappi's newsroom returned HTTP 403 during verification; Nippon's dynamic newsroom exposed no records without JavaScript; Arctic's site timed out. No guessed RSS URL, paid publication, private social account, or paywall bypass is used. RSS/Atom ingestion is supported for future independently verified official feeds.

The official pages and dated links were inspected through web browsing on 2026-10-09: [UPM releases](https://www.upm.com/news-and-stories/releases/), [Oji releases](https://www.ojiholdings.co.jp/en/news/release/) and [Hokuetsu releases](https://www.hokuetsucorp.com/en/newsrelease/). Runtime outbound requests could not be exercised in the implementation workspace because its network proxy was unavailable. HTML extraction was tested using representative local fixtures; validate the first network-enabled collection log and links before merging generated data. A source changing its HTML layout produces an error, preserves its previous entries, and requires a reviewed adapter update.

## Scheduling and activation

`.github/workflows/intelligence-update.yml` runs at 06:00 UTC daily and supports **Actions → Review official competitor intelligence → Run workflow**. GitHub scheduled workflows activate only after this workflow is merged into the repository's default branch; a draft implementation PR does not activate a schedule. GitHub may delay scheduled runs.

The collection job has read-only repository access. A separate job has permission to push a new `intelligence/update-*` branch and open a **draft review PR** containing only the two JSON snapshots. A reviewer must merge it before the deployed default-branch dashboard changes. There is no automatic merge or publication. Existing repository deployment behavior after a merge still applies. Repository settings must allow Actions to create pull requests, and branch policies must allow the review branch. If PR creation is blocked by repository policy, the snapshot remains available as a seven-day workflow artifact.

Every run can produce a review PR because collection status timestamps change even when no new announcements arrive. Resolve or close earlier update PRs before merging a newer snapshot to avoid overwriting newer evidence with an older snapshot.

## Evidence and failure semantics

Announcements contain `company`, `source`, `source_id`, `upstream_id`, `url`, `date`, `title`, `summary`, `evidence_type`, and `collected_at`. HTML adapters keep publisher titles, links and publication dates, with an empty summary. Feed summaries, if present, are publisher-supplied plain-text excerpts. The pipeline does not invent news or generate editorial summaries. A PDF announcement is an evidence link; its contents are not parsed into financial metrics.

Only HTTPS hosts registered for that source are accepted, including before following redirects. Responses are limited to 2 MiB and timed out; unsafe XML DTD/entity declarations, malformed feeds, unknown layouts, undated items and future publication dates are rejected. URLs lose tracking parameters and fragments before deduplication. Both canonical URLs and source-scoped publisher IDs are deduplicated. The snapshot retains up to 2,000 announcements, newest publication date first.

Each output file is replaced atomically. Failed sources preserve prior entries and their prior `last_success_at`. `last_attempt_at` says when a download was attempted, not when evidence was successfully collected. `latest_publication_date` reports source publication freshness, independently of fetch freshness. An old announcement retrieved today remains an old announcement. Manual sources have no download attempt or successful-fetch timestamp. Overall status is `success`, `partial`, `error`, or `manual`; an all-source error exits nonzero after saving diagnostic JSON. Partial failures are visible in the review log and do not discard successful sources.

The collector does not overwrite an unreadable previous JSON file: it stops before writing. Restore or repair the reviewed snapshot and rerun. Never replace a corrupt snapshot with an empty file to hide a failure.

## Audited business and financial facts

Scheduled collection currently refreshes announcement evidence only. Financial series, capacity, mill lists, product ranges, corporate entities, and forecast classifications require manual review of primary filings. Keep missing values missing; never extrapolate a forecast or treat publication date as the reporting period.

At least quarterly, and whenever a newly collected announcement identifies a closure, investment, transaction or result:

1. Open the company's official IR or exchange filing; for Korea use DART and check each legal entity separately. Moorim Paper and Moorim P&P are different issuers. For China check the company's matching exchange ticker; for Japan note March fiscal year ends; for Sappi check September fiscal year ends.
2. Confirm the reporting period, publication date, legal entity, original currency, unit, accounting basis and consolidated versus segment scope. A cumulative interim number is not a standalone quarter. Distinguish annual actuals from management forecasts.
3. Update only verified fields in `research/asia.json` or `research/global.json`, retaining the supporting document URL, publication date, period, evidence type and any limitation. Review the schema against existing research entries, then run `node research/build.cjs` to regenerate `intelligence_data.js`. Do not alter `data.js` operational history through this collector.
4. Run `npm test` and this pipeline suite, review the displayed company and financial views, then submit the research change in a separate review PR.

OpenDART API integration is not enabled. A future implementation would need a repository secret for its API key, explicit verified corporate-code mappings, report/period selection, unit/scope validation and meaningful filing-ingestion tests. Do not put credentials in source files or add a secret merely to imply financial automation.
