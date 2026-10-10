"""Behavioral tests for official-source ingestion; all fixtures are synthetic."""
import importlib.util
import io
import json
from pathlib import Path
import tempfile
import unittest
from datetime import datetime, timezone
from unittest.mock import patch

ROOT = Path(__file__).resolve().parents[1]
SCRIPT = ROOT / 'scripts/update_intelligence.py'
NOW = datetime(2026, 10, 9, 6, tzinfo=timezone.utc)
SOURCE = {'id': 'upm', 'company': 'upm', 'name': 'UPM official', 'mode': 'rss',
          'url': 'https://example.com/feed.xml', 'allowed_hosts': ['example.com']}
RSS = b'''<rss version="2.0"><channel><title>Official</title><item><title>Paper update</title>
<link>https://example.com/release?utm_source=rss#top</link><guid>release-1</guid>
<pubDate>Thu, 08 Oct 2026 10:00:00 GMT</pubDate><description>&lt;p&gt;Published fact&lt;/p&gt;</description>
</item></channel></rss>'''
ATOM = b'''<feed xmlns="http://www.w3.org/2005/Atom"><entry><id>release-2</id><title>Result</title>
<link href="https://example.com/results"/><published>2026-10-09T01:00:00+00:00</published>
<summary>Reported result</summary></entry></feed>'''

class PipelineTests(unittest.TestCase):
    @classmethod
    def setUpClass(cls):
        if SCRIPT.exists():
            spec = importlib.util.spec_from_file_location('update_intelligence', SCRIPT)
            cls.pipeline = importlib.util.module_from_spec(spec)
            spec.loader.exec_module(cls.pipeline)

    def setUp(self):
        self.assertTrue(SCRIPT.exists(), 'Official feed ingestion is not implemented')
        self.p = self.pipeline

    def test_rss_retains_dated_attributed_official_evidence(self):
        entry = self.p.parse_feed(RSS, SOURCE, NOW)[0]
        self.assertEqual(entry['date'], '2026-10-08')
        self.assertEqual(entry['company'], 'upm')
        self.assertEqual(entry['url'], 'https://example.com/release')
        self.assertEqual(entry['summary'], 'Published fact')
        self.assertEqual(entry['evidence_type'], 'official_release')

    def test_atom_supports_namespaces_and_published_date(self):
        entry = self.p.parse_feed(ATOM, SOURCE, NOW)[0]
        self.assertEqual(entry['title'], 'Result')
        self.assertEqual(entry['date'], '2026-10-09')

    def test_unsafe_urls_and_external_article_hosts_are_rejected(self):
        for url in ('javascript:alert(1)', 'file:///secret', 'http://example.com/a',
                    'https://other.com/a', 'https://user:pass@example.com/a', 'https://example.com:444/a'):
            body = RSS.replace(b'https://example.com/release?utm_source=rss#top', url.encode())
            self.assertEqual(self.p.parse_feed(body, SOURCE, NOW), [])

    def test_future_and_undated_items_are_not_collected(self):
        self.assertEqual(self.p.parse_feed(RSS.replace(b'08 Oct 2026', b'10 Oct 2026'), SOURCE, NOW), [])
        self.assertEqual(self.p.parse_feed(RSS.replace(b'<pubDate>Thu, 08 Oct 2026 10:00:00 GMT</pubDate>', b''), SOURCE, NOW), [])

    def test_duplicate_urls_and_source_ids_collapse(self):
        first = self.p.parse_feed(RSS, SOURCE, NOW)[0]
        second = dict(first, url='https://example.com/release?utm_medium=email', id='second')
        third = dict(first, url='https://example.com/renamed')
        self.assertEqual(len(self.p.merge_entries([first], [second, third])), 1)

    def test_malformed_and_entity_feeds_fail_closed(self):
        for body in (b'<rss><broken>', b'<!DOCTYPE rss [<!ENTITY x "boom">]><rss/>', b'<html/>',
                     b'\xff\xfe' + '<!DOCTYPE rss><rss/>'.encode('utf-16-le')):
            with self.assertRaises(ValueError):
                self.p.parse_feed(body, SOURCE, NOW)

    def test_parser_enforces_response_size(self):
        with self.assertRaises(ValueError):
            self.p.parse_feed(b'x' * (self.p.MAX_FEED_BYTES + 1), SOURCE, NOW)

    def test_reader_enforces_response_size_before_parse(self):
        response = io.BytesIO(b'x' * (self.p.MAX_FEED_BYTES + 1))
        with self.assertRaises(ValueError):
            self.p.read_response(response)

    def test_reader_enforces_deadline_even_when_source_keeps_sending(self):
        with patch.object(self.p.time, 'monotonic', side_effect=[0, self.p.TIMEOUT_SECONDS + 1]):
            with self.assertRaises(TimeoutError):
                self.p.read_response(io.BytesIO(RSS))

    def test_html_date_uses_first_publication_date_not_later_reporting_period(self):
        source = dict(SOURCE, mode='html', release_path_pattern=r'^/release/')
        html = b'<html><li><a href="/release/report.pdf">August 07, 2026 Financial results for 2026-06-30</a></li></html>'
        entry = self.p.parse_newsroom(html, source, NOW)[0]
        self.assertEqual(entry['date'], '2026-08-07')

    def test_failure_preserves_entries_and_last_success(self):
        old = self.p.parse_feed(RSS, SOURCE, NOW)[0]
        news = {'schema_version': 1, 'entries': [old], 'last_successful_fetch': '2026-10-08T06:00:00Z'}
        log = {'sources': [{'id': 'upm', 'last_success_at': '2026-10-08T06:00:00Z'}]}
        def failing_fetch(source):
            raise OSError('offline')
        updated, status = self.p.collect({'sources': [SOURCE]}, news, log, NOW, failing_fetch)
        self.assertEqual(updated['entries'], [old])
        self.assertEqual(updated['last_successful_fetch'], news['last_successful_fetch'])
        self.assertEqual(status['sources'][0]['last_success_at'], '2026-10-08T06:00:00Z')
        self.assertEqual(status['sources'][0]['last_attempt_at'], '2026-10-09T06:00:00Z')
        self.assertEqual(status['status'], 'error')

    def test_success_distinguishes_collection_and_publication_freshness(self):
        news, log = self.p.collect({'sources': [SOURCE]}, {}, {}, NOW, lambda source: RSS)
        self.assertEqual(news['last_successful_fetch'], '2026-10-09T06:00:00Z')
        self.assertEqual(log['sources'][0]['latest_publication_date'], '2026-10-08')
        self.assertEqual(log['sources'][0]['status'], 'success')

    def test_partial_failure_does_not_discard_successful_sources(self):
        other = dict(SOURCE, id='other', company='other')
        def fetch(source):
            if source['id'] == 'other':
                raise ValueError('malformed feed')
            return RSS
        news, log = self.p.collect({'sources': [SOURCE, other]}, {}, {}, NOW, fetch)
        self.assertEqual(len(news['entries']), 1)
        self.assertEqual(log['status'], 'partial')

    def test_manual_review_never_claims_a_fetch_success(self):
        source = dict(SOURCE, mode='manual', reason='No verified feed')
        news, log = self.p.collect({'sources': [source]}, {}, {}, NOW, lambda source: self.fail('manual fetched'))
        self.assertIsNone(news['last_successful_fetch'])
        self.assertEqual(log['sources'][0]['status'], 'manual')
        self.assertIsNone(log['sources'][0]['last_attempt_at'])

    def test_unregistered_feed_host_is_rejected_before_network(self):
        with self.assertRaises(ValueError):
            self.p.fetch_feed(dict(SOURCE, url='https://other.com/feed'))

    def test_html_newsroom_only_uses_dated_registered_release_links(self):
        source = dict(SOURCE, mode='html', release_path_pattern=r'^/release/')
        html = b'<html><li><time>8 October 2026</time><a href="/release/one">Official update</a></li><li><a href="/release/two">Undated</a></li><li><time>10 October 2026</time><a href="/release/future">Future</a></li></html>'
        entries = self.p.parse_newsroom(html, source, NOW)
        self.assertEqual([(x['title'], x['date']) for x in entries], [('Official update', '2026-10-08')])
        self.assertEqual(entries[0]['summary'], '')

    def test_html_layout_change_fails_instead_of_claiming_success(self):
        source = dict(SOURCE, mode='html', release_path_pattern=r'^/release/')
        with self.assertRaises(ValueError):
            self.p.parse_newsroom(b'<html><p>New unsupported layout</p></html>', source, NOW)

    def test_html_date_inside_anchor_and_pdf_links_are_supported(self):
        source = dict(SOURCE, mode='html', release_path_pattern=r'^/release/')
        html = b'<html><ul><li><a href="/release/report.pdf">August 07, 2026 Consolidated Financial Results [121 KB]</a></li></ul></html>'
        entry = self.p.parse_newsroom(html, source, NOW)[0]
        self.assertEqual(entry['date'], '2026-08-07')
        self.assertIn('Financial Results', entry['title'])

    def test_atomic_write_creates_readable_json(self):
        with tempfile.TemporaryDirectory() as directory:
            target = Path(directory) / 'news.json'
            self.p.atomic_write_json(target, {'entries': []})
            self.assertEqual(json.loads(target.read_text()), {'entries': []})
            self.assertEqual(list(Path(directory).iterdir()), [target])

    def test_redirect_is_validated_before_following(self):
        handler = self.p.OfficialRedirectHandler(['example.com'])
        from urllib.request import Request
        with self.assertRaises(ValueError):
            handler.redirect_request(Request(SOURCE['url']), None, 302, 'found', {}, 'http://example.com/feed')

if __name__ == '__main__':
    unittest.main()
