#!/usr/bin/env python3
"""Collect attributed official announcements. Never infer financial metrics."""
import argparse
from datetime import datetime, timezone
from email.utils import parsedate_to_datetime
import hashlib
from html import unescape
from html.parser import HTMLParser
import json
import os
from pathlib import Path
import re
import tempfile
import time
from urllib.parse import parse_qsl, urlencode, urljoin, urlsplit, urlunsplit
from urllib.request import HTTPRedirectHandler, Request, build_opener
import xml.etree.ElementTree as ET

MAX_FEED_BYTES = 2 * 1024 * 1024
TIMEOUT_SECONDS = 15
MAX_ENTRIES = 2000
ROOT = Path(__file__).resolve().parents[1]


def timestamp(now):
    return now.astimezone(timezone.utc).isoformat(timespec='seconds').replace('+00:00', 'Z')


def canonical_url(url, allowed_hosts=None):
    parts = urlsplit(url.strip())
    if (parts.scheme.lower() != 'https' or not parts.hostname or parts.username or parts.password
            or parts.port not in (None, 443) or any(ord(c) < 33 for c in url)):
        raise ValueError('Only public HTTPS links without credentials are accepted')
    host = parts.hostname.lower()
    if allowed_hosts is not None and host not in [h.lower() for h in allowed_hosts]:
        raise ValueError('URL host is not registered for this source')
    params = [(key, value) for key, value in parse_qsl(parts.query, keep_blank_values=True)
              if not key.lower().startswith('utm_') and key.lower() not in {'gclid', 'fbclid'}]
    return urlunsplit(('https', host, parts.path or '/', urlencode(sorted(params)), ''))


class OfficialRedirectHandler(HTTPRedirectHandler):
    def __init__(self, allowed_hosts):
        super().__init__()
        self.allowed_hosts = allowed_hosts

    def redirect_request(self, req, fp, code, msg, headers, newurl):
        canonical_url(newurl, self.allowed_hosts)
        return super().redirect_request(req, fp, code, msg, headers, newurl)


def read_response(response):
    chunks, size = [], 0
    # HTTPResponse.read1 returns available bytes instead of waiting to fill a chunk.
    # This lets the wall-clock deadline stop a slow, continuously streaming source.
    read_chunk = getattr(response, 'read1', response.read)
    deadline = time.monotonic() + TIMEOUT_SECONDS
    while True:
        if time.monotonic() > deadline:
            raise TimeoutError('Source response exceeded time limit')
        chunk = read_chunk(min(65536, MAX_FEED_BYTES + 1 - size))
        if not chunk:
            return b''.join(chunks)
        size += len(chunk)
        if size > MAX_FEED_BYTES:
            raise ValueError('Source response exceeds 2 MiB limit')
        chunks.append(chunk)


def fetch_feed(source):
    # Validate before constructing requests, including redirect destinations.
    canonical_url(source['url'], source['allowed_hosts'])
    opener = build_opener(OfficialRedirectHandler(source['allowed_hosts']))
    request = Request(source['url'], headers={'User-Agent': 'PaperIntelligenceDashboard/1.0',
                                            'Accept': 'application/rss+xml, application/atom+xml, text/html, application/xml',
                                            'Accept-Encoding': 'identity'})
    with opener.open(request, timeout=TIMEOUT_SECONDS) as response:
        canonical_url(response.geturl(), source['allowed_hosts'])
        length = response.headers.get('Content-Length')
        if length and int(length) > MAX_FEED_BYTES:
            raise ValueError('Source response exceeds 2 MiB limit')
        return read_response(response)


def publication_date(value, now):
    try:
        parsed = datetime.fromisoformat(value.strip().replace('Z', '+00:00'))
    except ValueError:
        try:
            parsed = parsedate_to_datetime(value.strip())
        except (ValueError, TypeError):
            return None
    if parsed.tzinfo is None:
        parsed = parsed.replace(tzinfo=timezone.utc)
    # Compare publication's stated calendar date; dates later than today are excluded.
    return parsed.date().isoformat() if parsed.date() <= now.date() else None


class PlainText(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.hidden = 0

    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style'):
            self.hidden += 1

    def handle_endtag(self, tag):
        if tag in ('script', 'style') and self.hidden:
            self.hidden -= 1
        if tag in ('p', 'br', 'div'):
            self.parts.append(' ')

    def handle_data(self, data):
        if not self.hidden:
            self.parts.append(data)


def plain_text(value, limit):
    parser = PlainText()
    parser.feed(value or '')
    return re.sub(r'\s+', ' ', unescape(' '.join(parser.parts))).strip()[:limit]


def make_entry(source, url, title, date, now, upstream_id='', summary=''):
    return {'id': hashlib.sha256(url.encode()).hexdigest()[:24], 'company': source['company'],
            'date': date, 'title': title, 'summary': summary, 'url': url,
            'source': source['name'], 'source_id': source['id'], 'upstream_id': upstream_id or url,
            'category': 'official', 'evidence_type': 'official_release', 'collected_at': timestamp(now)}


def parse_feed(body, source, now):
    if len(body) > MAX_FEED_BYTES:
        raise ValueError('Feed exceeds size limit')
    # Reject DTD/entity declarations even in UTF-16/32 XML before the stdlib parser sees them.
    inspected = body.replace(b'\x00', b'').upper()
    if b'<!DOCTYPE' in inspected or b'<!ENTITY' in inspected:
        raise ValueError('XML DTD and entity declarations are prohibited')
    try:
        root = ET.fromstring(body)
    except ET.ParseError as error:
        raise ValueError('Malformed XML feed') from error
    kind = root.tag.rsplit('}', 1)[-1]
    if kind not in ('rss', 'feed', 'RDF'):
        raise ValueError('Response is not RSS or Atom')
    entries = []
    for node in root.iter():
        if node.tag.rsplit('}', 1)[-1] not in ('item', 'entry'):
            continue
        values, link = {}, ''
        for child in node:
            key = child.tag.rsplit('}', 1)[-1]
            if key == 'link' and child.attrib.get('href') and child.attrib.get('rel', 'alternate') == 'alternate':
                link = child.attrib['href']
            else:
                values[key] = ''.join(child.itertext()).strip()
        link = link or values.get('link', '')
        date = publication_date(values.get('published') or values.get('pubDate') or values.get('date') or values.get('updated', ''), now)
        title = plain_text(values.get('title', ''), 500)
        if not date or not title:
            continue
        try:
            url = canonical_url(link, source['allowed_hosts'])
        except ValueError:
            continue
        # RSS descriptions remain explicitly publisher supplied and plain text, never inferred.
        summary = plain_text(values.get('summary') or values.get('description', ''), 300)
        entries.append(make_entry(source, url, title, date, now, values.get('guid') or values.get('id'), summary))
    return merge_entries([], entries)


DATE_PATTERNS = [
    (r'\b(\d{4})[-/](\d{1,2})[-/](\d{1,2})\b', ('%Y-%m-%d', '%Y/%m/%d')),
    (r'\b\d{1,2}\s+[A-Za-z]+\s+\d{4}\b', ('%d %B %Y', '%d %b %Y')),
    (r'\b[A-Za-z]+\s+\d{1,2},?\s+\d{4}\b', ('%B %d, %Y', '%b %d, %Y', '%B %d %Y', '%b %d %Y')),
]


def newsroom_date(text, now):
    candidates = [(match.start(), match, formats) for pattern, formats in DATE_PATTERNS
                  for match in re.finditer(pattern, text)]
    # Publication dates precede report-period dates on registered release cards.
    for _, match, formats in sorted(candidates, key=lambda item: item[0]):
        for fmt in formats:
            try:
                date = datetime.strptime(match.group(), fmt).date()
                return (date.isoformat() if date <= now.date() else None), match.span()
            except ValueError:
                pass
    return None, None


class NewsroomParser(HTMLParser):
    """Pair a date with exactly one allowed release link in its nearest HTML block."""
    BLOCKS = {'li', 'article', 'div', 'tr', 'dd', 'section', 'a'}

    def __init__(self, source, now):
        super().__init__(convert_charrefs=True)
        self.source, self.now = source, now
        self.frames, self.entries = [], []
        self.hidden = 0
        self.active_anchor = None
        self.pattern = re.compile(source['release_path_pattern'])

    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style'):
            self.hidden += 1
        if self.hidden:
            return
        if tag in self.BLOCKS:
            self.frames.append({'tag': tag, 'text': [], 'links': []})
        if tag == 'a':
            raw = dict(attrs).get('href', '')
            try:
                url = canonical_url(urljoin(self.source['url'], raw), self.source['allowed_hosts'])
            except ValueError:
                return
            if self.pattern.search(urlsplit(url).path):
                self.active_anchor = {'url': url, 'text': []}
                for frame in self.frames:
                    frame['links'].append(self.active_anchor)

    def handle_data(self, data):
        if self.hidden:
            return
        for frame in self.frames:
            frame['text'].append(data)
        if self.active_anchor:
            self.active_anchor['text'].append(data)

    def handle_endtag(self, tag):
        if tag in ('script', 'style') and self.hidden:
            self.hidden -= 1
            return
        if self.hidden:
            return
        if tag == 'a':
            self.active_anchor = None
        if tag not in self.BLOCKS:
            return
        index = next((i for i in range(len(self.frames) - 1, -1, -1) if self.frames[i]['tag'] == tag), None)
        if index is None:
            return
        closing = self.frames[index:]
        del self.frames[index:]
        for frame in closing:
            if len(frame['links']) != 1:
                continue
            link = frame['links'][0]
            date, _ = newsroom_date(' '.join(frame['text']), self.now)
            title = plain_text(' '.join(link['text']), 500)
            if date and title:
                self.entries.append(make_entry(self.source, link['url'], title, date, self.now))


def parse_newsroom(body, source, now):
    if len(body) > MAX_FEED_BYTES:
        raise ValueError('Newsroom exceeds size limit')
    parser = NewsroomParser(source, now)
    parser.feed(body.decode('utf-8-sig', errors='strict'))
    parser.close()
    entries = merge_entries([], parser.entries)
    if not entries:
        raise ValueError('No dated registered release links; newsroom layout needs review')
    return entries


def merge_entries(existing, incoming):
    result, seen_urls, seen_ids = [], set(), set()
    for item in list(existing) + list(incoming):
        try:
            url = canonical_url(item['url'])
        except (ValueError, KeyError):
            continue
        key = (item.get('source_id'), item.get('upstream_id') or item.get('id') or url)
        duplicate = url in seen_urls or key in seen_ids
        seen_urls.add(url)
        seen_ids.add(key)
        if not duplicate:
            result.append(dict(item, url=url))
    return sorted(result, key=lambda entry: (entry.get('date', ''), entry.get('id', '')), reverse=True)[:MAX_ENTRIES]


def collect(config, previous_news, previous_log, now, fetch=fetch_feed):
    news = {'schema_version': 1, 'last_successful_fetch': previous_news.get('last_successful_fetch'),
            'entries': previous_news.get('entries', [])}
    previous_statuses = {source['id']: source for source in previous_log.get('sources', [])}
    statuses, incoming, successes, failures = [], [], 0, 0
    for source in config['sources']:
        prior = previous_statuses.get(source['id'], {})
        state = {'id': source['id'], 'company': source['company'], 'name': source['name'],
                 'mode': source['mode'], 'url': source['url'], 'status': 'manual',
                 'last_attempt_at': None, 'last_success_at': prior.get('last_success_at'),
                 'latest_publication_date': prior.get('latest_publication_date'), 'error': None}
        if source['mode'] == 'manual':
            state['reason'] = source.get('reason', 'Manual official IR review required')
        else:
            state['last_attempt_at'] = timestamp(now)
            try:
                body = fetch(source)
                parser = parse_newsroom if source['mode'] == 'html' else parse_feed
                if source['mode'] not in ('html', 'rss', 'atom'):
                    raise ValueError('Unregistered source mode')
                entries = parser(body, source, now)
                incoming.extend(entries)
                state.update(status='success', last_success_at=timestamp(now), collected_count=len(entries))
                if entries:
                    state['latest_publication_date'] = max(entry['date'] for entry in entries)
                successes += 1
            except (OSError, ValueError, KeyError, TimeoutError) as error:
                state.update(status='error', error=f'{type(error).__name__}: {str(error)[:200]}')
                failures += 1
        statuses.append(state)
    # Existing snapshots survive errors. Never overwrite the successful timestamp on failed runs.
    if successes:
        news['entries'] = merge_entries(news['entries'], incoming)
        news['last_successful_fetch'] = timestamp(now)
    log = {'schema_version': 1, 'attempted_at': timestamp(now),
           'last_successful_fetch': news['last_successful_fetch'],
           'status': 'partial' if successes and failures else 'error' if failures else 'success' if successes else 'manual',
           'sources': statuses}
    return news, log


def atomic_write_json(path, data):
    path = Path(path)
    path.parent.mkdir(parents=True, exist_ok=True)
    temporary = None
    try:
        with tempfile.NamedTemporaryFile(mode='w', encoding='utf-8', dir=path.parent, delete=False) as file:
            temporary = file.name
            json.dump(data, file, ensure_ascii=False, indent=2)
            file.write('\n')
            file.flush()
            os.fsync(file.fileno())
        os.replace(temporary, path)
    finally:
        if temporary and os.path.exists(temporary):
            os.unlink(temporary)


def read_json(path, default):
    return json.loads(path.read_text(encoding='utf-8')) if path.exists() else default


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--config', type=Path, default=ROOT / 'config/intelligence-sources.json')
    parser.add_argument('--output-dir', type=Path, default=ROOT / 'intelligence')
    args = parser.parse_args()
    now = datetime.now(timezone.utc)
    news_path, log_path = args.output_dir / 'news.json', args.output_dir / 'update-log.json'
    config = read_json(args.config, {})
    news, log = collect(config, read_json(news_path, {}), read_json(log_path, {}), now)
    atomic_write_json(news_path, news)
    atomic_write_json(log_path, log)
    print(json.dumps({'status': log['status'], 'entries': len(news['entries']),
                      'successful_sources': sum(s['status'] == 'success' for s in log['sources']),
                      'failed_sources': sum(s['status'] == 'error' for s in log['sources'])}))
    # Workflow always retains failure logs in its review PR, then surfaces all-source failure.
    return 1 if log['status'] == 'error' else 0


if __name__ == '__main__':
    raise SystemExit(main())
