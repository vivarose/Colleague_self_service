"""Download and validate all public sources before replacing the published catalog."""

import argparse
from datetime import datetime, timedelta, timezone
from hashlib import sha256
import io
import json
import os
from pathlib import Path
import re
import shutil
import subprocess
import sys
import tempfile
import time
from urllib.parse import urljoin, urlparse
from urllib.request import Request, urlopen

from bs4 import BeautifulSoup
import pandas as pd
from pypdf import PdfReader

from course_info_to_html import CATALOG_TERM, CATALOG_TERM_LABEL
from seat_allocations import parse_allocations


ROOT = Path(__file__).resolve().parent
CATALOG = 'https://hamilton.smartcatalogiq.com'
SCHEDULE = 'https://www.hamilton.edu/admission/visiting/courses/printable-schedule?term=27%2FSP'
SEATS = 'https://www.hamilton.edu/offices/registrar/registration/class-year-seats'
SITEMAP = f'{CATALOG}/en/current/college-catalogue/site-map'
EST = timezone(timedelta(hours=-5))


def download(url, sources):
    with urlopen(Request(url, headers={'User-Agent': 'Hamilton-Course-Explorer/1.0'}), timeout=60) as response:
        body = response.read()
        if not body:
            raise ValueError(f'Empty source: {url}')
        sources.append({
            'url': url, 'resolved_url': response.url,
            'retrieved_at': datetime.now(timezone.utc).isoformat(),
            'sha256': sha256(body).hexdigest(),
        })
    return body


def lines(cell):
    return [line.strip() for line in cell.get_text('\n').splitlines() if line.strip()]


def parse_schedule(content):
    soup = BeautifulSoup(content, 'html.parser')
    if not any('27/SP' in heading.get_text() for heading in soup.find_all('h3')):
        raise ValueError('Schedule does not identify the configured Spring 2027 term.')
    table = next((table for table in soup.find_all('table')
                  if [cell.get_text(strip=True) for cell in table.find_all('th')] ==
                  ['Course', 'Title', 'Day', 'Time', 'Room', 'Credit', 'Instructor', 'Method', 'Max Cap', 'Restr']), None)
    if table is None:
        raise ValueError('Schedule table headers have changed.')
    rows = []
    seen = set()
    for tr in table.find_all('tr'):
        cells = tr.find_all('td', recursive=False)
        if not cells:
            continue
        if len(cells) != 10:
            raise ValueError('Schedule row does not contain 10 columns.')
        code = cells[0].get_text(strip=True)
        match = re.fullmatch(r'([A-Z]+)\s*(\d+[A-Z]*)\s*(\d{2})', code)
        if not match:
            raise ValueError(f'Invalid schedule section: {code}')
        subject, number, section = match.groups()
        key = subject, number, section
        if key in seen:
            raise ValueError(f'Duplicate schedule section: {key}')
        seen.add(key)
        title_lines = lines(cells[1])
        row = {
            'COURSE_NO': code, 'TITLE': title_lines[0],
            'COMMENTS': ' '.join(title_lines[1:]).removeprefix('NOTE:').strip(),
            'CREDIT': cells[5].get_text(strip=True),
            'TEACHINGMETHOD': cells[7].get_text(strip=True),
            'MAXCAP': cells[8].get_text(strip=True),
            'RESTRICTED': cells[9].get_text(strip=True),
        }
        days, times, rooms, instructors = (lines(cells[index]) for index in (2, 3, 4, 6))
        if max(map(len, (days, times, rooms, instructors))) > 4:
            raise ValueError(f'More than four meeting slots or instructors: {code}')
        if len(days) != len(times) or len(rooms) != len(times):
            raise ValueError(f'Misaligned meeting columns: {code}')
        for i in range(4):
            meeting = times[i] if i < len(times) else ''
            if meeting:
                time_match = re.fullmatch(r'(\d{2}:\d{2}[AP]M)[ -]+(\d{2}:\d{2}[AP]M)', meeting)
                if time_match:
                    meeting = f'{time_match[1]} {time_match[2]}'
                elif meeting.upper() not in ('TBA', 'ARR', 'ARRANGED'):
                    raise ValueError(f'Unrecognized meeting time for {code}: {meeting}')
            row.update({
                f'DAY{i + 1}': days[i] if i < len(days) else '',
                f'TIME{i + 1}': meeting,
                f'ROOM{i + 1}': rooms[i] if i < len(rooms) else '',
                f'INSTRUCTOR{i + 1}': instructors[i] if i < len(instructors) else '',
            })
        rows.append(row)
    if not rows:
        raise ValueError('Schedule is empty.')
    return rows


def catalog_keys(label):
    match = re.fullmatch(r'([A-Z]+)[ -]+(\d+[A-Z]*(?:/\d+[A-Z]*)*)', label.strip().upper())
    return [(match[1], number) for number in match[2].split('/')] if match else []


def parse_course(content, expected_key, url):
    main = BeautifulSoup(content, 'html.parser').select_one('#main')
    if main is None or main.h1 is None or main.select_one('.desc') is None:
        raise ValueError(f'Catalog course markup is incomplete: {url}')
    identity = main.h1.find('span')
    if identity is None or expected_key not in catalog_keys(identity.get_text(' ', strip=True)):
        raise ValueError(f'Catalog identity mismatch for {expected_key}: {url}')
    prereqs = main.select_one('.sc_prereqs')
    coreqs = main.select_one('.sc_coreqs')
    if prereqs is None or coreqs is None:
        raise ValueError(f'Catalog prerequisite/corequisite containers missing: {url}')
    title = main.h1.get_text(' ', strip=True).removeprefix(identity.get_text(' ', strip=True)).strip()
    def field(name):
        heading = next((h for h in main.find_all(['h2', 'h3'])
                        if h.get_text(' ', strip=True) == name), None)
        if heading is None:
            return ''
        parts = []
        for sibling in heading.next_siblings:
            if getattr(sibling, 'name', None) in ('h1', 'h2', 'h3', 'div'):
                break
            parts.append(sibling.get_text(' ', strip=True) if hasattr(sibling, 'get_text') else str(sibling))
        return ' '.join(' '.join(parts).split())
    return {
        'Subject Code': expected_key[0], 'Course Number': expected_key[1],
        'Course Name': title, 'Course Description': main.select_one('.desc').decode_contents(),
        'Prerequisite Narrative': re.sub(r'^Prerequisites?\s*:?\s*', '', prereqs.get_text(' ', strip=True)),
        'Corequisite Narrative': re.sub(r'^Corequisites?\s*:?\s*', '', coreqs.get_text(' ', strip=True)),
        'Cross Listed Courses Narrative': field('Cross Listed Courses'),
        'Notes': field('Notes'), 'Catalogue URL': url,
    }


def fetch_catalog(schedule, sources):
    soup = BeautifulSoup(download(SITEMAP, sources), 'html.parser')
    listings = sorted({
        urljoin(CATALOG, a['href']) for a in soup.select('a[href]')
        if '/academicprograms/' in a['href'] and 'courses' in a['href'].rsplit('/', 1)[-1]
    })
    if len(listings) < 20:
        raise ValueError('Catalog discovery is unexpectedly incomplete.')
    wanted = set()
    for row in schedule:
        match = re.fullmatch(r'([A-Z]+)\s*(\d+[A-Z]*)\s*(\d{2})', row['COURSE_NO'])
        wanted.add((match[1], match[2].removesuffix('W')))
    locations = {}
    for listing in listings:
        main = BeautifulSoup(download(listing, sources), 'html.parser').select_one('#main')
        if main is None:
            raise ValueError(f'Catalog listing missing: {listing}')
        for anchor in main.select('a.sc-courselink[href]'):
            url = urljoin(CATALOG, anchor['href'])
            if urlparse(url).hostname != urlparse(CATALOG).hostname:
                raise ValueError(f'Unexpected catalog host: {url}')
            for key in catalog_keys(anchor.get_text(' ', strip=True)):
                if key in wanted:
                    if key in locations and locations[key] != url:
                        raise ValueError(f'Ambiguous catalog identity: {key}')
                    locations[key] = url
        time.sleep(0.1)
    missing = wanted - locations.keys()
    print(f'Schedule course identities without catalog entries: {sorted(missing)}')
    if len(locations) < len(wanted) * 0.6:
        raise ValueError('Less than 60% of schedule courses have discoverable catalog entries.')
    records = []
    cached = {}
    for key, url in sorted(locations.items()):
        if url not in cached:
            cached[url] = download(url, sources)
            time.sleep(0.1)
        records.append(parse_course(cached[url], key, url))
    for key in sorted(missing):
        records.append({
            'Subject Code': key[0], 'Course Number': key[1],
            'Notes': 'No matching public catalog entry was found during this refresh.'
        })
    return records


def refresh(output, refresh_date):
    if refresh_date != datetime.now(EST).date().isoformat():
        raise ValueError('Refresh date must be today in fixed EST (UTC-05:00).')
    if CATALOG_TERM != '2027-spring' or CATALOG_TERM_LABEL != 'Spring 2027':
        raise ValueError('Update source URLs and validation when changing the configured term.')
    sources = []
    schedule = parse_schedule(download(SCHEDULE, sources))
    previous = pd.read_excel(output / 'merged_course_schedule.xlsx')
    previous_count = len(previous.drop_duplicates(['Subject Code', 'Course Number+', 'Section']))
    if len(schedule) < previous_count * 0.9:
        raise ValueError(f'Schedule shrank by over 10% ({previous_count} to {len(schedule)}); manual review required.')
    pdf = download(SEATS, sources)
    if not pdf.startswith(b'%PDF') or not re.search(r'/27SP(?:%20| )', sources[-1]['resolved_url']):
        raise ValueError('Seat allocation source is not a Spring 2027 PDF.')
    reader = PdfReader(io.BytesIO(pdf))
    allocation_text = '\n'.join(page.extract_text() for page in reader.pages)
    allocations = parse_allocations(allocation_text)
    if len(allocations) < 200:
        raise ValueError('Seat allocation feed has fewer than 200 rows; manual review required.')
    catalog = fetch_catalog(schedule, sources)
    with tempfile.TemporaryDirectory() as temporary:
        staging = Path(temporary)
        pd.DataFrame(schedule).to_excel(staging / 'courseschedule.xlsx', index=False)
        pd.DataFrame(catalog).to_csv(staging / 'courseExport.csv', index=False)
        (staging / 'seats.pdf').write_bytes(pdf)
        environment = {**os.environ, 'CATALOG_INPUT_DIR': str(staging),
                       'CATALOG_SEATS_FILE': str(staging / 'seats.pdf')}
        subprocess.run([sys.executable, str(ROOT / 'join-catalogue.py')], env=environment, check=True)
        merged = pd.read_excel(staging / 'merged_course_schedule.xlsx')
        if len(merged) != len(schedule):
            raise ValueError('Merging sources changed the schedule row count.')
        for name in ('course_info_to_html.py', 'catalogue_snapshot.py', 'htmltemplate.html', 'filter.js'):
            shutil.copy2(ROOT / name, staging / name)
        if (output / 'catalog-snapshots').exists():
            shutil.copytree(output / 'catalog-snapshots', staging / 'catalog-snapshots')
        provenance = {
            'refresh_date': refresh_date, 'timezone': 'UTC-05:00',
            'source_retrieved_at': datetime.now(timezone.utc).isoformat(), 'sources': sources
        }
        (staging / 'source-provenance.json').write_text(json.dumps(provenance, indent=2) + '\n')
        subprocess.run([sys.executable, str(staging / 'course_info_to_html.py')], check=True)
        # All network, parsing, and build checks finish before published files change.
        for name in ('merged_course_schedule.xlsx', 'courses.html', 'source-provenance.json'):
            shutil.copy2(staging / name, output / name)
        shutil.copytree(staging / 'catalog-snapshots', output / 'catalog-snapshots', dirs_exist_ok=True)


if __name__ == '__main__':
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--date', required=True, help='Fixed-EST request date, YYYY-MM-DD')
    parser.add_argument('--output', type=Path, default=ROOT)
    args = parser.parse_args()
    refresh(args.output.resolve(), args.date)
