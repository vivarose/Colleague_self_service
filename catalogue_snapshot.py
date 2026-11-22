from datetime import datetime, timezone
from hashlib import sha256
import json
import math
from clean_data import clean_html

SCHEMA_VERSION = 1 # matches filter.js: snapshot.schema_version
IDENTITY_FIELDS = ('Subject Code', 'Course Number+', 'Section')


def canonical_value(value):
    if value is None or isinstance(value, float) and math.isnan(value):
        return ''
    if isinstance(value, float):
        if not math.isfinite(value):
            raise ValueError('Catalog values must be finite.')
        if value.is_integer():
            return str(int(value))
    return str(value).strip()


def canonical_json(value):
    return json.dumps(value, sort_keys=True, ensure_ascii=True, separators=(',', ':'))



def record_id(row, term):
    fields = {name: canonical_value(value) for name, value in row.items()}
    identity = [fields.get(name, '') for name in IDENTITY_FIELDS]
    if not all(identity):
        raise ValueError(f'Missing section identity: {identity!r}')
    identity[0] = identity[0].upper()
    if identity[2].isdigit():
        identity[2] = str(int(identity[2]))
    return canonical_json([term, *identity])


SNAPSHOT_TRACKED_FIELDS = (
    'Subject Code', 'COURSE_NO', 'Section', 'Course Number', 'Course Number+',
    'Course Name', 'TITLE',
    'MAXCAP', 'Sec Capacity',
    '1st-years', '2nd-years', '3rd-years', '4th-years',
    'Fr Cap', 'So Cap', 'Jr Cap', 'Sr Cap',
    'INSTRUCTOR1', 'INSTRUCTOR2', 'INSTRUCTOR3', 'INSTRUCTOR4',
    'TEACHINGMETHOD', 'Writing-intensive', 'RESTRICTED',
    'CREDIT', 'Catalogue URL',
    'ROOM1', 'ROOM2', 'ROOM3', 'ROOM4',
    'DAY1', 'DAY2', 'DAY3', 'DAY4',
    'TIME1-start', 'TIME1-end',
    'TIME2-start', 'TIME2-end',
    'TIME3-start', 'TIME3-end',
    'TIME4-start', 'TIME4-end',
    'Prerequisite Narrative',
    'Corequisite Narrative',
    'Cross Listed Courses Narrative',
    'COMMENTS', 'Notes',
    'Course Description',
)


def prune_snapshot_variant(row: dict) -> dict:
    """Normalize split merge columns and keep only the 48 canonical snapshot fields."""
    normalized = {k: canonical_value(v) for k, v in row.items()}

    if not normalized.get('Course Number'):
        normalized['Course Number'] = (
            normalized.get('Course Number_x') # shouldn't have this anymore, but for back-compatibility
            or normalized.get('Course Number_y')
            or ''
        )

    if not normalized.get('DAY4'):
        normalized['DAY4'] = (
            normalized.get('DAY4_sched')
            or normalized.get('DAY4_colls')
            or ''
        )

    pruned = {field: normalized.get(field, '') for field in SNAPSHOT_TRACKED_FIELDS}
    pruned['Course Description'] = clean_html(pruned['Course Description'])
    return pruned


def build_snapshot(rows, term, term_label, generated_at=None, taken_seats_as_of_label=""):
    if not term or not term_label:
        raise ValueError('An explicit catalog term and label are required.')
    grouped = {}
    for row in rows:
        fields = prune_snapshot_variant(row)
        section_id = record_id(fields, term)
        identity = json.loads(section_id)[1:]
        for name, value in zip(IDENTITY_FIELDS, identity):
            fields[name] = value
        grouped.setdefault(section_id, {})[canonical_json(fields)] = fields
    if not grouped:
        raise ValueError('Cannot publish an empty catalog snapshot.')

    records = []
    identity_issues = []
    for section_id, variants in sorted(grouped.items()):
        records.append({'id': section_id, 'variants': [variants[key] for key in sorted(variants)]})
        if len(variants) > 1:
            identity_issues.append({'id': section_id, 'variant_count': len(variants)})

    content = {
        'schema_version': SCHEMA_VERSION,
        'term': term,
        'term_label': term_label,
        'records': records,
        'identity_issues': identity_issues,
    }
    content_hash = sha256(canonical_json(content).encode('utf-8')).hexdigest()
    return {
        **content,
        'as_of_label': taken_seats_as_of_label,
        'snapshot_id': content_hash,
        'generated_at': generated_at or datetime.now(timezone.utc).isoformat(),
    }


def publish_snapshot(snapshot, directory):
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f'{snapshot["snapshot_id"]}.json'
    if path.exists():
        stored = json.loads(path.read_text(encoding='utf-8'))
        content = {
            key: value
            for key, value in stored.items()
            if key not in ('snapshot_id', 'generated_at', 'as_of_label')
        }
        if sha256(canonical_json(content).encode('utf-8')).hexdigest() != snapshot['snapshot_id']:
            raise ValueError(f'Archived snapshot failed integrity validation: {path}')
        # Refresh the stored label/timestamp in place without creating a duplicate file
        if stored.get('as_of_label') != snapshot.get('as_of_label'):
            path.write_text(canonical_json(snapshot) + '\n', encoding='utf-8')
    else:
        path.write_text(canonical_json(snapshot) + '\n', encoding='utf-8')
    return path