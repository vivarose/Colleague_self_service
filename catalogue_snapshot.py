from datetime import datetime, timezone
from hashlib import sha256
import json
import math


SCHEMA_VERSION = 1
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


def build_snapshot(rows, term, term_label, generated_at=None):
    if not term or not term_label:
        raise ValueError('An explicit catalog term and label are required.')
    grouped = {}
    for row in rows:
        fields = {name: canonical_value(value) for name, value in row.items()}
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
        'snapshot_id': content_hash,
        'generated_at': generated_at or datetime.now(timezone.utc).isoformat(),
    }


def publish_snapshot(snapshot, directory):
    directory.mkdir(parents=True, exist_ok=True)
    path = directory / f'{snapshot["snapshot_id"]}.json'
    if path.exists():
        stored = json.loads(path.read_text(encoding='utf-8'))
        content = {key: value for key, value in stored.items() if key not in ('snapshot_id', 'generated_at')}
        if sha256(canonical_json(content).encode('utf-8')).hexdigest() != snapshot['snapshot_id']:
            raise ValueError(f'Archived snapshot failed integrity validation: {path}')
    else:
        path.write_text(canonical_json(snapshot) + '\n', encoding='utf-8')
    return path