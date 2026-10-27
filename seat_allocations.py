import re
import warnings


SECTION = r'[A-Z]+-[A-Z0-9]+-\d+'
CAP_FIELDS = ('Sec Capacity', 'Fr Cap', 'So Cap', 'Jr Cap', 'Sr Cap')


def allocation_key(subject, number, section):
    subject = str(subject).strip().upper()
    number = str(number).strip().upper()
    section = str(section).strip()
    if re.fullmatch(r'\d+(?:\.0)?', section):
        section = str(int(float(section)))
    if not subject or not number or not section:
        raise ValueError('Seat allocation identity is incomplete.')
    return subject, number, section


def parse_allocations(text):
    pattern = re.compile(
        rf'^({SECTION})\s+(.+?)\s+(\d+(?:\s+OPEN)?|OPEN)'
        r'\s+(\d+|OPEN)\s+(\d+|OPEN)\s+(\d+|OPEN)\s+(\d+|OPEN)\s*$'
    )
    rows = []
    seen = set()
    identifiers = re.findall(rf'\b{SECTION}\b', text)
    for line in text.splitlines():
        if not re.search(rf'\b{SECTION}\b', line):
            continue
        match = pattern.fullmatch(line.strip())
        if not match:
            raise ValueError(f'Unparsed seat allocation row: {line}')
        identity, title, *caps = match.groups()
        subject, number, section = identity.split('-')
        key = allocation_key(subject, number, section)
        if key in seen:
            raise ValueError(f'Duplicate normalized seat allocation: {identity}')
        seen.add(key)
        if all(value.isdigit() for value in caps) and int(caps[0]) != sum(map(int, caps[1:])):
            warnings.warn(f'Allocation sum differs from capacity for {identity}: {caps}', stacklevel=2)
        rows.append({
            'Subject Code': subject, 'Course Number': number, 'Section': section,
            **dict(zip(CAP_FIELDS, caps))
        })
    if not rows or len(rows) != len(identifiers):
        raise ValueError('Seat allocation extraction is empty or incomplete.')
    return rows


def merge_allocations(schedule, seats):
    schedule = schedule.copy()
    seats = seats.copy()
    for frame, number_column in ((schedule, 'Course Number+'), (seats, 'Course Number')):
        frame['_allocation_key'] = [
            allocation_key(row['Subject Code'], row[number_column], row['Section'])
            for _, row in frame.iterrows()
        ]
    if seats['_allocation_key'].duplicated().any():
        raise ValueError('Duplicate normalized seat allocation identities.')
    schedule_keys = set(schedule['_allocation_key'])
    seat_keys = set(seats['_allocation_key'])
    print(f'Schedule sections without explicit allocations: {len(schedule_keys - seat_keys)}')
    if seat_keys - schedule_keys:
        warnings.warn(f'PDF sections missing from schedule: {sorted(seat_keys - schedule_keys)}', stacklevel=2)
    return schedule.merge(
        seats[['_allocation_key', *CAP_FIELDS]],
        on='_allocation_key', how='left', validate='many_to_one'
    ).drop(columns='_allocation_key')
