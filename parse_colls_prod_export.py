# -*- coding: utf-8 -*-
"""
Created on Thu Nov 12 19:38:20 2026

@author: vhorowit with Gemini
"""

import re
import pandas as pd
import numpy as np

#%% parse availability

def parse_availability(df, col='Availability'):
    """
    Parses 'Available / Taken / Capacity / Waitlisted' (e.g., '8 / 12 / 20 / 4')
    into nullable integer columns and derived availability flags.
    """
    out_cols = ['Seats Available', 'Seats Taken', 'Seat Capacity', 'Seats Waitlisted']

    if col not in df.columns:
        for c in out_cols:
            df[c] = pd.Series(pd.NA, index=df.index, dtype='Int64')
        df['Has Open Seats'] = pd.Series(pd.NA, index=df.index, dtype='boolean')
        df['Has Waitlist'] = pd.Series(pd.NA, index=df.index, dtype='boolean')
        df['Fill Rate'] = np.nan
        return df

    extracted = (
        df[col]
        .astype('string')
        .str.extract(r'^\s*(\d+)\s*/\s*(\d+)\s*/\s*(\d+)\s*/\s*(\d+)\s*$')
    )

    for idx, name in enumerate(out_cols):
        df[name] = pd.to_numeric(extracted[idx], errors='coerce').astype('Int64')

    valid_mask = df['Seats Available'].notna() & df['Seat Capacity'].notna()
    df['Has Open Seats'] = pd.Series(pd.NA, index=df.index, dtype='boolean')
    df.loc[valid_mask, 'Has Open Seats'] = df.loc[valid_mask, 'Seats Available'] > 0

    wl_mask = df['Seats Waitlisted'].notna()
    df['Has Waitlist'] = pd.Series(pd.NA, index=df.index, dtype='boolean')
    df.loc[wl_mask, 'Has Waitlist'] = df.loc[wl_mask, 'Seats Waitlisted'] > 0

    # Guard against 0 / 0 / 0 / 0 entries
    cap = df['Seat Capacity'].astype('Float64')
    taken = df['Seats Taken'].astype('Float64')
    df['Fill Rate'] = np.where(cap > 0, (taken / cap).round(3), np.nan)

    return df

#%% parse meeting times

DAY_TOKEN_MAP = {
    'M': 'M', 'MON': 'M',
    'T': 'T', 'TU': 'T', 'TUE': 'T',
    'W': 'W', 'WED': 'W',
    'TH': 'R', 'THU': 'R', 'R': 'R',
    'F': 'F', 'FRI': 'F',
    'S': 'S', 'SAT': 'S',
    'U': 'U', 'SUN': 'U',
}

MEETING_RE = re.compile(
    r'(?P<days>(?:(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun|Th|Tu|M|T|W|R|F|S|U)\s*,\s*)*'
    r'(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun|Th|Tu|M|T|W|R|F|S|U))'
    r'\s+'
    r'(?P<start_h>\d{1,2}):(?P<start_m>\d{2})\s*(?P<start_p>AM|PM)?'
    r'\s*-\s*'
    r'(?P<end_h>\d{1,2}):(?P<end_m>\d{2})\s*(?P<end_p>AM|PM)'
    r'(?:\s+(?P<start_date>\d{1,2}/\d{1,2}/\d{4})\s*-\s*(?P<end_date>\d{1,2}/\d{1,2}/\d{4}))?'
    r'(?:\s+(?P<room>.*?)\s*\((?P<method>[^)]+)\))?',
    re.IGNORECASE,
)


def normalize_days(raw_days: str) -> str:
    """Converts 'T, Th' -> 'TR' and 'M, W' -> 'MW' to match DAY_MAP in course_info_to_html.py."""
    tokens = [t.strip().upper() for t in raw_days.split(',') if t.strip()]
    return ''.join(DAY_TOKEN_MAP.get(t, t) for t in tokens)


def resolve_periods(start_h: int, start_p: str | None, end_h: int, end_p: str) -> tuple[str, str]:
    """Infers missing start AM/PM (e.g., '8:30-9:45 AM' -> AM/AM, '11:00-12:15 PM' -> AM/PM)."""
    end_p = end_p.upper()
    if start_p:
        return start_p.upper(), end_p
    if end_p == 'AM':
        return 'AM', 'AM'
    # end_p is 'PM' and start_p was omitted:
    # Handle spans crossing noon (e.g., 11:00-12:15 PM or 11:30-1:00 PM)
    if (end_h == 12 and start_h != 12) or (end_h < 12 and start_h < 12 and start_h > end_h):
        return 'AM', 'PM'
    return 'PM', 'PM'


def parse_schedule_entry(entry: str, max_slots: int = 4) -> dict:
    """
    Parses one or more meeting blocks from a raw schedule string into columns
    compatible with course_info_to_html.py (DAY1..DAY4, TIME1-start..TIME4-end, ROOM1..ROOM4).
    """
    parsed = {}
    for i in range(1, max_slots + 1):
        parsed[f'DAY{i}'] = None
        parsed[f'TIME{i}-start'] = None
        parsed[f'TIME{i}-end'] = None
        parsed[f'ROOM{i}'] = None

    if not entry or pd.isna(entry):
        return parsed

    matches = list(MEETING_RE.finditer(str(entry).strip()))
    for idx, m in enumerate(matches[:max_slots], start=1):
        sh, sm = int(m.group('start_h')), m.group('start_m')
        eh, em = int(m.group('end_h')), m.group('end_m')
        sp, ep = resolve_periods(sh, m.group('start_p'), eh, m.group('end_p'))

        raw_room = (m.group('room') or '').strip()
        room = None if raw_room.upper() in {'', 'TBD', 'TBA', 'N/A'} else raw_room

        parsed[f'DAY{idx}'] = normalize_days(m.group('days'))
        parsed[f'TIME{idx}-start'] = f'{sh}:{sm} {sp}'
        parsed[f'TIME{idx}-end'] = f'{eh}:{em} {ep}'
        parsed[f'ROOM{idx}'] = room

    return parsed


def add_parsed_schedule_columns(df: pd.DataFrame, source_col: str) -> pd.DataFrame:
    """Applies parse_schedule_entry across a DataFrame column and attaches DAY*/TIME*/ROOM* columns."""
    schedule_df = df[source_col].apply(parse_schedule_entry).apply(pd.Series)
    for col in schedule_df.columns:
        df[col] = schedule_df[col]
    return df