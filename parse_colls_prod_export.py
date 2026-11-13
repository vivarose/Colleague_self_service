# -*- coding: utf-8 -*-
"""
Created on Thu Nov 12 19:38:20 2026

@author: vhorowit with Gemini

Reads colls-prod-export.csv and enriches it with parsed section, availability,
and meeting schedule columns.
"""

import re
import pandas as pd
import numpy as np
from pathlib import Path


BASE_DIR = Path(__file__).resolve().parent

#%% 1. Parse Section Name

SECTION_NAME_RE = re.compile(
    r'^\s*(?P<subject>[A-Za-z]+)'
    r'-(?P<course_plus>\d+[A-Za-z]*)'
    r'-(?P<section>[A-Za-z0-9]+)'
    r'(?:\s+(?P<designations>.*\S))?\s*$'
)


def parse_section_name(df: pd.DataFrame, col: str = 'Section Name') -> pd.DataFrame:
    """
    Parses 'AFRST-146W-01  Writing Intensive' into Subject Code, Course Number+,
    Course Number, Writing-intensive, Section, and Section Designations.
    """
    extracted = df[col].astype('string').str.extract(SECTION_NAME_RE)

    df['Subject Code'] = extracted['subject'].str.strip().str.upper()
    df['Course Number+'] = extracted['course_plus'].str.strip().str.upper()
    df['Section'] = extracted['section'].str.strip().str.upper()

    # Match join-catalogue.py conventions for catalog merging and W-flagging
    df['Writing-intensive'] = np.where(
        df['Course Number+'].str.contains('W', na=False), 'W', ''
    )
    df['Course Number'] = (
        df['Course Number+']
        .str.replace('W', '', regex=False)
        .str.strip()
    )
    df['Section Designations'] = (
        extracted['designations']
        .str.replace(r'\s{2,}', ', ', regex=True)
        .str.strip()
    )

    return df


#%% 2. Parse Availability

def parse_availability(df: pd.DataFrame, col: str = 'Availability') -> pd.DataFrame:
    """
    Parses 'Available / Taken / Capacity / Waitlisted' (e.g., '8 / 12 / 20 / 4')
    into nullable integer columns and derived availability flags.
    """
    out_cols = ['Seats Available', 'Seats Taken', 'Current Capacity', 'Seats Waitlisted']

    if col not in df.columns:
        for c in out_cols:
            df[c] = pd.Series(pd.NA, index=df.index, dtype='Int64')
        df['Has Open Seats'] = pd.Series(pd.NA, index=df.index, dtype='boolean')
        df['Has Waitlist'] = pd.Series(pd.NA, index=df.index, dtype='boolean')
        return df

    extracted = (
        df[col]
        .astype('string')
        .str.extract(r'^\s*(\d+)\s*/\s*(\d+)\s*/\s*(\d+)\s*/\s*(\d+)\s*$')
    )

    for idx, name in enumerate(out_cols):
        df[name] = pd.to_numeric(extracted[idx], errors='coerce').astype('Int64')

    valid_mask = df['Seats Available'].notna() & df['Current Capacity'].notna()
    df['Has Open Seats'] = pd.Series(pd.NA, index=df.index, dtype='boolean')
    df.loc[valid_mask, 'Has Open Seats'] = df.loc[valid_mask, 'Seats Available'] > 0

    wl_mask = df['Seats Waitlisted'].notna()
    df['Has Waitlist'] = pd.Series(pd.NA, index=df.index, dtype='boolean')
    df.loc[wl_mask, 'Has Waitlist'] = df.loc[wl_mask, 'Seats Waitlisted'] > 0

    return df


#%% 3. Parse Meeting Times

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


def add_parsed_schedule_columns(df: pd.DataFrame, source_col: str = 'Meeting Information') -> pd.DataFrame:
    """Applies parse_schedule_entry across a DataFrame column and attaches DAY*/TIME*/ROOM* columns."""
    schedule_df = pd.DataFrame(
        df[source_col].apply(parse_schedule_entry).tolist(),
        index=df.index,
    )
    for col in schedule_df.columns:
        df[col] = schedule_df[col]
    return df


#%% 4. Load and Parse Pipeline

def load_and_parse_colls_export(csv_path: str | Path = BASE_DIR / 'colls-prod-export.csv') -> pd.DataFrame:
    """Reads colls-prod-export.csv and applies all section, availability, and schedule parsers."""
    df = pd.read_csv(csv_path, encoding='utf-8-sig')
    df.columns = df.columns.str.strip()

    df = parse_section_name(df, col='Section Name')
    df = parse_availability(df, col='Availability')
    df = add_parsed_schedule_columns(df, source_col='Meeting Information')

    return df


def build_course_no(subj, course_num, section):
    """Builds an 11-character fixed-width COURSE_NO from subject, number, and section."""
    if pd.isna(subj) or pd.isna(course_num):
        return pd.NA

    subj_str = str(subj).strip().upper()
    num_str = str(course_num).strip().upper()
    if not subj_str or not num_str or subj_str == 'NAN' or num_str == 'NAN':
        return pd.NA

    sec_str = '' if pd.isna(section) else str(section).strip()
    # Guard against numeric sections read as floats (e.g., 1.0 -> '01')
    if sec_str.endswith('.0'):
        sec_str = sec_str[:-2]
    if sec_str.isdigit():
        sec_str = sec_str.zfill(2)

    return f"{subj_str:<5}{num_str:<4}{sec_str:<2}"

#%% if name == main


if __name__ == '__main__':
    colls_csv = BASE_DIR / 'colls-prod-export.csv'
    df_colls = load_and_parse_colls_export(colls_csv)

    print(f"Loaded {len(df_colls)} rows and {len(df_colls.columns)} total columns.")

    # Quick sanity checks on edge cases
    unparsed_sections = df_colls[df_colls['Subject Code'].isna()]
    if not unparsed_sections.empty:
        print(f"⚠️ Unparsed Section Names ({len(unparsed_sections)}):")
        print(unparsed_sections[['Section Name']].head(10))

    has_meeting_text = df_colls['Meeting Information'].notna() & (df_colls['Meeting Information'].astype(str).str.strip() != '')
    unparsed_meetings = df_colls[has_meeting_text & df_colls['DAY1'].isna()]
    if not unparsed_meetings.empty:
        print(f"⚠️ Meeting Information without parsed DAY1 ({len(unparsed_meetings)}):")
        print(unparsed_meetings[['Section Name', 'Meeting Information']].drop_duplicates().head(10))

    preview_cols = [
        'Subject Code', 'Course Number+', 'Section', 'Writing-intensive',
        'Faculty', 'DAY1', 'TIME1-start', 'TIME1-end', 'ROOM1',
        'Seats Available', 'Current Capacity', 'Seats Waitlisted', 
    ]
    print("\nSample parsed output:")
    print(df_colls[preview_cols].head(5).to_string(index=False))