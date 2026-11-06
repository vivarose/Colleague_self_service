# -*- coding: utf-8 -*-
"""
Created on Sat Oct 24 21:46:53 2026

@author: vhorowit with Gemini
"""

import html
import json
from hashlib import sha256
from pathlib import Path
import re
import pandas as pd
import os
import warnings

from catalogue_snapshot import build_snapshot, canonical_json, publish_snapshot, record_id

BASE_DIR = Path(__file__).resolve().parent
CATALOG_TERM = '2027-spring'
CATALOG_TERM_LABEL = 'Spring 2027'
CATALOG_TERM_LABEL_SHORT = "Spr '27"


TEACHING_METHODS = {
    'LEC': None,
    'LEC,Final Exam': None,
    'LEC Final Exam': None,
    'SEM': 'Seminar',
    'LAB': 'Laboratory',
    'IND': 'Independent Study',
    'STU': 'Studio',
    'CON': 'Concert / Performance',
    'REC': 'Recitation',
    'PRAC': 'Practicum',
    'DISC': 'Discussion',
    'WKS': 'Workshop',
    'STUDIO': 'Studio'
}

SUBJECT_MAP = {
    'AFRST': 'Africana Studies',
    'AIIS': 'American Indian and Indigenous Studies',
    'AMST': 'American Studies',
    'ANTHR': 'Anthropology',
    'ARABC': 'Arabic',
    'ARCH': 'Archaeology',
    'ART': 'Art',
    'ARTH': 'Art History',
    'ASNST': 'Asian Studies',
    'BICHM': 'Biochemistry/Molecular Biology',
    'BIO': 'Biology',
    'CHEM': 'Chemistry',
    'CHNSE': 'Chinese',
    'CLASC': 'Classical Studies',
    'CNMS': 'Cinema and Media Studies',
    'COLEG': 'College Courses and Seminars',
    'CPSCI': 'Computer Science',
    'CRWR': 'Creative Writing',
    'DANCE': 'Dance and Movement Studies',
    'DARTS': 'Digital Arts',
    'ECON': 'Economics',
    'EDUC': 'Education Studies',
    'ENVST': 'Environmental Studies',
    'FRNCH': 'French and Francophone Studies',
    'GEARC' : 'Geoarchaeology',
    'GEOSC': 'Geosciences',
    'GERMN': 'German Studies',
    'GOVT': 'Government',
    'GREEK': 'Greek',
    'HEBRW': 'Hebrew',
    'HIST': 'History',
    'HSPST': 'Hispanic Studies',
    'ITALN': 'Italian',
    'JAPN': 'Japanese',
    'JLJS': 'Jurisprudence, Law, and Justice Studies',
    'JWST': 'Jewish Studies',
    'LATIN': 'Latin',
    'LIT': 'Literature',
    'LTAM': 'Latin American Studies',
    'MATH': 'Mathematics and Statistics',
    'MDRST': 'Medieval and Renaissance Studies',
    'MEIWS': 'Middle East and Islamic World Studies',
    'MUSIC': 'Music',
    'NEURO': 'Neuroscience',
    'PHIL': 'Philosophy',
    'PHYS': 'Physics',
    'PPOL': 'Public Policy',
    'PSYCH': 'Psychology',
    'RELST': 'Religious Studies',
    'RSNST': 'Russian Studies',
    'RUSSN': 'Russian Language',
    'SOC': 'Sociology',
    'THETR': 'Theatre',
    'WMGST': "Women's and Gender Studies",
}


def build_dept_options(df, SUBJECT_MAP=SUBJECT_MAP):
    """Extracts unique subject codes, resolves names via dictionary, and builds HTML <option> tags."""
    if 'Subject Code' in df.columns:
        code_col = 'Subject Code'
        name_col = 'Subject Name' if 'Subject Name' in df.columns else None
    elif 'SUBJ' in df.columns:
        code_col = 'SUBJ'
        name_col = None
    else:
        return '<option value="all">All Departments</option>'

    # Get sorted unique subject codes present in the dataframe
    unique_codes = sorted(df[code_col].dropna().unique())

    options = ['<option value="all">All Departments</option>']
    for raw_code in unique_codes:
        code = str(raw_code).strip()
        if not code:
            continue

        # Lookup in dictionary, falling back to DataFrame's Subject Name or the code itself
        if code in SUBJECT_MAP:
            name = SUBJECT_MAP[code]
        elif name_col:
            matching_rows = df[df[code_col] == raw_code][name_col].dropna()
            name = str(matching_rows.iloc[0]).strip() if not matching_rows.empty else code
        else:
            name = code

        options.append(f'<option value="{code}">{code} — {name}</option>')
    
    return '\n'.join(options)

DAY_MAP = {'M': 'Mon', 'T': 'Tue', 'W': 'Wed', 'R': 'Thu', 'F': 'Fri'}


def time_to_minutes(t_str):
    """Converts '08:30AM' or '02:30PM' into integer minutes past midnight for JS filtering."""
    if not is_valid(t_str):
        return None
    s = str(t_str).strip().upper()
    m = re.match(r'^(\d{1,2}):(\d{2})\s*(AM|PM)?$', s)
    if not m:
        return None
    hr, mn, period = int(m.group(1)), int(m.group(2)), m.group(3)
    if period == 'PM' and hr < 12:
        hr += 12
    elif period == 'AM' and hr == 12:
        hr = 0
    return hr * 60 + mn


def is_valid(val):
    """Check if a value is non-empty and non-null."""
    if pd.isna(val) or val is None:
        return False
    s = str(val).strip()
    return s.lower() not in [
        '',
        'nan',
        'none',
        'null',
        '<p><br></p>',
        '&nbsp;',
        '<p></p>',
    ]


def clean_time(t_str):
    """Format time string (e.g., '08:30AM' -> '8:30 AM')."""
    if not is_valid(t_str):
        return ''
    s = re.sub(r'^0(\d)', r'\1', str(t_str).strip())
    s = re.sub(r'(AM|PM|am|pm)', r' \1', s)
    return re.sub(r'\s+', ' ', s).strip()


def clean_num(val):
    """Format numeric values cleanly (e.g., 6.0 -> 6)."""
    if not is_valid(val):
        return ''
    try:
        f = float(val)
        return str(int(f)) if f.is_integer() else f'{f:g}'
    except ValueError:
        return str(val).strip()


def format_days(day_str):
    """Format day abbreviations (e.g., 'MW' -> 'Mon/Wed')."""
    if not is_valid(day_str):
        return ''
    d_clean = str(day_str).strip()
    res = [DAY_MAP.get(c, c) for c in d_clean]
    return '/'.join(res)


def clean_room(room_val):
    """Clean room strings and drop 'TBA' entries."""
    if not is_valid(room_val):
        return ''
    r = str(room_val).strip()
    if r.upper() in ['TBA', 'TBA TBA', 'TBA/TBA', 'N/A']:
        return ''
    return r


def clean_html(text):
    """Strip raw HTML tags and decode HTML entities."""
    if not is_valid(text):
        return ''
    text_unescaped = html.unescape(str(text))
    clean = re.sub(r'<br\s*/?>', ' ', text_unescaped)
    clean = re.sub(r'<p[^>]*>', '', clean)
    clean = re.sub(r'</p>', ' ', clean)
    clean = re.sub(r'<[^>]+>', '', clean)
    return re.sub(r'\s+', ' ', clean).strip()


def clean_title(title_str):
    """Format titles nicely, preserving Roman numerals."""
    if not is_valid(title_str):
        return ''
    s = str(title_str).strip()
    if not s.isupper():
        return s

    small_words = {
        'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'in',
        'nor', 'of', 'on', 'or', 'so', 'the', 'to', 'up', 'yet',
    }
    roman_pattern = r'^(?i:M{0,4}(C[MD]|D?C{0,3})(X[CL]|L?X{0,3})(I[XV]|V?I{0,3}))$'

    words = s.split()
    cleaned = []
    for idx, w in enumerate(words):
        w_lower = w.lower()
        if re.match(roman_pattern, w) and len(w) > 1:
            cleaned.append(w.upper())
        elif w_lower == 'i':
            cleaned.append('I')
        elif idx > 0 and w_lower in small_words:
            cleaned.append(w_lower)
        else:
            cleaned.append(w.capitalize())

    return ' '.join(cleaned)


def clean_comments_notes(comments, notes):
    """Combines COMMENTS and Notes while deduplicating overlapping text."""
    raw_texts = []
    for text in [comments, notes]:
        if is_valid(text):
            cleaned = clean_html(text)
            if cleaned:
                raw_texts.append(cleaned)

    if not raw_texts:
        return None
    if len(raw_texts) == 1:
        return raw_texts[0]

    c_str, n_str = raw_texts[0], raw_texts[1]
    c_norm, n_norm = c_str.rstrip('.'), n_str.rstrip('.')

    if c_norm in n_norm:
        return n_str if n_str.endswith('.') else n_str + '.'
    if n_norm in c_norm:
        return c_str if c_str.endswith('.') else c_str + '.'

    sentences = []
    for text in raw_texts:
        chunks = re.split(r'(?<=[.!?])\s+', text)
        for chunk in chunks:
            chunk_s = chunk.strip()
            norm_chunk = chunk_s.rstrip('.')
            if chunk_s and not any(
                norm_chunk == existing.rstrip('.') for existing in sentences
            ):
                sentences.append(chunk_s)

    return ' '.join(sentences)

def get_base_num(course_num):
    """Strips trailing 'L' or section modifiers to get the primary course number (e.g., '270L' -> '270')."""
    if not is_valid(course_num):
        return ''
    # Remove trailing L, W, or section suffixes
    clean = re.sub(r'[LwLW]$', '', str(course_num).strip())
    return clean

def group_courses_and_labs(rows):
    """Groups laboratory courses under their parent main courses."""
    parents = []
    labs_by_key = {}  # Key: (Subject Code, Base Number) -> List of lab rows

    # First Pass: Separate Main Courses and Lab Courses
    for row in rows:
        subj = str(row.get('Subject Code', '')).strip()
        num_plus = str(row.get('Course Number+', row.get('Course Number', ''))).strip()
        
        is_lab = num_plus.endswith('L') # only look for labs affiliated with coursees
        base_num = get_base_num(num_plus)
        group_key = (subj, base_num)

        if is_lab:
            if group_key not in labs_by_key:
                labs_by_key[group_key] = []
            labs_by_key[group_key].append(row)
        else:
            parents.append((group_key, row))

    return parents, labs_by_key

def get_course_level(course_num):
    """Extracts the first digit of the course number to determine level (e.g., '146W' -> '100')."""
    if not is_valid(course_num):
        return ''
    match = re.search(r'\d', str(course_num))
    if match:
        first_digit = match.group(0)
        return f'{first_digit}00'
    return ''


#%% Describe historical availability


def build_historical_html(row):
    """Generates HTML for the single-term prior historical snapshot section."""
    # Check if course was offered or has history data
    hist_offered = row.get('Offered_Last_Year')
    if hist_offered is None:
        return None

    hist_term = row.get('Hist_Term_Label', 'Spring 2026')

    # Status when course was not offered in the prior term
    if hist_offered is False or str(hist_offered).strip().lower() in ['false', 'no', '0', 'not offered']:
        return f'''<div><span style="color: #6e7781;">Was not offered in {hist_term}</span></div>'''

    # Extract metrics safely from dataframe row
    fill_rate = row.get('Hist_Fill_Rate')
    enrolled = row.get('Hist_Total_Enrollment',0)
    capacity = row.get('Hist_Seats_Offered',0) 
    sections = row.get('Hist_Sections_Offered',0)
    waitlist = row.get('Hist_Total_Waitlist')
    open_seats = row.get('Hist_Open_Seats')

    # Return fallback if no metrics are present
    if fill_rate is None and enrolled is None and capacity is None:
        if is_valid(hist_offered):
            return f'''<div style="margin-top: 12px; padding: 8px 12px; background-color: #f6f8fa; border: 1px solid #e1e4e8; border-radius: 6px; font-size: 12px; color: #57606a;">
    📊 <b>{hist_term} History:</b> <span style="color: #6e7781;">Not offered in {hist_term}</span>
  </div>'''
        return None

    # Format values
    fill_str = f"{clean_num(fill_rate*100)}%" if is_valid(fill_rate) else ""
    counts_str = f" ({clean_num(enrolled)} / {clean_num(capacity)} seats)" if is_valid(enrolled) and is_valid(capacity) else ""
    fill_display = f"{fill_str}{counts_str}".strip() or "N/A"

    sections_display = clean_num(sections) if is_valid(sections) else "1"
    waitlist_display = f"{clean_num(waitlist)} student(s)" if is_valid(waitlist) else "0 students"
    open_seats_display = f"{clean_num(open_seats)} remaining" if is_valid(open_seats) else "0 remaining"

    hist_info_html = \
        f'''<p style="font-size: 12px; color: #57606a; margin: 8px 0;">
        <b>📊 Prior Term ({hist_term}):</b> 
        Fill Rate: <b style="color: #24292f;">{fill_display}</b> &bull; 
        Sections: <b style="color: #24292f;">{sections_display}</b> &bull; 
        Waitlist: <b style="color: #24292f;">{waitlist_display}</b> &bull; 
        Open Seats: <b style="color: #24292f;">{open_seats_display}</b>
        </p>'''

    return hist_info_html
  
  

#%% Build html!
def display_course_info(data, mode='html'):
    """Displays formatted course details for DataFrame or list of rows."""
    if isinstance(data, pd.DataFrame):
        rows = [r.to_dict() for _, r in data.iterrows()]
    elif isinstance(data, pd.Series):
        rows = [data.to_dict()]
    elif isinstance(data, dict):
        rows = [data]
    elif isinstance(data, list):
        rows = [r.to_dict() if isinstance(r, pd.Series) else r for r in data]
    else:
        raise TypeError(
            'Input must be a pandas DataFrame, Series, dict, or list of rows.'
        )

    if mode.lower() == 'html':
        parents, labs_by_key = group_courses_and_labs(rows) # group labs with parent courses
        html_cards = []

        for group_key, row in parents:
            matching_labs = labs_by_key.get(group_key, [])
            record_ids = [record_id(row, CATALOG_TERM)]
            record_ids.extend(record_id(lab, CATALOG_TERM) for lab in matching_labs)
            record_ids_attribute = html.escape(json.dumps(record_ids), quote=True)
            all_starts, all_ends, all_days = [], [], []
            meeting_slots = []

            for i in range(1, 5):
                st_min = time_to_minutes(row.get(f'TIME{i}-start'))
                en_min = time_to_minutes(row.get(f'TIME{i}-end'))
                d_code = str(row.get(f'DAY{i}', '')).strip()

                if st_min is not None and en_min is not None and is_valid(d_code):
                    all_starts.append(st_min)
                    all_ends.append(en_min)
                    all_days.append(d_code)
                    meeting_slots.append({'days': d_code, 'start': st_min, 'end': en_min})

            min_start = str(min(all_starts)) if all_starts else ''
            max_end = str(max(all_ends)) if all_ends else ''
            days_str = ''.join(all_days)
            slots_json_escaped = html.escape(json.dumps(meeting_slots))

            yr1 = str(row.get('1st-years', 'ok')).strip().lower()
            yr2 = str(row.get('2nd-years', 'ok')).strip().lower()
            yr3 = str(row.get('3rd-years', 'ok')).strip().lower()
            yr4 = str(row.get('4th-years', 'ok')).strip().lower()

            is_wi = (
                'true'
                if str(row.get('Writing-intensive', '')).strip().upper() == 'W'
                else 'false'
            )

            ## Prerequisite check.
            prereq = row.get('Prerequisite Narrative') or row.get('Prerequisites')
            prereq_display = str(prereq).strip() if is_valid(prereq) else None

            # Check Course Description for implicit prerequisite wording (e.g., "Prerequisite, 113 or placement by the department")
            desc_text = str(row.get('Course Description', ''))
            desc_prereq_match = re.search(
                r'Prerequisite,?\s+.*?\s+or\s+placement\s+by\s+the\s+department',
                desc_text,
                re.IGNORECASE,
            )

            if desc_prereq_match:
                has_prereq = 'true'
                if not prereq_display:
                    prereq_display = desc_prereq_match.group(0).strip()
            else:
                has_prereq = 'true' if prereq_display else 'false'

            subj = str(row.get('Subject Code', '')).strip()
            num_plus = str(
                row.get('Course Number+', row.get('Course Number', ''))
            ).strip()
            course_level = get_course_level(num_plus)
            sec = str(row.get('Section', '')).strip()

            code_sec = f'{subj} {num_plus}'
            if is_valid(sec) and sec != 'nan':
                code_sec += f"-{sec.zfill(2) if sec.isdigit() else sec}"

            raw_title = row.get('Course Name')
            if not is_valid(raw_title):
                raw_title = row.get('TITLE')
            title = clean_title(raw_title)
            full_title = f'{code_sec}: {title}' if title else code_sec
            
            cat_url = row.get('Catalogue URL')
            if is_valid(cat_url):
                url_str = str(cat_url).strip()
                title_html = f'<a href="{html.escape(url_str)}" target="_blank" style="color: inherit; text-decoration: underline;">{full_title}</a>'
            else:
                title_html = full_title

            sub_items = [] # set up subheader, starting with subject name
            subj_name_plus = row.get('Subject Name+')
            subj_name = row.get('Subject Name')
            
            # 1st Attempt: Subject Name+ (e.g. specific catalog/override name)
            if is_valid(subj_name_plus):
                sub_items.append(str(subj_name_plus).strip())
            
            # 2nd Attempt: Dictionary Map lookup using Subject Code
            elif is_valid(SUBJECT_MAP.get(subj)):
                sub_items.append(SUBJECT_MAP[subj].strip())
            
            # 3rd Attempt: General Subject Name from catalog
            elif is_valid(subj_name):
                sub_items.append(str(subj_name).strip())
            
            # 4th Attempt: Raw Subject Code (e.g. 'AFRST', 'BIO')
            elif is_valid(subj):
                sub_items.append(subj)

            cred = row.get('CREDIT')
            if is_valid(cred):
                c_val = clean_num(cred)
                sub_items.append(f"{c_val} Credit{'s' if c_val != '1' else ''}")

            tm = str(row.get('TEACHINGMETHOD', '')).strip()
            if is_valid(tm):
                method_name = TEACHING_METHODS.get(tm.upper(), tm)
                if is_valid(method_name):
                    sub_items.append(method_name)

            wi = str(row.get('Writing-intensive', '')).strip()
            if is_valid(wi) and wi.upper() == 'W':
                sub_items.append('Writing Intensive')

            subtitle_str = ' &nbsp;•&nbsp; '.join(sub_items)

            schedules = []
            for i in range(1, 5):
                days = format_days(row.get(f'DAY{i}'))
                start = clean_time(row.get(f'TIME{i}-start'))
                end = clean_time(row.get(f'TIME{i}-end'))
                room = clean_room(row.get(f'ROOM{i}'))

                if start and end:
                    time_str = f'{start} – {end}'
                    if days:
                        time_str = f'{days} {time_str}'
                    if room:
                        time_str += f' @ {room}'
                    schedules.append(time_str)

            schedule_display = ' &nbsp;|&nbsp; '.join(schedules) if schedules else 'TBA'

            instructors = []
            for i in range(1, 5):
                inst = str(row.get(f'INSTRUCTOR{i}', '')).strip()
                if is_valid(inst) and inst not in instructors:
                    instructors.append(inst)
            instructor_display = ', '.join(instructors) if instructors else 'TBA'

            cap_parts = []
            max_cap = row.get('MAXCAP') or row.get('Sec Capacity')
            if is_valid(max_cap):
                cap_parts.append(f'{clean_num(max_cap)} Seats')

            seat_breakdown = []
            for yr, label in [('Fr Cap', 'Fr'), ('So Cap', 'So'), ('Jr Cap', 'Jr'), ('Sr Cap', 'Sr')]:
                v = row.get(yr)
                if is_valid(v):
                    seat_breakdown.append(f'{label}: {clean_num(v)}')

            if seat_breakdown:
                cap_parts.append(f"[{', '.join(seat_breakdown)}]")

            capacity_display = ' '.join(cap_parts) if cap_parts else 'N/A'

            cross_list = row.get('Cross Listed Courses Narrative') or row.get('Cross Listed Courses')
            cross_display = str(cross_list).strip().strip(',') if is_valid(cross_list) else None

            comments_display = clean_comments_notes(row.get('COMMENTS'), row.get('Notes'))
                 
            ## BUILD THE CARD!
            card = [
                (
                    f'<div class="course-card" data-subject="{subj}" data-level="{course_level}"'
                    f' data-record-ids="{record_ids_attribute}"'
                    f' data-days="{days_str}" data-start-min="{min_start}"'
                    f' data-end-min="{max_end}" data-wi="{is_wi}"'
                    f' data-has-prereq="{has_prereq}"'
                    f' data-slots="{slots_json_escaped}"'
                    f' data-1st-years="{yr1}" data-2nd-years="{yr2}"'
                    f' data-3rd-years="{yr3}" data-4th-years="{yr4}"'
                    ' style="font-family: \'Segoe UI\', Arial, sans-serif; border: 1px solid #d0d7de;'
                    ' border-radius: 8px; padding: 18px 20px; margin-bottom: 20px; background-color: #ffffff;'
                    ' box-shadow: 0 2px 5px rgba(0,0,0,0.05);">'
                ),
                " <div style='border-bottom: 2px solid #002f6c; padding-bottom: 8px; margin-bottom: 14px; display: flex; justify-content: space-between; align-items: flex-start;'>",
                "   <div>",
                f"     <h3 style='margin: 0; color: #002f6c; font-size: 19px; font-weight: 700;'>{title_html}</h3>",
            ]

            if subtitle_str:
                card.append(
                    f"     <div style='margin-top: 4px; font-size: 13px; color: #57606a; font-weight: 600;'>{subtitle_str}</div>"
                )

            card.extend([
                "   </div>",
                "   <button onclick='hideCourseCard(this)' style='background: none; border: 1px solid #d0d7de; border-radius: 4px; padding: 2px 8px; font-size: 12px; color: #cf222e; cursor: pointer; font-weight: 600;' title='Hide this course'>✕ Hide</button>",
                " </div>",
                " <div style='font-size: 14px; line-height: 1.8; color: #24292f;'>",
                f'   <div>🕒 <b>Schedule:</b> {schedule_display}</div>',
                f'   <div>👤 <b>Instructor:</b> {instructor_display}</div>',
                f'   <div>👥 <b>Capacity:</b> {capacity_display}</div>',
            ])
            

            if cross_display:
                card.append(f'   <div>🔀 <b>Cross-listed:</b> {cross_display}</div>')

            if prereq_display:
                card.append(f'   <div>📋 <b>Prerequisites:</b> {prereq_display}</div>')

            if comments_display:
                card.append(f'   <div>💬 <b>Comments:</b> {comments_display}</div>')

            hist_info_html = build_historical_html(row)
            if is_valid(hist_info_html):
                card.append(f'{hist_info_html}')

            card.append(' </div>')

            desc = row.get('Course Description')
            if is_valid(desc):
                cleaned_desc = clean_html(desc)
                card.extend([
                    " <div style='margin-top: 14px;'>",
                    "   <div style='font-size: 12px; font-weight: 700; color: #57606a; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;'>📖 Description</div>",
                    "   <div style='font-size: 14px; color: #24292f; line-height: 1.6; background-color: #f6f8fa; border-left: 4px solid #002f6c; padding: 10px 14px; border-radius: 0 6px 6px 0;'>",
                    f"     {cleaned_desc}",
                    "   </div>",
                    " </div>",
                ])
                
            if matching_labs:
                card.append('  <div class="nested-labs-container" style="margin-top: 15px; padding: 12px; background-color: #f0f4f8; border-radius: 6px; border: 1px solid #d0d7de;">')
                card.append(f'    <strong style="color: #002f6c; font-size: 13px;">🔬 Required Co-requisite Lab Sections ({len(matching_labs)} available):</strong>')
                
                for lab in matching_labs:
                    lab_sec = str(lab.get('Section', '')).strip()
                    lab_code = f"{lab.get('Subject Code', '')} {lab.get('Course Number+', '')}-{lab_sec}"
                    
                    # Format lab schedule
                    lab_schedules = []
                    for i in range(1, 5):
                        d = format_days(lab.get(f'DAY{i}'))
                        st = clean_time(lab.get(f'TIME{i}-start'))
                        en = clean_time(lab.get(f'TIME{i}-end'))
                        rm = clean_room(lab.get(f'ROOM{i}'))
                        if st and en:
                            lab_schedules.append(f"{d} {st}–{en}" + (f" @ {rm}" if rm else ""))
                    
                    lab_sched_str = " | ".join(lab_schedules) if lab_schedules else "TBA"
                    lab_inst = lab.get('INSTRUCTOR1', 'TBA')
                    lab_cap = clean_num(lab.get('MAXCAP') or lab.get('Sec Capacity', ''))
    
                    card.append(f'''
                    <div class="lab-card" style="margin-top: 8px; padding: 10px; background: #ffffff; border-left: 3px solid #002f6c; border-radius: 4px; box-shadow: 0 1px 3px rgba(0,0,0,0.05);">
                        <div style="font-weight: bold; font-size: 13px; color: #24292f;">{lab_code}</div>
                        <div style="font-size: 12px; color: #57606a; margin-top: 2px;">
                            🕒 <b>Schedule:</b> {lab_sched_str} | 👤 <b>Instructor:</b> {lab_inst} | 👥 <b>Seats:</b> {lab_cap}
                        </div>
                    </div>
                    ''')
                card.append('  </div>') # Close nested-labs-container
            card.append('  </div>') # Close card-body

            html_cards.append('\n'.join(card))

        return '\n'.join(html_cards)


def load_file_content(filename):
    """Loads text content from a file relative to the script directory."""
    file_path = BASE_DIR / filename
    if not file_path.exists():
        raise FileNotFoundError(f"Required file '{file_path}' was not found.")
    with open(file_path, 'r', encoding='utf-8') as f:
        return f.read()




if __name__ == '__main__':
    excel_file = BASE_DIR / 'merged_course_schedule.xlsx'
    html_template_file = 'htmltemplate.html'
    js_file = 'filter.js'

    if not excel_file.exists():
        print(f"❌ Error: Could not find '{excel_file}'. Run join-catalogue.py first.")
        raise SystemExit(1)
    else:
        try:
            print("📂 Loading HTML template and JavaScript files...")
            template_text = load_file_content(html_template_file)
            js_text = load_file_content(js_file)
    
            print(f"📊 Reading dataset from '{excel_file.name}'...")
            df = pd.read_excel(excel_file)
            print(f"✅ Loaded {len(df)} course rows.")
    
            print("📸 Generating catalog snapshot structure...")
            snapshot = build_snapshot(
                df.where(pd.notna(df), None).to_dict('records'), 
                CATALOG_TERM, 
                CATALOG_TERM_LABEL
            )
    
            provenance_path = BASE_DIR / 'source-provenance.json'
            if provenance_path.exists():
                print("📜 Reading source provenance metadata...")
                provenance = json.loads(provenance_path.read_text())
            else:
                print("⚠️  No 'source-provenance.json' found. Proceeding with defaults.")
                provenance = {}
    
            metadata = {
                'schema_version': snapshot['schema_version'],
                'term': snapshot['term'],
                'term_label': snapshot['term_label'],
                'snapshot_id': snapshot['snapshot_id'],
                'generated_at': snapshot['generated_at'],
                'snapshot_path': f"catalog-snapshots/{snapshot['snapshot_id']}.json",
                'source_file_sha256': sha256(excel_file.read_bytes()).hexdigest(),
                'source_updated_at': None,
                'source_retrieved_at': provenance.get('source_retrieved_at'),
                'refresh_date': provenance.get('refresh_date'),
                'identity_issues': snapshot['identity_issues'],
            }
    
            # 1. Generate course card HTML string
            print("🎨 [1/3] Generating course HTML cards...")
            full_course_html = display_course_info(df)
    
            # 2. Build department dropdown options string
            print("🏢 [2/3] Building department dropdown menu options...")
            dept_options = build_dept_options(df)
    
            # 3. Replace all placeholders on template_text
            print("🧱 [3/3] Injecting cards, JavaScript, and metadata into HTML template...")
            final_html = (
                template_text.replace('{{FULL_TEXT_PLACEHOLDER}}', full_course_html)
                .replace('{{FILTER_JS_PLACEHOLDER}}', js_text)
                .replace('{{DEPT_OPTIONS_PLACEHOLDER}}', dept_options)
                .replace('{{CATALOG_METADATA_PLACEHOLDER}}', canonical_json(metadata).replace('<', '\\u003c'))
                .replace('{{CATALOG_SNAPSHOT_PLACEHOLDER}}', canonical_json(snapshot).replace('<', '\\u003c'))
                .replace('{{CATALOG_TERM_LABEL_PLACEHOLDER}}', html.escape(CATALOG_TERM_LABEL_SHORT))
            )
    
            print("💾 Publishing catalog snapshot to snapshot folder...")
            publish_snapshot(snapshot, BASE_DIR / 'catalog-snapshots')
    
            print("📝 Writing output file to 'courses.html'...")
            output_path = BASE_DIR / 'courses.html'
            temporary_output = output_path.with_suffix('.html.tmp')
            with open(temporary_output, 'w', encoding='utf-8') as f:
                f.write(final_html)
            temporary_output.replace(output_path)
    
            print("🏷️  Updating 'latest.json' metadata reference...")
            latest_path = BASE_DIR / 'catalog-snapshots' / 'latest.json'
            temporary_latest = latest_path.with_suffix('.json.tmp')
            temporary_latest.write_text(canonical_json(metadata) + '\n', encoding='utf-8')
            temporary_latest.replace(latest_path)
    
            print('🎉 Saved courses.html successfully using htmltemplate.html and filter.js!')
            if snapshot['identity_issues']:
                print(f"⚠️  Warning: {len(snapshot['identity_issues'])} ambiguous section identities retained in snapshot.")
    
        except Exception as e:
            print(f"❌ Error building courses.html: {e}")
            raise SystemExit(1) from e