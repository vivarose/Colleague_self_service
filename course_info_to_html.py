# -*- coding: utf-8 -*-
"""
Created on Sat Oct 24 21:46:53 2026

@author: vhorowit with Gemini
"""

import html
import json
import os
import re
import pandas as pd

os.chdir(r'C:\Users\vhorowit\Documents\GitHub\Colleague_self_service')

TEACHING_METHODS = {
    'LEC': None,
    'SEM': 'Seminar',
    'LAB': 'Laboratory',
    'IND': 'Independent Study',
    'STU': 'Studio',
    'CON': 'Concert / Performance',
    'REC': 'Recitation',
    'PRAC': 'Practicum',
    'DISC': 'Discussion',
    'WKS': 'Workshop',
}

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
        html_cards = []

        for row in rows:
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

            prereq = row.get('Prerequisite Narrative') or row.get('Prerequisites')
            prereq_display = str(prereq).strip() if is_valid(prereq) else None
            has_prereq = 'true' if prereq_display else 'false'

            subj = str(row.get('Subject Code', '')).strip()
            num_plus = str(
                row.get('Course Number+', row.get('Course Number', ''))
            ).strip()
            sec = str(row.get('Section', '')).strip()

            code_sec = f'{subj} {num_plus}'
            if is_valid(sec) and sec != 'nan':
                code_sec += f"-{sec.zfill(2) if sec.isdigit() else sec}"

            raw_title = row.get('Course Name') or row.get('TITLE')
            title = clean_title(raw_title)
            full_title = f'{code_sec}: {title}' if title else code_sec

            sub_items = []
            subj_name = row.get('Subject Name')
            if is_valid(subj_name):
                sub_items.append(str(subj_name).strip())
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

            card = [
                (
                    f'<div class="course-card" data-subject="{subj}"'
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
                f"     <h3 style='margin: 0; color: #002f6c; font-size: 19px; font-weight: 700;'>{full_title}</h3>",
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

            card.append('</div>')
            html_cards.append('\n'.join(card))

        return '\n'.join(html_cards)


def load_file_content(filename):
    """Loads text content from a file in the current directory."""
    if not os.path.exists(filename):
        raise FileNotFoundError(f"Required file '{filename}' was not found in the current directory.")
    with open(filename, 'r', encoding='utf-8') as f:
        return f.read()


if __name__ == '__main__':
    excel_file = 'merged_course_schedule.xlsx'
    html_template_file = 'htmltemplate.html'
    js_file = 'filter.js'

    if not os.path.exists(excel_file):
        print(f"Error: Could not find '{excel_file}'. Run join-catalogue.py first.")
    else:
        try:
            template_text = load_file_content(html_template_file)
            js_text = load_file_content(js_file)

            df = pd.read_excel(excel_file)
            print(f'Loaded {len(df)} course rows from {excel_file}.')

            full_course_html = display_course_info(df)

            final_html = template_text.replace(
                '{{FULL_TEXT_PLACEHOLDER}}', full_course_html
            ).replace(
                '{{FILTER_JS_PLACEHOLDER}}', js_text
            )

            with open('courses.html', 'w', encoding='utf-8') as f:
                f.write(final_html)

            print('🎉 Saved courses.html successfully using htmltemplate.html and filter.js!')

        except Exception as e:
            print(f"Error building courses.html: {e}")