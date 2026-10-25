# -*- coding: utf-8 -*-
"""
Created on Sat Oct 24 21:46:53 2026

@author: vhorowit with Gemini
"""
import re
import pandas as pd
import os


# Mapping table for teaching method codes
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

# Day code expansion dictionary
DAY_MAP = {'M': 'Mon', 'T': 'Tue', 'W': 'Wed', 'R': 'Thu', 'F': 'Fri'}


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
    if r.upper() in ['TBA', 'TBA TBA', 'TBA/TBA', 'N/A', 'TBA  TBA']:
        return ''
    return r


def clean_html(text):
    """Strip raw HTML tags for clean text rendering."""
    if not is_valid(text):
        return ''
    clean = re.sub(r'<br\s*/?>', ' ', str(text))
    clean = re.sub(r'<p[^>]*>', '', clean)
    clean = re.sub(r'</p>', ' ', clean)
    clean = re.sub(r'<[^>]+>', '', clean)
    return re.sub(r'\s+', ' ', clean).strip()


def clean_title(title_str):
    """Format titles nicely, preserving Roman numerals (e.g., 'Physics II Lab')."""
    if not is_valid(title_str):
        return ''
    s = str(title_str).strip()
    if not s.isupper():
        return s

    small_words = {
        'a',
        'an',
        'and',
        'as',
        'at',
        'but',
        'by',
        'for',
        'in',
        'nor',
        'of',
        'on',
        'or',
        'so',
        'the',
        'to',
        'up',
        'yet',
    }
    roman_pattern = (
        r'^(?i:M{0,4}(C[MD]|D?C{0,3})(X[CL]|L?X{0,3})(I[XV]|V?I{0,3}))$'
    )

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
    """Combines COMMENTS and Notes while stripping HTML tags and deduplicating

    overlapping substrings or identical sentences.
    """

    # 1. Clean HTML tags and gather non-empty strings
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

    # 2. Check if one text is entirely contained inside the other (ignoring trailing periods)
    c_norm = c_str.rstrip('.')
    n_norm = n_str.rstrip('.')

    if c_norm in n_norm:
        return n_str if n_str.endswith('.') else n_str + '.'
    if n_norm in c_norm:
        return c_str if c_str.endswith('.') else c_str + '.'

    # 3. Sentence-level deduplication
    sentences = []
    for text in raw_texts:
        chunks = re.split(r'(?<=[.!?])\s+', text)
        for chunk in chunks:
            chunk_s = chunk.strip()
            norm_chunk = chunk_s.rstrip('.')
            # Add sentence if it hasn't already been included
            if chunk_s and not any(
                norm_chunk == existing.rstrip('.') for existing in sentences
            ):
                sentences.append(chunk_s)

    return ' '.join(sentences)

def display_course_info(data, mode='html'):
    """Displays formatted course details for a DataFrame, Series, or list of rows."""
    if isinstance(data, pd.DataFrame):
        rows = [r.to_dict() for _, r in data.iterrows()]
    elif isinstance(data, pd.Series):
        rows = [data.to_dict()]
    elif isinstance(data, dict):
        rows = [data]
    elif isinstance(data, list):
        rows = [
            r.to_dict() if isinstance(r, pd.Series) else r for r in data
        ]
    else:
        raise TypeError(
            'Input must be a pandas DataFrame, Series, dict, or list of rows.'
        )

    if mode.lower() == 'html':
        html_cards = []

        for row in rows:
            subj = str(row.get('Subject Code', '')).strip()
            num_plus = str(
                row.get('Course Number+', row.get('Course Number', ''))
            ).strip()
            sec = str(row.get('Section', '')).strip()

            code_sec = f'{subj} {num_plus}'
            if is_valid(sec) and sec != 'nan':
                code_sec += f"-{sec.zfill(2) if sec.isdigit() else sec}"

            # Always assign full_title unconditionally
            raw_title = row.get('Course Name')
            if not is_valid(raw_title):
                raw_title = row.get('TITLE')

            title = clean_title(raw_title)
            full_title = f'{code_sec}: {title}' if title else code_sec

            # Subtitle Metadata Pills
            sub_items = []
            
            subj_name = row.get("Subject Name")
            if is_valid(subj_name):
                sub_items.append(str(subj_name).strip())
            elif is_valid(subj):
                sub_items.append(subj)

            cred = row.get('CREDIT')
            if is_valid(cred):
                c_val = clean_num(cred)
                sub_items.append(
                    f"{c_val} Credit{'s' if c_val != '1' else ''}"
                )
            
            tm = str(row.get('TEACHINGMETHOD', '')).strip()
            if is_valid(tm):
                method_name = TEACHING_METHODS.get(tm.upper(), tm)
                # Only append if the method name is valid and not None/empty
                if is_valid(method_name):
                    sub_items.append(method_name)

            wi = str(row.get('Writing-intensive', '')).strip()
            if is_valid(wi) and wi.upper() == 'W':
                sub_items.append('Writing Intensive')

            subtitle_str = ' &nbsp;•&nbsp; '.join(sub_items)

            # Schedule Assembly
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

            schedule_display = (
                ' &nbsp;|&nbsp; '.join(schedules) if schedules else 'TBA'
            )

            # Instructors
            instructors = []
            for i in range(1, 5):
                inst = str(row.get(f'INSTRUCTOR{i}', '')).strip()
                if is_valid(inst) and inst not in instructors:
                    instructors.append(inst)
            instructor_display = (
                ', '.join(instructors) if instructors else 'TBA'
            )

            # Capacity & Restrictions
            cap_parts = []
            max_cap = row.get('MAXCAP') or row.get('Sec Capacity')
            if is_valid(max_cap):
                cap_parts.append(f'{clean_num(max_cap)} Seats')

            '''
            restricted = str(row.get('RESTRICTED', '')).strip()
            if is_valid(restricted) and restricted.upper() in [
                'R',
                'YES',
                'TRUE',
                '1',
            ]:
                cap_parts.append('(Restricted)')
            '''

            seat_breakdown = []
            for yr, label in [
                ('Fr Cap', 'Fr'),
                ('So Cap', 'So'),
                ('Jr Cap', 'Jr'),
                ('Sr Cap', 'Sr'),
            ]:
                v = row.get(yr)
                if is_valid(v):
                    seat_breakdown.append(f'{label}: {clean_num(v)}')

            if seat_breakdown:
                cap_parts.append(f"[{', '.join(seat_breakdown)}]")

            capacity_display = ' '.join(cap_parts) if cap_parts else 'N/A'

            # Cross-listing
            cross_list = row.get(
                'Cross Listed Courses Narrative'
            ) or row.get('Cross Listed Courses')
            cross_display = (
                str(cross_list).strip().strip(',')
                if is_valid(cross_list)
                else None
            )

            # Comments & Notes
            comments = str(row.get('COMMENTS', '')).strip()
            notes = str(row.get('Notes', '')).strip()
            comment_parts = []
            if is_valid(comments):
                comment_parts.append(comments)
            if is_valid(notes) and notes != comments:
                comment_parts.append(notes)
            # Comments & Notes
            comments_display = clean_comments_notes(row.get('COMMENTS'), row.get('Notes'))

            # Prerequisites & Corequisites
            prereq = row.get('Prerequisite Narrative') or row.get(
                'Prerequisites'
            )
            prereq_display = str(prereq).strip() if is_valid(prereq) else None

            # Build Card HTML
            card = [
                "<div style=\"font-family: 'Segoe UI', Arial, sans-serif; border: 1px solid #d0d7de; border-radius: 8px; padding: 18px 20px; margin-bottom: 20px; background-color: #ffffff; box-shadow: 0 2px 5px rgba(0,0,0,0.05);\">",
                "  <div style='border-bottom: 2px solid #002f6c; padding-bottom: 8px; margin-bottom: 14px;'>",
                f"    <h3 style='margin: 0; color: #002f6c; font-size: 19px; font-weight: 700;'>{full_title}</h3>",
            ]

            if subtitle_str:
                card.append(
                    f"    <div style='margin-top: 4px; font-size: 13px; color: #57606a; font-weight: 600;'>{subtitle_str}</div>"
                )

            card.append('  </div>')
            card.append(
                "  <div style='font-size: 14px; line-height: 1.8; color: #24292f;'>"
            )
            card.append(
                f'    <div>🕒 <b>Schedule:</b> {schedule_display}</div>'
            )
            card.append(
                f'    <div>👤 <b>Instructor:</b> {instructor_display}</div>'
            )
            card.append(
                f'    <div>👥 <b>Capacity:</b> {capacity_display}</div>'
            )

            if cross_display:
                card.append(
                    f'    <div>🔀 <b>Cross-listed:</b> {cross_display}</div>'
                )

            if prereq_display:
                card.append(
                    f'    <div>📋 <b>Prerequisites:</b> {prereq_display}</div>'
                )

            if comments_display:
                card.append(
                    f'    <div>💬 <b>Comments:</b> {comments_display}</div>'
                )

            card.append('  </div>')

            # Description Block
            desc = row.get('Course Description')
            if is_valid(desc):
                cleaned_desc = clean_html(desc)
                card.append("  <div style='margin-top: 14px;'>")
                card.append(
                    "    <div style='font-size: 12px; font-weight: 700; color: #57606a; text-transform: uppercase; letter-spacing: 0.5px; margin-bottom: 4px;'>📖 Description</div>"
                )
                card.append(
                    f"    <div style='font-size: 14px; color: #24292f; line-height: 1.6; background-color: #f6f8fa; border-left: 4px solid #002f6c; padding: 10px 14px; border-radius: 0 6px 6px 0;'>{cleaned_desc}</div>"
                )
                card.append('  </div>')

            card.append('</div>')
            html_cards.append('\n'.join(card))

        full_html = '\n'.join(html_cards)

        # Attempt IPython display for Spyder/Jupyter
        try:
            from IPython.display import HTML, display

            display(HTML(full_html))
        except ImportError:
            pass

        return full_html

    # ==================== PLAIN TEXT CONSOLE MODE ====================
    else:
        text_blocks = []
        for row in rows:
            subj = str(row.get('Subject Code', '')).strip()
            num_plus = str(
                row.get('Course Number+', row.get('Course Number', ''))
            ).strip()
            sec = str(row.get('Section', '')).strip()

            raw_title = row.get('Course Name')
            if not is_valid(raw_title):
                raw_title = row.get('TITLE')
            title = clean_title(raw_title)

            header = f"=== {subj} {num_plus}-{sec}" + (
                f': {title}' if title else ''
            )
            lines = [header, '=' * len(header)]

            scheds = []
            for i in range(1, 5):
                d = format_days(row.get(f'DAY{i}'))
                st = clean_time(row.get(f'TIME{i}-start'))
                en = clean_time(row.get(f'TIME{i}-end'))
                rm = clean_room(row.get(f'ROOM{i}'))
                if st and en:
                    scheds.append(f"{d} {st}-{en}" + (f' @ {rm}' if rm else ''))

            lines.append(
                f"  • Schedule: {', '.join(scheds) if scheds else 'TBA'}"
            )
            lines.append(
                f"  • Instructor: {row.get('INSTRUCTOR1', 'TBA')}"
            )

            desc = row.get('Course Description')
            if is_valid(desc):
                lines.append(f'  • Description: {clean_html(desc)}')

            lines.append('')
            text_blocks.append('\n'.join(lines))

        full_text = '\n'.join(text_blocks)
        return full_text

#%%

if __name__ == "__main__":
    course_and_schedule_df = pd.read_excel(r'merged_course_schedule.xlsx')
    print('Loaded', len(course_and_schedule_df), 'course rows')
    
    # Pass a single row from your DataFrame
    #full_text = display_course_info(course_and_schedule_df.iloc[0]);
    
    #display_course_info(course_and_schedule_df[course_and_schedule_df['Subject Code'] == 'PHYS'])

    dept = "ANTHR"
    # 1. Generate the HTML string for your desired selection (e.g. Physics)
    #full_text = display_course_info(course_and_schedule_df[course_and_schedule_df['Subject Code'] == dept])
    
    full_text = display_course_info(course_and_schedule_df)
    
    # 2. Wrap in a clean HTML document structure so it opens properly in browsers
    html_document = f"""<!DOCTYPE html>
    <html lang="en">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Hamilton Course Export</title>
        <style>
            body {{
                font-family: 'Segoe UI', Arial, sans-serif;
                background-color: #f4f6f8;
                margin: 0;
                padding: 30px 15px;
            }}
            .container {{
                max-width: 860px;
                margin: 0 auto;
            }}
        </style>
    </head>
    <body>
        <div class="container">
            {full_text}
        </div>
    </body>
    </html>"""
    
    # 3. Save 
    filename = "courses.html"
    
    with open(filename, 'w', encoding='utf-8') as f:
        f.write(html_document)
    
    print('Successfully saved course cards to:',
          os.getcwd(),
          filename)