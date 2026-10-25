# -*- coding: utf-8 -*-
"""Created on Sat Oct 24 21:46:53 2026

@author: vhorowit with Gemini
"""

import html
import json
import os
import re
import pandas as pd

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
  if r.upper() in ['TBA', 'TBA TBA', 'TBA/TBA', 'N/A', 'TBA  TBA']:
    return ''
  return r


def clean_html(text):
  """Strip raw HTML tags and decode HTML entities (&rsquo;, &amp;) for clean text rendering."""
  if not is_valid(text):
    return ''
  text_unescaped = html.unescape(str(text))
  clean = re.sub(r'<br\s*/?>', ' ', text_unescaped)
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
  """Combines COMMENTS and Notes while stripping HTML tags and deduplicating overlapping sentences."""
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
  """Displays formatted course details for a DataFrame, Series, or list of rows."""
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

      is_wi = (
          'true'
          if str(row.get('Writing-intensive', '')).strip().upper() == 'W'
          else 'false'
      )

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

      schedule_display = (
          ' &nbsp;|&nbsp; '.join(schedules) if schedules else 'TBA'
      )

      instructors = []
      for i in range(1, 5):
        inst = str(row.get(f'INSTRUCTOR{i}', '')).strip()
        if is_valid(inst) and inst not in instructors:
          instructors.append(inst)
      instructor_display = (
          ', '.join(instructors) if instructors else 'TBA'
      )

      cap_parts = []
      max_cap = row.get('MAXCAP') or row.get('Sec Capacity')
      if is_valid(max_cap):
        cap_parts.append(f'{clean_num(max_cap)} Seats')

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

      cross_list = row.get('Cross Listed Courses Narrative') or row.get(
          'Cross Listed Courses'
      )
      cross_display = (
          str(cross_list).strip().strip(',') if is_valid(cross_list) else None
      )

      comments_display = clean_comments_notes(
          row.get('COMMENTS'), row.get('Notes')
      )

      prereq = row.get('Prerequisite Narrative') or row.get('Prerequisites')
      prereq_display = str(prereq).strip() if is_valid(prereq) else None

      card = [
          (
              f'<div class="course-card" data-subject="{subj}"'
              f' data-days="{days_str}" data-start-min="{min_start}"'
              f' data-end-min="{max_end}" data-wi="{is_wi}"'
              f' data-slots="{slots_json_escaped}" style="font-family:'
              " 'Segoe UI', Arial, sans-serif; border: 1px solid #d0d7de;"
              ' border-radius: 8px; padding: 18px 20px; margin-bottom: 20px;'
              ' background-color: #ffffff; box-shadow: 0 2px 5px'
              ' rgba(0,0,0,0.05);">'
          ),
          (
              "  <div style='border-bottom: 2px solid #002f6c; padding-bottom:"
              " 8px; margin-bottom: 14px;'>"
          ),
          (
              "    <h3 style='margin: 0; color: #002f6c; font-size: 19px;"
              f" font-weight: 700;'>{full_title}</h3>"
          ),
      ]

      if subtitle_str:
        card.append(
            "    <div style='margin-top: 4px; font-size: 13px; color:"
            f" #57606a; font-weight: 600;'>{subtitle_str}</div>"
        )

      card.append('  </div>')
      card.append(
          "  <div style='font-size: 14px; line-height: 1.8; color: #24292f;'>"
      )
      card.append(f'    <div>🕒 <b>Schedule:</b> {schedule_display}</div>')
      card.append(f'    <div>👤 <b>Instructor:</b> {instructor_display}</div>')
      card.append(f'    <div>👥 <b>Capacity:</b> {capacity_display}</div>')

      if cross_display:
        card.append(f'    <div>🔀 <b>Cross-listed:</b> {cross_display}</div>')

      if prereq_display:
        card.append(
            f'    <div>📋 <b>Prerequisites:</b> {prereq_display}</div>'
        )

      if comments_display:
        card.append(
            f'    <div>💬 <b>Comments:</b> {comments_display}</div>'
        )

      card.append('  </div>')

      desc = row.get('Course Description')
      if is_valid(desc):
        cleaned_desc = clean_html(desc)
        card.append("  <div style='margin-top: 14px;'>")
        card.append(
            "    <div style='font-size: 12px; font-weight: 700; color: #57606a;"
            ' text-transform: uppercase; letter-spacing: 0.5px;'
            " margin-bottom: 4px;'>📖 Description</div>"
        )
        card.append(
            "    <div style='font-size: 14px; color: #24292f; line-height:"
            ' 1.6; background-color: #f6f8fa; border-left: 4px solid #002f6c;'
            ' padding: 10px 14px; border-radius: 0 6px 6px'
            f" 0;'>{cleaned_desc}</div>"
        )
        card.append('  </div>')

      card.append('</div>')
      html_cards.append('\n'.join(card))

    full_html = '\n'.join(html_cards)
    return full_html
  else:
    text_blocks = []
    for row in rows:
      subj = str(row.get('Subject Code', '')).strip()
      num_plus = str(
          row.get('Course Number+', row.get('Course Number', ''))
      ).strip()
      sec = str(row.get('Section', '')).strip()

      raw_title = row.get('Course Name') or row.get('TITLE')
      title = clean_title(raw_title)

      header = f'=== {subj} {num_plus}-{sec}' + (f': {title}' if title else '')
      lines = [header, '=' * len(header)]

      scheds = []
      for i in range(1, 5):
        d = format_days(row.get(f'DAY{i}'))
        st = clean_time(row.get(f'TIME{i}-start'))
        en = clean_time(row.get(f'TIME{i}-end'))
        rm = clean_room(row.get(f'ROOM{i}'))
        if st and en:
          scheds.append(f'{d} {st}-{en}' + (f' @ {rm}' if rm else ''))

      lines.append(f"  • Schedule: {', '.join(scheds) if scheds else 'TBA'}")
      lines.append(f"  • Instructor: {row.get('INSTRUCTOR1', 'TBA')}")

      desc = row.get('Course Description')
      if is_valid(desc):
        lines.append(f'  • Description: {clean_html(desc)}')

      lines.append('')
      text_blocks.append('\n'.join(lines))

    full_text = '\n'.join(text_blocks)
    return full_text


# Standard template string using .replace() to avoid f-string brace conflicts
HTML_TEMPLATE = """<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Hamilton Course Schedule Explorer</title>
    <style>
        body {
            font-family: 'Segoe UI', Arial, sans-serif;
            background-color: #f4f6f8;
            margin: 0;
            padding: 20px 15px;
            color: #24292f;
        }
        .container {
            max-width: 900px;
            margin: 0 auto;
        }
        
        .filter-panel {
            background: #ffffff;
            border: 1px solid #d0d7de;
            border-radius: 8px;
            padding: 18px 22px;
            margin-bottom: 24px;
            box-shadow: 0 4px 12px rgba(0,0,0,0.06);
            position: sticky;
            top: 10px;
            z-index: 100;
        }
        .filter-header {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 14px;
            border-bottom: 1px solid #e1e4e8;
            padding-bottom: 8px;
        }
        .filter-header h2 {
            margin: 0;
            font-size: 18px;
            color: #002f6c;
        }
        .counter-badge {
            font-size: 13px;
            font-weight: 600;
            background-color: #002f6c;
            color: #ffffff;
            padding: 4px 10px;
            border-radius: 12px;
        }
        .filter-grid {
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 14px;
        }
        .filter-group {
            display: flex;
            flex-direction: column;
            gap: 4px;
        }
        .filter-group label {
            font-size: 12px;
            font-weight: 700;
            color: #57606a;
            text-transform: uppercase;
        }
        .filter-group input, .filter-group select {
            padding: 7px 10px;
            border: 1px solid #d0d7de;
            border-radius: 6px;
            font-size: 13px;
            outline: none;
        }
        .filter-group input:focus, .filter-group select:focus {
            border-color: #002f6c;
            box-shadow: 0 0 0 2px rgba(0,47,108,0.15);
        }
        .checkbox-group {
            display: flex;
            gap: 12px;
            align-items: center;
            height: 100%;
            padding-top: 18px;
        }
        .checkbox-group label {
            font-size: 13px;
            text-transform: none;
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 4px;
        }

        .conflict-section {
            margin-top: 16px;
            padding-top: 14px;
            border-top: 1px dashed #d0d7de;
        }
        .conflict-title {
            font-size: 13px;
            font-weight: 700;
            color: #002f6c;
            margin-bottom: 8px;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .day-checkboxes {
            display: flex;
            gap: 10px;
            align-items: center;
            font-size: 12px;
            font-weight: 600;
        }
        .day-checkboxes label {
            cursor: pointer;
            display: flex;
            align-items: center;
            gap: 2px;
            text-transform: none;
        }
        .add-block-row {
            display: flex;
            flex-wrap: wrap;
            gap: 10px;
            align-items: flex-end;
            margin-bottom: 10px;
        }
        .btn-add {
            padding: 7px 14px;
            background-color: #002f6c;
            color: white;
            border: none;
            border-radius: 6px;
            font-size: 13px;
            font-weight: 600;
            cursor: pointer;
            transition: background 0.2s;
        }
        .btn-add:hover {
            background-color: #001f48;
        }
        .busy-list {
            display: flex;
            flex-wrap: wrap;
            gap: 8px;
            margin-top: 8px;
        }
        .busy-chip {
            background: #ffebe9;
            border: 1px solid #ffc1c0;
            color: #cf222e;
            padding: 4px 10px;
            border-radius: 16px;
            font-size: 12px;
            font-weight: 600;
            display: flex;
            align-items: center;
            gap: 6px;
        }
        .busy-chip .remove-btn {
            cursor: pointer;
            font-weight: bold;
            color: #cf222e;
            border: none;
            background: none;
            padding: 0;
            font-size: 14px;
            line-height: 1;
        }

        .reset-btn {
            grid-column: 1 / -1;
            padding: 8px 14px;
            background-color: #f3f4f6;
            border: 1px solid #d0d7de;
            border-radius: 6px;
            cursor: pointer;
            font-size: 13px;
            font-weight: 600;
            color: #24292f;
            transition: background 0.2s;
            margin-top: 10px;
        }
        .reset-btn:hover {
            background-color: #e5e7eb;
        }
        .hidden {
            display: none !important;
        }
    </style>
</head>
<body>
    <div class="container">
        <!-- Interactive Control Bar -->
        <div class="filter-panel">
            <div class="filter-header">
                <h2>🔍 Course Filters</h2>
                <span class="counter-badge" id="courseCounter">Loading...</span>
            </div>
            
            <div class="filter-grid">
                <!-- Text Search -->
                <div class="filter-group">
                    <label>Search Keyword</label>
                    <input type="text" id="searchInput" placeholder="Course name, instructor, code...">
                </div>

                <!-- Exclude Before Time -->
                <div class="filter-group">
                    <label>Hide Starts Before</label>
                    <select id="timeBeforeSelect">
                        <option value="">-- No Restriction --</option>
                        <option value="510">9:00 AM</option>
                        <option value="600">10:00 AM</option>
                        <option value="660">11:00 AM</option>
                        <option value="720">12:00 PM (Noon)</option>
                        <option value="780">1:00 PM</option>
                    </select>
                </div>

                <!-- Exclude After Time -->
                <div class="filter-group">
                    <label>Hide Ends After</label>
                    <select id="timeAfterSelect">
                        <option value="">-- No Restriction --</option>
                        <option value="840">2:00 PM</option>
                        <option value="900">3:00 PM</option>
                        <option value="960">4:00 PM</option>
                        <option value="1020">5:00 PM</option>
                    </select>
                </div>

                <!-- Checkbox Toggles -->
                <div class="checkbox-group">
                    <label><input type="checkbox" id="wiOnly"> Writing Intensive Only</label>
                </div>
            </div>

            <!-- EXCLUDE BUSY TIME SLOTS (SCHEDULE CONFLICT FILTER) -->
            <div class="conflict-section">
                <div class="conflict-title">🚫 Exclude Specific Busy Time Blocks (Schedule Conflicts)</div>
                <div class="add-block-row">
                    <div class="filter-group">
                        <label>Days</label>
                        <div class="day-checkboxes">
                            <label><input type="checkbox" class="busy-day" value="M"> M</label>
                            <label><input type="checkbox" class="busy-day" value="T"> T</label>
                            <label><input type="checkbox" class="busy-day" value="W"> W</label>
                            <label><input type="checkbox" class="busy-day" value="R"> R</label>
                            <label><input type="checkbox" class="busy-day" value="F"> F</label>
                        </div>
                    </div>
                    <div class="filter-group">
                        <label>Start Time</label>
                        <input type="time" id="busyStartTime" value="10:00">
                    </div>
                    <div class="filter-group">
                        <label>End Time</label>
                        <input type="time" id="busyEndTime" value="10:50">
                    </div>
                    <button class="btn-add" onclick="addBusyBlock()">+ Exclude Time Block</button>
                </div>

                <!-- List of active busy blocks -->
                <div class="busy-list" id="busyList"></div>
            </div>

            <button class="reset-btn" onclick="resetFilters()">Reset All Filters</button>
        </div>

        <!-- Course Cards Output -->
        <div id="courseContainer">
            {{FULL_TEXT_PLACEHOLDER}}
        </div>
    </div>

    <!-- Client-Side Filtering Engine -->
    <script>
        let busyBlocks = [];

        function parseTimeStringToMinutes(timeStr) {
            if (!timeStr) return null;
            const parts = timeStr.split(':');
            if (parts.length < 2) return null;
            return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
        }

        function formatMinutesTo12H(minutes) {
            let h = Math.floor(minutes / 60);
            let m = minutes % 60;
            const ampm = h >= 12 ? 'PM' : 'AM';
            h = h % 12;
            if (h === 0) h = 12;
            const mStr = m < 10 ? '0' + m : m;
            return `${h}:${mStr} ${ampm}`;
        }

        function addBusyBlock() {
            const dayCheckboxes = document.querySelectorAll('.busy-day:checked');
            const selectedDays = Array.from(dayCheckboxes).map(cb => cb.value);

            const startStr = document.getElementById('busyStartTime').value;
            const endStr = document.getElementById('busyEndTime').value;

            const startMin = parseTimeStringToMinutes(startStr);
            const endMin = parseTimeStringToMinutes(endStr);

            if (selectedDays.length === 0) {
                alert('Please select at least one day (M, T, W, R, F).');
                return;
            }
            if (startMin === null || endMin === null || startMin >= endMin) {
                alert('Please enter a valid start and end time (start must be before end).');
                return;
            }

            const dayLabel = selectedDays.join('');
            const label = `${dayLabel} ${formatMinutesTo12H(startMin)} – ${formatMinutesTo12H(endMin)}`;

            busyBlocks.push({
                days: selectedDays,
                start: startMin,
                end: endMin,
                label: label
            });

            renderBusyChips();
            filterCourses();
        }

        function removeBusyBlock(index) {
            busyBlocks.splice(index, 1);
            renderBusyChips();
            filterCourses();
        }

        function renderBusyChips() {
            const container = document.getElementById('busyList');
            container.innerHTML = '';
            busyBlocks.forEach((block, idx) => {
                const chip = document.createElement('div');
                chip.className = 'busy-chip';
                chip.innerHTML = `🚫 ${block.label} <button class="remove-btn" onclick="removeBusyBlock(${idx})" title="Remove filter">&times;</button>`;
                container.appendChild(chip);
            });
        }

        function filterCourses() {
            const query = document.getElementById('searchInput').value.toLowerCase().trim();
            const hideBefore = parseInt(document.getElementById('timeBeforeSelect').value, 10);
            const hideAfter = parseInt(document.getElementById('timeAfterSelect').value, 10);
            const wiOnly = document.getElementById('wiOnly').checked;

            const cards = document.querySelectorAll('.course-card');
            let visibleCount = 0;

            cards.forEach(card => {
                let show = true;

                // 1. Keyword search
                if (query !== '') {
                    const cardText = card.innerText.toLowerCase();
                    if (!cardText.includes(query)) {
                        show = false;
                    }
                }

                // 2. Hide courses starting before selected time
                if (show && !isNaN(hideBefore)) {
                    const startMin = parseInt(card.getAttribute('data-start-min'), 10);
                    if (!isNaN(startMin) && startMin < hideBefore) {
                        show = false;
                    }
                }

                // 3. Hide courses ending after selected time
                if (show && !isNaN(hideAfter)) {
                    const endMin = parseInt(card.getAttribute('data-end-min'), 10);
                    if (!isNaN(endMin) && endMin > hideAfter) {
                        show = false;
                    }
                }

                // 4. Writing Intensive filter
                if (show && wiOnly) {
                    if (card.getAttribute('data-wi') !== 'true') {
                        show = false;
                    }
                }

                // 5. Exclude Schedule Conflicts (Busy Time Blocks)
                if (show && busyBlocks.length > 0) {
                    const slotsRaw = card.getAttribute('data-slots');
                    if (slotsRaw) {
                        try {
                            const courseSlots = JSON.parse(slotsRaw);
                            
                            for (let b of busyBlocks) {
                                for (let cSlot of courseSlots) {
                                    const cDays = cSlot.days.split('');
                                    const hasDayOverlap = cDays.some(d => b.days.includes(d));

                                    if (hasDayOverlap) {
                                        if (cSlot.start < b.end && cSlot.end > b.start) {
                                            show = false;
                                            break;
                                        }
                                    }
                                }
                                if (!show) break;
                            }
                        } catch (e) {
                            console.error('Error parsing course slots JSON:', e);
                        }
                    }
                }

                if (show) {
                    card.classList.remove('hidden');
                    visibleCount++;
                } else {
                    card.classList.add('hidden');
                }
            });

            document.getElementById('courseCounter').textContent = 
                `Showing ${visibleCount} of ${cards.length} courses`;
        }

        function resetFilters() {
            document.getElementById('searchInput').value = '';
            document.getElementById('timeBeforeSelect').value = '';
            document.getElementById('timeAfterSelect').value = '';
            document.getElementById('wiOnly').checked = false;
            
            document.querySelectorAll('.busy-day').forEach(cb => cb.checked = false);
            busyBlocks = [];
            renderBusyChips();

            filterCourses();
        }

        document.getElementById('searchInput').addEventListener('input', filterCourses);
        document.getElementById('timeBeforeSelect').addEventListener('change', filterCourses);
        document.getElementById('timeAfterSelect').addEventListener('change', filterCourses);
        document.getElementById('wiOnly').addEventListener('change', filterCourses);

        window.addEventListener('DOMContentLoaded', filterCourses);
    </script>
</body>
</html>"""


if __name__ == '__main__':
  course_and_schedule_df = pd.read_excel('merged_course_schedule.xlsx')
  print('Loaded', len(course_and_schedule_df), 'course rows')

  full_text = display_course_info(course_and_schedule_df)

  # Insert the course card HTML without f-string syntax collisions
  html_document = HTML_TEMPLATE.replace('{{FULL_TEXT_PLACEHOLDER}}', full_text)

  with open('courses.html', 'w', encoding='utf-8') as f:
    f.write(html_document)

  print('Saved courses.html successfully!')
