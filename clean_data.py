# -*- coding: utf-8 -*-
"""
Created on Sun Nov 22 12:27:37 2026

@author: vhorowit and gemini

A helper file for course_info_to_html, for regex stuff
"""
import html
import re
import pandas as pd
import sys
if sys.stdout.encoding.lower() != 'utf-8':
    sys.stdout.reconfigure(encoding='utf-8')

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


def clean_html(text, context="", verbose = True):
    """Strip raw HTML tags and decode HTML entities."""
    # I tried without this and it's really better with stripping.
    
    # Return an empty string immediately if the input is null, empty, or invalid
    if not is_valid(text):
        return ''
        
    # Convert HTML character entities (e.g., '&amp;', '&#39;') back into normal text characters
    text_unescaped = html.unescape(str(text))
    
    # Replace <br> or <br/> tags with a space so words on adjacent lines don't merge
    clean = re.sub(r'<br\s*/?>', ' ', text_unescaped)
    
    # Remove opening <p> tags entirely, including any attributes (like class or style)
    clean = re.sub(r'<p[^>]*>', '', clean) # this avoids an extra lineskip.
    
    # Replace closing </p> tags with a space to maintain separation between paragraphs
    clean = re.sub(r'</p>', ' ', clean)
    
    # Strip out any other remaining HTML tags (e.g., <b>, <a>, <span>)
    clean = re.sub(r'<[^>]+>', '', clean)
    
    # Fix missing space after a comma when followed immediately by a letter
    fixed_commas = re.sub(r',([a-zA-Z])', r', \1', clean)
    
    # Display a message whenever a comma spacing fix occurs
    if verbose and fixed_commas != clean:
        matches = re.findall(r'\b\w+,[a-zA-Z]\w*\b', clean)
        label = f"[{context}] " if context else ""
        print(f"🔧 Comma space fixed {label}-> {matches}")
        
    return re.sub(r'\s+', ' ', fixed_commas).strip()




def clean_title(title_str):
    """Format titles nicely, preserving Roman numerals and acronyms like U.S."""
    if not is_valid(title_str):
        return ''
    s = str(title_str).strip()
    if not s.isupper():
        return s

    small_words = {
        'a', 'an', 'and', 'as', 'at', 'but', 'by', 'for', 'in',
        'nor', 'of', 'on', 'or', 'so', 'the', 'to', 'up', 'yet',
    }
    preserve_upper = {'u.s.', 'u.s.a.'}
    roman_pattern = r'^(?=[MDCLXVI])(?i:M{0,4}(C[MD]|D?C{0,3})(X[CL]|L?X{0,3})(I[XV]|V?I{0,3}))$'

    words = s.split()
    cleaned = []
    for idx, w in enumerate(words):
        # Strip surrounding punctuation (e.g., parentheses, colons, commas) to inspect the core token
        match = re.match(r'^([^\w.]*)([\w.]+)([^\w.]*)$', w)
        if match:
            prefix, core, suffix = match.groups()
        else:
            prefix, core, suffix = '', w, ''

        core_lower = core.lower()

        if core_lower in preserve_upper:
            formatted = core.upper()
        elif re.match(roman_pattern, core) and len(core) > 1:
            formatted = core.upper()
        elif core_lower == 'i':
            formatted = 'I'
        elif idx > 0 and core_lower in small_words:
            formatted = core_lower
        else:
            formatted = core.capitalize()

        cleaned.append(f'{prefix}{formatted}{suffix}')

    return ' '.join(cleaned)


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



def _ensure_terminal_punct(text):
    text = text.strip()
    if text and text[-1] not in '.!?':
        return text + '.'
    return text


def clean_comments_notes(comments, notes):
    """Combines COMMENTS and Notes for display while deduplicating overlapping text."""
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
    c_norm, n_norm = c_str.rstrip('.!?'), n_str.rstrip('.!?')

    if c_norm in n_norm:
        return n_str
    if n_norm in c_norm:
        return c_str

    sentences = []
    for text in raw_texts:
        chunks = re.split(r'(?<=[.!?])\s+', text)
        for chunk in chunks:
            chunk_s = chunk.strip()
            norm_chunk = chunk_s.rstrip('.!?')
            if chunk_s and not any(
                norm_chunk == existing.rstrip('.!?') for existing in sentences
            ):
                sentences.append(_ensure_terminal_punct(chunk_s))

    return ' '.join(sentences)

def get_base_num(course_num):
    """Strips trailing 'L' or section modifiers to get the primary course number (e.g., '270L' -> '270')."""
    if not is_valid(course_num):
        return ''
    # Remove trailing L, W, or section suffixes
    clean = re.sub(r'[a-zA-Z]+$', '', str(course_num).strip())
    return clean


def get_course_level(course_num):
    """Extracts the first digit of the course number to determine level (e.g., '146W' -> '100')."""
    if not is_valid(course_num):
        return ''
    match = re.search(r'\d', str(course_num))
    if match:
        first_digit = match.group(0)
        return f'{first_digit}00'
    return ''