# -*- coding: utf-8 -*-
"""
Created on Sat Oct 24 19:24:02 2026

@author: vhorowit with Gemini

Instructions:
    1) Download the course schedule from:
    https://www.hamilton.edu/admission/visiting/courses/printable-schedule 
    (be sure to select the desired semester)
    2) Download the catalogue descriptions from 
    https://iq3.smartcatalogiq.com/catalogs/hamilton-college/home
    Tools  ->  Export courses
    3) Set the directory below where you have the files.
    4) Run this.

"""
#%pip install openpyxl
import re
import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

os.chdir(r'C:\Users\vhorowit\Downloads')

# 1. Load the two files
schedule_df = pd.read_excel('courseschedule.xlsx')
courses_df = pd.read_csv('courseExport.csv')

# 2. Convert 'Course Number' in catalog to string for safe joining
courses_df['Course Number'] = courses_df['Course Number'].astype(str).str.strip()
courses_df['Subject Code'] = courses_df['Subject Code'].astype(str).str.strip()

#%% Clean up courses_df
courses_df = courses_df.rename(columns={'Instructor': 'Usual instructor'})

# 1. Split 'Course Number' values on '/' into lists (e.g., ['145', '245'])
courses_df['Course Number'] = (
    courses_df['Course Number'].astype(str).str.split('/')
)
# 2. Duplicate rows by exploding the list elements into individual rows
courses_df = courses_df.explode('Course Number').reset_index(drop=True)
# 3. Strip any extra whitespace from the new course numbers
courses_df['Course Number'] = courses_df['Course Number'].str.strip()


#%% Split time columns into start and end time

# 1. Identify all columns that start with 'TIME' followed by a number
time_cols = [
    col
    for col in schedule_df.columns
    if col.startswith('TIME') and col[4:].isdigit()
]

for col in time_cols:
    # Convert column to string dtype safely, then extract start and end times
    times = (
        schedule_df[col]
        .astype('string')
        .str.extract(r'^(?P<start>\S+)\s+(?P<end>\S+)$')
    )

    # Assign start and end columns directly
    schedule_df[f'{col}-start'] = times['start']
    schedule_df[f'{col}-end'] = times['end']

# Verify the results on TIME1
#print(schedule_df[['TIME1', 'TIME1-start', 'TIME1-end']].head())

# Drop original TIME columns if you no longer need them
schedule_df.drop(columns=time_cols, inplace=True)

#%% 3. Extract 'Subject Code' and 'Course Number' from the schedule's 'COURSE_NO' column
# Example: 'AFRST102 01' -> Subject: 'AFRST', Number: '102'
extracted_keys = schedule_df['COURSE_NO'].astype(str).str.extract(
    r'([A-Za-z]+)\s*(\d+[A-Za-z]*)'
)
schedule_df['Subject Code'] = extracted_keys[0].str.strip()
schedule_df['Course Number+'] = extracted_keys[1].str.strip()

# 1. Populate 'Writing-intensive' with 'W' if 'W' is present in 'Course Number+' (otherwise empty string '')
schedule_df['Writing-intensive'] = schedule_df['Course Number+'].astype(
    str
).apply(lambda x: 'W' if 'W' in x else '')

# 2. Remove 'W' from 'Course Number+' to create 'Course Number'
schedule_df['Course Number'] = (
    schedule_df['Course Number+']
    .astype(str)
    .str.replace('W', '', regex=False)
    .str.strip()
)

# 4. Perform a left join keeping all rows from schedule_df, using indicator=True
merged_df = schedule_df.merge(
    courses_df, on=['Subject Code', 'Course Number'], how='left', indicator=True
)

# Identify courses in schedule_df that were not found in courses_df
unmatched = merged_df[merged_df['_merge'] == 'left_only'][
    ['Subject Code', 'Course Number', 'COURSE_NO']
].drop_duplicates()

if not unmatched.empty:
    print(
        f"\n⚠️ WARNING: {len(unmatched)} course(s) in the schedule were not found in the course catalog:"
    )
    for _, row in unmatched.iterrows():
        print(
            f"  - {row['Subject Code']} {row['Course Number']} (Schedule Code: {str(row['COURSE_NO']).strip()})"
        )
else:
    print("\n All schedule courses were successfully matched in the catalog!")

# Remove the temporary merge indicator column
merged_df.drop(columns=['_merge'], inplace=True)

#%% Include class seat restrictions

#%pip install requests pypdf

import re
import io
import requests
import pandas as pd
from pypdf import PdfReader

# 1. Download the PDF directly from the Hamilton website
url = "https://www.hamilton.edu/offices/registrar/registration/class-year-seats"
headers = {"User-Agent": "Mozilla/5.0"}
response = requests.get(url, headers=headers)
response.raise_for_status()

# 2. Extract text from the PDF pages
pdf_file = io.BytesIO(response.content)
reader = PdfReader(pdf_file)

pdf_text = ""
for page in reader.pages:
    pdf_text += page.extract_text() + "\n"

# 3. Parse table rows from the PDF text using regex
# PDF table columns: Section Name | Short Title | Sec Capacity | Fr Cap | So Cap | Jr Cap | Sr Cap
# Example row: "AFRST-146W-01  BLACK WOMEN'S LIT RENAISSANCE  18  9  9  0  0"
pattern = re.compile(
    r'([A-Z]+-\w+-\d+)\s+(.*?)\s+(\d+|\d+\s+OPEN|OPEN)\s+(\d+|OPEN)\s+(\d+|OPEN)\s+(\d+|OPEN)\s+(\d+|OPEN)'
)

parsed_data = []
for match in pattern.finditer(pdf_text):
    sec_name, title, sec_cap, fr_cap, so_cap, jr_cap, sr_cap = match.groups()
    
    # Extract Subject Code, Course Number, and Section Number from Section Name (e.g. AFRST-146W-01)
    # Group 1: Subject (e.g. AFRST)
    # Group 2: Course Number (e.g. 146W)
    # Group 3: Section (e.g. 01)
    parts = sec_name.split('-')
    if len(parts) >= 3:
        subj = parts[0].strip()
        course_num = parts[1].strip()
        section = parts[2].strip()
        
        parsed_data.append({
            'Subject Code': subj,
            'Course Number': course_num,
            'Section': section,
            'Sec Capacity': sec_cap,
            'Fr Cap': fr_cap,
            'So Cap': so_cap,
            'Jr Cap': jr_cap,
            'Sr Cap': sr_cap
        })

seats_df = pd.DataFrame(parsed_data)

# 4. Ensure Section column exists on merged_df for exact section-level matching
if 'Section' not in merged_df.columns:
    # Extract section (the trailing digits) from COURSE_NO (e.g., 'AFRST102 01' -> '01')
    merged_df['Section'] = merged_df['COURSE_NO'].astype(str).str.extract(r'(\d+)\s*$')[0].str.strip()

# Clean key columns in both DataFrames
for col in ['Subject Code', 'Course Number', 'Section']:
    merged_df[col] = merged_df[col].astype(str).str.strip()
    seats_df[col] = seats_df[col].astype(str).str.strip()

# 5. Left-join seat restriction information onto merged_df
final_df = merged_df.merge(
    seats_df[['Subject Code', 'Course Number', 'Section', 'Sec Capacity', 'Fr Cap', 'So Cap', 'Jr Cap', 'Sr Cap']],
    on=['Subject Code', 'Course Number', 'Section'],
    how='left'
)

#%% Determine Class Year Eligibility with Compact Audit Output

def determine_class_year_eligibility(row):
    """
    Evaluates eligibility ('ok' or 'No') for 1st, 2nd, 3rd, and 4th year students
    using explicit seat caps first, then falling back to text parsing.
    
    Ignores exclusion rules that target specific majors/concentrators (e.g., "Not open to Sr ECON concentrators")
    so general students in that class year remain marked as 'ok'.
    """
    year_cap_map = {
        '1st-years': row.get('Fr Cap'),
        '2nd-years': row.get('So Cap'),
        '3rd-years': row.get('Jr Cap'),
        '4th-years': row.get('Sr Cap')
    }
    
    status = {}
    has_explicit_caps = False
    
    # Step 1: Check explicit class seat caps from PDF
    for year_key, cap_val in year_cap_map.items():
        if pd.notna(cap_val) and str(cap_val).strip() != '':
            has_explicit_caps = True
            val_str = str(cap_val).strip().upper()
            try:
                num = float(val_str)
                status[year_key] = 'ok' if num > 0 else 'No'
            except ValueError:
                if 'OPEN' in val_str:
                    status[year_key] = 'ok'
                elif val_str == '0':
                    status[year_key] = 'No'
                else:
                    status[year_key] = 'ok'

    # If all explicit caps were found in PDF data, return immediately
    if has_explicit_caps and len(status) == 4:
        return pd.Series(status)

    # Step 2: Fall back to text parsing
    text_fields = {
        'COMMENTS': str(row.get('COMMENTS', '')),
        'Notes': str(row.get('Notes', '')),
        'Prerequisites': str(row.get('Prerequisites', '')),
        'Prerequisite Narrative': str(row.get('Prerequisite Narrative', '')),
        'Course Description': str(row.get('Course Description', ''))
    }
    
    raw_combined = ' '.join([v for v in text_fields.values() if v and v.lower() != 'nan'])
    cleaned = re.sub(r'<[^>]+>', ' ', raw_combined)
    cleaned = cleaned.replace('\xa0', ' ').replace('&nbsp;', ' ')
    combined_text = re.sub(r'\s+', ' ', cleaned).strip()
    
    if not combined_text:
        return pd.Series({'1st-years': 'ok', '2nd-years': 'ok', '3rd-years': 'ok', '4th-years': 'ok'})

    # Class year regex patterns (flexible punctuation & word boundaries)
    fr_pattern = r'(?:\bfirst[\s-]years?|\bfreshm[ae]n|\bfrs?\b\.?|\b1st\b)'
    so_pattern = r'(?:\bsophs?\b\.?|\bsophomores?|\bsos?\b\.?|\bsecond[\s-]years?|\b2nd\b)'
    jr_pattern = r'(?:\bjuniors?|\bjrs?\b\.?|\bthird[\s-]years?|\b3rd\b)'
    sr_pattern = r'(?:\bseniors?|\bsrs?\b\.?|\bfourth[\s-]years?|\b4th\b)'
    
    consent_pattern = r'(?:with|by)\s+(?:the\s+)?(?:instructor\'?s?\s+|department\'?s?\s+)?(?:consent|permission)'
    major_concentrator_pattern = r'\b(?:concentrators?|majors?|minors?)\b'
    
    text_status = {'1st-years': 'ok', '2nd-years': 'ok', '3rd-years': 'ok', '4th-years': 'ok'}
    matched_snippet = None
    
    pos_boundary = r'(?=\s*;\s*|\s*\n\s*|\.(?:\s+[A-Z]|\s*$)|(?:\s+not\s+open\s+to|\s+closed\s+to)|$)'
    excl_boundary = r'(?=\s*;\s*|\s*\n\s*|\.(?:\s+[A-Z]|\s*$)|(?:\s+open\s+to|\s+limited\s+to|\s+restricted\s+to|\s+must\s+|\bprerequisite:?)|$)'

    # 1. Positive restriction patterns
    positive_patterns = [
        r'(?:must\s+have|must\s+be|requires?|prerequisite:?)\s+([^.;\n]+?)\s+(?:standing|status|class\s+standing|class\s+year|class\s+of)',
        r'([^.;\n]+?)\s+(?:standing|status)\s+(?:required|necessary|needed)',
        r'\b(?<!not\s)(?<!never\s)open\s+to\s+(.*?)' + pos_boundary,
        r'(?:limited\s+to|restricted\s+to)\s+(.*?)' + pos_boundary,
        r'\b([^.;\n]+?)\s+only\b'
    ]
    
    for pat in positive_patterns:
        m = re.search(pat, combined_text, re.IGNORECASE)
        if m:
            group_str = m.group(1).lower()
            full_match = m.group(0)
            
            # Skip if 'only' / 'open to' is modified by consent/permission requirements
            if re.search(consent_pattern, full_match, re.IGNORECASE):
                continue
            
            has_fr = bool(re.search(fr_pattern, group_str, re.IGNORECASE))
            has_so = bool(re.search(so_pattern, group_str, re.IGNORECASE))
            has_jr = bool(re.search(jr_pattern, group_str, re.IGNORECASE))
            has_sr = bool(re.search(sr_pattern, group_str, re.IGNORECASE))
            
            if re.search(r'or\s+higher|or\s+above|and\s+above|\+', full_match, re.IGNORECASE):
                if has_so:
                    has_so = has_jr = has_sr = True
                elif has_jr:
                    has_jr = has_sr = True
                elif has_fr:
                    has_fr = has_so = has_jr = has_sr = True

            if has_fr or has_so or has_jr or has_sr:
                text_status['1st-years'] = 'ok' if has_fr else 'No'
                text_status['2nd-years'] = 'ok' if has_so else 'No'
                text_status['3rd-years'] = 'ok' if has_jr else 'No'
                text_status['4th-years'] = 'ok' if has_sr else 'No'
                matched_snippet = full_match
                break

    # 2. Exclusion patterns ("Not open to ...")
    not_open_match = re.search(
        r'(?:not\s+open\s+to|closed\s+to)\s+(.*?)' + excl_boundary, 
        combined_text, 
        re.IGNORECASE
    )
    if not_open_match:
        excluded_str = not_open_match.group(1).lower()
        
        # Guard: If the exclusion targets department majors/concentrators (e.g. "Sr ECON concentrators"),
        # general class-year eligibility remains 'ok' for that year.
        if not re.search(major_concentrator_pattern, excluded_str, re.IGNORECASE):
            has_excl = False
            if re.search(fr_pattern, excluded_str, re.IGNORECASE):
                text_status['1st-years'] = 'No'
                has_excl = True
            if re.search(so_pattern, excluded_str, re.IGNORECASE):
                text_status['2nd-years'] = 'No'
                has_excl = True
            if re.search(jr_pattern, excluded_str, re.IGNORECASE):
                text_status['3rd-years'] = 'No'
                has_excl = True
            if re.search(sr_pattern, excluded_str, re.IGNORECASE):
                text_status['4th-years'] = 'No'
                has_excl = True
                
            if has_excl and not matched_snippet:
                matched_snippet = not_open_match.group(0)

    # Step 3: Override text results with explicit PDF caps wherever available
    for key in ['1st-years', '2nd-years', '3rd-years', '4th-years']:
        if key in status:
            text_status[key] = status[key]

    # Step 4: Compact Audit Output
    if any(text_status[k] == 'No' for k in ['1st-years', '2nd-years', '3rd-years', '4th-years']):
        subj = str(row.get('Subject Code', '')).strip()
        num = str(row.get('Course Number+', row.get('Course Number', ''))).strip()
        course_id = f"{subj} {num}".strip()
        
        interpretation = ", ".join([f"{v}" for k, v in text_status.items()])
        snippet_display = matched_snippet if matched_snippet else combined_text
        
        print(f"🔍{course_id}", f"   \"{snippet_display}\"", f"   {interpretation}\n")

    return pd.Series(text_status)

# Apply evaluation function across DataFrame
year_cols = ['1st-years', '2nd-years', '3rd-years', '4th-years']
final_df[year_cols] = final_df.apply(determine_class_year_eligibility, axis=1)

# Apply evaluation function across DataFrame
year_cols = ['1st-years', '2nd-years', '3rd-years', '4th-years']
final_df[year_cols] = final_df.apply(determine_class_year_eligibility, axis=1)

#%% Re-order the output and save
# 1. Specify the desired leading columns
first_cols = ['Subject Code', 'Course Number+', 'Section', 
              'TITLE', 'Course Name',  'CREDIT',
              'DAY1', 'TIME1-start', 'TIME1-end',
              'DAY2', 'TIME2-start', 'TIME2-end',
              'DAY3', 'TIME3-start', 'TIME3-end',
              'TIME4-start', 'TIME4-end',
]
# 2. Collect all other existing columns in their original relative order
remaining_cols = [col for col in final_df.columns if col not in first_cols]
# 3. Combine and reorder the DataFrame
final_df = final_df[first_cols + remaining_cols]
# Verify the updated column order
#print(final_df.columns[:5].tolist())

# Save the updated merged DataFrame
final_df.to_excel('merged_course_schedule.xlsx', index=False)
print(f'Saved merged spreadsheet. Total rows: {len(final_df)}')
print(os.getcwd())

print(r'merged_course_schedule.xlsx');
