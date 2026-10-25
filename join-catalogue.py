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

# Rename
courses_df = courses_df.rename(columns={'Instructor': 'Usual instructor'})

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
schedule_df['Course Number'] = extracted_keys[1].str.strip()

# 4. Perform a left join keeping all rows from schedule_df
merged_df = schedule_df.merge(
    courses_df, on=['Subject Code', 'Course Number'], how='left'
)

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

# Re-order the output
# 1. Specify the desired leading columns
first_cols = ['Subject Code', 'Course Number', 'Section', 
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

# Save the updated merged DataFrame
final_df.to_excel('merged_course_schedule_with_seats.xlsx', index=False)
print(f"Successfully joined seat information! Total rows: {len(final_df)}")
