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
    3) Place the exports beside this script, or set CATALOG_INPUT_DIR.
    4) Run this.

"""
#%pip install openpyxl
import re
import numpy as np
import pandas as pd
import os
import warnings
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

os.chdir(os.environ.get('CATALOG_INPUT_DIR', os.path.dirname(os.path.abspath(__file__))))

# 1. Load the two files
schedule_df = pd.read_excel('courseschedule.xlsx')
courses_df = pd.read_csv('courseExport.csv')
try:
    prev_df = pd.read_pickle('merged_course_schedule.pkl')
except:
    prev_df = pd.DataFrame()

# 2. Convert 'Course Number' in catalog to string for safe joining
courses_df['Course Number'] = courses_df['Course Number'].astype(str).str.strip()
courses_df['Subject Code'] = courses_df['Subject Code'].astype(str).str.strip()


# Ensure 'Subject Name' column exists 
if 'Subject Name' not in courses_df.columns:
    courses_df['Subject Name'] = np.nan

courses_df = courses_df.rename(columns={'Subject Name': 'Subject Name+'})
courses_df['Subject Name'] = np.where(
    courses_df['Subject Code'] == 'MUSIC', 'Music', courses_df['Subject Name+']
)

#%% If registration is currently happening, then schedule_df is showing MAXCAP 
# only for current registering class

# Save the currently adjusted (one-class) capacity to a new column
schedule_df = schedule_df.rename(columns={'MAXCAP': 'current_capacity'})
# Check schedule for duplicates
duplicates = schedule_df[schedule_df.duplicated(subset=['COURSE_NO'], keep=False)]

if not duplicates.empty:
    print(f"⚠️ Found {len(duplicates)} rows of schedule_df with ambiguous section identities. Inspect these:")
    # Display the columns that help identify the duplicates
    cols_to_show = ['COURSE_NO']
    if 'Course Name' in schedule_df.columns:
        cols_to_show.append('Course Name')
    elif 'TITLE' in schedule_df.columns:
        cols_to_show.append('TITLE')
        
    print(duplicates[cols_to_show].sort_values(by=['Subject Code', 'Course Number+', 'Section']))

# (schedule_df[schedule_df['INSTRUCTOR1'] == 'Merrill H'])[['Course Number+','MAXCAP']]

# Merge the old 'Sec Capacity' into the new schedule_df
if not prev_df.empty and 'MAXCAP' in prev_df.columns:
    # Define the columns that uniquely identify a course section
    # (Adjust these if your raw Excel file uses different column names for the section ID)
    merge_keys = ['COURSE_NO']
    
    # Isolate only the keys and the column you want to pull over to avoid duplicating data
    capacity_df = prev_df[merge_keys + ['MAXCAP']]
    
    # Left merge to keep all new courses while appending the old capacities
    schedule_df = schedule_df.merge(capacity_df, on=merge_keys, how='left')
else:
    # Fallback in case the pickle file is missing or empty
    warnings.warn('Falling back')
    schedule_df['MAXCAP'] = schedule_df['current_capacity']

#%% Clean up courses_df

from course_info_to_html import is_valid
courses_df = courses_df.rename(columns={'Instructor': 'Usual instructor'})

# 1. Split 'Course Number' values on '/' into lists (e.g., ['145', '245'])
courses_df['Course Number'] = (
    courses_df['Course Number'].astype(str).str.split('/')
)
# 2. Duplicate rows by exploding the list elements into individual rows
courses_df = courses_df.explode('Course Number').reset_index(drop=True)
# 3. Strip any extra whitespace from the new course numbers
courses_df['Course Number'] = courses_df['Course Number'].str.strip()

courses_df = courses_df[courses_df['Course Description'].apply(is_valid)]

#%% Check courses_df for duplicates

cols_to_compare = ['Subject Code', 'Course Number']
duplicates_courses = courses_df[courses_df.duplicated(subset=cols_to_compare, keep=False)]
if not duplicates_courses.empty:
    print(f"⚠️ Found {len(duplicates)} rows of courses_df with ambiguous course identities. Inspect these:")
    # Display the columns that help identify the duplicates
    cols_to_show = cols_to_compare.copy()
        
    print(duplicates_courses[cols_to_show].sort_values(by=cols_to_compare))


#%% Add period to end of schedule_df['COMMNENTS']

# Create a mask for rows that are not null, not blank (even after stripping spaces), 
# and do not already end with a period.
mask = (
    schedule_df['COMMENTS'].notna() & 
    (schedule_df['COMMENTS'].str.strip() != '') & 
    ~schedule_df['COMMENTS'].str.strip().str.endswith('.', na=False)
)

# Apply the period only to the rows matching the mask
schedule_df.loc[mask, 'COMMENTS'] = \
    schedule_df.loc[mask, 'COMMENTS'].astype(str) + '.'

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
#schedule_df.drop(columns=time_cols, inplace=True)

#%% 3. Extract 'Subject Code' and 'Course Number' from the schedule's 'COURSE_NO' column
# Example: 'AFRST102 01' -> Subject: 'AFRST', Number: '102'
extracted_keys = schedule_df['COURSE_NO'].astype(str).str.extract(
    r'([A-Za-z]+)\s*(\d+[A-Za-z]*)'
)
# Now extracted_keys is a dataframe with two columns: the subject codes and 
# course number+
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

#%% Identify courses in schedule_df that were not found in courses_df
unmatched = merged_df[merged_df['_merge'] == 'left_only'][
    ['Subject Code', 'Course Number', 'COURSE_NO']
].drop_duplicates()

# Filter out lab sections (course numbers containing 'L')
unmatched = unmatched[
    ~unmatched['Course Number'].astype(str).str.contains('L', case=False, na=False)
]

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
if os.environ.get('CATALOG_SEATS_FILE'):
    with open(os.environ['CATALOG_SEATS_FILE'], 'rb') as seats_file:
        pdf_file = io.BytesIO(seats_file.read())
else:
    response = requests.get(url, headers=headers, timeout=60)
    response.raise_for_status()
    pdf_file = io.BytesIO(response.content)
reader = PdfReader(pdf_file)

pdf_text = ""
for page in reader.pages:
    pdf_text += page.extract_text() + "\n"

# 3. Parse table rows from the PDF text using regex
# PDF table columns: Section Name | Short Title | Sec Capacity | Fr Cap | So Cap | Jr Cap | Sr Cap
# Example row: "AFRST-146W-01  BLACK WOMEN'S LIT RENAISSANCE  18  9  9  0  0"
from seat_allocations import parse_allocations, merge_allocations

seats_df = pd.DataFrame(parse_allocations(pdf_text))
seats_df = seats_df.rename(columns={'Course Number': 'Course Number+'})

# 4. Ensure Section column exists on merged_df for exact section-level matching
if 'Section' not in merged_df.columns:
    # Extract section (the trailing digits) from COURSE_NO (e.g., 'AFRST102 01' -> '01')
    merged_df['Section'] = merged_df['COURSE_NO'].astype(str).str.extract(r'(\d+)\s*$')[0].str.strip()

final_df = merge_allocations(merged_df, seats_df)

#%% Check for discrepancies between MAXCAP and Sec Capacity

# 1. Safely convert 'Sec Capacity' to numeric (turns "" or text into NaN)
sec_cap_numeric = pd.to_numeric(final_df['Sec Capacity'], errors='coerce')

# 2. Create a mask checking for non-null values AND inequality
mask = (
    final_df['MAXCAP'].notna() & 
    sec_cap_numeric.notna() & 
    (final_df['MAXCAP'] != sec_cap_numeric)
)

check_me = final_df[mask].copy()

# Reorder columns to make reviewing easier
first_cols = ['Subject Code', 'Course Number+', 'Section', 'MAXCAP', 'Sec Capacity']

# Collect all other existing columns in their original relative order
remaining_cols = [col for col in check_me.columns if col not in first_cols]

# Combine and reorder the DataFrame
check_me = check_me[first_cols + remaining_cols]

# Print the DataFrame
print(check_me)

#%% Use the pdf values when available.
# Overwrite MAXCAP with Sec Capacity ONLY where they both exist and differ
final_df.loc[mask, 'MAXCAP'] = sec_cap_numeric[mask]

#%% Fix missing subject names
# 1. Convert empty strings or whitespace-only cells to NaN
final_df['Subject Name'] = final_df['Subject Name'].replace(r'^\s*$', np.nan, regex=True)

# 2. Fill missing Subject Name values using the known name for each Subject Code
final_df['Subject Name'] = final_df['Subject Name'].fillna(
    final_df.groupby('Subject Code')['Subject Name'].transform('first')
)


# Mask rows where Subject Code is 'CLNG' and Subject Name is NaN, empty, or whitespace
mask = (final_df['Subject Code'] == 'CLNG') & (
    final_df['Subject Name'].isna() | (final_df['Subject Name'].astype(str).str.strip() == '')
)

# Assign 'Critical Languages' to matching rows
final_df.loc[mask, 'Subject Name'] = 'Critical Languages'

## Set Subject Name+ (which is the same as Subject Name except for Music.)
# 1. Convert empty strings or whitespace-only cells in 'Subject Name+' to NaN
final_df['Subject Name+'] = final_df['Subject Name+'].replace(r'^\s*$', np.nan, regex=True)

# 2. Fill missing 'Subject Name+' values with the corresponding 'Subject Name'
final_df['Subject Name+'] = final_df['Subject Name+'].fillna(final_df['Subject Name'])

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

    # 500-Level Course Rule (Senior-only courses: 500–599)
    course_num = str(row.get('Course Number+') or row.get('Course Number') or row.get('COURSE') or '').strip()
    if re.search(r'\b5\d{2}\b', course_num):
        return pd.Series({'1st-years': 'No', '2nd-years': 'No', '3rd-years': 'No', '4th-years': 'ok'})
    
    # Check title rule (e.g., "Senior Thesis", "Senior Seminar", "Senior Project")
    raw_title = str(row.get('Course Name') or row.get('TITLE') or row.get('Title') or '')
    if re.search(r'\bseniors?\b', raw_title, re.IGNORECASE):
        return pd.Series({'1st-years': 'No', '2nd-years': 'No', '3rd-years': 'No', '4th-years': 'ok'})

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
        
        print(f"🔍{course_id}", f"    \"{snippet_display}\"", f"    {interpretation}\n")

    return pd.Series(text_status)

# Apply evaluation function across DataFrame
year_cols = ['1st-years', '2nd-years', '3rd-years', '4th-years']
final_df[year_cols] = final_df.apply(determine_class_year_eligibility, axis=1)

#%% Merge in URLs

if 'Catalogue URL' not in final_df.columns:
    url_df = pd.read_csv('URL_list.csv')
    
    # Deduplicate keys in url_df to prevent expanding final_df's row count
    url_df = url_df.drop_duplicates(subset=['Subject Code', 'Course Number+'])
    
    # Perform a left join to preserve all original rows in final_df
    final_df = final_df.merge(
        url_df, 
        on=['Subject Code', 'Course Number+'], 
        how='left'
    )

#%% Load historical seat availability

def load_prior_year_availability(
        csv_path=os.path.join("catalogue-history", 
                              "Hamilton_Courses_All_Pages_2026-spring.csv"),
        check_for_anomalies = True):
    """
    Parses Spring 2026 CSV and returns a dictionary mapping course section/base keys
    to historical enrollment stats.
    """
    df_history = pd.read_csv(csv_path)
    
    ## Parse the availability data
    # Define the new column names in order
    newcols = ['Open seats', 'Enrollment', 'Seats offered', 'Waitlist']
    
    # Split by '/' surrounded by optional whitespace, parse to numeric, and assign
    df_history[newcols] = (
        df_history['Availability']
        .str.split(r'\s*/\s*', expand=True)
        .apply(pd.to_numeric, errors='coerce')
        .astype('Int64')  # Keeps integer formatting while allowing NaN/missing values
    )
    
    if check_for_anomalies: # Check for anomalies in enrollment
        df = df_history.copy()
        open_col, enr_col, cap_col, _ = newcols

        if all(c in df.columns for c in [cap_col, enr_col, open_col]):
            df['Calculated_Capacity'] = df[open_col] + df[enr_col]
            df['Seat_Difference'] = df['Calculated_Capacity'] - df[cap_col]
            
            # Anomaly condition: absolute difference greater than zero
            anomalies = df[df['Seat_Difference'] != 0].copy()

            if anomalies.empty:
                print("✅ Open seats plus enrollment equals seats offered across all courses.")
            else:
                print(f"⚠️ Found {len(anomalies)} seat count anomalies:\n")
                
                # Columns to print for easy reading
                cols_to_print = [c for c in ['Subject Code', 'Course Number', 'Section', cap_col, enr_col, open_col, 'Calculated_Capacity', 'Seat_Difference'] if c in df.columns]
                print(anomalies[cols_to_print].to_string(index=False))
    
    # pars subject and course number
    df_history[['Subject Code', 'Course Number+', 'Section']] = (
        df_history['Section Name'].str.extract(r'^([A-Za-z]+)-([0-9]+[A-Za-z]*)-([A-Za-z0-9]+)')
)
    return df_history

course_keys = ['Subject Code', 'Course Number+']
df_hist = load_prior_year_availability()
# Aggregate historical data by course
hist_summary = (
    df_hist.groupby(course_keys)
    .agg(
        Hist_Sections_Offered=('Section', 'nunique'),
        Hist_Seats_Offered=('Seats offered', 'sum'),
        Hist_Total_Enrollment=('Enrollment', 'sum'),
        Hist_Total_Waitlist=('Waitlist', 'sum'),
        Hist_Open_Seats=('Open seats', 'sum'),
    )
    .reset_index()
)

# 3. Derive meaningful historical metrics
hist_summary['Hist_Fill_Rate'] = (
    hist_summary['Hist_Total_Enrollment'] / hist_summary['Hist_Seats_Offered']
).round(2)

hist_summary['Hist_Avg_Section_Enrollment'] = (
    hist_summary['Hist_Total_Enrollment']
    / hist_summary['Hist_Sections_Offered']
).round(1)

# 4. Merge aggregated history onto current offerings
final_df = pd.merge(final_df, hist_summary, on=course_keys, how='left')

final_df['Offered_Last_Year'] = final_df['Hist_Seats_Offered'].notna()
final_df['Hist_Status'] = final_df['Offered_Last_Year'].map(
    {True: 'Offered in Spring 2026', False: 'New / Not offered in Spring 2026'}
)


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

#%% Save the updated merged DataFrame
final_df.to_csv('merged_course_schedule.csv', index=False)
final_df.to_pickle('merged_course_schedule.pkl')

print(f'Saved merged spreadsheet. Total rows: {len(final_df)}')
print(os.getcwd())

print(r'merged_course_schedule.csv / pkl');


#%% Check for duplicates

df = merged_df

# Fill missing 'Course Number+' with 'Course Number' to match the script's logic
course_nums = df['Course Number+'].fillna(df.get('Course Number', ''))

# Create a temporary unique identifier for each section
df['section_id'] = df['Subject Code'].astype(str) + " " + course_nums.astype(str) + "-" + df['Section'].astype(str)

# Find all rows where this identifier is duplicated
duplicates = df[df.duplicated(subset=['section_id'], keep=False)]

if not duplicates.empty:
    print(f"⚠️ Found {len(duplicates)} rows with ambiguous section identities. Inspect these:")
    # Display the columns that help identify the duplicates
    cols_to_show = ['Subject Code', 'Course Number+', 'Section']
    if 'Course Name' in df.columns:
        cols_to_show.append('Course Name')
    elif 'TITLE' in df.columns:
        cols_to_show.append('TITLE')
        
    print(duplicates[cols_to_show].sort_values(by=['Subject Code', 'Course Number+', 'Section']))


