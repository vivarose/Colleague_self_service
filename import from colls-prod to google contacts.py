"""
2024, 2025
import from colls-prod to google contacts.py
Viva R. Horowitz
I use this code to import a list of students in a class 
into my google contacts.

To use it, 
1. Open the Class Roster in colls prod
2. Export -> Download csv
3. Edit the python code (folder = and file =) below to indicate where that csv file is on your computer.
4. Run the code. (You need to have software to run python scripts.)
5. It will prompt you for what you want to call the group of students and export a file. 
6. Open contacts.google.com and click "Import" and upload the new file.
7. When you want to email the class of students, type the tag in the To: field 
and all the names will populate.
"""

## save the csv file directly from 
## https://collss-prod.hamilton.edu/Student/Student/Faculty/FacultyNavigation/xxxxx
## to the Downloads folder (or wherever), 
## then run this code.

folder = r'C:\Users\vhorowit\Downloads' ## You'll need to set your own folder
file = 'section-rosters_PHYS-390W-01_8_2_2025_10_37 PM.csv'

## All user inputs are set now.

import pandas as pd
import os

## Request user gives group tag for Google Contacts.
print('The tag needs to be set each time to indicate the course year and number.')
print('The new file will also be saved to this filename.')
course = input(
    "Group tag in Google contacts (eg. 2022f-Phys290): ")
savename = course+'.csv'


df = pd.read_csv(os.path.join(folder,file), 
                 header=0)

if 'Name' in df.columns:
    df = df.rename({"Name": "Student Name"}, 
                   axis='columns')
if 'Advisee Preferred Email' in df.columns:
    df = df.rename({'Advisee Preferred Email' : 'Preferred Email'}, 
                   axis = 'columns')

if 'Class Level' not in df.columns:
    df['Class Level'] = None
    adviseeflag = True
else:
    adviseeflag = False


#display(df)

# Start editing that dataframe
df.loc[:,'Group Membership'] = [course] * len(df) 
df.loc[:,'E-mail 1 - Type'] = ['Hamilton'] * len(df)


# Extract names
firstnames = []
notes = []
names = []
names_and_year = []

for student,year in zip(df['Student Name'], df['Class Level']):
    if "cross-listed" in student:
        fullname, note = student.split(sep = ' cross-listed as ')
    else:
        fullname = student
        note = ''
    if ',' in fullname:
        firstname =  fullname.split(sep=',')[1].split()[0] 
    else:
        firstname = fullname.split(sep = ' ')[0] 
        if firstname[-1] == '.': ## if the student has a title, like "Ms."
            firstname = fullname.split(sep = ' ')[1] 
    names.append(fullname)
    if year is not None:
        yr = str(year)[-2] + str(year)[-1]
        names_and_year.append(fullname + " '" + yr)
    firstnames.append(firstname)
    notes.append(note)
df.loc[:,'Given Name'] = firstnames
df.loc[:,'Notes'] = notes
df.loc[:,'FullName'] = names
if adviseeflag:
    assert (len(df) == len(names))
    df.loc[:,'Name'] = names
else:
    df.loc[:,'Name'] = names_and_year



lastnames = []
for student in df['FullName']:
    if ',' in student:
        lastname =  student.split(',')[0]
    else:
        lastname =  student.rsplit(maxsplit=1)[-1]
    lastnames.append(lastname)
df.loc[:,'Family Name'] = lastnames

if not adviseeflag:
    df.loc[:,'Organization 1 - Title'] = \
        [("class of " + str(year)) for year in df['Class Level']]
    
    
#############
print("\nGmail filter string:")

gmail_filter_string = ''

for i in range(len(df)):
    if i > 0:
        gmail_filter_string = gmail_filter_string + ' OR '
    gmail_filter_string = gmail_filter_string + df.loc[i,'Preferred Email']
    
gmail_filter_string = 'to:(' + gmail_filter_string + ') OR from:(' + \
                        gmail_filter_string + ')'

print('\n', gmail_filter_string, '\n')
#############


## format the way google contacts likes
df = df.rename({"Preferred Email": "E-mail 1 - Value"}, 
               axis='columns')

## choose the relevant columns and save
if adviseeflag:
    df2 = df[['Name','Given Name','Family Name', 'Notes',  
              'Group Membership', 'E-mail 1 - Type', 'E-mail 1 - Value', 
              ]]
else:
    df2 = df[['Name','Given Name','Family Name', 'Notes', 'Organization 1 - Title', 
              'Group Membership', 'E-mail 1 - Type', 'E-mail 1 - Value', 
              ]]
df2.to_csv(os.path.join(folder,savename))

## Explain what to do with the exported file.
print('Exported file: ')
print(os.path.join(folder,savename))
print('Import this to contacts.google.com')

#print(df2)