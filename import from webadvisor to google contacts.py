"""
2017 - 2022 
import from webadvisor to google contacts.py
Viva R. Horowitz
I use this code to import a list of students in a class 
into my google contacts.

To use it, 
1. Open the Class Roster in webadvisor.hamilton.edu (not the Photo roster)
2. Save the html file. 
3. Edit the python code (folder =) below to indicate where that html file is on your computer.
4. Run the code. (You need to have software to run python scripts.)
5. It will prompt you for what you want to call the group of students and export a file. 
6. Open contacts.google.com and click "Import" and upload the new file.
7. When you want to email the class of students, type the tag in the To: field 
and all the names will populate.
"""

print('The tag needs to be set each time to indicate the course year and number.')
print('The new file will also be saved to this filename.')
course = input(
    "Group tag in Google contacts (eg. 2022f-Phys290): ")

## save the html file directly from 
## webadvisor > Student Roster (not photo roster) to 
## the Downloads folder (or wherever), 
## then run this code.

folder = r'C:\Users\vhorowit\Downloads' ## You'll need to set your own folder
file = 'Class Roster.html'
savename = course+'.csv'

## All user inputs are set now.

import pandas as pd
import os

# SettingWithCopyWarning is angry and I don't care
pd.options.mode.chained_assignment = None 
# reference: https://towardsdatascience.com/how-to-suppress-settingwithcopywarning-in-pandas-c0c759bd0f10

# read_html returns a list of dataframes.
lt = pd.read_html(os.path.join(folder,file), 
                  match = "Access", skiprows=1, header=0)

# Until webadvisor changes its formatting, this is the table you want.
whichdataframe = 2  

# Pull out one large dataframe with all the information
df = lt[whichdataframe].dropna(how='all', axis='columns')

#display(df)

# Start editing that dataframe
df.loc[:,'Group Membership'] = [course] * len(df) 
df.loc[:,'E-mail 1 - Type'] = ['Hamilton'] * len(df)

firstnames =  [student.split(sep=',')[1].split()[0] 
               for student in df['Student']]
df.loc[:,'Given Name'] = firstnames
lastnames =  [student.split(',')[0] for student in df['Student']]
df.loc[:,'Family Name'] = lastnames

phones =  [num.replace('(CEL)','')  for num in df['Phone Number']]
df.loc[:,'Phone 1 - Value'] = phones

df.loc[:,'Organization 1 - Title'] = \
    [("class of " + str(year)) for year in df['Class']]

## format the way google contacts likes
df = df.rename({"E-mail Address": "E-mail 1 - Value", 'Student':'Name'}, 
               axis='columns')

## choose the relevant columns and save
df2 = df[['Name','Given Name','Family Name', 'Organization 1 - Title', 
          'Group Membership', 'E-mail 1 - Type', 'E-mail 1 - Value', 
          'Phone 1 - Value' ]]
df2.to_csv(os.path.join(folder,savename))

## Explain what to do with the exported file.
print('Exported file: ')
print(os.path.join(folder,savename))
print('Import this to contacts.google.com')

#display(df2.head())
