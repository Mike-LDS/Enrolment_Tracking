#!/usr/bin/env python3
'''
Enrolment Reporting Code - Updated for 2022/23
'''
import csv
import pandas as pd
import numpy as np
from openpyxl import load_workbook

#--------------------------------------------------------------------------------- One to One Fee Table Compiling ---------------------------------------------------------------------------------------#


programs = ['Explicit Instruction', 'Homework Support','RISE Now', 'RISE TEAM', 'LDS Access', 'RISE at School']
other_pro = ['KTEA-3']

locations = ['East Van', 'North Van', 'RISE at Home', 'LDS Access']
loc_convert = ['RISE at LC: East Van', 'RISE at LC: North Van', 'RISE at Home', 'LDS Access']

lessons = pd.DataFrame({'ID':[], 'First Name':[], 'Last Name':[], 'Program':[], 'Location':[], 'Hours':[], '$/hr from Family':[], 'Date':[]})
other_lessons = pd.DataFrame({'ID':[], 'First Name':[], 'Last Name':[], 'Program':[], 'Location':[], 'Hours':[], '$/hr from Family':[], 'Date':[]})

students = pd.DataFrame({'ID':[], 'First Name':[], 'Last Name':[], 'Start Date':[], 'Program':[], 'Location':[], 'Hours/Week':[], '$/hr from Family':[], 'CKNW Submitted Date':[],
                            'CKNW Approval Date':[], 'CKNW Expiry Date':[], 'CKNW Funding':[], 'Variety Submitted Date':[], 'Variety Approval Date':[], 'Variety Expiry Date':[], 'Variety Funding':[],
                            'AFU Submitted Date':[], 'AFU Expiry Date':[], 'AFU Funding':[], 'JP or other':[], 'Other Funding':[], 'Status':[]})

## Compiling Lessons Data
with open('appointments.csv', newline='') as csvfile: # Lesson Export form Sept 7, 2022 to June 24, 2023
    reader = csv.DictReader(csvfile)
    for row in reader:
        loc = ''
        pro = ''
        if row['recipient_1'] != 'LDS Admin':
            name = row['recipient_1'].split(' ',1)
            for i in range(len(programs)):
                if programs[i].lower() in row['topic'].lower():
                    pro = programs[i]
            for i in  range(len(locations)):
                if locations[i] in row['location']:
                    loc = loc_convert[i]
            if loc == '':
                loc = row['location']

            date = pd.to_datetime(row['start'], format='%d/%m/%Y %I:%M %p')
            if pro != '':
                
                line = pd.DataFrame({'ID':row['recipient_id_1'], 'First Name': name[0], 'Last Name': name[1], 'Program':pro, 'Location':loc, 'Hours':float(row['units_raw']),
                                 '$/hr from Family':float(row['charge_rate_1']), 'Date':date}, index=[0])
                lessons = pd.concat([lessons,line])
            else:
                for i in range(len(other_pro)):
                    if other_pro[i].lower() in row['topic'].lower():
                        pro = other_pro[i]
                if pro != '':
                    line = pd.DataFrame({'ID':row['recipient_id_1'], 'First Name': name[0], 'Last Name': name[1], 'Program':pro, 'Location':loc, 'Hours':float(row['units_raw']),
                                 '$/hr from Family':float(row['charge_rate_1']), 'Date':date}, index=[0])
                    other_lessons = pd.concat([other_lessons,line])
                else:
                    print(row['topic'])


## Compiling Data for Unique Students
start_date = pd.to_datetime('2022-09-07 00:00:00')
end_date = pd.to_datetime('2023-06-25 00:00:00')
today = pd.to_datetime(pd.Timestamp.today().date())
uniques = lessons['ID'].unique()


num = 0

for student in uniques:
    stu_data = lessons.loc[(lessons['ID']==student)]
    first_name = stu_data['First Name'].mode()[0]
    if ')' in stu_data['Last Name'].mode()[0]:
        last_name = (stu_data['Last Name'].mode()[0]).split(') ',-1)[1]
    else:
        last_name = stu_data['Last Name'].mode()[0]
    start = min(stu_data['Date'])                               # Start Date
    end = max(stu_data['Date'])                                 # Date of last lesson

    # Filter out students with no future lessons - Live
    filtered = stu_data[(stu_data['Date'] > today) & (stu_data['Date'] < end_date)]
    hrs_sum = filtered['Hours'].sum()                           # Sum of hours

    if hrs_sum != 0:
        status = 'Live'
        pro = filtered['Program'].mode()[0]
        loc = filtered['Location'].mode()[0]
        nx_lesson = min(filtered['Date'])                       # Date of next lesson
               
        # Determining Weeks of Lessons   
        if nx_lesson < pd.to_datetime('2022-12-18 00:00:00'):   # Starts before Christmas Break
            if end > pd.to_datetime('2023-03-26 00:00:00'):     # and ends after Spring Break
                weeks = ((abs(end - nx_lesson).days)/7) - 4
            if end > pd.to_datetime('2023-01-01 00:00:00'):     # and ends after Christmas Break, but before Spring Break
                weeks = ((abs(end - nx_lesson).days)/7) - 2
            else:                                               # and ends before Christmas Break
                weeks = ((abs(end - nx_lesson).days)/7)
        elif nx_lesson < pd.to_datetime('2023-03-12 00:00:00'): # Starts before Spring Break, but after Christmas
            if end > pd.to_datetime('2023-03-26 00:00:00'):     # and ends after Spring Break
                weeks = ((abs(end - nx_lesson).days)/7) - 2
            else:                                               # and ends before Spring Break
                weeks = ((abs(end - nx_lesson).days)/7)
        else:                                                   # Starts after Spring Break
            weeks = ((abs(end - nx_lesson).days)/7)

        if weeks < 1:
            weeks = 1

        hrs = hrs_sum/weeks                                 # Calcualate hours per week
        hrs = round(hrs * 2) / 2                            # Round to the nearest .5
        if hrs == 0:
            hrs = ''
            
        try:
            rate = filtered['$/hr from Family'].mode()[0]
        except:
            rate = ''

    else:
        pro = stu_data['Program'].mode()[0]
        loc = stu_data['Location'].mode()[0]
        hrs = ''
        rate = ''
        status = 'Dormant'
        
    line = pd.DataFrame({'ID':student, 'First Name':first_name, 'Last Name':last_name, 'Start Date':start , 'Program':pro, 'Location':loc, 'Hours/Week':hrs, '$/hr from Family':rate, 'Status':status}, index=[num])
    students = pd.concat([students,line])

    # Funding Adjustments
    if loc == 'LDS Access':
        students.loc[(students['ID'] == student, 'JP or other')] = 'Grant Funded'
    if pro == 'RISE Now':
        students.loc[(students['ID'] == student, 'JP or other')] = 'Sponsored'
    if pro == 'RISE at School' and loc == 'Thunderbird Elementary School':
        students.loc[(students['ID'] == student, 'JP or other')] = 'Sponsored'
    if first_name == 'St.':
        students.loc[(students['ID'] == student, '$/hr from Family')] = 27.5

    # Student Specific Adjustments
    if student == '1785364':
        students.loc[(students['ID'] == student, 'Program')] = 'RISE at School'
        students.loc[(students['ID'] == student, 'Location')] = 'KLEOS'
        students.loc[(students['ID'] == student, 'JP or other')] = 'Sponsored'
    if student == '1379956':
        students.loc[(students['ID'] == student, 'Program')] = 'RISE at School'
    if student == '1850528':
        students.loc[(students['ID'] == student, 'Program')] = 'RISE at School'
        
            
## Compiling Funding Data
funding = load_workbook('THIRD PARTY COVERAGE - 2022-23.xlsx', data_only=True)

for j in range(len(uniques)):
    student = students.loc[(students['ID'] == uniques[j])]
    
# AFU Funding
    funding_ws = funding['AFU']
    surname = funding_ws['B:B']
    firstname = funding_ws['C:C']
    expires = funding_ws['H:H']
    AFU_funds = funding_ws['J:J']
    for i in range(len(surname)):
        try:
            fn = (firstname[i].value).replace(' ','')
            sn = (surname[i].value).replace(' ','')
        except:
            fn = firstname[i].value
            sn = firstname[i].value
            
        if (student['First Name'].values[0].upper()).replace(' ','') == fn and (student['Last Name'].values[0].upper()).replace(' ','') == sn and student['AFU Funding'].isnull()[0]:
            try:
                students.loc[(students['ID'] == uniques[j], 'AFU Funding')] = int(AFU_funds[i].value)
                students.loc[(students['ID'] == uniques[j], 'AFU Expiry Date')] = expires[i].value
            except:
                students.loc[(students['ID'] == uniques[j], 'AFU Funding')] = 0

# CKNW Funding
    cknw_ws = funding['CKNW']
    surname = cknw_ws['B:B']
    firstname = cknw_ws['C:C']
    granted = cknw_ws['H:H']
    expires = cknw_ws['I:I']
    cknw_fund = cknw_ws['J:J']
    for i in range(len(surname)):
        try:
            fn = (firstname[i].value).replace(' ','')
            sn = (surname[i].value).replace(' ','')
        except:
            fn = firstname[i].value
            sn = firstname[i].value

        if(student['First Name'].values[0].upper()).replace(' ','') == fn and (student['Last Name'].values[0].upper()).replace(' ','') == sn and student['CKNW Funding'].isnull()[0]:
            students.loc[(students['ID'] == uniques[j],'CKNW Expiry Date')] = expires[i].value
            students.loc[(students['ID'] == uniques[j],'CKNW Approval Date')] = granted[i].value
            students.loc[(students['ID'] == uniques[j],'CKNW Funding')] = cknw_fund[i].value

# Variety Funding
    variety_ws = funding['VARIETY']
    surname = variety_ws['B:B']
    firstname = variety_ws['C:C']
    granted = variety_ws['H:H']
    expires = variety_ws['I:I']
    variety_fund = variety_ws['K:K']
    for i in range(len(surname)):
        try:
            fn = (firstname[i].value).replace(' ','')
            sn = (surname[i].value).replace(' ','')
        except:
            fn = firstname[i].value
            sn = firstname[i].value

        if(student['First Name'].values[0].upper()).replace(' ','') == fn and (student['Last Name'].values[0].upper()).replace(' ','') == sn and student['Variety Funding'].isnull()[0]:
            students.loc[(students['ID'] == uniques[j],'Variety Expiry Date')] = expires[i].value
            students.loc[(students['ID'] == uniques[j],'Variety Approval Date')] = granted[i].value
            students.loc[(students['ID'] == uniques[j],'Variety Funding')] = variety_fund[i].value

# JP Funding
    jp_ws = funding['JORDANS PRINCIPLE']
    surname = jp_ws['B:B']
    firstname = jp_ws['C:C']
    granted = jp_ws['G:G']
    expires = jp_ws['H:H']
    jp_fund = jp_ws['I:I']
    for i in range(len(surname)):
        try:
            fn = (firstname[i].value).replace(' ','')
            sn = (surname[i].value).replace(' ','')
        except:
            fn = firstname[i].value
            sn = firstname[i].value

        if(student['First Name'].values[0].upper()).replace(' ','') == fn and (student['Last Name'].values[0].upper()).replace(' ','') == sn and student['Other Funding'].isnull()[0]:
            students.loc[(students['ID'] == uniques[j],'JP or other')] = expires[i].value
            students.loc[(students['ID'] == uniques[j],'Other Funding')] = jp_fund[i].value

# Submitted Applications
    apps_ws = funding['Funding Applications']
    surname = apps_ws['B:B']
    firstname = apps_ws['C:C']
    funder = apps_ws['D:D']
    submitted = apps_ws['E:E']
    notes = apps_ws['H:H']
    for i in range(len(surname)):
        try:
            fn = (firstname[i].value).replace(' ','')
            sn = (surname[i].value).replace(' ','')
        except:
            fn = firstname[i].value
            sn = firstname[i].value

        if(student['First Name'].values[0].upper()).replace(' ','') == fn and (student['Last Name'].values[0].upper()).replace(' ','') == sn:
            if funder[i].value == 'CKNW':
                if submitted[i].value != None and submitted[i].value != ' ' and submitted[i].value != '' and str(notes[i].value).lower() != 'declined':
                    students.loc[(students['ID'] == uniques[j],'CKNW Submitted Date')] = submitted[i].value
                else:
                    students.loc[(students['ID'] == uniques[j],'CKNW Submitted Date')] = notes[i].value
            elif funder[i].value  == 'Variety':
                if submitted[i].value != None and submitted[i].value != ' ' and submitted[i].value != '':
                    students.loc[(students['ID'] == uniques[j],'Variety Submitted Date')] = submitted[i].value
                else:
                    students.loc[(students['ID'] == uniques[j],'Variety Submitted Date')] = notes[i].value
            elif funder[i].value  == 'AFU':
                if submitted[i].value != None and submitted[i].value != ' ' and submitted[i].value != '':
                    students.loc[(students['ID'] == uniques[j],'AFU Submitted Date')] = submitted[i].value
                else:
                    students.loc[(students['ID'] == uniques[j],'AFU Submitted Date')] = notes[i].value
    

with open('users.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['\ufeffID'] in uniques and 'SD 5/ DL' in row['Labels']:
            students.loc[(students['ID'] == row['\ufeffID'], 'Program')] = 'RISE at School'
            students.loc[(students['ID'] == row['\ufeffID'], 'Location')] = 'SD 5/DL'


live_students = students[(students['Status'] == 'Live')]
export = live_students.drop(columns=['Status'])
export.to_csv('OnetoOne_feeTable.csv', index=False)

#--------------------------------------------------------------------------- Student Information & Ernolment Table Compiling ----------------------------------------------------------------------------#

student_info = pd.DataFrame({'ID':[], 'First Name':[], 'Last Name':[],'Status':[], 'New/Returning':[],'Date of Birth':[], 'Grade at Sept 2022':[], 'Family ID':[], 'Parent/Guardian':[],
                             'Family Email':[], 'Address':[], 'City':[], 'Postal Code':[], 'School':[], 'Diagnosis':[], 'BC Designation':[], 'One-to-One: East Van':[], 'One-to-One: North Van':[],
                             'One-to-One: RISE at Home':[], 'One-to-One: LDS Access':[], 'One-to-One: Pipeline':[],'RISE at School':[], 'SLP':[], 'RISE TEAM':[], 'RISE Now':[], 'Spring Break Camps':[],
                             'Early RISErs: Fall':[], 'KTEA-3':[]})

pipe_info = pd.DataFrame({'ID':[], 'First Name':[], 'Last Name':[],'Status':[], 'New/Returning':[],'Date of Birth':[], 'Grade at Sept 2022':[], 'Family ID':[], 'Parent/Guardian':[],
                        'Family Email':[], 'Address':[], 'City':[], 'Postal Code':[], 'School':[], 'Diagnosis':[], 'BC Designation':[]})

er_info = pd.DataFrame({'ID':[], 'First Name':[], 'Last Name':[],'Status':[], 'New/Returning':[],'Date of Birth':[], 'Grade at Sept 2022':[], 'Family ID':[], 'Parent/Guardian':[],
                        'Family Email':[], 'Address':[], 'City':[], 'Postal Code':[], 'School':[], 'Diagnosis':[], 'BC Designation':[]})

# AddStudents from 1 to 1
students = students.reset_index()
num = 0
for index, student in students.iterrows():
    line = pd.DataFrame({'ID':student['ID'], 'First Name':student['First Name'], 'Last Name':student['Last Name'], 'Status':student['Status']}, index=[num])
    num = num + 1
    student_info = pd.concat([student_info,line])

for index, student in students.iterrows():
    if student['Status'] == 'Live':
        pro_status = 'Enrolled'
    else:
        pro_status = 'Discontinued'
        
    if student['Program'] == 'Explicit Instruction' and student['Location'] == 'RISE at LC: East Van':
        student_info.loc[(student_info['ID'] == student['ID'],'One-to-One: East Van')] = pro_status
    elif student['Program'] == 'Explicit Instruction' and student['Location'] == 'RISE at LC: North Van':
        student_info.loc[(student_info['ID'] == student['ID'],'One-to-One: North Van')] = pro_status
    elif student['Program'] == 'Explicit Instruction' and student['Location'] == 'RISE at Home':
        student_info.loc[(student_info['ID'] == student['ID'],'One-to-One: RISE at Home')] = pro_status
    elif student['Program'] == 'Explicit Instruction' and student['Location'] == 'LDS Access':
        student_info.loc[(student_info['ID'] == student['ID'],'One-to-One: LDS Access')] = pro_status
    elif student['Program'] == 'Homework Support' and student['Location'] == 'RISE at LC: East Van':
        student_info.loc[(student_info['ID'] == student['ID'],'One-to-One: East Van')] = pro_status
    elif student['Program'] == 'Homework Support' and student['Location'] == 'RISE at LC: North Van':
        student_info.loc[(student_info['ID'] == student['ID'],'One-to-One: North Van')] = pro_status
    elif student['Program'] == 'Homework Support' and student['Location'] == 'RISE at Home':
        student_info.loc[(student_info['ID'] == student['ID'],'One-to-One: RISE at Home')] = pro_status
    elif student['Program'] == 'RISE at School':
        student_info.loc[(student_info['ID'] == student['ID'],'RISE at School')] = pro_status
    elif student['Program'] == 'RISE TEAM':
        student_info.loc[(student_info['ID'] == student['ID'],'RISE TEAM')] = pro_status
    elif student['Program'] == 'RISE Now':
        student_info.loc[(student_info['ID'] == student['ID'],'RISE Now')] = pro_status
    else:
        print(student)

# Adding Other Programs w/ Lessons
other_lessons = other_lessons.reset_index()
other_uni = other_lessons['ID'].unique()
stu_uni = student_info['ID'].unique()

for student in other_uni:
    if student not in stu_uni:
        fn = other_lessons.loc[(other_lessons['ID'] == student)]['First Name'][0]
        ln = other_lessons.loc[(other_lessons['ID'] == student)]['Last Name'][0]
        line = pd.DataFrame({'ID':student, 'First Name':fn, 'Last Name':ln, 'Status':'Live'}, index=[num])
        student_info = pd.concat([student_info,line])
        num = num + 1
    for pro in other_lessons.loc[(other_lessons['ID'] == student)]['Program']:
        if pro == 'KTEA-3':
            student_info.loc[(student_info['ID'] == student,'KTEA-3')] = 'Enrolled'


# Openning Student Export
with open('users.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['\ufeffID'] in student_info['ID'].unique():
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'Date of Birth')] = row['Date of birth']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'Grade at Sept 2022')] = row['Academic Year']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'Family ID')] = row['Client ID']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'Parent/Guardian')] = row['Client Name']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'Family Email')] = row['Client Email']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'Address')] = row['Street Address']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'City')] = row['Town']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'Postal Code')] = row['Zipcode/Postcode']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'School')] = row['School']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'Diagnosis')] = row['Diagnosis']
            student_info.loc[(student_info['ID'] == row['\ufeffID'],'BC Designation')] = row['BC Designation']
                
        elif '2022/23 One-to-one Instruction' in row['Labels']:
            line = pd.DataFrame({'ID':row['\ufeffID'], 'First Name': row['First name'], 'Last Name': row['Last name'], 'Status': 'Prospect (Pipeline)', 'Date of Birth':row['Date of birth'],
                                 'Grade at Sept 2022':row['Academic Year'],'Family ID':row['Client ID'], 'Parent/Guardian':row['Client Name'], 'Family Email':row['Client Email'],
                                 'Address':row['Street Address'], 'City':row['Town'], 'Postal Code':row['Zipcode/Postcode'], 'School':row['School'], 'Diagnosis':row['Diagnosis'],
                                 'BC Designation':row['BC Designation'], 'One-to-One: Pipeline':'Applied'}, index=[num])
            pipe_info = pd.concat([pipe_info,line])

                                                     
        if '2022 Early RISErs - Fall' in row['Labels']:           
            line = pd.DataFrame({'ID':row['\ufeffID'], 'First Name': row['First name'], 'Last Name': row['Last name'], 'Date of Birth':row['Date of birth'],
                                 'Grade at Sept 2022': row['Academic Year'],'Family ID':row['Client ID'], 'Parent/Guardian':row['Client Name'], 'Family Email':row['Client Email'],
                                 'Address':row['Street Address'], 'City':row['Town'], 'Postal Code':row['Zipcode/Postcode'], 'School':row['School'], 'Diagnosis':row['Diagnosis'],
                                 'BC Designation':row['BC Designation']}, index=[0])
            er_info = pd.concat([er_info,line])

# Openning the Client Export
families = student_info['Family ID'].unique()
pipe_fam = pipe_info['Family ID'].unique()
er_families = er_info['Family ID'].unique()

with open('users (1).csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['\ufeffID'] in families:
            # Fill in missing Address information
            if student_info.loc[(student_info['Family ID'] == row['\ufeffID'])]['Postal Code'].values[0] == '':
                student_info.loc[(student_info['Family ID'] == row['\ufeffID'], 'Postal Code')] = row['Zipcode/Postcode']
            if student_info.loc[(student_info['Family ID'] == row['\ufeffID'])]['City'].values[0] == '':
                student_info.loc[(student_info['Family ID'] == row['\ufeffID'], 'City')] = row['Town']
            if student_info.loc[(student_info['Family ID'] == row['\ufeffID'])]['Address'].values[0] == '':
                student_info.loc[(student_info['Family ID'] == row['\ufeffID'], 'Address')] = row['Street Address']

            # Setting Dormant Students Dormant    
            if row['Status'] == 'Dormant':
                student_info.loc[(student_info['Family ID'] == row['\ufeffID'],'Status')] = 'Dormant'

        # Inputting 1:1 Pipeline Students
        if row['\ufeffID'] in pipe_fam and row['Status'] == 'Prospect (Pipeline)':
            student_info = pd.concat([student_info,pipe_info.loc[(pipe_info['Family ID'] == row['\ufeffID'])]])

        # Inserting Active ER Students
        if row['\ufeffID'] in er_families and row['Status'] == 'Live':
            student_info = pd.concat([student_info,er_info.loc[(er_info['Family ID'] == row['\ufeffID'])]])
            for item in er_info.loc[(er_info['Family ID'] == row['\ufeffID'])]['ID']:
                student_info.loc[(student_info['ID'] == item, 'Status')] = 'Live'
                student_info.loc[(student_info['ID'] == item, 'Early RISErs: Fall')] = 'Enrolled'
    

## Checking in Returning Students
historic_lessons = pd.read_csv(r'/Users/lds/Documents/Student Data/Student Statistics/lessons.csv')

start_date = pd.to_datetime('2021-09-01 00:00:00')
end_date = pd.to_datetime('2022-09-03 00:00:00')

filtered = historic_lessons[(pd.to_datetime(historic_lessons['DateTime']) > start_date) & (pd.to_datetime(historic_lessons['DateTime']) < end_date) & (historic_lessons['Status'] == 'Complete')]
returning = filtered['ID'].unique()
uniques = student_info['ID']

for student in uniques:
    if student in str(returning):
        student_info.loc[(student_info['ID'] == student,'New/Returning')] = 'Returning'
    else:
        student_info.loc[(student_info['ID'] == student,'New/Returning')] = 'New'
        

export = student_info.drop(columns=['Family ID'])
export.to_csv('Enrollment.csv', index=False)


#----------------------------------------------------------------------------------------- Family Mapping -----------------------------------------------------------------------------------------------#

map_data = pd.DataFrame({'Program':[], 'New/Returning':[], 'Status':[], 'Address':[], 'City':[], 'Country':[], 'Postal Code':[]}) 

uniques = students['ID'].unique()
num = 0
for student in uniques:
    if student_info.loc[(student_info['ID'] == student)]['Postal Code'].values[0] != '':
        pro = students.loc[(students['ID'] == student)]['Program'].values[0]
        if pro == 'Explicit Instruction' or pro == 'Homework Support':
            pro = students.loc[(students['ID'] == student)]['Location'].values[0]

        line = pd.DataFrame({'Program':pro, 'New/Returning':student_info.loc[(student_info['ID'] == student)]['New/Returning'].values[0],
                         'Status':students.loc[(students['ID'] == student)]['Status'].values[0],
                         'Address':student_info.loc[(student_info['ID'] == student)]['Address'].values[0].upper(),
                         'City':student_info.loc[(student_info['ID'] == student)]['City'].values[0].upper(),
                         'Country':'CANADA',
                         'Postal Code':student_info.loc[(student_info['ID'] == student)]['Postal Code'].values[0][0:3].upper()}, index=[num])
        num = num + 1
        map_data = pd.concat([map_data,line])

# Adding Pipeline Students
one_to_one_pipeline = student_info.loc[(student_info['One-to-One: Pipeline'] == 'Applied')]
for student in one_to_one_pipeline['ID'].unique():
    if student_info.loc[(student_info['ID'] == student)]['Postal Code'].values[0] != '':
        line = pd.DataFrame({'Program':'One-to-One: Pipeline', 'New/Returning':student_info.loc[(student_info['ID'] == student)]['New/Returning'].values[0],
                             'Status':'Pipeline',
                             'Address':student_info.loc[(student_info['ID'] == student)]['Address'].values[0].upper(),
                             'City':student_info.loc[(student_info['ID'] == student)]['City'].values[0].upper(),
                             'Country':'CANADA',
                             'Postal Code':student_info.loc[(student_info['ID'] == student)]['Postal Code'].values[0][0:3].upper()}, index=[num])
        num = num + 1
        map_data = pd.concat([map_data,line])

map_export = map_data.drop(columns=['Address'])
map_export.to_csv('Map_Data.csv', index=False)

## Google Map Update
map_ev = map_data.loc[(map_data['Program'] == 'RISE at LC: East Van') & (map_data['Status'] == 'Live')]
map_ev.to_csv('Map_EV.csv', index=False)

map_nv = map_data.loc[(map_data['Program'] == 'RISE at LC: North Van') & (map_data['Status'] == 'Live')]
map_nv.to_csv('Map_NV.csv', index=False)

map_ah = map_data.loc[(map_data['Program'] == 'RISE at Home') & (map_data['Status'] == 'Live')]
map_ah.to_csv('Map_AH.csv', index=False)

map_as = map_data.loc[(map_data['Program'] == 'RISE at School') & (map_data['Status'] == 'Live')]
map_as.to_csv('Map_AS.csv', index=False)

map_ac = map_data.loc[(map_data['Program'] == 'LDS Access') & (map_data['Status'] == 'Live')]
map_ac.to_csv('Map_AC.csv', index=False)

map_pi = map_data.loc[(map_data['Program'] == 'One-to-One: Pipeline')]
map_pi.to_csv('Map_PI.csv', index=False)

























