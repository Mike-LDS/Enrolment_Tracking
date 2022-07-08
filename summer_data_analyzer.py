#!/usr/bin/env python3
'''
Summer Enrolment Reporting Code
'''
import csv
import pandas as pd
import numpy as np
from openpyxl import load_workbook

summer_programs = ['Summer Tutoring', 'RISE Now', 'RISE TEAM', 'Summer RISE Intensive', 'LDS Access']
locations = ['East Van', 'North Van', 'RISE @ Home', 'LDS Access']
loc_convert = ['RISE At LC: East Van', 'RISE At LC: North Van', 'RISE @ Home', 'LDS Access']

lessons = pd.DataFrame({'ID':[], 'First Name':[], 'Last Name':[], 'Program':[], 'Location':[], '1-to-1 Hours':[], '$/hr from Family':[], 'Date':[]})

students = pd.DataFrame({'ID':[], 'First Name':[], 'Last Name':[], 'Status':[], 'New/Returning':[], 'Program':[], 'Location':[], '1-to-1 Hours':[], '$/hr from Family':[], 'CKNW Approval Date':[], 'CKNW Expiry Date':[],
          'CKNW Funding':[], 'Variety Approval Date':[], 'Variety Expiry Date':[], 'Variety Funding':[], 'AFU Expiry Date':[], 'AFU Funding':[], 'JP or other':[], 'Other Funding':[],
            'Date of Birth':[], 'Grade @ Spet 2021':[], 'Family ID':[], 'Parent/Guardian':[], 'Family Email':[], 'Address':[], 'City':[], 'Postal Code':[], 'School':[], 'Diagnosis':[], 'BC Designation':[], 'Summer Camps - Live':[]})


## Compiling Lessons Data
with open('appointments.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['status'] != 'Cancelled':
            name = row['recipient_1'].split(' ',1)
            for i in range(len(summer_programs)):
                if summer_programs[i].lower() in row['topic'].lower():
                    pro = summer_programs[i]
            for i in  range(len(locations)):
                if locations[i] in row['location']:
                    loc = loc_convert[i]

            date = pd.to_datetime(row['start'], format='%d/%m/%Y %I:%M %p')
            
            line = pd.DataFrame({'ID':row['recipient_id_1'], 'First Name': name[0], 'Last Name': name[1], 'Program':pro, 'Location':loc, '1-to-1 Hours':float(row['units_raw']), '$/hr from Family':float(row['charge_rate_1']), 'Date':date}, index=[0])
            lessons = pd.concat([lessons,line])         

with open('2022_SummerCamps.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        if row['Status'] == 'Live':
            date = pd.to_datetime(row['Week'], format='%Y-%m-%d')
            line = pd.DataFrame({'ID':row['ID'], 'First Name': row['First Name'], 'Last Name':row['Last Name'], 'Program':'Summer Camps', 'Location':'RISE At LC: East Van', 'Date':date}, index=[0])
            lessons = pd.concat([lessons,line]) 

## Printing Total Enrolments
print('Summer Tutoring - East Van')
print(len(lessons[(lessons['Program']=='Summer Tutoring') & (lessons['Location']=='RISE At LC: East Van')]['ID'].unique()))
print('Summer Tutoring - North Van')
print(len(lessons[(lessons['Program']=='Summer Tutoring') & (lessons['Location']=='RISE At LC: North Van')]['ID'].unique()))
print('Summer Tutoring - RISE @ Home')
print(len(lessons[(lessons['Program']=='Summer Tutoring') & (lessons['Location']=='RISE @ Home')]['ID'].unique()))
print('LDS Access')
print(len(lessons[(lessons['Program']=='LDS Access') & (lessons['Location']=='LDS Access')]['ID'].unique()))
print('RISE Now')
print(len(lessons[(lessons['Program']=='RISE Now')]['ID'].unique()))
print('RISE Team')
print(len(lessons[(lessons['Program']=='RISE TEAM')]['ID'].unique()))
print('Summer RISE Intensive')
print(len(lessons[(lessons['Program']=='Summer RISE Intensive')]['ID'].unique()))
print('Summer Camps')
print(len(lessons[(lessons['Program']=='Summer Camps')]['ID'].unique()))


## Compiling Data for Unique Students
uniques = lessons['ID'].unique()
num = 0

for student in uniques:
    sc = ''
    stu_data = lessons.loc[(lessons['ID']==student)]
    first_name = stu_data['First Name'].mode()[0]
    last_name = stu_data['Last Name'].mode()[0]

    # Filtered Lesson Dates
    start_date = pd.to_datetime('2022-07-10 00:00:00')
    end_date = pd.to_datetime('2022-07-17 00:00:00')
    filtered = stu_data[(stu_data['Date'] > start_date) & (stu_data['Date'] < end_date)] 
    try:
        pro = filtered['Program'].mode()[0]
        loc = filtered['Location'].mode()[0]
        if pro == 'Summer Camps':
            sc = 1
    except:
        pro = stu_data['Program'].mode()[0]
        loc = stu_data['Location'].mode()[0]
        
    hrs = round(filtered['1-to-1 Hours'].sum(),1)
    
    if hrs == 0:
        hrs = ''
    elif pro == 'LDS Access':
        hrs = round(hrs, 0)

    try:
        rate = stu_data['$/hr from Family'].mode()[0]
    except:
        rate = ''
        
    line = pd.DataFrame({'ID':student, 'First Name':first_name, 'Last Name':last_name, 'Status':'Live' , 'Program':pro, 'Location':loc, '1-to-1 Hours':hrs, '$/hr from Family':rate, 'Summer Camps - Live':sc}, index=[num])
    students = pd.concat([students,line])

    if pro == 'LDS Access':
        students.loc[(students['ID'] == student, 'JP or other')] = 'Grant Funded'

## Compiling Data from Student Export
summer_labels = ['2022 Summer Tutoring', '2022 Summer Intensive Intervention', '2022 Summer Camps', 'Thunderbird Access - Summer 2022']
pro_convert = ['Summer Tutoring', 'Summer RISE Intensive', 'Summer Camps', 'LDS Access']

with open('users.csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        for i in range(len(summer_labels)):

            # Adding Students that do not have lesson in TC
            if summer_labels[i] in row['Labels'] and row['\ufeffID'] not in uniques:
                rate = ''
                if 'RISE @ Home' in row['Labels'] and pro_convert[i] == 'Summer Tutoring':
                    loc = 'RISE @ Home'
                elif 'North Vancouver' in row['Labels'] and pro_convert[i] == 'Summer Tutoring':
                    loc = 'RISE At LC: North Van'
                elif  'Thunderbird Access' in row['Labels']:
                    loc = 'LDS Access'
                else:
                    loc = 'RISE At LC: East Van'

                if '2021/22' in row['Labels'] or 'Spring Break Camps 2022' in row['Labels']:
                    cust = 'Returning'
                else:
                    cust = 'New'
                    
                line = pd.DataFrame({'ID':row['\ufeffID'], 'First Name':row['First name'], 'Last Name':row['Last name'], 'New/Returning':cust, 'Program':pro_convert[i], 'Location':loc, '$/hr from Family':rate,
                                     'Date of Birth':row['Date of birth'], 'Grade @ Spet 2021':row['Academic Year'], 'Family ID':row['Client ID'], 'Parent/Guardian':row['Client Name'], 'Family Email':row['Client Email'], 'Address':row['Street Address'],
                                     'City':row['Town'], 'Postal Code':row['Zipcode/Postcode'], 'School':row['School'], 'Diagnosis':row['Diagnosis'], 'BC Designation':row['BC Designation']}, index=[0])
                students = pd.concat([students,line])
                uniques = np.append(uniques, [row['\ufeffID']], axis=0)

            # Updating Information for Students with Lessons in TC     
            elif row['\ufeffID'] in uniques:
                if '2021/22' in row['Labels'] or 'Spring Break Camps 2022' in row['Labels']:
                    students.loc[(students['ID'] == row['\ufeffID'], 'New/Returning')] = 'Returning'
                else:
                    students.loc[(students['ID'] == row['\ufeffID'], 'New/Returning')] = 'New'

                students.loc[(students['ID'] == row['\ufeffID'], 'First Name')] = row['First name']
                students.loc[(students['ID'] == row['\ufeffID'], 'Last Name')] = row['Last name']
                students.loc[(students['ID'] == row['\ufeffID'], 'Date of Birth')] = row['Date of birth']
                students.loc[(students['ID'] == row['\ufeffID'], 'Grade @ Spet 2021')] = row['Academic Year']
                students.loc[(students['ID'] == row['\ufeffID'], 'Family ID')] = row['Client ID']
                students.loc[(students['ID'] == row['\ufeffID'], 'Parent/Guardian')] = row['Client Name']
                students.loc[(students['ID'] == row['\ufeffID'], 'Family Email')] = row['Client Email']
                students.loc[(students['ID'] == row['\ufeffID'], 'Address')] = row['Street Address']
                students.loc[(students['ID'] == row['\ufeffID'], 'City')] = row['Town']
                students.loc[(students['ID'] == row['\ufeffID'], 'Postal Code')] = row['Zipcode/Postcode']
                students.loc[(students['ID'] == row['\ufeffID'], 'School')] = row['School']
                students.loc[(students['ID'] == row['\ufeffID'], 'Diagnosis')] = row['Diagnosis']
                students.loc[(students['ID'] == row['\ufeffID'], 'BC Designation')] = row['BC Designation']

## Updating Client Status
with open('users (1).csv', newline='') as csvfile:
    reader = csv.DictReader(csvfile)
    for row in reader:
        family =  students.loc[(students['Family ID'] == row['\ufeffID'])]
        if not family.empty:
            for i in range(len(family)):
                stu_status = family['Status'].values[i]
                if stu_status != 'Live':
                    students.loc[(students['ID'] == family['ID'].values[i], 'Status')] = row['Status']
            
## Compiling Funding Data
funding = load_workbook('STUDENTS - 3RD PARTY COVERAGE - 2022.xlsx', data_only=True)

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
        
        if student['First Name'].values[0].upper() == fn and student['Last Name'].values[0].upper() == sn and student['AFU Funding'].isnull()[0]:
            try:
                students.loc[(students['ID'] == uniques[j], 'AFU Funding')] = int(AFU_funds[i].value)
                students.loc[(students['ID'] == uniques[j], 'AFU')] = expires[i].value
            except:
                students.loc[(students['ID'] == uniques[j], 'AFU Funding')] = 0

# CKNW Funding
    cknw_ws = funding['CKNW']
    surname = cknw_ws['B:B']
    firstname = cknw_ws['C:C']
    granted = cknw_ws['I:I']
    expires = cknw_ws['J:J']
    cknw_fund = cknw_ws['K:K']
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
            
export = students.drop(columns=['ID', 'Family ID'])
export.to_csv('test.csv', index=False)


                
