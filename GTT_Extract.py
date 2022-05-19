#!/usr/bin/env python
# coding: utf-8

# In[ ]:


#GTT Extract
import pandas as pd
import numpy as np
import datetime as dt
import time
import tkinter as tk
from tkinter import filedialog

root = tk.Tk()
root.withdraw()

print('initalizing.....for dramatic effect')
time.sleep(2)
print('importing custom notif extract')

#load GTT_notif extract
filename = filedialog.askopenfilename()
fTable = pd.read_excel(filename)
fTable['Title'] = fTable['Title'].astype(str)
fTable = fTable.rename(columns={'Title':'NotificationNumber'})
fTable = fTable[:-1]
fTable['NotifGrid'].loc[fTable['NotifGrid'] == 'San Jacinto'] = 'Eastern'
fTable = fTable[fTable['NotifGrid'].notna()]
fTable = fTable[~(fTable['NotifGrid'] == 'Distribution')]
#fTable = fTable[~(fTable['NotifGrid'] == 'TTC North West')]
#fTable = fTable[~(fTable['NotifGrid'] == 'TTC South East')]
fTable['Grid Pivot'] = fTable['NotifGrid']
fTable['Grid Pivot'].loc[fTable['NotifAOR'] == 'ROW'] = 'ROW'

print('importing GTT_Notif_Excep')
#load GTT_Excep
filename = filedialog.askopenfilename()
eTable = pd.read_excel(filename)
#FILTER OUT NON TCM and ROW
#break up exceptions and data cleanup into two seperate tables
eTable = eTable.dropna(subset = ['NotificationNumber'])
dTable = eTable[eTable['ExceptionType'] == 'Data Clean Up']
eTable = eTable[eTable['ExceptionType'] == 'Exception']

#load circuit ref table
#filename = 'Copy of 2021 Inspection Orders-v2.xlsx'
#cTable = pd.read_excel(filename)
#fTable = pd.merge(fTable,cTable[['Circuit FLOC','Description']],left_on = 'CircuitName',right_on = 'Circuit FLOC',how = 'left')

#data cleanup merge
print('merging data cleanup')
fTable = pd.merge(fTable,dTable[['NotificationNumber','ExceptionCategory','ExceptionStatus','IdentifiedDate','SubmittedBy','DuplicateNotification','ReviewDate','ClearDate','ExceptionComments','GridComments','Created']],on = 'NotificationNumber',how = 'left')
fTable = fTable.rename(columns={'ExceptionCategory':'DCU_Category','ExceptionStatus':'DCU_Status','IdentifiedDate':'DCU_IdentifiedDate','SubmittedBy':'DCU_SubmittedBy','ReviewDate':'DCU_ReviewDate','ClearDate':'DCU_ClearDate','ExceptionComments':'P&MD Comments','GridComments':'DCU_GridComments','Created_y':'Date DCU added to GTT'})

print('squeezing exceptions')
#squeeze exceptions into single row per notification
eTable = eTable.fillna('')
eTable['IdentifiedDate'] = pd.to_datetime(eTable['IdentifiedDate']).dt.strftime('%m/%d/%y')
eTable['IdentifiedDate'] = eTable['IdentifiedDate'].replace('NaT','')
eTable['ExpectedCompletionDate'] = pd.to_datetime(eTable['ExpectedCompletionDate']).dt.strftime('%m/%d/%y')
eTable['ExpectedCompletionDate'] = eTable['ExpectedCompletionDate'].replace('NaT','')
eTable['ReviewDate'] = pd.to_datetime(eTable['ReviewDate']).dt.strftime('%m/%d/%y')
eTable['ReviewDate'] = eTable['ReviewDate'].replace('NaT','')
eTable['ClearDate'] = pd.to_datetime(eTable['ClearDate']).dt.strftime('%m/%d/%y')
eTable['ClearDate'] = eTable['ClearDate'].replace('NaT','')

eTable = eTable.groupby(['NotificationNumber']).agg({'ExceptionCategory':'|'.join,'ExceptionSubCategory':'|'.join,'IdentifiedDate':'|'.join,'SubmittedBy':'|'.join,'ExpectedCompletionDate':'|'.join,'ReviewDate':'|'.join,'ClearDate':'|'.join,'GridComments':'|'.join}).reset_index()

print('adding fluff columns')
#Modify Schedule status
fTable['ScheduleStatus'].loc[(fTable['ScheduleDate'].dt.year == 2022)&(fTable['ScheduleDate'].dt.quarter == 1)] = 'Scheduled-Q1-2022'
fTable['ScheduleStatus'].loc[(fTable['ScheduleDate'].dt.year == 2022)&(fTable['ScheduleDate'].dt.quarter == 2)] = 'Scheduled-Q2-2022'
fTable['ScheduleStatus'].loc[(fTable['ScheduleDate'].dt.year == 2022)&(fTable['ScheduleDate'].dt.quarter == 3)] = 'Scheduled-Q3-2022'
fTable['ScheduleStatus'].loc[(fTable['ScheduleDate'].dt.year == 2022)&(fTable['ScheduleDate'].dt.quarter == 4)] = 'Scheduled-Q4-2022' 
fTable['ScheduleStatus'].loc[fTable['ScheduleDate'].dt.year > 2022] = 'Scheduled - 2023+'
fTable['ScheduleStatus'].loc[fTable['ScheduleDate'].dt.date < dt.date.today()] = 'Unscheduled'
fTable['ScheduleStatus'].loc[fTable['NotifCompletionDate'].notna()] = 'Field Complete'

#GTT link
fTable['GTT_Link'] = np.nan
gtt_url = 'https://apps.powerapps.com/play/0e9072fc-0f56-4aa8-8f7a-c197cc5bdcd1?tenantId=5b2a8fee-4c95-4bdc-8aae-196f8aacb1b6&&NotifLinkID='
fTable['GTT_Link'] = gtt_url + fTable['NotificationNumber'] 

#GO95 overall Status
eTable['GO95 Exception Status'] = np.nan
eTable['GO95 Exception Status'] = np.where(eTable['ExceptionSubCategory'].str.contains('GO95'),'GO95 Exception','Internal Exception')
eTable.to_excel('etest.xlsx',index = False)

#Req End Date Status
fTable['Req End Date Category'] = np.nan
#fTable['Req End Date Category'].loc[fTable['Notification_Required_End_Date'].dt.year == 2021] = '2021'
fTable['Req End Date Category'].loc[fTable['RequiredEndDate'].dt.date < dt.date(2022,7,1)] = 'Q1/Q2 2022'
fTable['Req End Date Category'].loc[fTable['RequiredEndDate'].dt.date >= dt.date(2022,7,1)] = 'Q3/Q4 2022'
fTable['Req End Date Category'].loc[fTable['RequiredEndDate'].dt.year > 2022] = '2023+'
fTable['Req End Date Category'].loc[fTable['RequiredEndDate'].dt.year < 2022] = 'Pre-2022'

#Goal Category
fTable['2022 Goal Category'] = ''
fTable['2022 Goal Category'].loc[fTable['RequiredEndDate'].dt.year > 2022] = 'Due 2023+'
fTable['2022 Goal Category'].loc[fTable['RequiredEndDate'].dt.year == 2022] = '2022 30 Day Goal'
fTable['2022 Goal Category'].loc[fTable['RequiredEndDate'].dt.year < 2022] = 'Pre-2022 Rollover'

#30 Day Opporutnity
fTable['30_Day_Opportunity'] = ''
fTable['30_Day_Opportunity'].loc[(fTable['2022 Goal Category'] == '2022 30 Day Goal')&(fTable['RequiredEndDate'].dt.date - dt.date.today() < dt.timedelta(days=30))] = 'Missed Opportunity'
fTable['30_Day_Opportunity'].loc[(fTable['2022 Goal Category'] == '2022 30 Day Goal')&(fTable['RequiredEndDate'].dt.date - dt.date.today() >= dt.timedelta(days=30))] = 'Pending Opporunity'

#Days after schedule
fTable['Sched Days until Compliance Date'] = ''
fTable['Sched Days until Compliance Date'] = fTable['RequiredEndDate'] - fTable['ScheduleDate']
fTable['Sched Days until Compliance Date'] = fTable['Sched Days until Compliance Date'].dt.days

#Schedule Pivot
fTable['Schedule 30Day Pivot'] = ''
fTable['Schedule 30Day Pivot'].loc[fTable['Sched Days until Compliance Date'] < 0] = 'Scheduled After Due Date'
fTable['Schedule 30Day Pivot'].loc[fTable['Sched Days until Compliance Date'] >= 0] = 'Scheduled < 30 Days From Due Date'
fTable['Schedule 30Day Pivot'].loc[fTable['Sched Days until Compliance Date'] > 30] = 'Scheduled > 30 Days From Due Date'

#break out exceptions into single columns
def splitadd(numCols,name):
    colnames = []
    for i in range(0,numCols):
        colnames.append(name + str(i+1))
    return colnames
          
print('breaking out exceptions and splicing columns')
#ignore the fact this can be a function i just don't care to figure it out anymore
#For each category, breakout into seperate columns based on which # exception for that notification
add = eTable['ExceptionCategory'].str.split('|',expand = True)
header = np.vectorize(splitadd)(len(add.columns),'ExceptionCategory').tolist()
add.columns = header
eTable = pd.merge(eTable,add,left_index = True,right_index = True)

add = eTable['ExceptionSubCategory'].str.split('|',expand = True)
header = np.vectorize(splitadd)(len(add.columns),'ExceptionSubCategory').tolist()
add.columns = header
eTable = pd.merge(eTable,add,left_index = True,right_index = True)

add = eTable['IdentifiedDate'].str.split('|',expand = True)
header = np.vectorize(splitadd)(len(add.columns),'IdentifiedDate').tolist()
add.columns = header
eTable = pd.merge(eTable,add,left_index = True,right_index = True)

add = eTable['SubmittedBy'].str.split('|',expand = True)
header = np.vectorize(splitadd)(len(add.columns),'SubmittedBy').tolist()
add.columns = header
eTable = pd.merge(eTable,add,left_index = True,right_index = True)

add = eTable['ExpectedCompletionDate'].str.split('|',expand = True)
header = np.vectorize(splitadd)(len(add.columns),'ExpectedCompletionDate').tolist()
add.columns = header
eTable = pd.merge(eTable,add,left_index = True,right_index = True)

add = eTable['ClearDate'].str.split('|',expand = True)
header = np.vectorize(splitadd)(len(add.columns),'ClearDate').tolist()
add.columns = header
eTable = pd.merge(eTable,add,left_index = True,right_index = True)

add = eTable['ReviewDate'].str.split('|',expand = True)
header = np.vectorize(splitadd)(len(add.columns),'ReviewDate').tolist()
add.columns = header
eTable = pd.merge(eTable,add,left_index = True,right_index = True)

add = eTable['GridComments'].str.split('|',expand = True)
header = np.vectorize(splitadd)(len(add.columns),'ExceptionComments').tolist()
add.columns = header
eTable = pd.merge(eTable,add,left_index = True,right_index = True)

#create max expect_comp_date
#countofExcep = eTable[eTable.columns[-1]]
#countofExcep = int(countofExcep.name[-1:])
#eCol = ['NotificationNumber']
#for i in range(1,countofExcep+1):
#    eCol.append('ExpectedCompletionDate'+str(i))
#ecd = eTable[eCol]
#ecd = ecd.fillna(np.nan)
#ecd.iloc[:,1:] = dt.datetime.strptime(ecd.iloc[:,1:],'%m-%d-%Y')
#ecd['max_expected_date'] = ecd.iloc[:,1:].max(axis = 1)

#display(ecd)

print('merging exception table')
#combine main table with exception table
eTable = eTable.sort_values(by = ['GO95 Exception Status'])
fTable = pd.merge(fTable,eTable,on='NotificationNumber',how='left')
fTable['GO95 Exception Status'] = fTable['GO95 Exception Status'].fillna('No GO95 Exception')
#drop grouped exception concat
fTable = fTable.drop(columns=['ExceptionCategory','ExceptionSubCategory','IdentifiedDate','SubmittedBy','ExpectedCompletionDate','ClearDate','ReviewDate','GridComments'])

print('shuffling columns')
#Select main table and first four exceptions
firstcols = ['NotificationNumber','NotifGrid','Grid Pivot','District','HighFireFlag','NotifFLOC','NotifAOR','NotifPriority','NotifShortText','ActivityText','CircuitName','WorkOrder','APN','PropertyType','GO95 Exception Status','ScheduleStatus','Req End Date Category','2022 Goal Category','30_Day_Opportunity','Sched Days until Compliance Date','Schedule 30Day Pivot','ScheduleDate','ScheduleNotes','RequiredEndDate','PastDueOrFuture','NotifLongText','sapEquipNumber','sapEquipStatus','sapPlannerGroup','NotifStatuses','PatrolmanAssigned','PatrolmanNotes','FLOCLat','FLOCLong','NotifCompletionDate','ConstructionResource','PorCorR','CapOrOM','Modified','Created_x','DCU_Category','DCU_Status','DCU_IdentifiedDate','DCU_SubmittedBy','DuplicateNotification','DCU_ReviewDate','DCU_ClearDate','P&MD Comments','DCU_GridComments','Date DCU added to GTT','ExceptionCategory1','ExceptionSubCategory1','IdentifiedDate1','SubmittedBy1','ExpectedCompletionDate1','ReviewDate1','ClearDate1','ExceptionComments1','ExceptionCategory2','ExceptionSubCategory2','IdentifiedDate2','SubmittedBy2','ExpectedCompletionDate2','ReviewDate2','ClearDate2','ExceptionComments2','ExceptionCategory3','ExceptionSubCategory3','IdentifiedDate3','SubmittedBy3','ExpectedCompletionDate3','ReviewDate3','ClearDate3','ExceptionComments3','ExceptionCategory4','ExceptionSubCategory4','IdentifiedDate4','SubmittedBy4','ExpectedCompletionDate4','ReviewDate4','ClearDate4','ExceptionComments4']

#add any additional exceptions at the end
lastcols = [col for col in fTable.columns if col not in firstcols]
fTable = fTable[firstcols+lastcols]

print('creating extract')
fTable.to_excel('GTT Extract.xlsx',index = False)
print('extract created')

print('done')


# In[ ]:




