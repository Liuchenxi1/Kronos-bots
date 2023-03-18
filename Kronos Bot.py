#!/usr/bin/env python
# coding: utf-8

# pip install selenium
# pip install powerbiclient #-> install the power BI
# pip install smartsheet-python-sdk #-> install the smart sheet

# In[43]:


from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time


# In[44]:


import smartsheet
import logging
import os


# In[45]:


#operating power BI from Jupyter notebooks. 
from powerbiclient import Report, models


# # auto polit the paycode file from the Kronos

# In[4]:


driver = webdriver.Chrome()
driver.get('https://ath04.prd.mykronos.com/authn/XUI/?realm=/cevalogistics_prd_01#login&goto=https%3A%2F%2Fcevalogistics.prd.mykronos.com%3A443%2Fwfd%2Fhome')
driver.maximize_window()
#put username in
#(driver, 10) means: Selenium will wait for a maximum of 10 seconds for an element matching the given criteria to be found

username = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idToken1"))).send_keys('-------')
password = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idToken2"))).send_keys('-----------')

#click "login"
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "loginButton_0"))).click()

time.sleep(6)

#click the 3 lines menu on the left 
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "navmenu-open-btn"))).click()

#click the dataviews&reports
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "585_category"))).click()

#click the data library
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "575_text"))).click()

#click the paycode summary TK
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "krn-slat-29"))).click()

#This part is all good! Do not delete anything. 


# In[5]:


#click the timeframe drop down list
time.sleep(10) ##The ptyhon runing to fast and need to get some time off. 

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ia.timeframe.selector_clock"))).click()

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ia.timeframe.selector_select_range"))).click()

#This part is all good! Do not delete anything. datae HAS to go liek as this


# In[6]:


#Pick the damn date!

weekstart = input("What the day you like to start? \nEnter form:** MMDDYY **")
weekend = input("What the day you like to end? \nEnter form:** MMDDYY **")


# In[7]:


time.sleep(2)

start_date_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "startDateTimeInput")))
start_date_field.send_keys(weekstart)

end_date_field = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "endDateTimeInput")))
end_date_field.send_keys(weekend)

time.sleep(5)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "tfsApplyButton"))).click()

#This part is all good! Do not delete anything. datae HAS to go liek as this


# In[8]:


#pick the all business units
time.sleep(2) #------> Need let the website to load, otherwise, it will crush.

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "hyperfindIcon"))).click()
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "selectLocations"))).click()

time.sleep(5) #------> Need let the website to load, otherwise, it will crush.
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "orgmap-select-all-link"))).click()

time.sleep(2)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mapSelectDone"))).click()

#This part is all good! Do not delete anything.


# In[9]:


time.sleep(8)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "com.kronos.dataview.share"))).click()
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "com.kronos.dataview.share.export"))).click()
#This part works, Do not delete anything!


# In[46]:


time.sleep(8) ##-> python too fast, need to slow down and let the file download. 

import pandas as pd
import numpy as np
import csv

timesheet = pd.read_csv('C:/Users/liuchen/Downloads/US Paycode summary TK.CSV', skiprows = 8)


# In[11]:


timesheet.head()


# In[12]:


timesheet.dtypes


# In[13]:


timesheet['Employee ID'].nunique()


# In[14]:


timesheet['Business Unit'].unique()


# In[15]:


timesheet['Employee Full Name'] = timesheet['Employee Full Name'].str.replace(r'\s+[A-Z]$', '', regex=True)
timesheet[['Last_Name', 'Name']] = timesheet['Employee Full Name'].str.split(', ', 2, expand=True)
timesheet['Employee Full Name'] = timesheet['Name']+" "+timesheet['Last_Name'].str.title()

#Name porivd is done!!! Do not change any!!


# In[16]:


timesheet['Employee Full Name']


# In[17]:


def fix_columns(df):
    for idx in range(len(df['Business Unit'])):
        #i = [] -> this one doesn't work becasue i is not list, it will change the unit names 
        #but shows error too: Index(...) must be called with a collection of some kind, None was passed
        # you will need to give the "idex" as list name
        if df['Business Unit'][idx] == 16800626:
            df['Business Unit'][idx] = 'NCR'
        elif df['Business Unit'][idx] == 16800681:
            df['Business Unit'][idx] = 'Ramsey'
        elif df['Business Unit'][idx] == 16800698:
            df['Business Unit'][idx] = 'Pentland'   
        elif df['Business Unit'][idx] == 16800699:
            df['Business Unit'][idx] = 'GP'     
        elif df['Business Unit'][idx] == 16800718:
            df['Business Unit'][idx] = 'Cabinet Health'
        elif df['Business Unit'][idx] == 16800720:
            df['Business Unit'][idx] = 'Sidewalk Labs'
        elif df['Business Unit'][idx] == 16800723:
            df['Business Unit'][idx] = 'Cybex'
        elif df['Business Unit'][idx] == 16800737:
            df['Business Unit'][idx] = 'Gathr'
        elif df['Business Unit'][idx] == 168780:
            df.drop(idx, inplace=True)
    return df


# In[18]:


fixed_timesheet = fix_columns(timesheet)
timesheet['Business Unit'].unique()


# In[19]:


timesheet = timesheet.drop(columns = ['Legal Entity','Product','Location','Last_Name','Name'])
timesheet['Paycode Name'].unique()


# In[20]:


paycode_to_delete = input("What paycode do you want to delete? (Enter 'done' to exit)""\U0001F612")
## -> this is the loop for delete no needed paycodes

while paycode_to_delete.lower() not in ['done']:
    rows_to_delete = timesheet[timesheet['Paycode Name'].str.lower() == paycode_to_delete.lower()].index
    timesheet.drop(rows_to_delete, inplace=True)
    paycode_to_delete = input("What paycode do you want to delete? (Enter 'done' to exit)""\U0001F612")

## this is all good, don't do anything


# In[21]:


timesheet['Paycode Name'].unique()


# hourworked = timesheet.grouupby(['Employee Full Name'])['Actual Total Hours (Include Corrections)'].sum()

# In[22]:


timesheet.to_csv(r'Z:/Analytics/Krons Employee hours Sum and B_S/xotimesheetox.csv')


# # Start to get the business swap from smart sheet info
# Can't do it like Kornos, Can't left click on the sheet to find the "ID"

# In[47]:


import smartsheet
import pandas as pd

# to read the smartsheet toktken in ptyhon
smart = smartsheet.Smartsheet('xmfhGE48v905xFQmWJ3tbb7uBrWV29NBdjLMv')

# Get the sheet by ID, to read the sheet
sheet = smart.Sheets.get_sheet(5852364868478852)

# Create a list of dictionaries representing the rows in the sheet
# This part is provid the values to readable!!!
data = []
for row in sheet.rows:
    row_data = {}
    for cell in row.cells:
        row_data[cell.column_id] = cell.value
    data.append(row_data)

# Convert the list of dictionaries to a pandas DataFrame
b_s = pd.DataFrame(data)

# View the resulting DataFrame, but the columns names are not clean
print(b_s.head())


# In[49]:


# change the names of columns 
b_s = b_s.rename(columns = {2028794436446084:'Date_moved',
                            2458906080372612:'Agancy',
                            5070412966061956:'Employee_ID',
                            5679954724710276:'Employee_Name',
                            3428154911025028:'Home_business',
                            7931754538395524:'Business_worked',
                            5117004771288964:'Lob/lob2',
                            1021014619514756:'Time_start',
                            2728113795819396:'Time_end',
                            2865204957603716:'Total_hours_worked',
                            7683539788425092:'Supervisor'})


# In[50]:


b_s.head()


# In[51]:


b_s = b_s.drop(['Employee_ID','Lob/lob2'],axis =1)


# In[52]:


weekstart = input("What the day you like to start? \nEnter form:** YYYY-MM-DD **")
weekend = input("What the day you like to end? \nEnter form:** YYYY-MM-DD **")


# In[53]:


b_s = b_s.loc[(b_s["Date_moved"] >= weekstart) | (b_s["Date_moved"] == weekend)]


# In[54]:


b_s.head(5)


# In[55]:


b_s['NCR -reglaur'],b_s['NCR -OT'],b_s['Ramsey -reglaur'],b_s['Ramsey -OT'],b_s['GP -reglaur'],b_s['GP -OT']=["","","","","",""]


# In[56]:


b_s['sidewalk_labs -reglaur'],b_s['sidewalk_labs -OT'],b_s['Cybex -reglaur'],b_s['Cybex -OT'],b_s['Gathr -reglaur'],b_s['Gathr  -OT']=["","","","","",""]


# In[57]:


b_s = b_s.loc[b_s["Agancy"] == 'CEVA']


# In[58]:


n = input("********Do you happy with the result?! ********\nyes or no? ")
while n not in ['yes']:
    n = input("Why not!? \nlet's do it again: Do you happy with the DAMN result OR NOT!? \n****************\nyes or no? ")
print("Alright ~\nAlright ~\nAlright ~","\N{winking face}")


# In[34]:


time.sleep(3)

b_s.to_excel(r'Z:/Analytics/Krons Employee hours Sum and B_S/Business Swap_clean.xlsx')
# export the files for now and it can fulfil by python automatically. 
# 不要自动填充，因为supervisor经常犯错，所以有些时间，BU还是需要手动来确认.


# In[59]:


##--> set up the bs names as dataframe.This part is important for next. 

bs_names = list(b_s['Employee_Name'].unique())
bs_names = pd.DataFrame(bs_names, columns=['name'])
bs_names


# In[61]:


#using these names to find the data in the timesheet.这个句子是可行的，然后发转一下
#bs_hours = bs_names['name'].isin(timesheet['Employee Full Name'])

#using these names to find the data in the timesheet.使用unique bs_names来找在timesheet里的value
timesheet_filtered = timesheet[timesheet['Employee Full Name'].isin(bs_names['name'])]

groupby_result = timesheet_filtered.groupby(['Employee Full Name', 'Paycode Name'])['Actual Total Hours (Include Corrections)'].agg('sum')

groupby_result


# In[41]:


print("NICE WORK!! NOW PUT THE NUMBERS ON THE SHEET!! MY FRIEND!")


# In[ ]:




