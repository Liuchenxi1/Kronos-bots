#!/usr/bin/env python
# coding: utf-8

# pip install selenium

# In[1]:


from selenium import webdriver

from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import time


# In[2]:


driver = webdriver.Chrome()
driver.get('https://ath04.prd.mykronos.com/authn/XUI/?realm=/cevalogistics_prd_01#login&goto=https%3A%2F%2Fcevalogistics.prd.mykronos.com%3A443%2Fwfd%2Fhome')
driver.maximize_window()
#put username in
#(driver, 10) means: Selenium will wait for a maximum of 10 seconds for an element matching the given criteria to be found

username = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idToken1"))).send_keys('xxxxxx')
password = WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "idToken2"))).send_keys('xxxxxx')

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


# In[3]:


#click the timeframe drop down list
time.sleep(10) ##The ptyhon runing to fast and need to get some time off. 

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ia.timeframe.selector_clock"))).click()

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "ia.timeframe.selector_select_range"))).click()

#datestart = input("What the day you like to start?") -> This part doesn't work yet

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "startDateTimeInput"))).send_keys('02122023')
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "endDateTimeInput"))).send_keys('02182023')
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "tfsApplyButton"))).click()

#This part is all good! Do not delete anything. datae HAS to go liek as this


# In[4]:


#pick the all business units
time.sleep(2) #------> Need let the website to load, otherwise, it will crush.

WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "hyperfindIcon"))).click()
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "selectLocations"))).click()

time.sleep(5) #------> Need let the website to load, otherwise, it will crush.
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "orgmap-select-all-link"))).click()

time.sleep(2)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "mapSelectDone"))).click()

#This part is all good! Do not delete anything.


# In[5]:


time.sleep(8)
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "com.kronos.dataview.share"))).click()
WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.ID, "com.kronos.dataview.share.export"))).click()
#This part works, Do not delete anything!


# In[6]:


time.sleep(8) ##-> python too fast, need to slow down and let the file download. 

import pandas as pd
import numpy as np
import csv

timesheet = pd.read_csv('C:/Users/liuchen/Downloads/US Paycode summary TK.CSV', skiprows = 8)


# In[7]:


timesheet.head()


# In[8]:


timesheet.dtypes


# In[9]:


timesheet['Employee ID'].nunique()


# In[10]:


timesheet['Employee ID'].unique()
#maybe I don't need this


# In[11]:


timesheet['Business Unit'].unique()


# In[12]:


timesheet['Employee Full Name'] = timesheet['Employee Full Name'].str.replace(r'\s+[A-Z]$', '', regex=True)
timesheet[['Last_Name', 'Name']] = timesheet['Employee Full Name'].str.split(', ', 2, expand=True)
timesheet['Employee Full Name'] = timesheet['Name']+" "+timesheet['Last_Name'].str.title()

#Name porivd is done!!! Do not change any!!


# In[13]:


timesheet['Employee Full Name']


# In[14]:


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


# In[15]:


fixed_timesheet = fix_columns(timesheet)
timesheet['Business Unit'].unique()


# In[17]:


timesheet = timesheet.drop(columns = ['Legal Entity','Product','Location','Last_Name','Name'])


# In[18]:


deletion = timesheet[(timesheet['Paycode Name'] == 'US Total Overtime') |  
                    (timesheet['Paycode Name'] == 'US Personal Unpaid') | 
                    (timesheet['Paycode Name'] == 'US Unplanned Absence')].index
timesheet.drop(deletion, inplace = True)


# In[19]:


timesheet.head(20)


# In[20]:


#timesheet.to_csv(r'C:/Users/liuchen/Desktop/timesheet_02232023.csv')


# In[ ]:




