''' This module processes ATC Raw Data'''

import os
import numpy as np
import pandas as pd
from datetime import datetime, timedelta
from filedialog import getfile
from filedialog import getfolder
from time import strftime


def round_minutes(dt, direction, resolution):
    new_minute = (dt.minute // resolution + (1 if direction == 'up' else 0)) * resolution
    return dt + timedelta(minutes=new_minute - dt.minute)

def roundTime(tm):

    return tm.minute \

curdir = os.getcwd()

# get file with raw data
# rawdatafile = getfile(curdir, "Please select file with raw data")

# test file
rawdatafile = 'C:/Test/Projects/ATC Report/ATC Report Template -Completed Example.xlsm'

print(rawdatafile)

data = pd.read_excel(rawdatafile, sheet_name='Raw Data', skiprows=8, usecols='A:C,E')#, converters = {'Time':pd.to_datetime})

print(data.dtypes)
print(data.head())

data['Date'] = data.apply(lambda r : pd.datetime.combine(r['Date'],r['Time']),1)

print(data.dtypes)
print(data.head())

data = data.drop('Time',1)

print(data.dtypes)
print(data.head())

data['Date'] = data['Date'].apply(lambda tm: round_minutes(tm, 'down', 15))

print(data.dtypes)
print(data.head())

data['Day'] =  data['Date'].apply(lambda dt: dt.weekday_name)

print(data.dtypes)
print(data.head())

data['Period'] = data['Date'].apply(lambda dt: f'{dt.hour:02d}:{dt.minute:02d}')

print(data.dtypes)
print(data.head())

data = data.drop('Date',1)

print(data.dtypes)
print(data.head())

classed = pd.pivot_table(data, values='Direction', index=['Day','Period'], columns='Class', aggfunc='count')

#print(classed.head())


classed.to_excel("test.xlsx")

""" 
with pd.ExcelWriter(r'test.xlsx') as writer:
    classed.to_excel(writer, sheet_name='Data', float_format='%.2f', index=False)
 """