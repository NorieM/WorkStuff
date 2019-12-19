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

data['Day'] =  data['Date'].apply(lambda dt: dt.day_name())

data['Period'] = data['Date'].apply(lambda dt: f'{dt.hour:02d}:{dt.minute:02d}')

data = data.drop('Date',1)

days = data['Day'].unique()

print(days)

with pd.ExcelWriter(r'test.xlsx') as writer:

    for rw, day in enumerate(days):
        classed = pd.pivot_table(data[data['Day']==day], values='Direction', index=['Day','Period'], columns='Class', aggfunc='count').fillna(0)
        classed.to_excel(writer, sheet_name='DayData', float_format='%.2f', startrow=(rw*100)+10, startcol=2)

    classed = pd.pivot_table(data, values='Direction', index=['Day','Period'], columns='Class', aggfunc='count')
    classed.to_excel(writer, sheet_name='Data', float_format='%.2f')


