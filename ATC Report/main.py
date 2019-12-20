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

data = pd.read_excel(rawdatafile, sheet_name='Raw Data', skiprows=8, usecols='A:C,E')#, converters = {'Time':pd.to_datetime})

data['Date'] = data.apply(lambda r : pd.datetime.combine(r['Date'],r['Time']),1)

data = data.drop('Time',1)

data['Date'] = data['Date'].apply(lambda tm: round_minutes(tm, 'down', 15))

data['Day'] =  data['Date'].apply(lambda dt: dt.day_name())

data['Period'] = data['Date'].apply(lambda dt: f'{dt.hour:02d}:{dt.minute:02d}')

data = data.drop('Date',1)

days = data['Day'].unique()

directions = data['Direction'].unique()

period_index =[f'{tm.hour:02d}:{tm.minute:02d}' for tm in pd.date_range(start='2020-01-01 00:00', end='2020-01-01 23:45', freq='15min')]

print(period_index)

print(type(data['Period'][0]))

with pd.ExcelWriter(r'test.xlsx') as writer:

    for col, direction in enumerate(directions):

        directiondata = data[data['Direction']==direction]        

        for rw, day in enumerate(days):            
            daydata = directiondata[directiondata['Day']==day]
            classed = pd.pivot_table(daydata, values='Direction', index=['Day','Period'], columns='Class', aggfunc='count').fillna(0)

            classed = classed.reindex(columns=[1,2,3,4,5,6,7,8,9,10,11,12], fill_value=0)

            """ classed = classed.unstack(level=0).reindex(period_index)
            print(classed.head())
            classed = classed.fillna(0)
             """#classed.stack('Day').sort_index()

            classed = classed.reindex(pd.MultiIndex.from_product([[day],period_index]), fill_value=0)

            if rw==0:
                print(classed.head())

            classed.to_excel(writer, sheet_name='DayData', float_format='%.2f', startrow=(rw*100)+10, startcol=col*15+2)

    classed = pd.pivot_table(data, values='Direction', index=['Day','Period'], columns='Class', aggfunc='count')
    classed.to_excel(writer, sheet_name='Data', float_format='%.2f')


