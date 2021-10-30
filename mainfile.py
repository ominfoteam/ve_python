import pandas as pd
from openpyxl import load_workbook
from datetime import date
import numpy as np

"""def addMonths(dateObj, num):
    return dateObj.setMonth(dateObj.getMonth() + num)"""

xl=pd.ExcelFile('Inputs.xlsx')

df= xl.parse('Sheet1')
premiumfrequency=df['Premium Frequency'][0]
modalpremium=df['Modal Premium'][0]
#premiumspaid=df['Modal Premium'][0]
perinstallment=modalpremium/12

#12

#8333.34

#100000
""""today=date.today()
rcd=df['RCD'][0]
totalpremiumpaidtilldate=0
premiumcount=0"""
numberofmonths=(pd.to_datetime('today')-pd.to_datetime(df['RCD'][0],format='%Y%m'))//np.timedelta64(1, 'M')+1
#numberofmonths=(pd.to_datetime(df['Last Premium Date'][0],format='%Y%m')-pd.to_datetime(df['RCD'][0],format='%Y%m'))//np.timedelta64(1, 'M')+1
"""for x in range[rcd , today]:
    premiumcount +=1
    totalpremiumpaidtilldate+=modalpremium
    plus_month_period = 3

    #calculate date + future period
    df['future_date'] = plus_month_period.astype("timedelta64[M]")"""
    



#numberofmonthstilldate=round(numberofdays/np.timedelta64(1,'M'),2)
#print(numberofmonthstilldate)
#121
totalamountpaidtilldate=perinstallment * numberofmonths

print("premium frequency "+str(premiumfrequency))
print("Annually "+str(modalpremium))
print("Per InstallMent "+str(round(perinstallment,2)))
print("No of months paid till now " + str(numberofmonths))
print("Final Amount "+str(round(totalamountpaidtilldate,2)))
#1008333.34


