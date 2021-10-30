import pandas as pd
from openpyxl import load_workbook
from datetime import *

import numpy as np
import numpy_financial as npf

xl=pd.ExcelFile('Inputs.xlsx')

df= xl.parse('Sheet1')
premiumfrequency=df['Premium Frequency'][0]
modalpremium=df['Modal Premium'][0]
rcd=df['RCD'][0]
lastpremiumdate=df['Last Premium Date'][0]
totalpremiumpaidtilldate=0
premiumcount=0
today = datetime.strptime("06-30-2020", '%m-%d-%Y')

print("First Premium Date - "+ str(rcd.date())) 
premiumarrayforirr=[]

while rcd <= today:
    if premiumcount == 0:
        premiumarrayforirr.append(modalpremium * -1)

    else:
         premiumarrayforirr.append(modalpremium)
    premiumcount +=1
    totalpremiumpaidtilldate+=modalpremium
    rcd=rcd + pd.DateOffset(months=premiumfrequency)
    
        


r = npf.irr(premiumarrayforirr)  
print("Array - "+str(premiumarrayforirr))
print("IRR - "+str(round(r,2)))   
print("Last Premium Date - "+ str(rcd.date()))
print("Total Premium Count - " + str(premiumcount))
print("Total Premium Amount - Rs." + str(totalpremiumpaidtilldate))
