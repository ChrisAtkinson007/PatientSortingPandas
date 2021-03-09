import pandas as pd
import numpy as np
import datetime as dt

data = pd.read_excel (r'C:\Users\TheCretin\Desktop\Palm Tree Excel\Stats.xlsx')
df = pd.DataFrame(data, columns= ['Ref No','Admis Date','Disch date','Age'])
df["present"]='0'


startdate='2020-01-05'
enddate='2020-07-14'
randomdate='2020-07-01'
df2 = pd.DataFrame({"Date": pd.date_range(startdate, enddate)})
df2["Occ"]=0

for i in range(len(df2)):
  for j in range(len(df)):
      if df.loc[j,"Admis Date"]<=df2.loc[i,"Date"] and df2.loc[i,"Date"]<df.loc[j,"Disch date"] and int(df.loc[j,"Age"])>17:
          df2.loc[i,"Occ"]=df2.loc[i,"Occ"]+1

print(df2)

df2.to_excel(r'C:\Users\TheCretin\Desktop\Palm Tree Excel\Occupancy.xlsx', index = False)

for i in range(len(df2)):
    if df2.loc[i,"Occ"]>11:
        df2.loc[i,"Occ"]=11

df2.to_excel(r'C:\Users\TheCretin\Desktop\Palm Tree Excel\OccupancyMax11.xlsx', index = False)