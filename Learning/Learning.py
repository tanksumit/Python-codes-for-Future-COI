import requests
import json
import numpy as np
import pandas as pd
from datetime import datetime, time, timedelta
from time import sleep
import xlwings as xw
import os



headers = {"user-agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Mobile Safari/537.36"}

Excel_File = "Sumit_Auto_Nifty_OI.xlsm"

wb = xw.Book(Excel_File)


OI_Sheet = wb.sheets("OI_Sheet")
OI_CH = wb.sheets("OI_CH")


def get_oi():
      Range = 800
      url = 'https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY'
      data = requests.get(url, headers=headers).json()
      CE=(data['CE'] for data in data['filtered']['data'] if 'CE' in data)
      PE=(data['PE'] for data in data['filtered']['data'] if 'PE' in data)
      CEDF = pd.DataFrame(CE)
      CEDF["Type"]="CE"
      CEDF.loc[CEDF['strikePrice'] > CEDF['underlyingValue'], 'OPTION TYPE'] = 'OTM'
      CEDF.loc[CEDF['strikePrice'] < CEDF['underlyingValue'], 'OPTION TYPE'] = 'ITM'
      CEDF = CEDF.loc[(CEDF['strikePrice'] > CEDF['underlyingValue']-Range) & (CEDF['strikePrice'] < CEDF['underlyingValue']+Range)]
      PEDF = pd.DataFrame(PE)
      PEDF["Type"]="PE"
      PEDF.loc[PEDF['strikePrice'] > PEDF['underlyingValue'], 'OPTION TYPE'] = 'ITM'
      PEDF.loc[PEDF['strikePrice'] < PEDF['underlyingValue'], 'OPTION TYPE'] = 'OTM'
      PEDF = PEDF.loc[(PEDF['strikePrice'] > PEDF['underlyingValue']-Range) & (PEDF['strikePrice'] < PEDF['underlyingValue']+Range)]
      CEPE=pd.concat([CEDF,PEDF])
      df=pd.DataFrame(CEPE).drop(['expiryDate', 'underlying','pchangeinOpenInterest','totalTradedVolume','pChange', 'totalBuyQuantity', 'totalSellQuantity', 'bidQty', 'bidprice', 'askQty',
      'askPrice','lastPrice', 'identifier','change','changeinOpenInterest','impliedVolatility'],axis=1)
      df['openInterest']=df['openInterest'].apply(pd.to_numeric)
      df['openInterest']=75*df['openInterest']
      df.sort_values(['strikePrice','Type'],ascending=False,axis=0,inplace=True)
      df['Time'] = datetime.now().strftime('%H:%M:%S')
      df['Identifier'] = df.apply(lambda x:'%s_%s' % (x['strikePrice'],x['Type']),axis=1)
      
      df['StrikeStr'] = df['strikePrice'].astype(str)
      df['StrikeStr1'] = df['StrikeStr'].str[2:].astype(int)
      
      df.drop(df[df['StrikeStr1'] > 00].index, inplace = True) 
      df = df[['Time','Identifier','strikePrice','Type','OPTION TYPE','underlyingValue','openInterest',]].set_index('Time')
      return df


def all_data():
      while time(9, 00) <= datetime.now().time() <= time(18, 31): 
            alldata=pd.DataFrame()
            sleep_time = 180
            while True:
                  print('Program Starts to check data at',datetime.now().strftime('%H:%M:%S'))
                  print()
                  try:
                        dt=get_oi()
                        df=pd.DataFrame(dt)
                        if len(alldata) == 0: 
                              alldata=alldata.append(dt)
                              print('First time data appended to DataFrame & excel')
                              print()
                              OI_Sheet.range("A1").options(index=True,headers=False).value = alldata
                              print('Please wait for seconds :',sleep_time)
                              print()
                              sleep(sleep_time)
                        elif (alldata['underlyingValue'][0] == dt['underlyingValue'][0]):
                              print('No new data found at',datetime.now().strftime('%H:%M:%S'))
                              print()
                              print('NIFTY LTP :',alldata['underlyingValue'][0],dt['underlyingValue'][0])
                              print()
                              print('Please wait for seconds :',sleep_time)
                              sleep(sleep_time)
                              
                        else:
                              alldata=alldata.append(dt)
                              print('New data appened to DataFrame & excel',datetime.now().strftime('%H:%M:%S'))
                              OI_Sheet.range("A1").options(index=True,headers=False).value = alldata
                              OICR=df[['openInterest','strikePrice','Type']][-32:]
                              OIPR=df[['openInterest','strikePrice','Type']][-65:-32]
                              OI_CH.range("A1").options(index=True,headers=False).value = OICR
                              OI_CH.range("E1").options(index=False,headers=False).value = OIPR

                              
                              print()
                              print('Please wait for seconds :',sleep_time)
                              sleep(sleep_time)
                        
                                      
                  
                  except Exception as error:
                        print('error {0}'.format(error))
                        print()
                        sleep(sleep_time)
                        continue
                        print()
      


all_data()