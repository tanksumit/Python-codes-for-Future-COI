import requests
import json
import numpy as np
import pandas as pd
from datetime import datetime, time, timedelta
from time import sleep
import xlwings as xw
import os



headers = {"user-agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Mobile Safari/537.36"}

Excel_File = "Sumit_Auto_Nifty_OI.xlsx"

wb = xw.Book(Excel_File)


OI_Sheet = wb.sheets("OI_Sheet")
OI_CH = wb.sheets("OI_CH")
CEPE_ALL = wb.sheets("CEPE_All")

#Folder path where you want to store EOD files
File_path = "D:\\Automate the Borring Stuff\\Learning\\Files"

#OI_Data_File (oi_filename)

Data_File = os.path.join(File_path,"OI_Data_Records_{0}.csv".format(datetime.now().strftime('%d-%m-%y-%H-%M')))


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
      
      

      CEDF=pd.DataFrame(CEDF).drop(['expiryDate','pchangeinOpenInterest', 'totalBuyQuantity', 'totalSellQuantity', 'bidQty', 'bidprice', 'askQty', 'askPrice','identifier','pChange'],axis=1)
      PEDF=pd.DataFrame(PEDF).drop(['expiryDate','pchangeinOpenInterest', 'totalBuyQuantity', 'totalSellQuantity', 'bidQty', 'bidprice', 'askQty','askPrice','identifier'],axis=1)
      
      CEDF[['openInterest','changeinOpenInterest']]=CEDF[['openInterest','changeinOpenInterest']].apply(pd.to_numeric)
      CEDF[['openInterest','changeinOpenInterest']]=75*CEDF[['openInterest','changeinOpenInterest']]
      
      PEDF[['openInterest','changeinOpenInterest']]=PEDF[['openInterest','changeinOpenInterest']].apply(pd.to_numeric)
      PEDF[['openInterest','changeinOpenInterest']]=75*PEDF[['openInterest','changeinOpenInterest']]

      CEDF.sort_values(['strikePrice','Type'],ascending=True,axis=0,inplace=True)
      PEDF.sort_values(['strikePrice','Type'],ascending=True,axis=0,inplace=True)
      
      CEDF['Identifier'] = CEDF.apply(lambda x:'%s_%s' % (x['strikePrice'],x['Type']),axis=1)
      PEDF['Identifier'] = PEDF.apply(lambda x:'%s_%s' % (x['strikePrice'],x['Type']),axis=1)
      
      CEDF['StrikeStr'] = CEDF['strikePrice'].astype(str)
      CEDF['StrikeStr1'] = CEDF['StrikeStr'].str[2:].astype(int)
      CEDF.drop(CEDF[CEDF['StrikeStr1'] > 00].index, inplace = True)

      PEDF['StrikeStr'] = PEDF['strikePrice'].astype(str)
      PEDF['StrikeStr1'] = PEDF['StrikeStr'].str[2:].astype(int)
      PEDF.drop(PEDF[PEDF['StrikeStr1'] > 00].index, inplace = True) 
      
      CEDF['Time'] = datetime.now().strftime('%H:%M:%S')
      CEDF = CEDF[['Time','Identifier','Type','OPTION TYPE','openInterest','changeinOpenInterest','totalTradedVolume','lastPrice','change','impliedVolatility','strikePrice','underlyingValue']].set_index('Time')
      PEDF['Time'] = datetime.now().strftime('%H:%M:%S')
      PEDF = PEDF[['Time','impliedVolatility','change','lastPrice','totalTradedVolume','changeinOpenInterest','openInterest','OPTION TYPE','Type','Identifier','strikePrice','underlyingValue']].set_index('Time')

      
      df = pd.concat([CEDF,PEDF])

      
      
      
      return df, CEDF,PEDF




def all_data():
      Start_Time = time(9, 00) 
      End_Time = time(15, 31)
      Time_now = datetime.now().time() 
      while  Start_Time <= datetime.now().time() <= End_Time: 
            alldata=pd.DataFrame()
            sleep_time = 170
            while True:
                  print('Program Starts to check data at',datetime.now().strftime('%H:%M:%S'))
                  print()
                  try:
                        CEPE, CEDF, PEDF=get_oi()
                        if len(alldata) == 0: 
                              alldata=alldata.append(CEPE)
                              print('First time data appended to DataFrame & excel')
                              print()
                              OI_Sheet.range("A1").options(index=True,headers=False).value = CEDF
                              OI_Sheet.range("L1").options(index=False,headers=False).value = PEDF
                              print('Please wait for seconds :',sleep_time)
                              print()
                              sleep(sleep_time)
                              
                        elif (alldata['underlyingValue'][0] == CEPE['underlyingValue'][0]):
                              print('No new data found at',datetime.now().strftime('%H:%M:%S'))
                              print()
                              print('NIFTY LTP :',alldata['underlyingValue'][0],CEPE['underlyingValue'][0])
                              print()
                              print('Please wait for seconds :',sleep_time)
                              sleep(sleep_time)
                              print()
                              
                        else:
                              alldata=alldata.append(CEPE)
                              print('New data appened to DataFrame & excel',datetime.now().strftime('%H:%M:%S'))
                              
                              OI_Sheet.range("A1").options(index=True,headers=False).value = CEDF
                              OI_Sheet.range("L1").options(index=False,headers=False).value = PEDF

                              CEPE_ALL.range("A1").options(index=True,headers=False).value = alldata
                              OICR=alldata[['strikePrice','Type','OPTION TYPE','openInterest']][-32:]
                              OIPR=alldata[['strikePrice','Type','OPTION TYPE','openInterest']][-64:-32]
                              
                              
                              OI_CH.range("A1").options(index=True,headers=False).value = OICR
                              OI_CH.range("G1").options(index=True,headers=False).value = OIPR

                              
                              print()
                              print('Please wait for seconds :',sleep_time)
                              print()
                              sleep(sleep_time)
                        
                                      
                  
                  except Exception as error:
                        print('error {0}'.format(error))
                        print()
                        sleep(sleep_time)
                        continue
                        print()

                  if Time_now > End_Time:
                        newdata = pd.concat([alldata,CEDF, PEDF])
                        newdata.to_csv(Data_File,index = True, header=True)
                        break



all_data()





