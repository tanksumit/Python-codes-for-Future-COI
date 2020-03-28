from bs4 import BeautifulSoup
import requests
import pandas as pd
from datetime import datetime, time
from time import sleep
import xlwings as xw
import os
import numpy as np

headers={'User-Agent':'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.149 Mobile Safari/537.36'}

Symbol=["NIFTY","BANKNIFTY"]
Expiry='30APR2020'
#Create required files

Excel_File = "Pair_Trading.xlsx"

wb = xw.Book(Excel_File)

#Data File 

Data_Sheet = wb.sheets("Data_Sheet")

def get_data():
      NFBNF={}
      for i in Symbol:
            url='https://www1.nseindia.com/live_market/dynaContent/live_watch/get_quote/GetQuoteFO.jsp?underlying='+i+'&instrument=FUTIDX&expiry='+Expiry+'&type=-&strike=-'
            data = requests.get(url,headers=headers)
            soup = BeautifulSoup(data.text,"html.parser")
            data_array=soup.find(id="responseDiv").getText().strip().split(":")
            with open ("data.text", "w") as file:
                  file.write(str(data_array))
            for item in data_array:
                  if 'pChange' in item:
                        index = data_array.index(item)+1
                        pChange = data_array[index].split('"')[1]
                  if 'lastPrice' in item:
                        index = data_array.index(item)+1
                        lastPrice = data_array[index].split('"')[1]
            NFBNF['Time']=datetime.now().strftime('%d-%m-%Y %H:%M:%S')
            NFBNF[i]={'% Change':pChange,'LTP':lastPrice}
            NFBNF[i]={'% Change':pChange,'LTP':lastPrice}
      df=pd.DataFrame(NFBNF)
      return df



def Trading_opporunity():
      while time(9, 00) <= datetime.now().time() <= time(19, 31): 
            alldata=[]
            sleep_time = 10
            while True:
                  print('Program Starts to check Trade at',datetime.now().strftime('%H:%M:%S'))
                  print()
                  try:
                        dt=get_data()
                        df=pd.DataFrame(dt)
                        if len(alldata) == 0: 
                              alldata.append(dt)
                        if (alldata[0]['NIFTY'][1] == dt['NIFTY'][1]) and (alldata[0]['BANKNIFTY'][1] == dt['BANKNIFTY'][1]):
                              print('No new data found at',datetime.now().strftime('%H:%M:%S'))
                              print()
                              print('NIFTY LTP :',alldata[0]['NIFTY'][1],dt['NIFTY'][1],'BANKNIFTY LTP :',alldata[0]['BANKNIFTY'][1],dt['BANKNIFTY'][1])
                              print()
                              sleep(sleep_time)
                        elif (float(df['NIFTY'][0])<=-1) and (float(df['BANKNIFTY'][0])>=1):
                              alldata.append(dt)
                              print('Buy NIFTY and Sell BankNIFTY at ',datetime.now().strftime('%H:%M:%S'))
                              print()
                              sleep(sleep_time)
                        elif (float(df['NIFTY'][0])>= 1) and (float(df['BANKNIFTY'][0])<=-1):
                              alldata.append(dt)
                              print('Buy BankNIFTY and Sell NIFTY at ',datetime.now().strftime('%H:%M:%S'))
                              print()
                              sleep(sleep_time)
                        else:
                              print('No Trading opportunity at ',datetime.now().strftime('%H:%M:%S'))
                              print()
                              print(alldata)
                              print()
                              dt=get_data()
                        for i in alldata:
                              fd=pd.DataFrame(i)
                              Data_Sheet.range("A1").options(index=False,headers=True).value = fd
                  except Exception as error:
                        print('error {0}'.format(error))
                        print()
                        sleep(sleep_time)
                        continue
                  finally:
                        print('Program Ends No Trade Opporunity Found till',datetime.now().strftime('%H:%M:%S'))
                        print()      
Trading_opporunity()