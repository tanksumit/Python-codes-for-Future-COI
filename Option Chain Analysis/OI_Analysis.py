#import required modules

import requests
import json
import pandas as pd
import xlwings as xw 
from time import sleep
from datetime import datetime, time, timedelta
import os
import numpy as np

#Variables

Nifty_URL = "https://www.nseindia.com/api/option-chain-indices?symbol=NIFTY"

headers = {"user-agent": "Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/80.0.3987.132 Mobile Safari/537.36"}

#Enter expiry date in 30-Apr-2020 formate if want data for specific expiry by default latest expiry

Expiry = ""

#Create required files

Excel_File = "Option_Chain.xlsx"

wb = xw.Book(Excel_File)

#Current_OI_Data (sheet_oi_single) 

Current_OI_Data = wb.sheets("OIData")

Dashboard_Data = wb.sheets("Dashboard")

Max_Pain_Data = wb.sheets("Max Pain")

Sheet_live_Data = wb.sheets("All_Data")

File_path = "D:\\Automate the Borring Stuff\\Option Chain Analysis\\Files"

#OI_Data_File (oi_filename)

OI_Data_File = os.path.join(File_path,"OI_Data_Records_{0}.json".format(datetime.now().strftime('%d%m%y')))

Max_Pain_Data_File = os.path.join(File_path,"Max_Pain_Data_Records_{0}.json".format(datetime.now().strftime('%d%m%y')))

#Empty list to check whether data store are updated and not duplicate data

#Data (df_list)
Data = []

#Max_Pain (mp_list)
Max_Pain = []

#Prog start from here

def Fetch_OI_Data(datas,Max_Pain_DF):
    Tries = 1
    Max_Retries = 3
    while Tries <= Max_Retries:
        try: 
            # webdata (r)
            Webdata = requests.get(Nifty_URL,headers=headers,timeout=100).json()
            with open ("OIData.json", "w") as file:
                file.write (json.dumps (Webdata, indent = 4, sort_keys = True))
            if Expiry:
                CE_Values = [data['CE']for data in Webdata ['records'] ['data'] if "CE" in data and str(data["expiryDate"].lower()) == str(Expiry.lower())]
                PE_Values = [data['PE']for data in Webdata ['records'] ['data'] if "PE" in data and str(data["expiryDate"].lower()) == str(Expiry.lower())]
                #CE_DataFrame (ce_data) PE_DataFrame (pe_data)  
                CE_DataFrame = pd.DataFrame(CE_Values)
                PE_DataFrame = pd.DataFrame(PE_Values)
                CE_DataFrame = CE_DataFrame.sort_values(['strikePrice'])
                PE_DataFrame = PE_DataFrame.sort_values(['strikePrice'])
                Current_OI_Data.range("A1").options(index=False,headers=True).value = CE_DataFrame.drop(['bidQty',
                    'bidprice','askQty','askPrice','totalBuyQuantity','totalSellQuantity','expiryDate','underlying',
                    'identifier','totalTradedVolume','underlyingValue','pChange','pchangeinOpenInterest'], 
                    axis = 1)[['openInterest','changeinOpenInterest','impliedVolatility','change','lastPrice','strikePrice']]
                Current_OI_Data.range("G1").options(index=False,headers=True).value = PE_DataFrame.drop(['bidQty','bidprice','askQty','askPrice',
                    'totalBuyQuantity','totalSellQuantity','expiryDate','underlying','identifier','totalTradedVolume',
                    'underlyingValue','pChange','pchangeinOpenInterest','strikePrice'], axis = 1)[['lastPrice','change',
                    'impliedVolatility','changeinOpenInterest','openInterest']]
                CE_DataFrame['Type'] = "CE"
                PE_DataFrame['Type'] = "PE"
                #CE_PEData (df1)
                CE_PEData = pd.concat([CE_DataFrame,PE_DataFrame],sort = True)    
                if len(Data) > 0:
                    CE_PEData['Time'] = Data[-1][0]['Time']
                if len(Data) > 0 and CE_PEData.to_dict('records') == Data[-1]: #Checking the data recieved from NSE with Data store in Data list
                    print("Duplicate Data not recording")
                    sleep(5)
                    Tries += 1
                    continue 
                CE_PEData['Time'] = datetime.now().strftime('%H:%M')
                # if not datas.empty:
                #     datas = [['strikePrice','expiryDate','underlying','identifier','openInterest','changeinOpenInterest',
                #     'pchangeinOpenInterest','totalTradedVolume','impliedVolatility','lastPrice','change','pChange','totalBuyQuantity',
                #     'totalSellQuantity','bidQty','bidprice','askQty','askPrice','underlyingValue','Type','Time']]
                # CE_PEData = CE_PEData [['strikePrice','expiryDate','underlying','identifier','openInterest','changeinOpenInterest',
                #     'pchangeinOpenInterest','totalTradedVolume','impliedVolatility','lastPrice','change','pChange','totalBuyQuantity',
                #     'totalSellQuantity','bidQty','bidprice','askQty','askPrice','underlyingValue','Type','Time']]
                PCR = PE_DataFrame['openInterest'].sum()/CE_DataFrame['openInterest'].sum()
                Max_Pain = Dashboard_Data.range("B1").value
                Max_Pain_Dict = {datetime.now().strftime('%H:%M'):{'ATime':datetime.now().strftime('%H:%M'),
                'Underlying':CE_PEData['underlyingValue'].iloc[-1],
                'Max Pain':Max_Pain,
                'PCR':PCR,
                'Call_Decay':CE_DataFrame.nlargest(15,'openInterest',keep = 'last')['change'].mean(),
                'Put_Decay':PE_DataFrame.nlargest(15,'openInterest',keep = 'last')['change'].mean()}}
                #Max_Pain_DataFrame (df3)
                Max_Pain_DataFrame = pd.DataFrame(Max_Pain_Dict).transpose()
                #Max_Pain_DF (mp_df)
                Max_Pain_DF = pd.concat([Max_Pain_DF,Max_Pain_DataFrame],sort = True)
                Max_Pain_Data.range('A2').options(header = False, index = False).value = Max_Pain_DF            
                with open (Max_Pain_Data_File, 'w') as file:
                    file.write(json.dumps(Max_Pain_DF.to_dict(), indent = 4, sort_keys = True))
                datas = pd.concat([datas, CE_PEData],sort = True)
                Data.append(CE_PEData.to_dict('records'))                        
                with open (OI_Data_File, 'w') as file:
                    file.write(json.dumps(Data, indent = 4, sort_keys = True))
                return datas, Max_Pain_DF
            else:
                CE_Values = [data['CE']for data in Webdata ['filtered'] ['data'] if "CE" in data]
                PE_Values = [data['PE']for data in Webdata ['filtered'] ['data'] if "PE" in data]
                #CE_DataFrame (ce_data) PE_DataFrame (pe_data)  
                CE_DataFrame = pd.DataFrame(CE_Values)
                PE_DataFrame = pd.DataFrame(PE_Values)
                CE_DataFrame = CE_DataFrame.sort_values(['strikePrice'])
                PE_DataFrame = PE_DataFrame.sort_values(['strikePrice'])
                Current_OI_Data.range("A1").options(index=False,headers=True).value = CE_DataFrame.drop(['bidQty',
                    'bidprice','askQty','askPrice','totalBuyQuantity','totalSellQuantity','expiryDate','underlying',
                    'identifier','totalTradedVolume','underlyingValue','pChange','pchangeinOpenInterest'], 
                    axis = 1)[['openInterest','changeinOpenInterest','impliedVolatility','change','lastPrice','strikePrice']]
                Current_OI_Data.range("G1").options(index=False,headers=True).value = PE_DataFrame.drop(['bidQty','bidprice','askQty','askPrice',
                    'totalBuyQuantity','totalSellQuantity','expiryDate','underlying','identifier','totalTradedVolume',
                    'underlyingValue','pChange','pchangeinOpenInterest','strikePrice'], axis = 1)[['lastPrice','change',
                    'impliedVolatility','changeinOpenInterest','openInterest']]
                CE_DataFrame['Type'] = "CE"
                PE_DataFrame['Type'] = "PE"
                #CE_PEData (df1)
                CE_PEData = pd.concat([CE_DataFrame,PE_DataFrame],sort = True)    
                if len(Data) > 0:
                    CE_PEData['Time'] = Data[-1][0]['Time']
                if len(Data) > 0 and CE_PEData.to_dict('records') == Data[-1]: #Checking the data recieved from NSE with Data store in Data list
                    print("Duplicate Data not recording")
                    sleep(5)
                    Tries += 1
                    continue 
                CE_PEData['Time'] = datetime.now().strftime('%H:%M')
                # if not datas.empty:
                #     datas = [['strikePrice','expiryDate','underlying','identifier','openInterest','changeinOpenInterest',
                #     'pchangeinOpenInterest','totalTradedVolume','impliedVolatility','lastPrice','change','pChange','totalBuyQuantity',
                #     'totalSellQuantity','bidQty','bidprice','askQty','askPrice','underlyingValue','Type','Time']]
                # CE_PEData = CE_PEData [['strikePrice','expiryDate','underlying','identifier','openInterest','changeinOpenInterest',
                #     'pchangeinOpenInterest','totalTradedVolume','impliedVolatility','lastPrice','change','pChange','totalBuyQuantity',
                #     'totalSellQuantity','bidQty','bidprice','askQty','askPrice','underlyingValue','Type','Time']]
                PCR = PE_DataFrame['openInterest'].sum()/CE_DataFrame['openInterest'].sum()
                Max_Pain = Dashboard_Data.range("B1").value
                Max_Pain_Dict = {datetime.now().strftime('%H:%M'):{'ATime':datetime.now().strftime('%H:%M'),
                'Underlying':CE_PEData['underlyingValue'].iloc[-1],
                'Max Pain':Max_Pain,
                'PCR':PCR,
                'Call_Decay':CE_DataFrame.nlargest(15,'openInterest',keep = 'last')['change'].mean(),
                'Put_Decay':PE_DataFrame.nlargest(15,'openInterest',keep = 'last')['change'].mean()}}
                #Max_Pain_DataFrame (df3)
                Max_Pain_DataFrame = pd.DataFrame(Max_Pain_Dict).transpose()
                #Max_Pain_DF (mp_df)
                Max_Pain_DF = pd.concat([Max_Pain_DF,Max_Pain_DataFrame],sort = True)
                Max_Pain_Data.range('A2').options(header = False, index = False).value = Max_Pain_DF            
                with open (Max_Pain_Data_File, 'w') as file:
                    file.write(json.dumps(Max_Pain_DF.to_dict(), indent = 4, sort_keys = True))
                datas = pd.concat([datas, CE_PEData],sort = True)
                Data.append(CE_PEData.to_dict('records'))                        
                with open (OI_Data_File, 'w') as file:
                    file.write(json.dumps(Data, indent = 4, sort_keys = True))
                return datas, Max_Pain_DF
        except Exception as error:
            print('error {0}'.format(error))
            Tries +=1
            sleep(10)
            continue   
    if Tries >= Max_Retries:
        print('Max retries exceeded. No new Data at time {0}'.format(datetime.now().strftime('%d-%m-%y-%H:%M')))
        return datas, Max_Pain_DF

def Main():
    global Data
    try:
        Data = json.loads (open(OI_Data_File).read())
    except Exception as error:
        print('Error in reading Data: {0}'.format(error))
        Data = []
    

    if Data:
        datas = pd.DataFrame()
        for item in Data: 
            datas = pd.concat([datas, pd.DataFrame(item)],sort = True)
    else:
        datas = pd.DataFrame()

    try:
        Max_Pain = json.loads (open(Max_Pain_Data_File).read())
        Max_Pain_DF = pd.DataFrame().from_dict(Max_Pain)
    except Exception as error:
        print('Error in reading Data: {0}'.format(error))
        Max_Pain = []
        Max_Pain_DF = pd.DataFrame() 

    timeframe = 3
    while time(9, 15) <= datetime.now().time() <= time(20, 31): 
        timenow = datetime.now()
        check = True if timenow.minute/timeframe in list (np.arange(0.0, 20.0)) else False #enter time interval of our choice under timeframe variable. 
        #based on that decide  arnage (20 is decided based on how much time interval you want. if interval of 5 the it will be 12)
        if check:
            nextscan = timenow + timedelta (minutes = timeframe)
            datas, Max_Pain_DF = Fetch_OI_Data(datas,Max_Pain_DF)
            if not datas.empty:
                datas['impliedVolatility'] = datas['impliedVolatility'].replace(to_replace = 0, method = 'bfill').values
                datas['identifier'] = datas['strikePrice'].astype(str) + datas['Type']
                Sheet_live_Data.range('A1').options(header = True, index = False).value = datas.drop(['underlying'],axis = 1) [['Time',
                'Type','strikePrice','identifier','underlyingValue','expiryDate','openInterest','changeinOpenInterest','pchangeinOpenInterest',
                'lastPrice','change','pChange','impliedVolatility','askPrice','askQty','bidQty','bidprice','totalBuyQuantity','totalSellQuantity','totalTradedVolume']]
                wb.api.RefreshAll()
                waitsecs = int((nextscan - datetime.now()).seconds) 
                print('wait for {0} seconds',format(waitsecs))
                sleep(waitsecs) if waitsecs > 0 else sleep(0)
            else:
                print('No data received')
                sleep(15)

if __name__ == '__main__':
    Main()


