# from io import BytesIO
# from urllib.request import urlopen
# from zipfile import ZipFile
# import pandas as pd
# import requests
# import urllib.request
# from bs4 import BeautifulSoup
# from datetime import date, timedelta, datetime  
# import xlwings as xw
# import json

# fno = ['RELIANCE','ZEEL','AMARAJABAT','COLPAL','IDEA','ASIANPAINT','NAUKRI','MUTHOOTFIN','JUSTDIAL','NESTLEIND','ACC','TATACONSUM','MFSL','EXIDEIND','SRF','MARUTI','MARICO','ESCORTS','HEROMOTOCO','APOLLOTYRE','SUNTV','HCLTECH','CUMMINSIND','TORNTPHARM','BRITANNIA','M&M','TCS','BAJAJ-AUTO','INFRATEL','BOSCHLTD','EICHERMOT','NIITTECH','PIDILITIND','PETRONET','ADANIPOWER','KOTAKBANK','HINDUNILVR','TECHM','NCC','INFY','GLENMARK','ASHOKLEY','ITC','HAVELLS','BERGEPAINT','AMBUJACEM','DRREDDY','NTPC','MOTHERSUMI','ADANIPORTS','RAMCOCEM','BEL','TATAPOWER','ADANIENT','AUROPHARMA','BHARTIARTL','TITAN','BIOCON','AXISBANK','GMRINFRA','PEL','SUNPHARMA','MCDOWELL-N','JINDALSTEL','INDUSINDBK','TORNTPOWER','UPL','CONCOR','SRTRANSFIN','GAIL','TATASTEEL','COALINDIA','ULTRACEMCO','BANDHANBNK','JUBLFOOD','DABUR','BATAINDIA','CHOLAFIN','LT','SBIN','BAJAJFINSV','WIPRO','MRF','JSWSTEEL','CIPLA','IDFCFIRSTB','HDFCBANK','TATACHEM','GRASIM','BAJFINANCE','HDFCLIFE','APOLLOHOSP','SHREECEM','IBULHSGFIN','LUPIN','CESC','HDFC','ICICIBANK','L&TFH','VOLTAS','DIVISLAB','PAGEIND','MANAPPURAM','HINDALCO','YESBANK','IOC','FEDERALBNK','TATAMOTORS','BHARATFORG','GODREJCP','POWERGRID','MINDTREE','NMDC','PNB','BPCL','SAIL','NATIONALUM','DLF','UBL','LICHSGFIN','PFC','BANKBARODA','BHEL','EQUITAS','RECLTD','M&MFIN','HINDPETRO','SIEMENS','TVSMOTOR','BALKRISIND','ICICIPRULI','GODREJPROP','MGL','IGL','CANBK','PVR','VEDL','CADILAHC','OIL','INDIGO','UJJIVAN',
# 'RBLBANK','ONGC','CENTURYTEX']

# Excel_File = "Cash_BC.xlsx"

# wb = xw.Book(Excel_File)

# Data = wb.sheets("Data")

# fnoData = wb.sheets("fno")

# cashBC = wb.sheets("cashbc")

# To_Date=date.today().strftime('%d-%m-%Y')
# No_of_Days=30
# FDate=date.today()
# Fm_Date=FDate-timedelta(days=No_of_Days)
# From_Date=Fm_Date.strftime('%d-%m-%Y')
# Days=(FDate.day-Fm_Date.day)
# now=datetime.now()
# dt_string=now.strftime("%d/%m/%Y %H:%M")

# Trading_Days = []
# url="https://www1.nseindia.com/products/dynaContent/equities/indices/historicalindices.jsp?indexType=NIFTY%2050&fromDate="+str(From_Date)+"&toDate="+str(To_Date)
# headers = {'User-Agent':'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Mobile Safari/537.36','Referer':'https://www1.nseindia.com/products/content/equities/indices/historical_index_data.htm'}
# response = requests.get(url,headers=headers)
# soup = BeautifulSoup(response.text,"html.parser")
# s=soup.findAll('nobr')
# for i in s:
#     for m in i:
#         Trading_Days.append(m)
        
# # # https://www1.nseindia.com/content/historical/EQUITIES/2020/APR/cm21APR2020bhav.csv.zip

# links = []
# for i in Trading_Days:
#     url = 'https://www1.nseindia.com/content/historical/EQUITIES/'
#     link=(url+(i[-4:len(i)])+'/'+(i[-8:len(i)-5].upper()+'/'+'cm'+i.upper().replace('-',"")+'bhav.csv.zip'))
#     links.append(link)

    



# url = links

# all = pd.DataFrame()

# count = 0
# while count < len(url):
#     with urlopen(url[count]) as zipresp:
#         with ZipFile(BytesIO(zipresp.read())) as zfile:
#             zfile.extractall('D:\Automate the Borring Stuff\CM_Bhavcopy\Bhavcopy_Files')
#             r=zfile.namelist()
#             for i in r:
#                 data=pd.read_csv('D:\Automate the Borring Stuff\CM_Bhavcopy\Bhavcopy_Files\{}'.format(i))
#                 filter=pd.DataFrame(data).drop(['OPEN','HIGH','LOW','CLOSE','LAST','PREVCLOSE','TOTTRDVAL','TOTALTRADES','TOTALTRADES','ISIN'],axis=1)
#                 EQ=(filter.loc[filter['SERIES']=='EQ'])
#                 i=EQ[EQ['SYMBOL'].isin(fno)].drop(['SERIES','Unnamed: 13'],axis = 1).set_index('SYMBOL')
#                 all=all.append(i)
#                 count = count + 1
#                 print('Data receivied for cash bhav copy {} Trading Day'.format(count))
#                 lastr = len(all['TIMESTAMP'])
                
#                 Data.range('A3').options(header = True, index = True).value = all[lastr-144:lastr]
#                 Data.range('D3').options(header = True, index = True).value = all[lastr-288:lastr-144]
#                 Data.range('G3').options(header = True, index = True).value = all[lastr-432:lastr-288]
#                 Data.range('J3').options(header = True, index = True).value = all[lastr-576:lastr-432]
#                 Data.range('M3').options(header = True, index = True).value = all[lastr-720:lastr-576]
#                 Data.range('P3').options(header = True, index = True).value = all[lastr-864:lastr-720]
#                 Data.range('S3').options(header = True, index = True).value = all[lastr-1008:lastr-864]
#                 Data.range('V3').options(header = True, index = True).value = all[lastr-1152:lastr-1008]
#                 Data.range('Y3').options(header = True, index = True).value = all[lastr-1296:lastr-1152]
#                 Data.range('AB3').options(header = True, index = True).value = all[lastr-1440:lastr-1296]
#                 Data.range('AE3').options(header = True, index = True).value = all[lastr-1584:lastr-1440]
#                 Data.range('AH3').options(header = True, index = True).value = all[lastr-1728:lastr-1584]


# fno_url = "https://www1.nseindia.com/live_market/dynaContent/live_watch/stock_watch/foSecStockWatch.json"
                    
# headers = {'User-Agent':'Mozilla/5.0 (Linux; Android 6.0; Nexus 5 Build/MRA58N) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/79.0.3945.130 Mobile Safari/537.36','Referer':'https://www1.nseindia.com/products/content/equities/indices/historical_index_data.htm'}

# response = requests.get(fno_url,headers=headers).json()
# data = pd.DataFrame(response['data']).drop(['ltP','ptsC','trdVolM','ntP', 'mVal', 'wkhi', 'wklo', 'wkhicm_adj', 'wklocm_adj','xDt', 'cAct', 'yPC', 'mPC'],axis = 1)
# data = data.sort_values('symbol')
# fnoData.range('A3').options(header = True, index = False).value = data





# links = []
# for i in Trading_Days:
#     url = 'https://www1.nseindia.com/content/historical/DERIVATIVES/'
#     link=(url+(i[-4:len(i)])+'/'+(i[-8:len(i)-5].upper()+'/'+'fo'+i.upper().replace('-',"")+'bhav.csv.zip'))
#     links.append(link)


# url = links

# all = pd.DataFrame()

# count = 0
# while count < len(url):
#     with urlopen(url[count]) as zipresp:
#         with ZipFile(BytesIO(zipresp.read())) as zfile:
#             zfile.extractall('D:\Automate the Borring Stuff\CM_Bhavcopy\Bhavcopy_Files')
#             r=zfile.namelist()
#             for i in r:
#                 data=pd.read_csv('D:\Automate the Borring Stuff\CM_Bhavcopy\Bhavcopy_Files\{}'.format(i))
#                 filter=pd.DataFrame(data).drop(['STRIKE_PR','OPTION_TYP','CONTRACTS','VAL_INLAKH'],axis=1)
#                 data=(filter.loc[filter['INSTRUMENT']=='FUTSTK']).set_index('SYMBOL')
                
#                 data['COI_OI']=data['OPEN_INT'].sum()
#                 data['COI_CHOI']=data['CHG_IN_OI'].sum()
                            
                
#                 all=all.append(data).drop(['INSTRUMENT','SETTLE_PR'],axis = 1)[['TIMESTAMP','EXPIRY_DT',
#                 'OPEN','HIGH','LOW','CLOSE','OPEN_INT','CHG_IN_OI','COI_OI','COI_CHOI']]
#                 count = count + 1
#                 print('Data receivied for fno bhav copy {} Trading Day'.format(count))           

# # all = all.sort_values('TIMESTAMP')
# cashBC.range('A3').options(header = True, index = True).value = all



import pandas as pd

data = pd.read_csv('allcombined.csv')
data_1 = data.groupby(['SYMBOL','TIMESTAMP']).sum().groupby('SYMBOL').cumsum()
print(data_1.head())