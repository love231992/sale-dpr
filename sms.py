import requests
import json

URL = 'https://www.sms4india.com/api/v1/sendCampaign'

# get request
def sendPostRequest(reqUrl, apiKey, secretKey, useType, phoneNo, senderId, textMessage):
  req_params = {
  'apikey':apiKey,
  'secret':secretKey,
  'usetype':useType,
  'phone': phoneNo,
  'message':textMessage,
  'senderid':senderId
  }
  return requests.post(reqUrl, req_params)

import pandas as pd
import datetime
from datetime import date
import time
today = date.today()
d1 = today.strftime("%d-%m")
def sms_total(cust_code):
    x = datetime.datetime.now()
    month = x.strftime("%B")
    path = r'\\10.9.32.2\adm\Ash\FY 2019-20\Sale detail sheet\SALE DETAIL SHEET JANUARY 2020.xlsx'
    today = date.today()
    d1 = today.strftime("%d-%m-%Y")
    tarik = d1 
    df = pd.read_excel(open(path,"rb"),sheet_name= month.upper(), index_col=None, header=  None)
    acc = df[df[2] == cust_code]
    a = acc[acc[6]== tarik][8].sum()
    return round(a,2)


def sum_all(li):
    all1 = []
    for i in li:
        x = sms_total(i)
        all1.append(x)
    return sum(all1)

li = [[20,31],[100125],[13,14,15,16,37],[17,18,27,28]]
cust = []
for i in li:
    cust.append(sum_all(i)) 
path = r'\\10.9.32.2\adm\Ash\FY 2019-20\Sale detail sheet\SALE DETAIL SHEET JANUARY 2020.xlsx'  
x = datetime.datetime.now()
month = x.strftime("%B")
df = pd.read_excel(open(path,"rb"),sheet_name= month.upper(), index_col=None, header=  None)  
today = date.today()
d1 = today.strftime("%d-%m-%Y")
tarik = d1
total= df[df[6]== tarik][8].sum()  
total_count= df[df[6]== tarik][8].count()  
msg1 = f'\"Dear Sir,\nTotal fly ash sold for the date of {d1}, quantity {total} MT with filled {total_count} nos. of bulker.\nAsian cement-{cust[0]}\nUltratech cement-{cust[3]}\nAmbuja - {cust[2]}\nACC-{cust[1]}\nSilo level is A,B,C-1,1,1 Mtr and empty bulkers inside the plant is 20 Nos.\nTotal pond ash trips for the date of 10th Jan is 00 Nos. with approx 00 MT.The total ash utilization is 97% till yesterday for the FY 19-20\nRegards,\nAsh Management.\"'
# msg ="hey hwllo" 
msg = msg1
print(msg)

# get response
# response = sendPostRequest(URL, '6SH3R50GCLSSSUWX27G1ZICFCJDIWCYV', 'L61JI3Z57XI7JNLX', 'stage', '8288004541', 'lovnpl1992@gmail.com', msg )
"""
  Note:-
    you must provide apikey, secretkey, usetype, mobile, senderid and message values
    and then requst to api
"""
# print response if you want
# print(response.text)