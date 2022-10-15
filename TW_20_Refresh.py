# -*- coding: utf-8 -*-
"""
Created on Tue Sep 20 10:56:31 2022

@author: user
"""

import pandas as pd
import requests
from bs4 import BeautifulSoup
import datetime
import time as time
import random

headers = {
    'Accept': 'text/html,application/xhtml+xml,application/xml;q=0.9,image/webp,image/apng,*/*;q=0.8,application/signed-exchange;v=b3;q=0.9',
    'Accept-Encoding': 'gzip, deflate, br',
    'Accept-Language': 'zh-TW,zh;q=0.9,en;q=0.8,en-GB;q=0.7,en-US;q=0.6',
    'Cache-Control': 'max-age=0',
    'Connection': 'keep-alive',
    'Cookie': '_ga=GA1.3.1014788092.1639653982; _gid=GA1.3.53874647.1663498028; JSESSIONID=B8EBDCB41BFE2A7142BA081B897F5791',
    'Host': 'www.twse.com.tw',
    'sec-ch-ua': '"Microsoft Edge";v="105", " Not;A Brand";v="99", "Chromium";v="105"',
    'sec-ch-ua-mobile': '?0',
    'sec-ch-ua-platform': "Windows",
    'Sec-Fetch-Dest': 'document',
    'Sec-Fetch-Mode': 'navigate',
    'Sec-Fetch-Site': 'none',
    'Sec-Fetch-User': '?1',
    'Upgrade-Insecure-Requests': '1',
    'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/105.0.0.0 Safari/537.36 Edg/105.0.1343.42'
}

#88年至今加權指數資料
TW_20 = pd.read_excel("D:/user/研究所/python/TW_88_111.xlsx")

#今天日期
today = datetime.date.today()
str_today = today.strftime('%Y-%m-%d')
Now_Year = str_today[0:4]
Now_Month = str_today[5:7]

#最近一筆資料之年月
Lastest_day =TW_20.iloc[-1].tolist()
lastest_year = int(Lastest_day[0][0:3]) + 1911 
#民國轉西元
lastest_month = int(Lastest_day[0][4:6])

month = ['01','02','03','04','05','06','07','08','09','10','11','12']
 
URL_list = []
if int(datetime.date.today().year) != lastest_year:
    gap = int(datetime.date.today().year) - lastest_year  #ex:2022 到 2024 ,gap = 2
    for i in range(lastest_year,lastest_year + 1 + gap):  # range(2022,2025) <- 2022到2024
        for j in month:
            URL = 'https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=html&date='+str(i)+j+'01'
            URL_list.append(URL)
else:
    for i in range(lastest_year,lastest_year + 1):  # range(2022,2023) <- 2022年
        for j in month:
            URL = 'https://www.twse.com.tw/indicesReport/MI_5MINS_HIST?response=html&date='+str(i)+j+'01'
            URL_list.append(URL)

all_list = []
for url in URL_list:
    resp = requests.get(url,headers=headers)
    time.sleep(random.uniform(1.1,5.5))
    soup = BeautifulSoup(resp.text, 'html5lib')
    try:
        prices = soup.find('tbody').find_all('tr')
    except Exception as e:
        print('Exception: {}'.format(e))
    for price in prices:
        Daily_price = [s for s in price.stripped_strings]
        Daily_price_drop = [i.replace(',','') for i in Daily_price] #把,消除
        Daily_price_real = [Daily_price_drop[0]] + [float(Daily_price_drop[j]) for j in range(1,5)] #將數字字串轉為float
        all_list.append(Daily_price_real)

update_df = pd.DataFrame(all_list, columns = ['Date', 'Open', 'Upper', 'Lower','Close'])
duplicate_df = TW_20.append(update_df,ignore_index=True)
TW_20_update = duplicate_df.drop_duplicates()

TW_20_update.to_excel('TW_88_111.xlsx',index=False)




