# -*- coding: utf-8 -*-
"""
Created on Mon Nov 23 14:27:46 2020

@author: HP
"""

import numpy as np
import datetime
import sys

import pandas as pd
from pathlib import Path
script_dir = Path(__file__).resolve().parent
project_dir = script_dir.parent
sys.path.insert(0, str(project_dir))

from common.db import get_engine
from common.config import load_ini

# ============================ read configuration parameters from ini file ==================================
Engine = get_engine('CROSS_HOUR')

config = load_ini('Config_MUSIC_GD.ini')
url_music = config['CONNECT_qhztj']['url']
user_music = config['CONNECT_qhztj']['user']
pwd_music = config['CONNECT_qhztj']['password']
interface_music = config['INTERFACE']['STAT_SURF_AWST_PRE_HOR_XTRM']
columns_music = config['COLUMNS_IN']['HYDRO_PRE_HOR']
columns_out = config['COLUMNS_OUT']['HYDRO_PRE_HOR']
# ============================ read config parameters from ini file ==================================

# thresholds = thresholds.split(',')
columns_out = columns_out.split(',')

Output_TB = 'awst_cli_mul_hor'

# ======================= 设定读取MDOS2数据的起止时间 =============================
time_end = datetime.datetime.now()
time_end = datetime.datetime(2019,1,1,0) #北京时
time_stt = time_end +datetime.timedelta(hours=-5)
time_stt = datetime.datetime(2019,1,1,0) #北京时

iyear_end = time_end.year; imonth_end = time_end.month; iday_end = time_end.day; ihour_end = time_end.hour
time_end = datetime.datetime(iyear_end,imonth_end,iday_end,ihour_end) + datetime.timedelta(hours=-8) #国际时
strtime_end = time_end.strftime('%Y-%m-%d %H:%M:%S')

iyear_stt = time_stt.year; imonth_stt = time_stt.month; iday_stt = time_stt.day; ihour_stt = time_stt.hour
time_stt = datetime.datetime(iyear_stt,imonth_stt,iday_stt,ihour_stt) + datetime.timedelta(hours=-8) #国际时
strtime_stt = time_stt.strftime('%Y-%m-%d %H:%M:%S')

time_range = pd.date_range(strtime_end,strtime_stt,freq='-1h')
# ======================= 设定读取MDOS2数据的起止时间 =============================


for time in time_range:
    strtime = time.strftime('%Y%m%d%H%M%S')  
    
    print('正在读取music接口：水文站小时数据(国际时)：' + strtime)
   
    # reading the accumulated rainfall data of Guangdong province from music
    baseUrl = url_music
    baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
    baseUrl += '&interfaceId=' + interface_music
    baseUrl += '&cols=' + columns_music
    baseUrl += '&ymdhms='+ strtime
    baseUrl += '&dataFormat=html'
    baseUrl += '&v02301=a'
    

    # reading the accumulated rainfall data of Guangdong province from music
    tmp = pd.read_html(baseUrl,encoding='utf-8')
    data_tmp = tmp[0]
    data_tmp.fillna('999999',inplace=True)
    data_a = data_tmp.iloc[1:]
    col_names = data_tmp.iloc[0]
    col_names = map(str.lower, col_names)
    data_a.columns = col_names

    
    strtime = (time + datetime.timedelta(hours=8)).strftime('%Y-%m-%d %H:%M:%S')
    stryear = (time + datetime.timedelta(hours=8)).strftime('%Y')
    

    if data_a.empty == True:
        print('该时次（北京时）水文站无数据：' + strtime)
        continue

    data_a.columns = columns_out
    data_a = data_a.copy()
    data_a.replace('999998','999999',inplace=True)
    data_a.replace('9999','999999',inplace=True)
    
    data_a['ddatetime'] = time + datetime.timedelta(hours=8)
    for column in data_a.columns[2:]:

        if column == 'r':
            data_a[column] = data_a[column].astype(float)
            data_a.loc[data_a[column] == 999999,column] = 0  
            
    data_a['d_iymdhm'] = datetime.datetime.now()
    data_a.drop_duplicates(subset=['stacode'],keep='first',inplace=True)
    
    sql = 'delete from ' + Output_TB + '_' + stryear
    sql += ' where DDATETIME = to_date(\'' + strtime + '\',\'yyyy-mm-dd hh24:mi:ss\')'
    sql += ' and length(stacode)=8'
    Engine.execute(sql)
    
    print('正在写入数据库：水文站小时降水数据（北京时）：' + strtime)
    # write the DataFrame data into Oracle database
    data_a.to_sql(Output_TB + '_' + stryear,Engine,index=False,if_exists='append',chunksize=1000)#,dtype=dtypedict
    
print('完成水文站小时降水数据更新')