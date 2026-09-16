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
interface_music = config['INTERFACE']['AWST_CLI_MUL_HOR_UPDATE']
columns_music = config['COLUMNS_IN']['AWST_CLI_MUL_HOR']
columns_out = config['COLUMNS_OUT']['AWST_CLI_MUL_HOR']
thresholds = config['THRESHOLD']['AWST_CLI_MUL_HOR']
# ============================ read config parameters from ini file ==================================

thresholds = thresholds.split(',')
columns_out = columns_out.split(',')

Output_TB = 'awst_cli_mul_hor'
csv_path = script_dir / 'awst_hor_update_list.csv'
pd.DataFrame(columns=['ddatetime', 'stacode']).to_csv(csv_path, index=False)

import urllib.parse
prov = '广东'
prov_code = urllib.parse.quote(prov.encode('gb2312'))

# ======================= 设定读取MDOS2数据的起止时间 =============================
time_end = datetime.datetime.now()
# time_end = datetime.datetime(2026,1,1,0) #北京时
time_stt = time_end +datetime.timedelta(hours=-5)
# time_stt = datetime.datetime(2026,1,1,0) #北京时

iyear_end = time_end.year; imonth_end = time_end.month; iday_end = time_end.day; ihour_end = time_end.hour
time_end = datetime.datetime(iyear_end,imonth_end,iday_end,ihour_end) + datetime.timedelta(hours=-8) #国际时
strtime_end = time_end.strftime('%Y-%m-%d %H:%M:%S')

iyear_stt = time_stt.year; imonth_stt = time_stt.month; iday_stt = time_stt.day; ihour_stt = time_stt.hour
time_stt = datetime.datetime(iyear_stt,imonth_stt,iday_stt,ihour_stt) + datetime.timedelta(hours=-8) #国际时
strtime_stt = time_stt.strftime('%Y-%m-%d %H:%M:%S')

time_range = pd.date_range(strtime_end,strtime_stt,freq='-1h')
# ======================= 设定读取MDOS2数据的起止时间 =============================


for time in time_range:
    strtime = (time + datetime.timedelta(hours=8)).strftime('%Y-%m-%d %H')  
    print('正在读取music接口：' + strtime + '时（北京时）更新的区域站小时数据')

    time_stt = datetime.datetime(time.year,time.month,time.day,time.hour,0,0)  # 当前时次的最初时刻
    time_end = datetime.datetime(time.year,time.month,time.day,time.hour,59,59)  # 当前时次的最后时刻
    s_ymdhms = time_stt.strftime('%Y%m%d%H%M%S') 
    e_ymdhms = time_end.strftime('%Y%m%d%H%M%S')

    # reading the hourly data of Guangdong province from music
    baseUrl = url_music
    baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
    baseUrl += '&interfaceId=' + interface_music
    baseUrl += '&s_ymdhms=' + s_ymdhms
    baseUrl += '&e_ymdhms=' + e_ymdhms
    baseUrl += '&cols=' + columns_music
    baseUrl += '&prov=' + prov_code
    baseUrl += '&dataFormat=html'    

    tmp = pd.read_html(baseUrl,encoding='utf-8')
    data_tmp = tmp[0]
    data_tmp.fillna('999999',inplace=True)
    df_hour = data_tmp.iloc[1:]
    if df_hour.empty == True:
        print( strtime +'（北京时）无数据更新')
        continue

    df_hour.columns = columns_out

    AWST_CLI_MUL_HOR = df_hour.copy()  
    AWST_CLI_MUL_HOR.replace('999998', '999999',inplace=True)
    AWST_CLI_MUL_HOR.replace('9999', '999999',inplace=True)
    AWST_CLI_MUL_HOR['ddatetime'] = AWST_CLI_MUL_HOR['ddatetime'].astype('datetime64') + datetime.timedelta(hours=8)
    
    # This is important: only keep records which ddatetime is earlier than the current time, while running Music_AWST_CLI_MUL_HOR is prerequisite.
    AWST_CLI_MUL_HOR = AWST_CLI_MUL_HOR[AWST_CLI_MUL_HOR['ddatetime'] < time + datetime.timedelta(hours=8)].copy()
    if AWST_CLI_MUL_HOR.empty == True:
        print('未更新 ' + strtime +' 时之前时次的数据')
        continue

    AWST_CLI_MUL_HOR.reset_index(drop=True, inplace=True)
    for column in AWST_CLI_MUL_HOR.columns[2:]:

        if column == 'r':
            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] == 999990,column] = 0  
        if column[0] == 'f' and column[-1] == 'x':
            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] == 999017,column] = 361
        
        if column[-4:] != 'time':

            threshold = thresholds[columns_out.index(column)]
            threshold_min = threshold.split('to')[0]
            threshold_max = threshold.split('to')[1]   

            AWST_CLI_MUL_HOR[column] = AWST_CLI_MUL_HOR[column].astype('float')
            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] == 999000,column] = np.nan
            # ======== 2022-6-27：修改阈值范围 ========
            AWST_CLI_MUL_HOR.loc[(AWST_CLI_MUL_HOR[column] < float(threshold_min)) | (AWST_CLI_MUL_HOR[column] > float(threshold_max)),column] = np.nan
            
        if column[-4:] == 'time':
            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] != '999999',column] = AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] != '999999',column].str.zfill(4)

            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] < '0',column] ='999999'              
            tmp_hh = AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] != '999999',column].str[0:2]
            tmp_hh[tmp_hh > '23'] = '23'  # 原始数据偶有错误，此处质量控制一下
            tmp_mm = AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] != '999999',column].str[-2:]
            tmp_mm[tmp_mm > '59'] = '59'  # 原始数据偶有错误，此处质量控制一下                    
            tmp_time = tmp_hh + ':' + tmp_mm
                    
            tmp_time = pd.to_datetime(tmp_time,format='%H:%M') + datetime.timedelta(hours=8)
            tmp_time = tmp_time.dt.strftime('%H:%M')
            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] != '999999',column] = tmp_time    
            
            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] == '999999',column] =''   

    AWST_CLI_MUL_HOR['d_iymdhm'] = datetime.datetime.now()
    AWST_CLI_MUL_HOR.drop_duplicates(subset=['ddatetime','stacode'],keep='first',inplace=True)

    for idx in AWST_CLI_MUL_HOR.index:
        tmp_df = AWST_CLI_MUL_HOR.loc[[idx]]

        tmp_stacode= tmp_df.loc[idx,'stacode']
        tmp_ddatetime = tmp_df.loc[idx,'ddatetime']
        strtmp_ddatetime = tmp_ddatetime.strftime('%Y-%m-%d %H')

        str_year = tmp_ddatetime.strftime('%Y') 
        print('正在写入数据库：' + strtmp_ddatetime + ', ' + tmp_stacode)
        
        sql = 'delete from ' + Output_TB + '_' +str_year
        sql += ' where DDATETIME = to_date(\'' + strtmp_ddatetime + '\',\'yyyy-mm-dd hh24\')'
        sql += ' and stacode =\'' + tmp_stacode + '\''
        Engine.execute(sql)

        # write the DataFrame data into Oracle database
        tmp_df.to_sql(Output_TB + '_' +str_year, Engine, index=False, if_exists='append', chunksize=100)#,dtype=dtypedict
    
    AWST_CLI_MUL_HOR[['ddatetime','stacode']].to_csv(csv_path, mode='a', header=False, index=False)

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ':完成')