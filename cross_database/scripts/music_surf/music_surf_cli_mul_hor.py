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
interface_music = config['INTERFACE']['SURF_CLI_MUL_HOR']
columns_music = config['COLUMNS_IN']['SURF_CLI_MUL_HOR']
columns_out = config['COLUMNS_OUT']['SURF_CLI_MUL_HOR']
thresholds = config['THRESHOLD']['SURF_CLI_MUL_HOR']
# ============================ read config parameters from ini file ==================================

thresholds = thresholds.split(',')
columns_out = columns_out.split(',')

Output_TB = 'surf_cli_mul_hor'

# ======================= 设定读取MDOS2数据的起止时间 =============================
time_end = datetime.datetime.now()
time_end = datetime.datetime(2026,5,11,16) #北京时
time_stt = time_end +datetime.timedelta(hours=-5)
time_stt = datetime.datetime(2026,5,11,12) #北京时

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
    
    print('正在读取music接口：国家站小时数据(国际时)：' + strtime)
    # reading the accumulated rainfall data of Guangdong province from music
    baseUrl = url_music
    baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
    baseUrl += '&interfaceId=' + interface_music
    baseUrl += '&ymdhms=' + strtime
    # baseUrl += '&e_ymdhms=' + strtime
    baseUrl += '&cols=' + columns_music
    baseUrl += '&prov=%B9%E3%B6%AB'
    baseUrl += '&dataFormat=html'
    
    strtime = (time + datetime.timedelta(hours=8)).strftime('%Y-%m-%d %H:%M:%S')
    stryear = (time + datetime.timedelta(hours=8)).strftime('%Y')    
    
    # reading the aws station info of Guangdong province from music
    tmp = pd.read_html(baseUrl,encoding='utf-8')
    data_tmp = tmp[0]
    data_tmp.fillna('999999',inplace=True)
    SURF_CLI_MUL_HOR = data_tmp.iloc[1:]
    if SURF_CLI_MUL_HOR.empty == True:
        print('该时次（北京时）国家站无数据：' + strtime)
        continue

#    col_names = data_tmp.iloc[0]
#    col_names = map(str.lower, col_names) #[col_name.lower() for col_name in col_names]
    SURF_CLI_MUL_HOR.columns = columns_out

    SURF_CLI_MUL_HOR = SURF_CLI_MUL_HOR.copy()
    SURF_CLI_MUL_HOR.replace('999998','999999',inplace=True)
    SURF_CLI_MUL_HOR.replace('9999', '999999', inplace=True)
    
    SURF_CLI_MUL_HOR['ddatetime'] = time + datetime.timedelta(hours=8)
    for column in SURF_CLI_MUL_HOR.columns[2:]:
        
        if column == 'r':
            SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] == 999990,column] = 0  
        if column[0] == 'f' and column[-1] == 'x':
            SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] == 999017,column] = 361

        if column[-4:] != 'time':

            threshold = thresholds[columns_out.index(column)]
            threshold_min = threshold.split('to')[0]
            threshold_max = threshold.split('to')[1]   

            # SURF_CLI_MUL_HOR.loc[:,column] = SURF_CLI_MUL_HOR.loc[:,column].astype('float')
            SURF_CLI_MUL_HOR[column] = SURF_CLI_MUL_HOR[column].astype('float')
            SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] >= 999000,column] = np.nan
            # ======== 2022-6-27：修改阈值范围 ========
            SURF_CLI_MUL_HOR.loc[(SURF_CLI_MUL_HOR[column] < float(threshold_min)) | (SURF_CLI_MUL_HOR[column] > float(threshold_max)),column] = np.nan            

        if column[-4:] == 'time':
            SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] != '999999',column] = SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] != '999999',column].str.zfill(4)
            
            SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] < '0',column] ='999999'              
            tmp_hh = SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] != '999999',column].str[0:2]
            tmp_hh[tmp_hh > '23'] = '23'  # 原始数据偶有错误，此处质量控制一下
            tmp_mm = SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] != '999999',column].str[-2:]
            tmp_mm[tmp_mm > '59'] = '59'  # 原始数据偶有错误，此处质量控制一下                    
            tmp_time = tmp_hh + ':' + tmp_mm
                    
            tmp_time = pd.to_datetime(tmp_time,format='%H:%M') + datetime.timedelta(hours=8)
            tmp_time = tmp_time.dt.strftime('%H:%M')
            SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] != '999999',column] = tmp_time    
            
            SURF_CLI_MUL_HOR.loc[SURF_CLI_MUL_HOR[column] == '999999',column] =''

    
    SURF_CLI_MUL_HOR['d_iymdhm'] = datetime.datetime.now()
    SURF_CLI_MUL_HOR.drop_duplicates(subset=['stacode'],keep='first',inplace=True)
    
    sql = 'delete from ' + Output_TB + '_' + stryear
    sql += ' where DDATETIME = to_date(\'' + strtime + '\',\'yyyy-mm-dd hh24:mi:ss\')'
    Engine.execute(sql)
    
    print('正在写入数据库：国家站小时数据（北京时）：' + strtime)
    # write the DataFrame data into Oracle database
    SURF_CLI_MUL_HOR.to_sql(Output_TB + '_' + stryear,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict
    
print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ':完成')