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
interface_music = config['INTERFACE']['AWST_CLI_MUL_HOR']
columns_music = config['COLUMNS_IN']['AWST_CLI_MUL_HOR']
columns_out = config['COLUMNS_OUT']['AWST_CLI_MUL_HOR']
thresholds = config['THRESHOLD']['AWST_CLI_MUL_HOR']
# ============================ read config parameters from ini file ==================================

thresholds = thresholds.split(',')
columns_out = columns_out.split(',')

stacodes = 'G37' #'G1654'
code_len = len(stacodes)
Output_TB = 'awst_cli_mul_hor'

# ======================= 设定读取MDOS2数据的起止时间 =============================
# time_end = datetime.datetime.now()
time_end = datetime.datetime(2025,8,2,0) #北京时
# time_stt = time_end +datetime.timedelta(hours=-2)
time_stt = datetime.datetime(2025,7,31,20) #北京时

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
    
    print('正在读取music接口：区域站小时数据(国际时)：' + strtime)
    # reading the accumulated rainfall data of Guangdong province from music
    baseUrl = url_music
    baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
    baseUrl += '&interfaceId=' + interface_music
    baseUrl += '&ymdhms=' + strtime
    baseUrl += '&prov=%B9%E3%B6%AB'
    # baseUrl += '&iiiii=' + stacode
    baseUrl += '&dataFormat=html'
    baseUrl += '&cols=' + columns_music
    
    strtime = (time + datetime.timedelta(hours=8)).strftime('%Y-%m-%d %H:%M:%S')
    stryear = (time + datetime.timedelta(hours=8)).strftime('%Y')
    
    # reading the aws station info of Guangdong province from music
    tmp = pd.read_html(baseUrl,encoding='utf-8')
    data_tmp = tmp[0]
    data_tmp.fillna('999999',inplace=True)
    AWST_CLI_MUL_HOR_PROV = data_tmp.iloc[1:]
    if AWST_CLI_MUL_HOR_PROV.empty == True:
        print('该时次（北京时）区域站无数据：' + strtime)
        continue

#    col_names = data_tmp.iloc[0]
#    col_names = map(str.lower, col_names) #[col_name.lower() for col_name in col_names]
    AWST_CLI_MUL_HOR_PROV.columns = columns_out
    AWST_CLI_MUL_HOR = AWST_CLI_MUL_HOR_PROV[AWST_CLI_MUL_HOR_PROV['stacode'].str[0:code_len]==stacodes]

    AWST_CLI_MUL_HOR = AWST_CLI_MUL_HOR.copy()
    AWST_CLI_MUL_HOR.replace('9999','999999',inplace=True)
    
    AWST_CLI_MUL_HOR['ddatetime'] = time + datetime.timedelta(hours=8)
    for column in AWST_CLI_MUL_HOR.columns[2:]:

        if column == 'r':
            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] == 999990,column] = 0  
        if column[0] == 'f' and column[-1] == 'x':
            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] == 999017,column] = 361
        
        if column[-4:] != 'time':

            threshold = thresholds[columns_out.index(column)]
            threshold_min = threshold.split('to')[0]
            threshold_max = threshold.split('to')[1]   

            # AWST_CLI_MUL_HOR.loc[:,column] = AWST_CLI_MUL_HOR.loc[:,column].astype('float')
            AWST_CLI_MUL_HOR[column] = AWST_CLI_MUL_HOR[column].astype('float')
            AWST_CLI_MUL_HOR.loc[AWST_CLI_MUL_HOR[column] == 999999,column] = np.nan
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
    AWST_CLI_MUL_HOR.drop_duplicates(subset=['stacode'],keep='first',inplace=True)
    
    sql = 'delete from ' + Output_TB + '_' + stryear
    sql += ' where DDATETIME = to_date(\'' + strtime + '\',\'yyyy-mm-dd hh24:mi:ss\')'
    sql += ' and stacode like \'%' + stacodes + '%\''
    Engine.execute(sql)
    
    print('正在写入数据库：区域站小时数据（北京时）：' + strtime)
    # write the DataFrame data into Oracle database
    AWST_CLI_MUL_HOR.to_sql(Output_TB + '_' + stryear,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict
    
print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ':完成')