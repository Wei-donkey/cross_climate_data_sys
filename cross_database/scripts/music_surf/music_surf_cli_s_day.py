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
Engine = get_engine('CROSS_CLIMATE')

config = load_ini('Config_MUSIC_GD.ini')
url_music = config['CONNECT_qhztj']['url']
user_music = config['CONNECT_qhztj']['user']
pwd_music = config['CONNECT_qhztj']['password']
interface_music = config['INTERFACE']['SURF_CLI_S_DAY']
columns_music = config['COLUMNS_IN']['SURF_CLI_S_DAY']
# columns_out = config['COLUMNS_OUT']['SURF_CLI_S_DAY']
# ============================ read config parameters from ini file ==================================


Output_TB = 'surf_cli_mul_day'
StaInfo_TB = 't_othe_station_meta_basic_tab'

# ======================= 设定读取数据的起止时间 =============================
date_end = datetime.datetime.now()
# date_end = datetime.datetime(2023,2,28) #北京时
date_stt = date_end +datetime.timedelta(days=-10)
date_stt = datetime.datetime(2026,8,10) #北京时

strdate_end = date_end.strftime('%Y-%m-%d')
strdate_stt = date_stt.strftime('%Y-%m-%d')
date_range = pd.date_range(strdate_end,strdate_stt,freq='-1d')
# ======================= 设定读取数据的起止时间 =============================

for date in date_range:
    strtime = date.strftime('%Y%m%d%H%M%S')
    
    print('正在读取music接口：国家站日照时数日数据(北京时)：' + strtime)
    baseUrl = url_music
    baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
    baseUrl += '&interfaceId=' + interface_music
    baseUrl += '&ymdhms=' + strtime
    baseUrl += '&cols=' + columns_music
    baseUrl += '&prov=%B9%E3%B6%AB'
    baseUrl += '&dataFormat=html'
    
    strdate = date.strftime('%Y-%m-%d')

    tmp = pd.read_html(baseUrl,encoding='utf-8')
    data_tmp = tmp[0]
    SURF_CLI_S_HOR = data_tmp.iloc[1:]
    SURF_CLI_S_HOR = SURF_CLI_S_HOR.copy()

    for col in SURF_CLI_S_HOR.columns[2:]:
        SURF_CLI_S_HOR[col] = SURF_CLI_S_HOR[col].astype(float)
    SURF_CLI_S_HOR.replace(99,np.NaN,inplace=True)

    # SURF_CLI_S_DAY.columns = columns_out.split(',')

    # SURF_CLI_S_DAY = SURF_CLI_S_DAY.copy()
    SURF_CLI_S_DAY = pd.DataFrame()
    SURF_CLI_S_DAY['stacode'] = SURF_CLI_S_HOR.iloc[:,1]
    SURF_CLI_S_DAY['ddate'] = date
    SURF_CLI_S_DAY['s'] = SURF_CLI_S_HOR.iloc[:,2:].sum(axis='columns')

    SURF_CLI_S_DAY.loc[(SURF_CLI_S_DAY['s'] < 150) & (SURF_CLI_S_DAY['s'] >=0) ,'s'] \
        = 0.1 * SURF_CLI_S_DAY.loc[(SURF_CLI_S_DAY['s'] < 150) & (SURF_CLI_S_DAY['s'] >=0) ,'s']    
    SURF_CLI_S_DAY.loc[(SURF_CLI_S_DAY['s'] >= 150) | (SURF_CLI_S_DAY['s'] <0),'s'] = np.nan
    SURF_CLI_S_DAY.dropna(subset=['s'],inplace=True)

    if SURF_CLI_S_DAY.empty == True:
        print('该日（北京时）无日照时数数据：' + strdate)
        continue
    
    sql = 'select * from ' + Output_TB + ' where DDATE = to_date(\'' + strdate + '\',\'yyyy-mm-dd\')'
    sql += ' and stacode in (select v01301 from ' + StaInfo_TB + ' where v_prcode = \'广东\')'  # 排除香港澳门数据
    SURF_CLI_MUL_DAY = pd.read_sql(sql,Engine) #SQLAchemy return columns name in lowercase letters  
    
    if SURF_CLI_MUL_DAY.empty == True:
        print('该日（北京时）无国家站日数据：' + strdate)
        continue
    
    SURF_CLI_MUL_DAY.drop(columns = ['s'],inplace=True)
    for column in SURF_CLI_MUL_DAY.columns:
        if column[-4:] != 'time':
            if column not in ('stacode','ddate','d_iymdhm'):
                SURF_CLI_MUL_DAY[column] = SURF_CLI_MUL_DAY[column].astype(float)
    
    SURF_CLI_MUL_DAY = pd.merge(SURF_CLI_MUL_DAY, SURF_CLI_S_DAY, on =['ddate','stacode'], how='left')
    SURF_CLI_MUL_DAY['d_iymdhm'] = datetime.datetime.now()

    print('正在写入数据库：国家站日数据（北京时）：' + strdate)
    sql = 'delete from ' + Output_TB
    sql += ' where DDATE = to_date(\'' + strdate + '\',\'yyyy-mm-dd\')'
    sql += ' and stacode in (select v01301 from ' + StaInfo_TB + ' where v_prcode = \'广东\')'
    Engine.execute(sql)
    
    # write the DataFrame data into Oracle database
    SURF_CLI_MUL_DAY.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict
    
print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ':完成')