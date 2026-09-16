# -*- coding: utf-8 -*-
"""
Created on Mon Jun 19 14:27:46 2023

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
interface_music = config['INTERFACE']['SURF_CLI_MUL_DAY']
columns_music = config['COLUMNS_IN']['SURF_CLI_MUL_DAY']
columns_out = config['COLUMNS_OUT']['SURF_CLI_MUL_DAY']
# ============================ read config parameters from ini file ==================================

columns_out = columns_out.split(',')

Output_TB = 'surf_cli_mul_day'
StaInfo_TB = 't_othe_station_meta_basic_tab'

# ======================= 设定读取数据的起止时间 =============================
date_end = datetime.datetime.now() + datetime.timedelta(days=-1)
# date_end = datetime.datetime(1975,4,5) #北京时
date_stt = date_end +datetime.timedelta(days=-2)
# date_stt = datetime.datetime(1975,4,5) #北京时

strdate_end = date_end.strftime('%Y-%m-%d')
strdate_stt = date_stt.strftime('%Y-%m-%d')
date_range = pd.date_range(strdate_end,strdate_stt,freq='-1d')
# ======================= 设定读取数据的起止时间 =============================

import urllib.parse
provs = ['广西','海南','福建','江西','湖南','贵州','云南']#
for prov in provs:
    prov_code = urllib.parse.quote(prov.encode('gb2312'))
    
    for date in date_range:
        strtime = date.strftime('%Y%m%d%H%M%S')
        strdate = date.strftime('%Y-%m-%d')
    
        print('正在读取music接口：国家站全要素日数据(北京时)：' + strdate + '：' + prov)
        baseUrl = url_music
        baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
        baseUrl += '&interfaceId=' + interface_music
        baseUrl += '&ymdhms=' + strtime
        baseUrl += '&cols=' + columns_music
        baseUrl += '&prov=' + prov_code
        baseUrl += '&dataFormat=html'
    
        # reading the aws station info of Guangdong province from music
        tmp = pd.read_html(baseUrl,encoding='utf-8')
        data_tmp = tmp[0]
        data_tmp.fillna('999999',inplace=True)
        DF_Day = data_tmp.iloc[1:]
        if DF_Day.empty == True:
            print('该日（北京时）接口无数据：' + strdate)
            continue
    
        DF_Day.columns = columns_out
        DF_Day = DF_Day.copy()
    
        DF_Day.replace('999998','999999',inplace=True)
        DF_Day.replace('9999', '999999', inplace=True)
        DF_Day.replace('999990','0',inplace=True)
    
        DF_vars = pd.DataFrame()
        DF_vars['stacode'] = DF_Day['stacode'].astype('str')
        # DF_vars['ddate'] = pd.to_datetime(DF_Day[['year','month','day']])
        DF_vars['ddate'] = pd.to_datetime(DF_Day['ddatetime'])
    
        # print('正在处理各要素特征值、格式和数据类型……')
        for column in DF_Day.columns[2:]:
    
            if column[-4:] == 'time':
                # DF_Day[column] = DF_Day[column].astype('str')
                DF_Day.loc[DF_Day[column] != '999999', column] = \
                    DF_Day.loc[DF_Day[column] != '999999', column].str.zfill(4)
    
                tmp_hh = DF_Day.loc[DF_Day[column] != '999999', column].str[0:2]
                tmp_hh[tmp_hh > '23'] = '23'  # 原始数据偶有错误，此处质量控制一下
                tmp_mm = DF_Day.loc[DF_Day[column] != '999999', column].str[-2:]
                tmp_mm[tmp_mm > '59'] = '59'  # 原始数据偶有错误，此处质量控制一下
                tmp_time = tmp_hh + ':' + tmp_mm
                tmp_time = pd.to_datetime(tmp_time, format='%H:%M') + datetime.timedelta(hours=8)
                tmp_time = tmp_time.dt.strftime('%H:%M')
    
                DF_Day.loc[DF_Day[column] != '999999', column] = tmp_time
                DF_Day.loc[DF_Day[column] == '999999', column] = ''
                DF_vars[column] = DF_Day[column]
    
            else:
                if column in ('u_min', 'fzx', 'fjx', 'v_min'):
                    DF_Day[column] = DF_Day[column].astype('int64')
                elif column in ('u'):
                    DF_Day[column] = DF_Day[column].astype('float')  # 需要先转成float
                    DF_Day[column] = DF_Day[column].round(0)  # 才能四舍五入取整
                    DF_Day[column] = DF_Day[column].astype('int64')  # 不带小数的才能转为int
                else:
                    DF_Day[column] = DF_Day[column].astype('float')
                DF_Day.loc[DF_Day[column] >= 99999, column] = np.nan
                DF_vars[column] = DF_Day[column]
    
        sql = 'select stacode,ddate,td,f10s,v from ' + Output_TB + ' where DDATE = to_date(\'' + strdate + '\',\'yyyy-mm-dd\')'#,s
        sql += ' and stacode in (select v01301 from ' + StaInfo_TB + ' where v_prcode = \'' + prov+ '\')'   # 排除香港澳门数据
        SURF_CLI_MUL_DAY = pd.read_sql(sql, Engine)  # SQLAchemy return columns name in lowercase letters
        SURF_CLI_MUL_DAY['td'] = SURF_CLI_MUL_DAY['td'].astype('float')
        SURF_CLI_MUL_DAY['f10s'] = SURF_CLI_MUL_DAY['f10s'].astype('float')
        SURF_CLI_MUL_DAY['v'].fillna(999999, inplace=True)
        SURF_CLI_MUL_DAY['v'] = SURF_CLI_MUL_DAY['v'].astype('int64')  # 空值不能转int，所以先填充
        SURF_CLI_MUL_DAY['v'].replace(999999,np.nan,inplace=True)
        # SURF_CLI_MUL_DAY['s'] = SURF_CLI_MUL_DAY['s'].astype('float')
    
        SURF_CLI_MUL_DAY = pd.merge(DF_vars, SURF_CLI_MUL_DAY, on=['ddate', 'stacode'], how='left')
        SURF_CLI_MUL_DAY['t_dtr'] = SURF_CLI_MUL_DAY['t_max'] - SURF_CLI_MUL_DAY['t_min']
    
        SURF_CLI_MUL_DAY['d_iymdhm'] = datetime.datetime.now()
        SURF_CLI_MUL_DAY.drop_duplicates(subset=['stacode'], keep='first', inplace=True)

        print('正在写入数据库：国家站接口日数据（北京时）：' + strdate)
        sql = 'delete from ' + Output_TB
        sql += ' where DDATE = to_date(\'' + strdate + '\',\'yyyy-mm-dd\')'
        sql += ' and (stacode in (select v01301 from ' + StaInfo_TB + ' where v_prcode = \'' + prov+ '\')' 
        sql += ' or stacode not in (select v01301 from ' + StaInfo_TB + ' where v02301 = \'%A%\'))' 
        Engine.execute(sql)        
    
        # write the DataFrame data into Oracle database
        SURF_CLI_MUL_DAY.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict
    
print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ':完成')