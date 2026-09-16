# -*- coding: utf-8 -*-
"""
Created on Mon Jun 21 14:27:46 2023

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
interface_music = config['INTERFACE']['SURF_CLI_MUL_DAY_UPDATE']
columns_music = config['COLUMNS_IN']['SURF_CLI_MUL_DAY']
columns_out = config['COLUMNS_OUT']['SURF_CLI_MUL_DAY']
# ============================ read config parameters from ini file ==================================

columns_out = columns_out.split(',')

Output_TB = 'surf_cli_mul_day'
StaInfo_TB = 't_othe_station_meta_basic_tab'

import urllib.parse
provs = ['广西','海南','福建','江西','湖南','贵州','云南']
for prov in provs:
    prov_code = urllib.parse.quote(prov.encode('gb2312'))
    # the url encoding for '广东' is %B9%E3%B6%AB
    
    for i in range(1,6):  #更新过去6个时次有变化的数据
        # ======================= 读取上一时次更新的数据 =============================
        time_1hour = datetime.datetime.now() + datetime.timedelta(hours=-i)
        # time_1hour = datetime.datetime(2023,6,11,12) #北京时
        date_renew = datetime.datetime.now() + datetime.timedelta(days=-1)  # 如果当天前的数据发生了更新，则这些数据重新写入数据库
        date_renew = datetime.datetime(date_renew.year,date_renew.month,date_renew.day)
    
    
        time_end = datetime.datetime(time_1hour.year,time_1hour.month,time_1hour.day,time_1hour.hour,59,59)  # 当前时次的最后时刻
        time_stt = datetime.datetime(time_1hour.year,time_1hour.month,time_1hour.day,time_1hour.hour,0,0)  # 当前时次的最初时刻
    
        strtime_1hour = time_1hour.strftime('%Y-%m-%d %H')
        # ======================= 读取当前更新的数据 =============================
    
        s_ymdhms = (time_stt + datetime.timedelta(hours=-8)).strftime('%Y%m%d%H%M%S')  # 国际时
        e_ymdhms = (time_end + datetime.timedelta(hours=-8)).strftime('%Y%m%d%H%M%S')  # 国际时
    
        print('正在读取music接口：国家站全要素(北京时)：' + strtime_1hour + '时更新的日数据：' + prov)
        baseUrl = url_music
        baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
        baseUrl += '&interfaceId=' + interface_music
        baseUrl += '&s_ymdhms=' + s_ymdhms
        baseUrl += '&e_ymdhms=' + e_ymdhms
        baseUrl += '&cols=' + columns_music
        baseUrl += '&prov=' + prov_code
        baseUrl += '&dataFormat=html'
    
        # reading the aws station info of Guangdong province from music
        tmp = pd.read_html(baseUrl,encoding='utf-8')
        data_tmp = tmp[0]
        data_tmp.fillna('999999',inplace=True)
        DF_Day = data_tmp.iloc[1:]
    
        if DF_Day.empty == True:
            print( strtime_1hour +'（北京时）无数据更新')
    
        else:
            DF_Day.columns = columns_out
            DF_Day = DF_Day.copy()
            DF_Day['ddatetime'] = DF_Day['ddatetime'].astype('datetime64')
            DF_Day = DF_Day.loc[DF_Day['ddatetime']<=date_renew]
    
            if DF_Day.empty == True:
                print( strtime_1hour + '（北京时）未更新1日前的数据')
    
            else:
                DF_Day.replace('999998','999999',inplace=True)
                DF_Day.replace('9999', '999999', inplace=True)
                DF_Day.replace('999990','0',inplace=True)
    
                DF_vars = pd.DataFrame()
                DF_vars['stacode'] = DF_Day['stacode'].astype('str')
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
    
                for idx in DF_vars.index:
                    tmp_DF = DF_vars.loc[[idx]] # []内为[](list)，则结果为DataFrame
                    for column in tmp_DF.columns[2:]:
                        if column[-4:] != 'time':
                            if column in ('u', 'u_min', 'fzx', 'fjx', 'v_min'):
                                tmp_DF[column].fillna(999999, inplace=True)
                                tmp_DF[column] = tmp_DF[column].astype('int64')
                                tmp_DF[column].replace(999999, np.nan, inplace=True)
                            else:
                                tmp_DF[column] = tmp_DF[column].astype('float')
    
                    tmp_stacode= DF_vars.loc[idx,'stacode']
                    tmp_ddate = DF_vars.loc[idx,'ddate']
                    strtmp_ddate = tmp_ddate.strftime('%Y-%m-%d')
    
                    sql = 'select stacode,ddate,td,f10s,v from ' + Output_TB#,s
                    sql += ' where DDATE = to_date(\'' + strtmp_ddate + '\',\'yyyy-mm-dd\')'
                    sql += ' and stacode =\'' + tmp_stacode + '\''
    
                    SURF_CLI_MUL_DAY = pd.read_sql(sql , Engine)
                    SURF_CLI_MUL_DAY['td'] = SURF_CLI_MUL_DAY['td'].astype('float')
                    SURF_CLI_MUL_DAY['f10s'] = SURF_CLI_MUL_DAY['f10s'].astype('float')
                    SURF_CLI_MUL_DAY['v'].fillna(999999, inplace=True)
                    SURF_CLI_MUL_DAY['v'] = SURF_CLI_MUL_DAY['v'].astype('int64')  # 空值不能转int，所以先填充
                    SURF_CLI_MUL_DAY['v'].replace(999999,np.nan,inplace=True)
                    # SURF_CLI_MUL_DAY['s'] = SURF_CLI_MUL_DAY['s'].astype('float')
    
                    SURF_CLI_MUL_DAY = pd.merge(tmp_DF, SURF_CLI_MUL_DAY, on=['ddate', 'stacode'], how='left')
                    SURF_CLI_MUL_DAY['t_dtr'] = SURF_CLI_MUL_DAY['t_max'] - SURF_CLI_MUL_DAY['t_min']
                    SURF_CLI_MUL_DAY['d_iymdhm'] = datetime.datetime.now()
    
                    print('正在写入数据库：：' + strtmp_ddate + '日数据（北京时）：' + tmp_stacode)
                    sql = 'delete from ' + Output_TB
                    sql += ' where DDATE = to_date(\'' + strtmp_ddate + '\',\'yyyy-mm-dd\')'
                    sql += ' and stacode =\'' + tmp_stacode + '\''
                    Engine.execute(sql)
    
                    # write the DataFrame data into Oracle database
                    SURF_CLI_MUL_DAY.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ':完成')