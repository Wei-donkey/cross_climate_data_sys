# -*- coding: utf-8 -*-
"""
Created on Thu Mar 18 14:15:51 2021

@author: Administrator
"""

import sys

from pathlib import Path
script_dir = Path(__file__).resolve().parent
project_dir = script_dir.parent
sys.path.insert(0, str(project_dir))

from common.db import get_engine

# ============================ read configuration parameters from ini file ==================================
Engine = get_engine('CROSS_CLIMATE')
# ============================ read configuration parameters from ini file ==================================

import numpy as np
import datetime
import pandas as pd

# sta_types = ['surf','awst']
StaInfo_TB = 't_othe_station_meta_basic_tab'
Output_TB = 'surf_awst_cli_ca' 
Input_TB1 = 'surf_cli_mul_day' 
Input_TB2 = 'awst_cli_mul_day'

sql_Stacode = 'select v01301 stacode,slm from ' + StaInfo_TB
sql_Stacode += ' order by v01301'
Stacodes = pd.read_sql(sql_Stacode,Engine,index_col='stacode')

# 统计这段时间内各站点降温日的相关指标
def Stat_CA(date_stt,date_end):
    date_stt = date_stt + datetime.timedelta(days=-2)

    strdate_stt = date_stt.strftime('%Y-%m-%d')
    strdate_end = date_end.strftime('%Y-%m-%d')

    # 输出的相关统计量
    DF_CA_out = pd.DataFrame(columns=['stacode', 'ddate', 't', 't_min', 'tv_24h', 'tv_48h'])  # ,'ilevel'])

    idx = pd.date_range(strdate_stt, strdate_end, freq='1d')
    idx_0d = idx[2:]  # 当前日序
    idx_1d = idx[1:-1]  # 前一天日序
    idx_2d = idx[:-2]  # 前两天日序

    sql = 'select ddate,stacode,t_min,t from ' + Input_TB1
    sql += ' where ddate between to_date(\'' + strdate_stt + '\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end + '\',\'yyyy-mm-dd\')'
    sql += ' UNION '
    sql += 'select ddate,stacode,t_min,t from ' + Input_TB2
    sql += ' where ddate between to_date(\'' + strdate_stt + '\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end + '\',\'yyyy-mm-dd\')'

    DF_tmp = pd.read_sql(sql, Engine, index_col='ddate')
    # ======================= 质量控制 ==============================
    DF_tmp.loc[(DF_tmp['t'] < -10) | (DF_tmp['t'] > 50), 't'] = np.nan
    DF_tmp.loc[(DF_tmp['t_min'] < -10) | (DF_tmp['t_min'] > 50), 't_min'] = np.nan
    DF_tmp.loc[(DF_tmp['t'] - DF_tmp['t_min'] <= 0) | (DF_tmp['t'] - DF_tmp['t_min'] > 15), ['t', 't_min']] = np.nan
    DF_tmp.dropna(inplace=True)  # t或t_min为空，则删除 2022-4-28

    for idx_sta in Stacodes.index:
        stacode = idx_sta
        # stacode = '57989'
        data_tmp = DF_tmp[DF_tmp['stacode'] == stacode]
        if data_tmp.empty == False:
            print('统计截至' + strdate_end + '每日CA：' + stacode)
            data_tmp = data_tmp.copy()
            data_tmp.replace(-9999, np.nan, inplace=True)
            Data_sta = data_tmp.reindex(idx)  # 完整序列数据

            Data_sta = data_tmp.reindex(idx_0d)  # 当日数据

            T_sta_1d = data_tmp.reindex(idx_1d)['t']  # 前1日平均气温
            T_sta_1d.index += datetime.timedelta(days=1)

            T_sta_2d = data_tmp.reindex(idx_2d)['t']  # 前2日平均气温
            T_sta_2d.index += datetime.timedelta(days=2)

            # 前一日24h降温=前2日平均气温-前1日平均气温
            TV_24h_1d = T_sta_2d - T_sta_1d
            # 前一日升温：则将前一日24h降温赋值为0（为了保证当前日降温情况下：当前日48h降温>0）
            TV_24h_1d[TV_24h_1d < 0] = 0
            # 当前日24h降温
            Data_sta['tv_24h'] = T_sta_1d - Data_sta['t']
            # 为了精度，对浮点数运算结果进行四舍五入
            Data_sta['tv_24h'] = round(Data_sta['tv_24h'] + 0.00001, 1)
            # 当前日48h降温
            Data_sta['tv_48h'] = TV_24h_1d + Data_sta['tv_24h']
            # 为了精度，对浮点数运算结果进行四舍五入
            Data_sta['tv_48h'] = round(Data_sta['tv_48h'] + 0.00001, 1)
            Data_sta['ca'] = 0  # 添加默认冷空气等级-无

            # 如果24h变温>0，则当日一定属于冷空气过程
            Data_sta.loc[Data_sta['tv_24h'] > 0, 'ca'] = 1
            # 如果24h变温为0，但前一日24h有降温，则当日依然算弱冷空气
            Data_sta.loc[(Data_sta['tv_24h'] == 0) & (TV_24h_1d > 0), 'ca'] = 1

            Data_sta.reset_index(level=0, inplace=True)
            Data_sta = Data_sta.rename(columns={'index': 'ddate'})
            Data_sta = Data_sta[Data_sta['ca'] != 0]  # 仅存储冷空气过程指标
            Data_sta = Data_sta.drop(columns=['ca'])
            # Data_sta.dropna(inplace=True)
            # 48h变温为空，则用24h变温替代（2022-4-28）
            Data_sta.loc[Data_sta['tv_48h'].isnull(), 'tv_48h'] = Data_sta.loc[Data_sta['tv_48h'].isnull(), 'tv_24h']

            DF_CA_out = pd.concat([DF_CA_out, Data_sta], ignore_index=True)

    DF_CA_out['d_iymdhm'] = datetime.datetime.now()

    date_stt = date_stt + datetime.timedelta(days=2)
    strdate_stt = date_stt.strftime('%Y-%m-%d')
    sql = 'delete from ' + Output_TB
    sql += ' where ddate between to_date(\'' + strdate_stt + '\',\'yyyy-mm-dd\') and  to_date(\'' + strdate_end + '\',\'yyyy-mm-dd\')'
    Engine.execute(sql)

    print('正在写入数据库：' + Output_TB)
    DF_CA_out.to_sql(Output_TB, Engine, index=False, if_exists='append', chunksize=100)  # ,dtype=dtypedict


# ======================= 历史统计：设定更新月数据的起止月份 =============================
imonths = 15 # 更新近1个月统计值（1947-2021共900月）
current_date = datetime.datetime.now()
# current_date =datetime.datetime(2022,8,31) # 手动设定日期
iyear = 2022  # current_date.year
imonth = 12  # current_date.month
for i in range(imonths):
    ThisMon_year = iyear + (imonth-i-1)//12; ThisMon_mon = (imonth-i)%12
    if ThisMon_mon == 0: ThisMon_mon = 12
    NextMon_year = iyear + (imonth-i)//12; NextMon_mon = (imonth-i+1)%12
    if NextMon_mon == 0: NextMon_mon = 12

    if ThisMon_mon in (6,7,8): continue

    date_stt = datetime.datetime(ThisMon_year,ThisMon_mon,1)
    date_end = datetime.datetime(NextMon_year,NextMon_mon,1) + datetime.timedelta(days=-1)
# ======================= 设定更新日数据的起止日期 =============================
    Stat_CA(date_stt,date_end)


    
print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')