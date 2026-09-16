# -*- coding: utf-8 -*-
"""
Created on Wed Feb 8 14:15:51 2023

@author: Administrator
"""

import sys

from pathlib import Path
script_dir = Path(__file__).resolve().parent
project_dir = script_dir.parent
sys.path.insert(0, str(project_dir))

from common.db import get_engine

# ============================ read configuration parameters from ini file ==================================
Engine_in = get_engine('CROSS_RAIN')
Engine_out = get_engine('CROSS_RAIN')
# ============================ read configuration parameters from ini file ==================================

import numpy as np
import datetime
import pandas as pd

current_year = datetime.datetime.now().year
years = range(current_year,current_year-2,-1)
# years = range(2023,1950,-1)
StaInfo_TB = 'climate.t_othe_station_meta_basic_tab'
Input_TB2 = 'climate.surf_cli_mul_day'
Input_TB1 = 'surf_cli_pre_rsod'
Output_TB = 'reg_cli_pre_rsd'
prov = '广东'

for year in years:
    flag = False # 是否开汛
    if year==current_year: 
        types = ['R', 'R08']  #
    else:
        types = ['R']
        
    print('正在统计 ' + str(year) + '年：广东省开汛日期')

    for type in types:

        sql = 'select RSOD_' + type + ' from ' + Input_TB1
        sql += ' where iyear =' + str(year)
        sql += ' order by RSOD_' + type
        DF_tmp = pd.read_sql(sql, Engine_in)

        rows, cols = DF_tmp.shape
        sta_cnt_50pct = round(rows/2)  # 50%的站点数量
        sta_cnt_10pct = round(rows / 10)  # 10%的站点数量

        RSOD = DF_tmp.iloc[sta_cnt_50pct-1,0]
        if not pd.isnull(RSOD):  # 已经有50%站点开汛了，接下来判断哪一天及其前一天有10%站点R或R08>=38

            # ======================= 设定滚动查询日雨量的起止日期 =============================
            date_stt = RSOD
            strdate_stt = date_stt.strftime('%Y-%m-%d')
            date_end = datetime.datetime(year,6,30) #北京时
            strdate_end = date_end.strftime('%Y-%m-%d')
            date_range = pd.date_range(strdate_stt, strdate_end, freq='1d')
            # ======================= 设定更新日数据的起止日期 =============================

            for date in date_range:
                strdate_stt2 = (date + datetime.timedelta(days=-1)).strftime('%Y-%m-%d')
                strdate_end2 = date.strftime('%Y-%m-%d')

                sql = 'select stacode from ' + Input_TB2
                sql += ' where ddate between to_date(\'' + strdate_stt2 + '\',\'yyyy-mm-dd\')'
                sql += ' and to_date(\'' + strdate_end2 + '\',\'yyyy-mm-dd\')'
                sql += ' and ' + type + '>=38'
                sql += ' and stacode IN'
                sql += ' (SELECT v01301 FROM ' + StaInfo_TB 
                sql += ' WHERE v_prcode=\'' +prov +'\' and stacode<>\'59486\')'
                sql += ' group by stacode'

                DF_tmp = pd.read_sql(sql, Engine_in)
                rows, cols = DF_tmp.shape
                if rows >= sta_cnt_10pct:
                    RSOD = date
                    flag = True
                    break

        if flag==True:  # 如果用R为标准已开汛，则不再判断R08
            break

    if flag==True:
        strRSOD = RSOD.strftime('%m-%d')
        DF_out = pd.DataFrame([[str(year),strRSOD,None,type]],columns=['IYEAR','RSOD','RSED','R_TYPE'])
    else:
        DF_out = pd.DataFrame([[str(year), None, None, None]], columns=['IYEAR', 'RSOD', 'RSED', 'R_TYPE'])

    DF_out['D_IYMDHM'] = datetime.datetime.now()
    DF_out.insert(loc=0,column='CREGION',value=prov)

    sql = 'delete from ' + Output_TB + ' where iyear=\'' + str(year) + '\' and cregion=\'' + prov + '\''
    Engine_out.execute(sql)

    print('正在写入数据库：' + Output_TB)
    DF_out.to_sql(Output_TB,Engine_out,index=False,if_exists='append',chunksize=100)

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')