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
years = range(current_year,current_year-1,-1)
# years = range(2023,1950,-1)
StaInfo_TB = 'climate.t_othe_station_meta_basic_tab'
Input_TB = 'climate.surf_cli_mul_day'
Output_TB = 'surf_cli_pre_rsod'
prov = '广东'

types = ['R', 'R08']

sql_Stacode = 'select v01301 stacode, stt_date from ' + StaInfo_TB
sql_Stacode += ' where V02301 like \'%A%\' and V02301 not like \'%B%\' and v_prcode=\'' + prov + '\''
sql_Stacode += ' and V01301<>\'59486\''
sql_Stacode += ' order by v01301'
Stacodes = pd.read_sql(sql_Stacode,Engine_in,index_col='stacode')

for year in years:
    print('正在统计 ' + str(year) + '年：各站点开汛日期')

    date_stt = datetime.datetime(year,3,1)  #北京时
    strdate_stt = date_stt.strftime('%Y-%m-%d')
    date_end = datetime.datetime(year,6,30)  #北京时
    strdate_end = date_end.strftime('%Y-%m-%d')

    for type in types:
        DF_out_tmp = pd.DataFrame(columns=['stacode', 'RSOD_'+type, type])

        sql = 'select ddate,stacode,' + type + ' from ' + Input_TB
        sql += ' where ddate between to_date(\'' + strdate_stt +'\',\'yyyy-mm-dd\')'
        sql += ' and to_date(\'' + strdate_end +'\',\'yyyy-mm-dd\')'
        sql += ' and ' + type + '>=38'
        sql += ' order by stacode, ddate'
        DF_tmp = pd.read_sql(sql, Engine_in)

        for index in Stacodes.index:
            stacode = index
            stt_year = Stacodes.loc[index,'stt_date'].strftime('%Y')

            data_sta = DF_tmp[DF_tmp['stacode'] == stacode]
            if year >= int(stt_year):
                if data_sta.empty == False:
                    RSOD = data_sta.iloc[0,0]
                    R = data_sta.iloc[0,2]
                    DF_out_tmp = DF_out_tmp.append({'stacode': stacode, 'RSOD_'+type: RSOD, type: R}, ignore_index=True)
                else:
                    # RSOD = pd.NaT
                    # R = float(np.nan)
                    DF_out_tmp = DF_out_tmp.append({'stacode': stacode}, ignore_index=True)

        if type=='R':
            DF_out = DF_out_tmp
        if type=='R08':
            DF_out = pd.merge(DF_out,DF_out_tmp,on='stacode')

    DF_out['RSOD_R'] = DF_out['RSOD_R'].astype('datetime64')
    DF_out['RSOD_R08'] = DF_out['RSOD_R08'].astype('datetime64')
    DF_out['R'] = DF_out['R'].astype('float')
    DF_out['R08'] = DF_out['R08'].astype('float')
    DF_out['D_IYMDHM'] = datetime.datetime.now()
    DF_out.insert(loc=1,column='IYEAR',value=year)

    sql = 'delete from ' + Output_TB + ' where iyear=\'' + str(year) + '\''
    Engine_out.execute(sql)

    print('正在写入数据库：' + Output_TB)
    DF_out.to_sql(Output_TB,Engine_out,index=False,if_exists='append',chunksize=100)

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')