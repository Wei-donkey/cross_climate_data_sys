# -*- coding: utf-8 -*-
"""
Created on Mon Nov 23 14:27:46 2020

@author: HP
"""



import datetime
import sys

import pandas as pd
import numpy as np

from pathlib import Path
script_dir = Path(__file__).resolve().parent
project_dir = script_dir.parent
sys.path.insert(0, str(project_dir))

from common.db import get_engine
from common.config import load_ini

# ============================ read configuration parameters from ini file ==================================
Engine_out = get_engine('CROSS_CLIMATE')
Engine_in = get_engine('CROSS_HOUR')

config = load_ini('Config_Vars.ini')
vars_in = config['COLUMNS_IN']['SURF_CLI_V_DAY']
vars_out = config['COLUMNS_OUT']['SURF_CLI_V_DAY']
stats = config['STAT']['SURF_CLI_V_DAY']
thresholds = config['THRESHOLD']['SURF_CLI_V_DAY']
# ============================ read config parameters from ini file ==================================

vars_in = vars_in.split(',')
vars_out = vars_out.split(',')
stats = stats.split(',')
thresholds = thresholds.split(',')

Input_TB = 'surf_cli_mul_hor' 
Output_TB = 'surf_cli_mul_day'

stacode = '59094'
# ======================= 设定更新日数据的起止日期 =============================
date_end = datetime.datetime.now()
date_end = datetime.datetime(1969,12,31) #北京时
strdate_end = date_end.strftime('%Y-%m-%d')
date_stt = date_end +datetime.timedelta(days=-1)
date_stt = datetime.datetime(1969,12,31) #北京时
strdate_stt = date_stt.strftime('%Y-%m-%d')
date_range = pd.date_range(strdate_end,strdate_stt,freq='-1d')
# ======================= 设定更新日数据的起止日期 =============================

for date in date_range:    
    strdate = date.strftime('%Y-%m-%d')
    stryear = date.strftime('%Y')
    DF_V = pd.DataFrame()
    
    for var_out in vars_out:

        strtime_stt = (date + datetime.timedelta(hours=-3)).strftime('%Y-%m-%d %H')
        strtime_end = (date + datetime.timedelta(hours=20)).strftime('%Y-%m-%d %H')
        stryear_stt = strtime_stt[0:4]
        stryear_end = strtime_end[0:4]
        
        var_in = vars_in[vars_out.index(var_out)]
        threshold = thresholds[vars_out.index(var_out)]
        threshold_min = threshold.split('to')[0]
        threshold_max = threshold.split('to')[1]  
        stat = stats[vars_out.index(var_out)]
        
        print(strdate + ': computing daily ' + '- ' + var_in + '- ' + stat)          
            

        sql = 'select stacode,' + stat + '(' + var_in +') ' + var_out 
        sql += ' from ' + Input_TB + '_' + stryear
        sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
        sql += ' and to_char(ddatetime,\'hh24\') in (\'02\',\'08\',\'14\',\'20\')'
        sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
        sql += ' and stacode=\'' + stacode +'\''
        sql += ' group by stacode'
        DF_var = pd.read_sql(sql,Engine_in) 
        DF_var = DF_var.drop_duplicates(['stacode',var_out])  
                
        if DF_var.empty == True: continue

        # 只有平均值需要四舍五入
        DF_var[var_in]=(DF_var[var_in]+0.00001).round(0)


    DF_V = DF_var.copy() 
    DF_V.insert(loc=0,column='ddate',value=date) 

    sql = 'select * from '+ Output_TB + ' where DDATE = to_date(\'' + strdate + '\',\'yyyy-mm-dd\')'
    sql += ' and stacode=\'' + stacode + '\''
    SURF_CLI_MUL_DAY = pd.read_sql(sql,Engine_out) #SQLAchemy return columns name in lowercase letters  
    
    if SURF_CLI_MUL_DAY.empty == True:
        print('该日（北京时）无国家站日数据：' + strdate)
        continue
    
    SURF_CLI_MUL_DAY.drop(columns = ['v'],inplace=True)
    for column in SURF_CLI_MUL_DAY.columns:
        if column[-4:] != 'time':
            if column not in ('stacode','ddate','d_iymdhm'):
                SURF_CLI_MUL_DAY[column] = SURF_CLI_MUL_DAY[column].astype(float)
    
    SURF_CLI_MUL_DAY = pd.merge(SURF_CLI_MUL_DAY, DF_V, on =['ddate','stacode'], how='left')
    
    SURF_CLI_MUL_DAY['d_iymdhm'] = datetime.datetime.now()
   
        
    print('正在写入数据库：' + Output_TB)
    sql = 'delete from ' + Output_TB + ' where ddate = to_date(\'' + strdate +'\',\'yyyy-mm-dd\')'
    sql += ' and stacode=\'' + stacode + '\''
    Engine_out.execute(sql)
    # write the DataFrame data into Oracle database
    SURF_CLI_MUL_DAY.to_sql(Output_TB,Engine_out,index=False,if_exists='append',chunksize=100)

Engine_in.dispose()
Engine_out.dispose()
print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')