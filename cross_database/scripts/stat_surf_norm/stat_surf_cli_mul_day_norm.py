# -*- coding: utf-8 -*-
"""
Created on Mon Nov 23 14:27:46 2020

@author: HP
"""

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

config = load_ini('Config_Vars.ini')
vars_in = config['COLUMNS_IN']['SURF_CLI_MUL_DAY_NORM']
vars_out = config['COLUMNS_OUT']['SURF_CLI_MUL_DAY_NORM']
stats = config['STAT']['SURF_CLI_MUL_DAY_NORM']
thresholds = config['THRESHOLD']['SURF_CLI_MUL_DAY_NORM']
# ============================ read configuration parameters from ini file ==================================

vars_in = vars_in.split(',')
vars_out = vars_out.split(',')
stats = stats.split(',')
thresholds = thresholds.split(',')

Input_TB = 'surf_cli_mul_day' 
Output_TB = 'surf_cli_mul_day_norm'

year_stt, year_end  = 1991, 2020

DF_vars = pd.DataFrame()
  
for var_out in vars_out:
    var_in = vars_in[vars_out.index(var_out)]
    threshold = thresholds[vars_out.index(var_out)]
    threshold_min = threshold.split('to')[0]
    threshold_max = threshold.split('to')[1]   

    print('computing daily climate normals: ' + '- ' + var_out)
    
    sql = 'select stacode,to_char(ddate,\'mm\') IMONTH, to_char(ddate,\'dd\') IDAY, avg(' + var_in +') ' + var_out 
    sql += ' from ' + Input_TB
    sql += ' where to_char(ddate,\'yyyy\') between \'' + str(year_stt) +'\' and \'' + str(year_end) +'\''
    sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
    sql += ' group by to_char(ddate,\'mm\'),to_char(ddate,\'dd\'),stacode'
    sql += ' order by stacode,imonth,iday'
    
    DF_var = pd.read_sql(sql,Engine) 
    
    DF_var[var_out]=(DF_var[var_out]+0.00001).round(1)
    if DF_vars.empty == True:
        DF_vars = DF_var.copy()
    else:
        DF_vars = pd.merge(DF_vars,DF_var,on=['stacode','imonth','iday'],how='outer')

DF_vars['d_iymdhm'] = datetime.datetime.now()
# DF_vars.fillna(-9999,inplace=True)

print('正在写入数据库：' + Output_TB)

DF_vars['stacode'] = DF_vars['stacode'].astype(str)
DF_vars['imonth'] = DF_vars['imonth'].astype(int)
DF_vars['iday'] = DF_vars['iday'].astype(int)
# 剔除2月29日的常年值
DF_vars.drop(DF_vars[(DF_vars['imonth'] == 2) & (DF_vars['iday'] == 29)].index,inplace=True)
for var_out in vars_out:
    DF_vars[var_out] = DF_vars[var_out].astype(float)
    
sql = 'delete from ' + Output_TB
Engine.execute(sql)      
# write the DataFrame data into Oracle database
DF_vars.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')