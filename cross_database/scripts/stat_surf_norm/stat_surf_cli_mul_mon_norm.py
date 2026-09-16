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
vars_in = config['COLUMNS_IN']['SURF_CLI_MUL_RESTAT_NORM']
vars_out = config['COLUMNS_OUT']['SURF_CLI_MUL_RESTAT_NORM']
stats = config['STAT']['SURF_CLI_MUL_RESTAT_NORM']
# ============================ read configuration parameters from ini file ==================================

vars_in = vars_in.split(',')
vars_out = vars_out.split(',')

Input_TB = 'surf_cli_mul_mon' 
Output_TB = 'surf_cli_mul_mon_norm'

year_stt, year_end  = 1991, 2020

DF_vars = pd.DataFrame()
  
for var_out in vars_out:
    var_in = vars_in[vars_out.index(var_out)]  

    print('computing monthly climate normals: ' + '- ' + var_out)
    
    sql = 'select stacode,imonth, ' + var_in +' ' + var_out
    sql += ' from ' + Input_TB
    sql += ' where iyear between \'' + str(year_stt) +'\' and \'' + str(year_end) +'\''

    DF_raw_data = pd.read_sql(sql,Engine)
    DF_raw_data[var_out]=(DF_raw_data[var_out]+0.00001).round(1)  # 对原始数据进行四舍五入

    group_tmp = DF_raw_data.groupby(['stacode','imonth'])
    DF_var = group_tmp.agg('mean').reset_index()  # 将多索引变成多列
    DF_var[var_out] = (DF_var[var_out] + 0.00001).round(1)  # 对结果数据进行四舍五入

    if DF_vars.empty == True:
        DF_vars = DF_var.copy()
    else:
        DF_vars = pd.merge(DF_vars,DF_var,on=['stacode','imonth'],how='outer')
 
DF_vars['d_iymdhm'] = datetime.datetime.now()

print('正在写入数据库：' + Output_TB)

DF_vars['stacode'] = DF_vars['stacode'].astype(str)
DF_vars['imonth'] = DF_vars['imonth'].astype(int)

for var_out in vars_out:
    DF_vars[var_out] = DF_vars[var_out].astype(float)
    
sql = 'delete from ' + Output_TB
Engine.execute(sql)      
# write the DataFrame data into Oracle database
DF_vars.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')