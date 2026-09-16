# -*- coding: utf-8 -*-
"""
Created on Mon Mar 4 14:27:46 2022

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
Engine = get_engine('CROSS_CLIMATE')

config = load_ini('Config_Vars.ini')
vars_in = config['COLUMNS_IN']['SURF_CLI_MUL_RESTAT']
vars_out = config['COLUMNS_OUT']['SURF_CLI_MUL_RESTAT']
stats = config['STAT']['SURF_CLI_MUL_RESTAT']
thresholds = config['THRESHOLD']['SURF_CLI_MUL_RESTAT']
# ============================ read config parameters from ini file ==================================

vars_in = vars_in.split(',')
vars_out = vars_out.split(',')
stats = stats.split(',')
thresholds = thresholds.split(',')

Input_TB = 'surf_cli_mul_day' 
Output_TB = 'surf_cli_mul_yer'

# ======================= 设定更新年数据的起止月份 =============================
iyears = 1
current_date = datetime.datetime.now()
current_date =datetime.datetime(2026,12,31) # 手动设定日期（包括本年）
iyear_end = current_date.year
# ======================= 设定更新年数据的起止月份 =============================

# ================== 构建一个所有站点对应的空DataFrame ====================
StaInfo_TB = 't_othe_station_meta_basic_tab'
sql = 'select v01301 stacode from '+ StaInfo_TB + ' where v02301 like \'%A%\' and v_prcode in (\'广东\',\'香港\',\'澳门\') order by v01301'
DF_stacode = pd.read_sql(sql,Engine)
# =======================================================================

for i in range(iyears):
    iyear = iyear_end - i
    stryyyy = str(iyear)    
    date_stt = datetime.datetime(iyear,1,1)
    date_end = datetime.datetime(iyear,12,31)       
        
    # 当前日期尚未进入该季度/季节，则统计下一个季度/季节
    if current_date <= date_stt: 
        continue
        
    strdate_stt = date_stt.strftime('%Y-%m-%d')
    strdate_end = date_end.strftime('%Y-%m-%d')

    DF_vars = DF_stacode.copy()
    
    for var_out in vars_out:
        var_in = vars_in[vars_out.index(var_out)]
        threshold = thresholds[vars_out.index(var_out)]
        threshold_min = threshold.split('to')[0]
        threshold_max = threshold.split('to')[1]  
        stat = stats[vars_out.index(var_out)]
        
        print(stryyyy  + ': computing yearly ' + '- ' + var_out)

        if stat not in ('max','min'):    
            sql = 'select stacode,' + stat + '(' + var_in +') ' + var_out 
            sql += ' from ' + Input_TB
            sql += ' where ddate between to_date(\'' + strdate_stt +'\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end +'\',\'yyyy-mm-dd\')'
            sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
            sql += ' group by stacode'
        elif stat in ('max','min'):
            sql = 'select a.stacode,a.' + var_out + ',to_char(b.ddate,\'mm-dd\') ' + var_out + 'date from'
            sql += ' (select stacode,' + stat + '(' + var_in +') ' + var_out 
            sql += ' from ' + Input_TB
            sql += ' where ddate between to_date(\'' + strdate_stt +'\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end +'\',\'yyyy-mm-dd\')'
            sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
            sql += ' group by stacode) a'
            sql += ' left join'
            sql += ' (select stacode,ddate,' + var_in 
            sql += ' from ' + Input_TB
            sql += ' where ddate between to_date(\'' + strdate_stt +'\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end +'\',\'yyyy-mm-dd\')'
            sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max + ') b'
            sql += ' on a.stacode=b.stacode and a.' + var_out + '= b.' + var_in                      
     
        DF_var = pd.read_sql(sql,Engine)
        if DF_var.empty != True:
            # 只有平均值需要四舍五入保留两位小数
            if stat == 'avg':
                DF_var[var_out]=(DF_var[var_out]+0.00001).round(1)
            if stat in ('max','min'):
                 DF_var.drop_duplicates(['stacode',var_out],inplace=True)
                 
        DF_vars = pd.merge(DF_vars,DF_var,on=['stacode'],how='left')
        # 该要素有效数据量为空的站点，值应该为0;2024-9-11：保留cnt字段为空，以便删除全为空的记录
        # if var_out[-4:]  == 'cnts':
        # DF_vars.loc[pd.isnull(DF_vars[var_out]),var_out] = 0

        if DF_var.empty == True:
            continue

    # 2024-9-11 设stacode为index，以便删除全为空值的站点
    DF_vars.set_index('stacode', inplace=True)
    DF_vars.dropna(axis='index', how='all', inplace=True)
    DF_vars.reset_index(inplace=True)

    # 2024-9-13 cnts字段不能为null，否则会导致cross出错
    for var_out in vars_out:
        if var_out[-4:] == 'cnts':
            DF_vars.loc[pd.isnull(DF_vars[var_out]), var_out] = 0

    # 判断var_days/sum对应的var_cnts是否为0，如果为0，则var_days/sum空值保持空；如果不为0，则var_days/sum变为0
    print('通过_cnts字段判断日数和累积量的正确赋值……')
    for var_out in vars_out:
        var_in = vars_in[vars_out.index(var_out)]
        if var_in in ['t','u','p','v','b','d'] : var_in = var_in + 'a'
        if var_in[-3:] == 'max': var_in = var_in[0] + 'x'
        if var_in[-3:] == 'min': var_in = var_in[0] + 'n'
        if (var_out[-4:] == 'days') or (var_out[-3:] == 'sum'):
            var_cnts = var_in + '_cnts'
            DF_vars.loc[pd.isnull(DF_vars[var_out]) & (DF_vars[var_cnts] == 0),var_out] = np.nan
            DF_vars.loc[pd.isnull(DF_vars[var_out]) & (DF_vars[var_cnts] != 0),var_out] = 0

        # if var_out[-4:] == 'cnts':
        #     DF_vars[var_out] = DF_vars[var_out].astype('int64')
        # else:
        DF_vars[var_out] = DF_vars[var_out].astype('float')

    DF_vars.insert(loc=1,column='iyear',value=int(stryyyy))
    DF_vars['d_iymdhm'] = datetime.datetime.now()      

    print('正在写入数据库：' + Output_TB)
    sql = 'delete from ' + Output_TB + ' where iyear = ' + stryyyy
    sql += ' and stacode in (select v01301 from ' + StaInfo_TB + ' where v_prcode in (\'广东\',\'香港\',\'澳门\'))'
    Engine.execute(sql)      
    # write the DataFrame data into Oracle database
    DF_vars.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')