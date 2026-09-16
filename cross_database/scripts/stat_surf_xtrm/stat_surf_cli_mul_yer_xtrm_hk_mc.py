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
vars_in = config['COLUMNS_IN']['SURF_CLI_MUL_RESTAT_XTRM']
vars_out = config['COLUMNS_OUT']['SURF_CLI_MUL_RESTAT_XTRM']
stats = config['STAT']['SURF_CLI_MUL_RESTAT_XTRM']
# thresholds = config['THRESHOLD']['SURF_CLI_MUL_RESTAT_XTRM']
# ============================ read configuration parameters from ini file ==================================

vars_in = vars_in.split(',')
vars_out = vars_out.split(',')
stats = stats.split(',')
# thresholds = thresholds.split(',')

StaInfo_TB = 't_othe_station_meta_basic_tab'
Input_TB = 'surf_cli_mul_yer'
Output_TB = 'surf_cli_mul_yer_xtrm'

# ======================= 设定更新日数据的起止年份 =============================
# year_stt, year_end  = 2025, 2025
year_stt = datetime.datetime.now().year-1; year_end = year_stt
# ======================= 设定更新日数据的起止年份 =============================

for iyear in range(year_stt,year_end+1):

    print('正在读取累年极值数据表……')
    sql = 'select * from ' + Output_TB
    sql += ' where stacode in (select v01301 from ' + StaInfo_TB + ' where v_prcode in (\'香港\',\'澳门\'))'

    DF_Xtrm = pd.read_sql(sql,Engine)  
    for col in DF_Xtrm.columns[1:-1]:
        DF_Xtrm[col] = DF_Xtrm[col].astype(float)   
    
    print('================================================')
    print('computing yearly historic extremes: ' + str(iyear))
    for var_out in vars_out:
        var_in = vars_in[vars_out.index(var_out)]
        stat = stats[vars_out.index(var_out)] 
        
        print('computing yearly historic extremes: ' + var_out)
        
        DF_tmp=DF_Xtrm.loc[:,['stacode',var_out,var_out + 'year']]
        # 删除极值为空的记录
        DF_tmp.dropna(axis='index',subset=[var_out],inplace=True)
    
        sql = 'select stacode, ' + stat + '(' + var_in +') ' + var_out
        sql += ' from ' + Input_TB
        sql += ' where iyear= ' + str(iyear)
        sql += ' and stacode in (select v01301 from ' + StaInfo_TB + ' where v_prcode in (\'香港\',\'澳门\'))'
        sql += ' group by stacode'
        
        DF_var = pd.read_sql(sql,Engine)
        DF_var.dropna(axis='index', subset=[var_out], inplace=True)
        if DF_var.empty == False:
            DF_var[var_out + 'year'] = iyear
            DF_var['stacode'] = DF_var['stacode'].astype(str)
            
            if DF_tmp.empty == True:
                DF_tmp = DF_var.copy()
            else:
                DF_var.rename(columns={var_out:var_out + '_new',var_out+'year':var_out + 'year_new'},inplace=True)
                DF_tmp = pd.merge(DF_tmp,DF_var,on=['stacode'],how='outer')
                # merge之后，之前的年份会出现空值（无法与当前年份值比大小），用当前年份值填充之前的空值
                DF_tmp.loc[pd.isnull(DF_tmp[var_out]),var_out+'year'] = DF_tmp[var_out+'year_new']
                DF_tmp.loc[pd.isnull(DF_tmp[var_out]),var_out] = DF_tmp[var_out+'_new']
            
                if stat == 'max':
                    DF_tmp.loc[DF_tmp[var_out+'_new']>=DF_tmp[var_out],var_out+'year']= iyear
                    DF_tmp.loc[DF_tmp[var_out+'_new']>=DF_tmp[var_out],var_out]= DF_tmp[var_out+ '_new']
                if stat == 'min':
                    DF_tmp.loc[DF_tmp[var_out+'_new']<=DF_tmp[var_out],var_out+'year']= iyear
                    DF_tmp.loc[DF_tmp[var_out+'_new']<=DF_tmp[var_out],var_out]= DF_tmp[var_out+ '_new']
                DF_tmp.drop(columns=[var_out+'_new',var_out+'year_new'],inplace=True)     
                
            DF_Xtrm.drop(columns=[var_out,var_out + 'year'],inplace=True)
            DF_Xtrm = pd.merge(DF_Xtrm,DF_tmp,on=['stacode'],how='outer')
 
    DF_Xtrm['d_iymdhm'] = datetime.datetime.now()
    print('正在写入数据库：' + Output_TB)
    
    DF_Xtrm['stacode'] = DF_Xtrm['stacode'].astype(str)
    for var_out in vars_out:
        DF_Xtrm[var_out] = DF_Xtrm[var_out].astype(float)
        DF_Xtrm[var_out + 'year'] = DF_Xtrm[var_out + 'year'].astype(float)   
    
    sql = 'delete from ' + Output_TB
    sql += ' where stacode in (select v01301 from ' + StaInfo_TB + ' where v_prcode in (\'香港\',\'澳门\'))'
    Engine.execute(sql)
    DF_Xtrm.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')