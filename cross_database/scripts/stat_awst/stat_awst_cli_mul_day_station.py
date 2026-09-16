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
vars_in = config['COLUMNS_IN']['AWST_CLI_MUL_DAY']
vars_out = config['COLUMNS_OUT']['AWST_CLI_MUL_DAY']
stats = config['STAT']['AWST_CLI_MUL_DAY']
thresholds = config['THRESHOLD']['AWST_CLI_MUL_DAY']
# ============================ read config parameters from ini file ==================================

vars_in = vars_in.split(',')
vars_out = vars_out.split(',')
stats = stats.split(',')
thresholds = thresholds.split(',')

Input_TB = 'awst_cli_mul_hor'
Output_TB = 'awst_cli_mul_day'

# ======================= 设定更新日数据的起止日期 =============================
date_end = datetime.datetime.now()
date_end = datetime.datetime(2025,9,19) #北京时
strdate_end = date_end.strftime('%Y-%m-%d')
date_stt = date_end +datetime.timedelta(days=-1)
date_stt = datetime.datetime(2025,9,19) #北京时
strdate_stt = date_stt.strftime('%Y-%m-%d')
date_range = pd.date_range(strdate_end,strdate_stt,freq='-1d')
# ======================= 设定更新日数据的起止日期 =============================

stacode = '81601595'
# ================== 构建一个所有站点对应的空DataFrame ====================
StaInfo_TB = 'climate.t_othe_station_meta_basic_tab'
sql = 'select v01301 stacode from '+ StaInfo_TB + ' where v01301=\'' + stacode + '\''
DF_stacode = pd.read_sql(sql,Engine_in)
# =======================================================================

for date in date_range:    
    strdate = date.strftime('%Y-%m-%d')
    stryear = date.strftime('%Y')
    DF_vars = DF_stacode.copy()
    
    for var_out in vars_out:
        if var_out == 'r08' : 
            strtime_stt = (date + datetime.timedelta(hours=9)).strftime('%Y-%m-%d %H')
            strtime_end = (date + datetime.timedelta(hours=32)).strftime('%Y-%m-%d %H')
        elif var_out == 'r20_08' : 
            strtime_stt = (date + datetime.timedelta(hours=-3)).strftime('%Y-%m-%d %H')
            strtime_end = (date + datetime.timedelta(hours=8)).strftime('%Y-%m-%d %H')    
        elif var_out == 'r08_20' : 
            strtime_stt = (date + datetime.timedelta(hours=9)).strftime('%Y-%m-%d %H')
            strtime_end = (date + datetime.timedelta(hours=20)).strftime('%Y-%m-%d %H')
        else:
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
            
        if stat == 'avg':
            sql = 'select stacode,' + stat + '(' + var_in +') ' + var_out 
            sql += ' from ' + Input_TB + '_' + stryear
            sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
            sql += ' and to_char(ddatetime,\'hh24\') in (\'02\',\'08\',\'14\',\'20\')'
            sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
            sql += ' and stacode=\'' + stacode + '\''
            sql += ' group by stacode'
            DF_var = pd.read_sql(sql,Engine_in) 
            DF_var = DF_var.drop_duplicates(['stacode',var_out])  

        elif stat == 'sum':     
            if stryear_stt == stryear_end: 
                sql = 'select stacode,' + stat + '(' + var_in +') ' + var_out 
                sql += ' from ' + Input_TB + '_' + stryear                
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' and stacode=\'' + stacode + '\''
                sql += ' group by stacode'
                DF_var = pd.read_sql(sql,Engine_in)
                DF_var = DF_var.drop_duplicates(['stacode',var_out])  

            if stryear_stt != stryear_end:
                sql = 'select stacode,' + stat + '(' + var_in +') ' + var_out
                sql += ' from ('
                sql += 'select stacode,ddatetime,' + var_in
                sql += ' from ' + Input_TB + '_' + stryear_stt
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and stacode=\'' + stacode + '\''
                sql += ' union '
                sql += 'select stacode,ddatetime,' + var_in
                sql += ' from ' + Input_TB + '_' + stryear_end
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and stacode=\'' + stacode + '\''
                sql += ')'
                sql += ' where ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' group by stacode'
                DF_var = pd.read_sql(sql,Engine_in)
                DF_var = DF_var.drop_duplicates(['stacode',var_out])  
                
        elif stat in ('max','min'):
            if var_in not in ('fjs','fzs'): var_in2 = var_in + 'time'
            if var_in == 'fzs' : var_in2 = var_in + 'time,fzx'
            if var_in == 'fjs' : var_in2 = var_in + 'time,fjx'
            
            if stryear_stt == stryear_end:    
                sql = 'select a.stacode,a.' + var_out + ',' + var_in2
                sql += ' from'
                sql += ' (select stacode,' + stat + '(' + var_in +') ' + var_out 
                sql += ' from ' + Input_TB + '_' + stryear
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' and stacode=\'' + stacode + '\''
                sql += ' group by stacode) a'
                sql += ' left join'
                sql += ' (select stacode,' + var_in + ',' +  var_in2     
                sql += ' from ' + Input_TB + '_' + stryear
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' and stacode=\'' + stacode + '\''
                sql += ') b'
                sql += ' on a.stacode=b.stacode and a.' + var_out + '= b.' + var_in      
                DF_var = pd.read_sql(sql,Engine_in)
                DF_var.drop_duplicates(['stacode',var_out],inplace=True)  
                
            if stryear_stt != stryear_end: # 分别读取前一年和当年的数据写入DF_var1和DF_var2    

                sql = 'select a.stacode,a.' + var_out + ',' + var_in2
                sql += ' from'
                sql += ' (select stacode,' + stat + '(' + var_in +') ' + var_out 
                sql += ' from ' + Input_TB + '_' + stryear_stt
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' and stacode=\'' + stacode + '\''
                sql += ' group by stacode) a'
                sql += ' left join'
                sql += ' (select stacode,' + var_in + ',' +  var_in2
                sql += ' from ' + Input_TB + '_' + stryear_stt
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' and stacode=\'' + stacode + '\''
                sql += ') b'
                sql += ' on a.stacode=b.stacode and a.' + var_out + '= b.' + var_in      
                DF_var1 = pd.read_sql(sql,Engine_in) 
                
                sql = 'select a.stacode,a.' + var_out + ',' + var_in2
                sql += ' from'
                sql += ' (select stacode,' + stat + '(' + var_in +') ' + var_out 
                sql += ' from ' + Input_TB + '_' + stryear_end
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' and stacode=\'' + stacode + '\''
                sql += ' group by stacode) a'
                sql += ' left join'
                sql += ' (select stacode,' + var_in + ',' +  var_in2
                sql += ' from ' + Input_TB + '_' + stryear_end
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' and stacode=\'' + stacode + '\''
                sql += ') b'
                sql += ' on a.stacode=b.stacode and a.' + var_out + '= b.' + var_in      
                DF_var2 = pd.read_sql(sql,Engine_in) 
                
                DF_var = pd.concat([DF_var1,DF_var2])  
                DF_var.sort_values(by=[var_out],ascending=False,inplace = True)
                DF_var.drop_duplicates(['stacode'],inplace=True) 
                
        if DF_var.empty == True: 
            DF_vars = pd.merge(DF_vars,DF_var,on=['stacode'],how='left')
            continue
        
        # ============================== 判断有效数据数量，数量不足的，统计结果用nan替代 ==================================
        if stat == 'avg':
            sql = 'select stacode,count(*) cnt from ' + Input_TB  + '_' + stryear  
            sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
            sql += ' and to_char(ddatetime,\'hh24\') in (\'02\',\'08\',\'14\',\'20\')'
            sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
            sql += ' and stacode=\'' + stacode + '\''
            sql += ' group by stacode'

        else:
            if stryear_stt == stryear_end:     
                sql = 'select stacode,count(*) cnt from ' + Input_TB  + '_' + stryear  
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' and stacode=\'' + stacode + '\''
                sql += ' group by stacode'
            if stryear_stt != stryear_end:    
                sql = 'select stacode,count(*) cnt'
                sql += ' from ('
                sql += 'select stacode,ddatetime,' + var_in
                sql += ' from ' + Input_TB + '_' + stryear_stt
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and stacode=\'' + stacode + '\''
                sql += ' union '
                sql += 'select stacode,ddatetime,' + var_in
                sql += ' from ' + Input_TB + '_' + stryear_end
                sql += ' where ddatetime between to_date(\'' + strtime_stt +'\',\'yyyy-mm-dd hh24\') and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24\')'
                sql += ' and stacode=\'' + stacode + '\''
                sql += ')'
                sql += ' where ' + var_in + ' between ' + threshold_min + ' and ' + threshold_max
                sql += ' group by stacode'            
        
        DF_cnt= pd.read_sql(sql,Engine_in)
        DF_var_cnt = pd.merge(DF_var,DF_cnt,on=['stacode'],how='left')
        if stat == 'avg':
            DF_var_cnt.loc[DF_var_cnt['cnt']<4,var_out] = np.nan
        else:
            DF_var_cnt.loc[DF_var_cnt['cnt']<1,var_out] = np.nan
        DF_var = DF_var_cnt.drop(columns='cnt')
        # ============================== 判断有效数据数量，数量不足的，统计结果用nan替代 ==================================
        
        # 只有平均值需要四舍五入
        if stat == 'avg':
            if var_in in ('u','v'): 
                DF_var[var_in]=(DF_var[var_in]+0.00001).round(0)
            else: 
                DF_var[var_in]=(DF_var[var_in]+0.00001).round(1)

        DF_vars = pd.merge(DF_vars,DF_var,on=['stacode'],how='left')

    # 2024-9-11 设stacode为index，以便删除全为空值的站点
    DF_vars.set_index('stacode',inplace=True)
    DF_vars.dropna(axis='index',how='all',inplace=True)
    DF_vars.reset_index(inplace=True)

    # 读取日照时数日数据
    sql = 'select stacode,s from '+ Output_TB + ' where ddate = to_date(\'' + strdate +'\',\'yyyy-mm-dd\')'
    sql += ' and stacode=\'' + stacode + '\''
    DF_S = pd.read_sql(sql,Engine_out) #SQLAchemy return columns name in lowercase letters  
    if DF_S.empty == True: DF_S['s']=DF_S['s'].astype(float)
    DF_vars = pd.merge(DF_vars, DF_S, on ='stacode', how='left')
    
    DF_vars['t_dtr'] = DF_vars['t_max'] - DF_vars['t_min']    
    DF_vars.insert(loc=0,column='ddate',value=date) 

    for column in DF_vars.columns[2:]: 
        if column[-4:] != 'time':
            DF_vars[column] = DF_vars[column].astype('float')

    DF_vars['d_iymdhm'] = datetime.datetime.now()    
        
    print('正在写入数据库：' + Output_TB)
    sql = 'delete from ' + Output_TB + ' where ddate = to_date(\'' + strdate +'\',\'yyyy-mm-dd\')'
    sql += ' and stacode=\'' + stacode + '\''
    Engine_out.execute(sql)      
    # write the DataFrame data into Oracle database
    DF_vars.to_sql(Output_TB,Engine_out,index=False,if_exists='append',chunksize=1000)

Engine_in.dispose()
Engine_out.dispose()
print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')