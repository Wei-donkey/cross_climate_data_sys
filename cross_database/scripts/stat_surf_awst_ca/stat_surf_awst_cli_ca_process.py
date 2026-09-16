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
Input_TB = 'surf_awst_cli_ca' 
Output_TB = 'surf_awst_cli_ca_process'

sql_Stacode = 'select v01301 stacode from ' + StaInfo_TB
sql_Stacode += ' order by v01301'
Stacodes = pd.read_sql(sql_Stacode,Engine,index_col='stacode')

if datetime.datetime.now().month >= 7: #  如果当前是下半年
    current_year = datetime.datetime.now().year
elif datetime.datetime.now().month <= 6: #  如果当前是上半年
    current_year = datetime.datetime.now().year - 1
years = range(current_year,current_year-1,-1)
# years = range(2022,2021,-1)

for year in years:
# ======================= 设定更新日数据的起止日期 =============================
    strdate_stt = datetime.datetime(year,7,1).strftime('%Y-%m-%d')
    strdate_end = datetime.datetime(year+1,6,30).strftime('%Y-%m-%d')
# ======================= 设定更新日数据的起止日期 =============================
        
    for idx_sta in Stacodes.index:    
        DF_CA_out = pd.DataFrame(columns=['stacode','date_stt','date_end','idays','tv_24h','tv_acc','t_min','t','tv_48h','ilevel'])#,'RS_degree'])        
    
        stacode = idx_sta
        sql = 'select * from ' + Input_TB + ' where stacode =\'' + stacode + '\''
        sql += ' and ddate between to_date(\'' + strdate_stt +'\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end +'\',\'yyyy-mm-dd\')'
        sql += ' order by ddate'
        data_tmp = pd.read_sql(sql,Engine)       
        
        if data_tmp.empty == False:
            print('统计'+ str(year) + '年度CA过程：' + stacode)
            data_tmp = data_tmp.copy()
            # data_tmp.replace(999999,np.nan,inplace=True)
            
            idx0 = data_tmp.index[0]
            stt_date = data_tmp.loc[idx0,'ddate']
            end_date = data_tmp.loc[idx0,'ddate']
            t_acc = data_tmp.loc[idx0,'t']
            tv_24h = data_tmp.loc[idx0,'tv_24h'] 
            tv_acc = data_tmp.loc[idx0,'tv_24h'] 
            t_min = data_tmp.loc[idx0,'t_min'] 
            tv_48h = data_tmp.loc[idx0,'tv_48h'] 
            # ilevel = data_tmp.loc[idx0,'ilevel'] 
            
            for idx in data_tmp.index[1:]:
                current_date = data_tmp.loc[idx,'ddate']
                
                if end_date + datetime.timedelta(days=1) == current_date: # 冷空气过程延续
                    end_date = current_date
                    t_acc += data_tmp.loc[idx, 't']
                    if tv_24h < data_tmp.loc[idx,'tv_24h']:
                        tv_24h = data_tmp.loc[idx,'tv_24h']
                    if data_tmp.loc[idx,'tv_24h']>0: 
                        tv_acc += data_tmp.loc[idx,'tv_24h'] 
                    if t_min > data_tmp.loc[idx,'t_min']:
                        t_min = data_tmp.loc[idx,'t_min']
                    if tv_48h < data_tmp.loc[idx,'tv_48h']:
                        tv_48h = data_tmp.loc[idx,'tv_48h']
                    # if ilevel < data_tmp.loc[idx,'ilevel']:
                    #     ilevel = data_tmp.loc[idx,'ilevel']
                        
                    if idx == data_tmp.index[-1]: # 最后一条记录
                        idays = (end_date - stt_date).days + 1
                        t = round(t_acc/idays + 0.00001,1) # 四舍五入
    
                        # ============ 不同等级冷空气标准有重叠，可能导致强冷空气被判断为中等，因此从弱到强进行判断 ============
                        if (tv_24h > 0) & (tv_24h < 2): ilevel = 0
                        if (tv_24h>=2) & (tv_24h<4): ilevel = 1
        
                        if (tv_24h>=4) & (tv_24h<6): ilevel = 2    
                        if (tv_24h>=6) & (tv_24h<8) & (t_min>7): ilevel = 2  
                        if (tv_48h>=8) & (tv_48h<10) & (t_min>7): ilevel = 2                
        
                        if (tv_24h>=8) & (t_min>5): ilevel = 3
                        if (tv_24h>=6) & (tv_24h<8) & (t_min<=7): ilevel = 3
                        if (tv_48h>=10) & (t_min>5): ilevel = 3
                        if (tv_48h>=8) & (tv_48h<10) & (t_min<=7): ilevel = 3  
                        
                        if (tv_24h>=8) & (t_min<=5): ilevel= 4
                        if (tv_48h>=10) & (t_min<=5): ilevel= 4
                        # =================================================================================================                    
    
                        print(stacode + ' 冷空气过程：' + stt_date.strftime('%Y-%m-%d'))
                        DF_CA_out = DF_CA_out.append({'stacode':stacode,'date_stt':stt_date,'date_end':end_date,'idays':idays\
                                          ,'tv_24h':tv_24h,'tv_acc':tv_acc,'t_min':t_min,'t':t\
                                          ,'tv_48h':tv_48h,'ilevel':ilevel},ignore_index=True)     
                        
                          
                else: # 冷空气过程结束
                    idays = (end_date - stt_date).days + 1
                    t = round(t_acc/idays + 0.00001,1) # 四舍五入
                    
                    # ============ 不同等级冷空气标准有重叠，可能导致强冷空气被判断为中等，因此从弱到强进行判断 ============
                    if (tv_24h > 0) & (tv_24h < 2): ilevel = 0
                    if (tv_24h >= 2) & (tv_24h<4): ilevel = 1
    
                    if (tv_24h>=4) & (tv_24h<6): ilevel = 2    
                    if (tv_24h>=6) & (tv_24h<8) & (t_min>7): ilevel = 2  
                    if (tv_48h>=8) & (tv_48h<10) & (t_min>7): ilevel = 2                
    
                    if (tv_24h>=8) & (t_min>5): ilevel = 3
                    if (tv_24h>=6) & (tv_24h<8) & (t_min<=7): ilevel = 3
                    if (tv_48h>=10) & (t_min>5): ilevel = 3
                    if (tv_48h>=8) & (tv_48h<10) & (t_min<=7): ilevel = 3  
                    
                    if (tv_24h>=8) & (t_min<=5): ilevel= 4
                    if (tv_48h>=10) & (t_min<=5): ilevel= 4
                    # =================================================================================================
                    
                    print(stacode + ' 冷空气过程：' + stt_date.strftime('%Y-%m-%d'))
                    DF_CA_out = DF_CA_out.append({'stacode':stacode,'date_stt':stt_date,'date_end':end_date,'idays':idays\
                                      ,'tv_24h':tv_24h,'tv_acc':tv_acc,'t_min':t_min,'t':t\
                                      ,'tv_48h':tv_48h,'ilevel':ilevel},ignore_index=True)
                    
                    # 新的冷空气过程参数初始化  
                    stt_date = current_date
                    end_date = current_date
                    t_acc = data_tmp.loc[idx,'t'] 
                    tv_24h = data_tmp.loc[idx,'tv_24h']
                    tv_acc = data_tmp.loc[idx,'tv_24h'] 
                    t_min = data_tmp.loc[idx,'t_min']  
                    tv_48h = data_tmp.loc[idx,'tv_48h'] 
                    # ilevel = data_tmp.loc[idx,'ilevel']
                    
                    if idx == data_tmp.index[-1]:
                        idays = (end_date - stt_date).days + 1
                        t = round(t_acc/idays + 0.00001,1) # 四舍五入
    
                        # ============ 不同等级冷空气标准有重叠，可能导致强冷空气被判断为中等，因此从弱到强进行判断 ============
                        if (tv_24h > 0) & (tv_24h < 2): ilevel = 0
                        if (tv_24h>=2) & (tv_24h<4): ilevel = 1
        
                        if (tv_24h>=4) & (tv_24h<6): ilevel = 2    
                        if (tv_24h>=6) & (tv_24h<8) & (t_min>7): ilevel = 2  
                        if (tv_48h>=8) & (tv_48h<10) & (t_min>7): ilevel = 2                
        
                        if (tv_24h>=8) & (t_min>5): ilevel = 3
                        if (tv_24h>=6) & (tv_24h<8) & (t_min<=7): ilevel = 3
                        if (tv_48h>=10) & (t_min>5): ilevel = 3
                        if (tv_48h>=8) & (tv_48h<10) & (t_min<=7): ilevel = 3  
                        
                        if (tv_24h>=8) & (t_min<=5): ilevel= 4
                        if (tv_48h>=10) & (t_min<=5): ilevel= 4
                        # =================================================================================================                    
                        
                        print(stacode + ' 冷空气过程：' + stt_date.strftime('%Y-%m-%d'))
                        DF_CA_out = DF_CA_out.append({'stacode':stacode,'date_stt':stt_date,'date_end':end_date,'idays':idays\
                                          ,'tv_24h':tv_24h,'tv_acc':tv_acc,'t_min':t_min,'t':t\
                                          ,'tv_48h':tv_48h,'ilevel':ilevel},ignore_index=True)
                    
    
            # for col in DF_CA_out.columns[2:]:
            #     DF_CA_out.loc[DF_CA_out[col].isnull(),col] = 999999
    
            DF_CA_out['d_iymdhm'] = datetime.datetime.now()
            
            sql = 'delete from ' + Output_TB 
            sql += ' where stacode =\'' + stacode + '\''
            sql += ' and date_stt between to_date(\'' + strdate_stt +'\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end +'\',\'yyyy-mm-dd\')'
            Engine.execute(sql)      
        
            print('正在写入数据库：' + Output_TB)
            # write the DataFrame data into Oracle database
            DF_CA_out.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')