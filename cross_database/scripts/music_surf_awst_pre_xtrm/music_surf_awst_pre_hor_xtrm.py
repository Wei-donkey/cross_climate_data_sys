# -*- coding: utf-8 -*-
"""
Created on Wed May 27 12:15:27 2020

@author: qhzx
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
Engine = get_engine('CROSS_RAIN')

config = load_ini('Config_MUSIC_GD.ini')
url_music = config['CONNECT_qhztj']['url']
user_music = config['CONNECT_qhztj']['user']
pwd_music = config['CONNECT_qhztj']['password']
interface_music = config['INTERFACE']['STAT_SURF_AWST_PRE_HOR_XTRM']
columns_music = config['COLUMNS_IN']['STAT_SURF_AWST_PRE_HOR_XTRM']
# ============================ read configuration parameters from ini file ==================================

StaInfo_TB = 'climate.t_othe_station_meta_basic_tab'
Output_TB = 'surf_awst_cli_pre_hor_extreme'
Threshold_TB = 'surf_awst_cli_pre_threshold'
RainQC_TB = 'surf_awst_cli_pre_hor_qc'

stacode = 'G1255'

Hour_types = ['1','3','6','12','24']
# 各站点每月输出的极值数量不一致，汛期多，非汛期少
rankings = [10,20,40,50,50,50,50,50,50,40,20,10]
# 小于阈值表雨量最小值的不参与极值统计（若阈值表无该站点记录，则设置<5mm不参与统计）
r_not_stat = 5             
        
# File_log = open('Stat_Rain_Hourly_realtime.log','w',encoding='utf-8')

# ======================= 设定数据统计的起止时间 =============================
time_end = datetime.datetime.now()
# time_end = datetime.datetime(2026,5,21,1) #北京时
time_stt = datetime.datetime.now()
time_stt = datetime.datetime(2026,8,5,8) #北京时

iyear_end = time_end.year; imonth_end = time_end.month; iday_end = time_end.day; ihour_end = time_end.hour
time_end = datetime.datetime(iyear_end,imonth_end,iday_end,ihour_end) + datetime.timedelta(hours=-8) #国际时
strtime_end = time_end.strftime('%Y-%m-%d %H:%M:%S')

iyear_stt = time_stt.year; imonth_stt = time_stt.month; iday_stt = time_stt.day; ihour_stt = time_stt.hour
time_stt = datetime.datetime(iyear_stt,imonth_stt,iday_stt,ihour_stt) + datetime.timedelta(hours=-8) #国际时
strtime_stt = time_stt.strftime('%Y-%m-%d %H:%M:%S')

time_range = pd.date_range(strtime_end,strtime_stt,freq='-1h')
# ======================= 设定数据统计的起止时间 =============================


for time in time_range:
    time_UTC = time    
    # 转换为北京时
    time = time + datetime.timedelta(hours=8)
    
    print('正在读取国家站+区域站滑动雨量：' + time.strftime('%Y-%m-%d %H:%M'))
    
    ddatetime_NewRec = time
    strtime_NewRec = time.strftime('%Y-%m-%d %H:%M')
    month_NewRec = time.month
    
    # 写入极值表的各站点该月极值数量
    ranking_rec = rankings[month_NewRec-1]
    
    strtime_UTC = time_UTC.strftime('%Y%m%d%H%M%S')
    
    # reading the accumulated rainfall data of Guangdong province from music
    baseUrl = url_music
    baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
    baseUrl += '&interfaceId=' + interface_music
    baseUrl += '&cols=' + columns_music
    baseUrl += '&ymdhms='+ strtime_UTC
    baseUrl += '&dataFormat=html'
    baseUrl_A = baseUrl + '&v02301=A'
    baseUrl_B = baseUrl + '&v02301=B'     

    # reading the accumulated rainfall data of Guangdong province from music
    tmp = pd.read_html(baseUrl_A,encoding='utf-8')
    data_tmp = tmp[0]
    data_A = data_tmp.iloc[1:]
    col_names = data_tmp.iloc[0]
    col_names = map(str.lower, col_names)
    data_A.columns = col_names

    tmp = pd.read_html(baseUrl_B,encoding='utf-8')
    data_tmp = tmp[0]
    data_B = data_tmp.iloc[1:]
    col_names = data_tmp.iloc[0]
    col_names = map(str.lower, col_names)
    data_B.columns = col_names
    
    data_NewRec = pd.concat([data_A,data_B],ignore_index= True)
    # 删除一模一样的记录
    data_NewRec = data_NewRec.drop_duplicates()
    
    # ============ 修改返回数据的各列类型======================
    data_NewRec['v13019']=data_NewRec['v13019'].astype('float'); data_NewRec['v13020']=data_NewRec['v13020'].astype('float')
    data_NewRec['v13021']=data_NewRec['v13021'].astype('float'); data_NewRec['v13022']=data_NewRec['v13022'].astype('float')
    data_NewRec['v13023']=data_NewRec['v13023'].astype('float')

    data_NewRec = data_NewRec[data_NewRec['v01301']==stacode].copy()

    if data_NewRec.empty == True:
        print('该时次数据为空：' + time.strftime('%Y-%m-%d %H:%M'))
        # File_log.writelines('该时次数据为空,' + time.strftime('%Y-%m-%d %H:%M:%S')+ '\n')   
        continue


    # ========= 删除1、3、6、12、24滑动雨量均为0的记录行 ===============
    all_zero = (data_NewRec['v13019'] ==0) & (data_NewRec['v13020'] ==0) \
                & (data_NewRec['v13021'] ==0) & (data_NewRec['v13022']==0) & (data_NewRec['v13023']==0) 
    data_NewRec = data_NewRec[~all_zero]
        
    if data_NewRec.empty == True:
        print('该时次数据全为 0：' + time.strftime('%Y-%m-%d %H:%M'))
        # File_log.writelines('该时次数据为空,' + time.strftime('%Y-%m-%d %H:%M:%S')+ '\n')   
        continue

    # ======================= 读取当前月份：所有站点、所有历时的雨量最小值 from 阈值表===========================
    print('正在读取当前月份：所有站点、所有历时最小值')
    sql = 'select hours_type,stacode,r_min,time_min from ' + Threshold_TB
    sql += ' where imonth=' + str(month_NewRec)
    sql += ' and stacode=\'' + stacode + '\''
    data_Min = pd.read_sql(sql,Engine)

    print('正在判断实时极值：' + time.strftime('%Y-%m-%d %H:%M'))
    
    data_DB = pd.DataFrame(columns=['hours_type','ddatetime','stacode','r']) 
    
    for idx in data_NewRec.index:
        stacode = data_NewRec.loc[idx,'v01301']

        for Hour_type in Hour_types:
#            D_IYMDHM = datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')
            
            if Hour_type == '1': Rain_NewRec = data_NewRec.loc[idx,'v13019']
            if Hour_type == '3': Rain_NewRec = data_NewRec.loc[idx,'v13020']
            if Hour_type == '6': Rain_NewRec = data_NewRec.loc[idx,'v13021']
            if Hour_type == '12': Rain_NewRec = data_NewRec.loc[idx,'v13022']
            if Hour_type == '24': Rain_NewRec = data_NewRec.loc[idx,'v13023']
            
            # 该记录为空值，忽略
            if float(Rain_NewRec) == 9999.0:
                continue      
            
            Rain_Min_tmp = data_Min.loc[(data_Min['hours_type']==int(Hour_type)) & (data_Min['stacode']==stacode)]            

            # 极值表中无该站点该月值（区域站可能存在该情形）
            if Rain_Min_tmp.empty == True:  
                Rain_Min = r_not_stat
            else:
                Rain_Min = Rain_Min_tmp['r_min'].values[0]
                strtime_Min =  pd.to_datetime(Rain_Min_tmp['time_min'].values[0]).strftime('%Y-%m-%d %H:%M')
                
            if float(Rain_NewRec) >= float(Rain_Min):
               
                # 同时段起止时间，如12小时降水，则前后分别扩展11个时次，24小时降水，则前后分别扩展23个时次
                time_END_SamePeriod = time + datetime.timedelta(hours=int(Hour_type)- 1) 
                time_STT_SamePeriod = time + datetime.timedelta(hours=-int(Hour_type)+ 1)   

                sql = 'select HOURS_TYPE,STACODE,DDATETIME,R from ' + Output_TB + ' where HOURS_TYPE=' + Hour_type
                sql += ' and STACODE=\'' + stacode + '\''
                sql += ' and DDATETIME between to_date(\'' + time_STT_SamePeriod.strftime('%Y-%m-%d %H:%M') +'\',\'yyyy-mm-dd hh24:mi\')'
                sql += ' and to_date(\'' + time_END_SamePeriod.strftime('%Y-%m-%d %H:%M') +'\',\'yyyy-mm-dd hh24:mi\')'
                result = Engine.execute(sql)
                data_SamePeriod = result.fetchall()
                        
                if len(data_SamePeriod) == 0:
                    # get the extreme value count at current station, current month from the Rain Extreme database
                    sql = 'select count(*) from ' + Output_TB + ' where HOURS_TYPE=' + Hour_type
                    sql += ' and STACODE=\'' + stacode + '\''
                    sql += ' and to_char(DDATETIME,\'mm\')=\'' + str(month_NewRec).zfill(2) + '\''
                    result = Engine.execute(sql)
                    tmp = result.fetchall()
                    count_rec = tmp[0][0]
                    
                    # 如果该站极值数量>=规定的数量，则删除最小值记录1条
                    if count_rec >= ranking_rec:                           
                        #ddatetime_Min = data_Min[0][2].strftime('%Y-%m-%d %H:%M')                                
                        print('删除' + Hour_type + '小时累积雨量最小极值：' + stacode + ': ' + str(Rain_Min) + ', ' + strtime_Min)
                        # File_log.writelines('删除' + Hour_type + '小时累积雨量旧极值,' + stacode + ',' + str(Rain_Min)+ '\n') 
                        
                        sql = 'delete from ' + Output_TB + ' where HOURS_TYPE=' + Hour_type
                        sql += ' and STACODE=\'' + stacode + '\''
                        sql += ' and to_char(DDATETIME,\'mm\')=\'' + str(month_NewRec).zfill(2) + '\''                        
                        sql += ' and R=' + str(Rain_Min)
#                        sql += ' and DDATETIME=to_date(\'' + ddatetime_Min +'\',\'yyyy-mm-dd hh24:mi\')'
                        Engine.execute(sql)   
                        
                    print('插入' + Hour_type + '小时累积雨量新极值：' + stacode + ': ' + str(Rain_NewRec)+ ', ' + strtime_NewRec)
                    # File_log.writelines('插入' + Hour_type + '小时累积雨量新极值,' + stacode  + ',' + str(Rain_NewRec)+ ',' + strtime_NewRec+ '\n') 
                    tmp_list = [Hour_type, ddatetime_NewRec, stacode, Rain_NewRec]
                    tmp_series = pd.Series(tmp_list,index=data_DB.columns)
                    data_DB = data_DB.append(tmp_series,ignore_index=True)

                elif len(data_SamePeriod) == 1: # 此前已存在同一降水过程的雨量极值
                    # 再判断旧极值与新极值哪个大
                    Rain_SamePeriod = data_SamePeriod[0][3]
                    ddatetime_SamePeriod = data_SamePeriod[0][2]#.strftime('%Y-%m-%d %H:%M')#; ihour_SamePeriod = data_SamePeriod[0][3]
                    strtime_SamePeriod = ddatetime_SamePeriod.strftime('%Y-%m-%d %H:%M')
                    
                    # 是该站该历时的同时次记录
                    if ddatetime_NewRec==ddatetime_SamePeriod:                     
                        print('更新' + Hour_type + '小时累积雨量极值：' + stacode + ': ' + str(Rain_NewRec) + ', ' + strtime_NewRec)
                        # File_log.writelines('更新' + Hour_type + '小时累积雨量极值,' + stacode + ',' + str(Rain_NewRec)+ ',' + strtime_NewRec + '\n')
                        tmp_list = [Hour_type, ddatetime_NewRec, stacode, Rain_NewRec]
                        tmp_series = pd.Series(tmp_list,index=data_DB.columns)
                        data_DB = data_DB.append(tmp_series,ignore_index=True)
                            
                    else:
                        if float(Rain_NewRec) > float(Rain_SamePeriod) :                                
                            print('删除' + Hour_type + '小时累积雨量同时段旧极值：' + stacode + ': ' + str(Rain_SamePeriod) + ', ' + strtime_SamePeriod)
                            # File_log.writelines('删除' + Hour_type + '小时累积雨量同时段旧极值,' + stacode + ',' + str(Rain_SamePeriod)+ '\n')
                        
                            # 删除该同一过程旧的降水记录，并插入新的降水记录
                            sql = 'delete from ' + Output_TB + ' where HOURS_TYPE=' + Hour_type
                            sql += ' and STACODE=\'' + stacode + '\''
                            sql += ' and DDATETIME'
                            sql += ' between to_date(\'' + time_STT_SamePeriod.strftime('%Y-%m-%d %H:%M') +'\',\'yyyy-mm-dd hh24:mi\')'
                            sql += ' and to_date(\'' + time_END_SamePeriod.strftime('%Y-%m-%d %H:%M') +'\',\'yyyy-mm-dd hh24:mi\')'
                            Engine.execute(sql)  

                            print('插入' + Hour_type + '小时累积雨量同时段新极值：' + stacode + ': ' + str(Rain_NewRec) + ', ' + strtime_NewRec)
                            # File_log.writelines('插入' + Hour_type + '小时累积雨量同时段新极值,' + stacode + ',' + str(Rain_NewRec) + ',' + strtime_NewRec+ '\n')   
                            tmp_list = [Hour_type, ddatetime_NewRec, stacode, Rain_NewRec]
                            tmp_series = pd.Series(tmp_list,index=data_DB.columns)
                            data_DB = data_DB.append(tmp_series,ignore_index=True)
                        elif float(Rain_NewRec) == float(Rain_SamePeriod): 
                            if ddatetime_NewRec < ddatetime_SamePeriod: 
                                print('删除' + Hour_type + '小时累积雨量同时段旧极值（出现晚）：' + stacode + ': ' + str(Rain_SamePeriod) + ', ' + strtime_SamePeriod)
                                # 删除该同一过程旧的降水记录，并插入新的降水记录
                                sql = 'delete from ' + Output_TB + ' where HOURS_TYPE=' + Hour_type
                                sql += ' and STACODE=\'' + stacode + '\''
                                sql += ' and DDATETIME'
                                sql += ' between to_date(\'' + time_STT_SamePeriod.strftime('%Y-%m-%d %H:%M') +'\',\'yyyy-mm-dd hh24:mi\')'
                                sql += ' and to_date(\'' + time_END_SamePeriod.strftime('%Y-%m-%d %H:%M') +'\',\'yyyy-mm-dd hh24:mi\')'
                                Engine.execute(sql)                               
                                
                                print('插入' + Hour_type + '小时累积雨量同时段新极值：' + stacode + ': ' + str(Rain_NewRec) + ', ' + strtime_NewRec)
                                # File_log.writelines('插入' + Hour_type + '小时累积雨量同时段新极值,' + stacode + ',' + str(Rain_NewRec) + ',' + strtime_NewRec+ '\n')   
                                tmp_list = [Hour_type, ddatetime_NewRec, stacode, Rain_NewRec]
                                tmp_series = pd.Series(tmp_list,index=data_DB.columns)
                                data_DB = data_DB.append(tmp_series,ignore_index=True)                        
                               
                        
                elif len(data_SamePeriod) ==2: # 前后相关历时已存在2条雨量极值
                    Rain_SamePeriod1 = data_SamePeriod[0][3]
                    ddatetime_SamePeriod1 = data_SamePeriod[0][2]
                    strtime_SamePeriod1 = ddatetime_SamePeriod1.strftime('%Y-%m-%d %H:%M')
                    
                    Rain_SamePeriod2 = data_SamePeriod[1][3]
                    ddatetime_SamePeriod2 = data_SamePeriod[1][2]
                    strtime_SamePeriod2 = ddatetime_SamePeriod2.strftime('%Y-%m-%d %H:%M')

                    if (float(Rain_NewRec) > float(Rain_SamePeriod1)) and (float(Rain_NewRec) > float(Rain_SamePeriod2)) :       
                        print('删除' + Hour_type + '小时累积雨量同时段旧极值:' + stacode + ': ' + str(Rain_SamePeriod1) + ', ' + strtime_SamePeriod1)       
                        print('删除' + Hour_type + '小时累积雨量同时段旧极值:' + stacode + ': ' + str(Rain_SamePeriod2) + ', ' + strtime_SamePeriod2)
                        # File_log.writelines('删除' + Hour_type + '小时累积雨量同时段旧极值,' + stacode + ','  + str(Rain_SamePeriod1) + '\n')
                        # File_log.writelines('删除' + Hour_type + '小时累积雨量同时段旧极值,' + stacode + ','  + str(Rain_SamePeriod2)+ '\n')
                    
                        # 删除该同一过程旧的降水记录，并插入新的降水记录
                        sql = 'delete from ' + Output_TB + ' where HOURS_TYPE=' + Hour_type
                        sql += ' and STACODE=\'' + stacode + '\''
                        sql += ' and DDATETIME'
                        sql += ' between to_date(\'' + time_STT_SamePeriod.strftime('%Y-%m-%d %H:%M') +'\',\'yyyy-mm-dd hh24:mi\')'
                        sql += ' and to_date(\'' + time_END_SamePeriod.strftime('%Y-%m-%d %H:%M') +'\',\'yyyy-mm-dd hh24:mi\')'                                    
                        Engine.execute(sql)  

                        print('插入' + Hour_type + '小时累积雨量同时段新极值：' + stacode  + ': ' + str(Rain_NewRec)+ ', ' + strtime_NewRec)
                        # File_log.writelines('插入' + Hour_type + '小时累积雨量同时段新极值,' + stacode  + ',' + str(Rain_NewRec)+ ',' + strtime_NewRec+ '\n') 
                        tmp_list = [Hour_type, ddatetime_NewRec, stacode, Rain_NewRec]
                        tmp_series = pd.Series(tmp_list,index=data_DB.columns)
                        data_DB = data_DB.append(tmp_series,ignore_index=True)                        

    if data_DB.empty == False: 
        data_DB['hours_type']=data_DB['hours_type'].astype('int')        
        data_DB['ddatetime']=data_DB['ddatetime'].astype('datetime64')
#        data_DB['d_iymdhm']=data_DB['d_iymdhm'].astype('datetime64')  
        
        data_DB['qc'] = '-9999'
        data_DB['qc_manual'] = 0
        data_DB['r_original'] = -9999.0
        data_DB['d_iymdhm'] = datetime.datetime.now()
   

        for Hour_type in Hour_types:
            width = int(Hour_type)
            data_DB.loc[data_DB['hours_type']==int(Hour_type),'qc']='0'.rjust(width,'0')      
      
        sql = 'delete from ' + Output_TB + ' where ddatetime= to_date(\'' + strtime_NewRec +'\',\'yyyy-mm-dd hh24:mi\')'
        sql += ' and stacode=\'' + stacode + '\''
        Engine.execute(sql)
        
        print('正在写入数据库：' + Output_TB)
        data_DB.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict


print('完成统计')
# File_log.close()