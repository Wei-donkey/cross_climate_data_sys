# -*- coding: utf-8 -*-
"""
Created on Wed May  6 08:46:21 2020
modified on April 6 2021
@author: qhzx
"""


import datetime
stt_now= datetime.datetime.now()
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
interface_music = config['INTERFACE']['QC_SURF_AWST_PRE_HOR']
columns_music = config['COLUMNS_IN']['QC_SURF_AWST_PRE_HOR']
# ============================ read configuration parameters from ini file ==================================

# 设置小时降水量的气候界限值
Rain_limit = 200

StaInfo_TB = 'climate.t_othe_station_meta_basic_tab'
StaNeigbors_TB = 'climate.t_othe_station_neigbors_tab'
RainQC_TB = 'surf_awst_cli_pre_hor_qc'

# ======================= 设定质量控制数据的起止时间 =============================
time_end = datetime.datetime.now()
# time_end = datetime.datetime(2020,6,9,18) #北京时
time_stt = time_end +datetime.timedelta(hours=-5)
# time_stt = datetime.datetime(2021,4,9,0) #北京时

iyear_end = time_end.year; imonth_end = time_end.month; iday_end = time_end.day; ihour_end = time_end.hour
time_end = datetime.datetime(iyear_end,imonth_end,iday_end,ihour_end) + datetime.timedelta(hours=-8) #国际时
strtime_end = time_end.strftime('%Y-%m-%d %H:%M:%S')

iyear_stt = time_stt.year; imonth_stt = time_stt.month; iday_stt = time_stt.day; ihour_stt = time_stt.hour
time_stt = datetime.datetime(iyear_stt,imonth_stt,iday_stt,ihour_stt) + datetime.timedelta(hours=-8) #国际时
strtime_stt = time_stt.strftime('%Y-%m-%d %H:%M:%S')

time_range = pd.date_range(strtime_stt,strtime_end,freq='1h')
# ======================= 设定质量控制数据的起止时间 =============================

# 删除 Q13019=0 的所有记录（这些记录源于被中断运行的本脚本）
sql = 'delete from ' + RainQC_TB + ' where q13019=0'
Engine.execute(sql)

# 首先从music读取待质量控制的多个时次小时雨量，存入数据库
sta_types = ['A','B']
for time in time_range:
    strtime = time.strftime('%Y-%m-%d %H:%M:%S')        
    
    # 删除当前时次所有记录
    sql = 'delete from ' + RainQC_TB + ' where ddatetime= to_date(\'' + strtime + '\',\'yyyy-mm-dd hh24:mi:ss\')'
    Engine.execute(sql)

    for sta_type in sta_types:
        strtime = time.strftime('%Y-%m-%d %H:%M:%S')
        
        if sta_type == 'A': print('正在读取国家站小时雨量(国际时)：' + strtime)
        if sta_type == 'B': print('正在读取区域站小时雨量(国际时)：' + strtime)
     
        strtime = time.strftime('%Y%m%d%H%M%S')
   
        # reading the accumulated rainfall data of Guangdong province from music
        baseUrl = url_music
        baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
        baseUrl += '&interfaceId=' + interface_music
        baseUrl += '&cols=' + columns_music
        baseUrl += '&ymdhms='+ strtime + '&v02301=' + sta_type
        baseUrl += '&dataFormat=html'
        
        # reading the accumulated rainfall data of Guangdong province from music
        tmp = pd.read_html(baseUrl,encoding='utf-8')
        data_tmp = tmp[0]
    
        col_names = data_tmp.iloc[0]
        data_tmp = data_tmp.iloc[1:]
        col_names = map(str.lower, col_names)
        data_tmp.columns = col_names    
        data_NewRec = data_tmp #[(data_tmp['v02301']=='A') | (data_tmp['v02301']=='B')]
        
    #            data_NewRec = data_NewRec[data_NewRec['V01301']=='G6820']            
        if data_NewRec.empty == True:
            if sta_type == 'A': print('该时次国家站数据为空：' + time.strftime('%Y-%m-%d %H:%M:%S'))   
            if sta_type == 'B': print('该时次区域站数据为空：' + time.strftime('%Y-%m-%d %H:%M:%S'))   
            continue
    
        data_HourRain = data_NewRec[['ddatetime','v01301','v13019']]
        data_HourRain['ddatetime']=data_HourRain['ddatetime'].astype('datetime64')     
        data_HourRain['v13019']=data_HourRain['v13019'].astype('float')
        
        #删除站号重复的记录，59107和59293
        data_HourRain = data_HourRain.drop_duplicates()
        data_HourRain.to_sql(RainQC_TB,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict

print('完成区域站+国家站小时雨量读取')            





# ======================= 质控多个时次小时雨量数据 =============================
sql_Stacode = 'select v01301 stacode from ' + StaInfo_TB + ' where v_prcode=\'广东\''
Stacodes = pd.read_sql(sql_Stacode,Engine)

sql_Stacode = 'select v01301 stacode,sta_Neigbors from ' + StaNeigbors_TB
Stacodes_Neigbors = pd.read_sql(sql_Stacode,Engine)


print('准备进行 ' + strtime_stt + ' ' + strtime_end + ' 小时雨量质量控制')   

idx = pd.date_range(strtime_stt,strtime_end,freq='1h')
Data_realtime = pd.DataFrame(index=idx)


# ======================================== read data  ==================================          
# -------- read precipitation data ------------
sql = 'select DDATETIME TIME,V01301 STACODE,V13019 R from ' + RainQC_TB
sql = sql + ' where DDATETIME between to_date(\'' + strtime_stt + '\',\'yyyy-mm-dd hh24:mi:ss\')'
sql = sql + ' and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24:mi:ss\')'         
Data_Rain = pd.read_sql(sql,Engine,index_col='time') #SQLAchemy return columns name in lowercase letters


# =============================================================================
# # ============================ 写入DF后，就删除该时段数据库记录 ==========================     
# sql = 'delete from ' + RainQC_TB
# sql += ' where DDATETIME between to_date(\'' + strtime_stt + '\',\'yyyy-mm-dd hh24:mi:ss\')'
# sql += ' and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24:mi:ss\')' 
# Engine.execute(sql)
# =============================================================================


print('构建全省：时间/站点二维数据')
for idx_sta in Stacodes.index:
    stacode = Stacodes.loc[idx_sta,'stacode']
    Data_sta = Data_Rain[Data_Rain['stacode']==stacode]
    
    if Data_sta.empty == False:
        Data_sta = Data_sta.reindex(idx)
        Data_realtime[stacode] = Data_sta['r']

if Data_realtime.empty == False:
    QC_realtime = Data_realtime.copy()
    # ----- 默认所有记录为正确 ------
    QC_realtime.loc[:,:] = 0                
#                QC_realtime[~QC_realtime.isnull()] = 0
    
    # ----- 1.缺测检查（所有9999或999999或null记录为缺测 QC=9） ------
    print('正在进行：缺测值检查') 
    QC_realtime[Data_realtime >= 9999] = 9
    QC_realtime[Data_realtime.isnull()] = 9
    
    # ----- 2.设备界限值检查（错误性判断QC=2） ------
    print('正在进行：设备界限值检查') 
    QC_realtime[(QC_realtime!=9) & ((Data_realtime < 0) | (Data_realtime > 999.9))] = 2
    
#        # ----- 3.气候界限值检查（错误性判断QC=3） ------
#        print('正在进行：气候界限值检查') 
#        QC_realtime[(QC_realtime!=9) & (QC_realtime!=2) & (Data_realtime > Rain_limit)] = 3
   
    # ===== 4.时间一致性检查（错误性判断QC=4） ==================
    print('正在进行：时间一致性检查') 
    # ---- create the monthly data of 1 hour before ------
    Data_tmp = Data_realtime.copy()
    Data_tmp.index += datetime.timedelta(hours=1)
    Data_realtime1 = Data_tmp.reindex(idx)                
    Data_Diff_realtime1 = Data_realtime - Data_realtime1
    
    # ---- create the monthly data of 2 hours before ------
    Data_tmp = Data_realtime.copy()
    Data_tmp.index += datetime.timedelta(hours=2)
    Data_realtime2 = Data_tmp.reindex(idx)                
    Data_Diff_realtime2 = Data_realtime - Data_realtime2
    
    # ---- create the monthly data of 3 hours before ------
    Data_tmp = Data_realtime.copy()
    Data_tmp.index += datetime.timedelta(hours=3)
    Data_realtime3 = Data_tmp.reindex(idx)                
    Data_Diff_realtime3 = Data_realtime - Data_realtime3
    
    # ----- 连续4小时雨量不变：QC4_realtime 'TRUE' values indicate the time and stacode with 4 hours same hourly rainfall ------- 
    QC4_realtime = (QC_realtime!=9) & (Data_realtime <= 10) & (Data_realtime > 1) & \
                 (Data_Diff_realtime1 == 0) & (Data_Diff_realtime2 == 0) & (Data_Diff_realtime3 == 0) 
    QC4_realtime[QC4_realtime == False] = np.nan
    QC4_realtime = QC4_realtime.dropna(how = 'all')
    if QC4_realtime.empty == False:
        for col in QC4_realtime.columns:
            QC4_tmp = QC4_realtime[col]
            QC4_tmp = QC4_tmp.dropna()
            for idx_QC4 in QC4_tmp.index:
                idx_tmp = idx_QC4 + datetime.timedelta(hours= 0); QC_realtime.loc[idx_tmp,col] = 4
                idx_tmp = idx_QC4 + datetime.timedelta(hours=-1); QC_realtime.loc[idx_tmp,col] = 4
                idx_tmp = idx_QC4 + datetime.timedelta(hours=-2); QC_realtime.loc[idx_tmp,col] = 4
                idx_tmp = idx_QC4 + datetime.timedelta(hours=-3); QC_realtime.loc[idx_tmp,col] = 4

    # ----- 连续4小时变化<=0.1：QC4_realtime 'TRUE' values indicate the time and stacode with 4 hours same hourly rainfall ------- 
    QC4_realtime = (QC_realtime!=9) & (Data_realtime > 10) & \
                 (Data_Diff_realtime1.abs() < 0.2) & (Data_Diff_realtime2.abs() < 0.2) & (Data_Diff_realtime3.abs() < 0.2)
#                             ((Data_Diff_realtime1 > -0.2) & (Data_Diff_realtime1 < 0.2)) & \
#                             ((Data_Diff_realtime2 > -0.2) & (Data_Diff_realtime2 < 0.2)) & \
#                             ((Data_Diff_realtime3 > -0.2) & (Data_Diff_realtime3 < 0.2)) 
    QC4_realtime[QC4_realtime == False] = np.nan
    QC4_realtime = QC4_realtime.dropna(how = 'all')
    if QC4_realtime.empty == False:
        for col in QC4_realtime.columns:
            QC4_tmp = QC4_realtime[col]
            QC4_tmp = QC4_tmp.dropna()
            for idx_QC4 in QC4_tmp.index:
                idx_tmp = idx_QC4 + datetime.timedelta(hours= 0); QC_realtime.loc[idx_tmp,col] = 4
                idx_tmp = idx_QC4 + datetime.timedelta(hours=-1); QC_realtime.loc[idx_tmp,col] = 4
                idx_tmp = idx_QC4 + datetime.timedelta(hours=-2); QC_realtime.loc[idx_tmp,col] = 4
                idx_tmp = idx_QC4 + datetime.timedelta(hours=-3); QC_realtime.loc[idx_tmp,col] = 4
            
    print('正在进行：空间一致性检查')                          
    # ===== 5.空间一致性检查（QC=5或QC=1） ==================
    # ------ 将数据前后各增加1个时次 --------
    time_stt2 = time_stt + datetime.timedelta(hours=-1)
    time_end2 = time_end + datetime.timedelta(hours=1)
    strtime_end2 = time_end2.strftime('%Y-%m-%d %H:%M:%S') 
    strtime_stt2 = time_stt2.strftime('%Y-%m-%d %H:%M:%S')            
    idx2 = pd.date_range(strtime_stt2,strtime_end2,freq='1h')
    Data_realtime = Data_realtime.reindex(idx2)
    QC_realtime = QC_realtime.reindex(idx2)

    sta_Province = list(Data_realtime.columns.values)
    
    for col in QC_realtime.columns:

        # ----- 逐列检查所有 QC!=9且R> 30mm 的降水记录 -------
#                QC_R30 = QC_realtime.loc[(QC_realtime[col]!=2) & (QC_realtime[col]!=3) & (QC_realtime[col]!=4) & (QC_realtime[col]!=9) & (Data_realtime[col] >30)]
        QC_R30 = QC_realtime.loc[(QC_realtime[col]!=9) & (Data_realtime[col] >30)]
        
        if QC_R30.empty == False:
            # -------- 创建无该站点数据的本地区月数据 ----------
            tmp = Stacodes_Neigbors[Stacodes_Neigbors['stacode'] == col]['sta_neigbors'].values 
            strtmp = tmp[0]
            # 若存在邻近站点则继续判断，否则不进行空间一致性检查
            if strtmp is not None:
                sta_neigbors = strtmp.split(',')
                
                # 判断邻近站与全省站的交集，避免错误 KeyError: "['G3410', 'G9525'] not in index"
                sta_tmp = [sta for sta in sta_neigbors if sta in sta_Province]                     
                Data_realtime_tmp = Data_realtime[sta_tmp]

                
                # -------- 逐行检查 ------------
                idx_QC_R30 = QC_R30.index
                for idx_QC in idx_QC_R30:
                    
#                            # ------- 默认设置 30mm以上降水为空间错误QC=5 -------
                    QC_realtime.loc[idx_QC,col] = 5                           
                    # -------- 待检查的雨量记录 --------
                    Rain_tmp = Data_realtime.loc[idx_QC,col]
                    
                    # ------ 该雨量的0.5倍量级以上降水（正确性判断） ----------
                    Rain_range_min = Rain_tmp * 0.5
                    Rain_range_max = Rain_limit                            
                    count_tmp = 0
                    # -------- 前1时次达到该降水量级的站点数 -----------
                    idx_tmp = idx_QC + datetime.timedelta(hours=-1)
                    count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                    # -------- 当前时次达到该降水量级的站点数 -----------
                    idx_tmp = idx_QC + datetime.timedelta(hours=0)              
                    count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                    # -------- 后1时次达到该降水量级的站点数 -----------
                    idx_tmp = idx_QC + datetime.timedelta(hours=1)
                    count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                    # --------- 若有其他2个以上同量级降水站点，则该记录判定为正确 ------------
                    if count_tmp >=2:
                        QC_realtime.loc[idx_QC,col] = 0; continue #break
                    else:

                        # ------ 该雨量的0.4倍量级以上降水（正确性判断） ----------
                        Rain_range_min = Rain_tmp * 0.4
                        Rain_range_max = Rain_limit                            
                        count_tmp = 0
                        idx_tmp = idx_QC + datetime.timedelta(hours=-1)
                        count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                        idx_tmp = idx_QC + datetime.timedelta(hours=0)              
                        count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                        idx_tmp = idx_QC + datetime.timedelta(hours=1)
                        count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                        # --------- 若有其他3个以上同量级降水站点，则该记录判定为正确 ------------
                        if count_tmp >=3:
                            QC_realtime.loc[idx_QC,col] = 0; continue #break
                        else:
                            
                            # ------ 该雨量的0.3倍量级以上降水（正确性判断） ----------
                            Rain_range_min = Rain_tmp * 0.3
                            Rain_range_max = Rain_limit                            
                            count_tmp = 0
                            idx_tmp = idx_QC + datetime.timedelta(hours=-1)
                            count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                            idx_tmp = idx_QC + datetime.timedelta(hours=0)              
                            count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                            idx_tmp = idx_QC + datetime.timedelta(hours=1)
                            count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                            # --------- 若有其他4个以上同量级降水站点，则该记录判定为正确 ------------
                            if count_tmp >=4:
                                QC_realtime.loc[idx_QC,col] = 0; continue #break
                            else:
                                
                                # ------ 该雨量的0.2倍量级以上降水（正确性判断） ----------
                                Rain_range_min = Rain_tmp * 0.2
                                Rain_range_max = Rain_limit                            
                                count_tmp = 0
                                idx_tmp = idx_QC + datetime.timedelta(hours=-1)
                                count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                                idx_tmp = idx_QC + datetime.timedelta(hours=0)              
                                count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                                idx_tmp = idx_QC + datetime.timedelta(hours=1)
                                count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                                # --------- 若有其他5个以上同量级降水站点，则该记录判定为正确 ------------
                                if count_tmp >=5:
                                    QC_realtime.loc[idx_QC,col] = 0; continue #break
                                else:  
                                                                               
                                    # ------ 该雨量的0.1倍量级以上降水（正确性 或 可疑性判断）----------
                                    Rain_range_min = Rain_tmp * 0.1
                                    Rain_range_max = Rain_limit
                                    count_tmp = 0
                                    idx_tmp = idx_QC + datetime.timedelta(hours=-1)
                                    count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                                    idx_tmp = idx_QC + datetime.timedelta(hours=0)              
                                    count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                                    idx_tmp = idx_QC + datetime.timedelta(hours=1)
                                    count_tmp += Data_realtime_tmp.loc[idx_tmp][Data_realtime_tmp.loc[idx_tmp].between(Rain_range_min,Rain_range_max)].count()
                                    # --------- 若有其他10个以上同量级降水站点，则该记录判定为正确 ------------
                                    if count_tmp >=10:
                                        QC_realtime.loc[idx_QC,col] = 0 # continue #break
                                    elif (count_tmp >=5) & (count_tmp<10):
                                        QC_realtime.loc[idx_QC,col] = 1

    QC_realtime_DB = pd.DataFrame(columns=['ddatetime','q13019','v01301'])    
    
    for col in QC_realtime.columns:              
        QC_tmp = pd.DataFrame({'ddatetime':QC_realtime.index,'q13019':QC_realtime[col].values})
        QC_tmp = QC_tmp.dropna()
        QC_tmp['v01301'] = col
        QC_realtime_DB = pd.concat([QC_realtime_DB,QC_tmp])
                            # write the DataFrame data into Oracle database
    
    Data_Rain['ddatetime'] = Data_Rain.index
    Data_Rain = Data_Rain.rename(columns={'stacode':'v01301','r':'v13019'})
    RainQC_DB = pd.merge(QC_realtime_DB, Data_Rain, on =['v01301','ddatetime'], how='left')
    
   
    RainQC_DB = RainQC_DB[(RainQC_DB['q13019'] != 0) & (RainQC_DB['q13019'] != 9)]
    RainQC_DB['D_IYMDHM'] = datetime.datetime.now()
    
    # 删除该时段数据库记录     
    sql = 'delete from ' + RainQC_TB
    sql += ' where DDATETIME between to_date(\'' + strtime_stt + '\',\'yyyy-mm-dd hh24:mi:ss\')'
    sql += ' and to_date(\'' + strtime_end +'\',\'yyyy-mm-dd hh24:mi:ss\')' 
    Engine.execute(sql)
    
    RainQC_DB.to_sql(RainQC_TB,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict
        
print('完成质量控制')
end_now= datetime.datetime.now()
print('耗时: ' + str(end_now-stt_now))