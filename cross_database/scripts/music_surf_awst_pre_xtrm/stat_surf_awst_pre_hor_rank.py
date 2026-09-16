# -*- coding: utf-8 -*-
"""
Created on Thu Jun  4 16:11:00 2020
modified on April 8
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

# ============================ read configuration parameters from ini file ==================================
Engine = get_engine('CROSS_RAIN')
# ============================ read configuration parameters from ini file ==================================

StaInfo_TB = 'climate.t_othe_station_meta_basic_tab'
Input_TB = 'surf_awst_cli_pre_hor_extreme'
Output_TB = 'surf_awst_cli_pre_hor_ranking'
Threshold_TB = 'surf_awst_cli_pre_threshold'

Hours_types = ['1','3','6','12','24']
regions = ['prov','city','county','town','sta']
# 破上述各区域下述各前n位极值才记录
rankings = [50,40,30,20,10]
# 各历时r<10mm的记录，均不参与排名统计
r_not_stat = 10


scales = ['MONTH','SEASON','QUARTER','ANNUAL']
scales_cn = ['月','季节','季度','年']
# '龙舟水'一定要放在第一位
scales2 = ['DBR','MONTH','SEASON','QUARTER','ANNUAL']#dbr- dragon-boat-rain
scales2_cn = ['龙舟水','月','季节','季度','年']


# ======================= 设定数据统计的起止时间（默认只统计当前时次，不更新过去的时次） =============================
time_end = datetime.datetime.now()
# time_end = datetime.datetime(2026,7,30,17) #北京时
time_stt = datetime.datetime.now()
# time_stt = datetime.datetime(2026,7,1,17) #北京时

iyear_end = time_end.year; imonth_end = time_end.month; iday_end = time_end.day; ihour_end = time_end.hour
time_end = datetime.datetime(iyear_end,imonth_end,iday_end,ihour_end)# + datetime.timedelta(hours=-8) #国际时
strtime_end = time_end.strftime('%Y-%m-%d %H:%M:%S')

iyear_stt = time_stt.year; imonth_stt = time_stt.month; iday_stt = time_stt.day; ihour_stt = time_stt.hour
time_stt = datetime.datetime(iyear_stt,imonth_stt,iday_stt,ihour_stt)# + datetime.timedelta(hours=-8) #国际时
strtime_stt = time_stt.strftime('%Y-%m-%d %H:%M:%S')

time_range = pd.date_range(strtime_end,strtime_stt,freq='-1h')
# ======================= 设定数据统计的起止时间 =============================


# 各记录在各等级地区（省-市-县-镇）的排名统计
def Stat_Rain_Record_Region(region_lvl,region,region_higher,ddatetime_NewRec,data_NewRec_region):

    if region_lvl == 'prov' : ranking = rankings[0]
    if region_lvl == 'city' : ranking = rankings[1]     
    if region_lvl == 'county' : ranking = rankings[2]
    if region_lvl == 'town' : ranking = rankings[3]       
    if region_lvl == 'sta' : ranking = rankings[4]     
    
    data_ranking_tmp = pd.DataFrame(columns=['hours_type','ddatetime','stacode','r','cperiod','ranking'])

    for Hours_type in Hours_types:
        data_NewRec_A2 = data_NewRec_region[data_NewRec_region['hours_type']==int(Hours_type)]
        if data_NewRec_A2.empty == True:
            continue        

        # 以历时过程的结束时次为判断依据
        time_END_RainPeriod = datetime.datetime.strptime(ddatetime_NewRec,'%Y-%m-%d %H:%M')
        
        # 如果时间处于龙舟水期间，则time_scale = scales2
        if (time_END_RainPeriod.strftime('%m-%d') >= '05-21') and  (time_END_RainPeriod.strftime('%m-%d') <= '06-20'):
            periods = scales2
            periods_cn = scales2_cn
        else:
            periods = scales
            periods_cn = scales_cn

        mm = time_END_RainPeriod.strftime('%m')
        for period in periods:
            if period == 'DBR':
                sql_period = ' and to_char(ddatetime,\'mm-dd\') between \'05-21\' and \'06-20\''
            elif period == 'MONTH':
                sql_period = ' and to_char(ddatetime,\'mm\') = \'' + mm + '\''
            elif period == 'SEASON':
                if mm in ('03','04','05') : sql_period = ' and to_char(ddatetime,\'mm\') in  (\'03\',\'04\',\'05\')'
                if mm in ('06','07','08') : sql_period = ' and to_char(ddatetime,\'mm\') in  (\'06\',\'07\',\'08\')'
                if mm in ('09','10','11') : sql_period = ' and to_char(ddatetime,\'mm\') in  (\'09\',\'10\',\'11\')'
                if mm in ('12','01','02') : sql_period = ' and to_char(ddatetime,\'mm\') in  (\'12\',\'01\',\'02\')'
            elif period == 'QUARTER':
                if mm in ('01','02','03') : sql_period = ' and to_char(ddatetime,\'mm\') in  (\'01\',\'02\',\'03\')'
                if mm in ('04','05','06') : sql_period = ' and to_char(ddatetime,\'mm\') in  (\'04\',\'05\',\'06\')'
                if mm in ('07','08','09') : sql_period = ' and to_char(ddatetime,\'mm\') in  (\'07\',\'08\',\'09\')'
                if mm in ('10','11','12') : sql_period = ' and to_char(ddatetime,\'mm\') in  (\'10\',\'11\',\'12\')'
            elif period == 'ANNUAL':
                sql_period = ''

            # get the first N records from the Rain Extreme database
            sql = 'select * from ('
            sql += 'select STACODE,DDATETIME,R from ' + Input_TB +' a'
            sql += ' left join '+ StaInfo_TB +' b'
            sql += ' on a.stacode=b.v01301'
            sql += ' where HOURS_TYPE=' + Hours_type
            sql += ' and DDATETIME <= to_date(\'' + ddatetime_NewRec +'\',\'yyyy-mm-dd hh24:mi\')' 
            sql += ' and QC_MANUAL=0 and INLAND=1'
            if region_lvl == 'prov' : sql += ' and V_PRCODE=\'' + region + '\''
            if region_lvl == 'sta' : sql += ' and V01301=\'' + region + '\''
            if region_lvl == 'city' : sql += ' and V_CITY=\'' + region + '\' and V_PRCODE=\'' + region_higher + '\''
            if region_lvl == 'county' : sql += ' and V_COUNTY=\'' + region + '\' and V_CITY=\'' + region_higher + '\''
            if region_lvl == 'town' : sql += ' and V_TOWN=\'' + region + '\' and V_COUNTY=\'' + region_higher + '\''
            # if region_lvl == 'sta' : sql += ' and V01301=\'' + region + '\' and V_TOWN=\'' + region_higher + '\''
            sql += sql_period
            sql += ' order by R desc) where rownum<=' + str(ranking)
            
            data_Extreme = pd.read_sql(sql,Engine)
            # rows, cols = data_Extreme.shape
            # if rows == 0:
            if data_Extreme.empty == True:
                continue 
            else:    
                R_threshold = data_Extreme.iloc[-1]['r']
            
            data_NewRec_A3 = data_NewRec_A2[data_NewRec_A2['r'] >= R_threshold]         
            if data_NewRec_A3.empty == True:
                if period == 'DBR': continue 
                if period != 'DBR': break      
            
            for idx in data_NewRec_A3.index:
                
                stacode = data_NewRec_A3.loc[idx,'stacode']
                R_NewRec = data_NewRec_A3.loc[idx,'r']
              
                data_tmp = data_Extreme[(data_Extreme['ddatetime']==ddatetime_NewRec) & (data_Extreme['stacode']==stacode)]
                # 2022-6-9 修复bug:不应该筛选ddatetime_NewRec时，stacode的记录，应该筛选r==R_NewRec的记录（可能有并列多条）
                data_tmp = data_Extreme[data_Extreme['r'] == R_NewRec]
                if data_tmp.empty == True: # 最后几名雨量和本站相等，于是极值集data_Extreme中无该站点
                    # 在data_Extreme中查找比R_NewRec小或相等的记录，则是R_NewRec的排名
                    for idx_extreme in data_Extreme.index:
                        if R_NewRec >= data_Extreme.loc[idx_extreme,'r']:
                            ranking_tmp = idx_extreme + 1
                            break
                else:    
                    ranking_tmp = data_tmp.index[0] + 1
                    
                print(stacode + ',' + Hours_type + '小时累积雨量 ' + str(R_NewRec) + ',破' + region + ' ' + periods_cn[periods.index(period)] + ' 极值第' + str(ranking_tmp) + '名')              
                
                data_ranking_tmp = data_ranking_tmp.append({'hours_type':Hours_type,'ddatetime':ddatetime_NewRec,'stacode':stacode\
                                                            ,'r':R_NewRec,'cperiod':period,'ranking':ranking_tmp},ignore_index=True)

    
    return data_ranking_tmp
                
#def Stat_Rain_Record(time_range):
for time in time_range:
    ddatetime_NewRec = time.strftime('%Y-%m-%d %H:%M')
    
    month = time.month
    sql = 'select hours_type,stacode,imonth,r_ranking10 from ' + Threshold_TB
    sql += ' where imonth =' + str(month)
    
    data_Ranking10 = pd.read_sql(sql,Engine)
   
    # 读取极值表当前时次极值（不包含QC标识错误的、INLAND标识不参与统计的记录）
    print('正在读取极值表中当前时次降水量记录：'  + ddatetime_NewRec)  
    sql = 'select a.HOURS_TYPE,a.STACODE,a.DDATETIME,a.R,b.V_PRCODE,b.V_CITY,b.V_COUNTY,b.V_TOWN from '+ Input_TB +' a'
    sql += ' left join '+ StaInfo_TB +' b'
    sql += ' on a.stacode=b.v01301'
    sql += ' where DDATETIME = to_date(\'' + ddatetime_NewRec +'\',\'yyyy-mm-dd hh24:mi\')'      
    sql += ' and QC_MANUAL=0 and INLAND=1'
    sql += ' order by V_PRCODE,V_CITY,V_COUNTY,V_TOWN'
    data_NewRec = pd.read_sql(sql,Engine) #SQLAchemy return columns name in lowercase letters  
    
    # 去掉 r< r_not_stat 的记录
    data_NewRec = data_NewRec[data_NewRec['r'] >= r_not_stat]

    # print('正在过滤小于各历时、各站点第10名的降水量记录')  
    for idx in data_NewRec.index:
        Hours_type = data_NewRec.loc[idx,'hours_type']
        stacode = data_NewRec.loc[idx,'stacode']
        r_Ranking10 = data_Ranking10[(data_Ranking10['hours_type']==Hours_type) & (data_Ranking10['stacode']==stacode)]
        if r_Ranking10.empty == False:
            r_Ranking10 = r_Ranking10['r_ranking10'].values[0]
            r_NewRec = data_NewRec.loc[idx,'r']
            # 如果该新记录未进入本站前10名，则不进行排名统计
            if r_NewRec < r_Ranking10:
                data_NewRec = data_NewRec.drop(idx)

    if data_NewRec.empty == True:              
        print('当前时次无破极值记录：' + ddatetime_NewRec)  
        continue
    
    # 全省一次性输出一次排名数据
    data_ranking_prov = pd.DataFrame(columns=['hours_type','ddatetime','stacode','r','cperiod','ranking_prov'])
    data_ranking_city = pd.DataFrame(columns=['hours_type','ddatetime','stacode','r','cperiod','ranking_city'])
    data_ranking_county = pd.DataFrame(columns=['hours_type','ddatetime','stacode','r','cperiod','ranking_county'])
    data_ranking_town = pd.DataFrame(columns=['hours_type','ddatetime','stacode','r','cperiod','ranking_town'])
    data_ranking_sta = pd.DataFrame(columns=['hours_type','ddatetime','stacode','r','cperiod','ranking_sta'])


    print('正在统计各站点该时次雨量破极值情况：' + ddatetime_NewRec)        
# ============================ 破省级记录统计 ===============================
    prov_list = ['广东']
    
    for prov in prov_list:
        if prov is None:
            continue 
        region_higher = ''
        data_NewRec_prov = data_NewRec[data_NewRec['v_prcode'] == prov]          
        print('正在统计破 ' + prov +' 极值情况：' + ddatetime_NewRec)  
        data_ranking_tmp = Stat_Rain_Record_Region('prov',prov,region_higher,ddatetime_NewRec,data_NewRec_prov)
        if data_ranking_tmp.empty == False:         
            data_ranking_tmp.rename(columns = {'ranking':'ranking_prov'},inplace=True)
            data_ranking_prov = pd.concat([data_ranking_prov,data_ranking_tmp])


# ============================ 市级排名统计 ===============================
        city_list = list(set(data_NewRec_prov['v_city'].values))
        for city in city_list:
            if city is None:
                continue
            region_higher = prov
            data_NewRec_city = data_NewRec_prov[data_NewRec_prov['v_city'] == city]
            print('正在统计破 ' + city +' 极值情况：' + ddatetime_NewRec)
            data_ranking_tmp = Stat_Rain_Record_Region('city',city,region_higher,ddatetime_NewRec,data_NewRec_city)
            if data_ranking_tmp.empty == False:
                data_ranking_tmp.rename(columns = {'ranking':'ranking_city'},inplace=True)
                data_ranking_city = pd.concat([data_ranking_city,data_ranking_tmp])

# ============================ 县级排名统计 ===============================
            county_list = list(set(data_NewRec_city['v_county'].values))
            for county in county_list:
                if county is None:
                    continue
                region_higher = city
                data_NewRec_county = data_NewRec_city[data_NewRec_city['v_county'] == county]
                print('正在统计破 ' + county +' 极值情况：' + ddatetime_NewRec)
                data_ranking_tmp = Stat_Rain_Record_Region('county',county,region_higher,ddatetime_NewRec,data_NewRec_county)
                if data_ranking_tmp.empty == False:
                    data_ranking_tmp.rename(columns = {'ranking':'ranking_county'},inplace=True)
                    data_ranking_county = pd.concat([data_ranking_county,data_ranking_tmp])

# ============================ 破镇级记录统计 ===============================
                town_list = list(set(data_NewRec_county['v_town'].values))
                for town in town_list:
                    if town is None:
                        continue
                    region_higher = county
                    data_NewRec_town = data_NewRec_county[data_NewRec_county['v_town'] == town]
                    print('正在统计破 ' + town +' 极值情况：' + ddatetime_NewRec)
                    data_ranking_tmp = Stat_Rain_Record_Region('town',town,region_higher,ddatetime_NewRec,data_NewRec_town)
                    if data_ranking_tmp.empty == False:
                        data_ranking_tmp.rename(columns = {'ranking':'ranking_town'},inplace=True)
                        data_ranking_town = pd.concat([data_ranking_town,data_ranking_tmp])


# ============================ 破站级记录统计 ===============================
        sta_list = list(set(data_NewRec_prov['stacode'].values))
        for sta in sta_list:
            if sta is None:
                continue                
            region_higher = ''
            data_NewRec_sta = data_NewRec_prov[data_NewRec_prov['stacode'] == sta]
            print('正在统计破 ' + sta +' 极值情况：' + ddatetime_NewRec)  
            data_ranking_tmp = Stat_Rain_Record_Region('sta',sta,region_higher,ddatetime_NewRec,data_NewRec_sta)
            data_ranking_tmp.rename(columns = {'ranking':'ranking_sta'},inplace=True)
            data_ranking_sta = pd.concat([data_ranking_sta,data_ranking_tmp])
        
                    
    data_ranking = data_ranking_prov
    data_ranking = pd.merge(data_ranking,data_ranking_city,on = ['hours_type','ddatetime','stacode','r','cperiod'],how='outer')
    data_ranking = pd.merge(data_ranking,data_ranking_county,on = ['hours_type','ddatetime','stacode','r','cperiod'],how='outer')
    data_ranking = pd.merge(data_ranking,data_ranking_town,on = ['hours_type','ddatetime','stacode','r','cperiod'],how='outer')
    data_ranking = pd.merge(data_ranking,data_ranking_sta,on = ['hours_type','ddatetime','stacode','r','cperiod'],how='outer')
    
    # data_Ranking = data_Ranking.dropna(axis=0,subset=['ranking_sta'])

    data_ranking['hours_type'] = data_ranking['hours_type'].astype('int')
    data_ranking['ddatetime'] = pd.to_datetime(data_ranking['ddatetime'])

    data_ranking = data_ranking.fillna(-9999)
    data_ranking['ranking_prov'] = data_ranking['ranking_prov'].astype('int')
    data_ranking['ranking_city'] = data_ranking['ranking_city'].astype('int')
    data_ranking['ranking_county'] = data_ranking['ranking_county'].astype('int')
    data_ranking['ranking_town'] = data_ranking['ranking_town'].astype('int')
    data_ranking.replace(-9999, np.nan, inplace=True)

    # delete existed data first
    sql = 'delete from ' + Output_TB
    sql += ' where DDATETIME = to_date(\'' + ddatetime_NewRec +'\',\'yyyy-mm-dd hh24:mi\')'  
    Engine.execute(sql)        

    data_ranking['D_IYMDHM'] = datetime.datetime.now()

    print('正在写入数据库：' + Output_TB)
    data_ranking.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)#,dtype=dtypedict

print('完成排名统计')
end_now= datetime.datetime.now()
print('耗时: ' + str(end_now-stt_now))