# -*- coding: utf-8 -*-
"""
Created on Tue Aug 11 20:43:04 2026

@author: Administrator
"""

import configparser
import datetime
import pandas as pd
import numpy as np

import sys

from pathlib import Path
script_dir = Path(__file__).resolve().parent
project_dir = script_dir.parent
sys.path.insert(0, str(project_dir))

from common.db import get_engine
Engine = get_engine('CROSS_RAIN')

stainfo_tb = 'climate.t_othe_station_meta_basic_tab'
input_tb1 = 'climate.surf_cli_mul_day'
input_tb2 = 'climate.awst_cli_mul_day'


def fetch_rhre_r_prov(prov,date_stt,date_end,statype):
    strdate_stt = date_stt.strftime('%Y-%m-%d')
    strdate_end = date_end.strftime('%Y-%m-%d')

    sql = 'select ddate,stacode,R from ' + input_tb1 + ' where R>=0.1'
    sql += ' and ddate between to_date(\'' + strdate_stt + '\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end + '\',\'yyyy-mm-dd\')'
    sql += ' and stacode in (select v01301 from ' + stainfo_tb + ' where V02301 like \'%A%\' and v_prcode=\'' + prov + '\')'
    if statype == 'both':
        sql += ' UNION '
        sql += 'select ddate,stacode,R from ' + input_tb2 + ' where R>=0.1'
        sql += ' and ddate between to_date(\'' + strdate_stt + '\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end + '\',\'yyyy-mm-dd\')'
        sql += ' and stacode in (select v01301 from ' + stainfo_tb + ' where V02301 like \'%B%\' and v_prcode=\'' + prov + '\' and inland=1)'
    df_rhre_r = pd.read_sql(sql, Engine, index_col='ddate')
    
    return df_rhre_r


def fetch_rhre_r_cities(cities,date_stt,date_end,statype):
    strdate_stt = date_stt.strftime('%Y-%m-%d')
    strdate_end = date_end.strftime('%Y-%m-%d')

    sql = 'select ddate,stacode,R from ' + input_tb1 + ' where R>=0.1'
    sql += ' and ddate between to_date(\'' + strdate_stt + '\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end + '\',\'yyyy-mm-dd\')'
    sql += ' and stacode in (select v01301 from ' + stainfo_tb + ' where V02301 like \'%A%\' and v_city in (' + cities + '))'
    if statype == 'both':
        sql += ' UNION '
        sql += 'select ddate,stacode,R from ' + input_tb2 + ' where R>=0.1'
        sql += ' and ddate between to_date(\'' + strdate_stt + '\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end + '\',\'yyyy-mm-dd\')'
        sql += ' and stacode in (select v01301 from ' + stainfo_tb + ' where V02301 like \'%B%\' and v_city in (' + cities + ') and inland=1)'
    df_rhre_r = pd.read_sql(sql, Engine, index_col='ddate')
    
    return df_rhre_r


def fetch_rhre_r_provinces(provinces,date_stt,date_end,statype):
    strdate_stt = date_stt.strftime('%Y-%m-%d')
    strdate_end = date_end.strftime('%Y-%m-%d')

    sql = 'select ddate,stacode,R from ' + input_tb1 + ' where R>=0.1'
    sql += ' and ddate between to_date(\'' + strdate_stt + '\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end + '\',\'yyyy-mm-dd\')'
    sql += ' and stacode in (select v01301 from ' + stainfo_tb + ' where V02301 like \'%A%\' and v_prcode in (' + provinces + '))'
    if statype == 'both':
        sql += ' UNION '
        sql += 'select ddate,stacode,R from ' + input_tb2 + ' where R>=0.1'
        sql += ' and ddate between to_date(\'' + strdate_stt + '\',\'yyyy-mm-dd\') and to_date(\'' + strdate_end + '\',\'yyyy-mm-dd\')'
        sql += ' and stacode in (select v01301 from ' + stainfo_tb + ' where V02301 like \'%B%\' and v_prcode in (' + provinces + ') and inland=1)'
    df_rhre_r = pd.read_sql(sql, Engine, index_col='ddate')

    return df_rhre_r


def fetch_year_hrain_prov(year,prov,stacodes):
    strtime_stt = datetime.datetime(year, 1, 1).strftime('%Y-%m-%d')
    strtime_end = datetime.datetime(year, 12, 31).strftime('%Y-%m-%d')
    idx_FullYear = pd.date_range(strtime_stt, strtime_end, freq='1d')
    df_year_hrain = pd.DataFrame(index=idx_FullYear)

    sql = 'select ddate,stacode,R from ' + input_tb1 + ' where to_char(ddate,\'yyyy\')=\'' + str(year) + '\''
    sql += ' and stacode in (select v01301 from ' + stainfo_tb
    sql += ' where v02301 like \'%A%\' and v_prcode=\'' + prov + '\')'
    sql += ' and R>=50'  # 2022-5-16 只取暴雨站日，先判断过程，后面再重新取过程所有雨量
    sql += ' order by ddate,stacode'
    DF_tmp = pd.read_sql(sql, Engine, index_col='ddate')

    print('构建 ' + prov + '：' + str(year) + '年- 日期/站点二维数据')
    for idx_sta in stacodes.index:
        stacode = idx_sta
        data_sta = DF_tmp[DF_tmp['stacode'] == stacode]
        if data_sta.empty == False:
            data_sta = data_sta.reindex(idx_FullYear)
            df_year_hrain[stacode] = data_sta['r']
        else:
            df_year_hrain[stacode] = np.nan

    df_year_hrain['RS_cnt'] = df_year_hrain.count(axis=1)
    df_year_hrain = df_year_hrain[df_year_hrain['RS_cnt'] != 0].copy()  # 2024-5-17:仅保留有暴雨站的日子

    return df_year_hrain


def fetch_year_hrain_cities(year,cities,Stacodes_SURFAWST,statype):
    strtime_stt = datetime.datetime(year, 1, 1).strftime('%Y-%m-%d')
    strtime_end = datetime.datetime(year, 12, 31).strftime('%Y-%m-%d')
    idx_FullYear = pd.date_range(strtime_stt, strtime_end, freq='1d')
    df_year_hrain = pd.DataFrame(index=idx_FullYear)

    sql = 'select ddate,stacode,R from ' + input_tb1 + ' where to_char(ddate,\'yyyy\')=\'' + str(year) + '\''
    sql += ' and stacode in (select v01301 from ' + stainfo_tb
    sql += ' where v02301 like \'%A%\' and v_city in (' + cities + '))'
    sql += ' and R>=50'  # 2022-5-16 只取暴雨站日，先判断过程，后面再重新取过程所有雨量
    if statype == 'both':
        sql += ' UNION '
        sql += 'select ddate,stacode,R from ' + input_tb2 + ' where to_char(ddate,\'yyyy\')=\'' + str(year) + '\''
        sql += ' and stacode in (select v01301 from ' + stainfo_tb
        sql += ' where v02301 like \'%%\' and v_city in (' + cities + ') and inland=1)'
        sql += ' and R>=50'  # 2022-5-16 只取暴雨站日，先判断过程，后面再重新取过程所有雨量    
    sql += ' order by ddate,stacode'
    DF_tmp = pd.read_sql(sql, Engine, index_col='ddate')

    print('构建 ' + cities + '：' + str(year) + '年- 日期/站点二维数据')
    for idx_sta in Stacodes_SURFAWST.index:
        stacode = idx_sta
        data_sta = DF_tmp[DF_tmp['stacode'] == stacode]
        if data_sta.empty == False:
            data_sta = data_sta.reindex(idx_FullYear)
            df_year_hrain[stacode] = data_sta['r']
        else:
            df_year_hrain[stacode] = np.nan

    df_year_hrain['RS_cnt'] = df_year_hrain.count(axis=1)
    df_year_hrain = df_year_hrain[df_year_hrain['RS_cnt'] != 0].copy()  # 2024-5-17:仅保留有暴雨站的日子
    
    return df_year_hrain


def fetch_year_hrain_provinces(year,provinces,stacodes,statype):
    strtime_stt = datetime.datetime(year, 1, 1).strftime('%Y-%m-%d')
    strtime_end = datetime.datetime(year, 12, 31).strftime('%Y-%m-%d')
    idx_FullYear = pd.date_range(strtime_stt, strtime_end, freq='1d')
    df_year_hrain = pd.DataFrame(index=idx_FullYear)

    sql = 'select ddate,stacode,R from ' + input_tb1 + ' where to_char(ddate,\'yyyy\')=\'' + str(year) + '\''
    sql += ' and stacode in (select v01301 from ' + stainfo_tb
    sql += ' where v02301 like \'%A%\' and v_prcode in (' + provinces + '))'
    sql += ' and R>=50'  # 2022-5-16 只取暴雨站日，先判断过程，后面再重新取过程所有雨量
    if statype == 'both':
        sql += ' UNION '
        sql += 'select ddate,stacode,R from ' + input_tb2 + ' where to_char(ddate,\'yyyy\')=\'' + str(year) + '\''
        sql += ' and stacode in (select v01301 from ' + stainfo_tb
        sql += ' where v02301 like \'%%\' and v_prcode in (' + provinces + ') and inland=1)'
        sql += ' and R>=50'  # 2022-5-16 只取暴雨站日，先判断过程，后面再重新取过程所有雨量
    sql += ' order by ddate,stacode'
    DF_tmp = pd.read_sql(sql, Engine, index_col='ddate')

    print('构建 ' + provinces + '：' + str(year) + '年- 日期/站点二维数据')
    for idx_sta in stacodes.index:
        stacode = idx_sta
        data_sta = DF_tmp[DF_tmp['stacode'] == stacode]
        if data_sta.empty == False:
            data_sta = data_sta.reindex(idx_FullYear)
            df_year_hrain[stacode] = data_sta['r']
        else:
            df_year_hrain[stacode] = np.nan

    df_year_hrain['RS_cnt'] = df_year_hrain.count(axis=1)
    df_year_hrain = df_year_hrain[df_year_hrain['RS_cnt'] != 0].copy()  # 2024-5-17:仅保留有暴雨站的日子

    return df_year_hrain


def fetch_lonlat_stations_prov(prov,statype):
    # 按从大到小排序取国家站86个站点经纬度
    sql_Stacode = 'select v01301 stacode,v06001 lon,v05001 lat, extract(year from stt_date) year_stt from ' + stainfo_tb
    if statype == 'surf': sql_Stacode += ' where (v02301 like \'%A%\')'
    if statype=='both': sql_Stacode += ' where (v02301 like \'%A%\' or v02301 like \'%B%\' and inland=1)'
    sql_Stacode += ' and v_prcode=\'' + prov + '\''
    sql_Stacode += ' order by v06001 desc, v05001 desc'
    Stacodes_SURF = pd.read_sql(sql_Stacode, Engine, index_col='stacode')
    LonLat_Stacodes = Stacodes_SURF.transpose()
    
    return LonLat_Stacodes, Stacodes_SURF


def fetch_lonlat_stations_cities(cities,statype):
    # 按从大到小排序取国家站86个站点经纬度
    sql_Stacode = 'select v01301 stacode,v06001 lon,v05001 lat, extract(year from stt_date) year_stt from ' + stainfo_tb
    if statype=='both': sql_Stacode += ' where (v02301 like \'%A%\' or v02301 like \'%B%\' and inland=1)'
    if statype=='surf': sql_Stacode += ' where (v02301 like \'%A%\')'
    sql_Stacode += ' and v_city in (' + cities + ')'
    sql_Stacode += ' order by v06001 desc, v05001 desc'
    Stacodes_SURFAWST = pd.read_sql(sql_Stacode, Engine, index_col='stacode')
    LonLat_Stacodes = Stacodes_SURFAWST.transpose()
    
    return LonLat_Stacodes, Stacodes_SURFAWST


def fetch_lonlat_stations_provinces(provinces,statype):
    # 按从大到小排序取国家站站点经纬度
    sql_Stacode = 'select v01301 stacode,v06001 lon,v05001 lat, extract(year from stt_date) year_stt from ' + stainfo_tb
    if statype=='both': sql_Stacode += ' where (v02301 like \'%A%\' or v02301 like \'%B%\' and inland=1)'
    if statype=='surf': sql_Stacode += ' where (v02301 like \'%A%\')'
    sql_Stacode += ' and v_prcode in (' + provinces + ')'
    sql_Stacode += ' order by v06001 desc, v05001 desc'
    Stacodes_SURFAWST = pd.read_sql(sql_Stacode, Engine, index_col='stacode')
    LonLat_Stacodes = Stacodes_SURFAWST.transpose()

    return LonLat_Stacodes, Stacodes_SURFAWST


def fetch_rhre_process(tb_name, region_name, year_stt=None, year_end=None):
    """读取指定过程表中某区域(省/区域/流域通用)的暴雨过程记录。

    year_stt/year_end 都不给：不限制年份；
    只给 year_stt：按单年查询（*_level.py 逐年重新定级时用）；
    两个都给：按 [year_stt, year_end] 年份区间查询（*_level.py 算30年气候基准时用）。
    """
    sql = 'select * from ' + tb_name + ' where cregion=\'' + region_name + '\''
    if year_stt is not None and year_end is not None:
        sql += ' and extract(year from date_stt) between ' + str(year_stt) + ' and ' + str(year_end)
    elif year_stt is not None:
        sql += ' and extract(year from date_stt) = ' + str(year_stt)
    sql += ' order by date_stt'

    return pd.read_sql(sql, Engine)


def write_db(data, tb_name, prov, year,flag):
    sql = 'delete from ' + tb_name + ' where cregion=\'' + prov 
    if flag=='process': sql += '\' and to_char(date_stt,\'yyyy\')=\'' + str(year) + '\''
    if flag=='days': sql += '\' and to_char(ddate,\'yyyy\')=\'' + str(year) + '\''
    Engine.execute(sql)
    print('正在写入数据库：' + tb_name)
    data.to_sql(tb_name, Engine, index=False, if_exists='append', chunksize=100)


def fetch_region_members(section):
    """读取 Config_Regions.ini 中指定 section 下每个区域对应的成员列表
    （[Reg_cities] 的成员是城市名，[Reg_provinces] 的成员是省份名）。

    返回 {区域名: 成员列表字符串}，成员列表字符串已经是可以直接拼进
    "v_city in (...)" 或 "v_prcode in (...)" 这类 SQL子句的形式
    （如 "'广州','佛山','肇庆'"），供 fetch_*_cities()/fetch_*_provinces()
    系列函数直接使用。
    """
    config = configparser.ConfigParser()
    config.read(script_dir / 'Config_Regions.ini', encoding='utf-8-sig')

    regions = {}
    for region_name, members_raw in config[section].items():
        member_list = [member.strip() for member in members_raw.split(',')]
        regions[region_name] = ','.join("'" + member + "'" for member in member_list)

    return regions