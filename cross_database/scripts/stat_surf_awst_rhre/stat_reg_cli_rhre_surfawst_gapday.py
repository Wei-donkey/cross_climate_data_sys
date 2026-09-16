# -*- coding: utf-8 -*-
"""
Created on Thu Mar 18 14:15:51 2024
考虑两场区域性暴雨之间的间隔日，若出现1站暴雨，则两场连做一场
单站日暴雨强度 = 暴雨总量/暴雨站次（国家站+区域站）；影响范围 = 国家站暴雨站数（国家站）
@author: Wei
"""

import datetime
import pandas as pd

from rhre_algorithms import rhre_assessment
from rhre_algorithms import rhre_recognition
from rhre_algorithms import compute_stanum_threshold

from data_port import fetch_rhre_r_cities
from data_port import fetch_year_hrain_cities
from data_port import fetch_lonlat_stations_cities
from data_port import fetch_region_members
from data_port import write_db


current_year = datetime.datetime.now().year
years = range(current_year, current_year - 1, -1)
years = range(2026, 1950, -1)

OUTPUT_TB1 = 'reg_cli_rhre_process_surfawst'
OUTPUT_TB2 = 'reg_cli_rhre_day_surfawst'
RATIO_THRESHOLD = 0.05  # 若5%的邻近站出现暴雨，则为一次区域性暴雨过程（不足4站时按4站算，见compute_stanum_threshold）
DISTANCE = 100  # 邻近站标准为距离不大于100公里
# 输出暴雨过程的相关统计量:开始日期、结束日期、持续日数、最大过程雨量、最大过程雨量站点、最大日雨量、最大日雨量站点、最大雨量日期
# 暴雨站数、暴雨站号、暴雨总量、暴雨站次、平均单站日站暴雨量、累计影响范围、综合强度指数
COLS_PROCESS = ['date_stt', 'date_end', 'rs_t', 'racc_max', 'raccmax_sta', 'r_max', 'rmax_sta', 'rmax_date',
                'rs_stanum', 'rs_stacodes', 'rs_ttl', 'rs_m', 'rs_i', 'rs_a', 'rs_z']
# 输出每日暴雨的相关统计量:日期、最大日雨量、最大日雨量站、暴雨站数、暴雨站号、总暴雨量、暴雨站次、平均单站日暴雨量、日影响范围、日综合强度指数
COLS_DAYS = ['ddate', 'r_max', 'rmax_sta', 'rs_stanum', 'rs_stacodes', 'rs_ttl', 'rs_m', 'rs_i', 'rs_a', 'rs_z']


def output_result():
    # =========== 3.输出所有区域性暴雨过程的相关指标 ==============
    df_rhres_process['d_iymdhm'] = datetime.datetime.now()
    df_rhres_process.insert(loc=0, column='cregion', value=region_name)
    df_rhres_process['rs_t'] = df_rhres_process['rs_t'].astype(int)
    df_rhres_process['rs_stanum'] = df_rhres_process['rs_stanum'].astype(int)
    df_rhres_process['rs_m'] = df_rhres_process['rs_m'].astype(float)

    write_db(df_rhres_process, OUTPUT_TB1, region_name, year, 'process')

    # =========== 4.输出所有区域性暴雨逐日的相关指标 ==============
    df_rhres_days['d_iymdhm'] = datetime.datetime.now()
    df_rhres_days.insert(loc=0, column='cregion', value=region_name)
    df_rhres_days['rs_stanum'] = df_rhres_days['rs_stanum'].astype(int)
    df_rhres_days['rs_m'] = df_rhres_days['rs_m'].astype(float)

    write_db(df_rhres_days, OUTPUT_TB2, region_name, year, 'days')


regions = fetch_region_members('Reg_cities')  # {区域名: 城市列表(SQL用)}，读取自 Config_Regions.ini

for region_name, cities in regions.items():
    # 区域范围小，国家站太稀疏，识别过程用国家站+区域站合并的站点密度（assessment阶段再按'surf'/'both'区分统计口径）
    lonlat_stations, stacodes = fetch_lonlat_stations_cities(cities, 'both')

    for year in years:
        # 该年份的国家站+区域站有效站点数量
        stacodes_year = stacodes.loc[stacodes['year_stt'] <= year]
        rows, cols = stacodes_year.shape
        stanum_year = rows
        stanum_threshold = compute_stanum_threshold(stanum_year, RATIO_THRESHOLD)  # 每一年区域性暴雨的数量阈值

        df_rhres_process = pd.DataFrame(columns=COLS_PROCESS)
        df_rhres_days = pd.DataFrame(columns=COLS_DAYS)

        df_year_hrain = fetch_year_hrain_cities(year, cities, stacodes, 'both')

        if not df_year_hrain.empty:
            # ========= 1 识别所有RS过程 ==============
            rhres_date, df_year_hrain = rhre_recognition(df_year_hrain, lonlat_stations, region_name, stanum_threshold, DISTANCE)

            if not rhres_date.empty:
                rows, cols = rhres_date.shape

                # ========= 2 统计所有RS过程的相关指标 ==============
                #  循环每一场区域性暴雨，统计该次区域性暴雨过程指标和区域性暴雨每日指标，并将今年的结果合并在一起
                for idx in rhres_date.index:
                    date_stt = rhres_date.loc[idx, 'date_stt']
                    date_end = rhres_date.loc[idx, 'date_end']
                    # 此次过程数据
                    df_single_process_hrain = df_year_hrain[
                        (df_year_hrain['date'] >= date_stt) & (df_year_hrain['date'] <= date_end)]

                    df_rhre_r = fetch_rhre_r_cities(cities, date_stt, date_end, 'both')
                    df_single_process, df_single_days = rhre_assessment(df_single_process_hrain, df_rhre_r, date_stt, date_end, stanum_year)

                    df_rhres_process = pd.concat([df_rhres_process, df_single_process], ignore_index=True)
                    df_rhres_days = pd.concat([df_rhres_days, df_single_days], ignore_index=True)

                output_result()

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')
