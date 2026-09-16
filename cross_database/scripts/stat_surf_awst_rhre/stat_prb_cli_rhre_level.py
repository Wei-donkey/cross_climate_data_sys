# -*- coding: utf-8 -*-
"""
Created on Thu Mar 18 14:15:51 2021
用国家站的30年气候背景值对国家站暴雨过程进行定级。
珠江流域范围内没有区域站数据，只有国家站('surf')一张过程表要处理，
不像省级/区域级那样还要再定级一遍 surfawst 版本。
@author: Wei
"""

import sys

from pathlib import Path
script_dir = Path(__file__).resolve().parent
project_dir = script_dir.parent
sys.path.insert(0, str(project_dir))

from data_port import fetch_region_members
from data_port import fetch_rhre_process
from data_port import write_db

import datetime

OUTPUT_TB = 'prb_cli_rhre_process_surf'

# 读取国家站最新30年区域性暴雨过程综合强度指数的最小值和最大值，采用百分位数法进行等级评判
YEAR_STT_30Y = 1991
YEAR_END_30Y = 2020

current_year = datetime.datetime.now().year
years = range(current_year, current_year - 1, -1)
years = range(2026, 1950, -1)

basins = fetch_region_members('Reg_provinces')  # {流域名: 省份列表(SQL用)}，读取自 Config_Regions.ini 的 [Reg_provinces]

for basin_name in basins:
    df_rs_30y = fetch_rhre_process(OUTPUT_TB, basin_name, YEAR_STT_30Y, YEAR_END_30Y)

    pcnt50 = df_rs_30y['rs_z'].quantile(0.5)
    pcnt80 = df_rs_30y['rs_z'].quantile(0.8)
    pcnt95 = df_rs_30y['rs_z'].quantile(0.95)

    for year in years:
        df_rs_year = fetch_rhre_process(OUTPUT_TB, basin_name, year)

        if df_rs_year.empty:
            continue

        print('统计 ' + basin_name + ' 区域性暴雨过程(国家站)综合强度等级：' + str(year))

        df_rs_year.loc[df_rs_year['rs_z'] <= pcnt50, 'clevel'] = '一般'
        df_rs_year.loc[(df_rs_year['rs_z'] <= pcnt80) & (df_rs_year['rs_z'] > pcnt50), 'clevel'] = '较强'
        df_rs_year.loc[(df_rs_year['rs_z'] <= pcnt95) & (df_rs_year['rs_z'] > pcnt80), 'clevel'] = '强'
        df_rs_year.loc[df_rs_year['rs_z'] > pcnt95, 'clevel'] = '特强'

        df_rs_year['d_iymdhm'] = datetime.datetime.now()
        write_db(df_rs_year, OUTPUT_TB, basin_name, year, 'process')

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')
