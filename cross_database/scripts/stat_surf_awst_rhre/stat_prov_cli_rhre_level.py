# -*- coding: utf-8 -*-
"""
Created on Thu Mar 18 14:15:51 2021
用国家站的30年气候背景值
对国家站暴雨过程和国家站+区域站暴雨过程进行定级
@author: Administrator
"""

import sys

from pathlib import Path
script_dir = Path(__file__).resolve().parent
project_dir = script_dir.parent
sys.path.insert(0, str(project_dir))

from data_port import fetch_rhre_process
from data_port import write_db

import datetime

Output_TB1 = 'prov_cli_rhre_process_surf'
Output_TB2 = 'prov_cli_rhre_process_surfawst'

prov = '广东'

# 读取国家站最新30年区域性暴雨过程综合强度指数的最小值和最大值，采用百分位数法进行等级评判
YEAR_STT_30Y = 1991
YEAR_END_30Y = 2020

DF_RS_30y = fetch_rhre_process(Output_TB1, prov, YEAR_STT_30Y, YEAR_END_30Y)

pcnt50 = DF_RS_30y['rs_z'].quantile(0.5)
pcnt80 = DF_RS_30y['rs_z'].quantile(0.8)
pcnt95 = DF_RS_30y['rs_z'].quantile(0.95)

current_year = datetime.datetime.now().year
years = range(current_year, current_year - 1, -1)
years = range(2026, 1950, -1)

for year in years:
    DF_RS_year_surf = fetch_rhre_process(Output_TB1, prov, year)
    DF_RS_year_surfawst = fetch_rhre_process(Output_TB2, prov, year)

    if DF_RS_year_surf.empty == False:
        print('统计 ' + prov + ' 区域性暴雨过程(国家站)综合强度等级：' + str(year))

        DF_RS_year_surf.loc[DF_RS_year_surf['rs_z'] <= pcnt50, 'clevel'] = '一般'
        DF_RS_year_surf.loc[(DF_RS_year_surf['rs_z'] <= pcnt80) & (DF_RS_year_surf['rs_z'] > pcnt50), 'clevel'] = '较强'
        DF_RS_year_surf.loc[(DF_RS_year_surf['rs_z'] <= pcnt95) & (DF_RS_year_surf['rs_z'] > pcnt80), 'clevel'] = '强'
        DF_RS_year_surf.loc[DF_RS_year_surf['rs_z'] > pcnt95, 'clevel'] = '特强'

        DF_RS_year_surf['d_iymdhm'] = datetime.datetime.now()
        write_db(DF_RS_year_surf, Output_TB1, prov, year, 'process')

    if DF_RS_year_surfawst.empty == False:
        print('统计 ' + prov + ' 区域性暴雨过程(国家站+区域站)综合强度等级：' + str(year))

        DF_RS_year_surfawst.loc[DF_RS_year_surfawst['rs_z'] <= pcnt50, 'clevel'] = '一般'
        DF_RS_year_surfawst.loc[(DF_RS_year_surfawst['rs_z'] <= pcnt80) & (DF_RS_year_surfawst['rs_z'] > pcnt50), 'clevel'] = '较强'
        DF_RS_year_surfawst.loc[(DF_RS_year_surfawst['rs_z'] <= pcnt95) & (DF_RS_year_surfawst['rs_z'] > pcnt80), 'clevel'] = '强'
        DF_RS_year_surfawst.loc[DF_RS_year_surfawst['rs_z'] > pcnt95, 'clevel'] = '特强'

        DF_RS_year_surfawst['d_iymdhm'] = datetime.datetime.now()
        write_db(DF_RS_year_surfawst, Output_TB2, prov, year, 'process')

print(datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S') + ': Finish')
