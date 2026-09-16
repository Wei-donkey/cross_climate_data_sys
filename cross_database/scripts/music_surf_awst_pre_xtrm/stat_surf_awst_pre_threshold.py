# -*- coding: utf-8 -*-
"""
Created on Thu Aug  6 09:09:30 2020

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

# ============================ read configuration parameters from ini file ==================================
Engine = get_engine('CROSS_RAIN')
# ============================ read configuration parameters from ini file ==================================

StaInfo_TB = 'climate.t_othe_station_meta_basic_tab'
Input_TB = 'surf_awst_cli_pre_hor_extreme'
Output_TB = 'surf_awst_cli_pre_threshold'

Hour_types = ['1','3','6','12','24']

ranking = 10

# 只更新当前月份阈值
month = datetime.datetime.now().month
# month = 3

sql = 'select V01301,INLAND from '+ StaInfo_TB + ' where INLAND=1'
sql += ' and v_prcode=\'广东\' and v02301 not like \'%O%\''
sql += ' order by V01301'
#StaInfo = pd.read_sql(sql,Engine) #SQLAchemy return columns name in lowercase letters  
result = Engine.execute(sql)
StaInfo = result.fetchall()

r_min_rows = []
r_ranking10_rows = []

# 服务端用窗口函数一次性算出该站每种历时的最小值(min)和第ranking名，
# 避免每站每历时各查一次数据库（原来是 每站 x 5种历时 x 2 = 10次查询）
sql_template = (
    "select hours_type, r, ddatetime, rn_asc, rn_desc, cnt from ("
    "select HOURS_TYPE hours_type, R r, ddatetime,"
    " row_number() over (partition by HOURS_TYPE order by R asc) rn_asc,"
    " row_number() over (partition by HOURS_TYPE order by R desc) rn_desc,"
    " count(*) over (partition by HOURS_TYPE) cnt"
    " from " + Input_TB + " a"
    " where HOURS_TYPE in (" + ",".join(Hour_types) + ")"
    " and QC_MANUAL=0"
    " and to_char(ddatetime,'mm') = '" + str(month).zfill(2) + "'"
    " and STACODE='{stacode}'"
    ") where rn_asc = 1 or rn_desc = least(cnt," + str(ranking) + ")"
)

for sta_info in StaInfo:
    stacode = sta_info[0]
    print('查询当前站各历时降水量极值（最后1名/第 ' + str(ranking) + ' 名）：' + stacode)
    sql = sql_template.format(stacode=stacode)
    df_sta = pd.read_sql(sql, Engine)
    if df_sta.empty:
        continue

    for _, row in df_sta.iterrows():
        if row['rn_asc'] == 1:
            r_min_rows.append({
                'hours_type': row['hours_type'], 'stacode': stacode, 'imonth': month,
                'r_min': row['r'], 'time_min': row['ddatetime'],
            })
        if row['rn_desc'] == min(row['cnt'], ranking):
            r_ranking10_rows.append({
                'hours_type': row['hours_type'], 'stacode': stacode, 'imonth': month,
                'r_ranking10': row['r'], 'time_ranking10': row['ddatetime'],
            })

r_min = pd.DataFrame(r_min_rows, columns=['hours_type','stacode','imonth','r_min','time_min'])
r_ranking10 = pd.DataFrame(r_ranking10_rows, columns=['hours_type','stacode','imonth','r_ranking10','time_ranking10'])

r_threshold = pd.merge(r_ranking10, r_min, on =['hours_type','stacode','imonth'],how='outer')
r_threshold['imonth']=r_threshold['imonth'].astype('int')
r_threshold['hours_type']=r_threshold['hours_type'].astype('int')
r_threshold['d_iymdhm'] = datetime.datetime.now()


# delete existed data first
sql = 'delete from ' + Output_TB
sql += ' where imonth = ' + str(month)
Engine.execute(sql)

print('正在写入数据库：' + Output_TB)
r_threshold.to_sql(Output_TB,Engine,index=False,if_exists='append',chunksize=100)

print('完成雨量阈值更新')