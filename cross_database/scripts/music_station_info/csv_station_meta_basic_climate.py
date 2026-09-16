# -*- coding: utf-8 -*-
"""
Created on Mon Jul 27 10:32:27 2020

@author: qhzx
"""

from datetime import datetime
import sys

import pandas as pd
from pathlib import Path
script_dir = Path(__file__).resolve().parent
project_dir = script_dir.parent
sys.path.insert(0, str(project_dir))

from common.db import get_engine

# ============================ read configuration parameters from ini file ==================================
Engine1 = get_engine('CROSS_CLIMATE')
# ============================ read config parameters from ini file ==================================

StaInfo_TB = 't_othe_station_meta_basic_tab'

import urllib.parse
prov='广东'
prov_code=urllib.parse.quote(prov,encoding='gb2312')

print('正在读取csv文件：站点信息')

StaInfo = pd.read_csv(script_dir / 'stainfo_bk.csv')

StaInfo['v06001']=StaInfo['v06001'].astype('float'); StaInfo['v05001']=StaInfo['v05001'].astype('float')
StaInfo['v07001']=StaInfo['v07001'].astype('float'); StaInfo['v08001']=StaInfo['v08001'].astype('float')
StaInfo['v02001']=StaInfo['v02001'].astype('float'); StaInfo['v01300']=StaInfo['v01300'].astype('float')
StaInfo['v01300_1']=StaInfo['v01300_1'].astype('float')

StaInfo['vertime']=StaInfo['vertime'].astype('datetime64'); StaInfo['d_update_time']=StaInfo['d_update_time'].astype('datetime64')
StaInfo['inland']=StaInfo['inland'].astype('int')
StaInfo['stt_date']=StaInfo['stt_date'].astype('datetime64')

sql = 'delete from ' + StaInfo_TB + ' where v_prcode=\'' +prov+ '\''
Engine1.execute(sql)

print('正在写入CLIMATE数据库：' + StaInfo_TB)
# write the DataFrame data into Oracle database
StaInfo.to_sql(StaInfo_TB,Engine1,index=False,if_exists='append',chunksize=1000)
print('完成 CLIMATE库 站点信息更新')