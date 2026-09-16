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
from common.config import load_ini

# ============================ read configuration parameters from ini file ==================================
Engine1 = get_engine('CROSS_CLIMATE')

config = load_ini('Config_MUSIC_GD.ini')
url_music = config['CONNECT_qhztj']['url']
user_music = config['CONNECT_qhztj']['user']
pwd_music = config['CONNECT_qhztj']['password']
interface_music = config['INTERFACE']['STATION_META_BASIC']
columns_music = config['COLUMNS_IN']['STATION_META_BASIC']
# ============================ read config parameters from ini file ==================================

StaInfo_TB = 't_othe_station_meta_basic_tab'
DataAWST_TB = 'awst_cli_mul_day'
DataSURF_TB = 'surf_cli_mul_day'
import urllib.parse
prov='广东'
prov_code=urllib.parse.quote(prov,encoding='gb2312')

print('正在读取music接口：区域站站点信息')
baseUrl = url_music
baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
baseUrl += '&interfaceId=' + interface_music
baseUrl += '&cols=' + columns_music
baseUrl += '&prov=' + prov_code
baseUrl += '&value=B'
baseUrl += '&dataFormat=html'

# reading the aws station info of Guangdong province from music
tmp = pd.read_html(baseUrl,encoding='utf-8')
data_tmp = tmp[0]
StaInfo_aws = data_tmp.iloc[1:]
col_names = data_tmp.iloc[0]
col_names = map(str.lower, col_names) #[col_name.lower() for col_name in col_names]
StaInfo_aws.columns = col_names

print('正在读取music接口：国家站站点信息')
baseUrl = url_music
baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
baseUrl += '&interfaceId=' + interface_music
baseUrl += '&cols=' + columns_music
baseUrl += '&prov=' + prov_code
baseUrl += '&value=A'
baseUrl += '&dataFormat=html'

# reading the aws station info of Guangdong province from music
tmp = pd.read_html(baseUrl,encoding='utf-8')
data_tmp = tmp[0]
StaInfo_surf = data_tmp.iloc[1:]
col_names = data_tmp.iloc[0]
col_names = map(str.lower, col_names) #[col_name.lower() for col_name in col_names]
StaInfo_surf.columns = col_names

StaInfo_New= pd.concat([StaInfo_aws,StaInfo_surf],ignore_index=True)

print('正在读取music接口：水文站站点信息')
baseUrl = url_music
baseUrl += '&userId=' + user_music + '&pwd=' + pwd_music
baseUrl += '&interfaceId=' + interface_music
baseUrl += '&cols=' + columns_music
baseUrl += '&prov=' + prov_code
baseUrl += '&value=O'
baseUrl += '&dataFormat=html'

# reading the aws station info of Guangdong province from music
tmp = pd.read_html(baseUrl,encoding='utf-8')
data_tmp = tmp[0]
StaInfo_hydro = data_tmp.iloc[1:]
col_names = data_tmp.iloc[0]
col_names = map(str.lower, col_names) #[col_name.lower() for col_name in col_names]
StaInfo_hydro.columns = col_names

StaInfo_New= pd.concat([StaInfo_New,StaInfo_hydro],ignore_index=True)

sql = 'select V01301,INLAND, STT_DATE from '+ StaInfo_TB + ' order by V01301'  #,GBA
StaInfo_Old = pd.read_sql(sql,Engine1) #SQLAchemy return columns name in lowercase letters  

StaInfo = pd.merge(StaInfo_New, StaInfo_Old, on ='v01301', how='left')

#StaInfo.loc[pd.isnull(StaInfo['d_iymdhm']),'d_iymdhm'] = datetime.datetime.now()
StaInfo.loc[pd.isnull(StaInfo['inland']),'inland']=1
StaInfo.loc[pd.isnull(StaInfo['stt_date']),'stt_date'] = datetime(1899,9,9)

# ======================= 判断资料的开始日期，写入站点信息表 ===========================
for idx in StaInfo.index:
    stacode = StaInfo.loc[idx,'v01301']
    statype = StaInfo.loc[idx,'v02301']
    if 'B' in statype: sql = 'select stacode, min(ddate) stt_date from ' + DataAWST_TB
    if 'A' in statype: sql = 'select stacode, min(ddate) stt_date from ' + DataSURF_TB
    if 'O' in statype: sql = 'select stacode, min(ddate) stt_date from ' + DataAWST_TB
    sql += ' where stacode=\'' + stacode + '\' group by stacode'
    stt_date = pd.read_sql(sql,Engine1)
    if stt_date.empty == False:
        if 'B' in statype:
            tmp_stt_date = pd.to_datetime(stt_date['stt_date'].values[0])
            print(stacode + ' 站点资料起始时间：' + tmp_stt_date.strftime('%Y-%m-%d'))
            StaInfo.loc[StaInfo['v01301']==stacode,'stt_date'] = tmp_stt_date
        if 'A' in statype:
            iyear = stt_date['stt_date'].dt.year[0]
            if stt_date['stt_date'].dt.strftime('%m-%d')[0] != '01-01':
                tmp_stt_date = datetime(iyear+1,1,1)
            else:
                tmp_stt_date = datetime(iyear,1,1)
            print(stacode + ' 站点资料起始时间：' + tmp_stt_date.strftime('%Y-%m-%d'))
            StaInfo.loc[StaInfo['v01301']==stacode,'stt_date'] = tmp_stt_date
        if 'O' in statype:
            tmp_stt_date = pd.to_datetime(stt_date['stt_date'].values[0])
            print(stacode + ' 站点资料起始时间：' + tmp_stt_date.strftime('%Y-%m-%d'))
            StaInfo.loc[StaInfo['v01301']==stacode,'stt_date'] = tmp_stt_date
# ======================= 判断资料的开始日期，写入站点信息表 ===========================    

StaInfo['v06001']=StaInfo['v06001'].astype('float'); StaInfo['v05001']=StaInfo['v05001'].astype('float')
StaInfo['v07001']=StaInfo['v07001'].astype('float'); StaInfo['v08001']=StaInfo['v08001'].astype('float')
StaInfo['v02001']=StaInfo['v02001'].astype('float'); StaInfo['v01300']=StaInfo['v01300'].astype('float')
StaInfo['v01300_1']=StaInfo['v01300_1'].astype('float')

StaInfo['vertime']=StaInfo['vertime'].astype('datetime64'); StaInfo['d_update_time']=StaInfo['d_update_time'].astype('datetime64')
StaInfo['inland']=StaInfo['inland'].astype('int')
StaInfo['stt_date']=StaInfo['stt_date'].astype('datetime64')

StaInfo = StaInfo.drop_duplicates()
# StaInfo = StaInfo[~pd.isnull(StaInfo['stt_date'])].copy()

cols_truncate=['v_county','v_city','v_prcode','vf01015_cn','v_village','v_town','slm']
cols_length=[21,21,21,66,21,21,21]
for i,col in enumerate(cols_truncate):
    length=cols_length[i]
    StaInfo[col]=StaInfo[col].str[:length]

StaInfo.to_csv(script_dir / 'stainfo_bk.csv',index=False)
print('完成站点信息csv文件更新')