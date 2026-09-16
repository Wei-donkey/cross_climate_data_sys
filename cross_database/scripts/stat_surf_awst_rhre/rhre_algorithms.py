# -*- coding: utf-8 -*-
"""
Created on Tue Aug 11 16:26:06 2026

@author: Administrator
"""

from __future__ import annotations

import pandas as pd
import numpy as np
import datetime

stanum_threshold = 4  # 站点数量阈值的下限（即使按比例算出来的数更小，也至少要这么多站）
ratio_threshold = 0.05
r_earth = 6371


def compute_stanum_threshold(stanum_year, ratio=ratio_threshold, floor=stanum_threshold):
    """站点数量阈值 = max(floor, round(当年有效站点数 * ratio))。

    stanum_year 是"当年"有效站点数（按 year_stt <= year 过滤），国家站网是逐年建设起来的，
    早期年份可用站点很少，5%算出来会小于4——floor=4 主要是为了兜住这些早期年份，
    省级、区域级都一样，并非因为区域站点数天生比省少。
    """
    return max(floor, round(stanum_year * ratio))

# 单次过程评估（国家站该过程数据、所有站该过程日雨量、开始日期、结束日期、当年国家站数量），返回-->过程数据和日数据
# rhre_assessment can evaluate the process based on either surf data or surf+awst data.
def rhre_assessment(df_rhre_surf, df_r_all_station, date_stt, date_end,stanum_surf_year):
    # ========================= 2022-5-13 重构此次过程 国家站+区域站 暴雨数据 ====================
    strdate_stt = date_stt.strftime('%Y-%m-%d')
    strdate_end = date_end.strftime('%Y-%m-%d')

    print('统计区域性暴雨过程指标：' + strdate_stt + ' ' + strdate_end)

    # ===== 逐日统计相关量及综合强度指数(2022-6-15添加) =======
    # 输出的相关统计量:日期、最大日雨量、最大日雨量站点、暴雨站数、暴雨站号、总暴雨量、暴雨站次、平均单站日暴雨量、影响范围、综合强度指数
    df_rhre_days = pd.DataFrame(
        columns=['ddate', 'r_max', 'rmax_sta', 'rs_stanum', 'rs_stacodes', 'rs_ttl', 'rs_m', 'rs_i', 'rs_a', 'rs_z'])
    date_range = pd.date_range(strdate_stt, strdate_end, freq='1d')

    for date in date_range:
        DF_RS = df_r_all_station[(df_r_all_station['r'] >= 50) & (df_r_all_station.index == date)]

        rs_ttl = DF_RS['r'].sum()  # ====== 当天暴雨站点总降水量（国家站+区域站） ======
        rs_m = DF_RS.count().values[0]  # ====== 当天总暴雨站数（国家站+区域站）======
        rs_i = rs_ttl / rs_m  # ====== 当天平均单站暴雨量（国家站+区域站） ======
        r_max = DF_RS['r'].max()  # ====== 当天最大日降水量（国家站+区域站）======
        rmax_sta = DF_RS.loc[DF_RS['r'] == r_max, 'stacode'].values[0]  # ====== 当天最大日降水站点（国家站+区域站）======
        rs_stanum = rs_m  # ====== 当天暴雨站数（国家站+区域站）======
        rs_stacodes = DF_RS[0:499]['stacode'].astype(str).str.cat(sep=',')  # ====== 当天暴雨站号（国家站+区域站）======

        # ====== 利用单站暴雨量（国家站+区域站）、影响范围（国家站）计算当天综合影响指数 ======
        rs_t = 1  # 设定为1日
        Data_RS_SURF_day = df_rhre_surf[df_rhre_surf['date'] == date]
        rs_stanum_SURF = Data_RS_SURF_day.iloc[0]['RS_cnt']
        rs_a = 100*rs_stanum_SURF / stanum_surf_year  # 2023-02-28：修改影响范围为暴雨面积占比

        rs_z = rs_i * rs_a ** 0.5 * rs_t ** 0.5  # 当天综合强度指数

        DF_RS_day = pd.DataFrame(
            {'ddate': date, 'r_max': r_max, 'rmax_sta': rmax_sta, 'rs_stanum': rs_stanum, 'rs_stacodes': rs_stacodes \
                , 'rs_ttl': rs_ttl, 'rs_m': rs_m, 'rs_i': round(rs_i,1), 'rs_a': round(rs_a,2), 'rs_z': round(rs_z,1)}, index=[0])

        df_rhre_days = pd.concat([df_rhre_days, DF_RS_day], ignore_index=True)

    rs_t = (date_end - date_stt).days + 1  # 持续天数

    # ====== 按站号分组统计（国家站+区域站） ======
    Racc_stations = df_r_all_station.groupby('stacode').agg('sum')
    racc_max = Racc_stations.max().values[0]
    raccmax_sta = Racc_stations.loc[Racc_stations['r'] == racc_max].index[0]

    # ====== 暴雨站数（国家站+区域站）======
    DF_RS = df_r_all_station.loc[df_r_all_station['r'] >= 50].copy()
    Rmax_stations = DF_RS.groupby('stacode').agg('max')
    rs_stanum = Rmax_stations.count()
    # ====== 影响的站号（国家站+区域站）======
    rs_stacodes = Rmax_stations[0:599].index.astype(str).str.cat(sep=',')

    # ====== 利用每日指标计算过程指标(每日综合强度累加得到过程综合强度) ======
    r_max = df_rhre_days['r_max'].max()  # 最大日雨量
    rmax_sta = df_rhre_days.loc[df_rhre_days['r_max'] == r_max, 'rmax_sta'].values[0]  # 最大日雨量出现站点
    rmax_date = df_rhre_days.loc[df_rhre_days['r_max'] == r_max, 'ddate'].values[0]  # 最大日雨量出现日期
    rs_ttl = df_rhre_days['rs_ttl'].sum()  # 总雨量
    rs_m = df_rhre_days['rs_m'].sum()  # 暴雨总站次
    rs_i = df_rhre_days['rs_i'].mean()  # 平均日雨强

    # rs_a = df_rhre_days['rs_a'].max()  # 最大影响范围
    # rs_stanum_SURF = df_rhre_surf.iloc[:, 1:-2].max(axis='index').count()
    # rs_a = 100*rs_stanum_SURF / stanum_surf_year  # 2023-02-28：修改影响范围为暴雨面积占比
    rs_stanum_SURF = df_rhre_surf['RS_cnt'].sum()  #多日总站数
    rs_a = 100 * rs_stanum_SURF / (rs_t*stanum_surf_year)  #平均影响范围


    df_rhre_days['x'] = df_rhre_days.index + 1
    df_rhre_days['impact_factor'] = 1 - 0.5 / (1 + np.exp(-1.4 * (df_rhre_days['x'] - 6)))
    df_rhre_days['rs_z2'] = df_rhre_days['rs_z'] * df_rhre_days['impact_factor']
    rs_z = df_rhre_days['rs_z2'].sum()
    df_rhre_days.drop(['x','impact_factor','rs_z2'],axis='columns',inplace=True)

    df_rhre_process = pd.DataFrame(
        {'date_stt': date_stt, 'date_end': date_end, 'rs_t': rs_t, 'racc_max': racc_max, 'raccmax_sta': raccmax_sta \
            , 'r_max': r_max, 'rmax_sta': rmax_sta, 'rmax_date': rmax_date, 'rs_stanum': rs_stanum,
         'rs_stacodes': rs_stacodes \
            , 'rs_ttl': rs_ttl, 'rs_m': rs_m, 'rs_i': rs_i, 'rs_a': rs_a, 'rs_z': rs_z})

    return df_rhre_process, df_rhre_days    


def rhre_recognition(df_year_hrain, lonlat_stations, region_name, stanum_threshold, distance_km):
    """识别区域性暴雨过程起止日期。

    stanum_threshold: 达到区域性暴雨标准所需的（邻近）站点数量阈值，由调用方算好传入
                       （建议用 compute_stanum_threshold() 算：比例阈值，不低于floor）。
    distance_km: 邻近站判定标准（公里）。省级/区域级范围差异很大，务必按调用场景传入，
                 不要依赖模块内的默认值——之前 distance 是模块级全局变量，调用方即使自己
                 算了别的值也不会真正生效，这里改成显式参数以避免同样的问题再次发生。
    """
    print('识别 ' + region_name + ' 区域性暴雨过程起止日期')

    # =========== 1.首先判断所有 RHRE 过程的起止日期（国家站） ==============
    rhres_date = pd.DataFrame(columns=['date_stt', 'date_end'])

    # print('统计 ' + region_name + ' 区域性暴雨过程起止日期：' + str(year))
    df_year_hrain['reg_RS'] = 0  # 默认所有日期区域性暴雨标记为：非

    # ======= 1.1 从第一个暴雨日开始循环判断是否达到区域性暴雨标准 ===========
    for idx_date in df_year_hrain.index:

        # 只读取当前日各站点数据，不读取 RS_cnt 和 reg_RS
        Data_RS_day = df_year_hrain.loc[idx_date][:-2]
        Data_RS_day.dropna(inplace=True)

        # 暴雨站不到阈值，不判断邻近站点数量（2024-05-17）
        if len(Data_RS_day) < stanum_threshold:
            continue  # 如果没有达到数量要求，则直接下一日

        # 从第一个站开始与其他各站（从最远的站开始）测量距离
        idx_stacodes = Data_RS_day.index
        for idx_staA in idx_stacodes:
            # 邻近站点数归1(它自己)
            sta_cnt = 1

            lon_a = lonlat_stations.loc['lon', idx_staA] * (np.pi / 180)
            lat_a = lonlat_stations.loc['lat', idx_staA] * (np.pi / 180)

            for idx_staB in idx_stacodes[::-1]:

                lon_b = lonlat_stations.loc['lon', idx_staB] * (np.pi / 180)
                lat_b = lonlat_stations.loc['lat', idx_staB] * (np.pi / 180)

                Dist = r_earth * np.arccos(
                    np.sin(lat_a) * np.sin(lat_b) + np.cos(lat_a) * np.cos(lat_b) * np.cos(lon_a - lon_b))
                if Dist <= distance_km: sta_cnt += 1
                if Dist == 0: sta_cnt -= 1  # 如果是与本站比，则暴雨站点要减去本站（2024-5-17）

                # 如果达到了区域性暴雨要求的站点数阈值，则标记该日为区域性暴雨日，并退出判断，否则继续判断
                if sta_cnt >= stanum_threshold:
                    df_year_hrain.loc[idx_date, 'reg_RS'] = 1
                    break
                else:
                    continue

            else:
                continue
            break

    df_year_hrain = df_year_hrain.copy()

    # ========= 1.2 从第一个区域性暴雨日开始判断该过程的起止日期 ============
    # 下面copy()是为了避免Warning: A value is trying to be set on a copy of a slice from a DataFrameA
    # df_year_hrain = df_year_hrain.copy()
    df_year_hrain.reset_index(level=0, inplace=True)
    df_year_hrain.rename(columns={'index': 'date'}, inplace=True)  # 将日期index变为列（ddate）

    # 对起始日期进行初始化；对是否合并过过程进行初始化
    date_stt = datetime.datetime(1901, 1, 1); Flag_combined = False
    for idx in df_year_hrain.index:

        if date_stt == datetime.datetime(1901, 1, 1):
            if df_year_hrain.loc[idx, 'reg_RS'] == 0: continue  # 如果起始的这天没有达到RS标准，则跳过
            date_stt = df_year_hrain.loc[idx, 'date']  # 设reg_RS==1当天为区域性暴雨起始日

        date_end = df_year_hrain.loc[idx, 'date']
        if df_year_hrain.loc[idx, 'reg_RS'] == 1: Flag_date_end = 'IsRS'  # 标记结束日是否达到RS标准（2024-5-17）
        if df_year_hrain.loc[idx, 'reg_RS'] == 0: Flag_date_end = 'NotRS'  # 标记结束日是否达到RS标准


        if idx == df_year_hrain.index[-1]:  # 已经是最后一条记录，则记录起止日期
            if df_year_hrain.loc[idx, 'reg_RS'] == 0:   # 如果最后这天没有达到RS标准，则前一天作为结束（前一天必须达RS标准）
                date_end = date_end - datetime.timedelta(days=1)
            rhres_date = rhres_date.append({'date_stt': date_stt, 'date_end': date_end}, ignore_index=True)
        else:

            # 下一记录与本条记录日期连续
            if df_year_hrain.loc[idx, 'date'] + datetime.timedelta(days=1) == df_year_hrain.loc[idx + 1]['date']:

                # 当前日为“间隔日”，判断下一日是否也是“间隔日”：若是，则结束
                if (Flag_date_end == 'NotRS') & (df_year_hrain.loc[idx+1, 'reg_RS'] == 0):
                    date_end = date_end - datetime.timedelta(days=1)  # 则前2天作为结束（因为前一天未达RS标准）
                    rhres_date = rhres_date.append({'date_stt': date_stt, 'date_end': date_end}, ignore_index=True)
                    date_stt = datetime.datetime(1901, 1, 1)  # 对起始日期进行初始化
                    Flag_combined = False
                    continue

                # 当前日为“间隔日”，判断当前日之前是否已经合并过两个过程：若是，则结束(2025-07-07添加)
                if (Flag_date_end == 'NotRS') & (Flag_combined == True):
                    date_end = date_end - datetime.timedelta(days=1)
                    rhres_date = rhres_date.append({'date_stt': date_stt, 'date_end': date_end}, ignore_index=True)
                    date_stt = datetime.datetime(1901, 1, 1)  # 对起始日期进行初始化
                    Flag_combined = False
                    continue

                if (Flag_date_end == 'NotRS') & (df_year_hrain.loc[idx+1, 'reg_RS'] == 1):  # 当前为RS后的第1间隔日
                    Flag_combined = True
                    continue

                if (Flag_date_end == 'IsRS'):
                    continue

            # 下一记录与本条记录日期上不连续，则本次区域性降水过程结束，统计此次过程相关指标
            if df_year_hrain.loc[idx, 'date'] + datetime.timedelta(days=1) \
                    != df_year_hrain.loc[idx + 1]['date']:
                if df_year_hrain.loc[idx, 'reg_RS'] == 0:  # 如果最后这天没有达到RS标准，则前一天作为结束（前一天必须达RS标准）
                    date_end = date_end - datetime.timedelta(days=1)
                rhres_date = rhres_date.append({'date_stt': date_stt, 'date_end': date_end}, ignore_index=True)
                date_stt = datetime.datetime(1901, 1, 1)  # 对起始日期进行初始化
                Flag_combined = False

    return rhres_date, df_year_hrain


if __name__=='__main__':
    print('This module is intended for import only - please do not run it directly.')