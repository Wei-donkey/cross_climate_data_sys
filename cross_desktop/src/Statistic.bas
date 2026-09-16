Attribute VB_Name = "Statistic"


'2026-8-6 重新定义各类型站点标记
Public SelectSURFExt As Single '大区域站点：1-通过省份筛选的国家站，0.9-登陆默认的国家站（86个，不包含深圳观象台），0.5-任意多选国家站 ==== 系统启动默认0.9

Public SelectSURF As Single '国家站点：1-通过城市筛选的国家站，0.9-登录默认的国家站（86个，不包含深圳观象台），0.5-任意多选国家站 ========= 系统启动默认0.9

Public SelectAWST As Single '区域站站点：1-通过城市筛选的区域站，0.9-登陆默认的区域站（不包含周边站），0.5-任意多选国家站 ================= 系统启动默认0.9（陆面站、水文站、海洋站共享该标识）
Public bCheckInland As Boolean, bCheckHydro As Boolean, bCheckOcean As Boolean   '是否选择响应复选框：陆面站、水文站、海洋站

'2026-8-6 定义 sttDate_valid = "stt_date<>to_date('1899-09-09','yyyy-mm-dd')"
Public sttDate_valid As String

'******站点信息相关变量******
Public Type StaInfo
  Prov As String
  Region As String '行政分区
  City As String
  County As String
  Town As String
  StaType As String '站点类型：区分类型，用于计算多年平均平均值
  stacode As String
  staname As String
  Longitude As Single
  Latitude As Single
''''  GBA As Integer  '是否属于大湾区（1-是，0-否）
  DateSTT As String
  YearSTT As Integer '判断有效资料所在的年份，如DatesSTT=1958-12-01，查询08-17至08-26的资料，有效起始年为1959
  Inland As Integer  '2023-02-14：判断站点是否在陆地（默认不勾选非陆地站点）
  Hydro As Integer '2026-03-23：判断站点是否为谁文章（默认不勾选水文站）
  Selected As Boolean
End Type

'全部国家站点；已选国家站点
Public SURFInfo1() As StaInfo, SURFInfo2() As StaInfo

'全部自动站点；已选自动站点
Public AWSTInfo1() As StaInfo, AWSTInfo2() As StaInfo
'用户已选择的国家站/自动站站点数量
Public StaNum_SURF%, StaNum_AWST%

'用户已选的全部站点（国家站 or 自动站 or 国家站+自动站）
Public StaInfo() As StaInfo
'用户已选的全部站点
Public StaNum%

Public FlagZone As String '统计的空间尺度：STA,TOWN,CNTY,CITY

'2025-06-05 添加：是否选择了当前城市的所有县区
Public Flag_AllCounty As Boolean


'2026-07-16 区域统计功能：区域自定义站点
Public ExtStaInfo1() As StaInfo, ExtStaInfo2() As StaInfo
Public ExtStaInfo() As StaInfo
Public ExtStaNum%

Public ExtStaInfoLoaded As Boolean

Public FlagZoneExt As String '区域统计功能的统计结果尺度：STA,PROV,CNTY,CITY
Public ShowLonLatExt As Boolean '判断是否输出经纬度


Public Zone_SUM() As Single '区域累计值（二维变量）
Public Zone_NUM() As Single '区域累加站点数（二维变量）
Public Zone_AVG() As Single '区域平均值（二维变量）
Public Zone_MAX() As Single '区域最大值（二维变量）
Public Zone_MIN() As Single '区域最小值（二维变量）


'2025-06-01 添加结果三列的区域统计结果
Public Zone_COLS_SUM() As Single '区域累计值（二维变量）
Public Zone_COLS_NUM() As Single '区域累加站点数（二维变量）
Public Zone_COLS_AVG() As Single '区域平均值（二维变量）
Public Zone_COLS_MAX() As Single '区域最大值（二维变量）
Public Zone_COLS_MIN() As Single '区域最小值（二维变量）


'******任意时段统计（所有统计量）各结果指标变量******
Public Type Period_Result
  stacode As String
  staname As String

  stat As Single '当年统计值
  stat_date As Date '当年极值出现日期

  stat_hist As Single '累年平均/极值
  stat_hist_date As Date '累年极值出现日期
  stat_diff1 As Single '与累年平均距平

  stat_year As Single '某年统计值
  stat_year_date As Date '某年极值出现日期
  stat_diff2 As Single '与某年平均距平

  rank1 As Integer '排名（大-小）
  rank2 As Integer '排名（小-大）
  rank_years As String '排名年份
  maxyear As Integer '最大值年
  maxvalue As Single '最大值
  minyear As Integer '最小值年
  minvalue As Single '最小值
End Type
'任意时段结果变量：站点，区域平均、最大、最小
Public PRD_STA() As Period_Result
Public PRD_SUM() As Period_Result, PRD_NUM() As Period_Result, PRD_AVG() As Period_Result, PRD_MAX() As Period_Result, PRD_MIN() As Period_Result
'对上述变量开展多行统计，形成三行的变量
Public ROWS_Stat_Period() As Period_Result


'******季节划分（起始日）结果各分量******
Public Type Season_Result
  stacode As String
  staname As String
  OnsetNorm As Date  '常年起始日
  T_IdxNorm As Long  '起始日在 T 序列中的序号

  OnsetCurr As Date  '当年初次起始日
  T_IdxCurr As Long '初次起始日在 T 序列中的序号

  OnsetComp As Date  '某年起始日
  T_IdxComp As Long '初次起始日在 T 序列中的序号

  diffNorm As Long  '与常年距平
  diffComp As Long  '与某年距平
  Town As String
  County As String
  City As String
End Type
'各选定站点的季节起始日结果
Public SS_STA() As Season_Result
Public SS_SUM() As Season_Result, SS_NUM() As Season_Result, SS_AVG() As Season_Result, SS_MAX() As Season_Result, SS_MIN() As Season_Result
Public SeasonNum%
'对上述变量开展多行统计，形成三行的变量
Public ROWS_Stat_Seasons() As Season_Result


'******气温5日滑动平均******
Public Type TSeries
  T As Single
  ddate As Date
End Type
'5日滑动平均气温：当年数据序列、某年数据序列、常年数据序列
Public T5d_Curr() As TSeries, T5d_Comp() As TSeries, T5d_Norm() As TSeries
'逐日气温：当年数据序列、某年数据序列、常年数据序列
Public T_Curr() As TSeries, T_Comp() As TSeries, T_Norm() As TSeries


'******某次降温过程统计结果指标变量******
Public Type ColdAir_Result
  stacode As String
  staname As String
  ilevel As Single  '冷空气等级：0-无，1-弱，2-中，3-强，4-寒潮
  BDate As Date  '降温第一日
  EDate As Date  '回温前一日
  idays As Single  '降温日数
  tv_24hmax As Single  '最大24小时降温
  tv_48hmax As Single  '最大48小时降温
  tv_acc As Single '累计降温
  T As Single  '降温过程的平均气温
  t_min As Single  '降温过程的最低气温
  Town As String
  County As String
  City As String
End Type
'某次过程结果变量：站点，区域平均、
'2022年4月25日，定义最强过程变量CA_STA()，所有过程变量CA_ALL()
Public CA_STA() As ColdAir_Result, CA_ALL() As ColdAir_Result
Public CA_SUM() As ColdAir_Result, CA_NUM() As ColdAir_Result, CA_AVG() As ColdAir_Result, CA_MAX() As ColdAir_Result, CA_MIN() As ColdAir_Result
'对上述变量开展多行统计，形成三行的变量
Public ROWS_Stat_ColdAir() As ColdAir_Result


'小时数据结果二维变量：站点，区域平均、最大、最小（镇、县、市共用）
Public HOR_STA() As Single, HOR_AVG() As Single, HOR_MAX() As Single, HOR_MIN() As Single
'日数据结果二维变量：站点，区域平均、最大、最小
Public DAY_STA() As Single, DAY_AVG() As Single, DAY_MAX() As Single, DAY_MIN() As Single


'旬数据结果二维变量：站点，区域平均、最大、最小
Public TEN_STA() As Single, TEN_AVG() As Single, TEN_MAX() As Single, TEN_MIN() As Single
'月数据结果二维变量：站点，区域平均、最大、最小
Public MON_STA() As Single, MON_AVG() As Single, MON_MAX() As Single, MON_MIN() As Single
'季数据结果二维变量：站点，区域平均、最大、最小
Public QTR_STA() As Single, QTR_AVG() As Single, QTR_MAX() As Single, QTR_MIN() As Single
'年/任意逐年数据结果二维变量：站点，区域平均、最大、最小
Public YER_STA() As Single, YER_AVG() As Single, YER_MAX() As Single, YER_MIN() As Single

'对上述变量开展多行统计，形成三行多列的二维变量
Public ROWS_Stat() As Single
'对上述变量开展多列统计，形成三列多行的二维变量
Public COLS_Stat() As Single
'对COLS_Stat开展多行统计，形成三行三列的二维变量
Public COLS_ROWS_Stat() As Single
'==================== NewType.bas 迁移内容结束 ====================


Private Sub Accumulate_COLS_ZONE(Data_sta!(), ByVal i As Integer, ByVal k As Integer)
  Dim j As Integer
  For j = 1 To 3
    If Data_sta(i, j) <> -9999 Then
      If Data_sta(i, j) > Zone_COLS_MAX(k, j) Then Zone_COLS_MAX(k, j) = Data_sta(i, j)
      If Data_sta(i, j) < Zone_COLS_MIN(k, j) Then Zone_COLS_MIN(k, j) = Data_sta(i, j)
      Zone_COLS_SUM(k, j) = Zone_COLS_SUM(k, j) + Data_sta(i, j)
      Zone_COLS_NUM(k, j) = Zone_COLS_NUM(k, j) + 1
    End If
  Next j
End Sub

Public Sub Stat_COLS_ZONE(Data_sta!()) '站点数据的 最后三列（平均或累计/最大/最小）
  Dim iRows%
  Dim i%, j%, k%, l%
  
  If FlagZone = "CITY" Then iRows = CityNum
  If FlagZone = "CNTY" Then iRows = CountyNum
  If FlagZone = "TOWN" Then iRows = TownNum
  
  ReDim Zone_COLS_SUM(iRows, 3): ReDim Zone_COLS_NUM(iRows, 3): ReDim Zone_COLS_AVG(iRows, 3): ReDim Zone_COLS_MAX(iRows, 3): ReDim Zone_COLS_MIN(iRows, 3)
  For k = 1 To iRows
    For j = 1 To 3
      Zone_COLS_SUM(k, j) = 0: Zone_COLS_NUM(k, j) = 0: Zone_COLS_AVG(k, j) = -9999: Zone_COLS_MAX(k, j) = -9999: Zone_COLS_MIN(k, j) = 999999
    Next j
  Next k
  
    
  '逐站点循环，统计各区域累计/最大/最小/计数值
  For i = 1 To StaNum
    For k = 1 To iRows
    
      If FlagZone = "CITY" Then
        If StaInfo(i).City = CityInfo2(k).City Then
          Call Accumulate_COLS_ZONE(Data_sta, i, k)
          Exit For
        End If
      
      ElseIf FlagZone = "CNTY" Then
        If StaInfo(i).City = CountyInfo2(k).City And StaInfo(i).County = CountyInfo2(k).County Then
          Call Accumulate_COLS_ZONE(Data_sta, i, k)
          Exit For
        End If
      
      ElseIf FlagZone = "TOWN" Then
        If StaInfo(i).City = TownInfo2(k).City And StaInfo(i).County = TownInfo2(k).County And StaInfo(i).Town = TownInfo2(k).Town Then
          Call Accumulate_COLS_ZONE(Data_sta, i, k)
          Exit For
        End If
      End If
      
    Next k
  Next i
      
  '计算各区域平均值，并将空值赋值为-9999
  For k = 1 To iRows
    For j = 1 To 3
'      If Zone_Cols_NUM(k, j) <> 0 Then Zone_Cols_AVG(k, j) = Round(Zone_Cols_SUM(k, j) / Zone_Cols_NUM(k, j), 1)
      If Zone_COLS_NUM(k, j) <> 0 Then Zone_COLS_AVG(k, j) = Format(Zone_COLS_SUM(k, j) / Zone_COLS_NUM(k, j), "0.0")
      If Zone_COLS_NUM(k, j) = 0 Then Zone_COLS_AVG(k, j) = -9999: Zone_COLS_SUM(k, j) = -9999
      If Zone_COLS_MAX(k, j) = -9999 Then Zone_COLS_MAX(k, j) = -9999
      If Zone_COLS_MIN(k, j) = 999999 Then Zone_COLS_MIN(k, j) = -9999
    Next j
  Next k
End Sub



Private Sub Accumulate_ZONE(Data_sta!(), ByVal i As Integer, ByVal k As Integer, ByVal iTimes As Integer)
  Dim j As Integer
  For j = 1 To iTimes
    If Data_sta(i, j) <> -9999 Then
      If Data_sta(i, j) > Zone_MAX(k, j) Then Zone_MAX(k, j) = Data_sta(i, j)
      If Data_sta(i, j) < Zone_MIN(k, j) Then Zone_MIN(k, j) = Data_sta(i, j)
      Zone_SUM(k, j) = Zone_SUM(k, j) + Data_sta(i, j)
      Zone_NUM(k, j) = Zone_NUM(k, j) + 1
    End If
  Next j
End Sub

Public Sub Stat_ZONE(Data_sta!(), iTimes%) '二维站点数据，小时数/日数/旬数/月数...、区域标识 + 最后三列（平均或累计/最大/最小）
  Dim iRows%
  Dim i%, j%, k%, l%
  
  If FlagZone = "CITY" Then iRows = CityNum
  If FlagZone = "CNTY" Then iRows = CountyNum
  If FlagZone = "TOWN" Then iRows = TownNum
  
  ReDim Zone_SUM(iRows, iTimes): ReDim Zone_NUM(iRows, iTimes): ReDim Zone_AVG(iRows, iTimes): ReDim Zone_MAX(iRows, iTimes): ReDim Zone_MIN(iRows, iTimes)
  For k = 1 To iRows
    For j = 1 To iTimes
      Zone_SUM(k, j) = 0: Zone_NUM(k, j) = 0: Zone_AVG(k, j) = -9999: Zone_MAX(k, j) = -9999: Zone_MIN(k, j) = 999999
    Next j
  Next k
  
    
  '逐站点循环，统计各区域累计/最大/最小/计数值
  For i = 1 To StaNum
    For k = 1 To iRows
      If FlagZone = "CITY" Then
        If StaInfo(i).City = CityInfo2(k).City Then
          Call Accumulate_ZONE(Data_sta, i, k, iTimes)
          Exit For
        End If
      
      ElseIf FlagZone = "CNTY" Then
        If StaInfo(i).City = CountyInfo2(k).City And StaInfo(i).County = CountyInfo2(k).County Then
          Call Accumulate_ZONE(Data_sta, i, k, iTimes)
          Exit For
        End If
      
      ElseIf FlagZone = "TOWN" Then
        If StaInfo(i).City = TownInfo2(k).City And StaInfo(i).County = TownInfo2(k).County And StaInfo(i).Town = TownInfo2(k).Town Then
          Call Accumulate_ZONE(Data_sta, i, k, iTimes)
          Exit For
        End If
      End If
      
    Next k
  Next i
      
  '计算各区域平均值，并将空值赋值为-9999
  For k = 1 To iRows
    For j = 1 To iTimes
'      If Zone_NUM(k, j) <> 0 Then Zone_AVG(k, j) = Round(Zone_SUM(k, j) / Zone_NUM(k, j), 1)
      If Zone_NUM(k, j) <> 0 Then Zone_AVG(k, j) = Format(Zone_SUM(k, j) / Zone_NUM(k, j), "0.0")
      If Zone_NUM(k, j) = 0 Then Zone_AVG(k, j) = -9999: Zone_SUM(k, j) = -9999
      If Zone_MAX(k, j) = -9999 Then Zone_MAX(k, j) = -9999
      If Zone_MIN(k, j) = 999999 Then Zone_MIN(k, j) = -9999
    Next j
  Next k

End Sub




'统计任意时段数据的区域平均/最大/最小值
Private Sub Accumulate_ZONE_Period(Data_sta() As Period_Result, ByVal i As Integer, ByVal k As Integer)
  Dim j As Integer
  For j = 1 To 5
    If j = 1 Then
      If Data_sta(i).stat <> -9999 Then
        If Data_sta(i).stat > PRD_MAX(k).stat Then PRD_MAX(k).stat = Data_sta(i).stat: PRD_MAX(k).stat_date = Data_sta(i).stat_date
        If Data_sta(i).stat < PRD_MIN(k).stat Then PRD_MIN(k).stat = Data_sta(i).stat: PRD_MIN(k).stat_date = Data_sta(i).stat_date
        PRD_SUM(k).stat = PRD_SUM(k).stat + Data_sta(i).stat
        PRD_NUM(k).stat = PRD_NUM(k).stat + 1
      End If
    ElseIf j = 2 Then
      If Data_sta(i).stat_hist <> -9999 Then
        If Data_sta(i).stat_hist > PRD_MAX(k).stat_hist Then PRD_MAX(k).stat_hist = Data_sta(i).stat_hist: PRD_MAX(k).stat_hist_date = Data_sta(i).stat_hist_date
        If Data_sta(i).stat_hist < PRD_MIN(k).stat_hist Then PRD_MIN(k).stat_hist = Data_sta(i).stat_hist: PRD_MIN(k).stat_hist_date = Data_sta(i).stat_hist_date
        PRD_SUM(k).stat_hist = PRD_SUM(k).stat_hist + Data_sta(i).stat_hist
        PRD_NUM(k).stat_hist = PRD_NUM(k).stat_hist + 1
      End If
    ElseIf j = 3 Then
      If Data_sta(i).stat_diff1 <> -9999 Then
        If Data_sta(i).stat_diff1 > PRD_MAX(k).stat_diff1 Then PRD_MAX(k).stat_diff1 = Data_sta(i).stat_diff1
        If Data_sta(i).stat_diff1 < PRD_MIN(k).stat_diff1 Then PRD_MIN(k).stat_diff1 = Data_sta(i).stat_diff1
        PRD_SUM(k).stat_diff1 = PRD_SUM(k).stat_diff1 + Data_sta(i).stat_diff1
        PRD_NUM(k).stat_diff1 = PRD_NUM(k).stat_diff1 + 1
      End If
    ElseIf j = 4 Then
      If Data_sta(i).stat_year <> -9999 Then
        If Data_sta(i).stat_year > PRD_MAX(k).stat_year Then PRD_MAX(k).stat_year = Data_sta(i).stat_year: PRD_MAX(k).stat_year_date = Data_sta(i).stat_year_date
        If Data_sta(i).stat_year < PRD_MIN(k).stat_year Then PRD_MIN(k).stat_year = Data_sta(i).stat_year: PRD_MIN(k).stat_year_date = Data_sta(i).stat_year_date
        PRD_SUM(k).stat_year = PRD_SUM(k).stat_year + Data_sta(i).stat_year
        PRD_NUM(k).stat_year = PRD_NUM(k).stat_year + 1
      End If
    ElseIf j = 5 Then
      If Data_sta(i).stat_diff2 <> -9999 Then
        If Data_sta(i).stat_diff2 > PRD_MAX(k).stat_diff2 Then PRD_MAX(k).stat_diff2 = Data_sta(i).stat_diff2
        If Data_sta(i).stat_diff2 < PRD_MIN(k).stat_diff2 Then PRD_MIN(k).stat_diff2 = Data_sta(i).stat_diff2
        PRD_SUM(k).stat_diff2 = PRD_SUM(k).stat_diff2 + Data_sta(i).stat_diff2
        PRD_NUM(k).stat_diff2 = PRD_NUM(k).stat_diff2 + 1
      End If
    End If
  Next j
End Sub

Public Sub Stat_ZONE_Period(Data_sta() As Period_Result, FlagStat$, SelField_Initial$)  '任意时段站点数据、区域标识、统计量标识
  Dim iRows%
  Dim i%, j%, k%, l%
  
  If FlagZone = "CITY" Then iRows = CityNum
  If FlagZone = "CNTY" Then iRows = CountyNum
  If FlagZone = "TOWN" Then iRows = TownNum
  
  ReDim PRD_SUM(1 To iRows): ReDim PRD_NUM(1 To iRows): ReDim PRD_AVG(1 To iRows): ReDim PRD_MAX(1 To iRows): ReDim PRD_MIN(1 To iRows)
  For k = 1 To iRows
    PRD_SUM(k).stat = 0: PRD_SUM(k).stat_hist = 0: PRD_SUM(k).stat_diff1 = 0: PRD_SUM(k).stat_year = 0: PRD_SUM(k).stat_diff2 = 0
    PRD_NUM(k).stat = 0: PRD_NUM(k).stat_hist = 0: PRD_NUM(k).stat_diff1 = 0: PRD_NUM(k).stat_year = 0: PRD_NUM(k).stat_diff2 = 0
    PRD_AVG(k).stat = -9999: PRD_AVG(k).stat_date = CDate("1899-09-09")
    PRD_AVG(k).stat_hist = -9999: PRD_AVG(k).stat_hist_date = CDate("1899-09-09")
    PRD_AVG(k).stat_diff1 = -9999
    PRD_AVG(k).stat_year = -9999: PRD_AVG(k).stat_year_date = CDate("1899-09-09")
    PRD_AVG(k).stat_diff2 = -9999
    PRD_AVG(k).rank1 = -9999: PRD_AVG(k).rank2 = -9999: PRD_AVG(k).rank_years = "-9999": PRD_AVG(k).maxyear = -9999: PRD_AVG(k).maxvalue = -9999: PRD_AVG(k).minyear = -9999: PRD_AVG(k).minvalue = -9999
    PRD_MAX(k).stat = -9999: PRD_MAX(k).stat_date = CDate("1899-09-09")
    PRD_MAX(k).stat_hist = -9999: PRD_MAX(k).stat_hist_date = CDate("1899-09-09")
    PRD_MAX(k).stat_diff1 = -9999:
    PRD_MAX(k).stat_year = -9999: PRD_MAX(k).stat_year_date = CDate("1899-09-09")
    PRD_MAX(k).stat_diff2 = -9999
    PRD_MAX(k).rank1 = -9999: PRD_MAX(k).rank2 = -9999: PRD_MAX(k).rank_years = "-9999": PRD_MAX(k).maxyear = -9999: PRD_MAX(k).maxvalue = -9999: PRD_MAX(k).minyear = -9999: PRD_MAX(k).minvalue = -9999
    PRD_MIN(k).stat = 999999: PRD_MIN(k).stat_date = CDate("1899-09-09")
    PRD_MIN(k).stat_hist = 999999: PRD_MIN(k).stat_hist_date = CDate("1899-09-09")
    PRD_MIN(k).stat_diff1 = 999999:
    PRD_MIN(k).stat_year = 999999: PRD_MIN(k).stat_year_date = CDate("1899-09-09")
    PRD_MIN(k).stat_diff2 = 999999
    PRD_MIN(k).rank1 = -9999: PRD_MIN(k).rank2 = -9999: PRD_MIN(k).rank_years = "-9999": PRD_MIN(k).maxyear = -9999: PRD_MIN(k).maxvalue = -9999: PRD_MIN(k).minyear = -9999: PRD_MIN(k).minvalue = -9999
  Next k
    
  '逐站点循环，统计各区域累计/最大/最小/计数值
  For i = 1 To StaNum
    For k = 1 To iRows
    
      If FlagZone = "CITY" Then
        If StaInfo(i).City = CityInfo2(k).City Then
          Call Accumulate_ZONE_Period(Data_sta, i, k)
          Exit For
        End If
      
      ElseIf FlagZone = "CNTY" Then
        If StaInfo(i).City = CountyInfo2(k).City And StaInfo(i).County = CountyInfo2(k).County Then
          Call Accumulate_ZONE_Period(Data_sta, i, k)
          Exit For
        End If
      
      ElseIf FlagZone = "TOWN" Then
        If StaInfo(i).City = TownInfo2(k).City And StaInfo(i).County = TownInfo2(k).County And StaInfo(i).Town = TownInfo2(k).Town Then
          Call Accumulate_ZONE_Period(Data_sta, i, k)
          Exit For
        End If
      End If
      
    Next k
  Next i
      
  '计算各区域平均值，并将空值赋值为-9999
  For k = 1 To iRows
    For j = 1 To 5
        If j = 1 Then
          If PRD_NUM(k).stat <> 0 Then PRD_AVG(k).stat = Format(PRD_SUM(k).stat / PRD_NUM(k).stat, "0.0")
          If PRD_NUM(k).stat = 0 Then PRD_AVG(k).stat = -9999: PRD_SUM(k).stat = -9999
          If PRD_MAX(k).stat = -9999 Then PRD_MAX(k).stat = -9999: PRD_MAX(k).stat_date = CDate("1899-09-09")
          If PRD_MIN(k).stat = 999999 Then PRD_MIN(k).stat = -9999: PRD_MIN(k).stat_date = CDate("1899-09-09")
        
        ElseIf j = 2 Then
          If PRD_NUM(k).stat_hist <> 0 Then PRD_AVG(k).stat_hist = Format(PRD_SUM(k).stat_hist / PRD_NUM(k).stat_hist, "0.0")
          If PRD_NUM(k).stat_hist = 0 Then PRD_AVG(k).stat_hist = -9999: PRD_SUM(k).stat_hist = -9999
          If PRD_MAX(k).stat_hist = -9999 Then PRD_MAX(k).stat_hist = -9999: PRD_MAX(k).stat_hist_date = CDate("1899-09-09")
          If PRD_MIN(k).stat_hist = 999999 Then PRD_MIN(k).stat_hist = -9999: PRD_MIN(k).stat_hist_date = CDate("1899-09-09")
        
        ElseIf j = 3 Then
          If PRD_AVG(k).stat <> -9999 And PRD_AVG(k).stat_hist <> -9999 Then
            PRD_AVG(k).stat_diff1 = Format(PRD_AVG(k).stat, "0.0") - Format(PRD_AVG(k).stat_hist, "0.0")
          End If
          
          If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Or SelField_Initial = "S" Then
            If FlagStat = "SUM" Then
              If PRD_AVG(k).stat <> -9999 And PRD_AVG(k).stat_hist <> -9999 Then
                If PRD_AVG(k).stat_hist <> 0 Then PRD_AVG(k).stat_diff1 = 100 * (PRD_AVG(k).stat - PRD_AVG(k).stat_hist) / PRD_AVG(k).stat_hist
                If PRD_AVG(k).stat_hist = 0 Then PRD_AVG(k).stat_diff1 = -9999
              End If
            End If
          End If
          
          If PRD_MAX(k).stat_diff1 = -9999 Then PRD_MAX(k).stat_diff1 = -9999
          If PRD_MIN(k).stat_diff1 = 999999 Then PRD_MIN(k).stat_diff1 = -9999
        
        ElseIf j = 4 Then
          If PRD_NUM(k).stat_year <> 0 Then PRD_AVG(k).stat_year = Format(PRD_SUM(k).stat_year / PRD_NUM(k).stat_year, "0.0")
          If PRD_NUM(k).stat_year = 0 Then PRD_AVG(k).stat_year = -9999: PRD_SUM(k).stat_year = -9999
          If PRD_MAX(k).stat_year = -9999 Then PRD_MAX(k).stat_year = -9999: PRD_MAX(k).stat_year_date = CDate("1899-09-09")
          If PRD_MIN(k).stat_year = 999999 Then PRD_MIN(k).stat_year = -9999: PRD_MIN(k).stat_year_date = CDate("1899-09-09")
        
        ElseIf j = 5 Then
          If PRD_AVG(k).stat <> -9999 And PRD_AVG(k).stat_year <> -9999 Then
            PRD_AVG(k).stat_diff2 = Format(PRD_AVG(k).stat, "0.0") - Format(PRD_AVG(k).stat_year, "0.0")
          End If
          
          If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Or SelField_Initial = "S" Then
            If FlagStat = "SUM" Then
              If PRD_AVG(k).stat <> -9999 And PRD_AVG(k).stat_year <> -9999 Then
                If PRD_AVG(k).stat_year <> 0 Then PRD_AVG(k).stat_diff2 = 100 * (PRD_AVG(k).stat - PRD_AVG(k).stat_year) / PRD_AVG(k).stat_year
                If PRD_AVG(k).stat_year = 0 Then PRD_AVG(k).stat_diff2 = -9999
              End If
            End If
          End If
          
          If PRD_MAX(k).stat_diff2 = -9999 Then PRD_MAX(k).stat_diff2 = -9999
          If PRD_MIN(k).stat_diff2 = 999999 Then PRD_MIN(k).stat_diff2 = -9999
        
        End If
    Next j
  Next k

End Sub





'统计任意时段数据的区域平均/最大/最小值
Public Sub Stat_ZONE_PeriodExt(Data_sta() As Period_Result, FlagStat$, SelField_Initial$) '任意时段站点数据、区域标识、统计量标识、要素字段名
  Dim iRows%
  Dim i%, j%, k%, l%
  
  If FlagZoneExt = "CITY" Then iRows = ExtCityNum
  If FlagZoneExt = "CNTY" Then iRows = ExtCountyNum
  If FlagZoneExt = "PROV" Then iRows = ExtProvNum
  
  ReDim PRD_SUM(1 To iRows): ReDim PRD_NUM(1 To iRows): ReDim PRD_AVG(1 To iRows): ReDim PRD_MAX(1 To iRows): ReDim PRD_MIN(1 To iRows)
  For k = 1 To iRows
    PRD_SUM(k).stat = 0: PRD_SUM(k).stat_hist = 0: PRD_SUM(k).stat_diff1 = 0: PRD_SUM(k).stat_year = 0: PRD_SUM(k).stat_diff2 = 0
    PRD_NUM(k).stat = 0: PRD_NUM(k).stat_hist = 0: PRD_NUM(k).stat_diff1 = 0: PRD_NUM(k).stat_year = 0: PRD_NUM(k).stat_diff2 = 0
    PRD_AVG(k).stat = -9999: PRD_AVG(k).stat_date = CDate("1899-09-09")
    PRD_AVG(k).stat_hist = -9999: PRD_AVG(k).stat_hist_date = CDate("1899-09-09")
    PRD_AVG(k).stat_diff1 = -9999
    PRD_AVG(k).stat_year = -9999: PRD_AVG(k).stat_year_date = CDate("1899-09-09")
    PRD_AVG(k).stat_diff2 = -9999
    PRD_AVG(k).rank1 = -9999: PRD_AVG(k).rank2 = -9999: PRD_AVG(k).rank_years = "-9999": PRD_AVG(k).maxyear = -9999: PRD_AVG(k).maxvalue = -9999: PRD_AVG(k).minyear = -9999: PRD_AVG(k).minvalue = -9999
    PRD_MAX(k).stat = -9999: PRD_MAX(k).stat_date = CDate("1899-09-09")
    PRD_MAX(k).stat_hist = -9999: PRD_MAX(k).stat_hist_date = CDate("1899-09-09")
    PRD_MAX(k).stat_diff1 = -9999:
    PRD_MAX(k).stat_year = -9999: PRD_MAX(k).stat_year_date = CDate("1899-09-09")
    PRD_MAX(k).stat_diff2 = -9999
    PRD_MAX(k).rank1 = -9999: PRD_MAX(k).rank2 = -9999: PRD_MAX(k).rank_years = "-9999": PRD_MAX(k).maxyear = -9999: PRD_MAX(k).maxvalue = -9999: PRD_MAX(k).minyear = -9999: PRD_MAX(k).minvalue = -9999
    PRD_MIN(k).stat = 999999: PRD_MIN(k).stat_date = CDate("1899-09-09")
    PRD_MIN(k).stat_hist = 999999: PRD_MIN(k).stat_hist_date = CDate("1899-09-09")
    PRD_MIN(k).stat_diff1 = 999999:
    PRD_MIN(k).stat_year = 999999: PRD_MIN(k).stat_year_date = CDate("1899-09-09")
    PRD_MIN(k).stat_diff2 = 999999
    PRD_MIN(k).rank1 = -9999: PRD_MIN(k).rank2 = -9999: PRD_MIN(k).rank_years = "-9999": PRD_MIN(k).maxyear = -9999: PRD_MIN(k).maxvalue = -9999: PRD_MIN(k).minyear = -9999: PRD_MIN(k).minvalue = -9999
  Next k
    
  '逐站点循环，统计各区域累计/最大/最小/计数值
  For i = 1 To ExtStaNum
    For k = 1 To iRows
      
      If FlagZoneExt = "PROV" Then
        If ExtStaInfo(i).Prov = ExtProvInfo2(k).Prov Then
          Call Accumulate_ZONE_Period(Data_sta, i, k)
          Exit For
        End If
      
      
      ElseIf FlagZoneExt = "CITY" Then
        If ExtStaInfo(i).Prov = ExtCityInfo2(k).Prov And ExtStaInfo(i).City = ExtCityInfo2(k).City Then
          Call Accumulate_ZONE_Period(Data_sta, i, k)
          Exit For
        End If
      
      ElseIf FlagZoneExt = "CNTY" Then
        If ExtStaInfo(i).Prov = ExtCountyInfo2(k).Prov And ExtStaInfo(i).City = ExtCountyInfo2(k).City And ExtStaInfo(i).County = ExtCountyInfo2(k).County Then
          Call Accumulate_ZONE_Period(Data_sta, i, k)
          Exit For
        End If
      
      End If
      
    Next k
  Next i
      
  '计算各区域平均值，并将空值赋值为-9999
  For k = 1 To iRows
    For j = 1 To 5
        If j = 1 Then
          If PRD_NUM(k).stat <> 0 Then PRD_AVG(k).stat = Format(PRD_SUM(k).stat / PRD_NUM(k).stat, "0.0")
          If PRD_NUM(k).stat = 0 Then PRD_AVG(k).stat = -9999: PRD_SUM(k).stat = -9999
          If PRD_MAX(k).stat = -9999 Then PRD_MAX(k).stat = -9999: PRD_MAX(k).stat_date = CDate("1899-09-09")
          If PRD_MIN(k).stat = 999999 Then PRD_MIN(k).stat = -9999: PRD_MIN(k).stat_date = CDate("1899-09-09")
        ElseIf j = 2 Then
          If PRD_NUM(k).stat_hist <> 0 Then PRD_AVG(k).stat_hist = Format(PRD_SUM(k).stat_hist / PRD_NUM(k).stat_hist, "0.0")
          If PRD_NUM(k).stat_hist = 0 Then PRD_AVG(k).stat_hist = -9999: PRD_SUM(k).stat_hist = -9999
          If PRD_MAX(k).stat_hist = -9999 Then PRD_MAX(k).stat_hist = -9999: PRD_MAX(k).stat_hist_date = CDate("1899-09-09")
          If PRD_MIN(k).stat_hist = 999999 Then PRD_MIN(k).stat_hist = -9999: PRD_MIN(k).stat_hist_date = CDate("1899-09-09")
        ElseIf j = 3 Then
          
          If PRD_AVG(k).stat <> -9999 And PRD_AVG(k).stat_hist <> -9999 Then
            PRD_AVG(k).stat_diff1 = Format(PRD_AVG(k).stat, "0.0") - Format(PRD_AVG(k).stat_hist, "0.0")
          End If
          
          If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Or SelField_Initial = "S" Then
            If FlagStat = "SUM" Then
              If PRD_AVG(k).stat <> -9999 And PRD_AVG(k).stat_hist <> -9999 Then
                If PRD_AVG(k).stat_hist <> 0 Then PRD_AVG(k).stat_diff1 = 100 * (PRD_AVG(k).stat - PRD_AVG(k).stat_hist) / PRD_AVG(k).stat_hist
                If PRD_AVG(k).stat_hist = 0 Then PRD_AVG(k).stat_diff1 = -9999
              End If
            End If
          End If
          
          If PRD_MAX(k).stat_diff1 = -9999 Then PRD_MAX(k).stat_diff1 = -9999
          If PRD_MIN(k).stat_diff1 = 999999 Then PRD_MIN(k).stat_diff1 = -9999
        
        ElseIf j = 4 Then
          If PRD_NUM(k).stat_year <> 0 Then PRD_AVG(k).stat_year = Format(PRD_SUM(k).stat_year / PRD_NUM(k).stat_year, "0.0")
          If PRD_NUM(k).stat_year = 0 Then PRD_AVG(k).stat_year = -9999: PRD_SUM(k).stat_year = -9999
          If PRD_MAX(k).stat_year = -9999 Then PRD_MAX(k).stat_year = -9999: PRD_MAX(k).stat_year_date = CDate("1899-09-09")
          If PRD_MIN(k).stat_year = 999999 Then PRD_MIN(k).stat_year = -9999: PRD_MIN(k).stat_year_date = CDate("1899-09-09")
        ElseIf j = 5 Then
          
          If PRD_AVG(k).stat <> -9999 And PRD_AVG(k).stat_year <> -9999 Then
            PRD_AVG(k).stat_diff2 = Format(PRD_AVG(k).stat, "0.0") - Format(PRD_AVG(k).stat_year, "0.0")
          End If
          
          If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Or SelField_Initial = "S" Then
            If FlagStat = "SUM" Then
              If PRD_AVG(k).stat <> -9999 And PRD_AVG(k).stat_year <> -9999 Then
                If PRD_AVG(k).stat_year <> 0 Then PRD_AVG(k).stat_diff2 = 100 * (PRD_AVG(k).stat - PRD_AVG(k).stat_year) / PRD_AVG(k).stat_year
                If PRD_AVG(k).stat_year = 0 Then PRD_AVG(k).stat_diff2 = -9999
              End If
            End If
          End If
          
          If PRD_MAX(k).stat_diff2 = -9999 Then PRD_MAX(k).stat_diff2 = -9999
          If PRD_MIN(k).stat_diff2 = 999999 Then PRD_MIN(k).stat_diff2 = -9999
        
        End If
    Next j
  Next k

End Sub




Private Sub Accumulate_ZONE_ColdAir(Data_in() As ColdAir_Result, ByVal i As Integer, ByVal k As Integer)
  Dim j As Integer
  For j = 1 To 9 '各字段循环
    If j = 1 Then
      If Data_in(i).ilevel <> -9999 Then
        If Data_in(i).ilevel > CA_MAX(k).ilevel Then CA_MAX(k).ilevel = Data_in(i).ilevel
        If Data_in(i).ilevel < CA_MIN(k).ilevel Then CA_MIN(k).ilevel = Data_in(i).ilevel
      End If
    ElseIf j = 2 Then
      If Data_in(i).BDate <> CDate("1899-09-09") Then
        If Data_in(i).BDate > CA_MAX(k).BDate Then CA_MAX(k).BDate = Data_in(i).BDate
        If Data_in(i).BDate < CA_MIN(k).BDate Then CA_MIN(k).BDate = Data_in(i).BDate
      End If
    ElseIf j = 3 Then
      If Data_in(i).EDate <> CDate("1899-09-09") Then
        If Data_in(i).EDate > CA_MAX(k).EDate Then CA_MAX(k).EDate = Data_in(i).EDate
        If Data_in(i).EDate < CA_MIN(k).EDate Then CA_MIN(k).EDate = Data_in(i).EDate
      End If
    ElseIf j = 4 Then
      If Data_in(i).idays <> -9999 Then
        If Data_in(i).idays > CA_MAX(k).idays Then CA_MAX(k).idays = Data_in(i).idays
        If Data_in(i).idays < CA_MIN(k).idays Then CA_MIN(k).idays = Data_in(i).idays
        CA_SUM(k).idays = CA_SUM(k).idays + Data_in(i).idays
        CA_NUM(k).idays = CA_NUM(k).idays + 1
      End If
    ElseIf j = 5 Then
      If Data_in(i).tv_24hmax <> -9999 Then
        If Data_in(i).tv_24hmax > CA_MAX(k).tv_24hmax Then CA_MAX(k).tv_24hmax = Data_in(i).tv_24hmax
        If Data_in(i).tv_24hmax < CA_MIN(k).tv_24hmax Then CA_MIN(k).tv_24hmax = Data_in(i).tv_24hmax
        CA_SUM(k).tv_24hmax = CA_SUM(k).tv_24hmax + Data_in(i).tv_24hmax
        CA_NUM(k).tv_24hmax = CA_NUM(k).tv_24hmax + 1
      End If
    ElseIf j = 6 Then
      If Data_in(i).tv_48hmax <> -9999 Then
        If Data_in(i).tv_48hmax > CA_MAX(k).tv_48hmax Then CA_MAX(k).tv_48hmax = Data_in(i).tv_48hmax
        If Data_in(i).tv_48hmax < CA_MIN(k).tv_48hmax Then CA_MIN(k).tv_48hmax = Data_in(i).tv_48hmax
        CA_SUM(k).tv_48hmax = CA_SUM(k).tv_48hmax + Data_in(i).tv_48hmax
        CA_NUM(k).tv_48hmax = CA_NUM(k).tv_48hmax + 1
      End If
    ElseIf j = 7 Then
      If Data_in(i).tv_acc <> -9999 Then
        If Data_in(i).tv_acc > CA_MAX(k).tv_acc Then CA_MAX(k).tv_acc = Data_in(i).tv_acc
        If Data_in(i).tv_acc < CA_MIN(k).tv_acc Then CA_MIN(k).tv_acc = Data_in(i).tv_acc
        CA_SUM(k).tv_acc = CA_SUM(k).tv_acc + Data_in(i).tv_acc
        CA_NUM(k).tv_acc = CA_NUM(k).tv_acc + 1
      End If
    ElseIf j = 8 Then
      If Data_in(i).T <> -9999 Then
        If Data_in(i).T > CA_MAX(k).T Then CA_MAX(k).T = Data_in(i).T
        If Data_in(i).T < CA_MIN(k).T Then CA_MIN(k).T = Data_in(i).T
        CA_SUM(k).T = CA_SUM(k).T + Data_in(i).T
        CA_NUM(k).T = CA_NUM(k).T + 1
      End If
    ElseIf j = 9 Then
      If Data_in(i).t_min <> -9999 Then
        If Data_in(i).t_min > CA_MAX(k).t_min Then CA_MAX(k).t_min = Data_in(i).t_min
        If Data_in(i).t_min < CA_MIN(k).t_min Then CA_MIN(k).t_min = Data_in(i).t_min
        CA_SUM(k).t_min = CA_SUM(k).t_min + Data_in(i).t_min
        CA_NUM(k).t_min = CA_NUM(k).t_min + 1
      End If
    End If
  Next j
End Sub


'统计冷空气数据的区域平均/最大/最小值
Public Sub Stat_ZONE_ColdAir(Data_in() As ColdAir_Result)
  Dim iRows%
  Dim i%, j%, k%, l%
  Dim irows_CA% '冷空气记录的条数
  irows_CA = UBound(Data_in)
  
  If FlagZone = "CITY" Then iRows = CityNum
  If FlagZone = "CNTY" Then iRows = CountyNum
  If FlagZone = "TOWN" Then iRows = TownNum
  
  ReDim CA_SUM(1 To iRows): ReDim CA_NUM(1 To iRows): ReDim CA_AVG(1 To iRows): ReDim CA_MAX(1 To iRows): ReDim CA_MIN(1 To iRows)
  For k = 1 To iRows
    CA_SUM(k).ilevel = 0: CA_SUM(k).idays = 0: CA_SUM(k).tv_24hmax = 0: CA_SUM(k).tv_48hmax = 0: CA_SUM(k).tv_acc = 0: CA_SUM(k).T = 0: CA_SUM(k).t_min = 0
    CA_NUM(k).ilevel = 0: CA_NUM(k).idays = 0: CA_NUM(k).tv_24hmax = 0: CA_NUM(k).tv_48hmax = 0: CA_NUM(k).tv_acc = 0: CA_NUM(k).T = 0: CA_NUM(k).t_min = 0
    
    CA_AVG(k).ilevel = -9999: CA_AVG(k).BDate = CDate("1899-09-09"): CA_AVG(k).EDate = CDate("1899-09-09"):
    CA_AVG(k).idays = -9999: CA_AVG(k).tv_24hmax = -9999: CA_AVG(k).tv_48hmax = -9999: CA_AVG(k).tv_acc = -9999: CA_AVG(k).T = -9999: CA_AVG(k).t_min = -9999
    
    CA_MAX(k).ilevel = -9999: CA_MAX(k).BDate = CDate("1899-09-09"): CA_MAX(k).EDate = CDate("1899-09-09"):
    CA_MAX(k).idays = -9999: CA_MAX(k).tv_24hmax = -9999: CA_MAX(k).tv_48hmax = -9999: CA_MAX(k).tv_acc = -9999: CA_MAX(k).T = -9999: CA_MAX(k).t_min = -9999
    
    CA_MIN(k).ilevel = 999999: CA_MIN(k).BDate = CDate("2099-09-09"): CA_MIN(k).EDate = CDate("2099-09-09"):
    CA_MIN(k).idays = 999999: CA_MIN(k).tv_24hmax = 999999: CA_MIN(k).tv_48hmax = 999999: CA_MIN(k).tv_acc = 999999: CA_MIN(k).T = 999999: CA_MIN(k).t_min = 999999
  Next k
    
  '逐冷空气记录循环，统计各区域累计/最大/最小/计数值
  For i = 1 To irows_CA
    For k = 1 To iRows
      If FlagZone = "CITY" Then
        If Data_in(i).City = CityInfo2(k).City Then
          Call Accumulate_ZONE_ColdAir(Data_in, i, k)
          Exit For
        End If
      
      ElseIf FlagZone = "CNTY" Then
        If Data_in(i).City = CountyInfo2(k).City And Data_in(i).County = CountyInfo2(k).County Then
          Call Accumulate_ZONE_ColdAir(Data_in, i, k)
          Exit For
        End If
      
      ElseIf FlagZone = "TOWN" Then
        If Data_in(i).City = TownInfo2(k).City And Data_in(i).County = TownInfo2(k).County And Data_in(i).Town = TownInfo2(k).Town Then
          Call Accumulate_ZONE_ColdAir(Data_in, i, k)
          Exit For
        End If
      End If
      
    Next k
  Next i
      
  '计算各区域平均值，并将空值赋值为-9999
  For k = 1 To iRows
    For j = 1 To 9
        If j = 1 Then
          CA_AVG(k).ilevel = -9999: CA_SUM(k).ilevel = -9999
          If CA_MAX(k).ilevel = -9999 Then CA_MAX(k).ilevel = -9999
          If CA_MIN(k).ilevel = 999999 Then CA_MIN(k).ilevel = -9999
        ElseIf j = 2 Then
          CA_AVG(k).BDate = CDate("1899-09-09"): CA_SUM(k).BDate = CDate("1899-09-09")
          If CA_MAX(k).BDate = CDate("1899-09-09") Then CA_MAX(k).BDate = CDate("1899-09-09")
          If CA_MIN(k).BDate = CDate("2099-09-09") Then CA_MIN(k).BDate = CDate("1899-09-09")
        ElseIf j = 3 Then
          CA_AVG(k).EDate = CDate("1899-09-09"): CA_SUM(k).EDate = CDate("1899-09-09")
          If CA_MAX(k).EDate = CDate("1899-09-09") Then CA_MAX(k).EDate = CDate("1899-09-09")
          If CA_MIN(k).EDate = CDate("2099-09-09") Then CA_MIN(k).EDate = CDate("1899-09-09")
        ElseIf j = 4 Then
'          If CA_NUM(k).iDays <> 0 Then CA_AVG(k).iDays = Round(CA_SUM(k).iDays / CA_NUM(k).iDays, 1)
          If CA_NUM(k).idays <> 0 Then CA_AVG(k).idays = Format(CA_SUM(k).idays / CA_NUM(k).idays, "0.0")
          If CA_NUM(k).idays = 0 Then CA_AVG(k).idays = -9999: CA_SUM(k).idays = -9999
          If CA_MAX(k).idays = -9999 Then CA_MAX(k).idays = -9999
          If CA_MIN(k).idays = 999999 Then CA_MIN(k).idays = -9999
        ElseIf j = 5 Then
'          If CA_NUM(k).tv_24hmax <> 0 Then CA_AVG(k).tv_24hmax = Round(CA_SUM(k).tv_24hmax / CA_NUM(k).tv_24hmax, 1)
          If CA_NUM(k).tv_24hmax <> 0 Then CA_AVG(k).tv_24hmax = Format(CA_SUM(k).tv_24hmax / CA_NUM(k).tv_24hmax, "0.0")
          If CA_NUM(k).tv_24hmax = 0 Then CA_AVG(k).tv_24hmax = -9999: CA_SUM(k).tv_24hmax = -9999
          If CA_MAX(k).tv_24hmax = -9999 Then CA_MAX(k).tv_24hmax = -9999
          If CA_MIN(k).tv_24hmax = 999999 Then CA_MIN(k).tv_24hmax = -9999
        ElseIf j = 6 Then
'          If CA_NUM(k).tv_48hmax <> 0 Then CA_AVG(k).tv_48hmax = Round(CA_SUM(k).tv_48hmax / CA_NUM(k).tv_48hmax, 1)
          If CA_NUM(k).tv_48hmax <> 0 Then CA_AVG(k).tv_48hmax = Format(CA_SUM(k).tv_48hmax / CA_NUM(k).tv_48hmax, "0.0")
          If CA_NUM(k).tv_48hmax = 0 Then CA_AVG(k).tv_48hmax = -9999: CA_SUM(k).tv_48hmax = -9999
          If CA_MAX(k).tv_48hmax = -9999 Then CA_MAX(k).tv_48hmax = -9999
          If CA_MIN(k).tv_48hmax = 999999 Then CA_MIN(k).tv_48hmax = -9999
        ElseIf j = 7 Then
'          If CA_NUM(k).tv_acc <> 0 Then CA_AVG(k).tv_acc = Round(CA_SUM(k).tv_acc / CA_NUM(k).tv_acc, 1)
          If CA_NUM(k).tv_acc <> 0 Then CA_AVG(k).tv_acc = Format(CA_SUM(k).tv_acc / CA_NUM(k).tv_acc, "0.0")
          If CA_NUM(k).tv_acc = 0 Then CA_AVG(k).tv_acc = -9999: CA_SUM(k).tv_acc = -9999
          If CA_MAX(k).tv_acc = -9999 Then CA_MAX(k).tv_acc = -9999
          If CA_MIN(k).tv_acc = 999999 Then CA_MIN(k).tv_acc = -9999
        ElseIf j = 8 Then
'          If CA_NUM(k).t <> 0 Then CA_AVG(k).t = Round(CA_SUM(k).t / CA_NUM(k).t, 1)
          If CA_NUM(k).T <> 0 Then CA_AVG(k).T = Format(CA_SUM(k).T / CA_NUM(k).T, "0.0")
          If CA_NUM(k).T = 0 Then CA_AVG(k).T = -9999: CA_SUM(k).T = -9999
          If CA_MAX(k).T = -9999 Then CA_MAX(k).T = -9999
          If CA_MIN(k).T = 999999 Then CA_MIN(k).T = -9999
        ElseIf j = 9 Then
'          If CA_NUM(k).t_min <> 0 Then CA_AVG(k).t_min = Round(CA_SUM(k).t_min / CA_NUM(k).t_min, 1)
          If CA_NUM(k).t_min <> 0 Then CA_AVG(k).t_min = Format(CA_SUM(k).t_min / CA_NUM(k).t_min, "0.0")
          If CA_NUM(k).t_min = 0 Then CA_AVG(k).t_min = -9999: CA_SUM(k).t_min = -9999
          If CA_MAX(k).t_min = -9999 Then CA_MAX(k).t_min = -9999
          If CA_MIN(k).t_min = 999999 Then CA_MIN(k).t_min = -9999
        End If
    Next j
  Next k

End Sub



Private Sub Accumulate_ZONE_Seasons(Data_in() As Season_Result, ByVal i As Integer, ByVal k As Integer)
  Dim j As Integer
  For j = 1 To 8 '各指标循环
    If j = 1 Then
      If Data_in(i).OnsetNorm <> CDate("1899-09-09") Then
        If Data_in(i).OnsetNorm > SS_MAX(k).OnsetNorm Then SS_MAX(k).OnsetNorm = Data_in(i).OnsetNorm
        If Data_in(i).OnsetNorm < SS_MIN(k).OnsetNorm Then SS_MIN(k).OnsetNorm = Data_in(i).OnsetNorm
      End If
    ElseIf j = 2 Then
      If Data_in(i).OnsetCurr <> CDate("1899-09-09") Then
        If Data_in(i).OnsetCurr > SS_MAX(k).OnsetCurr Then SS_MAX(k).OnsetCurr = Data_in(i).OnsetCurr
        If Data_in(i).OnsetCurr < SS_MIN(k).OnsetCurr Then SS_MIN(k).OnsetCurr = Data_in(i).OnsetCurr
      End If
    ElseIf j = 3 Then
      If Data_in(i).OnsetComp <> CDate("1899-09-09") Then
        If Data_in(i).OnsetComp > SS_MAX(k).OnsetComp Then SS_MAX(k).OnsetComp = Data_in(i).OnsetComp
        If Data_in(i).OnsetComp < SS_MIN(k).OnsetComp Then SS_MIN(k).OnsetComp = Data_in(i).OnsetComp
      End If
    ElseIf j = 4 Then
      If Data_in(i).T_IdxNorm <> -9999 Then
        If Data_in(i).T_IdxNorm > SS_MAX(k).T_IdxNorm Then SS_MAX(k).T_IdxNorm = Data_in(i).T_IdxNorm
        If Data_in(i).T_IdxNorm < SS_MIN(k).T_IdxNorm Then SS_MIN(k).T_IdxNorm = Data_in(i).T_IdxNorm
        SS_SUM(k).T_IdxNorm = SS_SUM(k).T_IdxNorm + Data_in(i).T_IdxNorm
        SS_NUM(k).T_IdxNorm = SS_NUM(k).T_IdxNorm + 1
      End If
    ElseIf j = 5 Then
      If Data_in(i).T_IdxCurr <> -9999 Then
        If Data_in(i).T_IdxCurr > SS_MAX(k).T_IdxCurr Then SS_MAX(k).T_IdxCurr = Data_in(i).T_IdxCurr
        If Data_in(i).T_IdxCurr < SS_MIN(k).T_IdxCurr Then SS_MIN(k).T_IdxCurr = Data_in(i).T_IdxCurr
        SS_SUM(k).T_IdxCurr = SS_SUM(k).T_IdxCurr + Data_in(i).T_IdxCurr
        SS_NUM(k).T_IdxCurr = SS_NUM(k).T_IdxCurr + 1
      End If
    ElseIf j = 6 Then
      If Data_in(i).T_IdxComp <> -9999 Then
        If Data_in(i).T_IdxComp > SS_MAX(k).T_IdxComp Then SS_MAX(k).T_IdxComp = Data_in(i).T_IdxComp
        If Data_in(i).T_IdxComp < SS_MIN(k).T_IdxComp Then SS_MIN(k).T_IdxComp = Data_in(i).T_IdxComp
        SS_SUM(k).T_IdxComp = SS_SUM(k).T_IdxComp + Data_in(i).T_IdxComp
        SS_NUM(k).T_IdxComp = SS_NUM(k).T_IdxComp + 1
      End If
    ElseIf j = 7 Then
      If Data_in(i).diffNorm <> -9999 Then
        If Data_in(i).diffNorm > SS_MAX(k).diffNorm Then SS_MAX(k).diffNorm = Data_in(i).diffNorm
        If Data_in(i).diffNorm < SS_MIN(k).diffNorm Then SS_MIN(k).diffNorm = Data_in(i).diffNorm
        SS_SUM(k).diffNorm = SS_SUM(k).diffNorm + Data_in(i).diffNorm
        SS_NUM(k).diffNorm = SS_NUM(k).diffNorm + 1
      End If
    ElseIf j = 8 Then
      If Data_in(i).diffComp <> -9999 Then
        If Data_in(i).diffComp > SS_MAX(k).diffComp Then SS_MAX(k).diffComp = Data_in(i).diffComp
        If Data_in(i).diffComp < SS_MIN(k).diffComp Then SS_MIN(k).diffComp = Data_in(i).diffComp
        SS_SUM(k).diffComp = SS_SUM(k).diffComp + Data_in(i).diffComp
        SS_NUM(k).diffComp = SS_NUM(k).diffComp + 1
      End If
    End If
  Next j
End Sub


'统计季节划分数据的区域平均/最大/最小值
Public Sub Stat_ZONE_Seasons(Data_in() As Season_Result, BDateNorm As Date, BDate1 As Date, BDate2 As Date)
  Dim iRows%
  Dim i%, j%, k%, l%
  Dim irows_Season% '季节划分记录的条数
  irows_Season = UBound(Data_in)
  
  If FlagZone = "CITY" Then iRows = CityNum
  If FlagZone = "CNTY" Then iRows = CountyNum
  If FlagZone = "TOWN" Then iRows = TownNum
    
  ReDim SS_SUM(1 To iRows): ReDim SS_NUM(1 To iRows): ReDim SS_AVG(1 To iRows): ReDim SS_MAX(1 To iRows): ReDim SS_MIN(1 To iRows)
  For k = 1 To iRows
    SS_SUM(k).T_IdxNorm = 0: SS_SUM(k).T_IdxCurr = 0: SS_SUM(k).T_IdxComp = 0: SS_SUM(k).diffNorm = 0: SS_SUM(k).diffComp = 0
    SS_NUM(k).T_IdxNorm = 0: SS_NUM(k).T_IdxCurr = 0: SS_NUM(k).T_IdxComp = 0: SS_NUM(k).diffNorm = 0: SS_NUM(k).diffComp = 0
    
    SS_AVG(k).OnsetNorm = CDate("1899-09-09"): SS_AVG(k).OnsetCurr = CDate("1899-09-09"): SS_AVG(k).OnsetComp = CDate("1899-09-09")
    SS_AVG(k).T_IdxNorm = -9999: SS_AVG(k).T_IdxCurr = -9999: SS_AVG(k).T_IdxComp = -9999: SS_AVG(k).diffNorm = -9999: SS_AVG(k).diffComp = -9999
    
    SS_MAX(k).OnsetNorm = CDate("1899-09-09"): SS_MAX(k).OnsetCurr = CDate("1899-09-09"): SS_MAX(k).OnsetComp = CDate("1899-09-09")
    SS_MAX(k).T_IdxNorm = -9999: SS_MAX(k).T_IdxCurr = -9999: SS_MAX(k).T_IdxComp = -9999: SS_MAX(k).diffNorm = -9999: SS_MAX(k).diffComp = -9999
    
    SS_MIN(k).OnsetNorm = CDate("2099-09-09"): SS_MIN(k).OnsetCurr = CDate("2099-09-09"): SS_MIN(k).OnsetComp = CDate("2099-09-09")
    SS_MIN(k).T_IdxNorm = 999999: SS_MIN(k).T_IdxCurr = 999999: SS_MIN(k).T_IdxComp = 999999: SS_MIN(k).diffNorm = 999999: SS_MIN(k).diffComp = 999999
  Next k
    
  '逐季节划分记录循环，统计各区域累计/最大/最小/计数值
  For i = 1 To irows_Season
    For k = 1 To iRows
      If FlagZone = "CITY" Then
        If Data_in(i).City = CityInfo2(k).City Then
          Call Accumulate_ZONE_Seasons(Data_in, i, k)
          Exit For
        End If
      
      ElseIf FlagZone = "CNTY" Then
        If Data_in(i).City = CountyInfo2(k).City And Data_in(i).County = CountyInfo2(k).County Then
          Call Accumulate_ZONE_Seasons(Data_in, i, k)
          Exit For
        End If
      
      ElseIf FlagZone = "TOWN" Then
        If Data_in(i).City = TownInfo2(k).City And Data_in(i).County = TownInfo2(k).County And Data_in(i).Town = TownInfo2(k).Town Then
          Call Accumulate_ZONE_Seasons(Data_in, i, k)
          Exit For
        End If
      End If
      
    Next k
  Next i
      
  '计算各区域平均值，并将空值赋值为-9999
  For k = 1 To iRows
    For j = 1 To 8
        If j = 1 Then
          If SS_NUM(k).T_IdxNorm <> 0 Then SS_AVG(k).T_IdxNorm = Format(SS_SUM(k).T_IdxNorm / SS_NUM(k).T_IdxNorm, "0.0")
          If SS_NUM(k).T_IdxNorm = 0 Then SS_AVG(k).T_IdxNorm = -9999: SS_SUM(k).T_IdxNorm = -9999
          If SS_MAX(k).T_IdxNorm = -9999 Then SS_MAX(k).T_IdxNorm = -9999
          If SS_MIN(k).T_IdxNorm = 999999 Then SS_MIN(k).T_IdxNorm = -9999
        ElseIf j = 2 Then
          If SS_NUM(k).T_IdxCurr <> 0 Then SS_AVG(k).T_IdxCurr = Format(SS_SUM(k).T_IdxCurr / SS_NUM(k).T_IdxCurr, "0.0")
          If SS_NUM(k).T_IdxCurr = 0 Then SS_AVG(k).T_IdxCurr = -9999: SS_SUM(k).T_IdxCurr = -9999
          If SS_MAX(k).T_IdxCurr = -9999 Then SS_MAX(k).T_IdxCurr = -9999
          If SS_MIN(k).T_IdxCurr = 999999 Then SS_MIN(k).T_IdxCurr = -9999
        ElseIf j = 3 Then
          If SS_NUM(k).T_IdxComp <> 0 Then SS_AVG(k).T_IdxComp = Format(SS_SUM(k).T_IdxComp / SS_NUM(k).T_IdxComp, "0.0")
          If SS_NUM(k).T_IdxComp = 0 Then SS_AVG(k).T_IdxComp = -9999: SS_SUM(k).T_IdxComp = -9999
          If SS_MAX(k).T_IdxComp = -9999 Then SS_MAX(k).T_IdxComp = -9999
          If SS_MIN(k).T_IdxComp = 999999 Then SS_MIN(k).T_IdxComp = -9999
        ElseIf j = 4 Then
          If SS_NUM(k).diffNorm <> 0 Then SS_AVG(k).diffNorm = Format(SS_SUM(k).diffNorm / SS_NUM(k).diffNorm, "0.0")
          If SS_NUM(k).diffNorm = 0 Then SS_AVG(k).diffNorm = -9999: SS_SUM(k).diffNorm = -9999
          If SS_MAX(k).diffNorm = -9999 Then SS_MAX(k).diffNorm = -9999
          If SS_MIN(k).diffNorm = 999999 Then SS_MIN(k).diffNorm = -9999
        ElseIf j = 5 Then
          If SS_NUM(k).diffComp <> 0 Then SS_AVG(k).diffComp = Format(SS_SUM(k).diffComp / SS_NUM(k).diffComp, "0.0")
          If SS_NUM(k).diffComp = 0 Then SS_AVG(k).diffComp = -9999: SS_SUM(k).diffComp = -9999
          If SS_MAX(k).diffComp = -9999 Then SS_MAX(k).diffComp = -9999
          If SS_MIN(k).diffComp = 999999 Then SS_MIN(k).diffComp = -9999
        ElseIf j = 6 Then
          If SS_AVG(k).T_IdxNorm <> -9999 Then SS_AVG(k).OnsetNorm = BDateNorm - 1 + SS_AVG(k).T_IdxNorm
          If SS_SUM(k).T_IdxNorm <> -9999 Then SS_SUM(k).OnsetNorm = BDateNorm - 1 + SS_SUM(k).T_IdxNorm
          If SS_MAX(k).OnsetNorm = CDate("1899-09-09") Then SS_MAX(k).OnsetNorm = CDate("1899-09-09")
          If SS_MIN(k).OnsetNorm = CDate("2099-09-09") Then SS_MIN(k).OnsetNorm = CDate("1899-09-09")
        ElseIf j = 7 Then
          If SS_AVG(k).T_IdxCurr <> -9999 Then SS_AVG(k).OnsetCurr = BDate1 - 1 + SS_AVG(k).T_IdxCurr
          If SS_SUM(k).T_IdxCurr <> -9999 Then SS_SUM(k).OnsetCurr = BDate1 - 1 + SS_SUM(k).T_IdxCurr
          If SS_MAX(k).OnsetCurr = CDate("1899-09-09") Then SS_MAX(k).OnsetCurr = CDate("1899-09-09")
          If SS_MIN(k).OnsetCurr = CDate("2099-09-09") Then SS_MIN(k).OnsetCurr = CDate("1899-09-09")
        ElseIf j = 8 Then
          If SS_AVG(k).T_IdxComp <> -9999 Then SS_AVG(k).OnsetComp = BDate2 - 1 + SS_AVG(k).T_IdxComp
          If SS_SUM(k).T_IdxComp <> -9999 Then SS_SUM(k).OnsetComp = BDate2 - 1 + SS_SUM(k).T_IdxComp
          If SS_MAX(k).OnsetComp = CDate("1899-09-09") Then SS_MAX(k).OnsetComp = CDate("1899-09-09")
          If SS_MIN(k).OnsetComp = CDate("2099-09-09") Then SS_MIN(k).OnsetComp = CDate("1899-09-09")
        
        
        End If
    Next j
  Next k

End Sub



'统计多行（即多站点）平均、最大、最小值
Public Sub Stat_ROWS(DataIn!(), iCols%, iRows%)
  Dim sngSum!, sngAvg!, sngMax!, sngMin!, iSta% '当年值的统计量
  Dim iNum% '用于计算多站平均的站点数
    ReDim ROWS_Stat(3, iCols)
    For j = 1 To iCols
      sngSum = 0: iNum = 0: sngAvg = -9999: sngMax = -9999: sngMin = 999999
      For i = 1 To iRows
        If DataIn(i, j) <> -9999 Then
          If DataIn(i, j) > sngMax Then sngMax = DataIn(i, j)
          If DataIn(i, j) < sngMin Then sngMin = DataIn(i, j)
          sngSum = sngSum + DataIn(i, j)
          iNum = iNum + 1
        End If
      Next i
'      If iNum <> 0 Then sngAvg = Round(sngSum / iNum, 1)
      If iNum <> 0 Then sngAvg = Format(sngSum / iNum, "0.0")
      If iNum = 0 Then sngAvg = -9999: sngSum = -9999
      If sngMax = -9999 Then sngMax = -9999
      If sngMin = 999999 Then sngMin = -9999
      ROWS_Stat(1, j) = sngAvg: ROWS_Stat(2, j) = sngMax: ROWS_Stat(3, j) = sngMin
    Next j
End Sub



'统计多列（即多日）平均/累计、最大、最小值
Public Sub Stat_COLS(DataIn!(), iCols%, iRows%, stat$)
  Dim sngSum!, sngAvg!, sngMax!, sngMin!, iSta% '当年值的统计量
  Dim iNum% '用于计算多站平均的站点数
    ReDim COLS_Stat(iRows, 3)
    For i = 1 To iRows
      sngSum = 0: iNum = 0: sngAvg = -9999: sngMax = -9999: sngMin = 999999
      For j = 1 To iCols
        If DataIn(i, j) <> -9999 Then
          If DataIn(i, j) > sngMax Then sngMax = DataIn(i, j)
          If DataIn(i, j) < sngMin Then sngMin = DataIn(i, j)
          sngSum = sngSum + DataIn(i, j)
          iNum = iNum + 1
        End If
      Next j
      'If iNum <> 0 Then sngAvg = Round(sngSum / iNum, 1) '四舍五入结果错误
      If iNum <> 0 Then sngAvg = Format(sngSum / iNum, "0.0") '四舍五入结果正确
      If iNum = 0 Then sngAvg = -9999: sngSum = -9999
      If sngMax = -9999 Then sngMax = -9999
      If sngMin = 999999 Then sngMin = -9999
      If stat = "AVE" Then COLS_Stat(i, 1) = sngAvg
      If stat = "SUM" Then COLS_Stat(i, 1) = sngSum
      COLS_Stat(i, 2) = sngMax: COLS_Stat(i, 3) = sngMin
    Next i '
End Sub



'统计多列统计结果（即多站点）的平均、最大、最小值
Public Sub Stat_COLS_ROWS(iRows%)
  Dim sngSum!, sngAvg!, sngMax!, sngMin!, iSta% '当年值的统计量
  Dim iNum% '用于计算多站平均的站点数
    ReDim COLS_ROWS_Stat(3, 3)
    For j = 1 To 3
      sngSum = 0: iNum = 0: sngAvg = -9999: sngMax = -9999: sngMin = 999999
      For i = 1 To iRows
        If COLS_Stat(i, j) <> -9999 Then
          If COLS_Stat(i, j) > sngMax Then sngMax = COLS_Stat(i, j)
          If COLS_Stat(i, j) < sngMin Then sngMin = COLS_Stat(i, j)
          sngSum = sngSum + COLS_Stat(i, j)
          iNum = iNum + 1
        End If
      Next i
'      If iNum <> 0 Then sngAvg = Round(sngSum / iNum, 1)
      If iNum <> 0 Then sngAvg = Format(sngSum / iNum, "0.0")
      If iNum = 0 Then sngAvg = -9999: sngSum = -9999
      If sngMax = -9999 Then sngMax = -9999
      If sngMin = 999999 Then sngMin = -9999
      COLS_ROWS_Stat(1, j) = sngAvg: COLS_ROWS_Stat(2, j) = sngMax: COLS_ROWS_Stat(3, j) = sngMin
    Next j
End Sub




'统计- 任意时段统计结果 -多行（即多站点）平均、最大、最小值
Public Sub Stat_ROWS_Period(DataIn() As Period_Result, iRows%, FlagStat$, SelField_Initial$)
  Dim sngSum!, sngAvg!, sngMax!, sngMin!, iSta% '当年值的统计量
  Dim iNum% '用于计算多站平均的站点数
  Dim tmpData!
  Dim tmpDate '2024-9-18：添加上表最大值/最小值对应的日期
  Dim dateMax As Date, dateMin As Date '2024-9-18：添加下表极大/极小对应的日期
  
  ReDim ROWS_Stat_Period(1 To 3) '1、2、3下标分别对应多站点平均、最大、最小值
  For i = 1 To 3
    ROWS_Stat_Period(i).stat = -9999 '当年统计值
    ROWS_Stat_Period(i).stat_date = CDate("1899-09-09") '极值出现日期
    
    ROWS_Stat_Period(i).stat_hist = -9999 '累年值
    ROWS_Stat_Period(i).stat_hist_date = CDate("1899-09-09") '累年极值出现日期
    ROWS_Stat_Period(i).stat_diff1 = -9999 '多年平均距平
    
    ROWS_Stat_Period(i).stat_year = -9999 '某年值
    ROWS_Stat_Period(i).stat_year_date = CDate("1899-09-09") '某年极值出现日期
    ROWS_Stat_Period(i).stat_diff2 = -9999  '某年距平
    
    ROWS_Stat_Period(i).rank1 = -9999 '排名（大-小）
    ROWS_Stat_Period(i).rank2 = -9999 '排名（小-大）
    ROWS_Stat_Period(i).rank_years = "-9999" '排名年份
    ROWS_Stat_Period(i).maxyear = -9999 '最大值年
    ROWS_Stat_Period(i).maxvalue = -9999 '最大值
    ROWS_Stat_Period(i).minyear = -9999 '最小值年
    ROWS_Stat_Period(i).minvalue = -9999 '最小值
  Next i
  
  
  '循环非日期值的5列
  For j = 1 To 5
    sngSum = 0: iNum = 0: sngAvg = -9999: sngMax = -9999: sngMin = 999999
      
    For i = 1 To iRows
      If j = 1 Then tmpData = DataIn(i).stat: tmpDate = DataIn(i).stat_date
      If j = 2 Then tmpData = DataIn(i).stat_hist: tmpDate = DataIn(i).stat_hist_date
      If j = 3 Then tmpData = DataIn(i).stat_diff1
      If j = 4 Then tmpData = DataIn(i).stat_year: tmpDate = DataIn(i).stat_year_date
      If j = 5 Then tmpData = DataIn(i).stat_diff2
      If tmpData <> -9999 Then
        If tmpData > sngMax Then sngMax = tmpData: dateMax = tmpDate
        If tmpData < sngMin Then sngMin = tmpData: dateMin = tmpDate
        sngSum = sngSum + tmpData
        iNum = iNum + 1
      End If
    Next i
      
'    If iNum <> 0 Then sngAvg = Round(sngSum / iNum, 1)
    If iNum <> 0 Then sngAvg = Format(sngSum / iNum, "0.0")
    If iNum = 0 Then sngAvg = -9999: sngSum = -9999
    If sngMax = -9999 Then sngMax = -9999: dateMax = CDate("1899-09-09")
    If sngMin = 999999 Then sngMin = -9999: dateMin = CDate("1899-09-09")
    
    If j = 1 Then
      ROWS_Stat_Period(1).stat = sngAvg
      ROWS_Stat_Period(2).stat = sngMax: ROWS_Stat_Period(2).stat_date = dateMax
      ROWS_Stat_Period(3).stat = sngMin: ROWS_Stat_Period(3).stat_date = dateMin
    ElseIf j = 2 Then
      ROWS_Stat_Period(1).stat_hist = sngAvg
      ROWS_Stat_Period(2).stat_hist = sngMax: ROWS_Stat_Period(2).stat_hist_date = dateMax
      ROWS_Stat_Period(3).stat_hist = sngMin: ROWS_Stat_Period(3).stat_hist_date = dateMin
      
    ElseIf j = 3 Then '计算地区平均值的距平 = 当年值 - 多年平均值
      
      If ROWS_Stat_Period(1).stat <> -9999 And ROWS_Stat_Period(1).stat_hist <> -9999 Then
'        ROWS_Stat_Period(1).stat_diff1 = Round(ROWS_Stat_Period(1).stat, 1) - Round(ROWS_Stat_Period(1).stat_hist, 1)
        ROWS_Stat_Period(1).stat_diff1 = Format(ROWS_Stat_Period(1).stat, "0.0") - Format(ROWS_Stat_Period(1).stat_hist, "0.0")
      End If
      If FlagStat = "SUM" Then
        If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Or SelField_Initial = "S" Then
          If ROWS_Stat_Period(1).stat <> -9999 And ROWS_Stat_Period(1).stat_hist <> -9999 Then
            If ROWS_Stat_Period(1).stat_hist <> 0 Then ROWS_Stat_Period(1).stat_diff1 = 100 * (ROWS_Stat_Period(1).stat - ROWS_Stat_Period(1).stat_hist) / ROWS_Stat_Period(1).stat_hist
            If ROWS_Stat_Period(1).stat_hist = 0 Then ROWS_Stat_Period(1).stat_diff1 = -9999
          End If
        End If
      End If
      ROWS_Stat_Period(2).stat_diff1 = sngMax: ROWS_Stat_Period(3).stat_diff1 = sngMin
      
    ElseIf j = 4 Then
      ROWS_Stat_Period(1).stat_year = sngAvg
      ROWS_Stat_Period(2).stat_year = sngMax: ROWS_Stat_Period(2).stat_year_date = dateMax
      ROWS_Stat_Period(3).stat_year = sngMin: ROWS_Stat_Period(3).stat_year_date = dateMin
    ElseIf j = 5 Then '计算地区平均值的距平 = 当年值 - 某年值
      
      If ROWS_Stat_Period(1).stat <> -9999 And ROWS_Stat_Period(1).stat_year <> -9999 Then
'        ROWS_Stat_Period(1).stat_diff2 = Round(ROWS_Stat_Period(1).stat, 1) - Round(ROWS_Stat_Period(1).stat_year, 1)
        ROWS_Stat_Period(1).stat_diff2 = Format(ROWS_Stat_Period(1).stat, "0.0") - Format(ROWS_Stat_Period(1).stat_year, "0.0")
      End If
      If FlagStat = "SUM" Then
        If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Or SelField_Initial = "S" Then
          If ROWS_Stat_Period(1).stat <> -9999 And ROWS_Stat_Period(1).stat_year <> -9999 Then
            If ROWS_Stat_Period(1).stat_year <> 0 Then ROWS_Stat_Period(1).stat_diff2 = 100 * (ROWS_Stat_Period(1).stat - ROWS_Stat_Period(1).stat_year) / ROWS_Stat_Period(1).stat_year
            If ROWS_Stat_Period(1).stat_year = 0 Then ROWS_Stat_Period(1).stat_diff2 = -9999
          End If
        End If
      End If
      ROWS_Stat_Period(2).stat_diff2 = sngMax: ROWS_Stat_Period(3).stat_diff2 = sngMin
    End If
  
  
  Next j
  
  '统计历年最大值/最小值 的 极大/极小值 （2026-02-26）
  For j = 1 To 2
    sngSum = 0: iNum = 0: sngAvg = -9999: sngMax = -9999: sngMin = 999999
    
    For i = 1 To iRows
      If j = 1 Then tmpData = DataIn(i).maxvalue
      If j = 2 Then tmpData = DataIn(i).minvalue

      
      If tmpData <> -9999 Then
        If tmpData > sngMax Then sngMax = tmpData
        If tmpData < sngMin Then sngMin = tmpData
        sngSum = sngSum + tmpData
        iNum = iNum + 1
      End If
    Next i

    If sngMax = -9999 Then sngMax = -9999
    If sngMin = 999999 Then sngMin = -9999
    
    If j = 1 Then
      ROWS_Stat_Period(2).maxvalue = sngMax
      ROWS_Stat_Period(3).maxvalue = sngMin
    ElseIf j = 2 Then
      ROWS_Stat_Period(2).minvalue = sngMax
      ROWS_Stat_Period(3).minvalue = sngMin
    End If
  Next j
  
End Sub



'统计- 任意时段统计各站点历年值AAA - 历年多行（即多站点）平均；'result1-AAA的多年平均值、result2-BBB的多年平均值
Public Sub Stat_ROWS_AAA(DataIn1!(), Result1!(), DataIn2!(), Result2!)
  Dim Result1_sum!(), Result2_sum!, iNum%
  Dim BYear_in%, EYear_in% 'DataIn1的起止年份
  
  BYear_in = LBound(DataIn1, 2): EYear_in = UBound(DataIn1, 2)
  
  ReDim Result1(BYear_in To EYear_in): ReDim Result1_sum(BYear_in To EYear_in)

    Result2 = -9999: Result2_sum = 0: iNum = 0
    For i = 1 To StaNum '当年的多站点平均值
      If DataIn2(i) <> -9999 Then Result2_sum = Result2_sum + DataIn2(i): iNum = iNum + 1
    Next i
'    If iNum <> 0 Then Result2 = Round(Result2_sum / iNum, 1)
    If iNum <> 0 Then Result2 = Format(Result2_sum / iNum, "0.0")

    For j = BYear_in To EYear_in '从第一年开始统计每一年的平均值
      Result1(j) = -9999: Result1_sum(j) = 0: iNum = 0
      For i = 1 To StaNum
        If DataIn1(i, j) <> -9999 Then Result1_sum(j) = Result1_sum(j) + DataIn1(i, j): iNum = iNum + 1
      Next i
'      If iNum <> 0 Then Result1(j) = Round(Result1_sum(j) / iNum, 1)
      If iNum <> 0 Then Result1(j) = Format(Result1_sum(j) / iNum, "0.0")
    Next j

End Sub


'统计- 任意时段统计各站点历年值AAA - 历年多行（即多站点）平均；'result1-AAA的多年平均值、result2-BBB的多年平均值
Public Sub Stat_ROWS_AAAExt(DataIn1!(), Result1!(), DataIn2!(), Result2!)
  Dim Result1_sum!(), Result2_sum!, iNum%
  Dim BYear_in%, EYear_in% 'DataIn1的起止年份
  
  BYear_in = LBound(DataIn1, 2): EYear_in = UBound(DataIn1, 2)
  
  ReDim Result1(BYear_in To EYear_in): ReDim Result1_sum(BYear_in To EYear_in)

    Result2 = -9999: Result2_sum = 0: iNum = 0
    For i = 1 To ExtStaNum '当年的多站点平均值
      If DataIn2(i) <> -9999 Then Result2_sum = Result2_sum + DataIn2(i): iNum = iNum + 1
    Next i
    If iNum <> 0 Then Result2 = Format(Result2_sum / iNum, "0.0")

    For j = BYear_in To EYear_in '从第一年开始统计每一年的平均值
      Result1(j) = -9999: Result1_sum(j) = 0: iNum = 0
      For i = 1 To ExtStaNum
        If DataIn1(i, j) <> -9999 Then Result1_sum(j) = Result1_sum(j) + DataIn1(i, j): iNum = iNum + 1
      Next i
      If iNum <> 0 Then Result1(j) = Format(Result1_sum(j) / iNum, "0.0")
    Next j

End Sub



Private Sub Accumulate_ZONE_AAA(AAA!(), ByRef AAA_ZONE_sum!(), ByRef AAA_ZONE_num%(), _
                                 BBB!(), ByRef BBB_ZONE_sum!(), ByRef BBB_ZONE_num%(), _
                                 ByVal i As Integer, ByVal k As Integer, ByVal BYear As Integer, ByVal EYear As Integer)
  Dim j As Integer
  If BBB(i) <> -9999 Then
    BBB_ZONE_sum(k) = BBB_ZONE_sum(k) + BBB(i)
    BBB_ZONE_num(k) = BBB_ZONE_num(k) + 1
  End If
  For j = BYear To EYear
    If AAA(i, j) <> -9999 Then
      AAA_ZONE_sum(k, j) = AAA_ZONE_sum(k, j) + AAA(i, j)
      AAA_ZONE_num(k, j) = AAA_ZONE_num(k, j) + 1
    End If
  Next j
End Sub


'统计- 任意时段统计各站点历年值AAA - 历年各区域（即市县镇）平均
Public Sub Stat_ZONE_AAA(AAA!(), AAA_ZONE!(), BBB!(), BBB_ZONE!())
  Dim AAA_ZONE_sum!(), AAA_ZONE_num%(), BBB_ZONE_sum!(), BBB_ZONE_num%()
  Dim BYear%, EYear% 'AAA的起止年份

  Dim iRows%
  Dim i%, j%, k%, l%

  If FlagZone = "CITY" Then iRows = CityNum
  If FlagZone = "CNTY" Then iRows = CountyNum
  If FlagZone = "TOWN" Then iRows = TownNum
  
  ReDim BBB_ZONE(1 To iRows): ReDim BBB_ZONE_sum(1 To iRows): ReDim BBB_ZONE_num(1 To iRows)
  
  BYear = LBound(AAA, 2): EYear = UBound(AAA, 2)
  ReDim AAA_ZONE(1 To iRows, BYear To EYear): ReDim AAA_ZONE_sum(1 To iRows, BYear To EYear): ReDim AAA_ZONE_num(1 To iRows, BYear To EYear)

  '逐站点循环，统计各区域累计/最大/最小/计数值
  For i = 1 To StaNum
    For k = 1 To iRows
      If FlagZone = "CITY" Then
        If StaInfo(i).City = CityInfo2(k).City Then
          Call Accumulate_ZONE_AAA(AAA, AAA_ZONE_sum, AAA_ZONE_num, BBB, BBB_ZONE_sum, BBB_ZONE_num, i, k, BYear, EYear)
          Exit For
        End If
        
      ElseIf FlagZone = "CNTY" Then
        If StaInfo(i).City = CountyInfo2(k).City And StaInfo(i).County = CountyInfo2(k).County Then
          Call Accumulate_ZONE_AAA(AAA, AAA_ZONE_sum, AAA_ZONE_num, BBB, BBB_ZONE_sum, BBB_ZONE_num, i, k, BYear, EYear)
          Exit For
        End If
      
      ElseIf FlagZone = "TOWN" Then
        If StaInfo(i).City = TownInfo2(k).City And StaInfo(i).County = TownInfo2(k).County And StaInfo(i).Town = TownInfo2(k).Town Then
          Call Accumulate_ZONE_AAA(AAA, AAA_ZONE_sum, AAA_ZONE_num, BBB, BBB_ZONE_sum, BBB_ZONE_num, i, k, BYear, EYear)
          Exit For
        End If
      
      End If
    
    Next k
  Next i
  
  For k = 1 To iRows
    If BBB_ZONE_num(k) <> 0 Then BBB_ZONE(k) = Format(BBB_ZONE_sum(k) / BBB_ZONE_num(k), "0.0")
    If BBB_ZONE_num(k) = 0 Then BBB_ZONE(k) = -9999
    For j = BYear To EYear
      If AAA_ZONE_num(k, j) <> 0 Then AAA_ZONE(k, j) = Format(AAA_ZONE_sum(k, j) / AAA_ZONE_num(k, j), "0.0")
      If AAA_ZONE_num(k, j) = 0 Then AAA_ZONE(k, j) = -9999
    Next j
  Next k

End Sub


'统计- 任意时段统计各站点历年值AAA - 历年各区域（即市县镇）平均
Public Sub Stat_ZONE_AAAExt(AAA!(), AAA_ZONE!(), BBB!(), BBB_ZONE!())
  Dim AAA_ZONE_sum!(), AAA_ZONE_num%(), BBB_ZONE_sum!(), BBB_ZONE_num%()
  Dim BYear%, EYear% 'AAA的起止年份

  Dim iRows%
  Dim i%, j%, k%, l%

  If FlagZoneExt = "CITY" Then iRows = ExtCityNum
  If FlagZoneExt = "CNTY" Then iRows = ExtCountyNum
  If FlagZoneExt = "PROV" Then iRows = ExtProvNum
  
  ReDim BBB_ZONE(1 To iRows): ReDim BBB_ZONE_sum(1 To iRows): ReDim BBB_ZONE_num(1 To iRows)
  
  BYear = LBound(AAA, 2): EYear = UBound(AAA, 2)
  ReDim AAA_ZONE(1 To iRows, BYear To EYear): ReDim AAA_ZONE_sum(1 To iRows, BYear To EYear): ReDim AAA_ZONE_num(1 To iRows, BYear To EYear)

  '逐站点循环，统计各区域累计/最大/最小/计数值
  For i = 1 To ExtStaNum
    For k = 1 To iRows
      
      
      If FlagZoneExt = "PROV" Then
        If ExtStaInfo(i).Prov = ExtProvInfo2(k).Prov Then
          Call Accumulate_ZONE_AAA(AAA, AAA_ZONE_sum, AAA_ZONE_num, BBB, BBB_ZONE_sum, BBB_ZONE_num, i, k, BYear, EYear)
          Exit For
        End If
      
      ElseIf FlagZoneExt = "CITY" Then
        If ExtStaInfo(i).Prov = ExtCityInfo2(k).Prov And ExtStaInfo(i).City = ExtCityInfo2(k).City Then
          Call Accumulate_ZONE_AAA(AAA, AAA_ZONE_sum, AAA_ZONE_num, BBB, BBB_ZONE_sum, BBB_ZONE_num, i, k, BYear, EYear)
          Exit For
        End If
        
      ElseIf FlagZoneExt = "CNTY" Then
        If ExtStaInfo(i).Prov = ExtCountyInfo2(k).Prov And ExtStaInfo(i).City = ExtCountyInfo2(k).City And ExtStaInfo(i).County = ExtCountyInfo2(k).County Then
          Call Accumulate_ZONE_AAA(AAA, AAA_ZONE_sum, AAA_ZONE_num, BBB, BBB_ZONE_sum, BBB_ZONE_num, i, k, BYear, EYear)
          Exit For
        End If
      
      End If
    
    Next k
  Next i
  
  For k = 1 To iRows
    If BBB_ZONE_num(k) <> 0 Then BBB_ZONE(k) = Format(BBB_ZONE_sum(k) / BBB_ZONE_num(k), "0.0")
    If BBB_ZONE_num(k) = 0 Then BBB_ZONE(k) = -9999
    For j = BYear To EYear
      If AAA_ZONE_num(k, j) <> 0 Then AAA_ZONE(k, j) = Format(AAA_ZONE_sum(k, j) / AAA_ZONE_num(k, j), "0.0")
      If AAA_ZONE_num(k, j) = 0 Then AAA_ZONE(k, j) = -9999
    Next j
  Next k

End Sub



'统计- 冷空气统计结果 -多行（即多站点）平均、最大、最小值
Public Sub Stat_ROWS_ColdAir(DataIn() As ColdAir_Result)
  Dim sngSum!, sngAvg!, sngMax!, sngMin!, iSta% '当年值的统计量
  Dim dateMax As Date, dateMin As Date
  Dim iNum% '用于计算多站平均的站点数
  Dim tmpData!, tmpDate As Date
  Dim iRows%
  
  iRows = UBound(DataIn)
  
  ReDim ROWS_Stat_ColdAir(1 To 3) '1、2、3下标分别对应多站点平均、最大、最小值
  For i = 1 To 3
    ROWS_Stat_ColdAir(i).ilevel = -9999 '冷空气等级：0-无，1-弱，2-中，3-强，4-寒潮
    ROWS_Stat_ColdAir(i).BDate = CDate("1899-09-09") '降温第一日
    ROWS_Stat_ColdAir(i).EDate = CDate("1899-09-09") '回温前一日
    ROWS_Stat_ColdAir(i).idays = -9999 '降温日数
    
    ROWS_Stat_ColdAir(i).tv_24hmax = -9999 '最大24小时降温
    ROWS_Stat_ColdAir(i).tv_48hmax = -9999 '最大48小时降温
    ROWS_Stat_ColdAir(i).tv_acc = -9999 '累计降温
    ROWS_Stat_ColdAir(i).T = -9999  '降温过程的平均气温
    ROWS_Stat_ColdAir(i).t_min = -9999 '降温过程的最低气温
  Next i
  
  '循环非日期的7列
  For j = 1 To 9
    sngSum = 0: iNum = 0: sngAvg = -9999: sngMax = -9999: sngMin = 999999
      
    For i = 1 To iRows
      If j = 1 Then tmpData = DataIn(i).ilevel
      If j = 2 Then tmpDate = DataIn(i).BDate
      If j = 3 Then tmpDate = DataIn(i).EDate
      
      If j = 4 Then tmpData = DataIn(i).idays
      If j = 5 Then tmpData = DataIn(i).tv_24hmax
      If j = 6 Then tmpData = DataIn(i).tv_48hmax
      If j = 7 Then tmpData = DataIn(i).tv_acc
      If j = 8 Then tmpData = DataIn(i).T
      If j = 9 Then tmpData = DataIn(i).t_min
      If tmpData <> -9999 Then
        If tmpData > sngMax Then sngMax = tmpData
        If tmpData < sngMin Then sngMin = tmpData
        sngSum = sngSum + tmpData
        iNum = iNum + 1
      End If
    Next i
      
'    If iNum <> 0 Then sngAvg = Round(sngSum / iNum, 1)
    If iNum <> 0 Then sngAvg = Format(sngSum / iNum, "0.0")
    If iNum = 0 Then sngAvg = -9999: sngSum = -9999
    If sngMax = -9999 Then sngMax = -9999
    If sngMin = 999999 Then sngMin = -9999
    
    If j = 1 Then
      ROWS_Stat_ColdAir(1).ilevel = sngAvg: ROWS_Stat_ColdAir(2).ilevel = sngMax: ROWS_Stat_ColdAir(3).ilevel = sngMin
    ElseIf j = 4 Then
      ROWS_Stat_ColdAir(1).idays = sngAvg: ROWS_Stat_ColdAir(2).idays = sngMax: ROWS_Stat_ColdAir(3).idays = sngMin
    ElseIf j = 5 Then
      ROWS_Stat_ColdAir(1).tv_24hmax = sngAvg: ROWS_Stat_ColdAir(2).tv_24hmax = sngMax: ROWS_Stat_ColdAir(3).tv_24hmax = sngMin
    ElseIf j = 6 Then
      ROWS_Stat_ColdAir(1).tv_48hmax = sngAvg: ROWS_Stat_ColdAir(2).tv_48hmax = sngMax: ROWS_Stat_ColdAir(3).tv_48hmax = sngMin
    ElseIf j = 7 Then
      ROWS_Stat_ColdAir(1).tv_acc = sngAvg: ROWS_Stat_ColdAir(2).tv_acc = sngMax: ROWS_Stat_ColdAir(3).tv_acc = sngMin
    ElseIf j = 8 Then
      ROWS_Stat_ColdAir(1).T = sngAvg: ROWS_Stat_ColdAir(2).T = sngMax: ROWS_Stat_ColdAir(3).T = sngMin
    ElseIf j = 9 Then
      ROWS_Stat_ColdAir(1).t_min = sngAvg: ROWS_Stat_ColdAir(2).t_min = sngMax: ROWS_Stat_ColdAir(3).t_min = sngMin
    End If
  Next j
  
 
  '循环日期2列
  For j = 1 To 2
    dateMax = CDate("1899-09-09"): dateMin = CDate("2099-09-09")
    For i = 1 To iRows
      If j = 1 Then tmpDate = DataIn(i).BDate
      If j = 2 Then tmpDate = DataIn(i).EDate
      If tmpDate <> CDate("1899-09-09") Then
        If tmpDate > dateMax Then dateMax = tmpDate
        If tmpDate < dateMin Then dateMin = tmpDate
      End If
    Next i
    If j = 1 Then
      ROWS_Stat_ColdAir(2).BDate = dateMax: ROWS_Stat_ColdAir(3).BDate = dateMin
    ElseIf j = 2 Then
      ROWS_Stat_ColdAir(2).EDate = dateMax: ROWS_Stat_ColdAir(3).EDate = dateMin
    End If
  Next j
  
  
End Sub

Public Sub StatCAData(DataIn() As ColdAir_Result, DataOut() As ColdAir_Result)
  Dim i%, j%, k%
  Dim LastSTA$
  Dim iRows%
  
  iRows = UBound(DataIn)
  k = 1: LastSTA = ""
  For i = 1 To iRows
    If DataIn(i).stacode <> LastSTA Then
      LastSTA = DataIn(i).stacode
      For j = k To StaNum
        If DataIn(i).stacode = DataOut(j).stacode Then
          If DataIn(i).ilevel >= DataOut(j).ilevel Then
            DataOut(j).ilevel = DataIn(i).ilevel
            DataOut(j).BDate = DataIn(i).BDate
            DataOut(j).EDate = DataIn(i).EDate
            DataOut(j).idays = DataIn(i).idays
            DataOut(j).tv_24hmax = DataIn(i).tv_24hmax
            DataOut(j).tv_48hmax = DataIn(i).tv_48hmax
            DataOut(j).tv_acc = DataIn(i).tv_acc
            DataOut(j).T = DataIn(i).T
            DataOut(j).t_min = DataIn(i).t_min
          End If
          k = k + 1: Exit For
        End If
      Next j
  
    ElseIf DataIn(i).stacode = LastSTA Then
      If DataIn(i).ilevel >= DataOut(j).ilevel Then
        DataOut(j).ilevel = DataIn(i).ilevel
        DataOut(j).BDate = DataIn(i).BDate
        DataOut(j).EDate = DataIn(i).EDate
        DataOut(j).idays = DataIn(i).idays
        DataOut(j).tv_24hmax = DataIn(i).tv_24hmax
        DataOut(j).tv_48hmax = DataIn(i).tv_48hmax
        DataOut(j).tv_acc = DataIn(i).tv_acc
        DataOut(j).T = DataIn(i).T
        DataOut(j).t_min = DataIn(i).t_min
      End If
    
    End If
  Next i

End Sub


'统计- 季节起始日统计结果 -多行（即多站点）平均、最大、最小值
Public Sub Stat_ROWS_Seasons(DataIn() As Season_Result, BDateNorm As Date, BDate1 As Date, BDate2 As Date)
  Dim sngSum!, sngAvg!, sngMax!, sngMin!, iSta% '当年值的统计量
  Dim dateMax As Date, dateMin As Date
  Dim iNum% '用于计算多站平均的站点数
  Dim tmpData!, tmpDate As Date
  Dim iRows%
  
  iRows = UBound(DataIn)
  
  ReDim ROWS_Stat_Seasons(1 To 3) '1、2、3下标分别对应多站点平均、最大、最小值
  For i = 1 To 3
    ROWS_Stat_Seasons(i).OnsetNorm = CDate("1899-09-09")
    ROWS_Stat_Seasons(i).T_IdxNorm = -9999
    ROWS_Stat_Seasons(i).OnsetCurr = CDate("1899-09-09")
    ROWS_Stat_Seasons(i).T_IdxCurr = -9999
    ROWS_Stat_Seasons(i).OnsetComp = CDate("1899-09-09")
    ROWS_Stat_Seasons(i).T_IdxComp = -9999
    ROWS_Stat_Seasons(i).diffNorm = -9999
    ROWS_Stat_Seasons(i).diffComp = -9999
  Next i
  
  
  '循环非日期的5列
  For j = 1 To 5
    sngSum = 0: iNum = 0: sngAvg = -9999: sngMax = -9999: sngMin = 999999
      
    For i = 1 To iRows
      If j = 1 Then tmpData = DataIn(i).diffNorm
      If j = 2 Then tmpData = DataIn(i).diffComp
      If j = 3 Then tmpData = DataIn(i).T_IdxNorm
      If j = 4 Then tmpData = DataIn(i).T_IdxCurr
      If j = 5 Then tmpData = DataIn(i).T_IdxComp
      If tmpData <> -9999 Then
        If tmpData > sngMax Then sngMax = tmpData
        If tmpData < sngMin Then sngMin = tmpData
        sngSum = sngSum + tmpData
        iNum = iNum + 1
      End If
    Next i
      
    If iNum <> 0 Then sngAvg = Format(sngSum / iNum, "0.0")
    If iNum = 0 Then sngAvg = -9999: sngSum = -9999
    If sngMax = -9999 Then sngMax = -9999
    If sngMin = 999999 Then sngMin = -9999
    
    If j = 1 Then
      ROWS_Stat_Seasons(1).diffNorm = sngAvg: ROWS_Stat_Seasons(2).diffNorm = sngMax: ROWS_Stat_Seasons(3).diffNorm = sngMin
    ElseIf j = 2 Then
      ROWS_Stat_Seasons(1).diffComp = sngAvg: ROWS_Stat_Seasons(2).diffComp = sngMax: ROWS_Stat_Seasons(3).diffComp = sngMin
    ElseIf j = 3 Then
      ROWS_Stat_Seasons(1).T_IdxNorm = sngAvg: ROWS_Stat_Seasons(2).T_IdxNorm = sngMax: ROWS_Stat_Seasons(3).T_IdxNorm = sngMin
    ElseIf j = 4 Then
      ROWS_Stat_Seasons(1).T_IdxCurr = sngAvg: ROWS_Stat_Seasons(2).T_IdxCurr = sngMax: ROWS_Stat_Seasons(3).T_IdxCurr = sngMin
    ElseIf j = 5 Then
      ROWS_Stat_Seasons(1).T_IdxComp = sngAvg: ROWS_Stat_Seasons(2).T_IdxComp = sngMax: ROWS_Stat_Seasons(3).T_IdxComp = sngMin
    End If
  Next j
  
 
  '循环日期3列
  For j = 1 To 3
    dateMax = CDate("1899-09-09"): dateMin = CDate("2099-09-09")
    For i = 1 To iRows
      If j = 1 Then tmpDate = DataIn(i).OnsetNorm
      If j = 2 Then tmpDate = DataIn(i).OnsetCurr
      If j = 3 Then tmpDate = DataIn(i).OnsetComp
      
      If tmpDate <> CDate("1899-09-09") Then
        If tmpDate > dateMax Then dateMax = tmpDate
        If tmpDate < dateMin Then dateMin = tmpDate
      End If
    Next i
    If j = 1 Then
      ROWS_Stat_Seasons(1).OnsetNorm = BDateNorm - 1 + Round(ROWS_Stat_Seasons(1).T_IdxNorm, 0): ROWS_Stat_Seasons(2).OnsetNorm = dateMax: ROWS_Stat_Seasons(3).OnsetNorm = dateMin
    ElseIf j = 2 Then
      ROWS_Stat_Seasons(1).OnsetCurr = BDate1 - 1 + Round(ROWS_Stat_Seasons(1).T_IdxCurr, 0): ROWS_Stat_Seasons(2).OnsetCurr = dateMax: ROWS_Stat_Seasons(3).OnsetCurr = dateMin
    ElseIf j = 3 Then
      ROWS_Stat_Seasons(1).OnsetComp = BDate2 - 1 + Round(ROWS_Stat_Seasons(1).T_IdxComp, 0): ROWS_Stat_Seasons(2).OnsetComp = dateMax: ROWS_Stat_Seasons(3).OnsetComp = dateMin
    End If
  Next j
  
  
End Sub


