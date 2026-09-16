Attribute VB_Name = "SQLGeoInfo"


'******2022-01-18******
'******镇乡信息******
Public Type TownInfo
  Prov As String
  Region As String
  City As String
  County As String
  Town As String
  Code As String '2023-2-16添加：区域编码
End Type
Public TownInfo1() As TownInfo, TownInfo2() As TownInfo '可选镇乡、已选镇乡
Public TownNum% '用户已选镇乡数

'******区县信息******
Public Type CountyInfo
  Prov As String
  Region As String
  City As String
  County As String
  Code As String '2023-2-16添加：区域编码
End Type
Public CountyInfo1() As CountyInfo, CountyInfo2() As CountyInfo '可选区县、已选区县
Public CountyNum% '已选区县数

'******地市信息******
Public Type CityInfo
  Prov As String
  Region As String
  City As String
  Code As String '2023-2-16添加：区域编码
End Type
Public CityInfo1() As CityInfo, CityInfo2() As CityInfo  '可选地市、已选地市
Public CityNum% '已选城市数


'******2026-7-16：添加区县、地市、省份信息******
Public Type ProvInfo
  Prov As String
  Code As String
End Type
Public ExtCountyInfo1() As CountyInfo, ExtCountyInfo2() As CountyInfo
Public ExtCountyNum%
Public ExtCityInfo1() As CityInfo, ExtCityInfo2() As CityInfo
Public ExtCityNum%
Public ExtProvInfo1() As ProvInfo, ExtProvInfo2() As ProvInfo
Public ExtProvNum%

Public strSURFExt$  '存放区域统计选择后的国家站站点sql script

'******区域编码信息(2023.2.16)******
Public Type TownCode
  Code As String
  Town As String
  County As String
  City As String
End Type
Public Town_Code() As TownCode

Public Type CountyCode
  Code As String
  County As String
  City As String
End Type
Public County_Code() As TownCode

Public Type CityCode
  Code As String
  City As String
End Type
Public City_Code() As TownCode
'==================== NewType.bas 迁移内容结束 ====================


Public Sub GetCityCode() '2023.2.16添加区域编码
  
  strSQL = "SELECT c.ADMIN_CODE ,c.ADMIN_ZONE city"
  strSQL = strSQL + " FROM T_OTHE_ZONE_CODE_TAB c"
  strSQL = strSQL + " WHERE ILEVEL = 2"
  strSQL = strSQL + " ORDER BY city"
  
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim City_Code(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    City_Code(i).Code = ORARst!ADMIN_CODE
    City_Code(i).City = ORARst!City
    i = i + 1
    ORARst.MoveNext
  Loop
  ORARst.Close

End Sub


Public Sub GetCountyCode() '2023.2.16添加区域编码
  
  strSQL = "SELECT b.ADMIN_CODE ,b.ADMIN_ZONE county, c.ADMIN_ZONE city"
  strSQL = strSQL + " FROM T_OTHE_ZONE_CODE_TAB b, T_OTHE_ZONE_CODE_TAB c"
  strSQL = strSQL + " Where substr(b.ADMIN_CODE, 0, 4) = substr(c.ADMIN_CODE, 0, 4)"
  strSQL = strSQL + " and b.ILEVEL = 3  and c.ILEVEL = 2"
  strSQL = strSQL + " Union"
  strSQL = strSQL + " SELECT c.ADMIN_CODE, c.ADMIN_ZONE || '市' county, c.ADMIN_ZONE city"
  strSQL = strSQL + " FROM T_OTHE_ZONE_CODE_TAB c"
  strSQL = strSQL + " where c.ILEVEL =2 AND c.ADMIN_ZONE IN ('中山','东莞')"
  strSQL = strSQL + " ORDER BY city,county"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim County_Code(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    County_Code(i).Code = ORARst!ADMIN_CODE
    County_Code(i).County = ORARst!County
    County_Code(i).City = ORARst!City
    i = i + 1
    ORARst.MoveNext
  Loop
  ORARst.Close

End Sub


Public Sub GetTownCode() '2023.2.16添加区域编码
  
  strSQL = "SELECT a.ADMIN_CODE, a.ADMIN_ZONE town, b.ADMIN_ZONE county, c.ADMIN_ZONE city"
  strSQL = strSQL + " FROM T_OTHE_ZONE_CODE_TAB a, T_OTHE_ZONE_CODE_TAB b, T_OTHE_ZONE_CODE_TAB c"
  strSQL = strSQL + " Where substr(a.ADMIN_CODE, 0, 6) = substr(b.ADMIN_CODE, 0, 6)"
  strSQL = strSQL + " and substr(b.ADMIN_CODE, 0, 4) = substr(c.ADMIN_CODE, 0, 4)"
  strSQL = strSQL + " and a.ILEVEL = 4 and b.ILEVEL = 3  and c.ILEVEL = 2"
  strSQL = strSQL + " Union"
  strSQL = strSQL + " SELECT a.ADMIN_CODE, a.ADMIN_ZONE town, c.ADMIN_ZONE || '市' county, c.ADMIN_ZONE city"
  strSQL = strSQL + " FROM T_OTHE_ZONE_CODE_TAB a, T_OTHE_ZONE_CODE_TAB c"
  strSQL = strSQL + " Where substr(a.ADMIN_CODE, 0, 4) = substr(c.ADMIN_CODE, 0, 4)"
  strSQL = strSQL + " and a.ILEVEL = 4 and c.ILEVEL = 2 AND c.ADMIN_ZONE IN ('中山','东莞')"
  strSQL = strSQL + " ORDER BY city,county,town"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim Town_Code(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    Town_Code(i).Code = ORARst!ADMIN_CODE
    Town_Code(i).Town = ORARst!Town
    Town_Code(i).County = ORARst!County
    Town_Code(i).City = ORARst!City
    i = i + 1
    ORARst.MoveNext
  Loop
  ORARst.Close

End Sub



Public Sub GetCityInfo1(tmpZones$) '传递数据范围 tmpZones：'城市1','城市2','城市3','城市4','城市5','城市6','城市7'
  '用户可以各地区平均统计，取出各个地区信息
  strSQL = "SELECT V_Prcode,region,v_city FROM"
  strSQL = strSQL + " (SELECT V_Prcode,v_city From T_OTHE_STATION_META_BASIC_TAB WHERE v_city in (" & tmpZones & ")"
  strSQL = strSQL + " GROUP BY V_Prcode,v_city) a"
'  strSQL = strSQL + " AND v_city IS NOT null GROUP BY V_Prcode,v_city) a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b on a.v_city=b.admin_zone"
  strSQL = strSQL + " order by NLSSORT(v_city, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim CityInfo1(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    CityInfo1(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!Region) Then CityInfo1(i).Region = ORARst!Region
    If Not IsNull(ORARst!v_city) Then CityInfo1(i).City = ORARst!v_city
    
    For j = 1 To UBound(City_Code)
      If CityInfo1(i).City = City_Code(j).City Then
        CityInfo1(i).Code = City_Code(j).Code: Exit For
      End If
    Next j
    
    i = i + 1
    ORARst.MoveNext
  Loop
  ORARst.Close

End Sub


Public Sub GetCountyInfo1(tmpZones$) '传递数据范围 tmpZones：'城市1','城市2','城市3','城市4','城市5','城市6','城市7'
  '用户可以各地区平均统计，取出各个地区信息
  strSQL = "SELECT V_Prcode,region,v_city,v_county FROM"
  strSQL = strSQL + " (SELECT V_Prcode,v_city,v_county From T_OTHE_STATION_META_BASIC_TAB WHERE v_city in (" & tmpZones & ")"
  strSQL = strSQL + " GROUP BY V_Prcode,v_city,v_county) a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b on a.v_city=b.admin_zone"
  strSQL = strSQL + " order by NLSSORT(v_city, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim CountyInfo1(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    CountyInfo1(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!Region) Then CountyInfo1(i).Region = ORARst!Region
    If Not IsNull(ORARst!v_city) Then CountyInfo1(i).City = ORARst!v_city
    If Not IsNull(ORARst!v_county) Then CountyInfo1(i).County = ORARst!v_county
    
    For j = 1 To UBound(County_Code)
      If CountyInfo1(i).City = County_Code(j).City And CountyInfo1(i).County = County_Code(j).County Then
        CountyInfo1(i).Code = County_Code(j).Code: Exit For
      End If
    Next j
    
    i = i + 1
    ORARst.MoveNext
  Loop
  ORARst.Close

End Sub


Public Sub GetTownInfo1(tmpZones$) '传递数据范围 tmpZones：'城市1','城市2','城市3','城市4','城市5','城市6','城市7'
  '用户可以各地区平均统计，取出各个地区信息
  strSQL = "SELECT V_Prcode,region,v_city,v_county,v_town FROM"
  strSQL = strSQL + " (SELECT V_Prcode,v_city,v_county,v_town From T_OTHE_STATION_META_BASIC_TAB WHERE v_city in (" & tmpZones & ")"
  strSQL = strSQL + " AND v_town IS NOT null GROUP BY V_Prcode,v_city,v_county,v_town) a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b on a.v_city=b.admin_zone" 'ORDER BY v_city
  strSQL = strSQL + " order by NLSSORT(v_city, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim TownInfo1(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    TownInfo1(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!Region) Then TownInfo1(i).Region = ORARst!Region
    If Not IsNull(ORARst!v_city) Then TownInfo1(i).City = ORARst!v_city
    If Not IsNull(ORARst!v_county) Then TownInfo1(i).County = ORARst!v_county
    If Not IsNull(ORARst!v_town) Then TownInfo1(i).Town = ORARst!v_town
    
    For j = 1 To UBound(Town_Code)
      If TownInfo1(i).City = Town_Code(j).City And TownInfo1(i).County = Town_Code(j).County And TownInfo1(i).Town = Town_Code(j).Town Then
        TownInfo1(i).Code = Town_Code(j).Code: Exit For
      End If
    Next j
    
    i = i + 1
    ORARst.MoveNext
  Loop
  ORARst.Close

End Sub


Public Sub GetCityInfo2(tmpZone$) '传递数据范围 tmpZone：'广东'或者 '城市'
  '用户可以各地区平均统计，取出各个地区信息
  strSQL = "SELECT V_Prcode,region,v_city FROM"
  
  If InStr(tmpZone, ",") = 0 Then
    If tmpZone = "广东" Or tmpZone = "粤港澳" Then
      strSQL = strSQL + " (SELECT V_Prcode,v_city From T_OTHE_STATION_META_BASIC_TAB WHERE v_prcode ='广东'"
    Else
      strSQL = strSQL + " (SELECT V_Prcode,v_city From T_OTHE_STATION_META_BASIC_TAB WHERE v_city ='" & tmpZone & "'"
    End If
  Else
    strSQL = strSQL + " (SELECT V_Prcode,v_city From T_OTHE_STATION_META_BASIC_TAB WHERE v_city in (" & tmpZone & ")"
  End If
  
  strSQL = strSQL + " GROUP BY V_Prcode,v_city) a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b on a.v_city=b.admin_zone"
  strSQL = strSQL + " order by NLSSORT(v_city, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim CityInfo2(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    CityInfo2(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!Region) Then CityInfo2(i).Region = ORARst!Region
    If Not IsNull(ORARst!v_city) Then CityInfo2(i).City = ORARst!v_city
    
    For j = 1 To UBound(City_Code)
      If CityInfo2(i).City = City_Code(j).City Then
        CityInfo2(i).Code = City_Code(j).Code: Exit For
      End If
    Next j
    
    i = i + 1
    ORARst.MoveNext
  Loop
  ORARst.Close

End Sub


Public Sub GetCountyInfo2(tmpZone$, tmpCounties$) '传递数据范围 tmpZone：'广东'或者 '城市'
  '用户可以各地区平均统计，取出各个地区信息
  strSQL = "SELECT V_Prcode,region,v_city,v_county FROM"
  
  If InStr(tmpZone, ",") = 0 Then
    If tmpZone = "广东" Or tmpZone = "粤港澳" Then
      strSQL = strSQL + " (SELECT V_Prcode,v_city,v_county From T_OTHE_STATION_META_BASIC_TAB WHERE v_prcode ='广东'"
    Else
      strSQL = strSQL + " (SELECT V_Prcode,v_city,v_county From T_OTHE_STATION_META_BASIC_TAB WHERE v_city ='" & tmpZone & "'"
    End If
  Else
    strSQL = strSQL + " (SELECT V_Prcode,v_city,v_county From T_OTHE_STATION_META_BASIC_TAB WHERE v_city in (" & tmpZone & ")"
  End If
  
  If tmpCounties <> "" Then
    If InStr(tmpCounties, ",") = 0 Then
      strSQL = strSQL + " AND v_county = '" & tmpCounties & "'"
    Else
      strSQL = strSQL + " AND v_county in (" & tmpCounties & ")"
    End If
  End If
  
  strSQL = strSQL + " GROUP BY V_Prcode,v_city,v_county) a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b on a.v_city=b.admin_zone"
  strSQL = strSQL + " order by NLSSORT(v_city, 'NLS_SORT=SCHINESE_PINYIN_M')"
  strSQL = strSQL + " ,NLSSORT(v_county, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim CountyInfo2(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    CountyInfo2(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!Region) Then CountyInfo2(i).Region = ORARst!Region
    If Not IsNull(ORARst!v_city) Then CountyInfo2(i).City = ORARst!v_city
    If Not IsNull(ORARst!v_county) Then CountyInfo2(i).County = ORARst!v_county
    
    For j = 1 To UBound(County_Code)
      If CountyInfo2(i).City = County_Code(j).City And CountyInfo2(i).County = County_Code(j).County Then
        CountyInfo2(i).Code = County_Code(j).Code: Exit For
      End If
    Next j
    
    i = i + 1
    ORARst.MoveNext
  Loop
  ORARst.Close

End Sub


Public Sub GetTownInfo2(tmpZone$, tmpCounties$, tmpTowns$)  '传递数据范围 tmpZone：'广东'或者 '城市'
  '用户可以各地区平均统计，取出各个地区信息
  strSQL = "SELECT V_Prcode,region,v_city,v_county,v_town FROM"
  If InStr(tmpZone, ",") = 0 Then
    If tmpZone = "广东" Or tmpZone = "粤港澳" Then
      strSQL = strSQL + " (SELECT V_Prcode,v_city,v_county,v_town From T_OTHE_STATION_META_BASIC_TAB WHERE v_prcode ='广东'"
    Else
      strSQL = strSQL + " (SELECT V_Prcode,v_city,v_county,v_town From T_OTHE_STATION_META_BASIC_TAB WHERE v_city ='" & tmpZone & "'"
    End If
  Else
    strSQL = strSQL + " (SELECT V_Prcode,v_city,v_county,v_town From T_OTHE_STATION_META_BASIC_TAB WHERE v_city in (" & tmpZone & ")"
  End If
    
  If tmpCounties <> "" Then
    If InStr(tmpCounties, ",") = 0 Then
      strSQL = strSQL + " AND v_county = '" & tmpCounties & "'"
    Else
      strSQL = strSQL + " AND v_county in (" & tmpCounties & ")"
    End If
  End If
  
  If tmpTowns <> "" Then
    If InStr(tmpTowns, ",") = 0 Then
      strSQL = strSQL + " AND v_town = '" & tmpTowns & "'"
    Else
      strSQL = strSQL + " AND v_town in (" & tmpTowns & ")"
    End If
  End If
    
  strSQL = strSQL + " GROUP BY V_Prcode,v_city,v_county,v_town) a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b on a.v_city=b.admin_zone"
  strSQL = strSQL + " order by NLSSORT(v_city, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim TownInfo2(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    TownInfo2(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!Region) Then TownInfo2(i).Region = ORARst!Region
    If Not IsNull(ORARst!v_city) Then TownInfo2(i).City = ORARst!v_city
    If Not IsNull(ORARst!v_county) Then TownInfo2(i).County = ORARst!v_county
    If Not IsNull(ORARst!v_town) Then TownInfo2(i).Town = ORARst!v_town
    
    For j = 1 To UBound(Town_Code)
      If TownInfo2(i).City = Town_Code(j).City And TownInfo2(i).County = Town_Code(j).County And TownInfo2(i).Town = Town_Code(j).Town Then
        TownInfo2(i).Code = Town_Code(j).Code: Exit For
      End If
    Next j
    
    i = i + 1
    ORARst.MoveNext
  Loop
  ORARst.Close

End Sub



Public Sub LoadExtStaInfo()
  
  '2026-07-16 区域统计：首次点击"区域统计"按钮时加载区域自定义站点信息，仅加载一次
  strSQL = "SELECT V01301 stacode, slm staname, V06001 longitude, V05001 latitude, V_COUNTY county, V_CITY city, V_PRCODE prov, STT_DATE DATESTT"
  strSQL = strSQL & " FROM T_OTHE_STATION_META_BASIC_TAB"
  strSQL = strSQL & " WHERE v02301 LIKE '%A%'" ' and v_prcode not in ('香港','澳门') and v01301<>'59486'" 'AND extract(YEAR FROM stt_date) < 1990
  strSQL = strSQL & " ORDER BY V01301"

  STARst.CursorLocation = adUseClient
  STARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  
  STARst.MoveFirst
  ReDim ExtStaInfo2(1 To STARst.RecordCount)
  ExtStaNum = UBound(ExtStaInfo2)
  i = 1
  Do Until STARst.EOF
    ExtStaInfo2(i).Prov = STARst!Prov
    ExtStaInfo2(i).City = STARst!City
    If Not IsNull(STARst!County) Then ExtStaInfo2(i).County = STARst!County
    ExtStaInfo2(i).StaType = "SURF"
    ExtStaInfo2(i).stacode = STARst!stacode
    If Not IsNull(STARst!staname) Then ExtStaInfo2(i).staname = STARst!staname
    ExtStaInfo2(i).Longitude = STARst!Longitude
    ExtStaInfo2(i).Latitude = STARst!Latitude
    ExtStaInfo2(i).DateSTT = STARst!DateSTT
    ExtStaInfo2(i).YearSTT = Year(STARst!DateSTT)
    ExtStaInfo2(i).Inland = 1
    ExtStaInfo2(i).Hydro = 0
    i = i + 1
    STARst.MoveNext
  Loop
  
  STARst.MoveFirst
  ReDim ExtStaInfo1(1 To STARst.RecordCount)
  i = 1
  Do Until STARst.EOF
    ExtStaInfo1(i).Prov = STARst!Prov
    ExtStaInfo1(i).City = STARst!City
    If Not IsNull(STARst!County) Then ExtStaInfo1(i).County = STARst!County
    ExtStaInfo1(i).StaType = "SURF"
    ExtStaInfo1(i).stacode = STARst!stacode
    If Not IsNull(STARst!staname) Then ExtStaInfo1(i).staname = STARst!staname
    ExtStaInfo1(i).Longitude = STARst!Longitude
    ExtStaInfo1(i).Latitude = STARst!Latitude
    ExtStaInfo1(i).DateSTT = STARst!DateSTT
    ExtStaInfo1(i).YearSTT = Year(STARst!DateSTT)
    ExtStaInfo1(i).Inland = 1
    ExtStaInfo1(i).Hydro = 0
    ExtStaInfo1(i).Selected = (ExtStaInfo1(i).Prov = "广东" And ExtStaInfo1(i).stacode <> "59486") '2026-07-17 默认仅选中广东站点(不包括59486)
    i = i + 1
    STARst.MoveNext
  Loop
  
  STARst.Close


End Sub

Public Sub RebuildExtSelFromSelected()
  '2026-07-18 根据 ExtStaInfo1().Selected 的当前状态，统一重建 ExtStaInfo2/ExtStaNum 及去重后的省份/地市/区县列表
  Dim m%, N%, j%, blnFound As Boolean

  ExtStaNum = 0
  For m = 1 To UBound(ExtStaInfo1)
    If ExtStaInfo1(m).Selected = True Then ExtStaNum = ExtStaNum + 1
  Next m

  ReDim ExtStaInfo2(1 To ExtStaNum)
  ReDim ExtProvInfo2(1 To ExtStaNum)
  ReDim ExtCityInfo2(1 To ExtStaNum)
  ReDim ExtCountyInfo2(1 To ExtStaNum)
  ExtProvNum = 0: ExtCityNum = 0: ExtCountyNum = 0

  j = 0
  For m = 1 To UBound(ExtStaInfo1)
    If ExtStaInfo1(m).Selected = True Then
      j = j + 1
      ExtStaInfo2(j) = ExtStaInfo1(m)

      blnFound = False
      For N = 1 To ExtProvNum
        If ExtProvInfo2(N).Prov = ExtStaInfo1(m).Prov Then blnFound = True: Exit For
      Next N
      If Not blnFound Then
        ExtProvNum = ExtProvNum + 1
        ExtProvInfo2(ExtProvNum).Prov = ExtStaInfo1(m).Prov
      End If

      blnFound = False
      For N = 1 To ExtCityNum
        If ExtCityInfo2(N).Prov = ExtStaInfo1(m).Prov And ExtCityInfo2(N).City = ExtStaInfo1(m).City Then blnFound = True: Exit For
      Next N
      If Not blnFound Then
        ExtCityNum = ExtCityNum + 1
        ExtCityInfo2(ExtCityNum).Prov = ExtStaInfo1(m).Prov
        ExtCityInfo2(ExtCityNum).City = ExtStaInfo1(m).City
      End If

      blnFound = False
      For N = 1 To ExtCountyNum
        If ExtCountyInfo2(N).Prov = ExtStaInfo1(m).Prov And ExtCountyInfo2(N).City = ExtStaInfo1(m).City And ExtCountyInfo2(N).County = ExtStaInfo1(m).County Then blnFound = True: Exit For
      Next N
      If Not blnFound Then
        ExtCountyNum = ExtCountyNum + 1
        ExtCountyInfo2(ExtCountyNum).Prov = ExtStaInfo1(m).Prov
        ExtCountyInfo2(ExtCountyNum).City = ExtStaInfo1(m).City
        ExtCountyInfo2(ExtCountyNum).County = ExtStaInfo1(m).County
      End If
    End If
  Next m

  ReDim Preserve ExtProvInfo2(1 To ExtProvNum)
  ReDim Preserve ExtCityInfo2(1 To ExtCityNum)
  ReDim Preserve ExtCountyInfo2(1 To ExtCountyNum)

  ExtStaInfo = ExtStaInfo2
End Sub


Public Sub loadExtZoneInfo()
  
 '***************2026-08-11 获取省份**********
  Call GetExtProvInfo1
  Call GetExtProvInfo2("")

'***************2026-08-11 获取地市**********
  Call GetExtCityInfo1
  Call GetExtCityInfo2("", "")

'***************2026-08-11 获取县区**********
  Call GetExtCountyInfo1
  Call GetExtCountyInfo2("", "", "")
    
  '2026-07-16 默认总是查询站点尺度数据(区域统计)
  FlagZoneExt = "STA"
  
  ExtProvNum = UBound(ExtProvInfo2)
  ExtCityNum = UBound(ExtCityInfo2)
  ExtCountyNum = UBound(ExtCountyInfo2)
  
End Sub



Public Sub GetExtProvInfo1()
  strSQL = "SELECT V_Prcode From T_OTHE_STATION_META_BASIC_TAB WHERE v02301 LIKE '%A%'"
  'and v_prcode not in ('香港','澳门') and v01301<>'59486'" 'AND extract(YEAR FROM stt_date) < 1990
  strSQL = strSQL + " GROUP BY V_Prcode"
  strSQL = strSQL + " order by NLSSORT(v_prcode, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  
  ORARst.MoveFirst
  ReDim ExtProvInfo1(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    ExtProvInfo1(i).Prov = ORARst!v_prcode
    i = i + 1
    ORARst.MoveNext
  Loop
  
  ORARst.Close
End Sub


Public Sub GetExtCityInfo1()
  strSQL = "SELECT V_Prcode,v_city From T_OTHE_STATION_META_BASIC_TAB WHERE v02301 LIKE '%A%'"
  'and v_prcode not in ('香港','澳门') and v01301<>'59486'" 'AND extract(YEAR FROM stt_date) < 1990
  strSQL = strSQL + " GROUP BY V_Prcode,v_city"
  strSQL = strSQL + " order by NLSSORT(v_city, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  
  ORARst.MoveFirst
  ReDim ExtCityInfo1(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    ExtCityInfo1(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!v_city) Then ExtCityInfo1(i).City = ORARst!v_city
    i = i + 1
    ORARst.MoveNext
  Loop
    
  ORARst.Close
End Sub


Public Sub GetExtCountyInfo1()
  strSQL = "SELECT V_Prcode,v_city,v_county From T_OTHE_STATION_META_BASIC_TAB WHERE v02301 LIKE '%A%'"
  strSQL = strSQL + " GROUP BY V_Prcode,v_city,v_county"
  strSQL = strSQL + " order by NLSSORT(v_county, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly

  ORARst.MoveFirst
  ReDim ExtCountyInfo1(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    ExtCountyInfo1(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!v_city) Then ExtCountyInfo1(i).City = ORARst!v_city
    If Not IsNull(ORARst!v_county) Then ExtCountyInfo1(i).County = ORARst!v_county
    i = i + 1
    ORARst.MoveNext
  Loop
    
  ORARst.Close
End Sub


Public Sub GetExtProvInfo2(tmpProvs$)
  strSQL = "SELECT V_Prcode From T_OTHE_STATION_META_BASIC_TAB WHERE v02301 LIKE '%A%'"
  If tmpProvs <> "" Then
    If InStr(tmpProvs, ",") = 0 Then
      strSQL = strSQL + " and v_prcode ='" + tmpProvs + "'"
    Else
      strSQL = strSQL + " and v_prcode in (" + tmpProvs + ")"
    End If
  End If
  
  strSQL = strSQL + " GROUP BY V_Prcode"
  strSQL = strSQL + " order by NLSSORT(v_prcode, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim ExtProvInfo2(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    ExtProvInfo2(i).Prov = ORARst!v_prcode
    i = i + 1
    ORARst.MoveNext
  Loop
    
  ORARst.Close
End Sub


Public Sub GetExtCityInfo2(tmpProvs$, tmpCities$)
  strSQL = "SELECT V_Prcode,v_city From T_OTHE_STATION_META_BASIC_TAB WHERE v02301 LIKE '%A%'"
  If tmpProvs <> "" Then
    If InStr(tmpProvs, ",") = 0 Then
      strSQL = strSQL + " and v_prcode ='" + tmpProvs + "'"
    Else
      strSQL = strSQL + " and v_prcode in (" + tmpProvs + ")"
    End If
  End If
  
  If tmpCities <> "" Then
    If InStr(tmpCities, ",") = 0 Then
      strSQL = strSQL + " and v_city ='" + tmpCities + "'"
    Else
      strSQL = strSQL + " and v_city in (" + tmpCities + ")"
    End If
  End If
  
  strSQL = strSQL + " GROUP BY V_Prcode,v_city"
  strSQL = strSQL + " order by NLSSORT(v_city, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim ExtCityInfo2(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    ExtCityInfo2(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!v_city) Then ExtCityInfo2(i).City = ORARst!v_city
    i = i + 1
    ORARst.MoveNext
  Loop
      
  ORARst.Close
End Sub


Public Sub GetExtCountyInfo2(tmpProvs$, tmpCities$, tmpCounties$)
  strSQL = "SELECT V_Prcode,v_city,v_county From T_OTHE_STATION_META_BASIC_TAB WHERE v02301 LIKE '%A%'"
  If tmpProvs <> "" Then
    If InStr(tmpProvs, ",") = 0 Then
      strSQL = strSQL + " and v_prcode ='" + tmpProvs + "'"
    Else
      strSQL = strSQL + " and v_prcode in (" + tmpProvs + ")"
    End If
  End If
  
  If tmpCities <> "" Then
    If InStr(tmpCities, ",") = 0 Then
      strSQL = strSQL + " and v_city ='" + tmpCities + "'"
    Else
      strSQL = strSQL + " and v_city in (" + tmpCities + ")"
    End If
  End If
  
  If tmpCounties <> "" Then
    If InStr(tmpCounties, ",") = 0 Then
      strSQL = strSQL + " and v_county ='" + tmpCounties + "'"
    Else
      strSQL = strSQL + " and v_county in (" + tmpCounties + ")"
    End If
  End If
  
  strSQL = strSQL + " GROUP BY V_Prcode,v_city,v_county"
  strSQL = strSQL + " order by NLSSORT(v_city, 'NLS_SORT=SCHINESE_PINYIN_M')"
  strSQL = strSQL + " ,NLSSORT(v_county, 'NLS_SORT=SCHINESE_PINYIN_M')"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  ORARst.MoveFirst
  ReDim ExtCountyInfo2(1 To ORARst.RecordCount)
  i = 1
  Do Until ORARst.EOF
    ExtCountyInfo2(i).Prov = ORARst!v_prcode
    If Not IsNull(ORARst!v_city) Then ExtCountyInfo2(i).City = ORARst!v_city
    If Not IsNull(ORARst!v_county) Then ExtCountyInfo2(i).County = ORARst!v_county
    i = i + 1
    ORARst.MoveNext
  Loop
    
  ORARst.Close
End Sub
