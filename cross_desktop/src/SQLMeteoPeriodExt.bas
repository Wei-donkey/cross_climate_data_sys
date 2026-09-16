Attribute VB_Name = "SQLMeteoPeriodExt"


'********************** 循环查询所有年份（Byear-Eyear）各站值（不包括极值出现年份）：开始  ****************************
'查询多年任意时段值（要素字段名，要素月值字段名，开始年度，结束年度，开始日月，结束日月，是否跨年，输出结果，统计量，是否条件查询，条件最小值，条件最大值）
Public Sub QueryMulYearDataExt(SelField_Day$, SelField_Restat$, SelField_Mon$, BYear%, EYear%, BMMDD$, EMMDD$, crossYear As Boolean, Result!(), ByVal FlagStat$, Min!, Max!)
  '如果结果将由几部分求平均获取，则：FlagStat由"AVE"替换为"AVE_SUM"
  Dim i%, j%, k%
  
  Dim date_STT As Date, date_END As Date   '开始日期、结束日期
  Dim date1 As Date, date2 As Date   '分割日期1、分割日期2
  Dim A1_STT As Date, A1_END As Date, A2_STT As Date, A2_END As Date, A3_STT As Date, A3_END As Date 'A1、A2、A3起止日期
  Dim temp_STT$, temp_END$ '起止月份或起止年月
  
  Dim iDays1%, iDays2%, iDays3%, idays% '被分割后的多段日期的日数，及总日数
  Dim AAA1!(), AAA2!(), AAA3!() ' 存放某年度所有站点：多日结果、多个整月结果、多日结果
  Dim ADays1%(), ADays2%(), ADays3%() ' 存放某年度所有站点条件日数：多日结果、多个整月结果、多日结果
  Dim imonths% '多月的月份数
  
  Dim iNum%, sngSum!, sngMax!, sngMin '求和的个数，和数，最大值、最小值
  
  Dim A1IsLastYearData As Boolean
  Dim A2IsLastYearData As Boolean
  Dim A3IsLastYearData As Boolean
  
  Dim tempYear%, tempEMMDD$, tempEDate$, tempBMMDD$, tempBDate$ '2017-4-10修复月值不存在时查询结果错误的bug
  Dim tempMMDD$, tempdate$
  Dim tempBBB!(), tempBBBDate() As Date  '当年值、当年极值日期
  Dim tempFlagStat$
  
  Dim TableDay$ '查询要素所在的数据表
  Dim TableMon$ '要素月值所在的数据表
     
  '******** 2022-04-01判断某年的月统计结果是否缺失月数据，若缺则标记FlagMonMiss()=True ********
'  Dim iCnts%() '每个站点每年的有效数据数量
'  ReDim iCnts(1 To extStaNum, BYear To EYear)
  Dim FlagMonMiss() As Boolean '每年月值缺失的标记
  Dim SelField_Cnts As String '月值表中要素数据量字段
  SelField_Cnts = SelField_Restat & "_CNTS"
  ReDim FlagMonMiss(BYear To EYear)
  '********************************************************************************************
     
     
  
'  If CrossYear = True Then EYear = EYear - 1
  ReDim AAA1(1 To ExtStaNum, BYear To EYear), AAA2(1 To ExtStaNum, BYear To EYear), AAA3(1 To ExtStaNum, BYear To EYear)
  ReDim ADays1(1 To ExtStaNum, BYear To EYear), ADays2(1 To ExtStaNum, BYear To EYear), ADays3(1 To ExtStaNum, BYear To EYear)
  For j = BYear To EYear
    For i = 1 To ExtStaNum
      AAA1(i, j) = -9999: AAA2(i, j) = -9999: AAA3(i, j) = -9999
      ADays1(i, j) = -9999: ADays2(i, j) = -9999: ADays3(i, j) = -9999
    Next i
  Next j
  
  
  If EMMDD <> "02-29" Then
    If crossYear = False Then
      date_STT = CDate("2011-" & BMMDD)
      date_END = CDate("2011-" & EMMDD)
    ElseIf crossYear = True Then
      date_STT = CDate("2010-" & BMMDD)
      date_END = CDate("2011-" & EMMDD)
    End If
    
  '2024-2-27 添加对闰年2-29的判断
  ElseIf EMMDD = "02-29" Then
    
    If crossYear = False Then
      date_STT = CDate("2012-" & BMMDD)
      date_END = CDate("2012-" & EMMDD)
    ElseIf crossYear = True Then
      date_STT = CDate("2011-" & BMMDD)
      date_END = CDate("2012-" & EMMDD)
    End If
    
  End If
    
    
  date1 = CDate(Format(DateAdd("m", 1, date_STT), "yyyy-mm-01")) - 1
  date2 = CDate(Format(date_END, "yyyy-mm-01"))
  
  A1IsLastYearData = False: A2IsLastYearData = False: A3IsLastYearData = False
  
'  If date1 > date_END Or (date1 = date_END And date2 < date_STT) _
'  Or (date1 < date_END And Format(date_STT, "dd") <> "01" And Format(date_END + 1, "dd") <> "01" And date1 + 1 = date2) Then '情形一：多日
  If date1 > date_END Or (date1 = date_END And date2 < date_STT) Then '情形一：多日
    A1_STT = date_STT: A1_END = date_END: iDays1 = A1_END - A1_STT + 1
    iDays2 = 0
    iDays3 = 0
    
  ElseIf (date1 = date_END And date2 = date_STT) _
  Or (date1 < date_END And Format(date_STT, "dd") = "01" And Format(date_END + 1, "dd") = "01") Then '情形二：多月
    iDays1 = 0
    A2_STT = date_STT: A2_END = date_END: iDays2 = A2_END - A2_STT + 1: imonths = DateDiff("m", A2_STT, A2_END) + 1
    iDays3 = 0
  
  ElseIf date1 < date_END And Format(date_STT, "dd") = "01" And Format(date_END + 1, "dd") <> "01" Then '情形三：多月+多日
    iDays1 = 0
    A2_STT = date_STT: A2_END = date2 - 1: iDays2 = A2_END - A2_STT + 1: imonths = DateDiff("m", A2_STT, A2_END) + 1
    A3_STT = date2: A3_END = date_END: iDays3 = A3_END - A3_STT + 1
    If FlagStat = "AVE" Then FlagStat = "AVE_SUM"
    If Year(A3_STT) <> Year(A2_STT) Then A3IsLastYearData = True
    
  ElseIf date1 < date_END And Format(date_STT, "dd") <> "01" And Format(date_END + 1, "dd") = "01" Then '情形四：多日+多月
    A1_STT = date_STT: A1_END = date1: iDays1 = A1_END - A1_STT + 1
    A2_STT = date1 + 1: A2_END = date_END: iDays2 = A2_END - A2_STT + 1: imonths = DateDiff("m", A2_STT, A2_END) + 1
    iDays3 = 0
    If FlagStat = "AVE" Then FlagStat = "AVE_SUM"
    If Year(A2_STT) <> Year(A1_STT) Then A2IsLastYearData = True
    
  ElseIf date1 < date_END And Format(date_STT, "dd") <> "01" And Format(date_END + 1, "dd") <> "01" And date1 + 1 <> date2 Then '情形五：多日+多月+多日
    A1_STT = date_STT: A1_END = date1: iDays1 = A1_END - A1_STT + 1
    A2_STT = date1 + 1: A2_END = date2 - 1: iDays2 = A2_END - A2_STT + 1: imonths = DateDiff("m", A2_STT, A2_END) + 1
    A3_STT = date2: A3_END = date_END: iDays3 = A3_END - A3_STT + 1
    If FlagStat = "AVE" Then FlagStat = "AVE_SUM"
    If Year(A2_STT) <> Year(A1_STT) Then A2IsLastYearData = True
    If Year(A3_STT) <> Year(A1_STT) Then A3IsLastYearData = True
  
  ElseIf date1 < date_END And Format(date_STT, "dd") <> "01" And Format(date_END + 1, "dd") <> "01" And date1 + 1 = date2 Then '情形六：多日+多日
    A1_STT = date_STT: A1_END = date1: iDays1 = A1_END - A1_STT + 1
    iDays2 = 0: imonths = 0
    A3_STT = date2: A3_END = date_END: iDays3 = A3_END - A3_STT + 1
    If FlagStat = "AVE" Then FlagStat = "AVE_SUM"
    If Year(A3_STT) <> Year(A1_STT) Then A3IsLastYearData = True
  
  End If

  If iDays1 <> 0 Then
    temp_STT = Format(A1_STT, "mm-dd"): temp_END = Format(A1_END, "mm-dd")
    '**************** 2024-2-28重要修复：强行把A1结束日期"02-28"修改为"02-29"，因为最终结果会利用ADays1()计算，不涉及到iDays1******************
    If temp_END = "02-28" Then temp_END = "02-29"
    Call StatMulYearDayDataExt(SelField_Day, BYear, EYear, temp_STT, temp_END, A1IsLastYearData, AAA1(), ADays1(), FlagStat, Min, Max)
  End If
  If iDays2 <> 0 Then
    temp_STT = Format(A2_STT, "mm"): temp_END = Format(A2_END, "mm")
    If A2IsLastYearData = False Then
      Call StatMulYearMonDataExt(SelField_Mon, BYear, EYear, temp_STT, temp_END, imonths, A2IsLastYearData, AAA2(), ADays2(), FlagStat, iDays2, FlagMonMiss(), SelField_Cnts) ', FlagLimit)，判断读取的月值有效数据量是否=iDays2
    ElseIf A2IsLastYearData = True Then
      Call StatMulYearMonDataExt(SelField_Mon, BYear + 1, EYear + 1, temp_STT, temp_END, imonths, A2IsLastYearData, AAA2(), ADays2(), FlagStat, iDays2, FlagMonMiss(), SelField_Cnts) ', FlagLimit)
    End If
  End If
  If iDays3 <> 0 Then
    temp_STT = Format(A3_STT, "mm-dd"): temp_END = Format(A3_END, "mm-dd")
    If A3IsLastYearData = False Then
      Call StatMulYearDayDataExt(SelField_Day, BYear, EYear, temp_STT, temp_END, A3IsLastYearData, AAA3(), ADays3(), FlagStat, Min, Max)
    ElseIf A3IsLastYearData = True Then
      Call StatMulYearDayDataExt(SelField_Day, BYear + 1, EYear + 1, temp_STT, temp_END, A3IsLastYearData, AAA3(), ADays3(), FlagStat, Min, Max)
    End If
  End If

  For j = BYear To EYear
    For i = 1 To ExtStaNum
      sngSum = 0: iNum = 0: sngMin = 999999: sngMax = -9999: idays = 0
      If FlagStat = "AVE" Then
        If iDays1 <> 0 Then Result(i, j) = AAA1(i, j)
        If iDays2 <> 0 Then Result(i, j) = AAA2(i, j)
      ElseIf FlagStat = "AVE_SUM" Then
'        iDays1 = ADays1(i, j): iDays2 = ADays2(i, j): iDays3 = ADays3(i, j)
        If AAA1(i, j) <> -9999 Then sngSum = sngSum + AAA1(i, j): idays = idays + ADays1(i, j)
        If AAA2(i, j) <> -9999 Then sngSum = sngSum + AAA2(i, j): idays = idays + ADays2(i, j)
        If AAA3(i, j) <> -9999 Then sngSum = sngSum + AAA3(i, j): idays = idays + ADays3(i, j)
        If idays <> 0 Then Result(i, j) = sngSum / idays '求均值
      ElseIf FlagStat = "MAX" Then '求大值
        If AAA1(i, j) <> -9999 And AAA1(i, j) >= sngMax Then sngMax = AAA1(i, j)
        If AAA2(i, j) <> -9999 And AAA2(i, j) >= sngMax Then sngMax = AAA2(i, j)
        If AAA3(i, j) <> -9999 And AAA3(i, j) >= sngMax Then sngMax = AAA3(i, j)
        If sngMax <> -9999 Then Result(i, j) = sngMax
      ElseIf FlagStat = "MIN" Then '求小值
        If AAA1(i, j) <> -9999 And AAA1(i, j) <= sngMin Then sngMin = AAA1(i, j)
        If AAA2(i, j) <> -9999 And AAA2(i, j) <= sngMin Then sngMin = AAA2(i, j)
        If AAA3(i, j) <> -9999 And AAA3(i, j) <= sngMin Then sngMin = AAA3(i, j)
        If sngMin <> 999999 Then Result(i, j) = sngMin
      ElseIf FlagStat = "SUM" Or FlagStat = "DAYS" Then '求和值
        If AAA1(i, j) <> -9999 Then sngSum = sngSum + AAA1(i, j): iNum = iNum + 1
        If AAA2(i, j) <> -9999 Then sngSum = sngSum + AAA2(i, j): iNum = iNum + 1
        If AAA3(i, j) <> -9999 Then sngSum = sngSum + AAA3(i, j): iNum = iNum + 1
        If iNum <> 0 Then Result(i, j) = sngSum
      End If
      
      '2022-03-28对结果进行四舍五入
      If Result(i, j) <> -9999 Then
'        Result(i, j) = Round(Result(i, j), 1)
        Result(i, j) = Format(Result(i, j), "0.0")
      End If
      
      '*****************2021-08-27添加：若该年份<站点起始年份，则输出结果为-9999*******************
      If j < ExtStaInfo(i).YearSTT Then Result(i, j) = -9999
    Next i
  
  
    '************************************2022-04-08修改月值不存在时查询结果的错误bug'************************************
    '************************************2022-04-01修改月值不存在时查询结果的错误bug'************************************
    '************************************2017-04-10修复月值不存在时查询结果的错误bug'************************************

      If FlagMonMiss(j) = True Then
        tempFlagStat = FlagStat
        If tempFlagStat = "AVE_SUM" Then tempFlagStat = "AVE"
        
        '重新查询该年度BMMDD和EMMDD之间的统计值：
        tempBDate = CStr(j) & "-" & BMMDD
          

  
        If EMMDD <> "02-29" Then
          If crossYear = False Then tempEDate = CStr(j) & "-" & EMMDD
          If crossYear = True Then tempEDate = CStr(j + 1) & "-" & EMMDD
        '2024-2-27 添加对闰年2-29的判断
        ElseIf EMMDD = "02-29" Then
          If crossYear = False Then
            If j Mod 4 = 0 Then tempEDate = CStr(j) & "-02-29"
            If j Mod 4 <> 0 Then tempEDate = CStr(j) & "-02-28"
          ElseIf crossYear = True Then
            If (j + 1) Mod 4 = 0 Then tempEDate = CStr(j + 1) & "-02-29"
            If (j + 1) Mod 4 <> 0 Then tempEDate = CStr(j + 1) & "-02-28"
          End If
        End If
        
        If tempFlagStat = "MAX" Or tempFlagStat = "MIN" Then '统计量：最大值、最小值查询
          ReDim tempBBB(1 To ExtStaNum), tempBBBDate(1 To ExtStaNum)
          For i = 1 To ExtStaNum
            tempBBB(i) = -9999: tempBBBDate(i) = CDate("1899-09-09")
          Next i
          Call StatExtremeDayDataExt(SelField_Day, tempBDate, tempEDate, tempBBB(), tempBBBDate(), tempFlagStat, Min, Max)
        
        Else '其他统计量：条件日数，非条件日数查询的平均值、累计值查询
          ReDim tempBBB(1 To ExtStaNum)
          For i = 1 To ExtStaNum
            tempBBB(i) = -9999
            If (tempFlagStat = "SUM" Or tempFlagStat = "DAYS") And FlagLimit = True Then '条件日数或条件累计值查询，无符合条件的都为0
              If j >= TempYearSTT(i) Then tempBBB(i) = 0 '当年(iyear)年份已经有资料了，则条件日数默认为0
            End If
          Next i
          Call StatMulDayDataExt(SelField_Day, tempBDate, tempEDate, tempBBB(), tempFlagStat, Min, Max)
        End If
        
        For i = 1 To ExtStaNum
          Result(i, j) = tempBBB(i)
        Next i
      End If

    '************************************2017-04-10修复月值不存在时查询结果的错误bug'************************************
  Next j
  
End Sub
''''''************************循环查询所有年份（Byear-Eyear）各站值（不包括极值出现年份）：结束  ***************************
''''''*************************************************************************************************************************




'统计某年度多日统计值（数据表名，要素字段名，开始日期，结束日期，输出统计值，统计量，条件最小值，条件最大值）-不再输入IsLastYearData、不再输出条件日数（2017-3-10）
Public Sub StatMulDayDataExt(SelField$, BDate$, ByVal EDate$, Result!(), FlagStat$, Min!, Max!)
  Dim i%, j%, k%
  Dim iYear%
  Dim LastSTA$

    TableName = "SURF_CLI_MUL_DAY"
       
    If FlagStat = "AVE" Then '平均值
      strSQL = "select TA.stacode STA,sum(TA." & SelField & ")/count(TA." & SelField & ") AAA"
    
    ElseIf FlagStat = "SUM" Then '累计值
      strSQL = "select TA.stacode STA,sum(TA." & SelField & ") AAA"
        
    ElseIf FlagStat = "MAX" Then '最大值
      strSQL = "select TA.stacode STA,max(TA." & SelField & ") AAA"
      
    ElseIf FlagStat = "MIN" Then '最小值
      strSQL = "select TA.stacode STA,min(TA." & SelField & ") AAA"
    
    ElseIf FlagStat = "DAYS" Then '条件日数查询
      strSQL = "select TA.stacode STA,count(TA." & SelField & ") AAA"
    
    End If
    strSQL = strSQL & " from " & TableName & " TA where TA.ddate between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
    
    If Not (Min = -9999 And Max = 999999) Then '***********条件查询***********
      strSQL = strSQL & " and TA." & SelField & ">=" & Min & " and TA." & SelField & "<=" & Max & ""
    End If
    
    strSQL = strSQL & " and " & strSURFExt
    strSQL = strSQL & " group by TA.stacode"
    strSQL = strSQL & " order by STA"
    
      
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  If ORARst.EOF <> True Then
    
    ORARst.MoveFirst
    '将该年度所有站点的统计值存入变量数组
    
    k = 1: LastSTA = ""
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To ExtStaNum
          If ORARst!STA = ExtStaInfo(i).stacode Then
             If Not IsNull(ORARst!AAA) Then Result(i) = Format(ORARst!AAA, "0.0")
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i) = Format(ORARst!AAA, "0.0")
      End If
           
      ORARst.MoveNext
    Loop
  End If
  ORARst.Close

End Sub






'**************************************************任意时段统计模块************************************************
'该子过程只提供最大值、最小值查询，条件查询无最大值最小值查询功能，即不可能是条件查询
'统计某年度某时段极值及出现日起（数据表名，要素字段名，开始日期，结束日期，是否属于上年度结果，输出统计值，输出日期，统计量）
Public Sub StatExtremeDayDataExt(SelField$, BDate$, ByVal EDate$, Result!(), Result2() As Date, FlagStat$, Min!, Max!)  '不再输入IsLastYearData(2017-3-10)
  Dim i%, j%, k%
  Dim LastSTA$


    TableName = "SURF_CLI_MUL_DAY"

      
    strSQL = "select TA.stacode STA,TB.AAA,TA.ddate dDate from " & TableName & " TA,"
    If FlagStat = "MAX" Then '最大值
      strSQL = strSQL & " (select TA.stacode STA,max(TA." & SelField & ") AAA"
    ElseIf FlagStat = "MIN" Then '最小值
      strSQL = strSQL & " (select TA.stacode STA,min(TA." & SelField & ") AAA"
    End If
    
    strSQL = strSQL & " from " & TableName & " TA"
    strSQL = strSQL & " where TA.ddate between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
    If Not (Min = -9999 And Max = 999999) Then '***********条件查询***********
      strSQL = strSQL & " and TA." & SelField & ">=" & Min & " and TA." & SelField & "<=" & Max & ""
    End If
    
    strSQL = strSQL & " and " & strSURFExt
    strSQL = strSQL & " group by TA.stacode) TB"
    
    strSQL = strSQL & " where TA.stacode=TB.STA and TA." & SelField & "=TB.AAA"
    strSQL = strSQL & " and ddate between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
    strSQL = strSQL & " order by STA,dDate"
  
      
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  If ORARst.EOF <> True Then
    
    ORARst.MoveFirst
    '将该年度所有站点的统计值存入变量数组
    
    k = 1: LastSTA = ""
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To ExtStaNum
          If ORARst!STA = ExtStaInfo(i).stacode Then
            If Not IsNull(ORARst!AAA) Then
              Result(i) = Format(ORARst!AAA, "0.0")
              Result2(i) = Format(ORARst!ddate, "yyyy-mm-dd")
            End If
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then
          Result(i) = Format(ORARst!AAA, "0.0")
          Result2(i) = Format(ORARst!ddate, "yyyy-mm-dd")
        End If
      End If
           
      ORARst.MoveNext
    Loop
  End If
  ORARst.Close


End Sub





'合计多年份-多日统计值（数据表名，要素字段名，开始年度，结束年度，开始月日，结束月日，是否属于上年度结果，输出统计值，输出日数，统计量，条件最小值，条件最大值）
Public Sub StatMulYearDayDataExt(SelField$, BYear%, EYear%, BMMDD$, EMMDD$, IsLastYearData As Boolean, Result!(), Result2%(), FlagStat$, Min!, Max!)
'如果不跨年度，则结果可能是本年度，也可能是上年度的（IsLastYearData仅用来区别此不跨年度情况）；如果跨年度，则第二次查询结果肯定是属于上一年度的；
  Dim BYear2%, EYear2%
  Dim BMMDD2$, EMMDD2$
  Dim ABB1!(), ABB2!() '存放两次取出的结果
  Dim iDays1%(), iDays2%() '存放两次读取的多个月份的天数
  Dim idays% '两次读取后天数的总和
  Dim i%, j%, k%
  Dim LastSTA$
  Dim iNum%, sngSum!, sngMax!, sngMin '求和的个数，和数，最大值、最小值
  
    TableName = "SURF_CLI_MUL_DAY"
  
  '= = = = = = = = = = = = = = = = = = 同年度：读一次，取出结果 = = = = = = = = = = = = = = = = = =
  If BMMDD <= EMMDD Then
  
          
        If FlagStat = "AVE" Then '平均值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,sum(TA." & SelField & ")/count(TA." & SelField & ") AAA"
           
        ElseIf FlagStat = "AVE_SUM" Then '累计值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,sum(TA." & SelField & ") AAA,count(TA." & SelField & ") DAYS"
          
        ElseIf FlagStat = "SUM" Then '累计值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,sum(TA." & SelField & ") AAA"
            
        ElseIf FlagStat = "MAX" Then '最大值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,max(TA." & SelField & ") AAA"
          
        ElseIf FlagStat = "MIN" Then '最小值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,min(TA." & SelField & ") AAA"
        
        ElseIf FlagStat = "DAYS" Then '条件日数查询
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,count(TA." & SelField & ") AAA"
        End If
        
        strSQL = strSQL & " from " & TableName & " TA"
        strSQL = strSQL & " where to_char(ddate,'yyyy') between '" & BYear & "' and '" & EYear & "'"
        strSQL = strSQL & " and to_char(ddate,'mm-dd') between '" & BMMDD & "' and '" & EMMDD & "'"
        
        If Not (Min = -9999 And Max = 999999) Then '***********条件查询***********
          strSQL = strSQL & " and TA." & SelField & ">=" & Min & " and TA." & SelField & "<=" & Max & ""
        End If
       
        strSQL = strSQL & " and " & strSURFExt
        strSQL = strSQL & " group by to_char(TA.ddate,'yyyy'),TA.stacode"
        strSQL = strSQL & " order by STA,YER"
        
    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
    If ORARst.EOF <> True Then
      
      ORARst.MoveFirst
      '将该年度所有站点的统计值存入变量数组
        
      k = 1: LastSTA = ""
      Do Until ORARst.EOF
        If ORARst!STA <> LastSTA Then
          LastSTA = ORARst!STA
          For i = k To ExtStaNum
            If ORARst!STA = ExtStaInfo(i).stacode Then
              If IsLastYearData = False Then
                If Not IsNull(ORARst!AAA) Then
                  Result(i, ORARst!YER) = ORARst!AAA
                  If FlagStat = "AVE_SUM" Then Result2(i, ORARst!YER) = ORARst!DAYS
                End If
              ElseIf IsLastYearData = True Then
                If Not IsNull(ORARst!AAA) Then
                  Result(i, ORARst!YER - 1) = ORARst!AAA
                  If FlagStat = "AVE_SUM" Then Result2(i, ORARst!YER - 1) = ORARst!DAYS
                End If
              End If
              k = i + 1: Exit For
            End If
          Next i
        ElseIf ORARst!STA = LastSTA Then
          If IsLastYearData = False Then
            If Not IsNull(ORARst!AAA) Then
              Result(i, ORARst!YER) = ORARst!AAA
              If FlagStat = "AVE_SUM" Then Result2(i, ORARst!YER) = ORARst!DAYS
            End If
          ElseIf IsLastYearData = True Then
            If Not IsNull(ORARst!AAA) Then
              Result(i, ORARst!YER - 1) = ORARst!AAA
              If FlagStat = "AVE_SUM" Then Result2(i, ORARst!YER - 1) = ORARst!DAYS
            End If
          End If
        End If
               
        ORARst.MoveNext
      Loop
    End If
    ORARst.Close




  '= = = = = = = = = = = = = = = = = = 跨年度：读两次，合并结果 = = = = = = = = = = = = = = = = = =
  ElseIf BMMDD > EMMDD Then
  
    ReDim ABB1(1 To ExtStaNum, BYear To EYear), ABB2(1 To ExtStaNum, BYear To EYear) '存放两次取出的结果
    ReDim iDays1(1 To ExtStaNum, BYear To EYear), iDays2(1 To ExtStaNum, BYear To EYear) '存放两次取出的结果
    For j = BYear To EYear
      For i = 1 To ExtStaNum
        ABB1(i, j) = -9999: ABB2(i, j) = -9999
        iDays1(i, j) = 0: iDays2(i, j) = 0
      Next i
    Next j

    '第一次查询
    BYear2 = BYear: EYear2 = EYear
    BMMDD2 = BMMDD: EMMDD2 = "12-31"

        If FlagStat = "AVE" Then '平均值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,sum(TA." & SelField & ")/count(TA." & SelField & ") AAA,count(TA." & SelField & ") DAYS"
           
        ElseIf FlagStat = "AVE_SUM" Then '累计值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,sum(TA." & SelField & ") AAA,count(TA." & SelField & ") DAYS"
          
        ElseIf FlagStat = "SUM" Then '累计值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,sum(TA." & SelField & ") AAA"
            
        ElseIf FlagStat = "MAX" Then '最大值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,max(TA." & SelField & ") AAA"
          
        ElseIf FlagStat = "MIN" Then '最小值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,min(TA." & SelField & ") AAA"
        
        ElseIf FlagStat = "DAYS" Then '条件日数查询
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,count(TA." & SelField & ") AAA"
        End If
        strSQL = strSQL & " from " & TableName & " TA"
        strSQL = strSQL & " where to_char(ddate,'yyyy') between '" & BYear2 & "' and '" & EYear2 & "'"
        strSQL = strSQL & " and to_char(ddate,'mm-dd') between '" & BMMDD2 & "' and '" & EMMDD2 & "'"
        
        If Not (Min = -9999 And Max = 999999) Then '***********条件查询***********
          strSQL = strSQL & " and TA." & SelField & ">=" & Min & " and TA." & SelField & "<=" & Max & ""
        End If
        strSQL = strSQL & " and " & strSURFExt
        
        strSQL = strSQL & " group by to_char(TA.ddate,'yyyy'),TA.stacode"
        strSQL = strSQL & " order by STA,YER"


    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
    If ORARst.EOF <> True Then
      ORARst.MoveFirst
      '将该年度所有站点的统计值存入变量数组

      k = 1: LastSTA = ""
      Do Until ORARst.EOF
        If ORARst!STA <> LastSTA Then
          LastSTA = ORARst!STA
          For i = k To ExtStaNum
            If ORARst!STA = ExtStaInfo(i).stacode Then
              If Not IsNull(ORARst!AAA) Then
                ABB1(i, ORARst!YER) = ORARst!AAA
                If FlagStat = "AVE" Or FlagStat = "AVE_SUM" Then iDays1(i, ORARst!YER) = ORARst!DAYS
              End If
              k = i + 1: Exit For
            End If
          Next i
        ElseIf ORARst!STA = LastSTA Then
          If Not IsNull(ORARst!AAA) Then
            ABB1(i, ORARst!YER) = ORARst!AAA
            If FlagStat = "AVE" Or FlagStat = "AVE_SUM" Then iDays1(i, ORARst!YER) = ORARst!DAYS
          End If
        End If

        ORARst.MoveNext
      Loop
    End If
    ORARst.Close

    '第二次查询
    BYear2 = BYear + 1: EYear2 = EYear + 1
    BMMDD2 = "01-01": EMMDD2 = EMMDD
    
    
        If FlagStat = "AVE" Then '平均值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,sum(TA." & SelField & ")/count(TA." & SelField & ") AAA,count(TA." & SelField & ") DAYS"
           
        ElseIf FlagStat = "AVE_SUM" Then '累计值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,sum(TA." & SelField & ") AAA,count(TA." & SelField & ") DAYS"
          
        ElseIf FlagStat = "SUM" Then '累计值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,sum(TA." & SelField & ") AAA"
            
        ElseIf FlagStat = "MAX" Then '最大值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,max(TA." & SelField & ") AAA"
          
        ElseIf FlagStat = "MIN" Then '最小值
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,min(TA." & SelField & ") AAA"
        
        ElseIf FlagStat = "DAYS" Then '条件日数查询
          strSQL = "select TA.stacode STA,to_char(TA.ddate,'yyyy') YER,count(TA." & SelField & ") AAA"
        End If
        strSQL = strSQL & " from " & TableName & " TA"
        strSQL = strSQL & " where to_char(ddate,'yyyy') between '" & BYear2 & "' and '" & EYear2 & "'"
        strSQL = strSQL & " and to_char(ddate,'mm-dd') between '" & BMMDD2 & "' and '" & EMMDD2 & "'"
        
        If Not (Min = -9999 And Max = 999999) Then '***********条件查询***********
          strSQL = strSQL & " and TA." & SelField & ">=" & Min & " and TA." & SelField & "<=" & Max & ""
        End If
        
        strSQL = strSQL & " and " & strSURFExt
        strSQL = strSQL & " group by to_char(TA.ddate,'yyyy'),TA.stacode"
        strSQL = strSQL & " order by STA,YER"
    

    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
    If ORARst.EOF <> True Then
      ORARst.MoveFirst
      '将该年度所有站点的统计值存入变量数组

      k = 1: LastSTA = ""
      Do Until ORARst.EOF
        If ORARst!STA <> LastSTA Then
          LastSTA = ORARst!STA
          For i = k To ExtStaNum
            If ORARst!STA = ExtStaInfo(i).stacode Then
              If Not IsNull(ORARst!AAA) Then
                ABB2(i, ORARst!YER - 1) = ORARst!AAA
                If FlagStat = "AVE" Or FlagStat = "AVE_SUM" Then iDays2(i, ORARst!YER - 1) = ORARst!DAYS
              End If
              k = i + 1: Exit For
            End If
          Next i
        ElseIf ORARst!STA = LastSTA Then
          If Not IsNull(ORARst!AAA) Then
            ABB2(i, ORARst!YER - 1) = ORARst!AAA
            If FlagStat = "AVE" Or FlagStat = "AVE_SUM" Then iDays2(i, ORARst!YER - 1) = ORARst!DAYS
          End If
        End If

        ORARst.MoveNext
      Loop
    End If
    ORARst.Close

    '统计两次取出的结果（平均、最大、最小）
    For j = BYear To EYear

      For i = 1 To ExtStaNum
        sngSum = 0: iNum = 0: sngMin = 999999: sngMax = -9999: idays = 0
        If FlagStat = "AVE" Then
          If ABB1(i, j) <> -9999 Then sngSum = sngSum + ABB1(i, j) * iDays1(i, j): idays = idays + iDays1(i, j)
          If ABB2(i, j) <> -9999 Then sngSum = sngSum + ABB2(i, j) * iDays2(i, j): idays = idays + iDays2(i, j)
          If idays <> 0 Then '求均值
            Result(i, j) = sngSum / idays
          End If
        ElseIf FlagStat = "AVE_SUM" Then
          If ABB1(i, j) <> -9999 Then sngSum = sngSum + ABB1(i, j): idays = idays + iDays1(i, j)
          If ABB2(i, j) <> -9999 Then sngSum = sngSum + ABB2(i, j): idays = idays + iDays2(i, j)
          If idays <> 0 Then '求均值
            Result(i, j) = sngSum: Result2(i, j) = idays
          End If

        ElseIf FlagStat = "MAX" Then  '求大值
          If ABB1(i, j) <> -9999 And ABB1(i, j) >= sngMax Then sngMax = ABB1(i, j)
          If ABB2(i, j) <> -9999 And ABB2(i, j) >= sngMax Then sngMax = ABB2(i, j)
          If sngMax <> -9999 Then Result(i, j) = sngMax
        ElseIf FlagStat = "MIN" Then  '求小值
          If ABB1(i, j) <> -9999 And ABB1(i, j) <= sngMin Then sngMin = ABB1(i, j)
          If ABB2(i, j) <> -9999 And ABB2(i, j) <= sngMin Then sngMin = ABB2(i, j)
          If sngMin <> 999999 Then Result(i, j) = sngMin
        ElseIf FlagStat = "SUM" Or FlagStat = "DAYS" Then '求和值
          If ABB1(i, j) <> -9999 Then sngSum = sngSum + ABB1(i, j): iNum = iNum + 1
          If ABB2(i, j) <> -9999 Then sngSum = sngSum + ABB2(i, j): iNum = iNum + 1
          If iNum <> 0 Then Result(i, j) = sngSum
        End If
      Next i
    Next j
    
  End If
End Sub


' 2021-12-29：合计多年份-多月统计值（数据表名，要素月值字段名，开始年度，结束年度，开始月份，结束月份，月份数量，是否属于上年结果，输出统计值，输出日数，统计量，是否条件查询）
Private Sub StatMulYearMonDataExt(SelField$, BYear%, EYear%, BMM$, EMM$, imonths%, IsLastYearData As Boolean, Result!(), Result2%(), FlagStat$, iMonsDays%, FlagMonMiss() As Boolean, SelField_Cnts$) ', FlagLimit As Boolean)'多月应有的总日数iMonsDays
'如果不跨年度，则结果可能是本年度，也可能是上年度的（IsLastYearData仅用来区别此不跨年度情况）；如果跨年度，则第二次查询结果肯定是属于上一年度的；
  Dim BYear2%, EYear2%
  Dim BMM2$, EMM2$
  Dim ABB1!(), ABB2!() '存放两次取出的结果
  Dim iDays1%(), iDays2%() '存放两次读取的多个月份的天数
  Dim idays% '两次读取后天数的总和
  Dim i%, j%, k%
  Dim LastSTA$
  Dim iNum%, sngSum!, sngMax!, sngMin '求和的个数，和数，最大值、最小值

  '******** 2022-04-01判断某年的月统计结果是否缺失月数据，若缺则标记FlagMonMiss()=True ********
  Dim iCnts%() '每个站点每年的有效数据数量
  Dim iCnts1%(), iCnts2%() '存放两次读取的多个月份的有效天数


  Dim Query$ '查询一次或两次的标记
  If imonths <= 12 Then
    If BMM <= EMM Then Query = "Once"
    If BMM > EMM Then Query = "Twice"
  ElseIf imonths > 12 Then
    Query = "Twice"
  End If

    TableName = "SURF_CLI_MUL_MON"

  '= = = = = = = = = = = = = = = = = = 同年度：读一次，取出结果 = = = = = = = = = = = = = = = = = =
  If Query = "Once" Then
    ReDim iCnts(1 To ExtStaNum, BYear To EYear)


        If FlagStat = "AVE" Then '求平均值
          strSQL = "select TA.stacode STA,TA.iyear YER"
          strSQL = strSQL & ",sum(TA." & SelField & "*TA." & SelField_Cnts & ")/sum(TA." & SelField_Cnts & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "AVE_SUM" Then '求总和，及总天数--用于计算总的平均数
          strSQL = "select TA.stacode STA,TA.iyear YER"
          strSQL = strSQL & ",sum(TA." & SelField & "*TA." & SelField_Cnts & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") DAYS"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "MAX" Then '最大值
          strSQL = "select TA.stacode STA,TA.iyear YER,max(TA." & SelField & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "MIN" Then '最小值
          strSQL = "select TA.stacode STA,TA.iyear YER,min(TA." & SelField & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "SUM" Or FlagStat = "DAYS" Then '累计值
          strSQL = "select TA.stacode STA,TA.iyear YER,sum(TA." & SelField & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"
        End If

        strSQL = strSQL & " from " & TableName & " TA"
        strSQL = strSQL & " where iyear between " & BYear & " and " & EYear & ""
        strSQL = strSQL & " and trim(to_char(TA.imonth,'00')) between '" & BMM & "' and '" & EMM & "'"

        strSQL = strSQL & " and " & strSURFExt
        strSQL = strSQL & " group by TA.iyear,TA.stacode"
        strSQL = strSQL & " order by STA,YER"

    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
    If ORARst.EOF <> True Then
      ORARst.MoveFirst
      '将该年度所有站点的统计值存入变量数组

      ORARst.MoveFirst
      k = 1: LastSTA = ""

      Do Until ORARst.EOF
        If ORARst!STA <> LastSTA Then
          LastSTA = ORARst!STA
          For i = k To ExtStaNum
            If ORARst!STA = ExtStaInfo(i).stacode Then
              If IsLastYearData = False Then
                If Not IsNull(ORARst!AAA) Then
                  Result(i, ORARst!YER) = ORARst!AAA
                  If FlagStat = "AVE_SUM" Then Result2(i, ORARst!YER) = ORARst!DAYS
                  iCnts(i, ORARst!YER) = ORARst!iCnts
                End If
              ElseIf IsLastYearData = True Then
                If Not IsNull(ORARst!AAA) Then
                  Result(i, ORARst!YER - 1) = ORARst!AAA
                  If FlagStat = "AVE_SUM" Then Result2(i, ORARst!YER - 1) = ORARst!DAYS
                  iCnts(i, ORARst!YER) = ORARst!iCnts
                End If
              End If
              k = i + 1: Exit For
            End If
          Next i
        ElseIf ORARst!STA = LastSTA Then
          If IsLastYearData = False Then
            If Not IsNull(ORARst!AAA) Then
              Result(i, ORARst!YER) = ORARst!AAA
              If FlagStat = "AVE_SUM" Then Result2(i, ORARst!YER) = ORARst!DAYS
              iCnts(i, ORARst!YER) = ORARst!iCnts
            End If
          ElseIf IsLastYearData = True Then
            If Not IsNull(ORARst!AAA) Then
              Result(i, ORARst!YER - 1) = ORARst!AAA
              If FlagStat = "AVE_SUM" Then Result2(i, ORARst!YER - 1) = ORARst!DAYS
              iCnts(i, ORARst!YER) = ORARst!iCnts
            End If
         End If
        End If

        ORARst.MoveNext
      Loop

    End If
    ORARst.Close


  '= = = = = = = = = = = = = = = = = = 跨年度：读两次，合并结果 = = = = = = = = = = = = = = = = = =
  ElseIf Query = "Twice" Then
    ReDim ABB1(1 To ExtStaNum, BYear To EYear), ABB2(1 To ExtStaNum, BYear To EYear) '存放两次取出的结果
    ReDim iDays1(1 To ExtStaNum, BYear To EYear), iDays2(1 To ExtStaNum, BYear To EYear) '存放两次取出的结果
    ReDim iCnts(1 To ExtStaNum, BYear To EYear), iCnts1(1 To ExtStaNum, BYear To EYear), iCnts2(1 To ExtStaNum, BYear To EYear)

    For j = BYear To EYear
      For i = 1 To ExtStaNum
        ABB1(i, j) = -9999: ABB2(i, j) = -9999
        iDays1(i, j) = 0: iDays2(i, j) = 0
      Next i
    Next j

    '第一次查询
    BYear2 = BYear: EYear2 = EYear
    BMM2 = BMM: EMM2 = "12"

        If FlagStat = "AVE" Then  '平均值
          strSQL = "select TA.stacode STA,TA.iyear YER"
'          SelField_Cnts = Replace(SelField, "AVE", "CNTS")
          strSQL = strSQL & ",sum(TA." & SelField & "*TA." & SelField_Cnts & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") DAYS"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "AVE_SUM" Then
          strSQL = "select TA.stacode STA,TA.iyear YER"
'          SelField_Cnts = Replace(SelField, "AVE", "CNTS")
          strSQL = strSQL & ",sum(TA." & SelField & "*TA." & SelField_Cnts & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") DAYS"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "MAX" Then  '最大值
          strSQL = "select TA.stacode STA,TA.iyear YER,max(TA." & SelField & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "MIN" Then  '最小值
          strSQL = "select TA.stacode STA,TA.iyear YER,min(TA." & SelField & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "SUM" Or FlagStat = "DAYS" Then '累计值
          strSQL = "select TA.stacode STA,TA.iyear YER,sum(TA." & SelField & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"
        End If

        strSQL = strSQL & " from " & TableName & " TA"
        strSQL = strSQL & " where iyear between " & BYear2 & " and " & EYear2 & ""
        strSQL = strSQL & " and trim(to_char(TA.imonth,'00')) between '" & BMM2 & "' and '" & EMM2 & "'"

        strSQL = strSQL & " and " & strSURFExt
        strSQL = strSQL & " group by TA.iyear,TA.stacode"
        strSQL = strSQL & " order by STA,YER"


    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
    If ORARst.EOF <> True Then
      ORARst.MoveFirst
      '将该年度所有站点的统计值存入变量数组

      k = 1: LastSTA = ""
      Do Until ORARst.EOF
        If ORARst!STA <> LastSTA Then
          LastSTA = ORARst!STA
          For i = k To ExtStaNum
            If ORARst!STA = ExtStaInfo(i).stacode Then
              If Not IsNull(ORARst!AAA) Then
                ABB1(i, ORARst!YER) = ORARst!AAA
                If FlagStat = "AVE" Or FlagStat = "AVE_SUM" Then iDays1(i, ORARst!YER) = ORARst!DAYS
                iCnts1(i, ORARst!YER) = ORARst!iCnts
              End If
              k = i + 1: Exit For
            End If
          Next i
        ElseIf ORARst!STA = LastSTA Then
          If Not IsNull(ORARst!AAA) Then
            ABB1(i, ORARst!YER) = ORARst!AAA
            If FlagStat = "AVE" Or FlagStat = "AVE_SUM" Then iDays1(i, ORARst!YER) = ORARst!DAYS
            iCnts1(i, ORARst!YER) = ORARst!iCnts
          End If
        End If

        ORARst.MoveNext
      Loop
    End If
    ORARst.Close

    '第二次查询
    BYear2 = BYear + 1: EYear2 = EYear + 1
    BMM2 = "01": EMM2 = EMM

        If FlagStat = "AVE" Then  '平均值
          strSQL = "select TA.stacode STA,TA.iyear YER"
          strSQL = strSQL & ",sum(TA." & SelField & "*TA." & SelField_Cnts & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") DAYS"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "AVE_SUM" Then
          strSQL = "select TA.stacode STA,TA.iyear YER"
          strSQL = strSQL & ",sum(TA." & SelField & "*TA." & SelField_Cnts & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") DAYS"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "MAX" Then  '最大值
          strSQL = "select TA.stacode STA,TA.iyear YER,max(TA." & SelField & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "MIN" Then  '最小值
          strSQL = "select TA.stacode STA,TA.iyear YER,min(TA." & SelField & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"

        ElseIf FlagStat = "SUM" Or FlagStat = "DAYS" Then '累计值
          strSQL = "select TA.stacode STA,TA.iyear YER,sum(TA." & SelField & ") AAA"
          strSQL = strSQL & ",sum(TA." & SelField_Cnts & ") ICNTS"
        End If
        strSQL = strSQL & " from " & TableName & " TA"

        strSQL = strSQL & " where iyear between " & BYear2 & " and " & EYear2 & ""
        strSQL = strSQL & " and trim(to_char(TA.imonth,'00')) between '" & BMM2 & "' and '" & EMM2 & "'"
        strSQL = strSQL & " and " & strSURFExt
        strSQL = strSQL & " group by TA.iyear,TA.stacode"
        strSQL = strSQL & " order by STA,YER"


    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
    If ORARst.EOF <> True Then
      ORARst.MoveFirst
      '将该年度所有站点的统计值存入变量数组

      k = 1: LastSTA = ""
      Do Until ORARst.EOF
        If ORARst!STA <> LastSTA Then
          LastSTA = ORARst!STA
          For i = k To ExtStaNum
            If ORARst!STA = ExtStaInfo(i).stacode Then
              If Not IsNull(ORARst!AAA) Then
                ABB2(i, ORARst!YER - 1) = ORARst!AAA
                If FlagStat = "AVE" Or FlagStat = "AVE_SUM" Then iDays2(i, ORARst!YER - 1) = ORARst!DAYS
                iCnts2(i, ORARst!YER - 1) = ORARst!iCnts
              End If
              k = i + 1: Exit For
            End If
          Next i
        ElseIf ORARst!STA = LastSTA Then
          If Not IsNull(ORARst!AAA) Then
            ABB2(i, ORARst!YER - 1) = ORARst!AAA
            If FlagStat = "AVE" Or FlagStat = "AVE_SUM" Then iDays2(i, ORARst!YER - 1) = ORARst!DAYS
            iCnts2(i, ORARst!YER - 1) = ORARst!iCnts
          End If
        End If

        ORARst.MoveNext
      Loop
    End If
    ORARst.Close

    '统计两次取出的结果（平均、最大、最小）
    For j = BYear To EYear

      For i = 1 To ExtStaNum
        sngSum = 0: iNum = 0: sngMin = 999999: sngMax = -9999: idays = 0
        If FlagStat = "AVE" Then
          If ABB1(i, j) <> -9999 Then sngSum = sngSum + ABB1(i, j): idays = idays + iDays1(i, j)
          If ABB2(i, j) <> -9999 Then sngSum = sngSum + ABB2(i, j): idays = idays + iDays2(i, j)
          If idays <> 0 Then '求均值
            Result(i, j) = sngSum / idays
          End If
        ElseIf FlagStat = "AVE_SUM" Then
          If ABB1(i, j) <> -9999 Then sngSum = sngSum + ABB1(i, j): idays = idays + iDays1(i, j)
          If ABB2(i, j) <> -9999 Then sngSum = sngSum + ABB2(i, j): idays = idays + iDays2(i, j)
          If idays <> 0 Then '求均值
            Result(i, j) = sngSum: Result2(i, j) = idays
          End If

        ElseIf FlagStat = "MAX" Then  '求大值
          If ABB1(i, j) <> -9999 And ABB1(i, j) >= sngMax Then sngMax = ABB1(i, j)
          If ABB2(i, j) <> -9999 And ABB2(i, j) >= sngMax Then sngMax = ABB2(i, j)
          If sngMax <> -9999 Then Result(i, j) = sngMax
        ElseIf FlagStat = "MIN" Then  '求小值
          If ABB1(i, j) <> -9999 And ABB1(i, j) <= sngMin Then sngMin = ABB1(i, j)
          If ABB2(i, j) <> -9999 And ABB2(i, j) <= sngMin Then sngMin = ABB2(i, j)
          If sngMin <> 999999 Then Result(i, j) = sngMin
        ElseIf FlagStat = "SUM" Or FlagStat = "DAYS" Then '求和值
          If ABB1(i, j) <> -9999 Then sngSum = sngSum + ABB1(i, j): iNum = iNum + 1
          If ABB2(i, j) <> -9999 Then sngSum = sngSum + ABB2(i, j): iNum = iNum + 1
          If iNum <> 0 Then Result(i, j) = sngSum
        End If
        iCnts(i, j) = iCnts1(i, j) + iCnts2(i, j)
      Next i

    Next j

  End If


  '******** 2022-04-01判断某年的月统计结果是否缺失月数据，若缺则标记FlagMonMiss()=True ********
  For j = BYear To EYear
    If IsLastYearData = False Then
      FlagMonMiss(j) = True
    ElseIf IsLastYearData = True Then
      FlagMonMiss(j - 1) = True
    End If

    For i = 1 To ExtStaNum
      '只要有一个站不缺月值，那么设定当年该时段不缺月值
      'iMonsDays只按照平年来计算日数，而iCnts分平年和闰年统计日数，因此只需要iCnt>=iMonsDays即可（2023-1-6）
      'iMonsDays分平年/闰年来计算日数，iCnts也分平年和闰年统计日数，因此只需要iCnt>=iMonsDays即可（2024-2-28）
      If EMM >= "02" And (BMM <= "02" Or BMM > EMM) Then '不跨年和跨年的两种情形包含了2月，需要判断平闰年
        If j Mod 4 = 0 Then '闰年
          If iCnts(i, j) >= iMonsDays Then
            If IsLastYearData = False Then
              FlagMonMiss(j) = False
            ElseIf IsLastYearData = True Then
              FlagMonMiss(j - 1) = False
            End If

            Exit For
          End If

        ElseIf j Mod 4 <> 0 Then '平年
          If iCnts(i, j) >= iMonsDays Then
            If IsLastYearData = False Then
              FlagMonMiss(j) = False
            ElseIf IsLastYearData = True Then
              FlagMonMiss(j - 1) = False
            End If
            Exit For
          End If

        End If

      Else '不包含2月

        If iCnts(i, j) >= iMonsDays Then
          If IsLastYearData = False Then
            FlagMonMiss(j) = False
          ElseIf IsLastYearData = True Then
            FlagMonMiss(j - 1) = False
          End If
          Exit For
        End If
      End If

    Next i
  Next j


End Sub



'*********************************从极值表SURF_CLI_MUL_DAY_XTRM中统计历史日极值及出现日期*********************************
'读取某年度某时段极值：要素极值字段，开始年月，结束年月，是否跨年，输出值（一维），出现日期，统计量
Public Sub StatMulDayXtrmExt(SelField$, BMMDD$, EMMDD$, crossYear As Boolean, Result!(), Result2() As Date, FlagStat$, Min!, Max!)
  Dim Extreme_Max$, ExtremeMax_Year$ '需要统计最大值的对应极值表要素及年份字段名称
  Dim Extreme_Min$, ExtremeMin_Year$ '需要统计最小值的对应极值表要素及年份字段名称
  Dim BDate As Date, EDate As Date
  Dim i%, j%, k%
  Dim SelField2$ '要素极值对应的年份字段
  
  SelField2 = SelField & "YEAR"
  
  BDate = Format(("2001-" & BMMDD), "yyyy-mm-dd")
  If EMMDD = "02-29" Then
    EDate = Format("2001-02-28", "yyyy-mm-dd")
  Else
    EDate = Format(("2001-" & EMMDD), "yyyy-mm-dd")
  End If
  

    TableName = "SURF_CLI_MUL_DAY_XTRM"

  
    strSQL = "select TA.stacode STA,TB.AAA," & SelField2 & " || '-' || TA.imonth || '-' || ta.iday dDate"
    strSQL = strSQL & " from " & TableName & " TA,"
    If FlagStat = "MAX" Then strSQL = strSQL & " (select TA.stacode STA,max(TA." & SelField & ") AAA"
    If FlagStat = "MIN" Then strSQL = strSQL & " (select TA.stacode STA,min(TA." & SelField & ") AAA"
    strSQL = strSQL & " from " & TableName & " TA"

    If crossYear = False Then  '起止日期同年度
      strSQL = strSQL & " where (to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd'))"
      If Not (Min = -9999 And Max = 999999) Then '***********条件查询***********
        strSQL = strSQL & " and TA." & SelField & ">=" & Min & " and TA." & SelField & "<=" & Max & ""
      End If
        
      strSQL = strSQL & " and " & strSURFExt
      strSQL = strSQL & " group by TA.stacode) TB"
      
      strSQL = strSQL & " where TA.stacode=TB.STA and TA." & SelField & "=TB.AAA"
      strSQL = strSQL & " and (to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd'))"
      strSQL = strSQL & " order by STA,dDate"

    ElseIf crossYear = True Then '起止日期跨年度
      strSQL = strSQL & " where ((to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-31','yyyy-mm-dd'))"
      strSQL = strSQL & " or (to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('2001-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')))"
      If Not (Min = -9999 And Max = 999999) Then '***********条件查询***********
        strSQL = strSQL & " and TA." & SelField & ">=" & Min & " and TA." & SelField & "<=" & Max & ""
      End If
        
      strSQL = strSQL & " and " & strSURFExt
      strSQL = strSQL & " group by TA.stacode) TB"
      
      strSQL = strSQL & " where TA.stacode=TB.STA and TA." & SelField & "=TB.AAA"
      strSQL = strSQL & " and ((to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-31','yyyy-mm-dd'))"
      strSQL = strSQL & " or (to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('2001-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')))"
      strSQL = strSQL & " order by STA,dDate"
  
    End If
  
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    ORARst.MoveFirst
    k = 1
    Do Until ORARst.EOF
      For i = k To ExtStaNum
        If ORARst!STA = ExtStaInfo(i).stacode Then
          If Not IsNull(ORARst!AAA) Then
            Result(i) = Format(ORARst!AAA, "0.0")
            Result2(i) = Format(ORARst!ddate, "yyyy-mm-dd")
          End If
          k = i + 1: Exit For
        End If
      Next i
      ORARst.MoveNext
    Loop
  End If
  
  ORARst.Close
End Sub
'*****************************从极值表SURF_MUL_DAY_EXTREME中取出历史极值及出现日期：结束*****************************



'*******************************************2023-01-30 仅R、S、L_S三要素的累积量分日、月查询*******************************************
'*******************************************2022-05-12 任意时段多年平均值查询模块*******************************************
Public Sub QueryPeriodNormDataExt(SelField_Restat$, BMMDD$, EMMDD$, crossYear As Boolean, Result!(), ByVal FlagStat$, Min!, Max!)
  '如果结果将由几部分求平均获取，则：FlagStat由"AVE"替换为"AVE_SUM"
  Dim i%, j%, k%
  
  Dim date_STT As Date, date_END As Date   '开始日期、结束日期
  Dim date1 As Date, date2 As Date   '分割日期1、分割日期2
  Dim C1_STT As Date, C1_END As Date, C2_STT As Date, C2_END As Date, C3_STT As Date, C3_END As Date 'C1、C2、C3起止日期
  Dim temp_STT$, temp_END$ '起止月份或起止年月
  
  Dim iDays1%, iDays2%, iDays3%, idays% '被分割后的多段日期的日数，及总日数
  Dim CCC1!(), CCC2!(), CCC3!() ' 存放所有站点：多日结果、多个整月结果、多日结果
  Dim CDays1%(), CDays2%(), CDays3%() ' 存放有站点条件日数：多日结果、多个整月结果、多日结果
  Dim imonths% '多月的月份数
  
  Dim iNum%, sngSum!, sngMax!, sngMin '求和的个数，和数，最大值、最小值
  
  Dim TableNormDay$ '查询要素所在的数据表
  Dim TableNormMon$ '要素月值所在的数据表
  Dim TableNormYer$ '要素年值所在的数据表
  
  
  If EMMDD <> "02-29" Then
    If crossYear = False Then
      date_STT = CDate("2011-" & BMMDD)
      date_END = CDate("2011-" & EMMDD)
    ElseIf crossYear = True Then
      date_STT = CDate("2010-" & BMMDD)
      date_END = CDate("2011-" & EMMDD)
    End If
    
  '2024-2-27 添加对闰年2-29的判断
  ElseIf EMMDD = "02-29" Then
    If crossYear = False Then
      date_STT = CDate("2012-" & BMMDD)
      date_END = CDate("2012-" & EMMDD)
    ElseIf crossYear = True Then
      date_STT = CDate("2011-" & BMMDD)
      date_END = CDate("2012-" & EMMDD)
    End If
    
  End If
   
  
  date1 = CDate(Format(DateAdd("m", 1, date_STT), "yyyy-mm-01")) - 1
  date2 = CDate(Format(date_END, "yyyy-mm-01"))
  
  If date1 > date_END Or (date1 = date_END And date2 < date_STT) Then '情形一：多日
    C1_STT = date_STT: C1_END = date_END: iDays1 = C1_END - C1_STT + 1
    iDays2 = 0
    iDays3 = 0
    
    ReDim CCC1(1 To ExtStaNum, 1 To iDays1)
    For j = 1 To iDays1
      For i = 1 To ExtStaNum
        CCC1(i, j) = -9999
      Next i
    Next j

    
  ElseIf (date1 = date_END And date2 = date_STT) _
  Or (date1 < date_END And Format(date_STT, "dd") = "01" And Format(date_END + 1, "dd") = "01") Then '情形二：多月
    iDays1 = 0
    C2_STT = date_STT: C2_END = date_END: iDays2 = C2_END - C2_STT + 1: imonths = DateDiff("m", C2_STT, C2_END) + 1
    iDays3 = 0
  
    ReDim CCC2(1 To ExtStaNum, 1 To imonths)
    For j = 1 To imonths
      For i = 1 To ExtStaNum
        CCC2(i, j) = -9999
      Next i
    Next j
    '判断月值的读取是否跨年度
    If Year(C2_STT) < Year(C2_END) Then
      crossYear = True
    Else
      crossYear = False
    End If
  
  ElseIf date1 < date_END And Format(date_STT, "dd") = "01" And Format(date_END + 1, "dd") <> "01" Then '情形三：多月+多日
    iDays1 = 0
    C2_STT = date_STT: C2_END = date2 - 1: iDays2 = C2_END - C2_STT + 1: imonths = DateDiff("m", C2_STT, C2_END) + 1
    C3_STT = date2: C3_END = date_END: iDays3 = C3_END - C3_STT + 1
    
    ReDim CCC2(1 To ExtStaNum, 1 To imonths)
    For j = 1 To imonths
      For i = 1 To ExtStaNum
        CCC2(i, j) = -9999
      Next i
    Next j
    ReDim CCC3(1 To ExtStaNum, 1 To iDays3)
    For j = 1 To iDays3
      For i = 1 To ExtStaNum
        CCC3(i, j) = -9999
      Next i
    Next j
    '判断月值的读取是否跨年度
    If Year(C2_STT) < Year(C2_END) Then
      crossYear = True
    Else
      crossYear = False
    End If
    
  ElseIf date1 < date_END And Format(date_STT, "dd") <> "01" And Format(date_END + 1, "dd") = "01" Then '情形四：多日+多月
    C1_STT = date_STT: C1_END = date1: iDays1 = C1_END - C1_STT + 1
    C2_STT = date1 + 1: C2_END = date_END: iDays2 = C2_END - C2_STT + 1: imonths = DateDiff("m", C2_STT, C2_END) + 1
    iDays3 = 0
    
    ReDim CCC1(1 To ExtStaNum, 1 To iDays1)
    For j = 1 To iDays1
      For i = 1 To ExtStaNum
        CCC1(i, j) = -9999
      Next i
    Next j
    ReDim CCC2(1 To ExtStaNum, 1 To imonths)
    For j = 1 To imonths
      For i = 1 To ExtStaNum
        CCC2(i, j) = -9999
      Next i
    Next j
    '判断月值的读取是否跨年度
    If Year(C2_STT) < Year(C2_END) Then
      crossYear = True
    Else
      crossYear = False
    End If
    
    
  ElseIf date1 < date_END And Format(date_STT, "dd") <> "01" And Format(date_END + 1, "dd") <> "01" And date1 + 1 <> date2 Then '情形五：多日+多月+多日
    C1_STT = date_STT: C1_END = date1: iDays1 = C1_END - C1_STT + 1
    C2_STT = date1 + 1: C2_END = date2 - 1: iDays2 = C2_END - C2_STT + 1: imonths = DateDiff("m", C2_STT, C2_END) + 1
    C3_STT = date2: C3_END = date_END: iDays3 = C3_END - C3_STT + 1
  
    ReDim CCC1(1 To ExtStaNum, 1 To iDays1)
    For j = 1 To iDays1
      For i = 1 To ExtStaNum
        CCC1(i, j) = -9999
      Next i
    Next j
    ReDim CCC2(1 To ExtStaNum, 1 To imonths)
    For j = 1 To imonths
      For i = 1 To ExtStaNum
        CCC2(i, j) = -9999
      Next i
    Next j
    ReDim CCC3(1 To ExtStaNum, 1 To iDays3)
    For j = 1 To iDays3
      For i = 1 To ExtStaNum
        CCC3(i, j) = -9999
      Next i
    Next j
  
    '判断月值的读取是否跨年度（2024-1-2添加）
    If Year(C2_STT) < Year(C2_END) Then
      crossYear = True
    Else
      crossYear = False
    End If
  
  
  ElseIf date1 < date_END And Format(date_STT, "dd") <> "01" And Format(date_END + 1, "dd") <> "01" And date1 + 1 = date2 Then '情形六：多日+多日
    C1_STT = date_STT: C1_END = date1: iDays1 = C1_END - C1_STT + 1
    Days2 = 0
    C3_STT = date2: C3_END = date_END: iDays3 = C3_END - C3_STT + 1
  
    ReDim CCC1(1 To ExtStaNum, 1 To iDays1)
    For j = 1 To iDays1
      For i = 1 To ExtStaNum
        CCC1(i, j) = -9999
      Next i
    Next j
    ReDim CCC3(1 To ExtStaNum, 1 To iDays3)
    For j = 1 To iDays3
      For i = 1 To ExtStaNum
        CCC3(i, j) = -9999
      Next i
    Next j
  
  
  End If



  Dim tmp() As Integer '调取QueryMulDayRstat必须传递的参数（本用于各站点的极值出现年份，此处留空即可）2022-5-13
  Dim tmpMonth As Integer '判断当前月份
  Dim tmpDays As Integer '当前月份对应的日数
  If iDays1 <> 0 Then
    temp_STT = Format(C1_STT, "mm-dd"): temp_END = Format(C1_END, "mm-dd")
    Call QueryMulDayRestatExt(SelField_Restat, temp_STT, temp_END, "Norm", False, CCC1(), tmp(), False)
  End If
  If iDays2 <> 0 Then
    temp_STT = Format(C2_STT, "mm"): temp_END = Format(C2_END, "mm")
    If FlagStat = "SUM" Then
      Call QueryMulMonRestatExt(SelField_Restat & "_SUM", temp_STT, temp_END, "Norm", crossYear, CCC2(), tmp(), False)
    End If
    If FlagStat = "AVE" Then
      Call QueryMulMonRestatExt(SelField_Restat, temp_STT, temp_END, "Norm", crossYear, CCC2(), tmp(), False)
    End If
  End If
  If iDays3 <> 0 Then
    temp_STT = Format(C3_STT, "mm-dd"): temp_END = Format(C3_END, "mm-dd")
    Call QueryMulDayRestatExt(SelField_Restat, temp_STT, temp_END, "Norm", False, CCC3(), tmp(), False)
  End If



  '********** 2023-1-30 修改如下 **********
  For i = 1 To ExtStaNum
    sngSum = 0: iNum = 0: sngMin = 999999: sngMax = -9999: idays = 0
    '**** 多日值累加 ****
    For j = 1 To iDays1
      If CCC1(i, j) <> -9999 Then
        sngSum = sngSum + CCC1(i, j)
        idays = idays + 1
      End If
    Next j
    '**** 多月值x日数累加 ****
    For j = 1 To imonths
      tmpMonth = Month(C2_STT) + j - 1
      Select Case tmpMonth
        Case 1, 3, 5, 7, 8, 10, 12
          tmpDays = 31
        Case 2
          tmpDays = 28
        Case 4, 6, 9, 11
          tmpDays = 30
      End Select
        
      If CCC2(i, j) <> -9999 Then
        If FlagStat = "AVE" Then sngSum = sngSum + CCC2(i, j) * tmpDays
        If FlagStat = "SUM" Or FlagStat = "DAYS" Then sngSum = sngSum + CCC2(i, j)
        idays = idays + tmpDays
      End If
    Next j
    '**** 多日值累加 ****
    For j = 1 To iDays3
      If CCC3(i, j) <> -9999 Then
        sngSum = sngSum + CCC3(i, j)
        idays = idays + 1
      End If
    Next j
    '**** 求最后的结果 ****
    If idays <> 0 Then
      If FlagStat = "AVE" Then
        Result(i) = sngSum / idays '求均值
      ElseIf FlagStat = "SUM" Or FlagStat = "DAYS" Then
        Result(i) = sngSum  '求和值
      End If
    End If
      
    '2022-05-13对结果进行四舍五入
    If Result(i) <> -9999 Then
      Result(i) = Format(Result(i), "0.0")
    End If
    
  Next i
  


End Sub




'*******************************************2023-01-30 仅查询日值而后进行再统计*******************************************
Public Sub QueryPeriodNormDataExt2(SelField_Restat$, BMMDD$, EMMDD$, crossYear As Boolean, Result!(), ByVal FlagStat$, Min!, Max!)
  '如果结果将由几部分求平均获取，则：FlagStat由"AVE"替换为"AVE_SUM"
  Dim i%, j%, k%
  
  Dim date_STT As Date, date_END As Date   '开始日期、结束日期
  Dim date1 As Date, date2 As Date   '分割日期1、分割日期2
  Dim C1_STT As Date, C1_END As Date, C2_STT As Date, C2_END As Date, C3_STT As Date, C3_END As Date 'C1、C2、C3起止日期
  Dim temp_STT$, temp_END$ '起止月份或起止年月
  
  Dim iDays1%, iDays2%, iDays3%, idays% '被分割后的多段日期的日数，及总日数
  Dim CCC1!(), CCC2!(), CCC3!() ' 存放所有站点：多日结果、多个整月结果、多日结果
  Dim CDays1%(), CDays2%(), CDays3%() ' 存放有站点条件日数：多日结果、多个整月结果、多日结果
  Dim imonths% '多月的月份数
  
  Dim iNum%, sngSum!, sngMax!, sngMin '求和的个数，和数，最大值、最小值
  
  Dim TableNormDay$ '查询要素所在的数据表
     
  
  If EMMDD <> "02-29" Then
    If crossYear = False Then
      date_STT = CDate("2011-" & BMMDD)
      date_END = CDate("2011-" & EMMDD)
    ElseIf crossYear = True Then
      date_STT = CDate("2010-" & BMMDD)
      date_END = CDate("2011-" & EMMDD)
    End If
    
  '2024-2-27 添加对闰年2-29的判断
  ElseIf EMMDD = "02-29" Then
    If crossYear = False Then
      date_STT = CDate("2012-" & BMMDD)
      date_END = CDate("2012-" & EMMDD)
    ElseIf crossYear = True Then
      date_STT = CDate("2011-" & BMMDD)
      date_END = CDate("2012-" & EMMDD)
    End If
    
  End If
  
  
  
  date1 = CDate(Format(DateAdd("m", 1, date_STT), "yyyy-mm-01")) - 1
  date2 = CDate(Format(date_END, "yyyy-mm-01"))
  
  C1_STT = date_STT: C1_END = date_END: iDays1 = C1_END - C1_STT + 1
  iDays2 = 0
  iDays3 = 0
    
  ReDim CCC1(1 To ExtStaNum, 1 To iDays1)
  For j = 1 To iDays1
    For i = 1 To ExtStaNum
      CCC1(i, j) = -9999
    Next i
  Next j
    

  Dim tmp() As Integer '调取QueryMulDayRstat必须传递的参数（本用于各站点的极值出现年份，此处留空即可）2022-5-13
  Dim tmpMonth As Integer '判断当前月份
  Dim tmpDays As Integer '当前月份对应的日数
  temp_STT = Format(C1_STT, "mm-dd"): temp_END = Format(C1_END, "mm-dd")
  Call QueryMulDayRestatExt(SelField_Restat, temp_STT, temp_END, "Norm", False, CCC1(), tmp(), False)

  '********** 2023-1-30 修改如下 **********
  For i = 1 To ExtStaNum
    sngSum = 0: iNum = 0: sngMin = 999999: sngMax = -9999: idays = 0
      
    '**** 多日值累加 ****
    For j = 1 To iDays1
      If CCC1(i, j) <> -9999 Then
        sngSum = sngSum + CCC1(i, j)
        idays = idays + 1
      End If
    Next j
    '**** 求最后的结果 ****
    If idays <> 0 Then
      If FlagStat = "AVE" Then
        Result(i) = sngSum / idays '求均值
      ElseIf FlagStat = "SUM" Or FlagStat = "DAYS" Then
        Result(i) = sngSum  '求和值
      End If
    End If
      
    '2022-05-13对结果进行四舍五入
    If Result(i) <> -9999 Then
      Result(i) = Format(Result(i), "0.0")
    End If
    
  Next i
  

End Sub






'***********************************************逐日平均态/极端态查询模块(2022-04-02)**********************************************
'读取多日数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维），极端值对应的年份，是否查询极值年份
Public Sub QueryMulDayRestatExt(SelField$, BMMDD$, EMMDD$, DataType$, crossYear As Boolean, Result!(), Result2%(), XtrmYear As Boolean)
  Dim LastSTA$
  Dim BDate As Date, EDate As Date
  Dim i%, j%, k%

  If EMMDD <> "02-29" Then
    If crossYear = False Then
      BDate = Format(("2001-" & BMMDD), "yyyy-mm-dd")
      EDate = Format(("2001-" & EMMDD), "yyyy-mm-dd")
    ElseIf crossYear = True Then
      BDate = Format(("2001-" & BMMDD), "yyyy-mm-dd")
      EDate = Format(("2002-" & EMMDD), "yyyy-mm-dd")
    End If
    
  '2024-2-27 添加对闰年2-29的判断
  ElseIf EMMDD = "02-29" Then
    If crossYear = False Then
      BDate = Format(("2001-" & BMMDD), "yyyy-mm-dd")
      EDate = Format("2001-02-28", "yyyy-mm-dd")
    ElseIf crossYear = True Then
      BDate = Format(("2001-" & BMMDD), "yyyy-mm-dd")
      EDate = Format("2002-02-28", "yyyy-mm-dd")
    End If
       
  End If

  
    TableName = "SURF_CLI_MUL_DAY_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName = TableName & "_CMA"

  
    If crossYear = False Then  '起止日期同年度
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURFExt
      
    ElseIf crossYear = True Then
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-31','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURFExt
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2002-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURFExt
      
    End If
    strSQL = strSQL & " order by sta,ddate"

  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To ExtStaNum
          If ORARst!STA = ExtStaInfo2(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!AAA
            If XtrmYear = True Then Result2(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!iYear
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!AAA
        If XtrmYear = True Then Result2(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!iYear
      End If
  
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub



'***********************************************逐月平均态/极端态查询模块(2022-04-13)**********************************************
'读取多日数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维）
Public Sub QueryMulMonRestatExt(SelField$, BMM$, EMM$, DataType$, crossYear As Boolean, Result!(), Result2%(), XtrmYear As Boolean)
  Dim LastSTA$
  Dim BDate As Date, EDate As Date
  Dim i%, j%, k%
  
  If crossYear = False Then
    BDate = Format(("2001-" & BMM & "-01"), "yyyy-mm-dd")
    EDate = Format(("2001-" & EMM & "-01"), "yyyy-mm-dd")
  ElseIf crossYear = True Then
    BDate = Format(("2001-" & BMM & "-01"), "yyyy-mm-dd")
    EDate = Format(("2002-" & EMM & "-01"), "yyyy-mm-dd")
  End If
  
    TableName = "SURF_CLI_MUL_MON_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName = TableName & "_CMA"
  
  
    If crossYear = False Then  '起止日期同年度
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURFExt
      
    ElseIf crossYear = True Then
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-01','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURFExt
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2002-' || imonth || '-01','yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURFExt
      
    End If
    strSQL = strSQL & " order by sta,ddate"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
   
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To ExtStaNum
          If ORARst!STA = ExtStaInfo2(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("m", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!AAA
            If XtrmYear = True Then Result2(i, DateDiff("m", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!iYear
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("m", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!AAA
        If XtrmYear = True Then Result2(i, DateDiff("m", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!iYear
      End If
    
      ORARst.MoveNext
    Loop

  End If
  
  ORARst.Close

End Sub



'***********************************************逐年平均态/极端态查询模块(2022-04-16)**********************************************
'读取年数据：查询要素，数据类型，数据结果，极端年份结果，是否查极端态数据
Public Sub QueryMulYerRestatExt(SelField$, DataType$, Result!(), Result2%(), XtrmYear As Boolean)
  Dim LastSTA$
  Dim i%, j%, k%
 
    TableName = "SURF_CLI_MUL_YER_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName = TableName & "_CMA"

  
    strSQL = "select stacode STA," & SelField & " AAA"
    If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
    strSQL = strSQL & " from " & TableName
    strSQL = strSQL & " where " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURFExt
    strSQL = strSQL & " order by sta"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
   
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To ExtStaNum
          If ORARst!STA = ExtStaInfo2(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, 1) = ORARst!AAA
            If XtrmYear = True Then Result2(i, 1) = ORARst!iYear
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, 1) = ORARst!AAA
        If XtrmYear = True Then Result2(i, 1) = ORARst!iYear
      End If
    
      ORARst.MoveNext
    Loop

  End If
  
  ORARst.Close

End Sub

