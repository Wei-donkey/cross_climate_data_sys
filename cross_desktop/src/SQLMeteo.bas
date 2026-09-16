Attribute VB_Name = "SQLMeteo"
Option Explicit


'***********************************************逐时查询模块(2024-12-2)**********************************************
'读取连续多时次数据：查询数据表，选择要素，开始时间，结束时间，输出值（二维）
Public Sub QueryMulHorData1(SelField$, BTime$, ETime$, Result!())
  Dim i%, j%, k%
  Dim LastSTA$
  Dim BYear$, EYear$
  
  BYear = Format(BTime, "yyyy"):  EYear = Format(ETime, "yyyy")
  
  If StaType <> "BOTH" And BYear = EYear Then
  
    strSQL = "select stacode STA,ddatetime," & SelField & " AAA from " & StaType & "_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    
    strSQL = strSQL & " order by sta,ddatetime"
  
  ElseIf StaType <> "BOTH" And BYear <> EYear Then
  
    strSQL = "select stacode STA,ddatetime," & SelField & " AAA from " & StaType & "_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddatetime," & SelField & " AAA from " & StaType & "_CLI_MUL_HOR"
    strSQL = strSQL & "_" & EYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    
    strSQL = strSQL & " order by sta,ddatetime"
  
  ElseIf StaType = "BOTH" And BYear = EYear Then
  
    strSQL = "select stacode STA,ddatetime," & SelField & " AAA from SURF_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddatetime," & SelField & " AAA from AWST_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
    
    strSQL = strSQL & " order by sta,ddatetime"
  
  ElseIf StaType = "BOTH" And BYear <> EYear Then
  
    strSQL = "select stacode STA,ddatetime," & SelField & " AAA from SURF_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddatetime," & SelField & " AAA from SURF_CLI_MUL_HOR"
    strSQL = strSQL & "_" & EYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddatetime," & SelField & " AAA from AWST_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddatetime," & SelField & " AAA from AWST_CLI_MUL_HOR"
    strSQL = strSQL & "_" & EYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
    
    strSQL = strSQL & " order by sta,ddatetime"
  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn3, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("h", CDate(BTime), CDate(ORARst!dDatetime)) + 1) = ORARst!AAA
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("h", CDate(BTime), CDate(ORARst!dDatetime)) + 1) = ORARst!AAA
      End If
  
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub






'***********************************************逐时查询模块(2022-06-19)**********************************************
'读取固定时次数据：选择要素，开始时间，结束时间，输出值（二维）
Public Sub QueryMulHorData2(SelField$, BTime$, ETime$, Result!())
  Dim i%, j%, k%
  Dim LastSTA$
  Dim BYear$, EYear$, strHour$
  
  BYear = Format(BTime, "yyyy"):  EYear = Format(ETime, "yyyy"): strHour = Format(BTime, "hh:mm")
  
  If StaType <> "BOTH" And BYear = EYear Then
  
    strSQL = "select stacode STA,ddatetime," & SelField & " AAA from " & StaType & "_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
    strSQL = strSQL & " and to_char(ddatetime,'hh24:mi') = '" & strHour & "'"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    
    strSQL = strSQL & " order by sta,ddatetime"
  
  ElseIf StaType <> "BOTH" And BYear <> EYear Then
  
    strSQL = "select stacode STA,ddatetime," & SelField & " AAA from " & StaType & "_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
    strSQL = strSQL & " and to_char(ddatetime,'hh24:mi') = '" & strHour & "'"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    
    strSQL = strSQL & " UNION "
    
    strSQL = "select stacode STA,ddatetime," & SelField & " AAA from " & StaType & "_CLI_MUL_HOR"
    strSQL = strSQL & "_" & EYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
    strSQL = strSQL & " and to_char(ddatetime,'hh24:mi') = '" & strHour & "'"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    
    strSQL = strSQL & " order by sta,ddatetime"
  
  ElseIf StaType = "BOTH" And BYear = EYear Then
  
    strSQL = "select stacode STA,ddatetime," & SelField & " AAA from SURF_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
    strSQL = strSQL & " and to_char(ddatetime,'hh24:mi') = '" & strHour & "'"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddatetime," & SelField & " AAA from AWST_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
    strSQL = strSQL & " and to_char(ddatetime,'hh24:mi') = '" & strHour & "'"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
    
    strSQL = strSQL & " order by sta,ddatetime"
  
  ElseIf StaType = "BOTH" And BYear <> EYear Then
  
    strSQL = "select stacode STA,ddatetime," & SelField & " AAA from SURF_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
    strSQL = strSQL & " and to_char(ddatetime,'hh24:mi') = '" & strHour & "'"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddatetime," & SelField & " AAA from SURF_CLI_MUL_HOR"
    strSQL = strSQL & "_" & EYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
    strSQL = strSQL & " and to_char(ddatetime,'hh24:mi') = '" & strHour & "'"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddatetime," & SelField & " AAA from AWST_CLI_MUL_HOR"
    strSQL = strSQL & "_" & BYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
    strSQL = strSQL & " and to_char(ddatetime,'hh24:mi') = '" & strHour & "'"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddatetime," & SelField & " AAA from AWST_CLI_MUL_HOR"
    strSQL = strSQL & "_" & EYear
    strSQL = strSQL & " where ddatetime between to_date('" & BTime & "','yyyy-mm-dd hh24:mi') and to_date('" & ETime & "','yyyy-mm-dd hh24:mi')"
    strSQL = strSQL & " and to_char(ddatetime,'hh24:mi') = '" & strHour & "'"
'    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
    
    strSQL = strSQL & " order by sta,ddatetime"
  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn3, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("d", CDate(BTime), CDate(ORARst!dDatetime)) + 1) = ORARst!AAA
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("d", CDate(BTime), CDate(ORARst!dDatetime)) + 1) = ORARst!AAA
      End If
  
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub


'***********************************************逐日查询模块(2022-02-23)**********************************************
'读取多日数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维）
Public Sub QueryMulDayData(SelField$, BDate$, EDate$, Result!())
  Dim i%, j%, k%
  Dim LastSTA$
  Dim BYear$, EYear$
  
'  BYear = Format(BDate, "yyyy"): EYear = Format(EDate, "yyyy")
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    strSQL = "select stacode STA,ddate," & SelField & " AAA from " & StaType & "_CLI_MUL_DAY"
    strSQL = strSQL & " where ddate between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    strSQL = strSQL & " order by sta,ddate"
  
  ElseIf StaType = "BOTH" Then
    
    strSQL = "select stacode STA,ddate," & SelField & " AAA from SURF_CLI_MUL_DAY"
    strSQL = strSQL & " where ddate between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ddate," & SelField & " AAA from AWST_CLI_MUL_DAY"
    strSQL = strSQL & " where ddate between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
        
    strSQL = strSQL & " order by sta,ddate"
  
  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!AAA
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1) = ORARst!AAA
      End If
  
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub




'***********************************************逐日平均态/极端态查询模块(2022-04-02)**********************************************
'读取多日数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维），极端值对应的年份，是否查询极值年份
Public Sub QueryMulDayRestat(SelField$, BMMDD$, EMMDD$, DataType$, crossYear As Boolean, Result!(), Result2%(), XtrmYear As Boolean)
  Dim LastSTA$
  Dim BDate As Date, EDate As Date
  Dim i%, j%, k%
  
'  If CrossYear = False Then
'    BDate = Format(("2001-" & BMMDD), "yyyy-mm-dd")
'    EDate = Format(("2001-" & EMMDD), "yyyy-mm-dd")
'  ElseIf CrossYear = True Then
'    BDate = Format(("2001-" & BMMDD), "yyyy-mm-dd")
'    EDate = Format(("2002-" & EMMDD), "yyyy-mm-dd")
'  End If
'
  
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
  
  
  
  
  If StaType = "SURF" Or StaType = "AWST" Then
    TableName = StaType & "_CLI_MUL_DAY_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName = TableName & "_CMA"
  ElseIf StaType = "BOTH" Then
    TableName1 = "SURF_CLI_MUL_DAY_" & DataType
    TableName2 = "AWST_CLI_MUL_DAY_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName1 = TableName1 & "_CMA"
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName2 = TableName2 & "_CMA"
  End If
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    If crossYear = False Then  '起止日期同年度
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSta
      
    ElseIf crossYear = True Then
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-31','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSta
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2002-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSta
      
    End If
    strSQL = strSQL & " order by sta,ddate"
  
  
  ElseIf StaType = "BOTH" Then
    
    If crossYear = False Then  '起止日期同年度
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName1
      strSQL = strSQL & " where to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURF
      
    ElseIf crossYear = True Then
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName1
      strSQL = strSQL & " where to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-31','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURF
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName1
      strSQL = strSQL & " where to_date('2002-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURF
    End If

    strSQL = strSQL & " UNION "
    
    If crossYear = False Then  '起止日期同年度
      strSQL = strSQL & "select stacode STA,to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName2
      strSQL = strSQL & " where to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strAWST
      
    ElseIf crossYear = True Then
      strSQL = strSQL & "select stacode STA,to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName2
      strSQL = strSQL & " where to_date('2001-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-31','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strAWST
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-' || iday,'yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName2
      strSQL = strSQL & " where to_date('2002-' || imonth || '-' || iday,'yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strAWST
      
    End If
    strSQL = strSQL & " order by sta,ddate"

  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
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




'***********************************************逐月查询模块(2022-02-23)**********************************************
'读取多日数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维）
Public Sub QueryMulTenData(SelField$, BYYMM$, EYYMM$, Result!())
  Dim i%, j%, k%
  Dim LastSTA$
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    strSQL = "select stacode STA,to_date(iyear || '-' || imonth,'yyyy-mm') YYMM,iten," & SelField & " AAA from " & StaType & "_CLI_MUL_TEN"
    strSQL = strSQL & " where to_date(iyear || '-' || imonth,'yyyy-mm') between to_date('" & BYYMM & "','yyyy-mm') and to_date('" & EYYMM & "','yyyy-mm')"
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    strSQL = strSQL & " order by sta,YYMM,iten"
  
  ElseIf StaType = "BOTH" Then
    
    strSQL = "select stacode STA,to_date(iyear || '-' || imonth,'yyyy-mm') YYMM,iten," & SelField & " AAA from SURF_CLI_MUL_TEN"
    strSQL = strSQL & " where to_date(iyear || '-' || imonth,'yyyy-mm') between to_date('" & BYYMM & "','yyyy-mm') and to_date('" & EYYMM & "','yyyy-mm')"
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,to_date(iyear || '-' || imonth,'yyyy-mm') YYMM,iten," & SelField & " AAA from AWST_CLI_MUL_TEN"
    strSQL = strSQL & " where to_date(iyear || '-' || imonth,'yyyy-mm') between to_date('" & BYYMM & "','yyyy-mm') and to_date('" & EYYMM & "','yyyy-mm')"
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
        
    strSQL = strSQL & " order by sta,YYMM,iTen"
  
  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, 3 * DateDiff("m", CDate(BYYMM), CDate(ORARst!YYMM)) + 1 + ORARst!iten - 1) = ORARst!AAA
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, 3 * DateDiff("m", CDate(BYYMM), CDate(ORARst!YYMM)) + 1 + ORARst!iten - 1) = ORARst!AAA
      End If
  
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub




'***********************************************逐旬平均态/极端态查询模块(2023-06-06)**********************************************
'读取多旬数据：查询数据表，选择要素，开始月份，结束月份，输出值（二维）
Public Sub QueryMulTenRestat(SelField$, BMM$, EMM$, DataType$, crossYear As Boolean, Result!(), Result2%(), XtrmYear As Boolean)
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
  
  If StaType = "SURF" Or StaType = "AWST" Then
    TableName = StaType & "_CLI_MUL_TEN_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName = TableName & "_CMA"
  ElseIf StaType = "BOTH" Then
    TableName1 = "SURF_CLI_MUL_TEN_" & DataType
    TableName2 = "AWST_CLI_MUL_TEN_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName1 = TableName1 & "_CMA"
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName2 = TableName2 & "_CMA"
  End If
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    If crossYear = False Then  '起止日期同年度
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate,iten," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSta
      
    ElseIf crossYear = True Then
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate,iten," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-01','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSta
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-01','yyyy-mm-dd') ddate,iten," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2002-' || imonth || '-01','yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSta
      
    End If
    strSQL = strSQL & " order by sta,ddate,iten"
  
  
  ElseIf StaType = "BOTH" Then
    
    If crossYear = False Then  '起止日期同年度
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate,iten," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName1
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURF
      
    ElseIf crossYear = True Then
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate,iten," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName1
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-01','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURF
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-01','yyyy-mm-dd') ddate,iten," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName1
      strSQL = strSQL & " where to_date('2002-' || imonth || '-01','yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURF
    End If

    strSQL = strSQL & " UNION "
    
    If crossYear = False Then  '起止日期同年度
      strSQL = strSQL & "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate,iten," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName2
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strAWST
      
    ElseIf crossYear = True Then
      strSQL = strSQL & "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate,iten," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName2
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-01','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strAWST
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-01','yyyy-mm-dd') ddate,iten," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName2
      strSQL = strSQL & " where to_date('2002-' || imonth || '-01','yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strAWST
      
    End If
    strSQL = strSQL & " order by sta,ddate"

  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
   
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
          
            If Not IsNull(ORARst!AAA) Then Result(i, 3 * DateDiff("m", CDate(BDate), CDate(ORARst!ddate)) + 1 + ORARst!iten - 1) = ORARst!AAA
            If XtrmYear = True Then Result2(i, 3 * DateDiff("m", CDate(BDate), CDate(ORARst!ddate)) + 1 + ORARst!iten - 1) = ORARst!iYear
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, 3 * DateDiff("m", CDate(BDate), CDate(ORARst!ddate)) + 1 + ORARst!iten - 1) = ORARst!AAA
        If XtrmYear = True Then Result2(i, 3 * DateDiff("m", CDate(BDate), CDate(ORARst!ddate)) + 1 + ORARst!iten - 1) = ORARst!iYear
      End If
    
      ORARst.MoveNext
    Loop

  End If
  
  ORARst.Close

End Sub



'***********************************************逐月查询模块(2022-02-23)**********************************************
'读取多日数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维）
Public Sub QueryMulMonData(SelField$, BYYMM$, EYYMM$, Result!())
  Dim i%, j%, k%
  Dim LastSTA$
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    strSQL = "select stacode STA,to_date(iyear || '-' || imonth,'yyyy-mm') YYMM," & SelField & " AAA from " & StaType & "_CLI_MUL_MON"
    strSQL = strSQL & " where to_date(iyear || '-' || imonth,'yyyy-mm') between to_date('" & BYYMM & "','yyyy-mm') and to_date('" & EYYMM & "','yyyy-mm')"
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    strSQL = strSQL & " order by sta,YYMM"
  
  ElseIf StaType = "BOTH" Then
    
    strSQL = "select stacode STA,to_date(iyear || '-' || imonth,'yyyy-mm') YYMM," & SelField & " AAA from SURF_CLI_MUL_MON"
    strSQL = strSQL & " where to_date(iyear || '-' || imonth,'yyyy-mm') between to_date('" & BYYMM & "','yyyy-mm') and to_date('" & EYYMM & "','yyyy-mm')"
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,to_date(iyear || '-' || imonth,'yyyy-mm') YYMM," & SelField & " AAA from AWST_CLI_MUL_MON"
    strSQL = strSQL & " where to_date(iyear || '-' || imonth,'yyyy-mm') between to_date('" & BYYMM & "','yyyy-mm') and to_date('" & EYYMM & "','yyyy-mm')"
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
        
    strSQL = strSQL & " order by sta,YYMM"
  
  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("m", CDate(BYYMM), CDate(ORARst!YYMM)) + 1) = ORARst!AAA
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, DateDiff("m", CDate(BYYMM), CDate(ORARst!YYMM)) + 1) = ORARst!AAA
      End If
  
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub


'***********************************************逐月平均态/极端态查询模块(2022-04-13)**********************************************
'读取多日数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维）
Public Sub QueryMulMonRestat(SelField$, BMM$, EMM$, DataType$, crossYear As Boolean, Result!(), Result2%(), XtrmYear As Boolean)
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
  
  If StaType = "SURF" Or StaType = "AWST" Then
    TableName = StaType & "_CLI_MUL_MON_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName = TableName & "_CMA"
  ElseIf StaType = "BOTH" Then
    TableName1 = "SURF_CLI_MUL_MON_" & DataType
    TableName2 = "AWST_CLI_MUL_MON_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName1 = TableName1 & "_CMA"
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName2 = TableName2 & "_CMA"
  End If
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    If crossYear = False Then  '起止日期同年度
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSta
      
    ElseIf crossYear = True Then
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-01','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSta
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName
      strSQL = strSQL & " where to_date('2002-' || imonth || '-01','yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSta
      
    End If
    strSQL = strSQL & " order by sta,ddate"
  
  
  ElseIf StaType = "BOTH" Then
    
    If crossYear = False Then  '起止日期同年度
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName1
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURF
      
    ElseIf crossYear = True Then
      strSQL = "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName1
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-01','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURF
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName1
      strSQL = strSQL & " where to_date('2002-' || imonth || '-01','yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strSURF
    End If

    strSQL = strSQL & " UNION "
    
    If crossYear = False Then  '起止日期同年度
      strSQL = strSQL & "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName2
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strAWST
      
    ElseIf crossYear = True Then
      strSQL = strSQL & "select stacode STA,to_date('2001-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName2
      strSQL = strSQL & " where to_date('2001-' || imonth || '-01','yyyy-mm-dd') between to_date('" & BDate & "','yyyy-mm-dd') and to_date('2001-12-01','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strAWST
      
      strSQL = strSQL & " UNION "
      
      strSQL = strSQL & "select stacode STA,to_date('2002-' || imonth || '-01','yyyy-mm-dd') ddate," & SelField & " AAA"
      If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
      strSQL = strSQL & " from " & TableName2
      strSQL = strSQL & " where to_date('2002-' || imonth || '-01','yyyy-mm-dd') between to_date('2002-01-01','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
      strSQL = strSQL & " and " & SelField & " is not Null"
      strSQL = strSQL & " and " & strAWST
      
    End If
    strSQL = strSQL & " order by sta,ddate"

  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
   
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
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







'***********************************************逐季查询模块(2022-04-16)**********************************************
'读取多季数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维）
Public Sub QueryMulQtrData(SelField$, BYYYY$, EYYYY$, strType$, Result!())
  Dim i%, j%, k%
  Dim LastSTA$
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    strSQL = "select stacode STA,iyear YYYY,cquarter," & SelField & " AAA from " & StaType & "_CLI_MUL_QTR"
    strSQL = strSQL & " where iyear between " & BYYYY & " and " & EYYYY
    strSQL = strSQL & " and " & SelField & " is not Null"
    If strType = "Qtr" Then strSQL = strSQL & " and cquarter in ('1st','2nd','3rd','4th')"
    If strType = "Ssn" Then strSQL = strSQL & " and cquarter in ('aut','smr','spr','win')"
    strSQL = strSQL & " and " & strSta
    strSQL = strSQL & " order by sta,YYYY,cquarter"
  
  ElseIf StaType = "BOTH" Then
    
    strSQL = "select stacode STA,iyear YYYY,cquarter," & SelField & " AAA from SURF_CLI_MUL_QTR"
    strSQL = strSQL & " where iyear between " & BYYYY & " and " & EYYYY
    strSQL = strSQL & " and " & SelField & " is not Null"
    If strType = "Qtr" Then strSQL = strSQL & " and cquarter in ('1st','2nd','3rd','4th')"
    If strType = "Ssn" Then strSQL = strSQL & " and cquarter in ('aut','smr','spr','win')"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,iyear YYYY,cquarter," & SelField & " AAA from AWST_CLI_MUL_QTR"
    strSQL = strSQL & " where iyear between " & BYYYY & " and " & EYYYY
    strSQL = strSQL & " and " & SelField & " is not Null"
    If strType = "Qtr" Then strSQL = strSQL & " and cquarter in ('1st','2nd','3rd','4th')"
    If strType = "Ssn" Then strSQL = strSQL & " and cquarter in ('aut','smr','spr','win')"
    strSQL = strSQL & " and " & strAWST
        
    strSQL = strSQL & " order by sta,YYYY,cquarter"
  
  End If
  
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If ORARst!cquarter = "1st" Or ORARst!cquarter = "spr" Then
              If Not IsNull(ORARst!AAA) Then Result(i, 4 * (CInt(ORARst!YYYY) - CInt(BYYYY)) + 1) = ORARst!AAA
            ElseIf ORARst!cquarter = "2nd" Or ORARst!cquarter = "smr" Then
              If Not IsNull(ORARst!AAA) Then Result(i, 4 * (CInt(ORARst!YYYY) - CInt(BYYYY)) + 2) = ORARst!AAA
            ElseIf ORARst!cquarter = "3rd" Or ORARst!cquarter = "aut" Then
              If Not IsNull(ORARst!AAA) Then Result(i, 4 * (CInt(ORARst!YYYY) - CInt(BYYYY)) + 3) = ORARst!AAA
            ElseIf ORARst!cquarter = "4th" Or ORARst!cquarter = "win" Then
              If Not IsNull(ORARst!AAA) Then Result(i, 4 * (CInt(ORARst!YYYY) - CInt(BYYYY)) + 4) = ORARst!AAA
            End If
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If ORARst!cquarter = "1st" Or ORARst!cquarter = "spr" Then
          If Not IsNull(ORARst!AAA) Then Result(i, 4 * (CInt(ORARst!YYYY) - CInt(BYYYY)) + 1) = ORARst!AAA
        ElseIf ORARst!cquarter = "2nd" Or ORARst!cquarter = "smr" Then
          If Not IsNull(ORARst!AAA) Then Result(i, 4 * (CInt(ORARst!YYYY) - CInt(BYYYY)) + 2) = ORARst!AAA
        ElseIf ORARst!cquarter = "3rd" Or ORARst!cquarter = "aut" Then
          If Not IsNull(ORARst!AAA) Then Result(i, 4 * (CInt(ORARst!YYYY) - CInt(BYYYY)) + 3) = ORARst!AAA
        ElseIf ORARst!cquarter = "4th" Or ORARst!cquarter = "win" Then
          If Not IsNull(ORARst!AAA) Then Result(i, 4 * (CInt(ORARst!YYYY) - CInt(BYYYY)) + 4) = ORARst!AAA
        End If
      End If
  
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close


End Sub






'***********************************************逐季平均态/极端态查询模块(2023-06-07)**********************************************
'读取年数据：查询要素，数据类型，数据结果，极端年份结果，是否查极端态数据
Public Sub QueryMulQtrRestat(SelField$, DataType$, strType$, Result!(), Result2%(), XtrmYear As Boolean)
  Dim LastSTA$
  Dim i%, j%, k%
 
  If StaType = "SURF" Or StaType = "AWST" Then
    TableName = StaType & "_CLI_MUL_QTR_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName = TableName & "_CMA"
  ElseIf StaType = "BOTH" Then
    TableName1 = "SURF_CLI_MUL_QTR_" & DataType
    TableName2 = "AWST_CLI_MUL_QTR_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName1 = TableName1 & "_CMA"
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName2 = TableName2 & "_CMA"
  End If
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    strSQL = "select stacode STA,cquarter," & SelField & " AAA"
    If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
    strSQL = strSQL & " from " & TableName
    strSQL = strSQL & " where " & SelField & " is not Null"
    If strType = "Qtr" Then strSQL = strSQL & " and cquarter in ('1st','2nd','3rd','4th')"
    If strType = "Ssn" Then strSQL = strSQL & " and cquarter in ('aut','smr','spr','win')"
    strSQL = strSQL & " and " & strSta
    strSQL = strSQL & " order by sta,cquarter"
  
  ElseIf StaType = "BOTH" Then
    
    strSQL = "select stacode STA,cquarter," & SelField & " AAA"
    If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
    strSQL = strSQL & " from " & TableName1
    strSQL = strSQL & " where " & SelField & " is not Null"
    If strType = "Qtr" Then strSQL = strSQL & " and cquarter in ('1st','2nd','3rd','4th')"
    If strType = "Ssn" Then strSQL = strSQL & " and cquarter in ('aut','smr','spr','win')"
    strSQL = strSQL & " and " & strSURF

    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,cquarter," & SelField & " AAA"
    If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
    strSQL = strSQL & " from " & TableName2
    strSQL = strSQL & " where " & SelField & " is not Null"
    If strType = "Qtr" Then strSQL = strSQL & " and cquarter in ('1st','2nd','3rd','4th')"
    If strType = "Ssn" Then strSQL = strSQL & " and cquarter in ('aut','smr','spr','win')"
    strSQL = strSQL & " and " & strAWST
    
    strSQL = strSQL & " order by sta,cquarter"

  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
   
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If ORARst!cquarter = "1st" Or ORARst!cquarter = "spr" Then
              If Not IsNull(ORARst!AAA) Then Result(i, 1) = ORARst!AAA
            ElseIf ORARst!cquarter = "2nd" Or ORARst!cquarter = "smr" Then
              If Not IsNull(ORARst!AAA) Then Result(i, 2) = ORARst!AAA
            ElseIf ORARst!cquarter = "3rd" Or ORARst!cquarter = "aut" Then
              If Not IsNull(ORARst!AAA) Then Result(i, 3) = ORARst!AAA
            ElseIf ORARst!cquarter = "4th" Or ORARst!cquarter = "win" Then
              If Not IsNull(ORARst!AAA) Then Result(i, 4) = ORARst!AAA
            End If
            
            If XtrmYear = True Then
              If ORARst!cquarter = "1st" Or ORARst!cquarter = "spr" Then
                Result2(i, 1) = ORARst!iYear
              ElseIf ORARst!cquarter = "2nd" Or ORARst!cquarter = "smr" Then
                Result2(i, 2) = ORARst!iYear
              ElseIf ORARst!cquarter = "3rd" Or ORARst!cquarter = "aut" Then
                Result2(i, 3) = ORARst!iYear
              ElseIf ORARst!cquarter = "4th" Or ORARst!cquarter = "win" Then
                Result2(i, 4) = ORARst!iYear
              End If
            End If
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If ORARst!cquarter = "1st" Or ORARst!cquarter = "spr" Then
          If Not IsNull(ORARst!AAA) Then Result(i, 1) = ORARst!AAA
        ElseIf ORARst!cquarter = "2nd" Or ORARst!cquarter = "smr" Then
          If Not IsNull(ORARst!AAA) Then Result(i, 2) = ORARst!AAA
        ElseIf ORARst!cquarter = "3rd" Or ORARst!cquarter = "aut" Then
          If Not IsNull(ORARst!AAA) Then Result(i, 3) = ORARst!AAA
        ElseIf ORARst!cquarter = "4th" Or ORARst!cquarter = "win" Then
          If Not IsNull(ORARst!AAA) Then Result(i, 4) = ORARst!AAA
        End If
        
        If XtrmYear = True Then
          If ORARst!cquarter = "1st" Or ORARst!cquarter = "spr" Then
            Result2(i, 1) = ORARst!iYear
          ElseIf ORARst!cquarter = "2nd" Or ORARst!cquarter = "smr" Then
            Result2(i, 2) = ORARst!iYear
          ElseIf ORARst!cquarter = "3rd" Or ORARst!cquarter = "aut" Then
            Result2(i, 3) = ORARst!iYear
          ElseIf ORARst!cquarter = "4th" Or ORARst!cquarter = "win" Then
            Result2(i, 4) = ORARst!iYear
          End If
        End If
      End If
    
      ORARst.MoveNext
    Loop



  End If
  
  ORARst.Close

End Sub


'***********************************************逐年查询模块(2022-04-16)**********************************************
'读取多年数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维）
Public Sub QueryMulYerData(SelField$, BYYYY%, EYYYY%, Result!())
  Dim i%, j%, k%
  Dim LastSTA$
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    strSQL = "select stacode STA,iyear YYYY," & SelField & " AAA from " & StaType & "_CLI_MUL_YER"
    strSQL = strSQL & " where iyear between " & Str(BYYYY) & " and " & Str(EYYYY)
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    strSQL = strSQL & " order by sta,YYYY"
  
  ElseIf StaType = "BOTH" Then
    
    strSQL = "select stacode STA,iyear YYYY," & SelField & " AAA from SURF_CLI_MUL_YER"
    strSQL = strSQL & " where iyear between " & Str(BYYYY) & " and " & Str(EYYYY)
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,iyear YYYY," & SelField & " AAA from AWST_CLI_MUL_YER"
    strSQL = strSQL & " where iyear between " & Str(BYYYY) & " and " & Str(EYYYY)
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
        
    strSQL = strSQL & " order by sta,YYYY"
  
  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, CInt(ORARst!YYYY) - BYYYY + 1) = ORARst!AAA
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, CInt(ORARst!YYYY) - BYYYY + 1) = ORARst!AAA
      End If
  
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub


'***********************************************逐年平均态/极端态查询模块(2022-04-16)**********************************************
'读取年数据：查询要素，数据类型，数据结果，极端年份结果，是否查极端态数据
Public Sub QueryMulYerRestat(SelField$, DataType$, Result!(), Result2%(), XtrmYear As Boolean)
  Dim LastSTA$
  Dim i%, j%, k%
 
  If StaType = "SURF" Or StaType = "AWST" Then
    TableName = StaType & "_CLI_MUL_YER_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName = TableName & "_CMA"
  ElseIf StaType = "BOTH" Then
    TableName1 = "SURF_CLI_MUL_YER_" & DataType
    TableName2 = "AWST_CLI_MUL_YER_" & DataType
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName1 = TableName1 & "_CMA"
    If Norm_Src = "CMA" And DataType = "Norm" Then TableName2 = TableName2 & "_CMA"
  End If
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    strSQL = "select stacode STA," & SelField & " AAA"
    If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
    strSQL = strSQL & " from " & TableName
    strSQL = strSQL & " where " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSta
    strSQL = strSQL & " order by sta"
  
  ElseIf StaType = "BOTH" Then
    
    strSQL = "select stacode STA," & SelField & " AAA"
    If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
    strSQL = strSQL & " from " & TableName1
    strSQL = strSQL & " where " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURF

    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA," & SelField & " AAA"
    If XtrmYear = True Then strSQL = strSQL & "," & SelField & "YEAR IYEAR"
    strSQL = strSQL & " from " & TableName2
    strSQL = strSQL & " where " & SelField & " is not Null"
    strSQL = strSQL & " and " & strAWST
    
    strSQL = strSQL & " order by sta"

  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
   
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
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




'***********************************************逐日查询模块(2022-02-23)**********************************************
'读取多日数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维），等级-(1,2,3,4)
Public Sub QueryCAData(BDate$, EDate$, Result() As ColdAir_Result, strLvl$)
  Dim i%, j%, k%
  Dim LastSTA$
  Dim BYear$, EYear$
  
  If StaType = "SURF" Or StaType = "AWST" Then
  
    strSQL = "select stacode STA,ilevel,date_stt,date_end,idays,tv_24h,tv_48h,tv_acc,t,t_min from SURF_AWST_CLI_CA_PROCESS"
    strSQL = strSQL & " where date_stt >=to_date('" & BDate & "','yyyy-mm-dd') and date_stt<=to_date('" & EDate & "','yyyy-mm-dd')"
    strSQL = strSQL & " and " & strSta
    strSQL = strSQL & " and ilevel in " & strLvl
    strSQL = strSQL & " order by sta,date_stt"
  
  ElseIf StaType = "BOTH" Then
    
    strSQL = "select stacode STA,ilevel,date_stt,date_end,idays,tv_24h,tv_48h,tv_acc,t,t_min from SURF_AWST_CLI_CA_PROCESS"
    strSQL = strSQL & " where date_stt >=to_date('" & BDate & "','yyyy-mm-dd') and date_stt<=to_date('" & EDate & "','yyyy-mm-dd')"
    strSQL = strSQL & " and " & strSURF
    strSQL = strSQL & " and ilevel in " & strLvl
    
    strSQL = strSQL & " UNION "
    
    strSQL = strSQL & "select stacode STA,ilevel,date_stt,date_end,idays,tv_24h,tv_48h,tv_acc,t,t_min from SURF_AWST_CLI_CA_PROCESS"
    strSQL = strSQL & " where date_stt >=to_date('" & BDate & "','yyyy-mm-dd') and date_stt<=to_date('" & EDate & "','yyyy-mm-dd')"
    strSQL = strSQL & " and " & strAWST
    strSQL = strSQL & " and ilevel in " & strLvl
        
    strSQL = strSQL & " order by sta,date_stt"
  
  End If
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    ReDim Result(1 To ORARst.RecordCount)
    ORARst.MoveFirst: i = 1
    k = 1: LastSTA = ""
    Do Until ORARst.EOF
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For j = k To StaNum
          If ORARst!STA = StaInfo(j).stacode Then
            Result(i).ilevel = -9999 '冷空气等级：0-无，1-弱，2-中，3-强，4-寒潮
            Result(i).BDate = CDate("1899-09-09") '降温第一日
            Result(i).EDate = CDate("1899-09-09") '回温前一日
            Result(i).idays = -9999 '降温日数
            Result(i).tv_24hmax = -9999 '最大24小时降温
            Result(i).tv_48hmax = -9999 '最大48小时降温
            Result(i).tv_acc = -9999 '累计降温
            Result(i).T = -9999  '降温过程的平均气温
            Result(i).t_min = -9999 '降温过程的最低气温
            
            Result(i).ilevel = ORARst!ilevel
            Result(i).BDate = ORARst!date_STT
            Result(i).EDate = ORARst!date_END
            Result(i).idays = ORARst!idays
            Result(i).tv_24hmax = ORARst!tv_24h
            Result(i).tv_48hmax = ORARst!tv_48h
            Result(i).tv_acc = ORARst!tv_acc
            Result(i).T = ORARst!T
            Result(i).t_min = ORARst!t_min
            
            Result(i).stacode = StaInfo(j).stacode
            Result(i).staname = StaInfo(j).staname
            Result(i).Town = StaInfo(j).Town
            Result(i).County = StaInfo(j).County
            Result(i).City = StaInfo(j).City
            
            i = i + 1
            k = j + 1: Exit For
          End If
        Next j
      ElseIf ORARst!STA = LastSTA Then
      
        Result(i).ilevel = -9999 '冷空气等级：0-无，1-弱，2-中，3-强，4-寒潮
        Result(i).BDate = CDate("1899-09-09") '降温第一日
        Result(i).EDate = CDate("1899-09-09") '回温前一日
        Result(i).idays = -9999 '降温日数
        Result(i).tv_24hmax = -9999 '最大24小时降温
        Result(i).tv_48hmax = -9999 '最大48小时降温
        Result(i).tv_acc = -9999 '累计降温
        Result(i).T = -9999  '降温过程的平均气温
        Result(i).t_min = -9999 '降温过程的最低气温
        
        Result(i).ilevel = ORARst!ilevel
        Result(i).BDate = ORARst!date_STT
        Result(i).EDate = ORARst!date_END
        Result(i).idays = ORARst!idays
        Result(i).tv_24hmax = ORARst!tv_24h
        Result(i).tv_48hmax = ORARst!tv_48h
        Result(i).tv_acc = ORARst!tv_acc
        Result(i).T = ORARst!T
        Result(i).t_min = ORARst!t_min
        
        Result(i).stacode = StaInfo(j).stacode
        Result(i).staname = StaInfo(j).staname
        Result(i).Town = StaInfo(j).Town
        Result(i).County = StaInfo(j).County
        Result(i).City = StaInfo(j).City
          
        i = i + 1
      End If
      
      ORARst.MoveNext
    Loop

  End If
  
  ORARst.Close

End Sub



'***********************************************逐年查询模块(2026-07-16)**********************************************
'读取多年数据：查询数据表，选择要素，开始日期，结束日期，输出值（二维）
Public Sub QueryMulYerDataExt(SelField$, BYYYY%, EYYYY%, Result!())
  Dim i%, j%, k%
  Dim LastSTA$
  
    strSQL = "select stacode STA,iyear YYYY," & SelField & " AAA from SURF_CLI_MUL_YER"
    strSQL = strSQL & " where iyear between " & Str(BYYYY) & " and " & Str(EYYYY)
    strSQL = strSQL & " and " & SelField & " is not Null"
    strSQL = strSQL & " and " & strSURFExt
    strSQL = strSQL & " order by sta,YYYY"
  
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
          If ORARst!STA = ExtStaInfo(i).stacode Then
            If Not IsNull(ORARst!AAA) Then Result(i, CInt(ORARst!YYYY) - BYYYY + 1) = ORARst!AAA
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!AAA) Then Result(i, CInt(ORARst!YYYY) - BYYYY + 1) = ORARst!AAA
      End If
  
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub

