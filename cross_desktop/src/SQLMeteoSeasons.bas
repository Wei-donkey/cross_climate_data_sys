Attribute VB_Name = "SQLMeteoSeasons"
Option Explicit


'************ 当年/某年：查 SURF_CLI_MUL_DAY，sql 输出 T5D ************
Public Sub QueryT5dFromDay(Flag As String, BDate As Date, EDate As Date)
  Dim i%, k%
  Dim LastSTA$

  strSQL = "SELECT stacode STA, ddate,"
  strSQL = strSQL & " AVG(t) OVER (PARTITION BY stacode ORDER BY ddate ROWS BETWEEN 4 PRECEDING AND CURRENT ROW) t5d"
  strSQL = strSQL & " FROM SURF_CLI_MUL_DAY"
  strSQL = strSQL & " WHERE ddate BETWEEN TO_DATE('" & Format(BDate, "yyyy-mm-dd") & "','yyyy-mm-dd')"
  strSQL = strSQL & " AND TO_DATE('" & Format(EDate, "yyyy-mm-dd") & "','yyyy-mm-dd')"
  strSQL = strSQL & " and " & strSta
  strSQL = strSQL & " ORDER BY stacode, ddate"

  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
  
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!ddate < BDate + 4 Then GoTo Nextloop
      
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If Not IsNull(ORARst!T5d) Then
              If Flag = "Curr" Then
                T5d_Curr(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T5d, 1) 'BDate + 4 日才开始有5日滑动平均值
                T5d_Curr(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
              ElseIf Flag = "Comp" Then
                T5d_Comp(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T5d, 1)
                T5d_Comp(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
              End If
            End If
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!T5d) Then
          If Flag = "Curr" Then
            T5d_Curr(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T5d, 1)
            T5d_Curr(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
          ElseIf Flag = "Comp" Then
            T5d_Comp(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T5d, 1)
            T5d_Comp(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
          End If
        End If
      End If
      
Nextloop:
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub


'************ 常年：查 SURF_CLI_T_DAY_NORM（无年份，仅IMONTH/IDAY），SQL端按月日范围筛选（含跨年） ************
Public Sub QueryT5dFromNorm(BDate As Date, EDate As Date)
  
  Dim i%, k%
  Dim LastSTA$

  strSQL = "SELECT stacode STA, ddate,"
  strSQL = strSQL & " AVG(t) OVER (PARTITION BY stacode ORDER BY ddate ROWS BETWEEN 4 PRECEDING AND CURRENT ROW) t5d"
  strSQL = strSQL & " FROM SURF_CLI_T_DAY_NORM"
  strSQL = strSQL & " WHERE ddate BETWEEN TO_DATE('" & Format(BDate, "yyyy-mm-dd") & "','yyyy-mm-dd')"
  strSQL = strSQL & " AND TO_DATE('" & Format(EDate, "yyyy-mm-dd") & "','yyyy-mm-dd')"
  strSQL = strSQL & " and " & strSta
  strSQL = strSQL & " ORDER BY stacode, ddate"

  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
  
    '将该年度所有站点的统计值存入变量数组
    ORARst.MoveFirst
    k = 1: LastSTA = ""
  
    Do Until ORARst.EOF
      If ORARst!ddate < BDate + 4 Then GoTo Nextloop
      
      If ORARst!STA <> LastSTA Then
        LastSTA = ORARst!STA
        For i = k To StaNum
          If ORARst!STA = StaInfo(i).stacode Then
            If Not IsNull(ORARst!T5d) Then
              T5d_Norm(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T5d, 1)
              T5d_Norm(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
            End If
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!T5d) Then

            T5d_Norm(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T5d, 1)
            T5d_Norm(i, DateDiff("d", CDate(BDate + 4), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate

        End If
      End If
      
Nextloop:
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close

End Sub

'************ 当年/某年：查 SURF_CLI_MUL_DAY，sql 输出逐日 T ************
Public Sub QueryTFromDay(Flag As String, BDate As Date, EDate As Date)
  Dim i%, k%
  Dim LastSTA$

  strSQL = "SELECT stacode STA, ddate,T"
  strSQL = strSQL & " FROM SURF_CLI_MUL_DAY"
  strSQL = strSQL & " WHERE ddate BETWEEN TO_DATE('" & Format(BDate, "yyyy-mm-dd") & "','yyyy-mm-dd')"
  strSQL = strSQL & " AND TO_DATE('" & Format(EDate, "yyyy-mm-dd") & "','yyyy-mm-dd')"
  strSQL = strSQL & " and " & strSta
  strSQL = strSQL & " ORDER BY stacode, ddate"

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
            If Not IsNull(ORARst!T) Then
              If Flag = "Curr" Then
                T_Curr(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T, 1)
                T_Curr(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
              ElseIf Flag = "Comp" Then
                T_Comp(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T, 1)
                T_Comp(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
              End If
            End If
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!T) Then
          If Flag = "Curr" Then
            T_Curr(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T, 1)
            T_Curr(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
          ElseIf Flag = "Comp" Then
            T_Comp(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T, 1)
            T_Comp(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
          End If
        End If
      End If
      
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close
  

End Sub


'************ 当年/某年：查 SURF_CLI_T_DAY_NORM，sql 输出逐日 T ************
Public Sub QueryTFromNorm(BDate As Date, EDate As Date)
  Dim i%, k%
  Dim LastSTA$

  '前置4天，确保窗口第1天起T5D就可算
  strSQL = "SELECT stacode STA, ddate,T"
  strSQL = strSQL & " FROM SURF_CLI_T_DAY_NORM"
  strSQL = strSQL & " WHERE ddate BETWEEN TO_DATE('" & Format(BDate, "yyyy-mm-dd") & "','yyyy-mm-dd')"
  strSQL = strSQL & " AND TO_DATE('" & Format(EDate, "yyyy-mm-dd") & "','yyyy-mm-dd')"
  strSQL = strSQL & " and " & strSta
  strSQL = strSQL & " ORDER BY stacode, ddate"

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
            If Not IsNull(ORARst!T) Then
                T_Norm(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T, 1)
                T_Norm(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
            End If
            k = i + 1: Exit For
          End If
        Next i
      ElseIf ORARst!STA = LastSTA Then
        If Not IsNull(ORARst!T) Then
            T_Norm(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).T = Round(ORARst!T, 1)
            T_Norm(i, DateDiff("d", CDate(BDate), CDate(ORARst!ddate)) + 1).ddate = ORARst!ddate
        End If
      End If
      
      ORARst.MoveNext
    Loop

  End If
  ORARst.Close
  

End Sub





'************** 在5日滑动气温序列中，找到第一个通过阈值的日期 ***************
'************** 在第一个通过阈值的日期前9日的逐日气温序列中，找到第一个达到阈值的日期 ***************
Public Sub FindSteadyThresholdDateNorm(Flag As String, Data_T5d() As TSeries, Data_T() As TSeries, idays%, threshold_min!, threshold_max!)
  Dim i%, j%
  Dim tempdate As Date, tempIdx_T5d%, tempIdx_T% '起始日、所在5日滑动平均气温序列的序号、所在日平均气温序列中的序号
  Dim iSteadyDays%
  
  For i = 1 To StaNum
    tempdate = CDate("1899-09-09")
    iSteadyDays = 0
    
    For j = 1 To idays
    
      If Data_T5d(i, j).T >= threshold_min And Data_T5d(i, j).T < threshold_max Then
        iSteadyDays = iSteadyDays + 1
      ElseIf Data_T5d(i, j).T < threshold_min Or Data_T5d(i, j).T >= threshold_max Then
        iSteadyDays = 0
      End If
      
      If iSteadyDays = 5 Then
        tempdate = Data_T5d(i, j).ddate
        tempIdx_T5d = j
        tempIdx_T = j + 4
        Exit For  '找到第一个达到阈值的日期就退出循环
      End If
      
    Next j
    
    If tempdate <> CDate("1899-09-09") Then
      tempdate = tempdate + 1 '先往后推一日，便于在下面的循环中，每循环一次往前推一日（第一次循环即回到“稳定”日
      tempIdx_T = tempIdx_T + 1 '先往后推一日，便于在下面的循环中，每循环一次往前推一日（第一次循环即回到“稳定”日）
      
      For j = 0 To 8 '在 T 序列中往前循环 9 日，查找满足条件的日期
        If Data_T(i, tempIdx_T - 1).T >= threshold_min And Data_T(i, tempIdx_T - 1).T < threshold_max Then
          tempdate = tempdate - 1
          tempIdx_T = tempIdx_T - 1
        End If
      Next j

    End If

    SS_STA(i).OnsetNorm = tempdate
    SS_STA(i).T_IdxNorm = tempIdx_T
    
    
  Next i
End Sub




'************** 在5日滑动气温序列中，找到稳定通过阈值的日期 ***************
'************** 在稳定通过阈值的日期前9日的逐日气温序列中，找到第一个达到阈值的日期 ***************
'Flag：Curr/ Comp/ Norm -- 当年/某年/常年、Data_T5d：T5D序列、iDays1：T5D日数、Data_T：T序列、iDays2：T日数、threshold_min/max：阈值
Public Sub FindSteadyThresholdDate(Flag As String, Data_T5d() As TSeries, Data_T() As TSeries, idays%, threshold_min!, threshold_max!)
  Dim i%, j%
  Dim tempDate1 As Date, tempIdx1_T5d%, tempIdx1_T% '第一个起始日、所在5日滑动平均气温序列的序号、所在日平均气温序列中的序号
  Dim tempDate2 As Date, tempIdx2_T5d%, tempIdx2_T% '第二个起始日、所在5日滑动平均气温序列的序号、所在日平均气温序列中的序号
  Dim tempDate3 As Date, tempIdx3_T5d%, tempIdx3_T% '第二个起始日、所在5日滑动平均气温序列的序号、所在日平均气温序列中的序号
  Dim iSteadyDays%
  
  Dim OnsetFlag$ '采用第一起始日（First）或者第二起始日（Second）
  
  For i = 1 To StaNum
    OnsetFlag = "First"
    tempDate1 = CDate("1899-09-09")
    tempDate2 = CDate("1899-09-09")
    tempDate3 = CDate("1899-09-09")
    iSteadyDays = 0
    tempIdx1_T5d = 0
    tempIdx1_T = 0
    tempIdx2_T5d = 0
    tempIdx2_T = 0
    tempIdx3_T5d = 0
    tempIdx3_T = 0
    
    For j = 1 To idays
      If Data_T5d(i, j).T = -9999 Then
        GoTo Nextloop
      End If
    
      If Data_T5d(i, j).T >= threshold_min And Data_T5d(i, j).T < threshold_max Then
        iSteadyDays = iSteadyDays + 1
        
        If iSteadyDays = 5 Then
          '第一次满足阈值时，则写入初次起始日
          If tempDate1 = CDate("1899-09-09") Then
            tempDate1 = Data_T5d(i, j).ddate
            tempIdx1_T5d = j
            tempIdx1_T = j + 4
          
          '第二次满足阈值时，则写入二次起始日
          ElseIf tempDate2 = CDate("1899-09-09") Then
            tempDate2 = Data_T5d(i, j).ddate
            tempIdx2_T5d = j
            tempIdx2_T = j + 4
            
          '第三次满足阈值时，则写入三次起始日
          ElseIf tempDate3 = CDate("1899-09-09") Then
            tempDate3 = Data_T5d(i, j).ddate
            tempIdx3_T5d = j
            tempIdx3_T = j + 4
            
            Exit For '满足第三次起始日条件后，则退出本站循环
          End If
        End If
        
      ElseIf Data_T5d(i, j).T < threshold_min Or Data_T5d(i, j).T >= threshold_max Then
        iSteadyDays = 0
      End If
      
Nextloop:
    Next j
    
    If iSteadyDays < 5 Then
      tempDate1 = CDate("1899-09-09")
      tempIdx1_T5d = 0
      tempIdx1_T = 0
      
      tempDate2 = CDate("1899-09-09")
      tempIdx2_T5d = 0
      tempIdx2_T = 0
      
      tempDate3 = CDate("1899-09-09")
      tempIdx3_T5d = 0
      tempIdx3_T = 0
    End If


    If tempDate1 <> CDate("1899-09-09") Then
      
      tempDate1 = tempDate1 - 9 '从9天前的第一天往后查找，找到的第一个就是起始日
      tempIdx1_T = tempIdx1_T - 9  '从9天前的第一天往后查找，找到的第一个就是起始日
      
      For j = 0 To 8 '在 T 序列中往后循环 9 日，查找满足条件的日期
        tempDate1 = tempDate1 + 1
        tempIdx1_T = tempIdx1_T + 1
        If Data_T(i, tempIdx1_T).T >= threshold_min And Data_T(i, tempIdx1_T).T < threshold_max Then
          Exit For
        End If
      Next j
    End If

    If tempDate2 <> CDate("1899-09-09") Then
      tempDate2 = tempDate2 - 9 '从9天前的第一天往后查找，找到的第一个就是起始日
      tempIdx2_T = tempIdx2_T - 9 '从9天前的第一天往后查找，找到的第一个就是起始日
      
      For j = 0 To 8 '在 T 序列中往后循环 9 日，查找满足条件的日期
        tempDate2 = tempDate2 + 1
        tempIdx2_T = tempIdx2_T + 1
        If Data_T(i, tempIdx2_T + 1).T >= threshold_min And Data_T(i, tempIdx2_T + 1).T < threshold_max Then
          Exit For
        End If
      Next j
    End If

    If tempDate3 <> CDate("1899-09-09") Then
      tempDate3 = tempDate3 - 9 '从9天前的第一天往后查找，找到的第一个就是起始日
      tempIdx3_T = tempIdx3_T - 9 '从9天前的第一天往后查找，找到的第一个就是起始日
      
      For j = 0 To 8 '在 T 序列中往后循环 9 日，查找满足条件的日期
        tempDate3 = tempDate3 + 1
        tempIdx3_T = tempIdx3_T + 1
        If Data_T(i, tempIdx3_T + 1).T >= threshold_min And Data_T(i, tempIdx3_T + 1).T < threshold_max Then
          Exit For
        End If
      Next j
    End If
    

    
    '如果初次起始日没有偏早15天以上，则采用初次起始日为起始日 OnsetFlag = "First"
     
''''    ************** QXT 152-2012 ******************
''''    '如果初次起始日比常年起始日偏早15天，则进行二次判断
''''    If (tempIdx1_T - SS_STA(i).T_IdxNorm) < -15 Then
''''
''''      For j = tempIdx1_T To SS_STA(i).T_IdxNorm
''''        '如果初次起始日至常年起始日间有不满足阈值的情况，则 OnsetFlag="Second"
''''        If Data_T5d(i, j).T < threshold_min Or Data_T5d(i, j).T >= threshold_max Then
''''          OnsetFlag = "Second"
''''          Exit For
''''        End If
''''      Next j
''''
''''      '如果出现了上述情况，则继续比较初次和二次起始日之间，满足指标的日数和不满足指标的日数的大小
''''      If OnsetFlag = "Second" Then
''''        If (j - tempIdx1_T + 1) >= (tempIdx2_T - j - 1) Then
''''          OnsetFlag = "First"
''''        End If
''''      End If
''''
''''    End If
    
    
    '如果初次起始日没有偏早30天以上，则采用初次起始日为起始日 OnsetFlag = "First"
         
''''    ************** GBT 42074-2022 ******************
     '如果初次起始日比常年起始日偏早30天，则进行二次判断
    If (tempIdx1_T - SS_STA(i).T_IdxNorm) <= -30 Then
      
      If (tempIdx2_T - SS_STA(i).T_IdxNorm) > -30 Then '如果二次起始日比常年起始日偏早小于30天，则以二次起始日为季节起始日
        OnsetFlag = "Second"
      
      ElseIf (tempIdx2_T - SS_STA(i).T_IdxNorm) <= -30 Then '如果二次起始日比常年起始日偏早30天，则进行三次判断
      
        If (tempIdx3_T - SS_STA(i).T_IdxNorm) > -30 Then '如果三次起始日比常年起始日偏早小于30天，则以三次起始日为季节起始日
          OnsetFlag = "Third"
        End If
      
      End If

    End If
    

    If Flag = "Curr" Then
      If OnsetFlag = "First" Then
        SS_STA(i).OnsetCurr = tempDate1
        SS_STA(i).T_IdxCurr = tempIdx1_T
      ElseIf OnsetFlag = "Second" Then
        SS_STA(i).OnsetCurr = tempDate2
        SS_STA(i).T_IdxCurr = tempIdx2_T
      ElseIf OnsetFlag = "Third" Then
        SS_STA(i).OnsetCurr = tempDate3
        SS_STA(i).T_IdxCurr = tempIdx3_T
      End If
    
    ElseIf Flag = "Comp" Then
      If OnsetFlag = "First" Then
        SS_STA(i).OnsetComp = tempDate1
        SS_STA(i).T_IdxComp = tempIdx1_T
      ElseIf OnsetFlag = "Second" Then
        SS_STA(i).OnsetComp = tempDate2
        SS_STA(i).T_IdxComp = tempIdx2_T
      ElseIf OnsetFlag = "Third" Then
        SS_STA(i).OnsetComp = tempDate3
        SS_STA(i).T_IdxComp = tempIdx3_T
      End If
    
    End If
    
    
  Next i
  
End Sub
