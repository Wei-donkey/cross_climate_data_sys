Attribute VB_Name = "Output"
'小时模块填充表格1-通用
Public Sub HFGrid1_Fill_HOR(Frm As Object, DataFill!(), iTimes%)

    With Frm.HFGrid1
    .Redraw = False
    .Refresh
    For j = 1 To iTimes  '先按列循环
      For i = 1 To StaNum
        If DataFill(i, j) <> -9999 Then
          .TextMatrix(i, j + 2) = Format(DataFill(i, j), "0.0")
          If Frm.CheckHLight.Value = Checked Then '高亮阈值范围内数据
            If (DataFill(i, j) >= Frm.ComboMin_HL.Text) And (DataFill(i, j) <= Frm.ComboMax_HL.Text) Then
              .Row = i: .Col = j + 2: .CellBackColor = &HF0E0FF
            End If
          End If
        End If
      Next i
    Next j
    
    For j = 1 To 3
      For i = 1 To StaNum
        If COLS_Stat(i, j) <> -9999 Then
          .TextMatrix(i, j + 2 + iTimes) = Format(COLS_Stat(i, j), "0.0")
        End If
      Next i
    Next j
    
    .Refresh
    .Redraw = True
    End With

End Sub


'日旬月季年模块填充表格1-通用
Public Sub HFGrid1_Fill(Frm As Object, DataFill!(), iTimes%, XtrmYearFill%())

    With Frm.HFGrid1
    .Redraw = False
    .Refresh
    If Frm.XtrmYear = False Then
        For j = 1 To iTimes  '先按列循环
          For i = 1 To StaNum
            If DataFill(i, j) <> -9999 Then
              If Frm.FlagStat <> "DAYS" Then
                .TextMatrix(i, j + 2) = Format(DataFill(i, j), "0.0")
              ElseIf Frm.FlagStat = "DAYS" Then
                .TextMatrix(i, j + 2) = Format(DataFill(i, j), "0")
              End If
              If Frm.CheckHLight.Value = Checked Then '高亮阈值范围内数据
                If (DataFill(i, j) >= Frm.ComboMin_HL.Text) And (DataFill(i, j) <= Frm.ComboMax_HL.Text) Then
                  .Row = i: .Col = j + 2: .CellBackColor = &HF0E0FF
                End If
              End If
            End If
          Next i
        Next j
    
        For j = 1 To 3
          For i = 1 To StaNum
            If COLS_Stat(i, j) <> -9999 Then
              .TextMatrix(i, j + 2 + iTimes) = Format(COLS_Stat(i, j), "0.0")
            End If
          Next i
        Next j
    
    ElseIf Frm.XtrmYear = True Then
        For j = 1 To iTimes  '先按列循环
          For i = 1 To StaNum
            If DataFill(i, j) <> -9999 Then
              If Frm.FlagStat <> "DAYS" Then
                .TextMatrix(i * 2 - 1, j + 2) = Format(DataFill(i, j), "0.0")
              ElseIf Frm.FlagStat = "DAYS" Then
                .TextMatrix(i * 2 - 1, j + 2) = Format(DataFill(i, j), "0")
              End If
              If Frm.CheckHLight.Value = Checked Then '高亮阈值范围内数据
                If (DataFill(i, j) >= Frm.ComboMin_HL.Text) And (DataFill(i, j) <= Frm.ComboMax_HL.Text) Then
                  .Row = i * 2 - 1: .Col = j + 2: .CellBackColor = &HF0E0FF
                End If
              End If
            End If
            If XtrmYearFill(i, j) <> 1899 Then
              .TextMatrix(i * 2, j + 2) = XtrmYearFill(i, j)
            End If
          Next i
        Next j
    
        For j = 1 To 3
          For i = 1 To StaNum
            If COLS_Stat(i, j) <> -9999 Then
              .TextMatrix(i * 2 - 1, j + 2 + iTimes) = Format(COLS_Stat(i, j), "0.0")
            End If
          Next i
        Next j
    
    End If
    
        
    .Refresh
    .Redraw = True
    End With

End Sub


Public Sub HFGrid2_Fill(Frm As Object, DataFill!(), iTimes%)

    With Frm.HFGrid2
    .Redraw = False
    .Refresh
    For j = 1 To iTimes  '先按列循环
      For i = 1 To 3
        If DataFill(i, j) <> -9999 Then
          .TextMatrix(i, j + 2) = Format(DataFill(i, j), "0.0")
          If Frm.CheckHLight.Value = Checked Then '高亮阈值范围内数据
            If (DataFill(i, j) >= Frm.ComboMin_HL.Text) And (DataFill(i, j) <= Frm.ComboMax_HL.Text) Then
              .Row = i: .Col = j + 2: .CellBackColor = &HF0E0FF
            End If
          End If
        End If
      Next i
    Next j
    For j = 1 To 3
      For i = 1 To 3
        If COLS_ROWS_Stat(i, j) <> -9999 Then
          .TextMatrix(i, j + 2 + iTimes) = Format(COLS_ROWS_Stat(i, j), "0.0")
        End If
      Next i
    Next j
    .Refresh
    .Redraw = True
    End With

End Sub


Public Sub HFGrid3_fill(Frm As Object, DataFill!(), iTimes%, Index As Integer)
  Dim iRows%
  
  If FlagZone = "CITY" Then iRows = CityNum
  If FlagZone = "CNTY" Then iRows = CountyNum
  If FlagZone = "TOWN" Then iRows = TownNum
    
  With Frm.HFGrid3(Index)
  .Redraw = False
  .Refresh
  For j = 1 To iTimes  '先按列循环
    For k = 1 To iRows
      If DataFill(k, j) <> -9999 Then
'        If FlagZone = "CITY" Then .TextMatrix(k, j + 1) = Format(DataFill(k, j), "0.0")
'        If FlagZone = "CNTY" Then
        .TextMatrix(k, j + 2) = Format(DataFill(k, j), "0.0")
'        If FlagZone = "TOWN" Then .TextMatrix(k, j + 3) = Format(DataFill(k, j), "0.0")
        If Frm.CheckHLight.Value = Checked Then '高亮阈值范围内数据
          If (DataFill(k, j) >= Frm.ComboMin_HL.Text) And (DataFill(k, j) <= Frm.ComboMax_HL.Text) Then
            .Row = k
'            If FlagZone = "CITY" Then .Col = j + 1
'            If FlagZone = "CNTY" Then .Col = j + 2
'            If FlagZone = "TOWN" Then .Col = j + 3
            .Col = j + 2
            .CellBackColor = &HF0E0FF
          End If
        End If
      End If
    Next k
  Next j

  '统计多列的平均/最大/最小值
  Call Stat_COLS(DataFill, iTimes, iRows, "AVG")

  For j = 1 To 3
    For k = 1 To iRows
      If COLS_Stat(k, j) <> -9999 Then
'        If FlagZone = "CITY" Then .TextMatrix(k, j + 1 + iTimes) = Format(COLS_Stat(k, j), "0.0")
'        If FlagZone = "CNTY" Then
        .TextMatrix(k, j + 2 + iTimes) = Format(COLS_Stat(k, j), "0.0")
'        If FlagZone = "TOWN" Then .TextMatrix(k, j + 3 + iTimes) = Format(COLS_Stat(k, j), "0.0")
      End If
    Next k
  Next j

  
  .Refresh
  .Redraw = True
  End With
    
End Sub



'2025-06-01 添加小时和日模块的区域统计（主要是修复 多时次/多日累计列 统计错误的问题）,stat取值为AVG或SUM
Public Sub HFGrid3_fill_HorDay(Frm As Object, DataFill!(), DataFill2!(), iTimes%, Index As Integer)
  Dim iRows%
  
  If FlagZone = "CITY" Then iRows = CityNum
  If FlagZone = "CNTY" Then iRows = CountyNum
  If FlagZone = "TOWN" Then iRows = TownNum
    
  With Frm.HFGrid3(Index)
  .Redraw = False
  .Refresh
  For j = 1 To iTimes  '先按列循环
    For k = 1 To iRows
      If DataFill(k, j) <> -9999 Then
'        If FlagZone = "CITY" Then .TextMatrix(k, j + 1) = Format(DataFill(k, j), "0.0")
'        If FlagZone = "CNTY" Then
        .TextMatrix(k, j + 2) = Format(DataFill(k, j), "0.0")
'        If FlagZone = "TOWN" Then .TextMatrix(k, j + 3) = Format(DataFill(k, j), "0.0")
        If Frm.CheckHLight.Value = Checked Then '高亮阈值范围内数据
          If (DataFill(k, j) >= Frm.ComboMin_HL.Text) And (DataFill(k, j) <= Frm.ComboMax_HL.Text) Then
            .Row = k
'            If FlagZone = "CITY" Then .Col = j + 1
'            If FlagZone = "CNTY" Then .Col = j + 2
'            If FlagZone = "TOWN" Then .Col = j + 3
            .Col = j + 2
            .CellBackColor = &HF0E0FF
          End If
        End If
      End If
    Next k
  Next j

  
  For j = 1 To 3
    For k = 1 To iRows
      If DataFill2(k, j) <> -9999 Then
'        If FlagZone = "CITY" Then .TextMatrix(k, j + 1 + iTimes) = Format(COLS_Stat(k, j), "0.0")
'        If FlagZone = "CNTY" Then
        .TextMatrix(k, j + 2 + iTimes) = Format(DataFill2(k, j), "0.0")
'        If FlagZone = "TOWN" Then .TextMatrix(k, j + 3 + iTimes) = Format(COLS_Stat(k, j), "0.0")
      End If
    Next k
  Next j
  
  .Refresh
  .Redraw = True
  End With
    
End Sub




'填充任意时段数据进入HFGrid1表格控件：最大、最小统计结果
Public Sub HFGrid1_Fill_Period1(Frm As Object, DataFill() As Period_Result)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid1
  For j = 1 To 13  '先按列循环
    For i = 1 To StaNum
        
      If j = 1 Then tmpData = DataFill(i).stat
      If j = 2 Then tmpDate = DataFill(i).stat_date
      If j = 3 Then tmpData = DataFill(i).stat_hist
      If j = 4 Then tmpDate = DataFill(i).stat_hist_date
      If j = 5 Then tmpData = DataFill(i).stat_year
      If j = 6 Then tmpDate = DataFill(i).stat_year_date
      If j = 7 Then tmpData = DataFill(i).rank1
      If j = 8 Then tmpData = DataFill(i).rank2
      If j = 9 Then tmpStr = DataFill(i).rank_years
      If j = 10 Then tmpYear = DataFill(i).maxyear
      If j = 11 Then tmpData = DataFill(i).maxvalue
      If j = 12 Then tmpYear = DataFill(i).minyear
      If j = 13 Then tmpData = DataFill(i).minvalue
      
      If j = 1 Or j = 3 Or j = 5 Or j = 11 Or j = 13 Then
        If tmpData <> -9999 Then .TextMatrix(i, j + 2) = Format(tmpData, "0.0")
      ElseIf j = 2 Or j = 4 Or j = 6 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i, j + 2) = Format(tmpDate, "yyyy-mm-dd")
      ElseIf j = 7 Or j = 8 Then
        If tmpData <> -9999 Then .TextMatrix(i, j + 2) = tmpData
      ElseIf j = 9 Then
        If tmpStr <> "-9999" Then .TextMatrix(i, j + 2) = tmpStr
      ElseIf j = 10 Or j = 12 Then
        If tmpYear <> -9999 Then .TextMatrix(i, j + 2) = tmpYear
      End If
    Next i
  Next j
  
  End With

End Sub


'填充任意时段数据进入HFGrid1表格控件：平均、累计、条件查询的统计结果
Public Sub HFGrid1_Fill_Period2(Frm As Object, DataFill() As Period_Result)
  Dim tmpData!, tmpStr$, tmpYear%
  
  With Frm.HFGrid1
  For j = 1 To 12  '先按列循环
    For i = 1 To StaNum
        
      If j = 1 Then tmpData = DataFill(i).stat
      If j = 2 Then tmpData = DataFill(i).stat_hist
      If j = 3 Then tmpData = DataFill(i).stat_diff1
      If j = 4 Then tmpData = DataFill(i).stat_year
      If j = 5 Then tmpData = DataFill(i).stat_diff2
      If j = 6 Then tmpData = DataFill(i).rank1
      If j = 7 Then tmpData = DataFill(i).rank2
      If j = 8 Then tmpStr = DataFill(i).rank_years
      If j = 9 Then tmpYear = DataFill(i).maxyear
      If j = 10 Then tmpData = DataFill(i).maxvalue
      If j = 11 Then tmpYear = DataFill(i).minyear
      If j = 12 Then tmpData = DataFill(i).minvalue
      
      If j = 1 Or j = 2 Or j = 4 Or j = 10 Or j = 12 Then
        If tmpData <> -9999 Then .TextMatrix(i, j + 2) = Format(tmpData, "0.0")
      ElseIf j = 3 Or j = 5 Then
        If tmpData <> -9999 Then
          .TextMatrix(i, j + 2) = Format(tmpData, "0.0")
          If Frm.FlagStat = "SUM" Then
            If Frm.SelField_Initial = "R" Or Frm.SelField_Initial = "R08" Or Frm.SelField_Initial = "R20_08" Or Frm.SelField_Initial = "R08_20" Or Frm.SelField_Initial = "S" Then
              .TextMatrix(i, j + 2) = Format(tmpData, "0")
            End If
          End If
        End If
      ElseIf j = 6 Or j = 7 Then
        If tmpData <> -9999 Then .TextMatrix(i, j + 2) = tmpData
      ElseIf j = 8 Then
        If tmpStr <> "-9999" Then .TextMatrix(i, j + 2) = tmpStr
      ElseIf j = 9 Or j = 11 Then
        If tmpYear <> -9999 Then .TextMatrix(i, j + 2) = tmpYear
      End If
    Next i
  Next j
  
  End With

End Sub


'填充任意时段数据进入HFGrid2表格控件：最大、最小统计结果
Public Sub HFGrid2_Fill_Period1(Frm As Object, DataFill() As Period_Result)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid2
  For j = 1 To 13  '先按列循环
    For i = 1 To 3
        
      If j = 1 Then tmpData = DataFill(i).stat
      If j = 2 Then tmpDate = DataFill(i).stat_date
      If j = 3 Then tmpData = DataFill(i).stat_hist
      If j = 4 Then tmpDate = DataFill(i).stat_hist_date
      If j = 5 Then tmpData = DataFill(i).stat_year
      If j = 6 Then tmpDate = DataFill(i).stat_year_date
      If j = 7 Then tmpData = DataFill(i).rank1
      If j = 8 Then tmpData = DataFill(i).rank2
      If j = 9 Then tmpStr = DataFill(i).rank_years
      If j = 10 Then tmpYear = DataFill(i).maxyear
      If j = 11 Then tmpData = DataFill(i).maxvalue
      If j = 12 Then tmpYear = DataFill(i).minyear
      If j = 13 Then tmpData = DataFill(i).minvalue
      
      If j = 1 Or j = 3 Or j = 5 Or j = 11 Or j = 13 Then
        If tmpData <> -9999 Then .TextMatrix(i - 1, j + 2) = Format(tmpData, "0.0")
      ElseIf j = 2 Or j = 4 Or j = 6 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i - 1, j + 2) = Format(tmpDate, "yyyy-mm-dd")
      ElseIf j = 7 Or j = 8 Then
        If tmpData <> -9999 Then .TextMatrix(i - 1, j + 2) = tmpData
      ElseIf j = 9 Then
        If tmpStr <> "-9999" Then .TextMatrix(i - 1, j + 2) = tmpStr
      ElseIf j = 10 Or j = 12 Then
        If tmpYear <> -9999 Then .TextMatrix(i - 1, j + 2) = tmpYear
      End If
    Next i
  Next j
  
  End With

End Sub


'填充任意时段数据进入HFGrid1表格控件：最大、最小之外的统计结果
Public Sub HFGrid2_Fill_Period2(Frm As Object, DataFill() As Period_Result)
  Dim tmpData!, tmpStr$, tmpYear%
  
  With Frm.HFGrid2
  For j = 1 To 12  '先按列循环
    For i = 1 To 3
        
      If j = 1 Then tmpData = DataFill(i).stat
      If j = 2 Then tmpData = DataFill(i).stat_hist
      If j = 3 Then tmpData = DataFill(i).stat_diff1
      If j = 4 Then tmpData = DataFill(i).stat_year
      If j = 5 Then tmpData = DataFill(i).stat_diff2
      If j = 6 Then tmpData = DataFill(i).rank1
      If j = 7 Then tmpData = DataFill(i).rank2
      If j = 8 Then tmpStr = DataFill(i).rank_years
      If j = 9 Then tmpYear = DataFill(i).maxyear
      If j = 10 Then tmpData = DataFill(i).maxvalue
      If j = 11 Then tmpYear = DataFill(i).minyear
      If j = 12 Then tmpData = DataFill(i).minvalue
      
      If j = 1 Or j = 2 Or j = 4 Or j = 10 Or j = 12 Then
        If tmpData <> -9999 Then .TextMatrix(i - 1, j + 2) = Format(tmpData, "0.0")
      ElseIf j = 3 Or j = 5 Then
        If tmpData <> -9999 And Frm.FlagStat <> "SUM" Then .TextMatrix(i - 1, j + 2) = Format(tmpData, "0.0")
        If tmpData <> -9999 And Frm.FlagStat = "SUM" Then .TextMatrix(i - 1, j + 2) = Format(tmpData, "0")
      ElseIf j = 6 Or j = 7 Then
        If tmpData <> -9999 Then .TextMatrix(i - 1, j + 2) = tmpData
      ElseIf j = 8 Then
        If tmpStr <> "-9999" Then .TextMatrix(i - 1, j + 2) = tmpStr
      ElseIf j = 9 Or j = 11 Then
        If tmpYear <> -9999 Then .TextMatrix(i - 1, j + 2) = tmpYear
      End If
    Next i
  Next j
  
  End With

End Sub


'填充任意时段数据进入HFGrid3表格控件：最大、最小统计结果
Public Sub HFGrid3_fill_Period1(Frm As Object)
  Dim DataFill() As Period_Result  ' 需要填充到表格中的二维数据
  Dim tmpData!, iRows%
  Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
  
  If FlagZone = "CITY" Then iRows = CityNum ': StatLvl = 1
  If FlagZone = "CNTY" Then iRows = CountyNum ': StatLvl = 2
  If FlagZone = "TOWN" Then iRows = TownNum ': StatLvl = 3
  StatLvl = 2
  
    For l = 1 To 3
      If l = 1 Then DataFill = PRD_AVG
      If l = 2 Then DataFill = PRD_MAX
      If l = 3 Then DataFill = PRD_MIN
    
      With Frm.HFGrid3(l)
      For j = 1 To 10  '先按列循环
        For k = 1 To iRows
          If j = 1 Then
            If DataFill(k).stat <> -9999 Then .TextMatrix(k, StatLvl + 1) = Format(DataFill(k).stat, "0.0")
          ElseIf j = 2 Then
            If DataFill(k).stat_hist <> -9999 Then .TextMatrix(k, StatLvl + 3) = Format(DataFill(k).stat_hist, "0.0")
          ElseIf j = 3 Then
            If DataFill(k).stat_year <> -9999 Then .TextMatrix(k, StatLvl + 5) = Format(DataFill(k).stat_year, "0.0")
          ElseIf j = 4 Then
            If DataFill(k).rank1 <> -9999 Then .TextMatrix(k, StatLvl + 7) = DataFill(k).rank1
          ElseIf j = 5 Then
            If DataFill(k).rank2 <> -9999 Then .TextMatrix(k, StatLvl + 8) = DataFill(k).rank2
          ElseIf j = 6 Then
            If DataFill(k).rank_years <> "-9999" Then .TextMatrix(k, StatLvl + 9) = DataFill(k).rank_years
          ElseIf j = 7 Then
            If DataFill(k).maxyear <> -9999 Then .TextMatrix(k, StatLvl + 10) = DataFill(k).maxyear
          ElseIf j = 8 Then
            If DataFill(k).maxyear <> -9999 Then .TextMatrix(k, StatLvl + 11) = Format(DataFill(k).maxvalue, "0.0")
          ElseIf j = 9 Then
            If DataFill(k).minyear <> -9999 Then .TextMatrix(k, StatLvl + 12) = DataFill(k).minyear
          ElseIf j = 10 Then
            If DataFill(k).minyear <> -9999 Then .TextMatrix(k, StatLvl + 13) = Format(DataFill(k).minvalue, "0.0")
            
          End If
        
        Next k
      Next j
      
      For j = 1 To 3
        For k = 1 To iRows
          If j = 1 Then
            If DataFill(k).stat_date <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + 2) = Format(DataFill(k).stat_date, "yyyy-mm-dd")
          ElseIf j = 2 Then
            If DataFill(k).stat_hist_date <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + 4) = Format(DataFill(k).stat_hist_date, "yyyy-mm-dd")
          ElseIf j = 3 Then
            If DataFill(k).stat_year_date <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + 6) = Format(DataFill(k).stat_year_date, "yyyy-mm-dd")
          End If
        
        Next k
            
      Next j
      
        
      End With
    
    Next l
End Sub


'填充任意时段数据进入HFGrid3表格控件：平均、累计、条件日数等统计结果
Public Sub HFGrid3_fill_Period2(Frm As Object)
  Dim DataFill() As Period_Result  ' 需要填充到表格中的二维数据
  Dim tmpData!, iRows%
  Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
  
  If FlagZone = "CITY" Then iRows = CityNum ': StatLvl = 1
  If FlagZone = "CNTY" Then iRows = CountyNum ': StatLvl = 2
  If FlagZone = "TOWN" Then iRows = TownNum ': StatLvl = 3
  StatLvl = 2
  
    For l = 1 To 3
      If l = 1 Then DataFill = PRD_AVG
      If l = 2 Then DataFill = PRD_MAX
      If l = 3 Then DataFill = PRD_MIN
    
      With Frm.HFGrid3(l)
      For j = 1 To 12  '先按列循环
        For k = 1 To iRows
          If j = 1 Then
            If DataFill(k).stat <> -9999 Then .TextMatrix(k, StatLvl + 1) = Format(DataFill(k).stat, "0.0")
          ElseIf j = 2 Then
            If DataFill(k).stat_hist <> -9999 Then .TextMatrix(k, StatLvl + 2) = Format(DataFill(k).stat_hist, "0.0")
          ElseIf j = 3 Then
            If DataFill(k).stat_diff1 <> -9999 Then .TextMatrix(k, StatLvl + 3) = Format(DataFill(k).stat_diff1, "0.0")
          ElseIf j = 4 Then
            If DataFill(k).stat_year <> -9999 Then .TextMatrix(k, StatLvl + 4) = Format(DataFill(k).stat_year, "0.0")
          ElseIf j = 5 Then
            If DataFill(k).stat_diff2 <> -9999 Then .TextMatrix(k, StatLvl + 5) = Format(DataFill(k).stat_diff2, "0.0")
          ElseIf j = 6 Then
            If DataFill(k).rank1 <> -9999 Then .TextMatrix(k, StatLvl + 6) = DataFill(k).rank1
          ElseIf j = 7 Then
            If DataFill(k).rank2 <> -9999 Then .TextMatrix(k, StatLvl + 7) = DataFill(k).rank2
          ElseIf j = 8 Then
            If DataFill(k).rank_years <> "-9999" Then .TextMatrix(k, StatLvl + 8) = DataFill(k).rank_years
          ElseIf j = 9 Then
            If DataFill(k).maxyear <> -9999 Then .TextMatrix(k, StatLvl + 9) = DataFill(k).maxyear
          ElseIf j = 10 Then
            If DataFill(k).maxyear <> -9999 Then .TextMatrix(k, StatLvl + 10) = Format(DataFill(k).maxvalue, "0.0")
          ElseIf j = 11 Then
            If DataFill(k).minyear <> -9999 Then .TextMatrix(k, StatLvl + 11) = DataFill(k).minyear
          ElseIf j = 12 Then
            If DataFill(k).minyear <> -9999 Then .TextMatrix(k, StatLvl + 12) = Format(DataFill(k).minvalue, "0.0")
          End If
        Next k
      Next j
        
      End With
    
    Next l
End Sub




'填充任意时段数据进入HFGrid1表格控件：多要素统计结果：k标识所选要素的顺序
Public Sub HFGrid1_Fill_Period3(Frm As Object, DataFill() As Period_Result, k%, FlagStat$)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid1
  For j = 1 To 3  '先按列循环
    For i = 1 To StaNum
      If FlagStat = "MAX" Or FlagStat = "MIN" Then
        If j = 1 Then tmpData = DataFill(i).stat
        If j = 2 Then tmpDate = DataFill(i).stat_date
        If j = 3 Then tmpData = DataFill(i).stat_hist
        If j = 1 Or j = 3 Then
          If tmpData <> -9999 Then .TextMatrix(i, 3 * (k - 1) + j + 2) = Format(tmpData, "0.0")
        ElseIf j = 2 Then
          If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i, 3 * (k - 1) + j + 2) = Format(tmpDate, "yyyy-mm-dd")
        End If
      
      ElseIf FlagStat = "AVE" Then
        If j = 1 Then tmpData = DataFill(i).stat
        If j = 2 Then tmpData = DataFill(i).stat_hist
        If j = 3 Then tmpData = DataFill(i).stat_diff1
        If tmpData <> -9999 Then .TextMatrix(i, 3 * (k - 1) + j + 2) = Format(tmpData, "0.0")
      
      ElseIf FlagStat = "SUM" Then
        If j = 1 Then tmpData = DataFill(i).stat
        If j = 2 Then tmpData = DataFill(i).stat_hist
        If j = 3 Then tmpData = DataFill(i).stat_diff1
        If tmpData <> -9999 Then .TextMatrix(i, 3 * (k - 1) + j + 2) = Format(tmpData, "0.0")
        
        If j = 3 Then
          If Frm.SelField_Initial = "R" Or Frm.SelField_Initial = "R08" Or Frm.SelField_Initial = "R20_08" Or Frm.SelField_Initial = "R08_20" Or Frm.SelField_Initial = "S" Then
            .TextMatrix(i, 3 * (k - 1) + j + 2) = Format(tmpData, "0")
          End If
        End If
      End If
    Next i
  Next j
  
  End With

End Sub


'填充任意时段数据进入HFGrid2表格控件：多要素统计结果
Public Sub HFGrid2_Fill_Period3(Frm As Object, DataFill() As Period_Result, k%, FlagStat$)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid2
  For j = 1 To 3  '先按列循环
    For i = 1 To 3
      If FlagStat = "MAX" Or FlagStat = "MIN" Then
        If j = 1 Then tmpData = DataFill(i).stat
        If j = 2 Then tmpDate = DataFill(i).stat_date
        If j = 3 Then tmpData = DataFill(i).stat_hist
        If j = 1 Or j = 3 Then
          If tmpData <> -9999 Then .TextMatrix(i - 1, 3 * (k - 1) + j + 2) = Format(tmpData, "0.0")
        ElseIf j = 2 Then
          If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i - 1, 3 * (k - 1) + j + 2) = Format(tmpDate, "yyyy-mm-dd")
        End If
    
      ElseIf FlagStat = "AVE" Then
        If j = 1 Then tmpData = DataFill(i).stat
        If j = 2 Then tmpData = DataFill(i).stat_hist
        If j = 3 Then tmpData = DataFill(i).stat_diff1
        If tmpData <> -9999 Then .TextMatrix(i - 1, 3 * (k - 1) + j + 2) = Format(tmpData, "0.0")
      
      ElseIf FlagStat = "SUM" Then
        If j = 1 Then tmpData = DataFill(i).stat
        If j = 2 Then tmpData = DataFill(i).stat_hist
        If j = 3 Then tmpData = DataFill(i).stat_diff1
        If tmpData <> -9999 Then .TextMatrix(i - 1, 3 * (k - 1) + j + 2) = Format(tmpData, "0.0")
        
        If j = 3 Then
          If Frm.SelField_Initial = "R" Or Frm.SelField_Initial = "R08" Or Frm.SelField_Initial = "R20_08" Or Frm.SelField_Initial = "R08_20" Or Frm.SelField_Initial = "S" Then
            .TextMatrix(i - 1, 3 * (k - 1) + j + 2) = Format(tmpData, "0")
          End If
        End If
      End If
    Next i
  Next j
  
  End With

End Sub


'填充任意时段数据进入HFGrid3表格控件：多要素统计结果
Public Sub HFGrid3_fill_Period3(Frm As Object, m%, FlagStat$)
  Dim DataFill() As Period_Result  ' 需要填充到表格中的二维数据
  Dim tmpData!, iRows%
  Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
  
  If FlagZone = "CITY" Then iRows = CityNum ': StatLvl = 1
  If FlagZone = "CNTY" Then iRows = CountyNum ': StatLvl = 2
  If FlagZone = "TOWN" Then iRows = TownNum ': StatLvl = 3
  StatLvl = 2
  
    For l = 1 To 3
      If l = 1 Then DataFill = PRD_AVG
      If l = 2 Then DataFill = PRD_MAX
      If l = 3 Then DataFill = PRD_MIN
    
      With Frm.HFGrid3(l)
      For j = 1 To 3  '先按列循环
        For k = 1 To iRows
          If FlagStat = "MAX" Or FlagStat = "MIN" Then
            If j = 1 Then
              If DataFill(k).stat <> -9999 Then .TextMatrix(k, 3 * (m - 1) + StatLvl + 1) = Format(DataFill(k).stat, "0.0")
            ElseIf j = 2 Then
              If DataFill(k).stat_date <> CDate("1899-09-09") Then .TextMatrix(k, 3 * (m - 1) + StatLvl + 2) = Format(DataFill(k).stat_date, "yyyy-mm-dd")
            ElseIf j = 3 Then
              If DataFill(k).stat_hist <> -9999 Then .TextMatrix(k, 3 * (m - 1) + StatLvl + 3) = Format(DataFill(k).stat_hist, "0.0")
            End If
          ElseIf FlagStat = "AVE" Then
            If j = 1 Then
              If DataFill(k).stat <> -9999 Then .TextMatrix(k, 3 * (m - 1) + StatLvl + 1) = Format(DataFill(k).stat, "0.0")
            ElseIf j = 2 Then
              If DataFill(k).stat_hist <> -9999 Then .TextMatrix(k, 3 * (m - 1) + StatLvl + 2) = Format(DataFill(k).stat_hist, "0.0")
            ElseIf j = 3 Then
              If DataFill(k).stat_diff1 <> -9999 Then .TextMatrix(k, 3 * (m - 1) + StatLvl + 3) = Format(DataFill(k).stat_diff1, "0.0")
            End If
          ElseIf FlagStat = "SUM" Then
            If j = 1 Then
              If DataFill(k).stat <> -9999 Then .TextMatrix(k, 3 * (m - 1) + StatLvl + 1) = Format(DataFill(k).stat, "0.0")
            ElseIf j = 2 Then
              If DataFill(k).stat_hist <> -9999 Then .TextMatrix(k, 3 * (m - 1) + StatLvl + 2) = Format(DataFill(k).stat_hist, "0.0")
            ElseIf j = 3 Then
              If DataFill(k).stat_diff1 <> -9999 Then .TextMatrix(k, 3 * (m - 1) + StatLvl + 3) = Format(DataFill(k).stat_diff1, "0.0")
              If Frm.SelField_Initial = "R" Or Frm.SelField_Initial = "R08" Or Frm.SelField_Initial = "R20_08" Or Frm.SelField_Initial = "R08_20" Or Frm.SelField_Initial = "S" Then
                .TextMatrix(k, 3 * (m - 1) + StatLvl + 3) = Format(DataFill(k).stat_diff1, "0")
              End If
            End If
          End If

        Next k
      Next j
        
      End With
    
    Next l
End Sub


'填充冷空气数据进入HFGrid1表格控件-各站点最强过程
Public Sub HFGrid1_Fill_ColdAir(Frm As Object, DataFill() As ColdAir_Result)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  Dim iRows%
  
  iRows = UBound(DataFill)
  With Frm.HFGrid1
  For j = 1 To 9  '先按列循环
    For i = 1 To iRows
        
      If j = 1 Then tmpData = DataFill(i).ilevel
      If j = 2 Then tmpDate = DataFill(i).BDate
      If j = 3 Then tmpDate = DataFill(i).EDate
      If j = 4 Then tmpData = DataFill(i).idays
      If j = 5 Then tmpData = DataFill(i).tv_24hmax
      If j = 6 Then tmpData = DataFill(i).tv_48hmax
      If j = 7 Then tmpData = DataFill(i).tv_acc
      If j = 8 Then tmpData = DataFill(i).T
      If j = 9 Then tmpData = DataFill(i).t_min
      
      If j = 1 Then
        If tmpData = 0 Then
          If Frm.Check_lvl(0).Value = Checked Then .TextMatrix(i, j + 2) = "None"
          If Frm.Check_lvl(0).Value = Unchecked Then .TextMatrix(i, j + 2) = ""
        End If
        If tmpData = 1 Then .TextMatrix(i, j + 2) = "Weak"
        If tmpData = 2 Then .TextMatrix(i, j + 2) = "Moderate"
        If tmpData = 3 Then .TextMatrix(i, j + 2) = "Strong"
        If tmpData = 4 Then .TextMatrix(i, j + 2) = "Cold Wave"
      ElseIf j = 2 Or j = 3 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i, j + 2) = Format(tmpDate, "yyyy-mm-dd")
      ElseIf j = 4 Then
        If tmpData <> -9999 Then .TextMatrix(i, j + 2) = Format(tmpData, "0")
      Else
        If tmpData <> -9999 Then .TextMatrix(i, j + 2) = Format(tmpData, "0.0")
      End If
      
    Next i
  Next j
  
  End With

End Sub



'填充冷空气数据进入HFGrid4表格控件-所有过程
Public Sub HFGrid4_Fill_ColdAir(Frm As Object, DataFill() As ColdAir_Result)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid4
  .Redraw = False
  
  .Rows = UBound(DataFill) + 1
  For j = 1 To 15  '先按列循环
    For i = 1 To .Rows - 1
        .Row = i: .Col = j - 1
        .CellAlignment = flexAlignLeftCenter
        If .Row <> 0 Then
          If .Col = 1 Then .CellBackColor = &H80000005
          If .Col = 2 Then .CellBackColor = &H80000005
        End If
      
      If j = 1 Then .TextMatrix(i, j - 1) = i
      If j = 2 Then .TextMatrix(i, j - 1) = DataFill(i).stacode
      If j = 3 Then .TextMatrix(i, j - 1) = DataFill(i).staname
      
      If j = 4 Then tmpData = DataFill(i).ilevel
      If j = 5 Then tmpDate = DataFill(i).BDate
      If j = 6 Then tmpDate = DataFill(i).EDate
      If j = 7 Then tmpData = DataFill(i).idays
      If j = 8 Then tmpData = DataFill(i).tv_24hmax
      If j = 9 Then tmpData = DataFill(i).tv_48hmax
      If j = 10 Then tmpData = DataFill(i).tv_acc
      If j = 11 Then tmpData = DataFill(i).T
      If j = 12 Then tmpData = DataFill(i).t_min
      
      If j = 13 Then .TextMatrix(i, j - 1) = DataFill(i).Town
      If j = 14 Then .TextMatrix(i, j - 1) = DataFill(i).County
      If j = 15 Then .TextMatrix(i, j - 1) = DataFill(i).City
      
      
      If j = 4 Then
        If tmpData = 0 Then .TextMatrix(i, j - 1) = "None"
        If tmpData = 1 Then .TextMatrix(i, j - 1) = "Weak"
        If tmpData = 2 Then .TextMatrix(i, j - 1) = "Moderate"
        If tmpData = 3 Then .TextMatrix(i, j - 1) = "Strong"
        If tmpData = 4 Then .TextMatrix(i, j - 1) = "Cold Wave"
      ElseIf j = 5 Or j = 6 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i, j - 1) = Format(tmpDate, "yyyy-mm-dd")
      ElseIf j = 7 Then
        If tmpData <> -9999 Then .TextMatrix(i, j - 1) = Format(tmpData, "0")
      ElseIf j >= 8 And j <= 12 Then
        If tmpData <> -9999 Then .TextMatrix(i, j - 1) = Format(tmpData, "0.0")
      End If
      
    Next i
  Next j
        
  .Redraw = True
  End With

End Sub

'填充冷空气数据进入HFGrid2表格控件：最大、最小统计结果
Public Sub HFGrid2_Fill_ColdAir(Frm As Object, DataFill() As ColdAir_Result)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid2
  For j = 1 To 9  '先按列循环
    For i = 1 To 3
        
      If j = 1 Then tmpData = DataFill(i).ilevel
      If j = 2 Then tmpDate = DataFill(i).BDate
      If j = 3 Then tmpDate = DataFill(i).EDate
      If j = 4 Then tmpData = DataFill(i).idays
      If j = 5 Then tmpData = DataFill(i).tv_24hmax
      If j = 6 Then tmpData = DataFill(i).tv_48hmax
      If j = 7 Then tmpData = DataFill(i).tv_acc
      If j = 8 Then tmpData = DataFill(i).T
      If j = 9 Then tmpData = DataFill(i).t_min
      
      If j = 1 Then
        If i <> 1 And tmpData <> -9999 Then
          If tmpData = 0 Then .TextMatrix(i - 1, j + 2) = "None"
          If tmpData = 1 Then .TextMatrix(i - 1, j + 2) = "Weak"
          If tmpData = 2 Then .TextMatrix(i - 1, j + 2) = "Moderate"
          If tmpData = 3 Then .TextMatrix(i - 1, j + 2) = "Strong"
          If tmpData = 4 Then .TextMatrix(i - 1, j + 2) = "Cold Wave"
        End If
      ElseIf j = 2 Or j = 3 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i - 1, j + 2) = Format(tmpDate, "yyyy-mm-dd")
      Else
        If tmpData <> -9999 Then .TextMatrix(i - 1, j + 2) = Format(tmpData, "0.0")
      End If
    Next i
  Next j
  
  End With

End Sub



'填充冷空气数据进入HFGrid5表格控件：最大、最小统计结果
Public Sub HFGrid5_Fill_ColdAir(Frm As Object, DataFill() As ColdAir_Result)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid5
  For j = 1 To 9  '先按列循环
    For i = 1 To 3
        
      If j = 1 Then tmpData = DataFill(i).ilevel
      If j = 2 Then tmpDate = DataFill(i).BDate
      If j = 3 Then tmpDate = DataFill(i).EDate
      If j = 4 Then tmpData = DataFill(i).idays
      If j = 5 Then tmpData = DataFill(i).tv_24hmax
      If j = 6 Then tmpData = DataFill(i).tv_48hmax
      If j = 7 Then tmpData = DataFill(i).tv_acc
      If j = 8 Then tmpData = DataFill(i).T
      If j = 9 Then tmpData = DataFill(i).t_min
      
      If j = 1 Then
        If i <> 1 And tmpData <> -9999 Then
          If tmpData = 0 Then .TextMatrix(i - 1, j + 2) = "None"
          If tmpData = 1 Then .TextMatrix(i - 1, j + 2) = "Weak"
          If tmpData = 2 Then .TextMatrix(i - 1, j + 2) = "Moderate"
          If tmpData = 3 Then .TextMatrix(i - 1, j + 2) = "Strong"
          If tmpData = 4 Then .TextMatrix(i - 1, j + 2) = "Cold Wave"
        End If
      ElseIf j = 2 Or j = 3 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i - 1, j + 2) = Format(tmpDate, "yyyy-mm-dd")
      Else
        If tmpData <> -9999 Then .TextMatrix(i - 1, j + 2) = Format(tmpData, "0.0")
      End If
    Next i
  Next j
  
  End With

End Sub




'填充冷空气数据进入HFGrid3表格控件：最大、最小统计结果
Public Sub HFGrid3_fill_ColdAir(Frm As Object)
  Dim DataFill() As ColdAir_Result  ' 需要填充到表格中的二维数据
  Dim tmpData!, iRows%
  Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
  
  If FlagZone = "CITY" Then iRows = CityNum ': StatLvl = 1
  If FlagZone = "CNTY" Then iRows = CountyNum ': StatLvl = 2
  If FlagZone = "TOWN" Then iRows = TownNum ': StatLvl = 3
  StatLvl = 2
  
    For l = 1 To 3
      If l = 1 Then DataFill = CA_AVG
      If l = 2 Then DataFill = CA_MAX
      If l = 3 Then DataFill = CA_MIN
    
      With Frm.HFGrid3(l)
      For j = 1 To 9  '先按列循环
        For k = 1 To iRows
          If j = 1 Then
            If DataFill(k).ilevel <> -9999 Then
              If DataFill(k).ilevel = 0 Then .TextMatrix(k, StatLvl + 1) = "None"
              If DataFill(k).ilevel = 1 Then .TextMatrix(k, StatLvl + 1) = "Weak"
              If DataFill(k).ilevel = 2 Then .TextMatrix(k, StatLvl + 1) = "Moderate"
              If DataFill(k).ilevel = 3 Then .TextMatrix(k, StatLvl + 1) = "Strong"
              If DataFill(k).ilevel = 4 Then .TextMatrix(k, StatLvl + 1) = "Cold Wave"
            End If
          ElseIf j = 2 Then
            If DataFill(k).BDate <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + 2) = Format(DataFill(k).BDate, "yyyy-mm-dd")
          ElseIf j = 3 Then
            If DataFill(k).EDate <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + 3) = Format(DataFill(k).EDate, "yyyy-mm-dd")
          ElseIf j = 4 Then
            If DataFill(k).idays <> -9999 Then .TextMatrix(k, StatLvl + 4) = Format(DataFill(k).idays, "0.0")
          ElseIf j = 5 Then
            If DataFill(k).tv_24hmax <> -9999 Then .TextMatrix(k, StatLvl + 5) = Format(DataFill(k).tv_24hmax, "0.0")
          ElseIf j = 6 Then
            If DataFill(k).tv_48hmax <> -9999 Then .TextMatrix(k, StatLvl + 6) = Format(DataFill(k).tv_48hmax, "0.0")
          ElseIf j = 7 Then
            If DataFill(k).tv_acc <> -9999 Then .TextMatrix(k, StatLvl + 7) = Format(DataFill(k).tv_acc, "0.0")
          ElseIf j = 8 Then
            If DataFill(k).T <> -9999 Then .TextMatrix(k, StatLvl + 8) = Format(DataFill(k).T, "0.0")
          ElseIf j = 9 Then
            If DataFill(k).t_min <> -9999 Then .TextMatrix(k, StatLvl + 9) = Format(DataFill(k).t_min, "0.0")
          End If
        Next k
      Next j
        
      End With
    
    Next l
End Sub


'保存FrmHour窗体中的所有可见HFGrid表格
Public Sub Output_HFGrid_Hour(Frm As Object, SelField_Hor$)
  Dim i%, j%
  Dim strTime As Date, iTimes%
  Dim FlagSaveLimit As Boolean  '2023-10-25 存储限制标识
  
  Dim strFileName$, strFileType$, strTemp$, strStacode$

  FrmMain.CommonDialogSave.Filter = "Text Files (*.csv)|*.csv": strFileType = "csv"
  FrmMain.CommonDialogSave.FilterIndex = 1
  
  FrmMain.CommonDialogSave.FileName = "Hourly Data"
  
  ' 判断数据序列是否超出长度，超出则需要限制
  FlagSaveLimit = False
  Dim idays%
  idays = DateDiff("d", CDate(Frm.BTime), CDate(Frm.ETime)) + 1
  If idays >= 31 Then FlagSaveLimit = True
  
  If FlagSaveLimit = True Then '如果有限制下载，则提示
    If userID = "liuw" Then '超级管理员-IP无限制
      
    Else '其他所有用户-IP限制
        
      If IP_SaveData = "" Then
        MsgBox "This account has no registered IP for long-series data downloads -- cannot save data", vbInformation, "Notice"
        Exit Sub
      End If
        
      If IP_Ban = False Then 'IP地址匹配
        FrmMain.StatusBar1.Panels(1).Text = "This machine's IP is registered; saving data"
      ElseIf IP_Ban = True Then 'IP地址不匹配
        MsgBox "Registered account IP does not match this machine's IP -- cannot save data", vbInformation, "Notice"
        Exit Sub
      End If
        
    End If
  End If
  
  
  FrmMain.CommonDialogSave.InitDir = App.Path & "\Output"
  FrmMain.CommonDialogSave.Flags = &H2
  FrmMain.CommonDialogSave.ShowSave
  If FrmMain.CommonDialogSave.Flags = 2 Then Screen.MousePointer = 1: FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub
  
  frmWait.Label1.Caption = "Saving, please wait"
  frmWait.Show
  DoEvents
  Screen.MousePointer = 11
  FrmMain.StatusBar1.Panels(1).Text = "Saving"
  
  strFileName = FrmMain.CommonDialogSave.FileName
  If Len(Dir(strFileName)) <> 0 Then Kill strFileName '如果存在同名文件则删除同名文件


  iTimes = Frm.iHors

    Open strFileName For Output As #FileNum
    
    If Frm.HFGrid1.Visible = True Then
        With Frm.HFGrid1
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To iTimes
          If Frm.DataType = "Multi" Then strTime = DateAdd("h", j - 1, CDate(Frm.BTime))
          If Frm.DataType = "Single" Then strTime = DateAdd("d", j - 1, CDate(Frm.BTime))
          strTemp = strTemp & "," & Format(strTime, "yyyy-m-d h:mm")
        Next j
        For j = iTimes + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        
        End With
    End If
    
    If Frm.HFGrid2.Visible = True Then
        With Frm.HFGrid2
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To iTimes
          If Frm.DataType = "Multi" Then strTime = DateAdd("h", j - 1, CDate(Frm.BTime))
          If Frm.DataType = "Single" Then strTime = DateAdd("d", j - 1, CDate(Frm.BTime))
          strTemp = strTemp & "," & Format(strTime, "yyyy-m-d h:mm")
        Next j
        For j = iTimes + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
     
    If Frm.HFGrid3(1).Visible = True Then
        With Frm.HFGrid3(1)
        
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To iTimes
          If Frm.DataType = "Multi" Then strTime = DateAdd("h", j - 1, CDate(Frm.BTime))
          If Frm.DataType = "Single" Then strTime = DateAdd("d", j - 1, CDate(Frm.BTime))
          strTemp = strTemp & "," & Format(strTime, "yyyy-m-d h:mm")
        Next j
        For j = iTimes + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    If Frm.HFGrid3(2).Visible = True Then
        With Frm.HFGrid3(2)
        
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To iTimes
          If Frm.DataType = "Multi" Then strTime = DateAdd("h", j - 1, CDate(Frm.BTime))
          If Frm.DataType = "Single" Then strTime = DateAdd("d", j - 1, CDate(Frm.BTime))
          strTemp = strTemp & "," & Format(strTime, "yyyy-m-d h:mm")
        Next j
        For j = iTimes + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    If Frm.HFGrid3(3).Visible = True Then
        With Frm.HFGrid3(3)
        
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To iTimes
          If Frm.DataType = "Multi" Then strTime = DateAdd("h", j - 1, CDate(Frm.BTime))
          If Frm.DataType = "Single" Then strTime = DateAdd("d", j - 1, CDate(Frm.BTime))
          strTemp = strTemp & "," & Format(strTime, "yyyy-m-d h:mm")
        Next j
        For j = iTimes + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    
    Close #FileNum
    FrmMain.StatusBar1.Panels(1).Text = "Save result: Done"
 
  
    Dim DataGap$, Spatial$, Temporal$, DataType$, DataLength$
    Dim Savetime As Date
  
    Savetime = Format(Now, "yyyy-mm-dd hh:mm:ss")

    idays = DateDiff("d", CDate(Frm.BTime), CDate(Frm.ETime)) + 1
    If idays < 15 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
    DataLength = CStr(idays) & "日"
    DataGap = "时"
    Temporal = Format(Frm.BTime, "yyyy-mm-dd") & "至" & Format(Frm.ETime, "yyyy-mm-dd")
    DataType = "Data"
    
'    Spatial = Replace(DataZones, "'", "")
    Spatial = Replace(DataExtent, "'", "")
   
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
  
    strSQL = "insert into T_OTHE_CROSS_DATA_DOWNLOAD(USERID,DDATETIME,IP_SAVE,STATYPE,DATAGAP,DATATYPE,SPATIAL,TEMPORAL,DATALENGTH,VARIABLE)"
    strSQL = strSQL + " Values('" & userID & "',to_date('" & Savetime & "', 'yyyy-mm-dd hh24:mi:ss'),'" & UserIP & "','" & StaType & "','" & DataGap & "','" & DataType & "','" & Spatial & "','" & Temporal & "','" & DataLength & "','" & SelField_Hor & "')"
    
    ORAConn4.Execute strSQL
    ORAConn4.Close
  
  frmWait.Hide: DoEvents
  Screen.MousePointer = 1
End Sub



'保存FrmDay窗体中的所有可见HFGrid表格
Public Sub Output_HFGrid_Day(Frm As Object, SelField_Day$)
  Dim i%, j%
  Dim strDate As Date
  Dim FlagSaveLimit As Boolean  '2023-10-25 存储限制标识
  
  Dim strFileName$, strFileType$, strTemp$, strStacode$

  FrmMain.CommonDialogSave.Filter = "Text Files (*.csv)|*.csv": strFileType = "csv"
  FrmMain.CommonDialogSave.FilterIndex = 1
  
  FrmMain.CommonDialogSave.FileName = "Daily Data"
   
  ' 判断数据序列是否超出长度，超出则需要限制
  FlagSaveLimit = False
  If Frm.idays > 90 Then FlagSaveLimit = True
   
  If FlagSaveLimit = True Then '如果有限制下载，则提示
    If userID = "liuw" Then '超级管理员-IP无限制
      
    Else '其他所有用户-IP限制
        
      If IP_SaveData = "" Then
        MsgBox "This account has no registered IP for long-series data downloads -- cannot save data", vbInformation, "Notice"
        Exit Sub
      End If
        
      If IP_Ban = False Then 'IP地址匹配
        FrmMain.StatusBar1.Panels(1).Text = "This machine's IP is registered; saving data"
      ElseIf IP_Ban = True Then 'IP地址不匹配
        MsgBox "Registered account IP does not match this machine's IP -- cannot save data", vbInformation, "Notice"
        Exit Sub
      End If
        
    End If
  End If
  
  
  FrmMain.CommonDialogSave.InitDir = App.Path & "\Output"
  FrmMain.CommonDialogSave.Flags = &H2
  FrmMain.CommonDialogSave.ShowSave
  If FrmMain.CommonDialogSave.Flags = 2 Then Screen.MousePointer = 1: FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub
  
  frmWait.Label1.Caption = "Saving, please wait"
  frmWait.Show
  DoEvents
  Screen.MousePointer = 11
  FrmMain.StatusBar1.Panels(1).Text = "Saving"
  
  strFileName = FrmMain.CommonDialogSave.FileName
  If Len(Dir(strFileName)) <> 0 Then Kill strFileName '如果存在同名文件则删除同名文件

    Open strFileName For Output As #FileNum
    
    If Frm.HFGrid1.Visible = True Then
        With Frm.HFGrid1
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To Frm.idays
          tempdate = CDate(Frm.BDate) + j - 1
          strTemp = strTemp & "," & Format(tempdate, "yyyy-m-d")
        Next j
        For j = Frm.idays + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        
        End With
    End If
    
    If Frm.HFGrid2.Visible = True Then
        With Frm.HFGrid2
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To Frm.idays
          tempdate = CDate(Frm.BDate) + j - 1
          strTemp = strTemp & "," & Format(tempdate, "yyyy-m-d")
        Next j
        For j = Frm.idays + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
     
    If Frm.HFGrid3(1).Visible = True Then
        With Frm.HFGrid3(1)
        
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To Frm.idays
          tempdate = CDate(Frm.BDate) + j - 1
          strTemp = strTemp & "," & Format(tempdate, "yyyy-m-d")
        Next j
        For j = Frm.idays + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    If Frm.HFGrid3(2).Visible = True Then
        With Frm.HFGrid3(2)
        
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To Frm.idays
          tempdate = CDate(Frm.BDate) + j - 1
          strTemp = strTemp & "," & Format(tempdate, "yyyy-m-d")
        Next j
        For j = Frm.idays + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    If Frm.HFGrid3(3).Visible = True Then
        With Frm.HFGrid3(3)
        
        '输出表头行
        strTemp = .TextMatrix(0, 0)
        For j = 1 To 2
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        For j = 1 To Frm.idays
          tempdate = CDate(Frm.BDate) + j - 1
          strTemp = strTemp & "," & Format(tempdate, "yyyy-m-d")
        Next j
        For j = Frm.idays + 3 To .Cols - 1
          strTemp = strTemp & "," & .TextMatrix(0, j)
        Next j
        Print #FileNum, strTemp
        
        For i = 1 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    
    Close #FileNum
    FrmMain.StatusBar1.Panels(1).Text = "Save result: Done"
 
  
    Dim DataGap$, Spatial$, Temporal$, DataType$, DataLength$
    Dim Savetime As Date
  
    Savetime = Format(Now, "yyyy-mm-dd hh:mm:ss")

    If Frm.idays < 180 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
    DataLength = CStr(Frm.idays) & "日"
    DataGap = "日"
    Temporal = Frm.BDate & "至" & Frm.EDate
    DataType = Frm.DataType
    
'    Spatial = Replace(DataZones, "'", "")
    Spatial = Replace(DataExtent, "'", "")
    
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
  
    strSQL = "insert into T_OTHE_CROSS_DATA_DOWNLOAD(USERID,DDATETIME,IP_SAVE,STATYPE,DATAGAP,DATATYPE,SPATIAL,TEMPORAL,DATALENGTH,VARIABLE)"
    strSQL = strSQL + " Values('" & userID & "',to_date('" & Savetime & "', 'yyyy-mm-dd hh24:mi:ss'),'" & UserIP & "','" & StaType & "','" & DataGap & "','" & DataType & "','" & Spatial & "','" & Temporal & "','" & DataLength & "','" & SelField_Day & "')"
    
    ORAConn4.Execute strSQL
    ORAConn4.Close

  
  frmWait.Hide: DoEvents
  Screen.MousePointer = 1
End Sub




'保存Frm窗体中的所有可见HFGrid表格
Public Sub Output_HFGrid(Frm As Object, SelField$)
  Dim i%, j%
  Dim FlagSaveLimit As Boolean  '2023-10-25 存储限制标识
  
  Dim strFileName$, strFileType$, strTemp$, strStacode$

  FrmMain.CommonDialogSave.Filter = "Text Files (*.csv)|*.csv": strFileType = "csv"
  FrmMain.CommonDialogSave.FilterIndex = 1
  
  If Frm Is FrmMeteoTen Then
    FrmMain.CommonDialogSave.FileName = "Dekad Data"
  ElseIf Frm Is FrmMeteoMon Then
    FrmMain.CommonDialogSave.FileName = "Monthly Data"
  ElseIf Frm Is FrmMeteoQtr Then
    FrmMain.CommonDialogSave.FileName = "Seasonal Data"
  ElseIf Frm Is FrmMeteoYer Then
    FrmMain.CommonDialogSave.FileName = "Annual Data"
  ElseIf Frm Is FrmMeteoPeriodSin Then
    FrmMain.CommonDialogSave.FileName = "单时段数据"
  ElseIf Frm Is FrmMeteoPeriod Then
    FrmMain.CommonDialogSave.FileName = "任意时段数据"
  ElseIf Frm Is FrmMeteoPeriodYer Then
    FrmMain.CommonDialogSave.FileName = "任意时段逐年数据"
  ElseIf Frm Is FrmMeteoPeriodSin Then
    FrmMain.CommonDialogSave.FileName = "单时段数据"
  ElseIf Frm Is FrmMeteoColdAir Then
    FrmMain.CommonDialogSave.FileName = "冷空气统计数据"
  ElseIf Frm Is FrmMeteoSeasons Then
    FrmMain.CommonDialogSave.FileName = "气候季节划分数据"
  End If
  
  
  ' 判断数据序列是否超出长度，超出则需要限制
  FlagSaveLimit = False
  If (Not Frm Is FrmMeteoPeriod) And (Not Frm Is FrmMeteoPeriodSin) And (Not Frm Is FrmMeteoSeasons) Then    '仅有任意/单时段/季节划分数据不限制下载
    If Frm Is FrmMeteoTen Then
      If Frm.iMons >= 3 Then FlagSaveLimit = True
   ElseIf Frm Is FrmMeteoMon Then
      If Frm.iMons >= 3 Then FlagSaveLimit = True
    ElseIf Frm Is FrmMeteoQtr Then
      If Frm.iYers >= 1 Then FlagSaveLimit = True
    ElseIf Frm Is FrmMeteoYer Then
      If Frm.iYers >= 3 Then FlagSaveLimit = True
    ElseIf Frm Is FrmMeteoPeriodYer Then
      If Frm.iYers >= 3 Then FlagSaveLimit = True
    ElseIf Frm Is FrmMeteoColdAir Then
      If Frm.idays >= 90 Then FlagSaveLimit = True
    End If
  End If
  
  
  If FlagSaveLimit = True Then '如果有限制下载，则提示
    If userID = "liuw" Then '超级管理员-IP无限制
      
    Else '其他所有用户-IP限制
        
      If IP_SaveData = "" Then
        MsgBox "This account has no registered IP for long-series data downloads -- cannot save data", vbInformation, "Notice"
        Exit Sub
      End If
        
      If IP_Ban = False Then 'IP地址匹配
        FrmMain.StatusBar1.Panels(1).Text = "This machine's IP is registered; saving data"
      ElseIf IP_Ban = True Then 'IP地址不匹配
        MsgBox "Registered account IP does not match this machine's IP -- cannot save data", vbInformation, "Notice"
        Exit Sub
      End If
        
    End If
  End If
  
  
  FrmMain.CommonDialogSave.InitDir = App.Path & "\Output"
  FrmMain.CommonDialogSave.Flags = &H2
  FrmMain.CommonDialogSave.ShowSave
  If FrmMain.CommonDialogSave.Flags = 2 Then Screen.MousePointer = 1: FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub
  
  frmWait.Label1.Caption = "Saving, please wait"
  frmWait.Show
  DoEvents
  Screen.MousePointer = 11
  FrmMain.StatusBar1.Panels(1).Text = "Saving"
  
  strFileName = FrmMain.CommonDialogSave.FileName
  If Len(Dir(strFileName)) <> 0 Then Kill strFileName '如果存在同名文件则删除同名文件

    Open strFileName For Output As #FileNum
    
    If Frm.HFGrid1.Visible = True Then
        With Frm.HFGrid1
        For i = 0 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    If Frm.HFGrid2.Visible = True Then
        With Frm.HFGrid2
        For i = 0 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
     
    If Frm.HFGrid3(1).Visible = True Then
        With Frm.HFGrid3(1)
        For i = 0 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    If Frm.HFGrid3(2).Visible = True Then
        With Frm.HFGrid3(2)
        For i = 0 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    If Frm.HFGrid3(3).Visible = True Then
        With Frm.HFGrid3(3)
        For i = 0 To .Rows - 1
          strTemp = .TextMatrix(i, 0)
          For j = 1 To .Cols - 1
            strTemp = strTemp & "," & .TextMatrix(i, j)
          Next j
          Print #FileNum, strTemp
        Next i
        End With
    End If
    
    If Frm Is FrmMeteoColdAir Then
        If Frm.HFGrid4.Visible = True Then
            With Frm.HFGrid4
            For i = 0 To .Rows - 1
              strTemp = .TextMatrix(i, 0)
              For j = 1 To .Cols - 1
                strTemp = strTemp & "," & .TextMatrix(i, j)
              Next j
              Print #FileNum, strTemp
            Next i
            End With
        End If
        
        If Frm.HFGrid5.Visible = True Then
            With Frm.HFGrid5
            For i = 0 To .Rows - 1
              strTemp = .TextMatrix(i, 0)
              For j = 1 To .Cols - 1
                strTemp = strTemp & "," & .TextMatrix(i, j)
              Next j
              Print #FileNum, strTemp
            Next i
            End With
        End If
    End If
    
    Close #FileNum
    FrmMain.StatusBar1.Panels(1).Text = "Save result: Done"
 
  
  If (Not Frm Is FrmMeteoPeriod) And (Not Frm Is FrmMeteoPeriodSin) And (Not Frm Is FrmMeteoSeasons) Then    '仅有任意/单时段/季节划分数据不记录下载
    Dim DataGap$, Spatial$, Temporal$, DataType$, DataLength$
    Dim Savetime As Date
  
    Savetime = Format(Now, "yyyy-mm-dd hh:mm:ss")

    If Frm Is FrmMeteoTen Then
      If Frm.iMons < 6 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
      DataLength = CStr(Frm.iMons) & "月"
      DataGap = "旬"
      Temporal = Frm.BYYMM & "至" & Frm.EYYMM
      DataType = Frm.DataType
   ElseIf Frm Is FrmMeteoMon Then
      If Frm.iMons < 6 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
      DataLength = CStr(Frm.iMons) & "月"
      DataGap = "月"
      Temporal = Frm.BYYMM & "至" & Frm.EYYMM
      DataType = Frm.DataType
    ElseIf Frm Is FrmMeteoQtr Then
      If Frm.iYers < 3 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
      DataLength = CStr(Frm.iYers) & "年"
      DataGap = "季"
      Temporal = Frm.BYYYY & "至" & Frm.EYYYY
      DataType = Frm.DataType
    ElseIf Frm Is FrmMeteoYer Then
      If Frm.iYers < 6 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
      DataLength = CStr(Frm.iYers) & "年"
      DataGap = "年"
      Temporal = Frm.BYYYY & "至" & Frm.EYYYY
      DataType = Frm.DataType
    ElseIf Frm Is FrmMeteoPeriodYer Then
      If Frm.iYers < 6 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
      DataLength = CStr(Frm.iYers) & "年"
      DataGap = "逐年时段"
      Temporal = CStr(Frm.BYear) & "至" & CStr(Frm.EYear)
      DataType = "Data"
    ElseIf Frm Is FrmMeteoColdAir Then
      If Frm.idays < 300 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
      DataLength = CStr(Frm.idays) & "日"
      DataGap = "Cold Air Outbreak"
      Temporal = Frm.BDate & "至" & Frm.EDate
      DataType = "Data"
    End If
    
'    Spatial = Replace(DataZones, "'", "")
    Spatial = Replace(DataExtent, "'", "")
    
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
  
    strSQL = "insert into T_OTHE_CROSS_DATA_DOWNLOAD(USERID,DDATETIME,IP_SAVE,STATYPE,DATAGAP,DATATYPE,SPATIAL,TEMPORAL,DATALENGTH,VARIABLE)"
    strSQL = strSQL + " Values('" & userID & "',to_date('" & Savetime & "', 'yyyy-mm-dd hh24:mi:ss'),'" & UserIP & "','" & StaType & "','" & DataGap & "','" & DataType & "','" & Spatial & "','" & Temporal & "','" & DataLength & "','" & SelField & "')"
    
    ORAConn4.Execute strSQL
    ORAConn4.Close

  End If
  
  frmWait.Hide: DoEvents
  Screen.MousePointer = 1
End Sub



'利用HFGrid1中的数据进行等值线图绘制
Public Sub Map_Value(Frm As Object, Color$)
  Dim tempNum%, tempStaco$, tempStana$, tempLon!, tempLat!, TempVal$
  '序号，站号，站名，经度，维度，数值
  Dim i%, j%, k%
  Dim sngMax!, sngMin!, sngInterval!
  Dim tempStr$
  
  If FlagZone <> "STA" Then '非站点尺度则不可绘等值线图
    MsgBox "Contour maps require station-scale results", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If
  
  If Frm.HFGrid1.Row <> 1 Then
    MsgBox "Click a column header to select it for plotting", vbInformation, "Notice": Exit Sub
  End If
  
  If IsExcludedColumnName(Frm.HFGrid1.TextMatrix(0, Frm.HFGrid1.Col)) Then
'    IsDataColumn = False
    MsgBox "Non-data columns cannot be plotted", vbInformation, "Notice": Exit Sub
  End If
  
  
  With Frm.HFGrid1
  FileNum = FreeFile
  Open App.Path & "\Temp\temp.txt" For Output As #FileNum
  Close #FileNum
  FileNum = FreeFile
  Open App.Path & "\Temp\temp.txt" For Append As #FileNum
  k = 0
  For i = 1 To .Rows - 1
    tempNum = i
    tempStaco = .TextMatrix(i, 1): tempStana = .TextMatrix(i, 2)
    For j = 1 To StaNum
      If StaInfo(j).stacode = tempStaco Then
        tempLon = StaInfo(j).Longitude: tempLat = StaInfo(j).Latitude: Exit For
      End If
    Next j
    TempVal = .TextMatrix(i, .Col)
    If TempVal <> "" Then
      Print #FileNum, tempLon; ","; tempLat; ","; TempVal
      k = k + 1
    End If
  Next i
  Close #FileNum
  End With
  
  If k <= 2 Then MsgBox "Not enough valid data -- this would cause a Surfer error; please reselect", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
  Dim TempSplit$() '拆分该行记录，第六个拆分变量为待绘图的值
  FileNum = FreeFile
  i = 0
  Open App.Path & "\Temp\Temp.txt" For Input As #FileNum
  TempVal = ""
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    TempSplit = Split(tempStr, ",")
    If TempSplit(2) <> TempVal Then i = i + 1: TempVal = TempSplit(2)
  Loop
  Close #FileNum
  
  If i <= 1 Then MsgBox "No variation across stations -- this would cause a Surfer error; please reselect", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
  If (Frm Is FrmMeteoPeriod) Or (Frm Is FrmMeteoColdAir) Or (Frm Is FrmMeteoPeriodSin) Or (Frm Is FrmMeteoSeasons) Then
    sngMax = Int(Frm.HFGrid2.TextMatrix(1, Frm.HFGrid1.Col))
    sngMin = Int(Frm.HFGrid2.TextMatrix(2, Frm.HFGrid1.Col))
  Else
    sngMax = Int(Frm.HFGrid2.TextMatrix(2, Frm.HFGrid1.Col))
    sngMin = Int(Frm.HFGrid2.TextMatrix(3, Frm.HFGrid1.Col))
  End If
  sngInterval = (sngMax - sngMin) / 5
  
'****************************设置等值线图的相关参数**********************************
  FrmMapValue.sngMin = sngMin: FrmMapValue.sngMax = sngMax: FrmMapValue.sngInterval = sngInterval
  FrmMapValue.strColor = Color

  FrmMapValue.Show
  
End Sub






'利用HFGrid1中的数据进行等值线图绘制
Public Sub Map_ColdAir(Frm As Object)
  Dim tempNum%, tempStaco$, tempStana$, tempLon!, tempLat!, TempVal$
  '序号，站号，站名，经度，维度，数值
  Dim i%, j%, k%
  Dim sngMax!, sngMin!, sngInterval!, SrfVisible As Boolean

  Dim tempStr$
  
  Dim FlagTitle$
  Dim FlagLable$  '是否显示数据标签：Yes-是、No-否
  Dim strTitle$
  
  If FlagZone <> "STA" Then '非站点尺度则不可绘等值线图
    MsgBox "Contour maps require station-scale results", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If
  
  If Frm.HFGrid1.Row <> 1 Then
    MsgBox "Click a column header to select it for plotting", vbInformation, "Notice": Exit Sub
  End If
  
  With Frm.HFGrid1
  FileNum = FreeFile
  Open App.Path & "\Temp\temp.txt" For Output As #FileNum
  Close #FileNum
  FileNum = FreeFile
  Open App.Path & "\Temp\temp.txt" For Append As #FileNum
  k = 0
  For i = 1 To .Rows - 1
    tempNum = i
    tempStaco = .TextMatrix(i, 1): tempStana = .TextMatrix(i, 2)
    For j = 1 To StaNum
      If StaInfo(j).stacode = tempStaco Then
        tempLon = StaInfo(j).Longitude: tempLat = StaInfo(j).Latitude: Exit For
      End If
    Next j
    
    If .TextMatrix(i, .Col) = "None" Then
      TempVal = 0
    ElseIf .TextMatrix(i, .Col) = "Weak" Then
      TempVal = 1
    ElseIf .TextMatrix(i, .Col) = "Moderate" Then
      TempVal = 2
    ElseIf .TextMatrix(i, .Col) = "Strong" Then
      TempVal = 3
    ElseIf .TextMatrix(i, .Col) = "Cold Wave" Then
      TempVal = 4
    End If
    
    If TempVal <> "" Then
      Print #FileNum, tempLon; ","; tempLat; ","; TempVal
      k = k + 1
    End If
  Next i
  Close #FileNum
  End With
  
  If k <= 2 Then MsgBox "Not enough valid data -- this would cause a Surfer error; please reselect", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
  Dim TempSplit$() '拆分该行记录，第六个拆分变量为待绘图的值
  FileNum = FreeFile
  i = 0
  Open App.Path & "\Temp\Temp.txt" For Input As #FileNum
  TempVal = ""
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    TempSplit = Split(tempStr, ",")
    If TempSplit(2) <> TempVal Then i = i + 1: TempVal = TempSplit(2)
  Loop
  Close #FileNum
  
  If i <= 1 Then MsgBox "No variation across stations -- this would cause a Surfer error; please reselect", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
  
'  If (Frm Is FrmMeteoPeriod) Or (Frm Is FrmMeteoColdAir) Then
'    sngMax = Int(Frm.HFGrid2.TextMatrix(1, Frm.HFGrid1.Col))
'    sngMin = Int(Frm.HFGrid2.TextMatrix(2, Frm.HFGrid1.Col))
'  Else
'    sngMax = Int(Frm.HFGrid2.TextMatrix(2, Frm.HFGrid1.Col))
'    sngMin = Int(Frm.HFGrid2.TextMatrix(3, Frm.HFGrid1.Col))
'  End If
'  sngInterval = (sngMax - sngMin) / 5
'
''****************************设置等值线图的相关参数**********************************
'  FrmMapValue.sngMin = sngMin: FrmMapValue.sngMax = sngMax: FrmMapValue.sngInterval = sngInterval
'  FrmMapValue.strColor = Color
'
'  FrmMapValue.Show
  
  sngMax = 4: sngMin = 0: sngInterval = 1: SrfVisible = False: FlagTitle = "No": FlagLable = "No": strTitle = ""
  Call DrawContourMap(sngMax, sngMin, sngInterval, SrfVisible, "COLDAIR", FlagTitle, strTitle, FlagLable)
  
  
  
  
End Sub



Public Sub DrawContourMap(sngMax!, sngMin!, sngInterval!, SrfVisible As Boolean, Object$, FlagTitle$, Title$, FlagLable$)
'''  Dim iNumCols%, iNumRows% '由经纬度的差值计算
  Dim TempVal$, tempStr$
  Dim i%
  Dim tempSrfFile$
  Dim outWidth%, outHeight% '输出图形宽度、高度
  
  Screen.MousePointer = 11
  
  If Object = "COLDAIR" Then tempSrfFile = srfPath_ColdAir
  If Object = "VALUE" Then tempSrfFile = srfPath
  
  If SrfVisible = True Then
    SrfApp.Visible = True
  Else
    SrfApp.Visible = False
  End If
  FrmMain.StatusBar1.Panels(1).Text = "Defining Plot: opening .srf template file"
  Set Plot = SrfApp.Documents.Open(tempSrfFile)
  
  FrmMain.StatusBar1.Panels(1).Text = "Defining Shapes"
  Set Shapes = Plot.Shapes
  
  FrmMain.StatusBar1.Panels(1).Text = "Defining MapFrame"
  Set MapFrame = Shapes.Item("Map")
  
  FrmMain.StatusBar1.Panels(1).Text = "Defining ContourMap"
  Set ContourMap = MapFrame.Overlays.Item("Contours")
  
  FrmMain.StatusBar1.Panels(1).Text = "Defining ContourLevels"
  Set ContourLevels = ContourMap.Levels
  
    
    
  '调用SURFER进行Kriging插值
  FrmMain.StatusBar1.Panels(1).Text = "Calling Surfer to interpolate"
    
  If Object = "COLDAIR" Then
    SrfApp.GridData DataFile:=App.Path & "\Temp\Temp.txt", xCol:=1, yCol:=2, zCol:=3, _
    xMin:=sngxMin, xMax:=sngxMax, yMin:=sngyMin, yMax:=sngyMax, NumCols:=iNumCols, NumRows:=iNumRows, _
    algorithm:=Grid_Method, OutGrid:=App.Path & "\Temp\Temp.grd", ShowReport:=False
   
  ElseIf Object = "VALUE" Then
    SrfApp.GridData DataFile:=App.Path & "\Temp\Temp.txt", xCol:=1, yCol:=2, zCol:=3, _
    xMin:=sngxMin, xMax:=sngxMax, yMin:=sngyMin, yMax:=sngyMax, NumCols:=iNumCols, NumRows:=iNumRows, _
    algorithm:=Grid_Method, OutGrid:=App.Path & "\Temp\Temp.grd", ShowReport:=False
    
  End If
            
  FrmMain.StatusBar1.Panels(1).Text = "Masking the grid file"
  SrfApp.GridBlank InGrid:=App.Path & "\Temp\Temp.grd", BlankFile:=blnPath, _
  OutGrid:=App.Path & "\Temp\Temp.grd"
    
  FrmMain.StatusBar1.Panels(1).Text = "Replacing the contour source file with the masked temp.grd"
  ContourMap.GridFile = App.Path & "\Temp\Temp.grd"
    
  If Object = "VALUE" Then
    FrmMain.StatusBar1.Panels(1).Text = "Applying fill colors from the temp.lvl color-level file"
    ContourLevels.LoadFile (App.Path & "\Temp\Temp.lvl")
  End If

  If FlagTitle = "Yes" Then
    Plot.Shapes.Item("Title").Visible = True
    Plot.Shapes.Item("Title").Text = Plot.Shapes.Item("Title").Text & Title
  ElseIf FlagTitle = "No" Then
    Plot.Shapes.Item("Title").Visible = False
  End If
  
  
  '2023-11-29 添加数据标签图层控制
  If FlagLable = "Yes" Then
    FrmMain.StatusBar1.Panels(1).Text = "Replacing station data labels"
    Set PostMap = MapFrame.Overlays.Item("values")
    PostMap.SetInputData DataFileName:=App.Path & "\Temp\Temp.txt", xCol:=1, yCol:=2, LabCol:=3 ', SymCol:=4, AngleCol:=5
    PostMap.Visible = True
    
  End If
  
  
  '另存该srf文件至Output文件夹，使Template文件夹中始终存放原始GD.srf文件
  Plot.SaveAs App.Path & "\Temp\Temp.srf"
  FrmMain.StatusBar1.Panels(1).Text = "Exporting the contour map from Surfer"
  outWidth = CInt(Export_Width):  outHeight = CInt(Export_Width)
  Plot.Export FileName:=App.Path & "\Output\Map.png", Options:="Height:=" & outHeight & ",Width=" & outWidth

  If SrfVisible = False Then
    Shell "rundll32.exe C:\WINDOWS\system32\shimgvw.dll,ImageView_Fullscreen " & App.Path & "\Output\Map.png", vbNormalFocus
  End If
  
  If SrfVisible = True Then
  
  Else
    SrfApp.Quit
  End If
  
  
  FrmMain.StatusBar1.Panels(1).Text = "Contour map complete"
  Screen.MousePointer = 1
  
End Sub



'填充单时段数据进入HFGrid1表格控件：
Public Sub HFGrid1_Fill_PeriodSin(Frm As Object, DataFill() As Period_Result, FlagStat$)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid1
    For i = 1 To StaNum
      tmpData = DataFill(i).stat
      
      If FlagStat = "AVE" Then
        If tmpData <> -9999 Then .TextMatrix(i, 1 + 2) = Format(tmpData, "0.0")
      
      ElseIf FlagStat = "MAX" Then
        tmpDate = DataFill(i).stat_date
        If tmpData <> -9999 Then .TextMatrix(i, 2 + 2) = Format(tmpData, "0.0")
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i, 3 + 2) = Format(tmpDate, "yyyy-mm-dd")
      
      ElseIf FlagStat = "MIN" Then
        tmpDate = DataFill(i).stat_date
        If tmpData <> -9999 Then .TextMatrix(i, 4 + 2) = Format(tmpData, "0.0")
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i, 5 + 2) = Format(tmpDate, "yyyy-mm-dd")
      
      ElseIf FlagStat = "SUM" Then
        If tmpData <> -9999 Then .TextMatrix(i, 6 + 2) = Format(tmpData, "0.0")
        
      ElseIf FlagStat = "DAYS" Then
        If tmpData <> -9999 Then .TextMatrix(i, 7 + 2) = Format(tmpData, "0.0")
        
      End If
    Next i
  
  End With

End Sub


'填充单时段数据进入HFGrid2表格控件
Public Sub HFGrid2_Fill_PeriodSin(Frm As Object, DataFill() As Period_Result, FlagStat$)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid2
    For i = 1 To 3
      tmpData = DataFill(i).stat
      
      If FlagStat = "AVE" Then
        If tmpData <> -9999 Then .TextMatrix(i - 1, 1 + 2) = Format(tmpData, "0.0")
      
      ElseIf FlagStat = "MAX" Then
        tmpDate = DataFill(i).stat_date
        If tmpData <> -9999 Then .TextMatrix(i - 1, 2 + 2) = Format(tmpData, "0.0")
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i - 1, 3 + 2) = Format(tmpDate, "yyyy-mm-dd")
      
      ElseIf FlagStat = "MIN" Then
        tmpDate = DataFill(i).stat_date
        If tmpData <> -9999 Then .TextMatrix(i - 1, 4 + 2) = Format(tmpData, "0.0")
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i - 1, 5 + 2) = Format(tmpDate, "yyyy-mm-dd")
    
      ElseIf FlagStat = "SUM" Then
        If tmpData <> -9999 Then .TextMatrix(i - 1, 6 + 2) = Format(tmpData, "0.0")
      
      ElseIf FlagStat = "DAYS" Then
        If tmpData <> -9999 Then .TextMatrix(i - 1, 7 + 2) = Format(tmpData, "0.0")
        
      End If
    Next i
  
  End With

End Sub


'填充任意时段数据进入HFGrid3表格控件：多要素统计结果
Public Sub HFGrid3_fill_PeriodSin(Frm As Object, FlagStat$)
  Dim DataFill() As Period_Result  ' 需要填充到表格中的二维数据
  Dim tmpData!, tmpDate As Date, iRows%
  Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
  
  If FlagZone = "CITY" Then iRows = CityNum ': StatLvl = 1
  If FlagZone = "CNTY" Then iRows = CountyNum ': StatLvl = 2
  If FlagZone = "TOWN" Then iRows = TownNum ': StatLvl = 3
  StatLvl = 2
  
    For l = 1 To 3
      If l = 1 Then DataFill = PRD_AVG
      If l = 2 Then DataFill = PRD_MAX
      If l = 3 Then DataFill = PRD_MIN
    
      With Frm.HFGrid3(l)
        For k = 1 To iRows
          tmpData = DataFill(k).stat
          
          If FlagStat = "AVE" Then
            If tmpData <> -9999 Then .TextMatrix(k, StatLvl + 1) = Format(tmpData, "0.0")
          
          ElseIf FlagStat = "MAX" Then
              tmpDate = DataFill(k).stat_date
              If tmpData <> -9999 Then .TextMatrix(k, StatLvl + 2) = Format(tmpData, "0.0")
              If tmpDate <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + 3) = Format(tmpDate, "yyyy-mm-dd")
          
          
          ElseIf FlagStat = "MIN" Then
              tmpDate = DataFill(k).stat_date
              If tmpData <> -9999 Then .TextMatrix(k, StatLvl + 4) = Format(tmpData, "0.0")
              If tmpDate <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + 5) = Format(tmpDate, "yyyy-mm-dd")
          
          ElseIf FlagStat = "SUM" Then
              If tmpData <> -9999 Then .TextMatrix(k, StatLvl + 6) = Format(tmpData, "0.0")
          
          ElseIf FlagStat = "DAYS" Then
              If tmpData <> -9999 Then .TextMatrix(k, StatLvl + 7) = Format(tmpData, "0.0")
          
          End If

        Next k
        
      End With
    
    Next l
End Sub



'填充季节起始日数据进入HFGrid1表格
Public Sub HFGrid1_Fill_Seasons(Frm As Object, DataFill() As Season_Result)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  Dim iRows%
  
  iRows = UBound(DataFill)
  With Frm.HFGrid1
  For j = 1 To 5  '先按列循环
    For i = 1 To iRows
        
      If j = 1 Then tmpDate = DataFill(i).OnsetCurr
      If j = 2 Then tmpDate = DataFill(i).OnsetNorm
      If j = 3 Then tmpData = DataFill(i).diffNorm
      If j = 4 Then tmpDate = DataFill(i).OnsetComp
      If j = 5 Then tmpData = DataFill(i).diffComp
      
      If j = 1 Or j = 4 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i, j + 2) = Format(tmpDate, "yyyy-mm-dd")
      ElseIf j = 2 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i, j + 2) = Format(tmpDate, "mm-dd")
      ElseIf j = 3 Or j = 5 Then
        If tmpData <> -9999 Then .TextMatrix(i, j + 2) = Format(tmpData, "0")
      End If
      
    Next i
  Next j
  
  End With

End Sub



'填充季节起始日数据进入HFGrid2表格控件：最大、最小统计结果
Public Sub HFGrid2_Fill_Seasons(Frm As Object, DataFill() As Season_Result)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid2
  For j = 1 To 5  '先按列循环
    For i = 1 To 3

      If j = 1 Then tmpDate = DataFill(i).OnsetCurr
      If j = 2 Then tmpDate = DataFill(i).OnsetNorm
      If j = 3 Then tmpData = DataFill(i).diffNorm
      If j = 4 Then tmpDate = DataFill(i).OnsetComp
      If j = 5 Then tmpData = DataFill(i).diffComp

      If j = 1 Or j = 4 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i - 1, j + 2) = Format(tmpDate, "yyyy-mm-dd")
      ElseIf j = 2 Then
        If tmpDate <> CDate("1899-09-09") Then .TextMatrix(i - 1, j + 2) = Format(tmpDate, "mm-dd")
      ElseIf j = 3 Or j = 5 Then
        If tmpData <> -9999 Then .TextMatrix(i - 1, j + 2) = Format(tmpData, "0")
      End If
    Next i
  Next j
  
  End With

End Sub




'填充季节划分数据进入HFGrid3表格控件：最大、最小统计结果
Public Sub HFGrid3_fill_Seasons(Frm As Object)
  Dim DataFill() As Season_Result  ' 需要填充到表格中的二维数据
  Dim tmpData!, iRows%
  Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
  
  If FlagZone = "CITY" Then iRows = CityNum ': StatLvl = 1
  If FlagZone = "CNTY" Then iRows = CountyNum ': StatLvl = 2
  If FlagZone = "TOWN" Then iRows = TownNum ': StatLvl = 3
  StatLvl = 2
  
    For l = 1 To 3
      If l = 1 Then DataFill = SS_AVG
      If l = 2 Then DataFill = SS_MAX
      If l = 3 Then DataFill = SS_MIN
    
      With Frm.HFGrid3(l)
      For j = 1 To 5  '先按列循环
        For k = 1 To iRows
          If j = 1 Then
            If DataFill(k).OnsetCurr <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + j) = Format(DataFill(k).OnsetCurr, "yyyy-mm-dd")
          ElseIf j = 2 Then
            If DataFill(k).OnsetNorm <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + j) = Format(DataFill(k).OnsetNorm, "mm-dd")
          ElseIf j = 3 Then
            If DataFill(k).diffNorm <> -9999 Then .TextMatrix(k, StatLvl + j) = Format(DataFill(k).diffNorm, "0")
          ElseIf j = 4 Then
            If DataFill(k).OnsetComp <> CDate("1899-09-09") Then .TextMatrix(k, StatLvl + j) = Format(DataFill(k).OnsetComp, "yyyy-mm-dd")
          ElseIf j = 5 Then
            If DataFill(k).diffComp <> -9999 Then .TextMatrix(k, StatLvl + j) = Format(DataFill(k).diffComp, "0")
          End If
        Next k
      Next j
        
      End With
    
    Next l
End Sub


' 将不规范的分隔符修改为分号（;），避免导出的csv文件出乱子
Public Function SanitizeForCSV(ByVal s As String) As String
  s = Replace(s, vbCrLf, ";")
  s = Replace(s, Chr(13), ";")
  s = Replace(s, Chr(10), ";")
  s = Replace(s, ",", ";")
  SanitizeForCSV = s
End Function

