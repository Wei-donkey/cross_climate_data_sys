Attribute VB_Name = "OutputExt"




'填充任意时段数据进入HFGrid1表格控件：最大、最小统计结果
Public Sub HFGrid1_Fill_PeriodExt1(Frm As Object, DataFill() As Period_Result)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid1
  For j = 1 To 13  '先按列循环
    For i = 1 To ExtStaNum
        
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
Public Sub HFGrid1_Fill_PeriodExt2(Frm As Object, DataFill() As Period_Result)
  Dim tmpData!, tmpStr$, tmpYear%
  
  With Frm.HFGrid1
  For j = 1 To 12  '先按列循环
    For i = 1 To ExtStaNum
        
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
Public Sub HFGrid2_Fill_PeriodExt1(Frm As Object, DataFill() As Period_Result)
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
Public Sub HFGrid2_Fill_PeriodExt2(Frm As Object, DataFill() As Period_Result)
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
Public Sub HFGrid3_fill_PeriodExt1(Frm As Object)
  Dim DataFill() As Period_Result  ' 需要填充到表格中的二维数据
  Dim tmpData!, iRows%
  Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
  
  If FlagZoneExt = "CITY" Then iRows = ExtCityNum ': StatLvl = 1
  If FlagZoneExt = "CNTY" Then iRows = ExtCountyNum ': StatLvl = 2
  If FlagZoneExt = "PROV" Then iRows = ExtProvNum ': StatLvl = 3
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
Public Sub HFGrid3_fill_PeriodExt2(Frm As Object)
  Dim DataFill() As Period_Result  ' 需要填充到表格中的二维数据
  Dim tmpData!, iRows%
  Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3

  If FlagZoneExt = "CITY" Then iRows = ExtCityNum ': StatLvl = 1
  If FlagZoneExt = "CNTY" Then iRows = ExtCountyNum ': StatLvl = 2
  If FlagZoneExt = "PROV" Then iRows = ExtProvNum ': StatLvl = 3
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
Public Sub HFGrid1_Fill_PeriodExt3(Frm As Object, DataFill() As Period_Result, k%, FlagStat$)
  Dim tmpData!, tmpDate As Date, tmpStr$, tmpYear%
  
  With Frm.HFGrid1
  For j = 1 To 3  '先按列循环
    For i = 1 To ExtStaNum
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
Public Sub HFGrid2_Fill_PeriodExt3(Frm As Object, DataFill() As Period_Result, k%, FlagStat$)
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
Public Sub HFGrid3_fill_PeriodExt3(Frm As Object, m%, FlagStat$)
  Dim DataFill() As Period_Result  ' 需要填充到表格中的二维数据
  Dim tmpData!, iRows%
  Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
  
  If FlagZoneExt = "CITY" Then iRows = ExtCityNum ': StatLvl = 1
  If FlagZoneExt = "CNTY" Then iRows = ExtCountyNum ': StatLvl = 2
  If FlagZoneExt = "PROV" Then iRows = ExtProvNum ': StatLvl = 3
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
