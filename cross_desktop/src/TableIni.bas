Attribute VB_Name = "TableIni"
Public Sub ITimesInitial(Frm As Object)
    
  Dim BDate$, EDate$, idays% '用户自定义的当年查询起始日期
  Dim BDate2$, EDate2$ '用于判断有效资料的起止年份
  Dim BDate3$, EDate3$ '用户自定义对应的某年查询起始日期
  Dim BMMDD$, EMMDD$ '用户自定义的查询起止月日
  
  Dim BYYYY$, BMM$, BDD$, BHH$
  Dim EYYYY$, EMM$, EDD$, EHH$ '用户自定义的查询起止年、月、日
   
  Dim tempBDateStr$(), tempEDateStr$()
  
  Dim tempYear%
  Dim FlagRunBYear, FlagRunEYear As Boolean '判断起止年份是否为闰年
     
  FlagInputErr = False  '默认输入的时间没有错误
    
  If Frm Is FrmMeteoHour Then  '逐时资料查询窗体
    FrmMeteoHour.BTime = Format(FrmMeteoHour.ComboBDate.Text & " " & FrmMeteoHour.ComboBHour.Text & ":00", "yyyy-mm-dd hh:00")
    FrmMeteoHour.ETime = Format(FrmMeteoHour.ComboEDate.Text & " " & FrmMeteoHour.ComboEHour.Text & ":00", "yyyy-mm-dd hh:00")
    
    '************************************ 2023-10-10添加对输入日期的精准判断 **********************************
    tempBDateStr = Split(FrmMeteoHour.ComboBDate.Text, "-")
    BHH = Format(FrmMeteoHour.ComboBHour.Text, "00")
    tempEDateStr = Split(FrmMeteoHour.ComboEDate.Text, "-")
    EHH = Format(FrmMeteoHour.ComboEHour.Text, "00")
  
  ElseIf Frm Is FrmMeteoDay Then  '逐日资料查询窗体
    
    If FrmMeteoDay.DataType = "Data" Then
      FrmMeteoDay.BDate = Format(FrmMeteoDay.ComboBDate, "yyyy-mm-dd")
      FrmMeteoDay.EDate = Format(FrmMeteoDay.ComboEDate, "yyyy-mm-dd")
    Else '平均态和极端态没有2月29日，则设置为平年
      FrmMeteoDay.BDate = "2011-" & Format(FrmMeteoDay.ComboBDate, "mm-dd")
      FrmMeteoDay.EDate = "2011-" & Format(FrmMeteoDay.ComboEDate, "mm-dd")
    End If
 
    '************************************ 2023-10-10添加对输入日期的精准判断 **********************************
    tempBDateStr = Split(FrmMeteoDay.BDate, "-")
    tempEDateStr = Split(FrmMeteoDay.EDate, "-")
  
  ElseIf Frm Is FrmMeteoTen Then  '逐旬资料查询窗体
    FrmMeteoTen.BYYMM = Format(FrmMeteoTen.ComboBYear.Text & "-" & FrmMeteoTen.ComboBMM.Text, "yyyy-mm")
    FrmMeteoTen.EYYMM = Format(FrmMeteoTen.ComboEYear.Text & "-" & FrmMeteoTen.ComboEMM.Text, "yyyy-mm")
    FrmMeteoTen.BMM = Right(FrmMeteoTen.BYYMM, 2)
    FrmMeteoTen.EMM = Right(FrmMeteoTen.EYYMM, 2)

    '************************************ 2023-10-10添加对输入日期的精准判断 **********************************
    tempBDateStr = Split(FrmMeteoTen.BYYMM, "-")
    tempEDateStr = Split(FrmMeteoTen.EYYMM, "-")
  
  ElseIf Frm Is FrmMeteoMon Then  '逐月资料查询窗体
    FrmMeteoMon.BYYMM = Format(FrmMeteoMon.ComboBYear.Text & "-" & FrmMeteoMon.ComboBMM.Text, "yyyy-mm")
    FrmMeteoMon.EYYMM = Format(FrmMeteoMon.ComboEYear.Text & "-" & FrmMeteoMon.ComboEMM.Text, "yyyy-mm")
    FrmMeteoMon.BMM = Right(FrmMeteoMon.BYYMM, 2)
    FrmMeteoMon.EMM = Right(FrmMeteoMon.EYYMM, 2)

    '************************************ 2023-10-10添加对输入日期的精准判断 **********************************
    tempBDateStr = Split(FrmMeteoMon.BYYMM, "-")
    tempEDateStr = Split(FrmMeteoMon.EYYMM, "-")
  
  ElseIf Frm Is FrmMeteoQtr Then  '逐季资料查询窗体
    FrmMeteoQtr.BYYYY = FrmMeteoQtr.ComboBYear.Text
    FrmMeteoQtr.EYYYY = FrmMeteoQtr.ComboEYear.Text
    
    '************************************ 2023-10-10添加对输入日期的精准判断 **********************************
    tempBDateStr = Split(FrmMeteoQtr.BYYYY, "-")
    tempEDateStr = Split(FrmMeteoQtr.EYYYY, "-")
    
  ElseIf Frm Is FrmMeteoYer Then  '逐年资料查询窗体
    FrmMeteoYer.BYYYY = FrmMeteoYer.ComboBYear.Text
    FrmMeteoYer.EYYYY = FrmMeteoYer.ComboEYear.Text
    
    '************************************ 2023-10-10添加对输入日期的精准判断 **********************************
    tempBDateStr = Split(FrmMeteoYer.BYYYY, "-")
    tempEDateStr = Split(FrmMeteoYer.EYYYY, "-")
  
  ElseIf Frm Is FrmMeteoPeriodYer Then  '逐年任意时段资料查询窗体
    FrmMeteoPeriodYer.BYear = FrmMeteoPeriodYer.ComboBYear.Text
    FrmMeteoPeriodYer.EYear = FrmMeteoPeriodYer.ComboEYear.Text
    
    '************************************ 2023-10-10添加对输入日期的精准判断 **********************************
    tempBDateStr = Split(FrmMeteoPeriodYer.BYear, "-")
    tempEDateStr = Split(FrmMeteoPeriodYer.EYear, "-")
  
  End If
  
    
  '************************************ 2023-10-10添加对输入日期的精准判断 **********************************
  BYYYY = tempBDateStr(0)
  If Not (Frm Is FrmMeteoQtr Or Frm Is FrmMeteoYer Or Frm Is FrmMeteoPeriodYer) Then
    BMM = tempBDateStr(1)
  End If
  If Frm Is FrmMeteoHour Or Frm Is FrmMeteoDay Then
    BDD = tempBDateStr(2)
  End If
      
  EYYYY = tempEDateStr(0)
  If Not (Frm Is FrmMeteoQtr Or Frm Is FrmMeteoYer Or Frm Is FrmMeteoPeriodYer) Then
    EMM = tempEDateStr(1)
  End If
  If Frm Is FrmMeteoHour Or Frm Is FrmMeteoDay Then
    EDD = tempEDateStr(2)
  End If
  
'  BDate = Format(BYYYY, "0000") & "-" & Format(BMM, "00") & "-" & Format(BDD, "00")
'  EDate = Format(EYYYY, "0000") & "-" & Format(EMM, "00") & "-" & Format(EDD, "00")
    
'  BMMDD = Right(BDate, 5) '开始月日
'  EMMDD = Right(EDate, 5) '结束月日
  
  '判断起止年份是否为闰年
  FlagRunBYear = False: tempYear = CInt(BYYYY)
  If tempYear Mod 4 = 0 And tempYear Mod 100 <> 0 Or tempYear Mod 400 = 0 Then FlagRunBYear = True
  FlagRunEYear = False: tempYear = CInt(EYYYY)
  If tempYear Mod 4 = 0 And tempYear Mod 100 <> 0 Or tempYear Mod 400 = 0 Then FlagRunEYear = True
            
  If Not (Frm Is FrmMeteoQtr Or Frm Is FrmMeteoYer Or Frm Is FrmMeteoPeriodYer) Then
    If CInt(BMM) > 12 Or CInt(BMM) < 1 Then MsgBox "Invalid start month", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
    If CInt(EMM) > 12 Or CInt(EMM) < 1 Then MsgBox "Invalid end month", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
  End If
  
  If Frm Is FrmMeteoHour Or Frm Is FrmMeteoDay Then
    Select Case CInt(BMM)
      Case 1, 3, 5, 7, 8, 10, 12
        If CInt(BDD) > 31 Or CInt(BDD) < 0 Then MsgBox "Invalid start date", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
      Case 4, 6, 9, 11
        If CInt(BDD) > 30 Or CInt(BDD) < 0 Then MsgBox "Invalid start date", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
      Case 2
        If FlagRunBYear = True Then
          If CInt(BDD) > 29 Or CInt(BDD) < 0 Then MsgBox "Invalid start date", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
        ElseIf FlagRunBYear = False Then
          If CInt(BDD) > 28 Or CInt(BDD) < 0 Then MsgBox "Invalid start date", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
        End If
    End Select
        
    Select Case CInt(EMM)
      Case 1, 3, 5, 7, 8, 10, 12
        If CInt(EDD) > 31 Or CInt(EDD) < 0 Then MsgBox "Invalid end date", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
      Case 4, 6, 9, 11
        If CInt(EDD) > 30 Or CInt(EDD) < 0 Then MsgBox "Invalid end date", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
      Case 2
        If FlagRunEYear = True Then
          If CInt(EDD) > 29 Or CInt(EDD) < 0 Then MsgBox "Invalid end date", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
        ElseIf FlagRunEYear = False Then
          If CInt(EDD) > 28 Or CInt(EDD) < 0 Then MsgBox "Invalid end date", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
        End If
    End Select
  End If
        
  If Frm Is FrmMeteoHour Then  '逐时数据模块判断时次
    If CInt(BHH) > 23 Or CInt(BHH) < 0 Then MsgBox "Invalid start hour", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
    If CInt(EHH) > 23 Or CInt(EHH) < 0 Then MsgBox "Invalid end hour", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
  End If
  '************************************ 2023-10-10添加对输入日期的精准判断 **********************************
    
     
  If Frm Is FrmMeteoHour Then  '逐时资料查询窗体
    If CDate(FrmMeteoHour.ETime) < CDate(FrmMeteoHour.BTime) Then MsgBox "Invalid time", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
    If FrmMeteoHour.OptionHour(0).Value = True Then
      FrmMeteoHour.DataType = "Multi"
      FrmMeteoHour.iHors = DateDiff("h", CDate(FrmMeteoHour.BTime), CDate(FrmMeteoHour.ETime)) + 1
    ElseIf FrmMeteoHour.OptionHour(1).Value = True Then
      FrmMeteoHour.DataType = "Single"
      FrmMeteoHour.iHors = DateDiff("d", CDate(FrmMeteoHour.BTime), CDate(FrmMeteoHour.ETime)) + 1
    End If
  
  ElseIf Frm Is FrmMeteoDay Then  '逐日资料查询窗体
    If CDate(FrmMeteoDay.EDate) < CDate(FrmMeteoDay.BDate) Then MsgBox "Invalid date", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
    FrmMeteoDay.idays = DateDiff("d", CDate(FrmMeteoDay.BDate), CDate(FrmMeteoDay.EDate)) + 1
  
  ElseIf Frm Is FrmMeteoTen Then  '逐旬资料查询窗体
    If CDate(FrmMeteoTen.EYYMM) < CDate(FrmMeteoTen.BYYMM) Then MsgBox "Invalid time", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
    If FrmMeteoTen.DataType = "Data" Then
      FrmMeteoTen.iMons = DateDiff("m", CDate(FrmMeteoTen.BYYMM), CDate(FrmMeteoTen.EYYMM)) + 1
    Else
      If FrmMeteoTen.EMM >= FrmMeteoTen.BMM Then FrmMeteoTen.iMons = CInt(FrmMeteoTen.EMM) - CInt(FrmMeteoTen.BMM) + 1
      If FrmMeteoTen.EMM < FrmMeteoTen.BMM Then FrmMeteoTen.iMons = CInt(FrmMeteoTen.EMM) + (12 - CInt(FrmMeteoTen.BMM) + 1)
    End If
    FrmMeteoTen.iTens = 3 * FrmMeteoTen.iMons

  ElseIf Frm Is FrmMeteoMon Then  '逐月资料查询窗体
    If CDate(FrmMeteoMon.EYYMM) < CDate(FrmMeteoMon.BYYMM) Then MsgBox "Invalid time", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
    If FrmMeteoMon.DataType = "Data" Then
      FrmMeteoMon.iMons = DateDiff("m", CDate(FrmMeteoMon.BYYMM), CDate(FrmMeteoMon.EYYMM)) + 1
    Else
      If FrmMeteoMon.EMM >= FrmMeteoMon.BMM Then FrmMeteoMon.iMons = CInt(FrmMeteoMon.EMM) - CInt(FrmMeteoMon.BMM) + 1
      If FrmMeteoMon.EMM < FrmMeteoMon.BMM Then FrmMeteoMon.iMons = CInt(FrmMeteoMon.EMM) + (12 - CInt(FrmMeteoMon.BMM) + 1)
    End If
    
  ElseIf Frm Is FrmMeteoQtr Then  '逐季资料查询窗体
    If CDate(FrmMeteoQtr.EYYYY) < CDate(FrmMeteoQtr.BYYYY) Then MsgBox "Invalid year", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
    If FrmMeteoQtr.DataType = "Data" Then
      FrmMeteoQtr.iYers = CInt(FrmMeteoQtr.EYYYY) - CInt(FrmMeteoQtr.BYYYY) + 1
    Else
      FrmMeteoQtr.iYers = 1
    End If
    FrmMeteoQtr.iQtrs = 4 * FrmMeteoQtr.iYers
    
  ElseIf Frm Is FrmMeteoYer Then  '逐年资料查询窗体
    If CDate(FrmMeteoYer.EYYYY) < CDate(FrmMeteoYer.BYYYY) Then MsgBox "Invalid year", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
    If FrmMeteoYer.DataType = "Data" Then
      FrmMeteoYer.iYers = CInt(FrmMeteoYer.EYYYY) - CInt(FrmMeteoYer.BYYYY) + 1
    Else
      FrmMeteoYer.iYers = 1
    End If
  
  ElseIf Frm Is FrmMeteoPeriodYer Then  '逐年任意时段资料查询窗体
    If CDate(FrmMeteoPeriodYer.EYear) < CDate(FrmMeteoPeriodYer.BYear) Then MsgBox "Invalid year", vbInformation, "Notice": Screen.MousePointer = 1: FlagInputErr = True: Exit Sub
    FrmMeteoPeriodYer.iYers = CInt(FrmMeteoPeriodYer.EYear) - CInt(FrmMeteoPeriodYer.BYear) + 1
  
  End If
  
End Sub






'************************ 冷空气统计数据结果表格的初始化(所有过程) *****************************
Public Sub TableIni_MeteoColdAir(Frm As Object)
  Dim iItems_stat% '统计的项目数量
  Dim iItems_rank%
  Dim iItems_zone%
  
  iItems_stat = 9 '相关的项目数量：等级、开始日期、结束日期、日数、24小时降温、48小时降温、过程降温、平均气温、最低气温
  iItems_zone = 3 '区域信息数据：镇乡、区县、地市
  
  With Frm
    If FlagZone = "STA" Then '输出站点结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid4.Visible = True
      .HFGrid5.Visible = True
      .HFGrid3(1).Visible = False
      .HFGrid3(2).Visible = False
      .HFGrid3(3).Visible = False
    Else '输出区域统计结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid4.Visible = False
      .HFGrid5.Visible = False
      .HFGrid3(1).Visible = True
      .HFGrid3(2).Visible = True
      .HFGrid3(3).Visible = True
    End If
  End With

  '**************输出站点结果：设置HFGrid1和HFGrid2两个表格**************
  If FlagZone = "STA" Then
  
    '*********************初始化上方表格*******************************
      With Frm.HFGrid4
      .Redraw = False
      .Clear: .Refresh
''''      .GridLines = flexGridFlat: .GridLinesFixed = flexGridFlat  '如此设置才可以给GridColor赋值
''''      .GridColor = &HC0C0C0: .GridColorFixed = &HC0C0C0 '给Grid线条赋值浅灰色
      
      .Cols = 3 + iItems_stat + iItems_zone  '排名相关的7列、镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 600: .ColWidth(1) = 1000: .ColWidth(2) = 2500
      .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Station Code": .TextMatrix(0, 2) = "Station Name"
      
      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 2000
      Next i
      .TextMatrix(0, 3) = "Level": .TextMatrix(0, 4) = "Cold Onset Day": .TextMatrix(0, 5) = "Day Before Warm-Up": .TextMatrix(0, 6) = "Duration"
      .TextMatrix(0, 7) = "Max 24h Temp Drop": .TextMatrix(0, 8) = "Max 48h Temp Drop": .TextMatrix(0, 9) = "Temp Drop": .TextMatrix(0, 10) = "Avg Temp": .TextMatrix(0, 11) = "Min Temp"
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat) = 1000
        If i = 1 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "Township"
        If i = 2 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "County"
        If i = 3 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "City"
      Next i
      
      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
        .TextMatrix(0, 2 + iItems_stat + 4) = "Longitude": .TextMatrix(0, 2 + iItems_stat + 5) = "Latitude"
      End If
      
      '******************填充站点信息信息********************
'        If UBound(CA_All) = 0 Then
        If UBound(CA_ALL) = 0 Then .Rows = 2

        '******上方单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
      
      .Refresh
      .Redraw = True
      End With
      

    '*********************初始化下方表格*******************************
      With Frm.HFGrid5
      .Redraw = False
      .Clear: .Refresh
''''      .GridLines = flexGridFlat: .GridLinesFixed = flexGridFlat '如此设置才可以给GridColor赋值
''''      .GridColor = &HC0C0C0: .GridColorFixed = &HC0C0C0 '给Grid线条赋值浅灰色

      .Cols = 3 + iItems_stat + iItems_zone '排名相关的7列、镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 600: .ColWidth(1) = 1000: .ColWidth(2) = 2500
      
      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 2000
      Next i
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat) = 1000
      Next i
      
      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
      End If

      .Rows = 3
      .TextMatrix(0, 2) = "Average": .TextMatrix(1, 2) = "Max": .TextMatrix(2, 2) = "Min"
      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Col = 1 Then .CellBackColor = &H80000005
          If .Col = 2 Then .CellBackColor = &H80000005
        Next j
      Next i
      
      .Refresh
      .Redraw = True
    End With
    If FrmMeteoColdAir.HFGrid1.Rows > 2 Then FrmMeteoColdAir.HFGrid1.Row = 2

  Else '输出区域统计结果：设置HFGrid3、HFGrid4、HFGrid5三个表格
      Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
      
      '*********************循环初始化三个表格*******************************
      For k = 1 To 3
        With Frm.HFGrid3(k)
        .Redraw = False
        .Clear: .Refresh
        StatLvl = 2
          
        .Cols = 1 + StatLvl + iItems_stat '排名相关的5列
        .FixedCols = 1 + StatLvl
        .ColWidth(0) = 1000:
        If k = 1 Then .TextMatrix(0, 0) = "Regional Avg"
        If k = 2 Then .TextMatrix(0, 0) = "Regional Max"
        If k = 3 Then .TextMatrix(0, 0) = "Regional Min"
        For i = 1 To StatLvl
          .ColWidth(i) = 1500
          If i = 1 Then .TextMatrix(0, i) = "City"
          If i = 2 Then .TextMatrix(0, i) = "County"
          If i = 3 Then .TextMatrix(0, i) = "Township"
        Next i
        
        For i = StatLvl To StatLvl + iItems_stat
          .ColWidth(i) = 2000
        Next i
      

        .TextMatrix(0, StatLvl + 1) = "Level": .TextMatrix(0, StatLvl + 2) = "Cold Onset Day": .TextMatrix(0, StatLvl + 3) = "Day Before Warm-Up": .TextMatrix(0, StatLvl + 4) = "Duration": .TextMatrix(0, StatLvl + 5) = "Max 24h Temp Drop"
        .TextMatrix(0, StatLvl + 6) = "Max 48h Temp Drop": .TextMatrix(0, StatLvl + 7) = "Temp Drop": .TextMatrix(0, StatLvl + 8) = "Avg Temp": .TextMatrix(0, StatLvl + 9) = "Min Temp"
      
        
         '******************填充区域信息********************
        If FlagZone = "CITY" Then .Rows = CityNum + 1
        If FlagZone = "CNTY" Then .Rows = CountyNum + 1
        If FlagZone = "TOWN" Then .Rows = TownNum + 1
        For i = 1 To .Rows - 1
          .TextMatrix(i, 0) = i
            
          If FlagZone = "CITY" Then
            .TextMatrix(i, 1) = CityInfo2(i).Code
            .TextMatrix(i, 2) = CityInfo2(i).City
          ElseIf FlagZone = "CNTY" Then
            .TextMatrix(i, 1) = CountyInfo2(i).Code
            .TextMatrix(i, 2) = CountyInfo2(i).City & CountyInfo2(i).County
          ElseIf FlagZone = "TOWN" Then
            .TextMatrix(i, 1) = TownInfo2(i).Code
            .TextMatrix(i, 2) = TownInfo2(i).City & TownInfo2(i).County & TownInfo2(i).Town
          End If
        Next i
    
    
    
        '******单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
    
        .Refresh
        .Redraw = True
        End With
      
      Next k
  End If

          
End Sub





'************************ 冷空气统计数据结果表格的初始化(最强过程) *****************************
Public Sub TableIni_MeteoColdAir2(Frm As Object)
  Dim iItems_stat% '统计的项目数量
  Dim iItems_rank%
  Dim iItems_zone%
  
  iItems_stat = 9 '相关的项目数量：等级、开始日期、结束日期、日数、24小时降温、48小时降温、过程降温、平均气温、最低气温
  iItems_zone = 3 '区域信息数据：镇乡、区县、地市
  
  With Frm
    If FlagZone = "STA" Then '输出站点结果
      .HFGrid1.Visible = True
      .HFGrid2.Visible = True
      .HFGrid4.Visible = False
      .HFGrid5.Visible = False
      .HFGrid3(1).Visible = False
      .HFGrid3(2).Visible = False
      .HFGrid3(3).Visible = False
    Else '输出区域统计结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid4.Visible = False
      .HFGrid5.Visible = False
      .HFGrid3(1).Visible = True
      .HFGrid3(2).Visible = True
      .HFGrid3(3).Visible = True
    End If
  End With

  '**************输出站点结果：设置HFGrid1和HFGrid2两个表格**************
  If FlagZone = "STA" Then
  
    '*********************初始化上方表格*******************************
      With Frm.HFGrid1
      .Redraw = False
      .Clear: .Refresh
      
      .Cols = 3 + iItems_stat + iItems_zone  '排名相关的7列、镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 600: .ColWidth(1) = 1000: .ColWidth(2) = 2500
      .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Station Code": .TextMatrix(0, 2) = "Station Name"
      
      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 2000
      Next i
      .TextMatrix(0, 3) = "Level": .TextMatrix(0, 4) = "Cold Onset Day": .TextMatrix(0, 5) = "Day Before Warm-Up": .TextMatrix(0, 6) = "Duration"
      .TextMatrix(0, 7) = "Max 24h Temp Drop": .TextMatrix(0, 8) = "Max 48h Temp Drop": .TextMatrix(0, 9) = "Temp Drop": .TextMatrix(0, 10) = "Avg Temp": .TextMatrix(0, 11) = "Min Temp"
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat) = 1000
        If i = 1 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "Township"
        If i = 2 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "County"
        If i = 3 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "City"
      Next i
      
      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
        .TextMatrix(0, 2 + iItems_stat + 4) = "Longitude": .TextMatrix(0, 2 + iItems_stat + 5) = "Latitude"
      End If
      
      '******************填充站点信息信息********************
        .Rows = StaNum + 1
        For i = 1 To StaNum
          .TextMatrix(i, 0) = i
          .TextMatrix(i, 1) = StaInfo(i).stacode
          .TextMatrix(i, 2) = StaInfo(i).staname
          
          .TextMatrix(i, 1 + 2 + iItems_stat) = StaInfo(i).Town
          .TextMatrix(i, 2 + 2 + iItems_stat) = StaInfo(i).County
          .TextMatrix(i, 3 + 2 + iItems_stat) = StaInfo(i).City
                      
          If ShowLonLat = True Then
            .TextMatrix(i, 2 + iItems_stat + 4) = StaInfo(i).Longitude
            .TextMatrix(i, 2 + iItems_stat + 5) = StaInfo(i).Latitude
          End If
        Next i

        '******上方单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
      
      .Refresh
      .Redraw = True
      End With
      

    '*********************初始化下方表格*******************************
      With Frm.HFGrid2
      .Redraw = False
      .Clear: .Refresh
      .Cols = 3 + iItems_stat + iItems_zone '排名相关的7列、镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 600: .ColWidth(1) = 1000: .ColWidth(2) = 2500
      
      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 2000
      Next i
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat) = 1000
      Next i
      
      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
      End If

      .Rows = 3
      .TextMatrix(0, 2) = "Average": .TextMatrix(1, 2) = "Max": .TextMatrix(2, 2) = "Min"
      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Col = 1 Then .CellBackColor = &H80000005
          If .Col = 2 Then .CellBackColor = &H80000005
        Next j
      Next i
      
      .Refresh
      .Redraw = True
    End With
    If FrmMeteoColdAir.HFGrid1.Rows > 2 Then FrmMeteoColdAir.HFGrid1.Row = 2

  Else '输出区域统计结果：设置HFGrid3、HFGrid4、HFGrid5三个表格
      Dim StatLvl% '统计区域级别：城市-1，县区-2，镇街-3
      
      '*********************循环初始化三个表格*******************************
      For k = 1 To 3
        With Frm.HFGrid3(k)
        .Redraw = False
        .Clear: .Refresh
        StatLvl = 2
          
          

        .Cols = 1 + StatLvl + iItems_stat '排名相关的5列
        .FixedCols = 3
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Region Code": .TextMatrix(0, 2) = "Region Name"
        
        If k = 1 Then .TextMatrix(0, 0) = "Regional Avg"
        If k = 2 Then .TextMatrix(0, 0) = "Regional Max"
        If k = 3 Then .TextMatrix(0, 0) = "Regional Min"
          
        
        For i = StatLvl To StatLvl + iItems_stat
          .ColWidth(i) = 2000
        Next i
      

      .TextMatrix(0, StatLvl + 1) = "Level": .TextMatrix(0, StatLvl + 2) = "Cold Onset Day": .TextMatrix(0, StatLvl + 3) = "Day Before Warm-Up": .TextMatrix(0, StatLvl + 4) = "Duration": .TextMatrix(0, StatLvl + 5) = "Max 24h Temp Drop"
      .TextMatrix(0, StatLvl + 6) = "Max 48h Temp Drop": .TextMatrix(0, StatLvl + 7) = "Temp Drop": .TextMatrix(0, StatLvl + 8) = "Avg Temp": .TextMatrix(0, StatLvl + 9) = "Min Temp"
      

        
        
         '******************填充区域信息********************
        If FlagZone = "CITY" Then .Rows = CityNum + 1
        If FlagZone = "CNTY" Then .Rows = CountyNum + 1
        If FlagZone = "TOWN" Then .Rows = TownNum + 1
        For i = 1 To .Rows - 1
          .TextMatrix(i, 0) = i
            
          If FlagZone = "CITY" Then
            .TextMatrix(i, 1) = CityInfo2(i).Code
            .TextMatrix(i, 2) = CityInfo2(i).City
          ElseIf FlagZone = "CNTY" Then
            .TextMatrix(i, 1) = CountyInfo2(i).Code
            .TextMatrix(i, 2) = CountyInfo2(i).City & CountyInfo2(i).County
          ElseIf FlagZone = "TOWN" Then
            .TextMatrix(i, 1) = TownInfo2(i).Code
            .TextMatrix(i, 2) = TownInfo2(i).City & TownInfo2(i).County & TownInfo2(i).Town
          End If
        Next i
      
        '******单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
      
        .Refresh
        .Redraw = True
        End With
      
      Next k
  End If

          
End Sub



'************************ 固定时间间隔(时、日、旬、月、年)统计数据结果表格的初始化 *****************************
Public Sub TableIni_Meteo(Frm As Object, BTime$, iTimes%)  'iTimes为小时数、日数、旬数、月数、年份数等
  With Frm
    If FlagZone = "STA" Then '输出站点结果
      .HFGrid1.Visible = True
      .HFGrid2.Visible = True
      .HFGrid3(1).Visible = False
      .HFGrid3(2).Visible = False
      .HFGrid3(3).Visible = False
    Else '输出区域统计结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid3(1).Visible = True
      .HFGrid3(2).Visible = True
      .HFGrid3(3).Visible = True
    End If
  End With
  
  '**************输出站点结果：设置HFGrid1和HFGrid2两个表格**************
  If FlagZone = "STA" Then
  
    '*********************初始化上方表格*******************************
      With Frm.HFGrid1
      .Redraw = False
      .Clear: .Refresh

        .Cols = 3 + iTimes + 6   '统计区间均值、最大值、最小值，镇、县、市（国家站同样显示）
       
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Station Code": .TextMatrix(0, 2) = "Station Name"
      
      '******************将日期及相关统计量写入表头行********************
      For i = 1 To iTimes
        If Frm Is FrmMeteoHour Then
          If Frm.DataType = "Multi" Then .TextMatrix(0, 2 + i) = Format(DateAdd("h", i - 1, BTime), "yyyy-mm-dd hh:00")
          If Frm.DataType = "Single" Then .TextMatrix(0, 2 + i) = Format(DateAdd("d", i - 1, BTime), "yyyy-mm-dd hh:00")
        End If
        
        If Frm Is FrmMeteoDay Then .TextMatrix(0, 2 + i) = Format(DateAdd("d", i - 1, BTime), "yyyy-mm-dd")
        
       
        If Frm Is FrmMeteoMon Then
          If Frm.DataType = "Data" Then .TextMatrix(0, 2 + i) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm")
          If Frm.DataType <> "Data" Then .TextMatrix(0, 2 + i) = Format(DateAdd("m", i - 1, BTime), "mmm")
        End If
        
        If Frm Is FrmMeteoYer Then
          If Frm.DataType = "Data" Then .TextMatrix(0, 2 + i) = CStr(CInt(BTime) + i - 1) & ""
          If Frm.DataType = "Norm" Then .TextMatrix(0, 2 + i) = "Normals"
          If Frm.DataType = "Xtrm" Then .TextMatrix(0, 2 + i) = "Historical"
        End If
'        If Frm Is FrmMeteoYer Then .TextMatrix(0, 2 + i) = CStr(CInt(BTime) + i - 1) & ""
        If Frm Is FrmMeteoPeriodYer Then .TextMatrix(0, 2 + i) = CStr(CInt(BTime) + i - 1) & ""
        
        .ColWidth(i + 2) = 1500
      Next i
      
'      If Frm Is FrmMeteoDay Then
      If (Frm Is FrmMeteoDay) Or (Frm Is FrmMeteoHour) Then
        If Frm.OptionStat(1).Value = True Then .TextMatrix(0, 3 + iTimes) = "Average"
        If Frm.OptionStat(4).Value = True Then .TextMatrix(0, 3 + iTimes) = "Cumulative"
      Else
        .TextMatrix(0, 3 + iTimes) = "Average"
      End If
      
      .TextMatrix(0, 4 + iTimes) = "Max": .TextMatrix(0, 5 + iTimes) = "Min"
      .ColWidth(3 + iTimes) = 1500: .ColWidth(4 + iTimes) = 1500: .ColWidth(5 + iTimes) = 1500
      
      .TextMatrix(0, 6 + iTimes) = "Township": .TextMatrix(0, 7 + iTimes) = "County": .TextMatrix(0, 8 + iTimes) = "City"
      .ColWidth(6 + iTimes) = 1500: .ColWidth(7 + iTimes) = 1500: .ColWidth(8 + iTimes) = 1500
      
      If ShowLonLat = True Then
        .TextMatrix(0, 9 + iTimes) = "Longitude": .TextMatrix(0, 10 + iTimes) = "Latitude"
        .ColWidth(9 + iTimes) = 1500: .ColWidth(10 + iTimes) = 1500
      End If
      
      '******************填充站点信息信息********************
      If Frm Is FrmMeteoHour Then
          
        .Rows = StaNum + 1
        For i = 1 To StaNum
          .TextMatrix(i, 0) = i
          .TextMatrix(i, 1) = StaInfo(i).stacode
          .TextMatrix(i, 2) = StaInfo(i).staname
            
          '将平均、最大、最小值列涂灰色
          .Row = i: .Col = iTimes + 3
          If Not GuestMode Then .CellBackColor = &HE0E0E0
          .Row = i: .Col = iTimes + 4
          If Not GuestMode Then .CellBackColor = &HE0E0E0
          .Row = i: .Col = iTimes + 5
          If Not GuestMode Then .CellBackColor = &HE0E0E0
            
          .TextMatrix(i, iTimes + 6) = StaInfo(i).Town
          .TextMatrix(i, iTimes + 7) = StaInfo(i).County
          .TextMatrix(i, iTimes + 8) = StaInfo(i).City
                        
          If ShowLonLat = True Then
            .TextMatrix(i, iTimes + 9) = StaInfo(i).Longitude
            .TextMatrix(i, iTimes + 10) = StaInfo(i).Latitude
          End If
        Next i
      
      Else
        
        If Frm.XtrmYear = False Then '不显示累年极值出现的年份
          .Rows = StaNum + 1
          For i = 1 To StaNum
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = StaInfo(i).stacode
            .TextMatrix(i, 2) = StaInfo(i).staname
            
            '将平均、最大、最小值列涂灰色
            .Row = i: .Col = iTimes + 3
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            .Row = i: .Col = iTimes + 4
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            .Row = i: .Col = iTimes + 5
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            
            .TextMatrix(i, iTimes + 6) = StaInfo(i).Town
            .TextMatrix(i, iTimes + 7) = StaInfo(i).County
            .TextMatrix(i, iTimes + 8) = StaInfo(i).City
                        
            If ShowLonLat = True Then
              .TextMatrix(i, iTimes + 9) = StaInfo(i).Longitude
              .TextMatrix(i, iTimes + 10) = StaInfo(i).Latitude
            End If
          Next i
          
        ElseIf Frm.XtrmYear = True Then '显示累年极值出现的年份
          .Rows = StaNum * 2 + 1
          For i = 1 To StaNum
            .TextMatrix(i * 2 - 1, 0) = i
            .TextMatrix(i * 2 - 1, 1) = StaInfo(i).stacode
            .TextMatrix(i * 2 - 1, 2) = StaInfo(i).staname
            .TextMatrix(i * 2, 0) = i
            .TextMatrix(i * 2, 1) = StaInfo(i).stacode
            .TextMatrix(i * 2, 2) = StaInfo(i).staname
            
            '将年份行涂灰色
              For j = 1 To iTimes
                .Row = i * 2: .Col = 2 + j: .CellBackColor = &HE0E0E0
              Next j
            
            '将平均、最大、最小值列涂灰色
            .Row = i * 2 - 1: .Col = iTimes + 3
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            .Row = i * 2 - 1: .Col = iTimes + 4
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            .Row = i * 2 - 1: .Col = iTimes + 5
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            
            .TextMatrix(i * 2 - 1, iTimes + 6) = StaInfo(i).Town
            .TextMatrix(i * 2 - 1, iTimes + 7) = StaInfo(i).County
            .TextMatrix(i * 2 - 1, iTimes + 8) = StaInfo(i).City
            
            If ShowLonLat = True Then
              .TextMatrix(i * 2 - 1, iTimes + 9) = StaInfo(i).Longitude
              .TextMatrix(i * 2 - 1, iTimes + 10) = StaInfo(i).Latitude
            End If
          Next i
        
        End If
      
      End If
      
      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Row <> 0 Then
'            If .Col = 0 Then .CellBackColor = &H80000005
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          End If
        Next j
      Next i
      
      .Refresh
      .Redraw = True
      End With
          
    '*********************初始化下方表格*******************************
      With Frm.HFGrid2
      .Redraw = False
      .Clear: .Refresh
      
      .Cols = 3 + iTimes + 6   '统计区间均值、最大值、最小值，镇、县、市（国家站同样显示）
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      
      '******************将日期及相关统计量写入表头行********************
      For i = 1 To iTimes
        If Frm Is FrmMeteoHour Then
          If Frm.DataType = "Multi" Then .TextMatrix(0, 2 + i) = Format(DateAdd("h", i - 1, BTime), "yyyy-mm-dd hh:00")
          If Frm.DataType = "Single" Then .TextMatrix(0, 2 + i) = Format(DateAdd("d", i - 1, BTime), "yyyy-mm-dd hh:00")
        End If
        If Frm Is FrmMeteoDay Then .TextMatrix(0, 2 + i) = Format(DateAdd("d", i - 1, BTime), "yyyy-mm-dd")
        If Frm Is FrmMeteoMon Then
          If Frm.DataType = "Data" Then .TextMatrix(0, 2 + i) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm")
          If Frm.DataType <> "Data" Then .TextMatrix(0, 2 + i) = Format(DateAdd("m", i - 1, BTime), "mmm")
        End If
        If Frm Is FrmMeteoYer Then
          If Frm.DataType = "Data" Then .TextMatrix(0, 2 + i) = CStr(CInt(BTime) + i - 1) & ""
          If Frm.DataType = "Norm" Then .TextMatrix(0, 2 + i) = "Normals"
          If Frm.DataType = "Xtrm" Then .TextMatrix(0, 2 + i) = "Historical"
        End If
        If Frm Is FrmMeteoPeriodYer Then .TextMatrix(0, 2 + i) = CStr(CInt(BTime) + i - 1) & ""
        .ColWidth(i + 2) = 1500
      Next i
      
'      .TextMatrix(0, 3 + iTimes) = "Average": .TextMatrix(0, 4 + iTimes) = "Max": .TextMatrix(0, 5 + iTimes) = "Min"
      .ColWidth(3 + iTimes) = 1500: .ColWidth(4 + iTimes) = 1500: .ColWidth(5 + iTimes) = 1500
      .ColWidth(6 + iTimes) = 1500: .ColWidth(7 + iTimes) = 1500: .ColWidth(8 + iTimes) = 1500
      If ShowLonLat = True Then
        .ColWidth(9 + iTimes) = 1500: .ColWidth(10 + iTimes) = 1500
      End If
      
      '******************填充站点信息或地市信息********************
      .Rows = 3 + 1: .FixedRows = 1
      .TextMatrix(1, 2) = "Average"
      .TextMatrix(2, 2) = "Max"
      .TextMatrix(3, 2) = "Min"

      '******下方单元格左对齐、着色********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Row <> 0 Then
'            If .Col = 0 Then .CellBackColor = &H80000005
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          End If
        Next j
      Next i
      
      .Refresh
      .Redraw = True
      End With
      
      If Frm.HFGrid1.Rows > 2 Then Frm.HFGrid1.Row = 2
      
  Else '输出区域统计结果：设置HFGrid3、HFGrid4、HFGrid5三个表格
      
      '*********************循环初始化三个表格*******************************
      For k = 1 To 3
        With Frm.HFGrid3(k)
        .Redraw = False
        .Clear: .Refresh
          
        .Cols = 3 + iTimes + 3   '统计区间均值、最大值、最小值
      
        .FixedCols = 3
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Region Code": .TextMatrix(0, 2) = "Region Name"
        
        If k = 1 Then .TextMatrix(0, 0) = "Regional Avg"
        If k = 2 Then .TextMatrix(0, 0) = "Regional Max"
        If k = 3 Then .TextMatrix(0, 0) = "Regional Min"
                    
       '******************将日期及相关统计量写入表头行********************
        For i = 1 To iTimes
          If Frm Is FrmMeteoHour Then
            If Frm.DataType = "Multi" Then .TextMatrix(0, 2 + i) = Format(DateAdd("h", i - 1, BTime), "yyyy-mm-dd hh:00")
            If Frm.DataType = "Single" Then .TextMatrix(0, 2 + i) = Format(DateAdd("d", i - 1, BTime), "yyyy-mm-dd hh:00")
          End If
          If Frm Is FrmMeteoDay Then .TextMatrix(0, 2 + i) = Format(DateAdd("d", i - 1, BTime), "yyyy-mm-dd")
          If Frm Is FrmMeteoMon Then
            If Frm.DataType = "Data" Then .TextMatrix(0, 2 + i) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm")
            If Frm.DataType <> "Data" Then .TextMatrix(0, 2 + i) = Format(DateAdd("m", i - 1, BTime), "mmm")
          End If
          If Frm Is FrmMeteoYer Then
            If Frm.DataType = "Data" Then .TextMatrix(0, 2 + i) = CStr(CInt(BTime) + i - 1) & ""
            If Frm.DataType = "Norm" Then .TextMatrix(0, 2 + i) = "Normals"
            If Frm.DataType = "Xtrm" Then .TextMatrix(0, 2 + i) = "Historical"
          End If
          If Frm Is FrmMeteoPeriodYer Then .TextMatrix(0, 2 + i) = CStr(CInt(BTime) + i - 1) & ""
        
          .ColWidth(i + 2) = 1500
        Next i
                 
        If (Frm Is FrmMeteoDay) Or (Frm Is FrmMeteoHour) Then
          If Frm.OptionStat(1).Value = True Then .TextMatrix(0, 3 + iTimes) = "Average"
          If Frm.OptionStat(4).Value = True Then .TextMatrix(0, 3 + iTimes) = "Cumulative"
        Else
          .TextMatrix(0, 3 + iTimes) = "Average"
        End If
        
        .TextMatrix(0, 4 + iTimes) = "Max": .TextMatrix(0, 5 + iTimes) = "Min"
        .ColWidth(3 + iTimes) = 1500: .ColWidth(4 + iTimes) = 1500: .ColWidth(5 + iTimes) = 1500
          
          
        If FlagZone = "CITY" Then .Rows = CityNum + 1
        If FlagZone = "CNTY" Then .Rows = CountyNum + 1
        If FlagZone = "TOWN" Then .Rows = TownNum + 1
        For i = 1 To .Rows - 1
          .TextMatrix(i, 0) = i
          If FlagZone = "CITY" Then
            .TextMatrix(i, 1) = CityInfo2(i).Code
            .TextMatrix(i, 2) = CityInfo2(i).City
          ElseIf FlagZone = "CNTY" Then
            .TextMatrix(i, 1) = CountyInfo2(i).Code
            .TextMatrix(i, 2) = CountyInfo2(i).City & CountyInfo2(i).County
          ElseIf FlagZone = "TOWN" Then
            .TextMatrix(i, 1) = TownInfo2(i).Code
            .TextMatrix(i, 2) = TownInfo2(i).City & TownInfo2(i).County & TownInfo2(i).Town
          End If
          '将平均、最大、最小值列涂灰色
          .Row = i: .Col = iTimes + 3
          If Not GuestMode Then .CellBackColor = &HE0E0E0
          .Row = i: .Col = iTimes + 4
          If Not GuestMode Then .CellBackColor = &HE0E0E0
          .Row = i: .Col = iTimes + 5
          If Not GuestMode Then .CellBackColor = &HE0E0E0
        Next i
          
          
        '******单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
          
        .Refresh
        .Redraw = True
        End With
      
      Next k
  
  End If
  
End Sub



'************************ 旬/季统计数据结果表格的初始化（iStages-月份数/年份数，iTimes-旬数/季数） *****************************
Public Sub TableIni_MeteoTenQtr(Frm As Object, BTime$, iStages%, iTimes%)
  With Frm
    If FlagZone = "STA" Then '输出站点结果
      .HFGrid1.Visible = True
      .HFGrid2.Visible = True
      .HFGrid3(1).Visible = False
      .HFGrid3(2).Visible = False
      .HFGrid3(3).Visible = False
    Else '输出区域统计结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid3(1).Visible = True
      .HFGrid3(2).Visible = True
      .HFGrid3(3).Visible = True
    End If
  End With
  
  '**************输出站点结果：设置HFGrid1和HFGrid2两个表格**************
  If FlagZone = "STA" Then
  
    '*********************初始化上方表格*******************************
      With Frm.HFGrid1
      .Redraw = False
      .Clear: .Refresh
      .Cols = 3 + iTimes + 6   '统计区间均值、最大值、最小值，镇、县、市（国家站同样显示）
        
        
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Station Code": .TextMatrix(0, 2) = "Station Name"
      
      '******************将日期及相关统计量写入表头行********************
      For i = 1 To iStages
        If Frm Is FrmMeteoTen Then
          If Frm.DataType = "Data" Then
            .TextMatrix(0, 3 * i) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm" & "-1")
            .TextMatrix(0, 3 * i + 1) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm" & "-2")
            .TextMatrix(0, 3 * i + 2) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm" & "-3")
            
          ElseIf Frm.DataType <> "Data" Then
            .TextMatrix(0, 3 * i) = Format(DateAdd("m", i - 1, BTime), "mmm-1")
            .TextMatrix(0, 3 * i + 1) = Format(DateAdd("m", i - 1, BTime), "mmm-2")
            .TextMatrix(0, 3 * i + 2) = Format(DateAdd("m", i - 1, BTime), "mmm-3")
          
          End If
          .ColWidth(3 * i) = 1500
          .ColWidth(3 * i + 1) = 1500
          .ColWidth(3 * i + 2) = 1500
          
        ElseIf Frm Is FrmMeteoQtr Then

          If Frm.DataType = "Data" Then
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i - 1) = CStr(CInt(BTime) + i - 1) & " Spring"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i) = CStr(CInt(BTime) + i - 1) & " Summer"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 1) = CStr(CInt(BTime) + i - 1) & " Autumn"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 2) = CStr(CInt(BTime) + i - 1) & " Winter"
            
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i - 1) = CStr(CInt(BTime) + i - 1) & " Q1"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i) = CStr(CInt(BTime) + i - 1) & " Q2"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 1) = CStr(CInt(BTime) + i - 1) & " Q3"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 2) = CStr(CInt(BTime) + i - 1) & " Q4"
            
          ElseIf Frm.DataType <> "Data" Then
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i - 1) = "Spring"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i) = "Summer"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 1) = "Autumn"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 2) = "Winter"
          
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i - 1) = "Q1"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i) = "Q2"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 1) = "Q3"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 2) = "Q4"
          End If
          
          .ColWidth(4 * i - 1) = 1500
          .ColWidth(4 * i) = 1500
          .ColWidth(4 * i + 1) = 1500
          .ColWidth(4 * i + 2) = 1500
          
          
        End If
      Next i
      
      If (Frm Is FrmMeteoDay) Or (Frm Is FrmMeteoHour) Then
'        If Frm.OptionStat(1).Value = True Then .TextMatrix(0, 3 + iTimes) = "Average"
'        If Frm.OptionStat(4).Value = True Then .TextMatrix(0, 3 + iTimes) = "Cumulative"
      Else
        .TextMatrix(0, 3 + iTimes) = "Average"
      End If
      
      .TextMatrix(0, 4 + iTimes) = "Max": .TextMatrix(0, 5 + iTimes) = "Min"
      .ColWidth(3 + iTimes) = 1000: .ColWidth(4 + iTimes) = 1000: .ColWidth(5 + iTimes) = 1000
      
      .TextMatrix(0, 6 + iTimes) = "Township": .TextMatrix(0, 7 + iTimes) = "County": .TextMatrix(0, 8 + iTimes) = "City"
      .ColWidth(6 + iTimes) = 1000: .ColWidth(7 + iTimes) = 1000: .ColWidth(8 + iTimes) = 1000
      
      If ShowLonLat = True Then
        .TextMatrix(0, 9 + iTimes) = "Longitude": .TextMatrix(0, 10 + iTimes) = "Latitude"
        .ColWidth(9 + iTimes) = 1000: .ColWidth(10 + iTimes) = 1000
      End If
      
      '******************填充站点信息信息********************
      If Frm Is FrmMeteoHour Then
      
      Else
        
        If Frm.XtrmYear = False Then '不显示累年极值出现的年份
          .Rows = StaNum + 1
          For i = 1 To StaNum
            .TextMatrix(i, 0) = i
            .TextMatrix(i, 1) = StaInfo(i).stacode
            .TextMatrix(i, 2) = StaInfo(i).staname
            
            '将平均、最大、最小值列涂灰色
            .Row = i: .Col = iTimes + 3
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            .Row = i: .Col = iTimes + 4
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            .Row = i: .Col = iTimes + 5
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            
            .TextMatrix(i, iTimes + 6) = StaInfo(i).Town
            .TextMatrix(i, iTimes + 7) = StaInfo(i).County
            .TextMatrix(i, iTimes + 8) = StaInfo(i).City
                        
            If ShowLonLat = True Then
              .TextMatrix(i, iTimes + 9) = StaInfo(i).Longitude
              .TextMatrix(i, iTimes + 10) = StaInfo(i).Latitude
            End If
          Next i
          
        ElseIf Frm.XtrmYear = True Then '显示累年极值出现的年份
          .Rows = StaNum * 2 + 1
          For i = 1 To StaNum
            .TextMatrix(i * 2 - 1, 0) = i
            .TextMatrix(i * 2 - 1, 1) = StaInfo(i).stacode
            .TextMatrix(i * 2 - 1, 2) = StaInfo(i).staname
            .TextMatrix(i * 2, 0) = i
            .TextMatrix(i * 2, 1) = StaInfo(i).stacode
            .TextMatrix(i * 2, 2) = StaInfo(i).staname
            
            '将年份行涂灰色
              For j = 1 To iTimes
                .Row = i * 2: .Col = 2 + j: .CellBackColor = &HE0E0E0
              Next j
            
            '将平均、最大、最小值列涂灰色
            .Row = i * 2 - 1: .Col = iTimes + 3
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            .Row = i * 2 - 1: .Col = iTimes + 4
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            .Row = i * 2 - 1: .Col = iTimes + 5
            If Not GuestMode Then .CellBackColor = &HE0E0E0
            
            .TextMatrix(i * 2 - 1, iTimes + 6) = StaInfo(i).Town
            .TextMatrix(i * 2 - 1, iTimes + 7) = StaInfo(i).County
            .TextMatrix(i * 2 - 1, iTimes + 8) = StaInfo(i).City
            
            If ShowLonLat = True Then
              .TextMatrix(i * 2 - 1, iTimes + 9) = StaInfo(i).Longitude
              .TextMatrix(i * 2 - 1, iTimes + 10) = StaInfo(i).Latitude
            End If
          Next i
        
        End If
      
      End If
      
      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Row <> 0 Then
'            If .Col = 0 Then .CellBackColor = &H80000005
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          End If
        Next j
      Next i
      
      .Refresh
      .Redraw = True
      End With
          
    '*********************初始化下方表格*******************************
      With Frm.HFGrid2
      .Redraw = False
      .Clear: .Refresh
      
      
      .Cols = 3 + iTimes + 6   '统计区间均值、最大值、最小值，镇、县、市（国家站同样显示）
       If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      
      '******************将日期及相关统计量写入表头行********************
      For i = 1 To iStages
        If Frm Is FrmMeteoTen Then
          If Frm.DataType = "Data" Then
            .TextMatrix(0, 3 * i) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm" & "-1")
            .TextMatrix(0, 3 * i + 1) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm" & "-2")
            .TextMatrix(0, 3 * i + 2) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm" & "-3")
            
          ElseIf Frm.DataType <> "Data" Then
            .TextMatrix(0, 3 * i) = Format(DateAdd("m", i - 1, BTime), "mmm-1")
            .TextMatrix(0, 3 * i + 1) = Format(DateAdd("m", i - 1, BTime), "mmm-2")
            .TextMatrix(0, 3 * i + 2) = Format(DateAdd("m", i - 1, BTime), "mmm-3")
          
          End If
          .ColWidth(3 * i) = 1500
          .ColWidth(3 * i + 1) = 1500
          .ColWidth(3 * i + 2) = 1500
          
        ElseIf Frm Is FrmMeteoQtr Then

          If Frm.DataType = "Data" Then
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i - 1) = CStr(CInt(BTime) + i - 1) & " Spring"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i) = CStr(CInt(BTime) + i - 1) & " Summer"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 1) = CStr(CInt(BTime) + i - 1) & " Autumn"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 2) = CStr(CInt(BTime) + i - 1) & " Winter"
            
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i - 1) = CStr(CInt(BTime) + i - 1) & " Q1"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i) = CStr(CInt(BTime) + i - 1) & " Q2"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 1) = CStr(CInt(BTime) + i - 1) & " Q3"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 2) = CStr(CInt(BTime) + i - 1) & " Q4"
            
          ElseIf Frm.DataType <> "Data" Then
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i - 1) = "Spring"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i) = "Summer"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 1) = "Autumn"
            If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 2) = "Winter"
          
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i - 1) = "Q1"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i) = "Q2"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 1) = "Q3"
            If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 2) = "Q4"
          End If
          
          .ColWidth(4 * i - 1) = 1500
          .ColWidth(4 * i) = 1500
          .ColWidth(4 * i + 1) = 1500
          .ColWidth(4 * i + 2) = 1500
          
        End If
      Next i
      
'      .TextMatrix(0, 3 + iTimes) = "Average": .TextMatrix(0, 4 + iTimes) = "Max": .TextMatrix(0, 5 + iTimes) = "Min"
      .ColWidth(3 + iTimes) = 1000: .ColWidth(4 + iTimes) = 1000: .ColWidth(5 + iTimes) = 1000
      .ColWidth(6 + iTimes) = 1000: .ColWidth(7 + iTimes) = 1000: .ColWidth(8 + iTimes) = 1000
      If ShowLonLat = True Then
        .ColWidth(9 + iTimes) = 1000: .ColWidth(10 + iTimes) = 1000
      End If
      
      '******************填充站点信息或地市信息********************
      .Rows = 3 + 1: .FixedRows = 1
      .TextMatrix(1, 2) = "Average"
      .TextMatrix(2, 2) = "Max"
      .TextMatrix(3, 2) = "Min"

      '******下方单元格左对齐、着色********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Row <> 0 Then
'            If .Col = 0 Then .CellBackColor = &H80000005
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          End If
        Next j
      Next i
      
      .Refresh
      .Redraw = True
      End With
      
      If Frm.HFGrid1.Rows > 2 Then Frm.HFGrid1.Row = 2
      
  Else '输出区域统计结果：设置HFGrid3、HFGrid4、HFGrid5三个表格
      
      '*********************循环初始化三个表格*******************************
      For k = 1 To 3
        With Frm.HFGrid3(k)
        .Redraw = False
        .Clear: .Refresh
          
        .Cols = 3 + iTimes + 3   '统计区间均值、最大值、最小值
      
        .FixedCols = 3
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Region Code": .TextMatrix(0, 2) = "Region Name"
        
        If k = 1 Then .TextMatrix(0, 0) = "Regional Avg"
        If k = 2 Then .TextMatrix(0, 0) = "Regional Max"
        If k = 3 Then .TextMatrix(0, 0) = "Regional Min"
                    
                 
        '******************将日期及相关统计量写入表头行********************
        For i = 1 To iStages
        
           If Frm Is FrmMeteoTen Then
            If Frm.DataType = "Data" Then
              .TextMatrix(0, 3 * i) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm" & "-1")
              .TextMatrix(0, 3 * i + 1) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm" & "-2")
              .TextMatrix(0, 3 * i + 2) = Format(DateAdd("m", i - 1, BTime), "yyyy-mm" & "-3")
                
            ElseIf Frm.DataType <> "Data" Then
              .TextMatrix(0, 3 * i) = Format(DateAdd("m", i - 1, BTime), "mmm-1")
              .TextMatrix(0, 3 * i + 1) = Format(DateAdd("m", i - 1, BTime), "mmm-2")
              .TextMatrix(0, 3 * i + 2) = Format(DateAdd("m", i - 1, BTime), "mmm-3")
              
            End If
            .ColWidth(3 * i) = 1500
            .ColWidth(3 * i + 1) = 1500
            .ColWidth(3 * i + 2) = 1500
              
          ElseIf Frm Is FrmMeteoQtr Then
    
            If Frm.DataType = "Data" Then
              If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i - 1) = CStr(CInt(BTime) + i - 1) & " Spring"
              If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i) = CStr(CInt(BTime) + i - 1) & " Summer"
              If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 1) = CStr(CInt(BTime) + i - 1) & " Autumn"
              If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 2) = CStr(CInt(BTime) + i - 1) & " Winter"
              
              If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i - 1) = CStr(CInt(BTime) + i - 1) & " Q1"
              If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i) = CStr(CInt(BTime) + i - 1) & " Q2"
              If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 1) = CStr(CInt(BTime) + i - 1) & " Q3"
              If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 2) = CStr(CInt(BTime) + i - 1) & " Q4"
              
            ElseIf Frm.DataType <> "Data" Then
              If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i - 1) = "Spring"
              If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i) = "Summer"
              If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 1) = "Autumn"
              If Frm.OptionSsn.Value = True Then .TextMatrix(0, 4 * i + 2) = "Winter"
            
              If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i - 1) = "Q1"
              If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i) = "Q2"
              If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 1) = "Q3"
              If Frm.OptionQtr.Value = True Then .TextMatrix(0, 4 * i + 2) = "Q4"
            End If
            
            .ColWidth(4 * i - 1) = 1500
            .ColWidth(4 * i) = 1500
            .ColWidth(4 * i + 1) = 1500
            .ColWidth(4 * i + 2) = 1500
                        
          End If
        Next i
        
        .TextMatrix(0, 3 + iTimes) = "Max": .TextMatrix(0, 4 + iTimes) = "Max": .TextMatrix(0, 5 + iTimes) = "Min"
        .ColWidth(3 + iTimes) = 1000: .ColWidth(4 + iTimes) = 1000: .ColWidth(5 + iTimes) = 1000
          
          
        If FlagZone = "CITY" Then .Rows = CityNum + 1
        If FlagZone = "CNTY" Then .Rows = CountyNum + 1
        If FlagZone = "TOWN" Then .Rows = TownNum + 1
        For i = 1 To .Rows - 1
          .TextMatrix(i, 0) = i
          If FlagZone = "CITY" Then
            .TextMatrix(i, 1) = CityInfo2(i).Code
            .TextMatrix(i, 2) = CityInfo2(i).City
          ElseIf FlagZone = "CNTY" Then
            .TextMatrix(i, 1) = CountyInfo2(i).Code
            .TextMatrix(i, 2) = CountyInfo2(i).City & CountyInfo2(i).County
          ElseIf FlagZone = "TOWN" Then
            .TextMatrix(i, 1) = TownInfo2(i).Code
            .TextMatrix(i, 2) = TownInfo2(i).City & TownInfo2(i).County & TownInfo2(i).Town
          End If
          '将平均、最大、最小值列涂灰色
          .Row = i: .Col = iTimes + 3
          If Not GuestMode Then .CellBackColor = &HE0E0E0
          .Row = i: .Col = iTimes + 4
          If Not GuestMode Then .CellBackColor = &HE0E0E0
          .Row = i: .Col = iTimes + 5
          If Not GuestMode Then .CellBackColor = &HE0E0E0
        Next i
          
          
        '******单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
          
          
        .Refresh
        .Redraw = True
        End With
      
      Next k
  
  End If
  
End Sub




'************************ 任意时段统计数据结果表格的初始化(单要素) *****************************
Public Sub TableIni_MeteoPeriod(Frm As Object, SelField_Initial$)
  Dim iItems_stat% '统计的项目数量
  Dim iItems_rank%
  Dim iItems_zone%
  
  iItems_rank = 7 '排名相关的项目数量：排名(大-小)、排名(小-大)、排名年份、最大值年、最大值、最小值年、最小值
  iItems_zone = 3 '区域信息数据：镇乡、区县、地市
  
  If Frm.FlagStat = "AVE" Or Frm.FlagStat = "SUM" Or Frm.FlagStat = "DAYS" Then iItems_stat = 5
  If Frm.FlagStat = "MAX" Or Frm.FlagStat = "MIN" Then iItems_stat = 6
  
  With Frm
    If FlagZone = "STA" Then '输出站点结果
      .HFGrid1.Visible = True
      .HFGrid2.Visible = True
      .HFGrid3(1).Visible = False
      .HFGrid3(2).Visible = False
      .HFGrid3(3).Visible = False
    Else '输出区域统计结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid3(1).Visible = True
      .HFGrid3(2).Visible = True
      .HFGrid3(3).Visible = True
    End If
  End With

  '**************输出站点结果：设置HFGrid1和HFGrid2两个表格**************
  If FlagZone = "STA" Then
  
    '*********************初始化上方表格*******************************
      With Frm.HFGrid1
      .Redraw = False
      .Clear: .Refresh
      
      .Cols = 3 + iItems_stat + iItems_rank + iItems_zone '排名相关的7列、镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Station Code": .TextMatrix(0, 2) = "Station Name"
      
      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 1600
      Next i
      If Frm.FlagStat = "AVE" Or Frm.FlagStat = "SUM" Then
        If Frm.FlagStat = "AVE" Then .TextMatrix(0, 3) = "Average": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "Yearly Avg"
        If Frm.FlagStat = "SUM" Then .TextMatrix(0, 3) = "Cumulative": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "Yearly Total"
        If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R08_20" Or SelField_Initial = "R20_08" Or SelField_Initial = "S" Then
         .TextMatrix(0, 5) = "Anomaly %": .TextMatrix(0, 7) = "Anomaly %"
        Else
          .TextMatrix(0, 5) = "Anomaly": .TextMatrix(0, 7) = "Anomaly"
        End If
      
      ElseIf Frm.FlagStat = "MAX" Or Frm.FlagStat = "MIN" Then
        If Frm.FlagStat = "MAX" Then .TextMatrix(0, 3) = "Max": .TextMatrix(0, 5) = "MaxRecord": .TextMatrix(0, 7) = "Yearly Max":
        If Frm.FlagStat = "MIN" Then .TextMatrix(0, 3) = "Min": .TextMatrix(0, 5) = "MinRecord": .TextMatrix(0, 7) = "Yearly Min":
        .TextMatrix(0, 4) = "OccurrDate": .TextMatrix(0, 6) = "OccurrDate": .TextMatrix(0, 8) = "OccurrDate"
      
      ElseIf Frm.FlagStat = "DAYS" Then
        .TextMatrix(0, 3) = "ConditionDays": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "YearlyDays"
        .TextMatrix(0, 5) = "Anomaly": .TextMatrix(0, 7) = "Anomaly"
      End If
      
      .TextMatrix(0, 3 + iItems_stat) = "Rank(High-Low)"
      .TextMatrix(0, 4 + iItems_stat) = "Rank(Low-High)"
      
      .ColWidth(5 + iItems_stat) = 1800
      .TextMatrix(0, 5 + iItems_stat) = "Ranking Year"
      
      For i = 1 To iItems_rank - 3
        .ColWidth(i + 5 + iItems_stat) = 1000
        If i = 1 Then .TextMatrix(0, i + 5 + iItems_stat) = "MaxYear"
        If i = 2 Then .TextMatrix(0, i + 5 + iItems_stat) = "Max"
        If i = 3 Then .TextMatrix(0, i + 5 + iItems_stat) = "MinYear"
        If i = 4 Then .TextMatrix(0, i + 5 + iItems_stat) = "Min"
      Next i
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat + iItems_rank) = 1000
        If i = 1 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "Township"
        If i = 2 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "County"
        If i = 3 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "City"
      Next i
      
      
      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + iItems_rank + 4) = 1000: .ColWidth(2 + iItems_stat + iItems_rank + 5) = 1000
        .TextMatrix(0, 2 + iItems_stat + iItems_rank + 4) = "Longitude": .TextMatrix(0, 2 + iItems_stat + iItems_rank + 5) = "Latitude"
      End If
      
      '******************填充站点信息信息********************
        .Rows = StaNum + 1
        For i = 1 To StaNum
          .TextMatrix(i, 0) = i
          .TextMatrix(i, 1) = StaInfo(i).stacode
          .TextMatrix(i, 2) = StaInfo(i).staname
          

          .TextMatrix(i, 1 + 2 + iItems_stat + iItems_rank) = StaInfo(i).Town
          .TextMatrix(i, 2 + 2 + iItems_stat + iItems_rank) = StaInfo(i).County
          .TextMatrix(i, 3 + 2 + iItems_stat + iItems_rank) = StaInfo(i).City
                      
          If ShowLonLat = True Then
            .TextMatrix(i, 2 + iItems_stat + iItems_rank + 4) = StaInfo(i).Longitude
            .TextMatrix(i, 2 + iItems_stat + iItems_rank + 5) = StaInfo(i).Latitude
          End If
        Next i

      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Row <> 0 Then
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          End If
        Next j
      Next i
      
      .Refresh
      .Redraw = True
      End With
      

    '*********************初始化下方表格*******************************
      With Frm.HFGrid2
      .Redraw = False
      .Clear: .Refresh

      .Cols = 3 + iItems_stat + iItems_rank + iItems_zone '排名相关的7列、镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      
      For i = 1 To iItems_stat + 2 '排名（大-小）、排名（小-大）也设置宽1600
        .ColWidth(i + 2) = 1600
      Next i
      
      .ColWidth(5 + iItems_stat) = 1800 '排名年份设置宽1800
      
      For i = 1 To iItems_rank - 3 '最大值年、最大值、最小值年、最小值设宽1000
        .ColWidth(i + 5 + iItems_stat) = 1000
      Next i
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat + iItems_rank) = 1000
      Next i
      
      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + iItems_rank + 4) = 1000: .ColWidth(2 + iItems_stat + iItems_rank + 5) = 1000
      End If

      .Rows = 3
      .TextMatrix(0, 2) = "Average": .TextMatrix(1, 2) = "Max": .TextMatrix(2, 2) = "Min"
      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Col = 1 Then .CellBackColor = &H80000005
          If .Col = 2 Then .CellBackColor = &H80000005
        Next j
      Next i
      
      .Refresh
      .Redraw = True
    End With
    If FrmMeteoPeriod.HFGrid1.Rows > 2 Then FrmMeteoPeriod.HFGrid1.Row = 2

  Else '输出区域统计结果：设置HFGrid3、HFGrid4、HFGrid5三个表格
      
      '*********************循环初始化三个表格*******************************
      For k = 1 To 3
        With Frm.HFGrid3(k)
        .Redraw = False
        .Clear: .Refresh
      
        .Cols = 3 + iItems_stat + iItems_rank  '排名相关的7列
        .FixedCols = 3
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        If k = 1 Then .TextMatrix(0, 0) = "Regional Avg"
        If k = 2 Then .TextMatrix(0, 0) = "Regional Max"
        If k = 3 Then .TextMatrix(0, 0) = "Regional Min"
        .TextMatrix(0, 1) = "Region Code": .TextMatrix(0, 2) = "Region Name"
          
        For i = 1 To iItems_stat + 2
          .ColWidth(i + 2) = 1600
        Next i
        If Frm.FlagStat = "AVE" Or Frm.FlagStat = "SUM" Then
          If Frm.FlagStat = "AVE" Then .TextMatrix(0, 3) = "Average": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "Yearly Avg"
          If Frm.FlagStat = "SUM" Then .TextMatrix(0, 3) = "Cumulative": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "Yearly Total"
          If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R08_20" Or SelField_Initial = "R20_08" Or SelField_Initial = "S" Then
           .TextMatrix(0, 5) = "Anomaly %": .TextMatrix(0, 7) = "Anomaly %"
          Else
            .TextMatrix(0, 5) = "Anomaly": .TextMatrix(0, 7) = "Anomaly"
          End If
          
        ElseIf Frm.FlagStat = "MAX" Or Frm.FlagStat = "MIN" Then
          If Frm.FlagStat = "MAX" Then .TextMatrix(0, 3) = "Max": .TextMatrix(0, 5) = "MaxRecord": .TextMatrix(0, 7) = "Yearly Max":
          If Frm.FlagStat = "MIN" Then .TextMatrix(0, 3) = "Min": .TextMatrix(0, 5) = "MinRecord": .TextMatrix(0, 7) = "Yearly Min":
          .TextMatrix(0, 4) = "OccurrDate": .TextMatrix(0, 6) = "OccurrDate": .TextMatrix(0, 8) = "OccurrDate"
        
        ElseIf Frm.FlagStat = "DAYS" Then
          .TextMatrix(0, 3) = "ConditionDays": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "YearlyDays"
          .TextMatrix(0, 5) = "Anomaly": .TextMatrix(0, 7) = "Anomaly"
        End If
         
        If k = 1 Then
          .TextMatrix(0, 3 + iItems_stat) = "Rank(High-Low)"
          .TextMatrix(0, 4 + iItems_stat) = "Rank(Low-High)"
            
          .ColWidth(5 + iItems_stat) = 1800
          .TextMatrix(0, 5 + iItems_stat) = "Ranking Year"
              
          For i = 1 To iItems_rank - 3
            .ColWidth(i + 5 + iItems_stat) = 1000
            If i = 1 Then .TextMatrix(0, i + 5 + iItems_stat) = "MaxYear"
            If i = 2 Then .TextMatrix(0, i + 5 + iItems_stat) = "Max"
            If i = 3 Then .TextMatrix(0, i + 5 + iItems_stat) = "MinYear"
            If i = 4 Then .TextMatrix(0, i + 5 + iItems_stat) = "Min"
          Next i
        Else
          .ColWidth(5 + iItems_stat) = 1800
          For i = 1 To iItems_rank - 3
            .ColWidth(i + 5 + iItems_stat) = 1000
          Next i
        End If
          
        '******************填充区域信息********************
        If FlagZone = "CITY" Then .Rows = CityNum + 1
        If FlagZone = "CNTY" Then .Rows = CountyNum + 1
        If FlagZone = "TOWN" Then .Rows = TownNum + 1
        For i = 1 To .Rows - 1
          .TextMatrix(i, 0) = i
            
          If FlagZone = "CITY" Then
            .TextMatrix(i, 1) = CityInfo2(i).Code
            .TextMatrix(i, 2) = CityInfo2(i).City
          ElseIf FlagZone = "CNTY" Then
            .TextMatrix(i, 1) = CountyInfo2(i).Code
            .TextMatrix(i, 2) = CountyInfo2(i).City & CountyInfo2(i).County
          ElseIf FlagZone = "TOWN" Then
            .TextMatrix(i, 1) = TownInfo2(i).Code
            .TextMatrix(i, 2) = TownInfo2(i).City & TownInfo2(i).County & TownInfo2(i).Town
          End If
        Next i
    
        '******上方单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
   
        .Refresh
        .Redraw = True
        End With
      
      Next k
  End If

          
End Sub




'************************ 任意时段统计数据结果表格的初始化(多要素) 2022-4-17*****************************
Public Sub TableIni_MeteoPeriod2(Frm As Object) '要素初始ID、要素名称，统计量
  Dim iItems_stat% '统计的项目数量
  Dim iItems_zone%
  Dim Col_name$ '要素的中文名
  
  iItems_stat = UBound(SelFields_Initial) * 3 '任何统计量都对应三列结果
  iItems_zone = 3 '区域信息数据：镇乡、区县、地市
  
  With Frm
    If FlagZone = "STA" Then '输出站点结果
      .HFGrid1.Visible = True
      .HFGrid2.Visible = True
      .HFGrid3(1).Visible = False
      .HFGrid3(2).Visible = False
      .HFGrid3(3).Visible = False
    Else '输出区域统计结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid3(1).Visible = True
      .HFGrid3(2).Visible = True
      .HFGrid3(3).Visible = True
    End If
  End With

  '**************输出站点结果：设置HFGrid1和HFGrid2两个表格**************
  If FlagZone = "STA" Then
  
    '*********************初始化上方表格*******************************
      With Frm.HFGrid1
      .Redraw = False
      .Clear: .Refresh
      
      .Cols = 3 + iItems_stat + iItems_zone '镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Station Code": .TextMatrix(0, 2) = "Station Name"
      
      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 1400
      Next i
      
      For i = 1 To UBound(SelFields_Initial) '按选择要素循环，进行表格初始化
        .TextMatrix(0, 3 * (i - 1) + 2 + 1) = SelNames_Initial(i)
        If FlagStats(i) = "AVE" Then
            .TextMatrix(0, 3 * (i - 1) + 2 + 2) = "Normals"
            .TextMatrix(0, 3 * (i - 1) + 2 + 3) = "Anomaly"
        ElseIf FlagStats(i) = "SUM" Then
            .TextMatrix(0, 3 * (i - 1) + 2 + 2) = "Normals"
            If SelFields_Initial(i) = "R" Or SelFields_Initial(i) = "R08" Or SelFields_Initial(i) = "R08_20" Or SelFields_Initial(i) = "R20_08" Or SelFields_Initial(i) = "S" Then
              .TextMatrix(0, 3 * (i - 1) + 2 + 3) = "Anomaly %"
            Else
              .TextMatrix(0, 3 * (i - 1) + 2 + 3) = "Anomaly"
            End If
        ElseIf FlagStats(i) = "MAX" Or FlagStats(i) = "MIN" Then
            .TextMatrix(0, 3 * (i - 1) + 2 + 2) = "OccurrDate"
            .TextMatrix(0, 3 * (i - 1) + 2 + 3) = "Historical Extreme"
        End If
      Next i
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat + iItems_rank) = 1000
        If i = 1 Then .TextMatrix(0, i + 2 + iItems_stat) = "Township"
        If i = 2 Then .TextMatrix(0, i + 2 + iItems_stat) = "County"
        If i = 3 Then .TextMatrix(0, i + 2 + iItems_stat) = "City"
      Next i
      
      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
        .TextMatrix(0, 2 + iItems_stat + 4) = "Longitude": .TextMatrix(0, 2 + iItems_stat + 5) = "Latitude"
      End If
      
      '******************填充站点信息信息********************
      .Rows = StaNum + 1
      For i = 1 To StaNum
        .TextMatrix(i, 0) = i
        .TextMatrix(i, 1) = StaInfo(i).stacode
        .TextMatrix(i, 2) = StaInfo(i).staname
        
        .TextMatrix(i, 1 + 2 + iItems_stat) = StaInfo(i).Town
        .TextMatrix(i, 2 + 2 + iItems_stat) = StaInfo(i).County
        .TextMatrix(i, 3 + 2 + iItems_stat) = StaInfo(i).City
                      
        If ShowLonLat = True Then
          .TextMatrix(i, 2 + iItems_stat + 4) = StaInfo(i).Longitude
          .TextMatrix(i, 2 + iItems_stat + 5) = StaInfo(i).Latitude
        End If
      Next i

      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Row <> 0 Then
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          End If
        Next j
      Next i
      
      .Refresh
      .Redraw = True
      End With
      

    '*********************初始化下方表格*******************************
      With Frm.HFGrid2
      .Redraw = False
      .Clear: .Refresh
      
      .Cols = 3 + iItems_stat + iItems_zone  '排名相关的7列、镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度

      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500

      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 1400
      Next i


      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat) = 1000
      Next i

      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
      End If

      .Rows = 3
      .TextMatrix(0, 2) = "Average": .TextMatrix(1, 2) = "Max": .TextMatrix(2, 2) = "Min"
      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Col = 1 Then .CellBackColor = &H80000005
          If .Col = 2 Then .CellBackColor = &H80000005
        Next j
      Next i

      .Refresh
      .Redraw = True
    End With
    If FrmMeteoPeriod.HFGrid1.Rows > 2 Then FrmMeteoPeriod.HFGrid1.Row = 2

  Else '输出区域统计结果：设置HFGrid3、HFGrid4、HFGrid5三个表格
  
    '*********************循环初始化三个表格*******************************
    For k = 1 To 3
      With Frm.HFGrid3(k)
      .Redraw = False
      .Clear: .Refresh

      .Cols = 3 + iItems_stat + iItems_zone '镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
        
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      .TextMatrix(0, 1) = "Region Code": .TextMatrix(0, 2) = "Region Name"
      If k = 1 Then .TextMatrix(0, 0) = "Regional Avg"
      If k = 2 Then .TextMatrix(0, 0) = "Regional Max"
      If k = 3 Then .TextMatrix(0, 0) = "Regional Min"
        
      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 1400
      Next i
        
      For i = 1 To UBound(SelFields_Initial) '按选择要素循环，进行表格初始化
        .TextMatrix(0, 3 * (i - 1) + 2 + 1) = SelNames_Initial(i)
        If FlagStats(i) = "AVE" Then
            .TextMatrix(0, 3 * (i - 1) + 2 + 2) = "Normals"
            .TextMatrix(0, 3 * (i - 1) + 2 + 3) = "Anomaly"
        ElseIf FlagStats(i) = "SUM" Then
            .TextMatrix(0, 3 * (i - 1) + 2 + 2) = "Normals"
            If SelFields_Initial(i) = "R" Or SelFields_Initial(i) = "R08" Or SelFields_Initial(i) = "R08_20" Or SelFields_Initial(i) = "R20_08" Or SelFields_Initial(i) = "S" Then
              .TextMatrix(0, 3 * (i - 1) + 2 + 3) = "Anomaly %"
            Else
              .TextMatrix(0, 3 * (i - 1) + 2 + 3) = "Anomaly"
            End If
        ElseIf FlagStats(i) = "MAX" Or FlagStats(i) = "MIN" Then
            .TextMatrix(0, 3 * (i - 1) + 2 + 2) = "OccurrDate"
            .TextMatrix(0, 3 * (i - 1) + 2 + 3) = "Historical Extreme"
        End If
      Next i

      '******************填充区域信息********************
      If FlagZone = "CITY" Then .Rows = CityNum + 1
      If FlagZone = "CNTY" Then .Rows = CountyNum + 1
      If FlagZone = "TOWN" Then .Rows = TownNum + 1
      For i = 1 To .Rows - 1
        .TextMatrix(i, 0) = i
        
        If FlagZone = "CITY" Then
          .TextMatrix(i, 1) = CityInfo2(i).Code
          .TextMatrix(i, 2) = CityInfo2(i).City
        ElseIf FlagZone = "CNTY" Then
          .TextMatrix(i, 1) = CountyInfo2(i).Code
          .TextMatrix(i, 2) = CountyInfo2(i).City & CountyInfo2(i).County
        ElseIf FlagZone = "TOWN" Then
          .TextMatrix(i, 1) = TownInfo2(i).Code
          .TextMatrix(i, 2) = TownInfo2(i).City & TownInfo2(i).County & TownInfo2(i).Town
        End If
      Next i

      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Row <> 0 Then
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          End If
        Next j
      Next i

      .Refresh
      .Redraw = True
      End With

    Next k
  End If
          
End Sub


'************************ 单时段统计数据结果表格的初始化 *****************************
Public Sub TableIni_MeteoPeriodSin(Frm As Object, SelField_Initial$)
  Dim iItems_stat% '统计的项目数量
  Dim iItems_rank%
  Dim iItems_zone%
  
  iItems_zone = 3 '区域信息数据：镇乡、区县、地市
  
  iItems_stat = 7
  
  With Frm
    If FlagZone = "STA" Then '输出站点结果
      .HFGrid1.Visible = True
      .HFGrid2.Visible = True
      .HFGrid3(1).Visible = False
      .HFGrid3(2).Visible = False
      .HFGrid3(3).Visible = False
    Else '输出区域统计结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid3(1).Visible = True
      .HFGrid3(2).Visible = True
      .HFGrid3(3).Visible = True
    End If
  End With

  '**************输出站点结果：设置HFGrid1和HFGrid2两个表格**************
  If FlagZone = "STA" Then
  
    '*********************初始化上方表格*******************************
      With Frm.HFGrid1
      .Redraw = False
      .Clear: .Refresh
      
      .Cols = 3 + iItems_stat + iItems_zone  '镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Station Code": .TextMatrix(0, 2) = "Station Name"
      
      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 1600
      Next i
        
      .TextMatrix(0, 8) = "Cumulative": .TextMatrix(0, 9) = "Valid Count": .TextMatrix(0, 3) = "Average"
      .TextMatrix(0, 4) = "Max": .TextMatrix(0, 5) = "OccurrDate": .TextMatrix(0, 6) = "Min": .TextMatrix(0, 7) = "OccurrDate"
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat) = 1000
        If i = 1 Then .TextMatrix(0, i + 2 + iItems_stat) = "Township"
        If i = 2 Then .TextMatrix(0, i + 2 + iItems_stat) = "County"
        If i = 3 Then .TextMatrix(0, i + 2 + iItems_stat) = "City"
      Next i
      
      
      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
        .TextMatrix(0, 2 + iItems_stat + 4) = "Longitude": .TextMatrix(0, 2 + iItems_stat + iItems_rank + 5) = "Latitude"
      End If
      
      '******************填充站点信息信息********************
        .Rows = StaNum + 1
        For i = 1 To StaNum
          .TextMatrix(i, 0) = i
          .TextMatrix(i, 1) = StaInfo(i).stacode
          .TextMatrix(i, 2) = StaInfo(i).staname
          

          .TextMatrix(i, 1 + 2 + iItems_stat) = StaInfo(i).Town
          .TextMatrix(i, 2 + 2 + iItems_stat) = StaInfo(i).County
          .TextMatrix(i, 3 + 2 + iItems_stat) = StaInfo(i).City
                      
          If ShowLonLat = True Then
            .TextMatrix(i, 2 + iItems_stat + 4) = StaInfo(i).Longitude
            .TextMatrix(i, 2 + iItems_stat + 5) = StaInfo(i).Latitude
          End If
        Next i

      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Row <> 0 Then
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          End If
        Next j
      Next i
      
      .Refresh
      .Redraw = True
      End With
      

    '*********************初始化下方表格*******************************
      With Frm.HFGrid2
      .Redraw = False
      .Clear: .Refresh

      .Cols = 3 + iItems_stat + iItems_zone '排名相关的7列、镇县市3列
      If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      
      For i = 1 To iItems_stat + 2 '排名（大-小）、排名（小-大）也设置宽1600
        .ColWidth(i + 2) = 1600
      Next i
      
      .ColWidth(5 + iItems_stat) = 1800 '排名年份设置宽1800
      
      For i = 1 To iItems_rank - 3 '最大值年、最大值、最小值年、最小值设宽1000
        .ColWidth(i + 5 + iItems_stat) = 1000
      Next i
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat) = 1000
      Next i
      
      If ShowLonLat = True Then
        .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
      End If

      .Rows = 3
      .TextMatrix(0, 2) = "Average": .TextMatrix(1, 2) = "Max": .TextMatrix(2, 2) = "Min"
      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Col = 1 Then .CellBackColor = &H80000005
          If .Col = 2 Then .CellBackColor = &H80000005
        Next j
      Next i
      
      .Refresh
      .Redraw = True
    End With
    If FrmMeteoPeriodSin.HFGrid1.Rows > 2 Then FrmMeteoPeriodSin.HFGrid1.Row = 2

  Else '输出区域统计结果：设置HFGrid3、HFGrid4、HFGrid5三个表格
      
      '*********************循环初始化三个表格*******************************
      For k = 1 To 3
        With Frm.HFGrid3(k)
        .Redraw = False
        .Clear: .Refresh
      
        .Cols = 3 + iItems_stat  '排名相关的7列
        .FixedCols = 3
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        If k = 1 Then .TextMatrix(0, 0) = "Regional Avg"
        If k = 2 Then .TextMatrix(0, 0) = "Regional Max"
        If k = 3 Then .TextMatrix(0, 0) = "Regional Min"
        .TextMatrix(0, 1) = "Region Code": .TextMatrix(0, 2) = "Region Name"
          
        For i = 1 To iItems_stat
          .ColWidth(i + 2) = 1600
        Next i
                
        .TextMatrix(0, 8) = "Cumulative": .TextMatrix(0, 9) = "Valid Count": .TextMatrix(0, 3) = "Average"
        .TextMatrix(0, 4) = "Max": .TextMatrix(0, 5) = "OccurrDate": .TextMatrix(0, 6) = "Min": .TextMatrix(0, 7) = "OccurrDate"
      
      
        '******************填充区域信息********************
        If FlagZone = "CITY" Then .Rows = CityNum + 1
        If FlagZone = "CNTY" Then .Rows = CountyNum + 1
        If FlagZone = "TOWN" Then .Rows = TownNum + 1
        For i = 1 To .Rows - 1
          .TextMatrix(i, 0) = i
            
          If FlagZone = "CITY" Then
            .TextMatrix(i, 1) = CityInfo2(i).Code
            .TextMatrix(i, 2) = CityInfo2(i).City
          ElseIf FlagZone = "CNTY" Then
            .TextMatrix(i, 1) = CountyInfo2(i).Code
            .TextMatrix(i, 2) = CountyInfo2(i).City & CountyInfo2(i).County
          ElseIf FlagZone = "TOWN" Then
            .TextMatrix(i, 1) = TownInfo2(i).Code
            .TextMatrix(i, 2) = TownInfo2(i).City & TownInfo2(i).County & TownInfo2(i).Town
          End If
        Next i
    
        '******上方单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
   
        .Refresh
        .Redraw = True
        End With
      
      Next k
  End If

          
End Sub





'************************ 2026-7-16：区域统计数据结果表格的初始化 *****************************
Public Sub TableIni_MeteoPeriodExt(Frm As Object, SelField_Initial$)
  Dim iItems_stat% '统计的项目数量
  Dim iItems_rank%
  Dim iItems_zone%
  
  iItems_rank = 7 '排名相关的项目数量：排名(大-小)、排名(小-大)、排名年份、最大值年、最大值、最小值年、最小值
  iItems_zone = 3 '区域信息数据：区县、地市，省份
  
  If Frm.FlagStat = "AVE" Or Frm.FlagStat = "SUM" Or Frm.FlagStat = "DAYS" Then iItems_stat = 5
  If Frm.FlagStat = "MAX" Or Frm.FlagStat = "MIN" Then iItems_stat = 6
  
  With Frm
    If FlagZoneExt = "STA" Then '输出站点结果
      .HFGrid1.Visible = True
      .HFGrid2.Visible = True
      .HFGrid3(1).Visible = False
      .HFGrid3(2).Visible = False
      .HFGrid3(3).Visible = False
    Else '输出区域统计结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid3(1).Visible = True
      .HFGrid3(2).Visible = True
      .HFGrid3(3).Visible = True
    End If
  End With

  '**************输出站点结果：设置HFGrid1和HFGrid2两个表格**************
  If FlagZoneExt = "STA" Then
  
    '*********************初始化上方表格*******************************
      With Frm.HFGrid1
      .Redraw = False
      .Clear: .Refresh
      
      .Cols = 3 + iItems_stat + iItems_rank + iItems_zone '排名相关的7列、镇县市3列
      If ShowLonLatExt = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Station Code": .TextMatrix(0, 2) = "Station Name"
      
      For i = 1 To iItems_stat + 2
        .ColWidth(i + 2) = 1600
      Next i
      If Frm.FlagStat = "AVE" Or Frm.FlagStat = "SUM" Then
        If Frm.FlagStat = "AVE" Then .TextMatrix(0, 3) = "Average": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "Yearly Avg"
        If Frm.FlagStat = "SUM" Then .TextMatrix(0, 3) = "Cumulative": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "Yearly Total"
        If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R08_20" Or SelField_Initial = "R20_08" Or SelField_Initial = "S" Then
         .TextMatrix(0, 5) = "Anomaly %": .TextMatrix(0, 7) = "Anomaly %"
        Else
          .TextMatrix(0, 5) = "Anomaly": .TextMatrix(0, 7) = "Anomaly"
        End If
      
      ElseIf Frm.FlagStat = "MAX" Or Frm.FlagStat = "MIN" Then
        If Frm.FlagStat = "MAX" Then .TextMatrix(0, 3) = "Max": .TextMatrix(0, 5) = "MaxRecord": .TextMatrix(0, 7) = "Yearly Max":
        If Frm.FlagStat = "MIN" Then .TextMatrix(0, 3) = "Min": .TextMatrix(0, 5) = "MinRecord": .TextMatrix(0, 7) = "Yearly Min":
        .TextMatrix(0, 4) = "OccurrDate": .TextMatrix(0, 6) = "OccurrDate": .TextMatrix(0, 8) = "OccurrDate"
      
      ElseIf Frm.FlagStat = "DAYS" Then
        .TextMatrix(0, 3) = "ConditionDays": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "YearlyDays"
        .TextMatrix(0, 5) = "Anomaly": .TextMatrix(0, 7) = "Anomaly"
      End If
      
      .TextMatrix(0, 3 + iItems_stat) = "Rank(High-Low)"
      .TextMatrix(0, 4 + iItems_stat) = "Rank(Low-High)"
      
      .ColWidth(5 + iItems_stat) = 1800
      .TextMatrix(0, 5 + iItems_stat) = "Ranking Year"
      
      For i = 1 To iItems_rank - 3
        .ColWidth(i + 5 + iItems_stat) = 1000
        If i = 1 Then .TextMatrix(0, i + 5 + iItems_stat) = "MaxYear"
        If i = 2 Then .TextMatrix(0, i + 5 + iItems_stat) = "Max"
        If i = 3 Then .TextMatrix(0, i + 5 + iItems_stat) = "MinYear"
        If i = 4 Then .TextMatrix(0, i + 5 + iItems_stat) = "Min"
      Next i
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat + iItems_rank) = 1000
'        If i = 1 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "Township"
        If i = 1 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "County"
        If i = 2 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "City"
        If i = 3 Then .TextMatrix(0, i + 2 + iItems_stat + iItems_rank) = "Province"
      Next i
      
      
      If ShowLonLatExt = True Then
        .ColWidth(2 + iItems_stat + iItems_rank + iItems_zone + 1) = 1000: .ColWidth(2 + iItems_stat + iItems_rank + iItems_zone + 2) = 1000
        .TextMatrix(0, 2 + iItems_stat + iItems_rank + iItems_zone + 1) = "Longitude": .TextMatrix(0, 2 + iItems_stat + iItems_rank + iItems_zone + 2) = "Latitude"
      End If
      
      '******************填充站点信息信息********************
        .Rows = ExtStaNum + 1
        For i = 1 To ExtStaNum
          .TextMatrix(i, 0) = i
          .TextMatrix(i, 1) = ExtStaInfo2(i).stacode
          .TextMatrix(i, 2) = ExtStaInfo2(i).staname

'          .TextMatrix(i, 1 + 2 + iItems_stat + iItems_rank) = ExtStaInfo2(i).Town
          .TextMatrix(i, 1 + 2 + iItems_stat + iItems_rank) = ExtStaInfo2(i).County
          .TextMatrix(i, 2 + 2 + iItems_stat + iItems_rank) = ExtStaInfo2(i).City
          .TextMatrix(i, 3 + 2 + iItems_stat + iItems_rank) = ExtStaInfo2(i).Prov
                      
          If ShowLonLatExt = True Then
            .TextMatrix(i, 2 + iItems_stat + iItems_rank + iItems_zone + 1) = ExtStaInfo2(i).Longitude
            .TextMatrix(i, 2 + iItems_stat + iItems_rank + iItems_zone + 2) = ExtStaInfo2(i).Latitude
          End If
        Next i

      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Row <> 0 Then
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          End If
        Next j
      Next i
      
      .Refresh
      .Redraw = True
      End With
      

    '*********************初始化下方表格*******************************
      With Frm.HFGrid2
      .Redraw = False
      .Clear: .Refresh

      .Cols = 3 + iItems_stat + iItems_rank + iItems_zone '排名相关的7列、镇县市3列
      If ShowLonLatExt = True Then .Cols = .Cols + 2  '经度、纬度
      
      .FixedCols = 3
      .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
      
      For i = 1 To iItems_stat + 2 '排名（大-小）、排名（小-大）也设置宽1600
        .ColWidth(i + 2) = 1600
      Next i
      
      .ColWidth(5 + iItems_stat) = 1800 '排名年份设置宽1800
      
      For i = 1 To iItems_rank - 3 '最大值年、最大值、最小值年、最小值设宽1000
        .ColWidth(i + 5 + iItems_stat) = 1000
      Next i
      
      For i = 1 To 3
        .ColWidth(i + 2 + iItems_stat + iItems_rank) = 1000
      Next i
      
      If ShowLonLatExt = True Then
        .ColWidth(2 + iItems_stat + iItems_rank + 4) = 1000: .ColWidth(2 + iItems_stat + iItems_rank + 5) = 1000
      End If

      .Rows = 3
      .TextMatrix(0, 2) = "Average": .TextMatrix(1, 2) = "Max": .TextMatrix(2, 2) = "Min"
      '******上方单元格左对齐、前两列着色（白色）********
      For i = 0 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Col = 1 Then .CellBackColor = &H80000005
          If .Col = 2 Then .CellBackColor = &H80000005
        Next j
      Next i
      
      .Refresh
      .Redraw = True
    End With
    If FrmMeteoPeriodExt.HFGrid1.Rows > 2 Then FrmMeteoPeriodExt.HFGrid1.Row = 2

  Else '输出区域统计结果：设置HFGrid3、HFGrid4、HFGrid5三个表格
      
      '*********************循环初始化三个表格*******************************
      For k = 1 To 3
        With Frm.HFGrid3(k)
        .Redraw = False
        .Clear: .Refresh
      
        .Cols = 3 + iItems_stat + iItems_rank  '排名相关的7列
        
        If FlagZoneExt = "CITY" Then
          .Cols = .Cols + 1
        ElseIf FlagZoneExt = "CNTY" Then
          .Cols = .Cols + 2
        End If
        
        .FixedCols = 3
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        If k = 1 Then .TextMatrix(0, 0) = "Regional Avg"
        If k = 2 Then .TextMatrix(0, 0) = "Regional Max"
        If k = 3 Then .TextMatrix(0, 0) = "Regional Min"
        .TextMatrix(0, 1) = "Region Code": .TextMatrix(0, 2) = "Region Name"
          
        For i = 1 To iItems_stat + 2
          .ColWidth(i + 2) = 1600
        Next i
        If Frm.FlagStat = "AVE" Or Frm.FlagStat = "SUM" Then
          If Frm.FlagStat = "AVE" Then .TextMatrix(0, 3) = "Average": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "Yearly Avg"
          If Frm.FlagStat = "SUM" Then .TextMatrix(0, 3) = "Cumulative": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "Yearly Total"
          If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R08_20" Or SelField_Initial = "R20_08" Or SelField_Initial = "S" Then
           .TextMatrix(0, 5) = "Anomaly %": .TextMatrix(0, 7) = "Anomaly %"
          Else
            .TextMatrix(0, 5) = "Anomaly": .TextMatrix(0, 7) = "Anomaly"
          End If
          
        ElseIf Frm.FlagStat = "MAX" Or Frm.FlagStat = "MIN" Then
          If Frm.FlagStat = "MAX" Then .TextMatrix(0, 3) = "Max": .TextMatrix(0, 5) = "MaxRecord": .TextMatrix(0, 7) = "Yearly Max":
          If Frm.FlagStat = "MIN" Then .TextMatrix(0, 3) = "Min": .TextMatrix(0, 5) = "MinRecord": .TextMatrix(0, 7) = "Yearly Min":
          .TextMatrix(0, 4) = "OccurrDate": .TextMatrix(0, 6) = "OccurrDate": .TextMatrix(0, 8) = "OccurrDate"
        
        ElseIf Frm.FlagStat = "DAYS" Then
          .TextMatrix(0, 3) = "ConditionDays": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 6) = "YearlyDays"
          .TextMatrix(0, 5) = "Anomaly": .TextMatrix(0, 7) = "Anomaly"
        End If
         
        If k = 1 Then
          .TextMatrix(0, 3 + iItems_stat) = "Rank(High-Low)"
          .TextMatrix(0, 4 + iItems_stat) = "Rank(Low-High)"
            
          .ColWidth(5 + iItems_stat) = 1800
          .TextMatrix(0, 5 + iItems_stat) = "Ranking Year"
              
          For i = 1 To iItems_rank - 3
            .ColWidth(i + 5 + iItems_stat) = 1000
            If i = 1 Then .TextMatrix(0, i + 5 + iItems_stat) = "MaxYear"
            If i = 2 Then .TextMatrix(0, i + 5 + iItems_stat) = "Max"
            If i = 3 Then .TextMatrix(0, i + 5 + iItems_stat) = "MinYear"
            If i = 4 Then .TextMatrix(0, i + 5 + iItems_stat) = "Min"
          Next i
          
        Else
          .ColWidth(5 + iItems_stat) = 1800
          For i = 1 To iItems_rank - 3
            .ColWidth(i + 5 + iItems_stat) = 1000
          Next i
        End If
          
          
          If FlagZoneExt = "CITY" Then
            .TextMatrix(0, 3 + iItems_stat + iItems_rank) = "Province"
          ElseIf FlagZoneExt = "CNTY" Then
            .TextMatrix(0, 3 + iItems_stat + iItems_rank) = "City"
            .TextMatrix(0, 4 + iItems_stat + iItems_rank) = "Province"
         
          End If
          
  
        '******************填充区域信息********************
        If FlagZoneExt = "CITY" Then .Rows = ExtCityNum + 1
        If FlagZoneExt = "CNTY" Then .Rows = ExtCountyNum + 1
        If FlagZoneExt = "PROV" Then .Rows = ExtProvNum + 1
        For i = 1 To .Rows - 1
          .TextMatrix(i, 0) = i
            
          If FlagZoneExt = "CITY" Then
            .TextMatrix(i, 1) = ExtCityInfo2(i).Code
            .TextMatrix(i, 2) = ExtCityInfo2(i).City
            .TextMatrix(i, 3 + iItems_stat + iItems_rank) = ExtCityInfo2(i).Prov
          ElseIf FlagZoneExt = "CNTY" Then
            .TextMatrix(i, 1) = ExtCountyInfo2(i).Code
            .TextMatrix(i, 2) = ExtCountyInfo2(i).City & ExtCountyInfo2(i).County
            .TextMatrix(i, 3 + iItems_stat + iItems_rank) = ExtCountyInfo2(i).City
            .TextMatrix(i, 4 + iItems_stat + iItems_rank) = ExtCountyInfo2(i).Prov
          ElseIf FlagZoneExt = "PROV" Then
            .TextMatrix(i, 1) = ExtProvInfo2(i).Code
            .TextMatrix(i, 2) = ExtProvInfo2(i).Prov
          End If
        Next i
    
        '******上方单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
   
        .Refresh
        .Redraw = True
        End With
      
      Next k
  End If

          
End Sub



'2026-07-20 广东季节划分（起始日）
Public Sub TableIni_MeteoSeasons(Frm As Object)
  Dim i%
  Dim iItems_stat% '统计的项目数量
  Dim iItems_zone%
  
  iItems_zone = 3 '区域信息数据：镇乡、区县、地市
  
  iItems_stat = 5
  
   With Frm
    If FlagZone = "STA" Then '输出站点结果
      .HFGrid1.Visible = True
      .HFGrid2.Visible = True
      .HFGrid3(1).Visible = False
      .HFGrid3(2).Visible = False
      .HFGrid3(3).Visible = False
    Else '输出区域统计结果
      .HFGrid1.Visible = False
      .HFGrid2.Visible = False
      .HFGrid3(1).Visible = True
      .HFGrid3(2).Visible = True
      .HFGrid3(3).Visible = True
    End If
  End With
  
  
  If FlagZone = "STA" Then
      With Frm.HFGrid1
        .Redraw = False
        .Clear: .Refresh
    
        .Rows = StaNum + 1
    
        .Cols = 3 + iItems_stat + iItems_zone  '镇县市3列
        If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
        
        .FixedCols = 3
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        .ColWidth(3) = 1800: .ColWidth(4) = 1800: .ColWidth(5) = 1000: .ColWidth(6) = 1800: .ColWidth(7) = 1000
    
        .TextMatrix(0, 0) = "No.": .TextMatrix(0, 1) = "Station Code": .TextMatrix(0, 2) = "Station Name"
        .TextMatrix(0, 3) = "Onset Date": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 5) = "Anomaly"
        .TextMatrix(0, 6) = "Onset Date": .TextMatrix(0, 7) = "Anomaly"
          
        For i = 1 To 3
          .ColWidth(i + 2 + iItems_stat) = 1000
          If i = 1 Then .TextMatrix(0, i + 2 + iItems_stat) = "Township"
          If i = 2 Then .TextMatrix(0, i + 2 + iItems_stat) = "County"
          If i = 3 Then .TextMatrix(0, i + 2 + iItems_stat) = "City"
        Next i
        
        
        If ShowLonLat = True Then
          .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
          .TextMatrix(0, 2 + iItems_stat + 4) = "Longitude": .TextMatrix(0, 2 + iItems_stat + iItems_rank + 5) = "Latitude"
        End If
        
        '******************填充站点信息信息********************
        .Rows = StaNum + 1
        For i = 1 To StaNum
          .TextMatrix(i, 0) = i
          .TextMatrix(i, 1) = StaInfo(i).stacode
          .TextMatrix(i, 2) = StaInfo(i).staname
    
          .TextMatrix(i, 1 + 2 + iItems_stat) = StaInfo(i).Town
          .TextMatrix(i, 2 + 2 + iItems_stat) = StaInfo(i).County
          .TextMatrix(i, 3 + 2 + iItems_stat) = StaInfo(i).City
                      
          If ShowLonLat = True Then
            .TextMatrix(i, 2 + iItems_stat + 4) = StaInfo(i).Longitude
            .TextMatrix(i, 2 + iItems_stat + 5) = StaInfo(i).Latitude
          End If
        Next i
    
        '******上方单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
    
        .Refresh
        .Redraw = True
      End With
    
      With Frm.HFGrid2
        .Redraw = False
        .Clear: .Refresh
    
        .Rows = 3
    
        .Cols = 3 + iItems_stat + iItems_zone  '镇县市3列
        If ShowLonLat = True Then .Cols = .Cols + 2  '经度、纬度
        
        .FixedCols = 3
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        .ColWidth(3) = 1800: .ColWidth(4) = 1800: .ColWidth(5) = 1000: .ColWidth(6) = 1800: .ColWidth(7) = 1000
         
        For i = 1 To 3
          .ColWidth(i + 2 + iItems_stat) = 1000
        Next i
        
        If ShowLonLat = True Then
          .ColWidth(2 + iItems_stat + 4) = 1000: .ColWidth(2 + iItems_stat + 5) = 1000
        End If
    
        .TextMatrix(0, 2) = "Average": .TextMatrix(1, 2) = "Max": .TextMatrix(2, 2) = "Min"
    
        '******上方单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          Next j
        Next i
    
    
        .Refresh
        .Redraw = True
      End With
      
      
  Else '输出区域统计结果：设置HFGrid3、HFGrid4、HFGrid5三个表格
      
      '*********************循环初始化三个表格*******************************
      For k = 1 To 3
        With Frm.HFGrid3(k)
        .Redraw = False
        .Clear: .Refresh
      
        .Cols = 3 + iItems_stat  '排名相关的7列
        .FixedCols = 3
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        If k = 1 Then .TextMatrix(0, 0) = "Regional Avg"
        If k = 2 Then .TextMatrix(0, 0) = "Regional Max"
        If k = 3 Then .TextMatrix(0, 0) = "Regional Min"
        .TextMatrix(0, 1) = "Region Code": .TextMatrix(0, 2) = "Region Name"
    
        .ColWidth(0) = 1000: .ColWidth(1) = 1500: .ColWidth(2) = 2500
        .ColWidth(3) = 1800: .ColWidth(4) = 1800: .ColWidth(5) = 1000: .ColWidth(6) = 1800: .ColWidth(7) = 1000
        
        .TextMatrix(0, 3) = "Onset Date": .TextMatrix(0, 4) = "Normals": .TextMatrix(0, 5) = "Anomaly"
        .TextMatrix(0, 6) = "Onset Date": .TextMatrix(0, 7) = "Anomaly"
      
      
        '******************填充区域信息********************
        If FlagZone = "CITY" Then .Rows = CityNum + 1
        If FlagZone = "CNTY" Then .Rows = CountyNum + 1
        If FlagZone = "TOWN" Then .Rows = TownNum + 1
        For i = 1 To .Rows - 1
          .TextMatrix(i, 0) = i
            
          If FlagZone = "CITY" Then
            .TextMatrix(i, 1) = CityInfo2(i).Code
            .TextMatrix(i, 2) = CityInfo2(i).City
          ElseIf FlagZone = "CNTY" Then
            .TextMatrix(i, 1) = CountyInfo2(i).Code
            .TextMatrix(i, 2) = CountyInfo2(i).City & CountyInfo2(i).County
          ElseIf FlagZone = "TOWN" Then
            .TextMatrix(i, 1) = TownInfo2(i).Code
            .TextMatrix(i, 2) = TownInfo2(i).City & TownInfo2(i).County & TownInfo2(i).Town
          End If
        Next i
    
        '******上方单元格左对齐、前两列着色（白色）********
        For i = 0 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Row <> 0 Then
              If .Col = 1 Then .CellBackColor = &H80000005
              If .Col = 2 Then .CellBackColor = &H80000005
            End If
          Next j
        Next i
   
        .Refresh
        .Redraw = True
        End With
      
      Next k
  End If
      
      
End Sub
