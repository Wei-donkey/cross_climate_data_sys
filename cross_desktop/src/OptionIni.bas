Attribute VB_Name = "OptionIni"
    
'确定默认的统计量:任意时段（逐年）通用，返回结果至FlagStat中:2024-9-13
Public Sub Stat_MeteoPeriod(Frm As Object, SelField_Restat$, FlagStat$)
  With Frm
  
  '************************************************ 日间隔数据 ****************************************************
    If SelField_Restat = "FZS" Or SelField_Restat = "FJS" Or SelField_Restat = "F2S" Or SelField_Restat = "F10S" _
    Or SelField_Restat = "UX" Or SelField_Restat = "TX" Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX" Then
      .OptionStat(2).Value = True: FlagStat = "MAX"
    ElseIf SelField_Restat = "UN" Or SelField_Restat = "TN" Or SelField_Restat = "PN" Or SelField_Restat = "BN" Or SelField_Restat = "DN" Then
      .OptionStat(3).Value = True: FlagStat = "MIN"
    ElseIf SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Or SelField_Restat = "S" Or SelField_Restat = "L_B" Then
      .OptionStat(4).Value = True: FlagStat = "SUM"
    Else
      .OptionStat(1).Value = True: FlagStat = "AVE"
    End If
  
  End With

End Sub



'确定默认的统计量:旬、月、季、年通用，返回结果至FlagStat中
Public Sub Stat_Meteo(Frm As Object, SelField_Restat$, FlagStat$)
  With Frm
  
  
  '************************************************ 日间隔数据 ****************************************************
  If .ComboDataType.Text = .ComboDataType.List(0) Then
      If SelField_Restat = "FZS" Or SelField_Restat = "FJS" Or SelField_Restat = "F2S" Or SelField_Restat = "F10S" _
      Or SelField_Restat = "UX" Or SelField_Restat = "TX" Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX" Then
        .OptionStat(2).Value = True: FlagStat = "MAX"
      ElseIf SelField_Restat = "UN" Or SelField_Restat = "TN" Or SelField_Restat = "PN" Or SelField_Restat = "BN" Or SelField_Restat = "DN" Then
        .OptionStat(3).Value = True: FlagStat = "MIN"
      ElseIf SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Or SelField_Restat = "S" Or SelField_Restat = "L_B" Then
        .OptionStat(4).Value = True: FlagStat = "SUM"
      Else
        .OptionStat(1).Value = True: FlagStat = "AVE"
      End If
  
  '******************************************************* 平均态数据 **********************************************************
  '只能查平均值（除了雨量、蒸发量、日照）、累计值（只有平均气温、雨量、蒸发量、日照）、日数（积温日、高温、低温、雨日、阴天）、总量（积温、雨量）
  ElseIf .ComboDataType.Text = .ComboDataType.List(1) Then
      If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Or SelField_Restat = "S" Or SelField_Restat = "L_B" Then
        .OptionStat(4).Value = True: FlagStat = "SUM"
      Else
        .OptionStat(1).Value = True: FlagStat = "AVE"
      End If
  
  '******************************************************* 极端态数据 **********************************************************
  '平均值（除了雨量、蒸发量、日照）、累计值（只有平均气温、雨量、蒸发量、日照）、日数（积温日、高温、低温、雨日、阴天）、总量（积温、雨量）
  '最大值（最高气温、最高湿度、最大风速、极大风速、最高气压、最高草地温、最高裸地温）、最小值（最低气温、最小湿度、最低气压、最低能见度、最低草地温、最低裸地温）
  ElseIf .ComboDataType.Text = .ComboDataType.List(2) Then
    
      If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Or SelField_Restat = "S" Or SelField_Restat = "L_B" Then
        .OptionStat(4).Value = True: FlagStat = "SUM"
      ElseIf SelField_Restat = "TX" Or SelField_Restat = "UX" Or SelField_Restat = "FZS" Or SelField_Restat = "FJS" Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX" Then
        .OptionStat(2).Value = True: FlagStat = "MAX"
      ElseIf SelField_Restat = "TN" Or SelField_Restat = "UN" Or SelField_Restat = "PN" Or SelField_Restat = "VN" Or SelField_Restat = "BN" Or SelField_Restat = "DN" Then
        .OptionStat(3).Value = True: FlagStat = "MIN"
      Else
        .OptionStat(1).Value = True: FlagStat = "AVE"
      End If
  
  
  End If
  
  End With

End Sub


'****************************逐时模块****************************
Public Sub Option_MeteoHor(SelField_Hor$)
    With FrmMeteoHour
    .ComboMin_HL.Clear: .ComboMax_HL.Clear
    .ComboMin_HL.AddItem "-9999":  .ComboMax_HL.AddItem "999999"
    .ComboMin_HL.Text = "-9999": .ComboMax_HL.Text = "999999"
    
    If SelField_Hor = "R" Then  '雨量默认查累计值
      .OptionStat(4).Value = True
    Else
      .OptionStat(1).Value = True
    End If
    
    End With
End Sub


'****************************逐日模块****************************
Public Sub Option_MeteoDay(SelField_Initial$, SelField_Restat$)
    With FrmMeteoDay
    .ComboMin_HL.Clear: .ComboMax_HL.Clear
    .ComboMin_HL.AddItem "-9999":  .ComboMax_HL.AddItem "999999"
    .ComboMin_HL.Text = "-9999": .ComboMax_HL.Text = "999999"
    
    If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R08_20" Or SelField_Initial = "R20_08" Or SelField_Initial = "S" Then  '雨量和日照默认查累计值
      .OptionStat(4).Value = True
    Else
      .OptionStat(1).Value = True
    End If
    
    If .ComboDataType.Text = .ComboDataType.List(0) Then '日间隔数据
      .OptionStat(5).Enabled = False: .OptionStat(6).Enabled = False
      .CheckXtrmYear.Enabled = False: .CheckXtrmYear.Value = Unchecked
      
    ElseIf .ComboDataType.Text = .ComboDataType.List(1) Then '平均态数据
      .OptionStat(5).Enabled = False: .OptionStat(6).Enabled = False
      .CheckXtrmYear.Enabled = False: .CheckXtrmYear.Value = Unchecked
      .OptionStat(1).Value = True
    
    ElseIf .ComboDataType.Text = .ComboDataType.List(2) Then '极端态数据
      If FlagZone = "STA" Then .CheckXtrmYear.Enabled = True ': .CheckXtrmYear.Value = Checked
      If FlagZone <> "STA" Then .CheckXtrmYear.Enabled = False: .CheckXtrmYear.Value = Unchecked
      If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
      Or SelField_Restat = "F2S" Or SelField_Restat = "F10S" Or SelField_Restat = "FZS" Or SelField_Restat = "FJS" _
      Or SelField_Restat = "TX" Or SelField_Restat = "UX" Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX" Then  '雨量、最高气温只能查累年极大值
        .OptionStat(5).Enabled = True: .OptionStat(6).Enabled = False
        .OptionStat(5).Value = True
      ElseIf SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
      Or SelField_Restat = "F2S" Or SelField_Restat = "F10S" Or SelField_Restat = "FZS" Or SelField_Restat = "FJS" _
      Or SelField_Restat = "TN" Or SelField_Restat = "UN" Or SelField_Restat = "PN" Or SelField_Restat = "VN" Or SelField_Restat = "BN" Or SelField_Restat = "DN" Then   '最低气温只能查累年极小值
        .OptionStat(5).Enabled = False: .OptionStat(6).Enabled = True
        .OptionStat(6).Value = True
      Else
        .OptionStat(5).Enabled = True: .OptionStat(6).Enabled = True
      End If
      .OptionStat(1).Value = True

    End If
    
    End With
End Sub


'**************************** 逐旬月季年模块通用：确定哪些统计值可用/不可用 ****************************
Public Sub Option_Meteo(Frm As Object, SelField_Restat$)
    With Frm
      .ComboMin_HL.Clear: .ComboMax_HL.Clear
      .ComboMin_HL.AddItem "-9999":  .ComboMax_HL.AddItem "999999"
      .ComboMin_HL.Text = "-9999": .ComboMax_HL.Text = "999999"
    
    .ComboMin.Clear: .ComboMax.Clear
    If SelField_Restat = "TA" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "0": .ComboMin.AddItem "10": .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "TX" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "35": .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "TN" Then
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "5":  .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "S" Then
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "2":  .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "0.1":  .ComboMin.AddItem "10": .ComboMin.AddItem "25": .ComboMin.AddItem "50": .ComboMax.AddItem "999999"
    Else
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "999999"
    End If
    .ComboMin.Text = .ComboMin.List(0): .ComboMax.Text = .ComboMax.List(.ComboMax.ListCount - 1)
    
    
'    If Not (Frm Is FrmMeteoPeriod Or Frm Is FrmMeteoPeriodYer) Then  '旬月季年模块只能条件查询数据表中已有的字段
'        If Not (SelField_Restat = "TA" Or SelField_Restat = "TX" Or SelField_Restat = "TN" Or SelField_Restat = "R" Or SelField_Restat = "S") Then
'          .CheckLimit.Value = Unchecked
'          .CheckLimit.Enabled = False
'        Else
'          .CheckLimit.Enabled = True
'        End If
'    Else  '任意时段（逐年）模块可任意条件查询（2023-10-25）
'        .CheckLimit.Enabled = True
'    End If
    


    '************************************************ 常规间隔数据 ****************************************************
    If .ComboDataType.Text = .ComboDataType.List(0) Then
        
        
        '2024-9-19：平均值统计不可以条件查询
        If (.OptionStat(1).Value = True Or .OptionStat(2).Value = True Or .OptionStat(3).Value = True) Then
          .ComboMin.Enabled = False: .ComboMax.Enabled = False
          .ComboMin.Text = "-9999": .ComboMax.Text = "999999"
        Else
          .ComboMin.Enabled = True: .ComboMax.Enabled = True
          If (SelField_Restat = "S" And .OptionStat(4).Value = True) Then
            .ComboMin.Enabled = False: .ComboMax.Enabled = False
            .ComboMin.Text = "-9999": .ComboMax.Text = "999999"
          End If
        End If
            
        
        .OptionStat(5).Enabled = False: .OptionStat(6).Enabled = False
        .CheckXtrmYear.Enabled = False: .CheckXtrmYear.Value = Unchecked

        .OptionStat(1).Enabled = True: .OptionStat(2).Enabled = True: .OptionStat(3).Enabled = True: .OptionStat(4).Enabled = True: .OptionStat(7).Enabled = True

        '雨量、蒸发不可查平均值
        If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Or SelField_Restat = "L_B" Then
          .OptionStat(1).Enabled = False: .OptionStat(1).Value = False
        End If
               
        '最低温度、最低湿度、最低气压、最低地温、最低能见度不可查最大值
        If SelField_Restat = "UN" Or SelField_Restat = "TN" Or SelField_Restat = "PN" Or SelField_Restat = "BN" Or SelField_Restat = "DN" Or SelField_Restat = "VN" Then
          .OptionStat(2).Enabled = False: .OptionStat(2).Value = False
        End If
        
        '最高温度、最高气压、最高湿度、最高地温、日照、雨量、风速不可查最小值
        If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
        Or SelField_Restat = "FZS" Or SelField_Restat = "FJS" Or SelField_Restat = "F2S" Or SelField_Restat = "F10S" Or SelField_Restat = "S" _
        Or SelField_Restat = "UX" Or SelField_Restat = "TX" Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX" Then
          .OptionStat(3).Enabled = False: .OptionStat(3).Value = False
        End If

        '只有平均气温、日照、雨量、蒸发量可查累计值
        If Not (SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
        Or SelField_Restat = "L_B" Or SelField_Restat = "S" Or SelField_Restat = "TA") Then
          .OptionStat(4).Enabled = False: .OptionStat(4).Value = False
        End If

        '只有气温、日照、雨量可查日数
        If Not (SelField_Restat = "R" Or SelField_Restat = "TX" Or SelField_Restat = "TN" _
        Or SelField_Restat = "S" Or SelField_Restat = "TA") Then
          .OptionStat(7).Enabled = False: .OptionStat(7).Value = False
        End If

    
    '******************************************************* 平均态数据 **********************************************************
    '只能查平均值（除了雨量、蒸发量、日照）、累计值（只有平均气温、雨量、蒸发量、日照）、日数（积温日、高温、低温、雨日、阴天）、总量（积温、雨量）
    ElseIf .ComboDataType.Text = .ComboDataType.List(1) Then
        
        '2024-9-19：平均值统计不可以条件查询
        If .OptionStat(1).Value = True Then
          .ComboMin.Enabled = False: .ComboMax.Enabled = False
          .ComboMin.Text = "-9999": .ComboMax.Text = "999999"
        Else
          .ComboMin.Enabled = True: .ComboMax.Enabled = True
        End If
                    
        
        
        .OptionStat(5).Enabled = False: .OptionStat(6).Enabled = False
        .CheckXtrmYear.Enabled = False: .CheckXtrmYear.Value = Unchecked
        
        '所有要素都不可查询最大值/最小值
        .OptionStat(1).Enabled = True: .OptionStat(2).Enabled = False: .OptionStat(3).Enabled = False: .OptionStat(4).Enabled = True: .OptionStat(7).Enabled = True
          
        '雨量、蒸发不可查平均值
        If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Or SelField_Restat = "L_B" Then
          .OptionStat(1).Enabled = False: .OptionStat(1).Value = False
        End If
          
        '只有平均气温、日照、雨量、蒸发量可查累计值
        If Not (SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
        Or SelField_Restat = "L_B" Or SelField_Restat = "S" Or SelField_Restat = "TA") Then
          .OptionStat(4).Enabled = False: .OptionStat(4).Value = False
        End If
    
        '只有气温、日照、雨量可查日数
        If Not (SelField_Restat = "R" Or SelField_Restat = "TX" Or SelField_Restat = "TN" _
        Or SelField_Restat = "S" Or SelField_Restat = "TA") Then
          .OptionStat(7).Enabled = False: .OptionStat(7).Value = False
        End If
          

    '******************************************************* 极端态数据 **********************************************************
    '查平均值（除了雨量、蒸发量、日照）、累计值（只有平均气温、雨量、蒸发量、日照）、日数（积温日、高温、低温、雨日、阴天）、总量（积温、雨量）
    '最大值（最高气温、最高湿度、最大风速、极大风速、最高气压、最高草地温、最高裸地温）、最小值（最低气温、最小湿度、最低气压、最低能见度、最低草地温、最低裸地温）
    ElseIf .ComboDataType.Text = .ComboDataType.List(2) Then
      
        '2024-9-19：平均值统计不可以条件查询
        If .OptionStat(1).Value = True Then
          .ComboMin.Enabled = False: .ComboMax.Enabled = False
          .ComboMin.Text = "-9999": .ComboMax.Text = "999999"
        Else
          .ComboMin.Enabled = True: .ComboMax.Enabled = True
        End If
      
      If FlagZone = "STA" Then .CheckXtrmYear.Enabled = True ': .CheckXtrmYear.Value = Checked
      If FlagZone <> "STA" Then .CheckXtrmYear.Enabled = False: .CheckXtrmYear.Value = Unchecked

      .OptionStat(5).Enabled = True: .OptionStat(6).Enabled = True
      If (SelField_Restat = "TN" Or SelField_Restat = "UN" Or SelField_Restat = "PN" Or SelField_Restat = "VN" Or SelField_Restat = "BN" Or SelField_Restat = "DN") Then
        .OptionStat(5).Enabled = False: .OptionStat(6).Enabled = True:  .OptionStat(6).Value = True
        
      ElseIf (SelField_Restat = "TX" Or SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R20_08" Or SelField_Restat = "R08_20" Or SelField_Restat = "L_B" _
             Or SelField_Restat = "FZS" Or SelField_Restat = "FJS" Or SelField_Restat = "S" Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX") Then
        .OptionStat(5).Enabled = True: .OptionStat(6).Enabled = False:  .OptionStat(5).Value = True
      Else
        .OptionStat(5).Enabled = True: .OptionStat(6).Enabled = True:  .OptionStat(5).Value = True
      End If
      
    
      '默认所有要素所有统计量都可查询，下面再另行判断
      .OptionStat(1).Enabled = True: .OptionStat(2).Enabled = True: .OptionStat(3).Enabled = True: .OptionStat(4).Enabled = True: .OptionStat(7).Enabled = True
      
      '只有平均气温、气温日较差、相对湿度、2min/10min平均风速、平均气压、平均能见度、平均草温、地温、各层地温、海平面气压、湿球温度、水汽压、露点温度可选平均值，
      If Not (SelField_Restat = "TA" Or SelField_Restat = "TR" Or SelField_Restat = "UA" Or SelField_Restat = "F2S" Or SelField_Restat = "F10S" Or SelField_Restat = "PA" Or SelField_Restat = "VA" _
      Or SelField_Restat = "BA" Or SelField_Restat = "DA" Or SelField_Restat = "D5" Or SelField_Restat = "D10" Or SelField_Restat = "D15" Or SelField_Restat = "D20" Or SelField_Restat = "D40" _
      Or SelField_Restat = "D80" Or SelField_Restat = "D160" Or SelField_Restat = "D320" Or SelField_Restat = "P0" Or SelField_Restat = "I" Or SelField_Restat = "E" Or SelField_Restat = "TD") Then
        .OptionStat(1).Enabled = False: .OptionStat(1).Value = False
      End If
               
      '只有平均气温、日照、雨量、蒸发量可查累计值
      If Not (SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
      Or SelField_Restat = "L_B" Or SelField_Restat = "S" Or SelField_Restat = "TA") Then
        .OptionStat(4).Enabled = False: .OptionStat(4).Value = False
      End If

      '只有最高气温、最高湿度、最大风速、极大风速、最高气压、最高草地温、最高裸地温可查最大值
      If Not (SelField_Restat = "TX" Or SelField_Restat = "UX" Or SelField_Restat = "FZS" Or SelField_Restat = "FJS" _
      Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX") Then
        .OptionStat(2).Enabled = False: .OptionStat(2).Value = False
      End If

      '只有最高气温、最高湿度、最大风速、极大风速、最高气压、最高草地温、最高裸地温可查最大值
      If Not (SelField_Restat = "TN" Or SelField_Restat = "UN" Or SelField_Restat = "PN" _
      Or SelField_Restat = "VN" Or SelField_Restat = "BN" Or SelField_Restat = "DN") Then
        .OptionStat(3).Enabled = False: .OptionStat(3).Value = False
      End If

    
        '只有气温、日照、雨量可查日数
        If Not (SelField_Restat = "R" Or SelField_Restat = "TX" Or SelField_Restat = "TN" _
        Or SelField_Restat = "S" Or SelField_Restat = "TA") Then
          .OptionStat(7).Enabled = False: .OptionStat(7).Value = False
        End If
    
    End If

    
    End With
End Sub



'****************************任意时段（逐年）模块通用****************************
Public Sub Option_MeteoPeriod(Frm As Object, SelField_Restat$)
    With Frm
   
    .ComboMin.Clear: .ComboMax.Clear
    If SelField_Restat = "TA" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "0": .ComboMin.AddItem "10": .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "TX" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "35": .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "TN" Then
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "5":  .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "S" Then
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "2":  .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "0.1":  .ComboMin.AddItem "10": .ComboMin.AddItem "25": .ComboMin.AddItem "50": .ComboMax.AddItem "999999"
    Else
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "999999"
    End If
    .ComboMin.Text = .ComboMin.List(0): .ComboMax.Text = .ComboMax.List(.ComboMax.ListCount - 1)
    
      
        .OptionStat(1).Enabled = True: .OptionStat(2).Enabled = True: .OptionStat(3).Enabled = True: .OptionStat(4).Enabled = True: .OptionStat(7).Enabled = True
       
'        '雨量、蒸发不可查平均值
'        If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Or SelField_Restat = "L_B" Then
''            .OptionStat(1).Enabled = False: .OptionStat(1).Value = False
'        End If
        
        '最低温度、最低湿度、最低气压、最低地温、最低能见度不可查最大值
        If SelField_Restat = "UN" Or SelField_Restat = "TN" Or SelField_Restat = "PN" Or SelField_Restat = "BN" Or SelField_Restat = "DN" Or SelField_Restat = "VN" Then
          .OptionStat(2).Enabled = False: .OptionStat(2).Value = False
        End If
        
        '最高温度、最高气压、最高湿度、最高地温、日照、雨量、风速不可查最小值
        If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
        Or SelField_Restat = "FZS" Or SelField_Restat = "FJS" Or SelField_Restat = "F2S" Or SelField_Restat = "F10S" Or SelField_Restat = "S" _
        Or SelField_Restat = "UX" Or SelField_Restat = "TX" Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX" Then
          .OptionStat(3).Enabled = False: .OptionStat(3).Value = False
        End If
        
'        '只有平均气温、日照、雨量、蒸发量可查累计值
'        If Not (SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
'        Or SelField_Restat = "L_B" Or SelField_Restat = "S" Or SelField_Restat = "TA") Then
''            .OptionStat(4).Enabled = False: .OptionStat(4).Value = False
'        End If
    End With
End Sub
  

'**************************** 单时段模块****************************
Public Sub Option_MeteoPeriodSin(Frm As Object, SelField_Restat$)
    With Frm
    
    .ComboMin.Clear: .ComboMax.Clear
    If SelField_Restat = "TA" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "0": .ComboMin.AddItem "10": .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "TX" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "35": .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "TN" Then
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "5":  .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "S" Then
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "2":  .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "0.1":  .ComboMin.AddItem "10": .ComboMin.AddItem "25": .ComboMin.AddItem "50": .ComboMax.AddItem "999999"
    Else
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "999999"
    End If
    .ComboMin.Text = .ComboMin.List(0): .ComboMax.Text = .ComboMax.List(.ComboMax.ListCount - 1)
    
    End With
End Sub
  


'**************************** 条件查询模块专用 ****************************
Public Sub Option_ConditionQuery(Frm As Object, SelField_Restat$)
    With Frm
    
    
    .ComboMin.Clear: .ComboMax.Clear
    If SelField_Restat = "TA" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "0": .ComboMin.AddItem "10": .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "TX" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "35": .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "TN" Then
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "5":  .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "S" Then
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "2":  .ComboMax.AddItem "999999"
    ElseIf SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Then
      .ComboMin.AddItem "-9999": .ComboMin.AddItem "0.1":  .ComboMin.AddItem "10": .ComboMin.AddItem "25": .ComboMin.AddItem "50": .ComboMax.AddItem "999999"
    Else
      .ComboMin.AddItem "-9999":  .ComboMax.AddItem "999999"
    End If
    .ComboMin.Text = .ComboMin.List(0): .ComboMax.Text = .ComboMax.List(.ComboMax.ListCount - 1)
    
    
    '************************************************ 常规间隔数据 ****************************************************
        
    '2024-9-19：平均值统计不可以条件查询
    If (.OptionStat(1).Value = True Or .OptionStat(2).Value = True Or .OptionStat(3).Value = True) Then
      .ComboMin.Enabled = False: .ComboMax.Enabled = False
      .ComboMin.Text = "-9999": .ComboMax.Text = "999999"
    Else
      .ComboMin.Enabled = True: .ComboMax.Enabled = True
      If (SelField_Restat = "S" And .OptionStat(4).Value = True) Then
        .ComboMin.Enabled = False: .ComboMax.Enabled = False
        .ComboMin.Text = "-9999": .ComboMax.Text = "999999"
      End If
    End If


    .OptionStat(1).Enabled = True: .OptionStat(2).Enabled = True: .OptionStat(3).Enabled = True: .OptionStat(4).Enabled = True: .OptionStat(7).Enabled = True

    '雨量、蒸发不可查平均值
    If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Or SelField_Restat = "L_B" Then
      .OptionStat(1).Enabled = False: .OptionStat(1).Value = False
    End If
           
    '最低温度、最低湿度、最低气压、最低地温、最低能见度不可查最大值
    If SelField_Restat = "UN" Or SelField_Restat = "TN" Or SelField_Restat = "PN" Or SelField_Restat = "BN" Or SelField_Restat = "DN" Or SelField_Restat = "VN" Then
      .OptionStat(2).Enabled = False: .OptionStat(2).Value = False
    End If
    
    '最高温度、最高气压、最高湿度、最高地温、日照、雨量、风速不可查最小值
    If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
    Or SelField_Restat = "FZS" Or SelField_Restat = "FJS" Or SelField_Restat = "F2S" Or SelField_Restat = "F10S" Or SelField_Restat = "S" _
    Or SelField_Restat = "UX" Or SelField_Restat = "TX" Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX" Then
      .OptionStat(3).Enabled = False: .OptionStat(3).Value = False
    End If

    '只有平均气温、日照、雨量、蒸发量可查累计值
    If Not (SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" _
    Or SelField_Restat = "L_B" Or SelField_Restat = "S" Or SelField_Restat = "TA") Then
      .OptionStat(4).Enabled = False: .OptionStat(4).Value = False
    End If

    '只有气温、日照、雨量可查日数
    If Not (SelField_Restat = "R" Or SelField_Restat = "TX" Or SelField_Restat = "TN" _
    Or SelField_Restat = "S" Or SelField_Restat = "TA") Then
      .OptionStat(7).Enabled = False: .OptionStat(7).Value = False
    End If
    
    
    End With
End Sub



'确定默认的统计量:条件查询模块专用，返回结果至FlagStat中
Public Sub Stat_ConditionQuery(Frm As Object, SelField_Restat$, FlagStat$)
  With Frm
  
  '************************************************ 日间隔数据 ****************************************************
  If SelField_Restat = "FZS" Or SelField_Restat = "FJS" Or SelField_Restat = "F2S" Or SelField_Restat = "F10S" _
  Or SelField_Restat = "UX" Or SelField_Restat = "TX" Or SelField_Restat = "PX" Or SelField_Restat = "BX" Or SelField_Restat = "DX" Then
    .OptionStat(2).Value = True: FlagStat = "MAX"
  ElseIf SelField_Restat = "UN" Or SelField_Restat = "TN" Or SelField_Restat = "PN" Or SelField_Restat = "BN" Or SelField_Restat = "DN" Then
    .OptionStat(3).Value = True: FlagStat = "MIN"
  ElseIf SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R08_20" Or SelField_Restat = "R20_08" Or SelField_Restat = "S" Or SelField_Restat = "L_B" Then
    .OptionStat(4).Value = True: FlagStat = "SUM"
  Else
    .OptionStat(1).Value = True: FlagStat = "AVE"
  End If
   
  End With

End Sub


