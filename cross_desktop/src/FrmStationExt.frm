VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmStationExt 
   BackColor       =   &H80000005&
   Caption         =   "Regional - Station Selection"
   ClientHeight    =   11850
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   17160
   Icon            =   "FrmStationExt.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   11850
   ScaleWidth      =   17160
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   22
      Top             =   11595
      Width           =   17160
      _ExtentX        =   30268
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
      EndProperty
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Province Filter"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   120
      TabIndex        =   19
      Top             =   120
      Width           =   2500
      Begin VB.ListBox List_Prov 
         Appearance      =   0  'Flat
         Columns         =   3
         Height          =   930
         Left            =   0
         MultiSelect     =   1  'Simple
         TabIndex        =   20
         Top             =   240
         Width           =   2500
      End
   End
   Begin VB.CommandButton CmdAll 
      Appearance      =   0  'Flat
      Caption         =   "Select All"
      Height          =   400
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   18
      Top             =   1080
      Width           =   1680
   End
   Begin VB.CommandButton CmdClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
      Height          =   400
      Left            =   13200
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   1080
      Width           =   840
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh"
      Height          =   400
      Left            =   14160
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   1080
      Width           =   840
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   400
      Left            =   15120
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1080
      Width           =   840
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Result Spatial Scale"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   11400
      TabIndex        =   11
      Top             =   120
      Width           =   3855
      Begin VB.OptionButton OptionScale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Province"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   2760
         TabIndex        =   21
         Top             =   360
         Width           =   1040
      End
      Begin VB.OptionButton OptionScale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Station"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   14
         Top             =   360
         Width           =   975
      End
      Begin VB.OptionButton OptionScale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "County"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   1200
         TabIndex        =   13
         Top             =   360
         Width           =   860
      End
      Begin VB.OptionButton OptionScale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "City"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   2040
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "City Filter"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   2760
      TabIndex        =   9
      Top             =   120
      Width           =   4900
      Begin VB.ListBox List_City 
         Appearance      =   0  'Flat
         Columns         =   4
         Height          =   930
         Left            =   0
         MultiSelect     =   1  'Simple
         TabIndex        =   10
         Top             =   240
         Width           =   4900
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "National Station List"
      ForeColor       =   &H80000008&
      Height          =   9975
      Left            =   120
      TabIndex        =   5
      Top             =   1560
      Width           =   16815
      Begin VB.ListBox List_StaSURF 
         Appearance      =   0  'Flat
         Columns         =   5
         Height          =   9570
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   8
         Top             =   240
         Width           =   15615
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save List"
         Height          =   400
         Left            =   15840
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   840
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import List"
         Height          =   400
         Left            =   15840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   720
         Width           =   840
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Output"
      Height          =   735
      Left            =   15360
      TabIndex        =   3
      Top             =   120
      Width           =   1575
      Begin VB.CheckBox CheckLonLat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Lat/Lon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   360
         Width           =   1400
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "County Filter"
      ForeColor       =   &H80000008&
      Height          =   1335
      Left            =   7800
      TabIndex        =   1
      Top             =   120
      Width           =   3500
      Begin VB.ListBox List_County 
         Appearance      =   0  'Flat
         Columns         =   2
         Height          =   930
         Left            =   0
         MultiSelect     =   1  'Simple
         TabIndex        =   2
         Top             =   240
         Width           =   3500
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   16080
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   1080
      Width           =   840
   End
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   10560
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialogOpen 
      Left            =   11040
      Top             =   1200
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmStationExt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim SelectSURFExt_tmp As Single '进入模块时，将 selectSURFExt 存入临时变量，最后再读回

Dim SelectedProvs$ ' 2026-8-10 记录查询的镇街
Dim SelectedCities$ '2026-8-10 用户选择的城市
Dim SelectedCounties$  '2026-8-10 记录查询的区县（默认为当前城市所是区县）
Dim tempZone$ '2026-8-10 存放临时地区名

Dim Num_surf% '已选的国家站

Private IgnoreListSURF As Boolean
Private IgnoreListProvs As Boolean '不触发省份列表click过程


Private Sub Show_StaNum()
  Dim i%
 
  Num_surf = List_StaSURF.SelCount
  StatusBar1.Panels(1).Text = "Selected: " & Num_surf & " national stations"
End Sub




Private Sub Form_Load()
  Dim i%, j%
  Dim lngStyle As Long

  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = 11 * (Screen.Height - Me.Height) / 20

  SelectSURFExt_tmp = SelectSURFExt
  IgnoreListSURF = True

  If FlagZoneExt = "" Then FlagZoneExt = "STA"
  If FlagZoneExt = "STA" Then OptionScale(0).Value = True
  If FlagZoneExt = "PROV" Then OptionScale(1).Value = True
  If FlagZoneExt = "CNTY" Then OptionScale(2).Value = True
  If FlagZoneExt = "CITY" Then OptionScale(3).Value = True

  If ShowLonLatExt = True Then CheckLonLat.Value = Checked
  If ShowLonLatExt = False Then CheckLonLat.Value = Unchecked


  ' 如果ExtProvInfo2 被清空（Erase）了
  If (Not ExtProvInfo2) = True Then 'ExtProvInfo2被 erase了
    ExtProvNum = 0
  Else
    ExtProvNum = UBound(ExtProvInfo2)
  End If
  
  For i = 1 To UBound(ExtProvInfo1)
    List_Prov.AddItem ExtProvInfo1(i).Prov
    If ExtProvNum <> 0 Then 'ExtProvNum = 0时，不判断省份列表是否该选择
      For j = 1 To UBound(ExtProvInfo2)
        If ExtProvInfo1(i).Prov = ExtProvInfo2(j).Prov Then
          List_Prov.Selected(i - 1) = True: Exit For
        End If
      Next j
    End If
  Next i
  
  For i = 1 To UBound(ExtStaInfo1)
    List_StaSURF.AddItem ExtStaInfo1(i).stacode & " " & ExtStaInfo1(i).staname
    If ExtStaInfo1(i).Selected = True Then List_StaSURF.Selected(i - 1) = True
  Next i
    
  IgnoreListSURF = False
  SelectSURFExt = SelectSURFExt_tmp
  
  Call Show_StaNum

End Sub


Private Sub CheckLonLat_Click()
  If CheckLonLat.Value = Checked Then ShowLonLat = True
  If CheckLonLat.Value = Unchecked Then ShowLonLat = False
End Sub

Private Sub CmdCancel_Click()
  Unload Me
End Sub


Private Sub List_prov_Click()
  Dim i%, j%
  If IgnoreListProvs = True Then Exit Sub
  
  List_City.Clear
  For i = 1 To UBound(ExtCityInfo1)
    If ExtCityInfo1(i).Prov = List_Prov.Text Then List_City.AddItem ExtCityInfo1(i).City
  Next i

  For i = 0 To List_Prov.ListCount - 1
    If List_Prov.Selected(i) = True Then 'Select
        
      For j = 0 To List_StaSURF.ListCount - 1
        If ExtStaInfo1(j + 1).Prov = List_Prov.List(i) Then
          List_StaSURF.Selected(j) = True: ExtStaInfo1(j + 1).Selected = True
        End If
      Next j
        
    ElseIf List_Prov.Selected(i) = False Then  '不选择
      For j = 0 To List_StaSURF.ListCount - 1
        If ExtStaInfo1(j + 1).Prov = List_Prov.List(i) Then
          List_StaSURF.Selected(j) = False: ExtStaInfo1(j + 1).Selected = False
        End If
      Next j

    End If
  Next i

  '点击“全选”按钮的效果与选择（多个）城市是一样的，所以这里让选择状态变成 1：
  SelectSURFExt = 1
  
  Call Show_StaNum

End Sub


Private Sub List_city_Click()
  Dim i%, j%
  List_County.Clear
  For i = 1 To UBound(ExtCountyInfo1)
    If ExtCountyInfo1(i).City = List_City.Text Then List_County.AddItem ExtCountyInfo1(i).County
  Next i
  
  For i = 0 To List_City.ListCount - 1
''''    If List_地市.List(i) = "" Then Go送至 Next地市 '如果区县名称为空，不要跳过这个区县，免得与其相关的站点没是得到是效判断
    If List_City.Selected(i) = True Then 'Select
           
      For j = 0 To List_StaSURF.ListCount - 1
        If ExtStaInfo1(j + 1).City = List_City.List(i) Then
          List_StaSURF.Selected(j) = True: ExtStaInfo1(j + 1).Selected = True
        End If
      Next j
            
    ElseIf List_City.Selected(i) = False Then  '不选择
      For j = 0 To List_StaSURF.ListCount - 1
        If ExtStaInfo1(j + 1).City = List_City.List(i) Then
          List_StaSURF.Selected(j) = False: ExtStaInfo1(j + 1).Selected = False
        End If
      Next j
    End If
      
''''NextCity:
  Next i

  '选择（多个）县区让选择状态变成 0.5也无妨（因为站点少，不影响写入临时表）：
  SelectSURFExt = 0.5

  Call Show_StaNum

End Sub


Private Sub List_county_Click()
  Dim i%, j%, k%
  For i = 0 To List_County.ListCount - 1
    If List_County.Selected(i) = True Then 'Select
        
      For j = 0 To List_StaSURF.ListCount - 1
        If ExtStaInfo1(j + 1).County = List_County.List(i) Then
          For k = 0 To List_City.ListCount - 1 '判断该站点是否所属显示的县区
            If ExtStaInfo1(j + 1).City = List_City.List(k) Then
              List_StaSURF.Selected(j) = True: ExtStaInfo1(j + 1).Selected = True
            End If
          Next k
        End If
      Next j
      
    ElseIf List_County.Selected(i) = False Then  '不选择
      For j = 0 To List_StaSURF.ListCount - 1
        If ExtStaInfo1(j + 1).County = List_County.List(i) Then
          List_StaSURF.Selected(j) = False: ExtStaInfo1(j + 1).Selected = False
        End If
      Next j
    End If
      
  Next i

  '选择（多个）县区让选择状态变成 0.5也无妨（因为站点少，不影响写入临时表）：
  SelectSURFExt = 0.5

  Call Show_StaNum

End Sub


Private Sub List_StaSURF_Click()
  If IgnoreListSURF Then Exit Sub
  SelectSURFExt = 0.5
End Sub


Private Sub CmdRefresh_Click()
  Call Show_StaNum
End Sub


Private Sub CmdAll_Click()
  Dim i%
  For i = 0 To List_Prov.ListCount - 1
    List_Prov.Selected(i) = True
  Next i
  
  For i = 0 To List_StaSURF.ListCount - 1
      List_StaSURF.Selected(i) = True
  Next i
  
  SelectSURFExt = 1

  Call Show_StaNum
End Sub


Private Sub CmdClear_Click()
  Dim i%
  
  '2026-07-19 用 LB_SETSEL 一次性清空多选ListBox的选中项，替代逐项循环，加快"清空"按钮响应速度
  Call SendMessage(List_City.hWnd, LB_SETSEL, 0, -1)
  Call SendMessage(List_County.hWnd, LB_SETSEL, 0, -1)
  Call SendMessage(List_Prov.hWnd, LB_SETSEL, 0, -1)
  Call SendMessage(List_StaSURF.hWnd, LB_SETSEL, 0, -1)

  Erase ExtProvInfo2: Erase ExtCityInfo2: Erase ExtCountyInfo2
  
  SelectSURFExt = 0
  
  Call Show_StaNum
End Sub


Private Sub CmdOK_Click()
  Dim j%, i%
  Dim tempStr$()
  
  If List_StaSURF.SelCount = 0 Then
    MsgBox "Select at least one national station", vbInformation, "Notice"
    Exit Sub
  End If
  
  If OptionScale(0).Value = True Then FlagZoneExt = "STA"
  If OptionScale(1).Value = True Then FlagZoneExt = "PROV"
  If OptionScale(2).Value = True Then FlagZoneExt = "CNTY"
  If OptionScale(3).Value = True Then FlagZoneExt = "CITY"
  
  
  If FlagZoneExt <> "STA" And List_Prov.SelCount = 0 Then
    MsgBox "Not station-based statistics -- select at least one province", vbInformation, "Notice"
    Exit Sub
  End If
  
  Screen.MousePointer = 11
  
  
  Call Show_StaNum
  
  '基于省、县、市统计，则获取省、县、市信息
  '  If OptionScale(0).Value = False Then
  ' 2023-3-9 重新配置用户所选的数据区域DataZone
  ' 2026-8-6 配置用户所选的 SelectedProvs
  SelectedProvs = ""
  If List_Prov.SelCount = 1 Then
    For j = 0 To List_Prov.ListCount - 1
      If List_Prov.Selected(j) = True Then
        SelectedProvs = "" & List_Prov.List(j) & "" '只是一个省份的时候，一定不要加引号！
      End If
    Next j

  ElseIf List_Prov.SelCount >= 2 Then
    For j = 0 To List_Prov.ListCount - 1
      If List_Prov.Selected(j) = True Then
        If SelectedProvs = "" Then
          SelectedProvs = "'" & List_Prov.List(j) & "'"
        ElseIf SelectedProvs <> "" Then
          SelectedProvs = SelectedProvs & ",'" & List_Prov.List(j) & "'"
        End If
      End If
    Next j
  End If

  '2025-06-05：如果个别城市被选择
  SelectedCities = ""
  If List_City.SelCount = 1 Then
    For j = 0 To List_City.ListCount - 1
      If List_City.Selected(j) = True Then
        SelectedCities = "" & List_City.List(j) & "" '只是一个城市的时候，一定不要加引号！
      End If
    Next j

  ElseIf List_City.SelCount >= 2 Then
    For j = 0 To List_City.ListCount - 1
      If List_City.Selected(j) = True Then
        If SelectedCities = "" Then
          SelectedCities = "'" & List_City.List(j) & "'"
        ElseIf SelectedCities <> "" Then
          SelectedCities = SelectedCities & ",'" & List_City.List(j) & "'"
        End If
      End If
    Next j
  End If

  '2026-08-09：如果个别区县被选择
  SelectedCounties = ""
  If List_County.SelCount = 1 Then
    For j = 0 To List_County.ListCount - 1
      If List_County.Selected(j) = True Then
        SelectedCounties = "" & List_County.List(j) & "" '只是一个县区的时候，一定不要加引号！
      End If
    Next j

  ElseIf List_County.SelCount >= 2 Then
    For j = 0 To List_County.ListCount - 1
      If List_County.Selected(j) = True Then
        If SelectedCounties = "" Then
          SelectedCounties = "'" & List_County.List(j) & "'"
        ElseIf SelectedCounties <> "" Then
          SelectedCounties = SelectedCounties & ",'" & List_County.List(j) & "'"
        End If
      End If
    Next j
  End If

  If Not GuestMode Then
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    ORAConn2.Open
    ORAConn2.CursorLocation = 3

    If SelectedProvs <> "" Then
      Call GetExtCountyInfo2(SelectedProvs, SelectedCities, SelectedCounties)
      Call GetExtCityInfo2(SelectedProvs, SelectedCities)
      Call GetExtProvInfo2(SelectedProvs)
    End If

    ORAConn4.Close
    ORAConn2.Close
  End If


  If (Not ExtCountyInfo2) = True Then 'ext区县Info2被 erase了
    ExtCountyNum = 0
  Else
    ExtCountyNum = UBound(ExtCountyInfo2)
  End If

  If (Not ExtCityInfo2) = True Then 'ext地市Info2被 erase了
    ExtCityNum = 0
  Else
    ExtCityNum = UBound(ExtCityInfo2)
  End If

  If (Not ExtProvInfo2) = True Then 'ExtProvInfo2 被 erase了
    ExtProvNum = 0
  Else
    ExtProvNum = UBound(ExtProvInfo2)
  End If

  ExtStaNum = List_StaSURF.SelCount

  '************** 1. 把已选国家站站点存入站点信息变量 **************
  If ExtStaNum = 0 Then
    For i = 0 To List_StaSURF.ListCount - 1
      ExtStaInfo1(i + 1).Selected = False
    Next i

  Else
    ReDim ExtStaInfo2(1 To ExtStaNum)
    j = 1
    For i = 0 To List_StaSURF.ListCount - 1
      If List_StaSURF.Selected(i) = True Then
        ExtStaInfo1(i + 1).Selected = True
        ExtStaInfo2(j).Prov = ExtStaInfo1(i + 1).Prov ': ExtStaInfo2(j).Region = ExtStaInfo1(i + 1).Region
        ExtStaInfo2(j).City = ExtStaInfo1(i + 1).City: ExtStaInfo2(j).County = ExtStaInfo1(i + 1).County
        'ExtStaInfo2(j).Town = ExtStaInfo1(i + 1).Town
        ExtStaInfo2(j).StaType = "SURF"
        ExtStaInfo2(j).stacode = ExtStaInfo1(i + 1).stacode: ExtStaInfo2(j).staname = ExtStaInfo1(i + 1).staname
        ExtStaInfo2(j).Longitude = ExtStaInfo1(i + 1).Longitude: ExtStaInfo2(j).Latitude = ExtStaInfo1(i + 1).Latitude
        ExtStaInfo2(j).DateSTT = ExtStaInfo1(i + 1).DateSTT
        ExtStaInfo2(j).YearSTT = Year(CDate(ExtStaInfo1(i + 1).DateSTT))
        ExtStaInfo2(j).Inland = 1
        ExtStaInfo2(j).Hydro = 0

        j = j + 1
      Else
        ExtStaInfo1(i + 1).Selected = False
      End If
    Next i

  End If

    
''''  '************** 1. 记录已选国家站的选中状态（ExtStaInfo2 及省/市/县去重列表统一由 RebuildExtSel来自选择ed 重建） **************
''''  For i = 0 To List_StaSURF.ListCount - 1
''''    ExtStaInfo1(i + 1).Selected = (List_StaSURF.Selected(i) = True)
''''  Next i
''''

  If GuestMode = True Then
    '2026-07-18 统一通过 RebuildExtSel来自选择ed 重建 ExtStaInfo2/ExtStaNum 及省份/地市/区县去重列表（与 FrmMeteoPeriodExt.Form_Load 共用同一份逻辑）
    Call RebuildExtSelFromSelected
  End If
    
  If Not GuestMode Then
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    ORAConn2.Open
    ORAConn2.CursorLocation = 3
    
    If SelectSURFExt = 1 Then
      tempZone = SelectedProvs
      If InStr(tempZone, ",") = 0 Then
        tempZone = "'" & tempZone & "'"
      End If
    End If
                 
    If SelectSURFExt = 1 Then '单选/多选或全选省份，用prcode字段来做限定条件
      strSURFExt = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_prcode in (" & tempZone & ") and v02301 like '%A%'and " + sttDate_valid + ")"
      
    ElseIf SelectSURFExt = 0.9 Then '登陆时默认的状态，保持登陆时默认值
      strSURFExt = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_prcode ='广东' and v02301 like '%A%' and v01301<>'59486'and " + sttDate_valid + ")"
      
    ElseIf SelectSURFExt = 0.5 Then '随意多选，把已选站点站号存入临时站点信息表
      strSQL = "delete from T_OTHE_STATION_CODE_TMP_TAB where USERID ='" & userID & "' and STATYPE = 'SURFExt' and USERIP ='" & UserIP & "'"
      ORAConn2.Execute strSQL
            
      strSQL = "insert all"
      j = 1
      For i = 0 To List_StaSURF.ListCount - 1
        If List_StaSURF.Selected(i) = True Then
          strSQL = strSQL + " into T_OTHE_STATION_CODE_TMP_TAB Values('" & userID & "','SURFExt','" & ExtStaInfo1(i + 1).stacode & "','" & UserIP & "')"
        End If
      Next i
      strSQL = strSQL + " select 1 from dual"
        
      StatusBar1.Panels(1).Text = "Writing national station codes to temp table..."
      ORAConn2.Execute strSQL
      
      StatusBar1.Panels(1).Text = "National station codes written to temp table: Done"
      strSURFExt = "stacode in (select stacode from T_OTHE_STATION_CODE_TMP_TAB where USERID ='" & userID & "' and STATYPE = 'SURFExt' and USERIP ='" & UserIP & "')"
    End If
    
    ORAConn4.Close
    ORAConn2.Close
  
  End If
  

  ExtStaInfo = ExtStaInfo2
  
  Call TableIni_MeteoPeriodExt(FrmMeteoPeriodExt, FrmMeteoPeriodExt.SelField_Initial)
  
  FrmMeteoPeriodExt.WindowState = vbNormal

  FrmMeteoPeriodExt.StatusBar1.Panels(2).Text = "Selected: " & ExtStaNum & " national stations"

  Screen.MousePointer = 1
  
  Unload Me
  Exit Sub
End Sub



Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub



Private Sub CmdImport_Click()
  Dim i%, j%, k%, tempStr$, strFileName$, sta_list$(), strSplit$()

  CommonDialogOpen.Filter = "Text Files (*.csv)|*.csv"
  CommonDialogOpen.InitDir = App.Path & "\Ini"
  CommonDialogOpen.ShowOpen
  If CommonDialogOpen.Flags = 0 Then Exit Sub
  strFileName = CommonDialogOpen.FileName
  FileNum = FreeFile()
  tempStr = ""
  Open strFileName For Input As #FileNum
  Line Input #FileNum, tempStr
  If tempStr <> "" Then
    sta_list = Split(tempStr, ",")
  End If
  Close #FileNum

  For i = 0 To List_StaSURF.ListCount - 1
    List_StaSURF.Selected(i) = False
  Next i
'  j = 0
  For k = 0 To UBound(sta_list)
    For i = 0 To List_StaSURF.ListCount - 1
      strSplit = Split(List_StaSURF.List(i), " ")
      If sta_list(k) = strSplit(0) Then
        List_StaSURF.Selected(i) = True
'        j = i + 1
        Exit For
      End If
    Next i
  Next k
  StatusBar1.Panels(1).Text = "Selected: " & List_StaSURF.SelCount & " national stationsStation"

End Sub



Private Sub CmdSave_Click()
  Dim i%, j%
  Dim strFileName$, strTemp$, strSplit$()

  CommonDialogSave.Filter = "Text Files (*.csv)|*.csv"
  CommonDialogSave.FilterIndex = 1
  CommonDialogSave.FileName = "Station List"
  CommonDialogSave.InitDir = App.Path & "\Ini"
  CommonDialogSave.Flags = &H2
  CommonDialogSave.ShowSave
  If CommonDialogSave.Flags = 2 Then Screen.MousePointer = 1: Exit Sub

  Screen.MousePointer = 11
  StatusBar1.Panels(1).Text = "Saving selected stations"

  strFileName = CommonDialogSave.FileName
  If Len(Dir(strFileName)) <> 0 Then Kill strFileName

  FileNum = FreeFile()
  Open strFileName For Output As #FileNum
  strTemp = ""

  For i = 0 To List_StaSURF.ListCount - 1
    If List_StaSURF.Selected(i) = True Then
      strSplit = Split(List_StaSURF.List(i), " ")
      strTemp = strSplit(0)
      Exit For
    End If
  Next i
  If strTemp <> "" Then
    For j = i + 1 To List_StaSURF.ListCount - 1
      If List_StaSURF.Selected(j) = True Then
        strSplit = Split(List_StaSURF.List(j), " ")
        strTemp = strTemp & "," & strSplit(0)
      End If
    Next j
    Print #FileNum, strTemp
  End If

  Close #FileNum
  StatusBar1.Panels(1).Text = "Station list saved"
  Screen.MousePointer = 1

End Sub


'鼠标滚轮驱动表格上下滚动
Private Sub HFGrid1_GotFocus()
  On Error Resume Next
  Oldwinproc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
  SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub
    
'鼠标滚轮驱动表格上下滚动
Private Sub HFGrid1_LostFocus()
  On Error Resume Next
  SetWindowLong Me.hWnd, GWL_WNDPROC, Oldwinproc
End Sub

