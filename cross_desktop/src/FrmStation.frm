VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmStation 
   BackColor       =   &H80000005&
   Caption         =   "StationSelect"
   ClientHeight    =   11850
   ClientLeft      =   165
   ClientTop       =   510
   ClientWidth     =   16530
   Icon            =   "FrmStation.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   14401.21
   ScaleMode       =   0  'User
   ScaleWidth      =   16525.46
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   400
      Left            =   15480
      Style           =   1  'Graphical
      TabIndex        =   34
      Top             =   1800
      Width           =   840
   End
   Begin VB.Frame Frame9 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Township Filter"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   12720
      TabIndex        =   30
      Top             =   120
      Width           =   3601
      Begin VB.ListBox List_Town 
         Appearance      =   0  'Flat
         Columns         =   2
         Height          =   1290
         Left            =   0
         MultiSelect     =   1  'Simple
         TabIndex        =   31
         Top             =   240
         Width           =   3600
      End
   End
   Begin VB.Frame Frame8 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "County Filter"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   9000
      TabIndex        =   28
      Top             =   120
      Width           =   3601
      Begin VB.ListBox List_County 
         Appearance      =   0  'Flat
         Columns         =   2
         Height          =   1290
         Left            =   0
         MultiSelect     =   1  'Simple
         TabIndex        =   29
         Top             =   240
         Width           =   3600
      End
   End
   Begin VB.Frame Frame7 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Output"
      Height          =   735
      Left            =   2760
      TabIndex        =   25
      Top             =   120
      Width           =   1575
      Begin VB.CheckBox CheckLonLat 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Show Lat/Lon"
         ForeColor       =   &H80000008&
         Height          =   255
         Left            =   120
         TabIndex        =   26
         Top             =   360
         Width           =   1400
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " National Station List"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   16
      Top             =   2280
      Width           =   16215
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import List"
         Height          =   400
         Index           =   0
         Left            =   15240
         Style           =   1  'Graphical
         TabIndex        =   21
         Top             =   1560
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save List"
         Height          =   400
         Index           =   0
         Left            =   15240
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   1080
         Width           =   840
      End
      Begin VB.ListBox List_StaSURF 
         Appearance      =   0  'Flat
         Columns         =   5
         Height          =   1830
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   17
         Top             =   240
         Width           =   15015
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "City Filter"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   5880
      TabIndex        =   14
      Top             =   120
      Width           =   3001
      Begin VB.ListBox List_City 
         Appearance      =   0  'Flat
         Columns         =   3
         Height          =   1290
         Left            =   0
         MultiSelect     =   1  'Simple
         TabIndex        =   15
         Top             =   240
         Width           =   3000
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Result Spatial Scale"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   13
      Top             =   960
      Width           =   4215
      Begin VB.OptionButton OptionScale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "City"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   3360
         TabIndex        =   27
         Top             =   300
         Width           =   735
      End
      Begin VB.OptionButton OptionScale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "County"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   2400
         TabIndex        =   24
         Top             =   300
         Width           =   915
      End
      Begin VB.OptionButton OptionScale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Township"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   1200
         TabIndex        =   23
         Top             =   300
         Width           =   1095
      End
      Begin VB.OptionButton OptionScale 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "Station"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   22
         Top             =   300
         Width           =   975
      End
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   400
      Left            =   14520
      Style           =   1  'Graphical
      TabIndex        =   12
      Top             =   1800
      Width           =   840
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh"
      Height          =   400
      Left            =   13560
      Style           =   1  'Graphical
      TabIndex        =   11
      Top             =   1800
      Width           =   840
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   10
      Top             =   11595
      Width           =   16530
      _ExtentX        =   29157
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   5
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4411
            MinWidth        =   4411
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4411
            MinWidth        =   4411
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4411
            MinWidth        =   4411
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   4411
            MinWidth        =   4411
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   10859
            MinWidth        =   8822
            Text            =   "Avoid filtering stations via the AWS Station List where possible, to prevent long write times to the temp station table."
            TextSave        =   "Avoid filtering stations via the AWS Station List where possible, to prevent long write times to the temp station table."
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Zone Filter"
      ForeColor       =   &H80000008&
      Height          =   1575
      Left            =   4440
      TabIndex        =   5
      Top             =   120
      Width           =   1335
      Begin VB.CheckBox CheckRegion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "珠三角"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   975
      End
      Begin VB.CheckBox CheckRegion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "粤西"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   640
         Width           =   855
      End
      Begin VB.CheckBox CheckRegion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "粤东"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   7
         Top             =   920
         Width           =   975
      End
      Begin VB.CheckBox CheckRegion 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "粤北"
         ForeColor       =   &H80000008&
         Height          =   255
         Index           =   3
         Left            =   240
         TabIndex        =   6
         Top             =   1200
         Width           =   800
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   " AWS Station List"
      ForeColor       =   &H80000008&
      Height          =   6855
      Left            =   120
      TabIndex        =   3
      Top             =   4560
      Width           =   16215
      Begin VB.Frame Frame12 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1065
         Left            =   15240
         TabIndex        =   40
         Top             =   2400
         Width           =   855
         Begin VB.CheckBox CheckInland 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            Caption         =   "Check1"
            ForeColor       =   &H80000008&
            Height          =   255
            Left            =   300
            TabIndex        =   42
            Top             =   120
            Width           =   255
         End
         Begin VB.Label Label4 
            BackColor       =   &H8000000E&
            Caption         =   "Region Station"
            Height          =   375
            Left            =   120
            TabIndex        =   43
            Top             =   645
            Width           =   720
         End
         Begin VB.Label Label3 
            BackColor       =   &H8000000E&
            Caption         =   "Land"
            Height          =   255
            Left            =   250
            TabIndex        =   41
            Top             =   435
            Width           =   405
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   975
         Left            =   15240
         TabIndex        =   37
         Top             =   3480
         Width           =   855
         Begin VB.CheckBox CheckHydro 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   300
            TabIndex        =   38
            Top             =   120
            Width           =   225
         End
         Begin VB.Label Label2 
            BackColor       =   &H8000000E&
            Caption         =   "Hydro Station"
            Height          =   435
            Left            =   120
            TabIndex        =   39
            Top             =   420
            Width           =   720
         End
      End
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         ForeColor       =   &H80000008&
         Height          =   1185
         Left            =   15240
         TabIndex        =   35
         Top             =   4440
         Width           =   855
         Begin VB.CheckBox CheckOcean 
            Appearance      =   0  'Flat
            BackColor       =   &H80000005&
            ForeColor       =   &H80000008&
            Height          =   300
            Left            =   300
            TabIndex        =   36
            Top             =   120
            Width           =   225
         End
         Begin VB.Label Label5 
            BackColor       =   &H8000000E&
            Caption         =   "Region Station"
            Height          =   375
            Left            =   120
            TabIndex        =   45
            Top             =   645
            Width           =   640
         End
         Begin VB.Label Label1 
            BackColor       =   &H8000000E&
            Caption         =   "Marine"
            Height          =   255
            Left            =   200
            TabIndex        =   44
            Top             =   435
            Width           =   525
         End
      End
      Begin VB.CommandButton CmdImport 
         Caption         =   "Import List"
         Height          =   400
         Index           =   1
         Left            =   15240
         Style           =   1  'Graphical
         TabIndex        =   33
         Top             =   6240
         Width           =   840
      End
      Begin VB.CommandButton CmdSave 
         Caption         =   "Save List"
         Height          =   400
         Index           =   1
         Left            =   15240
         Style           =   1  'Graphical
         TabIndex        =   32
         Top             =   5760
         Width           =   840
      End
      Begin VB.ListBox List_StaAWST 
         Appearance      =   0  'Flat
         Columns         =   5
         Height          =   6510
         Left            =   120
         MultiSelect     =   1  'Simple
         TabIndex        =   4
         Top             =   240
         Width           =   15015
      End
   End
   Begin VB.CommandButton CmdClear 
      Appearance      =   0  'Flat
      Caption         =   "Clear"
      Height          =   400
      Left            =   12600
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1800
      Width           =   840
   End
   Begin VB.CommandButton CmdAll 
      Appearance      =   0  'Flat
      Caption         =   "Select All"
      Height          =   400
      Left            =   11400
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   1800
      Width           =   1080
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Station Type"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2535
      Begin VB.CheckBox CheckAWST 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "AWS Station"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   1440
         TabIndex        =   19
         Top             =   360
         Width           =   1020
      End
      Begin VB.CheckBox CheckSURF 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         Caption         =   "National Station"
         ForeColor       =   &H80000008&
         Height          =   300
         Left            =   240
         TabIndex        =   18
         Top             =   360
         Width           =   1215
      End
   End
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   8520
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComDlg.CommonDialog CommonDialogOpen 
      Left            =   9000
      Top             =   1800
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmStation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim SelectSURF_tmp As Single '进入模块时，将 selectSURF 存入临时变量，最后再读回
Dim SelectAWST_tmp As Single '进入模块时，将 selectSURF 存入临时变量，最后再读回

Dim SelectedCities$ '2026-8-6 用户选择的城市
Dim SelectedCounties$  '2025-06-05添加：记录查询的区县（默认为当前城市所是区县）
Dim SelectedTowns$ ' 2026-8-9：记录查询的镇街
Dim tempZone$ '2026-8-6 存放临时地区名

Dim Num_surf%, Num_inland%, Num_hydro%, Num_ocean% '已选的国家站、陆面区域站、水文站、海洋区域站

Private IgnoreCheckInland As Boolean, IgnoreCheckHydro As Boolean, IgnoreCheckOcean As Boolean   '不触发相应的复选框click过程（IgnoreCheckAWST As Boolean 不需要）
Private IgnoreListSURF As Boolean, IgnoreListAWST As Boolean '不触发国家站站点列表/区域站 Listclick过程
Private IgnoreListCities As Boolean '不触发城市列表click过程

Private Sub Show_StaNum()
  Dim i%
 
  Num_surf = List_StaSURF.SelCount
  
  Num_inland = 0: Num_hydro = 0: Num_ocean = 0
  If List_StaAWST.SelCount <> 0 Then
    For i = 1 To UBound(AWSTInfo1)
'      If AWSTInfo1(i).Selected = True Then
      If List_StaAWST.Selected(i - 1) = True Then '2026-8-19
        If AWSTInfo1(i).Inland = 1 And AWSTInfo1(i).Hydro = 0 Then Num_inland = Num_inland + 1
        If AWSTInfo1(i).Inland = 1 And AWSTInfo1(i).Hydro = 1 Then Num_hydro = Num_hydro + 1
        If AWSTInfo1(i).Inland = 0 Then Num_ocean = Num_ocean + 1
      End If
    Next i
  End If

  StatusBar1.Panels(1).Text = "Selected: " & Num_surf & " national stations"
  StatusBar1.Panels(2).Text = "Selected: " & Num_inland & " land zone stations"
  StatusBar1.Panels(3).Text = "Selected: " & Num_hydro & " hydrological stations"
  StatusBar1.Panels(4).Text = "Selected: " & Num_ocean & " marine zone stations"
End Sub




Private Sub Form_Load()
  Dim str1$, i%, j%
  
   '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = 11 * (Screen.Height - Me.Height) / 20
  
  SelectSURF_tmp = SelectSURF:  SelectAWST_tmp = SelectAWST
  IgnoreCheckInland = True: IgnoreCheckHydro = True: IgnoreCheckOcean = True
  IgnoreListSURF = True: IgnoreListAWST = True

  If StaType = "BOTH" Then
    CheckSURF.Value = Checked
    CheckAWST.Value = Checked
  ElseIf StaType = "SURF" Then
    CheckSURF.Value = Checked
    CheckAWST.Value = Unchecked
  ElseIf StaType = "AWST" Then
    CheckSURF.Value = Unchecked
    CheckAWST.Value = Checked
  End If
  
  If bCheckInland = True Then
    CheckInland.Value = Checked
  Else
    CheckInland.Value = Unchecked
  End If
  If bCheckHydro = True Then
    CheckHydro.Value = Checked
  Else
    CheckHydro.Value = Unchecked
  End If
  If bCheckOcean = True Then
    CheckOcean.Value = Checked
  Else
    CheckOcean.Value = Unchecked
  End If
  

  If FlagZone = "STA" Then OptionScale(0).Value = True
  If FlagZone = "TOWN" Then OptionScale(1).Value = True
  If FlagZone = "CNTY" Then OptionScale(2).Value = True
  If FlagZone = "CITY" Then OptionScale(3).Value = True
  
  If ShowLonLat = True Then CheckLonLat.Value = Checked
  If ShowLonLat = False Then CheckLonLat.Value = Unchecked
  

  ' 如果cityInfo2 被清空（Erase）了
  If (Not CityInfo2) = True Then '地市Info2被 erase了
    CityNum = 0
  Else
    CityNum = UBound(CityInfo2)
  End If
  
  For i = 1 To UBound(CityInfo1)
    List_City.AddItem CityInfo1(i).City
    If CityNum <> 0 Then '地市Num = 0时，不判断城市列表是否该选择
      For j = 1 To UBound(CityInfo2)
        If CityInfo1(i).City = CityInfo2(j).City Then
          List_City.Selected(i - 1) = True: Exit For
        End If
      Next j
    End If
  Next i
  
  For i = 1 To UBound(SURFInfo1)
    List_StaSURF.AddItem SURFInfo1(i).stacode & " " & SURFInfo1(i).staname
    If SURFInfo1(i).Selected = True Then List_StaSURF.Selected(i - 1) = True
  Next i
  
  For i = 1 To UBound(AWSTInfo1)
    List_StaAWST.AddItem AWSTInfo1(i).stacode & " " & AWSTInfo1(i).staname
    If AWSTInfo1(i).Selected = True Then List_StaAWST.Selected(i - 1) = True
  Next i
  
  IgnoreCheckInland = False: IgnoreCheckHydro = False: IgnoreCheckOcean = False
  IgnoreListSURF = False: IgnoreListAWST = False
  SelectSURF = SelectSURF_tmp:  SelectAWST = SelectAWST_tmp
  
  Call Show_StaNum

  
  CheckInland.Value = Checked '2026-8-18

End Sub

Private Sub CheckLonLat_Click()
  If CheckLonLat.Value = Checked Then ShowLonLat = True
  If CheckLonLat.Value = Unchecked Then ShowLonLat = False
End Sub

Private Sub CmdCancel_Click()
  Unload Me
End Sub

Private Sub CheckRegion_Click(Index As Integer)
  Dim i%
  IgnoreListCities = True
  
  If CheckRegion(Index).Value = Checked Then  'Select
    For i = 0 To List_City.ListCount - 1
      If CityInfo1(i + 1).Region = CheckRegion(Index).Caption Then List_City.Selected(i) = True
    Next i
  ElseIf CheckRegion(Index).Value = Unchecked Then  '不选
    For i = 0 To List_City.ListCount - 1
      If CityInfo1(i + 1).Region = CheckRegion(Index).Caption Then List_City.Selected(i) = False
    Next i
  End If
  
  IgnoreListCities = False
  Call List_city_Click
  
  Call Show_StaNum

End Sub


Private Sub List_city_Click()
  Dim i%, j%
  If IgnoreListCities = True Then Exit Sub
  
  List_County.Clear
  For i = 1 To UBound(CountyInfo1)
    If CountyInfo1(i).City = List_City.Text Then List_County.AddItem CountyInfo1(i).County
  Next i

  For i = 0 To List_City.ListCount - 1
    If List_City.Selected(i) = True Then 'Select
        
      For j = 0 To List_StaSURF.ListCount - 1
        If SURFInfo1(j + 1).City = List_City.List(i) Then
          List_StaSURF.Selected(j) = True: SURFInfo1(j + 1).Selected = True
        End If
      Next j
        
      For j = 0 To List_StaAWST.ListCount - 1
        If AWSTInfo1(j + 1).City = List_City.List(i) Then
            
          If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then 'LandRegional Station
            If CheckInland.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
            If CheckInland.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
          
          ElseIf AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then 'Hydrological Station
            If CheckHydro.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
            If CheckHydro.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
          
          ElseIf AWSTInfo1(j + 1).Inland = 0 Then 'MarineRegional Station
            If CheckOcean.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
            If CheckOcean.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
          End If
        End If
      Next j

    ElseIf List_City.Selected(i) = False Then  '不选择
      For j = 0 To List_StaSURF.ListCount - 1
        If SURFInfo1(j + 1).City = List_City.List(i) Then
          List_StaSURF.Selected(j) = False: SURFInfo1(j + 1).Selected = False
        End If
      Next j
      For j = 0 To List_StaAWST.ListCount - 1
        If AWSTInfo1(j + 1).City = List_City.List(i) Then
          List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
        End If
      Next j

    End If
  Next i

  '点击“全选”按钮的效果与选择（多个）城市是一样的，所以这里让选择状态变成 1：
  SelectSURF = 1: SelectAWST = 1
  
  Call Show_StaNum

End Sub


Private Sub List_county_Click()
  Dim i%, j%
  List_Town.Clear
  For i = 1 To UBound(TownInfo1)
    If TownInfo1(i).County = List_County.Text Then List_Town.AddItem TownInfo1(i).Town
  Next i
  
  For i = 0 To List_County.ListCount - 1
''''    If List_区县.List(i) = "" Then Go送至 Next区县 '如果区县名称为空，不要跳过这个区县，免得与其相关的站点没是得到是效判断
    If List_County.Selected(i) = True Then 'Select
           
      For j = 0 To List_StaSURF.ListCount - 1
        If SURFInfo1(j + 1).County = List_County.List(i) Then
          List_StaSURF.Selected(j) = True: SURFInfo1(j + 1).Selected = True
        End If
      Next j
      
      For j = 0 To List_StaAWST.ListCount - 1
        If AWSTInfo1(j + 1).County = List_County.List(i) Then
            
          If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then 'LandRegional Station
            If CheckInland.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
            If CheckInland.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
          
          ElseIf AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then 'Hydrological Station
            If CheckHydro.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
            If CheckHydro.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
          
          ElseIf AWSTInfo1(j + 1).Inland = 0 Then 'MarineRegional Station
            If CheckOcean.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
            If CheckOcean.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
          End If
        
        End If
        
        
      Next j
      
    ElseIf List_County.Selected(i) = False Then  '不选择
      For j = 0 To List_StaSURF.ListCount - 1
        If SURFInfo1(j + 1).County = List_County.List(i) Then
          List_StaSURF.Selected(j) = False: SURFInfo1(j + 1).Selected = False
        End If
      Next j
      For j = 0 To List_StaAWST.ListCount - 1
        If AWSTInfo1(j + 1).County = List_County.List(i) Then
          List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
        End If
      Next j
    End If
      
''''NextCounty:
  Next i

  '选择（多个）县区让选择状态变成 0.5也无妨（因为站点少，不影响写入临时表）：
  SelectSURF = 0.5: SelectAWST = 0.5

  Call Show_StaNum

End Sub


Private Sub List_Town_Click()
  Dim i%, j%, k%
  For i = 0 To List_Town.ListCount - 1
''''    If List_送至wn.List(i) = "" Then Go送至 Next送至wn '如果镇街名称为空，不要跳过这个镇街，免得与其相关的站点没是得到是效判断
    If List_Town.Selected(i) = True Then 'Select
        
      For j = 0 To List_StaSURF.ListCount - 1
        If SURFInfo1(j + 1).Town = List_Town.List(i) Then
          For k = 0 To List_County.ListCount - 1 '判断该站点是否所属显示的镇街
            If SURFInfo1(j + 1).County = List_County.List(k) Then
              List_StaSURF.Selected(j) = True: SURFInfo1(j + 1).Selected = True
            End If
          Next k
        End If
      Next j
      
      For j = 0 To List_StaAWST.ListCount - 1
        If AWSTInfo1(j + 1).Town = List_Town.List(i) Then
          For k = 0 To List_County.ListCount - 1 '判断站点是否属于镇街列表中的县区
            If AWSTInfo1(j + 1).County = List_County.List(k) Then
          
              If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then 'LandRegional Station
                If CheckInland.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                If CheckInland.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
                
              ElseIf AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then 'Hydrological Station
                If CheckHydro.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                If CheckHydro.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
                
              ElseIf AWSTInfo1(j + 1).Inland = 0 Then 'MarineRegional Station
                If CheckOcean.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                If CheckOcean.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
              End If
              
              GoTo NextAWST  '2026-8-9 已经找到站点所在镇街所属区县了，不需要继续判断，直接跳出
            End If
          Next k
        End If

NextAWST:
      Next j
      
    ElseIf List_Town.Selected(i) = False Then  '不选择
      For j = 0 To List_StaSURF.ListCount - 1
        If SURFInfo1(j + 1).Town = List_Town.List(i) Then
          List_StaSURF.Selected(j) = False: SURFInfo1(j + 1).Selected = False
        End If
      Next j
      For j = 0 To List_StaAWST.ListCount - 1
        If AWSTInfo1(j + 1).Town = List_Town.List(i) Then
          List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
        End If
      Next j
    End If
      
NextTown:
  Next i

  '选择（多个）镇街让选择状态变成 0.5也无妨（因为站点少，不影响写入临时表）：
  SelectSURF = 0.5: SelectAWST = 0.5

  Call Show_StaNum

End Sub


Private Sub List_StaSURF_Click()
  If IgnoreListSURF Then Exit Sub
  SelectSURF = 0.5
End Sub


Private Sub List_StaAWST_Click()
  If IgnoreListAWST Then Exit Sub
  SelectAWST = 0.5
End Sub


Private Sub CmdRefresh_Click()
  IgnoreListAWST = True '2026-8-19
  Call Show_StaNum
  IgnoreListAWST = False '2026-8-19
End Sub


Private Sub CmdAll_Click()
  Dim i%
  For i = 0 To List_City.ListCount - 1
    List_City.Selected(i) = True
  Next i
  
  For i = 0 To List_StaSURF.ListCount - 1
      List_StaSURF.Selected(i) = True
  Next i
  
  For i = 0 To List_StaAWST.ListCount - 1
''''    List_StaAWST.Selected(i) = True
    
    If CheckOcean.Value = Checked Then
      '若勾选了“海洋区域站”，则全选时全选海洋区域站
      If AWSTInfo1(i + 1).Inland = 0 Then List_StaAWST.Selected(i) = True: AWSTInfo1(i + 1).Selected = True
''''    Else
''''      '若没是勾选“海洋区域站”，则全选时清空海洋区域站
''''      If AWSTInfo1(i + 1).Inland = 0 Then List_StaAWST.Selected(i) = False: AWSTInfo1(i + 1).Selected = False
    End If
    
    If CheckHydro.Value = Checked Then
      '若勾选了“水文站”，则全选时全选水文站
      If AWSTInfo1(i + 1).Inland = 1 And AWSTInfo1(i + 1).Hydro = 1 Then List_StaAWST.Selected(i) = True: AWSTInfo1(i + 1).Selected = True
    End If
      
    If CheckInland.Value = Checked Then
      '若勾选了“陆面区域站”，则全选时全选陆面区域站
      If AWSTInfo1(i + 1).Inland = 1 And AWSTInfo1(i + 1).Hydro = 0 Then List_StaAWST.Selected(i) = True: AWSTInfo1(i + 1).Selected = True
    End If
  Next i
  
''''  If CheckSURF.Value = Checked Then
    SelectSURF = 1
''''  End If

    SelectAWST = 1
  
  Call Show_StaNum
End Sub


Private Sub CmdClear_Click()
  Dim i%
  For i = 0 To 3
    CheckRegion(i).Value = Unchecked
  Next i
  
  '2026-07-19 用 LB_SETSEL 一次性清空多选ListBox的选中项，替代逐项循环，加快"清空"按钮响应速度
  Call SendMessage(List_City.hWnd, LB_SETSEL, 0, -1)
  Call SendMessage(List_County.hWnd, LB_SETSEL, 0, -1)
  Call SendMessage(List_Town.hWnd, LB_SETSEL, 0, -1)
  Call SendMessage(List_StaSURF.hWnd, LB_SETSEL, 0, -1)
  Call SendMessage(List_StaAWST.hWnd, LB_SETSEL, 0, -1)

  Erase CityInfo2: Erase CountyInfo2: Erase TownInfo2
  
  SelectSURF = 0:  SelectAWST = 0
  
  Call Show_StaNum
  
End Sub


Private Sub CmdOK_Click()
  Dim j%, i%
  Dim tempStr$()
  
  If CheckSURF.Value = Unchecked And CheckAWST.Value = Unchecked Then
    MsgBox "Select at least one station type", vbInformation, "Notice"
    Exit Sub
  End If
  
  If CheckSURF.Value = Checked And List_StaSURF.SelCount = 0 Then
    MsgBox "National Stations is checked -- select at least one national station", vbInformation, "Notice"
    Exit Sub
  End If
  
  If CheckAWST.Value = Checked And List_StaAWST.SelCount = 0 Then
    MsgBox "AWS Stations is checked -- select at least one AWS station", vbInformation, "Notice"
    Exit Sub
  End If
  
  If OptionScale(0).Value = True Then FlagZone = "STA"
  If OptionScale(1).Value = True Then FlagZone = "TOWN"
  If OptionScale(2).Value = True Then FlagZone = "CNTY"
  If OptionScale(3).Value = True Then FlagZone = "CITY"
  
  If FlagZone <> "STA" And List_City.SelCount = 0 Then
    MsgBox "Not station-based statistics -- select at least one city", vbInformation, "Notice"
    Exit Sub
  End If
  
  Screen.MousePointer = 11
  
  
  If CheckInland.Value = Checked Then bCheckInland = True
  If CheckInland.Value = Unchecked Then bCheckInland = False
  If CheckHydro.Value = Checked Then bCheckHydro = True
  If CheckHydro.Value = Unchecked Then bCheckHydro = False
  If CheckOcean.Value = Checked Then bCheckOcean = True
  If CheckOcean.Value = Unchecked Then bCheckOcean = False
  
  Call Show_StaNum
  
  '基于镇、县、市统计，则获取镇、县、市信息
  '  If OptionScale(0).Value = False Then
  ' 2023-3-9 重新配置用户所选的数据区域DataZone
  ' 2026-8-6 配置用户所选的城市 SelectedCities
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
  
  '2025-06-05：如果个别区县被选择
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
    
  '2026-08-09：如果个别镇街被选择
  SelectedTowns = ""
  If List_Town.SelCount = 1 Then
'    If List_Town.SelCount <> List_Town.ListCount Or List_Town.SelCount <> 0 Then
    For j = 0 To List_Town.ListCount - 1
      If List_Town.Selected(j) = True Then
        SelectedTowns = "" & List_Town.List(j) & "" '只是一个镇街的时候，一定不要加引号！
      End If
    Next j
'    End If
  
  ElseIf List_Town.SelCount >= 2 Then
'    If List_Town.SelCount <> List_Town.ListCount Or List_Town.SelCount <> 0 Then
    For j = 0 To List_Town.ListCount - 1
      If List_Town.Selected(j) = True Then
        If SelectedTowns = "" Then
          SelectedTowns = "'" & List_Town.List(j) & "'"
        ElseIf SelectedTowns <> "" Then
          SelectedTowns = SelectedTowns & ",'" & List_Town.List(j) & "'"
        End If
      End If
    Next j
'    End If
  End If
  
  If Not GuestMode Then
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    ORAConn2.Open
    ORAConn2.CursorLocation = 3
    
    If SelectedCities <> "" Then '2023-3-14修补bug
      Call GetTownInfo2(SelectedCities, SelectedCounties, SelectedTowns)
      Call GetCountyInfo2(SelectedCities, SelectedCounties)
      Call GetCityInfo2(SelectedCities)
    End If
  
    ORAConn4.Close
    ORAConn2.Close
  End If
  
  If (Not TownInfo2) = True Then '送至wnInfo2被 erase了
    TownNum = 0
  Else
    TownNum = UBound(TownInfo2)
  End If
  
  If (Not CountyInfo2) = True Then '区县Info2被 erase了
    CountyNum = 0
  Else
    CountyNum = UBound(CountyInfo2)
  End If
  
  If (Not CityInfo2) = True Then '地市Info2被 erase了
    CityNum = 0
  Else
    CityNum = UBound(CityInfo2)
  End If
  
  StaNum_SURF = List_StaSURF.SelCount
  StaNum_AWST = List_StaAWST.SelCount

  '************** 1. 把已选国家站站点存入站点信息变量 **************
  If StaNum_SURF = 0 Then
    For i = 0 To List_StaSURF.ListCount - 1
      SURFInfo1(i + 1).Selected = False
    Next i
  
  Else
    ReDim SURFInfo2(1 To StaNum_SURF)
    j = 1
    For i = 0 To List_StaSURF.ListCount - 1
      If List_StaSURF.Selected(i) = True Then
        SURFInfo1(i + 1).Selected = True
        SURFInfo2(j).Prov = SURFInfo1(i + 1).Prov: SURFInfo2(j).Region = SURFInfo1(i + 1).Region
        SURFInfo2(j).City = SURFInfo1(i + 1).City: SURFInfo2(j).County = SURFInfo1(i + 1).County
        SURFInfo2(j).Town = SURFInfo1(i + 1).Town
        SURFInfo2(j).StaType = "SURF"
        SURFInfo2(j).stacode = SURFInfo1(i + 1).stacode: SURFInfo2(j).staname = SURFInfo1(i + 1).staname
        SURFInfo2(j).Longitude = SURFInfo1(i + 1).Longitude: SURFInfo2(j).Latitude = SURFInfo1(i + 1).Latitude
        SURFInfo2(j).DateSTT = SURFInfo1(i + 1).DateSTT
        SURFInfo2(j).YearSTT = Year(CDate(SURFInfo1(i + 1).DateSTT))
        SURFInfo2(j).Inland = SURFInfo1(i + 1).Inland
        SURFInfo2(j).Hydro = SURFInfo1(i + 1).Hydro
                
        j = j + 1
      Else
        SURFInfo1(i + 1).Selected = False
      End If
    Next i
    
  End If
    
  If Not GuestMode Then
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    ORAConn2.Open
    ORAConn2.CursorLocation = 3
    
    If CheckSURF.Value = Checked Then
      If SelectSURF = 1 Then
        tempZone = SelectedCities
        If InStr(tempZone, ",") = 0 Then
          tempZone = "'" & tempZone & "'"
        End If
      End If

                  
      If SelectSURF = 1 Then '单选/多选或全选城市，用city字段来做限定条件
        strSURF = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_city in (" & tempZone & ") and v02301 like '%A%'and " + sttDate_valid + ")"
        
      ElseIf SelectSURF = 0.9 Then '登陆时默认的状态，保持登陆时默认值
        If DataZone = "广东" Or DataZone = "粤港澳" Then
          strSURF = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_prcode ='广东' and v02301 like '%A%' and v01301<>'59486'and " + sttDate_valid + ")"
        Else
          strSURF = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_city ='" & DataZone & "' and v02301 like '%A%' and v01301<>'59486'and " + sttDate_valid + ")"
        End If
      
      ElseIf SelectSURF = 0.5 Then '随意多选，把已选站点站号存入临时站点信息表
        strSQL = "delete from T_OTHE_STATION_CODE_TMP_TAB where USERID ='" & userID & "' and STATYPE = 'SURF' and USERIP ='" & UserIP & "'"
        ORAConn2.Execute strSQL
            
        strSQL = "insert all"
        j = 1
        For i = 0 To List_StaSURF.ListCount - 1
          If List_StaSURF.Selected(i) = True Then
            strSQL = strSQL + " into T_OTHE_STATION_CODE_TMP_TAB Values('" & userID & "','SURF','" & SURFInfo1(i + 1).stacode & "','" & UserIP & "')"
          End If
        Next i
        strSQL = strSQL + " select 1 from dual"
        
        StatusBar1.Panels(1).Text = "Writing national station codes to temp table..."
        ORAConn2.Execute strSQL
        
        StatusBar1.Panels(1).Text = "National station codes written to temp table: Done"
        strSURF = "stacode in (select stacode from T_OTHE_STATION_CODE_TMP_TAB where USERID ='" & userID & "' and STATYPE = 'SURF' and USERIP ='" & UserIP & "')"
      End If
    End If
  
    ORAConn4.Close
    ORAConn2.Close
  
  End If
  
  
  '************** 2. 把已选自动站点存入站点信息变量 ************
  If StaNum_AWST = 0 Then
    For i = 0 To List_StaAWST.ListCount - 1
      AWSTInfo1(i + 1).Selected = False
    Next i
  
  Else
    ReDim AWSTInfo2(1 To StaNum_AWST)
    j = 1
    For i = 0 To List_StaAWST.ListCount - 1
      If List_StaAWST.Selected(i) = True Then
        AWSTInfo1(i + 1).Selected = True
        AWSTInfo2(j).Prov = AWSTInfo1(i + 1).Prov: AWSTInfo2(j).Region = AWSTInfo1(i + 1).Region
        AWSTInfo2(j).City = AWSTInfo1(i + 1).City: AWSTInfo2(j).County = AWSTInfo1(i + 1).County
        AWSTInfo2(j).Town = AWSTInfo1(i + 1).Town
        AWSTInfo2(j).StaType = "AWST"
        AWSTInfo2(j).stacode = AWSTInfo1(i + 1).stacode: AWSTInfo2(j).staname = AWSTInfo1(i + 1).staname
        AWSTInfo2(j).Longitude = AWSTInfo1(i + 1).Longitude: AWSTInfo2(j).Latitude = AWSTInfo1(i + 1).Latitude
        AWSTInfo2(j).DateSTT = AWSTInfo1(i + 1).DateSTT
        AWSTInfo2(j).YearSTT = Year(CDate(AWSTInfo1(i + 1).DateSTT))
        AWSTInfo2(j).Inland = AWSTInfo1(i + 1).Inland
        AWSTInfo2(j).Hydro = AWSTInfo1(i + 1).Hydro
        
        j = j + 1
      Else
        AWSTInfo1(i + 1).Selected = False
      End If
    Next i
    
  End If
  
    
  If Not GuestMode Then
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    ORAConn2.Open
    ORAConn2.CursorLocation = 3
    
    If CheckAWST.Value = Checked Then
      If SelectAWST >= 0.9 Then
        tempZone = SelectedCities
        If InStr(tempZone, ",") = 0 Then
          tempZone = "'" & tempZone & "'"
        End If
      End If

      
      If SelectAWST >= 0.9 Then '单选/多选/全选城市或登陆时默认的状态，用city字段来做限定条件
        strAWST = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_city in (" & tempZone & ")"
        
        If bCheckInland = True And bCheckHydro = True And bCheckOcean = True Then  'LandRegional Station+Hydrological Station+MarineRegional Station
          strAWST = strAWST & " and v02301 not like '%A%'"

        ElseIf bCheckInland = True And bCheckHydro = True And bCheckOcean = False Then  'LandRegional Station+Hydrological Station
          strAWST = strAWST & " and (v02301 not like '%A%'and inland=1)"

        ElseIf bCheckInland = False And bCheckHydro = True And bCheckOcean = True Then  'Hydrological Station+MarineRegional Station
          strAWST = strAWST & " and (v02301 like '%O%'or v02301 like '%B%'and inland=0)"

        ElseIf bCheckInland = True And bCheckHydro = False And bCheckOcean = True Then  'LandRegional Station+MarineRegional Station
          strAWST = strAWST & " and v02301 like '%B%'"

        ElseIf bCheckInland = True And bCheckHydro = False And bCheckOcean = False Then  'LandRegional Station
          strAWST = strAWST & " and (v02301 like '%B%'and inland=1)"

        ElseIf bCheckInland = False And bCheckHydro = False And bCheckOcean = True Then  'MarineRegional Station
          strAWST = strAWST & " and (v02301 like '%B%'and inland=0)"

        ElseIf bCheckInland = False And bCheckHydro = True And bCheckOcean = False Then 'Hydrological Station
          strAWST = strAWST & " and v02301 like '%O%'"
        End If
        
        strAWST = strAWST & " and " + sttDate_valid + ")"
           
      Else  '随意多选，把已选站点站号存入临时站点信息表
        strSQL = "delete from T_OTHE_STATION_CODE_TMP_TAB where USERID ='" & userID & "' and STATYPE = 'AWST' and USERIP ='" & UserIP & "'"
        ORAConn2.Execute strSQL
        strSQL = "insert all"
        For i = 0 To List_StaAWST.ListCount - 1
          If List_StaAWST.Selected(i) = True Then
            strSQL = strSQL + " into T_OTHE_STATION_CODE_TMP_TAB Values('" & userID & "','AWST','" & AWSTInfo1(i + 1).stacode & "','" & UserIP & "')"
          End If
        Next i
        strSQL = strSQL + " select 1 from dual"
        StatusBar1.Panels(2).Text = "Writing AWS station codes to temp table..."
        ORAConn2.Execute strSQL
        strAWST = "stacode in (select stacode from T_OTHE_STATION_CODE_TMP_TAB where USERID ='" & userID & "' and STATYPE = 'AWST' and USERIP ='" & UserIP & "')"
        StatusBar1.Panels(2).Text = "AWS station codes written to temp table: Done"
      
      End If
            
    End If
  
    ORAConn4.Close
    ORAConn2.Close
  End If


  If CheckSURF.Value = Checked And CheckAWST.Value = Checked Then '双选
    StaType = "BOTH"
  ElseIf CheckSURF.Value = Checked And CheckAWST.Value = Unchecked Then   'SelectNational Station
    StaType = "SURF"
  ElseIf CheckSURF.Value = Unchecked And CheckAWST.Value = Checked Then    'SelectAWS Station
    StaType = "AWST"
  End If
  
  
  '************ 2022-2-16 将站点信息写入通用的变量 StaInfo ************
  If StaType = "SURF" Then
    StaNum = StaNum_SURF
    ReDim StaInfo(1 To StaNum)
    StaInfo = SURFInfo2
    strSta = strSURF
  ElseIf StaType = "AWST" Then
    StaNum = StaNum_AWST
    ReDim StaInfo(1 To StaNum)
    StaInfo = AWSTInfo2
    strSta = strAWST
  ElseIf StaType = "BOTH" Then
    StaNum = StaNum_SURF + StaNum_AWST
    ReDim StaInfo(1 To StaNum)
    For i = 1 To StaNum_SURF
      StaInfo(i).Region = SURFInfo2(i).Region
      StaInfo(i).City = SURFInfo2(i).City
      StaInfo(i).County = SURFInfo2(i).County
      StaInfo(i).Town = SURFInfo2(i).Town
      StaInfo(i).StaType = "SURF"
      StaInfo(i).stacode = SURFInfo2(i).stacode
      StaInfo(i).staname = SURFInfo2(i).staname
      StaInfo(i).Longitude = SURFInfo2(i).Longitude
      StaInfo(i).Latitude = SURFInfo2(i).Latitude
      StaInfo(i).DateSTT = SURFInfo2(i).DateSTT
'      If StaInfo(i).DateSTT <> "" Then
      StaInfo(i).YearSTT = Year(CDate(StaInfo(i).DateSTT))
      StaInfo(i).Inland = SURFInfo2(i).Inland
      StaInfo(i).Hydro = SURFInfo2(i).Hydro
      
'      End If
    Next i
    For i = StaNum_SURF + 1 To StaNum
      StaInfo(i).Region = AWSTInfo2(i - StaNum_SURF).Region
      StaInfo(i).City = AWSTInfo2(i - StaNum_SURF).City
      StaInfo(i).County = AWSTInfo2(i - StaNum_SURF).County
      StaInfo(i).Town = AWSTInfo2(i - StaNum_SURF).Town
      StaInfo(i).StaType = "AWST"
      StaInfo(i).stacode = AWSTInfo2(i - StaNum_SURF).stacode
      StaInfo(i).staname = AWSTInfo2(i - StaNum_SURF).staname
      StaInfo(i).Longitude = AWSTInfo2(i - StaNum_SURF).Longitude
      StaInfo(i).Latitude = AWSTInfo2(i - StaNum_SURF).Latitude
      StaInfo(i).DateSTT = AWSTInfo2(i - StaNum_SURF).DateSTT
'      If StaInfo(i).DateSTT <> "" Then
      StaInfo(i).YearSTT = Year(CDate(StaInfo(i).DateSTT))
      StaInfo(i).Inland = AWSTInfo2(i - StaNum_SURF).Inland
      StaInfo(i).Hydro = AWSTInfo2(i - StaNum_SURF).Hydro

'      End If
    Next i
    
    '2022-2-28 将StaInfo按照stacode从小到大排序（选择排序法）
    Dim tempStaInfo As StaInfo
    For i = LBound(StaInfo) To UBound(StaInfo) - 1
      For j = i + 1 To UBound(StaInfo)
        If StaInfo(j).stacode < StaInfo(i).stacode Then
          tempStaInfo = StaInfo(j)
          StaInfo(j) = StaInfo(i)
          StaInfo(i) = tempStaInfo
        End If
      Next j
    Next i
  End If
  
  '************ 2022-2-16 将站点信息写入通用的变量 StaInfo ************
    If FormNum = 0 Then '如果没是打开的子窗体，则什么都不干
        
    Else '如果是打开的子窗体，则判断是哪一个子窗体，不需要重新配置option选项（2022-4-24），也不需要重新计算时次/日期/旬/月/年等（2022-6-22）
      If FrmMain.ActiveForm Is FrmMeteoHour Then
        Call TableIni_Meteo(FrmMeteoHour, FrmMeteoHour.BTime, FrmMeteoHour.iHors)
      ElseIf FrmMain.ActiveForm Is FrmMeteoDay Then
        Call TableIni_Meteo(FrmMeteoDay, FrmMeteoDay.BDate, FrmMeteoDay.idays)
      ElseIf FrmMain.ActiveForm Is FrmMeteoTen Then
        Call TableIni_MeteoTenQtr(FrmMeteoTen, FrmMeteoTen.BYYMM, FrmMeteoTen.iMons, FrmMeteoTen.iTens)
      
      ElseIf FrmMain.ActiveForm Is FrmMeteoMon Then
        Call TableIni_Meteo(FrmMeteoMon, FrmMeteoMon.BYYMM, FrmMeteoMon.iMons)
      
      ElseIf FrmMain.ActiveForm Is FrmMeteoQtr Then
        Call TableIni_MeteoTenQtr(FrmMeteoQtr, FrmMeteoQtr.BYYYY, FrmMeteoQtr.iYers, FrmMeteoQtr.iQtrs)
     
      ElseIf FrmMain.ActiveForm Is FrmMeteoYer Then
        Call TableIni_Meteo(FrmMeteoYer, FrmMeteoYer.BYYYY, FrmMeteoYer.iYers)
      ElseIf FrmMain.ActiveForm Is FrmMeteoPeriodYer Then
        Call TableIni_Meteo(FrmMeteoPeriodYer, FrmMeteoPeriodYer.BYear, FrmMeteoPeriodYer.iYers)
      ElseIf FrmMain.ActiveForm Is FrmMeteoPeriodSin Then
        Call TableIni_MeteoPeriodSin(FrmMeteoPeriodSin, FrmMeteoPeriodSin.SelField_Initial)
  
      ElseIf FrmMain.ActiveForm Is FrmMeteoPeriod Then
      
        If MultiSel = False Then Call TableIni_MeteoPeriod(FrmMeteoPeriod, FrmMeteoPeriod.SelField_Initial)
        If MultiSel = True Then Call TableIni_MeteoPeriod2(FrmMeteoPeriod)
    
    
      ElseIf FrmMain.ActiveForm Is FrmMeteoColdAir Then
        If FrmMeteoColdAir.FlagStat = "Multi" Then Call TableIni_MeteoColdAir(FrmMeteoColdAir)
        If FrmMeteoColdAir.FlagStat = "Single" Then Call TableIni_MeteoColdAir2(FrmMeteoColdAir)

      ElseIf FrmMain.ActiveForm Is FrmMeteoSeasons Then
        Call TableIni_MeteoSeasons(FrmMeteoSeasons)
        
      ElseIf FrmMain.ActiveForm Is FrmConditionQuery Then
        Call FrmConditionQuery.HFGrid1.Clear
        
      End If
    End If
  
           
  If StaType = "SURF" Then
    FrmMain.StatusBar1.Panels(2).Text = "Selected: " & Num_surf & " national stations"
    FrmMain.StatusBar1.Panels(3).Text = "No AWS stations selected"
  ElseIf StaType = "AWST" Then
    FrmMain.StatusBar1.Panels(2).Text = "No national stations selected"
    FrmMain.StatusBar1.Panels(3).Text = "Selected: " & Num_inland + Num_hydro + Num_ocean & " AWS stations"
  ElseIf StaType = "BOTH" Then
    FrmMain.StatusBar1.Panels(2).Text = "Selected: " & Num_surf & " national stations"
    FrmMain.StatusBar1.Panels(3).Text = "Selected: " & Num_inland + Num_hydro + Num_ocean & " AWS stations"
  End If


  '2026-8-26：数据空间范围（用于保存数据的时候写入数据表）
  If StaType = "SURF" Or StaType = "BOTH" Then
    If SelectSURF = 1 Then
      DataExtent = tempZone
    ElseIf SelectSURF = 0.9 Then
      If DataZone = "广东" Or DataZone = "粤港澳" Then
        DataExtent = "广东"
      Else
        DataExtent = DataZone
      End If
    ElseIf SelectSURF = 0.5 Then
      DataExtent = "Multiple Stations"
    End If
  
  ElseIf StaType = "AWST" Then
    If SelectAWST >= 0.9 Then
      DataExtent = tempZone
    Else
      DataExtent = "Multiple Stations"
    End If
  End If
  

 

  Screen.MousePointer = 1
  
  Unload Me
  Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub



Private Sub CmdImport_Click(Index As Integer)
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
  
  If Index = 0 Then
    For i = 0 To List_StaSURF.ListCount - 1
      List_StaSURF.Selected(i) = False
    Next i
    j = 0
    For k = 0 To UBound(sta_list)
      For i = j To List_StaSURF.ListCount - 1
        strSplit = Split(List_StaSURF.List(i), " ")
        If sta_list(k) = strSplit(0) Then
          
          List_StaSURF.Selected(i) = True
          j = i + 1
          Exit For
        End If
      Next i
    Next k
    StatusBar1.Panels(1).Text = "Selected: " & List_StaSURF.SelCount & " national stations"
    
  ElseIf Index = 1 Then
    For i = 0 To List_StaAWST.ListCount - 1
      List_StaAWST.Selected(i) = False
    Next i
    j = 0
    For k = 0 To UBound(sta_list)
      For i = j To List_StaAWST.ListCount - 1
        strSplit = Split(List_StaAWST.List(i), " ")
        If sta_list(k) = strSplit(0) Then
          
          List_StaAWST.Selected(i) = True
          j = i + 1
          Exit For
        End If
      Next i
    Next k
    StatusBar1.Panels(2).Text = "Selected: " & List_StaAWST.SelCount & " AWS stations"
  End If
  
End Sub


Private Sub CmdSave_Click(Index As Integer)
  Dim i%, j%
  Dim strFileName$, strFileType$, strTemp$, strStacode$, strSplit$()

  CommonDialogSave.Filter = "Text Files (*.csv)|*.csv"
  CommonDialogSave.FilterIndex = 1
  CommonDialogSave.FileName = "Station List"
  CommonDialogSave.InitDir = App.Path & "\Ini"
  CommonDialogSave.Flags = &H2
  CommonDialogSave.ShowSave
  If CommonDialogSave.Flags = 2 Then Screen.MousePointer = 1: Exit Sub

  frmWait.Label1.Caption = "Saving, please wait"
  Screen.MousePointer = 11
  FrmMain.StatusBar1.Panels(1).Text = "Saving selected stations"

  strFileName = CommonDialogSave.FileName
  If Len(Dir(strFileName)) <> 0 Then Kill strFileName '如果存在同名文件则删除同名文件

  FileNum = FreeFile()
  Open strFileName For Output As #FileNum
  strTemp = ""
  
  If Index = 0 Then 'National StationStation
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
    
  ElseIf Index = 1 Then 'AWS StationStation
    For i = 0 To List_StaAWST.ListCount - 1
      If List_StaAWST.Selected(i) = True Then
        strSplit = Split(List_StaAWST.List(i), " ")
        strTemp = strSplit(0)
        Exit For
      End If
    Next i
    If strTemp <> "" Then
      For j = i + 1 To List_StaAWST.ListCount - 1
        If List_StaAWST.Selected(j) = True Then
          strSplit = Split(List_StaAWST.List(j), " ")
          strTemp = strTemp & "," & strSplit(0)
        End If
      Next j
      Print #FileNum, strTemp
    End If
  
  End If

  Close #FileNum
  FrmMain.StatusBar1.Panels(1).Text = "Stations saved"
  Screen.MousePointer = 1

End Sub


' 2026-8-7 筛选自动站点列表中的陆面区域站，并根据CheckInland状态，选或不选（只管陆面区域站，不操作其他类型站点）
Private Sub Checkinland_Click()
  If IgnoreCheckInland Then Exit Sub
  
  Dim i%, j%, k%
    
  IgnoreListAWST = True ' List_StaAWST_Click() 动作失效
  SelectSURF_tmp = SelectSURF: SelectAWST_tmp = SelectAWST

  If GuestMode Then
    For j = 0 To List_StaAWST.ListCount - 1
      If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then
        If CheckInland.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
        If CheckInland.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
      End If
    Next j
    GoTo ENDSelection
  End If

  '没是选择任何地区，则陆面区域站点不选
  If List_City.SelCount = 0 And List_County.SelCount = 0 And List_Town.SelCount = 0 Then
    GoTo ENDSelection
  End If

  '****** 1.一定要首先判断镇街：若选择了某些镇街，则仅判断是否选择该些地区水文站
  If List_Town.SelCount > 0 Then
    For i = 0 To List_Town.ListCount - 1

      If List_Town.Selected(i) = True Then 'Select

        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).Town = List_Town.List(i) Then

            For k = 0 To List_County.ListCount - 1 '判断站点是否属于县区列表中的县区
              If AWSTInfo1(j + 1).County = List_County.List(k) Then

                If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then
                    If CheckInland.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                    If CheckInland.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
                End If
              End If
            Next k
          End If
        Next j
    
      ElseIf List_Town.Selected(i) = False Then  '不选择
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).Town = List_Town.List(i) Then
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then
              List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
    
      End If
    Next i

    GoTo ENDSelection
  End If
    
      
  '****** 2.然后判断区县：若选择了某些区县，则仅判断是否选择该些地区水文站
  If List_County.SelCount > 0 Then
    For i = 0 To List_County.ListCount - 1
    
      If List_County.Selected(i) = True Then 'Select
            
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).County = List_County.List(i) Then
                
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then '如果该entry是陆面区域站，则选或不选
                If CheckInland.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                If CheckInland.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
    
      ElseIf List_County.Selected(i) = False Then  '不选择
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).County = List_County.List(i) Then
          
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then
              List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
      End If

    Next i
    GoTo ENDSelection
  End If
  
  
  '最后判断城市：若选择了某些地市，则仅判断是否选择该些地区陆面区域站
  If List_City.SelCount > 0 Then
    For i = 0 To List_City.ListCount - 1
     
      If List_City.Selected(i) = True Then 'Select
            
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).City = List_City.List(i) Then
                
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then '如果该entry是陆面区域站，则选或不选
                If CheckInland.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                If CheckInland.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          
          End If
        Next j
    
      ElseIf List_City.Selected(i) = False Then  '不选择
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).City = List_City.List(i) Then
          
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 0 Then '如果该entry是陆面区域站，则不选
              List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
      End If
    Next i
      
    GoTo ENDSelection
  End If
    
ENDSelection:

  SelectSURF = SelectSURF_tmp: SelectAWST = SelectAWST_tmp
  
  Call Show_StaNum
  
  IgnoreListAWST = False ' List_StaAWST_Click() 恢复起效

End Sub



' 2026-8-7 筛选自动站点列表中的水文站，并根据CheckHydro状态，选或不选（只管水文站，不操作其他类型站点）(陆面区域站和海洋区域站都复制该click过程进行修改)
Private Sub CheckHydro_Click()
  If IgnoreCheckHydro Then Exit Sub
  
  
  Dim i%, j%, k%
    
  IgnoreListAWST = True ' List_StaAWST_Click() 动作失效
  SelectSURF_tmp = SelectSURF: SelectAWST_tmp = SelectAWST

  '2026-08-13 访客模式
  If GuestMode Then
    For j = 0 To List_StaAWST.ListCount - 1
      If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then
        If CheckHydro.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
        If CheckHydro.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
      End If
    Next j
    GoTo ENDSelection
  End If

  '没是选择任何地区，则水文站点不选
  If List_City.SelCount = 0 And List_County.SelCount = 0 And List_Town.SelCount = 0 Then
    GoTo ENDSelection
  End If
  
   '****** 1.一定要首先判断镇街：若选择了某些镇街，则仅判断是否选择该些地区水文站
  If List_Town.SelCount > 0 Then
    For i = 0 To List_Town.ListCount - 1
''''      If List_送至wn.List(i) = "" Then Go送至 Next送至wn '如果镇街名称为空，则跳过这个镇街
    
      If List_Town.Selected(i) = True Then 'Select
            
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).Town = List_Town.List(i) Then
          
            For k = 0 To List_County.ListCount - 1 '判断站点是否属于县区列表中的县区
              If AWSTInfo1(j + 1).County = List_County.List(k) Then
                
                If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then
                    If CheckHydro.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                    If CheckHydro.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
                End If
              End If
            Next k
          End If
        Next j
    
      ElseIf List_Town.Selected(i) = False Then  '不选择
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).Town = List_Town.List(i) Then
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then
              List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
    
      End If
      
''''NextTown:
    Next i

    GoTo ENDSelection
  End If
  
  
  '****** 2.然后判断区县：若选择了某些区县，则仅判断是否选择该些地区水文站
  If List_County.SelCount > 0 Then
    For i = 0 To List_County.ListCount - 1
      If List_County.Selected(i) = True Then 'Select
            
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).County = List_County.List(i) Then
                
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then '如果该entry是水文站，则选或不选
                If CheckHydro.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                If CheckHydro.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
    
      ElseIf List_County.Selected(i) = False Then  '不选择
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).County = List_County.List(i) Then
          
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then
              List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
      End If
      
    Next i
      
    GoTo ENDSelection
  End If
  
  
  '最后判断城市：若选择了某些地市，则仅判断是否选择该些地区水文站
  If List_City.SelCount > 0 Then
    For i = 0 To List_City.ListCount - 1
     
      If List_City.Selected(i) = True Then 'Select
            
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).City = List_City.List(i) Then
                
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then '如果该entry是水文站，则选或不选
                If CheckHydro.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                If CheckHydro.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          
          End If
        Next j
    
      ElseIf List_City.Selected(i) = False Then  '不选择
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).City = List_City.List(i) Then
          
            If AWSTInfo1(j + 1).Inland = 1 And AWSTInfo1(j + 1).Hydro = 1 Then '如果该entry是水文站，则不选
              List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
      End If
    Next i
      
    GoTo ENDSelection
  End If

 
ENDSelection:
 
  SelectSURF = SelectSURF_tmp: SelectAWST = SelectAWST_tmp
  

  Call Show_StaNum
  
  IgnoreListAWST = False ' List_StaAWST_Click() 恢复起效

End Sub


' 2026-8-7 筛选自动站点列表中的海洋区域站，并根据CheckInland状态，选或不选（只管海洋区域站，不操作其他类型站点）
Private Sub Checkocean_Click()
  If IgnoreCheckOcean Then Exit Sub

  Dim i%, j%, k%

  IgnoreListAWST = True ' List_StaAWST_Click() 动作失效
  SelectSURF_tmp = SelectSURF: SelectAWST_tmp = SelectAWST

  '2026-08-13 访客模式
  If GuestMode Then
    For j = 0 To List_StaAWST.ListCount - 1
      If AWSTInfo1(j + 1).Inland = 0 Then
        If CheckOcean.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
        If CheckOcean.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
      End If
    Next j
    GoTo ENDSelection
  End If

  '没是选择任何地区，则陆面区域站点不选
  If List_City.SelCount = 0 And List_County.SelCount = 0 And List_Town.SelCount = 0 Then
    GoTo ENDSelection
  End If
  
  '****** 1.一定要首先判断镇街：若选择了某些镇街，则仅判断是否选择该些地区水文站
  If List_Town.SelCount > 0 Then
    For i = 0 To List_Town.ListCount - 1

      If List_Town.Selected(i) = True Then 'Select
            
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).Town = List_Town.List(i) Then
          
            For k = 0 To List_County.ListCount - 1 '判断站点是否属于县区列表中的县区
              If AWSTInfo1(j + 1).County = List_County.List(k) Then
                
                If AWSTInfo1(j + 1).Inland = 0 Then
                  If CheckOcean.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
                  If CheckOcean.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
                End If
              End If
            Next k
          End If
        Next j
    
      ElseIf List_Town.Selected(i) = False Then  '不选择
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).Town = List_Town.List(i) Then
            If AWSTInfo1(j + 1).Inland = 0 Then
              List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
    
      End If

    Next i

    GoTo ENDSelection
  End If
  
  '****** 2.然后判断区县：若选择了某些地市，则仅判断是否选择该些地区陆面区域站
  If List_City.SelCount > 0 Then
    For i = 0 To List_City.ListCount - 1
      If List_City.Selected(i) = True Then 'Select
            
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).City = List_City.List(i) Then
                
            If AWSTInfo1(j + 1).Inland = 0 Then '如果该entry是海洋区域站，则选或不选
              If CheckOcean.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
              If CheckOcean.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
    
      ElseIf List_City.Selected(i) = False Then  '不选择
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).City = List_City.List(i) Then
          
            If AWSTInfo1(j + 1).Inland = 0 Then '如果该entry是海洋区域站，则不选
              List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
      End If

    Next i
      
    GoTo ENDSelection
  End If
  
  
  '3.最后判断城市：若选择了某些区县，则仅判断是否选择该些地区水文站
  If List_County.SelCount > 0 Then
    For i = 0 To List_County.ListCount - 1
    
      If List_County.Selected(i) = True Then 'Select
            
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).County = List_County.List(i) Then
                
            If AWSTInfo1(j + 1).Inland = 0 Then '如果该entry是海洋区域站，则选或不选
              If CheckOcean.Value = Checked Then List_StaAWST.Selected(j) = True: AWSTInfo1(j + 1).Selected = True
              If CheckOcean.Value = Unchecked Then List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          
          End If
        Next j
    
      ElseIf List_County.Selected(i) = False Then  '不选择
        For j = 0 To List_StaAWST.ListCount - 1
          If AWSTInfo1(j + 1).County = List_County.List(i) Then
          
            If AWSTInfo1(j + 1).Inland = 0 Then
              List_StaAWST.Selected(j) = False: AWSTInfo1(j + 1).Selected = False
            End If
          End If
        Next j
      End If
    Next i
     
    GoTo ENDSelection
  End If
 
ENDSelection:

  SelectSURF = SelectSURF_tmp: SelectAWST = SelectAWST_tmp
  

  Call Show_StaNum
  
  IgnoreListAWST = False ' List_StaAWST_Click() 恢复起效

End Sub

