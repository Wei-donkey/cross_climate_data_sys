VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmMeteoDay 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Daily Data Query"
   ClientHeight    =   13530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   28380
   Icon            =   "FrmMeteoDay.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13530
   ScaleWidth      =   28380
   WindowState     =   2  'Maximized
   Begin VB.TextBox TextEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0E0FF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   3360
      TabIndex        =   0
      Top             =   2400
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid1 
      Height          =   10320
      Left            =   120
      TabIndex        =   22
      Top             =   1560
      Width           =   28095
      _ExtentX        =   49556
      _ExtentY        =   18203
      _Version        =   393216
      BackColorFixed  =   14675935
      BackColorBkg    =   16777215
      GridColor       =   12632256
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFEFDF&
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   28095
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Data State"
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   14160
         TabIndex        =   20
         Top             =   120
         Width           =   3375
         Begin VB.CheckBox CheckXtrmYear 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Year"
            Height          =   255
            Left            =   2520
            TabIndex        =   26
            Top             =   840
            Width           =   735
         End
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "ExtrMin"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   25
            Top             =   840
            Width           =   1095
         End
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "ExtrMax"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   5
            Left            =   240
            TabIndex        =   24
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.ComboBox ComboDataType 
            Height          =   300
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   23
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "ElementSelect"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   15
         Top             =   120
         Width           =   12135
         Begin VB.ListBox List_Col 
            Columns         =   9
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   840
            Left            =   120
            TabIndex        =   19
            Top             =   240
            Width           =   11895
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Statistic"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   17520
         TabIndex        =   12
         Top             =   120
         Width           =   1455
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Avg"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   14
            Top             =   360
            Value           =   -1  'True
            Width           =   975
         End
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Cumula"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Width           =   975
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Highlight Results"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   18960
         TabIndex        =   6
         Top             =   120
         Width           =   2175
         Begin VB.CheckBox CheckHLight 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Yes"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   1440
            TabIndex        =   9
            Top             =   720
            Width           =   675
         End
         Begin VB.ComboBox ComboMin_HL 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   480
            TabIndex        =   8
            Top             =   360
            Width           =   855
         End
         Begin VB.ComboBox ComboMax_HL 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   480
            TabIndex        =   7
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label13 
            BackColor       =   &H00DFEFDF&
            Caption         =   "≥"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   180
            Left            =   195
            TabIndex        =   11
            Top             =   435
            Width           =   255
         End
         Begin VB.Label Label14 
            BackColor       =   &H00DFEFDF&
            Caption         =   "≤"
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   195
            Left            =   195
            TabIndex        =   10
            Top             =   795
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Start/End Date"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   12240
         TabIndex        =   2
         Top             =   120
         Width           =   1935
         Begin VB.ComboBox ComboBDate 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   4
            Text            =   "yyyy-mm-dd"
            Top             =   360
            Width           =   1455
         End
         Begin VB.ComboBox ComboEDate 
            BeginProperty Font 
               Name            =   "宋体"
               Size            =   9.75
               Charset         =   134
               Weight          =   400
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Left            =   240
            TabIndex        =   3
            Text            =   "yyyy-mm-dd"
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   21120
         TabIndex        =   5
         Top             =   120
         Width           =   6855
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid2 
      Height          =   1380
      Left            =   120
      TabIndex        =   21
      Top             =   11880
      Width           =   28095
      _ExtentX        =   49556
      _ExtentY        =   2434
      _Version        =   393216
      BackColorFixed  =   14675935
      BackColorBkg    =   16777215
      GridColor       =   12632256
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid3 
      Height          =   3900
      Index           =   1
      Left            =   120
      TabIndex        =   16
      Top             =   1560
      Width           =   28095
      _ExtentX        =   49556
      _ExtentY        =   6879
      _Version        =   393216
      BackColorFixed  =   14675935
      BackColorBkg    =   16777215
      GridColor       =   12632256
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   15240
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid3 
      Height          =   3900
      Index           =   2
      Left            =   120
      TabIndex        =   17
      Top             =   5460
      Width           =   28095
      _ExtentX        =   49556
      _ExtentY        =   6879
      _Version        =   393216
      BackColorFixed  =   14675935
      BackColorBkg    =   16777215
      GridColor       =   12632256
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid3 
      Height          =   3900
      Index           =   3
      Left            =   120
      TabIndex        =   18
      Top             =   9360
      Width           =   28095
      _ExtentX        =   49556
      _ExtentY        =   6879
      _Version        =   393216
      BackColorFixed  =   14675935
      BackColorBkg    =   16777215
      GridColor       =   12632256
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.Menu mnuGridPopup 
      Caption         =   "GridPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuGridCopy 
         Caption         =   "Copy Selection"
      End
   End
End
Attribute VB_Name = "FrmMeteoDay"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public BDate$, EDate$, idays%

Public mLastRightClickGrid As Object '最近一次被点击（任意按钮）的表格；供mnuGridCopy_Click及ChartData查找图表数据源时使用

Private mEditGrid As Object '判断当前正在被TextEdit编辑的是哪个Grid（HFGrid1/HFGrid2/HFGrid3(Index)之一）
Dim SelField_Initial$ '选择要素的初始名
Dim SelField_Restat$ '选择要素的再统计名

'Dim Sel字段_Day$ '选择要素的日间隔字段
Dim SelField_Xtrm$ '选择要素的极值表字段
Dim SelField_Norm$ '选择要素的平均态字段

Public FlagStat$ '统计量标识

Dim BMMDD$, EMMDD$ '用户自定义的查询起止月日

Dim i%, j%, k%, l%
'Dim TextMin!, TextMax!

Public DataType$ '查询的数据类型：日间隔数据、平均态、极端态
Public XtrmYear As Boolean '是否显示极值出现年份
Dim XtrmYear_STA() As Integer '各站点的极值出现年份
Dim crossYear As Boolean '查询时段是否跨年度


Private Sub CheckHLight_Click()
  If CheckHLight.Value = Checked Then
    ComboMin_HL.Enabled = True: ComboMax_HL.Enabled = True
  ElseIf CheckHLight.Value = Unchecked Then
    ComboMin_HL.Enabled = False: ComboMax_HL.Enabled = False
  End If
End Sub

Private Sub CheckXtrmYear_Click()
  If CheckXtrmYear.Value = Checked Then
    XtrmYear = True
  ElseIf CheckXtrmYear.Value = Unchecked Then
    XtrmYear = False
  End If
  
'  Call ITimesInitial(FrmMeteoDay)
  Call TableIni_Meteo(FrmMeteoDay, BDate, idays) ', SelField_Initial)
End Sub

Private Sub Form_Load()
  Dim i%
  Screen.MousePointer = 11
  
  FrmMain.FormAdd
  
  DataType = "Data"
  
  For i = 0 To 30
    ComboBDate.AddItem Format(Date - 30 + i, "yyyy-mm-dd")
    ComboEDate.AddItem Format(Date - 30 + i, "yyyy-mm-dd")
  Next i
  ComboBDate.Text = Format(Date - 1, "yyyy-mm-01")
  ComboEDate.Text = Format(Date - 1, "yyyy-mm-dd")
  
  Call ITimesInitial(FrmMeteoDay)
  
  If OptionStat(5).Value = True Then FlagStat = "MAX"
  If OptionStat(6).Value = True Then FlagStat = "MIN"
  If CheckXtrmYear.Value = Checked Then XtrmYear = True
  
  
  If FlagStat = "MAX" Or FlagStat = "MIN" Then
    SelField_Xtrm = SelField_Restat & "_" & FlagStat
  End If
  
  ComboDataType.AddItem "Daily Obs Value": ComboDataType.AddItem "Normals": ComboDataType.AddItem "Historical Extremes"
  ComboDataType.Text = ComboDataType.List(0)
  
  '****************************2022-03-08从ini配置文件中读取字段名*****************************
  Call ReadColInfo("COLS_Initial.ini")
  COL_Initial = COL_tmp
 
  Call ReadColInfo("COLS_ReStat.ini")
  COL_Restat = COL_tmp
 
  List_Col.Clear
  For i = 1 To UBound(COL_Initial)
    List_Col.AddItem COL_Initial(i).Col_name
  Next i
  
  List_Col.Selected(0) = True
  SelField_Initial = COL_Initial(1).Col_ID
  SelField_Restat = COL_Restat(1).Col_ID
  
  Screen.MousePointer = 1

End Sub

'2026-08-19 v2.14 界面自适应：主窗体尺寸变化时，让本窗体的网格和选项区跟着调整
Private Sub Form_Resize()
  Call ApplyLayout
End Sub

Public Sub ApplyLayout()
  On Error GoTo ResizeError
  If Me.WindowState = 1 Then Exit Sub '最小化时不处理

  Dim bottomAnchorTop As Long
  bottomAnchorTop = ResizeBottomAnchoredGrid(HFGrid2, Me.ScaleHeight, Me.ScaleWidth)
  Call ResizeFillGrid(HFGrid1, bottomAnchorTop, Me.ScaleWidth)
  Call ResizeStackedGridArray(Array(HFGrid3(1), HFGrid3(2), HFGrid3(3)), Me.ScaleHeight - MDI_RESIZE_BOTTOM_MARGIN, Me.ScaleWidth)
  Call ResizeFrame1Width(Frame1, Me.ScaleWidth)
  Call ResizeChildFrameWidth(Frame11, Frame1, 120)

  Exit Sub
ResizeError:
  Resume Next
End Sub


'*********************************************单要素逐日资料统计>>>开始*****************************************
'*********************逐日资料查询（可查询天气过程，如连续高温、低温、雨日等）**********************************
Public Sub CmdMeteoDay()
  
  Screen.MousePointer = 11
  
  If ComboDataType.Text = ComboDataType.List(0) Then
    DataType = "Data"
  ElseIf ComboDataType.Text = ComboDataType.List(1) Then
    DataType = "Norm"
  ElseIf ComboDataType.Text = ComboDataType.List(2) Then
    DataType = "Xtrm"
  End If
  
  Call ITimesInitial(FrmMeteoDay)
  If FlagInputErr = True Then Screen.MousePointer = 1: Exit Sub

  
  If idays <= 0 Then MsgBox "Invalid start/end date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  If idays > 366 Then MsgBox "Date range cannot exceed 366 days", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
 
  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart
  
  Call TableIni_Meteo(FrmMeteoDay, BDate, idays)
  
  '****************************2022-03-08*****************************
  For i = 0 To List_Col.ListCount - 1
    If List_Col.Selected(i) = True Then
      SelField_Initial = COL_Initial(i + 1).Col_ID
      SelField_Restat = COL_Restat(i + 1).Col_ID
      Exit For
    End If
  Next i
  
  If OptionStat(5).Value = True Then FlagStat = "MAX"
  If OptionStat(6).Value = True Then FlagStat = "MIN"
  
  If FlagStat = "MAX" Or FlagStat = "MIN" Then
    SelField_Xtrm = SelField_Restat & "_" & FlagStat
  End If
  
  BMMDD = Right(BDate, 5) '开始月日
  EMMDD = Right(EDate, 5) '结束月日
  
  If Left(EDate, 4) - Left(BDate, 4) = 0 Then '同年度
    crossYear = False
  ElseIf Left(EDate, 4) - Left(BDate, 4) = 1 Then '跨年度：年份数少一年
    crossYear = True
  End If
  
''''  '********2021-9-14 其他所是年份均不查02-29*************
''''  If BMMDD = "02-29" Then BMMDD = "02-28"
''''  If EMMDD = "02-29" Then EMMDD = "02-28"
  
  ORAConn2.Open
  ORAConn2.CursorLocation = 3
  
'************************************* 1、查询所是站点多日值  ***********************************
  ReDim DAY_STA(1 To StaNum, 1 To idays) '存放各站点、各日值
  If XtrmYear = True Then ReDim XtrmYear_STA(1 To StaNum, 1 To idays)
  For i = 1 To StaNum
    For j = 1 To idays
      DAY_STA(i, j) = -9999
      If XtrmYear = True Then XtrmYear_STA(i, j) = 1899 '默认极值年份1899年
    Next j
  Next i
  
  If ComboDataType.Text = ComboDataType.List(0) Then
    DataType = "Data"
  ElseIf ComboDataType.Text = ComboDataType.List(1) Then
    DataType = "Norm"
  ElseIf ComboDataType.Text = ComboDataType.List(2) Then
    DataType = "Xtrm"
  End If
  
  If DataType = "Data" Then
    Call QueryMulDayData(SelField_Initial, BDate, EDate, DAY_STA())
  ElseIf DataType = "Norm" Then
    Call QueryMulDayRestat(SelField_Restat, BMMDD, EMMDD, DataType, crossYear, DAY_STA(), XtrmYear_STA(), XtrmYear)
  ElseIf DataType = "Xtrm" Then
    Call QueryMulDayRestat(SelField_Xtrm, BMMDD, EMMDD, DataType, crossYear, DAY_STA(), XtrmYear_STA(), XtrmYear)
  End If


'************************************* 情形一：查询结果为站点尺度  ***********************************
  If FlagZone = "STA" Then '站点尺度
  
'************************************* 2、计算所是行和所是列的统计值  ***********************************
    '统计所是行的平均/最大/最小Value
    Call Stat_ROWS(DAY_STA, idays, StaNum)
    '统计所是列的平均（累计）/最大/最小Value
    If OptionStat(1).Value = True Then
      Call Stat_COLS(DAY_STA, idays, StaNum, "AVE")
    ElseIf OptionStat(4).Value = True Then
      Call Stat_COLS(DAY_STA, idays, StaNum, "SUM")
    End If
    
    '统计列的3列统计值的平均/最大/最小Value
    Call Stat_COLS_ROWS(StaNum)
    
'************************************* 3、填充上方表格  ***********************************
    Call HFGrid1_Fill(FrmMeteoDay, DAY_STA, idays, XtrmYear_STA)
'************************************* 4、填充下方表格 ***********************************
    Call HFGrid2_Fill(FrmMeteoDay, ROWS_Stat, idays)
  
  
'************************************* 情形二：查询结果为区域尺度  ***********************************
  Else
  
  
'************************************* 2、分区域统计所是行列值  ***********************************
    Call Stat_ZONE(DAY_STA, idays)
  
  
    '2025-06-01：统计所是列的 结果三列：平均（累计）/最大/最小Value：站点尺度
    If OptionStat(1).Value = True Then
      Call Stat_COLS(DAY_STA, idays, StaNum, "AVE")
    ElseIf OptionStat(4).Value = True Then
      Call Stat_COLS(DAY_STA, idays, StaNum, "SUM")
    End If
  
    '2025-06-01：统计结果3列的区域平均/最大/最小（2025-06-01新增），输入的数据为COL_Stat(i,1)、COL_Stat(i,2)、COL_Stat(i,3)
    Call Stat_COLS_ZONE(COLS_Stat)
  
'************************************* 3、填充上/中/下方表格值  ***********************************
      Call HFGrid3_fill_HorDay(FrmMeteoDay, Zone_AVG, Zone_COLS_AVG, idays, 1)
      Call HFGrid3_fill_HorDay(FrmMeteoDay, Zone_MAX, Zone_COLS_MAX, idays, 2)
      Call HFGrid3_fill_HorDay(FrmMeteoDay, Zone_MIN, Zone_COLS_MIN, idays, 3)

  End If


  ORAConn2.Close
  Call TimeStop
  frmWait.Hide: DoEvents
  
  Screen.MousePointer = 1
  
End Sub
'*********************************************单要素逐日资料统计>>>结束*****************************************


Private Sub ComboDataType_Click()
  If ComboDataType.Text = ComboDataType.List(0) Then
    DataType = "Data"
  ElseIf ComboDataType.Text = ComboDataType.List(1) Then
    DataType = "Norm"
  ElseIf ComboDataType.Text = ComboDataType.List(2) Then
    DataType = "Xtrm"
  End If
  
  CheckXtrmYear.Value = Unchecked
  
  Call Option_MeteoDay(SelField_Initial, SelField_Restat)
  Call TableIni_Meteo(FrmMeteoDay, BDate, idays)
End Sub


'********************对HFGrid1中的数据进行排序***************************
Private Sub HFGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid1
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols_Rows(FrmMeteoDay.HFGrid1, shift)
End Sub

Private Sub mnuGridCopy_Click()
  Call CopyGridSelectionToClipboard(mLastRightClickGrid)
End Sub

Private Sub HFGrid2_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid2
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Rows(FrmMeteoDay.HFGrid2, shift)
End Sub

'********************对HFGrid3中的数据进行排序***************************
Private Sub HFGrid3_MouseUp(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid3(Index)
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols_Rows(FrmMeteoDay.HFGrid3(Index), shift)
End Sub


'********************利用HFGrid中的数据绘制地图***************************
Public Sub CmdMap_MeteoDay()
  If List_Col.Selected(4) = True Or List_Col.Selected(5) = True Or List_Col.Selected(6) = True Or List_Col.Selected(7) = True Or List_Col.Selected(14) = True Or List_Col.Selected(15) = True Then
    Call Map_Value(FrmMeteoDay, "RYB")
  Else
    Call Map_Value(FrmMeteoDay, "BYR")
  End If
End Sub


'********************保存窗体内可见的HFGrid***************************
Public Sub CmdOutput_MeteoDay()
  Call Output_HFGrid_Day(FrmMeteoDay, SelField_Initial)
End Sub



Private Sub List_Col_Click()
  For i = 0 To List_Col.ListCount - 1
    If List_Col.Selected(i) = True Then
      SelField_Initial = COL_Initial(i + 1).Col_ID
      SelField_Restat = COL_Restat(i + 1).Col_ID
      Exit For
    End If
  Next i
  Call Option_MeteoDay(SelField_Initial, SelField_Restat)
  Call TableIni_Meteo(FrmMeteoDay, BDate, idays)
  CheckHLight.Value = Unchecked
  
  '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
  If GuestMode And List_Col.Selected(0) = True Then
    Call ImportData.LoadGuestCSV(FrmMeteoDay.HFGrid1, FrmMeteoDay.HFGrid2, "DataMeteoDay.csv", 3)
  End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FrmMain.FormDel
  FrmMain.StatusBar1.Panels(1).Text = ""
  Unload Me
End Sub

Private Sub OptionStat_Click(Index As Integer)
  If FlagZone = "STA" Then
    If OptionStat(1).Value = True Then HFGrid1.TextMatrix(0, 3 + idays) = "Average"
    If OptionStat(4).Value = True Then HFGrid1.TextMatrix(0, 3 + idays) = "Cumulative"
  Else
    If OptionStat(1).Value = True Then
      HFGrid3(1).TextMatrix(0, 3 + idays) = "Average"
      HFGrid3(2).TextMatrix(0, 3 + idays) = "Average"
      HFGrid3(3).TextMatrix(0, 3 + idays) = "Average"
    End If
    If OptionStat(4).Value = True Then
      HFGrid3(1).TextMatrix(0, 3 + idays) = "Cumulative"
      HFGrid3(2).TextMatrix(0, 3 + idays) = "Cumulative"
      HFGrid3(3).TextMatrix(0, 3 + idays) = "Cumulative"
    End If
  End If
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


'鼠标滚轮驱动表格上下滚动
Private Sub HFGrid3_GotFocus(Index As Integer)
  On Error Resume Next
  Oldwinproc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
  SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub
    
'鼠标滚轮驱动表格上下滚动
Private Sub HFGrid3_LostFocus(Index As Integer)
  On Error Resume Next
  SetWindowLong Me.hWnd, GWL_WNDPROC, Oldwinproc
End Sub


'********************双击单元格进行修改***************************
Private Sub HFGrid1_DblClick()
  Call GridEdit(HFGrid1)
End Sub

Private Sub HFGrid2_DblClick()
  Call GridEdit(HFGrid2)
End Sub

Private Sub HFGrid3_DblClick(Index As Integer)
  Call GridEdit(HFGrid3(Index))
End Sub

Private Sub GridEdit(ByVal Grd As Object)
  Set mEditGrid = Grd
  With Grd
    TextEdit.Left = .Left + .ColPos(.Col)
    TextEdit.Top = .Top + .RowPos(.Row)
    If .Appearance = flex3D Then
      TextEdit.Left = TextEdit.Left + 2 * Screen.TwipsPerPixelX
      TextEdit.Top = TextEdit.Top + 2 * Screen.TwipsPerPixelY
    End If
    TextEdit.Width = .ColWidth(.Col)
    TextEdit.Height = .RowHeight(.Row)
    TextEdit.Text = .Text
  End With
  TextEdit.Visible = True
  TextEdit.SelLength = Len(TextEdit.Text)
  TextEdit.SetFocus
End Sub

Private Sub TextEdit_Change()
  If Not mEditGrid Is Nothing Then mEditGrid.Text = TextEdit.Text
End Sub

Private Sub TextEdit_KeyPress(KeyAscii As Integer)
  If KeyAscii = vbKeyReturn Then
    TextEdit.Visible = False
    If Not mEditGrid Is Nothing Then mEditGrid.SetFocus
  End If
End Sub

Private Sub TextEdit_LostFocus()
  TextEdit.Visible = False
End Sub
