VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form FrmMeteoSeasons 
   BackColor       =   &H8000000B&
   Caption         =   "Climate Season Onset Date Statistics"
   ClientHeight    =   13530
   ClientLeft      =   120
   ClientTop       =   345
   ClientWidth     =   28380
   Icon            =   "FrmMeteoSeasons.frx":0000
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
      Left            =   3600
      TabIndex        =   22
      Top             =   3600
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFEFDF&
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   28095
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Prior-Year"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   6480
         TabIndex        =   14
         Top             =   120
         Width           =   1575
         Begin VB.ComboBox ComboYear2 
            Height          =   300
            Left            =   180
            TabIndex        =   15
            Text            =   "yyyy"
            Top             =   540
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Start/End Date"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   4560
         TabIndex        =   11
         Top             =   120
         Width           =   1935
         Begin VB.ComboBox ComboEDate 
            Height          =   300
            Left            =   240
            TabIndex        =   13
            Text            =   "yyyy-mm-dd"
            Top             =   720
            Width           =   1500
         End
         Begin VB.ComboBox ComboBDate 
            Height          =   300
            Left            =   240
            TabIndex        =   12
            Text            =   "yyyy-mm-dd"
            Top             =   360
            Width           =   1500
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Threshold Temp (C)"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   2760
         TabIndex        =   9
         Top             =   120
         Width           =   1815
         Begin VB.ComboBox ComboThreshold 
            Height          =   300
            Index           =   1
            Left            =   540
            TabIndex        =   18
            Top             =   720
            Width           =   1140
         End
         Begin VB.ComboBox ComboThreshold 
            Height          =   300
            Index           =   0
            Left            =   540
            TabIndex        =   10
            Top             =   360
            Width           =   1140
         End
         Begin VB.Label Label1 
            BackColor       =   &H00DFEFDF&
            Caption         =   "<"
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
            Left            =   160
            TabIndex        =   17
            Top             =   760
            Width           =   255
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
            Left            =   160
            TabIndex        =   16
            Top             =   400
            Width           =   255
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Season"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   4
         Top             =   120
         Width           =   2655
         Begin VB.OptionButton OptionSeason 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Winter"
            Height          =   255
            Index           =   3
            Left            =   1320
            TabIndex        =   8
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton OptionSeason 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Autumn"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   7
            Top             =   720
            Width           =   1095
         End
         Begin VB.OptionButton OptionSeason 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Summer"
            Height          =   255
            Index           =   1
            Left            =   1320
            TabIndex        =   6
            Top             =   360
            Width           =   1215
         End
         Begin VB.OptionButton OptionSeason 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Spring"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   5
            Top             =   360
            Value           =   -1  'True
            Width           =   1095
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   8040
         TabIndex        =   3
         Top             =   120
         Width           =   19935
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid1 
      Height          =   10320
      Left            =   120
      TabIndex        =   0
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid2 
      Height          =   1380
      Left            =   120
      TabIndex        =   1
      Top             =   11880
      Width           =   28095
      _ExtentX        =   49556
      _ExtentY        =   2434
      _Version        =   393216
      FixedRows       =   0
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
      TabIndex        =   19
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid3 
      Height          =   3900
      Index           =   2
      Left            =   120
      TabIndex        =   20
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
      TabIndex        =   21
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
Attribute VB_Name = "FrmMeteoSeasons"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mLastRightClickGrid As Object '最近一次被点击（任意按钮）的表格；供mnuGridCopy_Click及ChartData查找图表数据源时使用

Private mEditGrid As Object '判断当前正在被TextEdit编辑的是哪个Grid（HFGrid1/HFGrid2/HFGrid3(Index)之一）

'2026-07-20
'季节窗口默认值：根据当前日期自动判定本次默认展示的季节，并给出对应的临界温度、统计时段。
'常年取自 SURF_CLI_T_DAY_NORM
Private Sub Form_Load()
  Screen.MousePointer = 11
  FrmMain.FormAdd

  ComboThreshold(0).AddItem "-9999"
  ComboThreshold(0).AddItem "10"
  ComboThreshold(0).AddItem "22"
  
  ComboThreshold(1).AddItem "10"
  ComboThreshold(1).AddItem "22"
  ComboThreshold(1).AddItem "999999"

  Dim i%
  ComboYear2.Clear
  For i = 1951 To Year(Date)
    ComboYear2.AddItem i
  Next i
  ComboYear2.Text = Year(Date) - 1

  Dim defaultSeason%
  Select Case Month(Date)
    Case 2, 3
      defaultSeason = 0 'Spring
    Case 4, 5, 6, 7, 8
      defaultSeason = 1 'Summer
    Case 9, 10, 11
      defaultSeason = 2 'Autumn
    Case Else '11,12,1
      defaultSeason = 3 'Winter
  End Select

  OptionSeason(defaultSeason).Value = True '触发 Option季节_Click 以套用该季节的默认阈值与统计时段
  
  '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
  If GuestMode Then
    Call ImportData.LoadGuestCSV(FrmMeteoSeasons.HFGrid1, FrmMeteoSeasons.HFGrid2, "DataMeteoSeasons.csv", 3)
  End If

  Screen.MousePointer = 1

End Sub

'2026-08-19 v2.14 界面自适应：主窗体尺寸变化时，让本窗体的表格和选项区跟着调整
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

Private Sub OptionSeason_Click(Index As Integer)
  Call ApplySeasonDefault(Index)
  Call TableIni_MeteoSeasons(FrmMeteoSeasons)

End Sub

'2026-07-20 根据所选季节，套用该季节惯常起始月对应的默认阈值与统计时段
Private Sub ApplySeasonDefault(SeasonIdx As Integer)
  Dim baseY%
  baseY = Year(Date)
  If SeasonIdx = 3 And Month(Date) = 1 Then baseY = baseY - 1

  Select Case SeasonIdx
    Case 0 'Spring：临界10℃
      ComboThreshold(0).Text = "10": ComboThreshold(1).Text = "999999"
      ComboBDate.Text = Format(DateSerial(baseY, 1, 1), "yyyy-mm-dd")
      ComboEDate.Text = Format(DateSerial(baseY, 5, 31), "yyyy-mm-dd")
    Case 1 'Summer：临界22℃
      ComboThreshold(0).Text = "22": ComboThreshold(1).Text = "999999"
      ComboBDate.Text = Format(DateSerial(baseY, 3, 1), "yyyy-mm-dd")
      ComboEDate.Text = Format(DateSerial(baseY, 8, 31), "yyyy-mm-dd")
    Case 2 'Autumn：临界22℃
       ComboThreshold(0).Text = "-9999": ComboThreshold(1).Text = "22"
      ComboBDate.Text = Format(DateSerial(baseY, 9, 1), "yyyy-mm-dd")
      ComboEDate.Text = Format(DateSerial(baseY, 11, 30), "yyyy-mm-dd")
    Case 3 'Winter：临界10℃
       ComboThreshold(0).Text = "-9999": ComboThreshold(1).Text = "10"
      ComboBDate.Text = Format(DateSerial(baseY, 12, 1), "yyyy-mm-dd")
      ComboEDate.Text = Format(DateAdd("d", -1, DateSerial(baseY + 1, 3, 1)), "yyyy-mm-dd")
  End Select
End Sub


Public Sub CmdMeteoSeasons()

  Dim SeasonIdx%, i%, j%
  For i = 0 To 3
    If OptionSeason(i).Value = True Then SeasonIdx = i: Exit For
  Next i

  If Not IsNumeric(ComboThreshold(0).Text) Or Not IsNumeric(ComboThreshold(1).Text) Then
    MsgBox "Threshold temperature must be numeric", vbInformation, "Notice"
    Exit Sub
  End If
  If Not IsDate(ComboBDate.Text) Or Not IsDate(ComboEDate.Text) Then
    MsgBox "Invalid start/end date", vbInformation, "Notice"
    Exit Sub
  End If
  If CDate(ComboBDate.Text) > CDate(ComboEDate.Text) Then
    MsgBox "Invalid start/end date", vbInformation, "Notice"
    Exit Sub
  End If
  If Not IsNumeric(ComboYear2.Text) Then
    MsgBox "Prior-year value must be a valid year", vbInformation, "Notice"
    Exit Sub
  End If


  Screen.MousePointer = 11
  FrmMain.StatusBar1.Panels(1).Text = "Calculating season onset dates"
  DoEvents
    
  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart

  Call TableIni_MeteoSeasons(FrmMeteoSeasons) '2026-07-20 打开时即按当前已选站点填好站号/站名，起始日留空，与其他窗体一致

  Dim threshold_min!, threshold_max!
  threshold_min = CSng(ComboThreshold(0).Text)   '≥
  threshold_max = CSng(ComboThreshold(1).Text)   '<

  Dim BDate1 As Date, EDate1 As Date '当年日期
  BDate1 = CDate(ComboBDate.Text)
  EDate1 = CDate(ComboEDate.Text)
  
  Dim crossYear As Boolean
  If Year(EDate1) > Year(BDate1) Then
    crossYear = True
  End If

  Dim BMMDD, EMMDD As String
  BMMDD = Format(BDate1, "mm-dd")
  EMMDD = Format(EDate1, "mm-dd")
  
  Dim Year2%, BDate2 As Date, EDate2 As Date '某年日期
  Dim BDateNorm As Date, EDateNorm As Date '常年日期
  
  Year2 = CInt(ComboYear2.Text)

  If crossYear = False Then
    BDate2 = CDate(CStr(Year2) & "-" & BMMDD)
    EDate2 = CDate(CStr(Year2) & "-" & EMMDD)
    
    BDateNorm = CDate("2002-" & BMMDD)
    EDateNorm = CDate("2002-" & EMMDD)
    
  Else
    BDate2 = CDate(CStr(Year2) & "-" & BMMDD)
    EDate2 = CDate(CStr(Year2 + 1) & "-" & EMMDD)
    
    BDateNorm = CDate("2001-" & BMMDD)
    EDateNorm = CDate("2002-" & EMMDD)
  End If


  ReDim SS_STA(1 To StaNum)
  For i = 1 To StaNum
    SS_STA(i).stacode = StaInfo(i).stacode
    SS_STA(i).staname = StaInfo(i).staname
    SS_STA(i).OnsetNorm = CDate("1899-09-09")
    SS_STA(i).T_IdxNorm = 0
    
    SS_STA(i).OnsetCurr = CDate("1899-09-09")
    SS_STA(i).T_IdxCurr = 0
    
    SS_STA(i).OnsetComp = CDate("1899-09-09")
    SS_STA(i).T_IdxComp = 0
    
    SS_STA(i).diffNorm = 0
    SS_STA(i).diffComp = 0
    
    SS_STA(i).Town = StaInfo(i).Town
    SS_STA(i).County = StaInfo(i).County
    SS_STA(i).City = StaInfo(i).City
    
  Next i
  
  
  Dim iDays_Curr%, iDays_Comp%, iDays_Norm%
  iDays_Curr = EDate1 - BDate1 + 1
  iDays_Comp = EDate2 - BDate2 + 1
  iDays_Norm = EDateNorm - BDateNorm + 1
  
  ReDim T5d_Curr(1 To StaNum, 1 To iDays_Curr - 4)
  ReDim T5d_Comp(1 To StaNum, 1 To iDays_Comp - 4)
  ReDim T5d_Norm(1 To StaNum, 1 To iDays_Norm - 4)
  For i = 1 To StaNum
    For j = 1 To iDays_Curr - 4
      T5d_Curr(i, j).T = -9999
      T5d_Curr(i, j).ddate = CDate("1899-09-09")
    Next j
    For j = 1 To iDays_Comp - 4
      T5d_Comp(i, j).T = -9999
      T5d_Comp(i, j).ddate = CDate("1899-09-09")
    Next j
    For j = 1 To iDays_Norm - 4
      T5d_Norm(i, j).T = -9999
      T5d_Norm(i, j).ddate = CDate("1899-09-09")
    Next j
  Next i
  
  ReDim T_Curr(1 To StaNum, 1 To iDays_Curr)
  ReDim T_Comp(1 To StaNum, 1 To iDays_Comp)
  ReDim T_Norm(1 To StaNum, 1 To iDays_Norm)
  For i = 1 To StaNum
    For j = 1 To iDays_Curr
      T_Curr(i, j).T = -9999
      T_Curr(i, j).ddate = CDate("1899-09-09")
    Next j
    For j = 1 To iDays_Comp
      T_Comp(i, j).T = -9999
      T_Comp(i, j).ddate = CDate("1899-09-09")
    Next j
    For j = 1 To iDays_Norm
      T_Norm(i, j).T = -9999
      T_Norm(i, j).ddate = CDate("1899-09-09")
    Next j
  Next i

  ORAConn2.Open
  ORAConn2.CursorLocation = 3

'  Call QuerySeasonOnset(Threshold, AboveThreshold, BDateCurr, EDateCurr, BDateComp, EDateComp)
  
  ' 先统计常年季节起始日，用于当年/某年季节起始日的二次判断
  Call QueryT5dFromNorm(BDateNorm, EDateNorm)
  Call QueryTFromNorm(BDateNorm, EDateNorm)
  Call FindSteadyThresholdDateNorm("Norm", T5d_Norm, T_Norm, iDays_Norm - 4, threshold_min, threshold_max)


  Call QueryT5dFromDay("Curr", BDate1, EDate1)
  Call QueryTFromDay("Curr", BDate1, EDate1)
  Call FindSteadyThresholdDate("Curr", T5d_Curr, T_Curr, iDays_Curr - 4, threshold_min, threshold_max)
  
  Call QueryT5dFromDay("Comp", BDate2, EDate2)
  Call QueryTFromDay("Comp", BDate2, EDate2)
  Call FindSteadyThresholdDate("Comp", T5d_Comp, T_Comp, iDays_Comp - 4, threshold_min, threshold_max)


  ORAConn2.Close
  Call TimeStop
  frmWait.Hide: DoEvents
  

  For i = 1 To StaNum
    SS_STA(i).diffNorm = SS_STA(i).T_IdxCurr - SS_STA(i).T_IdxNorm
    SS_STA(i).diffComp = SS_STA(i).T_IdxCurr - SS_STA(i).T_IdxComp
  Next i

  
  '************************************* 2、计算所是行和所是列的统计值  ***********************************
  If FlagZone = "STA" Then '站点尺度
      '统计所是行的平均/最大/最小Value
      Call Stat_ROWS_Seasons(SS_STA, BDateNorm, BDate1, BDate2)
      '填充上方表格
      Call HFGrid1_Fill_Seasons(FrmMeteoSeasons, SS_STA)
      '填充下方表格
      Call HFGrid2_Fill_Seasons(FrmMeteoSeasons, ROWS_Stat_Seasons)
    
  Else
      '分区域统计所是行列值
      Call Stat_ZONE_Seasons(SS_STA, BDateNorm, BDate1, BDate2)
      '填充上/中/下方表格值
      Call HFGrid3_fill_Seasons(FrmMeteoSeasons)
   
  End If

  Screen.MousePointer = 1


End Sub

Public Sub CmdOutput_MeteoSeasons()
  Call Output_HFGrid(FrmMeteoSeasons, "")
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FrmMain.FormDel
End Sub


'********************对HFGrid中的数据进行排序、右键复制***************************
Private Sub HFGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid1
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols(FrmMeteoSeasons.HFGrid1, shift)
End Sub

Private Sub mnuGridCopy_Click()
  Call CopyGridSelectionToClipboard(mLastRightClickGrid)
End Sub


Private Sub HFGrid3_MouseUp(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid3(Index)
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols(FrmMeteoSeasons.HFGrid3(Index), shift)
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


'********************利用HFGrid中的数据绘制地图***************************
Public Sub CmdMap_MeteoSeasons()
    Call Map_Value(FrmMeteoSeasons, "BYR")
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


