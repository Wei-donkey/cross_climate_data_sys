VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmMeteoColdAir 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "Cold Air Statistics"
   ClientHeight    =   13530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   28380
   Icon            =   "FrmMeteoColdAir.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   Moveable        =   0   'False
   ScaleHeight     =   13530
   ScaleWidth      =   28380
   WindowState     =   2  'Maximized
   Begin VB.TextBox TextEdit 
      Appearance      =   0  'Flat
      BackColor       =   &H00F0E0FF&
      BorderStyle     =   0  'None
      Height          =   270
      Left            =   2400
      TabIndex        =   7
      Top             =   3240
      Visible         =   0   'False
      Width           =   975
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid2 
      Height          =   1380
      Left            =   120
      TabIndex        =   20
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFEFDF&
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   28095
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Start/End Date"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   2280
         TabIndex        =   4
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
            TabIndex        =   6
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
            TabIndex        =   5
            Text            =   "yyyy-mm-dd"
            Top             =   720
            Width           =   1455
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Objective"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   2175
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Single Event"
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
            Index           =   2
            Left            =   240
            TabIndex        =   3
            Top             =   360
            Width           =   1695
         End
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "All Events"
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
            TabIndex        =   2
            Top             =   720
            Width           =   1815
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Level Filter"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   4200
         TabIndex        =   9
         Top             =   120
         Width           =   5655
         Begin VB.CheckBox Check_lvl 
            BackColor       =   &H00DFEFDF&
            Caption         =   "None"
            Height          =   255
            Index           =   0
            Left            =   4200
            TabIndex        =   21
            Top             =   720
            Width           =   720
         End
         Begin VB.CheckBox Check_lvl 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Moderate Cold Air"
            Height          =   255
            Index           =   2
            Left            =   240
            TabIndex        =   13
            Top             =   720
            Value           =   1  'Checked
            Width           =   1935
         End
         Begin VB.CheckBox Check_lvl 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Cold Wave"
            Height          =   255
            Index           =   4
            Left            =   240
            TabIndex        =   12
            Top             =   360
            Value           =   1  'Checked
            Width           =   1215
         End
         Begin VB.CheckBox Check_lvl 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Strong Cold Air"
            Height          =   255
            Index           =   3
            Left            =   1560
            TabIndex        =   11
            Top             =   360
            Value           =   1  'Checked
            Width           =   1800
         End
         Begin VB.CheckBox Check_lvl 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Weak Cold Air"
            Height          =   255
            Index           =   1
            Left            =   2400
            TabIndex        =   10
            Top             =   720
            Value           =   1  'Checked
            Width           =   1800
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   7200
         TabIndex        =   8
         Top             =   120
         Width           =   20775
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid1 
      Height          =   10320
      Left            =   120
      TabIndex        =   17
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
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   13800
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid3 
      Height          =   3900
      Index           =   1
      Left            =   120
      TabIndex        =   14
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
      TabIndex        =   15
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
      TabIndex        =   16
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid4 
      Height          =   10320
      Left            =   120
      TabIndex        =   18
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid5 
      Height          =   1380
      Left            =   120
      TabIndex        =   19
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
   Begin VB.Menu mnuGridPopup 
      Caption         =   "GridPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuGridCopy 
         Caption         =   "Copy Selection"
      End
   End
End
Attribute VB_Name = "FrmMeteoColdAir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mLastRightClickGrid As Object '最近一次被点击（任意按钮）的表格；供mnuGridCopy_Click及ChartData查找图表数据源时使用
Private mEditGrid As Object '判断当前正在被TextEdit编辑的是哪个Grid（HFGrid1/HFGrid2/HFGrid3(Index)之一）

Dim SelField_Initial$ '选择要素的初始名
Dim i%, j%, k%
Public FlagStat '统计类型：多次过程Multi、单次过程Single

Dim BYYYY$, BMM$, BDD$, EYYYY$, EMM$, EDD$ '用户自定义的查询起止年、月、日

Public BDate$, EDate$, idays% '用户自定义的当年查询起始日期
Dim templevel_air As Integer  'Cold Air OutbreakLevel：0-None，1-弱，2-中，3-强，4-Cold Wave


Private Sub OptionStat_Click(Index As Integer)
  If Index = 1 Then
    FlagStat = "Multi"
  
    If FlagZone = "STA" Then '输出站点结果
      HFGrid1.Visible = False
      HFGrid2.Visible = False
      HFGrid4.Visible = True
      HFGrid5.Visible = True
      HFGrid3(1).Visible = False
      HFGrid3(2).Visible = False
      HFGrid3(3).Visible = False
    Else '输出区域统计结果
      HFGrid1.Visible = False
      HFGrid2.Visible = False
      HFGrid4.Visible = False
      HFGrid5.Visible = False
      HFGrid3(1).Visible = True
      HFGrid3(2).Visible = True
      HFGrid3(3).Visible = True
    End If
  
  ElseIf Index = 2 Then
    FlagStat = "Single"
  
    If FlagZone = "STA" Then '输出站点结果
      HFGrid1.Visible = True
      HFGrid2.Visible = True
      HFGrid4.Visible = False
      HFGrid5.Visible = False
      HFGrid3(1).Visible = False
      HFGrid3(2).Visible = False
      HFGrid3(3).Visible = False
    Else '输出区域统计结果
      HFGrid1.Visible = False
      HFGrid2.Visible = False
      HFGrid4.Visible = False
      HFGrid5.Visible = False
      HFGrid3(1).Visible = True
      HFGrid3(2).Visible = True
      HFGrid3(3).Visible = True
    End If
  
    '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
    If GuestMode Then
      Call ImportData.LoadGuestCSV(FrmMeteoColdAir.HFGrid1, FrmMeteoColdAir.HFGrid2, "DataMeteoColdAir.csv", 3)
      Exit Sub
    End If
  
  End If
End Sub


Private Sub Form_Load()
  Dim i%
  Screen.MousePointer = 11
 
  
  FrmMain.FormAdd
  For i = 0 To 30
    ComboBDate.AddItem Format(Date - 30 + i, "yyyy-mm-dd")
    ComboEDate.AddItem Format(Date - 30 + i, "yyyy-mm-dd")
  Next i
  ComboBDate.Text = Format(Date - 1, "yyyy-mm-01")
  ComboEDate.Text = Format(Date - 1, "yyyy-mm-dd")
 
  Call TableIni_MeteoColdAir2(FrmMeteoColdAir) ', SelField_Initial)
'  FlagStat = "Multi":
  ReDim CA_ALL(0)
  Call TableIni_MeteoColdAir(FrmMeteoColdAir) ', SelField_Initial)
  OptionStat(2).Value = True
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

  Dim bottomAnchorTop5 As Long
  bottomAnchorTop5 = ResizeBottomAnchoredGrid(HFGrid5, Me.ScaleHeight, Me.ScaleWidth)
  Call ResizeFillGrid(HFGrid4, bottomAnchorTop5, Me.ScaleWidth)

  Call ResizeStackedGridArray(Array(HFGrid3(1), HFGrid3(2), HFGrid3(3)), Me.ScaleHeight - MDI_RESIZE_BOTTOM_MARGIN, Me.ScaleWidth)
  Call ResizeFrame1Width(Frame1, Me.ScaleWidth)
  Call ResizeChildFrameWidth(Frame11, Frame1, 120)

  Exit Sub
ResizeError:
  Resume Next
End Sub



'*********************************************单要素逐日资料统计>>>开始*****************************************
'*********************逐日资料查询（可查询天气过程，如连续高温、低温、雨日等）**********************************
Public Sub CmdMeteoColdAir()
  Dim strLvl$ 'Cold Air OutbreakLevel
  
  Screen.MousePointer = 11
  
  '************2023-10-10添加对输入日期的精准判断**********
  Dim tempStr$()
  
  tempStr = Split(ComboBDate.Text, "-")
  BYYYY = tempStr(0): BMM = tempStr(1): BDD = tempStr(2)
  BDate = Format(BYYYY, "0000") & "-" & Format(BMM, "00") & "-" & Format(BDD, "00")
  
  tempStr = Split(ComboEDate.Text, "-")
  EYYYY = tempStr(0): EMM = tempStr(1): EDD = tempStr(2)
  EDate = Format(EYYYY, "0000") & "-" & Format(EMM, "00") & "-" & Format(EDD, "00")

  
  Dim FlagRunBYear, FlagRunEYear As Boolean '判断起止年份是否为闰年
  Dim tempYear%
  FlagRunBYear = False: tempYear = CInt(BYYYY)
  If tempYear Mod 4 = 0 And tempYear Mod 100 <> 0 Or tempYear Mod 400 = 0 Then FlagRunBYear = True
  FlagRunEYear = False: tempYear = CInt(EYYYY)
  If tempYear Mod 4 = 0 And tempYear Mod 100 <> 0 Or tempYear Mod 400 = 0 Then FlagRunEYear = True
  
  If CInt(BMM) > 12 Or CInt(BMM) < 1 Then MsgBox "Invalid start month", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  If CInt(EMM) > 12 Or CInt(EMM) < 1 Then MsgBox "Invalid end month", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
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

  idays = DateDiff("d", BDate, EDate) + 1
  If idays <= 0 Then MsgBox "Invalid start/end date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  If Left(EDate, 4) - Left(BDate, 4) > 1 Then
    MsgBox "Date range cannot exceed 1 year", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If
  
  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart
  

  ReDim CA_ALL(0) '清空冷空气内存变量
  If FlagStat = "Multi" Then
    Call TableIni_MeteoColdAir(FrmMeteoColdAir)
  ElseIf FlagStat = "Single" Then
    Call TableIni_MeteoColdAir2(FrmMeteoColdAir)
  End If
    
  ORAConn2.Open
  ORAConn2.CursorLocation = 3

  
'************************************* 一、查询所是站点所是冷空气过程值  ***********************************
  strLvl = "("
  For i = 0 To 4
    If Check_lvl(i).Value = Checked Then
      If strLvl = "(" Then
        strLvl = strLvl & Str(i)
      Else
        strLvl = strLvl & "," & Str(i)
      End If
    End If
  Next i
  strLvl = strLvl & ")"
  Call QueryCAData(BDate, EDate, CA_ALL, strLvl)
  
  If UBound(CA_ALL) <> 0 Then
  
      ' 统计这段时间的最强过程
      If FlagStat = "Single" Then
        ReDim CA_STA(1 To StaNum)
        For i = 1 To StaNum
          CA_STA(i).stacode = StaInfo(i).stacode
          CA_STA(i).staname = StaInfo(i).staname
          CA_STA(i).ilevel = 0 'Cold Air OutbreakLevel：0-None，1-弱，2-中，3-强，4-Cold Wave
          CA_STA(i).BDate = CDate("1899-09-09") '降温第一日
          CA_STA(i).EDate = CDate("1899-09-09") '回温前一日
          CA_STA(i).idays = -9999 '降温日数
          CA_STA(i).tv_24hmax = -9999 '最大24小时降温
          CA_STA(i).tv_48hmax = -9999 '最大48小时降温
          CA_STA(i).tv_acc = -9999 '累计降温
          CA_STA(i).T = -9999  '降温过程的平均气温
          CA_STA(i).t_min = -9999 '降温过程的最低气温
          CA_STA(i).Town = StaInfo(i).Town
          CA_STA(i).County = StaInfo(i).County
          CA_STA(i).City = StaInfo(i).City
          
        Next i
        Call StatCAData(CA_ALL, CA_STA)
      End If
      
      
    '************************************* 情形一：查询结果为站点尺度  ***********************************
      If FlagZone = "STA" Then '站点尺度
          
        If FlagStat = "Single" Then
          '统计所是行的平均/最大/最小Value
          Call Stat_ROWS_ColdAir(CA_STA)
          '填充上方表格
          Call HFGrid1_Fill_ColdAir(FrmMeteoColdAir, CA_STA)
          '填充下方表格
          Call HFGrid2_Fill_ColdAir(FrmMeteoColdAir, ROWS_Stat_ColdAir)
          
        ElseIf FlagStat = "Multi" Then
          '统计所是行的平均/最大/最小Value
          Call Stat_ROWS_ColdAir(CA_ALL)
          '填充上方表格
          Call HFGrid4_Fill_ColdAir(FrmMeteoColdAir, CA_ALL)
          '填充下方表格
          Call HFGrid5_Fill_ColdAir(FrmMeteoColdAir, ROWS_Stat_ColdAir)
        End If
        
      
    '************************************* 情形二：查询结果为区域尺度  ***********************************
      Else
        If FlagStat = "Single" Then
          '分区域统计所是行列值
          Call Stat_ZONE_ColdAir(CA_STA)
        ElseIf FlagStat = "Multi" Then
          '分区域统计所是行列值
          Call Stat_ZONE_ColdAir(CA_ALL)
        End If
        '填充上/中/下方表格值
        Call HFGrid3_fill_ColdAir(FrmMeteoColdAir)
      
      End If
  End If

  ORAConn2.Close
  Call TimeStop
  frmWait.Hide: DoEvents
  
  Screen.MousePointer = 1
  
End Sub
'*********************************************冷空气资料统计>>>结束*****************************************

'********************对HFGrid1中的数据进行排序***************************
Private Sub HFGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid1
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols(HFGrid1, shift)
End Sub

'********************对HFGrid3中的数据进行排序***************************
Private Sub HFGrid3_MouseUp(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid3(Index)
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols(HFGrid3(Index), shift)
End Sub

'********************对HFGrid4中的数据进行排序***************************
Private Sub HFGrid4_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid4
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols(HFGrid4, shift)
End Sub


Public Sub CmdMap_MeteoColdAir()
  If FlagStat = "Multi" Then
    MsgBox "results span all events -- cannot draw a contour map", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If
  If Right(HFGrid1.TextMatrix(0, HFGrid1.Col), 2) <> "Level" Then
    Call Map_Value(FrmMeteoColdAir, "BYR")
    
  ElseIf Right(HFGrid1.TextMatrix(0, HFGrid1.Col), 2) = "Level" Then
    If SetSurfer = "No" Then
      MsgBox "Please make sure Surfer11 is installed, and go to the CROSS menu: " & Chr(13) & "System >> Surfer Settings: Enable Surfer Plotting", vbInformation, "Notice": Exit Sub
    End If
    
    If District_ID <> "GD" Then MsgBox "Only province-level users can plot cold air distribution maps", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    If Len(Dir(srfPath_ColdAir)) = 0 Then
      MsgBox "    Local SURFER Template File   " & District_ID & "_COLDAIR.srf" & " does not exist, " & vbCrLf & _
             "请进入CROSS菜单：系统操作>>Surfer Settings，尝试重新下载", vbInformation, "Notice": Exit Sub
    End If

    Call Map_ColdAir(FrmMeteoColdAir)
    
  End If
    
End Sub



'********************保存窗体内可见的HFGrid***************************
Public Sub CmdOutput_MeteoColdAir()
  Call Output_HFGrid(FrmMeteoColdAir, "CA")
End Sub


Private Sub Form_Unload(Cancel As Integer)
  FrmMain.FormDel
  FrmMain.StatusBar1.Panels(1).Text = ""
  Unload Me
End Sub


Private Sub mnuGridCopy_Click()
  Call CopyGridSelectionToClipboard(mLastRightClickGrid)
End Sub


Private Sub HFGrid2_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid2
  If Button = 2 Then PopupMenu mnuGridPopup
End Sub

Private Sub HFGrid5_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid5
  If Button = 2 Then PopupMenu mnuGridPopup
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

'鼠标滚轮驱动表格上下滚动
Private Sub HFGrid4_GotFocus()
  On Error Resume Next
  Oldwinproc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
  SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub
    
'鼠标滚轮驱动表格上下滚动
Private Sub HFGrid4_LostFocus()
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

Private Sub HFGrid4_DblClick()
  Call GridEdit(HFGrid4)
End Sub

Private Sub HFGrid5_DblClick()
  Call GridEdit(HFGrid5)
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

