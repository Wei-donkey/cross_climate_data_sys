VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmMeteoPeriodSin 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Single-Period Data Statistics"
   ClientHeight    =   13530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   28380
   Icon            =   "FrmMeteoPeriodSin.frx":0000
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
      TabIndex        =   9
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
      TabIndex        =   8
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
      TabIndex        =   1
      Top             =   120
      Width           =   28095
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Threshold"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   14280
         TabIndex        =   13
         Top             =   120
         Width           =   1575
         Begin VB.ComboBox ComboMax 
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
            TabIndex        =   15
            Text            =   "ComboMax"
            Top             =   720
            Width           =   975
         End
         Begin VB.ComboBox ComboMin 
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
            TabIndex        =   14
            Text            =   "ComboMin"
            Top             =   360
            Width           =   975
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
            Left            =   240
            TabIndex        =   17
            Top             =   795
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
            Left            =   240
            TabIndex        =   16
            Top             =   435
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
         TabIndex        =   10
         Top             =   120
         Width           =   2055
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
            TabIndex        =   12
            Text            =   "yyyy-mm-dd"
            Top             =   720
            Width           =   1695
         End
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
            TabIndex        =   11
            Text            =   "yyyy-mm-dd"
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "ElementSelect"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   3
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
            TabIndex        =   7
            Top             =   240
            Width           =   11895
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   15840
         TabIndex        =   2
         Top             =   120
         Width           =   12135
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid3 
      Height          =   3900
      Index           =   1
      Left            =   120
      TabIndex        =   4
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
      TabIndex        =   5
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
      TabIndex        =   6
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
Attribute VB_Name = "FrmMeteoPeriodSin"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mLastRightClickGrid As Object '最近一次被点击（任意按钮）的表格；供mnuGridCopy_Click及ChartData查找图表数据源时使用
Private mEditGrid As Object '判断当前正在被TextEdit编辑的是哪个Grid（HFGrid1/HFGrid2/HFGrid3(Index)之一）

Dim BDate$, EDate$, idays% '用户自定义的当年查询起始日期
Dim BDate2$, EDate2$ '用于判断是效资料的起止年份
Dim BDate3$, EDate3$ '用户自定义对应的某年查询起始日期
Dim BMMDD$, EMMDD$ '用户自定义的查询起止月日


Dim BYYYY$, BMM$, BDD$, EYYYY$, EMM$, EDD$ '用户自定义的查询起止年、月、日

Dim iYear% '查询资料所在的年限，如起止日期：2012-12-31至2013-01-01，则iyear=2012；2013-08-01至2013-08-10，则iyear=2013
Dim iYear3% '同期值为某一年，该年份值变量，如2012

Dim BYear%, EYear%, iyears% 'B年份、E年份为从数据库取出的数据起止年份（1951年至今）,i年份s为读取资料起止年数
'Dim B年份2%, E年份2%, i年份s2% 'B年份2、E年份2为界面上选择的排名年份,i年份s2为排名年数
'Dim B年份_rank%, E年份_rank%, i年份s_rank% '实际排名年份
'Dim B年份3%, E年份3% '国家站多年平均值的起止年份
'Dim B年份4%, E年份4% '自动站多年平均值的起止年份

Dim AAA!() '按年度循环读oracle，统计值存入二维数组AAA
Dim AAA_tmp!() '2024-4-12Add
Dim ADays%() ' 存放多年度所是站点条件日数：多日结果

Dim BBB!(), BBBDate() As Date  ', BBBDays%() '当年值、当年极值日期，2014-07-24更改为二维数组,满足条件的日数
Dim CCC!(), CCCDate() As Date '多年平均值或历年极值、极值日期，2014-07-24定为一维数组
Dim DDD!(), DDDDate() As Date ', DDDDays%() '某年值、某年极值日期，2014-07-24更改为二维数组,满足条件的日数

Dim BBB_avg!, BBB_sum!, iNum% '当年多站点值的平均值、累计值
Dim AAA_avg!(), AAA_sum!() '历年多站点值的平均值、累计值
Dim AAA_ZONE!() '历年各区域的平均值
Dim BBB_ZONE!() '当年各区域的平均值


Dim Norm_Avg!(), Norm_Sum!(), Norm_num% '多年平均均值、多年总值、均值年份数

Dim AAARank!(), BBBRank! '某站点/地区平均的历年待排序值、当年值


Dim crossYear As Boolean '查询时段是否跨年度


Public SelField_Initial$ '选择要素的初始名
Dim SelField_Restat$ '选择要素的再统计名

Dim SelField_Day$ '选择要素的日间隔字段
'Dim Sel字段_Mon$ '选择要素的月值表字段
'
'Dim Sel字段_Xtrm$ '选择要素的极值表字段
'Dim Sel字段_Norm$ '选择要素的平均态字段

Public FlagStat$ '统计量标识
Dim FlagLimit As Boolean '条件查询标识


Dim i%, j%, k%, l%
Dim TextMin!, TextMax! '条件查询的上下限

Dim XtrmYear_STA() As Integer '各站点的极值出现年份
Dim SelField_Yer$ '选择要素的年间隔字段

Dim FlagTempStat As Boolean  '临时统计：不从月值表读取的条件查询--长耗时：2017-3-23
Dim FlagResult As String '2023-03-14：不符合条件的结果设为0或null，0参与多站点平均，null不参与平均



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
   
  Call ReadColInfo("COLS_Initial.ini")
  COL_Initial = COL_tmp
  
  Call ReadColInfo("COLS_ReStat.ini")
  COL_Restat = COL_tmp
 
  '****************************2022-03-08从ini配置文件中读取字段名*****************************
  List_Col.Clear
  For i = 1 To UBound(COL_Restat)
    List_Col.AddItem COL_Restat(i).Col_name
  Next i
  List_Col.Selected(0) = True
  
  SelField_Initial = COL_Initial(1).Col_ID
'  SelField_Restat = COL_Restat(1).Col_ID
    
''''''''  Call TableIni_MeteoPeriodSin(FrmMeteoPeriodSin, SelField_Initial)
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

'*********************************************单要素任意时段资料统计>>>开始*****************************************
Public Sub CmdMeteoPeriodSin()

  Screen.MousePointer = 11
  
  Call TableIni_MeteoPeriodSin(FrmMeteoPeriodSin, SelField_Initial)
  
  '************2023-1-18添加对输入日期的精准判断**********
  Dim tempStr$()
  
  tempStr = Split(ComboBDate.Text, "-")
  BYYYY = tempStr(0): BMM = tempStr(1): BDD = tempStr(2)
  BDate = Format(BYYYY, "0000") & "-" & Format(BMM, "00") & "-" & Format(BDD, "00")
  
  tempStr = Split(ComboEDate.Text, "-")
  EYYYY = tempStr(0): EMM = tempStr(1): EDD = tempStr(2)
  EDate = Format(EYYYY, "0000") & "-" & Format(EMM, "00") & "-" & Format(EDD, "00")
  
  BMMDD = Right(BDate, 5) '开始月日
  EMMDD = Right(EDate, 5) '结束月日
  
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
      
  '********2021-9-14 其他所是年份均不查02-29*************
  If BMMDD = "02-29" Then MsgBox "Start date cannot be Feb 29", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
  '********2024-2-27 添加对2月29日的判断*************
  If EMMDD = "02-29" Then
    If CInt(EYYYY) Mod 4 <> 0 Then MsgBox "Not a leap year -- invalid date input", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If
  
  idays = DateDiff("d", BDate, EDate) + 1
  If idays <= 0 Then MsgBox "Invalid start/end date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub

  If Left(EDate, 4) - Left(BDate, 4) > 10 Then
    MsgBox "Date range cannot exceed 10 years", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If

  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart

  '当年值年份
  iYear = Left(BDate, 4)

  
  ORAConn2.Open
  ORAConn2.CursorLocation = 3
 

  '****************************2022-03-08*****************************
  For i = 0 To List_Col.ListCount - 1
    If List_Col.Selected(i) = True Then
      SelField_Initial = COL_Initial(i + 1).Col_ID
'      SelField_Restat = COL_Restat(i + 1).Col_ID
      Exit For
    End If
  Next i
  
  TextMin = ComboMin.Text: TextMax = ComboMax.Text
  If TextMin = -9999 And TextMax = 999999 Then
    FlagLimit = False
  Else
    FlagLimit = True
  End If

  
  '按已选的要素循环统计
  For k = 1 To 5
    If k = 1 Then FlagStat = "AVE"
    If k = 2 Then FlagStat = "MAX"
    If k = 3 Then FlagStat = "MIN"
    If k = 4 Then FlagStat = "SUM"
    If k = 5 Then FlagStat = "DAYS"
  
    '************************************* 一、查询所是站点任意时段值  ***********************************
    ReDim PRD_STA(1 To StaNum)
    For i = 1 To StaNum
      PRD_STA(i).stacode = StaInfo(i).stacode
      PRD_STA(i).staname = StaInfo(i).staname
      
      PRD_STA(i).stat = -9999 '当年统计值
      PRD_STA(i).stat_date = CDate("1899-09-09") '极值出现日期
                
    Next i
    Call QueryPeriodSinData
  
  
    '************************************* 二、处理所是站点任意时段结果  ***********************************
    
    '************************************* 情形一：查询结果为站点尺度  ***********************************
      If FlagZone = "STA" Then '站点尺度
    
    '************************************* 1、计算所是行(站点)的统计值  ***********************************
        '统计所是行的平均/最大/最小Value
        Call Stat_ROWS_Period(PRD_STA, StaNum, FlagStat, SelField_Initial)
'    '************************************* 2、计算历年所是站的平均值  ***********************************
'        Call Stat_ROWS_AAA(AAA, AAA_avg, BBB, BBB_avg)
    '************************************* 3、填充上方表格  ***********************************
        Call HFGrid1_Fill_PeriodSin(FrmMeteoPeriodSin, PRD_STA, FlagStat)
    '************************************* 4、填充下方表格 ***********************************
        Call HFGrid2_Fill_PeriodSin(FrmMeteoPeriodSin, ROWS_Stat_Period, FlagStat)
    
    '************************************* 情形二：查询结果为区域尺度  ***********************************
      Else
    
    '************************************* 1、分区域统计所是行列值  ***********************************
        Call Stat_ZONE_Period(PRD_STA, FlagStat, SelField_Initial)
'    '************************************* 2、计算历年各区的平均值  ***********************************
'        Call Stat_ZONE_AAA(AAA, AAA_ZONE, BBB, BBB_ZONE)
    '************************************* 3、填充上/中/下方表格值  ***********************************
        Call HFGrid3_fill_PeriodSin(FrmMeteoPeriodSin, FlagStat)
        
      End If
  Next k

  ORAConn2.Close
  Call TimeStop
  frmWait.Hide: DoEvents
  
  Screen.MousePointer = 1
    
End Sub
'*********************************************单要素逐日资料统计>>>结束*****************************************


'********************************************* 查询、构建任意时段数据结果变量（2024-09-23）*****************************************
Private Sub QueryPeriodSinData()
  Dim i%, j%, k%, l%
  Dim sngSum!, sngMax!, sngMin! '当年值的统计量
  Dim sngSum2!, sngMax2!, sngMin2!  '统计往年值所需的中间变量
  Dim iNum%, iNum2% '用于计算多站平均的站点数
  

'*******************************************1、查询当年值：开始***********************************************************
  sngSum = 0: sngMax = -9999: sngMin = 999999: iNum = 0
    
  If FlagStat = "MAX" Or FlagStat = "MIN" Then '统计量：最大Value、最小Value查询
    '*******************************赋初始值-9999或0*************************
    ReDim BBB(1 To StaNum), BBBDate(1 To StaNum)
    For i = 1 To StaNum
      BBB(i) = -9999: BBBDate(i) = CDate("1899-09-09")
    Next i
    Call StatExtremeDayData(SelField_Initial, BDate, EDate, BBB(), BBBDate(), FlagStat, TextMin, TextMax)
    For i = 1 To StaNum
      
      '2026-4-1添加：站点开始日期前的数据设为-9999
      If iYear < StaInfo(i).YearSTT And BBB(i) <> -9999 Then
        BBB(i) = -9999
      End If
 
      PRD_STA(i).stat = BBB(i)
      PRD_STA(i).stat_date = BBBDate(i)
    Next i
        
  Else '其他统计量：条件日数，非条件日数查询的平均值、累计值查询
    '*******************************赋初始值-9999或0*************************
    ReDim BBB(1 To StaNum)
    For i = 1 To StaNum
      BBB(i) = -9999
    Next i
    Call StatMulDayData(SelField_Initial, BDate, EDate, BBB(), FlagStat, TextMin, TextMax)
    For i = 1 To StaNum
      
      '2026-4-1添加：站点开始日期前的数据设为-9999
      If iYear < StaInfo(i).YearSTT And BBB(i) <> -9999 Then
        BBB(i) = -9999
      End If
            
      If (FlagStat = "SUM" Or FlagStat = "DAYS") And FlagLimit = True Then '条件日数或条件累计值查询，无符合条件的都为0
        If iYear >= StaInfo(i).YearSTT Then
          If BBB(i) = -9999 Then
            If FlagResult = "0" Then BBB(i) = 0  '当年(iyear)年份已经是资料了，则条件日数默认为0
          ElseIf BBB(i) = 0 Then
            If FlagResult = "null" Then BBB(i) = -9999
          End If
        End If
      End If
      PRD_STA(i).stat = BBB(i)
    Next i
  End If
  
End Sub


'********************对HFGrid1中的数据进行排序***************************
Private Sub HFGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid1
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Period_Cols(HFGrid1, shift)
End Sub

'********************对HFGrid3中的数据进行排序***************************
Private Sub HFGrid3_MouseUp(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid3(Index)
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Period_Cols(HFGrid3(Index), shift)
End Sub


'********************利用HFGrid中的数据绘制地图***************************
Public Sub CmdMap_MeteoPeriodSin()
  If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Then
    Call Map_Value(FrmMeteoPeriodSin, "RYB")
  Else
    Call Map_Value(FrmMeteoPeriodSin, "BYR")
  End If
End Sub


'********************保存窗体内可见的HFGrid***************************
Public Sub CmdOutput_MeteoPeriodSin()
  Call Output_HFGrid(FrmMeteoPeriodSin, "")
End Sub

Private Sub List_Col_Click()
    For i = 0 To List_Col.ListCount - 1
      If List_Col.Selected(i) = True Then
        SelField_Initial = COL_Initial(i + 1).Col_ID
        SelField_Restat = COL_Restat(i + 1).Col_ID
        Exit For
      End If
    Next i
    Call Option_MeteoPeriodSin(FrmMeteoPeriodSin, SelField_Restat)
    Call TableIni_MeteoPeriodSin(FrmMeteoPeriodSin, SelField_Initial)
    
    '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
    If GuestMode And List_Col.Selected(0) = True Then
      Call ImportData.LoadGuestCSV(FrmMeteoPeriodSin.HFGrid1, FrmMeteoPeriodSin.HFGrid2, "DataMeteoPeriodSin.csv", 3)
    End If
 
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
