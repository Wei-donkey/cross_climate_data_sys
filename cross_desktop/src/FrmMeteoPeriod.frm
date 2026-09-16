VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmMeteoPeriod 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Custom-Period Data Statistics"
   ClientHeight    =   13530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   28380
   Icon            =   "FrmMeteoPeriod.frx":0000
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
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Start/End Date"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   12240
         TabIndex        =   30
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
            TabIndex        =   32
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
            TabIndex        =   31
            Text            =   "yyyy-mm-dd"
            Top             =   360
            Width           =   1695
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Ranking Year"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   22560
         TabIndex        =   27
         Top             =   120
         Width           =   1335
         Begin VB.ComboBox ComboEYear2 
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
            Left            =   120
            TabIndex        =   29
            Text            =   "yyyy"
            Top             =   720
            Width           =   855
         End
         Begin VB.ComboBox ComboBYear2 
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
            Left            =   120
            TabIndex        =   28
            Text            =   "yyyy"
            Top             =   360
            Width           =   855
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Prior-Year Value"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   20880
         TabIndex        =   24
         Top             =   120
         Width           =   1695
         Begin VB.ComboBox ComboYear3 
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
            Left            =   120
            TabIndex        =   25
            Text            =   "yyyy"
            Top             =   720
            Width           =   855
         End
         Begin VB.Label Label1 
            BackColor       =   &H00DFEFDF&
            Caption         =   "year"
            Height          =   255
            Left            =   240
            TabIndex        =   26
            Top             =   360
            Width           =   855
         End
         Begin VB.Image ImageSwitchNorm 
            Height          =   480
            Left            =   1080
            Picture         =   "FrmMeteoPeriod.frx":3482
            ToolTipText     =   "Switch Climate Normal Data Source"
            Top             =   480
            Width           =   480
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Condition not met"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   19080
         TabIndex        =   21
         Top             =   120
         Width           =   1815
         Begin VB.OptionButton OptionNull 
            BackColor       =   &H00DFEFDF&
            Caption         =   "is zero"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   23
            Top             =   360
            Value           =   -1  'True
            Width           =   1455
         End
         Begin VB.OptionButton OptionNull 
            BackColor       =   &H00DFEFDF&
            Caption         =   "is empty"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   22
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Threshold"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   17520
         TabIndex        =   16
         Top             =   120
         Width           =   1575
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
            TabIndex        =   18
            Text            =   "ComboMin"
            Top             =   360
            Width           =   975
         End
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
            TabIndex        =   17
            Text            =   "ComboMax"
            Top             =   720
            Width           =   975
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
            TabIndex        =   20
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
            Left            =   240
            TabIndex        =   19
            Top             =   795
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Statistic"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   14280
         TabIndex        =   10
         Top             =   120
         Width           =   3255
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Max"
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
            Left            =   1200
            TabIndex        =   15
            Top             =   360
            Width           =   975
         End
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Min"
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
            Index           =   3
            Left            =   2160
            TabIndex        =   14
            Top             =   360
            Width           =   975
         End
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
            TabIndex        =   13
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
            TabIndex        =   12
            Top             =   720
            Width           =   975
         End
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Day Count"
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
            Index           =   7
            Left            =   1200
            TabIndex        =   11
            Top             =   720
            Width           =   1815
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
            MultiSelect     =   2  'Extended
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
         Left            =   23040
         TabIndex        =   2
         Top             =   120
         Width           =   4935
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
Attribute VB_Name = "FrmMeteoPeriod"
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
Dim BYear2%, EYear2%, iYears2% 'B年份2、E年份2为界面上选择的排名年份,i年份s2为排名年数
Dim BYear_rank%, EYear_rank%, iYears_rank% '实际排名年份
Dim BYear3%, EYear3% '国家站多年平均值的起止年份
Dim BYear4%, EYear4% '自动站多年平均值的起止年份

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
Dim SelField_Mon$ '选择要素的月值表字段

Dim SelField_Xtrm$ '选择要素的极值表字段
Dim SelField_Norm$ '选择要素的平均态字段

Public FlagStat$ '统计量标识
Dim FlagLimit As Boolean '条件查询标识

Dim i%, j%, k%, l%
Dim TextMin!, TextMax! 'Day Count/总量查询的上下限

Dim XtrmYear_STA() As Integer '各站点的极值出现年份
Dim SelField_Yer$ '选择要素的年间隔字段

Dim FlagTempStat As Boolean  '临时统计：不从月值表读取的条件查询--长耗时：2017-3-23
Dim FlagResult As String '2023-03-14：不符合条件的结果设为0或null，0参与多站点平均，null不参与平均
Dim mLastSelIndex As Integer '要素选择列表中最后一次合法选择的索引，用于阻止取消唯一选中项

Private Sub OptionNull_Click(Index As Integer)
  If OptionNull(0).Value = True Then
    FlagResult = "0"
  ElseIf OptionNull(1).Value = True Then
    FlagResult = "null"
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
  
  ComboBYear2.Clear: ComboEYear2.Clear
  
  For i = 1951 To Year(Date)
    ComboYear3.AddItem i
    ComboBYear2.AddItem i
    ComboEYear2.AddItem i
  Next i
  ComboYear3.Text = Year(Date) - 1
  ComboBYear2.Text = 1951
  ComboEYear2.Text = Year(Date)
 
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
  mLastSelIndex = 0
  
  SelField_Initial = COL_Initial(1).Col_ID
  SelField_Restat = COL_Restat(1).Col_ID
  FlagStat = "AVE"
  OptionNull(1).Value = True
    
'''''''  Call TableIni_MeteoPeriod(FrmMeteoPeriod, SelField_Initial)
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
Public Sub CmdMeteoPeriod()

  Screen.MousePointer = 11
  
  Call TableIni_MeteoPeriod(FrmMeteoPeriod, SelField_Initial)
  
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
'  If EMMDD = "02-29" Then MsgBox "不支持2月29日", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
  '********2024-2-27 添加对2月29日的判断*************
  If EMMDD = "02-29" Then
    If CInt(EYYYY) Mod 4 <> 0 Then MsgBox "Not a leap year -- invalid date input", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If
  
  idays = DateDiff("d", BDate, EDate) + 1
  If idays <= 0 Then MsgBox "Invalid start/end date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub

  If Left(EDate, 4) - Left(BDate, 4) > 1 Then
    MsgBox "Date range cannot exceed 1 year", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If
  
  '排名起止年份（可自定义）
  BYear2 = ComboBYear2.Text  '起始年份
  EYear2 = ComboEYear2.Text  '结束年份
  If BYear2 > EYear2 Then MsgBox "Invalid ranking year", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub

    
  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart
  
  
  '当年值年份
  iYear = Left(BDate, 4)
  'a given yearValueYear
  iYear3 = ComboYear3.Text
  
  If Left(EDate, 4) - Left(BDate, 4) = 0 Then '同年度
    crossYear = False
    BDate3 = CStr(iYear3) & "-" & BMMDD
    If EMMDD <> "02-29" Then
      EDate3 = CStr(iYear3) & "-" & EMMDD
    
    '2024-2-27判断结束日期为02-29的某年是否为平年
    ElseIf EMMDD = "02-29" Then
      If iYear3 Mod 4 <> 0 Then EDate3 = CStr(iYear3) & "-02-28"
      If iYear3 Mod 4 = 0 Then EDate3 = CStr(iYear3) & "-02-29"
    End If
    
  ElseIf Left(EDate, 4) - Left(BDate, 4) = 1 Then '跨年度：年份数少一年
    crossYear = True
    BDate3 = CStr(iYear3) & "-" & BMMDD
    If EMMDD <> "02-29" Then
      EDate3 = CStr(iYear3 + 1) & "-" & EMMDD
  
    '2024-2-27判断结束日期为02-29的某年是否为平年
    ElseIf EMMDD = "02-29" Then
      If (iYear3 + 1) Mod 4 <> 0 Then EDate3 = CStr(iYear3 + 1) & "-02-28"
      If (iYear3 + 1) Mod 4 = 0 Then EDate3 = CStr(iYear3 + 1) & "-02-29"
    End If
  
  End If

  
  ORAConn2.Open
  ORAConn2.CursorLocation = 3
 

  '****************************2022-03-08*****************************
  For i = 0 To List_Col.ListCount - 1
    If List_Col.Selected(i) = True Then
      SelField_Initial = COL_Initial(i + 1).Col_ID
      SelField_Restat = COL_Restat(i + 1).Col_ID
      Exit For
    End If
  Next i
  
  TextMin = ComboMin.Text: TextMax = ComboMax.Text
  If TextMin = -9999 And TextMax = 999999 Then
    FlagLimit = False
  Else
    FlagLimit = True
  End If
  
  '平均、最大、最小统计结果不满足阈值时，只能显示：空
  If OptionStat(1).Value = True Then
    FlagStat = "AVE":  'OptionNull(1).Value = True
  ElseIf OptionStat(2).Value = True Then
    FlagStat = "MAX":  'OptionNull(1).Value = True
  ElseIf OptionStat(3).Value = True Then
    FlagStat = "MIN":  'OptionNull(1).Value = True
  ElseIf OptionStat(4).Value = True Then
    FlagStat = "SUM":  ' OptionNull(0).Value = True  2024-10-28 发现错误，注销
  ElseIf OptionStat(7).Value = True Then
    FlagStat = "DAYS": ' OptionNull(0).Value = True  2024-10-28 发现错误，注销
  End If
      
  If FlagLimit = False Then
    If (FlagStat = "SUM" Or FlagStat = "DAYS") Then
      FlagTempStat = True
      If List_Col.Selected(4) = True Or List_Col.Selected(5) = True Or List_Col.Selected(6) = True Or List_Col.Selected(7) = True _
      Or List_Col.Selected(32) = True Or List_Col.Selected(35) = True Then '20时/08时/20-08时/08-20时日雨量、蒸发、日照：SUM/DAYS直接读取月值表即可
        FlagTempStat = False
        SelField_Mon = SelField_Restat & "_" & FlagStat
      End If
    Else
      FlagTempStat = False
      SelField_Mon = SelField_Restat & "_" & FlagStat
    End If
    
  ' 2023-03-03 如果条件查询，则需要判断是否进行临时统计
  ElseIf FlagLimit = True Then
    
    If FlagStat = "DAYS" Then
        FlagTempStat = True '默认需要临时统计
    
        If List_Col.Selected(0) = True Then '平均气温
          If (TextMin = 10 Or TextMin = 0) And TextMax = 999999 Then FlagTempStat = False
        ElseIf List_Col.Selected(1) = True Then  '最高气温
          If TextMin = 35 And TextMax = 999999 Then FlagTempStat = False
        ElseIf List_Col.Selected(2) = True Then  '最低气温
          If TextMin = -9999 And TextMax = 5 Then FlagTempStat = False
        ElseIf List_Col.Selected(4) = True Then  '20时降水量
          If (TextMin = 0.1 Or TextMin = 10 Or TextMin = 25 Or TextMin = 50) And TextMax = 999999 Then FlagTempStat = False
        End If
        
    ElseIf FlagStat = "SUM" Then
        FlagTempStat = True '默认需要临时统计
    
        If List_Col.Selected(0) = True Then '平均气温
          If (TextMin = 10 Or TextMin = 0) And TextMax = 999999 Then FlagTempStat = False
        ElseIf List_Col.Selected(4) = True Then  '20时降水量
          If (TextMin = 0.1 Or TextMin = 10 Or TextMin = 25 Or TextMin = 50) And TextMax = 999999 Then FlagTempStat = False
        End If
        
    Else
        FlagTempStat = True '默认需要临时统计
    
    End If
        
    If FlagTempStat = False Then
      If TextMax <> 999999 Then SelField_Mon = SelField_Restat & "_" & TextMax & FlagStat
      If TextMin <> -9999 Then SelField_Mon = SelField_Restat & "_" & TextMin & FlagStat
      If TextMin = 0.1 Then SelField_Mon = SelField_Restat & "_" & FlagStat '雨量/雨日
    End If
  End If
    
  If FlagStat = "MAX" Or FlagStat = "MIN" Then
    SelField_Xtrm = SelField_Restat & "_" & FlagStat
  End If
  
'********************************2022-3-24修改阈值选项为list，不可以进行临时查询，FlagTempStat一定是False***************************
'********************************2022-3-06修改阈值选项为可修改，可以进行临时查询，FlagTempStat不一定是False***************************
  
  ' 判断从数据库中读取的数据年限Byear-Eyear
  BYear3 = Left(Years_NormSURF, 4)
  EYear3 = Right(Years_NormSURF, 4)
  
  BYear4 = Left(Years_NormAWST, 4)
  EYear4 = Right(Years_NormAWST, 4)

  BYear = BYear2
  If BYear3 < BYear Then BYear = BYear3
  If BYear4 < BYear Then BYear = BYear4
  EYear = EYear2
  If EYear3 > EYear Then EYear = EYear3
  If EYear4 > EYear Then EYear = EYear4
  
  iyears = CInt(EYear) - CInt(BYear) + 1
  
  '**************************判断是资料的第一个年份yyyy存入StaInfo(i).年份STT********************************************
  For i = 1 To StaNum
    For j = StaInfo(i).YearSTT To EYear '从该站是记录来的第一年开始判断
      If crossYear = False Then '同年度
        BDate2 = CStr(j) & "-" & BMMDD
        If EMMDD <> "02-29" Then
          EDate2 = CStr(j) & "-" & EMMDD
        ElseIf EMMDD = "02-29" Then
          If j Mod 4 <> 0 Then EDate2 = CStr(j) & "-02-28"
          If j Mod 4 = 0 Then EDate2 = CStr(j) & "-02-29"
        End If
          
      ElseIf crossYear = True Then '跨年度：年份数少一年
        BDate2 = CStr(j) & "-" & BMMDD
        If EMMDD <> "02-29" Then
          EDate2 = CStr(j + 1) & "-" & EMMDD
        ElseIf EMMDD = "02-29" Then
          If (j + 1) Mod 4 <> 0 Then EDate2 = CStr(j + 1) & "-02-28"
          If (j + 1) Mod 4 = 0 Then EDate2 = CStr(j + 1) & "-02-29"
        End If
      End If
      If CDate(BDate2) >= CDate(StaInfo(i).DateSTT) Then  '如果要查询的开始日期在是效资料日期之后，则该站点从这一年开始是资料可查
        StaInfo(i).YearSTT = j
        Exit For
      End If
    Next j
  Next i
  
  
'************************************* 一、查询所是站点任意时段值  ***********************************
  ReDim PRD_STA(1 To StaNum)
  For i = 1 To StaNum
    PRD_STA(i).stacode = StaInfo(i).stacode
    PRD_STA(i).staname = StaInfo(i).staname
      
    PRD_STA(i).stat = -9999 '当年统计值
    PRD_STA(i).stat_date = CDate("1899-09-09") '极值出现日期
    
    PRD_STA(i).stat_hist = -9999 '累年值
    PRD_STA(i).stat_hist_date = CDate("1899-09-09") '累年极值出现日期
    PRD_STA(i).stat_diff1 = -9999 '多年平均距平
    
    PRD_STA(i).stat_year = -9999 'a given yearValue
    PRD_STA(i).stat_year_date = CDate("1899-09-09") '某年极值出现日期
    PRD_STA(i).stat_diff2 = -9999  '某年距平
    
    PRD_STA(i).rank1 = -9999 '排名（大-小）
    PRD_STA(i).rank2 = -9999 '排名（小-大）
    PRD_STA(i).rank_years = "-9999" 'Ranking Year
    PRD_STA(i).maxyear = -9999 '最大Value年
    PRD_STA(i).maxvalue = -9999 'Max
    PRD_STA(i).minyear = -9999 '最小Value年
    PRD_STA(i).minvalue = -9999 'Min
  Next i
  Call QueryPeriodData

'************************************* 二、处理所是站点任意时段结果  ***********************************

'************************************* 情形一：查询结果为站点尺度  ***********************************
  If FlagZone = "STA" Then '站点尺度

'************************************* 1、对各站点历年值进行排序  ***********************************
    For i = 1 To StaNum
      ReDim AAARank(BYear To EYear)
      For j = BYear To EYear
        AAARank(j) = AAA(i, j)
      Next j
      BBBRank = BBB(i)
      Call Rank_PeriodData(AAARank, BBBRank, PRD_STA(i))
    Next i

'************************************* 2、计算所是行(站点)的统计值  ***********************************
    '统计所是行的平均/最大/最小Value
    Call Stat_ROWS_Period(PRD_STA, StaNum, FlagStat, SelField_Initial)
    
'************************************* 3、计算历年所是站的平均值  ***********************************
    Call Stat_ROWS_AAA(AAA, AAA_avg, BBB, BBB_avg)

'************************************* 4、对历年所是站平均值进行排序  ***********************************
    Call Rank_PeriodData(AAA_avg, BBB_avg, ROWS_Stat_Period(1))


'************************************* 5、填充上方表格  ***********************************
    If FlagStat = "MAX" Or FlagStat = "MIN" Then
      Call HFGrid1_Fill_Period1(FrmMeteoPeriod, PRD_STA)
    Else
      Call HFGrid1_Fill_Period2(FrmMeteoPeriod, PRD_STA)
    End If
'************************************* 6、填充下方表格 ***********************************
    If FlagStat = "MAX" Or FlagStat = "MIN" Then
      Call HFGrid2_Fill_Period1(FrmMeteoPeriod, ROWS_Stat_Period)
    Else
      Call HFGrid2_Fill_Period2(FrmMeteoPeriod, ROWS_Stat_Period)
    End If


'************************************* 情形二：查询结果为区域尺度  ***********************************
  Else

'************************************* 1、分区域统计所是行列值  ***********************************
    Call Stat_ZONE_Period(PRD_STA, FlagStat, SelField_Initial)

'************************************* 2、计算历年各区的平均值  ***********************************
    Call Stat_ZONE_AAA(AAA, AAA_ZONE, BBB, BBB_ZONE)

'************************************* 3、对各区域历年值进行排序  ***********************************
    Dim iRows%
    If FlagZone = "CITY" Then iRows = CityNum
    If FlagZone = "CNTY" Then iRows = CountyNum
    If FlagZone = "TOWN" Then iRows = TownNum
    
    For i = 1 To iRows
      ReDim AAARank(BYear To EYear)
      For j = BYear To EYear
        AAARank(j) = AAA_ZONE(i, j)
      Next j
      BBBRank = BBB_ZONE(i)
      Call Rank_PeriodData(AAARank, BBBRank, PRD_AVG(i))
    Next i

'************************************* 3、填充上/中/下方表格值  ***********************************
    If FlagStat = "MAX" Or FlagStat = "MIN" Then
      Call HFGrid3_fill_Period1(FrmMeteoPeriod)
    Else
      Call HFGrid3_fill_Period2(FrmMeteoPeriod)
    End If
    
  End If


  ORAConn2.Close
  Call TimeStop
  frmWait.Hide: DoEvents
  
  Screen.MousePointer = 1
  
End Sub
'*********************************************单要素逐日资料统计>>>结束*****************************************


'********************************************* 查询、构建任意时段数据结果变量（2022-03-14）*****************************************
Private Sub QueryPeriodData()
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

'************************************************2、查询某年值：开始****************************************************
  sngSum = 0: sngMax = -9999: sngMin = 999999: iNum = 0

  If FlagStat = "MAX" Or FlagStat = "MIN" Then '统计量：最大Value、最小Value查询
    '*******************************赋初始值-9999或0*************************
    ReDim DDD(1 To StaNum), DDDDate(1 To StaNum)
    For i = 1 To StaNum
      DDD(i) = -9999: DDDDate(i) = CDate("1899-09-09")
    Next i
    Call StatExtremeDayData(SelField_Initial, BDate3, EDate3, DDD(), DDDDate(), FlagStat, TextMin, TextMax)
    For i = 1 To StaNum
      
      '2026-4-1添加：站点开始日期前的数据设为-9999
      If iYear < StaInfo(i).YearSTT And DDD(i) <> -9999 Then
        DDD(i) = -9999
      End If
            
      PRD_STA(i).stat_year = DDD(i)
      PRD_STA(i).stat_year_date = DDDDate(i)
    Next i

  Else '其他包括：条件日数查询和非条件日数查询的平均值、累计值查询
    '*******************************赋初始值-9999或0*************************
    ReDim DDD(1 To StaNum) ', DDDDates(1 To StaNum, iYear3 To iYear3)
    For i = 1 To StaNum
      DDD(i) = -9999
    Next i
    Call StatMulDayData(SelField_Initial, BDate3, EDate3, DDD(), FlagStat, TextMin, TextMax)
    For i = 1 To StaNum
    
      '2026-4-1添加：站点开始日期前的数据设为-9999
      If iYear < StaInfo(i).YearSTT And DDD(i) <> -9999 Then
        DDD(i) = -9999
      End If
          
      If (FlagStat = "SUM" Or FlagStat = "DAYS") And FlagLimit = True Then '条件日数或条件累计值查询，无符合条件的都为0
        If iYear3 >= StaInfo(i).YearSTT Then
          If DDD(i) = -9999 Then
            If FlagResult = "0" Then DDD(i) = 0  'j年份已经是资料了，若用户选择条件结果设置为0，则条件结果默认为0
          ElseIf DDD(i) = 0 Then
            If FlagResult = "null" Then DDD(i) = -9999
          End If
        End If
      End If
      PRD_STA(i).stat_year = DDD(i)
    Next i
  End If

'***************************************** 3、查询所是年度（Byear-Eyear）各站值：开始  ***********************************
'*******************************赋初始值-9999或0*************************
  ReDim AAA(1 To StaNum, BYear To EYear) '存放各站点、各年份统计值（包括平均值、最大Value、最小Value、和值…）
  If FlagTempStat = True Then ReDim ADays(1 To StaNum, BYear To EYear)
  For i = 1 To StaNum
    For j = BYear To EYear
      AAA(i, j) = -9999
      If FlagTempStat = True Then ADays(i, j) = -9999
    Next j
  Next i


  If FlagTempStat = False Then
    If BMMDD = "01-01" And EMMDD = "12-31" Then '直接从年值数据表中取数据
    
      ReDim AAA_tmp(1 To StaNum, 1 To iyears) '存放各站点、各年值
      For i = 1 To StaNum
        For j = 1 To iyears
          AAA_tmp(i, j) = -9999
        Next j
      Next i
        
      SelField_Yer = SelField_Mon
      Call QueryMulYerData(SelField_Yer, BYear, EYear, AAA_tmp())
      
      For i = 1 To StaNum
        For j = BYear To EYear
          AAA(i, j) = AAA_tmp(i, j - BYear + 1)
        Next j
      Next i
      
    Else   '从日值+月值数据表中取数据完成统计
      Call QueryMulYearData(SelField_Initial, SelField_Restat, SelField_Mon, BYear, EYear, BMMDD, EMMDD, crossYear, AAA(), FlagStat, TextMin, TextMax)  '存入AAA(1 to StaNum)变量
    End If
      
  ElseIf FlagTempStat = True Then  '直接从日值数据表进行统计：2023-03-06
    
    '合计多年份-多日统计值（数据表名，要素字段名，开始年度，结束年度，开始月日，结束月日，是否属于上年度结果，输出统计值，输出日数，统计量，条件最小Value，条件最大Value）
    Call StatMulYearDayData(SelField_Initial, BYear, EYear, BMMDD, EMMDD, False, AAA(), ADays(), FlagStat, TextMin, TextMax)
  End If
  
  For i = 1 To StaNum
    For j = BYear To EYear
'      If FlagStat = "SUM" Or FlagStat = "DAYS" Then '条件日数或条件累计值查询，无符合条件的都为0
      If (FlagStat = "SUM" Or FlagStat = "DAYS") And FlagLimit = True Then '条件日数或条件累计值查询，无符合条件的都为0
        If j >= StaInfo(i).YearSTT Then
          If AAA(i, j) = -9999 Then
            If FlagResult = "0" Then AAA(i, j) = 0 'j年份已经是资料了，若用户选择条件结果设置为0，则条件结果默认为0
          ElseIf AAA(i, j) = 0 Then
            If FlagResult = "null" Then AAA(i, j) = -9999
          End If
        End If
      End If
    Next j
  Next i
  
'************************************* 3、查询所是年份（Byear-Eyear）或某年分各站值：结束 ***********************************


'************************************************4.1、查询累年极值：开始******************************************
  sngSum = 0: sngMax = -9999: sngMin = 999999: iNum = 0

  If FlagStat = "MAX" Or FlagStat = "MIN" Then '查询极值表--最大Value、最小Value、最值日期
    '*******************************赋初始值-9999或0*************************
    ReDim CCC(1 To StaNum), CCCDate(1 To StaNum)
    For i = 1 To StaNum
      CCC(i) = -9999: CCCDate(i) = CDate("1899-09-09")
    Next i
    Call StatMulDayXtrm(SelField_Xtrm, BMMDD, EMMDD, crossYear, CCC(), CCCDate(), FlagStat, TextMin, TextMax) '存入CCC(1 to StaNum),CCCDate(1 to StaNum)变量
    For i = 1 To StaNum
      PRD_STA(i).stat_hist = CCC(i)
      PRD_STA(i).stat_hist_date = CCCDate(i)
    Next i

'************************************************4.2、计算30年均值：开始*********************************************
  Else
    
    '****************** 多年平均值源于国家局下发常年值 *****************
    If Norm_Src = "CMA" Then
      ReDim CCC(1 To StaNum)
      For i = 1 To StaNum
        CCC(i) = -9999
      Next i
      
      ' 2023-3-6 如果不临时统计，则从国家局下发常年值表中获取，否则留空
      If FlagTempStat = False Then
        '**************** 整年度数据（2023-01-06） ******************
        If BMMDD = "01-01" And EMMDD = "12-31" Then
          ReDim YER_STA(1 To StaNum, 1) '存放各站点、常年平均值
          For i = 1 To StaNum
            YER_STA(i, 1) = -9999
          Next i
          
          If FlagLimit = False Then
            If FlagStat = "SUM" Or FlagStat = "DAYS" Then SelField_Yer = SelField_Restat & "_" & FlagStat
            If FlagStat <> "SUM" And FlagStat <> "DAYS" Then SelField_Yer = SelField_Restat
          ElseIf FlagLimit = True Then
            If TextMax <> 999999 Then SelField_Yer = SelField_Restat & "_" & TextMax & FlagStat
            If TextMin <> -9999 Then SelField_Yer = SelField_Restat & "_" & TextMin & FlagStat
            If TextMin = 0.1 Then SelField_Yer = SelField_Restat & "_" & FlagStat '雨量/雨日
          End If
          
'          If FlagStat = "SUM" Or FlagStat = "DAYS" Then SelField_Yer = SelField_Restat & "_" & FlagStat
'          If FlagStat <> "SUM" And FlagStat <> "DAYS" Then SelField_Yer = SelField_Restat
'
'          If TextMax <> 999999 Then SelField_Mon = SelField_Restat & "_" & TextMax & FlagStat
'          If TextMin <> -9999 Then SelField_Mon = SelField_Restat & "_" & TextMin & FlagStat
'          If TextMin = 0.1 Then SelField_Mon = SelField_Restat & "_" & FlagStat '雨量/雨日
          
          Call QueryMulYerRestat(SelField_Yer, "Norm", YER_STA(), XtrmYear_STA(), False)
          For i = 1 To StaNum
            CCC(i) = YER_STA(i, 1)
          Next i
          
        '**************** 非整年度数据（2023-01-30） ******************
        Else
          If FlagLimit = False Then
            If FlagStat = "DAYS" Then  '_DAYS统计量在非年表中无数据
              DoEvents
            ElseIf FlagStat = "SUM" And (SelField_Restat = "R" Or SelField_Restat = "S" Or SelField_Restat = "L_S") Then   'R、S、L_S通过多月+多日数据求和
              Call QueryPeriodNormData(SelField_Restat, BMMDD, EMMDD, crossYear, CCC(), FlagStat, TextMin, TextMax)
            
            ElseIf FlagStat = "AVE" And (SelField_Restat = "TA" Or SelField_Restat = "UA" Or SelField_Restat = "PA" Or SelField_Restat = "DA" Or SelField_Restat = "F10S") Then   '2023-9-14增加
              Call QueryPeriodNormData(SelField_Restat, BMMDD, EMMDD, crossYear, CCC(), FlagStat, TextMin, TextMax)
            
            Else  '其他统计量直接查日值进行统计
              Call QueryPeriodNormData2(SelField_Restat, BMMDD, EMMDD, crossYear, CCC(), FlagStat, TextMin, TextMax)
            End If
          
          ElseIf FlagLimit = True Then
          
            If TextMin = 0.1 Then  '2024-9-30 Add：雨量下限为0.1时，就当作不设限处理
              TextMin = -9999
              Call QueryPeriodNormData(SelField_Restat, BMMDD, EMMDD, crossYear, CCC(), FlagStat, TextMin, TextMax)
            Else
              DoEvents
            End If
          
          End If
        End If
      End If
        
        
      For i = 1 To StaNum
        PRD_STA(i).stat_hist = CCC(i)
      Next i
        
    '****************** 多年平均值源于CROSS实时统计数据 *****************
    Else
      ReDim Norm_Avg(1 To StaNum), Norm_Sum(1 To StaNum)
      For i = 1 To StaNum '从第一个站点开始统计每个站点的平均值
        Norm_Avg(i) = -9999: Norm_Sum(i) = 0: Norm_num = 0
        If StaInfo(i).StaType = "SURF" Then
          For j = BYear3 To EYear3
            If AAA(i, j) <> -9999 Then Norm_Sum(i) = Norm_Sum(i) + AAA(i, j): Norm_num = Norm_num + 1
          Next j
        ElseIf StaInfo(i).StaType = "AWST" Then
         For j = BYear4 To EYear4
            If AAA(i, j) <> -9999 Then Norm_Sum(i) = Norm_Sum(i) + AAA(i, j): Norm_num = Norm_num + 1
          Next j
        End If
'        If Norm_num <> 0 Then Norm_Avg(i) = Round(Norm_Sum(i) / Norm_num, 1)
        If Norm_num <> 0 Then Norm_Avg(i) = Format(Norm_Sum(i) / Norm_num, "0.0")
        PRD_STA(i).stat_hist = Norm_Avg(i)
      Next i
      
    End If
      
    '************************************************5、计算多年平均/某年距平（百分率）：开始*********************************************
    For i = 1 To StaNum
        If PRD_STA(i).stat <> -9999 And PRD_STA(i).stat_hist <> -9999 Then
          PRD_STA(i).stat_diff1 = PRD_STA(i).stat - PRD_STA(i).stat_hist
        End If
        If PRD_STA(i).stat <> -9999 And PRD_STA(i).stat_year <> -9999 Then
          PRD_STA(i).stat_diff2 = PRD_STA(i).stat - PRD_STA(i).stat_year
        End If
        
        If FlagStat = "SUM" Then
          If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Or SelField_Initial = "S" Then
              If PRD_STA(i).stat <> -9999 And PRD_STA(i).stat_hist <> -9999 Then
                If PRD_STA(i).stat_hist <> 0 Then PRD_STA(i).stat_diff1 = 100 * (PRD_STA(i).stat - PRD_STA(i).stat_hist) / PRD_STA(i).stat_hist
                If PRD_STA(i).stat_hist = 0 Then PRD_STA(i).stat_diff1 = -9999
              End If
              If PRD_STA(i).stat <> -9999 And PRD_STA(i).stat_year <> -9999 Then
                If PRD_STA(i).stat_year <> 0 Then PRD_STA(i).stat_diff2 = 100 * (PRD_STA(i).stat - PRD_STA(i).stat_year) / PRD_STA(i).stat_year
                If PRD_STA(i).stat_year = 0 Then PRD_STA(i).stat_diff2 = -9999
              End If
          End If
        End If
    Next i

  End If
  
End Sub



'********************************************* 多年任意时段数据结果排序（2022-03-28）*****************************************
Private Sub Rank_PeriodData(AAAIn!(), BBBIn!, RankOut As Period_Result) '输入某站多年一维数据、某站某年数据、结果输出至PRD_STA中
  Dim EEE!() '各站点多年值排序时，从小到大的值存入该变量
  Dim RankYear%() '各站点多年值排序时，从小到大的值对应年份存入该变量
  Dim TempAAA! '各站点多年值排序时，大小值交换时的中间变量
  Dim TempRankYear% '各站点多年值排序时，大小值对应的年份交换时的中间变量
  Dim RankA%, RankD%   '各站点从大到小排名，从小到大排名的排位序号


    '判断排序的起始年份
    For j = BYear2 To EYear2
      If AAAIn(j) <> -9999 Then BYear_rank = j: Exit For
    Next j
    
    k = 0
    
    For j = BYear2 To EYear2 '将排序年中不等于-9999的AAA(i,j)赋值给EEE(k)
      If AAAIn(j) <> -9999 Then
        k = k + 1
        ReDim Preserve EEE(1 To k): ReDim Preserve RankYear(1 To k)
        EEE(k) = AAAIn(j): RankYear(k) = j
      End If
    Next j

    '判断排序的结束年份
    If k <> 0 Then
      For j = EYear2 To BYear2 Step -1
        If AAAIn(j) <> -9999 Then EYear_rank = j: Exit For
      Next j
    ElseIf k = 0 Then
      Exit Sub  '该站点没是任何历年值
    End If

    iYears_rank = k  '排名年份数


    '将EEE(j)从小到大排序（选择排序法），存入变量EEE(j)；Rank年份(j)为从小到大值对应的年份
    For j = LBound(EEE) To UBound(EEE) - 1
      For k = j + 1 To UBound(EEE)
        If EEE(k) < EEE(j) Then
          TempAAA = EEE(k): TempRankYear = RankYear(k)
          EEE(k) = EEE(j): RankYear(k) = RankYear(j)
          EEE(j) = TempAAA: RankYear(j) = TempRankYear
        End If
      Next k
    Next j

    For j = LBound(EEE) To UBound(EEE)  '从小到大排序的排位
      If BBBIn <> -9999 Then
        If BBBIn <= EEE(j) Then
          RankD = j: Exit For
        End If
        RankD = j
      End If
    Next j

    For j = UBound(EEE) To LBound(EEE) Step -1 '从大到小排序的排位
      If BBBIn <> -9999 Then
        If BBBIn >= EEE(j) Then
          RankA = UBound(EEE) - j + 1: Exit For
        End If
        RankA = UBound(EEE) - j + 1
      End If
    Next j

    If BBBIn <> -9999 Then
      RankOut.rank1 = RankA
      RankOut.rank2 = RankD
      RankOut.rank_years = iYears_rank & "(" & BYear_rank & "-" & EYear_rank & ")"
      RankOut.maxyear = RankYear(UBound(RankYear))
      RankOut.maxvalue = EEE(UBound(EEE))
      RankOut.minyear = RankYear(LBound(RankYear))
      RankOut.minvalue = EEE(LBound(EEE))
    End If


End Sub



'*********************************************多要素任意时段资料统计>>>开始 2022-4-17*****************************************
Public Sub CmdMeteoPeriod2()

  Screen.MousePointer = 11
  
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

  If Left(EDate, 4) - Left(BDate, 4) > 1 Then
    MsgBox "Date range cannot exceed 1 year", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If
  
'  '排名起止年份（可自定义）
'  B年份2 = ComboB年份2.Text  '起始年份
'  E年份2 = ComboE年份2.Text  '结束年份
'  If BYear2 > EYear2 Then MsgBox "Invalid ranking year", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub

  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart

  '当年值年份
  iYear = Left(BDate, 4)
  'a given yearValueYear
  iYear3 = ComboYear3.Text
  
  If Left(EDate, 4) - Left(BDate, 4) = 0 Then '同年度
    crossYear = False
    BDate3 = CStr(iYear3) & "-" & BMMDD
    If EMMDD <> "02-29" Then
      EDate3 = CStr(iYear3) & "-" & EMMDD
    
    '2024-2-27判断结束日期为02-29的某年是否为平年
    ElseIf EMMDD = "02-29" Then
      If iYear3 Mod 4 <> 0 Then EDate3 = CStr(iYear3) & "-02-28"
      If iYear3 Mod 4 = 0 Then EDate3 = CStr(iYear3) & "-02-29"
    End If
    
  ElseIf Left(EDate, 4) - Left(BDate, 4) = 1 Then '跨年度：年份数少一年
    crossYear = True
    BDate3 = CStr(iYear3) & "-" & BMMDD
    If EMMDD <> "02-29" Then
      EDate3 = CStr(iYear3 + 1) & "-" & EMMDD
  
    '2024-2-27判断结束日期为02-29的某年是否为平年
    ElseIf EMMDD = "02-29" Then
      If (iYear3 + 1) Mod 4 <> 0 Then EDate3 = CStr(iYear3 + 1) & "-02-28"
      If (iYear3 + 1) Mod 4 = 0 Then EDate3 = CStr(iYear3 + 1) & "-02-29"
    End If
  
  End If
  
  
  ORAConn2.Open
  ORAConn2.CursorLocation = 3
 

  '****************************2022-04-17*****************************
  NumFields = 0
  For i = 0 To List_Col.ListCount - 1
    If List_Col.Selected(i) = True Then
      NumFields = NumFields + 1
      ReDim Preserve SelFields_Initial(1 To NumFields)
      ReDim Preserve SelNames_Initial(1 To NumFields)
      ReDim Preserve SelFields_Restat(1 To NumFields)
      ReDim Preserve FlagStats(1 To NumFields)
      SelFields_Initial(NumFields) = COL_Initial(i + 1).Col_ID
      SelNames_Initial(NumFields) = COL_Initial(i + 1).Col_name
      SelFields_Restat(NumFields) = COL_Restat(i + 1).Col_ID
        
      Call Stat_MeteoPeriod(FrmMeteoPeriod, SelFields_Restat(NumFields), FlagStats(NumFields))
    End If
  Next i

  If NumFields = 0 Then
    MsgBox "Please select at least one element", vbInformation, "Notice"
    ORAConn2.Close
    frmWait.Hide: DoEvents
    Screen.MousePointer = 1
    Exit Sub
  End If

  Call TableIni_MeteoPeriod2(FrmMeteoPeriod)

  
'  If CheckLimit.值 = Unchecked Then  '已经设定了条件查询为unchecked
    TextMin = ComboMin.Text: TextMax = ComboMax.Text
    If TextMin = -9999 And TextMax = 999999 Then
      FlagLimit = False
    Else
      FlagLimit = True
    End If
    If OptionStat(1).Value = True Then FlagStat = "AVE"
    If OptionStat(2).Value = True Then FlagStat = "MAX"
    If OptionStat(3).Value = True Then FlagStat = "MIN"
    If OptionStat(4).Value = True Then FlagStat = "SUM"
    If OptionStat(7).Value = True Then FlagStat = "DAYS"
'  ElseIf CheckLimit.Value = Checked Then
'    FlagLimit = True: TextMin = ComboMin.Text: TextMax = ComboMax.Text
'    If OptionLimit(2).值 = True Then '条件要素查询只能查累计值(条件总量)
'      FlagStat = "SUM"
'    ElseIf OptionLimit(1).值 = True Then '条件日数查询：只能查DAYS
'      FlagStat = "DAYS"
'    End If
'  End If

'********************************2022-3-24修改阈值选项为list，不可以进行临时查询，FlagTempStat一定是False***************************
  FlagTempStat = False

  
  ' 判断从数据库中读取的数据年限Byear-Eyear
  BYear3 = Left(Years_NormSURF, 4)
  EYear3 = Right(Years_NormSURF, 4)
  
  BYear4 = Left(Years_NormAWST, 4)
  EYear4 = Right(Years_NormAWST, 4)

  BYear = BYear3
  If BYear4 < BYear Then BYear = BYear4
  EYear = EYear3
  If EYear4 > EYear Then EYear = EYear4
  
  iyears = CInt(EYear) - CInt(BYear) + 1
  
  '**************************判断是资料的第一个年份yyyy存入StaInfo(i).年份STT********************************************
  For i = 1 To StaNum
    For j = StaInfo(i).YearSTT To EYear '从该站是记录来的第一年开始判断
      If crossYear = False Then '同年度
        BDate2 = CStr(j) & "-" & BMMDD
        If EMMDD <> "02-29" Then
          EDate2 = CStr(j) & "-" & EMMDD
        ElseIf EMMDD = "02-29" Then
          If j Mod 4 <> 0 Then EDate2 = CStr(j) & "-02-28"
          If j Mod 4 = 0 Then EDate2 = CStr(j) & "-02-29"
        End If
          
      ElseIf crossYear = True Then '跨年度：年份数少一年
        BDate2 = CStr(j) & "-" & BMMDD
        If EMMDD <> "02-29" Then
          EDate2 = CStr(j + 1) & "-" & EMMDD
        ElseIf EMMDD = "02-29" Then
          If (j + 1) Mod 4 <> 0 Then EDate2 = CStr(j + 1) & "-02-28"
          If (j + 1) Mod 4 = 0 Then EDate2 = CStr(j + 1) & "-02-29"
        End If
      End If
      If CDate(BDate2) >= CDate(StaInfo(i).DateSTT) Then  '如果要查询的开始日期在是效资料日期之后，则该站点从这一年开始是资料可查
        StaInfo(i).YearSTT = j
        Exit For
      End If
    Next j
  Next i
  
  
  '按已选的要素循环统计
  For k = 1 To UBound(SelFields_Restat)
    FlagStat = FlagStats(k)
    SelField_Initial = SelFields_Initial(k)
    SelField_Restat = SelFields_Restat(k)
    SelField_Mon = SelField_Restat & "_" & FlagStat
    If FlagStat = "MAX" Or FlagStat = "MIN" Then SelField_Xtrm = SelField_Restat & "_" & FlagStat
  
    '************************************* 一、查询所是站点任意时段值  ***********************************
    ReDim PRD_STA(1 To StaNum)
    For i = 1 To StaNum
      PRD_STA(i).stacode = StaInfo(i).stacode
      PRD_STA(i).staname = StaInfo(i).staname
      
      PRD_STA(i).stat = -9999 '当年统计值
      PRD_STA(i).stat_date = CDate("1899-09-09") '极值出现日期
        
      PRD_STA(i).stat_hist = -9999 '累年值
      PRD_STA(i).stat_hist_date = CDate("1899-09-09") '累年极值出现日期
      PRD_STA(i).stat_diff1 = -9999 '多年平均距平
        
    Next i
    Call QueryPeriodData
  
  
    '************************************* 二、处理所是站点任意时段结果  ***********************************
    
    '************************************* 情形一：查询结果为站点尺度  ***********************************
      If FlagZone = "STA" Then '站点尺度
    
    '************************************* 1、计算所是行(站点)的统计值  ***********************************
        '统计所是行的平均/最大/最小Value
        Call Stat_ROWS_Period(PRD_STA, StaNum, FlagStat, SelField_Initial)
    '************************************* 2、计算历年所是站的平均值  ***********************************
        Call Stat_ROWS_AAA(AAA, AAA_avg, BBB, BBB_avg)
    '************************************* 3、填充上方表格  ***********************************
        Call HFGrid1_Fill_Period3(FrmMeteoPeriod, PRD_STA, k, FlagStat)
    '************************************* 4、填充下方表格 ***********************************
        Call HFGrid2_Fill_Period3(FrmMeteoPeriod, ROWS_Stat_Period, k, FlagStat)
    
    '************************************* 情形二：查询结果为区域尺度  ***********************************
      Else
    
    '************************************* 1、分区域统计所是行列值  ***********************************
        Call Stat_ZONE_Period(PRD_STA, FlagStat, SelField_Initial)
    '************************************* 2、计算历年各区的平均值  ***********************************
        Call Stat_ZONE_AAA(AAA, AAA_ZONE, BBB, BBB_ZONE)
    '************************************* 3、填充上/中/下方表格值  ***********************************
        Call HFGrid3_fill_Period3(FrmMeteoPeriod, k, FlagStat)
        
      End If
  Next k

  ORAConn2.Close
  Call TimeStop
  frmWait.Hide: DoEvents
  
  Screen.MousePointer = 1
  
End Sub
'*********************************************单要素逐日资料统计>>>结束*****************************************


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
Public Sub CmdMap_MeteoPeriod()
  If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Then
    Call Map_Value(FrmMeteoPeriod, "RYB")
  Else
    Call Map_Value(FrmMeteoPeriod, "BYR")
  End If
End Sub


'********************保存窗体内可见的HFGrid***************************
Public Sub CmdOutput_MeteoPeriod()
  Call Output_HFGrid(FrmMeteoPeriod, "")
End Sub


Private Sub ImageSwitchNorm_Click()
  If Norm_Src = "CMA" Then
    Norm_Src = "GRMC"
    FrmMain.StatusBar1.Panels(4).Text = "National station normals: Climate Center " & Years_NormSURF & "Yr Avg"
    If FrmMeteoPeriodExt.Visible = True Then
      FrmMeteoPeriodExt.StatusBar1.Panels(3).Text = "National station normals: Climate Center " & Years_NormSURF & "Yr Avg"
    End If
  ElseIf Norm_Src = "GRMC" Then
    Norm_Src = "CMA"
    FrmMain.StatusBar1.Panels(4).Text = "National station normals: CMA-issued normals"
    If FrmMeteoPeriodExt.Visible = True Then
      FrmMeteoPeriodExt.StatusBar1.Panels(3).Text = "National station normals: CMA-issued normals"
    End If
  End If
End Sub

Private Sub List_Col_Click()
  If List_Col.SelCount = 0 Then
    List_Col.Selected(mLastSelIndex) = True
    MsgBox "Please select at least one element", vbInformation, "Notice"
    Exit Sub
  End If

  If List_Col.SelCount = 1 Then '单要素查询
    For i = 0 To List_Col.ListCount - 1
      If List_Col.Selected(i) = True Then
        SelField_Initial = COL_Initial(i + 1).Col_ID
        SelField_Restat = COL_Restat(i + 1).Col_ID
        mLastSelIndex = i
        Exit For
      End If
    Next i
    
    Call Stat_MeteoPeriod(FrmMeteoPeriod, SelField_Restat, FlagStat)
    Call Option_MeteoPeriod(FrmMeteoPeriod, SelField_Restat)
    Call TableIni_MeteoPeriod(FrmMeteoPeriod, SelField_Initial)
    MultiSel = False
    
    ComboMin.Enabled = True: ComboMax.Enabled = True
      
    
    '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
    If GuestMode And List_Col.Selected(0) = True Then
      Call ImportData.LoadGuestCSV(FrmMeteoPeriod.HFGrid1, FrmMeteoPeriod.HFGrid2, "DataMeteoPeriod.csv", 3)
    End If
    
  Else '多要素查询
'      CheckLimit.Value = Unchecked
      NumFields = 0
      For i = 0 To List_Col.ListCount - 1
        If List_Col.Selected(i) = True Then
          NumFields = NumFields + 1
          ReDim Preserve SelFields_Initial(1 To NumFields)
          ReDim Preserve SelNames_Initial(1 To NumFields)
          ReDim Preserve SelFields_Restat(1 To NumFields)
          ReDim Preserve FlagStats(1 To NumFields)
          SelFields_Initial(NumFields) = COL_Initial(i + 1).Col_ID
          SelNames_Initial(NumFields) = COL_Initial(i + 1).Col_name
          SelFields_Restat(NumFields) = COL_Restat(i + 1).Col_ID
          mLastSelIndex = i
          
          Call Stat_MeteoPeriod(FrmMeteoPeriod, SelFields_Restat(NumFields), FlagStats(NumFields))
          
        End If
      Next i
      Call TableIni_MeteoPeriod2(FrmMeteoPeriod)
      MultiSel = True
      
      ' 阈值筛选仅对单要素是意义（不同要素的量纲/取值范围不可比），一旦选择两个及以上要素，重置为不限并禁用
      ComboMin.Text = "-9999": ComboMax.Text = "999999"
      ComboMin.Enabled = False: ComboMax.Enabled = False
  
  End If
  
  
End Sub




Private Sub OptionStat_Click(Index As Integer)
  If List_Col.SelCount = 1 Then '单要素查询
    If OptionStat(1).Value = True Then
      FlagStat = "AVE"
      OptionNull(1).Value = True
'      CheckNull.Value = Checked
    ElseIf OptionStat(2).Value = True Then
      FlagStat = "MAX"
      OptionNull(1).Value = True
'      CheckNull.Value = Checked
    ElseIf OptionStat(3).Value = True Then
      FlagStat = "MIN"
      OptionNull(1).Value = True
'      CheckNull.Value = Checked
    ElseIf OptionStat(4).Value = True Then
      FlagStat = "SUM"
      OptionNull(0).Value = True
'      CheckNull.Value = Unchecked
    ElseIf OptionStat(7).Value = True Then
      FlagStat = "DAYS"
      OptionNull(0).Value = True
'      CheckNull.Value = Unchecked
    End If
  
'    Call Option_MeteoPeriod(FrmMeteoPeriod, SelField_Restat)
    Call TableIni_MeteoPeriod(FrmMeteoPeriod, SelField_Initial)
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

