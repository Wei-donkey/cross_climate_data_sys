VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmMeteoPeriodYer 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Year-by-Year Custom-Period Statistics"
   ClientHeight    =   13530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   28380
   Icon            =   "FrmMeteoPeriodYer.frx":0000
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
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid2 
      Height          =   1380
      Left            =   120
      TabIndex        =   17
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
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFEFDF&
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   28095
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Statistic"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   14280
         TabIndex        =   27
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
            TabIndex        =   32
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
            TabIndex        =   31
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
            TabIndex        =   30
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
            TabIndex        =   29
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
            TabIndex        =   28
            Top             =   720
            Width           =   1695
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Threshold"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   17520
         TabIndex        =   22
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
            TabIndex        =   24
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
            TabIndex        =   23
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
            TabIndex        =   26
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
            TabIndex        =   25
            Top             =   795
            Width           =   255
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Condition not met"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   19080
         TabIndex        =   19
         Top             =   120
         Width           =   1815
         Begin VB.OptionButton OptionNull 
            BackColor       =   &H00DFEFDF&
            Caption         =   "is zero"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   21
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
            TabIndex        =   20
            Top             =   720
            Width           =   1335
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Highlight Results"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   20880
         TabIndex        =   11
         Top             =   120
         Width           =   2175
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
            TabIndex        =   14
            Top             =   720
            Width           =   855
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
            TabIndex        =   13
            Top             =   360
            Width           =   855
         End
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
            TabIndex        =   12
            Top             =   720
            Width           =   675
         End
         Begin VB.Label Label2 
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
            TabIndex        =   16
            Top             =   795
            Width           =   255
         End
         Begin VB.Label Label1 
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
            TabIndex        =   15
            Top             =   435
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Start/End Year/Date"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   12240
         TabIndex        =   8
         Top             =   120
         Width           =   2055
         Begin VB.ComboBox ComboBMMDD 
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
            Left            =   1080
            TabIndex        =   34
            Text            =   "mm-dd"
            Top             =   360
            Width           =   900
         End
         Begin VB.ComboBox ComboEMMDD 
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
            Left            =   1080
            TabIndex        =   33
            Text            =   "mm-dd"
            Top             =   720
            Width           =   900
         End
         Begin VB.ComboBox ComboBYear 
            Appearance      =   0  'Flat
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
            TabIndex        =   10
            Text            =   "yyyy"
            Top             =   360
            Width           =   840
         End
         Begin VB.ComboBox ComboEYear 
            Appearance      =   0  'Flat
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
            TabIndex        =   9
            Text            =   "yyyy"
            Top             =   720
            Width           =   840
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
         Left            =   23040
         TabIndex        =   2
         Top             =   120
         Width           =   4935
      End
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
Attribute VB_Name = "FrmMeteoPeriodYer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mLastRightClickGrid As Object '最近一次被点击（任意按钮）的表格；供mnuGridCopy_Click及ChartData查找图表数据源时使用
Private mEditGrid As Object '判断当前正在被TextEdit编辑的是哪个Grid（HFGrid1/HFGrid2/HFGrid3(Index)之一）

'Public BYYYY$, EYYYY$, iYers%
Public BYear%, EYear%, iYers% 'B年份、E年份为从数据库取出的数据起止年份（1951年至今）,iYers为读取资料起止年数
Public XtrmYear As Boolean '是否显示极值出现年份
Dim XtrmYear_STA() As Integer '各站点的极值出现年份

'Dim BDate$, EDate$, iDays% '用户自定义的当年查询起始日期
Dim BDate2$, EDate2$ '用于判断是效资料的起止年份
'Dim BDate3$, EDate3$ '用户自定义对应的某年查询起始日期
Dim BMMDD$, EMMDD$ '用户自定义的查询起止月日


Dim BMM$, BDD$, EMM$, EDD$ '用户自定义的查询起止月、日


Dim AAA!() '按年度循环读oracle，统计值存入二维数组AAA
Dim AAA_tmp!() '2024-4-12Add
Dim ADays%() ' 存放多年度所是站点条件日数：多日结果


Dim crossYear As Boolean '查询时段是否跨年度


Dim SelField_Initial$ '选择要素的初始名
Dim SelField_Restat$ '选择要素的再统计名

Dim SelField_Day$ '选择要素的日间隔字段
Dim SelField_Mon$ '选择要素的月值表字段

Dim SelField_Yer$ '选择要素的年值表字段（2024-4-12添加）

Dim SelField_Xtrm$ '选择要素的极值表字段
Dim SelField_Norm$ '选择要素的平均态字段

Public FlagStat$ '统计量标识
Dim FlagLimit As Boolean '条件查询标识


Dim i%, j%, k%, l%
Dim TextMin!, TextMax! '条件查询的上下限

Dim FlagTempStat As Boolean  '临时统计：不从月值表读取的条件查询--长耗时：2017-3-23
Dim FlagResult As String '2023-03-14：不符合条件的结果设为0或null，0参与多站点平均，null不参与平均


Private Sub OptionNull_Click(Index As Integer)
  If OptionNull(0).Value = True Then
    FlagResult = "0"
  ElseIf OptionNull(1).Value = True Then
    FlagResult = "null"
  End If
End Sub

Private Sub CheckHLight_Click()
  If CheckHLight.Value = Checked Then
    ComboMin_HL.Enabled = True: ComboMax_HL.Enabled = True
  ElseIf CheckHLight.Value = Unchecked Then
    ComboMin_HL.Enabled = False: ComboMax_HL.Enabled = False
  End If
End Sub

Private Sub Form_Load()
  Dim i%
  Screen.MousePointer = 11
  FrmMain.FormAdd
   
  For i = 0 To 30
    ComboBMMDD.AddItem Format(Date - 30 + i, "mm-dd")
    ComboEMMDD.AddItem Format(Date - 30 + i, "mm-dd")
  Next i
  ComboBMMDD.Text = Format(Date - 1, "mm-01")
  ComboEMMDD.Text = Format(Date - 1, "mm-dd")
  
  ComboBYear.Clear: ComboEYear.Clear
  
  For i = 1951 To Year(Date)
    ComboBYear.AddItem i
    ComboEYear.AddItem i
  Next i
  ComboBYear.Text = Year(Date) - 29
  ComboEYear.Text = Year(Date)
 
  Call ReadColInfo("COLS_Initial.ini")
  COL_Initial = COL_tmp
  
  Call ReadColInfo("COLS_ReStat.ini")
  COL_Restat = COL_tmp
 
  '****************************2022-03-08从ini配置文件中读取字段名*****************************
  List_Col.Clear
  For i = 1 To UBound(COL_Restat)
    List_Col.AddItem COL_Restat(i).Col_name
  Next i
'  List_Col.Selected(0) = True
  SelField_Initial = COL_Initial(1).Col_ID
  SelField_Restat = COL_Restat(1).Col_ID
  FlagStat = "AVE"
  OptionNull(1).Value = True
   
  Call ITimesInitial(FrmMeteoPeriodYer)
  Call TableIni_Meteo(FrmMeteoPeriodYer, FrmMeteoPeriodYer.BYear, FrmMeteoPeriodYer.iYers)
  List_Col.Selected(0) = True
  
  Screen.MousePointer = 1

End Sub

'2026-08-19 v2.14 界面自适应：主窗体尺寸变化时，让本窗体的表格和选项区跟着调整）
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



'*********************************************单要素任意时段逐年资料统计>>>开始*****************************************
Public Sub CmdMeteoPeriodYer()

  Screen.MousePointer = 11
  
  Call ITimesInitial(FrmMeteoPeriodYer)
  If FlagInputErr = True Then Screen.MousePointer = 1: Exit Sub
  Call TableIni_Meteo(FrmMeteoPeriodYer, FrmMeteoPeriodYer.BYear, FrmMeteoPeriodYer.iYers)
  
  BYear = ComboBYear.Text  '起始年份
  EYear = ComboEYear.Text  '结束年份
    
   '************2023-1-18添加对输入日期的精准判断**********
  Dim tempStr$()
  
  tempStr = Split(ComboBMMDD.Text, "-")
  BMM = tempStr(0): BDD = tempStr(1)
  BMMDD = Format(BMM, "00") & "-" & Format(BDD, "00")
  
  tempStr = Split(ComboEMMDD.Text, "-")
  EMM = tempStr(0): EDD = tempStr(1)
  EMMDD = Format(EMM, "00") & "-" & Format(EDD, "00")
    
  If CInt(BMM) > 12 Or CInt(BMM) < 1 Then MsgBox "Invalid start month", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  If CInt(EMM) > 12 Or CInt(EMM) < 1 Then MsgBox "Invalid end month", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub

  Select Case CInt(BMM)
    Case 1, 3, 5, 7, 8, 10, 12
      If CInt(BDD) > 31 Or CInt(BDD) < 0 Then MsgBox "Invalid start date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    Case 4, 6, 9, 11
      If CInt(BDD) > 30 Or CInt(BDD) < 0 Then MsgBox "Invalid start date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    Case 2
      If CInt(BDD) > 29 Or CInt(BDD) < 0 Then MsgBox "Invalid start date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End Select
  
  Select Case CInt(EMM)
    Case 1, 3, 5, 7, 8, 10, 12
      If CInt(EDD) > 31 Or CInt(EDD) < 0 Then MsgBox "Invalid end date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    Case 4, 6, 9, 11
      If CInt(EDD) > 30 Or CInt(EDD) < 0 Then MsgBox "Invalid end date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    Case 2
      If CInt(EDD) > 29 Or CInt(EDD) < 0 Then MsgBox "Invalid end date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End Select
      
'  '********2021-9-14 其他所是年份均不查02-29*************
'  If BMMDD = "02-29" Then MsgBox "不支持2月29日", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
'  If EMMDD = "02-29" Then MsgBox "不支持2月29日", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
  '********2021-9-14 其他所是年份均不查02-29*************
  If BMMDD = "02-29" Then MsgBox "Start date cannot be Feb 29", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
  If BYear > EYear Then MsgBox "Invalid start/end year", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub

    
  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart

  
  If BMMDD <= EMMDD Then '同年度
    crossYear = False
  ElseIf BMMDD > EMMDD Then  '跨年度：年份数少一年
    crossYear = True
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
  
'  If FlagLimit = False Then
'    If (FlagStat = "SUM" Or FlagStat = "DAYS") Then
'      FlagTempStat = True
'      If List_Col.Selected(4) = True Or List_Col.Selected(5) = True Or List_Col.Selected(6) = True Or List_Col.Selected(7) = True _
'      Or List_Col.Selected(32) = True Or List_Col.Selected(35) = True Then '20时/08时/20-08时/08-20时日雨量、蒸发、日照
'        FlagTempStat = False  '非条件查询，不需要临时统计
'        SelField_Mon = SelField_Restat & "_" & FlagStat
'      End If
'    Else
'      FlagTempStat = False
'      SelField_Mon = SelField_Restat & "_" & FlagStat
'    End If
'
'  ' 2023-03-03 如果条件查询，则需要判断是否进行临时统计
'  ElseIf FlagLimit = True Then
'
'    FlagTempStat = True '默认需要临时统计
'
'    If List_Col.Selected(0) = True Then '平均气温
'      If TextMin = 10 And TextMax = 999999 Then FlagTempStat = False
'    ElseIf List_Col.Selected(1) = True Then  '最高气温
'      If TextMin = 35 And TextMax = 999999 Then FlagTempStat = False
'    ElseIf List_Col.Selected(2) = True Then  '最低气温
'      If TextMin = -9999 And TextMax = 5 Then FlagTempStat = False
'    ElseIf List_Col.Selected(4) = True Then  '20时日降水量
'      If (TextMin = 0.1 Or TextMin = 10 Or TextMin = 25 Or TextMin = 50) And TextMax = 999999 Then FlagTempStat = False
'    End If
'
'    If FlagTempStat = False Then
'      If TextMax <> 999999 Then SelField_Mon = SelField_Restat & "_" & TextMax & FlagStat
'      If TextMin <> -9999 Then SelField_Mon = SelField_Restat & "_" & TextMin & FlagStat
'      If TextMin = 0.1 Then SelField_Mon = SelField_Restat & "_" & FlagStat '雨量/雨日
'    End If
'  End If
    
    
    
        
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
    
  

'***************************************** 1、查询所是年度（Byear-Eyear）各站值：开始  ***********************************
'*******************************赋初始值-9999或0*************************
  ReDim AAA(1 To StaNum, BYear To EYear) '存放各站点、各年份统计值（包括平均值、最大Value、最小Value、和值…）
  If FlagTempStat = True Then ReDim ADays(1 To StaNum, BYear To EYear)
  For i = 1 To StaNum
    For j = BYear To EYear
      AAA(i, j) = -9999
      If FlagTempStat = True Then ADays(i, j) = -9999
    Next j
  Next i

  If FlagTempStat = False Then '从日值+月值数据表中取数据完成统计
    
    If BMMDD = "01-01" And EMMDD = "12-31" Then '直接从年值数据表中取数据
    
      ReDim AAA_tmp(1 To StaNum, 1 To iYers) '存放各站点、各年值
      For i = 1 To StaNum
        For j = 1 To iYers
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
 
  
 '************************************* 2、存放所是站点多年值  ***********************************
  ReDim YER_STA(1 To StaNum, 1 To iYers) '存放各站点、各年值
  For i = 1 To StaNum
    For j = 1 To iYers
      If (FlagStat = "SUM" Or FlagStat = "DAYS") And FlagLimit = True Then '条件日数或条件累计值查询，无符合条件的都为0
        If BYear + j - 1 >= StaInfo(i).YearSTT Then
          If AAA(i, BYear + j - 1) = -9999 Then
            If FlagResult = "0" Then AAA(i, BYear + j - 1) = 0 'j年份已经是资料了，若用户选择条件结果设置为0，则条件结果默认为0
          ElseIf AAA(i, BYear + j - 1) = 0 Then
            If FlagResult = "null" Then AAA(i, BYear + j - 1) = -9999
          End If
        End If
      End If

      YER_STA(i, j) = AAA(i, BYear + j - 1)
      If XtrmYear = True Then XtrmYear_STA(i, j) = 1899 '默认极值年份1899年
    Next j
  Next i
 

'************************************* 二、处理所是站点任意时段结果  ***********************************

'************************************* 情形一：查询结果为站点尺度  ***********************************
  If FlagZone = "STA" Then '站点尺度
  
'************************************* 2、计算所是行和所是列的统计值  ***********************************
    '统计所是行的平均/最大/最小Value
    Call Stat_ROWS(YER_STA, iYers, StaNum)
    '统计所是列的平均/最大/最小Value
    Call Stat_COLS(YER_STA, iYers, StaNum, "AVE")
    
    '统计列的3列统计值的平均/最大/最小Value
    Call Stat_COLS_ROWS(StaNum)
    
'************************************* 3、填充上方表格  ***********************************
    Call HFGrid1_Fill(FrmMeteoPeriodYer, YER_STA, iYers, XtrmYear_STA)
'************************************* 4、填充下方表格 ***********************************
    Call HFGrid2_Fill(FrmMeteoPeriodYer, ROWS_Stat, iYers)
  
'************************************* 情形二：查询结果为区域尺度  ***********************************
  Else
  
'************************************* 2、分区域统计所是行列值  ***********************************
    Call Stat_ZONE(YER_STA, iYers)
    
'************************************* 3、填充上/中/下方表格值  ***********************************
    Call HFGrid3_fill(FrmMeteoPeriodYer, Zone_AVG, iYers, 1)
    Call HFGrid3_fill(FrmMeteoPeriodYer, Zone_MAX, iYers, 2)
    Call HFGrid3_fill(FrmMeteoPeriodYer, Zone_MIN, iYers, 3)
  End If


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
  Call Sort_Cols_Rows(FrmMeteoPeriodYer.HFGrid1, shift)
End Sub

Private Sub HFGrid2_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid2
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Rows(FrmMeteoPeriodYer.HFGrid2, shift)
End Sub

'********************对HFGrid3中的数据进行排序***************************
Private Sub HFGrid3_MouseUp(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid3(Index)
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols_Rows(FrmMeteoPeriodYer.HFGrid3(Index), shift)
End Sub


'********************利用HFGrid中的数据绘制地图***************************
Public Sub CmdMap_MeteoPeriodYer()
  If SelField_Initial = "R" Or SelField_Initial = "R08" Or SelField_Initial = "R20_08" Or SelField_Initial = "R08_20" Then
    Call Map_Value(FrmMeteoPeriodYer, "RYB")
  Else
    Call Map_Value(FrmMeteoPeriodYer, "BYR")
  End If
End Sub


'********************保存窗体内可见的HFGrid***************************
Public Sub CmdOutput_MeteoPeriodYer()
  Call Output_HFGrid(FrmMeteoPeriodYer, SelField_Initial)
End Sub


Private Sub List_Col_Click()
  For i = 0 To List_Col.ListCount - 1
    If List_Col.Selected(i) = True Then
      SelField_Restat = COL_Restat(i + 1).Col_ID
      Exit For
    End If
  Next i
  
  Call Stat_MeteoPeriod(FrmMeteoPeriodYer, SelField_Restat, FlagStat)
  Call Option_MeteoPeriod(FrmMeteoPeriodYer, SelField_Restat)
  Call TableIni_Meteo(FrmMeteoPeriodYer, FrmMeteoPeriodYer.BYear, FrmMeteoPeriodYer.iYers)
  CheckHLight.Value = Unchecked
  
  '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
  If GuestMode And List_Col.Selected(0) = True Then
    Call ImportData.LoadGuestCSV(FrmMeteoPeriodYer.HFGrid1, FrmMeteoPeriodYer.HFGrid2, "DataMeteoPeriodYer.csv", 3)
  End If
  
End Sub


Private Sub OptionStat_Click(Index As Integer)
    If OptionStat(1).Value = True Then
      FlagStat = "AVE": OptionNull(1).Value = True
    ElseIf OptionStat(2).Value = True Then
      FlagStat = "MAX":  OptionNull(1).Value = True
    ElseIf OptionStat(3).Value = True Then
      FlagStat = "MIN":  OptionNull(1).Value = True
    ElseIf OptionStat(4).Value = True Then
      FlagStat = "SUM":  OptionNull(0).Value = True
    ElseIf OptionStat(7).Value = True Then
      FlagStat = "DAYS":  OptionNull(0).Value = True
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


