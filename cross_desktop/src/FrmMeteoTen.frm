VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmMeteoTen 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Dekad Data Query"
   ClientHeight    =   13530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   28380
   Icon            =   "FrmMeteoTen.frx":0000
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
      TabIndex        =   14
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
      TabIndex        =   13
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
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Highlight Results"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   24000
         TabIndex        =   34
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
            TabIndex        =   37
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
            TabIndex        =   36
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
            TabIndex        =   35
            Top             =   720
            Width           =   855
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
            TabIndex        =   39
            Top             =   435
            Width           =   255
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
            TabIndex        =   38
            Top             =   795
            Width           =   255
         End
      End
      Begin VB.Frame Frame7 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Data State"
         ForeColor       =   &H00000000&
         Height          =   1215
         Left            =   14160
         TabIndex        =   29
         Top             =   120
         Width           =   3375
         Begin VB.CheckBox CheckXtrmYear 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Year"
            Height          =   255
            Left            =   2520
            TabIndex        =   33
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
            TabIndex        =   32
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
            TabIndex        =   31
            Top             =   840
            Value           =   -1  'True
            Width           =   1095
         End
         Begin VB.ComboBox ComboDataType 
            Height          =   300
            Left            =   240
            Style           =   2  'Dropdown List
            TabIndex        =   30
            Top             =   360
            Width           =   2055
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Threshold"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   20760
         TabIndex        =   24
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
            Style           =   2  'Dropdown List
            TabIndex        =   26
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
            Style           =   2  'Dropdown List
            TabIndex        =   25
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
            TabIndex        =   28
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
            TabIndex        =   27
            Top             =   795
            Width           =   255
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Statistic Selection"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   17520
         TabIndex        =   18
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
            TabIndex        =   23
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
            TabIndex        =   22
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
            TabIndex        =   21
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
            TabIndex        =   20
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
            TabIndex        =   19
            Top             =   720
            Width           =   1575
         End
      End
      Begin VB.Frame Frame9 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Condition not met"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   22320
         TabIndex        =   15
         Top             =   120
         Width           =   1695
         Begin VB.OptionButton OptionNull 
            BackColor       =   &H00DFEFDF&
            Caption         =   "is zero"
            Height          =   255
            Index           =   0
            Left            =   240
            TabIndex        =   17
            Top             =   360
            Value           =   -1  'True
            Width           =   1215
         End
         Begin VB.OptionButton OptionNull 
            BackColor       =   &H00DFEFDF&
            Caption         =   "is empty"
            Height          =   255
            Index           =   1
            Left            =   240
            TabIndex        =   16
            Top             =   720
            Width           =   1215
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Start/End Year-Month-Dekad"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   12240
         TabIndex        =   8
         Top             =   120
         Width           =   1935
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
            TabIndex        =   12
            Text            =   "yyyy"
            Top             =   360
            Width           =   840
         End
         Begin VB.ComboBox ComboEMM 
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
            Left            =   1080
            TabIndex        =   11
            Text            =   "mm"
            Top             =   720
            Width           =   600
         End
         Begin VB.ComboBox ComboBMM 
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
            Left            =   1080
            TabIndex        =   10
            Text            =   "mm"
            Top             =   360
            Width           =   600
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
         Left            =   26160
         TabIndex        =   2
         Top             =   120
         Width           =   1815
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
Attribute VB_Name = "FrmMeteoTen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public mLastRightClickGrid As Object '最近一次被点击（任意按钮）的表格；供mnuGridCopy_Click及ChartData查找图表数据源时使用
Private mEditGrid As Object '判断当前正在被TextEdit编辑的是哪个Grid（HFGrid1/HFGrid2/HFGrid3(Index)之一）

Dim SelField_Restat$ '选择要素的再统计名

Dim SelField_Ten$ '选择要素的月间隔字段
Dim SelField_Xtrm$ '选择要素的极值表字段
Dim SelField_Norm$ '选择要素的平均态字段

Public FlagStat$ '统计量标识
Dim FlagLimit As Boolean '条件查询标识

Public BYYMM$, EYYMM$, iMons%, iTens%
Public BMM$, EMM$ '用户自定义的查询起止月

Dim i%, j%, k%, l%
Dim TextMin!, TextMax! '条件查询的上下限

Public DataType$ '查询的数据类型：日间隔数据、平均态、极端态
Public XtrmYear As Boolean '是否显示极值出现年份
Dim XtrmYear_STA() As Integer '各站点的极值出现年份
Dim crossYear As Boolean '查询时段是否跨年度

' 2022-04-17 添加多要素默认统计量的统计
Public MultiSel As Boolean   '是否多要素统计
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

Private Sub CheckXtrmYear_Click()
  If CheckXtrmYear.Value = Checked Then
    XtrmYear = True
  ElseIf CheckXtrmYear.Value = Unchecked Then
    XtrmYear = False
  End If
  
  Call TableIni_MeteoTenQtr(FrmMeteoTen, BYYMM, iMons, iTens) ', SelField_Restat)
End Sub

Private Sub Form_Load()
  Dim i%
  Screen.MousePointer = 11
  FrmMain.FormAdd
  
  MultiSel = False '默认为单要素统计
  
  For i = 1951 To Year(Date)
    ComboBYear.AddItem i
    ComboEYear.AddItem i
  Next i
  For i = 1 To 12
    ComboBMM.AddItem i
    ComboEMM.AddItem i
  Next i
  ComboBYear.Text = Year(DateAdd("m", -3, Date))
  ComboEYear.Text = Year(DateAdd("m", -1, Date))
  ComboBMM.Text = Month(DateAdd("m", -3, Date))
  ComboEMM.Text = Month(DateAdd("m", -1, Date))
    
  Call ITimesInitial(FrmMeteoTen)
 
  Call ReadColInfo("COLS_ReStat.ini")
  COL_Restat = COL_tmp
 
  If OptionStat(5).Value = True Then FlagStat = "MAX"
  If OptionStat(6).Value = True Then FlagStat = "MIN"
  If CheckXtrmYear.Value = Checked Then XtrmYear = True
 
    
  OptionNull(1).Value = True
  
  ComboDataType.AddItem "Dekad Obs Value": ComboDataType.AddItem "Normals": ComboDataType.AddItem "Historical Extremes"
  ComboDataType.Text = ComboDataType.List(0)
  
  
  '****************************2022-03-08从ini配置文件中读取字段名*****************************
  List_Col.Clear
  For i = 1 To UBound(COL_Restat)
    List_Col.AddItem COL_Restat(i).Col_name
  Next i
  List_Col.Selected(0) = True
  SelField_Restat = COL_Restat(1).Col_ID
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


'*********************************************单要素逐旬资料统计>>>开始*****************************************
Public Sub CmdMeteoTen()
    
  Screen.MousePointer = 11
  
  Call ITimesInitial(FrmMeteoTen)
  If FlagInputErr = True Then Screen.MousePointer = 1: Exit Sub
  
  If iMons <= 0 Then MsgBox "Invalid start/end year-month", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  If iMons > 240 Then MsgBox "Query range cannot exceed 240 months", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  TextMin = ComboMin.Text: TextMax = ComboMax.Text
  If TextMin = -9999 And TextMax = 999999 Then
    If Not ((List_Col.Selected(35) = True Or List_Col.Selected(4) = True Or List_Col.Selected(5) = True Or List_Col.Selected(6) = True Or List_Col.Selected(7) = True) And OptionStat(4).Value = True) Then   '2024-10-28：日照/降水可以查非条件的累计值
      If OptionStat(4).Value = True Or OptionStat(7).Value = True Then
          MsgBox "Please select a threshold value", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
      End If
    End If
  End If
    
  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart
  Call TableIni_MeteoTenQtr(FrmMeteoTen, BYYMM, iMons, iTens) ', SelField_Restat)
  
 
  ORAConn2.Open
  ORAConn2.CursorLocation = 3
 

  '****************************2022-03-08*****************************
  For i = 0 To List_Col.ListCount - 1
    If List_Col.Selected(i) = True Then
      SelField_Restat = COL_Restat(i + 1).Col_ID
      Exit For
    End If
  Next i
  
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
  
  If DataType = "Data" Then
    If ComboBYear.Text = ComboEYear.Text Then '同年度
      crossYear = False
    ElseIf ComboBYear.Text < ComboEYear.Text Then '跨年度
      crossYear = True
    End If
    
    If FlagLimit = False Then
      SelField_Ten = SelField_Restat & "_" & FlagStat
    ElseIf FlagLimit = True Then
      If TextMax <> 999999 Then SelField_Ten = SelField_Restat & "_" & TextMax & FlagStat
      If TextMin <> -9999 Then SelField_Ten = SelField_Restat & "_" & TextMin & FlagStat
      If TextMin = 0.1 Then SelField_Ten = SelField_Restat & "_" & FlagStat '雨量/雨日
    End If
  
  ElseIf DataType = "Norm" Then
    If ComboBMM.Text <= ComboEMM.Text Then '同年度
      crossYear = False
    ElseIf ComboBMM.Text > ComboEMM.Text Then '跨年度
      crossYear = True
    End If
    
    If FlagLimit = False Then
      If FlagStat = "SUM" Or FlagStat = "DAYS" Then SelField_Ten = SelField_Restat & "_" & FlagStat
      If FlagStat <> "SUM" And FlagStat <> "DAYS" Then SelField_Ten = SelField_Restat
    ElseIf FlagLimit = True Then
      If TextMax <> 999999 Then SelField_Ten = SelField_Restat & "_" & TextMax & FlagStat
      If TextMin <> -9999 Then SelField_Ten = SelField_Restat & "_" & TextMin & FlagStat
      If TextMin = 0.1 Then SelField_Ten = SelField_Restat & "_" & FlagStat '雨量/雨日
    End If

  ElseIf DataType = "Xtrm" Then
    If ComboBMM.Text <= ComboEMM.Text Then '同年度
      crossYear = False
    ElseIf ComboBMM.Text > ComboEMM.Text Then '跨年度
      crossYear = True
    End If
    If FlagLimit = False Then
      If FlagStat = "SUM" Then
        SelField_Ten = SelField_Restat & "_" & FlagStat
      Else
        SelField_Ten = SelField_Restat
      End If
      
    ElseIf FlagLimit = True Then
      If TextMax <> 999999 Then SelField_Ten = SelField_Restat & "_" & TextMax & FlagStat
      If TextMin <> -9999 Then SelField_Ten = SelField_Restat & "_" & TextMin & FlagStat
      If TextMin = 0.1 Then SelField_Ten = SelField_Restat & "_" & FlagStat '雨量/雨日
    End If
    If OptionStat(5).Value = True Then
      SelField_Ten = SelField_Ten & "_MAX"
    ElseIf OptionStat(6).Value = True Then
      SelField_Ten = SelField_Ten & "_MIN"
    End If
  End If
  
'************************************* 1、查询所是站点多日值  ***********************************
  ReDim TEN_STA(1 To StaNum, 1 To iTens) '存放各站点、各日值
  If XtrmYear = True Then ReDim XtrmYear_STA(1 To StaNum, 1 To iTens)
  For i = 1 To StaNum
    For j = 1 To iTens
      TEN_STA(i, j) = -9999
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
    Call QueryMulTenData(SelField_Ten, BYYMM, EYYMM, TEN_STA())
  ElseIf DataType = "Norm" Or DataType = "Xtrm" Then
    Call QueryMulTenRestat(SelField_Ten, BMM, EMM, DataType, crossYear, TEN_STA(), XtrmYear_STA(), XtrmYear)
'  ElseIf DataType = "Xtrm" Then
'    Call QueryMulTenRestat(SelField_Ten, BMM, EMM, DataType, CrossYear, TEN_STA(), XtrmYear_STA(), XtrmYear)
  End If
  
  
 '************************************* 2、对条件查询结果进行设置（0或空）  ***********************************
  For i = 1 To StaNum
    For j = 1 To iTens
      If (FlagStat = "SUM" Or FlagStat = "DAYS") And FlagLimit = True Then '条件日数或条件累计值查询，无符合条件的都为0
        If TEN_STA(i, j) = -9999 Then
          If FlagResult = "0" Then TEN_STA(i, j) = 0  'j年份已经是资料了，若用户选择条件结果设置为0，则条件结果默认为0
        ElseIf TEN_STA(i, j) = 0 Then
          If FlagResult = "null" Then TEN_STA(i, j) = -9999
        End If
      End If
    Next j
  Next i
  

'************************************* 情形一：查询结果为站点尺度  ***********************************
  If FlagZone = "STA" Then '站点尺度
  
'************************************* 2、计算所是行和所是列的统计值  ***********************************
    '统计所是行的平均/最大/最小Value
    Call Stat_ROWS(TEN_STA, iTens, StaNum)
    '统计所是列的平均/最大/最小Value
    Call Stat_COLS(TEN_STA, iTens, StaNum, "AVE")
    
    '统计列的3列统计值的平均/最大/最小Value
    Call Stat_COLS_ROWS(StaNum)
    
'************************************* 3、填充上方表格  ***********************************
    Call HFGrid1_Fill(FrmMeteoTen, TEN_STA, iTens, XtrmYear_STA)
'************************************* 4、填充下方表格 ***********************************
    Call HFGrid2_Fill(FrmMeteoTen, ROWS_Stat, iTens)
  
'************************************* 情形二：查询结果为区域尺度  ***********************************
  Else
  
'************************************* 2、分区域统计所是行列值  ***********************************
    Call Stat_ZONE(TEN_STA, iTens)
    
'************************************* 3、填充上/中/下方表格值  ***********************************
    Call HFGrid3_fill(FrmMeteoTen, Zone_AVG, iTens, 1)
    Call HFGrid3_fill(FrmMeteoTen, Zone_MAX, iTens, 2)
    Call HFGrid3_fill(FrmMeteoTen, Zone_MIN, iTens, 3)
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
  Call Sort_Cols_Rows(FrmMeteoTen.HFGrid1, shift)
End Sub

Private Sub HFGrid2_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid2
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Rows(FrmMeteoTen.HFGrid2, shift)
End Sub

'********************对HFGrid3中的数据进行排序***************************
Private Sub HFGrid3_MouseUp(Index As Integer, Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid3(Index)
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols_Rows(FrmMeteoTen.HFGrid3(Index), shift)
End Sub


'********************利用HFGrid中的数据绘制地图***************************
Public Sub CmdMap_MeteoTen()
  If SelField_Restat = "R" Or SelField_Restat = "R08" Or SelField_Restat = "R20_08" Or SelField_Restat = "R08_20" Then
    Call Map_Value(FrmMeteoTen, "RYB")
  Else
    Call Map_Value(FrmMeteoTen, "BYR")
  End If
End Sub


'********************保存窗体内可见的HFGrid***************************
Public Sub CmdOutput_MeteoTen()
  Call Output_HFGrid(FrmMeteoTen, SelField_Ten)
End Sub

Private Sub List_Col_Click()
  For i = 0 To List_Col.ListCount - 1
    If List_Col.Selected(i) = True Then
      SelField_Restat = COL_Restat(i + 1).Col_ID
      Exit For
    End If
  Next i
  
  If DataType = "Xtrm" Then
  
  End If
  
  Call Stat_Meteo(FrmMeteoTen, SelField_Restat, FlagStat)
  Call Option_Meteo(FrmMeteoTen, SelField_Restat)
  Call TableIni_MeteoTenQtr(FrmMeteoTen, BYYMM, iMons, iTens) ', SelField_Restat)
  CheckHLight.Value = Unchecked
  
  '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
  If GuestMode And List_Col.Selected(0) = True Then
    Call ImportData.LoadGuestCSV(FrmMeteoTen.HFGrid1, FrmMeteoTen.HFGrid2, "DataMeteoTen.csv", 3)
  End If

End Sub


Private Sub OptionStat_Click(Index As Integer)
  If Index = 5 Or Index = 6 Then
    Exit Sub
  End If

  If OptionStat(2).Value = True Then
    FlagStat = "MAX"
  ElseIf OptionStat(3).Value = True Then
    FlagStat = "MIN"
  ElseIf OptionStat(4).Value = True Then
    FlagStat = "SUM"
  ElseIf OptionStat(1).Value = True Then
    FlagStat = "AVE"
  ElseIf OptionStat(7).Value = True Then
    FlagStat = "DAYS"
  End If
    
'  If OptionStat(5).Enabled = False And OptionStat(6).Enabled = False Then
    Call Option_Meteo(FrmMeteoTen, SelField_Restat)
'  End If
End Sub


Private Sub ComboDataType_Click()
  If ComboDataType.Text = ComboDataType.List(0) Then
    DataType = "Data"
    OptionNull(0).Value = True
  ElseIf ComboDataType.Text = ComboDataType.List(1) Then
    DataType = "Norm"
    If Norm_Src = "CMA" Then OptionNull(1).Value = True
  ElseIf ComboDataType.Text = ComboDataType.List(2) Then
    DataType = "Xtrm"
    OptionNull(0).Value = True
  End If
  
  CheckXtrmYear.Value = Unchecked
  
  Call Option_Meteo(FrmMeteoTen, SelField_Restat)
  Call ITimesInitial(FrmMeteoTen)
  Call TableIni_MeteoTenQtr(FrmMeteoTen, BYYMM, iMons, iTens) ', SelField_Restat)
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


