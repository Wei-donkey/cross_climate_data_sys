VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Begin VB.Form FrmConditionQuery 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Multi-Element Conditional Query"
   ClientHeight    =   13530
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   28380
   Icon            =   "FrmConditionQuery.frx":0000
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   13530
   ScaleWidth      =   28380
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFEFDF&
      ForeColor       =   &H80000008&
      Height          =   1440
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   28095
      Begin VB.Frame Frame10 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Filter Output"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   24960
         TabIndex        =   2
         Top             =   120
         Width           =   1455
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
            Left            =   300
            TabIndex        =   4
            Text            =   "999999"
            Top             =   840
            Width           =   1000
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
            Left            =   300
            TabIndex        =   3
            Text            =   "-9999"
            Top             =   360
            Width           =   1000
         End
         Begin VB.Label Label14 
            BackColor       =   &H00DFEFDF&
            Caption         =   "≤"
            Height          =   195
            Left            =   75
            TabIndex        =   6
            Top             =   915
            Width           =   255
         End
         Begin VB.Label Label13 
            BackColor       =   &H00DFEFDF&
            Caption         =   "≥"
            Height          =   195
            Left            =   75
            TabIndex        =   5
            Top             =   435
            Width           =   255
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Threshold"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   18240
         TabIndex        =   25
         Top             =   120
         Width           =   1455
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
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   27
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
            Left            =   360
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   720
            Width           =   975
         End
         Begin VB.Label Label2 
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
            Left            =   120
            TabIndex        =   29
            Top             =   435
            Width           =   255
         End
         Begin VB.Label Label1 
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
            Left            =   120
            TabIndex        =   28
            Top             =   795
            Width           =   255
         End
      End
      Begin VB.Frame Frame6 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Statistic Selection"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   15240
         TabIndex        =   18
         Top             =   120
         Width           =   3015
         Begin VB.OptionButton OptionStat 
            BackColor       =   &H00DFEFDF&
            Caption         =   "DayCount"
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
            Left            =   1080
            TabIndex        =   23
            Top             =   720
            Width           =   1335
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
            Left            =   120
            TabIndex        =   22
            Top             =   720
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
            Left            =   120
            TabIndex        =   21
            Top             =   360
            Value           =   -1  'True
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
            Left            =   2040
            TabIndex        =   20
            Top             =   360
            Width           =   900
         End
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
            Left            =   1080
            TabIndex        =   19
            Top             =   360
            Width           =   975
         End
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   26400
         TabIndex        =   15
         Top             =   120
         Width           =   1575
         Begin VB.CommandButton CmdLimit 
            Caption         =   "AddElement"
            Height          =   855
            Left            =   120
            TabIndex        =   17
            Top             =   240
            Width           =   615
         End
         Begin VB.CommandButton CmdDel 
            Caption         =   "DelElement"
            Height          =   855
            Left            =   840
            TabIndex        =   16
            Top             =   240
            Width           =   615
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H00C0C0C0&
         Height          =   1215
         Left            =   27600
         TabIndex        =   14
         Top             =   120
         Width           =   375
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid2 
         Height          =   1125
         Left            =   19680
         TabIndex        =   12
         Top             =   210
         Width           =   5295
         _ExtentX        =   9340
         _ExtentY        =   1984
         _Version        =   393216
         BackColorFixed  =   16777215
         BackColorBkg    =   16777215
         Appearance      =   0
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Frame FrmMeteoHour 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Element List"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   1440
         TabIndex        =   10
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
            TabIndex        =   13
            Top             =   240
            Width           =   11895
         End
      End
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Start/End Date"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   13560
         TabIndex        =   7
         Top             =   120
         Width           =   1695
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
            Left            =   120
            TabIndex        =   9
            Text            =   "yyyy-mm-dd"
            Top             =   720
            Width           =   1455
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
            Left            =   120
            TabIndex        =   8
            Text            =   "yyyy-mm-dd"
            Top             =   360
            Width           =   1455
         End
      End
      Begin VB.Frame Frame5 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         Caption         =   "Data Table"
         ForeColor       =   &H80000008&
         Height          =   1215
         Left            =   120
         TabIndex        =   1
         Top             =   120
         Width           =   1335
         Begin VB.ComboBox ComboTable 
            Height          =   300
            ItemData        =   "FrmConditionQuery.frx":3482
            Left            =   120
            List            =   "FrmConditionQuery.frx":3484
            Style           =   2  'Dropdown List
            TabIndex        =   24
            Top             =   360
            Width           =   1095
         End
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid1 
      Height          =   11745
      Left            =   120
      TabIndex        =   11
      ToolTipText     =   "Alt(Ctrl)+Click: Sort Vertically"
      Top             =   1560
      Width           =   28095
      _ExtentX        =   49556
      _ExtentY        =   20717
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
Attribute VB_Name = "FrmConditionQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mLastRightClickGrid As Object '最近一次右键点击的表格控件，供 mnuGridCopy_Click 使用

Private Const SWP_NOSIZE = &H1
Private Const SWP_NOMOVE = &H2
Private Const HWND_TOPMOST = -1
Private Const HWND_NOTOPMOST = -2
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

Private DataGap$  'Data Type
Private BYYYY$, BMM$, BDD$, EYYYY$, EMM$, EDD$ '用户自定义的查询起止年、月、日
Private BDate$, EDate$ '用户自定义的当年查询起始日期
Private iyears%, imonths%, idays%

Public FlagStat$ '统计量标识
Dim FlagLimit As Boolean '条件查询标识
Dim TextMin!, TextMax! '条件查询的上下限
Dim SelField_Restat$ '选择要素的再统计名
Dim SelField_Data$ '选择要素的旬/月/季/年间隔字段
Private BYYMM$, EYYMM$ '用户自定义的查询起止年月（用于旬月查询）


Public Sub cmdConditionQuery()
  Dim BMMDD$, EMMDD$
  Dim DataTable$
  Dim DataTable1$, DataTable2$
  Dim strYear$
  Screen.MousePointer = 11
  
  If HFGrid2.Rows = 1 Then
    MsgBox "Please add an element", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  End If
  
   '************2023-10-10添加对输入日期的精准判断**********
  Dim tempStr$()
  
  tempStr = Split(ComboBDate.Text, "-")
  BYYYY = tempStr(0) '开始年份
  BMM = tempStr(1): BDD = tempStr(2)
  BDate = Format(BYYYY, "0000") & "-" & Format(BMM, "00") & "-" & Format(BDD, "00")
  
  tempStr = Split(ComboEDate.Text, "-")
  EYYYY = tempStr(0) '结束年份
  EMM = tempStr(1): EDD = tempStr(2)
  EDate = Format(EYYYY, "0000") & "-" & Format(EMM, "00") & "-" & Format(EDD, "00")

  BYYMM = Left(BDate, 7) '开始年月
  EYYMM = Left(EDate, 7) '结束年月
  
  BMMDD = Right(BDate, 5) '开始月日
  EMMDD = Right(EDate, 5) '结束月日
  
  iyears = Int(EYYYY) - Int(BYYYY) + 1
  imonths = DateDiff("m", BDate, EDate) + 1
  idays = DateDiff("d", BDate, EDate) + 1
  
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
      
  If idays <= 0 Then MsgBox "Invalid start/end date", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  
  '*************** 小Hourly Data查询 *******************
  If ComboTable.Text = "Hourly" Then
    If CInt(Format(EDate, "yyyy")) - CInt(Format(BDate, "yyyy")) = 1 Then  '跨年：年份数少一年
      MsgBox "Hourly data query cannot span years", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    End If
    strYear = Format(EDate, "yyyy")
      
    DataGap = "HOR"
    DataTable = StaType & "_CLI_MUL_" & DataGap & "_" & strYear
    DataTable1 = "SURF_CLI_MUL_" & DataGap & "_" & strYear
    DataTable2 = "AWST_CLI_MUL_" & DataGap & "_" & strYear
    
    Call cmdConditionQuery_HORDAY(DataTable, DataTable1, DataTable2)
    
  '*************** 日数据查询 *******************
  ElseIf ComboTable.Text = "Daily" Then
    If idays > 366 Then
      MsgBox "Daily data range cannot exceed 366 days", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    End If
    
    DataGap = "DAY"
    DataTable = StaType & "_CLI_MUL_" & DataGap
    DataTable1 = "SURF_CLI_MUL_" & DataGap
    DataTable2 = "AWST_CLI_MUL_" & DataGap
    
    Call cmdConditionQuery_HORDAY(DataTable, DataTable1, DataTable2)
    
  '*************** 2026-8-31：旬/月/季/年数据查询 *******************
  ElseIf ComboTable.Text = "Dekad" Then
    If iyears > 10 Then
      MsgBox "Dekad data range cannot exceed 10 years", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    End If
      
    DataGap = "TEN"
    DataTable = StaType & "_CLI_MUL_" & DataGap
    DataTable1 = "SURF_CLI_MUL_" & DataGap
    DataTable2 = "AWST_CLI_MUL_" & DataGap
    
    Call cmdConditionQuery_RESTAT(DataTable, DataTable1, DataTable2)
  
  ElseIf ComboTable.Text = "Monthly" Then
    If iyears > 30 Then
      MsgBox "Monthly data range cannot exceed 30 years", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    End If
    
    DataGap = "MON"
    DataTable = StaType & "_CLI_MUL_" & DataGap
    DataTable1 = "SURF_CLI_MUL_" & DataGap
    DataTable2 = "AWST_CLI_MUL_" & DataGap
  
    Call cmdConditionQuery_RESTAT(DataTable, DataTable1, DataTable2)
  
  ElseIf ComboTable.Text = "Seasonal" Then
    If iyears > 90 Then
      MsgBox "Seasonal data range cannot exceed 90 years", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    End If
    
    DataGap = "QTR"
    DataTable = StaType & "_CLI_MUL_" & DataGap
    DataTable1 = "SURF_CLI_MUL_" & DataGap
    DataTable2 = "AWST_CLI_MUL_" & DataGap
    
    Call cmdConditionQuery_RESTAT(DataTable, DataTable1, DataTable2)
  
  ElseIf ComboTable.Text = "Annual" Then
    
    DataGap = "YER"
    DataTable = StaType & "_CLI_MUL_" & DataGap
    DataTable1 = "SURF_CLI_MUL_" & DataGap
    DataTable2 = "AWST_CLI_MUL_" & DataGap
    
    Call cmdConditionQuery_RESTAT(DataTable, DataTable1, DataTable2)
  
  End If
  
End Sub


Private Sub cmdConditionQuery_RESTAT(DataTable$, DataTable1$, DataTable2$)

  Dim SelField$(), SelName$()
  Dim Min_HL$(), Max_HL$()
  Dim iElements% '条件查询的要素数量
  Dim i As Double, j%

  Screen.MousePointer = 11
  
  HFGrid1.Clear
  
  iElements = HFGrid2.Rows - 1
  ReDim SelField(1 To iElements): ReDim SelName(1 To iElements): ReDim Min_HL(1 To iElements): ReDim Max_HL(1 To iElements)
  For i = 1 To iElements
    SelField(i) = HFGrid2.TextMatrix(i, 2)
    SelName(i) = HFGrid2.TextMatrix(i, 1)
    Min_HL(i) = HFGrid2.TextMatrix(i, 3)
    Max_HL(i) = HFGrid2.TextMatrix(i, 4)
  Next i
  
  If StaType = "SURF" Or StaType = "AWST" Then
    strSQL = "select a.stacode 站号,b.slm 站名,b.v_city 地市,b.v_county 区县,b.v_town 镇街"
     
    If DataGap = "TEN" Then strSQL = strSQL & ",(TO_CHAR(TO_DATE(iyear || '-' || imonth || '-' || iten , 'yyyy-mm-dd'), 'yyyy-mm-dd')) 年月旬"
    If DataGap = "MON" Then strSQL = strSQL & ",(TO_CHAR(TO_DATE(iyear || '-' || imonth || '-01', 'yyyy-mm-dd'), 'yyyy-mm')) 年月"
    If DataGap = "QTR" Then strSQL = strSQL & ",(iyear || '-' || cquarter) 年季"
    If DataGap = "YER" Then strSQL = strSQL & ",iyear 年"
      
    For i = 1 To iElements  '20时日雨量这种名称是数字开头的，前后必须加引号
      strSQL = strSQL & "," & SelField(i) & " """
      strSQL = strSQL & SelName(i) & ""
      strSQL = strSQL & """"
    Next i
  
    strSQL = strSQL & " from " & DataTable & " a,T_OTHE_STATION_META_BASIC_TAB b"
    If DataGap = "TEN" Or DataGap = "MON" Then
      strSQL = strSQL & " where to_date(iyear || '-' || imonth,'yyyy-mm') between to_date('" & BYYMM & "','yyyy-mm') and to_date('" & EYYMM & "','yyyy-mm')"
    ElseIf DataGap = "QTR" Or DataGap = "YER" Then
      strSQL = strSQL & " where iyear between " & Str(BYYYY) & " and " & Str(EYYYY)
    End If
     
    For i = 1 To iElements
      If Not (Min_HL(i) = -9999 And Max_HL(i) = 999999) Then strSQL = strSQL & " and " & SelField(i) & " between " & Min_HL(i) & " and " & Max_HL(i)
    Next i
     
    strSQL = strSQL & " and a.stacode=b.v01301"
    If strSta <> "" Then strSQL = strSQL & " and a." & strSta & ""
  
    If DataGap = "TEN" Then strSQL = strSQL & " order by 站号,年月旬"
    If DataGap = "MON" Then strSQL = strSQL & " order by 站号,年月"
    If DataGap = "QTR" Then strSQL = strSQL & " order by 站号,年季"
    If DataGap = "YER" Then strSQL = strSQL & " order by 站号,年"
  
  ElseIf StaType = "BOTH" Then
    strSQL = "select a.stacode 站号,b.slm 站名,b.v_city 地市,b.v_county 区县,b.v_town 镇街"
     
    If DataGap = "TEN" Then strSQL = strSQL & ",(TO_CHAR(TO_DATE(iyear || '-' || imonth || '-' || iten , 'yyyy-mm-dd'), 'yyyy-mm-dd')) 年月旬"
    If DataGap = "MON" Then strSQL = strSQL & ",(TO_CHAR(TO_DATE(iyear || '-' || imonth || '-01', 'yyyy-mm-dd'), 'yyyy-mm')) 年月"
    If DataGap = "QTR" Then strSQL = strSQL & ",(iyear || '-' || cquarter) 年季"
    If DataGap = "YER" Then strSQL = strSQL & ",iyear 年"
      
    For i = 1 To iElements  '20时日雨量这种名称是数字开头的，前后必须加引号
      strSQL = strSQL & "," & SelField(i) & " """
      strSQL = strSQL & SelName(i) & ""
      strSQL = strSQL & """"
    Next i
     
    strSQL = strSQL & " from " & DataTable1 & " a,T_OTHE_STATION_META_BASIC_TAB b"
    If DataGap = "TEN" Or DataGap = "MON" Then
      strSQL = strSQL & " where to_date(iyear || '-' || imonth,'yyyy-mm') between to_date('" & BYYMM & "','yyyy-mm') and to_date('" & EYYMM & "','yyyy-mm')"
    ElseIf DataGap = "QTR" Or DataGap = "YER" Then
      strSQL = strSQL & " where iyear between " & Str(BYYYY) & " and " & Str(EYYYY)
    End If
     
    For i = 1 To iElements
      If Not (Min_HL(i) = -9999 And Max_HL(i) = 999999) Then strSQL = strSQL & " and " & SelField(i) & " between " & Min_HL(i) & " and " & Max_HL(i)
    Next i
     
    strSQL = strSQL & " and a.stacode=b.v01301"
    If strSURF <> "" Then strSQL = strSQL & " and a." & strSURF & ""
        
    strSQL = strSQL & " UNION "  '************* union 国家站和自动站结果 **********************
    
    strSQL = strSQL & "select a.stacode 站号,b.slm 站名,b.v_city 地市,b.v_county 区县,b.v_town 镇街"
     
    If DataGap = "TEN" Then strSQL = strSQL & ",(TO_CHAR(TO_DATE(iyear || '-' || imonth || '-' || iten , 'yyyy-mm-dd'), 'yyyy-mm-dd')) 年月旬"
    If DataGap = "MON" Then strSQL = strSQL & ",(TO_CHAR(TO_DATE(iyear || '-' || imonth || '-01', 'yyyy-mm-dd'), 'yyyy-mm')) 年月"
    If DataGap = "QTR" Then strSQL = strSQL & ",(iyear || '-' || cquarter) 年季"
    If DataGap = "YER" Then strSQL = strSQL & ",iyear 年"
      
    For i = 1 To iElements  '20时日雨量这种名称是数字开头的，前后必须加引号
      strSQL = strSQL & "," & SelField(i) & " """
      strSQL = strSQL & SelName(i) & ""
      strSQL = strSQL & """"
    Next i
     
    strSQL = strSQL & " from " & DataTable2 & " a,T_OTHE_STATION_META_BASIC_TAB b"
    If DataGap = "TEN" Or DataGap = "MON" Then
      strSQL = strSQL & " where to_date(iyear || '-' || imonth,'yyyy-mm') between to_date('" & BYYMM & "','yyyy-mm') and to_date('" & EYYMM & "','yyyy-mm')"
    ElseIf DataGap = "QTR" Or DataGap = "YER" Then
      strSQL = strSQL & " where iyear between " & Str(BYYYY) & " and " & Str(EYYYY)
    End If
     
    For i = 1 To iElements
      If Not (Min_HL(i) = -9999 And Max_HL(i) = 999999) Then strSQL = strSQL & " and " & SelField(i) & " between " & Min_HL(i) & " and " & Max_HL(i)
    Next i
     
    strSQL = strSQL & " and a.stacode=b.v01301"
    If strAWST <> "" Then strSQL = strSQL & " and a." & strAWST & ""
    
    If DataGap = "TEN" Then strSQL = strSQL & " order by 站号,年月旬"
    If DataGap = "MON" Then strSQL = strSQL & " order by 站号,年月"
    If DataGap = "QTR" Then strSQL = strSQL & " order by 站号,年季"
    If DataGap = "YER" Then strSQL = strSQL & " order by 站号,年"
    
  End If
  
  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart
  
  ORARst.CursorLocation = adUseClient
  ORAConn2.Open
  ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly

  If ORARst.EOF = True Then
    MsgBox "No results match the criteria", vbInformation, "Notice"
  ElseIf ORARst.EOF <> True Then
    Set HFGrid1.DataSource = ORARst
    For i = 1 To HFGrid1.Rows - 1
      HFGrid1.TextMatrix(i, 0) = i
    Next i
    For j = 1 To HFGrid1.Cols - 1
      HFGrid1.ColWidth(j) = 2000
    Next j
  End If
  ORARst.Close
  
  ORAConn2.Close
  
  Call TimeStop
  frmWait.Hide: DoEvents
  Screen.MousePointer = 1

End Sub



Private Sub cmdConditionQuery_HORDAY(DataTable$, DataTable1$, DataTable2$)
''''  Dim cross年份 As Boolean '查询时段是否跨年度
  Dim SelField$(), SelName$()
  Dim Min_HL$(), Max_HL$()
  Dim iElements% '条件查询的要素数量
  Dim i As Double, j%

  Screen.MousePointer = 11
  
  HFGrid1.Clear
  
  iElements = HFGrid2.Rows - 1
  ReDim SelField(1 To iElements): ReDim SelName(1 To iElements): ReDim Min_HL(1 To iElements): ReDim Max_HL(1 To iElements)
  For i = 1 To iElements
    SelField(i) = HFGrid2.TextMatrix(i, 2)
    SelName(i) = HFGrid2.TextMatrix(i, 1)
    Min_HL(i) = HFGrid2.TextMatrix(i, 3)
    Max_HL(i) = HFGrid2.TextMatrix(i, 4)
  Next i
  
  If StaType = "AWST" Or StaType = "SURF" Then
    strSQL = "select a.stacode 站号,b.slm 站名,b.v_city 地市,b.v_county 区县,b.v_town 镇街"
     
    If DataGap = "HOR" Then strSQL = strSQL & ",to_char(ddatetime,'yyyy-mm-dd hh24') 时间"
    If DataGap = "DAY" Then strSQL = strSQL & ",to_char(ddate,'yyyy-mm-dd') 日期"
      
    For i = 1 To iElements  '20时日雨量这种名称是数字开头的，前后必须加引号
      strSQL = strSQL & "," & SelField(i) & " """
      strSQL = strSQL & SelName(i) & ""
      strSQL = strSQL & """"
    Next i
     
    strSQL = strSQL & " from " & DataTable & " a,T_OTHE_STATION_META_BASIC_TAB b"
    If DataGap = "HOR" Then
      strSQL = strSQL & " where ddatetime between to_date('" & BDate & " 00','yyyy-mm-dd hh24') and to_date('" & EDate & " 23','yyyy-mm-dd hh24')"
    ElseIf DataGap = "DAY" Then
      strSQL = strSQL & " where ddate between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
    End If
     
    For i = 1 To iElements
      If Not (Min_HL(i) = -9999 And Max_HL(i) = 999999) Then strSQL = strSQL & " and " & SelField(i) & " between " & Min_HL(i) & " and " & Max_HL(i)
    Next i
     
    strSQL = strSQL & " and a.stacode=b.v01301"
    If strSta <> "" Then strSQL = strSQL & " and a." & strSta & ""
    
    If DataGap = "DAY" Then strSQL = strSQL & " order by 站号,日期"
    
  ElseIf StaType = "BOTH" Then
    strSQL = "select a.stacode 站号,b.slm 站名,b.v_city 地市,b.v_county 区县,b.v_town 镇街"
     
    If DataGap = "HOR" Then strSQL = strSQL & ",to_char(ddatetime,'yyyy-mm-dd hh24') 时间"
    If DataGap = "DAY" Then strSQL = strSQL & ",to_char(ddate,'yyyy-mm-dd') 日期"
      
    For i = 1 To iElements  '20时日雨量这种名称是数字开头的，前后必须加引号
      strSQL = strSQL & "," & SelField(i) & " """
      strSQL = strSQL & SelName(i) & ""
      strSQL = strSQL & """"
    Next i
     
    strSQL = strSQL & " from " & DataTable1 & " a,T_OTHE_STATION_META_BASIC_TAB b"
    If DataGap = "HOR" Then
      strSQL = strSQL & " where ddatetime between to_date('" & BDate & " 00','yyyy-mm-dd hh24') and to_date('" & EDate & " 23','yyyy-mm-dd hh24')"
    ElseIf DataGap = "DAY" Then
      strSQL = strSQL & " where ddate between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
    End If
     
    For i = 1 To iElements
      If Not (Min_HL(i) = -9999 And Max_HL(i) = 999999) Then strSQL = strSQL & " and " & SelField(i) & " between " & Min_HL(i) & " and " & Max_HL(i)
    Next i
     
    strSQL = strSQL & " and a.stacode=b.v01301"
    If strSURF <> "" Then strSQL = strSQL & " and a." & strSURF & ""
        
    strSQL = strSQL & " UNION "  '************* union 国家站和自动站结果 **********************
    
    strSQL = strSQL & "select a.stacode 站号,b.slm 站名,b.v_city 地市,b.v_county 区县,b.v_town 镇街"
     
    If DataGap = "HOR" Then strSQL = strSQL & ",to_char(ddatetime,'yyyy-mm-dd hh24') 时间"
    If DataGap = "DAY" Then strSQL = strSQL & ",to_char(ddate,'yyyy-mm-dd') 日期"
      
    For i = 1 To iElements  '20时日雨量这种名称是数字开头的，前后必须加引号
      strSQL = strSQL & "," & SelField(i) & " """
      strSQL = strSQL & SelName(i) & ""
      strSQL = strSQL & """"
    Next i
     
    strSQL = strSQL & " from " & DataTable2 & " a,T_OTHE_STATION_META_BASIC_TAB b"
    If DataGap = "HOR" Then
      strSQL = strSQL & " where ddatetime between to_date('" & BDate & " 00','yyyy-mm-dd hh24') and to_date('" & EDate & " 23','yyyy-mm-dd hh24')"
    ElseIf DataGap = "DAY" Then
      strSQL = strSQL & " where ddate between to_date('" & BDate & "','yyyy-mm-dd') and to_date('" & EDate & "','yyyy-mm-dd')"
    End If
     
    For i = 1 To iElements
      If Not (Min_HL(i) = -9999 And Max_HL(i) = 999999) Then strSQL = strSQL & " and " & SelField(i) & " between " & Min_HL(i) & " and " & Max_HL(i)
    Next i
     
    strSQL = strSQL & " and a.stacode=b.v01301"
    If strAWST <> "" Then strSQL = strSQL & " and a." & strAWST & ""
  
    If DataGap = "DAY" Then strSQL = strSQL & " order by 站号,日期"
  
  End If
  
  
  frmWait.Label1.Caption = "Querying, please wait"
  frmWait.Show: DoEvents
  Call TimeStart
  
  ORARst.CursorLocation = adUseClient
  If DataGap = "HOR" Then
    ORAConn3.Open
    ORARst.Open strSQL, ORAConn3, adOpenDynamic, adLockReadOnly
  ElseIf DataGap = "DAY" Then
    ORAConn2.Open
    ORARst.Open strSQL, ORAConn2, adOpenDynamic, adLockReadOnly
  End If

  If ORARst.EOF = True Then
    MsgBox "No results match the criteria", vbInformation, "Notice"
  ElseIf ORARst.EOF <> True Then
    Set HFGrid1.DataSource = ORARst
    For i = 1 To HFGrid1.Rows - 1
      HFGrid1.TextMatrix(i, 0) = i
    Next i
    For j = 1 To HFGrid1.Cols - 1
      HFGrid1.ColWidth(j) = 2000
    Next j
  End If
  ORARst.Close
  
  If DataGap = "HOR" Then
    ORAConn3.Close
  ElseIf DataGap = "DAY" Then
    ORAConn2.Close
  End If
  
  Call TimeStop
  frmWait.Hide: DoEvents
  Screen.MousePointer = 1

End Sub



Public Sub CmdSaveConditionQuery()
  Dim i As Long, j%  '2023-10-12 因为查询数据条数超过integer最大Value，溢出错误，因此改为long类型
  
  ' ************************ 判断是否是保存数据权限 ***********************************
  If userID = "liuw" Then '超级管理员-IP无限制
  
  Else '其他所是用户-IP限制
    
    If IP_SaveData = "" Then
      MsgBox "This account has no registered data-download IP -- cannot save data", vbInformation, "Notice"
      Exit Sub
    End If
    
    If IP_Ban = False Then 'IP地址匹配
      FrmMain.StatusBar1.Panels(1).Text = "This machine's IP is registered; saving data"
    ElseIf IP_Ban = True Then 'IP地址不匹配
      MsgBox "Registered account IP does not match this machine's IP -- cannot save data", vbInformation, "Notice"
      Exit Sub
    End If
    
  End If
  ' ************************ 判断是否是保存数据权限 ***********************************
  
  Dim strFileName$, strFileType$, strTemp$, strStacode$

  FrmMain.CommonDialogSave.Filter = "Text Files (*.csv)|*.csv": strFileType = "csv"
  FrmMain.CommonDialogSave.FilterIndex = 1
  
  FrmMain.CommonDialogSave.FileName = "Conditional Query Data"
  
  FrmMain.CommonDialogSave.InitDir = App.Path & "\Output"
  FrmMain.CommonDialogSave.Flags = &H2
  FrmMain.CommonDialogSave.ShowSave
  If FrmMain.CommonDialogSave.Flags = 2 Then Screen.MousePointer = 1: FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub
  
  frmWait.Label1.Caption = "Saving, please wait"
  frmWait.Show
  Screen.MousePointer = 11
  DoEvents
  FrmMain.StatusBar1.Panels(1).Text = "Saving"
  
  strFileName = FrmMain.CommonDialogSave.FileName
  If Len(Dir(strFileName)) <> 0 Then Kill strFileName '如果存在同名文件则删除同名文件

  Open strFileName For Output As #FileNum

  With HFGrid1
  For i = 0 To .Rows - 1
    strTemp = .TextMatrix(i, 0)
    For j = 1 To .Cols - 1
      strTemp = strTemp & "," & .TextMatrix(i, j)
    Next j
    Print #FileNum, strTemp
  Next i
  End With

  Close #FileNum
  FrmMain.StatusBar1.Panels(1).Text = "Save result: Done"


  Dim DataType$, Spatial$, Temporal$, DataLength$
  Dim Savetime As Date

  Savetime = Format(Now, "yyyy-mm-dd hh:mm:ss")

  
  If DataGap = "HOR" Or DataGap = "DAY" Then
    If idays < 180 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
    DataLength = CStr(idays) & "日"
    Temporal = BDate & "至" & EDate
    If DataGap = "HOR" Then DataGap = "Hourly Conditional Query"
    If DataGap = "DAY" Then DataGap = "Daily Conditional Query"
    
  ElseIf DataGap = "TEN" Then
    If imonths < 6 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
    DataLength = CStr(imonths) & "月"
    Temporal = BYYMM & "至" & EYYMM
    DataGap = "Dekad Conditional Query"
  ElseIf DataGap = "MON" Then
    If imonths < 6 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
    DataLength = CStr(imonths) & "月"
    Temporal = BYYMM & "至" & EYYMM
    DataGap = "Monthly Conditional Query"
  ElseIf DataGap = "QTR" Then
    If iyears < 3 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
    DataLength = CStr(iyears) & "年"
    Temporal = BYYYY & "至" & EYYYY
    DataGap = "Seasonal Conditional Query"
  ElseIf DataGap = "YER" Then
    If iyears < 6 Then frmWait.Hide: DoEvents: Screen.MousePointer = 1: Exit Sub
    DataLength = CStr(iyears) & "年"
    Temporal = BYYYY & "至" & EYYYY
    DataGap = "Annual Conditional Query"
  End If
  
  
  DataType = "Data"
  Spatial = Replace(DataExtent, "'", "")
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3

  strSQL = "insert into T_OTHE_CROSS_DATA_DOWNLOAD(USERID,DDATETIME,IP_SAVE,STATYPE,DATAGAP,DATATYPE,SPATIAL,TEMPORAL,DATALENGTH,VARIABLE)"
  strSQL = strSQL + " Values('" & userID & "',to_date('" & Savetime & "', 'yyyy-mm-dd hh24:mi:ss'),'" & UserIP & "','" & StaType & "','" & DataGap & "','" & DataType & "','" & Spatial & "','" & Temporal & "','" & DataLength & "','MULTI')"
  
  ORAConn4.Execute strSQL
  ORAConn4.Close
  
  frmWait.Hide: DoEvents
  Screen.MousePointer = 1


End Sub

Private Sub Form_Load()
  FrmMain.FormAdd

  Dim strYear$, i%, j%
  Call ReadColInfo("COLS_Hour.ini")
  COL_HOR = COL_tmp

  Call ReadColInfo("COLS_Initial.ini")
  COL_Initial = COL_tmp
  
  '2026-08-31 添加旬月季年条件查询
  Call ReadColInfo("COLS_ReStat.ini")
  COL_Restat = COL_tmp

  strYear = CStr(Year(Date) - 1)
  
  ComboBDate.Text = Format(strYear & "-01-01", "yyyy-mm-dd")
  ComboEDate.Text = Format(strYear & "-12-31", "yyyy-mm-dd")
  
  With HFGrid2
  .Cols = 5: .Rows = 1
  .TextMatrix(0, 1) = "Element Name": .TextMatrix(0, 2) = "Field Name": .TextMatrix(0, 3) = "Lower Limit": .TextMatrix(0, 4) = "Upper Limit"
  .ColWidth(0) = 300: .ColWidth(1) = 1400: .ColWidth(2) = 1000: .ColWidth(3) = 1000: .ColWidth(4) = 1000
  .Rows = 1
  End With
  
  
  ComboTable.AddItem "Hourly": ComboTable.AddItem "Daily": ComboTable.AddItem "Dekad"
  ComboTable.AddItem "Monthly": ComboTable.AddItem "Seasonal": ComboTable.AddItem "Annual"
  
  ComboTable.Text = "Daily":  Call ComboTable_click
  

  '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
  If GuestMode Then
    Call ImportData.LoadGuestCSV(FrmConditionQuery.HFGrid1, Nothing, "DataConditionQuery.csv", 3)
  
    For i = 1 To HFGrid1.Rows - 1
      HFGrid1.TextMatrix(i, 0) = i
    Next i
    For j = 1 To HFGrid1.Cols - 1
      HFGrid1.ColWidth(j) = 2000
    Next j

  End If

End Sub

'2026-08-19 v2.14 界面自适应：主窗体尺寸变化时，让本窗体的表格和选项区跟着调整。
'本窗体没是HFGrid2/'HFGrid3那种三联表格，HFGrid1直接贴到窗体底边距即可
Private Sub Form_Resize()
  Call ApplyLayout
End Sub


Public Sub ApplyLayout()
  On Error GoTo ResizeError
  If Me.WindowState = 1 Then Exit Sub '最小化时不处理

  Call ResizeFillGrid(HFGrid1, Me.ScaleHeight - MDI_RESIZE_BOTTOM_MARGIN, Me.ScaleWidth)
  Call ResizeFrame1Width(Frame1, Me.ScaleWidth)
  Call ResizeChildFrameWidth(Frame11, Frame1, 120)

  Exit Sub
ResizeError:
  Resume Next
End Sub


Private Sub CmdDel_Click()
  Dim i%
  If HFGrid2.Row = 0 Then MsgBox "Please select the row(s) to delete", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
  With HFGrid2
  
    If .Row < .Rows - 1 Then
      For i = 1 To .Rows - 1 - .Row
        .TextMatrix(.Row + i - 1, 1) = .TextMatrix(.Row + i, 1)
        .TextMatrix(.Row + i - 1, 2) = .TextMatrix(.Row + i, 2)
        .TextMatrix(.Row + i - 1, 3) = .TextMatrix(.Row + i, 3)
        .TextMatrix(.Row + i - 1, 4) = .TextMatrix(.Row + i, 4)
      Next i
    End If
  
  .Rows = .Rows - 1
  
  End With
End Sub

Private Sub CmdLimit_Click()
  Dim i%, j%
  Dim SelName$
  
  If ComboTable.Text <> "Hourly" And ComboTable.Text <> "Daily" Then
    TextMin = ComboMin.Text: TextMax = ComboMax.Text
    If TextMin = -9999 And TextMax = 999999 Then
      If Not ((List_Col.Selected(35) = True Or List_Col.Selected(4) = True Or List_Col.Selected(5) = True Or List_Col.Selected(6) = True Or List_Col.Selected(7) = True) And OptionStat(4).Value = True) Then   '2024-10-28：日照/降水可以查非条件的累计值
        If OptionStat(4).Value = True Or OptionStat(7).Value = True Then
            MsgBox "Please select a threshold value", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
        End If
      End If
    End If
  End If
  
  With HFGrid2
  .Rows = .Rows + 1
  .TextMatrix(.Rows - 1, 0) = .Rows - 1
  
  For i = 0 To List_Col.ListCount - 1
    If List_Col.Selected(i) = True Then
      SelName = List_Col.List(i)
      .TextMatrix(.Rows - 1, 1) = SelName
      .TextMatrix(.Rows - 1, 3) = ComboMin_HL.Text
      .TextMatrix(.Rows - 1, 4) = ComboMax_HL.Text
        
      If ComboTable.Text = "Hourly" Then
        For j = 1 To UBound(COL_HOR)
          If COL_HOR(j).Col_name = SelName Then .TextMatrix(.Rows - 1, 2) = COL_HOR(j).Col_ID: Exit For
        Next j
      
      ElseIf ComboTable.Text = "Daily" Then
        For j = 1 To UBound(COL_Initial)
          If COL_Initial(j).Col_name = SelName Then .TextMatrix(.Rows - 1, 2) = COL_Initial(j).Col_ID: Exit For
        Next j
      
      Else ' 其他再统计数据，需要判断是否为条件查询
      
        If TextMin = -9999 And TextMax = 999999 Then
          FlagLimit = False
        Else
          FlagLimit = True
        End If
      
        If FlagLimit = False Then
          SelField_Data = SelField_Restat & "_" & FlagStat
        ElseIf FlagLimit = True Then
          If TextMax <> 999999 Then SelField_Data = SelField_Restat & "_" & TextMax & FlagStat
          If TextMin <> -9999 Then SelField_Data = SelField_Restat & "_" & TextMin & FlagStat
          If TextMin = 0.1 Then SelField_Data = SelField_Restat & "_" & FlagStat '雨量/雨日
        End If
        
        .TextMatrix(.Rows - 1, 2) = SelField_Data
        
'        For j = 1 To UBound(COL_Restat)
'          If COL_Restat(j).Col_name = SelName Then .TextMatrix(.Rows - 1, 2) = COL_Restat(j).Col_ID: Exit For
'        Next j
      
      End If
        
      Exit For
    End If
  Next i
  
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)

  FrmMain.FormDel
  FrmMain.StatusBar1.Panels(1).Text = ""
  Unload Me
End Sub

Private Sub List_Col_Click()
  Dim i%
    
  If ComboTable.Text <> "Hourly" And ComboTable.Text <> "Daily" Then
    For i = 0 To List_Col.ListCount - 1
      If List_Col.Selected(i) = True Then
        SelField_Restat = COL_Restat(i + 1).Col_ID
        Exit For
      End If
    Next i
  
    Call Stat_ConditionQuery(FrmConditionQuery, SelField_Restat, FlagStat)
    Call Option_ConditionQuery(FrmConditionQuery, SelField_Restat)
  End If


End Sub


Private Sub ComboTable_click()

  Call FieldList
  HFGrid2.Clear
  With HFGrid2
  .Cols = 5: .Rows = 1
  .TextMatrix(0, 1) = "Element Name": .TextMatrix(0, 2) = "Field Name": .TextMatrix(0, 3) = "Lower Limit": .TextMatrix(0, 4) = "Upper Limit"
  End With

  If ComboTable.Text = "Hourly" Or ComboTable.Text = "Daily" Then
    OptionStat(1).Enabled = False: OptionStat(2).Enabled = False: OptionStat(3).Enabled = False
    OptionStat(4).Enabled = False: OptionStat(7).Enabled = False
    ComboMin.Enabled = False: ComboMax.Enabled = False
  Else
    OptionStat(1).Enabled = True: OptionStat(2).Enabled = True: OptionStat(3).Enabled = True
    OptionStat(4).Enabled = True: OptionStat(7).Enabled = True
    ComboMin.Enabled = True: ComboMax.Enabled = True
  End If

End Sub


Public Sub FieldList()
  Dim i%
  List_Col.Clear
  
   '****************************2022-06-14从ini配置文件中读取字段名*****************************
  If ComboTable.Text = ComboTable.List(0) Then
    For i = 1 To UBound(COL_HOR)
      List_Col.AddItem COL_HOR(i).Col_name
    Next i
          
  ElseIf ComboTable.Text = ComboTable.List(1) Then
    For i = 1 To UBound(COL_Initial)
      List_Col.AddItem COL_Initial(i).Col_name
    Next i
  
  Else
    For i = 1 To UBound(COL_Restat)
      List_Col.AddItem COL_Restat(i).Col_name
    Next i
    
  End If
  List_Col.Selected(0) = True

End Sub


Private Sub mnuGridCopy_Click()
  Call CopyGridSelectionToClipboard(mLastRightClickGrid)
End Sub


'********************对HFGrid1中的数据进行排序***************************
Private Sub HFGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid1
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols(FrmConditionQuery.HFGrid1, shift)
End Sub


Private Sub HFGrid2_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Set mLastRightClickGrid = HFGrid2
    PopupMenu mnuGridPopup
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


Private Sub OptionStat_Click(Index As Integer)

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
  
  Call Option_ConditionQuery(FrmConditionQuery, SelField_Restat)
End Sub

