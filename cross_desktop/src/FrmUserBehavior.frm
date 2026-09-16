VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmUserBehavior 
   Caption         =   "User Activity Analysis"
   ClientHeight    =   12285
   ClientLeft      =   150
   ClientTop       =   405
   ClientWidth     =   28230
   Icon            =   "FrmUserBehavior.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12285
   ScaleWidth      =   28230
   Begin VB.CommandButton cmdSavePNG 
      Caption         =   "SaveTime Series Chart"
      Height          =   350
      Index           =   1
      Left            =   21600
      TabIndex        =   18
      Top             =   240
      Width           =   2055
   End
   Begin VB.CommandButton CmdSaveData 
      Caption         =   "Save Time Series Data"
      Height          =   350
      Index           =   2
      Left            =   25800
      TabIndex        =   17
      Top             =   240
      Width           =   2280
   End
   Begin VB.PictureBox Picture1 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   5130
      Left            =   17160
      ScaleHeight     =   5070
      ScaleWidth      =   10860
      TabIndex        =   16
      Top             =   7080
      Width           =   10920
   End
   Begin VB.CommandButton cmdSavePNG 
      Caption         =   "Save Pie Chart"
      Height          =   350
      Index           =   0
      Left            =   20040
      TabIndex        =   15
      Top             =   240
      Width           =   1455
   End
   Begin VB.CommandButton CmdSaveData 
      Caption         =   "Save Pie Chart Data"
      Height          =   350
      Index           =   1
      Left            =   23760
      TabIndex        =   14
      Top             =   240
      Width           =   1920
   End
   Begin VB.PictureBox PicChart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   6210
      Left            =   17160
      ScaleHeight     =   6150
      ScaleWidth      =   10860
      TabIndex        =   4
      Top             =   720
      Width           =   10920
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid1 
      Height          =   11490
      Left            =   120
      TabIndex        =   12
      Top             =   705
      Width           =   16935
      _ExtentX        =   29871
      _ExtentY        =   20267
      _Version        =   393216
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      GridColor       =   12632256
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdRefresh 
      Caption         =   "Refresh"
      Height          =   350
      Left            =   17280
      TabIndex        =   11
      Top             =   240
      Width           =   960
   End
   Begin VB.CommandButton CmdSaveData 
      Caption         =   "Save Table Data"
      Height          =   350
      Index           =   0
      Left            =   18360
      TabIndex        =   13
      Top             =   240
      Width           =   1560
   End
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   0
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.ComboBox ComboIPSave 
      Height          =   300
      Left            =   14880
      Style           =   2  'Dropdown List
      TabIndex        =   9
      Top             =   245
      Width           =   2175
   End
   Begin VB.ComboBox ComboUserID 
      Height          =   300
      Left            =   11400
      Style           =   2  'Dropdown List
      TabIndex        =   7
      Top             =   245
      Width           =   1620
   End
   Begin VB.ComboBox ComboDept 
      Height          =   300
      Left            =   7560
      Style           =   2  'Dropdown List
      TabIndex        =   5
      Top             =   245
      Width           =   2580
   End
   Begin VB.ComboBox ComboEDate 
      Height          =   300
      Left            =   3600
      TabIndex        =   2
      Top             =   245
      Width           =   1300
   End
   Begin VB.ComboBox ComboBDate 
      Height          =   300
      Left            =   1200
      TabIndex        =   0
      Top             =   245
      Width           =   1300
   End
   Begin VB.Label LblIPSave 
      Caption         =   "Data-Download IP"
      Height          =   255
      Left            =   13200
      TabIndex        =   10
      Top             =   270
      Width           =   1620
   End
   Begin VB.Label LblUserID 
      Caption         =   "Username"
      Height          =   255
      Left            =   10320
      TabIndex        =   8
      Top             =   270
      Width           =   1020
   End
   Begin VB.Label LblDept 
      Caption         =   "Depart."
      Height          =   255
      Left            =   6840
      TabIndex        =   6
      Top             =   270
      Width           =   525
   End
   Begin VB.Label LblEDate 
      Caption         =   "End Date"
      Height          =   255
      Left            =   2760
      TabIndex        =   1
      Top             =   270
      Width           =   840
   End
   Begin VB.Label LblBDate 
      Caption         =   "Start Date"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   270
      Width           =   975
   End
   Begin VB.Menu mnuGridPopup 
      Caption         =   "GridPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyTable 
         Caption         =   "Copy Table"
      End
   End
   Begin VB.Menu mnuChartPopup 
      Caption         =   "ChartPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy Chart"
      End
   End
End
Attribute VB_Name = "FrmUserBehavior"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Private mLastRightClickGrid As Object
Private mLastRightClickChart As Object

Private mChartLabels() As String
Private mChartCounts() As Long
Private mChartN As Long

'2026-08-17 每日活跃IP柱状图（画在Picture1）：mIPChartLabels存日期，mIPChartCounts存当日活跃IP数
Private mIPChartLabels() As String
Private mIPChartCounts() As Long
Private mIPChartN As Long
Private mIPTotalActive As Long  '查询时段内（跨天去重）总活跃IP数
Private mIPTotalQueries As Long '查询时段内总查询次数

Private Sub Form_Load()
  Dim i%
  
   '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1

  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = 11 * (Screen.Height - Me.Height) / 20

  Call PopulateDateCombos
  '2026-07-22 访客模式下不查询真实的用户行为数据
  If Not GuestMode Then
    Call PopulateDeptCombo
    Call PopulateUserCombo
    Call PopulateIPCombo
    Call RefreshAll
  Else
    CmdRefresh.Enabled = False
  End If

  '2026-07-29 访客模式下直接从内置数据加载表格和饼图，不查询数据库
  If GuestMode Then
    Call ImportData.LoadGuestCSV(FrmUserBehavior.HFGrid1, Nothing, "UserBehaviorGrid.csv", 0)
    
    With HFGrid1
      .ColWidth(0) = 400
      .ColWidth(1) = 1800
      .ColWidth(2) = 900
      .ColWidth(3) = 1400
      .ColWidth(4) = 2000
      .ColWidth(5) = 1400
      .ColWidth(6) = 1600
      .ColWidth(7) = 3000
      .ColWidth(8) = 2400
      .ColWidth(9) = 900
      .ColWidth(10) = 800
      
      For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignLeftCenter
        .FixedAlignment(i) = flexAlignLeftCenter
      Next i
    End With
    
    Call LoadGuestChartCSV("UserBehaviorChart.csv")
    Call LoadGuestIPActivityCSV("UserActiveDaily.csv")
    Exit Sub
  End If
  
End Sub


Private Sub PopulateDateCombos()
  Dim i As Long
  ComboBDate.Clear
  ComboEDate.Clear
  For i = 0 To 60
    ComboBDate.AddItem Format(Date - 60 + i, "yyyy-mm-dd")
    ComboEDate.AddItem Format(Date - 60 + i, "yyyy-mm-dd")
  Next i
  ComboBDate.Text = Format(Date - 9, "yyyy-mm-dd")
  ComboEDate.Text = Format(Date, "yyyy-mm-dd")
End Sub


Private Function GetComboFilterValue(ByRef Combo As ComboBox) As String
  If Combo.ListIndex <= 0 Then
    GetComboFilterValue = ""
  Else
    GetComboFilterValue = Combo.Text
  End If
End Function


Private Sub PopulateDeptCombo()
  ComboDept.Clear
  ComboDept.AddItem "All"

  strSQL = "select distinct p.department from T_OTHE_CROSS_DATA_DOWNLOAD d"
  strSQL = strSQL & " join T_OTHE_CROSS_USERS_PROPERTIES p on d.userid = p.userid"
  strSQL = strSQL & " where p.department is not null order by p.department"

  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  Do Until ORARst.EOF
    ComboDept.AddItem CStr(ORARst!Department)
    ORARst.MoveNext
  Loop
  ORARst.Close
  ORAConn4.Close

  ComboDept.ListIndex = 0
End Sub


Private Sub PopulateUserCombo()
  ComboUserID.Clear
  ComboUserID.AddItem "All"

  Dim dept As String
  dept = GetComboFilterValue(ComboDept)

  strSQL = "select distinct d.userid from T_OTHE_CROSS_DATA_DOWNLOAD d"
  If dept <> "" Then
    strSQL = strSQL & " join T_OTHE_CROSS_USERS_PROPERTIES p on d.userid = p.userid"
    strSQL = strSQL & " where p.department = '" & EscapeSQLText(dept) & "'"
  End If
  strSQL = strSQL & " order by d.userid"

  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  Do Until ORARst.EOF
    ComboUserID.AddItem CStr(ORARst!userID)
    ORARst.MoveNext
  Loop
  ORARst.Close
  ORAConn4.Close

  ComboUserID.ListIndex = 0
End Sub


Private Sub PopulateIPCombo()
  ComboIPSave.Clear
  ComboIPSave.AddItem "All"

  Dim userID As String, dept As String
  userID = GetComboFilterValue(ComboUserID)
  dept = GetComboFilterValue(ComboDept)

  strSQL = "select distinct d.ip_save from T_OTHE_CROSS_DATA_DOWNLOAD d"
  If userID <> "" Then
    strSQL = strSQL & " where d.userid = '" & EscapeSQLText(userID) & "'"
  ElseIf dept <> "" Then
    strSQL = strSQL & " join T_OTHE_CROSS_USERS_PROPERTIES p on d.userid = p.userid"
    strSQL = strSQL & " where p.department = '" & EscapeSQLText(dept) & "'"
  End If
  strSQL = strSQL & " order by d.ip_save"

  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  Do Until ORARst.EOF
    ComboIPSave.AddItem CStr(ORARst!ip_save)
    ORARst.MoveNext
  Loop
  ORARst.Close
  ORAConn4.Close

  ComboIPSave.ListIndex = 0
End Sub

Private Sub ComboDept_Click()
  Call PopulateUserCombo
  Call PopulateIPCombo
End Sub

Private Sub ComboUserID_Click()
  Call PopulateIPCombo
End Sub

Private Function GetValidatedDateRange(ByRef BDate As String, ByRef EDate As String) As Boolean
  GetValidatedDateRange = False
  If Not IsDate(ComboBDate.Text) Or Not IsDate(ComboEDate.Text) Then
    MsgBox "Please enter a valid start/end date", vbInformation, "Notice"
    Exit Function
  End If
  BDate = Format(CDate(ComboBDate.Text), "yyyy-mm-dd")
  EDate = Format(CDate(ComboEDate.Text), "yyyy-mm-dd")
  GetValidatedDateRange = True
End Function

Private Function EscapeSQLText(ByVal s As String) As String
  EscapeSQLText = Replace(s, "'", "''")
End Function

Private Sub SetGridHeaders()
  With HFGrid1
    .ColWidth(0) = 400
    .TextMatrix(0, 1) = "Department": .ColWidth(1) = 1800
    .TextMatrix(0, 2) = "Username": .ColWidth(2) = 900
    .TextMatrix(0, 3) = "Data-Download IP": .ColWidth(3) = 1400
    .TextMatrix(0, 4) = "Data Download Time": .ColWidth(4) = 2000
    .TextMatrix(0, 5) = "Data Station Type": .ColWidth(5) = 1400
    .TextMatrix(0, 6) = "Data Type": .ColWidth(6) = 1600
    .TextMatrix(0, 7) = "Spatial Range": .ColWidth(7) = 3000
    .TextMatrix(0, 8) = "Time Range": .ColWidth(8) = 2400
    .TextMatrix(0, 9) = "Data Length": .ColWidth(9) = 900
    .TextMatrix(0, 10) = "Element": .ColWidth(10) = 800
  End With
End Sub

Private Sub CmdRefresh_Click()
  Call RefreshAll
End Sub


Private Sub CmdSaveData_Click(Index As Integer)
  Dim i As Long, j As Long
  Dim strFileName As String
  Dim strTemp As String

  CommonDialogSave.Filter = "Text Files (*.csv)|*.csv"
  CommonDialogSave.FilterIndex = 1
  CommonDialogSave.FileName = "User Activity"
  CommonDialogSave.InitDir = App.Path & "\Output"
  CommonDialogSave.Flags = &H2
  CommonDialogSave.ShowSave
  If CommonDialogSave.Flags = 2 Then Exit Sub

  strFileName = CommonDialogSave.FileName
'  If LCase(Right(strBaseName, 4)) = ".csv" Then strBaseName = Left(strBaseName, Len(strBaseName) - 4)
'  strGridFile = strBaseName & "Grid.csv"
'  strChartFile = strBaseName & "Chart.csv"

  frmWait.Label1.Caption = "Saving, please wait"
  frmWait.Show
  Screen.MousePointer = 11


  If Index = 0 Then
    If Len(Dir(strFileName)) <> 0 Then Kill strFileName
    Open strFileName For Output As #1
    With HFGrid1
      For i = 0 To .Rows - 1
        strTemp = SanitizeForCSV(.TextMatrix(i, 0))
        For j = 1 To .Cols - 1
          strTemp = strTemp & "," & SanitizeForCSV(.TextMatrix(i, j))
        Next j
        Print #1, strTemp
      Next i
    End With
    Close #1

  ElseIf Index = 1 Then
    If Len(Dir(strFileName)) <> 0 Then Kill strFileName
    Open strFileName For Output As #1
    Print #1, "QUERY_TYPE,CNT"
    For i = 0 To mChartN - 1
      Print #1, SanitizeForCSV(mChartLabels(i)) & "," & mChartCounts(i)
    Next i
    Close #1

  ElseIf Index = 2 Then
    If Len(Dir(strFileName)) <> 0 Then Kill strFileName
    Open strFileName For Output As #1
    Print #1, "QDAY,ACTIVE_IP_COUNT,TOTAL_ACTIVE_IP,TOTAL_QUERY_COUNT"

    For i = 0 To mIPChartN - 1
      Print #1, SanitizeForCSV(mIPChartLabels(i)) & "," & mIPChartCounts(i) & "," & mIPTotalActive & "," & mIPTotalQueries
    Next i
    Close #1

  End If
  
  Screen.MousePointer = 1
  frmWait.Hide
  FrmMain.StatusBar1.Panels(1).Text = "Save result: Done"
End Sub


Private Sub RefreshAll()
  Dim BDate As String, EDate As String
  If Not GetValidatedDateRange(BDate, EDate) Then Exit Sub

  Call LoadDownloadLog(BDate, EDate)
  Call LoadQueryStats(BDate, EDate)
  Call LoadIPActivityStats(BDate, EDate)
End Sub

'获取用户下载数据记录
Private Sub LoadDownloadLog(ByVal BDate As String, ByVal EDate As String)
  On Error Resume Next
  Set HFGrid1.DataSource = Nothing
  On Error GoTo 0

  strSQL = "select p.department, d.userid, d.ip_save, to_char(d.ddatetime,'yyyy-mm-dd hh24:mi:ss') ddatetime,"
  strSQL = strSQL & " case d.statype"
  strSQL = strSQL & " when 'SURF' then '" & ("National Station") & "'"
  strSQL = strSQL & " when 'AWST' then '" & ("AWS Station") & "'"
  strSQL = strSQL & " when 'BOTH' then '" & ("National Station") & "+" & ("AWS Station") & "'"
  strSQL = strSQL & " else d.statype end statype,"
  strSQL = strSQL & " case d.datagap"
  strSQL = strSQL & " when '" & ("Hourly Conditional Query") & "' then '" & ("Conditional - Hourly Data") & "'"
  strSQL = strSQL & " when '" & ("Daily Conditional Query") & "' then '" & ("Conditional - Daily Data") & "'"
  strSQL = strSQL & " when '" & ("Dekad Conditional Query") & "' then '" & ("Conditional - Dekad Data") & "'"
  strSQL = strSQL & " when '" & ("Monthly Conditional Query") & "' then '" & ("Conditional - Monthly Data") & "'"
  strSQL = strSQL & " when '" & ("Seasonal Conditional Query") & "' then '" & ("Conditional - Seasonal Data") & "'"
  strSQL = strSQL & " when '" & ("Annual Conditional Query") & "' then '" & ("Conditional - Annual Data") & "'"
  strSQL = strSQL & " when '" & ("Year-by-Year Custom Period") & "' then '" & ("Year-by-Year Custom-Period Data") & "'"
  strSQL = strSQL & " else d.datagap || '" & " Data" & "' end datagap,"
  strSQL = strSQL & " d.spatial, d.temporal, d.datalength"
  strSQL = strSQL & " ,d.variable"   ' 2026-8-18
  strSQL = strSQL & " from T_OTHE_CROSS_DATA_DOWNLOAD d"
  strSQL = strSQL & " left join T_OTHE_CROSS_USERS_PROPERTIES p on d.userid = p.userid"
  strSQL = strSQL & " where d.ddatetime >= to_date('" & BDate & "','yyyy-mm-dd')"
  strSQL = strSQL & " and d.ddatetime < to_date('" & EDate & "','yyyy-mm-dd') + 1"

  Dim deptFilter As String, userFilter As String, ipFilter As String
  deptFilter = GetComboFilterValue(ComboDept)
  userFilter = GetComboFilterValue(ComboUserID)
  ipFilter = GetComboFilterValue(ComboIPSave)

  If deptFilter <> "" Then
    strSQL = strSQL & " and p.department = '" & EscapeSQLText(deptFilter) & "'"
  End If
  If userFilter <> "" Then
    strSQL = strSQL & " and d.userid = '" & EscapeSQLText(userFilter) & "'"
  End If
  If ipFilter <> "" Then
    strSQL = strSQL & " and d.ip_save = '" & EscapeSQLText(ipFilter) & "'"
  End If
  strSQL = strSQL & " order by d.ddatetime desc"

  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly

  Set HFGrid1.DataSource = ORARst
  Call SetGridHeaders

  Dim i As Long
  For i = 1 To HFGrid1.Rows - 1
    HFGrid1.TextMatrix(i, 0) = i
  Next i

  ORARst.Close
  ORAConn4.Close
End Sub


' 获取用户查询数据次数
Private Sub LoadQueryStats(ByVal BDate As String, ByVal EDate As String)
  strSQL = "select query_type, count(*) cnt from T_OTHE_CROSS_DATA_QUERY"
  strSQL = strSQL & " where ddatetime >= to_date('" & BDate & "','yyyy-mm-dd')"
  strSQL = strSQL & " and ddatetime < to_date('" & EDate & "','yyyy-mm-dd') + 1"
  strSQL = strSQL & " group by query_type order by count(*) desc"

  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly

  ReDim mChartLabels(0 To 20)
  ReDim mChartCounts(0 To 20)
  mChartN = 0

  Do Until ORARst.EOF
    If mChartN <= UBound(mChartLabels) Then
      mChartLabels(mChartN) = TranslateQueryType(CStr(ORARst!query_type))
      mChartCounts(mChartN) = CLng(ORARst!cnt)
      mChartN = mChartN + 1
    End If
    ORARst.MoveNext
  Loop

  ORARst.Close
  ORAConn4.Close

  If mChartN > 0 Then
    ReDim Preserve mChartLabels(0 To mChartN - 1)
    ReDim Preserve mChartCounts(0 To mChartN - 1)
  End If

  Call DrawPieChart(mChartLabels, mChartCounts, mChartN)
End Sub


' 每日活跃IP统计+柱状图
Private Sub LoadIPActivityStats(ByVal BDate As String, ByVal EDate As String)
  strSQL = "select to_char(ddatetime,'yyyy-mm-dd') qday, query_ip, count(*) cnt"
  strSQL = strSQL & " from T_OTHE_CROSS_DATA_QUERY"
  strSQL = strSQL & " where ddatetime >= to_date('" & BDate & "','yyyy-mm-dd')"
  strSQL = strSQL & " and ddatetime < to_date('" & EDate & "','yyyy-mm-dd') + 1"
  strSQL = strSQL & " group by to_char(ddatetime,'yyyy-mm-dd'), query_ip"
  strSQL = strSQL & " order by qday"

  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly

  Dim totalRows As Long
  totalRows = ORARst.RecordCount
  If totalRows < 1 Then totalRows = 1

  Dim UniqueIP() As String, UniqueIPCount As Long
  ReDim mIPChartLabels(0 To totalRows)
  ReDim mIPChartCounts(0 To totalRows)
  ReDim UniqueIP(0 To totalRows)
  mIPChartN = 0
  UniqueIPCount = 0
  mIPTotalQueries = 0

  Dim curDay As String, thisIP As String, found As Boolean, j As Long
  curDay = ""
  Do Until ORARst.EOF
    If CStr(ORARst!qday) <> curDay Then
      curDay = CStr(ORARst!qday)
      mIPChartLabels(mIPChartN) = curDay
      mIPChartCounts(mIPChartN) = 0
      mIPChartN = mIPChartN + 1
    End If
    mIPChartCounts(mIPChartN - 1) = mIPChartCounts(mIPChartN - 1) + 1

    mIPTotalQueries = mIPTotalQueries + CLng(ORARst!cnt)

    thisIP = CStr(ORARst!query_ip)
    found = False
    For j = 0 To UniqueIPCount - 1
      If UniqueIP(j) = thisIP Then found = True: Exit For
    Next j
    If Not found Then
      UniqueIP(UniqueIPCount) = thisIP
      UniqueIPCount = UniqueIPCount + 1
    End If

    ORARst.MoveNext
  Loop

  ORARst.Close
  ORAConn4.Close

  mIPTotalActive = UniqueIPCount

  If mIPChartN > 0 Then
    ReDim Preserve mIPChartLabels(0 To mIPChartN - 1)
    ReDim Preserve mIPChartCounts(0 To mIPChartN - 1)
  End If

  Call DrawIPBarChart(mIPChartLabels, mIPChartCounts, mIPChartN)
End Sub


Private Sub DrawIPBarChart(ByRef Labels() As String, ByRef Counts() As Long, ByVal N As Long)
  Picture1.Cls
  Picture1.Font.Name = "宋体"
  Picture1.Font.Size = 10

  If N <= 0 Then
    Picture1.CurrentX = 100: Picture1.CurrentY = 100
    Picture1.Print "No valid data available for plotting"
    Exit Sub
  End If

  Dim maxCount As Long, i As Long
  maxCount = 0
  For i = 0 To N - 1
    If Counts(i) > maxCount Then maxCount = Counts(i)
  Next i
  If maxCount < 1 Then maxCount = 1

  Dim titleY As Single
  titleY = 60
  Picture1.CurrentX = 100: Picture1.CurrentY = titleY
  Picture1.Print "Daily Active IP Count"

  Dim summaryText As String
  summaryText = "Total active IPs in query period: " & mIPTotalActive & ", data query count: " & mIPTotalQueries & " times"
  Picture1.CurrentX = Picture1.ScaleWidth - Picture1.TextWidth(summaryText) - 100
  Picture1.CurrentY = titleY
  Picture1.Print summaryText



  Dim leftM As Single, rightM As Single, topM As Single, botM As Single
  leftM = 900: rightM = 100: topM = 500: botM = 700

  Dim plotW As Single, plotH As Single
  plotW = Picture1.ScaleWidth - leftM - rightM
  plotH = Picture1.ScaleHeight - topM - botM
  If plotW < 10 Then plotW = 10
  If plotH < 10 Then plotH = 10

  Dim axisMax As Double
  axisMax = maxCount * 1.15  '顶部留白，避免最高的柱子紧贴边框
  Call ChartData.DrawYAxis(Picture1, leftM, topM, plotH, 0, axisMax, "Active IP Count")

  '===== 2026-8-26 添加横向网格线 =====
  Dim tickValues() As Double
  Dim tickCount As Integer
  
  Dim niceMin As Double, niceMax As Double, tickStep As Double
  Call ComputeNiceTicks(0, axisMax, 10, niceMin, niceMax, tickStep)
  
  Dim gridY As Single
  Picture1.DrawWidth = 1
  Picture1.DrawStyle = 2 ' 虚线
  For i = niceMin To niceMax - tickStep Step tickStep
    If i <= axisMax Then
      gridY = topM + plotH - (plotH * (i / axisMax))
      If i > 0 And i < axisMax Then
        Picture1.Line (leftM, gridY)-(leftM + plotW, gridY), RGB(180, 180, 180)
      End If
    End If
  Next i

  ' 重置回实线
  Picture1.DrawStyle = 0


  ' 2026-8-27：先画横向网格线，再画边框，避免虚线压住实线
  Picture1.DrawWidth = 1
  Picture1.Line (leftM, topM)-(leftM, topM + plotH), vbBlack
  Picture1.Line (leftM, topM + plotH)-(leftM + plotW, topM + plotH), vbBlack
  Picture1.Line (leftM, topM)-(leftM + plotW, topM), vbBlack
  Picture1.Line (leftM + plotW, topM)-(leftM + plotW, topM + plotH), vbBlack


  '显示用的标签另存一份再做年份省略，不改Labels()本身
  Dim DispLabels() As String
  ReDim DispLabels(0 To N - 1)
  For i = 0 To N - 1
    DispLabels(i) = Labels(i)
  Next i
  Call ChartData.ApplyYearElision(DispLabels, N)

  Dim labelStep As Long
  labelStep = 1
  If N > 15 Then labelStep = (N \ 15) + 1

  Dim barW As Single, x1 As Single, barH As Single, y1 As Single
  barW = plotW / N
  Picture1.FillStyle = 0
  Picture1.FillColor = RGB(70, 130, 180)

  For i = 0 To N - 1
    x1 = leftM + i * barW
    barH = plotH * (CDbl(Counts(i)) / axisMax)
    y1 = topM + plotH - barH
    Picture1.Line (x1 + barW * 0.15, y1)-(x1 + barW * 0.85, topM + plotH), RGB(70, 130, 180), BF

    If i Mod labelStep = 0 Then
      Picture1.CurrentX = x1 + barW / 2 - Picture1.TextWidth(DispLabels(i)) / 2
      Picture1.CurrentY = topM + plotH + 60
      Picture1.Print DispLabels(i)
    End If
  Next i

  Picture1.CurrentX = leftM + (plotW - Picture1.TextWidth("Date")) / 2
  Picture1.CurrentY = topM + plotH + 420
  Picture1.Print "Date"
End Sub


Private Sub LoadGuestChartCSV(ByVal CsvFileName As String)
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String

  FilePath = App.Path & "\Preview\Data\" & CsvFileName
  If Dir(FilePath) = "" Then Exit Sub

  LineCount = 0
  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then LineCount = LineCount + 1
  Loop
  Close #FileNum

  If LineCount > 0 Then
    ReDim Lines(0 To LineCount - 1)
    LineCount = 0

    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then
        Lines(LineCount) = tempStr
        LineCount = LineCount + 1
      End If
    Loop
    Close #FileNum
  End If

  If LineCount < 2 Then Exit Sub

  ReDim mChartLabels(0 To LineCount - 2)
  ReDim mChartCounts(0 To LineCount - 2)
  mChartN = 0

  For i = 1 To LineCount - 1
    Fields = Split(Lines(i), ",")
    If UBound(Fields) >= 1 Then
      mChartLabels(mChartN) = Fields(0)
      mChartCounts(mChartN) = CLng(Fields(1))
      mChartN = mChartN + 1
    End If
  Next i

  Call DrawPieChart(mChartLabels, mChartCounts, mChartN)
End Sub


'2026-08-17 访客模式下每日活跃IP柱状图的数据来源
Private Sub LoadGuestIPActivityCSV(ByVal CsvFileName As String)
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String

  FilePath = App.Path & "\Preview\Data\" & CsvFileName
  If Dir(FilePath) = "" Then Exit Sub

  LineCount = 0
  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then LineCount = LineCount + 1
  Loop
  Close #FileNum

  If LineCount > 0 Then
    ReDim Lines(0 To LineCount - 1)
    LineCount = 0

    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then
        Lines(LineCount) = tempStr
        LineCount = LineCount + 1
      End If
    Loop
    Close #FileNum
  End If

  If LineCount < 2 Then Exit Sub

  ReDim mIPChartLabels(0 To LineCount - 2)
  ReDim mIPChartCounts(0 To LineCount - 2)
  mIPChartN = 0
  mIPTotalActive = 0
  mIPTotalQueries = 0

  For i = 1 To LineCount - 1
    Fields = Split(Lines(i), ",")
    If UBound(Fields) >= 1 Then
      mIPChartLabels(mIPChartN) = Fields(0)
      mIPChartCounts(mIPChartN) = CLng(Fields(1))
      If mIPChartN = 0 And UBound(Fields) >= 3 Then
        mIPTotalActive = CLng(Fields(2))
        mIPTotalQueries = CLng(Fields(3))
      End If
      mIPChartN = mIPChartN + 1
    End If
  Next i

  Call DrawIPBarChart(mIPChartLabels, mIPChartCounts, mIPChartN)
End Sub

Private Function TranslateQueryType(ByVal RawType As String) As String
  Select Case RawType
    Case "Hour"
      TranslateQueryType = "Hourly Data"
    Case "Day"
      TranslateQueryType = "Daily Data"
    Case "Ten"
      TranslateQueryType = "Dekad Data"
    Case "Mon"
      TranslateQueryType = "Monthly Data"
    Case "Quarter"
      TranslateQueryType = "Seasonal Data"
    Case "Year"
      TranslateQueryType = "Annual Data"
    Case "Period"
      TranslateQueryType = "Custom Period"
    Case "SinglePeriod"
      TranslateQueryType = "Single Period"
    Case "MultiPeriod"
      TranslateQueryType = "Year-by-Year Custom Period"
    Case "ColdAir"
      TranslateQueryType = "Cold Air Outbreak"
    Case "Season"
      TranslateQueryType = "Season Classification Dates"
    Case "PeriodExt"
      TranslateQueryType = "Regional Statistics"
    Case "Condition"
      TranslateQueryType = "Conditional Query"
      
      
    Case Else
      TranslateQueryType = RawType
  End Select
End Function

Private Sub DrawPieChart(ByRef Labels() As String, ByRef Counts() As Long, ByVal N As Long)
  PicChart.Cls

  If N <= 0 Then
    PicChart.CurrentX = 100: PicChart.CurrentY = 100
    PicChart.Print "No valid data available for plotting"
    Exit Sub
  End If

  Dim total As Long, i As Long
  total = 0
  For i = 0 To N - 1
    total = total + Counts(i)
  Next i
  If total <= 0 Then Exit Sub

  Dim Colors(0 To 9) As Long
  Colors(0) = RGB(70, 130, 180)
  Colors(1) = RGB(220, 20, 60)
  Colors(2) = RGB(60, 179, 113)
  Colors(3) = RGB(255, 165, 0)
  Colors(4) = RGB(147, 112, 219)
  Colors(5) = RGB(184, 134, 11)
  Colors(6) = RGB(0, 191, 255)
  Colors(7) = RGB(199, 21, 133)
  Colors(8) = RGB(154, 205, 50)
  Colors(9) = RGB(105, 105, 105)

  PicChart.CurrentX = 100: PicChart.CurrentY = 100
  PicChart.Print "Feature Module Query Count"

  Const PI As Double = 3.14159265358979
  '2026-08-17 图例改到右侧、饼图相应左移：右边留出legendW宽度专门画图例，
  '饼图只在剩下的左侧区域里居中，不再跟图例上下叠放
  Dim titleH As Single, legendW As Single, pieAreaW As Single
  titleH = 400
  legendW = PicChart.ScaleWidth * 0.32
  pieAreaW = PicChart.ScaleWidth - legendW

  Dim cx As Single, cy As Single, radius As Single
  cx = pieAreaW / 2
  cy = titleH + (PicChart.ScaleHeight - titleH) / 2
  radius = pieAreaW * 0.42
  If (PicChart.ScaleHeight - titleH) * 0.45 < radius Then radius = (PicChart.ScaleHeight - titleH) * 0.45

  PicChart.FillStyle = 0
  Dim cumAngle As Double, sweep As Double, a1 As Double, a2 As Double
  cumAngle = 0
  For i = 0 To N - 1
    sweep = (CDbl(Counts(i)) / CDbl(total)) * 2 * PI
    a1 = cumAngle
    a2 = cumAngle + sweep
    ' VB6 不能使用-0，但是负角度可以强制圆闭合，所以在这里使用一个很小的负数。
    If a1 <= 0.0001 Then a1 = 0.0001
    If a2 >= 2 * PI - 0.0001 Then a2 = 2 * PI - 0.0001
    PicChart.FillColor = Colors(i Mod 10)
    PicChart.Circle (cx, cy), radius, Colors(i Mod 10), -a1, -a2
    cumAngle = a2
  Next i

  Dim legendX As Single, legendTop As Single, rowH As Single, sw As Single
  legendX = pieAreaW + 200
  sw = 200
  rowH = (PicChart.ScaleHeight - titleH - 200) / N
  If rowH > 380 Then rowH = 380
  If rowH < 200 Then rowH = 200
  '图例整体在纵向居中，而不是从顶部贴着往下排
  legendTop = titleH + (PicChart.ScaleHeight - titleH - N * rowH) / 2
  If legendTop < titleH Then legendTop = titleH

  Dim pct As Double, Y As Single
  For i = 0 To N - 1
    Y = legendTop + i * rowH
    PicChart.Line (legendX, Y)-(legendX + sw, Y + sw), Colors(i Mod 10), BF
    pct = CDbl(Counts(i)) / CDbl(total) * 100
    PicChart.CurrentX = legendX + sw + 200
    PicChart.CurrentY = Y
    PicChart.Print Labels(i) & "  " & CStr(Counts(i)) & "  (" & Format(pct, "0.0") & "%)"
  Next i
End Sub



Private Sub HFGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid1
  If Button = 2 Then PopupMenu mnuGridPopup
  Dim i%
  With HFGrid1
  If shift = 2 Then
    .Sort = 7
    If .Col = .Cols - 1 Then .Sort = 3
  ElseIf shift = 4 Then
    .Sort = 8
    If .Col = .Cols - 1 Then .Sort = 4
  End If
  For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
  Next i
  .Refresh
  End With
End Sub

Private Sub PicChart_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickChart = PicChart
  If Button = 2 Then PopupMenu mnuChartPopup
End Sub


Private Sub Picture1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickChart = Picture1
  If Button = 2 Then PopupMenu mnuChartPopup
End Sub

Private Sub mnuCopyChart_Click()
  If mLastRightClickChart Is Nothing Then Exit Sub
  Clipboard.Clear
  Clipboard.SetData mLastRightClickChart.image, vbCFBitmap
End Sub

Private Sub mnuCopyTable_Click()
  Call CopyTableToClipboard(HFGrid1)
End Sub

Private Sub CopyTableToClipboard(ByRef Grid As Object)
  Dim r As Long, c As Long, LineText As String, OutText As String
  With Grid
    For r = 0 To .Rows - 1
      LineText = ""
      For c = 0 To .Cols - 1
        If c > 0 Then LineText = LineText & vbTab
        LineText = LineText & .TextMatrix(r, c)
      Next c
      OutText = OutText & LineText & vbCrLf
    Next r
  End With
  Clipboard.Clear
  Clipboard.SetText OutText
End Sub


Private Sub cmdSavePNG_Click(Index As Integer)
  Dim TargetPic As Object

  If Index = 0 Then
    Set TargetPic = PicChart
    CommonDialogSave.FileName = "User Activity Chart"
  Else
    Set TargetPic = Picture1
    CommonDialogSave.FileName = "Daily Active IP Chart"
  End If

  CommonDialogSave.Filter = "PNG Images (*.png)|*.png"
  CommonDialogSave.FilterIndex = 1
  CommonDialogSave.InitDir = App.Path & "\Output"
  CommonDialogSave.Flags = &H2
  CommonDialogSave.ShowSave
  If CommonDialogSave.Flags = 2 Then Exit Sub

  Call SaveChartAsPNG(TargetPic, CommonDialogSave.FileName)
  FrmMain.StatusBar1.Panels(1).Text = "Chart saved"
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


