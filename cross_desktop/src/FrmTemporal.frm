VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmTemporal 
   Caption         =   "Time Series Chart"
   ClientHeight    =   9420
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   21390
   Icon            =   "FrmTemporal.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   21390
   StartUpPosition =   2  '屏幕中心
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CheckBox CheckRegression 
      Caption         =   "Trend Equation"
      Height          =   255
      Left            =   19080
      TabIndex        =   12
      Top             =   1200
      Width           =   1935
   End
   Begin VB.CommandButton cmdSaveCSV 
      Caption         =   "Save Table"
      Height          =   450
      Left            =   20040
      TabIndex        =   11
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSavePNG 
      Caption         =   "Save Chart"
      Height          =   450
      Left            =   18720
      TabIndex        =   10
      Top             =   8760
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Manual Adjustment"
      Height          =   3135
      Left            =   16800
      TabIndex        =   7
      Top             =   1680
      Width           =   4455
      Begin VB.TextBox TextCircleSize 
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         Top             =   2520
         Width           =   615
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   3720
         TabIndex        =   28
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox TextFontSize 
         Height          =   375
         Left            =   1800
         TabIndex        =   27
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox TextExportHeight 
         Height          =   375
         Left            =   600
         TabIndex        =   26
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox TextYaxisStep 
         Height          =   300
         Left            =   2520
         TabIndex        =   20
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisMax 
         Height          =   300
         Left            =   2520
         TabIndex        =   19
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisMin 
         Height          =   300
         Left            =   2520
         TabIndex        =   18
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TextXaxisStep 
         Height          =   300
         Left            =   600
         TabIndex        =   17
         Top             =   2040
         Width           =   1815
      End
      Begin VB.ComboBox ComboXaxisMax 
         Height          =   300
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   1560
         Width           =   1815
      End
      Begin VB.ComboBox ComboXaxisMin 
         Height          =   300
         Left            =   600
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisName 
         Height          =   375
         Left            =   2520
         TabIndex        =   9
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TextXaxisName 
         Height          =   375
         Left            =   600
         TabIndex        =   8
         Top             =   480
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   " Point Size"
         Height          =   375
         Left            =   2520
         TabIndex        =   31
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Font Size"
         Height          =   375
         Left            =   1320
         TabIndex        =   30
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label LblExportHeight 
         Caption         =   "Export Image Height"
         Height          =   375
         Left            =   120
         TabIndex        =   29
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1150
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "End"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   1600
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Inter"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   2100
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Y-axis"
         Height          =   255
         Left            =   3120
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "X-axis"
         Height          =   255
         Left            =   1200
         TabIndex        =   13
         Top             =   240
         Width           =   615
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridStats 
      Height          =   3510
      Left            =   16800
      TabIndex        =   5
      Top             =   5040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6191
      _Version        =   393216
      Rows            =   10
      FixedRows       =   0
      FixedCols       =   0
      BackColorFixed  =   14675935
      BackColorBkg    =   16777215
      GridColor       =   12632256
      GridLinesFixed  =   1
      AllowUserResizing=   1
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.PictureBox PicChart 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      Height          =   9060
      Left            =   120
      ScaleHeight     =   9000
      ScaleWidth      =   16440
      TabIndex        =   4
      Top             =   225
      Width           =   16500
   End
   Begin VB.PictureBox PicExport 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8160
      Left            =   -30000
      ScaleHeight     =   8160
      ScaleWidth      =   16500
      TabIndex        =   25
      Top             =   225
      Visible         =   0   'False
      Width           =   16500
   End
   Begin VB.CheckBox CheckLine 
      Caption         =   "Add Trend Line"
      Height          =   255
      Left            =   16800
      TabIndex        =   3
      Top             =   1200
      Width           =   1575
   End
   Begin VB.ComboBox ComboChartType 
      Height          =   300
      Left            =   19080
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox ComboSeries 
      Height          =   300
      Left            =   16800
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label LblChartType 
      Caption         =   "Chart Type: "
      Height          =   255
      Left            =   19080
      TabIndex        =   1
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label LblMode 
      Caption         =   "Row/column for the data: "
      Height          =   255
      Left            =   16800
      TabIndex        =   6
      Top             =   360
      Width           =   2175
   End
   Begin VB.Menu mnuChartPopup 
      Caption         =   "Chart Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyChart 
         Caption         =   "Copy Chart"
      End
   End
   Begin VB.Menu mnuTablePopup 
      Caption         =   "Table Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuCopyTable 
         Caption         =   "Copy Table"
      End
   End
End
Attribute VB_Name = "FrmTemporal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGrid As Object
Private mFixedRows As Long
Private mFixedCols As Long
Private mYAxisTitle As String
Private mChartType As String
Private mShowTrend As Boolean
Private mIsRainfall As Boolean

Private mLastValues() As Double
Private mLastLabels() As String
Private mLastIsMissing() As Boolean
Private mLastN As Long


Private Sub CheckLine_Click()
  Call RefreshChart(False)

End Sub

Private Sub CheckRegression_Click()
  Call RefreshChart(False)
End Sub

Private Sub ComboChartType_click()
  Call RefreshChart(False)

End Sub


Private Sub ComboSeries_Click()
  '2026-08-19 换了行/站点，Y轴数据范围整个变了，之前手动填的min/max/step不再适用
  TextYaxisMin.Text = ""
  TextYaxisMax.Text = ""
  TextYaxisStep.Text = ""
  Call RefreshChart(True)
End Sub


Private Sub CmdRefresh_Click()
  Call RefreshChart(False)
End Sub

Private Sub Form_Load()
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
End Sub

' 由ChartData.ShowTemporalChartForActiveForm在这个实例创建时调用一次
Public Sub InitFromSelection(ByVal SrcGrid As Object, ByVal SrcCaption As String, ByVal YAxisTitle As String)
  Set mGrid = SrcGrid
  mFixedRows = mGrid.FixedRows
  mFixedCols = mGrid.FixedCols
  mYAxisTitle = YAxisTitle
  mIsRainfall = (InStr(mYAxisTitle, "雨量") > 0)

  Call PopulateComboSeries
  Call SelectComboIndex(mGrid.Row)
  Call PopulateComboChartType

  Me.Caption = SrcCaption & " - " & ("Time Series Chart")
  CheckLine.Caption = "Add Trend Line"
  LblMode.Caption = "Select Row: "
  LblChartType.Caption = "Chart Type: "
  mnuCopyChart.Caption = "Copy Chart"
  mnuCopyTable.Caption = "Copy Table"
  TextFontSize.Text = "12"

  TextExportHeight.Text = CStr(PicChart.ScaleHeight \ Screen.TwipsPerPixelY)

  TextCircleSize.Text = "60"
  Call ApplyDefaultTitles
  Call RefreshChart(True)
End Sub


Private Sub ApplyDefaultTitles()
  TextXaxisName.Text = "Time"
  TextYaxisName.Text = mYAxisTitle
End Sub

Private Sub PopulateComboSeries()
  Dim i As Long
  ComboSeries.Clear
  For i = mFixedRows To mGrid.Rows - 1
    ComboSeries.AddItem ChartData.BuildRowLabel(mGrid, i)
    ComboSeries.ItemData(ComboSeries.NewIndex) = i
  Next i
End Sub

Private Sub SelectComboIndex(ByVal DataIdx As Long)
  Dim j As Long
  For j = 0 To ComboSeries.ListCount - 1
    If ComboSeries.ItemData(j) = DataIdx Then
      ComboSeries.ListIndex = j
      Exit Sub
    End If
  Next j
  If ComboSeries.ListCount > 0 Then ComboSeries.ListIndex = 0
End Sub

Private Sub PopulateComboChartType()
  ComboChartType.Clear
  ComboChartType.AddItem "Line"
  ComboChartType.AddItem "Bar"
  ' 降水量是逐日/逐时段的离散总量，用柱状图表达比连成一条折线更符合直觉
  If mIsRainfall Then
    ComboChartType.ListIndex = 1
  Else
    ComboChartType.ListIndex = 0
  End If
End Sub


Private Sub RefreshChart(ByVal ResetXRange As Boolean)
  If mGrid Is Nothing Then Exit Sub
  If ComboSeries.ListIndex < 0 Then Exit Sub

  If ComboChartType.ListIndex = 1 Then
    mChartType = "BAR"
  Else
    mChartType = "LINE"
  End If
  mShowTrend = (CheckLine.Value = 1)

  Dim DataIdx As Long
  DataIdx = ComboSeries.ItemData(ComboSeries.ListIndex)

  Dim FullLabels() As String, FullValues() As Double, FullIsMissing() As Boolean, FullN As Long
  Call ChartData.ExtractTemporalSeries(mGrid, DataIdx, FullLabels, FullValues, FullIsMissing, FullN)

  If ResetXRange Then Call PopulateComboXRange(FullLabels, FullN)

  ' X轴起止点/间隔天数是从完整序列里截取出一个子集，供绘图和统计表共用同一份数据
  Dim Labels() As String, Values() As Double, IsMissing() As Boolean, N As Long
  Call ApplyXRangeFilter(FullLabels, FullValues, FullIsMissing, FullN, Labels, Values, IsMissing, N)
  Call ChartData.ApplyYearElision(Labels, N)

  Call DrawTemporalChart(Values, Labels, IsMissing, N)
  Call FillStatsTable(Values, IsMissing, N)

  mLastValues = Values
  mLastLabels = Labels
  mLastIsMissing = IsMissing
  mLastN = N
End Sub


Private Sub PopulateComboXRange(ByRef FullLabels() As String, ByVal FullN As Long)
  Dim i As Long
  ComboXaxisMin.Clear
  ComboXaxisMax.Clear
  For i = 0 To FullN - 1
    ComboXaxisMin.AddItem FullLabels(i): ComboXaxisMin.ItemData(ComboXaxisMin.NewIndex) = i
    ComboXaxisMax.AddItem FullLabels(i): ComboXaxisMax.ItemData(ComboXaxisMax.NewIndex) = i
  Next i
  If ComboXaxisMin.ListCount > 0 Then ComboXaxisMin.ListIndex = 0
  If ComboXaxisMax.ListCount > 0 Then ComboXaxisMax.ListIndex = ComboXaxisMax.ListCount - 1
End Sub


Private Sub ApplyXRangeFilter(ByRef FullLabels() As String, ByRef FullValues() As Double, ByRef FullIsMissing() As Boolean, ByVal FullN As Long, _
                               ByRef OutLabels() As String, ByRef OutValues() As Double, ByRef OutIsMissing() As Boolean, ByRef OutN As Long)
  If FullN <= 0 Then
    ReDim OutLabels(0 To 0)
    ReDim OutValues(0 To 0)
    ReDim OutIsMissing(0 To 0)
    OutN = 0
    Exit Sub
  End If

  Dim startIdx As Long, endIdx As Long
  startIdx = 0
  endIdx = FullN - 1
  If ComboXaxisMin.ListIndex >= 0 Then startIdx = ComboXaxisMin.ItemData(ComboXaxisMin.ListIndex)
  If ComboXaxisMax.ListIndex >= 0 Then endIdx = ComboXaxisMax.ItemData(ComboXaxisMax.ListIndex)
  If startIdx > endIdx Then
    Dim tmp As Long
    tmp = startIdx: startIdx = endIdx: endIdx = tmp
  End If

  Dim cnt As Long, i As Long, j As Long
  cnt = endIdx - startIdx + 1

  ReDim OutLabels(0 To cnt - 1)
  ReDim OutValues(0 To cnt - 1)
  ReDim OutIsMissing(0 To cnt - 1)
  j = 0
  For i = startIdx To endIdx
    OutLabels(j) = FullLabels(i)
    OutValues(j) = FullValues(i)
    OutIsMissing(j) = FullIsMissing(i)
    j = j + 1
  Next i
  OutN = cnt
End Sub


Private Sub DrawTemporalChart(ByRef Values() As Double, ByRef Labels() As String, ByRef IsMissing() As Boolean, ByVal N As Long, Optional ByVal TargetPic As PictureBox = Nothing)
  '2026-08-18 TargetPic必须声明成具体的PictureBox类型，不能是泛型Object——Line/Circle
  Dim DrawPic As PictureBox
  If TargetPic Is Nothing Then
    Set DrawPic = Me.PicChart
  Else
    Set DrawPic = TargetPic
  End If

  DrawPic.Cls

  Dim chartFontSize As Single, scaleFactor As Single
  chartFontSize = ChartData.GetValidFontSize(TextFontSize.Text, 12)
  DrawPic.Font.Name = "宋体"
  DrawPic.Font.Size = chartFontSize
  scaleFactor = chartFontSize / 12
  Dim dMin As Double, dMax As Double, i As Long, hasValid As Boolean
  hasValid = False
  dMin = 0: dMax = 0
  For i = 0 To N - 1
    If Not IsMissing(i) Then
      If Not hasValid Then
        dMin = Values(i): dMax = Values(i): hasValid = True
      Else
        If Values(i) < dMin Then dMin = Values(i)
        If Values(i) > dMax Then dMax = Values(i)
      End If
    End If
  Next i

  If N <= 0 Or Not hasValid Then
    DrawPic.CurrentX = 100: DrawPic.CurrentY = 100
    DrawPic.Print "NoneYes效数据可供绘图"
    Exit Sub
  End If

  Dim pad As Double
  If dMax = dMin Then
    If Abs(dMax) < 1 Then pad = 1 Else pad = Abs(dMax) * 0.1
  Else
    pad = (dMax - dMin) * 0.1
  End If
  Dim axisMin As Double, axisMax As Double
  If mIsRainfall Then
    axisMin = 0
    axisMax = dMax + pad
  Else
    axisMin = dMin - pad
    axisMax = dMax + pad
  End If
  If axisMax = axisMin Then axisMax = axisMin + 1


  Dim userAxisVal As Double, yForcedStep As Double
  If ChartData.TryParseAxisValue(TextYaxisMin.Text, userAxisVal) Then axisMin = userAxisVal
  If ChartData.TryParseAxisValue(TextYaxisMax.Text, userAxisVal) Then axisMax = userAxisVal
  If axisMax <= axisMin Then axisMax = axisMin + 1
  yForcedStep = 0
  If ChartData.TryParseAxisValue(TextYaxisStep.Text, userAxisVal) And userAxisVal > 0 Then yForcedStep = userAxisVal


  Dim dispYMin As Double, dispYMax As Double, dispYStep As Double
  If yForcedStep > 0 Then
    dispYStep = yForcedStep
  Else
    Call ChartData.ComputeNiceTicks(axisMin, axisMax, 10, dispYMin, dispYMax, dispYStep)
  End If
  TextYaxisMin.Text = Format(axisMin, "0.0")
  TextYaxisMax.Text = Format(axisMax, "0.0")
  TextYaxisStep.Text = Format(dispYStep, "0.0")

  Dim leftM As Single, rightM As Single, topM As Single, botM As Single
  leftM = 1100 * scaleFactor: rightM = 500 * scaleFactor: topM = 200 * scaleFactor: botM = 700 * scaleFactor

  Dim plotW As Single, plotH As Single
  plotW = DrawPic.ScaleWidth - leftM - rightM
  plotH = DrawPic.ScaleHeight - topM - botM
  If plotW < 10 Then plotW = 10
  If plotH < 10 Then plotH = 10

  DrawPic.DrawWidth = 1
  DrawPic.Line (leftM, topM)-(leftM, topM + plotH), vbBlack
  DrawPic.Line (leftM, topM + plotH)-(leftM + plotW, topM + plotH), vbBlack
  DrawPic.Line (leftM, topM)-(leftM + plotW, topM), vbBlack
  DrawPic.Line (leftM + plotW, topM)-(leftM + plotW, topM + plotH), vbBlack

  Call ChartData.DrawYAxis(DrawPic, leftM, topM, plotH, axisMin, axisMax, TextYaxisName.Text, yForcedStep)


  Dim xPos() As Single
  ReDim xPos(0 To N - 1)
  If mChartType = "BAR" Then
    Dim slot As Single
    slot = plotW / N
    For i = 0 To N - 1
      xPos(i) = leftM + slot * i + slot / 2
    Next i
  Else
    Dim stepX As Single
    If N > 1 Then stepX = plotW / (N - 1) Else stepX = 0
    For i = 0 To N - 1
      xPos(i) = leftM + stepX * i
    Next i
  End If

  Dim yPos() As Single
  ReDim yPos(0 To N - 1)
  For i = 0 To N - 1
    yPos(i) = topM + plotH - CSng((Values(i) - axisMin) / (axisMax - axisMin) * plotH)
  Next i

  If mChartType = "BAR" Then
    Dim barW As Single, zeroY As Single
    barW = (plotW / N) * 0.7
    zeroY = topM + plotH - CSng((0 - axisMin) / (axisMax - axisMin) * plotH)
    If zeroY < topM Then zeroY = topM
    If zeroY > topM + plotH Then zeroY = topM + plotH
    DrawPic.FillStyle = 0
    For i = 0 To N - 1
      If Not IsMissing(i) Then
        DrawPic.Line (xPos(i) - barW / 2, yPos(i))-(xPos(i) + barW / 2, zeroY), RGB(70, 130, 180), BF
      End If
    Next i
  Else
    DrawPic.DrawWidth = 2
    For i = 1 To N - 1
      If Not IsMissing(i - 1) And Not IsMissing(i) Then
        DrawPic.Line (xPos(i - 1), yPos(i - 1))-(xPos(i), yPos(i)), RGB(70, 130, 180)
      End If
    Next i

    Dim circleSize As Double
    circleSize = 80
    If ChartData.TryParseAxisValue(TextCircleSize.Text, userAxisVal) And userAxisVal > 0 Then circleSize = userAxisVal

    DrawPic.FillStyle = 0
    DrawPic.FillColor = RGB(255, 255, 255)
    DrawPic.DrawWidth = 2
    For i = 0 To N - 1
      If Not IsMissing(i) Then
        DrawPic.Circle (xPos(i), yPos(i)), circleSize, RGB(70, 130, 180)
      End If
    Next i
    DrawPic.DrawWidth = 1
  End If

  ' 趋势线的起止点是第一个到最后一个非缺测点——用它们在原序列里的真实列号
  Dim firstValid As Long, lastValid As Long
  firstValid = -1: lastValid = -1
  For i = 0 To N - 1
    If Not IsMissing(i) Then
      If firstValid = -1 Then firstValid = i
      lastValid = i
    End If
  Next i

  Dim trendValid As Boolean
  Dim tSlope As Double, tIntercept As Double, tR2 As Double
  trendValid = (firstValid >= 0 And lastValid > firstValid)
  If trendValid Then Call ComputeLinearTrendWithGaps(Values, IsMissing, N, tSlope, tIntercept, tR2)

  If trendValid And mShowTrend Then
    Dim yStart As Double, yEnd As Double
    yStart = tSlope * firstValid + tIntercept
    yEnd = tSlope * lastValid + tIntercept
    Dim ty0 As Single, ty1 As Single
    ty0 = topM + plotH - CSng((yStart - axisMin) / (axisMax - axisMin) * plotH)
    ty1 = topM + plotH - CSng((yEnd - axisMin) / (axisMax - axisMin) * plotH)
    If ty0 < topM Then ty0 = topM
    If ty0 > topM + plotH Then ty0 = topM + plotH
    If ty1 < topM Then ty1 = topM
    If ty1 > topM + plotH Then ty1 = topM + plotH
    DrawPic.DrawStyle = 1
    DrawPic.DrawWidth = 2
    DrawPic.Line (xPos(firstValid), ty0)-(xPos(lastValid), ty1), vbRed
    DrawPic.DrawStyle = 0
    DrawPic.DrawWidth = 1
  End If

  If trendValid And CheckRegression.Value = 1 Then
    Dim eqText As String, signStr As String
    signStr = "+"
    If tIntercept < 0 Then signStr = "-"
    eqText = "y = " & Format(tSlope, "0.000") & "x " & signStr & " " & Format(Abs(tIntercept), "0.000")

    Dim eqOffsetX As Single, eqOffsetY As Single
    eqOffsetX = 200
    eqOffsetY = 40

    Dim lineGap As Single
    lineGap = DrawPic.TextHeight("Ag") + 20

    DrawPic.CurrentX = leftM + eqOffsetX
    DrawPic.CurrentY = topM + eqOffsetY
    DrawPic.Print eqText

    Dim normalSize As Single, r2LineY As Single
    normalSize = DrawPic.Font.Size
    r2LineY = topM + eqOffsetY + lineGap

    DrawPic.CurrentX = leftM + eqOffsetX
    DrawPic.CurrentY = r2LineY
    DrawPic.Print "R";

    DrawPic.Font.Size = normalSize * 0.7
    DrawPic.CurrentY = r2LineY - DrawPic.TextHeight("Ag") * 0.3
    DrawPic.Print "2";

    DrawPic.Font.Size = normalSize
    DrawPic.CurrentY = r2LineY
    DrawPic.Print " = " & Format(tR2, "0.000")
  End If

  Dim labelStep As Long
  labelStep = 1
  If N > 15 Then labelStep = (N \ 15) + 1
  If ChartData.TryParseAxisValue(TextXaxisStep.Text, userAxisVal) And userAxisVal >= 1 Then labelStep = CLng(userAxisVal)
  TextXaxisStep.Text = CStr(labelStep)

  For i = 0 To N - 1
    If i Mod labelStep = 0 Then
      DrawPic.Line (xPos(i), topM + plotH)-(xPos(i), topM + plotH + 60), vbBlack
      DrawPic.CurrentX = xPos(i) - DrawPic.TextWidth(Labels(i)) / 2
      DrawPic.CurrentY = topM + plotH + 80
      DrawPic.Print Labels(i)
    End If
  Next i

  DrawPic.CurrentX = leftM + (plotW - DrawPic.TextWidth(TextXaxisName.Text)) / 2
  DrawPic.CurrentY = topM + plotH + 420 * scaleFactor
  DrawPic.Print TextXaxisName.Text
End Sub


Private Sub ComputeLinearTrendWithGaps(ByRef Values() As Double, ByRef IsMissing() As Boolean, ByVal N As Long, _
                                        ByRef Slope As Double, ByRef Intercept As Double, ByRef RSquared As Double)
  Slope = 0: Intercept = 0: RSquared = 0
  Dim i As Long, cnt As Long
  Dim sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double
  cnt = 0
  For i = 0 To N - 1
    If Not IsMissing(i) Then
      sumX = sumX + i
      sumY = sumY + Values(i)
      sumXY = sumXY + i * Values(i)
      sumX2 = sumX2 + i * i
      cnt = cnt + 1
    End If
  Next i
  If cnt < 2 Then
    If cnt = 1 Then Intercept = sumY
    Exit Sub
  End If

  Dim denom As Double
  denom = cnt * sumX2 - sumX * sumX
  If denom = 0 Then Exit Sub

  Slope = (cnt * sumXY - sumX * sumY) / denom
  Intercept = (sumY - Slope * sumX) / cnt

  Dim meanY As Double, ssTot As Double, ssRes As Double, predicted As Double
  meanY = sumY / cnt
  For i = 0 To N - 1
    If Not IsMissing(i) Then
      predicted = Slope * i + Intercept
      ssTot = ssTot + (Values(i) - meanY) * (Values(i) - meanY)
      ssRes = ssRes + (Values(i) - predicted) * (Values(i) - predicted)
    End If
  Next i
  If ssTot > 0 Then RSquared = 1 - ssRes / ssTot
End Sub

Private Sub FillStatsTable(ByRef Values() As Double, ByRef IsMissing() As Boolean, ByVal N As Long)
  Dim CompactValues() As Double, cN As Long, j As Long
  ReDim CompactValues(0 To N)
  cN = 0
  For j = 0 To N - 1
    If Not IsMissing(j) Then
      CompactValues(cN) = Values(j)
      cN = cN + 1
    End If
  Next j

  Dim dMean As Double, dMax As Double, dMin As Double, dSum As Double, dStd As Double
  Call ChartData.ComputeStats(CompactValues, cN, dMean, dMax, dMin, dSum, dStd)

  Dim tSlope As Double, tIntercept As Double, tR2 As Double
  Call ComputeLinearTrendWithGaps(Values, IsMissing, N, tSlope, tIntercept, tR2)

  Dim dRange As Double, dDelta As Double
  dRange = dMax - dMin
  If cN > 0 Then dDelta = CompactValues(cN - 1) - CompactValues(0)

  Dim FStat As Double, PValue As Double, isSig As Boolean
  Call ChartData.ComputeTrendSignificance(cN, tR2, FStat, PValue, isSig)

  With GridStats
    .Rows = 13: .Cols = 2
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .TextMatrix(0, 0) = "Sample Size": .TextMatrix(0, 1) = CStr(cN)
    .TextMatrix(1, 0) = "Average": .TextMatrix(1, 1) = Format(dMean, "0.0")
    .TextMatrix(2, 0) = "Max": .TextMatrix(2, 1) = Format(dMax, "0.0")
    .TextMatrix(3, 0) = "Min": .TextMatrix(3, 1) = Format(dMin, "0.0")
    .TextMatrix(4, 0) = "Sum": .TextMatrix(4, 1) = Format(dSum, "0.0")
    .TextMatrix(5, 0) = "Std Dev": .TextMatrix(5, 1) = Format(dStd, "0.0")
    .TextMatrix(6, 0) = "Range": .TextMatrix(6, 1) = Format(dRange, "0.0")
    .TextMatrix(7, 0) = "First-Last Change": .TextMatrix(7, 1) = Format(dDelta, "0.0")
    .TextMatrix(8, 0) = "Trend Slope": .TextMatrix(8, 1) = Format(tSlope, "0.000")
    .TextMatrix(9, 0) = "Determination Coeff": .TextMatrix(9, 1) = Format(tR2, "0.000")
    .TextMatrix(10, 0) = "F Value": .TextMatrix(10, 1) = Format(FStat, "0.00")
    .TextMatrix(11, 0) = "P Value": .TextMatrix(11, 1) = Format(PValue, "0.0000")
    .TextMatrix(12, 0) = "0.05 Significance"
    If isSig Then
      .TextMatrix(12, 1) = "Significant"
    Else
      .TextMatrix(12, 1) = "Not Significant"
    End If
    .ColWidth(0) = 2000
    Dim col2Width As Long
    col2Width = .Width - .ColWidth(0) - 200
    If col2Width < 500 Then col2Width = 500
    .ColWidth(1) = col2Width
  End With
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set mGrid = Nothing
End Sub

Private Sub PicChart_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuChartPopup
End Sub

Private Sub mnuCopyChart_Click()
  Clipboard.Clear
  Clipboard.SetData PicChart.image, vbCFBitmap
End Sub

Private Sub GridStats_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  If Button = 2 Then PopupMenu mnuTablePopup
End Sub

Private Sub mnuCopyTable_Click()
  Call CopyStatsTableToClipboard
End Sub

Private Sub CopyStatsTableToClipboard()
  Dim r As Long, c As Long, LineText As String, OutText As String
  With GridStats
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

'TextExportHeight留空或填的不是合法数字时，沿用当前屏幕上PicChart的像素高度
Private Function GetExportHeightPixels() As Long
  Dim v As Double
  If ChartData.TryParseAxisValue(TextExportHeight.Text, v) And v >= 100 Then
    GetExportHeightPixels = CLng(v)
  Else
    GetExportHeightPixels = PicChart.ScaleHeight \ Screen.TwipsPerPixelY
  End If
End Function

Private Sub cmdSavePNG_Click()
  CommonDialogSave.Filter = "PNG Images (*.png)|*.png"
  CommonDialogSave.FilterIndex = 1
  CommonDialogSave.FileName = "Time Series Chart"
  CommonDialogSave.InitDir = App.Path & "\Output"
  CommonDialogSave.Flags = &H2
  CommonDialogSave.ShowSave
  If CommonDialogSave.Flags = 2 Then Exit Sub

  '宽度按屏幕上PicChart当前的宽高比例换算
  Dim exportH As Long, exportW As Long
  exportH = GetExportHeightPixels()
  exportW = CLng(exportH * (PicChart.ScaleWidth / PicChart.ScaleHeight))

  PicExport.Height = exportH * Screen.TwipsPerPixelY
  PicExport.Width = exportW * Screen.TwipsPerPixelX
  Call DrawTemporalChart(mLastValues, mLastLabels, mLastIsMissing, mLastN, PicExport)

  Call SaveChartAsPNG(PicExport, CommonDialogSave.FileName)
  FrmMain.StatusBar1.Panels(1).Text = "Chart saved"
End Sub

Private Sub cmdSaveCSV_Click()
  CommonDialogSave.Filter = "Text Files (*.csv)|*.csv"
  CommonDialogSave.FilterIndex = 1
  CommonDialogSave.FileName = "Time Series Statistics Table"
  CommonDialogSave.InitDir = App.Path & "\Output"
  CommonDialogSave.Flags = &H2
  CommonDialogSave.ShowSave
  If CommonDialogSave.Flags = 2 Then Exit Sub

  Call SaveGridAsCSV(GridStats, CommonDialogSave.FileName)
  FrmMain.StatusBar1.Panels(1).Text = "Table saved"
End Sub
