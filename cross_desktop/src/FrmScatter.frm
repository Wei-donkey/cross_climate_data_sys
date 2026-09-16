VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmScatter 
   Caption         =   "Scatter Plot"
   ClientHeight    =   9420
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   14430
   Icon            =   "FrmScatter.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   9420
   ScaleWidth      =   14430
   StartUpPosition =   2  '屏幕中心
   Begin VB.Frame Frame1 
      Caption         =   "Manual Adjustment"
      Height          =   3135
      Left            =   9840
      TabIndex        =   10
      Top             =   1680
      Width           =   4455
      Begin VB.TextBox TextCircleSize 
         Height          =   375
         Left            =   3000
         TabIndex        =   32
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox TextExportHeight 
         Height          =   375
         Left            =   600
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
      Begin VB.TextBox TextYaxisStep 
         Height          =   300
         Left            =   2520
         TabIndex        =   19
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisMax 
         Height          =   300
         Left            =   2520
         TabIndex        =   18
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisMin 
         Height          =   300
         Left            =   2520
         TabIndex        =   17
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TextXaxisStep 
         Height          =   300
         Left            =   600
         TabIndex        =   16
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisName 
         Height          =   375
         Left            =   2520
         TabIndex        =   15
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TextXaxisName 
         Height          =   375
         Left            =   600
         TabIndex        =   14
         Top             =   480
         Width           =   1815
      End
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   3720
         TabIndex        =   13
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox TextXaxisMin 
         Height          =   300
         Left            =   600
         TabIndex        =   12
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TextXaxisMax 
         Height          =   300
         Left            =   600
         TabIndex        =   11
         Top             =   1560
         Width           =   1815
      End
      Begin VB.Label Label7 
         Caption         =   "Point Size"
         Height          =   375
         Left            =   2520
         TabIndex        =   31
         Top             =   2520
         Width           =   495
      End
      Begin VB.Label LblExportHeight 
         Caption         =   "Export Image Height"
         Height          =   375
         Left            =   120
         TabIndex        =   30
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label3 
         Caption         =   "Font Size"
         Height          =   375
         Left            =   1320
         TabIndex        =   29
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   25
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   24
         Top             =   1150
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "End"
         Height          =   255
         Left            =   120
         TabIndex        =   23
         Top             =   1600
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Inter"
         Height          =   255
         Left            =   120
         TabIndex        =   22
         Top             =   2100
         Width           =   495
      End
      Begin VB.Label Label2 
         Caption         =   "Y-axis"
         Height          =   255
         Left            =   3120
         TabIndex        =   21
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label8 
         Caption         =   "X-axis"
         Height          =   255
         Left            =   1200
         TabIndex        =   20
         Top             =   240
         Width           =   615
      End
   End
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdSaveCSV 
      Caption         =   "Save Table"
      Height          =   450
      Left            =   13080
      TabIndex        =   9
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSavePNG 
      Caption         =   "Save Chart"
      Height          =   450
      Left            =   11760
      TabIndex        =   8
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CheckBox CheckEquation 
      Caption         =   "Regression Equation"
      Height          =   495
      Left            =   12120
      TabIndex        =   7
      Top             =   1080
      Width           =   2175
   End
   Begin VB.CheckBox CheckLine 
      Caption         =   "Add Regression Line"
      Height          =   255
      Left            =   9840
      TabIndex        =   6
      Top             =   1200
      Width           =   2055
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridStats 
      Height          =   3495
      Left            =   9840
      TabIndex        =   4
      Top             =   5040
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   6165
      _Version        =   393216
      Rows            =   12
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
      ScaleWidth      =   9540
      TabIndex        =   3
      Top             =   225
      Width           =   9600
   End
   Begin VB.PictureBox PicExport 
      AutoRedraw      =   -1  'True
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   8160
      Left            =   -20000
      ScaleHeight     =   8160
      ScaleWidth      =   8715
      TabIndex        =   26
      Top             =   225
      Visible         =   0   'False
      Width           =   8715
   End
   Begin VB.ComboBox ComboY 
      Height          =   300
      Left            =   12120
      Style           =   2  'Dropdown List
      TabIndex        =   2
      Top             =   720
      Width           =   2175
   End
   Begin VB.ComboBox ComboX 
      Height          =   300
      Left            =   9840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2175
   End
   Begin VB.Label LblYAxis 
      Caption         =   "Row/column for Y-axis: "
      Height          =   255
      Left            =   12120
      TabIndex        =   5
      Top             =   360
      Width           =   2145
   End
   Begin VB.Label LblXAxis 
      Caption         =   "Row/column for X-axis: "
      Height          =   255
      Left            =   9840
      TabIndex        =   0
      Top             =   360
      Width           =   2145
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
Attribute VB_Name = "FrmScatter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGrid As Object
Private mMode As String
Private mFixedRows As Long
Private mFixedCols As Long

Private mLastX() As Double
Private mLastY() As Double
Private mLastN As Long


Private Sub ComboX_Click()
  '2026-08-19 换了X轴数据列，X轴数值范围整个变了，之前手动填的min/max/step不再适用
  TextXaxisMin.Text = ""
  TextXaxisMax.Text = ""
  TextXaxisStep.Text = ""
  Call ApplyDefaultTitles
  Call RefreshChart
End Sub

Private Sub ComboY_Click()
  TextYaxisMin.Text = ""
  TextYaxisMax.Text = ""
  TextYaxisStep.Text = ""
  Call ApplyDefaultTitles
  Call RefreshChart
End Sub

Private Sub CmdRefresh_Click()
  Call RefreshChart
End Sub

Private Sub CheckLine_Click()
  Call RefreshChart
End Sub

Private Sub checkEquation_Click()
  Call RefreshChart
End Sub

Private Sub Form_Load()
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
End Sub


Public Sub InitFromSelection(ByVal SrcGrid As Object, ByVal SrcCaption As String, ByVal Mode As String)
  Set mGrid = SrcGrid
  mFixedRows = mGrid.FixedRows
  mFixedCols = mGrid.FixedCols
  mMode = Mode

  Call PopulateCombos

  Dim idx As Long
  idx = ChartData.DetermineSelectionIndex(mGrid, mMode)
  Call SelectDefaultPair(idx)

  Me.Caption = SrcCaption & " - " & ("Scatter Plot")
  mnuCopyChart.Caption = "Copy Chart"
  mnuCopyTable.Caption = "Copy Table"
  If mMode = "ROW" Then
    LblXAxis.Caption = "Select Row" & "X" & "："
    LblYAxis.Caption = "Select Row" & "Y" & "："
  Else
    LblXAxis.Caption = "Select Column" & "X" & "："
    LblYAxis.Caption = "Select Column" & "Y" & "："
  End If


  TextFontSize.Text = "12"

  TextExportHeight.Text = CStr(PicChart.ScaleHeight \ Screen.TwipsPerPixelY)

  TextCircleSize.Text = "60"
  Call ApplyDefaultTitles
  Call RefreshChart
End Sub

' 根据当前的X/Y选择重新判断默认的坐标轴标题文本，并写入可编辑的文本框
Private Sub ApplyDefaultTitles()
  TextXaxisName.Text = ComboX.Text
  TextYaxisName.Text = ComboY.Text
End Sub

Private Sub PopulateCombos()
  Dim i As Long, lbl As String
  ComboX.Clear
  ComboY.Clear
  If mMode = "ROW" Then
    For i = mFixedRows To mGrid.Rows - 1
      lbl = ChartData.BuildRowLabel(mGrid, i)
      ComboX.AddItem lbl: ComboX.ItemData(ComboX.NewIndex) = i
      ComboY.AddItem lbl: ComboY.ItemData(ComboY.NewIndex) = i
    Next i
  Else
    For i = mFixedCols To mGrid.Cols - 1
      If ChartData.IsDataColumn(mGrid, i) Then
        lbl = mGrid.TextMatrix(mFixedRows - 1, i)
        ComboX.AddItem lbl: ComboX.ItemData(ComboX.NewIndex) = i
        ComboY.AddItem lbl: ComboY.ItemData(ComboY.NewIndex) = i
      End If
    Next i
  End If
End Sub


Private Sub SelectDefaultPair(ByVal DataIdx As Long)
  Dim xPos As Long, yPos As Long, j As Long
  xPos = -1
  For j = 0 To ComboX.ListCount - 1
    If ComboX.ItemData(j) = DataIdx Then xPos = j: Exit For
  Next j
  If xPos < 0 Then xPos = 0
  If ComboX.ListCount > 0 Then ComboX.ListIndex = xPos

  yPos = xPos + 1
  If yPos >= ComboY.ListCount Then yPos = xPos - 1
  If yPos < 0 Then yPos = 0
  If ComboY.ListCount > 0 Then ComboY.ListIndex = yPos
End Sub


Private Sub RefreshChart()
  If mGrid Is Nothing Then Exit Sub
  If ComboX.ListIndex < 0 Or ComboY.ListIndex < 0 Then Exit Sub

  Dim Idx1 As Long, Idx2 As Long
  Idx1 = ComboX.ItemData(ComboX.ListIndex)
  Idx2 = ComboY.ItemData(ComboY.ListIndex)

  Dim X() As Double, Y() As Double, N As Long
  Call ChartData.ExtractPairedSeries(mGrid, (mMode = "ROW"), Idx1, Idx2, X, Y, N)

  Call DrawScatter(X, Y, N, TextXaxisName.Text, TextYaxisName.Text)
  Call FillStatsTable(X, Y, N)

  mLastX = X
  mLastY = Y
  mLastN = N
End Sub


Private Sub DrawScatter(ByRef X() As Double, ByRef Y() As Double, ByVal N As Long, ByVal XTitle As String, ByVal YTitle As String, Optional ByVal TargetPic As PictureBox = Nothing)
  '2026-08-18 TargetPic必须声明成具体的PictureBox类型，不能是泛型Object——Line/Circle
  Dim DrawPic As PictureBox
  If TargetPic Is Nothing Then
    Set DrawPic = Me.PicChart
  Else
    Set DrawPic = TargetPic
  End If

  DrawPic.Cls
  ' 这里的Font.Size决定了图表上每一处文字元素的大小——刻度标签、旋转的Y轴标题
  Dim chartFontSize As Single, scaleFactor As Single
  chartFontSize = ChartData.GetValidFontSize(TextFontSize.Text, 12)
  DrawPic.Font.Name = "宋体"
  DrawPic.Font.Size = chartFontSize
  scaleFactor = chartFontSize / 12
  If N <= 0 Then
    DrawPic.CurrentX = 100: DrawPic.CurrentY = 100
    DrawPic.Print "NoneYes效数据可供绘图"
    Exit Sub
  End If

  Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double, i As Long
  xMin = X(0): xMax = X(0): yMin = Y(0): yMax = Y(0)
  For i = 1 To N - 1
    If X(i) < xMin Then xMin = X(i)
    If X(i) > xMax Then xMax = X(i)
    If Y(i) < yMin Then yMin = Y(i)
    If Y(i) > yMax Then yMax = Y(i)
  Next i

  Dim xPad As Double, yPad As Double
  If xMax = xMin Then
    If Abs(xMax) < 1 Then xPad = 1 Else xPad = Abs(xMax) * 0.1
  Else
    xPad = (xMax - xMin) * 0.1
  End If
  If yMax = yMin Then
    If Abs(yMax) < 1 Then yPad = 1 Else yPad = Abs(yMax) * 0.1
  Else
    yPad = (yMax - yMin) * 0.1
  End If

  Dim axisXMin As Double, axisXMax As Double, axisYMin As Double, axisYMax As Double
  axisXMin = xMin - xPad: axisXMax = xMax + xPad
  axisYMin = yMin - yPad: axisYMax = yMax + yPad


  Dim userAxisVal As Double
  If ChartData.TryParseAxisValue(TextXaxisMin.Text, userAxisVal) Then axisXMin = userAxisVal
  If ChartData.TryParseAxisValue(TextXaxisMax.Text, userAxisVal) Then axisXMax = userAxisVal
  If ChartData.TryParseAxisValue(TextYaxisMin.Text, userAxisVal) Then axisYMin = userAxisVal
  If ChartData.TryParseAxisValue(TextYaxisMax.Text, userAxisVal) Then axisYMax = userAxisVal
  If axisXMax <= axisXMin Then axisXMax = axisXMin + 1
  If axisYMax <= axisYMin Then axisYMax = axisYMin + 1

  Dim leftM As Single, rightM As Single, topM As Single, botM As Single
  leftM = 1100 * scaleFactor: rightM = 300 * scaleFactor: topM = 200 * scaleFactor: botM = 700 * scaleFactor

  Dim availW As Single, availH As Single, plotSize As Single
  availW = DrawPic.ScaleWidth - leftM - rightM
  availH = DrawPic.ScaleHeight - topM - botM
  plotSize = availW
  If availH < plotSize Then plotSize = availH
  If plotSize < 10 Then plotSize = 10

  DrawPic.DrawWidth = 1
  DrawPic.Line (leftM, topM)-(leftM, topM + plotSize), vbBlack
  DrawPic.Line (leftM, topM + plotSize)-(leftM + plotSize, topM + plotSize), vbBlack
  DrawPic.Line (leftM, topM)-(leftM + plotSize, topM), vbBlack
  DrawPic.Line (leftM + plotSize, topM)-(leftM + plotSize, topM + plotSize), vbBlack

  Dim dummyMin As Double, dummyMax As Double, tickStepX As Double, tickStepY As Double, commonStep As Double
  Call ChartData.ComputeNiceTicks(axisXMin, axisXMax, 10, dummyMin, dummyMax, tickStepX)
  Call ChartData.ComputeNiceTicks(axisYMin, axisYMax, 10, dummyMin, dummyMax, tickStepY)
  commonStep = tickStepX
  If tickStepY > commonStep Then commonStep = tickStepY

  Dim stepX As Double, stepY As Double
  stepX = commonStep
  stepY = commonStep
  If ChartData.TryParseAxisValue(TextXaxisStep.Text, userAxisVal) And userAxisVal > 0 Then stepX = userAxisVal
  If ChartData.TryParseAxisValue(TextYaxisStep.Text, userAxisVal) And userAxisVal > 0 Then stepY = userAxisVal

  TextXaxisMin.Text = Format(axisXMin, "0.0")
  TextXaxisMax.Text = Format(axisXMax, "0.0")
  TextXaxisStep.Text = Format(stepX, "0.0")
  TextYaxisMin.Text = Format(axisYMin, "0.0")
  TextYaxisMax.Text = Format(axisYMax, "0.0")
  TextYaxisStep.Text = Format(stepY, "0.0")

  Call ChartData.DrawYAxis(DrawPic, leftM, topM, plotSize, axisYMin, axisYMax, YTitle, stepY)
  Call ChartData.DrawXAxisNumeric(DrawPic, leftM, topM + plotSize, plotSize, axisXMin, axisXMax, stepX)

  Dim circleSize As Double
  circleSize = 80
  If ChartData.TryParseAxisValue(TextCircleSize.Text, userAxisVal) And userAxisVal > 0 Then circleSize = userAxisVal

  Dim px As Single, py As Single
  DrawPic.FillStyle = 1
  DrawPic.FillColor = RGB(255, 255, 255)
  DrawPic.DrawWidth = 2
  For i = 0 To N - 1
    px = leftM + CSng((X(i) - axisXMin) / (axisXMax - axisXMin) * plotSize)
    py = topM + plotSize - CSng((Y(i) - axisYMin) / (axisYMax - axisYMin) * plotSize)
    DrawPic.Circle (px, py), circleSize, RGB(70, 130, 180)
  Next i
  DrawPic.DrawWidth = 1

  Dim rSlope As Double, rIntercept As Double, rR2 As Double, rCorr As Double, rValid As Boolean
  Call ChartData.ComputeLinearRegressionXY(X, Y, N, rSlope, rIntercept, rR2, rCorr, rValid)
  If rValid And CheckLine.Value = 1 Then
    Dim ly1 As Double, ly2 As Double, lpy1 As Single, lpy2 As Single
    ly1 = rSlope * axisXMin + rIntercept
    ly2 = rSlope * axisXMax + rIntercept
    lpy1 = topM + plotSize - CSng((ly1 - axisYMin) / (axisYMax - axisYMin) * plotSize)
    lpy2 = topM + plotSize - CSng((ly2 - axisYMin) / (axisYMax - axisYMin) * plotSize)
    If lpy1 < topM Then lpy1 = topM
    If lpy1 > topM + plotSize Then lpy1 = topM + plotSize
    If lpy2 < topM Then lpy2 = topM
    If lpy2 > topM + plotSize Then lpy2 = topM + plotSize
    DrawPic.DrawStyle = 1
    DrawPic.DrawWidth = 2
    DrawPic.Line (leftM, lpy1)-(leftM + plotSize, lpy2), vbRed
    DrawPic.DrawStyle = 0
    DrawPic.DrawWidth = 1
  End If

  ' 与回归线本身相互独立——即使不显示可视化的趋势线，方程也可以显示，反之亦然。
  If rValid And CheckEquation.Value = 1 Then
    Dim eqText As String, signStr As String
    signStr = "+"
    If rIntercept < 0 Then signStr = "-"
    eqText = "y = " & Format(rSlope, "0.000") & "x " & signStr & " " & Format(Abs(rIntercept), "0.000")

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
    DrawPic.Print " = " & Format(rR2, "0.000")
  End If

  If Len(XTitle) > 0 Then
    DrawPic.CurrentX = leftM + (plotSize - DrawPic.TextWidth(XTitle)) / 2
    DrawPic.CurrentY = topM + plotSize + 420 * scaleFactor
    DrawPic.Print XTitle
  End If
End Sub

Private Sub FillStatsTable(ByRef X() As Double, ByRef Y() As Double, ByVal N As Long)
  Dim xMean As Double, xMax As Double, xMin As Double, xSum As Double, xStd As Double
  Dim yMean As Double, yMax As Double, yMin As Double, ySum As Double, yStd As Double
  Call ChartData.ComputeStats(X, N, xMean, xMax, xMin, xSum, xStd)
  Call ChartData.ComputeStats(Y, N, yMean, yMax, yMin, ySum, yStd)

  Dim rSlope As Double, rIntercept As Double, rR2 As Double, rCorr As Double, rValid As Boolean
  Call ChartData.ComputeLinearRegressionXY(X, Y, N, rSlope, rIntercept, rR2, rCorr, rValid)

  Dim FStat As Double, PValue As Double, isSig As Boolean
  If rValid Then Call ChartData.ComputeTrendSignificance(N, rR2, FStat, PValue, isSig)

  With GridStats
    .Rows = 12: .Cols = 2
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .TextMatrix(0, 0) = "Sample Size": .TextMatrix(0, 1) = CStr(N)
    .TextMatrix(1, 0) = "X" & " Average": .TextMatrix(1, 1) = Format(xMean, "0.0")
    .TextMatrix(2, 0) = "X" & " Std Deviation": .TextMatrix(2, 1) = Format(xStd, "0.0")
    .TextMatrix(3, 0) = "Y" & " Average": .TextMatrix(3, 1) = Format(yMean, "0.0")
    .TextMatrix(4, 0) = "Y" & " Std Deviation": .TextMatrix(4, 1) = Format(yStd, "0.0")

    If rValid Then
      .TextMatrix(5, 0) = "Correlation Coeff": .TextMatrix(5, 1) = Format(rCorr, "0.000")
      .TextMatrix(6, 0) = "Determination Coeff": .TextMatrix(6, 1) = Format(rR2, "0.000")
      .TextMatrix(7, 0) = "Regression Slope": .TextMatrix(7, 1) = Format(rSlope, "0.000")
      .TextMatrix(8, 0) = "Regression Intercept": .TextMatrix(8, 1) = Format(rIntercept, "0.000")
      .TextMatrix(9, 0) = "F Value": .TextMatrix(9, 1) = Format(FStat, "0.00")
      .TextMatrix(10, 0) = "P Value": .TextMatrix(10, 1) = Format(PValue, "0.0000")
      .TextMatrix(11, 0) = "0.05 Significance"
      If isSig Then
        .TextMatrix(11, 1) = "Significant"
      Else
        .TextMatrix(11, 1) = "Not Significant"
      End If
    Else
      .TextMatrix(5, 0) = "Correlation Coeff": .TextMatrix(5, 1) = "--"
      .TextMatrix(6, 0) = "Determination Coeff": .TextMatrix(6, 1) = "--"
      .TextMatrix(7, 0) = "Regression Slope": .TextMatrix(7, 1) = "--"
      .TextMatrix(8, 0) = "Regression Intercept": .TextMatrix(8, 1) = "--"
      .TextMatrix(9, 0) = "F Value": .TextMatrix(9, 1) = "--"
      .TextMatrix(10, 0) = "P Value": .TextMatrix(10, 1) = "--"
      .TextMatrix(11, 0) = "0.05 Significance": .TextMatrix(11, 1) = "--"
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
  CommonDialogSave.FileName = "Scatter Plot"
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
  Call DrawScatter(mLastX, mLastY, mLastN, TextXaxisName.Text, TextYaxisName.Text, PicExport)

  Call SaveChartAsPNG(PicExport, CommonDialogSave.FileName)
  FrmMain.StatusBar1.Panels(1).Text = "Chart saved"
End Sub

Private Sub cmdSaveCSV_Click()
  CommonDialogSave.Filter = "Text Files (*.csv)|*.csv"
  CommonDialogSave.FilterIndex = 1
  CommonDialogSave.FileName = "Scatter Plot Statistics Table"
  CommonDialogSave.InitDir = App.Path & "\Output"
  CommonDialogSave.Flags = &H2
  CommonDialogSave.ShowSave
  If CommonDialogSave.Flags = 2 Then Exit Sub

  Call SaveGridAsCSV(GridStats, CommonDialogSave.FileName)
  FrmMain.StatusBar1.Panels(1).Text = "Table saved"
End Sub
