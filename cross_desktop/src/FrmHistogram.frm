VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmHistogram 
   Caption         =   "Histogram"
   ClientHeight    =   9420
   ClientLeft      =   195
   ClientTop       =   405
   ClientWidth     =   14430
   Icon            =   "FrmHistogram.frx":0000
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
      TabIndex        =   6
      Top             =   1200
      Width           =   4455
      Begin VB.CommandButton CmdRefresh 
         Caption         =   "Refresh"
         Height          =   375
         Left            =   3360
         TabIndex        =   24
         Top             =   2520
         Width           =   975
      End
      Begin VB.TextBox TextFontSize 
         Height          =   375
         Left            =   1800
         TabIndex        =   23
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox TextExportHeight 
         Height          =   375
         Left            =   600
         TabIndex        =   22
         Top             =   2520
         Width           =   615
      End
      Begin VB.TextBox TextXaxisMax 
         Height          =   300
         Left            =   600
         TabIndex        =   20
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox TextXaxisMin 
         Height          =   300
         Left            =   600
         TabIndex        =   19
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TextXaxisName 
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisName 
         Height          =   375
         Left            =   2520
         TabIndex        =   11
         Top             =   480
         Width           =   1815
      End
      Begin VB.TextBox TextXaxisStep 
         Height          =   300
         Left            =   600
         TabIndex        =   10
         Top             =   2040
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisMin 
         Height          =   300
         Left            =   2520
         TabIndex        =   9
         Top             =   1080
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisMax 
         Height          =   300
         Left            =   2520
         TabIndex        =   8
         Top             =   1560
         Width           =   1815
      End
      Begin VB.TextBox TextYaxisStep 
         Height          =   300
         Left            =   2520
         TabIndex        =   7
         Top             =   2040
         Width           =   1815
      End
      Begin VB.Label Label3 
         Caption         =   "Font Size"
         Height          =   375
         Left            =   1320
         TabIndex        =   26
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label LblExportHeight 
         Caption         =   "Export Image Height"
         Height          =   375
         Left            =   120
         TabIndex        =   25
         Top             =   2520
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Name"
         Height          =   255
         Left            =   120
         TabIndex        =   18
         Top             =   600
         Width           =   495
      End
      Begin VB.Label Label4 
         Caption         =   "Start"
         Height          =   255
         Left            =   120
         TabIndex        =   17
         Top             =   1150
         Width           =   495
      End
      Begin VB.Label Label5 
         Caption         =   "End"
         Height          =   255
         Left            =   120
         TabIndex        =   16
         Top             =   1600
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "Inter"
         Height          =   255
         Left            =   120
         TabIndex        =   15
         Top             =   2100
         Width           =   495
      End
      Begin VB.Label Label8 
         Caption         =   "X-axis"
         Height          =   255
         Left            =   1200
         TabIndex        =   14
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Y-axis"
         Height          =   255
         Left            =   3120
         TabIndex        =   13
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
      TabIndex        =   5
      Top             =   8760
      Width           =   1215
   End
   Begin VB.CommandButton cmdSavePNG 
      Caption         =   "Save Chart"
      Height          =   450
      Left            =   11760
      TabIndex        =   4
      Top             =   8760
      Width           =   1215
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid GridStats 
      Height          =   4110
      Left            =   9840
      TabIndex        =   3
      Top             =   4500
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   7250
      _Version        =   393216
      Rows            =   6
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
      TabIndex        =   2
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
      TabIndex        =   21
      Top             =   225
      Visible         =   0   'False
      Width           =   8715
   End
   Begin VB.ComboBox ComboSeries 
      Height          =   300
      Left            =   9840
      Style           =   2  'Dropdown List
      TabIndex        =   1
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label LblMode 
      Caption         =   "Row/column for the data: "
      Height          =   255
      Left            =   9840
      TabIndex        =   0
      Top             =   360
      Width           =   2655
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
Attribute VB_Name = "FrmHistogram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mGrid As Object
Private mMode As String
Private mFixedRows As Long
'2026-08-17 缓存最近一次刷新Chart用来画图的数据，供cmd保存PNG_Click按导出尺寸重新画图
Private mLastValues() As Double
Private mLastN As Long
Private mFixedCols As Long
Private mXAxisTitle As String


Private Sub ComboSeries_Click()
  '2026-08-19 换了行/站点，X轴数值分布和Y轴频数都变了，之前手动填的min/max/step
  '不再适用于新数据，清空让它们按新数据重新自动计算
  TextXaxisMin.Text = ""
  TextXaxisMax.Text = ""
  TextXaxisStep.Text = ""
  TextYaxisMin.Text = ""
  TextYaxisMax.Text = ""
  TextYaxisStep.Text = ""
  Call RefreshChart
End Sub

Private Sub CmdRefresh_Click()
  Call RefreshChart
End Sub

Private Sub Form_Load()
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
End Sub

' 在本实例创建后由ChartData.Show直方图ForActiveForm调用一次：
' 把这个窗口跟源表格关联起来，判断用户选的是行还是列，并画出初始图表
Public Sub InitFromSelection(ByVal SrcGrid As Object, ByVal SrcCaption As String, ByVal XAxisTitle As String, ByVal Mode As String)
  Set mGrid = SrcGrid
  mFixedRows = mGrid.FixedRows
  mFixedCols = mGrid.FixedCols
  mMode = Mode
  mXAxisTitle = XAxisTitle

  Dim idx As Long
  idx = ChartData.DetermineSelectionIndex(mGrid, mMode)

  Call PopulateComboSeries
  Call SelectComboIndex(idx)

  Me.Caption = SrcCaption & " - " & ("Histogram")
  mnuCopyChart.Caption = "Copy Chart"
  mnuCopyTable.Caption = "Copy Table"
  If mMode = "ROW" Then
    LblMode.Caption = "Select Row: "
  Else
    LblMode.Caption = "Select Column: "
  End If
  TextFontSize.Text = "12"
  '2026-08-18 窗口刚打开时也回填一次，用PicChart当前的像素高度当默认值
  TextExportHeight.Text = CStr(PicChart.ScaleHeight \ Screen.TwipsPerPixelY)
  Call ApplyDefaultTitles
  Call RefreshChart
End Sub

' 根据当前选择重新生成默认的坐标轴标题文本，并写入可编辑的文本框
Private Sub ApplyDefaultTitles()
  TextXaxisName.Text = mXAxisTitle
  TextYaxisName.Text = "Frequency"
End Sub

Private Sub PopulateComboSeries()
  Dim i As Long, pos As Long
  ComboSeries.Clear
  If mMode = "ROW" Then
    For i = mFixedRows To mGrid.Rows - 1
      ComboSeries.AddItem ChartData.BuildRowLabel(mGrid, i)
      ComboSeries.ItemData(ComboSeries.NewIndex) = i
    Next i
  Else
    For i = mFixedCols To mGrid.Cols - 1
      ' 跳过非数据列（比如乡/县/市这类行政区划标签列）
      If ChartData.IsDataColumn(mGrid, i) Then
        ComboSeries.AddItem mGrid.TextMatrix(mFixedRows - 1, i)
        ComboSeries.ItemData(ComboSeries.NewIndex) = i
      End If
    Next i
  End If
End Sub

Private Sub SelectComboIndex(ByVal DataIdx As Long)
  Dim j As Long
  For j = 0 To ComboSeries.ListCount - 1
    If ComboSeries.ItemData(j) = DataIdx Then
      ComboSeries.ListIndex = j
      Exit Sub
    End If
  Next j
  ' 原本点击的那一行/列被过滤掉了（非数据列），退回到第一个可用项
  If ComboSeries.ListCount > 0 Then ComboSeries.ListIndex = 0
End Sub

Private Sub RefreshChart()
  If mGrid Is Nothing Then Exit Sub
  If ComboSeries.ListIndex < 0 Then Exit Sub

  Dim DataIdx As Long
  DataIdx = ComboSeries.ItemData(ComboSeries.ListIndex)

  Dim Labels() As String, Values() As Double, N As Long
  Call ChartData.ExtractSeries(mGrid, (mMode = "ROW"), DataIdx, Labels, Values, N)

  Call DrawHistogram(Values, N)
  Call FillStatsTable(Values, N)

  mLastValues = Values
  mLastN = N
End Sub


Private Sub DrawHistogram(ByRef Values() As Double, ByVal N As Long, Optional ByVal TargetPic As PictureBox = Nothing)
  '2026-08-18 TargetPic必须声明成具体的PictureBox类型
  Dim DrawPic As PictureBox
  If TargetPic Is Nothing Then
    Set DrawPic = Me.PicChart
  Else
    Set DrawPic = TargetPic
  End If

  DrawPic.Cls
  'Font.Size决定了图表上每一处文字元素的大小——刻度标签、旋转的Y轴标题
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

  Dim dMin As Double, dMax As Double, i As Long
  dMin = Values(0): dMax = Values(0)
  For i = 1 To N - 1
    If Values(i) < dMin Then dMin = Values(i)
    If Values(i) > dMax Then dMax = Values(i)
  Next i

  Dim targetK As Long
  If N < 2 Then
    targetK = 1
  Else
    targetK = CLng(1 + 3.32193 * Log(N) / Log(10#))
    If targetK < 5 Then targetK = 5
    If targetK > 15 Then targetK = 15
  End If

  ' 源数据只保留1位小数（比如气温精确到0.1）
  Dim binWidth As Double, k As Long
  dMin = Int(dMin * 10 + 0.0000001) / 10
  If dMax = dMin Then
    binWidth = 0.1
  Else
    binWidth = (dMax - dMin) / targetK
    binWidth = Int(binWidth * 10 + 0.9999) / 10  ' 向上取整到0.1
    If binWidth < 0.1 Then binWidth = 0.1
  End If

  Dim userAxisVal As Double
  If ChartData.TryParseAxisValue(TextXaxisMin.Text, userAxisVal) Then dMin = userAxisVal
  If ChartData.TryParseAxisValue(TextXaxisMax.Text, userAxisVal) Then dMax = userAxisVal
  If ChartData.TryParseAxisValue(TextXaxisStep.Text, userAxisVal) And userAxisVal > 0 Then binWidth = userAxisVal
  If dMax <= dMin Then dMax = dMin + binWidth  ' 防止用户把最小Value填得比最大Value还大

  k = Int((dMax - dMin) / binWidth + 0.0000001) + 1
  If k < 1 Then k = 1

  TextXaxisMin.Text = Format(dMin, "0.0")
  TextXaxisMax.Text = Format(dMax, "0.0")
  TextXaxisStep.Text = Format(binWidth, "0.0")

  Dim Freq() As Long
  ReDim Freq(0 To k - 1)
  Dim b As Long
  For i = 0 To N - 1
    b = Int((Values(i) - dMin) / binWidth + 0.0000001)
    If b >= k Then b = k - 1
    If b < 0 Then b = 0
    Freq(b) = Freq(b) + 1
  Next i

  Dim maxFreq As Long
  maxFreq = 0
  For i = 0 To k - 1
    If Freq(i) > maxFreq Then maxFreq = Freq(i)
  Next i
  If maxFreq = 0 Then maxFreq = 1

  ' 柱子高度按各柱频次占maxFreq的比例画
  Dim maxFreqPct As Double
  maxFreqPct = CDbl(maxFreq) / CDbl(N) * 100

  Dim leftM As Single, rightM As Single, topM As Single, botM As Single
  leftM = 1100 * scaleFactor: rightM = 300 * scaleFactor: topM = 200 * scaleFactor: botM = 700 * scaleFactor

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

  Dim axisYMin As Double, axisYMax As Double, yForcedStep As Double
  axisYMin = 0
  axisYMax = maxFreqPct
  If ChartData.TryParseAxisValue(TextYaxisMin.Text, userAxisVal) Then axisYMin = userAxisVal
  If ChartData.TryParseAxisValue(TextYaxisMax.Text, userAxisVal) Then axisYMax = userAxisVal
  If axisYMax <= axisYMin Then axisYMax = axisYMin + 1
  yForcedStep = 0
  If ChartData.TryParseAxisValue(TextYaxisStep.Text, userAxisVal) And userAxisVal > 0 Then yForcedStep = userAxisVal

  Dim dispYMin As Double, dispYMax As Double, dispYStep As Double
  If yForcedStep > 0 Then
    dispYStep = yForcedStep
  Else
    Call ChartData.ComputeNiceTicks(axisYMin, axisYMax, 10, dispYMin, dispYMax, dispYStep)
  End If
  TextYaxisMin.Text = Format(axisYMin, "0.0")
  TextYaxisMax.Text = Format(axisYMax, "0.0")
  TextYaxisStep.Text = Format(dispYStep, "0.0")

  Call ChartData.DrawYAxis(DrawPic, leftM, topM, plotH, axisYMin, axisYMax, TextYaxisName.Text, yForcedStep)

  Dim barW As Single
  barW = plotW / k

  DrawPic.FillStyle = 0  ' 0 = vbFSSolid
  DrawPic.FillColor = RGB(70, 130, 180)

  Dim x1 As Single, barH As Single, y1 As Single, barPct As Double
  For i = 0 To k - 1
    x1 = leftM + i * barW
    barPct = CDbl(Freq(i)) / CDbl(N) * 100
    barH = plotH * (barPct - axisYMin) / (axisYMax - axisYMin)
    If barH < 0 Then barH = 0
    If barH > plotH Then barH = plotH
    y1 = topM + plotH - barH
    DrawPic.Line (x1 + 20, y1)-(x1 + barW - 20, topM + plotH), RGB(70, 130, 180), BF

    If k <= 12 Or i Mod 2 = 0 Then
      DrawPic.Line (x1, topM + plotH)-(x1, topM + plotH + 60), vbBlack
      DrawPic.CurrentX = x1
      DrawPic.CurrentY = topM + plotH + 80
      DrawPic.Print Format(dMin + i * binWidth, "0.0")
    End If
  Next i

  If Len(TextXaxisName.Text) > 0 Then
    DrawPic.CurrentX = leftM + (plotW - DrawPic.TextWidth(TextXaxisName.Text)) / 2
    DrawPic.CurrentY = topM + plotH + 420 * scaleFactor
    DrawPic.Print TextXaxisName.Text
  End If
End Sub

Private Sub FillStatsTable(ByRef Values() As Double, ByVal N As Long)
  Dim dMean As Double, dMax As Double, dMin As Double, dSum As Double, dStd As Double
  Call ChartData.ComputeStats(Values, N, dMean, dMax, dMin, dSum, dStd)

  Dim dMedian As Double, dMode As Double, dQ1 As Double, dQ3 As Double
  Dim dP90 As Double, dP95 As Double, dP99 As Double
  Call ChartData.ComputePercentileStats(Values, N, dMedian, dMode, dQ1, dQ3, dP90, dP95, dP99)

  With GridStats
    .Rows = 13: .Cols = 2
    .ColAlignment(0) = flexAlignLeftCenter
    .ColAlignment(1) = flexAlignLeftCenter
    .TextMatrix(0, 0) = "Sample Size": .TextMatrix(0, 1) = CStr(N)
    .TextMatrix(1, 0) = "Average": .TextMatrix(1, 1) = Format(dMean, "0.0")
    .TextMatrix(2, 0) = "Max": .TextMatrix(2, 1) = Format(dMax, "0.0")
    .TextMatrix(3, 0) = "Min": .TextMatrix(3, 1) = Format(dMin, "0.0")
    .TextMatrix(4, 0) = "Sum": .TextMatrix(4, 1) = Format(dSum, "0.0")
    .TextMatrix(5, 0) = "Std Dev": .TextMatrix(5, 1) = Format(dStd, "0.0")
    .TextMatrix(6, 0) = "Mode": .TextMatrix(6, 1) = Format(dMode, "0.0")
    .TextMatrix(7, 0) = "Quartile" & "1": .TextMatrix(7, 1) = Format(dQ1, "0.0")
    .TextMatrix(8, 0) = "Median": .TextMatrix(8, 1) = Format(dMedian, "0.0")
    .TextMatrix(9, 0) = "Quartile" & "2": .TextMatrix(9, 1) = Format(dQ3, "0.0")
    .TextMatrix(10, 0) = "90" & "Percentile": .TextMatrix(10, 1) = Format(dP90, "0.0")
    .TextMatrix(11, 0) = "95" & "Percentile": .TextMatrix(11, 1) = Format(dP95, "0.0")
    .TextMatrix(12, 0) = "99" & "Percentile": .TextMatrix(12, 1) = Format(dP99, "0.0")
    .ColWidth(0) = 2000
    ' 用第2列填满表格自身剩余的宽度，而不是在数值右边留出一段空白
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
  CommonDialogSave.FileName = "Histogram"
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
  Call DrawHistogram(mLastValues, mLastN, PicExport)

  Call SaveChartAsPNG(PicExport, CommonDialogSave.FileName)
  FrmMain.StatusBar1.Panels(1).Text = "Chart saved"
End Sub

Private Sub cmdSaveCSV_Click()
  CommonDialogSave.Filter = "Text Files (*.csv)|*.csv"
  CommonDialogSave.FilterIndex = 1
  CommonDialogSave.FileName = "Histogram Statistics Table"
  CommonDialogSave.InitDir = App.Path & "\Output"
  CommonDialogSave.Flags = &H2
  CommonDialogSave.ShowSave
  If CommonDialogSave.Flags = 2 Then Exit Sub

  Call SaveGridAsCSV(GridStats, CommonDialogSave.FileName)
  FrmMain.StatusBar1.Panels(1).Text = "Table saved"
End Sub
