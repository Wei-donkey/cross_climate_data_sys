Attribute VB_Name = "ChartData"
Option Explicit

' VB6的PictureBox.Font是一个没有Orientation属性的StdFont对象——旋转文字
Private Type LOGFONT
    lfHeight As Long
    lfWidth As Long
    lfEscapement As Long
    lfOrientation As Long
    lfWeight As Long
    lfItalic As Byte
    lfUnderline As Byte
    lfStrikeOut As Byte
    lfCharSet As Byte
    lfOutPrecision As Byte
    lfClipPrecision As Byte
    lfQuality As Byte
    lfPitchAndFamily As Byte
    lfFaceName As String * 32
End Type

Private Declare Function CreateFontIndirect Lib "gdi32" Alias "CreateFontIndirectA" (lpLogFont As LOGFONT) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function TextOut Lib "gdi32" Alias "TextOutA" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As String, ByVal nCount As Long) As Long
Private Declare Function SetBkMode Lib "gdi32" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Declare Function GetDeviceCaps Lib "gdi32" (ByVal hdc As Long, ByVal nIndex As Long) As Long
Private Declare Function MulDiv Lib "kernel32" (ByVal nNumber As Long, ByVal nNumerator As Long, ByVal nDenominator As Long) As Long
Private Const TRANSPARENT As Long = 1
Private Const LOGPIXELSY As Long = 90
Private Declare Function SetTextColor Lib "gdi32" (ByVal hdc As Long, ByVal crColor As Long) As Long
Private Const DEFAULT_CHARSET As Long = 1
Private Const GB2312_CHARSET As Long = 134

'2026-08-19 用于让旋转的Y轴标题跟横排文字
Private Const CLEARTYPE_QUALITY As Long = 5
Private Declare Function TextOutW Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal lpString As Long, ByVal nCount As Long) As Long
Private Const OUT_TT_PRECIS As Long = 4

Private Const MISSING_VALUE As Double = -9999

' 供绘图窗体使用，使其在关闭之前始终置顶于其他所有窗口之上
Public Const HWND_TOPMOST As Long = -1
Public Const SWP_NOMOVE As Long = &H2
Public Const SWP_NOSIZE As Long = &H1


' 支持绘图的10个数据模块窗体类名（VB_Name）
Public Function IsSupportedDataForm(ByVal Frm As Object) As Boolean
  Select Case TypeName(Frm)
    Case "FrmMeteoHour", "FrmMeteoDay", "FrmMeteoTen", "FrmMeteoMon", "FrmMeteoQtr", "FrmMeteoYer", _
         "FrmMeteoPeriodSin", "FrmMeteoPeriod", "FrmMeteoPeriodYer", "FrmMeteoColdAir", "FrmImportData", _
         "FrmMeteoSeasons", "FrmMeteoPeriodExt"
      IsSupportedDataForm = True
    Case Else
      IsSupportedDataForm = False
  End Select
End Function

' 判断当前网格选区是整行选中还是整列选中
Public Function DetermineSelectionMode(ByVal Grid As Object) As String
  Dim colIsFull As Boolean, rowIsFull As Boolean
  colIsFull = (Grid.Col <= Grid.FixedCols And Grid.ColSel >= Grid.Cols - 1)
  rowIsFull = (Grid.Row <= Grid.FixedRows And Grid.RowSel >= Grid.Rows - 1)
  If rowIsFull And Not colIsFull Then
    DetermineSelectionMode = "COL"
  Else
    DetermineSelectionMode = "ROW"
  End If
End Function

Public Function DetermineSelectionIndex(ByVal Grid As Object, ByVal Mode As String) As Long
  If Mode = "ROW" Then
    DetermineSelectionIndex = Grid.Row
  Else
    DetermineSelectionIndex = Grid.Col
  End If
End Function


Public Function IsRowSelectionUnsupported(ByVal Frm As Object) As Boolean
  Select Case TypeName(Frm)
    Case "FrmMeteoPeriodSin", "FrmMeteoPeriod", "FrmMeteoColdAir", "FrmMeteoSeasons", "FrmMeteoPeriodExt"
      IsRowSelectionUnsupported = True
    Case Else
      IsRowSelectionUnsupported = False
  End Select
End Function

' 从Grid的Index处取出一行（UseRow=True）或一列（UseRow=False）数值数据，跳过非数字单元格以及-9999缺测值代码。
Public Sub ExtractSeries(ByVal Grid As Object, ByVal UseRow As Boolean, ByVal Index As Long, _
                          ByRef Labels() As String, ByRef Values() As Double, ByRef N As Long)
  Dim total As Long, i As Long, v As Double, s As String, HeaderText As String
  Dim colFrom As Long, colTo As Long
  N = 0

  colFrom = Grid.FixedCols
  colTo = Grid.Cols - 1


  If UseRow And TypeName(Grid.Parent) = "FrmImportData" Then
    Dim selStart As Long, selEnd As Long
    selStart = Grid.Col
    selEnd = Grid.ColSel
    If selEnd < selStart Then
      Dim tmpCol As Long
      tmpCol = selStart: selStart = selEnd: selEnd = tmpCol
    End If
    If selStart < colFrom Then selStart = colFrom
    If selEnd > colTo Then selEnd = colTo
    If selEnd >= selStart Then
      colFrom = selStart
      colTo = selEnd
    End If
  End If

  If UseRow Then
    total = colTo - colFrom + 1
  Else
    total = Grid.Rows - Grid.FixedRows
  End If
  If total < 1 Then total = 1
  ReDim Labels(0 To total - 1)
  ReDim Values(0 To total - 1)

  If UseRow Then
    For i = colFrom To colTo

      If Grid.FixedRows >= 1 Then
        HeaderText = Grid.TextMatrix(Grid.FixedRows - 1, i)
      Else
        HeaderText = ""
      End If
      If Not (IsRowSummaryColumn(HeaderText) Or Trim(HeaderText) = "") Then
        s = Trim(Grid.TextMatrix(Index, i))
        If IsNumeric(s) Then
          v = CDbl(s)
          If v <> MISSING_VALUE Then
            Labels(N) = Grid.TextMatrix(Grid.FixedRows - 1, i)
            Values(N) = v
            N = N + 1
          End If
        End If
      End If
    Next i
  Else
    For i = Grid.FixedRows To Grid.Rows - 1
      s = Trim(Grid.TextMatrix(i, Index))
      If IsNumeric(s) Then
        v = CDbl(s)
        If v <> MISSING_VALUE Then
          Labels(N) = Grid.TextMatrix(i, 0)
          Values(N) = v
          N = N + 1
        End If
      End If
    Next i
  End If

  If N > 0 Then
    ReDim Preserve Labels(0 To N - 1)
    ReDim Preserve Values(0 To N - 1)
  Else
    ReDim Labels(0 To 0)
    ReDim Values(0 To 0)
  End If
End Sub

' 供时序图使用的行模式序列提取
Public Sub ExtractTemporalSeries(ByVal Grid As Object, ByVal Index As Long, _
                                 ByRef Labels() As String, ByRef Values() As Double, _
                                 ByRef IsMissing() As Boolean, ByRef N As Long)
  Dim total As Long, i As Long, s As String, HeaderText As String, v As Double
  Dim colFrom As Long, colTo As Long
  N = 0
  colFrom = Grid.FixedCols
  colTo = Grid.Cols - 1
  total = colTo - colFrom + 1
  If total < 1 Then total = 1
  ReDim Labels(0 To total - 1)
  ReDim Values(0 To total - 1)
  ReDim IsMissing(0 To total - 1)

  For i = colFrom To colTo
    HeaderText = Grid.TextMatrix(Grid.FixedRows - 1, i)
    If Not (IsRowSummaryColumn(HeaderText) Or Trim(HeaderText) = "") Then
      Labels(N) = HeaderText
      s = Trim(Grid.TextMatrix(Index, i))
      If IsNumeric(s) Then
        v = CDbl(s)
        If v <> MISSING_VALUE Then
          Values(N) = v
          IsMissing(N) = False
        Else
          Values(N) = 0
          IsMissing(N) = True
        End If
      Else
        Values(N) = 0
        IsMissing(N) = True
      End If
      N = N + 1
    End If
  Next i

  If N > 0 Then
    ReDim Preserve Labels(0 To N - 1)
    ReDim Preserve Values(0 To N - 1)
    ReDim Preserve IsMissing(0 To N - 1)
  Else
    ReDim Labels(0 To 0)
    ReDim Values(0 To 0)
    ReDim IsMissing(0 To 0)
  End If
End Sub

' 从Grid中取出两行（UseRow=True）或两列（UseRow=False），组成散点图
Public Sub ExtractPairedSeries(ByVal Grid As Object, ByVal UseRow As Boolean, ByVal Index1 As Long, ByVal Index2 As Long, _
                                ByRef X() As Double, ByRef Y() As Double, ByRef N As Long)
  Dim total As Long, i As Long, v1 As Double, v2 As Double, s1 As String, s2 As String, HeaderText As String
  Dim colFrom As Long, colTo As Long
  N = 0

  colFrom = Grid.FixedCols
  colTo = Grid.Cols - 1

  ' 与上面ExtractSeries相同的FrmImportData行模式限制。
  If UseRow And TypeName(Grid.Parent) = "FrmImportData" Then
    Dim selStart As Long, selEnd As Long
    selStart = Grid.Col
    selEnd = Grid.ColSel
    If selEnd < selStart Then
      Dim tmpCol As Long
      tmpCol = selStart: selStart = selEnd: selEnd = tmpCol
    End If
    If selStart < colFrom Then selStart = colFrom
    If selEnd > colTo Then selEnd = colTo
    If selEnd >= selStart Then
      colFrom = selStart
      colTo = selEnd
    End If
  End If

  If UseRow Then
    total = colTo - colFrom + 1
  Else
    total = Grid.Rows - Grid.FixedRows
  End If
  If total < 1 Then total = 1
  ReDim X(0 To total - 1)
  ReDim Y(0 To total - 1)

  If UseRow Then
    For i = colFrom To colTo
      If Grid.FixedRows >= 1 Then
        HeaderText = Grid.TextMatrix(Grid.FixedRows - 1, i)
      Else
        HeaderText = ""
      End If
      If Not (IsRowSummaryColumn(HeaderText) Or Trim(HeaderText) = "") Then
        s1 = Trim(Grid.TextMatrix(Index1, i))
        s2 = Trim(Grid.TextMatrix(Index2, i))
        If IsNumeric(s1) And IsNumeric(s2) Then
          v1 = CDbl(s1): v2 = CDbl(s2)
          If v1 <> MISSING_VALUE And v2 <> MISSING_VALUE Then
            X(N) = v1: Y(N) = v2
            N = N + 1
          End If
        End If
      End If
    Next i
  Else
    For i = Grid.FixedRows To Grid.Rows - 1
      s1 = Trim(Grid.TextMatrix(i, Index1))
      s2 = Trim(Grid.TextMatrix(i, Index2))
      If IsNumeric(s1) And IsNumeric(s2) Then
        v1 = CDbl(s1): v2 = CDbl(s2)
        If v1 <> MISSING_VALUE And v2 <> MISSING_VALUE Then
          X(N) = v1: Y(N) = v2
          N = N + 1
        End If
      End If
    Next i
  End If

  If N > 0 Then
    ReDim Preserve X(0 To N - 1)
    ReDim Preserve Y(0 To N - 1)
  Else
    ReDim X(0 To 0)
    ReDim Y(0 To 0)
  End If
End Sub

' 通用的Y对X最小二乘回归
Public Sub ComputeLinearRegressionXY(ByRef X() As Double, ByRef Y() As Double, ByVal N As Long, _
                                      ByRef Slope As Double, ByRef Intercept As Double, _
                                      ByRef RSquared As Double, ByRef Correlation As Double, ByRef Valid As Boolean)
  Slope = 0: Intercept = 0: RSquared = 0: Correlation = 0: Valid = False
  If N <= 0 Then Exit Sub
  If N = 1 Then
    Intercept = Y(0)
    Valid = True
    Exit Sub
  End If

  Dim i As Long, sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double, sumY2 As Double
  For i = 0 To N - 1
    sumX = sumX + X(i)
    sumY = sumY + Y(i)
    sumXY = sumXY + X(i) * Y(i)
    sumX2 = sumX2 + X(i) * X(i)
    sumY2 = sumY2 + Y(i) * Y(i)
  Next i

  Dim denomX As Double
  denomX = N * sumX2 - sumX * sumX
  If denomX = 0 Then Exit Sub

  Slope = (N * sumXY - sumX * sumY) / denomX
  Intercept = (sumY - Slope * sumX) / N

  Dim denomY As Double, denomXY As Double
  denomY = N * sumY2 - sumY * sumY
  denomXY = denomX * denomY
  If denomXY > 0 Then
    Correlation = (N * sumXY - sumX * sumY) / Sqr(denomXY)
    RSquared = Correlation * Correlation
  End If
  Valid = True
End Sub

' 为一行构建展示用的标签，用于直方图/散点图/时序图的行选择下拉框，也用作图表坐标轴标题
Public Function BuildRowLabel(ByVal Grid As Object, ByVal RowIdx As Long) As String
  BuildRowLabel = Trim(Grid.TextMatrix(RowIdx, Grid.FixedCols - 1))
End Function

' 判断某列是否为真正可绘图的数据列：至少有一个数字单元格，且表头不是日期/排名/等级类标签
Public Function IsDataColumn(ByVal Grid As Object, ByVal ColIndex As Long) As Boolean
  Dim r As Long
  If Grid.FixedRows >= 1 Then
    If IsExcludedColumnName(Grid.TextMatrix(Grid.FixedRows - 1, ColIndex)) Then
      IsDataColumn = False
      Exit Function
    End If
  End If
  For r = Grid.FixedRows To Grid.Rows - 1
    If IsNumeric(Trim(Grid.TextMatrix(r, ColIndex))) Then
      IsDataColumn = True
      Exit Function
    End If
  Next r
  IsDataColumn = False
End Function


Public Function IsExcludedColumnName(ByVal HeaderText As String) As Boolean
  Dim kw As Variant
  For Each kw In Array("日期", "Date", "Level", "Onset", "Before") '"排名",
    If InStr(HeaderText, kw) > 0 Then
      IsExcludedColumnName = True
      Exit Function
    End If
  Next kw
  IsExcludedColumnName = False
End Function

' 将表头文字与平均值/最大值/最小值等汇总列标签进行匹配
Private Function IsRowSummaryColumn(ByVal HeaderText As String) As Boolean
  Select Case HeaderText
    Case "Average", "Cumulative", "Max", "Min", "Township", "County", "City", "Longitude", "Latitude"
      IsRowSummaryColumn = True
    Case Else
      IsRowSummaryColumn = False
  End Select
End Function

Public Sub ComputeStats(ByRef Values() As Double, ByVal N As Long, _
                         ByRef Mean As Double, ByRef Max As Double, ByRef Min As Double, _
                         ByRef Sum As Double, ByRef StdDev As Double)
  Dim i As Long, sq As Double
  Mean = 0: Max = 0: Min = 0: Sum = 0: StdDev = 0
  If N <= 0 Then Exit Sub

  Max = Values(0): Min = Values(0)
  For i = 0 To N - 1
    Sum = Sum + Values(i)
    If Values(i) > Max Then Max = Values(i)
    If Values(i) < Min Then Min = Values(i)
  Next i
  Mean = Sum / N

  If N > 1 Then
    sq = 0
    For i = 0 To N - 1
      sq = sq + (Values(i) - Mean) * (Values(i) - Mean)
    Next i
    StdDev = Sqr(sq / (N - 1))
  End If
End Sub

' 原地升序插入排序，供下面的ComputePercentileStats使用。
Private Sub SortAscending(ByRef Arr() As Double, ByVal N As Long)
  Dim i As Long, j As Long, tmp As Double
  For i = 1 To N - 1
    tmp = Arr(i)
    j = i - 1
    Do While j >= 0
      If Arr(j) <= tmp Then Exit Do
      Arr(j + 1) = Arr(j)
      j = j - 1
    Loop
    Arr(j + 1) = tmp
  Next i
End Sub

' 中位数、众数、第1/3四分位数，以及90/95/99百分位数
Public Sub ComputePercentileStats(ByRef Values() As Double, ByVal N As Long, _
                                   ByRef Median As Double, ByRef Mode As Double, _
                                   ByRef Q1 As Double, ByRef Q3 As Double, _
                                   ByRef P90 As Double, ByRef P95 As Double, ByRef P99 As Double)
  Median = 0: Mode = 0: Q1 = 0: Q3 = 0: P90 = 0: P95 = 0: P99 = 0
  If N <= 0 Then Exit Sub

  Dim Sorted() As Double
  ReDim Sorted(0 To N - 1)
  Dim i As Long
  For i = 0 To N - 1
    Sorted(i) = Values(i)
  Next i
  Call SortAscending(Sorted, N)

  Median = PercentileOf(Sorted, N, 0.5)
  Q1 = PercentileOf(Sorted, N, 0.25)
  Q3 = PercentileOf(Sorted, N, 0.75)
  P90 = PercentileOf(Sorted, N, 0.9)
  P95 = PercentileOf(Sorted, N, 0.95)
  P99 = PercentileOf(Sorted, N, 0.99)
  Mode = ModeOf(Sorted, N)
End Sub

Private Function PercentileOf(ByRef SortedValues() As Double, ByVal N As Long, ByVal p As Double) As Double
  If N = 1 Then
    PercentileOf = SortedValues(0)
    Exit Function
  End If
  Dim idx As Double, lo As Long, hi As Long, frac As Double
  idx = p * (N - 1)
  lo = Int(idx)
  frac = idx - lo
  hi = lo + 1
  If hi > N - 1 Then hi = N - 1
  PercentileOf = SortedValues(lo) + frac * (SortedValues(hi) - SortedValues(lo))
End Function

' 出现频率最高的值（并列时取较小的那个）
Private Function ModeOf(ByRef SortedValues() As Double, ByVal N As Long) As Double
  Dim i As Long, curVal As Double, curCount As Long
  Dim bestVal As Double, bestCount As Long
  curVal = SortedValues(0): curCount = 1
  bestVal = SortedValues(0): bestCount = 1
  For i = 1 To N - 1
    If SortedValues(i) = curVal Then
      curCount = curCount + 1
    Else
      curVal = SortedValues(i)
      curCount = 1
    End If
    If curCount > bestCount Then
      bestCount = curCount
      bestVal = curVal
    End If
  Next i
  ModeOf = bestVal
End Function

' Gamma函数自然对数的Lanczos近似
Private Function LogGamma(ByVal xx As Double) As Double
  Dim X As Double, Y As Double, tmp As Double, ser As Double
  Dim cof(0 To 5) As Double
  Dim j As Integer
  cof(0) = 76.1800917294715
  cof(1) = -86.5053203294168
  cof(2) = 24.0140982408309
  cof(3) = -1.23173957245015
  cof(4) = 1.20865097386618E-03
  cof(5) = -5.395239384953E-06
  Y = xx
  X = xx
  tmp = X + 5.5
  tmp = tmp - (X + 0.5) * Log(tmp)
  ser = 1.00000000019001
  For j = 0 To 5
    Y = Y + 1
    ser = ser + cof(j) / Y
  Next j
  LogGamma = -tmp + Log(2.506628274631 * ser / X)
End Function

' BetaI使用的连分数展开
Private Function BetaCF(ByVal X As Double, ByVal a As Double, ByVal b As Double) As Double
  Const MAXIT As Long = 200
  Const EPS As Double = 0.0000003
  Const FPMIN As Double = 1E-30
  Dim m As Long, m2 As Long
  Dim AA As Double, c As Double, d As Double, del As Double, h As Double
  Dim qab As Double, qam As Double, qap As Double

  qab = a + b
  qap = a + 1
  qam = a - 1
  c = 1
  d = 1 - qab * X / qap
  If Abs(d) < FPMIN Then d = FPMIN
  d = 1 / d
  h = d

  For m = 1 To MAXIT
    m2 = 2 * m
    AA = m * (b - m) * X / ((qam + m2) * (a + m2))
    d = 1 + AA * d
    If Abs(d) < FPMIN Then d = FPMIN
    c = 1 + AA / c
    If Abs(c) < FPMIN Then c = FPMIN
    d = 1 / d
    h = h * d * c

    AA = -(a + m) * (qab + m) * X / ((a + m2) * (qap + m2))
    d = 1 + AA * d
    If Abs(d) < FPMIN Then d = FPMIN
    c = 1 + AA / c
    If Abs(c) < FPMIN Then c = FPMIN
    d = 1 / d
    del = d * c
    h = h * del
    If Abs(del - 1) < EPS Then Exit For
  Next m

  BetaCF = h
End Function

' 正则化不完全beta函数I_x(a,b)
Private Function BetaI(ByVal X As Double, ByVal a As Double, ByVal b As Double) As Double
  Dim bt As Double
  If X <= 0 Then
    BetaI = 0
    Exit Function
  ElseIf X >= 1 Then
    BetaI = 1
    Exit Function
  End If

  bt = Exp(LogGamma(a + b) - LogGamma(a) - LogGamma(b) + a * Log(X) + b * Log(1 - X))

  If X < (a + 1) / (a + b + 2) Then
    BetaI = bt * BetaCF(X, a, b) / a
  Else
    BetaI = 1 - bt * BetaCF(1 - X, b, a) / b
  End If
End Function

' F(1, df2)统计量的双侧p值
Private Function FTestPValue(ByVal FStat As Double, ByVal df2 As Long) As Double
  If df2 <= 0 Or FStat <= 0 Then
    FTestPValue = 1
    Exit Function
  End If
  Dim X As Double
  X = df2 / (df2 + FStat)
  FTestPValue = BetaI(X, df2 / 2#, 0.5)
End Function

' 检验趋势线的斜率是否显著不为零（H0：斜率=0），使用F(1, N-2)检验
Public Sub ComputeTrendSignificance(ByVal N As Long, ByVal RSquared As Double, _
                                     ByRef FStat As Double, ByRef PValue As Double, ByRef IsSignificant As Boolean)
  FStat = 0: PValue = 1: IsSignificant = False
  Dim df2 As Long
  df2 = N - 2
  If df2 <= 0 Then Exit Sub

  If RSquared >= 0.9999999 Then
    FStat = 999999
    PValue = 0
  ElseIf RSquared <= 0 Then
    FStat = 0
    PValue = 1
  Else
    FStat = (RSquared / (1 - RSquared)) * df2
    PValue = FTestPValue(FStat, df2)
  End If

  IsSignificant = (PValue < 0.05)
End Sub

' Values对其序数位置（0, 1, 2, ...）的普通最小二乘线性回归
Public Sub ComputeLinearTrend(ByRef Values() As Double, ByVal N As Long, _
                               ByRef Slope As Double, ByRef Intercept As Double, ByRef RSquared As Double)
  Slope = 0: Intercept = 0: RSquared = 0
  If N <= 0 Then Exit Sub
  If N = 1 Then
    Intercept = Values(0)
    Exit Sub
  End If

  Dim i As Long, sumX As Double, sumY As Double, sumXY As Double, sumX2 As Double
  For i = 0 To N - 1
    sumX = sumX + i
    sumY = sumY + Values(i)
    sumXY = sumXY + i * Values(i)
    sumX2 = sumX2 + i * i
  Next i

  Dim denom As Double
  denom = N * sumX2 - sumX * sumX
  If denom = 0 Then Exit Sub

  Slope = (N * sumXY - sumX * sumY) / denom
  Intercept = (sumY - Slope * sumX) / N

  Dim meanY As Double, ssTot As Double, ssRes As Double, predicted As Double
  meanY = sumY / N
  For i = 0 To N - 1
    predicted = Slope * i + Intercept
    ssTot = ssTot + (Values(i) - meanY) * (Values(i) - meanY)
    ssRes = ssRes + (Values(i) - predicted) * (Values(i) - predicted)
  Next i
  If ssTot > 0 Then RSquared = 1 - ssRes / ssTot
End Sub

' 读取源窗体"List_Col"要素选择列表框中当前选中的项
Public Function GetSelectedElementName(ByVal ActForm As Object) As String
  Dim lst As Object, i As Long
  GetSelectedElementName = ""
  On Error Resume Next
  Set lst = ActForm.List_Col
  On Error GoTo 0
  If lst Is Nothing Then
    GetSelectedElementName = ActForm.Caption
    Exit Function
  End If

  On Error Resume Next
  For i = 0 To lst.ListCount - 1
    If lst.Selected(i) Then
      GetSelectedElementName = lst.List(i)
      Exit Function
    End If
  Next i
  GetSelectedElementName = ActForm.Caption
End Function

' 由FrmMain直方图工具栏菜单项调用的入口
Public Sub ShowHistogramForActiveForm()
  Dim ActForm As Object
  On Error Resume Next
  Set ActForm = FrmMain.ActiveForm
  On Error GoTo 0

  If ActForm Is Nothing Then
    MsgBox "Please open a query window first", vbInformation, "Notice"
    Exit Sub
  End If
  If Not IsSupportedDataForm(ActForm) Then
    MsgBox "This window does not support chart plotting", vbInformation, "Notice"
    Exit Sub
  End If

  Dim Grid As Object
  On Error Resume Next
  Set Grid = ActForm.mLastRightClickGrid
  On Error GoTo 0
  If Grid Is Nothing Then
    On Error Resume Next
    Set Grid = ActForm.HFGrid1
    On Error GoTo 0
  End If
  If Grid Is Nothing Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If
  If Grid.Rows <= Grid.FixedRows Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If

  Dim ModeStr As String
  ModeStr = DetermineSelectionMode(Grid)
  If ModeStr = "ROW" And IsRowSelectionUnsupported(ActForm) Then
    ModeStr = "COL"
  End If

  Dim ElementName As String
  ElementName = GetSelectedElementName(ActForm)

  Dim NewChart As New FrmHistogram
  Call NewChart.InitFromSelection(Grid, ActForm.Caption, ElementName, ModeStr)
  NewChart.Show vbModeless
End Sub

' 由FrmMain时序图工具栏菜单项调用的入口
Public Sub ShowTemporalChartForActiveForm()
  Dim ActForm As Object
  On Error Resume Next
  Set ActForm = FrmMain.ActiveForm
  On Error GoTo 0

  If ActForm Is Nothing Then
    MsgBox "Please open a query window first", vbInformation, "Notice"
    Exit Sub
  End If
  If Not IsSupportedDataForm(ActForm) Then
    MsgBox "This window does not support chart plotting", vbInformation, "Notice"
    Exit Sub
  End If
  If IsRowSelectionUnsupported(ActForm) Then
    MsgBox "This module's row data does not form a time series -- Time Series is not supported here", vbInformation, "Notice"
    Exit Sub
  End If

  Dim Grid As Object
  On Error Resume Next
  Set Grid = ActForm.mLastRightClickGrid
  On Error GoTo 0
  If Grid Is Nothing Then
    On Error Resume Next
    Set Grid = ActForm.HFGrid1
    On Error GoTo 0
  End If
  If Grid Is Nothing Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If
  If Grid.Rows <= Grid.FixedRows Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If

  Dim ElementName As String
  ElementName = GetSelectedElementName(ActForm)

  Dim NewChart2 As New FrmTemporal
  Call NewChart2.InitFromSelection(Grid, ActForm.Caption, ElementName)
  NewChart2.Show vbModeless
End Sub

' 由FrmMain散点图工具栏菜单项调用的入口
Public Sub ShowScatterForActiveForm()
  Dim ActForm As Object
  On Error Resume Next
  Set ActForm = FrmMain.ActiveForm
  On Error GoTo 0

  If ActForm Is Nothing Then
    MsgBox "Please open a query window first", vbInformation, "Notice"
    Exit Sub
  End If
  If Not IsSupportedDataForm(ActForm) Then
    MsgBox "This window does not support chart plotting", vbInformation, "Notice"
    Exit Sub
  End If

  Dim Grid As Object
  On Error Resume Next
  Set Grid = ActForm.mLastRightClickGrid
  On Error GoTo 0
  If Grid Is Nothing Then
    On Error Resume Next
    Set Grid = ActForm.HFGrid1
    On Error GoTo 0
  End If
  If Grid Is Nothing Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If
  If Grid.Rows <= Grid.FixedRows Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If

  Dim ModeStr As String
  ModeStr = DetermineSelectionMode(Grid)
  If ModeStr = "ROW" And IsRowSelectionUnsupported(ActForm) Then
    ModeStr = "COL"
  End If

  Dim NewChart3 As New FrmScatter
  Call NewChart3.InitFromSelection(Grid, ActForm.Caption, ModeStr)
  NewChart3.Show vbModeless
End Sub




Public Sub ShowHistogramForPeriodExt()
  Dim ActForm As Object
  On Error Resume Next
  Set ActForm = FrmMeteoPeriodExt
  On Error GoTo 0

  If ActForm Is Nothing Then
    MsgBox "Please open a query window first", vbInformation, "Notice"
    Exit Sub
  End If
  If Not IsSupportedDataForm(ActForm) Then
    MsgBox "This window does not support chart plotting", vbInformation, "Notice"
    Exit Sub
  End If

  ' 使用用户最近点击过的那个网格（HFGrid1/2/3/4/5）
  Dim Grid As Object
  On Error Resume Next
  Set Grid = ActForm.mLastRightClickGrid
  On Error GoTo 0
  If Grid Is Nothing Then
    On Error Resume Next
    Set Grid = ActForm.HFGrid1
    On Error GoTo 0
  End If
  If Grid Is Nothing Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If
  If Grid.Rows <= Grid.FixedRows Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If

  Dim ModeStr As String
  ModeStr = DetermineSelectionMode(Grid)
  If ModeStr = "ROW" And IsRowSelectionUnsupported(ActForm) Then
    ModeStr = "COL"
  End If

  Dim ElementName As String
  ElementName = GetSelectedElementName(ActForm)

  Dim NewChart As New FrmHistogram
  Call NewChart.InitFromSelection(Grid, ActForm.Caption, ElementName, ModeStr)
  NewChart.Show vbModeless
End Sub


Public Sub ShowScatterForPeriodExt()
  Dim ActForm As Object
  On Error Resume Next
  Set ActForm = FrmMeteoPeriodExt
  On Error GoTo 0

  If ActForm Is Nothing Then
    MsgBox "Please open a query window first", vbInformation, "Notice"
    Exit Sub
  End If
  If Not IsSupportedDataForm(ActForm) Then
    MsgBox "This window does not support chart plotting", vbInformation, "Notice"
    Exit Sub
  End If

  Dim Grid As Object
  On Error Resume Next
  Set Grid = ActForm.mLastRightClickGrid
  On Error GoTo 0
  If Grid Is Nothing Then
    On Error Resume Next
    Set Grid = ActForm.HFGrid1
    On Error GoTo 0
  End If
  If Grid Is Nothing Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If
  If Grid.Rows <= Grid.FixedRows Then
    MsgBox "Please run a query before plotting a chart", vbInformation, "Notice"
    Exit Sub
  End If

  Dim ModeStr As String
  ModeStr = DetermineSelectionMode(Grid)
  If ModeStr = "ROW" And IsRowSelectionUnsupported(ActForm) Then
    ModeStr = "COL"
  End If

  Dim NewChart3 As New FrmScatter
  Call NewChart3.InitFromSelection(Grid, ActForm.Caption, ModeStr)
  NewChart3.Show vbModeless
End Sub

' 计算一个接近rawRange的"整齐"数（1/2/5 x 10^n）
Private Function NiceNum(ByVal rawRange As Double, ByVal doRound As Boolean) As Double
  If rawRange <= 0 Then rawRange = 1
  Dim exponent As Double, fraction As Double, niceFraction As Double
  exponent = Int(Log(rawRange) / Log(10#))
  fraction = rawRange / (10# ^ exponent)
  If doRound Then
    If fraction < 1.5 Then
      niceFraction = 1
    ElseIf fraction < 3 Then
      niceFraction = 2
    ElseIf fraction < 7 Then
      niceFraction = 5
    Else
      niceFraction = 10
    End If
  Else
    If fraction <= 1 Then
      niceFraction = 1
    ElseIf fraction <= 2 Then
      niceFraction = 2
    ElseIf fraction <= 5 Then
      niceFraction = 5
    Else
      niceFraction = 10
    End If
  End If
  NiceNum = niceFraction * (10# ^ exponent)
End Function

' 给定数据范围[dataMin, dataMax]
Public Function GetValidFontSize(ByVal RawText As String, ByVal DefaultSize As Single) As Single
  Dim s As String
  s = Trim(RawText)
  If Not IsNumeric(s) Then
    GetValidFontSize = DefaultSize
    Exit Function
  End If
  Dim v As Single
  v = CSng(s)
  If v < 6 Then v = 6
  If v > 72 Then v = 72
  GetValidFontSize = v
End Function

' 解析坐标轴min/max/step文本框的原始文本
Public Function TryParseAxisValue(ByVal RawText As String, ByRef OutValue As Double) As Boolean
  Dim s As String
  s = Trim(RawText)
  If s = "" Or Not IsNumeric(s) Then
    TryParseAxisValue = False
    Exit Function
  End If
  OutValue = CDbl(s)
  TryParseAxisValue = True
End Function

Public Sub ComputeNiceTicks(ByVal dataMin As Double, ByVal dataMax As Double, ByVal targetCount As Long, _
                             ByRef niceMin As Double, ByRef niceMax As Double, ByRef tickStep As Double)
  If dataMax <= dataMin Then dataMax = dataMin + 1
  If targetCount < 2 Then targetCount = 2

  Dim rawStep As Double
  rawStep = NiceNum((dataMax - dataMin) / (targetCount - 1), True)
  niceMin = Int(dataMin / rawStep) * rawStep
  niceMax = Int(dataMax / rawStep + 1) * rawStep
  tickStep = rawStep
End Sub

' 在[axisMin, axisMax]（绘图区自身已经留白过的坐标范围）内，按整齐的刻度值画Y轴刻度线+标签
Public Sub DrawYAxis(ByVal Pic As Object, ByVal leftM As Single, ByVal topM As Single, ByVal plotH As Single, _
                      ByVal axisMin As Double, ByVal axisMax As Double, ByVal YTitle As String, _
                      Optional ByVal ForcedTickStep As Double = 0)
  Dim niceMin As Double, niceMax As Double, tickStep As Double
  If ForcedTickStep > 0 Then
    tickStep = ForcedTickStep
    niceMin = Int(axisMin / tickStep) * tickStep
    niceMax = Int(axisMax / tickStep + 1) * tickStep
  Else
    Call ComputeNiceTicks(axisMin, axisMax, 10, niceMin, niceMax, tickStep)
  End If

  Dim v As Double, yPix As Single, lbl As String, maxLabelWidth As Single
  maxLabelWidth = 0
  v = niceMin
  Do While v <= niceMax + tickStep * 0.001
    If v >= axisMin - tickStep * 0.001 And v <= axisMax + tickStep * 0.001 Then
      yPix = topM + plotH - CSng((v - axisMin) / (axisMax - axisMin) * plotH)
      Pic.Line (leftM - 60, yPix)-(leftM, yPix), vbBlack
      lbl = Format(v, "0.0")
      If Pic.TextWidth(lbl) > maxLabelWidth Then maxLabelWidth = Pic.TextWidth(lbl)
      Pic.CurrentX = leftM - 100 - Pic.TextWidth(lbl)
      Pic.CurrentY = yPix - Pic.TextHeight(lbl) / 2
      Pic.Print lbl
    End If
    v = v + tickStep
  Loop

  If Len(YTitle) > 0 Then

    Dim titleX As Single, titleLeftShift As Single

    titleLeftShift = 100
    titleX = leftM - 100 - maxLabelWidth - 100 - Pic.TextHeight(YTitle) - titleLeftShift

    If titleX < Pic.TextHeight(YTitle) * 0.5 Then titleX = Pic.TextHeight(YTitle) * 0.5
    Call DrawRotatedText(Pic, titleX, topM + plotH / 2 + Pic.TextWidth(YTitle) / 2, YTitle, 900)
  End If
End Sub


Private Sub DrawRotatedText(ByVal Pic As Object, ByVal X As Single, ByVal Y As Single, ByVal Text As String, ByVal AngleTenths As Long)
  Dim lf As LOGFONT
  Dim hFont As Long, hOldFont As Long, oldBkMode As Long, oldColor As Long
  Dim pxHeight As Long, pxX As Long, pxY As Long

  pxHeight = Pic.ScaleY(Pic.TextHeight("Ag"), vbTwips, vbPixels)
  pxX = Pic.ScaleX(X, vbTwips, vbPixels)
  pxY = Pic.ScaleY(Y, vbTwips, vbPixels)

  lf.lfHeight = -pxHeight
  lf.lfWeight = IIf(Pic.Font.Bold, 700, 400)
  lf.lfItalic = IIf(Pic.Font.Italic, 1, 0)
  lf.lfUnderline = IIf(Pic.Font.Underline, 1, 0)
  lf.lfStrikeOut = IIf(Pic.Font.Strikethrough, 1, 0)
  lf.lfCharSet = GB2312_CHARSET
  lf.lfOutPrecision = OUT_TT_PRECIS
  lf.lfQuality = CLEARTYPE_QUALITY
  lf.lfFaceName = Pic.Font.Name & vbNullChar
  lf.lfEscapement = AngleTenths
  lf.lfOrientation = AngleTenths
  hFont = CreateFontIndirect(lf)

  hOldFont = SelectObject(Pic.hdc, hFont)
  oldBkMode = SetBkMode(Pic.hdc, TRANSPARENT)
  oldColor = SetTextColor(Pic.hdc, vbBlack)

  Call TextOutW(Pic.hdc, pxX, pxY, strPtr(Text), Len(Text))

  Call SetTextColor(Pic.hdc, oldColor)
  Call SelectObject(Pic.hdc, hOldFont)
  Call DeleteObject(hFont)
  Call SetBkMode(Pic.hdc, oldBkMode)
  Pic.Refresh
End Sub

Public Sub DrawXAxisNumeric(ByVal Pic As Object, ByVal leftM As Single, ByVal bottomPix As Single, ByVal plotW As Single, _
                             ByVal axisMin As Double, ByVal axisMax As Double, _
                             Optional ByVal ForcedTickStep As Double = 0)
  Dim niceMin As Double, niceMax As Double, tickStep As Double
  If ForcedTickStep > 0 Then
    tickStep = ForcedTickStep
    niceMin = Int(axisMin / tickStep) * tickStep
    niceMax = Int(axisMax / tickStep + 1) * tickStep
  Else
    Call ComputeNiceTicks(axisMin, axisMax, 10, niceMin, niceMax, tickStep)
  End If

  Dim v As Double, xPix As Single, lbl As String
  v = niceMin
  Do While v <= niceMax + tickStep * 0.001
    If v >= axisMin - tickStep * 0.001 And v <= axisMax + tickStep * 0.001 Then
      xPix = leftM + CSng((v - axisMin) / (axisMax - axisMin) * plotW)
      Pic.Line (xPix, bottomPix)-(xPix, bottomPix + 60), vbBlack
      lbl = Format(v, "0.0")
      Pic.CurrentX = xPix - Pic.TextWidth(lbl) / 2
      Pic.CurrentY = bottomPix + 80
      Pic.Print lbl
    End If
    v = v + tickStep
  Loop
End Sub

' 对按时间顺序排列的日期/时段标签重新格式化
Public Sub ApplyYearElision(ByRef Labels() As String, ByVal N As Long)
  Dim i As Long
  Dim Yr As String, Rest As String, NoDash As Boolean

  Dim firstYr As String, allSameYear As Boolean, anyDated As Boolean
  firstYr = ""
  allSameYear = True
  anyDated = False
  For i = 0 To N - 1
    Call SplitYearAndRest(Labels(i), Yr, Rest, NoDash)
    If Yr <> "" And Rest <> "" Then
      anyDated = True
      If firstYr = "" Then
        firstYr = Yr
      ElseIf Yr <> firstYr Then
        allSameYear = False
      End If
    End If
  Next i
  If Not anyDated Then allSameYear = False

  Dim lastYr As String
  lastYr = ""
  For i = 0 To N - 1
    Call SplitYearAndRest(Labels(i), Yr, Rest, NoDash)
    If Yr <> "" Then
      If Rest <> "" And allSameYear Then
        Labels(i) = Rest
      ElseIf Yr <> lastYr Then
        If Rest = "" Then
          Labels(i) = Yr
        ElseIf NoDash Then
          Labels(i) = Yr & Rest
        Else
          Labels(i) = Yr & "-" & Rest
        End If
        lastYr = Yr
      Else
        If Rest <> "" Then Labels(i) = Rest
      End If
    End If
  Next i
End Sub

' 将原始网格表头标签拆分为开头的4位年份和剩余部分
Private Sub SplitYearAndRest(ByVal Label As String, ByRef Yr As String, ByRef Rest As String, ByRef NoDash As Boolean)
  Yr = ""
  Rest = ""
  NoDash = False
  If Len(Label) < 5 Then Exit Sub
  If Not IsNumeric(Left(Label, 4)) Then Exit Sub

  Dim sep As String
  sep = Mid(Label, 5, 1)
  If sep <> "-" And sep <> "年" Then Exit Sub

  Yr = Left(Label, 4)
  Rest = Mid(Label, 6)

  If sep = "-" Then
    If Right(Rest, 3) = ":00" Then Rest = Left(Rest, Len(Rest) - 3)
  ElseIf sep = "年" Then
    If Len(Rest) >= 3 Then
      If IsNumeric(Mid(Rest, 1, 2)) And Mid(Rest, 3, 1) = "月" Then
        Dim monthNum As String, afterMonth As String
        monthNum = CStr(CLng(Mid(Rest, 1, 2)))
        afterMonth = Mid(Rest, 4)
        If afterMonth = "" Then
          Rest = monthNum
        Else
          Rest = monthNum & "月" & afterMonth
        End If
      ElseIf InStr(Rest, "Quarter") > 0 Then
        NoDash = True
      End If
    End If
  End If
End Sub


