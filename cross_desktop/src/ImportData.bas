Attribute VB_Name = "ImportData"
Option Explicit


Public Sub Map_ImportedData(ByVal Grid As Object, ByVal DataCol As Long, _
                             ByVal colStacode As Long, ByVal colLon As Long, ByVal colLat As Long, _
                             ByVal Color As String)
  Dim i As Long, j As Long, k As Long
  Dim tempLon As Single, tempLat As Single, TempVal As String
  Dim sngMax As Single, sngMin As Single, sngInterval As Single
  Dim FoundCoord As Boolean

  FileNum = FreeFile
  Open App.Path & "\Temp\temp.txt" For Output As #FileNum
  Close #FileNum
  FileNum = FreeFile
  Open App.Path & "\Temp\temp.txt" For Append As #FileNum

  k = 0
  sngMax = -9999: sngMin = 999999

  For i = Grid.FixedRows To Grid.Rows - 1
    FoundCoord = False

    If colLon >= 0 And colLat >= 0 Then
      If IsNumeric(Grid.TextMatrix(i, colLon)) And IsNumeric(Grid.TextMatrix(i, colLat)) Then
        tempLon = CSng(Grid.TextMatrix(i, colLon))
        tempLat = CSng(Grid.TextMatrix(i, colLat))
        FoundCoord = True
      End If
    ElseIf colStacode >= 0 Then
      For j = 1 To StaNum
        If StaInfo(j).stacode = Trim(Grid.TextMatrix(i, colStacode)) Then
          tempLon = StaInfo(j).Longitude
          tempLat = StaInfo(j).Latitude
          FoundCoord = True
          Exit For
        End If
      Next j
    End If

    If FoundCoord Then
      TempVal = Trim(Grid.TextMatrix(i, DataCol))
      If TempVal <> "" And IsNumeric(TempVal) Then
        Print #FileNum, tempLon; ","; tempLat; ","; TempVal
        k = k + 1
        If CSng(TempVal) > sngMax Then sngMax = CSng(TempVal)
        If CSng(TempVal) < sngMin Then sngMin = CSng(TempVal)
      End If
    End If
  Next i

  Close #FileNum

  If k <= 2 Then
    MsgBox "Not enough valid data points to draw a contour map", vbInformation, "Notice"
    Exit Sub
  End If
  If sngMin = sngMax Then
    MsgBox "No variation across stations -- cannot draw a contour map", vbInformation, "Notice"
    Exit Sub
  End If

  sngInterval = (sngMax - sngMin) / 5

  FrmMapValue.sngMin = sngMin: FrmMapValue.sngMax = sngMax: FrmMapValue.sngInterval = sngInterval
  FrmMapValue.strColor = Color
  FrmMapValue.Show
End Sub


Public Sub LoadGuestCSV(ByVal Grid1 As Object, ByVal Grid2 As Object, ByVal CsvFileName As String, ByVal FixedCols1 As Integer)
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long, j As Long
  Dim SplitRow As Long
  Dim Fields() As String
  Dim MaxCols1 As Long, MaxCols2 As Long
  Dim HasGrid2 As Boolean

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

  ' 表头之后第一个第一列为空的数据行，标志着Grid2部分的开始
  SplitRow = -1
  For i = 1 To LineCount - 1
    Fields = Split(Lines(i), ",")
    If Trim(Fields(0)) = "" Then
      SplitRow = i
      Exit For
    End If
  Next i

  Dim Grid1LastRow As Long
  If SplitRow = -1 Then
    Grid1LastRow = LineCount - 1
    HasGrid2 = False
  Else
    Grid1LastRow = SplitRow - 1
    HasGrid2 = (Not Grid2 Is Nothing)
  End If

  ' ---- Grid1 ----
  MaxCols1 = 0
  For i = 0 To Grid1LastRow
    Fields = Split(Lines(i), ",")
    If UBound(Fields) > MaxCols1 Then MaxCols1 = UBound(Fields)
  Next i

  With Grid1
    .Redraw = False
    .Rows = Grid1LastRow + 1
    .Cols = MaxCols1 + 1
    .FixedRows = 1
    .FixedCols = FixedCols1
    For i = 0 To Grid1LastRow
      Fields = Split(Lines(i), ",")
      For j = 0 To UBound(Fields)
        .TextMatrix(i, j) = Fields(j)
      Next j
    Next i
  
    If FixedCols1 = 3 Then
       '******上方单元格左对齐、前两列着色（白色）********
      For i = 1 To .Rows - 1
        For j = 0 To .Cols - 1
          .Row = i: .Col = j
          .CellAlignment = flexAlignLeftCenter
          If .Col = 1 Then .CellBackColor = &H80000005
          If .Col = 2 Then .CellBackColor = &H80000005
        Next j
      Next i
        For j = 0 To .Cols - 1
          .Row = 0: .Col = j
          .CellAlignment = flexAlignLeftCenter
        Next j
    End If
    
    .Refresh
    .Redraw = True
  End With

  ' ---- Grid2 ----
  If HasGrid2 Then
    Dim isHeaderRepeat As Boolean
    Dim Grid2HeaderRow As Long
    Dim Grid2DataStart As Long

    isHeaderRepeat = (InStr(Lines(SplitRow), "Average") = 0 And InStr(Lines(SplitRow), "Max") = 0 And InStr(Lines(SplitRow), "Min") = 0)
    If isHeaderRepeat Then
      Grid2HeaderRow = SplitRow
      Grid2DataStart = SplitRow + 1
    Else
      Grid2HeaderRow = -1 ' 没有重复的表头行，表示这个导出本身没写第二表格的表头
      Grid2DataStart = SplitRow
    End If

    MaxCols2 = 0
    For i = Grid2DataStart To LineCount - 1
      Fields = Split(Lines(i), ",")
      If UBound(Fields) > MaxCols2 Then MaxCols2 = UBound(Fields)
    Next i
    If Grid2HeaderRow >= 0 Then
      Fields = Split(Lines(Grid2HeaderRow), ",")
      If UBound(Fields) > MaxCols2 Then MaxCols2 = UBound(Fields)
    End If

    With Grid2
     .Redraw = False
    
      If Grid2HeaderRow >= 0 Then
        .Rows = (LineCount - 1) - Grid2DataStart + 1 + 1 ' 有单独表头行时要多加1行（给表头本身留出空间），否则只需要数据行本身的行数
      Else
        .Rows = (LineCount - 1) - Grid2DataStart + 1
      End If
      .Cols = MaxCols2 + 1
'      .FixedRows = 1
      .FixedCols = FixedCols1

      If Grid2HeaderRow >= 0 Then
        Fields = Split(Lines(Grid2HeaderRow), ",")
        For j = 0 To UBound(Fields)
          .TextMatrix(0, j) = Fields(j)
        Next j
      End If
      ' 2026-07-29 本次导出没有单独的第二表格标题行，表示表头是静态的（表单已自带），保留原有表头，不用第一表格的表头覆盖它

      Dim r As Long
      If Grid2HeaderRow >= 0 Then
        r = 1
      Else
        r = 0
      End If
      
      For i = Grid2DataStart To LineCount - 1
        Fields = Split(Lines(i), ",")
        For j = 0 To UBound(Fields)
          .TextMatrix(r, j) = Fields(j)
        Next j
        r = r + 1
      Next i
      
      If FixedCols1 = 3 Then
         '******上方单元格左对齐、前两列着色（白色）********
        For i = 1 To .Rows - 1
          For j = 0 To .Cols - 1
            .Row = i: .Col = j
            .CellAlignment = flexAlignLeftCenter
            If .Col = 1 Then .CellBackColor = &H80000005
            If .Col = 2 Then .CellBackColor = &H80000005
          Next j
        Next i
        
        For j = 0 To .Cols - 1
          .Row = 0: .Col = j
          .CellAlignment = flexAlignLeftCenter
        Next j
        
      End If
    
      .Refresh
      .Redraw = True
    End With
  End If
End Sub


Public Sub LoadGuestTxt(ByVal TextBox As Object, ByVal TxtFileName As String)
  Dim FilePath As String
  Dim i%, tempStr$
  i = 0
  
  FilePath = App.Path & "\Preview\Data\" & TxtFileName
  If Dir(FilePath) = "" Then Exit Sub

  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    
    tempStr = Chr(10) & Str(i + 1) & "、" & tempStr
    TextBox.Text = TextBox.Text & tempStr
    i = i + 1
  Loop
  Close #FileNum
  
End Sub
