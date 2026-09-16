VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmImportData 
   Caption         =   "Import Data"
   ClientHeight    =   13530
   ClientLeft      =   150
   ClientTop       =   345
   ClientWidth     =   28380
   Icon            =   "FrmImportData.frx":0000
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
      Left            =   3000
      TabIndex        =   16
      Top             =   2520
      Visible         =   0   'False
      Width           =   975
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H00DFEFDF&
      ForeColor       =   &H80000008&
      Height          =   960
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   28095
      Begin VB.Frame Frame4 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   16680
         TabIndex        =   11
         Top             =   120
         Width           =   4695
         Begin VB.ComboBox ComboLon 
            Height          =   300
            Left            =   1200
            Style           =   2  'Dropdown List
            TabIndex        =   13
            Top             =   240
            Width           =   1095
         End
         Begin VB.ComboBox ComboLat 
            Height          =   300
            Left            =   3480
            Style           =   2  'Dropdown List
            TabIndex        =   12
            Top             =   240
            Width           =   1095
         End
         Begin VB.Label Label2 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Longitude"
            Height          =   255
            Left            =   240
            TabIndex        =   15
            Top             =   300
            Width           =   855
         End
         Begin VB.Label Label3 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Latitude"
            Height          =   255
            Left            =   2640
            TabIndex        =   14
            Top             =   285
            Width           =   855
         End
      End
      Begin VB.Frame Frame3 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   9240
         TabIndex        =   6
         Top             =   120
         Width           =   7455
         Begin VB.CommandButton CmdRefresh 
            Caption         =   "Refresh"
            Height          =   375
            Left            =   6120
            TabIndex        =   9
            Top             =   240
            Width           =   1215
         End
         Begin VB.CheckBox ChkNoHeader 
            BackColor       =   &H00DFEFDF&
            Caption         =   "No header in first row"
            Height          =   255
            Left            =   3720
            TabIndex        =   8
            Top             =   300
            Width           =   2295
         End
         Begin VB.TextBox TxtFixedCols 
            Height          =   345
            Left            =   2760
            TabIndex        =   7
            Text            =   "0"
            Top             =   240
            Width           =   615
         End
         Begin VB.Label LblFixedCols 
            BackColor       =   &H00DFEFDF&
            Caption         =   "Number of Attribute Columns"
            Height          =   255
            Left            =   240
            TabIndex        =   10
            Top             =   300
            Width           =   2460
         End
      End
      Begin VB.Frame Frame11 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   18720
         TabIndex        =   5
         Top             =   120
         Width           =   9255
      End
      Begin VB.Frame Frame2 
         Appearance      =   0  'Flat
         BackColor       =   &H00DFEFDF&
         ForeColor       =   &H00C0C0C0&
         Height          =   735
         Left            =   120
         TabIndex        =   2
         Top             =   120
         Width           =   9135
         Begin VB.CommandButton CmdOpen 
            Caption         =   "Import"
            Height          =   375
            Left            =   240
            TabIndex        =   4
            Top             =   240
            Width           =   1215
         End
         Begin VB.TextBox TextFile 
            Height          =   345
            Left            =   1560
            Locked          =   -1  'True
            TabIndex        =   3
            Top             =   240
            Width           =   7335
         End
         Begin MSComDlg.CommonDialog CommonDialogOpen 
            Left            =   0
            Top             =   0
            _ExtentX        =   847
            _ExtentY        =   847
            _Version        =   393216
         End
      End
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid1 
      Height          =   12225
      Left            =   120
      TabIndex        =   0
      Top             =   1080
      Width           =   28095
      _ExtentX        =   49556
      _ExtentY        =   21564
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
         Caption         =   "Copy Selected Cells"
      End
   End
End
Attribute VB_Name = "FrmImportData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim strFile$
Private mEditGrid As Object '判断当前正在被TextEdit编辑的是哪个Grid（HFGrid1/HFGrid2/HFGrid3(Index)之一）
Public mLastRightClickGrid As Object

Dim mColStacode As Long
Dim mColLon As Long
Dim mColLat As Long

Dim mLines() As String
Dim mLineCount As Long

Private Sub Form_Load()
  FrmMain.FormAdd

  mColStacode = -1
  mColLon = -1
  mColLat = -1

  mLineCount = 0
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


Private Sub CmdOpen_Click()
  CommonDialogOpen.Filter = "csv" & "File" & "(*.csv)|*.csv"
  CommonDialogOpen.InitDir = App.Path & "\Output"
  CommonDialogOpen.CancelError = False
  CommonDialogOpen.FileName = ""
  CommonDialogOpen.ShowOpen
  If CommonDialogOpen.FileName = "" Then Exit Sub

  If ReadCSVLines(CommonDialogOpen.FileName) Then
    TextFile.Text = CommonDialogOpen.FileName
    strFile = CommonDialogOpen.FileName
    Call RenderGrid
  End If
End Sub

' 用当前的固定列数和无表头设置重新渲染已加载的文件
Private Sub CmdRefresh_Click()
  If mLineCount = 0 Then
    MsgBox "Please import a data file first", vbInformation, "Notice"
    Exit Sub
  End If
  Call RenderGrid
End Sub

Private Function ReadCSVLines(ByVal FilePath As String) As Boolean
  Dim tempStr As String

  ReadCSVLines = False
  mLineCount = 0

  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then mLineCount = mLineCount + 1
  Loop
  Close #FileNum

  If mLineCount = 0 Then
    MsgBox "File is empty or contains no valid data", vbInformation, "Notice"
    Exit Function
  End If

  ReDim mLines(0 To mLineCount - 1)
  mLineCount = 0

  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then
      mLines(mLineCount) = tempStr
      mLineCount = mLineCount + 1
    End If
  Loop
  Close #FileNum

  ReadCSVLines = True
End Function

' 根据内存中的mLines()/mLineCount重建HFGrid1，遵循当前的固定列数（TxtFixedCols）和无表头标志
Private Sub RenderGrid()
  Dim i As Long, j As Long
  Dim HasHeader As Boolean
  Dim FixedColCount As Long
  Dim MaxCols As Long
  Dim DataRowCount As Long
  Dim StartLine As Long
  Dim Heads() As String
  Dim Fields() As String

  HasHeader = (ChkNoHeader.Value = vbUnchecked)

  FixedColCount = 0
  If IsNumeric(TxtFixedCols.Text) Then
    FixedColCount = CLng(TxtFixedCols.Text)
    If FixedColCount < 0 Then FixedColCount = 0
  End If

  MaxCols = 0
  For i = 0 To mLineCount - 1
    Fields = Split(mLines(i), ",")
    If UBound(Fields) > MaxCols Then MaxCols = UBound(Fields)
  Next i

  If HasHeader Then
    DataRowCount = mLineCount - 1
    StartLine = 1
  Else
    DataRowCount = mLineCount
    StartLine = 0
  End If

  If DataRowCount <= 0 Then
    MsgBox "File is empty or contains no valid data", vbInformation, "Notice"
    Exit Sub
  End If

  HFGrid1.Rows = DataRowCount + 1
  HFGrid1.Cols = MaxCols + 2
  HFGrid1.FixedRows = 1

  If FixedColCount > HFGrid1.Cols - 2 Then FixedColCount = HFGrid1.Cols - 2
  If FixedColCount < 0 Then FixedColCount = 0
  HFGrid1.FixedCols = 1 + FixedColCount

  If HasHeader Then
    Heads = Split(mLines(0), ",")
    For j = 0 To MaxCols
      If j <= UBound(Heads) Then
        HFGrid1.TextMatrix(0, j + 1) = Trim(Heads(j))
      End If
    Next j
    Call DetectSpecialColumns(Heads)
  Else
    For j = 0 To MaxCols
      HFGrid1.TextMatrix(0, j + 1) = "Field" & CStr(j + 1)
    Next j
    mColStacode = -1
  End If

  Call PopulateLonLatCombos

  For i = 0 To DataRowCount - 1
    Fields = Split(mLines(StartLine + i), ",")
    HFGrid1.TextMatrix(i + 1, 0) = i + 1
    For j = 0 To MaxCols
      If j <= UBound(Fields) Then
        HFGrid1.TextMatrix(i + 1, j + 1) = Trim(Fields(j))
      End If
    Next j
  Next i

  Set mLastRightClickGrid = HFGrid1
End Sub

' 根据表头文本（中文或英文）在CSV表头中查找站号
Private Sub DetectSpecialColumns(ByRef Heads() As String)
  Dim j As Long, h As String

  mColStacode = -1

  For j = 0 To UBound(Heads)
    h = Trim(Heads(j))
    If h = "Station Code" Or UCase(h) = "STACODE" Or UCase(h) = "STATION" Then
      mColStacode = j + 1
    End If
  Next j
End Sub


Private Sub PopulateLonLatCombos()
  Dim j As Long, i As Long
  Dim prevLon As String, prevLat As String
  Dim h As String
  Dim restored As Boolean

  prevLon = ComboLon.Text
  prevLat = ComboLat.Text

  ComboLon.Clear
  ComboLat.Clear
  For j = 1 To HFGrid1.Cols - 1
    h = Trim(HFGrid1.TextMatrix(0, j))
    If h <> "" Then
      ComboLon.AddItem h
      ComboLon.ItemData(ComboLon.NewIndex) = j
      ComboLat.AddItem h
      ComboLat.ItemData(ComboLat.NewIndex) = j
    End If
  Next j

  restored = False
  If prevLon <> "" Then
    For i = 0 To ComboLon.ListCount - 1
      If ComboLon.List(i) = prevLon Then ComboLon.ListIndex = i: restored = True: Exit For
    Next i
  End If
  If Not restored Then
    For i = 0 To ComboLon.ListCount - 1
      If ComboLon.List(i) = "Longitude" Or UCase(ComboLon.List(i)) = "LONGITUDE" Or UCase(ComboLon.List(i)) = "LON" Then
        ComboLon.ListIndex = i
        Exit For
      End If
    Next i
  End If

  restored = False
  If prevLat <> "" Then
    For i = 0 To ComboLat.ListCount - 1
      If ComboLat.List(i) = prevLat Then ComboLat.ListIndex = i: restored = True: Exit For
    Next i
  End If
  If Not restored Then
    For i = 0 To ComboLat.ListCount - 1
      If ComboLat.List(i) = "Latitude" Or UCase(ComboLat.List(i)) = "LATITUDE" Or UCase(ComboLat.List(i)) = "LAT" Then
        ComboLat.ListIndex = i
        Exit For
      End If
    Next i
  End If

  Call ApplyLonLatSelection
End Sub


Private Sub ApplyLonLatSelection()
  If ComboLon.ListIndex >= 0 Then
    mColLon = ComboLon.ItemData(ComboLon.ListIndex)
  Else
    mColLon = -1
  End If
  If ComboLat.ListIndex >= 0 Then
    mColLat = ComboLat.ItemData(ComboLat.ListIndex)
  Else
    mColLat = -1
  End If
End Sub

Private Sub ComboLon_Click()
  Call ApplyLonLatSelection
End Sub

Private Sub ComboLat_Click()
  Call ApplyLonLatSelection
End Sub


Public Sub CmdMap_ImportData()
  Dim DataCol As Long
  DataCol = HFGrid1.Col
  If DataCol < HFGrid1.FixedCols Then
    MsgBox "Please select a data column first", vbInformation, "Notice"
    Exit Sub
  End If
  If DataCol = mColStacode Or DataCol = mColLon Or DataCol = mColLat Then
    MsgBox "Please select a data column first", vbInformation, "Notice"
    Exit Sub
  End If
  If mColStacode < 0 And (mColLon < 0 Or mColLat < 0) Then
    MsgBox "Please first " & Chr(34) & "Longitude Column" & Chr(34) & "/" & Chr(34) & "Latitude Column" & Chr(34) & "specify the latitude/longitude columns in the dropdowns, or make sure the data includes" & Chr(34) & "Station Code" & Chr(34) & " columns to draw a contour map", vbInformation, "Notice"
    Exit Sub
  End If

  Call ImportData.Map_ImportedData(HFGrid1, DataCol, mColStacode, mColLon, mColLat, "BYR")
End Sub

Private Sub HFGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid1
  If Button = 2 Then PopupMenu mnuGridPopup
  Call Sort_Cols_Rows(FrmImportData.HFGrid1, shift)
End Sub

Private Sub mnuGridCopy_Click()
  Call CopyGridSelectionToClipboard(mLastRightClickGrid)
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FrmMain.FormDel
  FrmMain.StatusBar1.Panels(1).Text = ""
  Unload Me
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


'********************双击单元格进行修改***************************
Private Sub HFGrid1_DblClick()
  Call GridEdit(HFGrid1)
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
