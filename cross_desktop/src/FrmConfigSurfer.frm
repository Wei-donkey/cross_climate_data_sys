VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmConfigSurfer 
   Caption         =   "Surfer Settings"
   ClientHeight    =   8295
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5790
   Icon            =   "FrmConfigSurfer.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   30368.14
   ScaleMode       =   0  'User
   ScaleWidth      =   5790
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Gridding Method"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   16
      Top             =   480
      Width           =   2895
      Begin VB.ComboBox ComboGrid_Method 
         Height          =   300
         Left            =   240
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   360
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      Caption         =   "Output Pixels"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   4560
      TabIndex        =   11
      Top             =   480
      Width           =   1095
      Begin VB.TextBox TxtExport_Width 
         Height          =   270
         Left            =   120
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      Caption         =   "Plot Region"
      ForeColor       =   &H80000008&
      Height          =   6375
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   5535
      Begin VB.Frame Frame1 
         Appearance      =   0  'Flat
         ForeColor       =   &H80000008&
         Height          =   2745
         Left            =   1080
         TabIndex        =   19
         Top             =   3480
         Width           =   4335
         Begin VB.CommandButton CmdblnOpen 
            Caption         =   ".bln"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3360
            TabIndex        =   31
            Top             =   600
            Width           =   855
         End
         Begin VB.TextBox TextblnFile 
            Height          =   270
            Left            =   120
            TabIndex        =   30
            Top             =   650
            Width           =   3135
         End
         Begin VB.TextBox TxtLat 
            Height          =   270
            Index           =   1
            Left            =   3120
            TabIndex        =   28
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox TxtLat 
            Height          =   270
            Index           =   0
            Left            =   3120
            TabIndex        =   26
            Top             =   2040
            Width           =   975
         End
         Begin VB.TextBox TxtLon 
            Height          =   270
            Index           =   1
            Left            =   1080
            TabIndex        =   24
            Top             =   2280
            Width           =   975
         End
         Begin VB.TextBox TxtLon 
            Height          =   270
            Index           =   0
            Left            =   1080
            TabIndex        =   22
            Top             =   2040
            Width           =   975
         End
         Begin VB.CommandButton CmdsrfOpen 
            Caption         =   ".srf"
            Enabled         =   0   'False
            Height          =   375
            Left            =   3360
            TabIndex        =   21
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox TextsrfFile 
            Height          =   270
            Left            =   120
            TabIndex        =   20
            Top             =   300
            Width           =   3135
         End
         Begin VB.Label Label6 
            Caption         =   $"FrmConfigSurfer.frx":3482
            Height          =   735
            Left            =   120
            TabIndex        =   32
            Top             =   1080
            Width           =   4095
         End
         Begin VB.Label Label5 
            Caption         =   "LatitudeEnd"
            Height          =   255
            Left            =   2280
            TabIndex        =   29
            Top             =   2310
            Width           =   735
         End
         Begin VB.Label Label4 
            Caption         =   "LatitudeStart"
            Height          =   255
            Left            =   2280
            TabIndex        =   27
            Top             =   2070
            Width           =   735
         End
         Begin VB.Label Label3 
            Caption         =   "LongitudeEnd"
            Height          =   255
            Left            =   240
            TabIndex        =   25
            Top             =   2310
            Width           =   735
         End
         Begin VB.Label Label2 
            Caption         =   "LongitudeStart"
            Height          =   255
            Left            =   240
            TabIndex        =   23
            Top             =   2070
            Width           =   735
         End
      End
      Begin VB.OptionButton OptionDistrict 
         Caption         =   "Custom"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   18
         Top             =   3600
         Width           =   975
      End
      Begin VB.ListBox ListProv 
         Columns         =   4
         Height          =   780
         Left            =   1080
         TabIndex        =   15
         Top             =   360
         Width           =   4335
      End
      Begin VB.ListBox ListCounty 
         Columns         =   4
         Height          =   780
         Left            =   1080
         TabIndex        =   14
         Top             =   2650
         Width           =   4335
      End
      Begin VB.ListBox ListCity 
         Columns         =   4
         Height          =   1500
         Left            =   1080
         TabIndex        =   13
         Top             =   1140
         Width           =   4335
      End
      Begin VB.OptionButton OptionDistrict 
         Caption         =   "County-level"
         Height          =   600
         Index           =   2
         Left            =   120
         TabIndex        =   8
         Top             =   2650
         Width           =   975
      End
      Begin VB.OptionButton OptionDistrict 
         Caption         =   "City-level"
         Height          =   735
         Index           =   1
         Left            =   120
         TabIndex        =   7
         Top             =   1140
         Width           =   855
      End
      Begin VB.OptionButton OptionDistrict 
         Caption         =   "Province-level"
         Height          =   615
         Index           =   0
         Left            =   120
         TabIndex        =   6
         Top             =   360
         Width           =   855
      End
      Begin MSComDlg.CommonDialog CommonDialogOpen 
         Left            =   120
         Top             =   4920
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      Caption         =   "Grid Size"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3120
      TabIndex        =   4
      Top             =   480
      Width           =   1335
      Begin VB.TextBox TxtGrid_Size 
         Height          =   270
         Left            =   120
         TabIndex        =   9
         Top             =   360
         Width           =   615
      End
      Begin VB.Label Label1 
         Caption         =   "km"
         Height          =   255
         Left            =   840
         TabIndex        =   10
         Top             =   360
         Width           =   360
      End
   End
   Begin VB.FileListBox FileLvl 
      Height          =   270
      Left            =   4200
      Pattern         =   "*.lvl"
      TabIndex        =   3
      Top             =   120
      Visible         =   0   'False
      Width           =   495
   End
   Begin VB.CheckBox CheckSurfer 
      Caption         =   "Enable Surfer Plotting"
      Height          =   255
      Left            =   120
      TabIndex        =   2
      ToolTipText     =   "Please make sure Surfer is installed"
      Top             =   120
      Width           =   3540
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   4680
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   7800
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   3720
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   7800
      Width           =   855
   End
End
Attribute VB_Name = "FrmConfigSurfer"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub CheckSurfer_Click()

  If CheckSurfer.Value = Unchecked Then
    
    ComboGrid_Method.Enabled = False
    TxtGrid_Size.Enabled = False
    TxtExport_Width.Enabled = False
    Frame4.Enabled = False
  Else
  
    ComboGrid_Method.Enabled = True
    TxtGrid_Size.Enabled = True
    TxtExport_Width.Enabled = True
    Frame4.Enabled = True
  
  End If
End Sub

Private Sub CmdblnOpen_Click()
  Dim i%, j%
  Dim tempStr$()
  Dim Heads$() '定义表头
  
  CommonDialogOpen.Filter = "BLN Files (*.bln)|*.bln"
  CommonDialogOpen.InitDir = App.Path & "\Template\bln"
  CommonDialogOpen.ShowOpen
  If CommonDialogOpen.Flags = 0 Then FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub

  TextblnFile.Text = CommonDialogOpen.FileName
  Call FillLonLatFromBln(TextblnFile.Text)
End Sub

Private Sub CmdsrfOpen_Click()
  Dim i%, j%
  Dim tempStr$()
  Dim Heads$() '定义表头
  
  CommonDialogOpen.Filter = "srfFile(*.srf)|*.srf"
  CommonDialogOpen.InitDir = App.Path & "\Template\srf"
  CommonDialogOpen.ShowOpen
  If CommonDialogOpen.Flags = 0 Then FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub

  Dim strFileTitle$, strFileName$, strFileDir$, strParentDir$
  
  strFileTitle = CommonDialogOpen.FileTitle
  strFileName = CommonDialogOpen.FileName
  TextsrfFile.Text = strFileName
  strFileDir = Left(strFileName, InStrRev(strFileName, "\") - 1)
  strParentDir = Left(strFileDir, InStrRev(strFileDir, "\") - 1)
  
  tempStr = Split(strFileTitle, ".")
  TextblnFile.Text = strParentDir & "\bln\" & tempStr(0) & ".bln"
  If Len(Dir(TextblnFile.Text)) > 0 Then Call FillLonLatFromBln(TextblnFile.Text)

End Sub

'2026-08-15 根据已选中的bln边界文件，自动计算其经纬度范围并填入TxtLon/TxtLat
Private Sub FillLonLatFromBln(ByVal blnFilePath As String)
  Dim fnum As Integer
  Dim strLine As String
  Dim parts() As String
  Dim nPts As Long, k As Long
  Dim X As Double, Y As Double
  Dim xMin As Double, xMax As Double, yMin As Double, yMax As Double
  Dim flagFirst As Boolean

  If Len(Dir(blnFilePath)) = 0 Then Exit Sub

  flagFirst = True
  fnum = FreeFile
  On Error GoTo ReadErr
  Open blnFilePath For Input As #fnum

  Do While Not EOF(fnum)
    Line Input #fnum, strLine
    If Trim(strLine) = "" Then GoTo NextHeader

    parts = Split(strLine, ",")
    nPts = CLng(Trim(parts(0)))

    For k = 1 To nPts
      If EOF(fnum) Then Exit For
      Line Input #fnum, strLine
      parts = Split(strLine, ",")
      X = CDbl(Trim(parts(0)))
      Y = CDbl(Trim(parts(1)))

      If flagFirst Then
        xMin = X: xMax = X: yMin = Y: yMax = Y
        flagFirst = False
      Else
        If X < xMin Then xMin = X
        If X > xMax Then xMax = X
        If Y < yMin Then yMin = Y
        If Y > yMax Then yMax = Y
      End If
    Next k

NextHeader:
  Loop

  Close #fnum

  If Not flagFirst Then
    TxtLon(0).Text = Format(FloorTo2(xMin), "0.00")
    TxtLon(1).Text = Format(CeilTo2(xMax), "0.00")
    TxtLat(0).Text = Format(FloorTo2(yMin), "0.00")
    TxtLat(1).Text = Format(CeilTo2(yMax), "0.00")
  End If
  Exit Sub

ReadErr:
  Close #fnum
  MsgBox "Failed to read the .bln boundary file -- please enter latitude/longitude manually: " & Err.Description, vbExclamation, "Notice"
End Sub

'向下取整到小数点后2位，如 109.1252452 -> 109.12
Private Function FloorTo2(ByVal v As Double) As Double
  FloorTo2 = Int(v * 100 + 0.0000001) / 100
End Function

'向上取整到小数点后2位，如 117.562425 -> 117.57
Private Function CeilTo2(ByVal v As Double) As Double
  CeilTo2 = -Int(-v * 100 + 0.0000001) / 100
End Function

Private Sub Form_Unload(Cancel As Integer)
  Unload Me
End Sub

Private Sub CmdCancel_Click()
  Unload Me
End Sub

Private Sub CmdOK_Click()
  Dim FileName$, i%
  Dim Cancel%
  
  If CheckSurfer.Value = Unchecked Then
    SetSurfer = "No"
    Call ReadWriteIni("Config.ini", "[Set_Surfer]", SetSurfer, "w")
    FrmMain.StatusBar1.Panels(6).Text = "Not confirmed whether Surfer11 is installed"
    Unload Me
    Exit Sub
  Else
    If GuestMode = True Then ' 如果是访客，又不曾确定是否已安装surfer11，则进行下面的判断，并加载 SrfApp 对象
      If FlagSrfInstalled = False Then
  
        Cancel = MsgBox("Confirm Surfer11 is installed?", vbYesNo + vbQuestion, "Notice")
        If Cancel = vbNo Then  '不确定
          Exit Sub
        ElseIf Cancel = vbYes Then '确认已安装
          FlagSrfInstalled = True
          
        End If
      End If
    End If
  End If
  
  '设置出图分辨率
  Export_Width = CInt(TxtExport_Width.Text)
  Call ReadWriteIni("Config.ini", "[Export_Width]", Export_Width, "w")
    
  '设置画图地区
  PROV_Name = ListProv.Text: CITY_Name = ListCity.Text: COUNTY_Name = ListCounty.Text
  If OptionDistrict(0).Value = True Then
    District_Level = "PROV":  District_Name = PROV_Name
    If PROV_Name = "" Then MsgBox "A province-level region or custom region must be selected", vbInformation, "Notice": Exit Sub
  ElseIf OptionDistrict(1).Value = True Then
    District_Level = "CITY": District_Name = CITY_Name
    If CITY_Name = "" Then MsgBox "A city-level region or custom region must be selected", vbInformation, "Notice": Exit Sub
  ElseIf OptionDistrict(2).Value = True Then
    District_Level = "COUNTY": District_Name = COUNTY_Name
    If COUNTY_Name = "" Then MsgBox "A county-level region or custom region must be selected", vbInformation, "Notice": Exit Sub
    
  ElseIf OptionDistrict(3).Value = True Then
    District_Level = "CUSTOM":    District_Name = "Custom"
    If TextsrfFile.Text = "" Then MsgBox "An .srf template file must be selected", vbInformation, "Notice": Exit Sub
    If TextblnFile.Text = "" Then MsgBox "A .bln mask file must be selected", vbInformation, "Notice": Exit Sub
    
    If Len(Dir(TextsrfFile.Text)) = 0 Then
      MsgBox TextsrfFile.Text & " file not found, please manually select the corresponding .srf file", vbInformation, "Notice": Exit Sub
    End If
    If Len(Dir(TextblnFile.Text)) = 0 Then
      MsgBox TextblnFile.Text & " file not found, please manually select the corresponding .bln file", vbInformation, "Notice": Exit Sub
    End If

    If TxtLon(0).Text = "" Or TxtLon(1).Text = "" Or TxtLat(0).Text = "" Or TxtLat(1).Text = "" Then
      MsgBox "The full latitude/longitude range must be set", vbInformation, "Notice": Exit Sub
    End If
    
  End If
  
  If OptionDistrict(3).Value = False Then
    Call ReadWriteIni("Config.ini", "[District_Level]", District_Level, "w")
    Call ReadWriteIni("Config.ini", "[PROV_Name]", PROV_Name, "w")
    Call ReadWriteIni("Config.ini", "[CITY_Name]", CITY_Name, "w")
    Call ReadWriteIni("Config.ini", "[COUNTY_Name]", COUNTY_Name, "w")
  End If
  
  If OptionDistrict(3).Value = False Then
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    
    strSQL = "select * from T_OTHE_ZONE_META_BASIC_TAB where ADMIN_ZONE='" & District_Name & "'"
    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
    District_ID = ORARst!id
    sngxMin = ORARst!xMin: sngxMax = ORARst!xMax
    sngyMin = ORARst!yMin: sngyMax = ORARst!yMax
    ORARst.Close
    ORAConn4.Close
    
  Else
    sngxMin = TxtLon(0).Text: sngxMax = TxtLon(1).Text
    sngyMin = TxtLat(0).Text: sngyMax = TxtLat(1).Text
  
  End If
    
  '设置网格大小
  Grid_Size = TxtGrid_Size.Text
  Call ReadWriteIni("Config.ini", "[Grid_Size]", Grid_Size, "w")
  '1公里相当于0.01度
  iNumCols = CInt(100 * (sngxMax - sngxMin) / CSng(Grid_Size)): iNumRows = CInt(100 * (sngyMax - sngyMin) / CSng(Grid_Size))
  
  If ComboGrid_Method = ComboGrid_Method.List(0) Then
    Grid_Method = srfInverseDistance
    Call ReadWriteIni("Config.ini", "[Grid_Method]", "srfInverseDistance", "w")
  ElseIf ComboGrid_Method = ComboGrid_Method.List(1) Then
    Grid_Method = srfKriging
    Call ReadWriteIni("Config.ini", "[Grid_Method]", "srfKriging", "w")
  ElseIf ComboGrid_Method = ComboGrid_Method.List(2) Then
    Grid_Method = srfMinCurvature
    Call ReadWriteIni("Config.ini", "[Grid_Method]", "srfMinCurvature", "w")
  ElseIf ComboGrid_Method = ComboGrid_Method.List(3) Then
    Grid_Method = srfNaturalNeighbor
    Call ReadWriteIni("Config.ini", "[Grid_Method]", "srfNaturalNeighbor", "w")
  ElseIf ComboGrid_Method = ComboGrid_Method.List(4) Then
    Grid_Method = srfNearestNeighbor
    Call ReadWriteIni("Config.ini", "[Grid_Method]", "srfNearestNeighbor", "w")
  ElseIf ComboGrid_Method = ComboGrid_Method.List(5) Then
    Grid_Method = srfRegression
    Call ReadWriteIni("Config.ini", "[Grid_Method]", "srfRegression", "w")
  ElseIf ComboGrid_Method = ComboGrid_Method.List(6) Then
    Grid_Method = srfRadialBasis
    Call ReadWriteIni("Config.ini", "[Grid_Method]", "srfRadialBasis", "w")
  ElseIf ComboGrid_Method = ComboGrid_Method.List(7) Then
    Grid_Method = srfTriangulation
    Call ReadWriteIni("Config.ini", "[Grid_Method]", "srfTriangulation", "w")
  End If
    
  
  SrfEdition = "11EN"
  
  If OptionDistrict(3).Value = False Then
'    srfName = District_ID & "_VALUE11.srf"
    srfPath = App.Path & "\Template\srf\" & District_ID & "_VALUE11.srf"
    srfPath_ColdAir = App.Path & "\Template\srf\" & District_ID & "_COLDAIR.srf"
    blnName = District_ID & ".bln"
    blnPath = App.Path & "\Template\bln\" & blnName
  Else
    srfPath = TextsrfFile.Text
    blnPath = TextblnFile.Text
  End If
  
  Call ReadWriteIni("Config.ini", "[Surfer_Edition]", SrfEdition, "w")
  
  If CheckSurfer.Value = Checked Then
    SetSurfer = "Yes"
    FrmMain.StatusBar1.Panels(1).Text = "Loading Surfer plotting object": DoEvents
    Set SrfApp = CreateObject("Surfer.Application")
  ElseIf CheckSurfer.Value = Unchecked Then
    SetSurfer = "No"
    Set SrfApp = Nothing
  End If
  Call ReadWriteIni("Config.ini", "[Set_Surfer]", SetSurfer, "w")
  
  If SetSurfer = "Yes" And OptionDistrict(3).Value = False Then
    
    If Dir(blnPath) = "" Then '如果本地没是bln，则下载
      FrmMain.StatusBar1.Panels(1).Text = "Download .bln mask file via FTP"
      
      FileNum = FreeFile
      Open App.Path & "\Temp\getFTP_" & blnName & ".bat" For Output As #FileNum
        Print #FileNum, "ftp -s:" & App.Path & "\Temp\getFTP_" & blnName & ".ftp"
      Close #FileNum
      
      Call get_FTP_File("Template/bln", App.Path & "\Template\bln", blnName)
      Shell "cmd.exe /c " & App.Path + "\Temp\getFTP_" & blnName & ".bat", vbNormalFocus
    End If
    
    If Dir(srfPath) = "" Then '如果本地没是值.srf，则下载
      FrmMain.StatusBar1.Panels(1).Text = "Download .srf template via FTP"
      
      FileNum = FreeFile
      Open App.Path & "\Temp\getFTP_" & District_ID & "_VALUE11.srf" & ".bat" For Output As #FileNum
        Print #FileNum, "ftp -s:" & App.Path & "\Temp\getFTP_" & District_ID & "_VALUE11.srf" & ".ftp"
      Close #FileNum
      
      Call get_FTP_File("Template/srf", App.Path & "\Template\srf", District_ID & "_VALUE11.srf")
      Shell "cmd.exe /c " & App.Path + "\Temp\getFTP_" & District_ID & "_VALUE11.srf" & ".bat", vbNormalFocus
    End If

    If District_ID = "GD" Then '如果绘图区域是广东，则下载
        If Dir(srfPath_ColdAir) = "" Then '如果本地没是值.srf，则下载
          FrmMain.StatusBar1.Panels(1).Text = "Download GD_COLDAIR.srf template via FTP"
          
          FileNum = FreeFile
          Open App.Path & "\Temp\getFTP_" & District_ID & "_COLDAIR.srf" & ".bat" For Output As #FileNum
            Print #FileNum, "ftp -s:" & App.Path & "\Temp\getFTP_" & District_ID & "_COLDAIR.srf" & ".ftp"
          Close #FileNum
          
          Call get_FTP_File("Template/srf", App.Path & "\Template\srf", District_ID & "_COLDAIR.srf")
          Shell "cmd.exe /c " & App.Path + "\Temp\getFTP_" & District_ID & "_COLDAIR.srf" & ".bat", vbNormalFocus
        End If
    End If


  End If
  
  FrmMain.StatusBar1.Panels(1).Text = "Finish Setup"
  
  If SetSurfer = "Yes" Then
    FrmMain.StatusBar1.Panels(6).Text = "Plot region: " & District_Name
  ElseIf SetSurfer = "No" Then
    FrmMain.StatusBar1.Panels(6).Text = "Surfer plotting not enabled"
  End If
  
  Unload Me
  
End Sub

Private Sub Form_Load()
  Dim i%
  
  '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  
  Me.Left = FrmMain.Left + 500
  Me.Top = FrmMain.Top + FrmMain.Height - Me.Height - 600
  
  If UserLevel >= 2 Then OptionDistrict(0).Enabled = False
  
  FileLvl.Path = App.Path & "\Template\lvl"
  
  TxtExport_Width.Text = Export_Width
  TxtGrid_Size.Text = Grid_Size
  
  '2026-07-22 访客模式下禁止连接数据库
  If Not GuestMode Then
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    
    '************填充省级分布图地区列表*********
    If UserLevel = 1 Then
      ListProv.Clear
      strSQL = "select ADMIN_ZONE from T_OTHE_ZONE_META_BASIC_TAB where ilevel=1 order by ADMIN_ZONE"
      ORARst.CursorLocation = adUseClient
      ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
      Do Until ORARst.EOF
        ListProv.AddItem ORARst!ADMIN_ZONE
        ORARst.MoveNext
      Loop
      ORARst.Close
    End If
    
    '************填充市级分布图地区列表*********
    If UserLevel = 1 Then '省级用户
      strSQL = "select DATAZONE from T_OTHE_ZONE_META_BASIC_TAB where ilevel=2 order by DATAZONE"
    ElseIf UserLevel >= 2 Then  '市县用户
      strSQL = "select DATAZONE from T_OTHE_ZONE_META_BASIC_TAB where ADMIN_ZONE='" & UserZone & "' order by DATAZONE"
    End If
    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
    Do Until ORARst.EOF
      ListCity.AddItem ORARst!DataZone
      ORARst.MoveNext
    Loop
    ORARst.Close
    ORAConn4.Close

  End If
   
  If UserLevel = 1 Then
    If District_Level = "" Then OptionDistrict(0).Value = True
  ElseIf UserLevel = 2 Then
    If District_Level = "" Then OptionDistrict(1).Value = True
  ElseIf UserLevel = 3 Then
    If District_Level = "" Then OptionDistrict(2).Value = True
  End If
  If District_Level = "PROV" Then OptionDistrict(0).Value = True
  If District_Level = "CITY" Then OptionDistrict(1).Value = True
  If District_Level = "COUNTY" Then OptionDistrict(2).Value = True
  If District_Level = "CUSTOM" Then OptionDistrict(3).Value = True
   
   
      
  If GuestMode = True Then '2026-8-13：如果是访客
    OptionDistrict(0).Enabled = False
    OptionDistrict(1).Enabled = False
    OptionDistrict(2).Enabled = False
    OptionDistrict(3).Value = True
  End If
   
  If OptionDistrict(0).Value = True Then
    For i = 0 To ListProv.ListCount - 1
      If ListProv.List(i) = PROV_Name Then ListProv.Selected(i) = True: Exit For
    Next i
  
  ElseIf OptionDistrict(1).Value = True Then
    For i = 0 To ListProv.ListCount - 1
      If ListProv.List(i) = CITY_Name Then ListProv.Selected(i) = True: Exit For
    Next i
  
  ElseIf OptionDistrict(3).Value = True Then
    TextsrfFile.Text = srfPath
    CmdsrfOpen.Enabled = True
    TextblnFile.Text = blnPath
    CmdblnOpen.Enabled = True
    TxtLon(0).Text = sngxMin: TxtLon(1).Text = sngxMax
    TxtLat(0).Text = sngyMin: TxtLat(1).Text = sngyMax
    
  End If
   
  ComboGrid_Method.Clear
  ComboGrid_Method.AddItem "加权反距离": ComboGrid_Method.AddItem "克里格": ComboGrid_Method.AddItem "最小曲率": ComboGrid_Method.AddItem "自然邻点": ComboGrid_Method.AddItem "最近邻点"
  ComboGrid_Method.AddItem "多项式回归": ComboGrid_Method.AddItem "径向基函数": ComboGrid_Method.AddItem "带线性插Value的三角剖分"
  If Grid_Method = srfInverseDistance Then ComboGrid_Method = ComboGrid_Method.List(0) '2对应：srfKriging
  If Grid_Method = srfKriging Then ComboGrid_Method = ComboGrid_Method.List(1) '1对应：srfInverseDistance
  If Grid_Method = srfMinCurvature Then ComboGrid_Method = ComboGrid_Method.List(2) '2对应：srfKriging
  If Grid_Method = srfNaturalNeighbor Then ComboGrid_Method = ComboGrid_Method.List(3) '1对应：srfInverseDistance
  If Grid_Method = srfNearestNeighbor Then ComboGrid_Method = ComboGrid_Method.List(4) '2对应：srfKriging
  If Grid_Method = srfRegression Then ComboGrid_Method = ComboGrid_Method.List(5) '1对应：srfInverseDistance
  If Grid_Method = srfRadialBasis Then ComboGrid_Method = ComboGrid_Method.List(6) '2对应：srfKriging
  If Grid_Method = srfTriangulation Then ComboGrid_Method = ComboGrid_Method.List(7) '1对应：srfInverseDistance

  If SetSurfer = "Yes" Then
    CheckSurfer.Value = Checked
  ElseIf SetSurfer = "No" Then
    CheckSurfer.Value = Unchecked
  End If
  
End Sub


Private Sub ListCity_click()
  ListCounty.Clear
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  
  strSQL = "select ADMIN_ZONE from T_OTHE_ZONE_META_BASIC_TAB where ilevel=3 and DATAZONE='" & ListCity.Text & "' order by ADMIN_ZONE"
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  Do Until ORARst.EOF
    ListCounty.AddItem ORARst!ADMIN_ZONE
    ORARst.MoveNext
  Loop
  ORARst.Close
  ORAConn4.Close
End Sub

Private Sub OptionDistrict_Click(Index As Integer)
  If Index = 3 Then
    CmdsrfOpen.Enabled = True
    CmdblnOpen.Enabled = True
  Else
    CmdsrfOpen.Enabled = False
    CmdblnOpen.Enabled = False
  End If
End Sub
