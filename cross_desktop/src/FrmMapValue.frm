VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.Form FrmMapValue 
   Caption         =   "Contour Map Settings"
   ClientHeight    =   6120
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6645
   Icon            =   "FrmMapValue.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   10491.43
   ScaleMode       =   0  'User
   ScaleWidth      =   6645
   Begin VB.Frame Frame7 
      Height          =   615
      Left            =   120
      TabIndex        =   33
      Top             =   5280
      Width           =   2895
      Begin VB.CheckBox Check2 
         Caption         =   "Data Labels (Beta)"
         Height          =   255
         Left            =   240
         TabIndex        =   34
         Top             =   240
         Width           =   2535
      End
   End
   Begin VB.Frame Frame6 
      Appearance      =   0  'Flat
      Caption         =   "Grid Size"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   5280
      TabIndex        =   30
      Top             =   4440
      Width           =   1215
      Begin VB.TextBox TxtGrid_Size 
         Height          =   270
         Left            =   120
         TabIndex        =   31
         Top             =   360
         Width           =   495
      End
      Begin VB.Label Label6 
         Caption         =   "km"
         Height          =   255
         Left            =   720
         TabIndex        =   32
         Top             =   360
         Width           =   375
      End
   End
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      Caption         =   "Gridding Method"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   3120
      TabIndex        =   28
      Top             =   4440
      Width           =   2055
      Begin VB.ComboBox ComboGrid_Method 
         Height          =   300
         Left            =   120
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      Caption         =   "Color Level Selection"
      ForeColor       =   &H80000008&
      Height          =   3495
      Left            =   120
      TabIndex        =   8
      Top             =   840
      Width           =   6375
      Begin VB.Frame Frame3 
         BorderStyle     =   0  'None
         Height          =   1095
         Left            =   360
         TabIndex        =   25
         Top             =   1200
         Width           =   1335
         Begin VB.OptionButton OptionBYR 
            Caption         =   "Cool-Warm"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   240
            Width           =   1215
         End
         Begin VB.OptionButton OptionRYB 
            Caption         =   "Warm-Cool"
            Height          =   255
            Left            =   120
            TabIndex        =   26
            Top             =   600
            Width           =   1215
         End
      End
      Begin VB.PictureBox PictureRYB 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         ScaleHeight     =   375
         ScaleWidth      =   4395
         TabIndex        =   21
         Top             =   1560
         Visible         =   0   'False
         Width           =   4400
         Begin VB.Image Image2 
            Height          =   420
            Left            =   0
            Picture         =   "FrmMapValue.frx":3482
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4392
         End
      End
      Begin VB.TextBox TextMin 
         Height          =   270
         Left            =   1320
         TabIndex        =   17
         Top             =   800
         Width           =   735
      End
      Begin VB.TextBox TextMax 
         Height          =   270
         Left            =   3120
         TabIndex        =   16
         Top             =   800
         Width           =   735
      End
      Begin VB.TextBox TextInterval 
         Height          =   270
         Left            =   5400
         TabIndex        =   15
         Top             =   800
         Width           =   735
      End
      Begin VB.CommandButton CmdCustLvl 
         Caption         =   "Select"
         Height          =   375
         Left            =   5400
         TabIndex        =   12
         Top             =   2900
         Width           =   735
      End
      Begin MSComDlg.CommonDialog CommonDialogOpen 
         Left            =   5880
         Top             =   2520
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.TextBox TextCustLvl 
         Height          =   270
         Left            =   600
         TabIndex        =   11
         Top             =   3000
         Width           =   4575
      End
      Begin VB.OptionButton OptionAutoLvl 
         Caption         =   "Auto Color Levels"
         Height          =   255
         Left            =   240
         TabIndex        =   10
         Top             =   360
         Width           =   2175
      End
      Begin VB.OptionButton OptionCustLvl 
         Caption         =   "Custom Color Levels"
         Height          =   375
         Left            =   240
         TabIndex        =   9
         Top             =   2520
         Width           =   2655
      End
      Begin VB.PictureBox PictureBYR 
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BorderStyle     =   0  'None
         ForeColor       =   &H80000008&
         Height          =   375
         Left            =   1800
         ScaleHeight     =   375
         ScaleWidth      =   4395
         TabIndex        =   22
         Top             =   1560
         Visible         =   0   'False
         Width           =   4400
         Begin VB.Image Image1 
            Height          =   420
            Left            =   0
            Picture         =   "FrmMapValue.frx":39F8
            Stretch         =   -1  'True
            Top             =   0
            Width           =   4392
         End
      End
      Begin MSComctlLib.Slider SliderMax 
         Height          =   630
         Left            =   2160
         TabIndex        =   23
         Top             =   1200
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   1111
         _Version        =   393216
         LargeChange     =   2
         Min             =   2
         Max             =   13
         SelStart        =   9
         Value           =   9
      End
      Begin MSComctlLib.Slider SliderMin 
         Height          =   495
         Left            =   1800
         TabIndex        =   24
         Top             =   1800
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   873
         _Version        =   393216
         LargeChange     =   2
         Min             =   1
         Max             =   12
         SelStart        =   5
         TickStyle       =   1
         Value           =   5
      End
      Begin VB.Line Line1 
         BorderStyle     =   3  'Dot
         X1              =   0
         X2              =   6360
         Y1              =   2400
         Y2              =   2400
      End
      Begin VB.Label Label1 
         Caption         =   "Min"
         Height          =   255
         Left            =   600
         TabIndex        =   20
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label2 
         Caption         =   "Max"
         Height          =   255
         Left            =   2400
         TabIndex        =   19
         Top             =   840
         Width           =   615
      End
      Begin VB.Label Label3 
         Caption         =   "Interval"
         Height          =   255
         Left            =   4320
         TabIndex        =   18
         Top             =   840
         Width           =   975
      End
   End
   Begin VB.Frame Frame4 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   5
      Top             =   0
      Width           =   6375
      Begin VB.TextBox TextTitle 
         Height          =   270
         Left            =   1920
         TabIndex        =   14
         Top             =   240
         Width           =   4215
      End
      Begin VB.CheckBox Check1 
         Caption         =   "Check1"
         Height          =   255
         Left            =   240
         TabIndex        =   13
         Top             =   240
         Width           =   255
      End
      Begin VB.Label Label5 
         Height          =   255
         Left            =   3000
         TabIndex        =   7
         Top             =   240
         Width           =   1095
      End
      Begin VB.Label Label4 
         Caption         =   "Image Title: "
         Height          =   255
         Left            =   600
         TabIndex        =   6
         Top             =   285
         Width           =   1215
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   5520
      TabIndex        =   4
      Top             =   5400
      Width           =   975
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   4440
      TabIndex        =   3
      Top             =   5400
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      Caption         =   "Control Options"
      ForeColor       =   &H80000008&
      Height          =   735
      Left            =   120
      TabIndex        =   0
      Top             =   4440
      Width           =   2895
      Begin VB.OptionButton OptionManual 
         Caption         =   "Manual"
         Height          =   375
         Left            =   1560
         TabIndex        =   2
         Top             =   240
         Width           =   1095
      End
      Begin VB.OptionButton OptionAuto 
         Caption         =   "Auto"
         Height          =   375
         Left            =   240
         TabIndex        =   1
         Top             =   240
         Value           =   -1  'True
         Width           =   1095
      End
   End
End
Attribute VB_Name = "FrmMapValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public sngMin!, sngMax!, sngInterval!, strColor$, strTitle$
  
Dim tempHead$(1 To 2), tempLevel$() '文件头两行、所是等级行
Dim iLevels '等级文件中的等级数


Private Sub Form_Load()
  Dim LvlType$, LvlFile$, FlagTitle$
  
  '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  
  Me.Left = FrmMain.Left + 500
  Me.Top = FrmMain.Top + FrmMain.Height - Me.Height - 600
  
  TextTitle.Text = strTitle
  Call ReadWriteIni("Config.ini", "[Title_Value]", FlagTitle, "r")
  If FlagTitle = "Yes" Then
    Check1.Value = Checked
  ElseIf FlagTitle = "No" Then
    Check1.Value = Unchecked
  End If
  
  TextMin.Text = sngMin: TextMax.Text = sngMax: TextInterval.Text = sngInterval
'  TextInterval.SetFocus
  
  Call ReadWriteIni("Config.ini", "[LvlType_Value]", LvlType, "r")
  If LvlType = "Custom" Then
    OptionCustLvl.Value = True
  ElseIf LvlType = "Auto" Then
    OptionAutoLvl.Value = True
  End If
  Call ReadWriteIni("Config.ini", "[LvlFile_Value]", LvlFile, "r")
  TextCustLvl.Text = LvlFile
  
  If strColor = "BYR" Then OptionBYR.Value = True
  If strColor = "RYB" Then OptionRYB.Value = True
  
  TxtGrid_Size.Text = Grid_Size
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

  OptionAuto.Value = True
  
  If userID <> "liuw" Then
    Frame7.Visible = False
  End If
  
End Sub

Private Sub CmdCancel_Click()
  Unload Me
End Sub

'根据slider的值进行temp.lvl文件的制作
Private Sub CmdOK_Click()
  'On Error GoTo ErrMsg
  
  Dim SrfVisible As Boolean
  Dim i%, j%, k%
  Dim sngLevel!()
  Dim iMin%, iMax% ', iNum%
  Dim strTemp$() ', strLineTemp$()
   
'  Dim iWidth%, iHeight%
  Dim LvlFile$
  Dim LvlType$
  Dim FlagTitle$
  
  Dim FlagLable$  '是否显示数据标签：是-是、No-否
  
  If SetSurfer = "No" Then
    MsgBox "Please make sure Surfer11 is installed, and go to the CROSS menu: " & Chr(13) & "System >> Surfer Settings: Enable Surfer Plotting", vbInformation, "Notice": Exit Sub
  End If
  
  
  If GuestMode = True Then ' 如果是访客，又不曾确定是否已安装surfer11，则进行下面的判断，并加载 SrfApp 对象
    If FlagSrfInstalled = False Then
  
        Cancel = MsgBox("Confirm Surfer11 is installed?", vbYesNo + vbQuestion, "Notice")
        If Cancel = vbNo Then  '不确定
          Exit Sub
        ElseIf Cancel = vbYes Then '确认已安装
          FlagSrfInstalled = True
          
          FrmMain.StatusBar1.Panels(1).Text = "Loading Surfer plotting object": DoEvents
          Set SrfApp = CreateObject("Surfer.Application")
          
        End If
    End If
  End If
  
  If srfPath = "" Then '2026-08-18
    MsgBox "No SURFER template selected -- go to CROSS menu: System >> Surfer Settings to select one", vbInformation, "Notice": Exit Sub
  End If
  
  
  If Len(Dir(srfPath)) = 0 Then
    MsgBox "    Local SURFER Template File   " & District_ID & "_VALUE11.srf" & " does not exist, " & vbCrLf & _
           "请进入CROSS菜单：系统操作>>Surfer Settings，尝试重新下载", vbInformation, "Notice": Exit Sub
  End If
  If Check1.Value = Checked Then
    FlagTitle = "Yes"
  ElseIf Check1.Value = Unchecked Then
    FlagTitle = "No"
  End If
  Call ReadWriteIni("Config.ini", "[Title_Value]", FlagTitle, "w")
  
  If FlagTitle = "Yes" And TextTitle.Text <> "" Then
    strTitle = TextTitle.Text
  ElseIf FlagTitle = "Yes" And TextTitle.Text = "" Then
    MsgBox "Please enter a chart title", vbInformation, "Notice": Exit Sub
  End If
  
  If OptionCustLvl.Value = True Then
    LvlType = "Custom"
  ElseIf OptionAutoLvl.Value = True Then
    LvlType = "Auto"
  End If
  Call ReadWriteIni("Config.ini", "[LvlType_Value]", LvlType, "w")
  
  Grid_Size = TxtGrid_Size.Text
  iNumCols = CInt(100 * (sngxMax - sngxMin) / CSng(Grid_Size)): iNumRows = CInt(100 * (sngyMax - sngyMin) / CSng(Grid_Size))
  If ComboGrid_Method = ComboGrid_Method.List(0) Then
    Grid_Method = srfInverseDistance
  ElseIf ComboGrid_Method = ComboGrid_Method.List(1) Then
    Grid_Method = srfKriging
  ElseIf ComboGrid_Method = ComboGrid_Method.List(2) Then
    Grid_Method = srfMinCurvature
  ElseIf ComboGrid_Method = ComboGrid_Method.List(3) Then
    Grid_Method = srfNaturalNeighbor
  ElseIf ComboGrid_Method = ComboGrid_Method.List(4) Then
    Grid_Method = srfNearestNeighbor
  ElseIf ComboGrid_Method = ComboGrid_Method.List(5) Then
    Grid_Method = srfRegression
  ElseIf ComboGrid_Method = ComboGrid_Method.List(6) Then
    Grid_Method = srfRadialBasis
  ElseIf ComboGrid_Method = ComboGrid_Method.List(7) Then
    Grid_Method = srfTriangulation
  End If
  
  If OptionAutoLvl.Value = True Then 'Auto Color Levels
  
    If OptionBYR.Value = True Then strColor = "BYR"
    If OptionRYB.Value = True Then strColor = "RYB"
    
    sngMin = TextMin.Text: sngMax = TextMax.Text: sngInterval = TextInterval.Text
    iMin = SliderMin.Value:  iMax = SliderMax.Value ': iNum = iMax - iMin + 1
      
    ReDim sngLevel(iMin To iMax)
    
    For i = iMin To iMax
      sngLevel(i) = sngMin + sngInterval * (i - iMin)
    Next i

'****2015-8-28添加：读取lvl文件的头和各等级****
    FileNum = FreeFile
    Open lvlPath_Value For Input As #FileNum
    Do Until EOF(FileNum)
      j = 1
      For i = 1 To 2 '读取前两行文件头
        Line Input #FileNum, tempHead(j): j = j + 1
      Next i
      j = 1
      For i = 1 To 13
        ReDim Preserve tempLevel(1 To j)
        Line Input #FileNum, tempLevel(j): j = j + 1
      Next i
      
      iLevels = UBound(tempLevel)
      Exit Do
    Loop
    Close #FileNum
'****2015-8-28添加：读取lvl文件的头和各等级****
    
    FileNum = FreeFile '将读出的等级信息写入新文件
    Open App.Path & "\Temp\Temp.lvl" For Output As #FileNum
    Close #FileNum
    Open App.Path & "\Temp\Temp.lvl" For Append As #FileNum
      For i = 1 To 2
        Print #FileNum, tempHead(i)
      Next i
      If strColor = "BYR" Then 'Cool-Warm
        For i = iMin To iMax
          strTemp = Split(tempLevel(i), " ")
          strTemp(0) = sngLevel(i)
          tempLevel(i) = strTemp(0)
          For j = 1 To UBound(strTemp)
            tempLevel(i) = tempLevel(i) & " " & strTemp(j)
          Next j
          Print #FileNum, tempLevel(i)
        Next i
          
      ElseIf strColor = "RYB" Then 'Warm-Cool
        For i = iMin To iMax
          strTemp = Split(tempLevel(iLevels - i + 1), " ")
          strTemp(0) = sngLevel(i)
          tempLevel(iLevels - i + 1) = strTemp(0)
          For j = 1 To UBound(strTemp)
            tempLevel(iLevels - i + 1) = tempLevel(iLevels - i + 1) & " " & strTemp(j)
          Next j
          Print #FileNum, tempLevel(iLevels - i + 1)
        Next i
      End If
    Close #FileNum
  
  ElseIf OptionCustLvl.Value = True Then  'Custom Color Levels
    LvlFile = TextCustLvl.Text
    If LvlFile = "" Or Right(LvlFile, 3) <> "lvl" Then
      MsgBox ("Please select a *.lvl file"): Exit Sub
    End If
    If Dir(LvlFile) = "" Then
      MsgBox ("This *.lvl file does not exist"): Exit Sub
    End If
    Call ReadWriteIni("Config.ini", "[LvlFile_Value]", LvlFile, "w")
    
    FileCopy LvlFile, App.Path & "\Temp\Temp.lvl"
  End If
  FrmMain.StatusBar1.Panels(1).Text = "Temp.lvl file created"
  
  
  If OptionAuto.Value = True Then SrfVisible = False
  If OptionManual.Value = True Then SrfVisible = True
  
  If Check2.Value = Checked Then
    FlagLable = "Yes"
  ElseIf Check2.Value = Unchecked Then
    FlagLable = "No"
  End If
  
  
  Call DrawContourMap(sngMax, sngMin, sngInterval, SrfVisible, "VALUE", FlagTitle, strTitle, FlagLable)
   
End Sub

Private Sub CmdCustLvl_Click()
  Dim i%, j%
  
  CommonDialogOpen.Filter = "lvl(*.lvl)|*.lvl"
  CommonDialogOpen.InitDir = App.Path & "\Template\lvl"
  CommonDialogOpen.ShowOpen
  If CommonDialogOpen.Flags = 0 Then Screen.MousePointer = 1: FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub
  
  TextCustLvl.Text = CommonDialogOpen.FileName
  
End Sub

Private Sub OptionBYR_Click()
  PictureRYB.Visible = False: PictureBYR.Visible = True
End Sub

Private Sub OptionRYB_Click()
  PictureBYR.Visible = False: PictureRYB.Visible = True
End Sub

Private Sub OptionAutoLvl_Click()
  TextCustLvl.Enabled = False
  CmdCustLvl.Enabled = False
  TextMin.Enabled = True: TextMax.Enabled = True: TextInterval.Enabled = True
  OptionBYR.Enabled = True: OptionRYB.Enabled = True
  SliderMin.Enabled = True: SliderMax.Enabled = True
End Sub

Private Sub OptionCustLvl_Click()
  TextCustLvl.Enabled = True
  CmdCustLvl.Enabled = True
  TextMin.Enabled = False: TextMax.Enabled = False: TextInterval.Enabled = False
  OptionBYR.Enabled = False: OptionRYB.Enabled = False
  SliderMin.Enabled = False: SliderMax.Enabled = False
End Sub

Private Sub SliderMax_Scroll()
  Dim iMin%, iMax%, iNum%
  Dim sngMin!, sngMax!, sngInterval!
  sngMin = CSng(TextMin.Text): sngMax = CSng(TextMax.Text) ': sngInterval = CSng(TextInterval.Text)
  
  If SliderMax.Value <= SliderMin.Value Then
    SliderMin.Value = SliderMax.Value - 1
  End If
  iMin = SliderMin.Value: iMax = SliderMax.Value: iNum = iMax - iMin
  TextInterval.Text = Format((sngMax - sngMin) / iNum, "0.0")
End Sub

Private Sub SliderMin_Scroll()
  Dim sngMin!, sngMax!, sngInterval!
  sngMin = CSng(TextMin.Text): sngMax = CSng(TextMax.Text): sngInterval = CSng(TextInterval.Text)

  If SliderMin.Value + Int((sngMax - sngMin) / sngInterval) > SliderMax.Max Then
    SliderMax.Value = SliderMax.Max
    SliderMin.Value = SliderMax.Value - Int((sngMax - sngMin) / sngInterval)
  Else
    SliderMax.Value = SliderMin.Value + Int((sngMax - sngMin) / sngInterval)
  End If

End Sub

Private Sub TextInterval_LostFocus()
  Dim sngMin!, sngMax!, sngInterval!
  sngMin = CSng(TextMin.Text): sngMax = CSng(TextMax.Text): sngInterval = CSng(TextInterval.Text)
  If sngInterval = 0 Then MsgBox "Contour interval cannot be 0", vbInformation, "Notice": Exit Sub
  
  If SliderMin.Value + Int((sngMax - sngMin) / sngInterval) > SliderMax.Max Then
    SliderMax.Value = SliderMax.Max
    If SliderMax.Value - Int((sngMax - sngMin) / sngInterval) < SliderMin.Min Then
      MsgBox "Contour interval too small -- please reset", vbInformation, "Notice": Exit Sub ': SliderMin.Value = SliderMin.Min
    Else
      SliderMin.Value = SliderMax.Value - Int((sngMax - sngMin) / sngInterval)
    End If
  Else
    SliderMax.Value = SliderMin.Value + Int((sngMax - sngMin) / sngInterval)
  End If
End Sub

Private Sub TextMax_LostFocus()
  Dim sngMin!, sngMax!, sngInterval!
  sngMin = CSng(TextMin.Text): sngMax = CSng(TextMax.Text): sngInterval = CSng(TextInterval.Text)
  If sngInterval = 0 Then MsgBox "Contour interval cannot be 0", vbInformation, "Notice": Exit Sub
  
  If sngMax < sngMin Then MsgBox "Max cannot be less than Min -- please reset", vbInformation, "Notice": Exit Sub
  
  If SliderMin.Value + Int((sngMax - sngMin) / sngInterval) > SliderMax.Max Then
    SliderMax.Value = SliderMax.Max
    If SliderMax.Value - Int((sngMax - sngMin) / sngInterval) < SliderMin.Min Then
      MsgBox "Contour interval too small -- please reset", vbInformation, "Notice": Exit Sub ': SliderMin.Value = SliderMin.Min
    Else
      SliderMin.Value = SliderMax.Value - Int((sngMax - sngMin) / sngInterval)
    End If
  Else
    SliderMax.Value = SliderMin.Value + Int((sngMax - sngMin) / sngInterval)
  End If
 
End Sub


Private Sub TextMin_LostFocus()
  Dim sngMin!, sngMax!, sngInterval!
  sngMin = CSng(TextMin.Text): sngMax = CSng(TextMax.Text): sngInterval = CSng(TextInterval.Text)
  If sngInterval = 0 Then MsgBox "Contour interval cannot be 0", vbInformation, "Notice": Exit Sub
  
  If sngMax < sngMin Then MsgBox "Min cannot be greater than Max -- please reset", vbInformation, "Notice": Exit Sub
  
  If SliderMin.Value + Int((sngMax - sngMin) / sngInterval) > SliderMax.Max Then
    SliderMax.Value = SliderMax.Max
    If SliderMax.Value - Int((sngMax - sngMin) / sngInterval) < SliderMin.Min Then
      MsgBox "Contour interval too small -- please reset", vbInformation, "Notice": Exit Sub ': SliderMin.Value = SliderMin.Min
    Else
      SliderMin.Value = SliderMax.Value - Int((sngMax - sngMin) / sngInterval)
    End If
  Else
    SliderMax.Value = SliderMin.Value + Int((sngMax - sngMin) / sngInterval)
  End If
 
End Sub
