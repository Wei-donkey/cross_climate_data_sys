VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmConfig 
   Caption         =   "System Settings"
   ClientHeight    =   4095
   ClientLeft      =   9750
   ClientTop       =   3390
   ClientWidth     =   6390
   Icon            =   "FrmConfig.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4095
   ScaleWidth      =   6390
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Climate Baseline Source"
      Height          =   735
      Left            =   120
      TabIndex        =   15
      Top             =   1080
      Width           =   6135
      Begin VB.OptionButton OptionNorm_Src 
         BackColor       =   &H8000000E&
         Caption         =   "Climate Center Calculated Data"
         Height          =   255
         Index           =   1
         Left            =   2880
         TabIndex        =   17
         ToolTipText     =   "Computes climate normals in real time for the selected normal-period years"
         Top             =   360
         Width           =   3135
      End
      Begin VB.OptionButton OptionNorm_Src 
         BackColor       =   &H8000000E&
         Caption         =   "CMA-Issued Data"
         Height          =   255
         Index           =   0
         Left            =   600
         TabIndex        =   16
         ToolTipText     =   "CMA-issued data covers only the most recent 30-year normals"
         Top             =   360
         Width           =   1935
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00FFFFFF&
      Caption         =   "Climate Baseline Period"
      Height          =   855
      Left            =   120
      TabIndex        =   10
      Top             =   120
      Width           =   6135
      Begin VB.ComboBox Comboyears_AWST 
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
         Left            =   4680
         TabIndex        =   14
         Text            =   "yyyy-yyyy"
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox Comboyears_SURF 
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
         Left            =   1560
         TabIndex        =   13
         Text            =   "Comboyears_SURF"
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label6 
         BackColor       =   &H00FFFFFF&
         Caption         =   "National Station："
         Height          =   375
         Left            =   600
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
      Begin VB.Label Label5 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Regional Station："
         Height          =   375
         Left            =   3720
         TabIndex        =   11
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00FFFFFF&
      Caption         =   "System Properties"
      Height          =   1575
      Left            =   120
      TabIndex        =   2
      Top             =   1920
      Width           =   6135
      Begin VB.CommandButton CmdMain 
         Caption         =   "…"
         Height          =   255
         Left            =   5640
         TabIndex        =   9
         Top             =   1080
         Width           =   375
      End
      Begin VB.TextBox TxtMainPath 
         Height          =   270
         Left            =   1560
         TabIndex        =   7
         Top             =   1080
         Width           =   3975
      End
      Begin VB.TextBox TxtDepartment 
         Height          =   270
         Left            =   1560
         TabIndex        =   6
         Top             =   720
         Width           =   4455
      End
      Begin VB.TextBox TxtSys_Caption 
         Height          =   270
         Left            =   1560
         TabIndex        =   5
         Top             =   360
         Width           =   4455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFFFFF&
         Caption         =   "Win Bkground: "
         Height          =   255
         Left            =   120
         TabIndex        =   8
         Top             =   1080
         Width           =   1335
      End
      Begin VB.Label Label2 
         BackColor       =   &H00FFFFFF&
         Caption         =   "User's Unit:"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   720
         Width           =   1455
      End
      Begin VB.Label Label1 
         BackColor       =   &H00FFFFFF&
         Caption         =   "System Name: "
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
   End
   Begin VB.CommandButton CmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5400
      TabIndex        =   1
      Top             =   3600
      Width           =   855
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   4440
      TabIndex        =   0
      Top             =   3600
      Width           =   855
   End
   Begin MSComDlg.CommonDialog CommonDialogOpen 
      Left            =   120
      Top             =   2760
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
End
Attribute VB_Name = "FrmConfig"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim strMainPic$
Dim strFileName$

Private Sub CmdCancel_Click()
  Unload Me
End Sub

Private Sub CmdMain_Click()
  Dim i%, j%
  
  CommonDialogOpen.Filter = "jpg(*.jpg)|*.jpg"
  CommonDialogOpen.InitDir = App.Path & "\Ini"
  CommonDialogOpen.ShowOpen
  If CommonDialogOpen.Flags = 0 Then Screen.MousePointer = 1: FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub
  
  strFileName = CommonDialogOpen.FileName
  MainPic = CommonDialogOpen.FileTitle
  TxtMainPath.Text = strFileName

End Sub

Private Sub CmdOK_Click()
  Dim tempFlag$ '标识：当父模块已选，任意子模块已选，则值为True
  
  Sys_Caption = TxtSys_Caption.Text
  Department = TxtDepartment.Text
  Call ReadWriteIni("Config.ini", "[Caption]", Sys_Caption, "w")
  Call ReadWriteIni("Config.ini", "[Department]", Department, "w")
  
  strFileName = TxtMainPath.Text
  If strFileName <> "" Then
    If strFileName <> App.Path & "\Ini\" & MainPic Then '选择的jpg不在Ini文件夹内
      If Len(Dir(App.Path & "\Ini\" & MainPic)) <> 0 Then  'Ini内是同名文件
        Kill App.Path & "\Ini\" & MainPic
      End If
      FileCopy strFileName, App.Path & "\Ini\" & MainPic
    End If
  End If
  
  If TxtMainPath.Text = "" Then MainPic = "" '可以删除主界面底图
  Call ReadWriteIni("Config.ini", "[MainPic]", MainPic, "w")


  Call FrmMain.Picture1_Paint

  Years_NormSURF = Comboyears_SURF.Text
  Years_NormAWST = Comboyears_AWST.Text
  Call ReadWriteIni("Config.ini", "[Years_NormSURF]", Years_NormSURF, "w")
  Call ReadWriteIni("Config.ini", "[Years_NormAWST]", Years_NormAWST, "w")
  If OptionNorm_Src(0).Value = True Then
    Norm_Src = "CMA"
  ElseIf OptionNorm_Src(1).Value = True Then
    Norm_Src = "GRMC"
  End If

  If OptionNorm_Src(0).Value = True Then
    Norm_Src = "CMA"
  ElseIf OptionNorm_Src(1).Value = True Then
    Norm_Src = "GRMC"
  End If
  Call ReadWriteIni("Config.ini", "[Norm_Src]", Norm_Src, "w")
          
  If Norm_Src = "CMA" Then
    FrmMain.StatusBar1.Panels(4).Text = "National station normals: CMA-issued normals"

  ElseIf Norm_Src = "GRMC" Then
    FrmMain.StatusBar1.Panels(4).Text = "National station normals: Climate Center " & Years_NormSURF & "Yr Avg"
    If FrmMeteoPeriodExtStatus = "OPEN" Then
      FrmMeteoPeriodExt.StatusBar1.Panels(3).Text = "National station normals: Climate Center " & Years_NormSURF & "Yr Avg"
    End If
  End If
  FrmMain.StatusBar1.Panels(5).Text = "AWS station normals: Climate Center " & Years_NormAWST & "Yr Avg"
  
  
  Unload Me
  
End Sub

Private Sub Form_Load()
  Dim str1$, i%

'   '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  
  Me.Left = FrmMain.Left + 500
  Me.Top = FrmMain.Top + FrmMain.Height - Me.Height - 600
  
  Comboyears_SURF.Clear
  Dim strYear_stt$, strYear_end$
  strYear_stt = CStr(Year(Date) - 30)
  strYear_end = CStr(Year(Date) - 1)
    
  Comboyears_SURF.AddItem strYear_stt & "-" & strYear_end
    
  Comboyears_SURF.AddItem "1971-2000"
  Comboyears_SURF.AddItem "1981-2010"
  Comboyears_SURF.AddItem "1991-2020"
  Comboyears_SURF.Text = Years_NormSURF
  
  Comboyears_AWST.Clear
  Comboyears_AWST.AddItem Years_NormAWST
  Comboyears_AWST.Text = Years_NormAWST
    
  If Norm_Src = "CMA" Then
    OptionNorm_Src(0).Value = True
  ElseIf Norm_Src = "GRMC" Then
    OptionNorm_Src(1).Value = True
  End If
  
  TxtSys_Caption.Text = Sys_Caption
  TxtDepartment.Text = Department
  If MainPic <> "" Then TxtMainPath.Text = App.Path & "\Ini\" & MainPic
  
  strFileName = TxtMainPath.Text

End Sub

Private Sub OptionNorm_Src_Click(Index As Integer)
  If Index = 0 Then
    Comboyears_SURF.Text = Comboyears_SURF.List(Comboyears_SURF.ListCount - 1)
    Comboyears_SURF.Enabled = False
    Comboyears_AWST.Enabled = False
  ElseIf Index = 1 Then
    Comboyears_SURF.Enabled = True
    Comboyears_AWST.Enabled = True
  End If
  
End Sub
