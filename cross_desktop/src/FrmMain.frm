VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "mscomctl.OCX"
Begin VB.MDIForm FrmMain 
   Appearance      =   0  'Flat
   BackColor       =   &H00FFFFFF&
   ClientHeight    =   15030
   ClientLeft      =   315
   ClientTop       =   900
   ClientWidth     =   28365
   Icon            =   "FrmMain.frx":0000
   LinkTopic       =   "MDIForm1"
   Begin VB.Timer Timer1 
      Left            =   960
      Top             =   5640
   End
   Begin VB.PictureBox Picture1 
      Align           =   1  'Align Top
      BackColor       =   &H00FFFFFF&
      Height          =   1215
      Left            =   0
      ScaleHeight     =   1155
      ScaleWidth      =   28305
      TabIndex        =   2
      Top             =   1080
      Width           =   28365
   End
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   720
      Top             =   3480
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.Toolbar Toolbar1 
      Align           =   1  'Align Top
      Height          =   1080
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   28365
      _ExtentX        =   50033
      _ExtentY        =   1905
      ButtonWidth     =   2117
      ButtonHeight    =   1905
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   31
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Hourly"
            Key             =   "KeyMeteoHor"
            ImageIndex      =   1
         EndProperty
         BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Daily"
            Key             =   "KeyMeteoDay"
            ImageIndex      =   2
         EndProperty
         BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Dekad"
            Key             =   "KeyMeteoTen"
            ImageIndex      =   3
         EndProperty
         BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Monthly"
            Key             =   "KeyMeteoMon"
            ImageIndex      =   4
         EndProperty
         BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Seasonal"
            Key             =   "KeyMeteoQtr"
            ImageIndex      =   5
         EndProperty
         BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Annual"
            Key             =   "KeyMeteoYer"
            ImageIndex      =   6
         EndProperty
         BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SngPeriod"
            Key             =   "KeyMeteoPeriodSin"
            ImageIndex      =   7
         EndProperty
         BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ArbPeriod"
            Key             =   "KeyMeteoPeriod"
            ImageIndex      =   8
         EndProperty
         BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "YerPeriod"
            Key             =   "KeyMeteoPeriodYer"
            ImageIndex      =   9
         EndProperty
         BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Conditional"
            Key             =   "KeyConditionQuery"
            ImageIndex      =   10
         EndProperty
         BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Season"
            Key             =   "KeyMeteoSeasons"
            ImageIndex      =   11
         EndProperty
         BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ColdAir"
            Key             =   "KeyMeteoColdAir"
            ImageIndex      =   12
         EndProperty
         BeginProperty Button16 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button17 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Station"
            Key             =   "KeyStation"
            ImageIndex      =   13
         EndProperty
         BeginProperty Button18 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Query"
            Key             =   "KeyQueryData"
            ImageIndex      =   14
         EndProperty
         BeginProperty Button19 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button20 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Contour"
            Key             =   "KeyContour"
            ImageIndex      =   15
         EndProperty
         BeginProperty Button21 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Histogram"
            Key             =   "KeyHistgram"
            ImageIndex      =   16
         EndProperty
         BeginProperty Button22 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "TimeChart"
            Key             =   "KeyTemporal"
            ImageIndex      =   17
         EndProperty
         BeginProperty Button23 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "Scatter"
            Key             =   "KeyScatter"
            ImageIndex      =   18
         EndProperty
         BeginProperty Button24 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button25 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "SaveData"
            Key             =   "KeySaveData"
            ImageIndex      =   19
         EndProperty
         BeginProperty Button26 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "ImportData"
            Key             =   "KeyImportData"
            ImageIndex      =   20
         EndProperty
         BeginProperty Button27 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button28 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "System"
            Key             =   "KeySystem"
            ImageIndex      =   21
            Style           =   5
            BeginProperty ButtonMenus {66833FEC-8583-11D1-B16A-00C0F0283628} 
               NumButtonMenus  =   10
               BeginProperty ButtonMenu1 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeySurfer"
                  Text            =   "Surfer Settings"
               EndProperty
               BeginProperty ButtonMenu2 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeySysConfig"
                  Text            =   "System Settings"
               EndProperty
               BeginProperty ButtonMenu3 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyPwdConfig"
                  Text            =   "Password Settings"
               EndProperty
               BeginProperty ButtonMenu4 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyUserInfo"
                  Text            =   "View Users"
               EndProperty
               BeginProperty ButtonMenu5 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyUserBehavior"
                  Text            =   "User Activity"
               EndProperty
               BeginProperty ButtonMenu6 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyAdvice"
                  Text            =   "Message Center"
               EndProperty
               BeginProperty ButtonMenu7 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyUpdate"
                  Text            =   "System Update"
               EndProperty
               BeginProperty ButtonMenu8 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyAbout"
                  Text            =   "About"
               EndProperty
               BeginProperty ButtonMenu9 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyLogOff"
                  Text            =   "Log Off"
               EndProperty
               BeginProperty ButtonMenu10 {66833FEE-8583-11D1-B16A-00C0F0283628} 
                  Key             =   "KeyQuit"
                  Text            =   "Exit"
               EndProperty
            EndProperty
         EndProperty
         BeginProperty Button29 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
         BeginProperty Button30 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Caption         =   "RegionalStat"
            Key             =   "KeyMeteoPeriodExt"
            ImageIndex      =   22
         EndProperty
         BeginProperty Button31 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Style           =   3
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.StatusBar StatusBar1 
      Align           =   2  'Align Bottom
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   14775
      Width           =   28365
      _ExtentX        =   50033
      _ExtentY        =   450
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   6
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel3 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   7056
            MinWidth        =   7056
         EndProperty
         BeginProperty Panel4 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel5 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   10583
            MinWidth        =   10583
         EndProperty
         BeginProperty Panel6 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            Object.Width           =   8819
            MinWidth        =   8819
         EndProperty
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   360
      Top             =   8040
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   48
      ImageHeight     =   48
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   22
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":3482
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":515C
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":6E36
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":8B10
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":A7EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":C4C4
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":E19E
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":FE78
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":11B52
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1382C
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":15506
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":171E0
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":18EBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1AB94
            Key             =   ""
         EndProperty
         BeginProperty ListImage15 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1C86E
            Key             =   ""
         EndProperty
         BeginProperty ListImage16 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":1E548
            Key             =   ""
         EndProperty
         BeginProperty ListImage17 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":20222
            Key             =   ""
         EndProperty
         BeginProperty ListImage18 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":21EFC
            Key             =   ""
         EndProperty
         BeginProperty ListImage19 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":23BD6
            Key             =   ""
         EndProperty
         BeginProperty ListImage20 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":258B0
            Key             =   ""
         EndProperty
         BeginProperty ListImage21 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":2758A
            Key             =   ""
         EndProperty
         BeginProperty ListImage22 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "FrmMain.frx":29264
            Key             =   ""
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "FrmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


'*************读取pdf文档******************
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" _
(ByVal hWnd As Long, ByVal lpOperation As String, _
ByVal lpFile As String, ByVal lpParameters As String, _
ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'*************读取pdf文档******************

Dim LogOff% '是否处于注销状态
'定义存储系统开始运行时所处的时间
Dim BeginTime As Date
  
Public Sub FormAdd()
  Dim i%
  FormNum = FormNum + 1
  If FormNum > 0 Then
    Picture1.Visible = False
  ElseIf FormNum = 0 Then
    Picture1.Visible = True
  End If
End Sub

Public Sub FormDel()
  Dim i%
  FormNum = FormNum - 1
  If FormNum > 0 Then
    Picture1.Visible = False
  ElseIf FormNum = 0 Then
    Picture1.Visible = True
  End If
End Sub

Private Sub MDIForm_Load()
  Dim i%, j%
  Dim StartTime As Date
  Dim Menu_Ban() As String
  
  
  '************记得每次更新的时候更新Version号**********************
  Vers_self = "2.14":  Vers_date = "2026.9.15"
  '************记得每次更新的时候更新Version号**********************
  
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2 - 300
  
  
  '2026-3-24 读取是否自动注销（隐藏设置）
  Call ReadWriteIni("Config.ini", "[Auto_LogOff]", Auto_LogOff, "r")
  
  '2026-3-24 定时Log Off
  If Auto_LogOff = "Yes" Then
    Timer1.Interval = 60000 '1分钟判断一 times
    Timer1.Enabled = True
  Else
    Timer1.Enabled = False
  End If
  
  If GuestMode = True Then
    Timer1.Enabled = False
  End If
    
  
  '系统名称
  Call ReadWriteIni("Config.ini", "[Caption]", Sys_Caption, "r")
  '单位名称
  Call ReadWriteIni("Config.ini", "[Department]", Department, "r")
  '主界面底图文件名
  Call ReadWriteIni("Config.ini", "[MainPic]", MainPic, "r")

  FormNum = 0 '初始化时打开的子窗体数量为0
  LogOff = vbNo
  
  Me.Caption = Sys_Caption & " " & Vers_self
  
  Dim clientW As Long, clientH As Long
  Call GetMDIClientSize(clientW, clientH)
  Picture1.Move 0, Toolbar1.Height, clientW, clientH - Toolbar1.Height - StatusBar1.Height
  Call Picture1_Paint
  
  FrmMain.StatusBar1.Panels(1).Text = "Please enter username and password"


End Sub


'2026-07-20 取得 MDI 主窗体客户区（不含标题栏、菜单、边框）的真实尺寸，单位为缇（twip），
'与 Toolbar1.Height/StatusBar1.Height 等控件属性一致，供 Picture1 精确填满可用区域使用。
Private Sub GetMDIClientSize(ByRef clientWidthTwips As Long, ByRef clientHeightTwips As Long)
  Dim rc As RECT
  Call GetClientRect(Me.hWnd, rc)
  clientWidthTwips = rc.Right * Screen.TwipsPerPixelX
  clientHeightTwips = rc.Bottom * Screen.TwipsPerPixelY
End Sub

' 2026-03-26 按照Deepseek建议修改如下处理过程
Private Sub MDIForm_Resize()
    On Error GoTo ResizeError
    
    ' 窗体最小化时忽略
    If Me.WindowState = 1 Then Exit Sub
    
    ' 校验控件是否存在
    If Toolbar1 Is Nothing Or StatusBar1 Is Nothing Or Picture1 Is Nothing Then Exit Sub
    
    ' 仅在窗体正常或最大化时才调整大小
    If Me.WindowState = 0 Or Me.WindowState = 2 Then

        Dim clientW As Long, clientH As Long
        Call GetMDIClientSize(clientW, clientH)
        If Toolbar1.Height > 0 And clientH > 0 And clientW > 0 Then
            Picture1.Move 0, Toolbar1.Height, clientW, clientH - Toolbar1.Height - StatusBar1.Height
            Call Picture1_Paint
        End If

        If clientW > 0 Then
            Call ResizeStatusBarPanelsProportionally(StatusBar1, clientW)
        End If

        Dim Frm As Object
        For Each Frm In Forms
            On Error Resume Next
            Call Frm.ApplyLayout
            On Error GoTo ResizeError
        Next Frm
    End If

    Exit Sub
    
ResizeError:
    Resume Next
End Sub


Private Sub MDIForm_Unload(Cancel As Integer)
'  On Error Resume Next
  Dim tempStr$
  
  If Dir(App.Path + "\*.bat") <> "" Then Kill App.Path + "\*.bat"
  If Dir(App.Path + "\Temp\*.bat") <> "" Then Kill App.Path + "\Temp\*.bat"
  If Dir(App.Path + "\Temp\*.ftp") <> "" Then Kill App.Path + "\Temp\*.ftp"
  
  If LogOff = vbYes Then
    
    '2026-07-22 访客模式下从未写入过登录日志，不需要也不允许连接数据库
    If Not GuestMode Then
      ORAConn4.Open
      strSQL = "update T_OTHE_CROSS_USERS_LOG set STATUS=0 where UserID= '" & userID & "' and UserIP='" & UserIP & "' and LOGTIME= to_date('" & Logtime & "', 'yyyy-mm-dd hh24:mi:ss')"
      ORAConn4.Execute strSQL
      ORAConn4.Close
      Set FrmMain = Nothing
    End If
    
    If SetSurfer = "Yes" Then
      Set SrfApp = Nothing
    End If
    
    frmStart.Show
    
    '2026-07-22 访客模式注销后复位状态，方便重新进入访客模式或正式登录
    GuestMode = False
    frmStart.CmdGuest.Enabled = True
    
    Call ReadWriteIni("Config.ini", "[User]", tempStr, "r")
    frmStart.txtUserID.Text = tempStr
    Call ReadWriteIni("Config.ini", "[Password]", tempStr, "r")
    If tempStr = "不保存" Then
      frmStart.txtPassword.Text = ""
    End If
    
  Else
    Cancel = MsgBox("Are you sure you want to exit?", vbYesNo + vbQuestion, "Notice")
    If Cancel = vbNo Then
      Exit Sub
    ElseIf Cancel = vbYes Then
      
      If Not GuestMode Then
        ORAConn4.Open
        strSQL = "update T_OTHE_CROSS_USERS_LOG set STATUS=0 where UserID= '" & userID & "' and UserIP='" & UserIP & "' and LOGTIME= to_date('" & Logtime & "', 'yyyy-mm-dd hh24:mi:ss')"
        ORAConn4.Execute strSQL
        ORAConn4.Close
        
      End If
        
      If SetSurfer = "Yes" Then
        Set SrfApp = Nothing
      End If
      
      Set FrmMain = Nothing
      End
    End If
  End If
End Sub


Public Sub Picture1_Paint()
  Dim pp As StdPicture
  If MainPic <> "" Then
    Set pp = LoadPicture(App.Path & "\Ini\" & MainPic)
    Picture1.PaintPicture pp, 0, 0, Picture1.Width, Picture1.Height
  Else
    Set pp = Nothing
  End If
End Sub


Private Sub Toolbar1_ButtonMenuClick(ByVal ButtonMenu As MSComctlLib.ButtonMenu)

  Select Case ButtonMenu.Key
    
    Case "KeyQuit"
      LogOff = vbNo
      Unload Me
    Case "KeyLogOff"
      LogOff = vbYes
      Unload Me
    Case "KeyAdvice"
      FrmCommunicate.Show vbNormal
      FrmCommunicate.SetFocus
    Case "KeyAbout"
      FrmAboutCROSS.Show vbNormal
      FrmAboutCROSS.SetFocus
    Case "KeyUserInfo"
      FrmUserInfo.Show vbNormal
      FrmUserInfo.SetFocus
    Case "KeyUserBehavior"
      frmWait.Label1.Caption = "Querying, please wait"
      frmWait.Show: DoEvents
      FrmUserBehavior.Show vbNormal
      FrmUserBehavior.SetFocus
      frmWait.Hide: DoEvents
      
    Case "KeyUpdate"
      FrmUpdate.Show vbNormal
      FrmUpdate.SetFocus
    Case "KeySurfer"
      FrmConfigSurfer.Show vbNormal
      FrmConfigSurfer.SetFocus
    Case "KeyPwdConfig"
      FrmPassword.Show vbNormal
      FrmPassword.SetFocus
    Case "KeySysConfig"
      FrmConfig.Show vbNormal
      FrmConfig.SetFocus

    End Select
      
End Sub

Private Sub Toolbar1_ButtonClick(ByVal Button As MSComctlLib.Button)
  Select Case Button.Key
    Case "KeyImportData"
      FrmImportData.Show vbNormal
      FrmImportData.SetFocus
    Case "KeyConditionQuery"
      FrmConditionQuery.Show vbNormal
      FrmConditionQuery.SetFocus
    Case "KeyMeteoHor"
      FrmMeteoHour.Show vbNormal
      FrmMeteoHour.SetFocus
    Case "KeyMeteoDay"
      FrmMeteoDay.Show vbNormal
      FrmMeteoDay.SetFocus
    Case "KeyMeteoTen"
      FrmMeteoTen.Show vbNormal
      FrmMeteoTen.SetFocus
    Case "KeyMeteoMon"
      FrmMeteoMon.Show vbNormal
      FrmMeteoMon.SetFocus
    Case "KeyMeteoQtr"
      FrmMeteoQtr.Show vbNormal
      FrmMeteoQtr.SetFocus
    Case "KeyMeteoYer"
      FrmMeteoYer.Show vbNormal
      FrmMeteoYer.SetFocus
    Case "KeyMeteoPeriod"
      FrmMeteoPeriod.Show vbNormal
      FrmMeteoPeriod.SetFocus
    Case "KeyMeteoPeriodYer"
      FrmMeteoPeriodYer.Show vbNormal
      FrmMeteoPeriodYer.SetFocus

    Case "KeyMeteoPeriodExt"
      FrmMeteoPeriodExt.Show vbNormal
      FrmMeteoPeriodExt.SetFocus

    Case "KeyMeteoPeriodSin"
      FrmMeteoPeriodSin.Show vbNormal
      FrmMeteoPeriodSin.SetFocus
    
    Case "KeyStation"
      FrmStation.Show vbNormal
      FrmStation.SetFocus
      
    Case "KeyHistgram"
      Call ChartData.ShowHistogramForActiveForm
    Case "KeyTemporal"
      Call ChartData.ShowTemporalChartForActiveForm
    Case "KeyScatter"
      Call ChartData.ShowScatterForActiveForm
      
    Case "KeyMeteoColdAir"
      FrmMeteoColdAir.Show vbNormal
      FrmMeteoColdAir.SetFocus
      
    Case "KeyMeteoSeasons"
      FrmMeteoSeasons.Show vbNormal
      FrmMeteoSeasons.SetFocus
            
      
      
    Case "KeyQueryData"
      If GuestMode Then MsgBox "Data query is unavailable in Guest Mode. Please log in to use this feature.", vbInformation, "Notice": Exit Sub
      
      '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
      '2024-12-6 添加判断-是否离线，以便可以在服务器端修改用户在线状态，迫使其下线重新登录（适用于账号密  码被盗情况，或定期清理伪在线终端-如弹出错误提示后依然显示在线、实际并不在线的终端）。
      ORAConn4.Open
      ORAConn4.CursorLocation = 3 '****************很重要！此代码可以使取得ORARst的recordcount属性！！************************
        
      strSQL = "select * from T_OTHE_CROSS_USERS_LOG where UserID= '" & userID & "' and UserIP='" & UserIP & "' and LOGTIME= to_date('" & Logtime & "', 'yyyy-mm-dd hh24:mi:ss')"
      ORARst.CursorLocation = adUseClient
      ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
        
      If ORARst!Status = 0 Then
        ORARst.Close
        ORAConn4.Close
        MsgBox "You have been logged off. Please log in again.", vbInformation, "Notice"
        Exit Sub
      Else
        ORARst.Close
        ORAConn4.Close
      End If
      '= = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
    
      If FormNum = 0 Then
        MsgBox "Please open a query window first", vbInformation, "Notice"
      Else
        Call QueryData
      End If
    Case "KeyContour"
      If FormNum = 0 Then
        MsgBox "Please open a query window first", vbInformation, "Notice"
      Else
        Call Contour
      End If
    Case "KeySaveData"
      If FormNum = 0 Then
        MsgBox "Please open a query window first", vbInformation, "Notice"
      Else
        Call SaveData
      End If
  End Select

End Sub

Private Sub QueryData()
  
  If FormNum = 0 Then '如果没是打开的子窗体，则什么都不干
    
  Else '如果是打开的子窗体，则判断是哪一个子窗体
    
    Dim QueryType$, Querytime As Date
 
    Querytime = Format(Now, "yyyy-mm-dd hh:mm:ss")
    
    If FrmMain.ActiveForm Is FrmMeteoHour Then
      QueryType = "Hour"
    ElseIf FrmMain.ActiveForm Is FrmMeteoDay Then
      QueryType = "Day"
    ElseIf FrmMain.ActiveForm Is FrmMeteoTen Then
      QueryType = "Ten"
    ElseIf FrmMain.ActiveForm Is FrmMeteoMon Then
      QueryType = "Mon"
    ElseIf FrmMain.ActiveForm Is FrmMeteoQtr Then
      QueryType = "Quarter"
    ElseIf FrmMain.ActiveForm Is FrmMeteoYer Then
      QueryType = "Year"
    ElseIf FrmMain.ActiveForm Is FrmMeteoPeriod Then
      QueryType = "Period"
    ElseIf FrmMain.ActiveForm Is FrmMeteoPeriodSin Then
      QueryType = "SinglePeriod"
    ElseIf FrmMain.ActiveForm Is FrmMeteoPeriodYer Then
      QueryType = "MultiPeriod"
    ElseIf FrmMain.ActiveForm Is FrmMeteoColdAir Then
      QueryType = "ColdAir"
    ElseIf FrmMain.ActiveForm Is FrmMeteoSeasons Then
      QueryType = "Season"
    ElseIf FrmMain.ActiveForm Is FrmConditionQuery Then
      QueryType = "Condition"
    End If
    
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
  
    strSQL = "insert into T_OTHE_CROSS_DATA_QUERY(USERID,DDATETIME,QUERY_IP,QUERY_TYPE)"
    strSQL = strSQL + " Values('" & userID & "',to_date('" & Querytime & "', 'yyyy-mm-dd hh24:mi:ss'),'" & UserIP & "','" & QueryType & "')"
    
    ORAConn4.Execute strSQL
    ORAConn4.Close
    
    
    If FrmMain.ActiveForm Is FrmMeteoHour Then Call FrmMeteoHour.CmdMeteoHour
    If FrmMain.ActiveForm Is FrmMeteoDay Then Call FrmMeteoDay.CmdMeteoDay
    If FrmMain.ActiveForm Is FrmMeteoTen Then Call FrmMeteoTen.CmdMeteoTen
    If FrmMain.ActiveForm Is FrmMeteoMon Then Call FrmMeteoMon.CmdMeteoMon
    If FrmMain.ActiveForm Is FrmMeteoQtr Then Call FrmMeteoQtr.CmdMeteoQtr
    If FrmMain.ActiveForm Is FrmMeteoYer Then Call FrmMeteoYer.CmdMeteoYer
    If FrmMain.ActiveForm Is FrmMeteoPeriod Then
      If MultiSel = False Then Call FrmMeteoPeriod.CmdMeteoPeriod '单要素统计
      If MultiSel = True Then Call FrmMeteoPeriod.CmdMeteoPeriod2 '多要素统计
    End If
    If FrmMain.ActiveForm Is FrmMeteoPeriodSin Then Call FrmMeteoPeriodSin.CmdMeteoPeriodSin
    If FrmMain.ActiveForm Is FrmMeteoPeriodYer Then Call FrmMeteoPeriodYer.CmdMeteoPeriodYer
    If FrmMain.ActiveForm Is FrmMeteoColdAir Then Call FrmMeteoColdAir.CmdMeteoColdAir
    If FrmMain.ActiveForm Is FrmMeteoSeasons Then Call FrmMeteoSeasons.CmdMeteoSeasons
    If FrmMain.ActiveForm Is FrmConditionQuery Then Call FrmConditionQuery.cmdConditionQuery
  
  End If

End Sub

Private Sub Contour()
  If FormNum = 0 Then '如果没是打开的子窗体，则什么都不干
    
  Else '如果是打开的子窗体，则判断是哪一个子窗体
    If FrmMain.ActiveForm Is FrmMeteoHour Then Call FrmMeteoHour.CmdMap_MeteoHor
    If FrmMain.ActiveForm Is FrmMeteoDay Then Call FrmMeteoDay.CmdMap_MeteoDay
    If FrmMain.ActiveForm Is FrmMeteoTen Then Call FrmMeteoTen.CmdMap_MeteoTen
    If FrmMain.ActiveForm Is FrmMeteoMon Then Call FrmMeteoMon.CmdMap_MeteoMon
    If FrmMain.ActiveForm Is FrmMeteoQtr Then Call FrmMeteoQtr.CmdMap_MeteoQtr
    If FrmMain.ActiveForm Is FrmMeteoYer Then Call FrmMeteoYer.CmdMap_MeteoYer
    If FrmMain.ActiveForm Is FrmMeteoPeriod Then Call FrmMeteoPeriod.CmdMap_MeteoPeriod
    If FrmMain.ActiveForm Is FrmMeteoPeriodSin Then Call FrmMeteoPeriodSin.CmdMap_MeteoPeriodSin
    If FrmMain.ActiveForm Is FrmMeteoPeriodYer Then Call FrmMeteoPeriodYer.CmdMap_MeteoPeriodYer
    If FrmMain.ActiveForm Is FrmMeteoColdAir Then Call FrmMeteoColdAir.CmdMap_MeteoColdAir
    If FrmMain.ActiveForm Is FrmMeteoSeasons Then Call FrmMeteoSeasons.CmdMap_MeteoSeasons
    If FrmMain.ActiveForm Is FrmImportData Then Call FrmImportData.CmdMap_ImportData
  End If

End Sub


Private Sub SaveData()
  If FormNum = 0 Then '如果没是打开的子窗体，则什么都不干
  
  Else '如果是打开的子窗体，则判断是哪一个子窗体
    If FrmMain.ActiveForm Is FrmMeteoHour Then Call FrmMeteoHour.CmdOutput_MeteoHor
    If FrmMain.ActiveForm Is FrmMeteoDay Then Call FrmMeteoDay.CmdOutput_MeteoDay
    If FrmMain.ActiveForm Is FrmMeteoTen Then Call FrmMeteoTen.CmdOutput_MeteoTen
    If FrmMain.ActiveForm Is FrmMeteoMon Then Call FrmMeteoMon.CmdOutput_MeteoMon
    If FrmMain.ActiveForm Is FrmMeteoQtr Then Call FrmMeteoQtr.CmdOutput_MeteoQtr
    If FrmMain.ActiveForm Is FrmMeteoYer Then Call FrmMeteoYer.CmdOutput_MeteoYer
    If FrmMain.ActiveForm Is FrmMeteoPeriod Then Call FrmMeteoPeriod.CmdOutput_MeteoPeriod
    If FrmMain.ActiveForm Is FrmMeteoPeriodSin Then Call FrmMeteoPeriodSin.CmdOutput_MeteoPeriodSin
    If FrmMain.ActiveForm Is FrmMeteoPeriodYer Then Call FrmMeteoPeriodYer.CmdOutput_MeteoPeriodYer
    If FrmMain.ActiveForm Is FrmMeteoColdAir Then Call FrmMeteoColdAir.CmdOutput_MeteoColdAir
    If FrmMain.ActiveForm Is FrmMeteoSeasons Then Call FrmMeteoSeasons.CmdOutput_MeteoSeasons
    If FrmMain.ActiveForm Is FrmConditionQuery Then Call FrmConditionQuery.CmdSaveConditionQuery
  End If
End Sub



Private Sub Timer1_Timer()
    ' On Error Resume Next
    Dim tempStr$
    Static bExecuted As Boolean
    
    ' 判断是否为周日凌晨1:00-2:00
    If Weekday(Now) = vbSunday And Format(Now, "HH:MM") >= "01:00" And Format(Now, "HH:MM") < "01:05" Then
        If Not bExecuted Then
            ' 执行处理代码块
            If Dir(App.Path + "\*.bat") <> "" Then Kill App.Path + "\*.bat"
            If Dir(App.Path + "\Temp\*.bat") <> "" Then Kill App.Path + "\Temp\*.bat"
            If Dir(App.Path + "\Temp\*.ftp") <> "" Then Kill App.Path + "\Temp\*.ftp"
            
            ORAConn4.Open
            strSQL = "update T_OTHE_CROSS_USERS_LOG set STATUS=0 where UserID= '" & userID & "' and UserIP='" & UserIP & "' and LOGTIME= to_date('" & Logtime & "', 'yyyy-mm-dd hh24:mi:ss')"
            ORAConn4.Execute strSQL
            ORAConn4.Close
            
            If SetSurfer = "Yes" Then
                Set SrfApp = Nothing
            End If
                        
            FrmMain.Enabled = False
            
            frmStart.Show
            
            Call ReadWriteIni("Config.ini", "[User]", tempStr, "r")
            frmStart.txtUserID.Text = tempStr
            Call ReadWriteIni("Config.ini", "[Password]", tempStr, "r")
            If tempStr = "不保存" Then
                frmStart.txtPassword.Text = ""
            End If
            
            frmWait.Label1.Caption = "Logged off. Please log in again."
            frmWait.Show: DoEvents
            
            bExecuted = True  ' 标记为已执行，防止重复执行
        End If
    Else
        bExecuted = False   ' 过了周日1:00之后，重置标记
    End If
End Sub
