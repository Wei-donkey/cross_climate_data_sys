VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "richtx32.Ocx"
Begin VB.Form FrmUpdate 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "System Update"
   ClientHeight    =   5355
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   9075
   Icon            =   "FrmUpdate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   5355
   ScaleWidth      =   9075
   Begin VB.ComboBox ComboVersion 
      Appearance      =   0  'Flat
      Height          =   300
      Left            =   8040
      TabIndex        =   12
      Top             =   1320
      Width           =   850
   End
   Begin VB.CommandButton CmdUpload 
      Caption         =   "Upload New Version"
      Height          =   375
      Left            =   3840
      TabIndex        =   10
      Top             =   4920
      Width           =   1815
   End
   Begin VB.CheckBox CheckUpdate 
      Caption         =   "update when a new version available"
      Height          =   255
      Left            =   240
      TabIndex        =   9
      Top             =   4920
      Width           =   3615
   End
   Begin VB.CommandButton CmdLater 
      Caption         =   "Remind Me Later"
      Height          =   375
      Left            =   7320
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   4920
      Width           =   1575
   End
   Begin VB.CommandButton CmdUpdate 
      Caption         =   "System Update"
      Height          =   375
      Left            =   5760
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4920
      Width           =   1455
   End
   Begin RichTextLib.RichTextBox TextWhatsNew 
      Height          =   3135
      Left            =   240
      TabIndex        =   2
      Top             =   1680
      Width           =   8655
      _ExtentX        =   15266
      _ExtentY        =   5530
      _Version        =   393217
      BorderStyle     =   0
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      Appearance      =   0
      TextRTF         =   $"FrmUpdate.frx":3482
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
   Begin VB.Label Label6 
      Caption         =   "You can also open http://10.148.15.110/cross to download the installer and reinstall."
      Height          =   255
      Left            =   1200
      TabIndex        =   13
      Top             =   480
      Width           =   7695
   End
   Begin VB.Label Label5 
      Caption         =   "Release notes for all versions: "
      Height          =   255
      Left            =   4440
      TabIndex        =   11
      Top             =   1440
      Width           =   3375
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   240
      X2              =   9000
      Y1              =   1200
      Y2              =   1200
   End
   Begin VB.Label Label4 
      Caption         =   "Latest version release notes: "
      Height          =   255
      Left            =   240
      TabIndex        =   8
      Top             =   1440
      Width           =   3855
   End
   Begin VB.Label Version_Old 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   8040
      TabIndex        =   7
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Version_New 
      BackColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   2160
      TabIndex        =   6
      Top             =   840
      Width           =   795
   End
   Begin VB.Label Label3 
      Caption         =   "Update Notice"
      Height          =   255
      Left            =   1200
      TabIndex        =   5
      Top             =   240
      Width           =   5535
   End
   Begin VB.Image Image1 
      Height          =   720
      Left            =   240
      Picture         =   "FrmUpdate.frx":351F
      Top             =   240
      Width           =   720
   End
   Begin VB.Label Label2 
      Caption         =   "Currently installed version: "
      Height          =   195
      Left            =   6600
      TabIndex        =   1
      Top             =   885
      Width           =   1455
   End
   Begin VB.Label Label1 
      Caption         =   "Latest Version: "
      Height          =   195
      Left            =   1200
      TabIndex        =   0
      Top             =   885
      Width           =   975
   End
End
Attribute VB_Name = "FrmUpdate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim LocalPath_exe$, LocalPath_txt$ 'CROSS_Installer下载本地路径和WhatsNew下载本地路径


Private Sub CheckUpdate_Click()
  If CheckUpdate.Value = Unchecked Then
    Update = "Off"
  Else
    Update = "On"
  End If
  Call ReadWriteIni("Config.ini", "[Update]", Update, "w")
End Sub

Private Sub CmdUpload_Click()
  FrmUpload.Show
  FrmUpload.SetFocus
'  Unload Me
'  Exit Sub
End Sub

Private Sub ComboVersion_Click()
 Call showVersionInfo(ComboVersion.Text)
End Sub

Private Sub Form_Load()
   '使窗体始终居前
  
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  
  Me.Left = FrmMain.Left + FrmMain.Width - Me.Width - 500
  Me.Top = FrmMain.Top + FrmMain.Height - Me.Height - 600
  
  If userID <> "liuw" Then '超级管理员
    CmdUpload.Visible = False
  End If

  If Update = "On" Then CheckUpdate.Value = Checked
    
  '2026-07-22 访客模式下禁止连接数据库，不检查系统更新
  If Not GuestMode Then
    If Vers_New = "" Then '如果系统没是在启动时读取最新Version号，那么在此读取
      FrmMain.StatusBar1.Panels(1).Text = "Checking for the latest CROSS version"

      ORAConn4.Open
      ORAConn4.CursorLocation = 3 '****************很重要！此代码可以使取得ORARst的recordcount属性！！************************
      
      strSQL = "select max(version) version_new from T_OTHE_CROSS_UPDATE_FILE"
      ORARst.CursorLocation = adUseClient
      ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
      
      If ORARst.EOF <> True Then Vers_New = ORARst!Version_New
      
      ORARst.Close
      ORAConn4.Close
          
      Update_day = Format(Date, "yyyy-mm-dd")
      Call ReadWriteIni("Config.ini", "[Update_day]", Update_day, "w")
    End If
    
    Version_New.Caption = Vers_New
    Version_Old.Caption = Vers_self
    
    CmdUpdate.Enabled = False
    CmdLater.Enabled = False
    If CSng(Vers_self) >= CSng(Vers_New) Then
      Label3.Caption = "The installed CROSS is already up to date -- no update needed"
      Label6.Visible = False
    
    ElseIf CSng(Vers_self) < CSng(Vers_New) Then
      Label3.Caption = "A new version of CROSS is available -- please click the System Update button!"
      Label6.Visible = True
      CmdLater.Enabled = True
      CmdUpdate.Enabled = True
    End If
    
    ORAConn4.Open
    ORAConn4.CursorLocation = 3 '****************很重要！此代码可以使取得ORARst的recordcount属性！！************************
    
    strSQL = "select distinct version from T_OTHE_CROSS_UPDATE_FILE order by version desc"
    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
    
    If ORARst.EOF <> True Then
      ORARst.MoveFirst
      Do Until ORARst.EOF
        ComboVersion.AddItem ORARst!version
        ORARst.MoveNext
      Loop
    End If
    ORARst.Close
    ORAConn4.Close
    ComboVersion.Text = ComboVersion.List(0)
    
    Call showVersionInfo(ComboVersion.List(0))
    
    FrmMain.StatusBar1.Panels(1).Text = "CROSS version check complete"
    
    CmdUpdate.Enabled = True
  Else
    TextWhatsNew.Text = "System UpdateContent："

    Call LoadGuestTxt(TextWhatsNew, "VersionInfo.csv")
    ComboVersion.Enabled = False
    Version_New.Caption = ""
    Version_Old.Caption = Vers_self
    Label3.Caption = "System update check is not available in Guest Mode"
    CmdUpdate.Enabled = False
    CmdLater.Enabled = False
  End If
  
End Sub

Private Sub showVersionInfo(Vers_Show$)
  Dim i%, tempStr$
  TextWhatsNew.Text = ""
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3 '****************很重要！此代码可以使取得ORARst的recordcount属性！！************************
    
  strSQL = "select * from T_OTHE_CROSS_VERSION_INFO where version='" & Vers_Show & "' order by ddatetime desc"
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
    
  If ORARst.EOF <> True Then
    TextWhatsNew.Text = Vers_Show & "Version更新Content："
    ORARst.MoveFirst: i = 0
    Do Until ORARst.EOF
      tempStr = Chr(10) & Str(i + 1) & "、" & ORARst!content
      TextWhatsNew.Text = TextWhatsNew.Text & tempStr
      ORARst.MoveNext: i = i + 1
    Loop
      
  End If
    
  ORARst.Close
  ORAConn4.Close
End Sub
  
Private Sub CmdUpdate_Click()
  Dim i%, Response%
 
   
  Response = MsgBox("The system will close automatically, then update and relaunch. Continue?", vbOKCancel)
  If Response = 2 Then Exit Sub
  
  ' ******* 写 getFTP_update.bat 批处理程序，依次关闭CROSS、通过ftp下载相应升级文件、打开CROSS
  ' ********* 1）关闭 CROSS *******
  FileNum = FreeFile
  Open App.Path & "\update_CROSS.bat" For Output As #FileNum
    Print #FileNum, "taskkill /f /im CROSS.exe"
  Close #FileNum

  '********** 2) 通过数据库看需要下载哪些升级文件（包括 CROSS.exe），写入 getFTP_*.ftp 文件中 ************
  Dim tempFTPPath$, tempLocalPath$, tempFTPFile$
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
    
  strSQL = "select * from T_OTHE_CROSS_UPDATE_FILE where version='" & Vers_New & "'"
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
    
  If ORARst.EOF <> True Then
    ORARst.MoveFirst
    Do Until ORARst.EOF
      If Not IsNull(ORARst!file_path) Then
        tempFTPPath = ORARst!file_path
      Else
        tempFTPPath = ""
      End If
      tempLocalPath = App.Path & ORARst!file_path
      tempFTPFile = ORARst!File_Name
      
'      Call Append_update_CROSS(tempFTPPath$, tempLocalPath$, tempFTPFile$)
      
      Call get_FTP_File(tempFTPPath, tempLocalPath, tempFTPFile)
      
      FileNum = FreeFile
      Open App.Path & "\update_CROSS.bat" For Append As #FileNum
        Print #FileNum, "ftp -s:" & App.Path & "\Temp\getFTP_" & tempFTPFile & ".ftp"
      Close #FileNum
  
      ORARst.MoveNext
    Loop
  End If
    
  ORARst.Close
  ORAConn4.Close
    
  ' ********* 3）打开 CROSS *******
  FileNum = FreeFile
  Open App.Path & "\update_CROSS.bat" For Append As #FileNum
    Print #FileNum, "start "; "" & App.Path & "\CROSS.exe"
  Close #FileNum

  ' ******* 执行批处理程序，依次关闭CROSS、通过ftp下载相应升级文件、打开CROSS
  Shell "cmd.exe /c " & App.Path + "\update_CROSS.bat", vbNormalFocus
    
End Sub


Private Sub CmdLater_Click()
  Unload Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
  FrmMain.StatusBar1.Panels(1).Text = ""
  Unload Me
End Sub


