VERSION 5.00
Begin VB.Form FrmUserManage 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Management"
   ClientHeight    =   3075
   ClientLeft      =   9930
   ClientTop       =   5775
   ClientWidth     =   9885
   Icon            =   "FrmUserManage.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3246.895
   ScaleMode       =   0  'User
   ScaleWidth      =   10011.62
   Begin VB.Frame Frame5 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User Info"
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   9615
      Begin VB.TextBox TextIP 
         Height          =   1335
         Left            =   960
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   1320
         Width           =   8535
      End
      Begin VB.CommandButton cmdOK 
         Caption         =   "Submit"
         Height          =   345
         Left            =   8400
         Style           =   1  'Graphical
         TabIndex        =   15
         Top             =   840
         Width           =   1020
      End
      Begin VB.TextBox TextPassword 
         Height          =   270
         Left            =   3360
         TabIndex        =   2
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TextDepartment 
         Height          =   270
         Left            =   8160
         TabIndex        =   4
         Top             =   360
         Width           =   1335
      End
      Begin VB.TextBox TextName 
         Height          =   270
         Left            =   5760
         TabIndex        =   3
         Top             =   360
         Width           =   1335
      End
      Begin VB.ComboBox ComboUserZone 
         Height          =   300
         Left            =   960
         TabIndex        =   6
         Text            =   "广东"
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox ComboDataZone 
         Height          =   300
         Left            =   3360
         TabIndex        =   7
         Text            =   "广东"
         Top             =   840
         Width           =   1335
      End
      Begin VB.ComboBox ComboLevel 
         Height          =   300
         Left            =   5760
         TabIndex        =   5
         Text            =   "Province-level"
         Top             =   840
         Width           =   1335
      End
      Begin VB.TextBox TextUser 
         Height          =   270
         Left            =   960
         TabIndex        =   1
         Top             =   360
         Width           =   1335
      End
      Begin VB.Label Label8 
         BackColor       =   &H80000005&
         Caption         =   "Data Saving IP"
         Height          =   615
         Left            =   240
         TabIndex        =   17
         Top             =   1365
         Width           =   615
      End
      Begin VB.Label Label7 
         BackColor       =   &H80000005&
         Caption         =   "Work Unit"
         Height          =   255
         Left            =   7320
         TabIndex        =   14
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Username"
         Height          =   255
         Left            =   4920
         TabIndex        =   13
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label6 
         BackColor       =   &H80000005&
         Caption         =   "Data Range"
         Height          =   375
         Left            =   2880
         TabIndex        =   12
         Top             =   765
         Width           =   495
      End
      Begin VB.Label Label5 
         BackColor       =   &H80000005&
         Caption         =   "Admin Region"
         Height          =   375
         Left            =   240
         TabIndex        =   11
         Top             =   840
         Width           =   735
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Level"
         Height          =   255
         Left            =   5160
         TabIndex        =   10
         Top             =   885
         Width           =   615
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Caption         =   "Password"
         Height          =   255
         Left            =   2520
         TabIndex        =   9
         Top             =   405
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Account"
         Height          =   255
         Left            =   240
         TabIndex        =   8
         Top             =   405
         Width           =   735
      End
   End
End
Attribute VB_Name = "FrmUserManage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit


Private Sub CmdOK_Click()
  'On Error GoTo ErrMsg
  
  Dim Cancel%
  
  If TextUser.Text = "" Then MsgBox "Account cannot be empty", vbInformation, "Notice": Exit Sub

  If userID <> "liuw" Then
    If TextUser.Text <> userID Then MsgBox "Cannot modify another account's information", vbInformation, "Notice": Exit Sub
  End If

  ORAConn4.Open
  ORAConn4.CursorLocation = 3

  strSQL = "select * from T_OTHE_CROSS_USERS_PROPERTIES where userid='" & TextUser.Text & "'"
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockOptimistic
  
  With ORARst
  If .EOF Then
    Cancel = MsgBox("Create the new user?", vbYesNo + vbQuestion, "Notice")
    If Cancel = vbNo Then
      .Close: ORAConn4.Close
      Exit Sub
    ElseIf Cancel = vbYes Then
      If TextPassword.Text = "" Then MsgBox "Password cannot be empty", vbInformation, "Notice": Exit Sub
      If TextName.Text = "" Then MsgBox "Username cannot be empty", vbInformation, "Notice": Exit Sub
      .AddNew
    End If

  Else
    Cancel = MsgBox("Submit the user information?", vbYesNo + vbQuestion, "Notice")
    If Cancel = vbNo Then
      .Close: ORAConn4.Close
      Exit Sub
    ElseIf Cancel = vbYes Then
    
    End If
  End If
    
  !userID = TextUser.Text
  
  If (TextPassword.Text <> "" And TextPassword.Text <> "********") Then
    !Password = to_MD5(TextPassword.Text, 32)
  End If
  
  !FullName = TextName.Text
  !Department = TextDepartment.Text
  If ComboLevel.Text = "Province-level" Then !UserLevel = 1
  If ComboLevel.Text = "City-level" Then !UserLevel = 2
  If ComboLevel.Text = "County-level" Then !UserLevel = 3
  !UserZone = ComboUserZone.Text '用户所在广东省或某市、县
  !DataZone = ComboDataZone.Text '用户所在广东省或某市
  !IP_SaveData = TextIP.Text
  
  
'  !IP_Whitelist = TextIPwhite.Text
  

  .Update
  End With
  ORARst.Close
  ORAConn4.Close
  
  FrmMain.StatusBar1.Panels(1).Text = "User information added/updated successfully"
  
  
  FrmUserInfo.Show
  Call FrmUserInfo.ShowUserInfo
  FrmUserInfo.SetFocus
  
  Unload Me

  Exit Sub
ErrMsg:
  MsgBox Err.Description, vbCritical, "Error Message"
  
End Sub

Private Sub Form_Load()
   '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  Dim i%

  
  Me.Left = FrmMain.Left + 500
  Me.Top = FrmMain.Top + FrmMain.Height - Me.Height - 600
  
  
  
  ComboLevel.AddItem "Province-level": ComboLevel.AddItem "City-level": ComboLevel.AddItem "County-level"

  If userID <> "liuw" Then
    TextUser.Enabled = False
    TextPassword.Enabled = False
    TextName.Enabled = False
    TextDepartment.Enabled = False
    ComboLevel.Enabled = False
    ComboUserZone.Enabled = False
    ComboDataZone.Enabled = False
    TextIP.Enabled = False
  End If


End Sub

