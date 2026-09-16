VERSION 5.00
Begin VB.Form FrmPassword 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Password Settings"
   ClientHeight    =   3030
   ClientLeft      =   2760
   ClientTop       =   3750
   ClientWidth     =   3600
   Icon            =   "FrmPassword.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12271.5
   ScaleMode       =   0  'User
   ScaleWidth      =   3600
   ShowInTaskbar   =   0   'False
   Begin VB.CheckBox Check2 
      Caption         =   "Change Pwd"
      Height          =   255
      Left            =   2160
      TabIndex        =   9
      ToolTipText     =   "No need to re-enter manually at login"
      Top             =   120
      Width           =   1335
   End
   Begin VB.CommandButton CmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   2640
      Style           =   1  'Graphical
      TabIndex        =   8
      Top             =   2520
      Width           =   855
   End
   Begin VB.CheckBox Check1 
      Caption         =   "Remember Pwd"
      Height          =   255
      Left            =   240
      TabIndex        =   7
      ToolTipText     =   "No need to re-enter manually at login"
      Top             =   120
      Width           =   1455
   End
   Begin VB.Frame Frame1 
      Caption         =   "Change Password"
      Height          =   1815
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   3375
      Begin VB.TextBox txtNewPassword 
         Enabled         =   0   'False
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   3
         Top             =   840
         Width           =   1845
      End
      Begin VB.TextBox txtOldPassWord 
         Enabled         =   0   'False
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   2
         Top             =   360
         Width           =   1845
      End
      Begin VB.TextBox TxtConfirm 
         Enabled         =   0   'False
         Height          =   345
         IMEMode         =   3  'DISABLE
         Left            =   1320
         PasswordChar    =   "*"
         TabIndex        =   1
         Top             =   1320
         Width           =   1845
      End
      Begin VB.Label lblLabels 
         Caption         =   "New Pwd"
         Height          =   270
         Index           =   1
         Left            =   240
         TabIndex        =   6
         Top             =   960
         Width           =   960
      End
      Begin VB.Label lblLabels 
         Caption         =   "Current Pwd"
         Height          =   270
         Index           =   0
         Left            =   240
         TabIndex        =   5
         Top             =   480
         Width           =   1080
      End
      Begin VB.Label lblLabels 
         Caption         =   "Confirm Pwd"
         Height          =   270
         Index           =   2
         Left            =   240
         TabIndex        =   4
         Top             =   1440
         Width           =   1080
      End
   End
End
Attribute VB_Name = "FrmPassword"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit


Private Sub Check2_Click()
  If Check2.Value = Checked Then
    txtOldPassWord.Enabled = True
    txtNewPassword.Enabled = True
    TxtConfirm.Enabled = True
    
  ElseIf Check2.Value = Unchecked Then
    txtOldPassWord.Enabled = False
    txtNewPassword.Enabled = False
    TxtConfirm.Enabled = False
  
  End If
  
End Sub

Private Sub CmdOK_Click()

  If Check2.Value = Checked Then
    If txtOldPassWord.Text <> "" Then
    
      ORAConn4.Open
      ORAConn4.CursorLocation = 3
      
      strSQL = "select * from T_OTHE_CROSS_USERS_PROPERTIES where UserID='" & userID & "'"
      ORARst.CursorLocation = adUseClient
      ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockOptimistic
    
    
    
      If to_MD5(txtOldPassWord.Text, 32) = ORARst!Password Then
        If txtNewPassword.Text = TxtConfirm.Text Then
          ORARst!Password = to_MD5(txtNewPassword.Text, 32)
          ORARst.Update
          ORARst.Close
          ORAConn4.Close
          MsgBox "Password changed successfully", vbInformation, "Notice"
        ElseIf txtNewPassword.Text <> TxtConfirm.Text Then
          MsgBox "The two new passwords do not match -- please re-enter", vbInformation, "Notice"
          ORARst.Close
          ORAConn4.Close
          txtOldPassWord.SetFocus: txtOldPassWord.SelLength = Len(txtOldPassWord.Text): Exit Sub
        End If
      ElseIf txtOldPassWord.Text <> ORARst!Password Then
        MsgBox "Incorrect current password -- please re-enter", vbInformation, "Notice"
        ORARst.Close
        ORAConn4.Close
        txtOldPassWord.SetFocus: txtOldPassWord.SelLength = Len(txtOldPassWord.Text): Exit Sub
      End If
        
    End If
  End If
  
  If Check1.Value = Checked Then
    Call ReadWriteIni("Config.ini", "[Password]", UserPassword, "w")
  ElseIf Check1.Value = Unchecked Then
    Call ReadWriteIni("Config.ini", "[Password]", "不保存", "w")
  End If

  Unload Me
  
  Exit Sub
ErrMsg:
  MsgBox Err.Description, vbCritical, "Error Message"
  
End Sub

Private Sub Form_Load()
   '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1

  Dim tempStr$

  
  Me.Left = FrmMain.Left + 500
  Me.Top = FrmMain.Top + FrmMain.Height - Me.Height - 600
  
  
  txtOldPassWord.Text = ""
  txtNewPassword.Text = ""
  TxtConfirm.Text = ""

  Call ReadWriteIni("Config.ini", "[Password]", tempStr, "r")
  If tempStr <> "不保存" Then
    Check1.Value = Checked
  End If
  
  '2026-07-22 访客模式下禁止修改密  码
  If GuestMode Then CmdOK.Enabled = False
  
  Exit Sub
ErrMsg:
  MsgBox Err.Description, vbCritical, "Error Message"
  
End Sub
