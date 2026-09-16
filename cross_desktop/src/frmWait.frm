VERSION 5.00
Begin VB.Form frmWait 
   BackColor       =   &H80000005&
   BorderStyle     =   3  'Fixed Dialog
   ClientHeight    =   405
   ClientLeft      =   255
   ClientTop       =   1410
   ClientWidth     =   2460
   ClipControls    =   0   'False
   ControlBox      =   0   'False
   Icon            =   "frmWait.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   405
   ScaleWidth      =   2460
   ShowInTaskbar   =   0   'False
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H80000005&
      Caption         =   "Querying, please wait"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Option Explicit

Private Sub Form_KeyPress(KeyAscii As Integer)
    Unload Me
End Sub

Private Sub Form_Load()
   '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = (Screen.Height - Me.Height) / 2.2
End Sub
