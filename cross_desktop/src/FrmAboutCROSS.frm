VERSION 5.00
Begin VB.Form FrmAboutCROSS 
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   5745
   ClientLeft      =   2295
   ClientTop       =   1500
   ClientWidth     =   13770
   ClipControls    =   0   'False
   Icon            =   "FrmAboutCROSS.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   12576.9
   ScaleMode       =   0  'User
   ScaleWidth      =   12930.74
   ShowInTaskbar   =   0   'False
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      ForeColor       =   &H80000008&
      Height          =   5535
      Left            =   120
      TabIndex        =   0
      Top             =   0
      Width           =   13455
      Begin VB.CommandButton cmdOK 
         Cancel          =   -1  'True
         Caption         =   "OK"
         Default         =   -1  'True
         Height          =   345
         Left            =   11280
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   5040
         Width           =   1140
      End
      Begin VB.PictureBox picIcon 
         AutoSize        =   -1  'True
         ClipControls    =   0   'False
         Height          =   780
         Left            =   240
         Picture         =   "FrmAboutCROSS.frx":3482
         ScaleHeight     =   505.68
         ScaleMode       =   0  'User
         ScaleWidth      =   505.68
         TabIndex        =   1
         Top             =   300
         Width           =   780
      End
      Begin VB.Label Label5 
         Caption         =   "Service URL: http://10.148.15.110/cross"
         Height          =   180
         Left            =   1920
         TabIndex        =   16
         Top             =   4680
         Width           =   6255
      End
      Begin VB.Label Label3 
         Caption         =   $"FrmAboutCROSS.frx":6904
         Height          =   180
         Left            =   1920
         TabIndex        =   15
         Top             =   4440
         Width           =   9255
      End
      Begin VB.Label Label9 
         Caption         =   "derived statistics such as real-time climate state, long-term average state, and historical extreme state."
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   14
         Top             =   2400
         Width           =   12765
      End
      Begin VB.Label Label8 
         Caption         =   "spatial scales such as station / township / county / city / province / custom region, "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   13
         Top             =   1920
         Width           =   12525
      End
      Begin VB.Label Label7 
         Caption         =   "statistics such as average / cumulative / max / min / extreme max / extreme min / conditional day count / conditional total, "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   12
         Top             =   2160
         Width           =   12525
      End
      Begin VB.Label Label6 
         Caption         =   "time scales such as daily / dekad / monthly / seasonal / annual / custom period, "
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   11
         Top             =   1680
         Width           =   12645
      End
      Begin VB.Label Label1 
         Caption         =   $"FrmAboutCROSS.frx":699D
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   240
         TabIndex        =   10
         Top             =   1320
         Width           =   12645
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   0
         X1              =   100
         X2              =   13200
         Y1              =   4920
         Y2              =   4920
      End
      Begin VB.Label LblTitle 
         Height          =   195
         Left            =   1200
         TabIndex        =   9
         Top             =   360
         Width           =   5055
      End
      Begin VB.Label lbl_date 
         Height          =   195
         Left            =   1200
         TabIndex        =   8
         Top             =   840
         Width           =   3135
      End
      Begin VB.Label lblVersion 
         Caption         =   "Copyright: Guangdong Climate Center, Wei Liu"
         Height          =   195
         Left            =   8760
         TabIndex        =   7
         Top             =   840
         Width           =   4005
      End
      Begin VB.Label lblDescription 
         Caption         =   "Climatic Re-statistical-data Operation and Service System, CROSS"
         ForeColor       =   &H00000000&
         Height          =   195
         Left            =   1200
         TabIndex        =   6
         Top             =   600
         Width           =   7005
      End
      Begin VB.Line Line1 
         BorderColor     =   &H00808080&
         BorderStyle     =   6  'Inside Solid
         Index           =   1
         X1              =   100
         X2              =   13200
         Y1              =   4095
         Y2              =   4095
      End
      Begin VB.Label lblDisclaimer 
         Caption         =   $"FrmAboutCROSS.frx":6A34
         ForeColor       =   &H00000000&
         Height          =   465
         Left            =   1920
         TabIndex        =   5
         Top             =   5040
         Width           =   8895
      End
      Begin VB.Label Label2 
         Caption         =   $"FrmAboutCROSS.frx":6AFB
         ForeColor       =   &H00000000&
         Height          =   1275
         Left            =   240
         TabIndex        =   4
         Top             =   2760
         Width           =   12645
      End
      Begin VB.Label Label4 
         Caption         =   "Tel: 020-39456377             QQ Group: 667885872             Personal QQ: 6822762"
         Height          =   180
         Left            =   1920
         TabIndex        =   3
         Top             =   4200
         Width           =   9375
      End
   End
End
Attribute VB_Name = "FrmAboutCROSS"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CmdOK_Click()
  Unload Me
End Sub

Private Sub Form_Load()
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  
  Me.Left = FrmMain.Left + FrmMain.Width - Me.Width - 500
  Me.Top = FrmMain.Top + FrmMain.Height - Me.Height - 600
  
  LblTitle.Caption = Sys_Caption
  lbl_date.Caption = "CROSS " & Vers_self & "  " & Vers_date
End Sub
