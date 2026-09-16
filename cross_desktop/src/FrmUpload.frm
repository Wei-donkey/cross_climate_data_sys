VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmUpload 
   Caption         =   "Upgrade Management"
   ClientHeight    =   12285
   ClientLeft      =   60
   ClientTop       =   405
   ClientWidth     =   9255
   Icon            =   "FrmUpload.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   12285
   ScaleWidth      =   9255
   Begin VB.Frame FrmDataMap 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Files to be updated in this upgrade"
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   120
      TabIndex        =   17
      Top             =   120
      Width           =   4500
      Begin VB.CommandButton CmdUpload 
         Caption         =   "Upload"
         Height          =   345
         Left            =   2160
         Style           =   1  'Graphical
         TabIndex        =   20
         Top             =   2160
         Width           =   2100
      End
      Begin VB.CommandButton CmdDelFile 
         Caption         =   "Delete"
         Height          =   345
         Left            =   1200
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2160
         Width           =   900
      End
      Begin VB.CommandButton CmdAddFile 
         Caption         =   "Add"
         Height          =   345
         Left            =   240
         Style           =   1  'Graphical
         TabIndex        =   18
         Top             =   2160
         Width           =   900
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid1 
         Height          =   1695
         Left            =   240
         TabIndex        =   21
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   2990
         _Version        =   393216
         BackColorBkg    =   16777215
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin MSComDlg.CommonDialog CommonDialogOpen 
         Left            =   3960
         Top             =   120
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   3960
         TabIndex        =   22
         Top             =   9240
         Width           =   735
      End
   End
   Begin VB.Frame Frame3 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "All updated versions"
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   2655
      Left            =   4640
      TabIndex        =   14
      Top             =   120
      Width           =   4500
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid2 
         Height          =   2055
         Left            =   240
         TabIndex        =   15
         Top             =   360
         Width           =   4095
         _ExtentX        =   7223
         _ExtentY        =   3625
         _Version        =   393216
         BackColorBkg    =   16777215
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label3 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   3960
         TabIndex        =   16
         Top             =   9240
         Width           =   735
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "All feature entries across all version updates"
      ClipControls    =   0   'False
      ForeColor       =   &H80000008&
      Height          =   6975
      Left            =   120
      TabIndex        =   5
      Top             =   5160
      Width           =   9015
      Begin VB.TextBox Text4 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   8
         Top             =   5880
         Width           =   8655
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   345
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   5400
         Width           =   900
      End
      Begin VB.CommandButton CmdRead 
         Caption         =   "View"
         Height          =   345
         Left            =   6840
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   5400
         Width           =   900
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid3 
         Height          =   4935
         Left            =   240
         TabIndex        =   9
         Top             =   360
         Width           =   8655
         _ExtentX        =   15266
         _ExtentY        =   8705
         _Version        =   393216
         BackColorBkg    =   16777215
         AllowUserResizing=   1
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "宋体"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         _NumberOfBands  =   1
         _Band(0).Cols   =   2
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Update entry details: "
         Height          =   255
         Left            =   240
         TabIndex        =   11
         Top             =   5640
         Width           =   3975
      End
      Begin VB.Label LabelLength2 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   3960
         TabIndex        =   10
         Top             =   9240
         Width           =   735
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "features in this version update"
      ForeColor       =   &H80000008&
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   2880
      Width           =   9015
      Begin VB.TextBox TxtVersion 
         Height          =   270
         Left            =   7680
         TabIndex        =   13
         Top             =   360
         Width           =   1095
      End
      Begin VB.TextBox TxtContent 
         Appearance      =   0  'Flat
         Height          =   855
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   2
         Top             =   720
         Width           =   8655
      End
      Begin VB.CommandButton cmdAddInfo 
         Caption         =   "Add"
         Height          =   345
         Left            =   7800
         Style           =   1  'Graphical
         TabIndex        =   1
         Top             =   1680
         Width           =   900
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Version: "
         Height          =   255
         Left            =   6840
         TabIndex        =   12
         Top             =   360
         Width           =   735
      End
      Begin VB.Label LabelHint1 
         BackColor       =   &H80000005&
         Caption         =   "Up to 250 characters allowed. Characters entered: "
         Height          =   255
         Left            =   240
         TabIndex        =   4
         Top             =   1800
         Width           =   7455
      End
      Begin VB.Label LabelLength 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   3960
         TabIndex        =   3
         Top             =   1680
         Width           =   735
      End
   End
   Begin VB.Menu mnuGridPopup 
      Caption         =   "GridPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuGridCopy 
         Caption         =   "Copy Selection"
      End
   End
End
Attribute VB_Name = "FrmUpload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private mLastRightClickGrid As Object '最近一次被右键点击的表格，供mnuGridCopy_Click使用


Private Sub Form_Load()
  Dim i%
  
  '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
    
  Me.Left = FrmMain.Left + 500
  Me.Top = FrmMain.Top + FrmMain.Height - Me.Height - 600
  
  With HFGrid1
    .Cols = 4
    .ColWidth(0) = 350
    .ColWidth(1) = 800
    .ColWidth(2) = 1800
    .ColWidth(3) = 900
    
    .TextMatrix(0, 1) = "Version"
    .TextMatrix(0, 2) = "File"
    .TextMatrix(0, 3) = "Relative Path"
  End With
    
  Call showVersionInfo
  Call showUpdateFile
  
End Sub


Private Sub showUpdateFile()
  HFGrid2.Clear
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  
  With HFGrid2
  strSQL = "select * from T_OTHE_CROSS_UPDATE_FILE order by version desc"

  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
  
    Set .DataSource = ORARst
    .Cols = 4
      
    .ColWidth(0) = 350
    .ColWidth(1) = 800
    .ColWidth(2) = 1800
    .ColWidth(3) = 900
      
    .TextMatrix(0, 1) = "Version"
    .TextMatrix(0, 2) = "File"
    .TextMatrix(0, 3) = "Relative Path"
    
    
  End If
  
  ORARst.Close
  End With
  
  ORAConn4.Close
End Sub


Private Sub CmdAddFile_Click()

  Dim i%, j%
  Dim tempStr$
  Dim Heads$() '定义表头
  
'  CommonDialogOpen.Filter = "csvFile(*.csv)|*.csv"
  CommonDialogOpen.InitDir = App.Path
  CommonDialogOpen.ShowOpen
  If CommonDialogOpen.Flags = 0 Then FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub

  With HFGrid1
    If .TextMatrix(.Rows - 1, 1) <> "" Then .Rows = .Rows + 1
    .Row = .Rows - 1
    .TextMatrix(.Row, 1) = Vers_self
    .TextMatrix(.Row, 2) = CommonDialogOpen.FileTitle
    tempStr = CommonDialogOpen.FileName
    tempStr = Replace(tempStr, App.Path, "")
    tempStr = Replace(tempStr, CommonDialogOpen.FileTitle, "")
    
    '把最右边的“\”删掉
    Dim tempLen%
    tempLen = Len(tempStr)
    tempStr = Left(tempStr, tempLen - 1)
    
    .TextMatrix(.Row, 3) = tempStr
    
  End With

End Sub

Private Sub CmdDelFile_Click()
  With HFGrid1

    .TextMatrix(.Row, 1) = ""
    .TextMatrix(.Row, 2) = ""
    .TextMatrix(.Row, 3) = ""
  End With
End Sub


Private Sub CmdUpload_Click()
  Dim i%

  ' ******* 写 sendFTP_update.bat 批处理程序，通过ftp上传相应升级文件，并记录进T_OTHE_CROSS_UPDATE_FILE数据表
  Dim tempFilePath$, tempVersion$, tempFileName$
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  
  FileNum = FreeFile
  Open App.Path & "\upload_CROSS.bat" For Output As #FileNum
  Close #FileNum
  
  With HFGrid1
  
  For i = 1 To .Rows - 1
    If .TextMatrix(i, 1) <> "" Then
      tempVersion = .TextMatrix(i, 1)
      tempFileName = .TextMatrix(i, 2)
      tempFilePath = .TextMatrix(i, 3)
        
      Call send_FTP_File(tempFilePath, App.Path & tempFilePath$, tempFileName$)
      Open App.Path & "\upload_CROSS.bat" For Append As #FileNum
        Print #FileNum, "ftp -s:" & App.Path & "\Temp\sendFTP_" & tempFileName & ".ftp"
      Close #FileNum
      
      strSQL = "select * from T_OTHE_CROSS_UPDATE_FILE"
      strSQL = strSQL + " where VERSION = '" & tempVersion & "' and FILE_NAME = '" & tempFileName & "'"
      If tempFilePath <> "" Then
        strSQL = strSQL + " and FILE_PATH= '" & tempFilePath & "'"
      Else
        strSQL = strSQL + " and FILE_PATH is null"
      End If
      ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
      If ORARst.EOF = True Then
        strSQL = "insert into T_OTHE_CROSS_UPDATE_FILE(VERSION,FILE_NAME,FILE_PATH)"
        strSQL = strSQL + " Values('" & tempVersion & "','" & tempFileName & "','" & tempFilePath & "')"
        ORAConn4.Execute strSQL
      End If
      ORARst.Close
    
    End If
  Next i
  
  ORAConn4.Close

  End With

  ' ******* 执行批处理程序，依次关闭CROSS、通过ftp下载相应升级文件、打开CROSS
  Shell "cmd.exe /c " & App.Path + "\upload_CROSS.bat", vbNormalFocus



End Sub



Private Sub showVersionInfo()
  HFGrid3.Clear
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  
  With HFGrid3
  
  strSQL = "select * from T_OTHE_CROSS_VERSION_INFO order by version desc, ddatetime desc"

  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
  
    Set .DataSource = ORARst
    .ColWidth(0) = 350
    .ColWidth(1) = 1550
    .ColWidth(2) = 700
    .ColWidth(3) = 5700
    
    .TextMatrix(0, 1) = "Time"
    .TextMatrix(0, 2) = "Version"
    .TextMatrix(0, 3) = "Content"
    
  End If
  
  ORARst.Close
  End With
  
  ORAConn4.Close
End Sub


Private Sub cmdAddInfo_Click()
  Dim Addtime As Date
  Dim strTemp$(), i%
  Dim strContent$
  Dim version$ 'Version
  
  version = TxtVersion.Text
  strContent$ = TxtContent.Text
  Addtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  
  strSQL = "insert into T_OTHE_CROSS_VERSION_INFO(DDATETIME,VERSION,CONTENT)"
  strSQL = strSQL + " Values(to_date('" & Addtime & "', 'yyyy-mm-dd hh24:mi:ss'),'" & version & "','" & strContent & "')"
  
  ORAConn4.Execute strSQL
  ORAConn4.Close
  TxtContent.Text = ""
  Call showVersionInfo
  
  FrmMain.StatusBar1.Panels(1).Text = "Update entry added successfully"

End Sub



Private Sub CmdRead_Click()
  
  With HFGrid3
  If .TextMatrix(.Row, 1) = "" Then Exit Sub
    
  Text4.Text = .TextMatrix(.Row, 3)
  End With
  
End Sub


Private Sub CmdDelete_Click()
  Dim i%, Response%
  
  With HFGrid3
  If .TextMatrix(.Row, 1) = "" Then Exit Sub
  
    
  Response = MsgBox("Delete this update entry?", vbOKCancel)
  If Response = 2 Then Exit Sub
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  ORAConn4.Execute "delete from T_OTHE_CROSS_VERSION_INFO where DDATETIME=to_date('" & .TextMatrix(.Row, 1) & "', 'yyyy-mm-dd hh24:mi:ss')  and version='" & .TextMatrix(.Row, 2) & "'"
  ORAConn4.Close
    
  Call showVersionInfo
  FrmMain.StatusBar1.Panels(1).Text = "Deleted successfully"
    
  End With
  
  Exit Sub
  
End Sub
Private Sub mnuGridCopy_Click()
  Call CopyGridSelectionToClipboard(mLastRightClickGrid)
End Sub

Private Sub HFGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Set mLastRightClickGrid = HFGrid1
    PopupMenu mnuGridPopup
  End If
End Sub

Private Sub HFGrid2_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Set mLastRightClickGrid = HFGrid2
    PopupMenu mnuGridPopup
  End If
End Sub

Private Sub HFGrid3_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  If Button = 2 Then
    Set mLastRightClickGrid = HFGrid3
    PopupMenu mnuGridPopup
  End If
End Sub
