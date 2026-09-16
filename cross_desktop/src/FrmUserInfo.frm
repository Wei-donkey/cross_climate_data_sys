VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmUserInfo 
   Caption         =   "User List"
   ClientHeight    =   12285
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   22470
   Icon            =   "FrmUserInfo.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   123.158
   ScaleMode       =   0  'User
   ScaleWidth      =   133.869
   Begin VB.CommandButton CmdLogOutAll 
      Caption         =   "Log Off All Sessions"
      Height          =   375
      Left            =   15360
      TabIndex        =   5
      Top             =   120
      Width           =   2175
   End
   Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid1 
      Height          =   11535
      Left            =   120
      TabIndex        =   4
      Top             =   600
      Width           =   22215
      _ExtentX        =   39185
      _ExtentY        =   20346
      _Version        =   393216
      BackColorFixed  =   16777215
      BackColorBkg    =   16777215
      Appearance      =   0
      _NumberOfBands  =   1
      _Band(0).Cols   =   2
   End
   Begin VB.CommandButton CmdSave 
      Caption         =   "Save"
      Height          =   375
      Left            =   21360
      TabIndex        =   3
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton CmdAdd 
      Caption         =   "Add New"
      Height          =   375
      Left            =   18240
      TabIndex        =   2
      Top             =   120
      Width           =   975
   End
   Begin VB.CommandButton CmdDel 
      Caption         =   "Delete"
      Height          =   375
      Left            =   19320
      TabIndex        =   1
      Top             =   120
      Width           =   855
   End
   Begin VB.CommandButton CmdModify 
      Caption         =   "View"
      Height          =   375
      Left            =   20280
      TabIndex        =   0
      Top             =   120
      Width           =   975
   End
   Begin MSComDlg.CommonDialog CommonDialogSave 
      Left            =   0
      Top             =   0
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu mnuGridPopup 
      Caption         =   "GridPopup"
      Visible         =   0   'False
      Begin VB.Menu mnuGridCopy 
         Caption         =   "Copy Selection"
      End
   End
End
Attribute VB_Name = "FrmUserInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public mLastRightClickGrid As Object '最近一次被点击（任意按钮）的表格；供mnuGridCopy_Click及ChartData查找图表数据源时使用

  
Private Sub CmdAdd_Click()
  FrmUserManage.Show
  FrmUserManage.SetFocus
End Sub

Private Sub CmdDel_Click()
  Dim tempID$, Response
  
  tempID = HFGrid1.TextMatrix(HFGrid1.Row, 1)
  
  
  Response = MsgBox("Delete user " & tempID & " ？", vbOKCancel, "Confirm")
  If Response = 2 Then Exit Sub
  
  If tempID = "" Then Exit Sub
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3

  ORAConn4.Execute "delete from T_OTHE_CROSS_USERS_PROPERTIES where userid='" & tempID & "'"
  
  ORAConn4.Close
  
  Call ShowUserInfo
  

  
End Sub


'2024-12-6Add：可以定期于月初周末注销一 times
Private Sub CmdLogOutAll_Click()
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    
    strSQL = "update T_OTHE_CROSS_USERS_LOG"
    strSQL = strSQL + " set status=0"
    ORAConn4.Execute strSQL
    ORAConn4.Close
    
    MsgBox "All users' login status has been reset", vbInformation, "Notice": Screen.MousePointer = 1: Exit Sub
    
End Sub

Private Sub CmdModify_Click()
  With FrmUserManage
  
  .TextUser.Text = HFGrid1.TextMatrix(HFGrid1.Row, 1)
'  .TextPassword.Text = HFGrid1.TextMatrix(.Row, 2)
  .TextName.Text = HFGrid1.TextMatrix(HFGrid1.Row, 2)
  .TextDepartment.Text = HFGrid1.TextMatrix(HFGrid1.Row, 3)
  If HFGrid1.TextMatrix(HFGrid1.Row, 4) = 1 Then .ComboLevel.Text = "Province-level"
  If HFGrid1.TextMatrix(HFGrid1.Row, 4) = 2 Then .ComboLevel.Text = "City-level"
  If HFGrid1.TextMatrix(HFGrid1.Row, 4) = 3 Then .ComboLevel.Text = "County-level"
  .ComboUserZone.Text = HFGrid1.TextMatrix(HFGrid1.Row, 5)  '用户所在广东省或某市、县
  .ComboDataZone.Text = HFGrid1.TextMatrix(HFGrid1.Row, 6) '用户所在广东省或某市（地区）
  
  .TextIP.Text = HFGrid1.TextMatrix(HFGrid1.Row, 7)
'  .TextIPwhite.Text = HFGrid1.TextMatrix(HFGrid1.Row, 8)
  
  .TextPassword = "********"
  
  End With
  
  FrmUserManage.Show
  FrmUserManage.SetFocus
End Sub

Private Sub CmdSave_Click()
  
  Dim i%, j%
  Dim strFileName$, strFileType$, strTemp$, strStacode$

  CommonDialogSave.Filter = "Text Files (*.csv)|*.csv": strFileType = "csv"
  CommonDialogSave.FilterIndex = 1
  CommonDialogSave.FileName = "User Profile"
  CommonDialogSave.InitDir = App.Path & "\Output"
  CommonDialogSave.Flags = &H2
  CommonDialogSave.ShowSave
  If CommonDialogSave.Flags = 2 Then Screen.MousePointer = 1: FrmMain.StatusBar1.Panels(1).Text = "": Exit Sub
  
  frmWait.Label1.Caption = "Saving, please wait"
  frmWait.Show
  Screen.MousePointer = 11
  FrmMain.StatusBar1.Panels(1).Text = "Saving"
  
  strFileName = CommonDialogSave.FileName
  If Len(Dir(strFileName)) <> 0 Then Kill strFileName '如果存在同名文件则删除同名文件
    
    Open strFileName For Output As #1
    With HFGrid1
    For i = 0 To .Rows - 1
      strTemp = SanitizeForCSV(.TextMatrix(i, 0)) ' 部分单元格（如保存的多个IP)含换行符，直接写出会把一行拆成多行，先去除换行符再写
      strStacode = .TextMatrix(i, 1)
      For j = 1 To .Cols - 1
        strTemp = strTemp & "," & SanitizeForCSV(.TextMatrix(i, j))
      Next j
      Print #1, strTemp
    Next i
    End With
    Close #1
  
  frmWait.Hide: DoEvents
  Screen.MousePointer = 1
  FrmMain.StatusBar1.Panels(1).Text = "Save result: Done"

End Sub

Private Sub Form_Load()
  Dim i%
  
   '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1

  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = 11 * (Screen.Height - Me.Height) / 20
  
  '2026-07-22 访客模式下不查询真实用户信息
  If Not GuestMode Then
    Call ShowUserInfo
  End If
    
  CmdModify.Visible = True
  If userID <> "liuw" Then
    CmdAdd.Visible = False: CmdDel.Visible = False
    CmdLogOutAll.Visible = False
  End If

  '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
  If GuestMode Then
    Call ImportData.LoadGuestCSV(FrmUserInfo.HFGrid1, Nothing, "UserInfo.csv", 0)
    
    With HFGrid1
      .ColWidth(0) = 600
      .ColWidth(1) = 1100
      .ColWidth(2) = 1100
      .ColWidth(3) = 2000
      .ColWidth(4) = 600
      .ColWidth(5) = 1000
      .ColWidth(6) = 1000
      .ColWidth(7) = 7500
      
      .ColWidth(8) = 1000
      .ColWidth(9) = 2000
      .ColWidth(10) = 1500
      .ColWidth(11) = 1000
      .ColWidth(12) = 1000
    
      For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignLeftCenter
        .FixedAlignment(i) = flexAlignLeftCenter
      Next i
    
    End With
      
    Exit Sub
  End If

End Sub

Public Sub ShowUserInfo()
  Dim i%
  ORAConn4.Open
  ORAConn4.CursorLocation = 3

  
  strSQL = "SELECT a.userid,a.fullname,a.department,a.userlevel,a.userzone,a.DataZone,a.IP_savedata,b.LOGCOUNTS,b.LOGTIME,d.userip ,c.PCCOUNTS,a.CROSS_ver"
  strSQL = strSQL & " from T_OTHE_CROSS_USERS_PROPERTIES a"
  strSQL = strSQL & " left join (SELECT userid,max(logtime) LOGTIME,count(*) LOGCOUNTS from T_OTHE_CROSS_USERS_LOG group by userid) b"
  strSQL = strSQL & " on a.userid=b.userid"
  strSQL = strSQL & " left join (SELECT userid,count(*) PCCOUNTS from T_OTHE_CROSS_USERS_LOG where status=1 group by userid) c"
  strSQL = strSQL & " on a.userid=c.userid"
  strSQL = strSQL & " left join (SELECT logtime, userip FROM T_OTHE_CROSS_USERS_LOG) d"
  strSQL = strSQL & " on b.logtime = d.logtime"
  strSQL = strSQL & " order by logtime desc nulls last"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  
  With HFGrid1
    Set .DataSource = ORARst
    .TextMatrix(0, 0) = "No.": .ColWidth(0) = 600
    .TextMatrix(0, 1) = "Account": .ColWidth(1) = 1100
    .TextMatrix(0, 2) = "Username": .ColWidth(2) = 1100
    .TextMatrix(0, 3) = "User's Work Unit": .ColWidth(3) = 2000
    .TextMatrix(0, 4) = "Level": .ColWidth(4) = 600
    .TextMatrix(0, 5) = "Administrative Region": .ColWidth(5) = 1000
    .TextMatrix(0, 6) = "Data Range": .ColWidth(6) = 1000
    .TextMatrix(0, 7) = "Data-Download IP": .ColWidth(7) = 7500
    .TextMatrix(0, 8) = "Login Count": .ColWidth(8) = 1000
    .TextMatrix(0, 9) = "Last Login Time": .ColWidth(9) = 2000
    .TextMatrix(0, 10) = "Last Login IP": .ColWidth(10) = 1500
    .TextMatrix(0, 11) = "Online Sessions": .ColWidth(11) = 1000
    .TextMatrix(0, 12) = "Version": .ColWidth(12) = 1000
    
    For i = 1 To .Rows - 1
      .TextMatrix(i, 0) = i
    Next i
    
    For i = 0 To .Cols - 1
      .ColAlignment(i) = flexAlignLeftCenter
      .FixedAlignment(i) = flexAlignLeftCenter
    Next i
    
    
  End With
  
  ORARst.Close
  ORAConn4.Close

End Sub


Private Sub HFGrid1_MouseUp(Button As Integer, shift As Integer, X As Single, Y As Single)
  Set mLastRightClickGrid = HFGrid1
  If Button = 2 Then PopupMenu mnuGridPopup
  Dim i%
  With HFGrid1
  If shift = 2 Then
    .Sort = 7
    If .Col = .Cols - 1 Then .Sort = 3
  ElseIf shift = 4 Then
    .Sort = 8
    If .Col = .Cols - 1 Then .Sort = 4
  End If
  For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
  Next i
  .Refresh
  End With
End Sub


'鼠标滚轮驱动表格上下滚动
Private Sub HFGrid1_GotFocus()
  On Error Resume Next
  Oldwinproc = GetWindowLong(Me.hWnd, GWL_WNDPROC)
  SetWindowLong Me.hWnd, GWL_WNDPROC, AddressOf FlexScroll
End Sub
    
'鼠标滚轮驱动表格上下滚动
Private Sub HFGrid1_LostFocus()
  On Error Resume Next
  SetWindowLong Me.hWnd, GWL_WNDPROC, Oldwinproc
End Sub


Private Sub mnuGridCopy_Click()
  Call CopyGridSelectionToClipboard(mLastRightClickGrid)
End Sub
