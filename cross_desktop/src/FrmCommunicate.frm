VERSION 5.00
Object = "{0ECD9B60-23AA-11D0-B351-00A0C9055D8E}#6.0#0"; "MShflxgd.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "ComDlg32.OCX"
Begin VB.Form FrmCommunicate 
   Caption         =   "Message Center"
   ClientHeight    =   8145
   ClientLeft      =   3780
   ClientTop       =   1905
   ClientWidth     =   18135
   Icon            =   "FrmCommunicate.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8145
   ScaleWidth      =   18135
   Begin VB.Frame Frame4 
      BackColor       =   &H8000000E&
      Caption         =   "Reply Details"
      Height          =   2775
      Left            =   11880
      TabIndex        =   16
      Top             =   2280
      Width           =   6135
      Begin VB.CommandButton CmdReply 
         Caption         =   "Reply"
         Height          =   345
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   19
         Top             =   2280
         Width           =   900
      End
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   1935
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label LabelLength2 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   3840
         TabIndex        =   22
         Top             =   2400
         Width           =   855
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Up to 250 characters allowed. Characters entered: "
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   4695
      End
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H8000000E&
      Caption         =   "Content Details"
      Height          =   2055
      Left            =   11880
      TabIndex        =   14
      Top             =   120
      Width           =   6135
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   1575
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   15
         Top             =   360
         Width           =   5895
      End
   End
   Begin VB.Frame Frame2 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "User Feedback / Admin Messages"
      ForeColor       =   &H80000008&
      Height          =   7815
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   11655
      Begin VB.CommandButton CmdSaveData 
         Caption         =   "Save Table"
         Height          =   350
         Left            =   4680
         TabIndex        =   20
         Top             =   240
         Width           =   1440
      End
      Begin VB.ComboBox ComboType 
         Height          =   300
         Index           =   1
         Left            =   3360
         TabIndex        =   11
         Top             =   300
         Width           =   1215
      End
      Begin VB.ComboBox ComboType 
         Height          =   300
         Index           =   0
         Left            =   1080
         TabIndex        =   10
         Top             =   300
         Width           =   1215
      End
      Begin VB.CommandButton CmdMarkUnread 
         Caption         =   "MarkAsUnhandled"
         Height          =   345
         Left            =   8880
         Style           =   1  'Graphical
         TabIndex        =   9
         Top             =   240
         Width           =   1620
      End
      Begin VB.CommandButton CmdMarkRead 
         Caption         =   "MarkAsRead"
         Height          =   345
         Left            =   7680
         Style           =   1  'Graphical
         TabIndex        =   8
         Top             =   240
         Width           =   1140
      End
      Begin VB.CommandButton CmdRead 
         Caption         =   "View"
         Height          =   345
         Left            =   6720
         Style           =   1  'Graphical
         TabIndex        =   7
         Top             =   240
         Width           =   900
      End
      Begin VB.CommandButton CmdDelete 
         Caption         =   "Delete"
         Height          =   345
         Left            =   10560
         Style           =   1  'Graphical
         TabIndex        =   6
         Top             =   240
         Width           =   900
      End
      Begin MSHierarchicalFlexGridLib.MSHFlexGrid HFGrid1 
         Height          =   6975
         Left            =   120
         TabIndex        =   5
         ToolTipText     =   "Double-click any row to preview the feedback and its reply"
         Top             =   720
         Width           =   11415
         _ExtentX        =   20135
         _ExtentY        =   12303
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
      Begin MSComDlg.CommonDialog CommonDialogSave 
         Left            =   6240
         Top             =   240
         _ExtentX        =   847
         _ExtentY        =   847
         _Version        =   393216
      End
      Begin VB.Label Label4 
         BackColor       =   &H80000005&
         Caption         =   "Status"
         Height          =   255
         Left            =   2520
         TabIndex        =   13
         Top             =   360
         Width           =   735
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Msg Type"
         Height          =   255
         Left            =   240
         TabIndex        =   12
         Top             =   360
         Width           =   855
      End
   End
   Begin VB.Frame Frame1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "Messages between users and admins"
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   11880
      TabIndex        =   0
      Top             =   5160
      Width           =   6135
      Begin VB.CommandButton cmdSend 
         Caption         =   "Send"
         Height          =   345
         Left            =   5040
         Style           =   1  'Graphical
         TabIndex        =   2
         Top             =   2280
         Width           =   900
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   1815
         Left            =   120
         MultiLine       =   -1  'True
         TabIndex        =   1
         Top             =   360
         Width           =   5895
      End
      Begin VB.Label LabelLength 
         BackColor       =   &H80000005&
         Height          =   255
         Left            =   3840
         TabIndex        =   21
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label LabelHint1 
         BackColor       =   &H80000005&
         Caption         =   "Up to 250 characters allowed. Characters entered: "
         Height          =   255
         Left            =   240
         TabIndex        =   3
         Top             =   2400
         Width           =   4695
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
Attribute VB_Name = "FrmCommunicate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private mLastRightClickGrid As Object '最近一次右键点击的表格控件，供 mnuGridCopy_Click 使用
Private OraFilter As String
  
Private Sub CmdSaveData_Click()
  
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
      strTemp = SanitizeForCSV(.TextMatrix(i, 0))
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

Private Sub ComboType_Click(Index As Integer)
  Dim strMsgType$
  Dim strStatus$
  
  strMsgType = ComboType(0).Text
  If ComboType(1).Text = "Read" Then strStatus = "2"
  If ComboType(1).Text = "Replied" Then strStatus = "1"
  If ComboType(1).Text = "Unhandled" Then strStatus = "0"
  If ComboType(1).Text = "All" Then strStatus = "9"
    
  If strMsgType <> "All" Then
    If strStatus = "9" Then
      OraFilter = "msg_type='" & strMsgType & "'"
    Else
      OraFilter = "msg_type='" & strMsgType & "' and handled =" & strStatus
    End If
  
  ElseIf strMsgType = "All" Then
    If strStatus <> "9" Then
      OraFilter = "handled =" & strStatus
    Else
      OraFilter = ""
    End If
  End If
  
  Call showInfo(OraFilter)
  
End Sub


Private Sub Form_Load()
  Dim i%
  '
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1

  Me.Left = (Screen.Width - Me.Width) / 2
  Me.Top = 10 * (Screen.Height - Me.Height) / 20

  '2026-07-22 访客模式下不查询消息
  OraFilter = ""
  If Not GuestMode Then Call showInfo(OraFilter)
  ORARst.Filter = ""
  
    
  ComboType(0).AddItem "All": ComboType(0).AddItem "System Notice": ComboType(0).AddItem "User Feedback": ComboType(0).Text = "All"
  ComboType(1).AddItem "All": ComboType(1).AddItem "Read": ComboType(1).AddItem "Replied": ComboType(1).AddItem "Unhandled": ComboType(1).Text = "All"
  
  '2026-07-29 访客模式下直接从内置数据加载表格，不查询数据库
  If GuestMode Then
    ComboType(0).Enabled = False: ComboType(1).Enabled = False:
    CmdMarkRead.Enabled = False: CmdMarkUnread.Enabled = False: CmdDelete.Enabled = False
    
    Call ImportData.LoadGuestCSV(FrmCommunicate.HFGrid1, Nothing, "UserMessages.csv", 0)
    
    With HFGrid1
      .ColWidth(0) = 500
      .ColWidth(1) = 1550
      .ColWidth(2) = 800
      .ColWidth(3) = 800
      .ColWidth(4) = 900
      .ColWidth(5) = 3000
      .ColWidth(6) = 2500
      .ColWidth(7) = 1000
    
      For i = 0 To .Cols - 1
        .ColAlignment(i) = flexAlignLeftCenter
        .FixedAlignment(i) = flexAlignLeftCenter
      Next i
    
    End With
      
    Exit Sub
  End If
  
End Sub

Private Sub showInfo(OraFilter As String)
  HFGrid1.Clear
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  
  ORARst.Filter = OraFilter
  
  
  With HFGrid1
  
  If userID <> "liuw" Then
    strSQL = "select ddatetime,cfrom,cto,msg_type,content,reply,handled from T_OTHE_CROSS_COMMUNICATION where cfrom='liuw' and cto = '" & userID & "'" '来自管理员给自己的消息
    strSQL = strSQL & " UNION "
    strSQL = strSQL & "select ddatetime,cfrom,cto,msg_type,content,reply,handled from T_OTHE_CROSS_COMMUNICATION where cto = 'liuw'" '所是人发送给管理员的消息
    strSQL = strSQL & " order by ddatetime desc"
  Else
    strSQL = "select ddatetime,cfrom,cto,msg_type,content,reply,handled from T_OTHE_CROSS_COMMUNICATION where cto = 'liuw'" '所是人发送给管理员的消息
    strSQL = strSQL & " UNION "
    strSQL = strSQL & "select ddatetime,cfrom,cto,msg_type,content,reply,handled from T_OTHE_CROSS_COMMUNICATION where cfrom='liuw'" '来自管理员给自己的消息
    strSQL = strSQL & " order by ddatetime desc"
  End If

  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
  
    Set .DataSource = ORARst
    .ColWidth(0) = 500
    .ColWidth(1) = 1550
    .ColWidth(2) = 800
    .ColWidth(3) = 800
    .ColWidth(4) = 900
    .ColWidth(5) = 3000
    .ColWidth(6) = 2500
    .ColWidth(7) = 1000
    
    .TextMatrix(0, 1) = "Time"
    .TextMatrix(0, 2) = "From"
    .TextMatrix(0, 3) = "To"
    .TextMatrix(0, 4) = "Message Type"
    .TextMatrix(0, 5) = "Content"
    .TextMatrix(0, 6) = "Reply"
    .TextMatrix(0, 7) = "Status"
    
    For i = 1 To .Rows - 1
      .TextMatrix(i, 0) = i
      If .TextMatrix(i, 7) = "0" Then
        .TextMatrix(i, 7) = "Unhandled"
        If .TextMatrix(i, 3) = userID Then
          .Row = i
          For j = 0 To .Cols - 1
            .Col = j
            .CellBackColor = &H80C0FF
          Next j
        End If
        
      ElseIf .TextMatrix(i, 7) = "1" Then
        .TextMatrix(i, 7) = "Replied"
      ElseIf .TextMatrix(i, 7) = "2" Then
        .TextMatrix(i, 7) = "Read"
      End If
    Next i
  End If
  
  ORARst.Close
  End With
  
  ORAConn4.Close
End Sub

Private Sub cmdSend_Click()
  Dim Sendtime As Date
  Dim strTemp$(), i%
  Dim strContent$
  
  strContent = Text1.Text
  Sendtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  
  If userID <> "liuw" Then
    strSQL = "insert into T_OTHE_CROSS_COMMUNICATION(DDATETIME,CFROM,CTO,MSG_TYPE,CONTENT)"
    strSQL = strSQL + " Values(to_date('" & Sendtime & "', 'yyyy-mm-dd hh24:mi:ss'),'" & userID & "','liuw','User Feedback','" & strContent & "')"
  Else
    
    '只读取活跃用户
    strSQL = "select a.userid,b.LOGTIME" '
    strSQL = strSQL & " from T_OTHE_CROSS_USERS_PROPERTIES a"
    strSQL = strSQL & " right join (select userid,max(logtime) LOGTIME from T_OTHE_CROSS_USERS_LOG t group by userid) b"
    strSQL = strSQL & " on a.userid=b.userid where a.userid<>'liuw'"
    
    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
    ReDim strTemp(1 To ORARst.RecordCount)
    i = 1
    Do Until ORARst.EOF
      strTemp(i) = ORARst!userID
      i = i + 1
      ORARst.MoveNext
    Loop
    ORARst.Close
    
    '批量插入给所是用户的消息
    strSQL = "insert all"
    For i = 1 To UBound(strTemp)
      strSQL = strSQL + " into T_OTHE_CROSS_COMMUNICATION(DDATETIME,CFROM,CTO,MSG_TYPE,CONTENT,REPLY)"
      strSQL = strSQL + " Values(to_date('" & Sendtime & "', 'yyyy-mm-dd hh24:mi:ss'),'liuw','" & strTemp(i) & "','System Notice','" & strContent & "','System NoticeNone需Reply')"
    Next i
    strSQL = strSQL + " select 1 from dual"
  End If
  
  ORAConn4.Execute strSQL
  ORAConn4.Close
  Call showInfo(OraFilter)
  
  If userID <> "liuw" Then
    FrmMain.StatusBar1.Panels(1).Text = "Feedback submitted successfully"
  Else
    FrmMain.StatusBar1.Panels(1).Text = "System message sent successfully"
  End If
  
End Sub


Private Sub CmdRead_Click()
  Dim i%, Response%
  
  With HFGrid1
  If .TextMatrix(.Row, 1) = "" Then Exit Sub
    
  Text2.Text = .TextMatrix(.Row, 5)
  Text3.Text = .TextMatrix(.Row, 6)
  End With
  
End Sub


Private Sub CmdDelete_Click()
  Dim i%, Response%
  
  With HFGrid1
  If .TextMatrix(.Row, 1) = "" Then Exit Sub
  
  If (userID = .TextMatrix(.Row, 2)) Or (userID = "liuw") Then
    Response = MsgBox("Delete this message?", vbOKCancel)
    If Response = 2 Then Exit Sub
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    
    strSQL = "delete from T_OTHE_CROSS_COMMUNICATION"
    strSQL = strSQL + " where DDATETIME=to_date('" & .TextMatrix(.Row, 1) & "', 'yyyy-mm-dd hh24:mi:ss')  and cfrom='" & .TextMatrix(.Row, 2) & "' and cto='" & .TextMatrix(.Row, 3) & "'"
    ORAConn4.Execute strSQL
    ORAConn4.Close
    FrmMain.StatusBar1.Panels(1).Text = "Deleted successfully"
  Else
    MsgBox ("Cannot delete another user's messages"): Exit Sub
  End If
  
  End With
  Call showInfo(OraFilter)
  
End Sub

Private Sub CmdReply_Click()
  With HFGrid1
  
  If .TextMatrix(.Row, 4) = "System Notice" Then
    MsgBox ("System notices don't require a reply -- send a new message if you have questions"): Exit Sub
  End If
  
  If userID <> "liuw" Then
    MsgBox ("Please send a new message"): Exit Sub
  Else
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
    
    strSQL = "update T_OTHE_CROSS_COMMUNICATION"
    strSQL = strSQL + " set reply='" & Text3.Text & "',handled=1"
    strSQL = strSQL + " where DDATETIME=to_date('" & .TextMatrix(.Row, 1) & "', 'yyyy-mm-dd hh24:mi:ss')  and cfrom='" & .TextMatrix(.Row, 2) & "' and cto='" & .TextMatrix(.Row, 3) & "'"
    ORAConn4.Execute strSQL
    ORAConn4.Close
  End If
    
  End With
  Call showInfo(OraFilter)
  
  FrmMain.StatusBar1.Panels(1).Text = "Reply sent successfully"

End Sub

Private Sub CmdMarkRead_Click()
  
  With HFGrid1
  If userID = .TextMatrix(.Row, 3) Then  '消息接收用户 = 当前用户
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
  
    strSQL = "update T_OTHE_CROSS_COMMUNICATION"
    strSQL = strSQL + " set handled=2"
    strSQL = strSQL + " where DDATETIME=to_date('" & .TextMatrix(.Row, 1) & "', 'yyyy-mm-dd hh24:mi:ss')  and cfrom='" & .TextMatrix(.Row, 2) & "' and cto='" & .TextMatrix(.Row, 3) & "'"
    ORAConn4.Execute strSQL
    ORAConn4.Close
    FrmMain.StatusBar1.Panels(1).Text = "Mark as Read"
    
  Else
    MsgBox ("Cannot mark another user's messages"): Exit Sub
  End If
  
  End With
  Call showInfo(OraFilter)

End Sub


Private Sub CmdMarkUnread_Click()
  
  With HFGrid1
  If userID = .TextMatrix(.Row, 3) Then  '消息接收用户 = 当前用户
    ORAConn4.Open
    ORAConn4.CursorLocation = 3
  
    strSQL = "update T_OTHE_CROSS_COMMUNICATION"
    strSQL = strSQL + " set handled=0"
    strSQL = strSQL + " where DDATETIME=to_date('" & .TextMatrix(.Row, 1) & "', 'yyyy-mm-dd hh24:mi:ss')  and cfrom='" & .TextMatrix(.Row, 2) & "' and cto='" & .TextMatrix(.Row, 3) & "'"
    ORAConn4.Execute strSQL
    ORAConn4.Close
    FrmMain.StatusBar1.Panels(1).Text = "Mark as Unread"
    
  Else
    MsgBox ("Cannot mark another user's messages"): Exit Sub
  End If
  
  End With
  Call showInfo(OraFilter)

End Sub


Private Sub HFGrid1_DblClick()
  Call CmdRead_Click
End Sub

Private Sub Text1_Change()
  LabelLength.Caption = Len(Text1.Text)
End Sub

Private Sub Text3_Change()
  LabelLength2.Caption = Len(Text3.Text)
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
