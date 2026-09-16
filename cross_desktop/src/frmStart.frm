VERSION 5.00
Begin VB.Form frmStart 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "System Login"
   ClientHeight    =   1890
   ClientLeft      =   795
   ClientTop       =   11100
   ClientWidth     =   3120
   ClipControls    =   0   'False
   Icon            =   "frmStart.frx":0000
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1890
   ScaleWidth      =   3120
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton CmdGuest 
      Appearance      =   0  'Flat
      Caption         =   "Guest"
      Height          =   345
      Left            =   1800
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   1320
      Width           =   1080
   End
   Begin VB.CommandButton CmdLogin 
      Appearance      =   0  'Flat
      Caption         =   "Log In"
      Default         =   -1  'True
      Height          =   345
      Left            =   120
      MaskColor       =   &H000000FF&
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   1320
      Width           =   1560
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000005&
      Height          =   975
      Left            =   120
      TabIndex        =   4
      Top             =   120
      Width           =   2775
      Begin VB.TextBox txtUserID 
         Appearance      =   0  'Flat
         BackColor       =   &H00FFFFFF&
         Height          =   270
         Left            =   840
         TabIndex        =   0
         ToolTipText     =   "Please enter a username"
         Top             =   240
         Width           =   1800
      End
      Begin VB.TextBox txtPassword 
         Appearance      =   0  'Flat
         Height          =   270
         IMEMode         =   3  'DISABLE
         Left            =   840
         PasswordChar    =   "*"
         TabIndex        =   1
         ToolTipText     =   "Please enter a password"
         Top             =   525
         Width           =   1800
      End
      Begin VB.Label Label2 
         BackColor       =   &H80000005&
         Caption         =   "Password"
         Height          =   255
         Left            =   120
         TabIndex        =   6
         Top             =   600
         Width           =   615
      End
      Begin VB.Label Label1 
         BackColor       =   &H80000005&
         Caption         =   "Username"
         Height          =   255
         Left            =   120
         TabIndex        =   5
         Top             =   240
         Width           =   615
      End
   End
End
Attribute VB_Name = "frmStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Form_Load()
  Dim tempStr$
  
   '使窗体始终居前
  SetWindowPos Me.hWnd, &HFFFF, 200, 200, Me.Width, Me.Height, &H1
  Me.Left = FrmMain.Left + 500
  Me.Top = FrmMain.Top + FrmMain.Height - Me.Height - 600
  
  If Dir(App.Path + "\*.bat") <> "" Then Kill App.Path + "\*.bat"
  If Dir(App.Path + "\Temp\*.bat") <> "" Then Kill App.Path + "\Temp\*.bat"
  If Dir(App.Path + "\Temp\*.ftp") <> "" Then Kill App.Path + "\Temp\*.ftp"
  
  FrmMain.Enabled = False

  Call ReadWriteIni("Config.ini", "[User]", tempStr, "r")
  txtUserID.Text = tempStr
  Call ReadWriteIni("Config.ini", "[Password]", tempStr, "r")
  If tempStr <> "不保存" Then
    txtPassword.Text = tempStr
  End If
  
  txtUserID.SelLength = Len(txtUserID.Text)

  
  Call ReadWriteIni("Config.ini", "[Set_Surfer]", SetSurfer, "r")
  
  Call ReadWriteIni("Config.ini", "[Years_NormSURF]", Years_NormSURF, "r")
  Call ReadWriteIni("Config.ini", "[Years_NormAWST]", Years_NormAWST, "r")
  Call ReadWriteIni("Config.ini", "[Norm_Src]", Norm_Src, "r")
'  Norm_Src = "GRMC" 'Real_stat
  
  '判断是否自动更新CROSS系统：是-是，No-否
  Call ReadWriteIni("Config.ini", "[Update]", Update, "r")
  
  '服务器1的IP地址
  Call ReadWriteIni("Config.ini", "[ServerIP1]", tempStr, "r")
  ServerIP1 = tempStr
  
  '服务器2的IP地址
  Call ReadWriteIni("Config.ini", "[ServerIP2]", tempStr, "r")
  ServerIP2 = tempStr
  
  '系统数据查询的省份
  Call ReadWriteIni("Config.ini", "[Province]", tempStr, "r")
'  txtProv.Text = tempStr
  Province = tempStr
  
  'ftp端口
  Call ReadWriteIni("Config.ini", "[FTPPort]", tempStr, "r")
  FTPPort = tempStr

  FTPServer = ServerIP1 & " " & FTPPort
  FTPUser = GetFTPUser(): FTPPwd = GetFTPPwd()

  str1 = "(DESCRIPTION = (ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = " & ServerIP1 & ")(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = orclnyqx)))"
  ORAConn1 = ORAConnect(str1, GetOracleUser(1), GetOraclePwd(1))

  str1 = "(DESCRIPTION = (ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = " & ServerIP2 & ")(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = climate)))"
  ORAConn2 = ORAConnect(str1, GetOracleUser(2), GetOraclePwd(2))
 
  str1 = "(DESCRIPTION = (ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = " & ServerIP2 & ")(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = climate)))"
  ORAConn3 = ORAConnect(str1, GetOracleUser(3), GetOraclePwd(3))
   
  str1 = "(DESCRIPTION = (ADDRESS_LIST =(ADDRESS = (PROTOCOL = TCP)(HOST = " & ServerIP2 & ")(PORT = 1521)))(CONNECT_DATA =(SERVICE_NAME = climate)))"
  ORAConn4 = ORAConnect(str1, GetOracleUser(4), GetOraclePwd(4))
  
End Sub


Private Sub Cmdlogin_Click()
 Dim Response%
  
 If txtUserID.Text = "" Then
   Response = MsgBox("No username entered -- enter Guest Mode?", vbOKCancel)
   If Response = 2 Then
     Exit Sub
   Else
     Call CmdGuest_Click
     Exit Sub
   End If
 End If
  
  Dim tempStr$, tempdate$, tempCounts As Long
  Dim i%, j%
  
  Dim tempStr2$()
  
  CmdLogin.Enabled = False '点了以后就不能再点了
  userID = txtUserID.Text
  UserPassword = txtPassword.Text
  
  FrmMain.StatusBar1.Panels(1).Text = "Connecting to server": DoEvents
    
  ORAConn1.Open  '2026-7-17 保留该110服务器连接，用于白名单阻隔
  ORAConn1.Close
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3 '****************很重要！此代码可以使取得ORARst的recordcount属性！！************************
  
  strSQL = "select c.userid,c.password,c.fullname,c.department,c.userlevel,c.userzone,c.DataZone,c.CROSS_ver,c.ip_savedata,max(c.logtime) LOGTIME,count(*) LOGCOUNTS"
  strSQL = strSQL & " from (select a.*, b.LOGTIME from T_OTHE_CROSS_USERS_PROPERTIES a left join T_OTHE_CROSS_USERS_LOG b on a.userid=b.userid where a.userid='" & userID & "') c"
  strSQL = strSQL & " group by c.userid,c.password,c.fullname,c.department,c.userlevel,c.userzone,c.DataZone,c.CROSS_ver,c.ip_savedata"
  
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockOptimistic
   
  
  If ORARst.EOF = True Then
    MsgBox "User does not exist", vbInformation, "Notice": txtUserID.SetFocus: txtUserID.SelLength = Len(txtUserID.Text)
    ORARst.Close
    ORAConn4.Close
    FrmMain.StatusBar1.Panels(1).Text = "Account does not exist"
    CmdLogin.Enabled = True
    Exit Sub
  End If

  FrmMain.StatusBar1.Panels(1).Text = "Verifying account info": DoEvents
  Call ReadWriteIni("Config.ini", "[User]", userID, "w")

  If to_MD5(UserPassword, 32) <> ORARst!Password Then '将密  码转换成32位MD5密  码
    LoginSucceeded = False
    MsgBox "Incorrect password -- please try again", vbInformation, "Notice":  txtPassword.SetFocus: txtPassword.SelLength = Len(txtPassword.Text)
    FrmMain.StatusBar1.Panels(1).Text = "Incorrect password -- please try again"
    CmdLogin.Enabled = True
    ORARst.Close
    ORAConn4.Close
    Exit Sub
  End If

  txtUserID.selStart = 0
  txtUserID.SelLength = Len(txtUserID.Text)
  txtUserID.SetFocus

  FrmMain.StatusBar1.Panels(1).Text = "Login successful": DoEvents
  LoginSucceeded = True
  GuestMode = False
  ExtStaInfoLoaded = False
  
  '*************** 2026-8-6 添加各类型站点选择状态默认值 ***************
  SelectSURFExt = 0.9
  SelectSURF = 0.9
  SelectAWST = 0.9
  
  sttDate_valid = "stt_date<>to_date('1899-09-09','yyyy-mm-dd')"
  
  '***************获取登录用户相关信息**********
  UserName = ORARst!FullName
  UserDept = ORARst!Department
  
  UserLevel = ORARst!UserLevel
  UserZone = ORARst!UserZone
  tempStr = ORARst!DataZone
  DataZone = tempStr  '授权给用户的数据区域：粤港澳、广东、21地市（县区与本地市同权）
  If ORARst!Logtime <> "" Then tempdate = ORARst!Logtime
  If ORARst!IP_SaveData <> "" Then IP_SaveData = ORARst!IP_SaveData
  tempCounts = ORARst!logcounts + 1
  ORARst.Update
  ORARst.Close
    

 '*********查询本机IP***********
  Dim strComputer$
  Dim objWMI      As Object
  Dim colIP       As Object
  Dim IP          As Object
  Dim tempIPStr$()
  
  IP_Ban = True '默认为true，若IP匹配对了，则赋值false

  strComputer = "."
  Set objWMI = GetObject("winmgmts://" & strComputer & "/root/cimv2")
  Set colIP = objWMI.ExecQuery("Select * from Win32_NetworkAdapterConfiguration where IPEnabled=TRUE")
  For Each IP In colIP
    If Not IsNull(IP.IPAddress) Then
      
      For i = LBound(IP.IPAddress) To UBound(IP.IPAddress)
        If Len(IP.IPAddress(i)) <= 15 Then  '长度不超过15位最长ip地址
          tempIPStr = Split(IP.IPAddress(i), ".")
          If UBound(tempIPStr) = 3 Then  '以“.”分隔共4段
            UserIP = IP.IPAddress(i)
          End If
        End If
        If InStr(IP_SaveData, IP.IPAddress(i)) <> 0 Then 'IP地址匹配
          IP_Ban = False
        End If
      Next
     
    End If
  Next
     
''***************记录登录用户登录信息:2017-01-22  记录登录用户IP，并 status 设为1:2017-01-22**********
  FrmMain.StatusBar1.Panels(1).Text = "Recording account login info": DoEvents
  Logtime = Format(Now, "yyyy-mm-dd hh:mm:ss")
  
  strSQL = "insert into T_OTHE_CROSS_USERS_LOG(UserID,UserIP,LOGTIME,STATUS) values('" & userID & "','" & UserIP & "',to_date('" & Logtime & "', 'yyyy-mm-dd hh24:mi:ss'),1)"
  ORAConn4.Execute strSQL
  
'***************2022-01-24 省市县用户获取可查询的数据范围（以city为条件）**********
'***************2015-09-06 市县用户获取登录用户可查询的数据范围**********
  FrmMain.StatusBar1.Panels(1).Text = "Getting user's data query range": DoEvents

  strSQL = "select * from T_OTHE_ZONE_RANGE_BASIC_TAB where ADMIN_ZONE='" & DataZone & "'"
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  tempStr = ORARst!DATARANGE
  tempStr2 = Split(tempStr, ",")
  DataZones = "'" & tempStr2(0) & "'"
  For i = 1 To UBound(tempStr2)
    DataZones = DataZones & ",'" & tempStr2(i) & "'"
  Next i
  ORARst.Close

'***************获取登录用户所在地区相关信息**********
  FrmMain.StatusBar1.Panels(1).Text = "Reading account region info": DoEvents
  strSQL = "select * from T_OTHE_ZONE_META_BASIC_TAB where ADMIN_ZONE='" & UserZone & "'"
  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  District_ID = ORARst!id
  sngxMin = ORARst!xMin: sngxMax = ORARst!xMax: sngyMin = ORARst!yMin: sngyMax = ORARst!yMax
  ORARst.Close
'*******************************************************
  
  blnName = District_ID & ".bln"
  blnPath = App.Path & "\Template\bln\" & blnName
  
  Call ReadWriteIni("Config.ini", "[Surfer_Edition]", SrfEdition, "r")
  Call ReadWriteIni("Config.ini", "[Grid_Size]", Grid_Size, "r")
  Call ReadWriteIni("Config.ini", "[Export_Width]", Export_Width, "r")
  iNumCols = CInt(100 * (sngxMax - sngxMin) / CSng(Grid_Size)): iNumRows = CInt(100 * (sngyMax - sngyMin) / CSng(Grid_Size))
  
  Call ReadWriteIni("Config.ini", "[Grid_Method]", tempStr, "r")
  If tempStr = "srfInverseDistance" Then Grid_Method = srfInverseDistance
  If tempStr = "srfKriging" Then Grid_Method = srfKriging
  If tempStr = "srfMinCurvature" Then Grid_Method = srfMinCurvature
  If tempStr = "srfNaturalNeighbor" Then Grid_Method = srfNaturalNeighbor
  If tempStr = "srfNearestNeighbor" Then Grid_Method = srfNearestNeighbor
  If tempStr = "srfRegression" Then Grid_Method = srfRegression
  If tempStr = "srfRadialBasis" Then Grid_Method = srfRadialBasis
  If tempStr = "srfTriangulation" Then Grid_Method = srfTriangulation

  srfPath = App.Path & "\Template\srf\" & District_ID & "_VALUE11.srf"
  srfPath_ColdAir = App.Path & "\Template\srf\" & District_ID & "_COLDAIR.srf"

  '2018-04-24对lvl加入语言的判定：
  lvlPath_Value = App.Path & "\Template\lvl\VALUE_EN.lvl"
         
      
  FrmMain.StatusBar1.Panels(1).Text = "Loading station info": DoEvents
  StaType = "SURF"  ' 2023-10-11 系统每次启动均默认为国家站数据
  bCheckInland = True '2026-8-10 虽然默认不选区域站，但是陆面区域站复选框勾选


'************************** 1.1 获取本地区国家站站点信息存入StaSURF2变量（2015-09-07）**************************
  strSQL = "select V_Prcode,b.REGION,V_CITY city,V_COUNTY county,v_town town,v01301 stacode,slm staname,v06001 longitude,v05001 latitude,STT_DATE DATESTT,Inland"
  strSQL = strSQL + ",v02301"  ' 2026-02-23：添加v02301用来判断水文站
  strSQL = strSQL + " from T_OTHE_STATION_META_BASIC_TAB a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b"
  strSQL = strSQL + " on a.v_city=b.admin_zone"
  If UserLevel = 1 Then strSQL = strSQL & " where V_PRCODE='广东'"
  If UserLevel >= 2 Then strSQL = strSQL & " where V_CITY='" & DataZone & "'"
  strSQL = strSQL + " and v02301 like '%A%'"
  strSQL = strSQL + " and " + sttDate_valid
  strSQL = strSQL + " and v01301<>'59486'"  '2026-7-13 排除59486号站，不参与初始加载
  strSQL = strSQL + " order by stacode"
  STARst.CursorLocation = adUseClient
  STARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  STARst.MoveFirst
  
  ReDim SURFInfo2(1 To STARst.RecordCount)
  StaNum_SURF = UBound(SURFInfo2)
  i = 1
  Do Until STARst.EOF
    SURFInfo2(i).Prov = STARst!v_prcode
    SURFInfo2(i).Region = STARst!Region
    SURFInfo2(i).City = STARst!City
    SURFInfo2(i).stacode = STARst!stacode
    SURFInfo2(i).StaType = "SURF"
    If Not IsNull(STARst!County) Then SURFInfo2(i).County = STARst!County
    If Not IsNull(STARst!Town) Then SURFInfo2(i).Town = STARst!Town
    If Not IsNull(STARst!staname) Then SURFInfo2(i).staname = STARst!staname
    SURFInfo2(i).Longitude = STARst!Longitude
    SURFInfo2(i).Latitude = STARst!Latitude
    SURFInfo2(i).DateSTT = STARst!DateSTT
    SURFInfo2(i).YearSTT = Year(STARst!DateSTT)
    SURFInfo2(i).Inland = STARst!Inland
    If InStr(STARst!v02301, "O") > 0 Then
      SURFInfo2(i).Hydro = 1
    Else
      SURFInfo2(i).Hydro = 0
    End If
    i = i + 1
    STARst.MoveNext
  Loop
  STARst.Close


'************************** 1.2 获取邻近地区国家站站点信息存入StaSURF1变量（2015-09-07）**************************
  strSQL = "select V_Prcode,b.REGION,V_CITY city,V_COUNTY county,v_town town,v01301 stacode,slm staname,v06001 longitude,v05001 latitude,STT_DATE DATESTT,Inland"
  strSQL = strSQL + ",v02301"  ' 2026-02-23：添加v02301用来判断水文站
  strSQL = strSQL + " from T_OTHE_STATION_META_BASIC_TAB a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b"
  strSQL = strSQL + " on a.v_city=b.admin_zone"
  strSQL = strSQL & " where v_city in (" & DataZones & ")" '2022-1-24 省市县均用city字段来限定条件
  strSQL = strSQL + " and v02301 like '%A%'"
  strSQL = strSQL + " and " + sttDate_valid
  strSQL = strSQL + " order by stacode"
  
  STARst.CursorLocation = adUseClient
  STARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  STARst.MoveFirst
  
  ReDim SURFInfo1(1 To STARst.RecordCount)
  i = 1
  Do Until STARst.EOF
    SURFInfo1(i).Prov = STARst!v_prcode
    SURFInfo1(i).Region = STARst!Region
    SURFInfo1(i).City = STARst!City
    SURFInfo1(i).stacode = STARst!stacode
    SURFInfo1(i).StaType = "SURF"
    If Not IsNull(STARst!County) Then SURFInfo1(i).County = STARst!County
    If Not IsNull(STARst!Town) Then SURFInfo1(i).Town = STARst!Town
    If Not IsNull(STARst!staname) Then SURFInfo1(i).staname = STARst!staname
    SURFInfo1(i).Longitude = STARst!Longitude
    SURFInfo1(i).Latitude = STARst!Latitude
    SURFInfo1(i).DateSTT = STARst!DateSTT
    SURFInfo1(i).YearSTT = Year(STARst!DateSTT)
    SURFInfo1(i).Inland = STARst!Inland
    If InStr(STARst!v02301, "O") > 0 Then
      SURFInfo1(i).Hydro = 1
    Else
      SURFInfo1(i).Hydro = 0
    End If
    For j = 1 To UBound(SURFInfo2)
      If SURFInfo1(i).stacode = SURFInfo2(j).stacode Then SURFInfo1(i).Selected = True: Exit For
    Next j
    If SURFInfo1(i).Inland = 0 Then
      SURFInfo1(i).Selected = False '默认不勾选非陆地站点（2023.2.14）
    End If
    i = i + 1
    STARst.MoveNext
  Loop
  STARst.Close
  
'************************** 2.1 获取本地区域站站点信息**************************
  strSQL = "select V_Prcode,b.REGION,V_CITY city,V_COUNTY county,v_town town,v01301 stacode,vf01015_cn staname,v06001 longitude,v05001 latitude,STT_DATE DATESTT,Inland"
  strSQL = strSQL + ",v02301"  ' 2026-02-23：添加v02301用来判断水文站
  strSQL = strSQL + " from T_OTHE_STATION_META_BASIC_TAB a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b"
  strSQL = strSQL + " on a.v_city=b.admin_zone"
  If UserLevel = 1 Then strSQL = strSQL & " where V_PRCODE='广东'"
  If UserLevel >= 2 Then strSQL = strSQL & " where V_CITY='" & DataZone & "'"
  strSQL = strSQL + " and v02301 like '%B%' and inland=1"
  strSQL = strSQL + " and " + sttDate_valid
  strSQL = strSQL + " order by stacode"
  STARst.CursorLocation = adUseClient
  STARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  STARst.MoveFirst
  
  ReDim AWSTInfo2(1 To STARst.RecordCount)
  StaNum_AWST = UBound(AWSTInfo2)
  i = 1
  Do Until STARst.EOF
    AWSTInfo2(i).Prov = STARst!v_prcode
    AWSTInfo2(i).Region = STARst!Region
    AWSTInfo2(i).City = STARst!City
    AWSTInfo2(i).stacode = STARst!stacode
    AWSTInfo2(i).StaType = "AWST"
    If Not IsNull(STARst!County) Then AWSTInfo2(i).County = STARst!County
    If Not IsNull(STARst!Town) Then AWSTInfo2(i).Town = STARst!Town
    If Not IsNull(STARst!staname) Then AWSTInfo2(i).staname = STARst!staname
    AWSTInfo2(i).Longitude = STARst!Longitude
    AWSTInfo2(i).Latitude = STARst!Latitude
    AWSTInfo2(i).DateSTT = STARst!DateSTT
    AWSTInfo2(i).YearSTT = Year(STARst!DateSTT)
    AWSTInfo2(i).Inland = STARst!Inland
    If InStr(STARst!v02301, "O") > 0 Then
      AWSTInfo2(i).Hydro = 1
    Else
      AWSTInfo2(i).Hydro = 0
    End If
    i = i + 1
    STARst.MoveNext
  Loop
  STARst.Close


'************************** 2.2 获取邻近地区区域站站点信息**************************
  strSQL = "select V_Prcode,b.REGION,V_CITY city,V_COUNTY county,v_town town,v01301 stacode,vf01015_cn staname,v06001 longitude,v05001 latitude,STT_DATE DATESTT,Inland"
  strSQL = strSQL + ",v02301"  ' 2026-02-23：添加v02301用来判断水文站
  strSQL = strSQL + " from T_OTHE_STATION_META_BASIC_TAB a"
  strSQL = strSQL + " left join (SELECT ADMIN_ZONE,MAX(region) region FROM T_OTHE_ZONE_RANGE_BASIC_TAB GROUP BY ADMIN_ZONE) b"
  strSQL = strSQL + " on a.v_city=b.admin_zone"
  strSQL = strSQL & " where v_city in (" & DataZones & ")"
  strSQL = strSQL + " and (v02301 like '%B%' or v02301 like '%O%')"  '2026-02-23：AddSelectHydrological Station
  strSQL = strSQL + " and " + sttDate_valid
  strSQL = strSQL + " order by stacode"
  STARst.CursorLocation = adUseClient
  STARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  STARst.MoveFirst
  
  ReDim AWSTInfo1(1 To STARst.RecordCount)
  i = 1
  Do Until STARst.EOF
    AWSTInfo1(i).Prov = STARst!v_prcode
    AWSTInfo1(i).Region = STARst!Region
    AWSTInfo1(i).City = STARst!City
    AWSTInfo1(i).stacode = STARst!stacode
    AWSTInfo1(i).StaType = "AWST"
    If Not IsNull(STARst!County) Then AWSTInfo1(i).County = STARst!County
    If Not IsNull(STARst!Town) Then AWSTInfo1(i).Town = STARst!Town
    If Not IsNull(STARst!staname) Then AWSTInfo1(i).staname = STARst!staname
    AWSTInfo1(i).Longitude = STARst!Longitude
    AWSTInfo1(i).Latitude = STARst!Latitude
    AWSTInfo1(i).DateSTT = STARst!DateSTT
    AWSTInfo1(i).YearSTT = Year(STARst!DateSTT)
    AWSTInfo1(i).Inland = STARst!Inland
    If InStr(STARst!v02301, "O") > 0 Then
      AWSTInfo1(i).Hydro = 1
    Else
      AWSTInfo1(i).Hydro = 0
    End If
      
    For j = 1 To UBound(AWSTInfo2)
      If AWSTInfo1(i).stacode = AWSTInfo2(j).stacode Then AWSTInfo1(i).Selected = True: Exit For
    Next j
    If AWSTInfo1(i).Inland = 0 Then
      AWSTInfo1(i).Selected = False '默认不勾选非陆地站点
    End If
    If AWSTInfo1(i).Hydro = 1 Then
      AWSTInfo1(i).Selected = False '默认不勾选水文站点（2026-03-23）
    End If
      
      
    i = i + 1
    STARst.MoveNext
  Loop
  STARst.Close

  If DataZone = "广东" Or DataZone = "粤港澳" Then
    strSURF = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_prcode ='广东' and v02301 like '%A%' and v01301<>'59486'and " + sttDate_valid + ")"
    strAWST = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_prcode ='广东' and v02301 like '%B%' and inland=1 and " + sttDate_valid + ")"
  Else
    strSURF = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_city ='" & DataZone & "' and v02301 like '%A%' and v01301<>'59486'and " + sttDate_valid + ")"
    strAWST = "stacode in (select v01301 from T_OTHE_STATION_META_BASIC_TAB where v_city ='" & DataZone & "' and v02301 like '%B%' and inland=1 and " + sttDate_valid + ")"
  End If
  
  
  '2026-8-19：数据空间范围（用于保存数据的时候写入数据表）
  If DataZone = "广东" Or DataZone = "粤港澳" Then
    DataExtent = "广东"
  Else
    DataExtent = DataZone
  End If
  

''''''''  '2023-2-15 登陆时，默认全选香港、澳门、海上站点外的所是站点
''''''''  FlagSURF_All = True
''''''''  FlagAWST_All = True
''''''''
''''''''  '2026-3-23 登陆时，默认不选水文站/海上站
''''''''  FlagHydro = False
''''''''  FlagOcean = False

  '************ 2022-2-16 将站点信息写入通用的变量 StaInfo ************
  If StaType = "SURF" Then
    StaNum = StaNum_SURF
    ReDim StaInfo(1 To StaNum)
    StaInfo = SURFInfo2
    strSta = strSURF
  ElseIf StaType = "AWST" Then
    StaNum = StaNum_AWST
    ReDim StaInfo(1 To StaNum)
    StaInfo = AWSTInfo2
    strSta = strAWST
  '************ 2022-2-25 需要给 StaInfo 重新排序，否则国家站和区域站站号混乱了 ************
  ElseIf StaType = "BOTH" Then
    StaNum = StaNum_SURF + StaNum_AWST
    ReDim StaInfo(1 To StaNum)
    For i = 1 To StaNum_SURF
      StaInfo(i).Region = SURFInfo2(i).Region
      StaInfo(i).City = SURFInfo2(i).City
      StaInfo(i).County = SURFInfo2(i).County
      StaInfo(i).Town = SURFInfo2(i).Town
      StaInfo(i).StaType = "SURF"
      StaInfo(i).stacode = SURFInfo2(i).stacode
      StaInfo(i).staname = SURFInfo2(i).staname
      StaInfo(i).Longitude = SURFInfo2(i).Longitude
      StaInfo(i).Latitude = SURFInfo2(i).Latitude
      StaInfo(i).DateSTT = SURFInfo2(i).DateSTT
      StaInfo(i).YearSTT = SURFInfo2(i).YearSTT
      StaInfo(i).Inland = SURFInfo2(i).Inland
      StaInfo(i).Hydro = SURFInfo2(i).Hydro
    Next i
    For i = StaNum_SURF + 1 To StaNum
      StaInfo(i).Region = AWSTInfo2(i - StaNum_SURF).Region
      StaInfo(i).City = AWSTInfo2(i - StaNum_SURF).City
      StaInfo(i).County = AWSTInfo2(i - StaNum_SURF).County
      StaInfo(i).Town = AWSTInfo2(i - StaNum_SURF).Town
      StaInfo(i).StaType = "AWST"
      StaInfo(i).stacode = AWSTInfo2(i - StaNum_SURF).stacode
      StaInfo(i).staname = AWSTInfo2(i - StaNum_SURF).staname
      StaInfo(i).Longitude = AWSTInfo2(i - StaNum_SURF).Longitude
      StaInfo(i).Latitude = AWSTInfo2(i - StaNum_SURF).Latitude
      StaInfo(i).DateSTT = AWSTInfo2(i - StaNum_SURF).DateSTT
      StaInfo(i).YearSTT = AWSTInfo2(i - StaNum_SURF).YearSTT
      StaInfo(i).Inland = AWSTInfo2(i - StaNum_SURF).Inland
      StaInfo(i).Hydro = AWSTInfo2(i - StaNum_SURF).Hydro
    Next i

    '2022-2-28 将StaInfo按照stacode从小到大排序（选择排序法）
    '定义临时站点信息类型变量，用于站点排序
    Dim tempStaInfo As StaInfo
    For i = LBound(StaInfo) To UBound(StaInfo) - 1
      For j = i + 1 To UBound(StaInfo)
        If StaInfo(j).stacode < StaInfo(i).stacode Then
          tempStaInfo = StaInfo(j)
          StaInfo(j) = StaInfo(i)
          StaInfo(i) = tempStaInfo
        End If
      Next j
    Next i

  End If
  '************ 2022-2-16 将站点信息写入通用的变量 StaInfo ************

 '***************2023-02-16 获取全省各级区域编码**********
  FrmMain.StatusBar1.Panels(1).Text = "Getting region codes": DoEvents
  Call GetTownCode
  Call GetCountyCode
  Call GetCityCode
  
 '***************2022-01-19 获取本地区用户可查询的镇乡**********
  FrmMain.StatusBar1.Panels(1).Text = "Loading township info": DoEvents
  Call GetTownInfo2(DataZone, "", "")

'***************2022-01-18 获取邻近地区用户可查询的镇乡**********
  Call GetTownInfo1(DataZones)

'***************2022-01-19 获取本地区的县区**********
  FrmMain.StatusBar1.Panels(1).Text = "Loading county info": DoEvents
  Call GetCountyInfo2(DataZone, "")

'***************2022-01-18 获取用户可查询的县区**********
  Call GetCountyInfo1(DataZones)

'***************2022-01-19 获取用户本地区的地市**********
  FrmMain.StatusBar1.Panels(1).Text = "Loading city info": DoEvents
  Call GetCityInfo2(DataZone)

'***************2022-01-17 获取用户可查询的地市**********
  Call GetCityInfo1(DataZones)
   
  '2022-01-18 默认总是查询站点尺度数据
  FlagZone = "STA" ': Stat_TOWN = False: Stat_CNTY = False: Stat_CITY = False
    
  '2026-07-16 默认总是查询站点尺度数据(区域统计)
  FlagZoneExt = "STA"
  
  TownNum = UBound(TownInfo2)
  CountyNum = UBound(CountyInfo2)
  CityNum = UBound(CityInfo2)
  
  
  If SetSurfer = "Yes" Then
    
    If Dir(blnPath) = "" Then '如果本地没是bln，则下载
      FrmMain.StatusBar1.Panels(1).Text = "Download .bln mask file via FTP"
      
      FileNum = FreeFile
      Open App.Path & "\Temp\getFTP_" & blnName & ".bat" For Output As #FileNum
        Print #FileNum, "ftp -s:" & App.Path & "\Temp\getFTP_" & blnName & ".ftp"
      Close #FileNum
      
      Call get_FTP_File("Template/bln", App.Path & "\Template\bln", blnName)
      Shell "cmd.exe /c " & App.Path + "\Temp\getFTP_" & blnName & ".bat", vbNormalFocus
    End If
    
    If Dir(srfPath) = "" Then '如果本地没是值.srf，则下载
      FrmMain.StatusBar1.Panels(1).Text = "Download .srf template via FTP"
      
      FileNum = FreeFile
      Open App.Path & "\Temp\getFTP_" & District_ID & "_VALUE11.srf" & ".bat" For Output As #FileNum
        Print #FileNum, "ftp -s:" & App.Path & "\Temp\getFTP_" & District_ID & "_VALUE11.srf" & ".ftp"
      Close #FileNum
      
      Call get_FTP_File("Template/srf", App.Path & "\Template\srf", District_ID & "_VALUE11.srf")
      Shell "cmd.exe /c " & App.Path + "\Temp\getFTP_" & District_ID & "_VALUE11.srf" & ".bat", vbNormalFocus
    End If

    
    If District_ID = "GD" Then '如果绘图区域是广东，则下载
        If Dir(srfPath_ColdAir) = "" Then '如果本地没是值.srf，则下载
          FrmMain.StatusBar1.Panels(1).Text = "Download GD_COLDAIR.srf template via FTP"
          
          FileNum = FreeFile
          Open App.Path & "\Temp\getFTP_" & District_ID & "_COLDAIR.srf" & ".bat" For Output As #FileNum
            Print #FileNum, "ftp -s:" & App.Path & "\Temp\getFTP_" & District_ID & "_COLDAIR.srf" & ".ftp"
          Close #FileNum
          
          Call get_FTP_File("Template/srf", App.Path & "\Template\srf", District_ID & "_COLDAIR.srf")
          Shell "cmd.exe /c " & App.Path + "\Temp\getFTP_" & District_ID & "_COLDAIR.srf" & ".bat", vbNormalFocus
        End If
    End If

    Set SrfApp = CreateObject("Surfer.Application")
  End If
  
  
  FrmMain.Show
  FrmMain.Enabled = True
  FrmMain.WindowState = vbMaximized '2026-08-19 v2.14 登录成功后主窗体默认最大化，兼顾大小屏用户

  FrmMain.Caption = userID & "@" & Sys_Caption & " " & Vers_self
  FrmMain.StatusBar1.Panels(1).Text = "Welcome to CROSS": DoEvents
            

  If StaType = "SURF" Then '每次重启都只选国家站
    FrmMain.StatusBar1.Panels(2).Text = "Selected: " & StaNum_SURF & " national stations"
    FrmMain.StatusBar1.Panels(3).Text = "No AWS stations selected"
  End If
            
    
  If Norm_Src = "CMA" Then
    FrmMain.StatusBar1.Panels(4).Text = "National station normals: CMA-issued normals"
  ElseIf Norm_Src = "GRMC" Then
    FrmMain.StatusBar1.Panels(4).Text = "National station normals: Climate Center " & Years_NormSURF & "Yr Avg"
  End If
  FrmMain.StatusBar1.Panels(5).Text = "AWS station normals: Climate Center " & Years_NormAWST & "Yr Avg"
                
        
  If SetSurfer = "Yes" Then
    FrmMain.StatusBar1.Panels(6).Text = "Plot region: " & UserZone
  ElseIf SetSurfer = "No" Then
    FrmMain.StatusBar1.Panels(6).Text = "Surfer plotting not enabled"
  End If
              
  Call ReadWriteIni("Config.ini", "[Password]", tempStr, "r")
  If tempStr <> "不保存" Then
    tempStr = txtPassword.Text
    Call ReadWriteIni("Config.ini", "[Password]", tempStr, "w")
  End If
              
  CmdLogin.Enabled = True '加载完以后就能再点了
  frmStart.Hide
 
 '**********************用户CROSSVersion号存入数据库*********************************
  ORAConn4.Execute "update T_OTHE_CROSS_USERS_PROPERTIES set CROSS_VER= '" & Vers_self & "' where UserID='" & userID & "'"
  ORAConn4.Close
  
  Call CheckMsg

 '**********************判断是否要进行系统更新检查*********************************
  If Update = "On" Then
    FrmMain.StatusBar1.Panels(1).Text = "Checking for the latest CROSS version"

    ORAConn4.Open
    ORAConn4.CursorLocation = 3 '****************很重要！此代码可以使取得ORARst的recordcount属性！！************************
    
    strSQL = "select max(version) version_new from T_OTHE_CROSS_UPDATE_FILE"
    ORARst.CursorLocation = adUseClient
    ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
    
    If ORARst.EOF <> True Then
      If Not IsNull(ORARst!Version_New) Then Vers_New = ORARst!Version_New
    End If
    
    ORARst.Close
    ORAConn4.Close
        
    Update_day = Format(Date, "yyyy-mm-dd")
    Call ReadWriteIni("Config.ini", "[Update_day]", Update_day, "w")
    
    If CSng(Vers_self) < CSng(Vers_New) Then
      FrmMain.StatusBar1.Panels(1).Text = "A new version of CROSS is available"
      FrmUpdate.Show
      FrmUpdate.SetFocus
    End If
    
    FrmMain.StatusBar1.Panels(1).Text = "Welcome to CROSS"
    
  End If

  frmWait.Hide

End Sub


Private Sub CmdGuest_Click()
  '2026-07-22 访客模式：不需要账号密  码，仅加载templdate\data\站点Info里的示例站点
  '地市Num等计数变量记为0，避免其他窗体的 UBound(...) 等代码因数组未初始化而报错

  Dim tempStr$

  CmdGuest.Enabled = False '点了以后就不能再点了，避免重复进入访客模式

  GuestMode = True
  userID = "Guest"
  UserPassword = ""
  UserName = "Guest"
  UserDept = ""
  UserLevel = 1
  UserZone = "广东"
  DataZone = ""
  DataZones = "''"
  UserIP = ""
  IP_SaveData = ""
  IP_Ban = True

  StaType = "SURF"
  FlagZone = "STA"
  FlagZoneExt = "STA"

  ReDim StaInfo(1 To 1)
  ReDim SURFInfo1(1 To 1): ReDim SURFInfo2(1 To 1)
  ReDim AWSTInfo1(1 To 1): ReDim AWSTInfo2(1 To 1)
  ReDim CityInfo1(1 To 1): ReDim CityInfo2(1 To 1)
  ReDim CountyInfo1(1 To 1): ReDim CountyInfo2(1 To 1)
  ReDim TownInfo1(1 To 1): ReDim TownInfo2(1 To 1)

  '2026-07-22 大区域统计功能是自己的站点缓存，直接置为已加载以跳过其查询逻辑
  ReDim ExtStaInfo1(1 To 1): ReDim ExtStaInfo2(1 To 1): ReDim ExtStaInfo(1 To 1)
  ReDim ExtProvInfo1(1 To 1): ReDim ExtProvInfo2(1 To 1)
  ReDim ExtCityInfo1(1 To 1): ReDim ExtCityInfo2(1 To 1)
  ReDim ExtCountyInfo1(1 To 1): ReDim ExtCountyInfo2(1 To 1)
  ExtStaNum = 0: ExtProvNum = 0: ExtCityNum = 0: ExtCountyNum = 0
  ExtStaInfoLoaded = True

  StaNum = 0: StaNum_SURF = 0: StaNum_AWST = 0
  TownNum = 0: CountyNum = 0: CityNum = 0

  '2026-07-29 若内置站点数据文件存在，则用真实SURF站点覆盖上面的占位数组；否则保持占位状态，不报错
  Call LoadGuestStationInfo

  '"1=0"确保任何拼接了这些片段的SQL都只会查出0行，而不会因为空字符串导致SQL语句语法错误
  strSta = "1=0"
  strSURF = "1=0"
  strAWST = "1=0"

  Call ReadWriteIni("Config.ini", "[Set_Surfer]", SetSurfer, "r")
  Call ReadWriteIni("Config.ini", "[Years_NormSURF]", Years_NormSURF, "r")
  Call ReadWriteIni("Config.ini", "[Years_NormAWST]", Years_NormAWST, "r")
  Call ReadWriteIni("Config.ini", "[Norm_Src]", Norm_Src, "r")

  Call ReadWriteIni("Config.ini", "[Surfer_Edition]", SrfEdition, "r")
  Call ReadWriteIni("Config.ini", "[Grid_Size]", Grid_Size, "r")
  Call ReadWriteIni("Config.ini", "[Export_Width]", Export_Width, "r")
  iNumCols = CInt(100 * (sngxMax - sngxMin) / CSng(Grid_Size)): iNumRows = CInt(100 * (sngyMax - sngyMin) / CSng(Grid_Size))

  Call ReadWriteIni("Config.ini", "[Grid_Method]", tempStr, "r")
  If tempStr = "srfInverseDistance" Then Grid_Method = srfInverseDistance
  If tempStr = "srfKriging" Then Grid_Method = srfKriging
  If tempStr = "srfMinCurvature" Then Grid_Method = srfMinCurvature
  If tempStr = "srfNaturalNeighbor" Then Grid_Method = srfNaturalNeighbor
  If tempStr = "srfNearestNeighbor" Then Grid_Method = srfNearestNeighbor
  If tempStr = "srfRegression" Then Grid_Method = srfRegression
  If tempStr = "srfRadialBasis" Then Grid_Method = srfRadialBasis
  If tempStr = "srfTriangulation" Then Grid_Method = srfTriangulation


  FrmMain.Show
  FrmMain.Enabled = True
  FrmMain.WindowState = vbMaximized '2026-08-19 v2.14 访客模式初始化后主窗体默认最大化，兼顾大小屏用户
  FrmMain.Caption = "Guest@" & Sys_Caption & " " & Vers_self
  FrmMain.StatusBar1.Panels(1).Text = "Welcome to CROSS Guest Mode": DoEvents
  FrmMain.StatusBar1.Panels(2).Text = "Guest Mode: Loaded " & StaNum & " national stations"
  FrmMain.StatusBar1.Panels(3).Text = "Guest Mode: AWS stations not loaded"
  FrmMain.StatusBar1.Panels(4).Text = ""
  FrmMain.StatusBar1.Panels(5).Text = ""


'  '2026-08-18 访客模式下完全不连接数据库，画图范围默认为广东，使用简化模板GD_tmp.srf
  sngxMin = 109.66: sngxMax = 117.19: sngyMin = 20.21: sngyMax = 25.53
  blnPath = App.Path & "\Template\bln\GD_tmp.bln"
  srfPath = App.Path & "\Template\srf\GD_tmp.srf"
  
  lvlPath_Value = App.Path & "\Template\lvl\VALUE_EN.lvl"

  If SetSurfer = "Yes" Then
    FrmMain.StatusBar1.Panels(6).Text = "Plot Region: Custom"
  ElseIf SetSurfer = "No" Then
    FrmMain.StatusBar1.Panels(6).Text = "Not confirmed whether Surfer11 is installed"
  End If
  FlagSrfInstalled = False '2026-8-13：默认认为访客没是安装Surfer11，等在画等值线图的时候确定

  '访客模式下也隐藏frm起点，不然用户以为没反应
  frmStart.Hide
End Sub

' 从 Preview\Data 下的内置数据加载访客模式的站点/区域信息：
' - Info站点SURF.csv / Info站点AWST.csv：站点信息（StaInfo 结构），
'   分别填充 SURFInfo1/SURFInfo2 和 AWSTInfo1/AWSTInfo2；StaInfo/ExtStaInfo*
'   仍只取 SURF（与 StaType 默认为 "SURF" 一致）。
' - Info送至wns.csv / InfoCounties.csv / InfoCities.csv：乡镇/区县/地市列表。
' 各 Info1/Info2 内容相同，因为访客没是"本地"与"跨区"之分。
Private Sub LoadGuestStationInfo()
  Dim N As Long

  N = LoadGuestStaInfoCSV("InfoStationSURF.csv", "SURF", SURFInfo2)
  If N > 0 Then
    StaNum_SURF = N
    SURFInfo1 = SURFInfo2
    StaInfo = SURFInfo2
    StaNum = UBound(StaInfo)
  End If

  N = LoadGuestStaInfoCSV("InfoStationAWST.csv", "AWST", AWSTInfo2)
  If N > 0 Then
    StaNum_AWST = N
    AWSTInfo1 = AWSTInfo2
  End If

  Call LoadGuestTownInfo("InfoTowns.csv")
  Call LoadGuestCountyInfo("InfoCounties.csv")
  Call LoadGuestCityInfo("InfoCities.csv")
  Call LoadGuestZoneCodes

  Call LoadGuestExtStationInfo
End Sub

'2026-08-13 简易CSV逐字符解析：正确处理带引号且引号内含逗号的字段（如 "B,E"），
'比 Split(Line, ",") 更可靠。原先靠"从某个固定列号往后的都算v02301"来兜底拼回最后一列，
'一旦后面又插入了新列（实际发生过：Info站点AWST.csv把DATESTT挪到了V02301后面），
'就会整行错位——V02301的内容被当成INLAND解析，IsNumeric判断失败，Inland静默变成0
Private Function ParseCSVLine(ByVal Line As String) As String()
  Dim Result() As String
  Dim N As Long
  Dim inQuotes As Boolean
  Dim cur As String
  Dim ch As String
  Dim p As Long

  N = 0
  ReDim Result(0 To 0)
  inQuotes = False
  cur = ""

  For p = 1 To Len(Line)
    ch = Mid(Line, p, 1)
    If ch = Chr(34) Then
      inQuotes = Not inQuotes
    ElseIf ch = "," And Not inQuotes Then
      ReDim Preserve Result(0 To N)
      Result(N) = cur
      N = N + 1
      cur = ""
    Else
      cur = cur & ch
    End If
  Next p
  ReDim Preserve Result(0 To N)
  Result(N) = cur

  ParseCSVLine = Result
End Function

'按列名（忽略大小写）在表头数组中查找对应下标，找不到返回-1
Private Function ColIndex(ByRef Headers() As String, ByVal ColName As String) As Integer
  Dim i As Integer
  For i = 0 To UBound(Headers)
    If UCase(Trim(Headers(i))) = UCase(ColName) Then
      ColIndex = i
      Exit Function
    End If
  Next i
  ColIndex = -1
End Function

' 从 Preview\Data\<Csv文件名称> 读取站点信息，写入 Info2()，StaType 字段填StaTypeStr（"SURF"/"AWST"）
Private Function LoadGuestStaInfoCSV(ByVal CsvFileName As String, ByVal StaTypeStr As String, ByRef Info2() As StaInfo) As Long
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String, Headers() As String
  Dim v02301 As String
  Dim colProv As Integer, colRegion As Integer, colCity As Integer, colCounty As Integer, colTown As Integer
  Dim colStacode As Integer, colStaname As Integer, colLon As Integer, colLat As Integer
  Dim colInland As Integer, colV02301 As Integer, colDateStt As Integer

  LoadGuestStaInfoCSV = 0

  FilePath = App.Path & "\Preview\Data\" & CsvFileName
  If Dir(FilePath) = "" Then Exit Function

  LineCount = 0
  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then LineCount = LineCount + 1
  Loop
  Close #FileNum
  ' 先数实际非空白行数，再按精确行数ReDim

  If LineCount > 0 Then
    ReDim Lines(0 To LineCount - 1)
    LineCount = 0

    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then
        Lines(LineCount) = tempStr
        LineCount = LineCount + 1
      End If
    Loop
    Close #FileNum
  End If

  If LineCount < 2 Then Exit Function

  Headers = ParseCSVLine(Lines(0))
  colProv = ColIndex(Headers, "V_PRCODE")
  colRegion = ColIndex(Headers, "REGION")
  colCity = ColIndex(Headers, "CITY")
  colCounty = ColIndex(Headers, "COUNTY")
  colTown = ColIndex(Headers, "TOWN")
  colStacode = ColIndex(Headers, "STACODE")
  colStaname = ColIndex(Headers, "STANAME")
  colLon = ColIndex(Headers, "LONGITUDE")
  colLat = ColIndex(Headers, "LATITUDE")
  colInland = ColIndex(Headers, "INLAND")
  colV02301 = ColIndex(Headers, "V02301")
  colDateStt = ColIndex(Headers, "DATESTT")

  ReDim Info2(1 To LineCount - 1)

  For i = 1 To LineCount - 1
    Fields = ParseCSVLine(Lines(i))

    Info2(i).Prov = Fields(colProv)
    Info2(i).Region = Fields(colRegion)
    Info2(i).City = Fields(colCity)
    Info2(i).County = Fields(colCounty)
    Info2(i).Town = Fields(colTown)
    Info2(i).StaType = StaTypeStr
    Info2(i).stacode = Fields(colStacode)
    Info2(i).staname = Fields(colStaname)
    If IsNumeric(Fields(colLon)) Then Info2(i).Longitude = CSng(Fields(colLon))
    If IsNumeric(Fields(colLat)) Then Info2(i).Latitude = CSng(Fields(colLat))
    Info2(i).DateSTT = Fields(colDateStt)
    If IsDate(Info2(i).DateSTT) Then Info2(i).YearSTT = Year(CDate(Info2(i).DateSTT))
    If IsNumeric(Fields(colInland)) Then Info2(i).Inland = CInt(Fields(colInland))

    v02301 = Fields(colV02301)
    If InStr(v02301, "O") > 0 Then
      Info2(i).Hydro = 1
    Else
      Info2(i).Hydro = 0
    End If

    '2026-08-13 自动站(AWST)现含陆面区域站/海洋区域站/水文站三种子类型
    If StaTypeStr = "AWST" Then
      Info2(i).Selected = False
    Else
      Info2(i).Selected = (Info2(i).Inland <> 0)
    End If
  Next i

  LoadGuestStaInfoCSV = UBound(Info2)
End Function

' 从 Preview\Data\ 读取乡镇列表（V_PRCODE,REGION,V_CITY,V_COUNTY,V_TOWN
Private Sub LoadGuestTownInfo(ByVal CsvFileName As String)
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String

  FilePath = App.Path & "\Preview\Data\" & CsvFileName
  If Dir(FilePath) = "" Then Exit Sub

  LineCount = 0
  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then LineCount = LineCount + 1
  Loop
  Close #FileNum
  ' 先数实际非空白行数，再按精确行数ReDim

  If LineCount > 0 Then
    ReDim Lines(0 To LineCount - 1)
    LineCount = 0

    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then
        Lines(LineCount) = tempStr
        LineCount = LineCount + 1
      End If
    Loop
    Close #FileNum
  End If

  If LineCount < 2 Then Exit Sub

  ReDim TownInfo2(1 To LineCount - 1)
  For i = 1 To LineCount - 1
    Fields = Split(Lines(i), ",")
    TownInfo2(i).Prov = Replace(Fields(0), Chr(34), "")
    TownInfo2(i).Region = Replace(Fields(1), Chr(34), "")
    TownInfo2(i).City = Replace(Fields(2), Chr(34), "")
    TownInfo2(i).County = Replace(Fields(3), Chr(34), "")
    TownInfo2(i).Town = Replace(Fields(4), Chr(34), "")
  Next i

  TownInfo1 = TownInfo2
  TownNum = UBound(TownInfo2)
End Sub

' 从 Preview\Data\ 读取区县列表（V_PRCODE,REGION,V_CITY,V_COUNTY），填充 区县Info1/区县Info2/区县Num
Private Sub LoadGuestCountyInfo(ByVal CsvFileName As String)
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String

  FilePath = App.Path & "\Preview\Data\" & CsvFileName
  If Dir(FilePath) = "" Then Exit Sub

  LineCount = 0
  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then LineCount = LineCount + 1
  Loop
  Close #FileNum
  ' 先数实际非空白行数，再按精确行数ReDim
  
  If LineCount > 0 Then
    ReDim Lines(0 To LineCount - 1)
    LineCount = 0

    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then
        Lines(LineCount) = tempStr
        LineCount = LineCount + 1
      End If
    Loop
    Close #FileNum
  End If

  If LineCount < 2 Then Exit Sub

  ReDim CountyInfo2(1 To LineCount - 1)
  For i = 1 To LineCount - 1
    Fields = Split(Lines(i), ",")
    CountyInfo2(i).Prov = Replace(Fields(0), Chr(34), "")
    CountyInfo2(i).Region = Replace(Fields(1), Chr(34), "")
    CountyInfo2(i).City = Replace(Fields(2), Chr(34), "")
    CountyInfo2(i).County = Replace(Fields(3), Chr(34), "")
  Next i

  CountyInfo1 = CountyInfo2
  CountyNum = UBound(CountyInfo2)
End Sub

' 从 Preview\Data\ 读取地市列表（V_PRCODE,REGION,V_CITY）填充 地市Info1/地市Info2/地市Num
Private Sub LoadGuestCityInfo(ByVal CsvFileName As String)
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String

  FilePath = App.Path & "\Preview\Data\" & CsvFileName
  If Dir(FilePath) = "" Then Exit Sub

  LineCount = 0
  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then LineCount = LineCount + 1
  Loop
  Close #FileNum
  ' 先数实际非空白行数，再按精确行数ReDim

  If LineCount > 0 Then
    ReDim Lines(0 To LineCount - 1)
    LineCount = 0

    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then
        Lines(LineCount) = tempStr
        LineCount = LineCount + 1
      End If
    Loop
    Close #FileNum
  End If

  If LineCount < 2 Then Exit Sub

  ReDim CityInfo2(1 To LineCount - 1)
  For i = 1 To LineCount - 1
    Fields = Split(Lines(i), ",")
    CityInfo2(i).Prov = Replace(Fields(0), Chr(34), "")
    CityInfo2(i).Region = Replace(Fields(1), Chr(34), "")
    CityInfo2(i).City = Replace(Fields(2), Chr(34), "")
  Next i

  CityInfo1 = CityInfo2
  CityNum = UBound(CityInfo2)
End Sub

' 从 Preview\Data\ 读取乡镇/区县/地市行政区划代码，分别填充 送至wn_Code/区县_Code/地市_Code
Private Sub LoadGuestZoneCodes()
  ReDim Town_Code(1 To 1)
  ReDim County_Code(1 To 1)
  ReDim City_Code(1 To 1)

  Call LoadGuestTownCode("CodeTowns.csv")
  Call LoadGuestCountyCode("CodeCounties.csv")
  Call LoadGuestCityCode("CodeCities.csv")

  Dim i As Long, j As Long

  For i = 1 To UBound(CityInfo2)
    For j = 1 To UBound(City_Code)
      If CityInfo2(i).City = City_Code(j).City Then
        CityInfo2(i).Code = City_Code(j).Code
        Exit For
      End If
    Next j
  Next i
  CityInfo1 = CityInfo2

  For i = 1 To UBound(CountyInfo2)
    For j = 1 To UBound(County_Code)
      If CountyInfo2(i).City = County_Code(j).City And CountyInfo2(i).County = County_Code(j).County Then
        CountyInfo2(i).Code = County_Code(j).Code
        Exit For
      End If
    Next j
  Next i
  CountyInfo1 = CountyInfo2

  For i = 1 To UBound(TownInfo2)
    For j = 1 To UBound(Town_Code)
      If TownInfo2(i).City = Town_Code(j).City And TownInfo2(i).County = Town_Code(j).County And TownInfo2(i).Town = Town_Code(j).Town Then
        TownInfo2(i).Code = Town_Code(j).Code
        Exit For
      End If
    Next j
  Next i
  TownInfo1 = TownInfo2
End Sub

' 从 Preview\Data\ 读取乡镇代码（ADMIN_CODE,TOWN,COUNTY,CITY），填充 Town_Code
Private Sub LoadGuestTownCode(ByVal CsvFileName As String)
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String

  FilePath = App.Path & "\Preview\Data\" & CsvFileName
  If Dir(FilePath) = "" Then Exit Sub

  LineCount = 0
  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then LineCount = LineCount + 1
  Loop
  Close #FileNum
  ' 先数实际非空白行数，再按精确行数ReDim

  If LineCount > 0 Then
    ReDim Lines(0 To LineCount - 1)
    LineCount = 0

    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then
        Lines(LineCount) = tempStr
        LineCount = LineCount + 1
      End If
    Loop
    Close #FileNum
  End If

  If LineCount < 2 Then Exit Sub

  ReDim Town_Code(1 To LineCount - 1)
  For i = 1 To LineCount - 1
    Fields = Split(Lines(i), ",")
    Town_Code(i).Code = Replace(Fields(0), Chr(34), "")
    Town_Code(i).Town = Replace(Fields(1), Chr(34), "")
    Town_Code(i).County = Replace(Fields(2), Chr(34), "")
    Town_Code(i).City = Replace(Fields(3), Chr(34), "")
  Next i
End Sub

' 从 Preview\Data\ 读取区县代码（ADMIN_CODE,COUNTY,CITY），填充区县_Code
Private Sub LoadGuestCountyCode(ByVal CsvFileName As String)
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String

  FilePath = App.Path & "\Preview\Data\" & CsvFileName
  If Dir(FilePath) = "" Then Exit Sub

  LineCount = 0
  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then LineCount = LineCount + 1
  Loop
  Close #FileNum
  ' 先数实际非空白行数，再按精确行数ReDim

  If LineCount > 0 Then
    ReDim Lines(0 To LineCount - 1)
    LineCount = 0

    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then
        Lines(LineCount) = tempStr
        LineCount = LineCount + 1
      End If
    Loop
    Close #FileNum
  End If

  If LineCount < 2 Then Exit Sub

  ReDim County_Code(1 To LineCount - 1)
  For i = 1 To LineCount - 1
    Fields = Split(Lines(i), ",")
    County_Code(i).Code = Replace(Fields(0), Chr(34), "")
    County_Code(i).County = Replace(Fields(1), Chr(34), "")
    County_Code(i).City = Replace(Fields(2), Chr(34), "")
  Next i
End Sub

' 从 Preview\Data\ 读取地市代码（ADMIN_CODE,CITY），填充地市_Code
Private Sub LoadGuestCityCode(ByVal CsvFileName As String)
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String

  FilePath = App.Path & "\Preview\Data\" & CsvFileName
  If Dir(FilePath) = "" Then Exit Sub

  LineCount = 0
  FileNum = FreeFile
  Open FilePath For Input As #FileNum
  Do Until EOF(FileNum)
    Line Input #FileNum, tempStr
    If Trim(tempStr) <> "" Then LineCount = LineCount + 1
  Loop
  Close #FileNum
  ' 先数实际非空白行数，再按精确行数ReDim

  If LineCount > 0 Then
    ReDim Lines(0 To LineCount - 1)
    LineCount = 0

    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then
        Lines(LineCount) = tempStr
        LineCount = LineCount + 1
      End If
    Loop
    Close #FileNum
  End If

  If LineCount < 2 Then Exit Sub

  ReDim City_Code(1 To LineCount - 1)
  For i = 1 To LineCount - 1
    Fields = Split(Lines(i), ",")
    City_Code(i).Code = Replace(Fields(0), Chr(34), "")
    City_Code(i).City = Replace(Fields(1), Chr(34), "")
  Next i
End Sub

' 大区域统计（FrmMeteoPeriodExt/Frm站点Ext）的跨省站点/省市县信息
Private Sub LoadGuestExtStationInfo()
  Dim FilePath As String
  Dim Lines() As String
  Dim LineCount As Long
  Dim tempStr As String
  Dim i As Long
  Dim Fields() As String

  ' ---- Station：InfoStationSURFExt.csv（STACODE,STANAME,LONGITUDE,LATITUDE,COUNTY,CITY,PROV,DATESTT） ----
  FilePath = App.Path & "\Preview\Data\InfoStationSURFExt.csv"
  If Dir(FilePath) <> "" Then
    LineCount = 0
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then LineCount = LineCount + 1
    Loop
    Close #FileNum

    If LineCount > 0 Then
      ReDim Lines(0 To LineCount - 1)
      LineCount = 0

      FileNum = FreeFile
      Open FilePath For Input As #FileNum
      Do Until EOF(FileNum)
        Line Input #FileNum, tempStr
        If Trim(tempStr) <> "" Then
          Lines(LineCount) = tempStr
          LineCount = LineCount + 1
        End If
      Loop
      Close #FileNum
    End If

    If LineCount >= 2 Then
      ReDim ExtStaInfo1(1 To LineCount - 1)
      For i = 1 To LineCount - 1
        Fields = Split(Lines(i), ",")
        ExtStaInfo1(i).stacode = Replace(Fields(0), Chr(34), "")
        ExtStaInfo1(i).staname = Replace(Fields(1), Chr(34), "")
        If IsNumeric(Fields(2)) Then ExtStaInfo1(i).Longitude = CSng(Fields(2))
        If IsNumeric(Fields(3)) Then ExtStaInfo1(i).Latitude = CSng(Fields(3))
        ExtStaInfo1(i).County = Replace(Fields(4), Chr(34), "")
        ExtStaInfo1(i).City = Replace(Fields(5), Chr(34), "")
        ExtStaInfo1(i).Prov = Replace(Fields(6), Chr(34), "")
        ExtStaInfo1(i).DateSTT = Replace(Fields(7), Chr(34), "")
        If IsDate(ExtStaInfo1(i).DateSTT) Then ExtStaInfo1(i).YearSTT = Year(CDate(ExtStaInfo1(i).DateSTT))
        ExtStaInfo1(i).StaType = "SURF"
        ExtStaInfo1(i).Inland = 1
        ExtStaInfo1(i).Hydro = 0
        ExtStaInfo1(i).Selected = (ExtStaInfo1(i).Prov = "广东")
      Next i
    End If
  End If

  ' ---- 省：InfoProvsExt.csv（V_PRCODE） ----
  FilePath = App.Path & "\Preview\Data\InfoProvsExt.csv"
  If Dir(FilePath) <> "" Then
    LineCount = 0
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then LineCount = LineCount + 1
    Loop
    Close #FileNum

    If LineCount > 0 Then
      ReDim Lines(0 To LineCount - 1)
      LineCount = 0

      FileNum = FreeFile
      Open FilePath For Input As #FileNum
      Do Until EOF(FileNum)
        Line Input #FileNum, tempStr
        If Trim(tempStr) <> "" Then
          Lines(LineCount) = tempStr
          LineCount = LineCount + 1
        End If
      Loop
      Close #FileNum
    End If

    If LineCount >= 2 Then
      ReDim ExtProvInfo1(1 To LineCount - 1)
      For i = 1 To LineCount - 1
        Fields = Split(Lines(i), ",")
        ExtProvInfo1(i).Prov = Replace(Fields(0), Chr(34), "")
      Next i
    End If
  End If

  ' ---- 市：InfoCitiesExt.csv（V_PRCODE,V_CITY） ----
  FilePath = App.Path & "\Preview\Data\InfoCitiesExt.csv"
  If Dir(FilePath) <> "" Then
    LineCount = 0
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then LineCount = LineCount + 1
    Loop
    Close #FileNum

    If LineCount > 0 Then
      ReDim Lines(0 To LineCount - 1)
      LineCount = 0

      FileNum = FreeFile
      Open FilePath For Input As #FileNum
      Do Until EOF(FileNum)
        Line Input #FileNum, tempStr
        If Trim(tempStr) <> "" Then
          Lines(LineCount) = tempStr
          LineCount = LineCount + 1
        End If
      Loop
      Close #FileNum
    End If

    If LineCount >= 2 Then
      ReDim ExtCityInfo1(1 To LineCount - 1)
      For i = 1 To LineCount - 1
        Fields = Split(Lines(i), ",")
        ExtCityInfo1(i).Prov = Replace(Fields(0), Chr(34), "")
        ExtCityInfo1(i).City = Replace(Fields(1), Chr(34), "")
      Next i
    End If
  End If

  ' ---- 县：InfoCountiesExt.csv（V_PRCODE,V_CITY,V_COUNTY） ----
  FilePath = App.Path & "\Preview\Data\InfoCountiesExt.csv"
  If Dir(FilePath) <> "" Then
    LineCount = 0
    FileNum = FreeFile
    Open FilePath For Input As #FileNum
    Do Until EOF(FileNum)
      Line Input #FileNum, tempStr
      If Trim(tempStr) <> "" Then LineCount = LineCount + 1
    Loop
    Close #FileNum

    If LineCount > 0 Then
      ReDim Lines(0 To LineCount - 1)
      LineCount = 0

      FileNum = FreeFile
      Open FilePath For Input As #FileNum
      Do Until EOF(FileNum)
        Line Input #FileNum, tempStr
        If Trim(tempStr) <> "" Then
          Lines(LineCount) = tempStr
          LineCount = LineCount + 1
        End If
      Loop
      Close #FileNum
    End If

    If LineCount >= 2 Then
      ReDim ExtCountyInfo1(1 To LineCount - 1)
      For i = 1 To LineCount - 1
        Fields = Split(Lines(i), ",")
        ExtCountyInfo1(i).Prov = Replace(Fields(0), Chr(34), "")
        ExtCountyInfo1(i).City = Replace(Fields(1), Chr(34), "")
        ExtCountyInfo1(i).County = Replace(Fields(2), Chr(34), "")
      Next i
    End If
  End If

  Call RebuildExtSelFromSelected
End Sub

Private Sub CheckMsg()
  Dim Cancel%
  
  ORAConn4.Open
  ORAConn4.CursorLocation = 3
  
  '****************针对普通用户*************
  If userID <> "liuw" Then
    strSQL = "select ddatetime,cfrom,cto,content,reply,handled from T_OTHE_CROSS_COMMUNICATION where cto = 'liuw' and cfrom = '" & userID & "'" '意见反馈
    strSQL = strSQL & " and handled=1" '管理员已答复，但用户未读
    strSQL = strSQL & " UNION "
    strSQL = strSQL & "select ddatetime,cfrom,cto,content,reply,handled from T_OTHE_CROSS_COMMUNICATION where cfrom='liuw' and cto = '" & userID & "'" 'System Notice
    strSQL = strSQL & " and handled=0" '未答复或未读
    strSQL = strSQL & " order by ddatetime desc"
    
  '****************针对管理员*************
  Else
    strSQL = "select ddatetime,cfrom,cto,content,reply,handled from T_OTHE_CROSS_COMMUNICATION where cto = 'liuw'" '意见反馈
    strSQL = strSQL & " and handled=0" '未答复或未读
    strSQL = strSQL & " order by ddatetime desc"
  End If

  ORARst.CursorLocation = adUseClient
  ORARst.Open strSQL, ORAConn4, adOpenDynamic, adLockReadOnly
  
  If ORARst.EOF <> True Then
    If userID <> "liuw" Then
      Cancel = MsgBox("Yes" & CStr(ORARst.RecordCount) & " new message(s)/system replies -- please check the Message Center.", vbYesNo + vbQuestion, "Notice")
    Else
      Cancel = MsgBox("Yes" & CStr(ORARst.RecordCount) & " new message(s) -- please check the Message Center.", vbYesNo + vbQuestion, "Notice")
    End If
      
  End If
  
  ORARst.Close
  ORAConn4.Close
  
  If Cancel = vbYes Then
    FrmCommunicate.Show
    FrmCommunicate.SetFocus
  End If
  
  
End Sub


Private Sub Form_Unload(Cancel As Integer)
  '设置全局变量为 false
  '不提示失败的登录
  LoginSucceeded = False
  End
End Sub

Private Sub txtUserID_GotFocus()
  txtUserID.selStart = 0
  txtUserID.SelLength = Len(txtUserID.Text)
End Sub
Private Sub txtPassword_GotFocus()
  txtPassword.selStart = 0
  txtPassword.SelLength = Len(txtPassword.Text)
End Sub


