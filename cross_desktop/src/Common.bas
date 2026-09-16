Attribute VB_Name = "Common"
Option Explicit


' 2022-04-17 添加多要素默认统计量的统计
Public MultiSel As Boolean   '是否多要素统计
Public NumFields%  '选择的多要素数量
Public SelFields_Initial$() '选择的多要素的初始名
Public SelNames_Initial$() '选择的多要素的中文名
Public SelFields_Restat$() '选择的多要素名称
Public FlagStats$() '选择的各要素对应的统计量


'******2022-3-8 定义字段名和字段意义******
Public Type COL_INFO
  Col_ID As String
  Col_name As String
End Type
Public COL_HOR() As COL_INFO
Public COL_Initial() As COL_INFO
Public COL_Restat() As COL_INFO
Public COL_tmp() As COL_INFO
'Public COL_QTR() As COL_INFO
'Public COL_YER() As COL_INFO
'Public COL_TEN() As COL_INFO
'Public COL_lst() As COL_INFO


Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, _
  ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Const GWL_STYLE = (-16)
Public Const WS_MAXIMIZEBOX = &H10000


'支持滚轮鼠标API---------------------------------
Public Const GWL_WNDPROC = (-4)
Public Const WM_COMMAND = &H111
Public Const WM_MBUTTONDOWN = &H207
Public Const WM_MBUTTONUP = &H208
Public Const WM_MOUSEWHEEL = &H20A

Public Oldwinproc As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Const LB_SETSEL = &H185 '2026-07-19 一次性设置/清除多选ListBox全部选中项，避免逐项循环

'2026-07-20 MDIForm 没有 ScaleHeight/ScaleWidth/ScaleMode 属性
'需要用 GetClientRect 直接取得 MDI 主窗体客户区的真实像素尺寸，
'再用 Screen.TwipsPerPixelX/Y 换算成与 Toolbar1.Height 等控件属性一致的缇（twip）单位
Public Type RECT
  Left As Long
  Top As Long
  Right As Long
  Bottom As Long
End Type
Public Declare Function GetClientRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long



'******使用Windows API函数可以获所经过的时间（毫秒级别）******
Declare Sub GetSystemTime Lib "kernel32" (lpSystemTime As SYSTEMTIME)
Public Type SYSTEMTIME
  wYear As Integer
  wMonth As Integer
  wDayOfWeek As Integer
  wDay As Integer
  wHour As Integer
  wMinute As Integer
  wSecond As Integer
  wMilliseconds As Integer
End Type
Public TimeNow As SYSTEMTIME
Public sngBTime As Single, sngETime As Single, sngDuration As Single      '看看查询时长



'多年平均值年限
Public Years_NormSURF$, Years_NormAWST$
'多年平均值来源：Norm_data(下发数据)或Real_data(实时统计)
Public Norm_Src$


Public FlagInputErr As Boolean  '2023-10-10 日期输入错误标识

Public crossID$, crossPWD$ 'CROSS系统对应的数据库、用户名、密码

Public FTPServer$, FTPPort$, FTPUser$, FTPPwd$ ' CROSS的ftp服务对应的服务器、端口、ftp用户名、密码

Public FileNum%, FileNum2%
Public ChangeCover As String '自动更新主界面图像
Public iFlagFrm% '判断最后一次加载的是哪一个窗口
Public FormNum% '主窗体中打开的子窗体的数量

Public ServerIP1$ '服务器1的IP地址(10.148.15.110)
Public ServerIP2$ '服务器1的IP地址(10.148.15.149)

Public WebRoot$ '服务器CROSS的根目录地址
Public Province$ '数据应用的省份




'******登录用户的相关信息*******
Public LoginSucceeded As Boolean
Public GuestMode As Boolean '2026-07-22 访客模式标志
Public userID$, UserPassword$, UserLevel%, UserZone$ '登录用户的相关信息:户名、密码、级别、区县市名
Public DataZone$, DataZones$   '所在地市名、可查询数据的地市列表
Public UserName$, UserDept$ '用户名，用户单位
Public IP_SaveData$ '2023-5-9：保存数据的ip，防止账号被盗用
Public IP_Ban As Boolean 'True-不可保存，False-可以保存
Public UserIP$ ' 2024-01-12：用户本机ip，用于临时站点表多用户同账号的识别
''''Public strZoneID$ '登录用户所在地区及边界

Public Logtime As Date '2024-10-9：某个ip用户的登录时间，记录在数据库表T_OTHE_CROSS_USERS_LOG

'******系统更新的相关信息*******
Public Update$ '是否自动更新
Public Update_day$ '上次更新日期
Public Vers_self$ '当前版本
Public Vers_date$ '当前版本日期
Public Vers_New$ '最新版本
Public Sys_Caption$ '系统名称
Public Department$ '单位名称
Public Path_Full$, Path_Patch$, Path_txt$ 'CROSS_Installer、CROSS_Patch和WhatsNew的网络相对路径
Public MainPic$ 'StartPic$,Ini文件夹下的主界面底图文件名


'*******调用surfer的相关变量*******
Public FlagSrfInstalled As Boolean ' 2026-8-13添加：访客登陆时，设为False，画等值线图时与用户交互，可变为True
Public SrfApp As Object
Public Plot As Object, Shapes As Object, MapFrame As Object
Public ContourMap As Object, ContourLevels As Object, Window As Object
Public ImageMap As Object, ColorMap As Object, BaseMap As Object
Public PostMap As Object '2023-11-29添加

''''Public FirstSet$ 'Yes-之前没设置过，No-之前设置过

Public SrfEdition$ 'surfer的版本，如8EN,11EN,8CN,11CN
Public lvlPath_Value$ ', srfName$
Public srfPath$, srfPath_ColdAir$
Public blnPath$, blnName$
Public SetSurfer$ 'NO-不启用surfer，YES-启用surfer
Public iWidth%, iHeight% 'surfer出图分辨率
Public iNumCols%, iNumRows% 'surfer插值的纵横网格数：由经纬度的差值计算

Public Export_Width$ 'surfer输出图形宽度
Public PROV_Name$, CITY_Name$, COUNTY_Name$ '用户选择的画图区域的省市县名称
Public District_Level$ '用户选择画图地区的级别
Public Grid_Size$ 'surfer插值的格网大小
Public Grid_Method As Variant 'surfer插值方法-srfKriging、srfInverseDistance，不能定义为string类型
Public District_ID$, District_Name$ '用户选择的图画区域ID及名称，如GD-广东；GZhou-高州
Public sngxMin!, sngyMin!, sngxMax!, sngyMax! '用户选择的图画区域的边界


'******调用数据库对象的相关变量******
Public SQLConn As New ADODB.Connection
Public SQLRst As New ADODB.Recordset

Public ORAConn1 As New ADODB.Connection
Public ORAConn2 As New ADODB.Connection
Public ORAConn3 As New ADODB.Connection  '2026-4-12 转移小时数据至111服务器
Public ORAConn4 As New ADODB.Connection  '2026-7-17 转移cross系统管理数据至111服务器
Public ORARst As New ADODB.Recordset
'Public ORARst2 As New ADODB.Recordset
'Public ORARst3 As New ADODB.Recordset

Public STARst As New ADODB.Recordset  '国家站信息
Public STARst2 As New ADODB.Recordset  '区域站信息

Public PARRst As New ADODB.Recordset

Public MDBConn As New ADODB.Connection
Public MDBRst As New ADODB.Recordset
Public MDBConn2 As New ADODB.Connection
Public MDBRst2 As New ADODB.Recordset

Public strSQL$, str1$
Public strSta$  '若站点类型为单选，则用来存放选择站点后的站点列表
Public strSURF$, strAWST$ '若站点类型为双选，则分别存放已选的国家站/自动站站点列表2017-01-24
Public strAllSta$ '同时输出两类型站点的站点列表字符串


Public StaType$  '用于判断查询国家站资料或自动站资料：SURF-国家站；AWST-自动站。

Public DataExtent$ ' 2026-8-19：数据空间范围

Public Auto_LogOff$ '2023-3-24：定时注销（1-默认每周日凌晨1点注销；0-不注销）

Public tempStacode$(), tempStaname$(), TempLongitude$(), TempLatitude$()
Public TempDateSTT$(), TempYearSTT%()
'气候分区，镇乡，县区，城市
Public tempArea$(), tempTown$(), tempCounty$(), tempCity$()


Public ShowLonLat As Boolean '保存结果时判断是否输出经纬度

Public TableName$, TableName1$, TableName2$ '至查询一个表时，写TableName，同时查询两个表时，写TableName1和TableName2

Public FrmMeteoPeriodExtStatus$ '定义FrmMeteoPeriodExt窗体状态

Public WebServiceURL, InputParameters, txtResult As String



'==================== 2026-08-13 图表另存为PNG图片 ====================
Private Type GUID
  Data1 As Long
  Data2 As Integer
  Data3 As Integer
  Data4(0 To 7) As Byte
End Type

Private Type GdiplusStartupInput
  GdiplusVersion As Long
  DebugEventCallback As Long
  SuppressBackgroundThread As Long
  SuppressExternalCodecs As Long
End Type

Private Declare Function GdiplusStartup Lib "gdiplus" (Token As Long, inputbuf As GdiplusStartupInput, ByVal outputbuf As Long) As Long
Private Declare Sub GdiplusShutdown Lib "gdiplus" (ByVal Token As Long)
Private Declare Function GdipCreateBitmapFromHBITMAP Lib "gdiplus" (ByVal hbm As Long, ByVal hPal As Long, bitmap As Long) As Long
Private Declare Function GdipSaveImageToFile Lib "gdiplus" (ByVal image As Long, ByVal fileNamePtr As Long, ByRef clsidEncoder As GUID, ByVal encoderParams As Long) As Long
Private Declare Function GdipDisposeImage Lib "gdiplus" (ByVal image As Long) As Long
Private Declare Function CLSIDFromString Lib "ole32" (ByVal strPtr As Long, ByRef clsid As GUID) As Long


'2026-08-19 MDI子窗体随主窗体缩放自动调整控件尺寸的逻辑（v2.14）
Public Const MDI_RESIZE_RIGHT_MARGIN As Long = 150
Public Const MDI_RESIZE_BOTTOM_MARGIN As Long = 150

Sub Main()
  FrmMain.Show
  frmStart.Show
  frmStart.SetFocus
End Sub
 
 
 
Public Function FlexScroll(ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
  On Error Resume Next
  Select Case wMsg
    Case WM_MOUSEWHEEL
      Select Case wParam
        Case -7864320         '向下滚
          SendKeys "{PGDN}"
        Case 7864320           '向上滚
          SendKeys "{PGUP}"
      End Select
  End Select
  FlexScroll = CallWindowProc(Oldwinproc, hWnd, wMsg, wParam, lParam)
End Function
'支持滚轮鼠标API---------------------------------



'ORACLE数据库无需参数Initial Catalog
Public Function ORAConnect(ServerName As String, userID As String, Password As String)
  ORAConnect = "Provider=MSDAORA; Data Source=" & ServerName & ";User ID=" & userID & ";Password=" & Password & ";Persist Security Info=False"
End Function



'读写"col_info.ini"文件
Public Sub ReadColInfo(FileName As String)
  Dim strTemp As String
  Dim strId As String
  Dim strName As String
  Dim i%
  
  If Dir(App.Path & "\Ini\" & FileName, vbNormal) = "" Then
    FrmMain.Show
    MsgBox App.Path & "\Ini" & "folder: file not found: " & FileName & "!" & Chr(10) + Chr(13) _
         & "Please contact the system developer", vbExclamation, "Notice"
    End
  End If
  
  FileNum = FreeFile
  Open App.Path & "\Ini\" & FileName For Input As #FileNum
    i = 0
    Do While Not EOF(FileNum)
      Line Input #FileNum, strTemp
        
      If Left(strTemp, 1) = "[" And Right(strTemp, 1) = "]" Then
        i = i + 1
        ReDim Preserve COL_tmp(1 To i)
        strTemp = Replace(strTemp, "[", "")
        strId = Replace(strTemp, "]", "")
        COL_tmp(i).Col_ID = strId
        Line Input #FileNum, strName
        COL_tmp(i).Col_name = strName
      End If
    Loop
  Close #FileNum
End Sub





'读写".ini"文件
Public Sub ReadWriteIni(FileName As String, FlaBring As String, strIniData, FlagRW As String)
  Dim strTemp As String
  
  If Dir(App.Path & "\Ini\" & FileName, vbNormal) = "" Then
    FrmMain.Show
    MsgBox App.Path & "\Ini" & "folder: file not found: " & FileName & "!" & Chr(10) + Chr(13) _
         & "Please contact the system developer", vbExclamation, "Notice"
    End
  End If
  
  If FlagRW = "r" Then
    FileNum = FreeFile
    Open App.Path & "\Ini\" & FileName For Input As #FileNum
      Do
        Line Input #FileNum, strTemp
      Loop Until strTemp = FlaBring
      Line Input #FileNum, strIniData
    Close #FileNum
  
  ElseIf FlagRW = "w" Then
    Dim iRow As Integer
    Dim strIniTemp() As String
    Dim i As Integer
        
    '将ini文件中原所有行读出到变量数组
    iRow = 0
    FileNum = FreeFile
    Open App.Path & "\Ini\" & FileName For Input As #FileNum
      Do While Not EOF(FileNum)
        Line Input #FileNum, strTemp
        iRow = iRow + 1
      Loop
    Close #FileNum
    ReDim strIniTemp(0 To iRow - 1)
    FileNum = FreeFile
    Open App.Path & "\Ini\" & FileName For Input As #FileNum
      For i = 0 To iRow - 1
        Line Input #FileNum, strIniTemp(i)
      Next i
    Close #FileNum
    '对原所有行中的值进行新值的替换
    For i = 0 To iRow - 1
      If strIniTemp(i) = FlaBring Then
        strIniTemp(i + 1) = strIniData
        Exit For
      End If
    Next i
    '将所有行再重新写入一个空白的ini文件
    FileNum = FreeFile
    Open App.Path & "\Ini\" & FileName For Output As #FileNum
      For i = 0 To iRow - 1
        Print #FileNum, strIniTemp(i)
      Next i
    Close #FileNum
  End If
End Sub


Public Sub send_FTP_File(FTPPath$, LocalPath$, FTPFile$)
  Dim i%
  
  FileNum = FreeFile
  Open App.Path & "\Temp\sendFTP_" & FTPFile & ".ftp" For Output As #FileNum
    Print #FileNum, "ftp"
    Print #FileNum, "open " & FTPServer
    Print #FileNum, FTPUser
    Print #FileNum, FTPPwd
    Print #FileNum, "prompt off"
    Print #FileNum, "binary"
    Print #FileNum, "cd " & FTPPath
    Print #FileNum, "lcd " & LocalPath
    Print #FileNum, "send " & FTPFile
    Print #FileNum, "disconnect"
    Print #FileNum, "bye"

  Close #FileNum
End Sub




Public Sub get_FTP_File(FTPPath$, LocalPath$, FTPFile$)
  Dim i%
'  Dim tempStr$()
'  tempStr = Split(FTPFile, ".")

  FileNum = FreeFile
  Open App.Path & "\Temp\getFTP_" & FTPFile & ".ftp" For Output As #FileNum
    Print #FileNum, "ftp"
    Print #FileNum, "open " & FTPServer
    Print #FileNum, FTPUser
    Print #FileNum, FTPPwd
    Print #FileNum, "prompt off"
    Print #FileNum, "binary"
    Print #FileNum, "cd " & FTPPath
    Print #FileNum, "lcd " & LocalPath
    Print #FileNum, "get " & FTPFile
    Print #FileNum, "disconnect"
    Print #FileNum, "bye"
  
  Close #FileNum
End Sub



Public Sub TimeStart() '计时开始
  GetSystemTime TimeNow
  sngBTime = 3600 * CSng(TimeNow.wHour)
  sngBTime = sngBTime + 60 * TimeNow.wMinute + TimeNow.wSecond + TimeNow.wMilliseconds / 1000
  FrmMain.StatusBar1.Panels(1).Text = "Querying, please wait"
End Sub

Public Sub TimeStop() '计时结束
  GetSystemTime TimeNow
  sngETime = 3600 * CSng(TimeNow.wHour)
  sngETime = sngETime + 60 * TimeNow.wMinute + TimeNow.wSecond + TimeNow.wMilliseconds / 1000
  sngDuration = Format(sngETime - sngBTime, "0.0")
  FrmMain.StatusBar1.Panels(1).Text = "Query time: " & sngDuration & " sec"
End Sub


Public Sub Sort_Cols_Rows(FrmGrid As Object, shift As Integer)
  On Error Resume Next
  Dim i%, j%, k%, l%
  Dim iRow%, iCols% '横向排序
  Dim iCol%, iRows% '纵向排序
  
  With FrmGrid
  
  If FrmGrid Is FrmMeteoHour.HFGrid1 Or FrmGrid Is FrmMeteoDay.HFGrid1 Or FrmGrid Is FrmMeteoTen.HFGrid1 Or FrmGrid Is FrmMeteoMon.HFGrid1 _
  Or FrmGrid Is FrmMeteoQtr.HFGrid1 Or FrmGrid Is FrmMeteoYer.HFGrid1 Or FrmGrid Is FrmMeteoPeriodYer.HFGrid1 Then
    iCols = .Cols - 7
  Else
    iCols = .Cols - 4
  End If
  
  
  '2026-08-11 空值哨兵按排序方向区分：升序用大数把空值推到末尾，降序用小数把空值推到末尾（末尾=最小值之后）
  If shift = 2 Then
    For i = 1 To .Rows - 1
      If .TextMatrix(i, .Col) = "" Then .TextMatrix(i, .Col) = "999999"
    Next i
  ElseIf shift = 4 Then
    For i = 1 To .Rows - 1
      If .TextMatrix(i, .Col) = "" Then .TextMatrix(i, .Col) = "-999999"
    Next i
  ElseIf shift = 1 Then
    For j = 3 To iCols
      If .TextMatrix(.Row, j) = "" Then .TextMatrix(.Row, j) = "999999"
    Next j
  End If


  If shift = 2 Then ' 按Ctrl键：从上到下升序排序
    If .TextMatrix(0, .Col) = "Township" Or .TextMatrix(0, .Col) = "County" Or .TextMatrix(0, .Col) = "City" Or .TextMatrix(0, .Col) = "Province" Then
      .Sort = 0: .Sort = 7: GoTo THE_END
    End If
  ElseIf shift = 4 Then ' 按Alt键：从上到下降序排序
    If .TextMatrix(0, .Col) = "Township" Or .TextMatrix(0, .Col) = "County" Or .TextMatrix(0, .Col) = "City" Or .TextMatrix(0, .Col) = "Province" Then
      .Sort = 0: .Sort = 8: GoTo THE_END
    End If
  End If


  If shift = 2 Then ' 按Ctrl键：从上到下升序排序
    .Sort = 0: .Sort = 3

  ElseIf shift = 4 Then ' 按Alt键：从上到下降序排序
    .Sort = 0: .Sort = 4

  ElseIf shift = 1 Then  '按SHIFT键：横向-从左到右升序排序（MSHFlexGrid的Sort属性不支持按行横向排序，仍需手工比较）
    For i = 3 To iCols - 1
      k = i
      For j = i + 1 To iCols
          If CSng(.TextMatrix(.Row, j)) < CSng(.TextMatrix(.Row, k)) Then
            k = j
          End If
      Next j
      .ColPosition(k) = i
    Next i
  End If

THE_END:
  If shift = 2 Or shift = 4 Then
    For i = 1 To .Rows - 1
      .TextMatrix(i, 0) = i
      If .TextMatrix(i, .Col) = "999999" Or .TextMatrix(i, .Col) = "-999999" Then .TextMatrix(i, .Col) = ""
    Next i
  ElseIf shift = 1 Then
    For j = 3 To iCols
      If .TextMatrix(.Row, j) = "999999" Then .TextMatrix(.Row, j) = ""
    Next j
  End If

  .Refresh
  End With

End Sub


Public Sub Sort_Cols(Grid As Object, shift As Integer)
  On Error Resume Next
  Dim i%, j%, k%, l%
  Dim iRow%, iCols% '横向排序
  Dim iCol%, iRows% '纵向排序
  
  With Grid
  
  If .TextMatrix(0, .Col) = "Level" Then
    For i = 1 To .Rows - 1
      If .TextMatrix(i, .Col) = "None" Then .TextMatrix(i, .Col) = 0
      If .TextMatrix(i, .Col) = "Weak" Then .TextMatrix(i, .Col) = 1
      If .TextMatrix(i, .Col) = "Moderate" Then .TextMatrix(i, .Col) = 2
      If .TextMatrix(i, .Col) = "Strong" Then .TextMatrix(i, .Col) = 3
      If .TextMatrix(i, .Col) = "Cold Wave" Then .TextMatrix(i, .Col) = 4
    Next i
  End If
    
  '2026-08-11 空值哨兵按排序方向区分：升序用大数、降序用小数，确保空值始终排在最小值之后
  If shift = 2 Then
    For i = 1 To .Rows - 1
      If .TextMatrix(i, .Col) = "" Then .TextMatrix(i, .Col) = "999999"
    Next i
  ElseIf shift = 4 Then
    For i = 1 To .Rows - 1
      If .TextMatrix(i, .Col) = "" Then .TextMatrix(i, .Col) = "-999999"
    Next i
  End If

  If shift = 2 Then ' 按Ctrl键：从上到下升序排序
    If .TextMatrix(0, .Col) = "Township" Or .TextMatrix(0, .Col) = "County" Or .TextMatrix(0, .Col) = "City" Or .TextMatrix(0, .Col) = "Province" Or Right(.TextMatrix(0, .Col), 1) = "日" Or Right(.TextMatrix(0, .Col), 2) = "日期" Then
      .Sort = 0: .Sort = 7: GoTo THE_END
    End If
  ElseIf shift = 4 Then ' 按Alt键：从上到下降序排序
    If .TextMatrix(0, .Col) = "Township" Or .TextMatrix(0, .Col) = "County" Or .TextMatrix(0, .Col) = "City" Or .TextMatrix(0, .Col) = "Province" Or Right(.TextMatrix(0, .Col), 1) = "日" Or Right(.TextMatrix(0, .Col), 2) = "日期" Then
      .Sort = 0: .Sort = 8: GoTo THE_END
    End If
  End If

  If shift = 2 Then ' 按Ctrl键：从上到下升序排序
    .Sort = 0: .Sort = 3

  ElseIf shift = 4 Then ' 按Alt键：从上到下降序排序
    .Sort = 0: .Sort = 4

  End If

THE_END:
  If .TextMatrix(0, .Col) = "Level" Then
    For i = 1 To .Rows - 1
      If .TextMatrix(i, .Col) = 0 Then .TextMatrix(i, .Col) = "None"
      If .TextMatrix(i, .Col) = 1 Then .TextMatrix(i, .Col) = "Weak"
      If .TextMatrix(i, .Col) = 2 Then .TextMatrix(i, .Col) = "Moderate"
      If .TextMatrix(i, .Col) = 3 Then .TextMatrix(i, .Col) = "Strong"
      If .TextMatrix(i, .Col) = 4 Then .TextMatrix(i, .Col) = "Cold Wave"
    Next i
  End If
  
  For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
    If .TextMatrix(i, .Col) = "999999" Or .TextMatrix(i, .Col) = "-999999" Then .TextMatrix(i, .Col) = ""
  Next i

  .Refresh
  End With
End Sub


Public Sub Sort_Rows(FrmGrid As Object, shift As Integer)
  Dim i%, j%, k%
  Dim iRow%, iCols%
  
  With FrmGrid
  iCols = .Cols - 7
        
  If shift = 1 Then  '按SHIFT键：横向-从左到右升序排序

    '2026-08-11 空值哨兵改为大数（与 Sort_Cols_Rows 一致），原先的 -9999 在升序时会把空值排到最前面，与"空值排到末尾"的规则相反
    For j = 3 To iCols
      If .TextMatrix(.Row, j) = "" Then .TextMatrix(.Row, j) = "999999"
    Next j

    For i = 3 To iCols - 1
      k = i
      For j = i + 1 To iCols
        If CSng(.TextMatrix(.Row, j)) < CSng(.TextMatrix(.Row, k)) Then
          k = j
        End If
      Next j
      .ColPosition(k) = i
    Next i

    For j = 3 To iCols
      If .TextMatrix(.Row, j) = "999999" Then .TextMatrix(.Row, j) = ""
    Next j
  End If
  .Refresh
  End With
End Sub


Public Sub Sort_Period_Cols(Grid As Object, shift As Integer)
  Dim i%, j%, k%, l%
  Dim iRow%, iCols% '横向排序
  Dim iCol%, iRows% '纵向排序

  With Grid

  '2026-08-11 空值哨兵按排序方向区分：升序用大数、降序用小数，确保空值始终排在最小值之后
  If shift = 2 Then
    For i = 1 To .Rows - 1
      If .TextMatrix(i, .Col) = "" Then .TextMatrix(i, .Col) = "999999"
    Next i
  ElseIf shift = 4 Then
    For i = 1 To .Rows - 1
      If .TextMatrix(i, .Col) = "" Then .TextMatrix(i, .Col) = "-999999"
    Next i
  End If

  If shift = 2 Then ' 按Ctrl键：从上到下升序排序
    If .TextMatrix(0, .Col) = "Township" Or .TextMatrix(0, .Col) = "County" Or .TextMatrix(0, .Col) = "City" Or .TextMatrix(0, .Col) = "Province" Or Right(.TextMatrix(0, .Col), 2) = "日期" Or Right(.TextMatrix(0, .Col), 2) = "Year" Then
      .Sort = 0: .Sort = 7: GoTo THE_END
    End If
  ElseIf shift = 4 Then ' 按Alt键：从上到下降序排序
    If .TextMatrix(0, .Col) = "Township" Or .TextMatrix(0, .Col) = "County" Or .TextMatrix(0, .Col) = "City" Or .TextMatrix(0, .Col) = "Province" Or Right(.TextMatrix(0, .Col), 2) = "日期" Or Right(.TextMatrix(0, .Col), 2) = "Year" Then
      .Sort = 0: .Sort = 8: GoTo THE_END
    End If
  End If

  If shift = 2 Then ' 按Ctrl键：从上到下升序排序
    .Sort = 0: .Sort = 3

  ElseIf shift = 4 Then ' 按Alt键：从上到下降序排序
    .Sort = 0: .Sort = 4

  End If

THE_END:
  For i = 1 To .Rows - 1
    .TextMatrix(i, 0) = i
    If .TextMatrix(i, .Col) = "999999" Or .TextMatrix(i, .Col) = "-999999" Then .TextMatrix(i, .Col) = ""
  Next i

  .Refresh
  End With
End Sub



'把PictureBox当前显示内容保存为PNG文件
Public Sub SaveChartAsPNG(Pic As Object, FileName As String)
  Dim gdipToken As Long
  Dim startupInput As GdiplusStartupInput
  Dim gdipBitmap As Long
  Dim pngClsid As GUID

  startupInput.GdiplusVersion = 1
  If GdiplusStartup(gdipToken, startupInput, 0) <> 0 Then
    MsgBox "GDI+ initialization failed -- cannot save PNG", vbExclamation, "Notice"
    Exit Sub
  End If

  If CLSIDFromString(strPtr("{557CF406-1A04-11D3-9A73-0000F81EF32E}"), pngClsid) <> 0 Then
    GdiplusShutdown gdipToken
    MsgBox "Failed to get the PNG encoder", vbExclamation, "Notice"
    Exit Sub
  End If

  If GdipCreateBitmapFromHBITMAP(Pic.image.Handle, 0, gdipBitmap) <> 0 Then
    GdiplusShutdown gdipToken
    MsgBox "Failed to read the chart bitmap", vbExclamation, "Notice"
    Exit Sub
  End If

  If GdipSaveImageToFile(gdipBitmap, strPtr(FileName), pngClsid, 0) <> 0 Then
    MsgBox "Failed to save the PNG file", vbExclamation, "Notice"
  End If

  GdipDisposeImage gdipBitmap
  GdiplusShutdown gdipToken
End Sub


'==================== 2026-08-13 表格另存为CSV ====================
Public Sub SaveGridAsCSV(Grid As Object, FileName As String)
  Dim i As Long, j As Long
  Dim strTemp As String

  If Len(Dir(FileName)) <> 0 Then Kill FileName
  Open FileName For Output As #1
  With Grid
    For i = 0 To .Rows - 1
      strTemp = SanitizeForCSV(.TextMatrix(i, 0))
      For j = 1 To .Cols - 1
        strTemp = strTemp & "," & SanitizeForCSV(.TextMatrix(i, j))
      Next j
      Print #1, strTemp
    Next i
  End With
  Close #1
End Sub


' 将MSHFlexGrid当前所选单元格以制表符分隔的文本形式复制到剪贴板
Public Sub CopyGridSelectionToClipboard(ByRef Grid As MSHFlexGrid)

  Dim RowTop As Long, RowBottom As Long
  Dim ColLeft As Long, ColRight As Long
  Dim RowList() As Long, ColList() As Long
  Dim RowCount As Long, ColCount As Long
  Dim r As Long, c As Long
  Dim OutText As String
  Dim LineText As String

  If Grid Is Nothing Then Exit Sub
  If Grid.Rows <= 0 Or Grid.Cols <= 0 Then Exit Sub

  RowTop = Grid.Row: RowBottom = Grid.RowSel
  If RowBottom < RowTop Then Swap_Long RowTop, RowBottom

  ColLeft = Grid.Col: ColRight = Grid.ColSel
  If ColRight < ColLeft Then Swap_Long ColLeft, ColRight

  ' 构建行索引列表：先放固定表头行，再放被选中的行
  ReDim RowList(0 To Grid.Rows - 1)
  RowCount = 0
  For r = 0 To Grid.FixedRows - 1
    RowList(RowCount) = r: RowCount = RowCount + 1
  Next r
  For r = RowTop To RowBottom
    If r >= Grid.FixedRows Then
      RowList(RowCount) = r: RowCount = RowCount + 1
    End If
  Next r

  ReDim ColList(0 To Grid.Cols - 1)
  ColCount = 0
  For c = 0 To Grid.FixedCols - 1
    ColList(ColCount) = c: ColCount = ColCount + 1
  Next c
  For c = ColLeft To ColRight
    If c >= Grid.FixedCols Then
      ColList(ColCount) = c: ColCount = ColCount + 1
    End If
  Next c

  If RowCount = 0 Or ColCount = 0 Then Exit Sub

  OutText = ""
  For r = 0 To RowCount - 1
    LineText = ""
    For c = 0 To ColCount - 1
      If c > 0 Then LineText = LineText & vbTab
      LineText = LineText & Grid.TextMatrix(RowList(r), ColList(c))
    Next c
    OutText = OutText & LineText & vbCrLf
  Next r

  Clipboard.Clear
  Clipboard.SetText OutText

End Sub


Private Sub Swap_Long(ByRef a As Long, ByRef b As Long)
  Dim Temp As Long
  Temp = a: a = b: b = Temp
End Sub


Public Sub ResizeFillGrid(ByVal Grid As Object, ByVal bottomAnchorTop As Long, ByVal NewClientWidth As Long)
  Dim newHeight As Long
  newHeight = bottomAnchorTop - Grid.Top
  If newHeight < 200 Then newHeight = 200 '极端窄小窗口下的保底，避免出现负高度报错
  Grid.Move Grid.Left, Grid.Top, NewClientWidth - MDI_RESIZE_RIGHT_MARGIN - Grid.Left, newHeight
End Sub


Public Function ResizeBottomAnchoredGrid(ByVal Grid As Object, ByVal NewClientHeight As Long, ByVal NewClientWidth As Long) As Long
  Dim newTop As Long
  newTop = NewClientHeight - MDI_RESIZE_BOTTOM_MARGIN - Grid.Height

  Grid.Move Grid.Left, newTop, NewClientWidth - MDI_RESIZE_RIGHT_MARGIN - Grid.Left, Grid.Height
  ResizeBottomAnchoredGrid = newTop
End Function


Public Sub ResizeStackedGridArray(ByVal Grids As Variant, ByVal bottomAnchorTop As Long, ByVal NewClientWidth As Long)
  Dim N As Integer, i As Integer, eachHeight As Long, curTop As Long, firstTop As Long
  N = UBound(Grids) - LBound(Grids) + 1
  firstTop = Grids(LBound(Grids)).Top
  eachHeight = (bottomAnchorTop - firstTop) \ N
  If eachHeight < 100 Then eachHeight = 100 '极端窄小窗口下的保底
  curTop = firstTop
  For i = LBound(Grids) To UBound(Grids)
    Grids(i).Move Grids(i).Left, curTop, NewClientWidth - MDI_RESIZE_RIGHT_MARGIN - Grids(i).Left, eachHeight
    curTop = curTop + eachHeight
  Next i
End Sub


Public Sub ResizeFrame1Width(ByVal Frame1 As Object, ByVal NewClientWidth As Long)
  Dim newWidth As Long
  newWidth = NewClientWidth - MDI_RESIZE_RIGHT_MARGIN - Frame1.Left
  If newWidth < 200 Then newWidth = 200 '极端窄小窗口下的保底
  Frame1.Width = newWidth
End Sub


Public Sub ResizeChildFrameWidth(ByVal ChildFrame As Object, ByVal ParentFrame As Object, ByVal MarginToParentRight As Long)
  Dim newWidth As Long
  newWidth = ParentFrame.Width - MarginToParentRight - ChildFrame.Left
  If newWidth < 100 Then newWidth = 100 '极端窄小窗口下的保底
  ChildFrame.Width = newWidth
End Sub


Public Sub ResizeStatusBarPanelsProportionally(ByVal SB As Object, ByVal NewTotalWidth As Long)
  Dim N As Integer, i As Integer
  N = SB.Panels.Count
  If N <= 0 Then Exit Sub

  Dim oldW() As Long, sumOld As Long
  ReDim oldW(1 To N)
  sumOld = 0
  For i = 1 To N
    oldW(i) = SB.Panels(i).Width
    sumOld = sumOld + oldW(i)
  Next i
  If sumOld <= 0 Then Exit Sub

  Dim sumNew As Long, w As Long
  sumNew = 0
  For i = 1 To N - 1
    w = CLng(CDbl(NewTotalWidth) * oldW(i) / sumOld)
    If w < 100 Then w = 100 '每个Panel的最小宽度保底
    SB.Panels(i).Width = w
    sumNew = sumNew + w
  Next i

  Dim lastW As Long
  lastW = NewTotalWidth - sumNew
  If lastW < 100 Then lastW = 100
  SB.Panels(N).Width = lastW
End Sub
