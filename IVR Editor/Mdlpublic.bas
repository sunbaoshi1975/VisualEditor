Attribute VB_Name = "Mdlpublic"
Option Explicit

'Windows declarations
Public Declare Function SetCapture Lib "USER32" (ByVal hWnd As Long) As Long
Public Declare Function ClipCursor Lib "USER32" (lpRect As Any) As Long
Public Declare Function ReleaseCapture Lib "USER32" () As Long
Public Declare Function GetWindowRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetClientRect Lib "USER32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetCursorPos Lib "USER32" (lpPoint As POINTAPI) As Long
Public Declare Function SetCursorPos Lib "USER32" (ByVal x As Long, ByVal y As Long) As Long
Public Declare Function GetDC Lib "USER32" (ByVal hWnd As Long) As Long
Public Declare Function ReleaseDC Lib "USER32" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Public Declare Function GetStockObject Lib "GDI32" (ByVal nIndex As Long) As Long
Public Declare Function CreatePen Lib "GDI32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Public Declare Function SetROP2 Lib "GDI32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Public Declare Function Rectangle Lib "GDI32" (ByVal hdc As Long, ByVal x1 As Long, ByVal y1 As Long, ByVal x2 As Long, ByVal y2 As Long) As Long
Public Declare Function LineTo Lib "GDI32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Public Declare Function MoveToEx Lib "GDI32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Public Declare Function ClientToScreen Lib "USER32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long
Public Declare Function ScreenToClient Lib "USER32" (ByVal hWnd As Long, lpPoint As POINTAPI) As Long

Public Const NULL_BRUSH = 5
Public Const PS_SOLID = 0
Public Const R2_NOT = 6

'===================================

Public Declare Function GetSystemMetrics Lib "USER32" (ByVal nIndex As Long) As Long
Public Const SM_CXHSCROLL = 21
Public Const SM_CYHSCROLL = 3
Public Const SM_CYMENU = 15

'==============
' Owner menu
'==============
Public Declare Sub CopyMem Lib "KERNEL32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Public Declare Function GetSystemMenu Lib "USER32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Public Declare Function AppendMenu Lib "USER32" Alias "AppendMenuA" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal wIDNewItem As Long, ByVal lpNewItem As Any) As Long
Public Const MF_SEPARATOR = &H800&
Public Const MF_STRING = &H0&
Public Const IDM_ABOUT = 10000

'============================================
' Open and Save flow data common file dialog
'============================================
Public Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Declare Function GetOpenFileName Lib "comdlg32.dll" Alias "GetOpenFileNameA" (pOpenfilename As OPENFILENAME) As Long
Public Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
Public Const OFN_EXTENSIONDIFFERENT = &H400
Public Const OFN_FILEMUSTEXIST = &H1000

Public Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long
Public Declare Sub GetSystemTime Lib "KERNEL32" (lpSystemTime As SYSTEMTIME)
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

'' Sun added 2002-04-02
Public Declare Function sndPlaySound Lib "winmm.dll" Alias "sndPlaySoundA" (ByVal lpszSoundName As String, ByVal uFlags As Long) As Long
Public Declare Function CHOOSECOLOR Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

'**********************************************************
'系统环境常量
''' Sun added 2002-06-11
Public Const MAX_LANGUAGE_TYPE = 10

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ 目录及文件
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'' 临时文件
Public Const Def_Temp_FN = "tempfile"

'' Color
Public Const Def_LBL_PlayColor = &HFF&
Public Const Def_LBL_StopColor = &H80000012

''公司网址
Public Const Def_Web_NetSCSI = "http://www.pcsg.net/"
Public Const Def_Web_Unicom = "http://www.chinaunicom.com/"
Public Const Def_Web_MailToMe = "Mailto:sunboss@263.net"
Public Const Def_Web_MailToSCSI = "Mailto:tony@pcsg.net"

'' PI
Public Const Def_PI = 3.14159265358979
Public Const Def_GOLDEN_DIV = 0.618

Public Const Def_TWIPS_PER_CM = 567
Public Const Def_TWIPS_PER_INCH = 1440

Public Const Move_Big_Step = 150
Public Const Move_Small_Step = 15

Public Enum ControlState
    StateNothing = 0
    StateDragging
    StateSizing
End Enum

''配置文件
Public Const Def_INI_Config = "VEDIT.INI"

''INI 文件段名及行名
'''系统段
Public Const Def_INI_SEC_SYS = "SYSTEM"
Public Const Def_INI_ENTRY_VoxPath = "VoxPath"                  '' 系统语音路径

'查询条件
'Michael Added @ Sep,4,07
Public Const Def_INI_SEC_SearchCon = "SearchCondition"
Public Const Def_INI_ENTRY_SecName = "Name"
Public Const Def_INI_ENTRY_SecOP = "OP"
Public Const Def_INI_ENTRY_SecExp = "EXP"

'''数据库段
'''''''OLEDB
Public Const Def_INI_SEC_ODBC = "ODBC"
Public Const Def_INI_SEC_TAR_ODBC = "Tar_ODBC"
Public Const Def_INI_ENTRY_TYPE = "TYPE"
Public Const Def_INI_ENTRY_DBSERVER = "SQLSERVER"
Public Const Def_INI_ENTRY_DBNAME = "DATABASE"
'''''''SQLDB
Public Const Def_INI_ENTRY_DSN = "DSN"
Public Const Def_INI_ENTRY_USERID = "USERID"
Public Const Def_INI_ENTRY_PWD = "PASSWORD"

'''远程访问段
Public Const Def_INI_SEC_Remote = "REMOTE"
Public Const Def_INI_ENTRY_ServerIP = "SERVERIP"                '' 服务器地址
Public Const Def_INI_ENTRY_ServerPort = "SERVERPORT"            '' 服务器侦听端口号

Public Const Def_INI_SEC_Default = "DEFAULT"
Public Const Def_INI_ENTRY_Key_repeat = "Key_repeat"
Public Const Def_INI_ENTRY_Key_return = "Key_return"
Public Const Def_INI_ENTRY_Key_root = "Key_root"
Public Const Def_INI_ENTRY_Nd_parent = "Nd_parent"
Public Const Def_INI_ENTRY_Nd_root = "Nd_root"
Public Const Def_INI_ENTRY_Uservar = "Uservar"
Public Const Def_INI_ENTRY_Var = "Var"
'''本地段
Public Const Def_INI_SEC_Local = "LOCAL"
Public Const Def_INI_ENTRY_LocalPort = "LOCALPORT"              '' 本地机侦听端口号

'''用户工作环境段
Public Const Def_INI_SEC_OPTION = "OPTION"
Public Const Def_INI_ENTRY_OPT_FRAME_BG = "FrameBKColor"        '' Main Frame backgroud color
Public Const Def_INI_ENTRY_OPT_PAGE_HE = "PageHeight"           '' Work page height
Public Const Def_INI_ENTRY_OPT_PAGE_WD = "PageWidth"            '' Work page width
Public Const Def_INI_ENTRY_OPT_PAGE_BG = "PageBKColor"          '' Work page backgroud color
Public Const Def_INI_ENTRY_OPT_SHOW_NODE_CAPTION = "ShowNodeCaption"  '' Whether show Nodes' Caption
'Mike added @ 2008-1-29
Public Const Def_INI_ENTRY_OPT_SHOW_NODE_TAG = "ShowNodeTag"    '' Whether show Nodes' Tag
Public Const Def_INI_ENTRY_OPT_MAXCLIPSTACKS = "MaxClipStacks"  '' Maximum Stacks of ClipBoard
Public Const Def_INI_ENTRY_OPT_OpenInNewWindow = "OpenInNewWindow"    '' 在新窗口创建或打开流程
Public Const Def_INI_ENTRY_OPT_VoiceEditorPath = "VoiceEditorPath"    '' 语音编辑外接程序路径
Public Const Def_INI_ENTRY_OPT_ShowSysNodes = "ShowSysNodes"    '' 是否显示系统节点
''Michael Add @ 06-28-07 For Record File Type     0 - VOX, 1 - WAV
Public Const Def_INI_ENTRY_OPT_RecordFileType = "RecordFileType" '' 录音文件类型
Public Const Def_INI_ENTRY_OPT_VoiceFileType = "VoiceFileType"   '' 语音文件类型
'**********************************************************

'' MRU Items
'' Sun added 2012-05-07
Public Const Def_INI_SEC_MRUList = "MRUList"

'' TTS 设置
'' Michael Added @ 2007-11-29
Public Const Def_INI_SEC_TTS = "TTS"
Public Const Def_INI_ENTRY_TTSVOICE = "TTS_Voice"
Public Const Def_INI_ENTRY_TTSRATE = "TTS_Rate"
Public Const Def_INI_ENTRY_TTSVolume = "TTS_Volume"
Public Const Def_INI_ENTRY_TTSFormat = "TTS_Fromat"
' **********************************************************
'系统缺省值
Public Const Def_Default_TYPE1 = "ODBC"                         '' 使用ODBC
Public Const Def_Default_TYPE0 = "OLEDB"                        '' 使用OLEDB
Public Const Def_Default_DBServer = "PowerVoice"                '' 缺省 DB 服务器名称/IP地址
Public Const Def_Default_Database = "dbCallcenter"              '' 缺省 Database
Public Const Def_Default_DSN = "dbCallcenter"                   '' 缺省 ODBC DSN Name
Public Const Def_Default_USERID = "sa"                          '' 缺省 ODBC User ID
Public Const Def_Default_PWD = ""                               '' 缺省 ODBC Password

Public Const Def_Default_MaxClipBoardStacks = 15                '' 缺省粘贴板最大堆栈数

Public Const Def_Default_OpenInNewWindow = vbUnchecked          '' 缺省在新窗口创建或打开流程

Public Const Def_Default_ServerIP = "Callcenter"                '' 缺省服务器名称/IP地址
Public Const Def_Default_ServerPort = 10020                     '' 缺省服务器Port
Public Const Def_Default_LocalPort = 0                          '' 缺省本机Port
Public Const Def_Default_Key_repeat = 0
Public Const Def_Default_Key_return = 1
Public Const Def_Default_Key_root = 2
Public Const Def_Default_Nd_parent = 0
Public Const Def_Default_Nd_root = 256
Public Const Def_Default_Uservar = 0
Public Const Def_Default_Var = 0

'Michael Added @ 2007-11-29 for default TTS value
'Michael Modified @ 2007-12-5
#If Language = 0 Then
    Public Const Def_Default_TTSVoice = "Microsoft Simplified Chinese"  '' 默认发声引擎(Chinese)
#ElseIf Language = 1 Then
    Public Const Def_Default_TTSVoice = "Microsoft Sam"                 '' 默认发声引擎(English)
#End If
Public Const Def_Default_TTSRate = 0            '' 默认速率(-10 ~ 10)
Public Const Def_Default_TTSVolume = 100        '' 默认音量(Max)
Public Const Def_Default_TTSFormat = "SAFT8kHz16BitMono"

'**********************************************************
' TCP/IP
'' 消息包大小
Public Const Def_PKGLENGTH = 100

'' 用户类型指消息包头中的发送者&接收者
Public Const USER_MSG = 99                                      '' CTI SERVER
Public Const USER_PROGRAM = 50                                  '' 流程编辑子系统

'' 消息包类型
Public Const PKGTYP_CONTROL = 0                                 '' 控制包
Public Const PKGTYP_CALL = 1
Public Const PKGTYP_DATA = 2                                    '' 数据包
Public Const PKGTYP_STATUS = 3                                  '' 状态包
Public Const PKGTYP_TEST = 4                                    '' 状态包
Public Const PKGTYP_BILL = 5                                    '' 状态包
Public Const PKGTYP_ACK = 10                                    '' 响应包

'' EasyResource -> CTI 消息定义
Public Const sckEasyRS_Start = 5                                '' 系统启动
Public Const sckEasyRS_End = 7                                  '' 系统结束
Public Const sckEasyRS_Load = 1                                 '' 加载新资源

'' CTI -> EasyResource 消息定义
Public Const sckCTI_StartAck = 1005                             '' 系统启动响应
Public Const sckCTI_EndAck = 1007                               '' 系统结束响应
Public Const sckCTI_LoadAck = 1001                              '' 加载新资源响应

'****************************************************************
'*
'* 全局变量 & 结构体
'*
'****************************************************************
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++系统环境                                                    ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type typSystemEnvironment
    strVersionInfo       As String                              '' 版本信息
    strPath_Working      As String                              '' 系统工作目录
    strINI_File          As String                              '' 系统初始化文件目录
    strServerIP          As String                              '' 服务器地址
    intServerPort        As Integer                             '' 服务器侦听端口号
    strLocalIP           As String                              '' 本地机地址
    intLocalPort         As Integer                             '' 本地机侦听端口号
    intCurStep           As Integer                             '' 当前步骤
    crlCurItem           As VB.TextBox                          '' 当前条目
    strPath_SysVox       As String                              '' 语音资源根目录
    strOSUser            As String                              '' Windows用户
    blnUserODBC          As Boolean                             '' 使用OLE DB--0，或者ODBC--1
    strDSN               As String                              '' ODBC DSN Name
    strDBServer          As String                              '' SQL Server
    strDBName            As String                              '' SQL Server Database Name
    strUserID            As String                              '' ODBC User ID
    strPWD               As String                              '' ODBC Password
    strConString         As String                              '' OLEDB Connection String
    intConfigSet         As Integer                             '' Need Automatic Set Config
    intOpenInNewWindow   As Integer                             '' 在新窗口创建或打开
    strVoiceEditorPath   As String                              '' 语音编辑外接程序路径
    intShowSysNodes      As Integer                             '' 是否显示系统节点
    'Michael add @ 06-28-07 For Record File Type
    intRecFileType       As Integer                             ''录音文件类型
    intVoiFileType       As Integer                             ''语音文件类型
    
    ' Sun added 2007-10-20
    intPageHeight        As Long                                '' 打印纸张高度
    intPageWidth         As Long                                '' 打印纸张宽度
    ' Michael Added @ 2007-11-29 For TTS Setting
    intTTSVoice          As Integer
    strTTSVoice          As String                              ''TTS声音类型
    intTTSRate           As Integer                             ''TTS速率
    intTTSVolume         As Byte                                ''TTS音量
    strTTSFormat         As String                              ''TTS码率
    ' Michael Added @ 2007-12-10 for default page size
    intDefPageHeight     As Long                                ''默认打印纸张高度
    intDefPageWidth      As Long                                ''默认打印纸张宽度
    
End Type
Public gSystem           As typSystemEnvironment                '' 系统环境

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''MRU Menu Sun added 2012/05/07
Public Const PCS_MAX_MRUITEMS = 5                               '' 控制包
Private Type PCS_MRU_ITEM
    blnVisible          As Boolean                              '' 是否可见
    strItemName         As String                               '' 菜单项名
    intItemID           As Integer                              '' 流程ID
End Type
Public gMRUMenu(PCS_MAX_MRUITEMS) As PCS_MRU_ITEM

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''流程结构体 modify:Scott Data:2001/08/23
Private Type GETUSERFLOW
       IntP_id As Byte                                          ''流程编号
       StrP_name As String                                      ''流程名称
       StrP_description As String * 50                          ''流程描述
       StrP_version As String                                   ''流程版本号
       StrP_auther As String                                    ''流程作者
       StrP_user As String                                      ''流程用户
       DateP_createtime As Date                                 ''流程创建时间
       DataP_modifytime As Date                                 ''流程修改时间
End Type
Public gGetUserFlow As GETUSERFLOW

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++资源编号分类条件                                            ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type typRIDClassify
    LBound               As Long                                '' 下限
    UBound               As Long                                '' 上限
    Caption              As String                              '' 描述
End Type
Public gRID(4) As typRIDClassify                                '' 资源编号分类

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++查询条件                                                    ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type typQueryCondition
    intScale             As Integer                              '' 范围类别 0:分类 1:编号范围 2:指定编号
    intClass             As Integer                              '' 分类组合码
    lngStartID           As Long                                 '' 起始编号
    lngEndID             As Long                                 '' 终止编号
    lngThisID            As Long                                 '' 指定编号
End Type
Public gCondition        As typQueryCondition                   '' 查询条件

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ TCP/IP 消息                                                ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type SMsgHeader
    Sender               As Byte                                '' 发送者
    Receiver             As Byte                                '' 接收者
    PackageNo            As Byte                                '' 包号
    PackageType          As Byte                                '' 包类型
    PackageLen           As Integer                             '' 包长
End Type

''通信消息包结构定义
Public Type SCtiMsi_Package
    Msgheader            As SMsgHeader                           '' 消息包头
    command              As Integer                             '' 命令字
    sChannel             As Integer                             '' 发送通道号（主控通道）
    bytData              As Byte                                '' 组
    intData              As Integer                             '' 流程
    strdata              As String * 50                         '' 数据
End Type

'CHOOSECOLOR
Public Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

'CHOOSECOLOR Constants
Public Const CC_RGBINIT = &H1
Public Const CC_FULLOPEN = &H2
Public Const CC_ANYCOLOR = &H100

'================================
' Works Page properties
'================================
'Page size
Public gPageWidth As Long
Public gPageHeight As Long
'Page backcolor
Public gPageBackColor As OLE_COLOR

'=================================
'Works frame properties
'=================================
'Frame backcolor
Public gFrameBackColor As OLE_COLOR

'=================================
' Node properties
'=================================
'变量传递
Public Type GETUSERVAR
       IntP_id As Byte                                          ''流程编号
       IntN_index As Integer                                    ''节点Index
       IntN_id As Integer                                       ''节点ID
       IntN_no As Byte                                          ''节点编号
       IntN_page As Integer                                     ''节点所在页
       IntN_left As Integer                                     ''左
       IntN_top As Integer                                      ''上
       IntN_height As Integer                                   ''高
       IntN_width As Integer                                    ''宽
       StrN_description As String                               ''描述
       StrN_data1 As String                                     ''数据段1
       StrN_data2 As String                                     ''数据段2
End Type
Public gGetUserVar As GETUSERVAR

'=================================
' Node handles properties
'=================================
'Handles color
Public gNodeHandColor As OLE_COLOR

'=================================
' Global flow and node information variable
Public gCallFlow As clsIVRProgram

' Global clipboard variable
Public gClipBoard As clsIVRClipBoard

' Voice resouce id for playing
Public gintSoundResourceID As Long

' Node List Display Control Variable
Public g_NodeListShowType As Byte

' Node Caption Switch Variable
Public gblnShowNodeCaption As Boolean

' Mike : Node Tag Switch variable @ 08-1-29
Public gblnShowNodeTag As Boolean

'' 上次资源路径
Public gStrLastResPath As String

'Michael add
Public gbCallFromPro As Byte
Public gstrSQL As String
Public gbSearchFlag As Byte
