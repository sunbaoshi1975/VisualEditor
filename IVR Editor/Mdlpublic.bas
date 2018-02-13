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
'ϵͳ��������
''' Sun added 2002-06-11
Public Const MAX_LANGUAGE_TYPE = 10

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ Ŀ¼���ļ�
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'' ��ʱ�ļ�
Public Const Def_Temp_FN = "tempfile"

'' Color
Public Const Def_LBL_PlayColor = &HFF&
Public Const Def_LBL_StopColor = &H80000012

''��˾��ַ
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

''�����ļ�
Public Const Def_INI_Config = "VEDIT.INI"

''INI �ļ�����������
'''ϵͳ��
Public Const Def_INI_SEC_SYS = "SYSTEM"
Public Const Def_INI_ENTRY_VoxPath = "VoxPath"                  '' ϵͳ����·��

'��ѯ����
'Michael Added @ Sep,4,07
Public Const Def_INI_SEC_SearchCon = "SearchCondition"
Public Const Def_INI_ENTRY_SecName = "Name"
Public Const Def_INI_ENTRY_SecOP = "OP"
Public Const Def_INI_ENTRY_SecExp = "EXP"

'''���ݿ��
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

'''Զ�̷��ʶ�
Public Const Def_INI_SEC_Remote = "REMOTE"
Public Const Def_INI_ENTRY_ServerIP = "SERVERIP"                '' ��������ַ
Public Const Def_INI_ENTRY_ServerPort = "SERVERPORT"            '' �����������˿ں�

Public Const Def_INI_SEC_Default = "DEFAULT"
Public Const Def_INI_ENTRY_Key_repeat = "Key_repeat"
Public Const Def_INI_ENTRY_Key_return = "Key_return"
Public Const Def_INI_ENTRY_Key_root = "Key_root"
Public Const Def_INI_ENTRY_Nd_parent = "Nd_parent"
Public Const Def_INI_ENTRY_Nd_root = "Nd_root"
Public Const Def_INI_ENTRY_Uservar = "Uservar"
Public Const Def_INI_ENTRY_Var = "Var"
'''���ض�
Public Const Def_INI_SEC_Local = "LOCAL"
Public Const Def_INI_ENTRY_LocalPort = "LOCALPORT"              '' ���ػ������˿ں�

'''�û�����������
Public Const Def_INI_SEC_OPTION = "OPTION"
Public Const Def_INI_ENTRY_OPT_FRAME_BG = "FrameBKColor"        '' Main Frame backgroud color
Public Const Def_INI_ENTRY_OPT_PAGE_HE = "PageHeight"           '' Work page height
Public Const Def_INI_ENTRY_OPT_PAGE_WD = "PageWidth"            '' Work page width
Public Const Def_INI_ENTRY_OPT_PAGE_BG = "PageBKColor"          '' Work page backgroud color
Public Const Def_INI_ENTRY_OPT_SHOW_NODE_CAPTION = "ShowNodeCaption"  '' Whether show Nodes' Caption
'Mike added @ 2008-1-29
Public Const Def_INI_ENTRY_OPT_SHOW_NODE_TAG = "ShowNodeTag"    '' Whether show Nodes' Tag
Public Const Def_INI_ENTRY_OPT_MAXCLIPSTACKS = "MaxClipStacks"  '' Maximum Stacks of ClipBoard
Public Const Def_INI_ENTRY_OPT_OpenInNewWindow = "OpenInNewWindow"    '' ���´��ڴ����������
Public Const Def_INI_ENTRY_OPT_VoiceEditorPath = "VoiceEditorPath"    '' �����༭��ӳ���·��
Public Const Def_INI_ENTRY_OPT_ShowSysNodes = "ShowSysNodes"    '' �Ƿ���ʾϵͳ�ڵ�
''Michael Add @ 06-28-07 For Record File Type     0 - VOX, 1 - WAV
Public Const Def_INI_ENTRY_OPT_RecordFileType = "RecordFileType" '' ¼���ļ�����
Public Const Def_INI_ENTRY_OPT_VoiceFileType = "VoiceFileType"   '' �����ļ�����
'**********************************************************

'' MRU Items
'' Sun added 2012-05-07
Public Const Def_INI_SEC_MRUList = "MRUList"

'' TTS ����
'' Michael Added @ 2007-11-29
Public Const Def_INI_SEC_TTS = "TTS"
Public Const Def_INI_ENTRY_TTSVOICE = "TTS_Voice"
Public Const Def_INI_ENTRY_TTSRATE = "TTS_Rate"
Public Const Def_INI_ENTRY_TTSVolume = "TTS_Volume"
Public Const Def_INI_ENTRY_TTSFormat = "TTS_Fromat"
' **********************************************************
'ϵͳȱʡֵ
Public Const Def_Default_TYPE1 = "ODBC"                         '' ʹ��ODBC
Public Const Def_Default_TYPE0 = "OLEDB"                        '' ʹ��OLEDB
Public Const Def_Default_DBServer = "PowerVoice"                '' ȱʡ DB ����������/IP��ַ
Public Const Def_Default_Database = "dbCallcenter"              '' ȱʡ Database
Public Const Def_Default_DSN = "dbCallcenter"                   '' ȱʡ ODBC DSN Name
Public Const Def_Default_USERID = "sa"                          '' ȱʡ ODBC User ID
Public Const Def_Default_PWD = ""                               '' ȱʡ ODBC Password

Public Const Def_Default_MaxClipBoardStacks = 15                '' ȱʡճ��������ջ��

Public Const Def_Default_OpenInNewWindow = vbUnchecked          '' ȱʡ���´��ڴ����������

Public Const Def_Default_ServerIP = "Callcenter"                '' ȱʡ����������/IP��ַ
Public Const Def_Default_ServerPort = 10020                     '' ȱʡ������Port
Public Const Def_Default_LocalPort = 0                          '' ȱʡ����Port
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
    Public Const Def_Default_TTSVoice = "Microsoft Simplified Chinese"  '' Ĭ�Ϸ�������(Chinese)
#ElseIf Language = 1 Then
    Public Const Def_Default_TTSVoice = "Microsoft Sam"                 '' Ĭ�Ϸ�������(English)
#End If
Public Const Def_Default_TTSRate = 0            '' Ĭ������(-10 ~ 10)
Public Const Def_Default_TTSVolume = 100        '' Ĭ������(Max)
Public Const Def_Default_TTSFormat = "SAFT8kHz16BitMono"

'**********************************************************
' TCP/IP
'' ��Ϣ����С
Public Const Def_PKGLENGTH = 100

'' �û�����ָ��Ϣ��ͷ�еķ�����&������
Public Const USER_MSG = 99                                      '' CTI SERVER
Public Const USER_PROGRAM = 50                                  '' ���̱༭��ϵͳ

'' ��Ϣ������
Public Const PKGTYP_CONTROL = 0                                 '' ���ư�
Public Const PKGTYP_CALL = 1
Public Const PKGTYP_DATA = 2                                    '' ���ݰ�
Public Const PKGTYP_STATUS = 3                                  '' ״̬��
Public Const PKGTYP_TEST = 4                                    '' ״̬��
Public Const PKGTYP_BILL = 5                                    '' ״̬��
Public Const PKGTYP_ACK = 10                                    '' ��Ӧ��

'' EasyResource -> CTI ��Ϣ����
Public Const sckEasyRS_Start = 5                                '' ϵͳ����
Public Const sckEasyRS_End = 7                                  '' ϵͳ����
Public Const sckEasyRS_Load = 1                                 '' ��������Դ

'' CTI -> EasyResource ��Ϣ����
Public Const sckCTI_StartAck = 1005                             '' ϵͳ������Ӧ
Public Const sckCTI_EndAck = 1007                               '' ϵͳ������Ӧ
Public Const sckCTI_LoadAck = 1001                              '' ��������Դ��Ӧ

'****************************************************************
'*
'* ȫ�ֱ��� & �ṹ��
'*
'****************************************************************
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ϵͳ����                                                    ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type typSystemEnvironment
    strVersionInfo       As String                              '' �汾��Ϣ
    strPath_Working      As String                              '' ϵͳ����Ŀ¼
    strINI_File          As String                              '' ϵͳ��ʼ���ļ�Ŀ¼
    strServerIP          As String                              '' ��������ַ
    intServerPort        As Integer                             '' �����������˿ں�
    strLocalIP           As String                              '' ���ػ���ַ
    intLocalPort         As Integer                             '' ���ػ������˿ں�
    intCurStep           As Integer                             '' ��ǰ����
    crlCurItem           As VB.TextBox                          '' ��ǰ��Ŀ
    strPath_SysVox       As String                              '' ������Դ��Ŀ¼
    strOSUser            As String                              '' Windows�û�
    blnUserODBC          As Boolean                             '' ʹ��OLE DB--0������ODBC--1
    strDSN               As String                              '' ODBC DSN Name
    strDBServer          As String                              '' SQL Server
    strDBName            As String                              '' SQL Server Database Name
    strUserID            As String                              '' ODBC User ID
    strPWD               As String                              '' ODBC Password
    strConString         As String                              '' OLEDB Connection String
    intConfigSet         As Integer                             '' Need Automatic Set Config
    intOpenInNewWindow   As Integer                             '' ���´��ڴ������
    strVoiceEditorPath   As String                              '' �����༭��ӳ���·��
    intShowSysNodes      As Integer                             '' �Ƿ���ʾϵͳ�ڵ�
    'Michael add @ 06-28-07 For Record File Type
    intRecFileType       As Integer                             ''¼���ļ�����
    intVoiFileType       As Integer                             ''�����ļ�����
    
    ' Sun added 2007-10-20
    intPageHeight        As Long                                '' ��ӡֽ�Ÿ߶�
    intPageWidth         As Long                                '' ��ӡֽ�ſ��
    ' Michael Added @ 2007-11-29 For TTS Setting
    intTTSVoice          As Integer
    strTTSVoice          As String                              ''TTS��������
    intTTSRate           As Integer                             ''TTS����
    intTTSVolume         As Byte                                ''TTS����
    strTTSFormat         As String                              ''TTS����
    ' Michael Added @ 2007-12-10 for default page size
    intDefPageHeight     As Long                                ''Ĭ�ϴ�ӡֽ�Ÿ߶�
    intDefPageWidth      As Long                                ''Ĭ�ϴ�ӡֽ�ſ��
    
End Type
Public gSystem           As typSystemEnvironment                '' ϵͳ����

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''MRU Menu Sun added 2012/05/07
Public Const PCS_MAX_MRUITEMS = 5                               '' ���ư�
Private Type PCS_MRU_ITEM
    blnVisible          As Boolean                              '' �Ƿ�ɼ�
    strItemName         As String                               '' �˵�����
    intItemID           As Integer                              '' ����ID
End Type
Public gMRUMenu(PCS_MAX_MRUITEMS) As PCS_MRU_ITEM

'+++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
''���̽ṹ�� modify:Scott Data:2001/08/23
Private Type GETUSERFLOW
       IntP_id As Byte                                          ''���̱��
       StrP_name As String                                      ''��������
       StrP_description As String * 50                          ''��������
       StrP_version As String                                   ''���̰汾��
       StrP_auther As String                                    ''��������
       StrP_user As String                                      ''�����û�
       DateP_createtime As Date                                 ''���̴���ʱ��
       DataP_modifytime As Date                                 ''�����޸�ʱ��
End Type
Public gGetUserFlow As GETUSERFLOW

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++��Դ��ŷ�������                                            ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type typRIDClassify
    LBound               As Long                                '' ����
    UBound               As Long                                '' ����
    Caption              As String                              '' ����
End Type
Public gRID(4) As typRIDClassify                                '' ��Դ��ŷ���

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++��ѯ����                                                    ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type typQueryCondition
    intScale             As Integer                              '' ��Χ��� 0:���� 1:��ŷ�Χ 2:ָ�����
    intClass             As Integer                              '' ���������
    lngStartID           As Long                                 '' ��ʼ���
    lngEndID             As Long                                 '' ��ֹ���
    lngThisID            As Long                                 '' ָ�����
End Type
Public gCondition        As typQueryCondition                   '' ��ѯ����

'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
'++ TCP/IP ��Ϣ                                                ++
'++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Private Type SMsgHeader
    Sender               As Byte                                '' ������
    Receiver             As Byte                                '' ������
    PackageNo            As Byte                                '' ����
    PackageType          As Byte                                '' ������
    PackageLen           As Integer                             '' ����
End Type

''ͨ����Ϣ���ṹ����
Public Type SCtiMsi_Package
    Msgheader            As SMsgHeader                           '' ��Ϣ��ͷ
    command              As Integer                             '' ������
    sChannel             As Integer                             '' ����ͨ���ţ�����ͨ����
    bytData              As Byte                                '' ��
    intData              As Integer                             '' ����
    strdata              As String * 50                         '' ����
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
'��������
Public Type GETUSERVAR
       IntP_id As Byte                                          ''���̱��
       IntN_index As Integer                                    ''�ڵ�Index
       IntN_id As Integer                                       ''�ڵ�ID
       IntN_no As Byte                                          ''�ڵ���
       IntN_page As Integer                                     ''�ڵ�����ҳ
       IntN_left As Integer                                     ''��
       IntN_top As Integer                                      ''��
       IntN_height As Integer                                   ''��
       IntN_width As Integer                                    ''��
       StrN_description As String                               ''����
       StrN_data1 As String                                     ''���ݶ�1
       StrN_data2 As String                                     ''���ݶ�2
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

'' �ϴ���Դ·��
Public gStrLastResPath As String

'Michael add
Public gbCallFromPro As Byte
Public gstrSQL As String
Public gbSearchFlag As Byte
