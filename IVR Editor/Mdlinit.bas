Attribute VB_Name = "Mdlinit"
Option Explicit

Public Sub SysInit()
Dim lv_Str As String
Dim lv_Rtn
On Error GoTo errFind

'检查应用程序是否已经执行，避免重入
If App.PrevInstance = True Then
    Call Message("M001")
    End
End If

'设置系统各工作路径
gSystem.strPath_Working = App.path & "\"
gSystem.strINI_File = gSystem.strPath_Working & Def_INI_Config
gSystem.strVersionInfo = "1.0"
gSystem.intConfigSet = 0
Set gClipBoard = New clsIVRClipBoard
    
' initialize system variables
gNodeHandColor = vbWhite
gblnShowNodeCaption = False

    '获取数据库访问参数
    '' MS SQL or Other Database
    ''TYPE
    lv_Str = GetIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_TYPE, gSystem.strINI_File)
    ''Michael Add : make connect type is ODBC or not
    If Trim(lv_Str) <> "" Then
        If Trim(lv_Str) = "ODBC" Then
            gSystem.blnUserODBC = True
        Else
            gSystem.blnUserODBC = False
        End If
    ''Michael Add : Make the default connect type ODBC
    Else
        gSystem.blnUserODBC = True
        Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_TYPE, Def_Default_TYPE1, gSystem.strINI_File)
    End If
    
    '' MS SQL Server Name
    '' Michael Add : IP address
    lv_Str = GetIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_DBSERVER, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        gSystem.strDBServer = lv_Str
    Else
        gSystem.strDBServer = Def_Default_DBServer
        Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_DBSERVER, Def_Default_DBServer, gSystem.strINI_File)
    End If

    '' MS SQL Database Name
    lv_Str = GetIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_DBNAME, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        gSystem.strDBName = lv_Str
    Else
        gSystem.strDBName = Def_Default_Database
        Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_DBNAME, Def_Default_Database, gSystem.strINI_File)
    End If

    ''ODBC DSN
    lv_Str = GetIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_DSN, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        gSystem.strDSN = lv_Str
    Else
        gSystem.strDSN = Def_Default_DSN
        Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_DSN, Def_Default_DSN, gSystem.strINI_File)
    End If

    ''ODBC User ID
    lv_Str = GetIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_USERID, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        gSystem.strUserID = lv_Str
    Else
        gSystem.strUserID = Def_Default_USERID
        Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_USERID, Def_Default_USERID, gSystem.strINI_File)
    End If

    ''ODBC Password
    lv_Str = GetIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_PWD, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        gSystem.strPWD = lv_Str
    Else
    ''Michael Add ; the default password is NULL password
        gSystem.strPWD = Def_Default_PWD
        Call WriteIniFileString(Def_INI_SEC_ODBC, Def_INI_ENTRY_PWD, Def_Default_PWD, gSystem.strINI_File)
    End If
    
    If Not gSystem.blnUserODBC Then
        gSystem.strConString = "Provider=SQLOLEDB;Integrated Security=SSPI;Data Source=" & gSystem.strDBServer & ";Initial Catalog=" & gSystem.strDBName
    Else
        gSystem.strConString = "DSN=" & gSystem.strDSN & ";UID=" & gSystem.strUserID & ";PWD=" & gSystem.strPWD & ";Initial Catalog=" & gSystem.strDBName
    End If

'获取系统工作参数
' 系统语音路径
On Error Resume Next

' Michael Modified @ Aug,7,07
lv_Str = GetIniFileString(Def_INI_SEC_SYS, Def_INI_ENTRY_VoxPath, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    ''Michael Add : add a '\' to the end of the sting
    gSystem.strPath_SysVox = mdlcommon.AddDirSepMark(lv_Str)
    'If the default dir is not exist then create it
    
    If Dir(gSystem.strPath_SysVox, vbDirectory) = "" Then
    
        '' Sun replaced 2008-02-18
        ''' From
        ''MkDir (gSystem.strPath_SysVox)
        ''' To
        If CreateNewDir(gSystem.strPath_SysVox) > 0 Then
            Call Message("E144")
        End If

    End If
Else
    gSystem.strPath_SysVox = "C:\VOX Path"
    Call WriteIniFileString(Def_INI_SEC_SYS, Def_INI_ENTRY_VoxPath, gSystem.strPath_SysVox, gSystem.strINI_File)
    gSystem.strPath_SysVox = mdlcommon.AddDirSepMark(gSystem.strPath_SysVox)
    
    '' Sun replaced 2008-02-18
    ''' From
    ''MkDir (gSystem.strPath_SysVox)
    ''' To
    If CreateNewDir(gSystem.strPath_SysVox) > 0 Then
        Call Message("E144")
    End If
    
End If
gStrLastResPath = gSystem.strPath_SysVox
On Error GoTo errFind

'' 获取TTS参数
'' Michael Added @ 2007-11-29
'' TTS Voice
lv_Str = GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVOICE, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.strTTSVoice = Trim(lv_Str)
Else
    gSystem.strTTSVoice = Def_Default_TTSVoice
    Call WriteIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVOICE, Def_Default_TTSVoice, gSystem.strINI_File)
End If

'' TTS Rate
lv_Str = GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSRATE, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.intTTSRate = CInt(Trim(lv_Str))
Else
    gSystem.intTTSRate = Def_Default_TTSRate
    Call WriteIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSRATE, Def_Default_TTSRate, gSystem.strINI_File)
End If

'' TTS Volume
lv_Str = GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVolume, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.intTTSVolume = CByte(Val(lv_Str) Mod 256)
Else
    gSystem.intTTSVolume = Def_Default_TTSVolume
    Call WriteIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSVolume, Def_Default_TTSVolume, gSystem.strINI_File)
End If

'' TTS Format
lv_Str = GetIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSFormat, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.strTTSFormat = Trim(lv_Str)
Else
    gSystem.strTTSFormat = Def_Default_TTSFormat
    Call WriteIniFileString(Def_INI_SEC_TTS, Def_INI_ENTRY_TTSFormat, Def_Default_TTSFormat, gSystem.strINI_File)
End If


'获取网络工作模式
''服务器IP
lv_Str = GetIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_ServerIP, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.strServerIP = lv_Str
Else
    gSystem.strServerIP = Def_Default_ServerIP
    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_ServerIP, Def_Default_ServerIP, gSystem.strINI_File)
End If

''服务器Port  --- default:10020
lv_Str = GetIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_ServerPort, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.intServerPort = Val(lv_Str)
Else
    gSystem.intServerPort = Def_Default_ServerPort
    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_ServerPort, Def_Default_ServerPort, gSystem.strINI_File)
End If

''本地Port   --- defaule : 0
lv_Str = GetIniFileString(Def_INI_SEC_Local, Def_INI_ENTRY_LocalPort, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.intLocalPort = Val(lv_Str)
Else
    gSystem.intLocalPort = Def_Default_LocalPort
    Call WriteIniFileString(Def_INI_SEC_Local, Def_INI_ENTRY_LocalPort, Def_Default_LocalPort, gSystem.strINI_File)
End If
 
'' IDE Environmental Option
''' Frame Backgroup Color
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_FRAME_BG, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gFrameBackColor = Val(lv_Str)
Else
    
    '' Sun added 2008-02-18
    gFrameBackColor = &H808080
    
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_FRAME_BG, Str(gFrameBackColor), gSystem.strINI_File)
End If

''' Work Page Height
''' Michael Added @ 2007-12-10
gSystem.intDefPageHeight = 1150 * Screen.TwipsPerPixelY
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_HE, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gPageHeight = Val(lv_Str)
Else
    
    '' Sun added 2008-02-18
    gPageHeight = gSystem.intDefPageHeight
    
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_HE, Str(gPageHeight), gSystem.strINI_File)
End If

''' Work Page Width
''' Michael Added @ 2007-12-10
gSystem.intDefPageWidth = 800 * Screen.TwipsPerPixelX
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_WD, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gPageWidth = Val(lv_Str)
Else

    '' Sun added 2008-02-18
    gPageWidth = gSystem.intDefPageWidth
    
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_WD, Str(gPageWidth), gSystem.strINI_File)
End If

''' Work Page Backgroup Color
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_BG, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gPageBackColor = Val(lv_Str)
Else
    
    '' Sun added 2008-02-18
    gPageBackColor = &HFFFFFF
    
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_PAGE_BG, Str(gPageBackColor), gSystem.strINI_File)
End If

''' 显示节点说明
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_SHOW_NODE_CAPTION, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gblnShowNodeCaption = (lv_Str = "1")
Else
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_SHOW_NODE_CAPTION, "0", gSystem.strINI_File)
End If

''' 显示节点标签   added by Mike @08-1-29
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_SHOW_NODE_TAG, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gblnShowNodeTag = (lv_Str = "1")
Else
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_SHOW_NODE_TAG, "0", gSystem.strINI_File)
End If

''' 粘贴板最大堆栈数
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_MAXCLIPSTACKS, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gClipBoard.MaxStacks = CByte(Val(lv_Str))
Else
    gClipBoard.MaxStacks = Def_Default_MaxClipBoardStacks
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_MAXCLIPSTACKS, Str(gClipBoard.MaxStacks), gSystem.strINI_File)
End If

'' Sun added 2007-03-25
'' 在新窗口创建或打开流程
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_OpenInNewWindow, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    If Val(lv_Str) = 1 Then
        gSystem.intOpenInNewWindow = vbChecked
    Else
        gSystem.intOpenInNewWindow = vbUnchecked
    End If
Else
    gSystem.intOpenInNewWindow = Def_Default_OpenInNewWindow
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_OpenInNewWindow, Str(gSystem.intOpenInNewWindow), gSystem.strINI_File)
End If

'' Sun added 2007-03-25
'' 语音编辑外接程序路径
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_VoiceEditorPath, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.strVoiceEditorPath = Trim(lv_Str)
Else
    gSystem.strVoiceEditorPath = ""
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_VoiceEditorPath, gSystem.strVoiceEditorPath, gSystem.strINI_File)
End If

'' Sun added 2007-03-25
'' 是否显示系统节点
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_ShowSysNodes, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.intShowSysNodes = Val(lv_Str)
Else
    gSystem.intShowSysNodes = False
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_ShowSysNodes, Str(gSystem.intShowSysNodes), gSystem.strINI_File)
End If

''重复当前节点
lv_Str = GetIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Key_repeat, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    Default.key_repeat = Asc(lv_Str)
    Node0_Data2.key_repeat = Asc(lv_Str)
Else
    Default.key_repeat = Def_Default_Key_repeat
    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Key_repeat, Def_Default_Key_repeat, gSystem.strINI_File)
End If

''回到上一级菜单
lv_Str = GetIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Key_return, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    Default.key_return = Asc(lv_Str)
    Node0_Data2.key_return = Asc(lv_Str)
Else
    Default.key_return = Def_Default_Key_return
    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Key_return, Def_Default_Key_return, gSystem.strINI_File)
End If
''回到主菜单
lv_Str = GetIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Key_root, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    Default.key_root = Asc(lv_Str)
    Node0_Data2.key_root = Asc(lv_Str)
Else
    Default.key_root = Def_Default_Key_root
    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Key_root, Def_Default_Key_root, gSystem.strINI_File)
End If
''父节点
lv_Str = GetIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Nd_parent, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    Default.nd_parent = lv_Str
    Node0_Data2.nd_parent = lv_Str
Else
    Default.nd_parent = Def_Default_Nd_parent
    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Nd_parent, Def_Default_Nd_parent, gSystem.strINI_File)
End If

''根节点  Default : 256
lv_Str = GetIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Nd_root, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    Default.nd_root = lv_Str
    Node0_Data2.nd_root = lv_Str
Else
    Default.nd_root = Def_Default_Nd_root
    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Nd_root, Def_Default_Nd_root, gSystem.strINI_File)
End If

''用户定义变量数 Default : 0
lv_Str = GetIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Uservar, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    Defaultuservar.uservars = lv_Str
    Node1_Data1.uservars = lv_Str
Else
    Defaultuservar.uservars = Def_Default_Uservar
    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Uservar, Def_Default_Uservar, gSystem.strINI_File)
End If

'? Michael -- System var ? Default : 0
lv_Str = GetIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Var, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    Defaulttype.uservar(0) = lv_Str
    Node2_Data2.uservar(0) = lv_Str
Else
    Defaulttype.uservar(0) = Def_Default_Var
    Call WriteIniFileString(Def_INI_SEC_Remote, Def_INI_ENTRY_Var, Def_Default_Var, gSystem.strINI_File)
End If

''Michael add @06-28-07 For Record File Types *******************************
lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_RecordFileType, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.intRecFileType = Val(lv_Str)
Else
    gSystem.intRecFileType = 0
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_RecordFileType, Str(gSystem.intRecFileType), gSystem.strINI_File)
End If

lv_Str = GetIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_VoiceFileType, gSystem.strINI_File)
If Trim(lv_Str) <> "" Then
    gSystem.intVoiFileType = Val(lv_Str)
Else
    gSystem.intVoiFileType = 0
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_VoiceFileType, Str(gSystem.intVoiFileType), gSystem.strINI_File)
End If
'**********************   Add End ****************************************

'工作环境设定
ChDrive Left(gSystem.strPath_Working, 3)
ChDir gSystem.strPath_Working                       '' 进入工作目录

' Sun added 2002-12-04
g_NodeListShowType = 0
Call LoadNodeTypeNameList     'From mdlNode

Call InitResoureIDClassify    '资源ID分类

errFind:
    
    If Err = -2147467259 Then
        WriteLogMessage Err.Number, enu_Warnning, LoadResString(1976)
        Resume Next
        
        Err.Clear
        Message "E028"
    
    ElseIf Err <> 0 Then ' 其他的错误
        MsgBox LoadNationalResString(1484) & Err.Description
        'Mike added @ 2008-7-4
        WriteLogMessage Err.Number, enu_Error, "System init fail", Err.Description
        End
    End If
End Sub

'资源ID分类
'
Private Sub InitResoureIDClassify()
    
    gRID(0).LBound = 0
    gRID(0).UBound = 1000
    gRID(0).Caption = LoadNationalResString(1485)
    
    gRID(1).LBound = 1000
    gRID(1).UBound = 20000
    gRID(1).Caption = LoadNationalResString(1486)
    
    gRID(2).LBound = 20000
    gRID(2).UBound = 25000
    gRID(2).Caption = LoadNationalResString(1487)
    
    gRID(3).LBound = 25000
    gRID(3).UBound = 26000
    gRID(3).Caption = LoadNationalResString(1488)
    
    gRID(4).LBound = 26000
    gRID(4).UBound = 27000
    gRID(4).Caption = LoadNationalResString(1489)
    
End Sub

'' Sun added 2012-05-07
''' Read MRU Data
Public Sub ReadMRUMData()

Dim lv_sCaption As String
Dim lv_sData As String
Dim lv_Loop As Integer

For lv_Loop = 1 To PCS_MAX_MRUITEMS
    lv_sCaption = "MRUItemID" & Trim(Str(lv_Loop))
    lv_sData = Trim(GetIniFileString(Def_INI_SEC_MRUList, lv_sCaption, gSystem.strINI_File))
    If Val(lv_sData) > 0 Then
        gMRUMenu(lv_Loop).intItemID = Val(lv_sData)
        gMRUMenu(lv_Loop).blnVisible = True
        
        lv_sCaption = "MRUCaption" & Trim(Str(lv_Loop))
        gMRUMenu(lv_Loop).strItemName = Trim(GetIniFileString(Def_INI_SEC_MRUList, lv_sCaption, gSystem.strINI_File))
    Else
        gMRUMenu(lv_Loop).blnVisible = False
        gMRUMenu(lv_Loop).strItemName = ""
        gMRUMenu(lv_Loop).intItemID = 0
    End If
Next

End Sub

'' Sun added 2012-05-07
''' Save MRU Data
Public Sub SaveMRUMData()

Dim lv_sCaption As String
Dim lv_sData As String
Dim lv_Loop As Integer

For lv_Loop = 1 To PCS_MAX_MRUITEMS
    lv_sCaption = "MRUItemID" & Trim(Str(lv_Loop))
    If gMRUMenu(lv_Loop).blnVisible Then
        lv_sData = Trim(Str(gMRUMenu(lv_Loop).intItemID))
    Else
        lv_sData = ""
    End If
    Call WriteIniFileString(Def_INI_SEC_MRUList, lv_sCaption, lv_sData, gSystem.strINI_File)
    lv_sCaption = "MRUCaption" & Trim(Str(lv_Loop))
    Call WriteIniFileString(Def_INI_SEC_MRUList, lv_sCaption, Trim(gMRUMenu(lv_Loop).strItemName), gSystem.strINI_File)
Next

End Sub

'' Sun added 2012-05-07
''' Push MRU Item
Public Sub PushMRUItem(bytItemID As Byte, strName As String)

Dim lv_sCaption As String
Dim lv_nPos As Integer
Dim lv_Loop As Integer

'' 查找重复ID 或者 未使用的位置
For lv_Loop = 1 To PCS_MAX_MRUITEMS

    If gMRUMenu(lv_Loop).intItemID = bytItemID Or gMRUMenu(lv_Loop).blnVisible = False Then
        Exit For
    End If
Next

If lv_Loop > PCS_MAX_MRUITEMS Then
    lv_nPos = PCS_MAX_MRUITEMS
Else
    lv_nPos = lv_Loop
End If

For lv_Loop = lv_nPos To 2 Step -1
    gMRUMenu(lv_Loop).blnVisible = gMRUMenu(lv_Loop - 1).blnVisible
    gMRUMenu(lv_Loop).intItemID = gMRUMenu(lv_Loop - 1).intItemID
    gMRUMenu(lv_Loop).strItemName = gMRUMenu(lv_Loop - 1).strItemName
Next

gMRUMenu(1).blnVisible = True
gMRUMenu(1).intItemID = bytItemID
gMRUMenu(1).strItemName = strName

End Sub
