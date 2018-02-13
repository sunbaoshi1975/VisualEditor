Attribute VB_Name = "mdlcommon"
Option Explicit
'***************************************************************
'   Common1.bas
'   共同过程、函数
'***************************************************************

'***************************************************************
'   API声明
'***************************************************************
Declare Function GetPrivateProfileInt Lib "kernel32" Alias "GetPrivateProfileIntA" (ByVal lpApplicationName As String, ByVal lpKeyName As String, ByVal nDefault As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function GetPrivateProfileSection Lib "kernel32" Alias "GetPrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
Declare Function WritePrivateProfileSection Lib "kernel32" Alias "WritePrivateProfileSectionA" (ByVal lpAppName As String, ByVal lpString As String, ByVal lpFileName As String) As Long
Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal X As Long, ByVal Y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

'In W95 & NT, @f sleep is better than @f Doevents
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

'Communication parts for send ascii 127-255 thar can't properly be sent by MSCOMM control
Declare Function WriteFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToWrite As Long, lpNumberOfBytesWritten As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, lpOverlapped As OVERLAPPED) As Long
Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Public Type OVERLAPPED
        Internal As Long
        InternalHigh As Long
        offset As Long
        OffsetHigh As Long
        hEvent As Long
End Type

'API SetWindowPos
Public Const WP_HWNDTOPMOST = -1
Public Const WP_HWNDNOTOPMOST = -2
Public Const WP_NOMOVE = &H2
Public Const WP_NOSIZE = &H1
Public Const WP_NOACTIVATE = &H10
Public Const WP_SHOWWINDOW = &H40

'Sub SeparateString()
Global Const SEPARATE_DATABASE = vbKeyReturn
Global Const SEPARATE_INIFILE = 44
Global Const SEPARATE_FORMAT = 47
Global Const SEPARATE_PUTATMARK = 64
'Function GetNowDate()
Global Const NOW_YEAR = 0
Global Const NOW_MONTH = 1
Global Const NOW_DAY = 2
Global Const NOW_YOUBI = 3
Global Const NOW_YYYYMMDD = 4
Global Const NOW_HHMMSS = 5
'File Operation
Global Const FO_DIR_SEPMARK = "\"

'Add "\"  to a string
'
Public Function AddDirSepMark(ByVal strTargetItem As String) As String

    If Not Right(strTargetItem, 1) = FO_DIR_SEPMARK Then
        strTargetItem = strTargetItem & FO_DIR_SEPMARK
    End If
    
    AddDirSepMark = strTargetItem

End Function

'".05" -> "0.05"
'
Public Function BeforePutZero(ByVal varValue As Variant) As String

    Dim strValue As String
    
    strValue = CStr(varValue)
    
    If Left(strValue, 1) = "." Then strValue = "0" & strValue
    
    BeforePutZero = strValue

End Function

Public Function EraseInvalidCharacters(ByVal varValue As Variant) As String
On Error Resume Next

    Dim strValue As String
    Dim lv_nMid As Integer
    
    strValue = Trim(CStr(varValue))
    
    lv_nMid = InStr(varValue, Chr(0))
    If lv_nMid > 0 Then
        strValue = Left(strValue, lv_nMid - 1)
    End If
    EraseInvalidCharacters = strValue
    
On Error GoTo 0
End Function

Public Sub ControlCenterPosition(ctrlFront As Control, ctrlBack As Form)
    ctrlFront.Top = (ctrlBack.Height - ctrlFront.Height) / 2
    ctrlFront.Left = (ctrlBack.Width - ctrlFront.Width) / 2
End Sub

'Change mouse pointer
'
Public Sub ChangeMousePointer(lngPointer As Long, okFlag As Boolean)

    If okFlag Then Screen.MousePointer = lngPointer
    DoEvents

End Sub

'If specificated file exists, then return True
'
Public Function CheckExistFile(ByVal strFilePath As String) As Boolean

    Dim intFileNum As Integer
    Dim strPathName As String

    On Error Resume Next

    If Right$(strFilePath, 1) = FO_DIR_SEPMARK Then
        strPathName = Left$(strFilePath, Len(strFilePath) - 1)
    Else
        strPathName = strFilePath
    End If

    intFileNum = FreeFile
    Open strPathName For Input As intFileNum
    CheckExistFile = IIf(Err, False, True)
    Close intFileNum

    Err = 0

End Function

'按指定格式返回字符串型当前日期
'
Public Function GetNowDateString(intDateType As Integer) As String

    Dim strDateFormat As String
    
    Select Case intDateType
    Case NOW_YEAR
        strDateFormat = "yyyy"
        
    Case NOW_MONTH
        strDateFormat = "mm"
        
    Case NOW_DAY
        strDateFormat = "dd"
    
    Case NOW_YOUBI
        strDateFormat = "aaa"
    
    Case NOW_YYYYMMDD
        strDateFormat = "yyyymmdd"
    
    Case NOW_HHMMSS
        strDateFormat = "hhmmss"
    
    Case Else
        GetNowDateString = ""
        Exit Function
        
    End Select
    
    GetNowDateString = Format(Now, strDateFormat)

End Function

Public Function GetOptionButtonIndex(optButton As Object) As Integer

    On Error GoTo GetOptionButtonIndex_Err
    
    Dim intIndex As Integer

    intIndex = 0
    Do
        If optButton(intIndex).value Then
            GetOptionButtonIndex = intIndex
            Exit Do
        End If
        intIndex = intIndex + 1
    Loop

Exit Function
GetOptionButtonIndex_Err:
    GetOptionButtonIndex = -1

End Function

Public Function IsApplicationRunning() As Boolean

    IsApplicationRunning = True

    If App.PrevInstance Then
        Call MsgBox(LoadNationalResString(1479), vbCritical)
        IsApplicationRunning = False
    End If

End Function

Public Sub ResizeCenterControl(baseForm As Form, ResizeCtrl As Control)

    ResizeCtrl.Width = baseForm.Width - (ResizeCtrl.Left) * 3

End Sub

'按指定格式返回整数型当前日期
'
Public Function GetNowDate(dateType As Integer) As Integer

    Dim fmtStr As String

    Select Case dateType
    Case NOW_YEAR
        fmtStr = "yyyy"
        
    Case NOW_MONTH
        fmtStr = "m"
        'fmtStr = "mm"
        
    Case NOW_DAY
        fmtStr = "d"
        'fmtStr = "dd"
    
    Case NOW_YYYYMMDD
        fmtStr = "yyyymmdd"
    
    Case NOW_YOUBI
        fmtStr = "w"
    
    Case Else
        GetNowDate = -1
        Exit Function
        
    End Select
    
    GetNowDate = Val(Format(Now, fmtStr))

End Function

'从指定列表控件中查找特定项，并返回其序号；
'如果没有找到，则返回缺省序号
'
Public Function SearchListIndex(objCtrl As Control, objItem As String, defaultIndex As Integer) As Integer

    Dim ilp As Integer

    For ilp = 0 To objCtrl.ListCount - 1
        If objCtrl.List(ilp) = objItem Then
            SearchListIndex = ilp
            Exit Function
        End If
    Next ilp
    
    SearchListIndex = defaultIndex

End Function

'从指定列表控件中查找特定数据项，并返回其序号；
'如果没有找到，则返回缺省序号
'
Public Function SearchItemDataIndex(objCtrl As Control, ByVal objItem As Long, ByVal defaultIndex As Integer) As Integer

    Dim ilp As Integer

    For ilp = 0 To objCtrl.ListCount - 1
        If objCtrl.ItemData(ilp) = objItem Then
            SearchItemDataIndex = ilp
            Exit Function
        End If
    Next ilp
    
    SearchItemDataIndex = defaultIndex

End Function

'从指定单选控件(optionbox)数组中，按 TAG 属性查找特定项，并返回其序号；
'如果没有找到，则返回缺省序号
'
Public Function SearchOptionIndex(objCtrl As Variant, defaultIndex As Integer) As Integer
Dim Lv_Control As Control

    For Each Lv_Control In objCtrl
        If Lv_Control.value Then
            SearchOptionIndex = Lv_Control.Index
            Exit Function
        End If
    Next
    
    SearchOptionIndex = defaultIndex

End Function

'从指定单选控件(optionbox)数组中，按 TAG 属性查找特定项，并返回其序号；
'如果没有找到，则返回缺省序号
'
Public Function SearchOptionTagIndex(objCtrl As Variant, objItem As String, defaultIndex As Integer) As Integer
Dim Lv_Control As Control

    For Each Lv_Control In objCtrl
        If Trim(Lv_Control.Tag) = Trim(objItem) Then
            SearchOptionTagIndex = Lv_Control.Index
            Exit Function
        End If
    Next
    
    SearchOptionTagIndex = defaultIndex

End Function

'Add specific letter before a string upto fix length
'
'e.g.
'   strData = "256"
'   intDigit = 5
'   strPutChr = "0"
'   return: "00256"
Public Function BeforePutChr(strdata As String, intDigit As Integer, strPutChr As String) As String

    Dim ii%
    Dim str1$

    str1 = ""

    If Len(strdata) >= intDigit Then
        BeforePutChr = strdata
        Exit Function
    End If

    For ii = 1 To intDigit - Len(strdata)
        str1 = str1 + strPutChr
    Next ii

    BeforePutChr = str1 + strdata

End Function

'Center a form
Public Sub FormCenterPosition(appointForm As Form)
    appointForm.Top = (Screen.Height - appointForm.Height) / 2
    appointForm.Left = (Screen.Width - appointForm.Width) / 2
End Sub

'Maxize a form
Public Sub FormStrechWholeScreen(appointForm As Form)
    appointForm.Move 0, 0, Screen.Width, Screen.Height
End Sub

'IniFile 读操作
'
Public Function GetIniFileString(SecName As String, EntName As String, FilePath As String) As String

    Const LEN_GETDATA = 256
    Dim rtnStr As String
    
    rtnStr = Space(LEN_GETDATA)
    
    If GetPrivateProfileString(SecName, EntName, "", rtnStr, Len(rtnStr), FilePath) <> 0 Then
        GetIniFileString = Left(rtnStr, InStr(rtnStr, Chr(0)) - 1)
    Else
        GetIniFileString = ""
    End If

End Function

'IniFile 写操作
'
Public Function WriteIniFileString(SecName As String, EntName As String, WriteData As String, FilePath As String) As Boolean

    WriteIniFileString = True

    If WritePrivateProfileString(SecName, EntName, WriteData, FilePath) = 0 Then
        'Debug.Print Err.Description
        WriteIniFileString = False
        Exit Function
    End If

End Function

'对指定字符串进行加密处理，返回密文
'
Public Function strEncription(ByVal f_str As String) As String
Dim lv_Lpv As Byte, lv_Len As Byte
Dim lv_Source_Char As Integer, lv_Target_Char As Integer
Dim lv_Str As String

'明文长度
lv_Len = Len(f_str)

'遍历明文
For lv_Lpv = 1 To lv_Len
''取一个字符
    lv_Source_Char = Asc(Mid(f_str, lv_Lpv, 1))
''进行变换
    lv_Target_Char = (lv_Source_Char + lv_Lpv * 37 + lv_Len * 7) Mod 128
''反向连接放入密文中
    lv_Str = Chr(lv_Target_Char) + lv_Str
Next

'返回密文
strEncription = lv_Str

End Function

'对指定字符串进行解密处理，返回明文
'
Public Function strDiscription(ByVal f_str As String) As String
Dim lv_Lpv As Byte, lv_Len As Byte
Dim lv_Source_Char As Integer, lv_Target_Char As Integer
Dim lv_Str As String

'密文长度
lv_Len = Len(f_str)

'遍历密文
For lv_Lpv = 1 To lv_Len
''反向取一个字符
    lv_Source_Char = Asc(Mid(f_str, lv_Len - lv_Lpv + 1, 1))
''进行变换
    lv_Target_Char = (128 * lv_Len + lv_Source_Char - lv_Lpv * 37 - lv_Len * 7) Mod 128
''连接放入明文中
    lv_Str = lv_Str + Chr(lv_Target_Char)
Next

'返回明文
strDiscription = lv_Str

End Function

'用键盘模拟IME切换:ON
'
Public Sub EnableIMEStatus()
    If IMEStatus <> vbIMEModeOn Then
        SendKeys "^ "
    End If
End Sub

'用键盘模拟IME切换:OFF
'
Public Sub DisableIMEStatus()
    If IMEStatus <> vbIMEModeOff Then
        SendKeys "^ "
    End If
End Sub

'删除指定文件
'
Public Function RemoveFile(ByVal f_FilePath As String) As Boolean
    If Dir(f_FilePath) <> "" Then
        On Error Resume Next
        Kill f_FilePath
        If Err.Number = 0 Then RemoveFile = True
        Err.Clear
        On Error GoTo 0
    End If
End Function

'将 "YYYYMMDD" 格式字符串转换成日期型
'
Public Function ConvertDate(ByVal f_str As String) As Date
    
    f_str = Left(f_str & "00000000", 8)
    ConvertDate = DateSerial(Left(f_str, 4), Mid(f_str, 5, 2), Right(f_str, 2))
    
End Function

'返回指定 "YYYYMMDD" 格式字符串的当前周日日期
'
Public Function GetThisSunday(ByVal f_str As String) As Date
    Dim lv_date As Date
    Dim lv_Week
    
    '' 转换指定日期
    lv_date = ConvertDate(f_str)
    
    '' 返回指定日期的星期号
    lv_Week = Val(Format(lv_date, "w"))
    
    '' 推算周日
    lv_date = lv_date - lv_Week + 1
    
    '' 返回值
    GetThisSunday = lv_date
    
End Function

'是否是闰年
'
Public Function IsLeapYear(ByVal f_Year As Integer) As Boolean

    If f_Year / 100 = Int(f_Year / 100) Then
        If f_Year / 400 = Int(f_Year / 400) Then
            IsLeapYear = True
        Else
            IsLeapYear = False
        End If
    Else
        If f_Year / 4 = Int(f_Year / 4) Then
            IsLeapYear = True
        Else
            IsLeapYear = False
        End If
    End If
    
End Function

'返回某月的天数
'
Public Function GetDaysInMonth(ByVal f_Year As Integer, ByVal f_Month As Integer) As Integer

    If f_Month = 2 Then
        If IsLeapYear(f_Year) Then
            GetDaysInMonth = 29
        Else
            GetDaysInMonth = 28
        End If
    Else
        Select Case f_Month
        Case 1, 3, 5, 7, 8, 10, 12
            GetDaysInMonth = 31
        Case Else
            GetDaysInMonth = 30
        End Select
    End If
    
End Function

'移动窗口到指定位置
'
Public Function SetWindowPosVB(ByVal f_hwnd As Long, ByVal f_hWndInsertAfter As Long, ByVal f_X As Long, ByVal f_Y As Long, ByVal f_Cx As Long, ByVal f_Cy As Long, ByVal f_wFlags As Long) As Long
    SetWindowPosVB = SetWindowPos(f_hwnd, f_hWndInsertAfter, f_X, f_Y, f_Cx, f_Cy, f_wFlags)
End Function

'创建新目录
'
Public Function CreateNewDir(DirName As String) As Long
    
    '---- 变量 ----
    Dim lRet As Long
    
    '---- 变量初始化 ----
    lRet = 0
    
    '---- 没有该目录则新建 ----
    If IsNull(Dir(DirName, vbDirectory)) = True Or Dir(DirName, vbDirectory) = "" Then
        On Error GoTo CreateNewDir_Err
        MkDir DirName
        On Error GoTo 0
    End If
    
    '---- 終了处理 ----
    CreateNewDir = lRet
    Exit Function
    
CreateNewDir_Err:
    lRet = Err
    On Error Resume Next
    
End Function

'文件拷贝
'
Public Function FileCP(FromName As String, ToName As String) As Long

''返回值
Dim ERR_Cnt As Long

ERR_Cnt = 0

On Error GoTo FileCP_Err
    FileCopy FromName, ToName
On Error GoTo 0

FileCP = ERR_Cnt

Exit Function

FileCP_Err:
    ERR_Cnt = Err
    Resume Next
    
End Function

'后台处理其他进程
'
Public Sub SleepVB(ByVal f_dwMilliseconds As Long)
    Call Sleep(f_dwMilliseconds)
End Sub

'返回指定位的0，1值
'76543210
'
Public Function Get_Bit_Value(ByVal f_Byte As Byte, f_Bit As Byte) As Byte
Dim lv_Lpv As Byte

For lv_Lpv = 0 To f_Bit
    Get_Bit_Value = f_Byte Mod 2
    f_Byte = Int(f_Byte / 2)
Next lv_Lpv

End Function

'设定指定位的0，1值
'76543210
'
Public Sub Set_Bit_Value(f_Byte As Byte, f_Bit As Byte, f_Value As Byte)
Dim lv_ByteH As Byte, lv_ByteL As Byte
Dim lv_Base

lv_Base = 2 ^ f_Bit
lv_ByteH = Int(f_Byte / lv_Base / 2)
lv_ByteL = f_Byte Mod lv_Base
f_Byte = lv_ByteH * lv_Base * 2 + f_Value * lv_Base + lv_ByteL

End Sub

'十进制到二进制
'
Public Function Bin(ByVal f_Byte As Byte) As String
Dim lv_Timer$
Dim lv_lvp As Byte
    
lv_Timer = ""
For lv_lvp = 0 To 7
    lv_Timer = lv_Timer & IIf(Get_Bit_Value(f_Byte, 7 - lv_lvp) = 1, "1", "0")
Next

Bin = lv_Timer

End Function


' 交换两个变量的值
'
Public Sub VarSwap(f_Source As Variant, f_Target As Variant)
On Error Resume Next

    Dim lv_Temp
    
    lv_Temp = f_Target
    f_Target = f_Source
    f_Source = lv_Temp
    
On Error GoTo 0
End Sub
