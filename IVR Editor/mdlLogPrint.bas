Attribute VB_Name = "mdlLogPrint"
'Michael Added This Module @ 2008-7-1
'Function : Write Log to a text file
'Usage    : call WriteLogMsg() by on error statement when there might be an exception
'
'Last Modified : 2007-7-4
'*************************************************************************
'/////////////////////////////////////////////////////////////////////////
'Err.Number [LogLevel] Date Time ThreadID : Err.Description
'Sample :
'Error-1105[D] 2008-05-28 10:18:21.453 (0x0B08):ADDRESSLIST UPDATE FINISH
'
'LogLevel - N:None, E:Error, W:Warning, I:Information, D:Debug, M:Max
'///////////////////////////////////////////////////////////////////////////
Option Explicit

Declare Function GetThreadLocale Lib "KERNEL32" () As Long
Declare Function GetCurrentThreadId Lib "KERNEL32" () As Long 'get Current thread id
Private gfso            As Object 'File System Object used for write file, check file or folder exist
Private g_logFolderPath As String
Private g_logFilePath   As String
Private Const ForReading = 1, ForWriting = 2, ForAppending = 8

Public Enum LogLevel
    enu_None = 0
    enu_Error = 1
    enu_Warnning = 2
    enu_Information = 3
    enu_Debug = 4
    enu_Max = 5
End Enum

Sub LogFile()
    g_logFolderPath = App.path & "\Log"   'Log folder
    
    If Not (CheckLogFolder(g_logFolderPath)) Then
        CreateLogFolder g_logFolderPath
    End If
End Sub

' check Log Folder exist or not
Function CheckLogFolder(ByVal fdir As String) As Boolean
    Set gfso = CreateObject("Scripting.FileSystemObject")
    CheckLogFolder = gfso.FolderExists(fdir)
End Function

' if Log Folder do not exist then Create it
Sub CreateLogFolder(mkdir As String)
On Error GoTo ErrorHandler

    Dim obj_dir As Object
    Set obj_dir = gfso.CreateFolder(mkdir)
   
Exit Sub
    
ErrorHandler:
    If Err.Number > 0 Then MsgBox Err.Description & ":Log Module Error", vbCritical + vbOKOnly, App.Title
    On Error Resume Next
End Sub

' check the log file exist or not
Function CheckLogFile(ByVal ffile As String) As Boolean
    CheckLogFile = gfso.FileExists(ffile)
End Function

Sub setLogfile()

    g_logFilePath = g_logFolderPath & "\" & Format(Date, "yyyymmdd") & ".log"
    CreateLogFile g_logFilePath
    
End Sub

Sub CreateLogFile(mklogfile)
On Error GoTo ErrorHandler

    Dim lf_log As Object

    If (CheckLogFile(mklogfile)) Then
        Set lf_log = gfso.OpenTextFile(mklogfile, ForAppending)
    Else
        Set lf_log = gfso.CreateTextFile(mklogfile, False)
    End If
    
    lf_log.Close
    
Exit Sub
    
ErrorHandler:
    If Err.Number > 0 Then MsgBox "Write Log Fail:" & Err.Description, vbCritical + vbOKOnly, App.Title
    On Error Resume Next
End Sub

'write the log to the file
Public Sub WriteLogMsg(errDesp As String, _
                       errNo As String, _
                       errLocation As String, _
                       Optional debugInfo As String)
On Error GoTo ErrorHandler

    Dim obj_log As Object
    Dim str_ErrorInfo As String
    
    str_ErrorInfo = Format(Time, "h:mm:ss") & " IN " & errLocation
        
    If (CheckLogFile(g_logFilePath)) Then
        Set obj_log = gfso.OpenTextFile(g_logFilePath, ForAppending)
    Else
        Set obj_log = gfso.CreateTextFile(g_logFilePath, False)
    End If
    
    obj_log.writeline "CurrentThreadId " & StandHex(GetCurrentThreadId)
    obj_log.writeline "Time : " & PreciseTime
    obj_log.writeline (str_ErrorInfo)
    obj_log.writeline ("ErrNumber : " & errNo)
    obj_log.writeline ("ErrDesprition : " & errDesp)
    If Len(debugInfo) > 0 Then obj_log.writeline ("Debug Info : " & debugInfo)
    obj_log.writeline  'write a blank line
    obj_log.Close
    
Exit Sub
    
ErrorHandler:
    If Err.Number > 0 Then MsgBox "Write Log Fail:" & Err.Description, vbCritical + vbOKOnly, App.Title
    On Error Resume Next
End Sub

'delete the blank log file when program shutdown
Public Sub DelBlankLog()
On Error GoTo 0
    
    Dim obj_LogContect As Object
    If CheckLogFile(g_logFilePath) Then
        Set obj_LogContect = gfso.OpenTextFile(g_logFilePath, ForReading)
        If obj_LogContect.AtEndOfLine Then
            obj_LogContect.Close
            Call gfso.DeleteFile(g_logFilePath, False)
        End If
    End If
    
    Set gfso = Nothing
End Sub

'Get precise time - HH:MM:SS.ms
Public Function PreciseTime() As String
    Dim Tss As Single, HM As Integer, SS As Integer, MM As Integer, HH As Integer
    Tss = Timer() * 1000
    HM = Tss Mod 1000               'get millisecond
    Tss = Tss \ 1000                'Total Seconds
    HH = Tss \ 3600                 'get hour
    Tss = Tss Mod 3600              'Total Minutes
    MM = Tss \ 60                   'Get Minute
    SS = Tss Mod 60                 'Get Second
    PreciseTime = Format(HH, "00") & ":" & Format(MM, "00") & ":" & Format(SS, "00") & "." & Format(HM, "000")
End Function

'Format a decimal number to a hex number string like "0xABCD"
Public Function StandHex(iDex As Integer) As String
    Dim strTemp As String, iLoop As Integer
On Error Resume Next
    strTemp = Hex(iDex)
    
    If Len(strTemp) < 4 And Len(strTemp) >= 0 Then
        Do
            strTemp = 0 & strTemp
        Loop Until (Len(strTemp) = 4)
    End If
    
    StandHex = "0x" & strTemp
End Function

'write the log to the file
'Error[1105][D] 2008-05-28 10:18:21.453 (0x0B08):ADDRESSLIST UPDATE FINISH
Public Sub WriteLogMessage(errNo As String, _
                           enuLL As LogLevel, _
                           Description As String, _
                           Optional ExtendInfo As String = ".")
On Error GoTo ErrorHandler

    Dim obj_log As Object
    
    'add the extend information to Description
    'If ExtendInfo <> "" Then Description = Description & " ," & ExtendInfo
    If ExtendInfo <> "." Then
        Description = Description & "," & ExtendInfo
    Else
        Description = Description & ExtendInfo
    End If
    
    If (CheckLogFile(g_logFilePath)) Then
        Set obj_log = gfso.OpenTextFile(g_logFilePath, ForAppending)
    Else
        Set obj_log = gfso.CreateTextFile(g_logFilePath, False)
    End If
    
    obj_log.writeline ParseLogType(enuLL) & " " & _
                      Format(Date, "yyyy-mm-dd") & " " & _
                      PreciseTime & " " & _
                      "(" & StandHex(GetCurrentThreadId) & "):" & _
                      "ErrorCode[" & errNo & "]_" & _
                      Description
    obj_log.Close
Exit Sub
    
ErrorHandler:
    If Err.Number > 0 Then MsgBox "Write Log Fail:" & Err.Description, vbCritical + vbOKOnly, App.Title
    On Error Resume Next
End Sub

Private Function ParseLogType(enuLL As LogLevel) As String
    Select Case enuLL
        Case 0:
            ParseLogType = "[N]"
        Case 1:
            ParseLogType = "[E]"
        Case 2:
            ParseLogType = "[W]"
        Case 3:
            ParseLogType = "[I]"
        Case 4:
            ParseLogType = "[D]"
        Case 5:
            ParseLogType = "[M]"
        Case Else
            ParseLogType = "[*]"
    End Select
End Function

