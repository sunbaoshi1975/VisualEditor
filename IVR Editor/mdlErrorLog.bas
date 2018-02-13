Attribute VB_Name = "mdlErrorLog"
'Michael Added This Module @ 2008-7-1
'Function : Write Error Log to a text file
'Usage    : call WriteLogMsg(err) by on error statement when there might be an exception
' argument err should better format like : "functionname() & err.Description"
'*************************************************************************
Option Explicit

Public gfso         As Object
Public gs_logname   As String
Public Const ForReading = 1, ForWriting = 2, ForAppending = 8

Sub LogFile()
    Dim str_path As String
        
    str_path = App.path & "\ErrorLog"
    
    If Not (CheckLogFolder(str_path)) Then
        CreateLogFolder str_path
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
    If Err.Number > 0 Then MsgBox Err.Description, vbCritical + vbOKOnly, "Log Mod Error"
    On Error Resume Next
End Sub

' check the log file exist or not
Function CheckLogFile(ByVal ffile As String) As Boolean
    CheckLogFile = gfso.FileExists(ffile)
End Function

Sub setLogfile()
    Dim str_logfile As String

    str_logfile = App.path & "\ErrorLog\" & Format(Date, "yyyymmdd") & ".log"
    gs_logname = str_logfile
    CreateLogFile str_logfile
    
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
    If Err.Number > 0 Then MsgBox Err.Description, vbCritical + vbOKOnly, "Log Mod Error"
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
        
    If (CheckLogFile(gs_logname)) Then
        Set obj_log = gfso.OpenTextFile(gs_logname, ForAppending)
    Else
        Set obj_log = gfso.CreateTextFile(gs_logname, False)
    End If

    obj_log.writeline (str_ErrorInfo)
    obj_log.writeline ("ErrNumber : " & errNo)
    obj_log.writeline ("ErrDesprition : " & errDesp)
    If Len(debugInfo) > 0 Then obj_log.writeline ("Debug Info : " & debugInfo)
    obj_log.writeline  'write a blank line
    obj_log.Close
    
Exit Sub
    
ErrorHandler:
    If Err.Number > 0 Then MsgBox Err.Description, vbCritical + vbOKOnly, "Log Mod Error"
    On Error Resume Next
End Sub

'delete the empty log when program shutdown
Public Sub DelBlankLog()
On Error GoTo 0
    
    Dim obj_LogContect As Object
    If CheckLogFile(gs_logname) Then
        Set obj_LogContect = gfso.OpenTextFile(gs_logname, ForReading)
        If obj_LogContect.AtEndOfLine Then
            obj_LogContect.Close
            Call gfso.DeleteFile(gs_logname, False)
        End If
    End If
    
    Set gfso = Nothing
End Sub
