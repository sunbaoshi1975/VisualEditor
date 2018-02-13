Attribute VB_Name = "Module1"
'*** SDI 记事本应用程序示例全局模块               ***
'****************************************************
Option Explicit

' 存储关于子窗体信息的用户定义类型
Type FormState
    Deleted As Integer
    Dirty As Integer
    Color As Long
End Type
   Public dd As Long
   Public GdeleteP_ID As Long
   Public GdeleteN_id As Long
   Public GN_id As Long
   Public GN_no As Long
   Public gstrData1 As String
   Public gstrData2 As String
Public M_Str_P_ModifyTime As String '状态栏中流程修改时间
Public M_Str_P_Version As String '状态栏中流程版本号
Public M_Str_P_User As String '状态栏中流程用户名称

Public M_Cn As ADODB.Connection
Public FState As FormState              ' 用户定义型数组
Public gFindString As String            ' 保存搜索文本
Public gFindCase As Integer             ' 区分大小写标志
Public gFindDirection As Integer        ' 搜索方向标志
Public gCurPos As Integer               ' 保存当前光标位置
Public gFirstTime As Integer            ' 起始位置
Public Const ThisApp = "MDINote"        ' 注册表 App 常量。
Public Const ThisKey = "Recent Files"   ' 注册表 Key 常量。
Sub GetRecentFiles()
    ' 本过程演示 GetAllSettings 函数的用法，它从 Windows 注册表中返回值的数组。
    ' 在这种情况下，注册表包含最近打开的文件列表。使用 SaveSetting 语句记下最近使用的文件名。
   ' 该语句在 WriteRecentFiles 过程中使用
    Dim i As Integer
    Dim varFiles As Variant ' 存储返回的数组的变量
    
    ' 用 GetAllSettings 语句从注册表中返回最近使用的文件。
    '  ThisApp 和 ThisKey是模块中定义的常数
   
    If GetSetting(ThisApp, ThisKey, "RecentFile1") = Empty Then Exit Sub
    
    varFiles = GetAllSettings(ThisApp, ThisKey)
    
   ' For i = 0 To UBound(varFiles, 1)
   '     frmSDI.mnuRecentFile(0).Visible = True
   '     frmSDI.mnuRecentFile(i + 1).Caption = varFiles(i, 1)
    '    frmSDI.mnuRecentFile(i + 1).Visible = True
   ' Next i
End Sub
Sub ResizeNote()
    ' 文本框充满窗体的内部区域
    If frmSDI.picToolbar.Visible Then
        frmSDI.tvwDB.Height = frmSDI.ScaleHeight - frmSDI.picToolbar.Height - 300
        frmSDI.tvwDB.Width = frmSDI.ScaleWidth - 7000
        frmSDI.tvwDB.Top = frmSDI.picToolbar.Height
        frmSDI.DataGrid1.Height = frmSDI.ScaleHeight - frmSDI.picToolbar.Height - 300
        frmSDI.DataGrid1.Left = frmSDI.ScaleWidth - 7000 + 150
        frmSDI.DataGrid1.Top = frmSDI.picToolbar.Height

    Else
        frmSDI.tvwDB.Height = frmSDI.ScaleHeight - 300
        frmSDI.tvwDB.Width = frmSDI.ScaleWidth - 7000
        frmSDI.tvwDB.Top = 0
        frmSDI.DataGrid1.Height = frmSDI.ScaleHeight - 300
        frmSDI.DataGrid1.Top = 0

    End If
End Sub

Sub WriteRecentFiles(OpenFileName)
    ' 本过程使用 SaveSettings 语句将最近打开的文件名写入系统注册表。
    ' SaveSettings 语句要求三个参数其中两个存储为常数并在本模块内定义。
    ' GetRecentFiles 过程中使用 GetAllSettings 函数来检索这个过程中存储的文件名。
    Dim i As Integer
    Dim strFile As String
    Dim strKey As String

    ' 将文件 RecentFile1 复制给 RecentFile2，等等
    For i = 3 To 1 Step -1
        strKey = "RecentFile" & i
        strFile = GetSetting(ThisApp, ThisKey, strKey)
        If strFile <> "" Then
            strKey = "RecentFile" & (i + 1)
            SaveSetting ThisApp, ThisKey, strKey, strFile
        End If
    Next i
  
    ' 将正在打开的文件写到最近使用文件列表的第一项
    SaveSetting ThisApp, ThisKey, "RecentFile1", OpenFileName
End Sub

