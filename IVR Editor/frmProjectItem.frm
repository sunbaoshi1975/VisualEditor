VERSION 5.00
Begin VB.Form frmProjectItem 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "项目编辑"
   ClientHeight    =   3165
   ClientLeft      =   6105
   ClientTop       =   6060
   ClientWidth     =   7485
   Icon            =   "frmProjectItem.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3165
   ScaleWidth      =   7485
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1910"
   Begin VB.TextBox txtPUser 
      Height          =   375
      Left            =   4680
      TabIndex        =   5
      Top             =   1080
      Width           =   2535
   End
   Begin VB.CommandButton cmdQuit 
      Caption         =   "取消(&Q)"
      Height          =   375
      Left            =   6000
      TabIndex        =   10
      Tag             =   "1909"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&O)"
      Height          =   375
      Left            =   4680
      TabIndex        =   9
      Tag             =   "1372"
      Top             =   2640
      Width           =   1215
   End
   Begin VB.TextBox txtPDes 
      Height          =   375
      Left            =   1080
      TabIndex        =   8
      Top             =   2040
      Width           =   6135
   End
   Begin VB.TextBox txtPMTime 
      Enabled         =   0   'False
      Height          =   375
      Left            =   4680
      TabIndex        =   7
      Top             =   1590
      Width           =   2535
   End
   Begin VB.TextBox txtPCTime 
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   6
      Top             =   1590
      Width           =   2535
   End
   Begin VB.TextBox txtPAuthor 
      Height          =   375
      Left            =   1080
      TabIndex        =   4
      Top             =   1080
      Width           =   2535
   End
   Begin VB.TextBox txtPName 
      Height          =   375
      Left            =   1080
      TabIndex        =   3
      Top             =   600
      Width           =   6135
   End
   Begin VB.TextBox txtPVersion 
      Height          =   375
      Left            =   6000
      TabIndex        =   2
      Top             =   150
      Width           =   1215
   End
   Begin VB.TextBox txtPID 
      BackColor       =   &H008080FF&
      Enabled         =   0   'False
      Height          =   375
      Left            =   1080
      TabIndex        =   1
      Top             =   150
      Width           =   1215
   End
   Begin VB.Label lblPDes 
      AutoSize        =   -1  'True
      Caption         =   "项目描述:"
      Height          =   195
      Left            =   120
      TabIndex        =   17
      Tag             =   "1942"
      Top             =   2160
      Width           =   765
   End
   Begin VB.Label lblPMTime 
      AutoSize        =   -1  'True
      Caption         =   "修改时间:"
      Height          =   195
      Left            =   3840
      TabIndex        =   16
      Tag             =   "1120"
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label lblPCTime 
      AutoSize        =   -1  'True
      Caption         =   "创建时间:"
      Height          =   195
      Left            =   120
      TabIndex        =   15
      Tag             =   "1106"
      Top             =   1680
      Width           =   765
   End
   Begin VB.Label lblPUser 
      AutoSize        =   -1  'True
      Caption         =   "项目用户:"
      Height          =   195
      Left            =   3840
      TabIndex        =   14
      Tag             =   "1941"
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label lblPAuthor 
      AutoSize        =   -1  'True
      Caption         =   "项目作者:"
      Height          =   195
      Left            =   120
      TabIndex        =   13
      Tag             =   "1940"
      Top             =   1170
      Width           =   765
   End
   Begin VB.Label lblPName 
      AutoSize        =   -1  'True
      Caption         =   "项目名称"
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Tag             =   "1938"
      Top             =   720
      Width           =   720
   End
   Begin VB.Label lblVersion 
      AutoSize        =   -1  'True
      Caption         =   "项目版本:"
      Height          =   195
      Left            =   4920
      TabIndex        =   11
      Tag             =   "1939"
      Top             =   240
      Width           =   765
   End
   Begin VB.Label lblPID 
      AutoSize        =   -1  'True
      Caption         =   "项目编号:"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Tag             =   "1937"
      Top             =   240
      Width           =   765
   End
End
Attribute VB_Name = "frmProjectItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/File Ver : V1.0
'/Author   : Michael S
'/Date     : July,13,07
'/Name     : frmProjectItem
'//////////////////////////////////////////////////////////////////////

Option Explicit
Dim strSQL As String
Public bCopyOpt As Boolean
Public lv_strOriFront As String, lv_strOriBack As String
Public lv_strNewFront As String, lv_strNewBack As String

Private Sub cmdQuit_Click()
    Unload Me
End Sub


Private Sub cmdOK_Click()
    'New Add Operation
    If txtPID.Enabled Then
        'Project ID must be needed
        If txtPID.Text = "" Then
            Message "E141"
            Exit Sub
        End If
        
        'Project Name must be needed
        If txtPName.Text = "" Then
            Message "E142"
            Exit Sub
        End If
        
        'Project ID must an integer between 1 to 255
        If Val(txtPID.Text) > 255 Or Val(txtPID.Text) < 1 Then
            Message "E143"
            Exit Sub
        End If
        
        strSQL = "Insert into tbrefivr (P_ID,P_Name,P_Version,P_Auther,P_User,P_CreateTime,P_ModifyTime" & _
                    ",P_Description,P_Type) values ('" & txtPID.Text & "','" & txtPName.Text & "','" & txtPVersion & "'," & _
                    "'" & txtPAuthor.Text & "','" & txtPUser.Text & "','" & txtPCTime.Text & "','" & txtPMTime.Text & "', " & _
                    "'" & txtPDes.Text & "','R')"
        'gCallFlow.AddCallFlowIntoList txtPID, txtPName, txtPDes, txtPAuthor, txtPUser, txtPVersion
        
        'Michael Added Oct,29,2K7 - Copy Resources when Project Copied
        If ((frmProjectList.bCopyRes) And frmProjectList.lstProject.SelectedItem.Text <> "") Then
            CopyResourceofPro frmProjectList.lstProject.SelectedItem.Text, txtPID
        End If
        '***** Added End *****
        'Mike added @ 2008-7-7
        Call WriteLogMessage(0, enu_Information, "Add a new Resource Project:" & txtPID)
    
    'Modify Operation
    Else
        'Project Name must be needed
        If IsNull(txtPName.Text) Then
            Message "E142"
            Exit Sub
        End If
    
        strSQL = "Update tbRefIVR SET P_Name='" & txtPName.Text & "',P_Version='" & txtPVersion.Text & "',P_auther='" & txtPAuthor.Text & "'," & _
                    "P_User='" & txtPUser.Text & "',P_createtime='" & txtPCTime.Text & "',p_ModifyTime='" & txtPMTime.Text & "'," & _
                    "P_Description='" & txtPDes.Text & "' where P_ID='" & txtPID.Text & "'"
        'gCallFlow.UpdateCallFlowInfo txtPName, txtPDes, txtPAuthor, txtPUser, txtPVersion
        'Mike added @ 2008-7-7
        Call WriteLogMessage(0, enu_Information, "Update Resource Project:" & txtPID)
        
    End If
    
    FillProListView (strSQL)
    Unload Me
End Sub


'Insert or Update the data of database
Private Sub FillProListView(strTempSQL As String)
On Error GoTo BackDoor

    Dim itmX As ListItem
    Dim intCount As Integer
    
' Add the contect
With gCallFlow.RS_Project
    
    .CursorType = adOpenKeyset
    .LockType = adLockReadOnly
    .CursorLocation = adUseClient
    .Open strTempSQL, gSystem.strConString
    
End With
            
frmProjectList.FillProListView

BackDoor:
    'MsgBox (Err.Description)
    On Error GoTo 0
End Sub

'------------------------------------------------------------------------------------
'Michael Added @ Oct,29,2K7
Private Sub CopyResourceofPro(strSce_RID As String, strDes_RID As String)
On Error GoTo ErrHandle

    'Step 1 -> 复制数据表中数据
    Dim RS_Res As ADODB.Recordset
    Set RS_Res = New ADODB.Recordset
    Dim strCopyRes As String
    
    Dim strSQLRes As String
    strSQLRes = "SELECT R_ID, L_ID, R_Type, R_Path, R_Description,CreateTime, ModifyTime, R_Note, P_ID " & _
                "FROM tbResource WHERE " & _
                "P_ID= '" & strSce_RID & "' Order By P_ID,R_ID,L_ID"
    
    With RS_Res
        If .State = 1 Then
            .Close
        End If
        
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open strSQLRes, gSystem.strConString
        Set RS_Res = Nothing
    
        If .RecordCount > 0 Then
            Dim iLoopCount As Integer
            Dim iLoopInter As Integer
            Dim iLoopOuter As Integer
            iLoopCount = .RecordCount
            Dim arrRes() As String
            ReDim arrRes(1 To iLoopCount, 1 To 8)
            
            .MoveFirst
            For iLoopOuter = 1 To iLoopCount
                For iLoopInter = 1 To 8
                    If iLoopInter = 1 And Not IsNull(.Fields("R_ID")) Then
                        arrRes(iLoopOuter, iLoopInter) = Trim(.Fields("R_ID"))
                    ElseIf iLoopInter = 2 And Not IsNull(.Fields("L_ID")) Then
                        arrRes(iLoopOuter, iLoopInter) = Trim(.Fields("L_ID"))
                    ElseIf iLoopInter = 3 And Not IsNull(.Fields("R_Type")) Then
                        arrRes(iLoopOuter, iLoopInter) = Trim(.Fields("R_Type"))
                    ElseIf iLoopInter = 4 And Not IsNull(.Fields("R_Path")) Then
                        'Michael Modified @ 2007-11-28
                        If bCopyOpt = False Then
                            arrRes(iLoopOuter, iLoopInter) = Trim(.Fields("R_Path"))
                        Else
                            '前缀,后缀替换
                            arrRes(iLoopOuter, iLoopInter) = ReplaceResPath(Trim(.Fields("R_Path")))
                        End If
                    ElseIf iLoopInter = 5 And Not IsNull(.Fields("R_Description")) Then
                        arrRes(iLoopOuter, iLoopInter) = Trim(.Fields("R_Description"))
                    ElseIf iLoopInter = 6 And Not IsNull(.Fields("CreateTime")) Then
                        arrRes(iLoopOuter, iLoopInter) = Trim(.Fields("CreateTime"))
                    ElseIf iLoopInter = 7 And Not IsNull(.Fields("ModifyTime")) Then
                        arrRes(iLoopOuter, iLoopInter) = Trim(.Fields("ModifyTime"))
                    ElseIf iLoopInter = 8 And Not IsNull(.Fields("R_Note")) Then
                        arrRes(iLoopOuter, iLoopInter) = Trim(.Fields("R_Note"))
                    Else
                        arrRes(iLoopOuter, iLoopInter) = ""
                    End If
                Next iLoopInter
                .MoveNext
            Next iLoopOuter
            
            For iLoopOuter = 1 To iLoopCount
                iLoopInter = 1
                strCopyRes = "INSERT INTO tbResource (P_ID,R_ID,L_ID,R_Type,R_Path,R_Description,CreateTime" & _
                             ",ModifyTime,R_Note) VALUES ('" & strDes_RID & "','" & arrRes(iLoopOuter, iLoopInter) & "','" & arrRes(iLoopOuter, iLoopInter + 1) & "'," & _
                             "'" & arrRes(iLoopOuter, iLoopInter + 2) & "','" & arrRes(iLoopOuter, iLoopInter + 3) & "','" & arrRes(iLoopOuter, iLoopInter + 4) & "','" & arrRes(iLoopOuter, iLoopInter + 5) & "', " & _
                             "'" & arrRes(iLoopOuter, iLoopInter + 6) & "','" & arrRes(iLoopOuter, iLoopInter + 7) & "')"
                If .State = 1 Then
                    .Close
                End If
                .Open strCopyRes, gSystem.strConString
            Next iLoopOuter
        End If
    End With
        
ErrHandle:
    If Err.Number > 0 Then MsgBox Err.Description
    Exit Sub
    
End Sub

'项目资源前后缀替换 (Michael Modified @ 2007-11-28)
Private Function ReplaceResPath(strOriPath As String) As String
    ReplaceResPath = strOriPath
    '处理后缀
    If lv_strNewBack <> lv_strOriBack Then
        ReplaceResPath = Left(ReplaceResPath, Len(ReplaceResPath) - 3) & lv_strNewBack
    End If
    
    '处理前缀
    If lv_strNewFront <> lv_strOriFront Then
        If lv_strOriFront = "" Then
            '直接添加前缀
            ReplaceResPath = lv_strNewFront & ReplaceResPath
        ElseIf lv_strNewFront = "" Then
            '直接截取前缀
            ReplaceResPath = Right(ReplaceResPath, Len(ReplaceResPath) - Len(lv_strOriFront))
        Else
            '替換前綴
            ReplaceResPath = lv_strNewFront & Right(ReplaceResPath, Len(ReplaceResPath) - Len(lv_strOriFront))
        End If
    End If

End Function

'Michael Added @ 2007-12-5
Private Sub Form_Load()
    LoadResStrings Me
End Sub
