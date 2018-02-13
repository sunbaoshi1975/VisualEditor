VERSION 5.00
Begin VB.Form frmSyncCallFlow 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "流程同步"
   ClientHeight    =   3750
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4140
   Icon            =   "frmSyncCallFlow.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3750
   ScaleWidth      =   4140
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1621"
   Begin VB.Frame famOpt 
      Caption         =   "目标数据库"
      Height          =   3105
      Left            =   60
      TabIndex        =   7
      Tag             =   "1623"
      Top             =   60
      Width           =   4005
      Begin VB.OptionButton optDataSource 
         Caption         =   "使用 ODBC"
         Height          =   195
         Index           =   1
         Left            =   150
         TabIndex        =   3
         Tag             =   "1433"
         Top             =   1560
         Width           =   1455
      End
      Begin VB.OptionButton optDataSource 
         Caption         =   "使用 OLE DB"
         Height          =   180
         Index           =   0
         Left            =   150
         TabIndex        =   0
         Tag             =   "1430"
         Top             =   360
         Value           =   -1  'True
         Width           =   1335
      End
      Begin VB.TextBox txtDBName 
         Height          =   270
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   2
         Top             =   1050
         Width           =   2055
      End
      Begin VB.TextBox txtServer 
         Height          =   270
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   1
         Top             =   690
         Width           =   2055
      End
      Begin VB.TextBox txtDSN 
         Height          =   285
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1845
         Width           =   2055
      End
      Begin VB.TextBox txtPassword 
         Height          =   285
         IMEMode         =   3  'DISABLE
         Left            =   1650
         MaxLength       =   50
         PasswordChar    =   "*"
         TabIndex        =   6
         Top             =   2565
         Width           =   2055
      End
      Begin VB.TextBox txtUserID 
         Height          =   285
         Left            =   1650
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2205
         Width           =   2055
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "DB Name"
         Height          =   180
         Index           =   5
         Left            =   570
         TabIndex        =   14
         Tag             =   "1432"
         Top             =   1095
         Width           =   630
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ServerIP"
         Height          =   195
         Index           =   4
         Left            =   570
         TabIndex        =   13
         Tag             =   "1431"
         Top             =   735
         Width           =   615
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "ODBC DSN"
         Height          =   195
         Index           =   3
         Left            =   570
         TabIndex        =   12
         Tag             =   "1434"
         Top             =   1890
         Width           =   840
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "口  令"
         Height          =   180
         Index           =   2
         Left            =   570
         TabIndex        =   11
         Tag             =   "1436"
         Top             =   2610
         Width           =   540
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "用户名"
         Height          =   180
         Index           =   1
         Left            =   570
         TabIndex        =   10
         Tag             =   "1435"
         Top             =   2250
         Width           =   540
      End
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   315
      Left            =   2190
      TabIndex        =   9
      Tag             =   "1144"
      Top             =   3330
      Width           =   1065
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&C)"
      Default         =   -1  'True
      Height          =   315
      Left            =   750
      TabIndex        =   8
      Tag             =   "1372"
      Top             =   3330
      Width           =   1065
   End
End
Attribute VB_Name = "frmSyncCallFlow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()

    Dim lv_msgResult As VbMsgBoxResult
    Dim lv_strConnStr As String
    Dim lv_strMessage As String
    Dim msgresult As Integer
    Dim lv_blnSucc As Boolean
    
    If gCallFlow.CallFlowID > 0 Then
    
        '' 流程是否先保存
        If Not gCallFlow.SavedMark Then
            lv_msgResult = MsgBox(LoadNationalResString(1102) & gCallFlow.CallFlowName & "(" & Str(gCallFlow.CallFlowID) & ") ?", vbYesNo + vbApplicationModal + vbQuestion)
            If lv_msgResult = vbYes Then
                gCallFlow.UpdateIvrTable
            End If
        End If
    
        '' 同步数据
        If optDataSource(0).value Then
            lv_strConnStr = "Provider=SQLOLEDB;Integrated Security=SSPI;Data Source=" & txtServer & ";Initial Catalog=" & txtDBName
        Else
            lv_strConnStr = "DSN=" & txtDSN & ";UID=" & txtUserID & ";PWD=" & txtPassword
        End If
        
        '' Data Source must differ from Data Target
        If StrComp(gSystem.strConString, lv_strConnStr, vbTextCompare) = 0 Then
            Call Message("M012")
            Exit Sub
        End If
        
        '' Confirm
        If gSystem.intConfigSet = 0 Then
            lv_strMessage = LoadNationalResString(1624) & gCallFlow.CallFlowName & "(" & Format(Str(gCallFlow.CallFlowID)) & ")" & _
                            LoadNationalResString(1625) & vbCrLf & "    [" & gSystem.strConString & "]" & vbCrLf & _
                            LoadNationalResString(1626) & vbCrLf & "    [" & lv_strConnStr & "]" & vbCrLf & _
                            LoadNationalResString(1627)
        Else
            lv_strMessage = LoadNationalResString(1714) & gCallFlow.ResourceName & "(" & Format(Str(gCallFlow.ResourceID)) & ")" & _
                            LoadNationalResString(1625) & vbCrLf & "    [" & gSystem.strConString & "]" & vbCrLf & _
                            LoadNationalResString(1626) & vbCrLf & "    [" & lv_strConnStr & "]" & vbCrLf & _
                            LoadNationalResString(1627)
        End If
        msgresult = MsgBox(lv_strMessage, vbYesNo + vbApplicationModal + vbQuestion + vbDefaultButton2)
        If msgresult <> vbYes Then
            Exit Sub
        End If
        
        lv_blnSucc = False
        If gSystem.intConfigSet = 0 Then
            If gCallFlow.SynchronizeCallFlow(lv_strConnStr) Then
                lv_blnSucc = True
                Call Message("M011")
            Else
                Call Message("E063")
            End If
                
        Else
            If gCallFlow.SynchronizeResource(lv_strConnStr) Then
                lv_blnSucc = True
                Call Message("M013")
            Else
                Call Message("E064")
            End If
        End If
    
        '' 成功
        If lv_blnSucc Then
            ''' 保存设置
            If optDataSource(1).value Then
                Call WriteIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_TYPE, Def_Default_TYPE1, gSystem.strINI_File)
            Else
                Call WriteIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_TYPE, Def_Default_TYPE0, gSystem.strINI_File)
            End If
        
            Call WriteIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_DBNAME, txtDBName, gSystem.strINI_File)
            Call WriteIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_DBSERVER, txtServer, gSystem.strINI_File)
            Call WriteIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_DSN, txtDSN, gSystem.strINI_File)
            Call WriteIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_USERID, txtUserID, gSystem.strINI_File)
            Call WriteIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_PWD, txtPassword, gSystem.strINI_File)
                            
            ''' Exit
            cmdExit_Click
        End If
        
    Else
        Call Message("M131")
    End If
    
End Sub

Private Sub Form_Load()
On Error Resume Next

    Dim lv_Str As String
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
   
    '获取数据库访问参数
    '' MS SQL or Other Database
    ''TYPE
    optDataSource(0).value = True
    lv_Str = GetIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_TYPE, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        If Trim(lv_Str) = "ODBC" Then
            optDataSource(0).value = False
        End If
    End If
    optDataSource(1).value = Not optDataSource(0).value
    
    '' MS SQL Server Name
    lv_Str = GetIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_DBSERVER, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        txtServer = lv_Str
    Else
        txtServer = Def_Default_DBServer
    End If
    txtServer.Enabled = optDataSource(0).value
    
    '' MS SQL Database Name
    lv_Str = GetIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_DBNAME, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        txtDBName = lv_Str
    Else
        txtDBName = Def_Default_Database
    End If
    txtDBName.Enabled = optDataSource(0).value
    
    ''ODBC DSN
    lv_Str = GetIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_DSN, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        txtDSN = lv_Str
    Else
        txtDSN = Def_Default_DSN
    End If
    txtDSN.Enabled = optDataSource(1).value
    
    ''ODBC User ID
    lv_Str = GetIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_USERID, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        txtUserID = lv_Str
    Else
        txtUserID = Def_Default_USERID
    End If
    txtUserID.Enabled = optDataSource(1).value
    
    ''ODBC Password
    lv_Str = GetIniFileString(Def_INI_SEC_TAR_ODBC, Def_INI_ENTRY_PWD, gSystem.strINI_File)
    If Trim(lv_Str) <> "" Then
        txtPassword = lv_Str
    Else
        txtPassword = Def_Default_PWD
    End If
    txtPassword.Enabled = optDataSource(1).value
    
    '改变鼠标指针形状->箭头光标
    ChangeMousePointer vbDefault, True
    
    If gSystem.intConfigSet = 0 Then
        Me.Tag = 1621
    Else
        Me.Tag = 1714
    End If
    
    LoadResStrings Me

On Error GoTo 0
End Sub

Private Sub optDataSource_Click(Index As Integer)
     
    If optDataSource(0).value Then
        txtServer.Enabled = True
        txtDBName.Enabled = True
        txtUserID.Enabled = False
        txtPassword.Enabled = False
        txtDSN.Enabled = False
    Else
        txtServer.Enabled = False
        txtDBName.Enabled = False
        txtUserID.Enabled = True
        txtPassword.Enabled = True
        txtDSN.Enabled = True
    End If

End Sub

Private Sub txtDBName_GotFocus()
    txtDBName.SelStart = 0
    txtDBName.SelLength = Len(txtDBName)
End Sub

Private Sub txtDSN_GotFocus()
    txtDSN.SelStart = 0
    txtDSN.SelLength = Len(txtDSN)
End Sub

Private Sub txtPassword_GotFocus()
    txtPassword.SelStart = 0
    txtPassword.SelLength = Len(txtPassword)
End Sub

Private Sub txtServer_GotFocus()
    txtServer.SelStart = 0
    txtServer.SelLength = Len(txtServer)
End Sub

Private Sub txtUserID_GotFocus()
    txtUserID.SelStart = 0
    txtUserID.SelLength = Len(txtUserID)
End Sub

