VERSION 5.00
Begin VB.Form Frm_FlowCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "新建流程"
   ClientHeight    =   4560
   ClientLeft      =   4350
   ClientTop       =   2460
   ClientWidth     =   4485
   Icon            =   "Frm_FlowCreate.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4560
   ScaleWidth      =   4485
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1361"
   Begin VB.CheckBox chkOpenNew 
      Caption         =   "在新窗口中打开"
      Height          =   375
      Left            =   90
      TabIndex        =   8
      Tag             =   "1664"
      Top             =   3540
      Width           =   3225
   End
   Begin VB.CommandButton CommandCreate 
      Caption         =   "确定(&C)"
      Default         =   -1  'True
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Tag             =   "1372"
      Top             =   4140
      Width           =   1065
   End
   Begin VB.CommandButton exit_Command 
      Caption         =   "退出(&E)"
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Tag             =   "1144"
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "流程属性"
      BeginProperty Font 
         Name            =   "宋体"
         Size            =   9.75
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   3405
      Left            =   45
      TabIndex        =   11
      Tag             =   "1362"
      Top             =   60
      Width           =   4395
      Begin VB.ComboBox Cmb_pid 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1710
         TabIndex        =   0
         Top             =   300
         Width           =   2505
      End
      Begin VB.TextBox Txt_pname 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   1
         Top             =   690
         Width           =   2505
      End
      Begin VB.TextBox Txt_pdescription 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1710
         MaxLength       =   200
         TabIndex        =   2
         Top             =   1050
         Width           =   2505
      End
      Begin VB.TextBox Txt_pversion 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1710
         MaxLength       =   10
         TabIndex        =   3
         Top             =   1410
         Width           =   2505
      End
      Begin VB.TextBox Txt_pauther 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   4
         Top             =   1770
         Width           =   2505
      End
      Begin VB.TextBox Txt_puser 
         BeginProperty Font 
            Name            =   "Fixedsys"
            Size            =   12
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   360
         Left            =   1710
         MaxLength       =   50
         TabIndex        =   5
         Top             =   2130
         Width           =   2505
      End
      Begin VB.TextBox Txt_pcreatetime 
         Height          =   345
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   6
         Top             =   2490
         Width           =   2505
      End
      Begin VB.TextBox Txt_pmodifytime 
         Height          =   345
         Left            =   1710
         Locked          =   -1  'True
         TabIndex        =   7
         Top             =   2850
         Width           =   2505
      End
      Begin VB.Label p_id_lable 
         AutoSize        =   -1  'True
         Caption         =   "流程ID:"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Tag             =   "1363"
         Top             =   420
         Width           =   1410
      End
      Begin VB.Label p_name_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程名称:"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Tag             =   "1364"
         Top             =   780
         Width           =   1410
      End
      Begin VB.Label p_desc_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程描述:"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Tag             =   "1365"
         Top             =   1140
         Width           =   1410
      End
      Begin VB.Label p_ver_Label 
         AutoSize        =   -1  'True
         Caption         =   "版本号:"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Tag             =   "1366"
         Top             =   1500
         Width           =   1410
      End
      Begin VB.Label p_auth_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程作者:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Tag             =   "1367"
         Top             =   1860
         Width           =   1410
      End
      Begin VB.Label p_user_Label 
         AutoSize        =   -1  'True
         Caption         =   "流程用户:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Tag             =   "1368"
         Top             =   2220
         Width           =   1410
      End
      Begin VB.Label p_crea_Label 
         AutoSize        =   -1  'True
         Caption         =   "创建时间:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Tag             =   "1369"
         Top             =   2550
         Width           =   1410
      End
      Begin VB.Label p_modi_Label 
         AutoSize        =   -1  'True
         Caption         =   "修改时间:"
         Height          =   195
         Left            =   150
         TabIndex        =   12
         Tag             =   "1370"
         Top             =   2940
         Width           =   1410
      End
   End
End
Attribute VB_Name = "Frm_FlowCreate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'//             Create New Flow and Update old Flow
'//文件名：  Frm_FlowCreate.frm
'//用途：    创建新的流程
'//作者:     Scott
'//创建日期：2001/08/29
'//修改日期：2005/08/20
'//文件描述：Create New Flows, So it is saved into tbRefIVR table
'//////////////////////////////////////////////////////////////////
Option Explicit

Dim lv_RS_FS As ADODB.Recordset
Dim m_blnDataChanged As Boolean

Private Sub Cmb_pid_Click()
On Error Resume Next

    Dim lv_FlowNo As Integer
    
    '判断流程是否为空
    If Cmb_pid.Text <> "" And Not IsNull(Cmb_pid) Then
    
        lv_FlowNo = CByte(Val(Trim(Cmb_pid.Text)))
        lv_RS_FS.MoveFirst
        lv_RS_FS.Find "p_id=" & lv_FlowNo
        
        If lv_RS_FS.EOF Then '新流程信息
           
            Txt_pname = "N/A"          '流程名称
            Txt_pdescription = "N/A"   '流程描述
            Txt_pauther = "N/A"        '流程作者
            Txt_puser = "N/A"          '流程用户
            Txt_pversion = "V1.0"      '版本号
            Txt_pmodifytime = Format(Now(), "yyyy-mm-dd hh:nn:ss")   '创建时间
            Txt_pcreatetime = Format(Now(), "yyyy-mm-dd hh:nn:ss")   '修改时间
           
        Else
           
            '打开已有流程
            Txt_pname = Trim(lv_RS_FS!P_Name)                   '流程名称
            Txt_pdescription = Trim(lv_RS_FS!P_Description)     '流程描述
            Txt_pauther = Trim(lv_RS_FS!P_Auther)               '流程作者
            Txt_puser = Trim(lv_RS_FS!P_User)                   '流程用户
            Txt_pversion = Trim(lv_RS_FS!P_Version)             '版本号
            Txt_pmodifytime = Format(lv_RS_FS!P_ModifyTime, "yyyy-mm-dd hh:nn:ss")             '修改时间
            Txt_pcreatetime = Format(lv_RS_FS!P_CreateTime, "yyyy-mm-dd hh:nn:ss")             '创建时间
           
        End If
   
        Frm_FlowCreate.Caption = LoadNationalResString(1373) & Txt_pname & "(" & Trim(Str(lv_FlowNo)) & ")"

    End If
    
    m_blnDataChanged = False
    
End Sub

Private Sub Cmb_pid_GotFocus()
    Cmb_pid.SelStart = 0
    Cmb_pid.SelLength = Len(Cmb_pid)

End Sub

Private Sub Cmb_pid_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
End Sub

Private Sub Cmb_pid_LostFocus()
On Error Resume Next

    Dim lv_FlowNo As Integer

    '判断流程是否为空
    If Trim(Cmb_pid.Text) = "" Or IsNull(Cmb_pid) Then
        'MsgBox "提示：请选择流程或输入新流程!"
    Else
        
        lv_FlowNo = CByte(Val(Trim(Cmb_pid.Text)))
        lv_RS_FS.MoveFirst
        lv_RS_FS.Find "p_id=" & lv_FlowNo
        
        If lv_RS_FS.EOF Then
        
            Frm_FlowCreate.Caption = LoadNationalResString(1373) & Trim(Str(lv_FlowNo))
            
            Txt_pname = "N/A"
            Txt_pdescription = "N/A"
            Txt_pauther = gSystem.strOSUser
            Txt_puser = "N/A"
            Txt_pversion = "V1.0"
            Txt_pmodifytime = Format(Now(), "yyyy-mm-dd hh:nn:ss")
            Txt_pcreatetime = Format(Now(), "yyyy-mm-dd hh:nn:ss")
        
        Else
        
            
        End If
     
     End If
     
End Sub

Public Sub CommandCreate_Click()
On Error Resume Next

    '创建新流程
    Dim lv_msgResult As VbMsgBoxResult
    Dim lv_nFlowID As Byte
    Dim lv_loop As Integer
    Dim lv_SubLoop As Integer
        
    If Trim(Cmb_pid) = "" Then
        Message "E089"
        Cmb_pid.SetFocus
        Exit Sub
    End If
    lv_nFlowID = CByte(Val(Trim(Cmb_pid.Text)))
    
    ' 需要判断流程是否已经打开
    If frmMain.IsCallFlowOpened(lv_nFlowID) Then
        Unload Me
        Message "M142"
        frmMain.SwitchToMDIForm lv_nFlowID
        Exit Sub
    End If
    
    '在新窗口中打开
    If chkOpenNew.value = vbChecked Then
        '' Open New CallFlow Window
        If Not frmMain.CreateNewMDIForm(lv_nFlowID) Then
            Unload Me
            Exit Sub
        End If
    Else
        If gCallFlow.CallFlowID > 0 Then
    
            frmMain.DeassignMDIForm gCallFlow.CallFlowID
            If Not gCallFlow.SavedMark Then
                lv_msgResult = MsgBox(LoadNationalResString(1102) & gCallFlow.CallFlowName & "(" & Str(gCallFlow.CallFlowID) & ") ?", vbYesNo + vbApplicationModal + vbQuestion)
                If lv_msgResult = vbYes Then
                    gCallFlow.UpdateIvrTable
                End If
            End If
        End If
        frmMain.AssignMDIForm lv_nFlowID
    
        '清屏
        gCallFlow.DestroyAllNodes
        gCallFlow.ClearWorkPage
    End If

    ' 打开新流程
    Call gCallFlow.OpenIvrRecordSet(lv_nFlowID)

    ' 判断流程是否在数据库里存在
    If gCallFlow.IsNewCallFlow Then
        
        '' 增加新流程
        Call gCallFlow.AddCallFlowIntoList(lv_nFlowID, Trim(Txt_pname), Trim(Txt_pdescription), Trim(Txt_pauther), Trim(Txt_puser), Trim(Txt_pversion))
        
        '' Create Default System Nodes
        If gCallFlow.NewNodeID <= 0 Then
            gCallFlow.CreateSysDefaultNodes
        End If
    Else

        '' 更新流程信息
        If m_blnDataChanged Then
            Call gCallFlow.UpdateCallFlowInfo(Trim(Txt_pname), Trim(Txt_pdescription), Trim(Txt_pauther), Trim(Txt_puser), Trim(Txt_pversion))
        End If
    End If

    frmMain.ShowCallFlowOnScreen
    
    Unload Me
   
End Sub

Private Sub exit_Command_Click()
'If Message("Q005") = vbYes Then
   Unload Me
'Else
'   Exit Sub
'End If
End Sub

Private Sub Form_Load()
    '从tbrefivr表中读取流程号,并填充到流程ID下拉列表中
    LoadProjectTable
    LoadResStrings Me
    
    '' 当前界面是否是空流程
    If gCallFlow.NewNodeID > 0 Then
        chkOpenNew.value = gSystem.intOpenInNewWindow
        chkOpenNew.Enabled = True
    Else
        chkOpenNew.value = vbUnchecked
        chkOpenNew.Enabled = False
    End If
    
    m_blnDataChanged = False
        
End Sub

Private Sub LoadProjectTable()
On Error GoTo BackDoor

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    Set lv_RS_FS = New ADODB.Recordset
   
    With lv_RS_FS
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open "Select * from tbrefivr where P_Type = 'P' Order By P_ID", gSystem.strConString
    End With
 
    '提取流程编号到Cmb_pid中
    Cmb_pid.Clear
    Do While Not lv_RS_FS.EOF
        If Not IsNull(lv_RS_FS!P_ID) Then
            With Cmb_pid
                .AddItem Trim(Str(lv_RS_FS!P_ID)) & " - " & Trim(lv_RS_FS!P_Name)
                .ItemData(.ListCount - 1) = lv_RS_FS!P_ID
            End With
            lv_RS_FS.MoveNext
        End If
    Loop
 
    gSystem.intConfigSet = 0
    
    '改变鼠标指针形状->箭头光标
    ChangeMousePointer vbDefault, True

    Exit Sub
    
BackDoor:
    
    '改变鼠标指针形状->箭头光标
    ChangeMousePointer vbDefault, True
    
    Debug.Print Err.Description
    
    Err.Clear
    Message "E016"
    gSystem.intConfigSet = 1        '' ODBC
        
End Sub

'流程作者
Private Sub Txt_pauther_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_pauther_GotFocus()
    Txt_pauther.SelStart = 0
    Txt_pauther.SelLength = Len(Txt_pauther)
End Sub

'流程描述
Private Sub Txt_pdescription_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_pdescription_GotFocus()
    Txt_pdescription.SelStart = 0
    Txt_pdescription.SelLength = Len(Txt_pdescription)
End Sub

'流程名称
Private Sub Txt_pname_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_pname_GotFocus()
    Txt_pname.SelStart = 0
    Txt_pname.SelLength = Len(Txt_pname)
End Sub

'流程用户
Private Sub Txt_puser_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_puser_GotFocus()
    Txt_puser.SelStart = 0
    Txt_puser.SelLength = Len(Txt_puser)
End Sub

'版本号
Private Sub Txt_pversion_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_pversion_GotFocus()
    Txt_pversion.SelStart = 0
    Txt_pversion.SelLength = Len(Txt_pversion)
End Sub
