VERSION 5.00
Begin VB.Form Frm_FlowCreate 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "�½�����"
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
      Caption         =   "���´����д�"
      Height          =   375
      Left            =   90
      TabIndex        =   8
      Tag             =   "1664"
      Top             =   3540
      Width           =   3225
   End
   Begin VB.CommandButton CommandCreate 
      Caption         =   "ȷ��(&C)"
      Default         =   -1  'True
      Height          =   315
      Left            =   1500
      TabIndex        =   9
      Tag             =   "1372"
      Top             =   4140
      Width           =   1065
   End
   Begin VB.CommandButton exit_Command 
      Caption         =   "�˳�(&E)"
      Height          =   315
      Left            =   2880
      TabIndex        =   10
      Tag             =   "1144"
      Top             =   4140
      Width           =   1065
   End
   Begin VB.Frame Frame1 
      Caption         =   "��������"
      BeginProperty Font 
         Name            =   "����"
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
         Caption         =   "����ID:"
         Height          =   195
         Left            =   150
         TabIndex        =   19
         Tag             =   "1363"
         Top             =   420
         Width           =   1410
      End
      Begin VB.Label p_name_Label 
         AutoSize        =   -1  'True
         Caption         =   "��������:"
         Height          =   195
         Left            =   150
         TabIndex        =   18
         Tag             =   "1364"
         Top             =   780
         Width           =   1410
      End
      Begin VB.Label p_desc_Label 
         AutoSize        =   -1  'True
         Caption         =   "��������:"
         Height          =   195
         Left            =   150
         TabIndex        =   17
         Tag             =   "1365"
         Top             =   1140
         Width           =   1410
      End
      Begin VB.Label p_ver_Label 
         AutoSize        =   -1  'True
         Caption         =   "�汾��:"
         Height          =   195
         Left            =   150
         TabIndex        =   16
         Tag             =   "1366"
         Top             =   1500
         Width           =   1410
      End
      Begin VB.Label p_auth_Label 
         AutoSize        =   -1  'True
         Caption         =   "��������:"
         Height          =   195
         Left            =   150
         TabIndex        =   15
         Tag             =   "1367"
         Top             =   1860
         Width           =   1410
      End
      Begin VB.Label p_user_Label 
         AutoSize        =   -1  'True
         Caption         =   "�����û�:"
         Height          =   195
         Left            =   150
         TabIndex        =   14
         Tag             =   "1368"
         Top             =   2220
         Width           =   1410
      End
      Begin VB.Label p_crea_Label 
         AutoSize        =   -1  'True
         Caption         =   "����ʱ��:"
         Height          =   195
         Left            =   150
         TabIndex        =   13
         Tag             =   "1369"
         Top             =   2550
         Width           =   1410
      End
      Begin VB.Label p_modi_Label 
         AutoSize        =   -1  'True
         Caption         =   "�޸�ʱ��:"
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
'//�ļ�����  Frm_FlowCreate.frm
'//��;��    �����µ�����
'//����:     Scott
'//�������ڣ�2001/08/29
'//�޸����ڣ�2005/08/20
'//�ļ�������Create New Flows, So it is saved into tbRefIVR table
'//////////////////////////////////////////////////////////////////
Option Explicit

Dim lv_RS_FS As ADODB.Recordset
Dim m_blnDataChanged As Boolean

Private Sub Cmb_pid_Click()
On Error Resume Next

    Dim lv_FlowNo As Integer
    
    '�ж������Ƿ�Ϊ��
    If Cmb_pid.Text <> "" And Not IsNull(Cmb_pid) Then
    
        lv_FlowNo = CByte(Val(Trim(Cmb_pid.Text)))
        lv_RS_FS.MoveFirst
        lv_RS_FS.Find "p_id=" & lv_FlowNo
        
        If lv_RS_FS.EOF Then '��������Ϣ
           
            Txt_pname = "N/A"          '��������
            Txt_pdescription = "N/A"   '��������
            Txt_pauther = "N/A"        '��������
            Txt_puser = "N/A"          '�����û�
            Txt_pversion = "V1.0"      '�汾��
            Txt_pmodifytime = Format(Now(), "yyyy-mm-dd hh:nn:ss")   '����ʱ��
            Txt_pcreatetime = Format(Now(), "yyyy-mm-dd hh:nn:ss")   '�޸�ʱ��
           
        Else
           
            '����������
            Txt_pname = Trim(lv_RS_FS!P_Name)                   '��������
            Txt_pdescription = Trim(lv_RS_FS!P_Description)     '��������
            Txt_pauther = Trim(lv_RS_FS!P_Auther)               '��������
            Txt_puser = Trim(lv_RS_FS!P_User)                   '�����û�
            Txt_pversion = Trim(lv_RS_FS!P_Version)             '�汾��
            Txt_pmodifytime = Format(lv_RS_FS!P_ModifyTime, "yyyy-mm-dd hh:nn:ss")             '�޸�ʱ��
            Txt_pcreatetime = Format(lv_RS_FS!P_CreateTime, "yyyy-mm-dd hh:nn:ss")             '����ʱ��
           
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

    '�ж������Ƿ�Ϊ��
    If Trim(Cmb_pid.Text) = "" Or IsNull(Cmb_pid) Then
        'MsgBox "��ʾ����ѡ�����̻�����������!"
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

    '����������
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
    
    ' ��Ҫ�ж������Ƿ��Ѿ���
    If frmMain.IsCallFlowOpened(lv_nFlowID) Then
        Unload Me
        Message "M142"
        frmMain.SwitchToMDIForm lv_nFlowID
        Exit Sub
    End If
    
    '���´����д�
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
    
        '����
        gCallFlow.DestroyAllNodes
        gCallFlow.ClearWorkPage
    End If

    ' ��������
    Call gCallFlow.OpenIvrRecordSet(lv_nFlowID)

    ' �ж������Ƿ������ݿ������
    If gCallFlow.IsNewCallFlow Then
        
        '' ����������
        Call gCallFlow.AddCallFlowIntoList(lv_nFlowID, Trim(Txt_pname), Trim(Txt_pdescription), Trim(Txt_pauther), Trim(Txt_puser), Trim(Txt_pversion))
        
        '' Create Default System Nodes
        If gCallFlow.NewNodeID <= 0 Then
            gCallFlow.CreateSysDefaultNodes
        End If
    Else

        '' ����������Ϣ
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
    '��tbrefivr���ж�ȡ���̺�,����䵽����ID�����б���
    LoadProjectTable
    LoadResStrings Me
    
    '' ��ǰ�����Ƿ��ǿ�����
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

    '�ı����ָ����״->ɳ©���
    mdlcommon.ChangeMousePointer vbHourglass, True

    Set lv_RS_FS = New ADODB.Recordset
   
    With lv_RS_FS
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open "Select * from tbrefivr where P_Type = 'P' Order By P_ID", gSystem.strConString
    End With
 
    '��ȡ���̱�ŵ�Cmb_pid��
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
    
    '�ı����ָ����״->��ͷ���
    ChangeMousePointer vbDefault, True

    Exit Sub
    
BackDoor:
    
    '�ı����ָ����״->��ͷ���
    ChangeMousePointer vbDefault, True
    
    Debug.Print Err.Description
    
    Err.Clear
    Message "E016"
    gSystem.intConfigSet = 1        '' ODBC
        
End Sub

'��������
Private Sub Txt_pauther_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_pauther_GotFocus()
    Txt_pauther.SelStart = 0
    Txt_pauther.SelLength = Len(Txt_pauther)
End Sub

'��������
Private Sub Txt_pdescription_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_pdescription_GotFocus()
    Txt_pdescription.SelStart = 0
    Txt_pdescription.SelLength = Len(Txt_pdescription)
End Sub

'��������
Private Sub Txt_pname_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_pname_GotFocus()
    Txt_pname.SelStart = 0
    Txt_pname.SelLength = Len(Txt_pname)
End Sub

'�����û�
Private Sub Txt_puser_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_puser_GotFocus()
    Txt_puser.SelStart = 0
    Txt_puser.SelLength = Len(Txt_puser)
End Sub

'�汾��
Private Sub Txt_pversion_Change()
    m_blnDataChanged = True
End Sub

Private Sub Txt_pversion_GotFocus()
    Txt_pversion.SelStart = 0
    Txt_pversion.SelLength = Len(Txt_pversion)
End Sub
