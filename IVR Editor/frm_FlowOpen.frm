VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frm_FlowOpen 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "������"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   11310
   Icon            =   "frm_FlowOpen.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   6630
   ScaleWidth      =   11310
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1699"
   Begin VB.CommandButton cmdDelete 
      Caption         =   "ɾ��(&D)"
      Height          =   345
      Left            =   4200
      TabIndex        =   2
      Tag             =   "1371"
      Top             =   6210
      Width           =   1065
   End
   Begin VB.CheckBox chkOpenNew 
      Caption         =   "���´����д�"
      Height          =   375
      Left            =   240
      TabIndex        =   1
      Tag             =   "1664"
      Top             =   6150
      Width           =   3225
   End
   Begin VB.CommandButton exit_Command 
      Caption         =   "�˳�(&E)"
      Height          =   345
      Left            =   6960
      TabIndex        =   3
      Tag             =   "1144"
      Top             =   6210
      Width           =   1065
   End
   Begin VB.CommandButton cmdOpen 
      Caption         =   "ȷ��(&C)"
      Default         =   -1  'True
      Height          =   345
      Left            =   5580
      TabIndex        =   0
      Tag             =   "1372"
      Top             =   6210
      Width           =   1095
   End
   Begin FPSpreadADO.fpSpread vasFlows 
      Height          =   6105
      Left            =   -60
      TabIndex        =   4
      Top             =   -30
      Width           =   11355
      _Version        =   458752
      _ExtentX        =   20029
      _ExtentY        =   10769
      _StockProps     =   64
      DisplayRowHeaders=   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      GridShowVert    =   0   'False
      MaxCols         =   8
      MaxRows         =   1
      OperationMode   =   3
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "frm_FlowOpen.frx":000C
      UserResize      =   1
   End
End
Attribute VB_Name = "frm_FlowOpen"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/////////////////////////////////////////////////////////////////
'// View Call Flow List and Select to Open
'// �ļ�����  Frm_FlowOpen.frm
'// ��;��    ѡ�������
'// ����:     Tony Sun
'// �������ڣ�2007/03/25
'// �޸����ڣ�2007/03/25
'// �ļ�������ѡ�����������
'//////////////////////////////////////////////////////////////////
Option Explicit

Private Sub chkOpenNew_Click()
    gSystem.intOpenInNewWindow = chkOpenNew.value
    Call WriteIniFileString(Def_INI_SEC_OPTION, Def_INI_ENTRY_OPT_OpenInNewWindow, Str(gSystem.intOpenInNewWindow), gSystem.strINI_File)
End Sub

' ɾ������
'
Private Sub cmdDelete_Click()
    Dim lv_nFlowID As Byte
    Dim lv_nOldRow As Integer
    
    If vasFlows.Row <= 0 Then Exit Sub
    
    If Message("Q008") <> vbYes Then Exit Sub
    
    lv_nOldRow = vasFlows.Row
    vasFlows.Col = 1
    lv_nFlowID = CByte(vasFlows.value Mod 256)
    
    ' ��Ҫ�ж������Ƿ��Ѿ���
    If frmMain.IsCallFlowOpened(lv_nFlowID) Then
        Message "M143"
        frmMain.SwitchToMDIForm lv_nFlowID
        Exit Sub
    End If
    
    If DeleteIVRProgram(lv_nFlowID) Then
        
        '' Reload project table
        LoadProjectTable
        vasFlows.Row = lv_nOldRow
        
        Message "M010"
    Else
        Message "E026"
    End If
    
End Sub

Private Sub cmdOpen_Click()
'On Error Resume Next

    '����������
    Dim lv_msgResult As VbMsgBoxResult
    Dim lv_nFlowID As Byte
    Dim lv_loop As Integer
    Dim lv_SubLoop As Integer
        
    If vasFlows.Row <= 0 Then
        Message "E089"
        vasFlows.SetFocus
        Exit Sub
    End If
    
    '' Sun added 2007-12-10
    vasFlows.Row = vasFlows.ActiveRow
    
    vasFlows.Col = 1
    lv_nFlowID = CByte(vasFlows.value Mod 256)
    
    ' ��Ҫ�ж������Ƿ��Ѿ���
    If frmMain.IsCallFlowOpened(lv_nFlowID) Then
        Unload Me
        Message "M142"
        frmMain.SwitchToMDIForm lv_nFlowID
        Exit Sub
    End If
    
    If chkOpenNew.value = vbChecked Then
        
        '' Sun added 2007-10-20
        If gCallFlow.CallFlowID > 0 Then
            '' Enable Original Call Flow Form
            SetMainFormItemsEnableWhenPropertyShow True
        End If
        
        '' Open New CallFlow Window
        If Not frmMain.CreateNewMDIForm(lv_nFlowID) Then
            Unload Me
            Exit Sub
        End If
    Else
        If gCallFlow.CallFlowID > 0 Then
        
            '' Enable Original Call Flow Form
            SetMainFormItemsEnableWhenPropertyShow True

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
    
    SetMainFormItemsEnableWhenPropertyShow False
    
    '' ���ؽ�����Դ
    LoadResStrings Me
    
    '' ���������б�
    LoadProjectTable
    
    ''����Ĭ��ѡ����
    If vasFlows.MaxRows > 0 Then
        vasFlows.Row = 1
    Else
        vasFlows.Row = 0
    End If
    
    '' ��ǰ�����Ƿ��ǿ�����
    If gCallFlow.NewNodeID > 0 Then
        chkOpenNew.value = gSystem.intOpenInNewWindow
        chkOpenNew.Enabled = True
    Else
        chkOpenNew.value = vbUnchecked
        chkOpenNew.Enabled = False
    End If
        
End Sub

Private Sub LoadProjectTable()
On Error GoTo BackDoor

    Dim lv_RS_FS As ADODB.Recordset
    Dim lv_nRow As Integer
    
    '�ı����ָ����״->ɳ©���
    mdlcommon.ChangeMousePointer vbHourglass, True

    Set lv_RS_FS = New ADODB.Recordset
   
    With lv_RS_FS
        .CursorType = adOpenKeyset
        .LockType = adLockReadOnly
        .CursorLocation = adUseClient
        .Open "Select * from tbrefivr where P_Type = 'P' Order By P_ID", gSystem.strConString
    End With
 
    '������Ϣ������
    vasFlows.MaxRows = lv_RS_FS.RecordCount
    lv_nRow = 1
    Do While Not lv_RS_FS.EOF
        If Not IsNull(lv_RS_FS!P_ID) Then
            With vasFlows
                
                .Row = lv_nRow
                
                .Col = 1
                .Text = Trim(Str(lv_RS_FS!P_ID))
                If Not IsNull(lv_RS_FS!P_Name) Then
                    .Col = 2
                    .Text = Trim(lv_RS_FS!P_Name)                       '��������
                End If
                If Not IsNull(lv_RS_FS!P_Description) Then
                    .Col = 3
                    .Text = Trim(lv_RS_FS!P_Description)                '��������
                End If
                If Not IsNull(lv_RS_FS!P_Version) Then
                    .Col = 4
                    .Text = Trim(lv_RS_FS!P_Version)                    '�汾��
                End If
                If Not IsNull(lv_RS_FS!P_Auther) Then
                    .Col = 5
                    .Text = Trim(lv_RS_FS!P_Auther)                     '��������
                End If
                If Not IsNull(lv_RS_FS!P_User) Then
                    .Col = 6
                    .Text = Trim(lv_RS_FS!P_User)                       '�����û�
                End If
                If Not IsNull(lv_RS_FS!P_CreateTime) Then
                    .Col = 7
                    .Text = Format(lv_RS_FS!P_CreateTime, "yyyy-mm-dd hh:nn:ss")             '����ʱ��
                End If
                If Not IsNull(lv_RS_FS!P_ModifyTime) Then
                    .Col = 8
                    .Text = Format(lv_RS_FS!P_ModifyTime, "yyyy-mm-dd hh:nn:ss")             '�޸�ʱ��
                End If
                
            End With
            
            lv_RS_FS.MoveNext
            lv_nRow = lv_nRow + 1
            
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

' Delete Call Flow with specific pid
Private Function DeleteIVRProgram(ByVal f_PID As Integer) As Boolean
On Error GoTo BackDoor
    
    Dim lv_CN As ADODB.Connection        '' ����
    Dim lv_SQL As String
    Dim lv_InTrans As Boolean

    '�ı����ָ����״->ɳ©���
    mdlcommon.ChangeMousePointer vbHourglass, True

    DeleteIVRProgram = False
    lv_InTrans = False
    
    Set lv_CN = New ADODB.Connection
    lv_CN.ConnectionString = gSystem.strConString
    lv_CN.CursorLocation = adUseClient
    lv_CN.Open
    
    '' Begin Transaction
    lv_CN.BeginTrans
    lv_InTrans = True
    
    '' Delete IVR Project Title from project table
    lv_SQL = "Delete from tbRefIVR Where P_ID = " & Str(f_PID) & " And P_Type = 'P'"
    lv_CN.Execute lv_SQL, 1, adCmdText
    
    '' Delete IVR Project Details from call flow table
    'Update Jeremy 2004-07-08
    #If Programnew = 0 Then    '
        lv_SQL = "Delete from tbIVRProgramnew Where P_ID = " & Str(f_PID)
    #Else
        lv_SQL = "Delete from tbIVRProgram Where P_ID = " & Str(f_PID)
    #End If
    'Update End
    lv_CN.Execute lv_SQL, -1, adCmdText
    
    '' Commit Transaction
    lv_CN.CommitTrans
    lv_InTrans = False
    DeleteIVRProgram = True

BackDoor:
    
    ' Roll back if we are in transaction
    If lv_InTrans Then lv_CN.RollbackTrans
    
    '�ı����ָ����״->��ͷ���
    mdlcommon.ChangeMousePointer vbDefault, True
    
    On Error GoTo 0

End Function

Private Sub Form_Unload(Cancel As Integer)
    SetMainFormItemsEnableWhenPropertyShow True
End Sub

' ˫������Ŀ
'
Private Sub vasFlows_DblClick(ByVal Col As Long, ByVal Row As Long)
    vasFlows.Row = Row
    cmdOpen_Click
End Sub

Private Sub vasFlows_EnterRow(ByVal Row As Long, ByVal RowIsLast As Long)
    Debug.Print Row
End Sub

Private Sub vasFlows_Click(ByVal Col As Long, ByVal Row As Long)
    vasFlows.Row = Row
End Sub
