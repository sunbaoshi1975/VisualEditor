VERSION 5.00
Begin VB.Form frmProjectSearch 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "��Ŀ��ѯ"
   ClientHeight    =   810
   ClientLeft      =   6105
   ClientTop       =   5040
   ClientWidth     =   6360
   Icon            =   "frmProjectSearch.frx":0000
   LinkTopic       =   "��Ŀ��ѯ"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   810
   ScaleWidth      =   6360
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdCancel 
      Caption         =   "ȡ��"
      Height          =   255
      Left            =   5400
      TabIndex        =   7
      Top             =   480
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "ȷ��"
      Height          =   255
      Left            =   5400
      TabIndex        =   6
      Top             =   120
      Width           =   735
   End
   Begin VB.TextBox txtSchState 
      Height          =   285
      Left            =   2880
      TabIndex        =   5
      Text            =   "Text1"
      Top             =   360
      Width           =   2175
   End
   Begin VB.ComboBox cboSchOpt 
      Height          =   315
      Left            =   1680
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   360
      Width           =   975
   End
   Begin VB.ComboBox cboSchName 
      Height          =   315
      Left            =   120
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label lblSearchSta 
      AutoSize        =   -1  'True
      Caption         =   "���ʽ"
      Height          =   195
      Left            =   3600
      TabIndex        =   2
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblSearchOpt 
      AutoSize        =   -1  'True
      Caption         =   "�����"
      Height          =   195
      Left            =   1800
      TabIndex        =   1
      Top             =   120
      Width           =   540
   End
   Begin VB.Label lblSearchName 
      AutoSize        =   -1  'True
      Caption         =   "�ֶ�����"
      Height          =   195
      Left            =   240
      TabIndex        =   0
      Top             =   120
      Width           =   720
   End
End
Attribute VB_Name = "frmProjectSearch"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'/ File Version : V1.0
'/ Author   : Michael S
'/ Date     : July,13,07
'///////////////////////////////////////////////////

Option Explicit
Dim strSQL As String

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If cboSchName.Text = "" Then
        Message "E134"
        cboSchName.SetFocus
        Exit Sub
    End If
    
    If cboSchOpt.Text = "" Then
        Message "E135"
        cboSchOpt.SetFocus
        Exit Sub
    End If
    
    If txtSchState.Text = "" Then
        Message "E136"
        txtSchState.SetFocus
        Exit Sub
    End If
    
    If cboSchName.ListIndex = 0 Then
        strSQL = "Select * from tbrefivr where P_Type = 'R' and P_ID  " & cboSchOpt.Text & " " & txtSchState.Text & " Order By P_ID"
    ElseIf cboSchName.ListIndex = 1 Then
        strSQL = "Select * from tbrefivr where P_Type = 'R' and P_Version  " & cboSchOpt.Text & " '" & txtSchState.Text & "' Order By P_ID"
    ElseIf cboSchName.ListIndex = 2 Then
        strSQL = "Select * from tbrefivr where P_Type = 'R' and P_Name  " & cboSchOpt.Text & " '" & txtSchState.Text & "' Order By P_ID"
    ElseIf cboSchName.ListIndex = 3 Then
        strSQL = "Select * from tbrefivr where P_Type = 'R' and P_Auther  " & cboSchOpt.Text & " '" & txtSchState.Text & "' Order By P_ID"
    ElseIf cboSchName.ListIndex = 4 Then
        strSQL = "Select * from tbrefivr where P_Type = 'R' and P_User  " & cboSchOpt.Text & " '" & txtSchState.Text & "' Order By P_ID"
    ElseIf cboSchName.ListIndex = 5 Then
        strSQL = "Select * from tbrefivr where P_Type = 'R' and P_CreateTime  " & cboSchOpt.Text & " '" & txtSchState.Text & "' Order By P_ID"
    ElseIf cboSchName.ListIndex = 6 Then
        strSQL = "Select * from tbrefivr where P_Type = 'R' and P_ModifyTime  " & cboSchOpt.Text & " '" & txtSchState.Text & "' Order By P_ID"
    ElseIf cboSchName.ListIndex = 7 Then
        strSQL = "Select * from tbrefivr where P_Type = 'R' and P_Description  " & cboSchOpt.Text & " '" & txtSchState.Text & "' Order By P_ID"
    'ElseIf cboSchName.ListIndex = 8 Then
    End If
    ReFillProList
    Unload Me
End Sub

Private Sub Form_Load()
    ' Fill and init the combo box - �ֶ�����,�����
    cboSchName.Clear
    cboSchName.AddItem "��Ŀ���", 0
    cboSchName.AddItem "��Ŀ�汾", 1
    cboSchName.AddItem "��Ŀ����", 2
    cboSchName.AddItem "��Ŀ����", 3
    cboSchName.AddItem "��Ŀ�û�", 4
    cboSchName.AddItem "����ʱ��", 5
    cboSchName.AddItem "�޸�ʱ��", 6
    cboSchName.AddItem "��Ŀ����", 7
    'cboSchName.AddItem "��Ŀ����", 8
    
    cboSchOpt.Clear
    cboSchOpt.AddItem ">", 0
    cboSchOpt.AddItem ">=", 1
    cboSchOpt.AddItem "=", 2
    cboSchOpt.AddItem "<", 3
    cboSchOpt.AddItem "<=", 4
    cboSchOpt.AddItem "like", 5
    
    txtSchState.Text = ""
End Sub

Public Sub ReFillProList()
    On Error GoTo BackDoor

    Dim itmX As ListItem
    Dim intCount As Integer
    
'Clear the old informations
frmProjectList.lstProject.ListItems.Clear

' Add the contect
With gCallFlow.RS_Project
    
    .CursorType = adOpenKeyset
    .LockType = adLockReadOnly
    .CursorLocation = adUseClient
    .Open strSQL, gSystem.strConString

    If .RecordCount > 0 Then
        .MoveFirst
        While Not .EOF
            Set itmX = frmProjectList.lstProject.ListItems.Add(, , .Fields("P_ID"))
            
            '��Ŀ����
            If Not IsNull(.Fields("P_Name")) Then
                itmX.SubItems(1) = Trim(.Fields("P_Name"))
            End If
            
            '��Ŀ����
            If Not IsNull(.Fields("P_Description")) Then
                itmX.SubItems(2) = Trim(.Fields("P_Description"))
            End If
            
            '��Ŀ�汾
            If Not IsNull(.Fields("P_Version")) Then
                itmX.SubItems(3) = Trim(.Fields("P_Version"))
            End If
            
            '��Ŀ����
            If Not IsNull(.Fields("P_Auther")) Then
                itmX.SubItems(4) = Trim(.Fields("P_Auther"))
            End If
            
            '��Ŀ�û�
            If Not IsNull(.Fields("P_User")) Then
                itmX.SubItems(5) = Trim(.Fields("P_User"))
            End If
            
            '����ʱ��
            If Not IsNull(.Fields("P_CreateTime")) Then
                itmX.SubItems(6) = Format(.Fields("P_CreateTime"), "yyyy-mm-dd hh:nn:ss")
            End If
            
            '�޸�ʱ��
            If Not IsNull(.Fields("P_ModifyTime")) Then
                itmX.SubItems(7) = Format(.Fields("P_ModifyTime"), "yyyy-mm-dd hh:nn:ss")
            End If
            
            '��Ŀ����
            If Not IsNull(.Fields("P_Type")) Then
                itmX.SubItems(8) = Trim(.Fields("P_Type"))
            End If
            
            .MoveNext
        Wend
    End If
    .Close
End With
            
frmProjectList.lstProject.Refresh

BackDoor:
    On Error GoTo 0
End Sub
