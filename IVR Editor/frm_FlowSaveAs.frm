VERSION 5.00
Begin VB.Form frm_FlowSaveAs 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "流程另存"
   ClientHeight    =   1110
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5280
   Icon            =   "frm_FlowSaveAs.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1110
   ScaleWidth      =   5280
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1715"
   Begin VB.CommandButton exit_Command 
      Caption         =   "退出(&E)"
      Height          =   315
      Left            =   3630
      TabIndex        =   3
      Tag             =   "1144"
      Top             =   660
      Width           =   1065
   End
   Begin VB.CommandButton CommandCreate 
      Caption         =   "确定(&C)"
      Default         =   -1  'True
      Height          =   315
      Left            =   2250
      TabIndex        =   2
      Tag             =   "1372"
      Top             =   660
      Width           =   1065
   End
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
      Left            =   2250
      TabIndex        =   0
      Top             =   150
      Width           =   2895
   End
   Begin VB.Label lblNewPID 
      AutoSize        =   -1  'True
      Caption         =   "输入新流程ID"
      Height          =   195
      Left            =   150
      TabIndex        =   1
      Tag             =   "1716"
      Top             =   210
      Width           =   1065
   End
End
Attribute VB_Name = "frm_FlowSaveAs"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim lv_RS_FS As ADODB.Recordset

Private Sub Cmb_pid_GotFocus()
    Cmb_pid.SelStart = 0
    Cmb_pid.SelLength = Len(Cmb_pid)
End Sub

Private Sub Cmb_pid_KeyPress(KeyAscii As Integer)
    KeyPress KeyAscii
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

Private Sub CommandCreate_Click()
On Error Resume Next

    '流程另存
    Dim lv_msgResult As VbMsgBoxResult
    Dim lv_nFlowID As Byte
    Dim lv_loop As Integer
    Dim lv_SubLoop As Integer
    
    ' 新流程ID
    If Trim(Cmb_pid) = "" Then
        Message "E127"
        Cmb_pid.SetFocus
        Exit Sub
    End If
    lv_nFlowID = CByte(Val(Trim(Cmb_pid.Text)) Mod 256)
    
    ' 需要判断流程是否已存在
    lv_RS_FS.MoveFirst
    lv_RS_FS.Find "p_id=" & lv_nFlowID
    If Not lv_RS_FS.EOF Then '新流程信息
        Message "E128"
        Cmb_pid.SetFocus
    Else
        If gCallFlow.SaveCallFlowAs(lv_nFlowID) Then
            Message "M145"
            Unload Me
        Else
            Message "E129"
            Cmb_pid.SetFocus
        End If
    End If

End Sub

Private Sub exit_Command_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    
    LoadProjectTable
    LoadResStrings Me

End Sub
