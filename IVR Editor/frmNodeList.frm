VERSION 5.00
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmNodeList 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "结点列表"
   ClientHeight    =   4920
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9255
   Icon            =   "frmNodeList.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4920
   ScaleWidth      =   9255
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1415"
   Begin VB.CommandButton cmdQuit 
      Caption         =   "取消(&C)"
      Height          =   375
      Left            =   8280
      TabIndex        =   2
      Tag             =   "1144"
      Top             =   4440
      Width           =   975
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "保存(&S)"
      Height          =   375
      Left            =   7200
      TabIndex        =   1
      Tag             =   "1007"
      Top             =   4440
      Width           =   975
   End
   Begin FPSpreadADO.fpSpread vasNodes 
      Height          =   4365
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9225
      _Version        =   458752
      _ExtentX        =   16272
      _ExtentY        =   7699
      _StockProps     =   64
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MaxCols         =   7
      MaxRows         =   20
      OperationMode   =   2
      ScrollBarExtMode=   -1  'True
      SelectBlockOptions=   0
      SpreadDesigner  =   "frmNodeList.frx":038A
      UserResize      =   1
   End
End
Attribute VB_Name = "frmNodeList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Mike added @ 2008-1-29
Private bChanged As Boolean

'Mike add this button event @ 2008-1-29
Private Sub cmdOK_Click()
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    'TODO : save the Node Tag to datesheet
    Call SaveNodeSpread
    
    bChanged = False
    cmdOK.Enabled = bChanged
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
End Sub

'Mike add this button event @ 2008-1-29
Private Sub cmdQuit_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Debug.Print gSystem.intRecFileType
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    Call RefreshNodesSpread
    
    'Mike add this @ 2008-1-29
    cmdOK.Enabled = bChanged
        
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    LoadResStrings Me
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        bChanged = False
        Unload Me
    End If
End Sub

Private Sub RefreshNodesSpread()
On Error Resume Next

    Dim lv_Index As Integer
    Dim lv_Rows As Integer
    Dim lv_seleRow As Integer

    vasNodes.Enabled = False
    vasNodes.MaxRows = 1
    lv_Rows = 1
    lv_seleRow = -1
        
    For lv_Index = 1 To gCallFlow.NewNodeID

        If gCallFlow.Node(lv_Index).NodeNo <> 255 And (g_NodeListShowType = 0 Or gCallFlow.Node(lv_Index).NodeNo = g_NodeListShowType) Then
            vasNodes.MaxRows = lv_Rows
            vasNodes.Row = lv_Rows
            lv_Rows = lv_Rows + 1
            
            If gCallFlow.Node(lv_Index).NodeID = gCallFlow.RootNodedID Then
                vasNodes.Col = 0
                vasNodes.Text = "Root"
            End If
            'Mike comment : 节点ID
            vasNodes.Col = 1
            vasNodes.Text = Trim(Str(gCallFlow.Node(lv_Index).NodeID))
            'Mike comment : 节点页码
            vasNodes.Col = 2
            vasNodes.Text = Trim(Str(gCallFlow.Node(lv_Index).InPage))
            'Mike comment : 节点类型
            vasNodes.Col = 3
            vasNodes.Text = gNodeTypeNameList(gCallFlow.Node(lv_Index).NodeNo)
            'Michael Added @ 2008-1-28 for <Node Tag>
            vasNodes.Col = 4
            vasNodes.Text = Trim(gCallFlow.Node(lv_Index).NodeTag)
            'Mike added @ 2008-2-2
            If gCallFlow.Node(lv_Index).NodeID < 255 Then vasNodes.Lock = True
            
            'Mike added @ 2008-2-19 for <Node Log>
            vasNodes.Col = 5
            vasNodes.Text = gCallFlow.GetNodeLogged(lv_Index)
            
            'Mike comment : 节点描述
            'Mike modify 4 to 5
            'Mike modified 5 to 6,2008-2-19
            vasNodes.Col = 6
            If Not IsNull(gCallFlow.Node(lv_Index).Description) Then
                vasNodes.Text = Trim(gCallFlow.Node(lv_Index).Description)
                'vasNodes.CellType = CellTypeEdit
                'vasNodes.Lock = False
            End If
            
            'Mike modify 5 to 6
            'Mike modified 6 to 7,2008-2-19
            vasNodes.Col = 7
            vasNodes.Text = Str(lv_Index)
            
            '' Select Node Item
            If Not gSystem.crlCurItem Is Nothing Then
                If gCallFlow.Node(lv_Index).NodeID = Val(gSystem.crlCurItem.Text) Then
                    lv_seleRow = vasNodes.Row
                End If
            End If
        End If
    Next

    vasNodes.Enabled = True
    
    If Not gSystem.crlCurItem Is Nothing Then
        vasNodes.ToolTipText = LoadNationalResString(1418)
    Else
        vasNodes.ToolTipText = LoadNationalResString(1419)
    End If
    
    If lv_seleRow > 0 Then
        'vasNodes.SelModeIndex = lv_seleRow
        'Mike Modified @ 2008-4-24
        'Reason:
        vasNodes.SetSelection 1, lv_seleRow, 6, lv_seleRow
    End If
    
    'Mike add this code for sorting  @ 2008-1-29
    vasNodes.UserColAction = UserColActionSort
    
On Error GoTo 0
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
    Dim msgUser As VbMsgBoxResult
    
    If bChanged = True Then
        msgUser = MsgBox(LoadNationalResString(1949) & "  ?", vbYesNoCancel + vbApplicationModal + vbQuestion)
        If msgUser = vbYes Then
            Call cmdOK_Click
        End If
        
        If msgUser = vbCancel Then Cancel = True
    End If
    
End Sub

' Goto Node / Edit Tag
Private Sub vasNodes_DblClick(ByVal Col As Long, ByVal Row As Long)
On Error GoTo ErrHandle:
    Dim lv_Index As Integer
                
    With vasNodes
        
        'Mike add this for user edit the tag of node 08-1-29
        If Col = 4 Then
            .Row = Row
            .Col = 4
            If .Lock = True Then
                Message ("M152")
                Exit Sub
            End If
        Else
            .EditMode = False
            .Row = Row
            'Mike changed index from 5 to 6 @ 2008-1-29
            'Mike changed index from 6 to 7 @ 2008-2-19
            .Col = 7
            lv_Index = Val(.Text)
            
            If lv_Index >= 1 And lv_Index <= gCallFlow.NewNodeID Then
                If gSystem.crlCurItem Is Nothing Then
                    If gCallFlow.Node(lv_Index).InPage <> gCallFlow.CurrentPage Then
                        Mdlfunction.GotoAnotherPage gCallFlow.Node(lv_Index).InPage
                    End If
                    gCallFlow.Node(lv_Index).SelectThisNode
                Else
                    If gCallFlow.Node(lv_Index).NodeID <> Val(gSystem.crlCurItem.Text) Then
                        gSystem.crlCurItem.Text = Trim(Str(gCallFlow.Node(lv_Index).NodeID))
                    End If
                End If
                Unload Me
            End If
        End If
        
    End With
Exit Sub

ErrHandle:
    Debug.Print Err.Description
End Sub

'Mike added @ 2008-1-30
Private Sub vasNodes_EditChange(ByVal Col As Long, ByVal Row As Long)
    If Col = 4 Then
        bChanged = True
        cmdOK.Enabled = bChanged
    End If
End Sub

' Mike Added @08-1-29 for input check(0~9,A~Z,a~z)
Private Sub vasNodes_KeyPress(KeyAscii As Integer)
    With vasNodes
        If (Not (KeyAscii >= 48 And KeyAscii <= 57)) And _
           (Not (KeyAscii >= 97 And KeyAscii <= 122)) And _
           (Not (KeyAscii >= 65 And KeyAscii <= 90)) And _
            KeyAscii <> 8 And KeyAscii <> 7 Then KeyAscii = 0
    End With
End Sub

Private Sub SaveNodeSpread()
On Error GoTo ErrHandle
    Dim lv_loop As Integer
    'Dim lv_row As Integer
    Dim lv_NodeID As Integer
    Dim lv_NodeIndex As Byte
    
    'Mike modified @ 2008-2-2, fixed sorting Tags save
    'lv_row = 1
        'For lv_loop = 1 To gCallFlow.NewNodeID
'            If gCallFlow.Node(lv_loop).NodeNo <> 255 And _
'            (g_NodeListShowType = 0 Or _
'            gCallFlow.Node(lv_loop).NodeNo = g_NodeListShowType) Then
'                .Row = lv_row
'                lv_row = lv_row + 1
'                .Col = 4
'                If gCallFlow.Node(lv_loop).NodeTag <> Trim(.Text) Then
'                    gCallFlow.Node(lv_loop).NodeTag = Trim(.Text)
'                    'gCallFlow.UpdateIvrRecord CInt(gCallFlow.Node(lv_loop).NodeID), _
'                    '                         CByte(gCallFlow.Node(lv_loop).NodeNo)
'                    gCallFlow.Node(lv_loop).ModifyFlag = DEF_OPERATION_MODIFY
'                End If
'            End If
            '*******************************************************
    With vasNodes
        For lv_loop = 1 To .MaxRows
            .Col = 1
            .Row = lv_loop
            lv_NodeID = .Text
            If lv_NodeID > 255 Then
                lv_NodeIndex = gCallFlow.SearchNodeIndexWithID(lv_NodeID)
                If gCallFlow.Node(lv_NodeIndex).NodeNo <> 255 And _
                   (g_NodeListShowType = 0 Or _
                   gCallFlow.Node(lv_NodeIndex).NodeNo = g_NodeListShowType) Then
                    .Col = 4
                    If gCallFlow.Node(lv_NodeIndex).NodeTag <> Trim(.Text) Then
                        gCallFlow.Node(lv_NodeIndex).NodeTag = Trim(.Text)
                        gCallFlow.UpdateAnotherIVRRecord (lv_NodeIndex)
                    End If
                End If
            End If
        Next lv_loop
    End With

Exit Sub
ErrHandle:
Debug.Print "SaveTagEror : " & Err.Description
    
End Sub
