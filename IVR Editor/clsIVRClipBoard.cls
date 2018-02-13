VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIVRClipBoard"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------
' IVR clipboard class
'-----------------------------------------------
Option Explicit

'' Private variable
Private m_ClipBoard() As New clsIVRClipItem

'' Properties
Private m_bytMaxStacks As Byte                                  '' 粘贴板最大堆栈数
Private m_blnClipboardMark As Boolean                           '' 粘贴板是否有内容
Private m_CopyedItem As New clsIVRClipItem                      '' 复制的内容
Private m_nClipItem As Byte                                     '' 粘贴板当前内容指针
Private m_nReDoTimes As Byte                                    '' 粘贴板回滚内容指针

Private Sub Class_Initialize()

    '' 变量初始化
    InitLocalVariables

End Sub

Private Sub Class_Terminate()

    '' 结束
    
End Sub

' 变量初始化
'
Private Sub InitLocalVariables()
    
    m_bytMaxStacks = 0
    m_blnClipboardMark = False
    m_CopyedItem.Operation = 0
    m_nClipItem = 0
    m_nReDoTimes = 0
    
End Sub

Public Property Let MaxStacks(f_bytStacks As Byte)
On Error Resume Next

    If m_bytMaxStacks <> f_bytStacks Then
        ReDim Preserve m_ClipBoard(f_bytStacks) As New clsIVRClipItem
        m_bytMaxStacks = f_bytStacks
    End If

End Property

Public Property Get MaxStacks() As Byte
On Error Resume Next

    MaxStacks = m_bytMaxStacks
        
End Property

Public Property Let ClipboardMark(f_blnMark As Boolean)
On Error Resume Next
  
    m_blnClipboardMark = f_blnMark
        
End Property

Public Property Get ClipboardMark() As Boolean
On Error Resume Next

    ClipboardMark = m_blnClipboardMark
        
End Property

Public Property Let CopyedItem(f_ClipItem As clsIVRClipItem)
On Error Resume Next
  
    Set m_CopyedItem = f_ClipItem
        
End Property

Public Property Get CopyedItem() As clsIVRClipItem
On Error Resume Next

    Set CopyedItem = m_CopyedItem
        
End Property

Public Property Let ClipItem(f_bytItem As Byte)
On Error Resume Next
  
    m_nClipItem = f_bytItem

End Property

Public Property Get ClipItem() As Byte
On Error Resume Next

    ClipItem = m_nClipItem
        
End Property

Public Property Let ReDoTimes(f_bytTimes As Byte)
On Error Resume Next
  
    m_nReDoTimes = f_bytTimes

End Property

Public Property Get ReDoTimes() As Byte
On Error Resume Next

    ReDoTimes = m_nReDoTimes
        
End Property

Public Property Get ClipboardItems(f_bytItemIndex As Byte) As clsIVRClipItem
On Error Resume Next

    If f_bytItemIndex >= 0 And f_bytItemIndex <= m_bytMaxStacks Then
        Set ClipboardItems = m_ClipBoard(f_bytItemIndex)
    End If
        
End Property

' ============================================
' 单节点粘贴板压栈操作
' 1 - Modify
' 2 - New
' 3 - Delete
'
' ============================================
Public Sub PushClipBoardStack(Optional ByVal f_Op As Byte = 1)

    '' Parameter Assert
    If f_Op < DEF_OPERATION_INIT Or f_Op > DEF_OPERATION_DELETE Then Exit Sub
    
    '' Arrange Old Stacks
    Dim lv_clipindex As Byte
    If m_nClipItem > m_bytMaxStacks Then
        m_nClipItem = m_bytMaxStacks
        For lv_clipindex = 1 To m_bytMaxStacks
            CopyClipBoardItem lv_clipindex - 1, lv_clipindex
        Next
    End If
    
    '' Add New Stack
    m_ClipBoard(m_nClipItem).NodeCount = 1
    m_ClipBoard(m_nClipItem).Operation = f_Op
    CopyNode2ClipBoard m_nClipItem, 0, gCallFlow.NodeSelectedID
    
    '' Change Mene Status
    F_SetUnDoMenuState 1
    
End Sub

' ============================================
' 多节点粘贴板压栈初始化
' 1 - Modify
' 2 - New
' 3 - Delete
'
' ============================================
Public Function MultiPushInitialize(ByVal f_NodeCount As Byte, ByVal f_Op As Byte) As Boolean
On Error Resume Next

    Dim lv_clipindex As Byte
    
    MultiPushInitialize = False
    
    '' Parameter Assert
    If f_Op < DEF_OPERATION_INIT Or f_Op > DEF_OPERATION_DELETE Then Exit Function
    
    '' How many selected nodes
    If f_NodeCount = 0 Then Exit Function
    
    '' Arrange Old Stacks
    If m_nClipItem > m_bytMaxStacks Then
        m_nClipItem = m_bytMaxStacks
        For lv_clipindex = 1 To m_bytMaxStacks
            CopyClipBoardItem lv_clipindex - 1, lv_clipindex
        Next
    End If
    
    m_ClipBoard(m_nClipItem).NodeCount = f_NodeCount
    m_ClipBoard(m_nClipItem).Operation = f_Op
    m_ClipBoard(m_nClipItem).NodePointer = 0
    
    '' Change Mene Status
    F_SetUnDoMenuState 1
    
    MultiPushInitialize = True
    
End Function

' ============================================
' 多节点粘贴板压栈操作
' 必须跟在MultiPushInitialize后面，连续使用
'
' ============================================
Public Sub MultiPushClipBoardStack(ByVal f_nNodeIndex As Integer)
On Error Resume Next

    Dim lv_Item As Byte
    
    lv_Item = m_nClipItem - 1
    
    If m_ClipBoard(lv_Item).NodePointer < m_ClipBoard(lv_Item).NodeCount Then
        CopyNode2ClipBoard lv_Item, m_ClipBoard(lv_Item).NodePointer, f_nNodeIndex
        m_ClipBoard(lv_Item).NodePointer = m_ClipBoard(lv_Item).NodePointer + 1
    End If
    
End Sub

Public Sub GetCopyedItem()
On Error Resume Next
    
    Dim lv_nSelectedCount As Integer
    Dim lv_Index As Integer
    Dim lv_clipindex As Byte
    
    '' How many selected nodes
    lv_nSelectedCount = gCallFlow.SelectedCount
    If lv_nSelectedCount = 0 Then Exit Sub
    
    CopyedItem.NodeCount = lv_nSelectedCount
    ClipboardMark = True
    For lv_Index = 1 To gCallFlow.NewNodeID
        If gCallFlow.Node(lv_Index).IsSelected Then
            CopyNodeData CopyedItem.Contents(lv_clipindex), gCallFlow.Node(lv_Index)
            lv_clipindex = lv_clipindex + 1
        End If
    Next
    
End Sub

Public Sub ReDoOperation()
On Error Resume Next

    Dim lv_nItemNode As Byte
    Dim lv_Item As Byte
    Dim lv_nNodeIndex As Integer
    
    Do While ReDoTimes > 0
        
        lv_Item = ClipItem
        If m_ClipBoard(lv_Item).Operation > DEF_OPERATION_INIT Then
            If m_ClipBoard(lv_Item).NodeCount > 0 Then
                
                '' Select Call Flow
                frmMain.SwitchToMDIForm m_ClipBoard(lv_Item).Contents(0).IntP_id
            
                For lv_nItemNode = 0 To m_ClipBoard(lv_Item).NodeCount - 1
    
                    lv_nNodeIndex = m_ClipBoard(lv_Item).Contents(lv_nItemNode).IntN_index
                    Select Case m_ClipBoard(lv_Item).Operation
                    Case DEF_OPERATION_MODIFY       '' ReModify
                        CopyClipBoard2Node lv_Item, lv_nItemNode
                        gCallFlow.UpdateAnotherIVRRecord lv_nNodeIndex
                        
                    Case DEF_OPERATION_NEW          '' ReNew
                        If gCallFlow.CreateNode(lv_nNodeIndex) Then
                            gCallFlow.Node(lv_nNodeIndex).FlowID = gCallFlow.CallFlowID
                            CopyClipBoard2Node lv_Item, lv_nItemNode, False
                            If gCallFlow.Node(lv_nNodeIndex).NodeNo = 255 Then
                                Call gCallFlow.Node(lv_nNodeIndex).AddLine
                                gCallFlow.Node(lv_nNodeIndex).MoveImageWithLine
                            End If
                            gCallFlow.Node(lv_nNodeIndex).Visible = True
                            '' Sun added 2002-03-30
                            gCallFlow.Node(lv_nNodeIndex).NodeCaptionVisible = gblnShowNodeCaption
                            
                            '' Sun added 2008-01-18
                            gCallFlow.Node(lv_nNodeIndex).NodeTagVisible = gblnShowNodeTag
                            
                            gCallFlow.AddNewIvrRecord lv_nNodeIndex
                        End If
                        
                    Case DEF_OPERATION_DELETE       '' ReDelete
                        If Not gCallFlow.Node(lv_nNodeIndex) Is Nothing Then
                            If Not gCallFlow.Node(lv_nNodeIndex).NodeDestroyed Then
                                gCallFlow.Node(lv_nNodeIndex).Destroy
                                If lv_nNodeIndex = gCallFlow.NodeSelectedID Then
                                    gCallFlow.NodeSelectedID = 0
                                End If
                                gCallFlow.DeleteIvrRecord lv_nNodeIndex
                            End If
                        End If
                        
                    End Select
                    
                Next
            
            End If
            
            If ReDoTimes > 1 Then
                F_SetUnDoMenuState 4
            Else
                F_SetUnDoMenuState 1
            End If
            
            Exit Do
            
        Else
        
            If ReDoTimes > 1 Then
                F_SetUnDoMenuState 4
            Else
                F_SetUnDoMenuState 1
            End If
            
        End If

    Loop
    
End Sub

Public Sub UnDoOperation()
On Error Resume Next
    
    Dim lv_nItemNode As Byte
    Dim lv_Item As Byte
    Dim lv_nNodeIndex As Integer
    Dim lv_bytOldModifyFlag As Byte
    
    Do While ClipItem > 0
        
        lv_Item = ClipItem - 1
        If m_ClipBoard(lv_Item).Operation > DEF_OPERATION_INIT Then
            If m_ClipBoard(lv_Item).NodeCount > 0 Then
            
                '' Select Call Flow
                frmMain.SwitchToMDIForm m_ClipBoard(lv_Item).Contents(0).IntP_id
                
                For lv_nItemNode = 0 To m_ClipBoard(lv_Item).NodeCount - 1
                
                    lv_nNodeIndex = m_ClipBoard(lv_Item).Contents(lv_nItemNode).IntN_index
                    lv_bytOldModifyFlag = gCallFlow.Node(lv_nNodeIndex).ModifyFlag
                    Select Case m_ClipBoard(lv_Item).Operation
                    Case DEF_OPERATION_MODIFY          '' UnModify
                        CopyClipBoard2Node lv_Item, lv_nItemNode
                        If lv_bytOldModifyFlag = DEF_OPERATION_INIT Then
                            gCallFlow.Node(lv_nNodeIndex).ModifyFlag = DEF_OPERATION_MODIFY
                            gCallFlow.UpdateAnotherIVRRecord lv_nNodeIndex
                        End If
                                            
                    Case DEF_OPERATION_NEW             '' UnNew
                        If Not gCallFlow.Node(lv_nNodeIndex) Is Nothing Then
                            If Not gCallFlow.Node(lv_nNodeIndex).NodeDestroyed Then
                                gCallFlow.Node(lv_nNodeIndex).Destroy
                                If lv_nNodeIndex = gCallFlow.NodeSelectedID Then
                                    gCallFlow.NodeSelectedID = 0
                                End If
                                If lv_bytOldModifyFlag <> DEF_OPERATION_NEW Then
                                    gCallFlow.DeleteIvrRecord lv_nNodeIndex
                                    m_ClipBoard(lv_Item).Contents(lv_nItemNode).ModifyFlag = DEF_OPERATION_INIT
                                Else
                                    gCallFlow.Node(lv_nNodeIndex).ModifyFlag = DEF_OPERATION_INIT
                                End If
                            End If
                        End If
                        
                    Case DEF_OPERATION_DELETE       '' UnDelete
                        If gCallFlow.CreateNode(lv_nNodeIndex) Then
                            gCallFlow.Node(lv_nNodeIndex).FlowID = gCallFlow.CallFlowID
                            CopyClipBoard2Node lv_Item, lv_nItemNode, False
                            If gCallFlow.Node(lv_nNodeIndex).NodeNo = 255 Then
                                Call gCallFlow.Node(lv_nNodeIndex).AddLine
                                gCallFlow.Node(lv_nNodeIndex).MoveImageWithLine
                            End If
                            gCallFlow.Node(lv_nNodeIndex).Visible = True
                            '' Sun added 2002-03-30
                            gCallFlow.Node(lv_nNodeIndex).NodeCaptionVisible = gblnShowNodeCaption
                            
                            '' Sun added 2008-01-18
                            gCallFlow.Node(lv_nNodeIndex).NodeTagVisible = gblnShowNodeTag
                            
                            If lv_bytOldModifyFlag = DEF_OPERATION_INIT Then
                                gCallFlow.AddNewIvrRecord lv_nNodeIndex
                            End If
                        End If
                        
                    End Select
                    
                Next
                
            End If
    
            If ClipItem > 1 Then
                F_SetUnDoMenuState 3
            Else
                F_SetUnDoMenuState 2
            End If
        
            Exit Do
            
        Else
            
            If ClipItem > 1 Then
                F_SetUnDoMenuState 3
            Else
                F_SetUnDoMenuState 2
            End If
        
        End If
        
    Loop

End Sub

Public Sub CopyNode2ClipBoard(ByVal f_ItemIndex As Byte, ByVal f_ItemNode As Byte, ByVal f_NodeIndex As Integer)
On Error Resume Next

'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntP_id = gCallFlow.Node(f_NodeIndex).FlowID
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_index = f_NodeIndex
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_id = gCallFlow.Node(f_NodeIndex).NodeID
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_no = gCallFlow.Node(f_NodeIndex).NodeNo
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_page = gCallFlow.Node(f_NodeIndex).InPage
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_top = gCallFlow.Node(f_NodeIndex).Top
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_left = gCallFlow.Node(f_NodeIndex).Left
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_height = gCallFlow.Node(f_NodeIndex).Height
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_width = gCallFlow.Node(f_NodeIndex).Width
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).StrN_data1 = gCallFlow.Node(f_NodeIndex).Data1
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).StrN_data2 = gCallFlow.Node(f_NodeIndex).Data2
'    m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).StrN_description = gCallFlow.Node(f_NodeIndex).Description
    CopyNodeData m_ClipBoard(f_ItemIndex).Contents(f_ItemNode), gCallFlow.Node(f_NodeIndex)
                
End Sub

Private Sub CopyNodeData(f_ItemData As clsIVRClipItemData, f_Node As CNode)
On Error Resume Next

    f_ItemData.IntP_id = f_Node.FlowID
    f_ItemData.IntN_index = f_Node.NodeIndex
    
    f_ItemData.IntN_id = f_Node.NodeID
    f_ItemData.IntN_no = f_Node.NodeNo
    f_ItemData.IntN_page = f_Node.InPage
    f_ItemData.IntN_top = f_Node.Top
    f_ItemData.IntN_left = f_Node.Left
    f_ItemData.IntN_height = f_Node.Height
    f_ItemData.IntN_width = f_Node.Width
    f_ItemData.StrN_data1 = f_Node.Data1
    f_ItemData.StrN_data2 = f_Node.Data2
    f_ItemData.StrN_description = f_Node.Description
    f_ItemData.ModifyFlag = f_Node.ModifyFlag
    
    If f_Node.NodeNo = 255 Then
        f_ItemData.Line_X1 = f_Node.Line_X1
        f_ItemData.Line_X2 = f_Node.Line_X2
        f_ItemData.Line_Y1 = f_Node.Line_Y1
        f_ItemData.Line_Y2 = f_Node.Line_Y2
        f_ItemData.Line_Color = f_Node.Line_Color
        f_ItemData.Line_Width = f_Node.Line_Width
        f_ItemData.Line_Style = f_Node.Line_Style
    End If
                
End Sub

Public Sub CopyClipBoard2Node(ByVal f_ItemIndex As Byte, ByVal f_ItemNode As Byte, Optional ByVal f_Exchange As Boolean = True)
On Error Resume Next

    Dim lv_nNodeIndex As Integer
    
    lv_nNodeIndex = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_index
    RestoreNodeData m_ClipBoard(f_ItemIndex).Contents(f_ItemNode), gCallFlow.Node(lv_nNodeIndex), f_Exchange
    
'    lv_nTmp = gCallFlow.Node(lv_nNodeIndex).NodeID
'    gCallFlow.Node(lv_nNodeIndex).NodeID = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_id
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_id = lv_nTmp
'
'    lv_nTmp = gCallFlow.Node(lv_nNodeIndex).NodeNo
'    gCallFlow.Node(lv_nNodeIndex).NodeNo = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_no
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_no = lv_nTmp
'
'    lv_nTmp = gCallFlow.Node(lv_nNodeIndex).InPage
'    gCallFlow.Node(lv_nNodeIndex).InPage = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_page
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_page = lv_nTmp
'
'    lv_nTmp = gCallFlow.Node(lv_nNodeIndex).Left
'    gCallFlow.Node(lv_nNodeIndex).Left = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_left
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_left = lv_nTmp
'
'    lv_nTmp = gCallFlow.Node(lv_nNodeIndex).Top
'    gCallFlow.Node(lv_nNodeIndex).Top = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_top
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_top = lv_nTmp
'
'    lv_nTmp = gCallFlow.Node(lv_nNodeIndex).Height
'    gCallFlow.Node(lv_nNodeIndex).Height = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_height
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_height = lv_nTmp
'
'    lv_nTmp = gCallFlow.Node(lv_nNodeIndex).Width
'    gCallFlow.Node(lv_nNodeIndex).Width = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_width
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).IntN_width = lv_nTmp
'
'    lv_sTmp = gCallFlow.Node(lv_nNodeIndex).Data1
'    gCallFlow.Node(lv_nNodeIndex).Data1 = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).StrN_data1
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).StrN_data1 = lv_sTmp
'
'    lv_sTmp = gCallFlow.Node(lv_nNodeIndex).Data2
'    gCallFlow.Node(lv_nNodeIndex).Data2 = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).StrN_data2
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).StrN_data2 = lv_sTmp
'
'    lv_sTmp = gCallFlow.Node(lv_nNodeIndex).Description
'    gCallFlow.Node(lv_nNodeIndex).Description = m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).StrN_description
'    If f_Exchange Then m_ClipBoard(f_ItemIndex).Contents(f_ItemNode).StrN_description = lv_sTmp
                    
End Sub

Public Sub RestoreNodeData(f_ItemData As clsIVRClipItemData, f_Node As CNode, Optional ByVal f_Exchange As Boolean = True, Optional ByVal f_MoveLinesWithNode As Boolean = True)
On Error Resume Next

    Dim lv_nTmp, lv_sTmp
    
    lv_nTmp = f_Node.NodeID
    f_Node.NodeID = f_ItemData.IntN_id
    If f_Exchange Then f_ItemData.IntN_id = lv_nTmp
    
    lv_nTmp = f_Node.NodeNo
    f_Node.NodeNo = f_ItemData.IntN_no
    If f_Exchange Then f_ItemData.IntN_no = lv_nTmp

    lv_nTmp = f_Node.InPage
    f_Node.InPage = f_ItemData.IntN_page
    If f_Exchange Then f_ItemData.IntN_page = lv_nTmp

    lv_nTmp = f_Node.Left
    f_Node.Left = f_ItemData.IntN_left
    If f_Exchange Then f_ItemData.IntN_left = lv_nTmp

    lv_nTmp = f_Node.Top
    f_Node.Top = f_ItemData.IntN_top
    If f_Exchange Then f_ItemData.IntN_top = lv_nTmp

    lv_nTmp = f_Node.Height
    f_Node.Height = f_ItemData.IntN_height
    If f_Exchange Then f_ItemData.IntN_height = lv_nTmp

    lv_nTmp = f_Node.Width
    f_Node.Width = f_ItemData.IntN_width
    If f_Exchange Then f_ItemData.IntN_width = lv_nTmp

    lv_sTmp = f_Node.Data1
    f_Node.Data1 = f_ItemData.StrN_data1
    If f_Exchange Then f_ItemData.StrN_data1 = lv_sTmp

    lv_sTmp = f_Node.Data2
    f_Node.Data2 = f_ItemData.StrN_data2
    If f_Exchange Then f_ItemData.StrN_data2 = lv_sTmp

    lv_sTmp = f_Node.Description
    f_Node.Description = f_ItemData.StrN_description
    If f_Exchange Then f_ItemData.StrN_description = lv_sTmp
    
    lv_nTmp = f_Node.ModifyFlag
    f_Node.ModifyFlag = f_ItemData.ModifyFlag
    If f_Exchange Then f_ItemData.ModifyFlag = lv_nTmp
    
    If f_Node.NodeNo = 255 Then
        f_Node.Line_X1 = f_ItemData.Line_X1
        f_Node.Line_X2 = f_ItemData.Line_X2
        f_Node.Line_Y1 = f_ItemData.Line_Y1
        f_Node.Line_Y2 = f_ItemData.Line_Y2
        f_Node.Line_Color = f_ItemData.Line_Color
        f_Node.Line_Width = f_ItemData.Line_Width
        f_Node.Line_Style = f_ItemData.Line_Style
        f_Node.MoveImageWithLine
    Else
        If f_MoveLinesWithNode Then
            f_Node.MoveLinesOnNode
        End If
    End If
                
End Sub

Public Sub CopyClipBoardItem(ByVal f_ItemIndex As Byte, ByVal f_SourceItem As Byte)
On Error Resume Next

    Dim lv_nNodeCount As Byte
    Dim lv_nIndex As Byte
    
    lv_nNodeCount = m_ClipBoard(f_SourceItem).NodeCount
    m_ClipBoard(f_ItemIndex).NodeCount = lv_nNodeCount
    m_ClipBoard(f_ItemIndex).Operation = m_ClipBoard(f_SourceItem).Operation
    
    For lv_nIndex = 0 To lv_nNodeCount - 1
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntP_id = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).IntP_id
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntN_index = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).IntN_index
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntN_id = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).IntN_id
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntN_no = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).IntN_no
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntN_page = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).IntN_page
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntN_left = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).IntN_left
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntN_top = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).IntN_top
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntN_height = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).IntN_height
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntN_width = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).IntN_width
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).StrN_description = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).StrN_description
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).StrN_data1 = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).StrN_data1
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).StrN_data2 = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).StrN_data2
        If m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).IntN_no = 255 Then
            m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).Line_X1 = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).Line_X1
            m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).Line_X2 = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).Line_X2
            m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).Line_Y1 = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).Line_Y1
            m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).Line_Y2 = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).Line_Y2
            m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).Line_Color = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).Line_Color
            m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).Line_Width = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).Line_Width
            m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).Line_Style = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).Line_Style
        End If
        m_ClipBoard(f_ItemIndex).Contents(lv_nIndex).ModifyFlag = m_ClipBoard(f_SourceItem).Contents(lv_nIndex).ModifyFlag
    Next
    
End Sub

Public Sub ClearClipboardItemOnFlow(ByVal f_PID As Byte)
On Error Resume Next

    Dim lv_clipindex As Byte
    
    For lv_clipindex = 0 To m_bytMaxStacks
        If m_ClipBoard(lv_clipindex).Contents(0).IntP_id = f_PID Then
            m_ClipBoard(lv_clipindex).Operation = DEF_OPERATION_INIT
        End If
    Next
    
End Sub
