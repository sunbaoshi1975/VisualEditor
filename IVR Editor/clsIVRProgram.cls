VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIVRProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Private variable
Private m_ctlNode() As New CNode            '' 1 Based

'' Properties
Private m_objRootWnd As Object              '' 菜单窗口
Private m_objParent As Object               '' 母窗口
Private m_objContainer As Object            '' 容器
Private m_blnSavedMark As Boolean           '' 是否已经保存

Private m_nFlowID As Byte                   '' 当前流程项目ID
Private m_strFlowName As String             '' 当前流程项目名称

Private m_nResourceID As Integer            '' 当前资源项目ID
Private m_strResourceName As String         '' 当前资源项目名称
Private m_strResDescription As String       '' 当前资源描述
Private m_strResAuther As String            '' 当前资源作者
Private m_strResUser As String              '' 当前资源用户
Private m_strResVersion As String           '' 当前资源版本号
Private m_tmResCreateTime As Date           '' 当前资源创建时间
Private m_tmResModifyTime As Date           '' 当前资源修改时间

Private m_nLanguageID As Integer            '' 当前服务语言ID
Private m_nSysNodeID As Integer             '' 最大系统节点ID
Private m_nUserNodeID As Integer            '' 最大用户节点ID
Private m_nNodeID As Integer                '' 新增节点编号
Private m_nNodeSelectedID As Integer        '' 当前选中的节点Index
Private m_nVarNodeCount As Integer          '' 变量节点数量
Private m_nRootNode As Integer              '' Root Node ID

Private m_nCurrentPage As Integer           '' 当前Page
Private m_nPageCount As Integer             '' Page Count

'' Class Variables
Private m_strConString As String            '' 连接串
Private m_adoCN As ADODB.Connection         '' 连接
Private m_rsFS As ADODB.Recordset           '' 流程记录
Private m_rsRS As ADODB.Recordset           '' 节点记录
Private m_rsResource As ADODB.Recordset     '' 资源列表

'Michael Add @ Jul,12,07
Private m_rsProject As ADODB.Recordset      '' 项目列表
Private m_rsFlashRS As ADODB.Recordset
Public m_conResource As ADODB.Connection

Private Sub Class_Initialize()
On Error GoTo BackDoor

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    '' 变量初始化
    InitLocalVariables
    
    m_strConString = gSystem.strConString
    
    '' Open Connection
    Set m_adoCN = New ADODB.Connection
    m_adoCN.ConnectionString = m_strConString
    m_adoCN.CursorLocation = adUseClient
    m_adoCN.IsolationLevel = adXactReadCommitted
    m_adoCN.Mode = adModeShareDenyNone + adModeReadWrite
    m_adoCN.Open
    
    Set m_rsFS = New ADODB.Recordset
    Set m_rsRS = New ADODB.Recordset
    Set m_rsResource = New ADODB.Recordset
    'Michael Add @ Jul,12,07
    Set m_rsProject = New ADODB.Recordset
    Set m_rsFlashRS = New ADODB.Recordset
    'Mike added @ 2008-7-7
    WriteLogMessage 0, enu_Information, "Connect to Database successful! ", "ConnectionString:" & m_strConString
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    Exit Sub
    
BackDoor:

    'Message "E016"
    'Mike @ 2008-7-7
    WriteLogMessage Err.Number, enu_Error, "Connect to Database failed! ConnectionString:" & m_strConString, Err.Description
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
    On Error GoTo 0

End Sub

Private Sub Class_Terminate()
On Error GoTo BackDoor

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    '' Close Recordsets
    If Not m_rsFS Is Nothing Then
        If m_rsFS.State = ADODB.adStateOpen Then
            m_rsFS.Close
            Set m_rsFS = Nothing
        End If
    End If
    If Not m_rsRS Is Nothing Then
        If m_rsRS.State = ADODB.adStateOpen Then
            m_rsRS.Close
            Set m_rsRS = Nothing
        End If
    End If
    If Not m_rsResource Is Nothing Then
        If m_rsResource.State = ADODB.adStateOpen Then
            m_rsResource.Close
            Set m_rsResource = Nothing
        End If
    End If
    'Michael Add @ Jul,12,07
    If Not m_rsProject Is Nothing Then
        If m_rsProject.State = ADODB.adStateOpen Then
            m_rsProject.Close
            Set m_rsProject = Nothing
        End If
    End If
    
    If Not m_rsFlashRS Is Nothing Then
        If m_rsFlashRS.State = ADODB.adStateOpen Then
            m_rsFlashRS.Close
            Set m_rsFlashRS = Nothing
        End If
    End If
    
    '' Close CN
    If Not m_adoCN Is Nothing Then
        If m_adoCN.State = ADODB.adStateOpen Then
            m_adoCN.Close
            Set m_adoCN = Nothing
        End If
    End If
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    Exit Sub

BackDoor:

    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
    Message "E899"
    If Err.Description <> "" Then
        MsgBox Err.Description
    End If
    
    On Error GoTo 0
    
End Sub

Private Sub InitLocalVariables()
    
    Set m_objRootWnd = Nothing
    Set m_objParent = Nothing
    Set m_objContainer = Nothing
    
    m_nFlowID = 0
    m_nResourceID = 0
    m_nLanguageID = 0
    m_nSysNodeID = 0
    m_nUserNodeID = 0
    m_nNodeID = 0
    m_nNodeSelectedID = 0
    m_nVarNodeCount = 0
    m_nCurrentPage = 1
    m_nPageCount = 0
    m_nRootNode = 255
    
    m_strFlowName = ""
    
    m_strResourceName = ""
    m_strResDescription = ""
    m_strResAuther = gSystem.strOSUser
    m_strResUser = ""
    m_strResVersion = ""
    m_tmResCreateTime = Now
    m_tmResModifyTime = Now
    
    m_blnSavedMark = True
    
End Sub

Public Property Get RS_Resource() As ADODB.Recordset
On Error Resume Next
    
    Set RS_Resource = m_rsResource
    
End Property

'Michael Added @ Jul,12,07
Public Property Get RS_Project() As ADODB.Recordset
On Error Resume Next
    Set RS_Project = m_rsProject
End Property
Public Property Get RS_FlashRS() As ADODB.Recordset
On Error Resume Next
    Set RS_FlashRS = m_rsFlashRS
End Property
'Add End

Public Property Let SavedMark(f_bSaved As Boolean)
On Error Resume Next
  
    m_blnSavedMark = f_bSaved
        
End Property

Public Property Get SavedMark() As Boolean
On Error Resume Next

    SavedMark = m_blnSavedMark
        
End Property

Public Property Get CallFlowID() As Byte
On Error Resume Next

    CallFlowID = m_nFlowID
        
End Property

Public Property Let CallFlowID(f_bytCallFlowID As Byte)
On Error Resume Next

    m_nFlowID = f_bytCallFlowID
        
End Property

Public Property Get CallFlowName() As String
On Error Resume Next

    CallFlowName = m_strFlowName
        
End Property

Public Property Get Description() As String
On Error Resume Next
    
    If m_rsFS.RecordCount > 0 Then
        Description = m_rsFS.Fields("P_Description")
    End If
        
End Property

Public Property Get CallFlowAuther() As String
On Error Resume Next

    If m_rsFS.RecordCount > 0 Then
        CallFlowAuther = m_rsFS.Fields("P_Auther")
    Else
        CallFlowAuther = gSystem.strOSUser
    End If
            
End Property

Public Property Get CallFlowUser() As String
On Error Resume Next

    If m_rsFS.RecordCount > 0 Then
        CallFlowUser = m_rsFS.Fields("P_User")
    End If
        
End Property

Public Property Get CallFlowVersion() As String
On Error Resume Next

    If m_rsFS.RecordCount > 0 Then
        CallFlowVersion = m_rsFS.Fields("P_version")
    End If
        
End Property

Public Property Get CreateTime() As Date
On Error Resume Next

    If m_rsFS.RecordCount > 0 Then
        CreateTime = m_rsFS.Fields("P_CreateTime")
    Else
        CreateTime = Now
    End If
        
End Property

Public Property Get ModifyTime() As Date
On Error Resume Next

    If m_rsFS.RecordCount > 0 Then
        ModifyTime = m_rsFS.Fields("P_ModifyTime")
    Else
        ModifyTime = Now
    End If
        
End Property

'Michael Add @ Jul,12,07
Public Property Get CallFlowType() As String
On Error Resume Next
    If m_rsFS.RecordCount > 0 Then
        CallFlowType = m_rsFS.Fields("P_Type")
    End If
End Property

Public Property Get ResourceID() As Integer
On Error Resume Next

    ResourceID = m_nResourceID
        
End Property

Public Property Get ResourceName() As String
On Error Resume Next

    ResourceName = m_strResourceName
        
End Property

Public Property Get LanguageID() As Integer
On Error Resume Next

    LanguageID = m_nLanguageID
        
End Property

Public Property Let SysNodeID(f_nID As Integer)
On Error Resume Next
  
    m_nSysNodeID = f_nID
        
End Property

Public Property Get SysNodeID() As Integer
On Error Resume Next

    SysNodeID = m_nSysNodeID
        
End Property

Public Property Let UserNodeID(f_nID As Integer)
On Error Resume Next
  
    m_nUserNodeID = f_nID
        
End Property

Public Property Get UserNodeID() As Integer
On Error Resume Next

    UserNodeID = m_nUserNodeID
        
End Property

Public Property Let NewNodeID(f_nID As Integer)
On Error Resume Next
  
    m_nNodeID = f_nID
        
End Property

Public Property Get NewNodeID() As Integer
On Error Resume Next

    NewNodeID = m_nNodeID

End Property

Public Property Let VarNodeCount(f_nCount As Integer)
On Error Resume Next

    m_nVarNodeCount = f_nCount

End Property

Public Property Get VarNodeCount() As Integer
On Error Resume Next

    VarNodeCount = m_nVarNodeCount

End Property

Public Property Let NodeSelectedID(f_nID As Integer)
On Error Resume Next

    m_nNodeSelectedID = f_nID
    If m_nNodeSelectedID = 0 Then
        m_objParent.ShowHandles False
        SetAllNodeSelect False
    End If

End Property

Public Property Get NodeSelectedID() As Integer
On Error Resume Next

    NodeSelectedID = m_nNodeSelectedID
    
End Property

Public Property Let RootNodedID(f_nID As Integer)
On Error Resume Next

    m_nRootNode = f_nID

End Property

Public Property Get RootNodedID() As Integer
On Error Resume Next

    RootNodedID = m_nRootNode

End Property

Public Property Let CurrentPage(f_nPageNo As Integer)
On Error Resume Next

    If m_nCurrentPage <> f_nPageNo Then
        m_nCurrentPage = f_nPageNo
        GotoThisPage f_nPageNo
    End If

End Property

Public Property Get CurrentPage() As Integer
On Error Resume Next

    CurrentPage = m_nCurrentPage

End Property

Public Property Let PageCount(f_nPageCount As Integer)
On Error Resume Next

    m_nPageCount = f_nPageCount

End Property

Public Property Get PageCount() As Integer
On Error Resume Next

    PageCount = m_nPageCount

End Property

Public Function GetActualPageCount() As Integer
On Error Resume Next
    
    Dim lv_nPageCount As Integer
    Dim lv_loop As Integer
    
    lv_nPageCount = 0
    For lv_loop = 1 To m_nNodeID
        If m_ctlNode(lv_loop).InPage > lv_nPageCount Then lv_nPageCount = m_ctlNode(lv_loop).InPage
    Next

    GetActualPageCount = lv_nPageCount
End Function

Public Property Get Node(ByVal f_nIndex As Integer) As CNode
On Error Resume Next
    
    Set Node = m_ctlNode(f_nIndex)
    
End Property

' 在画布上显示流程
'
' Mike Modified @ 2008-1-30 ,add f_blnShowTag
'Public Sub ShowCallFlowOnScreen(objRootWnd As Object, objParent As Object, objContainer As Object, ByVal f_blnShowCaption As Boolean, Optional ByVal f_blnShowSysNode As Boolean = False)
Public Sub ShowCallFlowOnScreen(objRootWnd As Object, _
                                objParent As Object, _
                                objContainer As Object, _
                                ByVal f_blnShowCaption As Boolean, _
                                ByVal f_blnShowTag As Boolean, _
                                Optional ByVal f_blnShowSysNode As Boolean = False)
'On Error Resume Next
On Error GoTo BackDoor

    Dim lv_Gdata1(DEF_NODE_DATA1_LEN) As Byte
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_loop As Integer
    Dim lv_SubLoop As Integer
    
    ReDim Preserve m_ctlNode(m_nNodeID) As New CNode
    
    Set m_objRootWnd = objRootWnd
    Set m_objParent = objParent
    Set m_objContainer = objContainer
    
    '' Init Var
    m_nCurrentPage = 1
    m_nPageCount = 0
    
    '' First Row
    m_rsRS.MoveFirst
    
    For lv_loop = 1 To m_nNodeID
    
        If m_ctlNode(lv_loop).Create(objRootWnd, objParent, objContainer, lv_loop) Then
            
            m_ctlNode(lv_loop).FlowID = CallFlowID
            m_ctlNode(lv_loop).NodeID = m_rsRS.Fields("N_ID")
            m_ctlNode(lv_loop).NodeNo = m_rsRS.Fields("N_NO")
            m_ctlNode(lv_loop).InPage = m_rsRS.Fields("N_Page")
            m_ctlNode(lv_loop).Height = m_rsRS.Fields("N_Height")
            m_ctlNode(lv_loop).Width = m_rsRS.Fields("N_Width")
            m_ctlNode(lv_loop).Top = m_rsRS.Fields("N_Top")
            m_ctlNode(lv_loop).Left = m_rsRS.Fields("N_Left")
            'Mike add this Node Tag @ 2008-1-30
            m_ctlNode(lv_loop).NodeTag = m_rsRS.Fields("N_Tag")
            
            If m_ctlNode(lv_loop).InPage > m_nPageCount Then m_nPageCount = m_ctlNode(lv_loop).InPage
            
            If Not IsNull(m_rsRS.Fields("N_Description")) Then
                m_ctlNode(lv_loop).Description = Trim(m_rsRS.Fields("N_Description"))
            End If
          
#If Programnew = 0 Then
            
            m_ctlNode(lv_loop).Data1 = CharToStr(m_rsRS, DEF_NODE_DATA1_LEN)
            m_ctlNode(lv_loop).Data2 = CharToStr(m_rsRS, DEF_NODE_DATA2_LEN)

#Else
                
            If Not IsNull(m_rsRS.Fields("n_data1")) Then
                m_ctlNode(lv_loop).Data1 = LeftB(m_rsRS.Fields("n_data1"), DEF_NODE_DATA1_LEN)
            End If
            If Not IsNull(m_rsRS.Fields("n_data2")) Then
                m_ctlNode(lv_loop).Data2 = LeftB(m_rsRS.Fields("n_data2"), DEF_NODE_DATA2_LEN)
            End If

#End If
            
            Select Case m_ctlNode(lv_loop).NodeNo
            Case 0
                If m_ctlNode(lv_loop).Data2 <> "" Then
                        
                    For lv_SubLoop = 0 To 12
                        lv_Gdata1(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_loop).Data1, lv_SubLoop + 1, 1))
                    Next
                    CopyMemory Node0_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                    
                    For lv_SubLoop = 0 To 63
                        lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_loop).Data2, lv_SubLoop + 1, 1))
                    Next
                        
                    CopyMemory Node0_Data2.key_repeat, lv_Gdata2(0), DEF_NODE_DATA2_LEN
                    
                    '' Resource Project ID
                    Call OpenResourceRecordSet(Node0_Data2.ResourceProject, 0)
                    
                    '' Root Node
                    m_nRootNode = Node0_Data2.nd_root
                        
                End If
                
            'mike added 2008-8-20
            'init uservars of NODE1_Data1
            Case 1
                If m_ctlNode(lv_loop).Data1 <> "" Then
                    For lv_SubLoop = 0 To 12
                        lv_Gdata1(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_loop).Data1, lv_SubLoop + 1, 1))
                    Next
                    CopyMemory Node1_Data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                End If
                    
            Case 2
                m_nVarNodeCount = m_nVarNodeCount + 1
            
            Case 255
                Call m_ctlNode(lv_loop).AddLine
                m_ctlNode(lv_loop).MoveImageWithLine
            End Select
            
            '' Sun modified 2002-03-30 for multipage
            m_ctlNode(lv_loop).VerifyVisible m_nCurrentPage
            
            '' Sun added 2002-12-04
            If m_nRootNode > 255 Then
                If m_ctlNode(lv_loop).NodeID = m_nRootNode Then
                    m_ctlNode(lv_loop).IsRootNode = True
                End If
            End If
            
            '' Sun added 2002-03-30
            m_ctlNode(lv_loop).NodeCaptionVisible = f_blnShowCaption
            
            '' Mike Added @ 2008-1-30 show/hide node tag on node image
            m_ctlNode(lv_loop).NodeTagVisible = f_blnShowTag
            
            m_ctlNode(lv_loop).ModifyFlag = DEF_OPERATION_INIT
            
            '' Next
            m_rsRS.MoveNext
              
        Else
            Exit Sub
        End If
    
    Next
    
Exit Sub
    
BackDoor:
    'Debug.Print "ShowCallFlowOnScreen(" & Str(m_nFlowID) & ") got error: " & Err.Description & ", on loop: " & Str(lv_loop)
    WriteLogMessage Err.Number, enu_Error, "Show Call flow on Screen got error, Callflow ID:" & Str(m_nFlowID), Err.Description
    Err.Clear
    Resume Next
End Sub

Public Sub SetWorkPageScale(ByVal f_nLeft As Long, ByVal f_nTop As Long, ByVal f_nWidth As Long, ByVal f_nHeight As Long)
    If Not m_objContainer Is Nothing Then
        m_objContainer.Move 0, 0, gPageWidth, gPageHeight
    End If
End Sub

Public Sub SetWorkPageBackColor(ByVal f_Color As OLE_COLOR)
    If Not m_objContainer Is Nothing Then
        m_objContainer.BackColor = f_Color
    End If
End Sub

Public Sub ClearWorkPage()
    If Not m_objContainer Is Nothing Then
        m_objContainer.Cls
    End If
    If Not m_objParent Is Nothing Then
        m_objParent.ShowHandles False
    End If
End Sub

' 选中的节点数量
'
Public Function SelectedCount() As Integer
On Error Resume Next

    Dim lv_Index As Integer
    Dim lv_nCount As Integer

    lv_nCount = 0
    For lv_Index = 1 To m_nNodeID
        If m_ctlNode(lv_Index).IsSelected Then
            lv_nCount = lv_nCount + 1
        End If
    Next
    
    SelectedCount = lv_nCount
    
End Function

''' Search Node with NodeID, Return Index
Public Function SearchNodeIndexWithID(f_NodeID As Integer) As Integer
On Error Resume Next

    Dim lv_Index As Integer
    
    For lv_Index = 1 To m_nNodeID
        If m_ctlNode(lv_Index).NodeID = f_NodeID And m_ctlNode(lv_Index).NodeNo <> 0 Then
            SearchNodeIndexWithID = lv_Index
            Exit Function
        End If
    Next
    
    SearchNodeIndexWithID = -1
    
End Function

' 增加用户节点
'
Public Sub AddUserNode()
On Error Resume Next

    m_nNodeID = m_nNodeID + 1
    m_nUserNodeID = m_nUserNodeID + 1
    
    ReDim Preserve m_ctlNode(m_nNodeID) As New CNode
    
    m_ctlNode(m_nNodeID).NodeID = m_nUserNodeID
    SavedMark = False
    
End Sub

' 增加变量节点
'
Public Sub AddVariableNode()
On Error Resume Next

    Dim lv_nPrevIndex As Integer
    
    m_nNodeID = m_nNodeID + 1
    m_nVarNodeCount = m_nVarNodeCount + 1
    m_nSysNodeID = m_nSysNodeID + 1
    
    ReDim Preserve m_ctlNode(m_nNodeID) As New CNode
    m_nNodeSelectedID = m_nSysNodeID
    lv_nPrevIndex = SearchNodeIndexWithID(m_nSysNodeID - 1)
    
    '变量1
    '变量长度
    Node2_Data2.uservar(0) = 0
    '变量数据类型
    Node2_Data2.uservar(1) = 2
    Node2_Data2.uservar(2) = Asc("0")
    
    '变量2
    '变量长度
    Node2_Data2.uservar(16) = 0
    '变量数据类型
    Node2_Data2.uservar(17) = 2
    Node2_Data2.uservar(18) = Asc("0")
    
    
    '变量3
    '变量长度
    Node2_Data2.uservar(32) = 0
    '变量数据类型
    Node2_Data2.uservar(33) = 2
    Node2_Data2.uservar(34) = Asc("0")
    
    '变量4
    '变量长度
    Node2_Data2.uservar(48) = 0
    '变量数据类型
    Node2_Data2.uservar(49) = 2
    Node2_Data2.uservar(50) = Asc("0")
    
    If m_ctlNode(m_nNodeID).Create(m_objRootWnd, m_objParent, m_objContainer, m_nNodeID) Then
    
        m_ctlNode(m_nNodeID).FlowID = m_nFlowID
        m_ctlNode(m_nNodeID).NodeID = m_nSysNodeID
        m_ctlNode(m_nNodeID).NodeNo = 2
        m_ctlNode(m_nNodeID).InPage = 1
        m_ctlNode(m_nNodeID).Left = 320 + m_ctlNode(lv_nPrevIndex).Left
        m_ctlNode(m_nNodeID).Top = 320 + m_ctlNode(lv_nPrevIndex).Top
        m_ctlNode(m_nNodeID).Height = Screen.TwipsPerPixelY * 32          '' 480
        m_ctlNode(m_nNodeID).Width = Screen.TwipsPerPixelX * 32           '' 480
        m_ctlNode(m_nNodeID).Description = "N/A"
        m_ctlNode(m_nNodeID).VerifyVisible m_nCurrentPage
       
        '' Sun added 2002-03-30
        m_ctlNode(m_nNodeID).NodeCaptionVisible = gblnShowNodeCaption
        
        F_NodeData m_nNodeID, 2
        AddNewIvrRecord m_nNodeID
        
    End If
    
    SavedMark = False

End Sub

' 删除选中节点
'
Public Sub DeleteSelectdNode()
    
    If NodeSelectedID > 0 Then
        DeleteNodeWithIndex NodeSelectedID
    End If
    SavedMark = False
   
End Sub

' 根据节点ID删除一个节点
'
Public Sub DeleteNode(f_NodeID As Integer)
On Error Resume Next
    
    Dim lv_nIndexID As Integer
    
    lv_nIndexID = SearchNodeIndexWithID(f_NodeID)
    If lv_nIndexID > 0 Then
        DeleteNodeWithIndex lv_nIndexID
        SavedMark = False
    End If

End Sub

' 根据节点Index删除一个节点
'
Public Sub DeleteNodeWithIndex(f_nIndex As Integer)
On Error Resume Next
    
    If f_nIndex > 0 Then
        DeleteIvrRecord f_nIndex
        m_ctlNode(f_nIndex).Destroy
        
        If m_ctlNode(f_nIndex).NodeID < 256 Then
            If m_ctlNode(f_nIndex).NodeID > 2 Then
                m_nVarNodeCount = m_nVarNodeCount - 1
            End If
            m_nSysNodeID = m_nSysNodeID - 1
        End If
        
    End If
    'Mike added @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Delete Node Successful! ", "Node Index:" & f_nIndex)
    SavedMark = False
    
End Sub

' 重建节点
'
Public Function CreateNode(f_nIndex As Integer) As Boolean
On Error Resume Next

    CreateNode = m_ctlNode(f_nIndex).Create(m_objRootWnd, m_objParent, m_objContainer, f_nIndex)
    
End Function

Private Sub GotoThisPage(ByVal f_nPageNo As Integer)
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    Dim lv_Index As Integer
    
    For lv_Index = 1 To m_nNodeID
        m_ctlNode(lv_Index).VerifyVisible f_nPageNo
    Next
            
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
End Sub

' 设置系统节点的显示状态
'
Public Sub SetAllSysNodesVisible()
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    Dim lv_Index As Integer
    
    For lv_Index = 1 To m_nNodeID
        ''Mike modified 2008-8-20
        'm_ctlNode(lv_Index).VerifyVisible m_nCurrentPage
        If m_ctlNode(lv_Index).VerifyVisible(m_nCurrentPage) Then
            m_ctlNode(lv_Index).NodeCaptionVisible = True
        End If
    Next
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
End Sub

' 设置所有节点的标签显示状态
'
Public Sub SetAllNodeCaptionVisible(ByVal f_Visible As Boolean)
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    Dim lv_Index As Integer
    
    For lv_Index = 1 To m_nNodeID
        m_ctlNode(lv_Index).NodeCaptionVisible = f_Visible
    Next
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
End Sub

'**************************************************
'Mike Added this sub @ 2008-1-29
' 设置所有节点的标签显示状态
Public Sub SetAllNodeTagVisible(ByVal f_Visible As Boolean)
On Error Resume Next
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    Dim lv_Index As Integer
    
    For lv_Index = 1 To m_nNodeID
        m_ctlNode(lv_Index).NodeTagVisible = f_Visible
    Next
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
End Sub
'***************************************************

' 设置一个指定区域内的节点选中状态
'
Public Sub SetRegionNodeSelect(f_Rect As CRect, ByVal f_Selected As Boolean)
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    Dim lv_Index As Integer
    
    For lv_Index = 1 To m_nNodeID
        If f_Rect.PtInRect(m_ctlNode(lv_Index).Left, m_ctlNode(lv_Index).Top) And _
              m_ctlNode(lv_Index).InPage = CurrentPage Then
            m_ctlNode(lv_Index).IsSelected = f_Selected
        End If
    Next
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
End Sub

' 设置某页内的节点选中状态
'
Public Sub SetPageNodeSelect(ByVal f_PageNo As Integer, ByVal f_Selected As Boolean)
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    Dim lv_Index As Integer
    
    For lv_Index = 1 To m_nNodeID
        If m_ctlNode(lv_Index).InPage = f_PageNo Then
            m_ctlNode(lv_Index).IsSelected = f_Selected
        End If
    Next
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
End Sub

Public Sub SetAllNodeSelect(ByVal f_Selected As Boolean)
On Error Resume Next

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    Dim lv_Index As Integer
    
    For lv_Index = 1 To m_nNodeID
        m_ctlNode(lv_Index).IsSelected = f_Selected
    Next
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
End Sub

Public Function GetSelectItemRect(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
On Error Resume Next

    Dim lv_Index As Integer
    
    GetSelectItemRect = False
    For lv_Index = 1 To m_nNodeID
    
        If m_ctlNode(lv_Index).IsSelected And m_ctlNode(lv_Index).InPage = CurrentPage Then
            x1 = m_ctlNode(lv_Index).Left
            y1 = m_ctlNode(lv_Index).Top
            x2 = m_ctlNode(lv_Index).Left + m_ctlNode(lv_Index).Width
            y2 = m_ctlNode(lv_Index).Top + m_ctlNode(lv_Index).Height
            GetSelectItemRect = True
            lv_Index = lv_Index + 1
            Exit For
        End If
    Next
    
    Do While lv_Index <= m_nNodeID
        If m_ctlNode(lv_Index).IsSelected And m_ctlNode(lv_Index).InPage = CurrentPage Then
            If x1 > m_ctlNode(lv_Index).Left Then
                x1 = m_ctlNode(lv_Index).Left
            End If
            If y1 > m_ctlNode(lv_Index).Top Then
                y1 = m_ctlNode(lv_Index).Top
            End If
            If x2 < m_ctlNode(lv_Index).Left + m_ctlNode(lv_Index).Width Then
                x2 = m_ctlNode(lv_Index).Left + m_ctlNode(lv_Index).Width
            End If
            If y2 < m_ctlNode(lv_Index).Top + m_ctlNode(lv_Index).Height Then
                y2 = m_ctlNode(lv_Index).Top + m_ctlNode(lv_Index).Height
            End If
        End If
        lv_Index = lv_Index + 1
    Loop

End Function

' 销毁所有节点
'
Public Sub DestroyAllNodes()
On Error Resume Next

    Dim lv_Index As Integer

    NodeSelectedID = 0
    For lv_Index = 1 To m_nNodeID
        m_ctlNode(lv_Index).Destroy
    Next
    
    ReDim m_ctlNode(0) As New CNode
    
    gClipBoard.ClearClipboardItemOnFlow m_nFlowID
    
    '' 变量初始化
    InitLocalVariables
    
    '' Close Recordsets
    m_rsRS.Close
    m_rsFS.Close
    m_rsResource.Close
    
End Sub

' Sun added 2007-03-25
' 流程另存
Public Function SaveCallFlowAs(ByVal f_bytNewPID As Byte) As Boolean
On Error GoTo BackDoor

    Dim lv_CN As ADODB.Connection        '' 连接
    Dim lv_RS As ADODB.Recordset
    Dim lv_SQL As String
    Dim lv_InTrans As Boolean
    Dim lv_loop As Integer
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    SaveCallFlowAs = False
    lv_InTrans = False

    Set lv_CN = New ADODB.Connection
    lv_CN.ConnectionString = gSystem.strConString
    lv_CN.CursorLocation = adUseClient
    lv_CN.Open

    '' Begin Transaction
    lv_CN.BeginTrans
    lv_InTrans = True
    
    Set lv_RS = New ADODB.Recordset
    lv_RS.CursorLocation = adUseClient
    lv_RS.CursorType = adOpenKeyset
    lv_RS.LockType = adLockOptimistic
    
    ''------------------------------------------------
    '' 另存流程清单
    ''------------------------------------------------
    '' Open Recordsets
    lv_SQL = "Select * from tbRefIVR Where P_ID = " & Str(f_bytNewPID) & " AND P_TYPE = 'P'"
    lv_RS.Open lv_SQL, lv_CN
    
    '' Insert New Ref IVR Data from Current Recordset
    If Not (m_rsFS.EOF And m_rsFS.BOF) Then
        
        lv_RS.AddNew
        
        For lv_loop = 0 To m_rsFS.Fields.Count - 1
            If UCase(lv_RS.Fields(lv_loop).name) = "P_ID" Then
                lv_RS.Fields(lv_loop) = f_bytNewPID
            Else
                lv_RS.Fields(lv_loop) = m_rsFS.Fields(lv_loop)
            End If
        Next
        
        
        lv_RS.Update
    
    End If
    
    '' Close Recordsets
    lv_RS.Close
    
    ''------------------------------------------------
    '' 另存流程数据
    ''------------------------------------------------
    #If Programnew = 0 Then
        lv_SQL = "Delete from tbIVRProgramnew Where P_ID = " & Str(f_bytNewPID)
    #Else
        lv_SQL = "Delete from tbIVRProgram Where P_ID = " & Str(f_bytNewPID)
    #End If
    
    '' Delete Old Data
    lv_CN.Execute lv_SQL, -1, adCmdText

    '' Open Recordsets
    #If Programnew = 0 Then
        lv_SQL = "Select * from tbIVRProgramnew Where P_ID = " & Str(f_bytNewPID) & " ORDER BY N_ID, N_NO"
    #Else
        lv_SQL = "Select * from tbIVRProgram Where P_ID = " & Str(f_bytNewPID) & " ORDER BY N_ID, N_NO"
    #End If
    lv_RS.Open lv_SQL, lv_CN
        
    '' Insert New Data
    If Not (m_rsRS.EOF And m_rsRS.BOF) Then
    
        m_rsRS.MoveFirst
        While Not m_rsRS.EOF
            
            lv_RS.AddNew
            
            For lv_loop = 0 To m_rsRS.Fields.Count - 1
                If UCase(lv_RS.Fields(lv_loop).name) = "P_ID" Then
                    lv_RS.Fields(lv_loop) = f_bytNewPID
                Else
                    lv_RS.Fields(lv_loop) = m_rsRS.Fields(lv_loop)
                End If
            Next
            
            lv_RS.Update
            
            m_rsRS.MoveNext
            
        Wend
        
    End If
    
    '' Close Recordsets
    lv_RS.Close
    
    '' Commit Transaction
    lv_CN.CommitTrans
    lv_InTrans = False
    SaveCallFlowAs = True
    
    'Mike Added @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Successful save CallFlow as Callflow:" & f_bytNewPID)
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    Exit Function
    
BackDoor:
    
    ' Roll back if we are in transaction
    If lv_InTrans Then lv_CN.RollbackTrans
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
    Message "E899"
    'Mike added @ 2008-7-7
    Call WriteLogMessage(Err.Number, enu_Error, "Failed to save CallFlow as " & f_bytNewPID & " Error of updating data in server!", Err.Description)
    
    If Err.Description <> "" Then
        MsgBox Err.Description
    End If
    
    On Error GoTo 0
End Function

' Sun added 2005-08-01
' 同步呼叫流程
Public Function SynchronizeCallFlow(ByVal f_strConnStr As String) As Boolean
On Error GoTo BackDoor

    Dim lv_CN As ADODB.Connection        '' 连接
    Dim lv_RS As ADODB.Recordset
    Dim lv_SQL As String
    Dim lv_InTrans As Boolean
    Dim lv_loop As Integer
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    SynchronizeCallFlow = False
    lv_InTrans = False

    Set lv_CN = New ADODB.Connection
    lv_CN.ConnectionString = f_strConnStr
    lv_CN.CursorLocation = adUseClient
    lv_CN.Open

    '' Begin Transaction
    lv_CN.BeginTrans
    lv_InTrans = True
    
    Set lv_RS = New ADODB.Recordset
    lv_RS.CursorLocation = adUseClient
    lv_RS.CursorType = adOpenKeyset
    lv_RS.LockType = adLockOptimistic
    
    ''------------------------------------------------
    '' 同步流程清单
    ''------------------------------------------------
    '' Delete Ref. IVR
    lv_SQL = "Delete from tbRefIVR Where P_ID = " & Str(CallFlowID) & " AND P_TYPE = 'P'"
    lv_CN.Execute lv_SQL, -1, adCmdText
    
    '' Open Recordsets
    lv_SQL = "Select * from tbRefIVR Where P_ID = " & Str(CallFlowID) & " AND P_TYPE = 'P'"
    lv_RS.Open lv_SQL, lv_CN
    
    '' Insert New Ref IVR Data from Current Recordset
    If Not (m_rsFS.EOF And m_rsFS.BOF) Then
        
        lv_RS.AddNew
        
        For lv_loop = 0 To m_rsFS.Fields.Count - 1
            lv_RS.Fields(lv_loop) = m_rsFS.Fields(lv_loop)
        Next
        
        lv_RS.Update
    
    End If
    
    '' Close Recordsets
    lv_RS.Close
        
    ''------------------------------------------------
    '' 同步流程数据
    ''------------------------------------------------
    #If Programnew = 0 Then
        lv_SQL = "Delete from tbIVRProgramnew Where P_ID = " & Str(CallFlowID)
    #Else
        lv_SQL = "Delete from tbIVRProgram Where P_ID = " & Str(CallFlowID)
    #End If
    
    '' Delete Old Data
    lv_CN.Execute lv_SQL, -1, adCmdText

    '' Open Recordsets
    #If Programnew = 0 Then
        lv_SQL = "Select * from tbIVRProgramnew Where P_ID = " & Str(CallFlowID) & " ORDER BY N_ID, N_NO"
    #Else
        lv_SQL = "Select * from tbIVRProgram Where P_ID = " & Str(CallFlowID) & " ORDER BY N_ID, N_NO"
    #End If
    lv_RS.Open lv_SQL, lv_CN
        
    '' Insert New Data
    If Not (m_rsRS.EOF And m_rsRS.BOF) Then
    
        m_rsRS.MoveFirst
        While Not m_rsRS.EOF
            
            lv_RS.AddNew
            
            For lv_loop = 0 To m_rsRS.Fields.Count - 1
                lv_RS.Fields(lv_loop) = m_rsRS.Fields(lv_loop)
            Next
            
            lv_RS.Update
            
            m_rsRS.MoveNext
            
        Wend
        
    End If
    
    '' Close Recordsets
    lv_RS.Close
    
    '' Commit Transaction
    lv_CN.CommitTrans
    lv_InTrans = False
    SynchronizeCallFlow = True
    
    'Mike @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Successful Synchronized CallFlow! CallFlow ID:" & CallFlowID)
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    Exit Function
    
BackDoor:
    
    ' Roll back if we are in transaction
    If lv_InTrans Then lv_CN.RollbackTrans
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
    Message "E899"
    If Err.Description <> "" Then
        MsgBox Err.Description
    End If
    
    'Mike Added @ 2008-7-7
    Call WriteLogMessage(Err.Number, enu_Warnning, "Failed to Synchronize callflows! CallFlow ID:" & CallFlowID, Err.Description)
    
    On Error GoTo 0
End Function

' 同步资源
'
Public Function SynchronizeResource(ByVal f_strConnStr As String) As Boolean
On Error GoTo BackDoor

    Dim lv_CN As ADODB.Connection        '' 连接
    Dim lv_RS As ADODB.Recordset
    Dim lv_SourceRS As ADODB.Recordset
    Dim lv_SQL As String
    Dim lv_InTrans As Boolean
    Dim lv_loop As Integer
    
    '改变鼠标指针形状->沙漏光标
    ChangeMousePointer vbHourglass, True

    SynchronizeResource = False
    lv_InTrans = False

    Set lv_CN = New ADODB.Connection
    lv_CN.ConnectionString = f_strConnStr
    lv_CN.CursorLocation = adUseClient
    lv_CN.Open

    '' Begin Transaction
    lv_CN.BeginTrans
    lv_InTrans = True
    
    Set lv_RS = New ADODB.Recordset
    Set lv_SourceRS = New ADODB.Recordset
    lv_RS.CursorLocation = adUseClient
    lv_RS.CursorType = adOpenKeyset
    lv_RS.LockType = adLockOptimistic
    
    ''------------------------------------------------
    '' 同步资源清单
    ''------------------------------------------------
    '' Delete Ref. IVR
    lv_SQL = "Delete from tbRefIVR Where P_ID = " & Str(ResourceID) & " AND P_TYPE = 'R'"
    lv_CN.Execute lv_SQL, -1, adCmdText
    
    '' Open Recordsets
    lv_SQL = "Select * from tbRefIVR Where P_ID = " & Str(ResourceID) & " AND P_TYPE = 'R'"
    lv_RS.Open lv_SQL, lv_CN
    lv_SourceRS.Open lv_SQL, gSystem.strConString, adOpenForwardOnly, adLockReadOnly
    
    '' Insert New Ref IVR Data from Current Recordset
    If Not (lv_SourceRS.EOF And lv_SourceRS.BOF) Then
        
        lv_RS.AddNew
        
        For lv_loop = 0 To lv_SourceRS.Fields.Count - 1
            lv_RS.Fields(lv_loop) = lv_SourceRS.Fields(lv_loop)
        Next
        
        lv_RS.Update
    
    End If
    
    '' Close Recordsets
    lv_SourceRS.Close
    lv_RS.Close
        
    ''------------------------------------------------
    '' 同步资源数据
    ''------------------------------------------------
    lv_SQL = "Delete from tbResource Where P_ID = " & Str(ResourceID)
    
    '' Delete Old Data
    lv_CN.Execute lv_SQL, -1, adCmdText

    '' Open Recordsets
    lv_SQL = "Select * from tbResource Where P_ID = " & Str(ResourceID) & " ORDER BY R_ID, L_ID"
    lv_RS.Open lv_SQL, lv_CN
        
    '' Insert New Data
    If Not (m_rsResource.EOF And m_rsResource.BOF) Then
    
        m_rsResource.MoveFirst
        While Not m_rsResource.EOF
            
            lv_RS.AddNew
            
            For lv_loop = 0 To m_rsResource.Fields.Count - 1
                lv_RS.Fields(lv_loop) = m_rsResource.Fields(lv_loop)
            Next
            
            lv_RS.Update
            
            m_rsResource.MoveNext
        Wend
        
    End If
    
    '' Close Recordsets
    lv_RS.Close
    
    '' Commit Transaction
    lv_CN.CommitTrans
    lv_InTrans = False
    SynchronizeResource = True
    
     'Mike Added @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Resource synchronized Successful! Project ID:" & ResourceID)
    
    '改变鼠标指针形状->箭头光标
    ChangeMousePointer vbDefault, True
    Exit Function
    
BackDoor:
    
    ' Roll back if we are in transaction
    If lv_InTrans Then lv_CN.RollbackTrans
    
    '改变鼠标指针形状->箭头光标
    ChangeMousePointer vbDefault, True
    
    If Err.Description <> "" Then
        MsgBox Err.Description
    End If
    
    'Mike Added @ 2008-7-7
    Call WriteLogMessage(Err.Number, enu_Error, "Failed to synchronize resource!" & f_strConnStr, Err.Description)
    
    On Error GoTo 0
End Function

'' Sun added 2007-04-08
' 资源导出到CSV文件
'
Public Function ExportResourceData(ByVal f_strFileName As String) As Boolean
On Error Resume Next

    Dim FileNumber
    Dim lv_strLine As String
    
    ' 缺省返回值
    ExportResourceData = False
        
    '*****************************************************
    '* Open File for Write
    '*****************************************************
    If CheckExistFile(f_strFileName) Then
        If MsgBox(LoadNationalResString(1669) & " '" & f_strFileName & LoadNationalResString(1109), vbYesNo + vbQuestion, App.Title) = vbNo Then
            Exit Function
        End If
        '' Kill the File
        Call RemoveFile(f_strFileName)
    End If
        
    '' Open file for write
    FileNumber = FreeFile   ' 取得未使用的文件号
    Open f_strFileName For Output Access Write Lock Write As #FileNumber
    If Err Then
        Err.Clear
        MsgBox LoadNationalResString(1110) & f_strFileName & LoadNationalResString(1111), vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
        
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
        
    '*****************************************************
    '* place saving flow information codes here          *
    '*****************************************************
    '' Section 1：IVR服务加载资源列表，3字段用逗号分割
    ''' L_ID, R_ID, R_Path
    If Not (m_rsResource.EOF And m_rsResource.BOF) Then
    
        m_rsResource.MoveFirst
        While Not m_rsResource.EOF
            
            lv_strLine = m_rsResource.Fields("L_ID")
            lv_strLine = lv_strLine & ", " & m_rsResource.Fields("R_ID")
            lv_strLine = lv_strLine & ", " & Trim(m_rsResource.Fields("R_Path"))
            lv_strLine = lv_strLine & vbLf
            
            Print #FileNumber, lv_strLine
            m_rsResource.MoveNext
            
        Wend
        
    End If
    
    '' Section 2：资源项目信息，10字段（比tbRefIVR表多一个Flag字段）用制表键分割
    ''' Flag(始终为-1)
    lv_strLine = "-1"
    lv_strLine = lv_strLine & vbTab & Trim(Str(m_nResourceID))
    lv_strLine = lv_strLine & vbTab & "R"
    lv_strLine = lv_strLine & vbTab & IIf(Trim(m_strResourceName) = "", "N/A", Trim(m_strResourceName))
    lv_strLine = lv_strLine & vbTab & IIf(Trim(m_strResDescription) = "", "N/A", Trim(m_strResDescription))
    lv_strLine = lv_strLine & vbTab & IIf(Trim(m_strResVersion) = "", "N/A", Trim(m_strResVersion))
    lv_strLine = lv_strLine & vbTab & IIf(Trim(m_strResAuther) = "", "N/A", Trim(m_strResAuther))
    lv_strLine = lv_strLine & vbTab & IIf(Trim(m_strResUser) = "", "N/A", Trim(m_strResUser))
    lv_strLine = lv_strLine & vbTab & Format(m_tmResCreateTime, "yyyy-mm-dd hh:nn:ss")
    lv_strLine = lv_strLine & vbTab & Format(m_tmResModifyTime, "yyyy-mm-dd hh:nn:ss")
    lv_strLine = lv_strLine & vbLf
    
    Print #FileNumber, lv_strLine
    
    '' Section 3：资源详细清单，10字段（比tbResource表多一个Flag字段）用制表键分割
    ''' Flag(始终为-2)
    If Not (m_rsResource.EOF And m_rsResource.BOF) Then
    
        m_rsResource.MoveFirst
        While Not m_rsResource.EOF
            
            lv_strLine = "-2"
            lv_strLine = lv_strLine & vbTab & Trim(Str(m_nResourceID))
            lv_strLine = lv_strLine & vbTab & Trim(Str(m_rsResource("L_ID")))
            lv_strLine = lv_strLine & vbTab & Trim(Str(m_rsResource("R_ID")))
            If Not IsNull(m_rsResource("R_Type")) Then
                lv_strLine = lv_strLine & vbTab & Trim(Str(m_rsResource("R_Type")))
            Else
                lv_strLine = lv_strLine & vbTab & "1"
            End If
            If Not IsNull(m_rsResource("R_Description")) Then
                lv_strLine = lv_strLine & vbTab & IIf(Trim(m_rsResource("R_Description")) = "", "N/A", Trim(m_rsResource("R_Description")))
            Else
                lv_strLine = lv_strLine & vbTab & "N/A"
            End If
            If Not IsNull(m_rsResource("R_Path")) Then
                lv_strLine = lv_strLine & vbTab & IIf(Trim(m_rsResource("R_Path")) = "", "N/A", Trim(m_rsResource("R_Path")))
            Else
                lv_strLine = lv_strLine & vbTab & "N/A"
            End If
            If Not IsNull(m_rsResource("R_Note")) Then
                lv_strLine = lv_strLine & vbTab & IIf(Trim(m_rsResource("R_Note")) = "", "N/A", Trim(m_rsResource("R_Note")))
            Else
                lv_strLine = lv_strLine & vbTab & "N/A"
            End If
            If Not IsNull(m_rsResource("CreateTime")) Then
                lv_strLine = lv_strLine & vbTab & Format(m_rsResource("CreateTime"), "yyyy-mm-dd hh:nn:ss")
            Else
                lv_strLine = lv_strLine & vbTab & Format(Now, "yyyy-mm-dd hh:nn:ss")
            End If
            If Not IsNull(m_rsResource("ModifyTime")) Then
                lv_strLine = lv_strLine & vbTab & Format(m_rsResource("ModifyTime"), "yyyy-mm-dd hh:nn:ss")
            Else
                lv_strLine = lv_strLine & vbTab & Format(Now, "yyyy-mm-dd hh:nn:ss")
            End If

            lv_strLine = lv_strLine & vbLf
            
            'Michael added @2008-9-9
            'Fix English comma might build bad CSV file when export the resource project
            lv_strLine = Replace(lv_strLine, ",", "，", , , vbTextCompare)
            
            Print #FileNumber, lv_strLine
            m_rsResource.MoveNext
            
        Wend
        
    End If
    
    'Mike Added @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Successful export Reource Date to the file!", "File Name:" & f_strFileName)
    
    ExportResourceData = True
    
BackDoor:
    
    ' Close File
    Close #FileNumber
    
    'Mike Added @ 2008-7-7
    Call WriteLogMessage(Err.Number, enu_Error, "Failed to export Resource Date to the file:" & f_strFileName & "!", Err.Description)
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

End Function

' 是否是新建项目
'
Public Function IsNewCallFlow() As Boolean
On Error GoTo BackDoor

    IsNewCallFlow = True
    
    If m_rsFS.RecordCount > 0 Then
        IsNewCallFlow = False
    End If
    
Exit Function
BackDoor:
    On Error GoTo 0
End Function

' 新建项目摘要
'
Public Function AddCallFlowIntoList(ByVal f_PID As Byte, ByVal f_Name As String, ByVal f_Desc As String, ByVal f_Auther As String, ByVal f_User As String, ByVal f_Version As String) As Boolean
On Error Resume Next

    If IsNewCallFlow Then
    
        m_rsFS.AddNew
        m_rsFS("p_id") = f_PID                              '流程编号
        m_rsFS("P_Type") = "P"                              '流程类型
        m_rsFS("p_name") = f_Name                           '流程名称
        m_rsFS("p_description") = f_Desc                    '流程描述
        m_rsFS("p_auther") = f_Auther                       '流程作者
        m_rsFS("p_user") = f_User                           '流程用户
        m_rsFS("p_version") = f_Version                     '版本号
        m_rsFS("P_modifytime") = Now()                      '修改时间
        m_rsFS("P_createtime") = Now()                      '创建时间
        m_rsFS.Update
        
        m_strFlowName = f_Name
        m_blnSavedMark = False
        
        'Michael Added @ 2008-7-4
        WriteLogMessage 0, enu_Information, "Create a new Callflow", "CallFlow ID:" & f_PID
    End If
    
    AddCallFlowIntoList = True
    
On Error GoTo 0
End Function

' 更新项目摘要
'
Public Function UpdateCallFlowInfo(ByVal f_Name As String, ByVal f_Desc As String, ByVal f_Auther As String, ByVal f_User As String, ByVal f_Version As String) As Boolean
On Error Resume Next

    If Not IsNewCallFlow Then
        
        ' 流程名称
        If IsNull(m_rsFS("p_name")) Then
            m_rsFS("p_name") = f_Name
            m_blnSavedMark = False
        ElseIf Trim(m_rsFS("p_name")) <> f_Name Then
            m_rsFS("p_name") = f_Name
            m_blnSavedMark = False
        End If
        
        ' 流程描述
        If IsNull(m_rsFS("p_description")) Then
            m_rsFS("p_description") = f_Desc
            m_blnSavedMark = False
        ElseIf Trim(m_rsFS("p_description")) <> f_Desc Then
            m_rsFS("p_description") = f_Desc
            m_blnSavedMark = False
        End If
        
        ' 流程作者
        If IsNull(m_rsFS("p_auther")) Then
            m_rsFS("p_auther") = f_Auther
            m_blnSavedMark = False
        ElseIf Trim(m_rsFS("p_auther")) <> f_Auther Then
            m_rsFS("p_auther") = f_Auther
            m_blnSavedMark = False
        End If
        
        ' 流程用户
        If IsNull(m_rsFS("p_user")) Then
            m_rsFS("p_user") = f_User
            m_blnSavedMark = False
        ElseIf Trim(m_rsFS("p_user")) <> f_User Then
            m_rsFS("p_user") = f_User
            m_blnSavedMark = False
        End If
        
        ' 版本号
        If IsNull(m_rsFS("p_version")) Then
            m_rsFS("p_version") = f_Version
            m_blnSavedMark = False
        ElseIf Trim(m_rsFS("p_version")) <> f_Version Then
            m_rsFS("p_version") = f_Version
            m_blnSavedMark = False
        End If
        
        m_strFlowName = f_Name
        
        ' 修改时间
        If m_blnSavedMark = False Then
            m_rsFS("P_modifytime") = Now()
            m_rsFS.Update
        End If
        
    End If
    
    UpdateCallFlowInfo = True
    'Michael Added @ 2008-7-4
    WriteLogMessage 0, enu_Information, "Update Callflow Information Successful!", "CallFlow Name:" & f_Name
On Error GoTo 0
End Function

' Create Default System Nodes ( 1, 2, 3 - 0, 1, 23 )
'
Public Function CreateSysDefaultNodes() As Boolean
On Error Resume Next

    '' Sun added 2002-05-23
    ''' Locate nodes at proper position
    Dim lv_NodeID As Integer
    Dim lv_NodeNo As Byte
    Dim lv_PageHeight As Long
    Dim lv_PageWidth As Long
    Dim lv_Left As Long
    Dim lv_Top As Long

    lv_PageHeight = Screen.TwipsPerPixelY * 40
    lv_PageWidth = Screen.TwipsPerPixelX * 500
    
    For m_nNodeID = 1 To 3
        
        Select Case m_nNodeID
        Case 1
            lv_NodeNo = 0
            lv_NodeID = 1
            lv_Left = lv_PageWidth + Screen.TwipsPerPixelX * 64 * m_nNodeID
            lv_Top = lv_PageHeight
        Case 2
            lv_NodeNo = 1
            lv_NodeID = 2
            lv_Left = lv_PageWidth + Screen.TwipsPerPixelX * 64 * m_nNodeID
            lv_Top = lv_PageHeight
        Case 3
            lv_NodeNo = 23
            lv_NodeID = 256
            lv_Left = lv_PageWidth / 2
            lv_Top = lv_PageHeight * 2
        End Select
        
        '' Default Properties
        F_NodeDefaultInfo m_nNodeID, lv_NodeNo
        
        m_rsRS.AddNew
        
        m_rsRS("p_id") = m_nFlowID
        m_rsRS("n_id") = lv_NodeID
        m_rsRS("n_no") = lv_NodeNo
        m_rsRS("n_page") = m_nCurrentPage
        m_rsRS("n_left") = lv_Left
        m_rsRS("n_top") = lv_Top
        m_rsRS("n_height") = Screen.TwipsPerPixelY * 32          '' 480
        m_rsRS("n_width") = Screen.TwipsPerPixelX * 32           '' 480
        m_rsRS("n_description") = "N/A"
        
        #If Programnew = 0 Then
            DateStrToAsc m_rsRS, gGetUserVar.StrN_data1, gGetUserVar.StrN_data2
        #Else
            m_rsRS.Fields("n_data1") = LeftB(gGetUserVar.StrN_data1, DEF_NODE_DATA1_LEN)
            m_rsRS.Fields("n_data2") = LeftB(gGetUserVar.StrN_data2, DEF_NODE_DATA2_LEN)
        #End If
        
        m_rsRS.Update
        
    Next
    
    m_nSysNodeID = 2
    m_nUserNodeID = 256
    
    'Michael Added @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Create Default System Node 1,2,256 for Callflow" & m_nFlowID)
End Function

' 打开流程和节点数据表
'
Public Sub OpenIvrRecordSet(ByVal f_PID As Byte)
On Error Resume Next
    
    Dim lv_rsSNodeMax As ADODB.Recordset
    Dim lv_rsUNodeMax As ADODB.Recordset
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    Set lv_rsSNodeMax = New ADODB.Recordset
    Set lv_rsUNodeMax = New ADODB.Recordset
    
    '节点
    m_rsRS.CursorLocation = adUseServer
    m_rsRS.Properties("Determine Key Columns For Rowset") = True
    '--------------------------------------------------
    m_rsRS.CursorType = adOpenKeyset
    m_rsRS.LockType = adLockPessimistic
    m_rsRS.CursorLocation = adUseClient
        
    #If Programnew = 0 Then
        m_rsRS.Open "SELECT * FROM tbivrprogramNew WHERE P_ID= " & f_PID & " ORDER BY N_ID, N_NO", m_adoCN
    #Else
        m_rsRS.Open "SELECT * FROM tbivrprogram WHERE P_ID= " & f_PID & " ORDER BY N_ID, N_NO", m_adoCN
    #End If
    m_nNodeID = m_rsRS.RecordCount
    
    '流程
    m_rsFS.CursorType = adOpenKeyset
    m_rsFS.LockType = adLockPessimistic
    m_rsFS.CursorLocation = adUseClient
    m_rsFS.Open "SELECT * FROM tbrefivr WHERE P_ID = " & f_PID & " AND P_Type = 'P'", m_adoCN
    m_nFlowID = f_PID
    m_strFlowName = Trim(m_rsFS.Fields("P_Name"))
    
    '最大系统节点ID
    lv_rsSNodeMax.CursorType = adOpenForwardOnly
    lv_rsSNodeMax.LockType = adLockReadOnly
    lv_rsSNodeMax.CursorLocation = adUseClient
    
    #If Programnew = 0 Then
        lv_rsSNodeMax.Open "SELECT max(N_id) FROM tbivrprogramnew WHERE N_ID>=0 and N_id<256 and P_ID=" & f_PID, m_adoCN
    #Else
        lv_rsSNodeMax.Open "SELECT max(N_id) FROM tbivrprogram WHERE N_ID>=0 and N_id<256 and P_ID=" & f_PID, , m_adoCN
    #End If
    
    If IsNull(lv_rsSNodeMax(0)) Then
       m_nSysNodeID = 0
    Else
       m_nSysNodeID = lv_rsSNodeMax(0)
    End If
    lv_rsSNodeMax.Close
    
    '最大用户节点ID
    lv_rsUNodeMax.CursorType = adOpenForwardOnly
    lv_rsUNodeMax.LockType = adLockReadOnly
    lv_rsUNodeMax.CursorLocation = adUseClient
    
    #If Programnew = 0 Then
        lv_rsUNodeMax.Open "SELECT max(N_id) FROM tbivrprogramnew WHERE N_ID>= 255 and N_ID<=32767 and P_ID=" & f_PID, m_adoCN
    #Else
        lv_rsUNodeMax.Open "SELECT max(N_id) FROM tbivrprogram WHERE N_ID>= 255 and N_ID<=32767 and P_ID=" & f_PID, m_adoCN
    #End If
    
    If IsNull(lv_rsUNodeMax(0)) Then
       m_nUserNodeID = 255
    Else
       m_nUserNodeID = lv_rsUNodeMax(0)
    End If
    lv_rsUNodeMax.Close
    
    m_blnSavedMark = True
    
    'Mike Added @ 2008-7-7
    WriteLogMessage 0, enu_Information, "Open CallFlow and Nodes Record...", "Callflow ID:" & m_nFlowID & " Callflow Name:" & m_strFlowName
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

    On Error GoTo 0
End Sub

' 打开资源数据表
'
Public Sub OpenResourceRecordSet(ByVal f_PID As Byte, ByVal f_LID As Integer)
On Error GoTo BackDoor
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    Dim lv_RS As ADODB.Recordset
    Dim lv_SQL As String
    
    '' Resource Items
    lv_SQL = "SELECT * FROM tbResource where p_id=" & f_PID & " ORDER BY R_ID, L_ID"
    
    If m_rsResource.State = ADODB.adStateOpen Then
        m_rsResource.Close
    End If
    
    m_rsResource.CursorType = adOpenStatic
    m_rsResource.LockType = adLockPessimistic
    m_rsResource.CursorLocation = adUseClient
    m_rsResource.Open lv_SQL, m_adoCN
    
    '' Sun removed 2007-04-18
    'm_rsResource.Filter = "L_ID = " & Trim(Str(f_LID))
    
    m_nResourceID = f_PID
    m_nLanguageID = f_LID
   
    '' Resource Project
    Set lv_RS = New ADODB.Recordset
    
    lv_SQL = "Select * From tbRefIVR Where P_Type='R' AND P_ID = " & Str(f_PID)
    
    lv_RS.CursorLocation = adUseClient
    lv_RS.LockType = adLockReadOnly
    lv_RS.CursorType = adOpenForwardOnly
    lv_RS.Open lv_SQL, m_adoCN
    If lv_RS.RecordCount > 0 Then
        If Not IsNull(lv_RS("P_Name")) Then
            m_strResourceName = Trim(lv_RS("P_Name"))
        Else
            m_strResourceName = ""
        End If
        If Not IsNull(lv_RS("P_Description")) Then
            m_strResDescription = Trim(lv_RS("P_Description"))
        Else
            m_strResDescription = ""
        End If
        If Not IsNull(lv_RS("P_Auther")) Then
            m_strResAuther = Trim(lv_RS("P_Auther"))
        Else
            m_strResAuther = gSystem.strOSUser
        End If
        If Not IsNull(lv_RS("P_User")) Then
            m_strResUser = Trim(lv_RS("P_User"))
        Else
            m_strResUser = ""
        End If
        If Not IsNull(lv_RS("P_Version")) Then
            m_strResVersion = Trim(lv_RS("P_Version"))
        Else
            m_strResVersion = ""
        End If
        If Not IsNull(lv_RS("P_CreateTime")) Then
            m_tmResCreateTime = lv_RS("P_CreateTime")
        Else
            m_tmResCreateTime = Now
        End If
        If Not IsNull(lv_RS("P_ModifyTime")) Then
            m_tmResModifyTime = lv_RS("P_ModifyTime")
        Else
            m_tmResModifyTime = Now
        End If
    
    End If
    lv_RS.Close
    
    'Michael Added @ 2008-7-4
    WriteLogMessage 0, enu_Information, "Open Resource Record...", "Project ID:" & f_PID
BackDoor:

    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
    If Err.Number > 0 Then
        MsgBox Err.Number & ": " & Err.Source & " - " & Err.Description
        'Mike Added @ 2008-7-7
        Call WriteLogMessage(Err.Number, enu_Error, "Failed to open resource Record! " & "ProjectID:" & f_PID, Err.Description & "-Source:" & Err.Source)
        Resume Next
    End If
    
    On Error GoTo 0
End Sub

Public Sub OpenResourceTable(ByVal f_PID As Byte, ByVal lv_SQL As String)
On Error GoTo BackDoor
    
     '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
    
    Dim lv_RS As ADODB.Recordset
    'Dim lv_SQL As String
    
    '' Resource Items
    'If lv_SQL = "" Then lv_SQL = "SELECT * FROM tbResource where p_id=" & f_PID & " ORDER BY R_ID, L_ID"
    
    If m_rsResource.State = ADODB.adStateOpen Then
        m_rsResource.Close
    End If
    
    m_rsResource.CursorType = adOpenStatic
    m_rsResource.LockType = adLockPessimistic
    m_rsResource.CursorLocation = adUseClient
    m_rsResource.Open lv_SQL, m_adoCN
    
    m_nResourceID = f_PID
    m_nLanguageID = 0
   
    '' Resource Project
    Set lv_RS = New ADODB.Recordset
    
    lv_SQL = "Select * From tbRefIVR Where P_Type='R' AND P_ID = " & Str(f_PID)
    
    lv_RS.CursorLocation = adUseClient
    lv_RS.LockType = adLockReadOnly
    lv_RS.CursorType = adOpenForwardOnly
    lv_RS.Open lv_SQL, m_adoCN
    If lv_RS.RecordCount > 0 Then
        If Not IsNull(lv_RS("P_Name")) Then
            m_strResourceName = Trim(lv_RS("P_Name"))
        Else
            m_strResourceName = ""
        End If
        If Not IsNull(lv_RS("P_Description")) Then
            m_strResDescription = Trim(lv_RS("P_Description"))
        Else
            m_strResDescription = ""
        End If
        If Not IsNull(lv_RS("P_Auther")) Then
            m_strResAuther = Trim(lv_RS("P_Auther"))
        Else
            m_strResAuther = gSystem.strOSUser
        End If
        If Not IsNull(lv_RS("P_User")) Then
            m_strResUser = Trim(lv_RS("P_User"))
        Else
            m_strResUser = ""
        End If
        If Not IsNull(lv_RS("P_Version")) Then
            m_strResVersion = Trim(lv_RS("P_Version"))
        Else
            m_strResVersion = ""
        End If
        If Not IsNull(lv_RS("P_CreateTime")) Then
            m_tmResCreateTime = lv_RS("P_CreateTime")
        Else
            m_tmResCreateTime = Now
        End If
        If Not IsNull(lv_RS("P_ModifyTime")) Then
            m_tmResModifyTime = lv_RS("P_ModifyTime")
        Else
            m_tmResModifyTime = Now
        End If
    
    End If
    lv_RS.Close
    
    'Michael Added @ 2008-7-4
    WriteLogMessage 0, enu_Information, "Open Resource Record...", "Project ID:" & f_PID
    
BackDoor:

    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
    If Err.Number > 0 Then
        MsgBox Err.Number & ": " & Err.Source & " - " & Err.Description
        'Mike added @ 2008-7-8
        Call WriteLogMessage(Err.Number, enu_Error, "Failed to open resource record, Project ID:" & f_PID, Err.Description)
        Resume Next
    End If
    
    On Error GoTo 0
End Sub

' Get Resource Path with specific id
'
Public Function GetResourceNameWithID(ByVal f_RID As Long, ByVal f_LID As Integer, f_RName As String) As Boolean
On Error GoTo BackDoor
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    GetResourceNameWithID = False

    If f_LID >= 0 Then
        m_rsResource.Filter = "L_ID = " & Trim(Str(f_LID))
    Else
        m_rsResource.Filter = "L_ID = " & Trim(Str(m_nLanguageID))
    End If
    
    If m_rsResource.RecordCount > 0 Then
        m_rsResource.MoveFirst
        m_rsResource.Find "R_ID = " & Trim(Str(f_RID))
        If Not m_rsResource.EOF Then
            If Not IsNull(m_rsResource("R_Path")) Then
                f_RName = Trim(m_rsResource("R_Path"))
                GetResourceNameWithID = True
            End If
        End If
    End If

BackDoor:
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
    On Error GoTo 0

End Function

' Get Resource Description with specific id
'
Public Function GetResourceDescriptionWithID(ByVal f_RID As Long, f_RName As String) As Boolean
On Error GoTo BackDoor
    
    Dim lv_Str As String
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    GetResourceDescriptionWithID = False
    
    If m_rsResource.RecordCount > 0 Then
        m_rsResource.MoveFirst
        m_rsResource.Find "R_ID = " & Trim(Str(f_RID))
        If Not m_rsResource.EOF Then
            If Not IsNull(m_rsResource("R_Description")) Then
                f_RName = Trim(m_rsResource("R_Description"))
                GetResourceDescriptionWithID = True
            End If
        End If
    End If

BackDoor:
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    
    On Error GoTo 0

End Function

'更新数据表
'
Public Function UpdateIvrTable() As Boolean
Dim lv_Index As Integer
Dim lv_InTrans As Boolean

Dim debug_nodeid As Integer, debug_Node As Byte

On Error GoTo BackDoor

    UpdateIvrTable = False
    lv_InTrans = False

    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    '' Begin Transaction
    m_adoCN.BeginTrans
    lv_InTrans = True

    '' Sun added 2003-04-27
    m_rsFS("P_modifytime") = Now()
    m_rsFS.Update
    
    For lv_Index = 1 To m_nNodeID
    
        If Not m_ctlNode(lv_Index) Is Nothing Then
#If debugMode = 1 Then
    debug_nodeid = m_ctlNode(lv_Index).NodeID
#End If
            Select Case m_ctlNode(lv_Index).ModifyFlag
            Case DEF_OPERATION_MODIFY
#If debugMode = 1 Then
    debug_Node = 1
#End If
                m_rsRS.MoveFirst
                m_rsRS.Find "n_id =" & m_ctlNode(lv_Index).NodeID
                If Not m_rsRS.EOF Then
                    
                    m_rsRS("n_left") = m_ctlNode(lv_Index).Left
                    m_rsRS("n_page") = m_ctlNode(lv_Index).InPage
                    m_rsRS("n_top") = m_ctlNode(lv_Index).Top
                    m_rsRS("n_height") = m_ctlNode(lv_Index).Height
                    m_rsRS("n_width") = m_ctlNode(lv_Index).Width
                
#If Programnew = 0 Then
                    DateStrToAsc m_rsRS, m_ctlNode(lv_Index).Data1, m_ctlNode(lv_Index).Data2
#Else
                    m_rsRS.Fields("n_data1") = LeftB(m_ctlNode(lv_Index).Data1, DEF_NODE_DATA1_LEN)
                    m_rsRS.Fields("n_data2") = LeftB(m_ctlNode(lv_Index).Data2, DEF_NODE_DATA2_LEN)
#End If
                  
                    m_rsRS("n_description") = Trim(m_ctlNode(lv_Index).Description)
                    ' Mike Added @ 2008-1-30  when update
                    m_rsRS("n_tag") = Trim(m_ctlNode(lv_Index).NodeTag)
                    m_rsRS.Update
                    m_rsRS.MoveFirst
               
                End If
                
            Case DEF_OPERATION_NEW
#If debugMode = 1 Then
    debug_Node = 2
#End If
                m_rsRS.MoveFirst
                m_rsRS.Find "n_id =" & m_ctlNode(lv_Index).NodeID
                If m_rsRS.EOF Then
                  
                    m_rsRS.AddNew
                    m_rsRS("p_id") = CallFlowID
                    m_rsRS("n_id") = m_ctlNode(lv_Index).NodeID
                    m_rsRS("n_no") = m_ctlNode(lv_Index).NodeNo
                    m_rsRS("n_left") = m_ctlNode(lv_Index).Left
                    m_rsRS("n_page") = m_ctlNode(lv_Index).InPage
                    m_rsRS("n_top") = m_ctlNode(lv_Index).Top
                    m_rsRS("n_height") = m_ctlNode(lv_Index).Height
                    m_rsRS("n_width") = m_ctlNode(lv_Index).Width
                  
#If Programnew = 0 Then
                    DateStrToAsc m_rsRS, m_ctlNode(lv_Index).Data1, m_ctlNode(lv_Index).Data2
#Else
                    m_rsRS.Fields("n_data1") = LeftB(m_ctlNode(lv_Index).Data1, DEF_NODE_DATA1_LEN)
                    m_rsRS.Fields("n_data2") = LeftB(m_ctlNode(lv_Index).Data2, DEF_NODE_DATA2_LEN)
#End If
            
                    m_rsRS("n_description") = Trim(m_ctlNode(lv_Index).Description)
                    ' Mike Added @ 2008-1-30
                    m_rsRS("n_tag") = Trim(m_ctlNode(lv_Index).NodeTag)
                    m_rsRS.Update
                    m_rsRS.MoveFirst
               
               End If
                   
            Case DEF_OPERATION_DELETE
#If debugMode = 1 Then
    debug_Node = 3
#End If
                m_rsRS.MoveFirst
                m_rsRS.Find "n_id=" & m_ctlNode(lv_Index).NodeID
                If m_rsRS.EOF And m_rsRS.RecordCount > 0 Then
                    m_rsRS.MoveFirst
                Else
                    If Not m_rsRS.EOF Then
                        m_rsRS.Delete
                        m_rsRS.MoveLast
                    End If
                End If
            End Select
            m_ctlNode(lv_Index).ModifyFlag = DEF_OPERATION_INIT
            
        End If
        
    Next
    
    '' Begin Transaction
    m_blnSavedMark = True
    
    '' Commit Transaction
    m_adoCN.CommitTrans
    lv_InTrans = False
    
    UpdateIvrTable = True
    
    'Mike added @ 2008-7-8
    Call WriteLogMessage(0, enu_Information, "Callflow saved successful!", "Callflow ID:" & CallFlowID)
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
Exit Function
           
BackDoor:

    Message "E899"
    If Err.Description <> "" Then
        MsgBox Err.Description
    End If
    
#If debugMode = 1 Then
    'This is for Debug; delete when release
    If debug_Node = 1 Then
        'Call WriteLogMsg(Err.Description, Err.Number, "UpdateIvrTable()", "Modify_NodeID = " & debug_nodeid)
        Call WriteLogMessage(Err.Number, enu_Debug, "Saving callflow" & CallFlowID & " error after modified Node:" & debug_nodeid, Err.Description)
    ElseIf debug_Node = 2 Then
        'Call WriteLogMsg(Err.Description, Err.Number, "UpdateIvrTable()", "New_NodeID = " & debug_nodeid)
        Call WriteLogMessage(Err.Number, enu_Debug, "Saving callflow" & CallFlowID & " error after added Node:" & debug_nodeid, Err.Description)
    ElseIf debug_Node = 3 Then
        'Call WriteLogMsg(Err.Description, Err.Number, "UpdateIvrTable()", "Delete_NodeID = " & debug_nodeid)
        Call WriteLogMessage(Err.Number, enu_Debug, "Saving callflow" & CallFlowID & " error when delete Node:" & debug_nodeid, Err.Description)
    End If
#Else
    'Call WriteLogMsg(Err.Description, Err.Number, "UpdateIvrTable()")
    Call WriteLogMessage(Err.Number, enu_Error, "Error occured when saving callflow" & CallFlowID & " !", Err.Description)
#End If
    
    ' Roll back if we are in transaction
    If lv_InTrans Then m_adoCN.RollbackTrans
    'Michael added @ 2008-7-4
    Call WriteLogMessage(0, enu_Warnning, "Can not save the callflow, Changes has been Rollback!")

    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
        
    On Error GoTo 0

End Function

' 增加记录
'
Public Function AddNewIvrRecord(f_nIndex As Integer) As Boolean

    m_ctlNode(f_nIndex).ModifyFlag = DEF_OPERATION_NEW
    SavedMark = False
    AddNewIvrRecord = True
    'Mike added @ 2008-7-8
    Call WriteLogMessage(0, enu_Information, "Add a new node", "Node Index:" & f_nIndex)
End Function

'删除记录
'
Public Function DeleteIvrRecord(f_nIndex As Integer) As Boolean

    m_ctlNode(f_nIndex).ModifyFlag = DEF_OPERATION_DELETE
    SavedMark = False
    DeleteIvrRecord = True
    'Mike added @ 2008-7-8
    Call WriteLogMessage(0, enu_Information, "Delete a node", "Node Index:" & f_nIndex)
End Function

Public Function UpdateIvrRecord(lv_nID As Integer, lv_nno As Byte)

    If m_ctlNode(NodeSelectedID).ModifyFlag <> DEF_OPERATION_NEW Then
        m_ctlNode(NodeSelectedID).ModifyFlag = DEF_OPERATION_MODIFY
    End If
    SavedMark = False
    UpdateIvrRecord = True
    'Mike added @ 2008-7-8
    Call WriteLogMessage(0, enu_Information, "Update node info", "Node ID:" & lv_nID)
End Function

Public Function UpdateAnotherIVRRecord(f_nIndex As Integer) As Boolean
On Error Resume Next

    If m_ctlNode(f_nIndex).ModifyFlag <> DEF_OPERATION_NEW Then
        m_ctlNode(f_nIndex).ModifyFlag = DEF_OPERATION_MODIFY
    End If
    SavedMark = False
    UpdateAnotherIVRRecord = True

End Function

' Change the node id properties of specified node
'
Public Sub UpdateNodeIDProperties(ByVal f_nIndex As Integer, f_NewNodeIDList() As Integer, f_OldNodeIDList() As Integer)
'On Error Resume Next

    Dim lv_Node6_data2 As SData2_006
    Dim lv_Node7_data2 As SData2_007
    Dim lv_Node8_data2 As SData2_008
    Dim lv_Node9_data2 As SData2_009
    Dim lv_Node10_data2 As SData2_010
    Dim lv_Node16_data2 As SData2_016
    Dim lv_Node17_data2 As SData2_017
    Dim lv_Node18_data2 As SData2_018
    Dim lv_Node19_data2 As SData2_019
    Dim lv_Node20_data2 As SData2_020
    Dim lv_Node21_data2 As SData2_021
    Dim lv_Node22_data2 As SData2_022
    Dim lv_Node23_data2 As SData2_023
    Dim lv_Node28_data2 As SData2_028
    Dim lv_Node40_data2 As SData2_040
    Dim lv_Node41_data2 As SData2_041
    Dim lv_Node50_data2 As SData2_050
    Dim lv_Node51_data2 As SData2_051
    Dim lv_Node55_data2 As SData2_055
    Dim lv_Node60_data2 As SData2_060
    Dim lv_Node61_data2 As SData2_061
    Dim lv_Node62_data2 As SData2_062
    Dim lv_Node63_data2 As SData2_063
    Dim lv_Node69_data2 As SData2_069
    Dim lv_Node70_data2 As SData2_070
    Dim lv_Node71_data2 As SData2_071
    Dim lv_Node80_data2 As SData2_080
    Dim lv_Node90_data2 As SData2_090
    Dim lv_Node91_data2 As SData2_091
    Dim lv_Node96_data2 As SData2_096
    Dim lv_Node100_data2 As SData2_100
    Dim lv_Node101_data2 As SData2_101
    Dim lv_Node102_data2 As SData2_102
    Dim lv_Node255_data2 As SData2_255
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_NodeID As Integer
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    Dim lv_Str As String
    Dim lv_SubLoop As Integer
    
    If f_nIndex > 0 Then

        For lv_SubLoop = 0 To 63
            lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(f_nIndex).Data2, lv_SubLoop + 1, 1))
        Next
        
        Select Case m_ctlNode(f_nIndex).NodeNo
        Case 6
             CopyMemory lv_Node6_data2.nd_goto, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node6_data2.nd_parent = UpdateNodeIDProperty(lv_Node6_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node6_data2.nd_goto = UpdateNodeIDProperty(lv_Node6_data2.nd_goto, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node6_data2.nd_goto, DEF_NODE_DATA2_LEN
             
        Case 7
             CopyMemory lv_Node7_data2.vox_userid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node7_data2.nd_parent = UpdateNodeIDProperty(lv_Node7_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node7_data2.nd_fail = UpdateNodeIDProperty(lv_Node7_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node7_data2.nd_succeed = UpdateNodeIDProperty(lv_Node7_data2.nd_succeed, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node7_data2.vox_userid, DEF_NODE_DATA2_LEN
        Case 8
             CopyMemory lv_Node8_data2.vox_password, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node8_data2.nd_parent = UpdateNodeIDProperty(lv_Node8_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node8_data2.nd_fail = UpdateNodeIDProperty(lv_Node8_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node8_data2.nd_succeed = UpdateNodeIDProperty(lv_Node8_data2.nd_succeed, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node8_data2.vox_password, DEF_NODE_DATA2_LEN
        Case 9
             CopyMemory lv_Node9_data2.workday, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node9_data2.nd_parent = UpdateNodeIDProperty(lv_Node9_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node9_data2.nd_sparetime = UpdateNodeIDProperty(lv_Node9_data2.nd_sparetime, f_NewNodeIDList(), f_OldNodeIDList())
             For lv_SubLoop = 0 To UBound(lv_Node9_data2.nd_timesec)
                lv_Node9_data2.nd_timesec(lv_SubLoop) = UpdateNodeIDProperty(lv_Node9_data2.nd_timesec(lv_SubLoop), f_NewNodeIDList(), f_OldNodeIDList())
             Next
             CopyMemory lv_Gdata2(0), lv_Node9_data2.workday, DEF_NODE_DATA2_LEN
        Case 10
             CopyMemory lv_Node10_data2.daytype(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node10_data2.nd_parent = UpdateNodeIDProperty(lv_Node10_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             For lv_SubLoop = 0 To UBound(lv_Node10_data2.nd_daysec)
                lv_Node10_data2.nd_daysec(lv_SubLoop) = UpdateNodeIDProperty(lv_Node10_data2.nd_daysec(lv_SubLoop), f_NewNodeIDList(), f_OldNodeIDList())
             Next
             CopyMemory lv_Gdata2(0), lv_Node10_data2.daytype(0), DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2004-12-30
        Case 16
             CopyMemory lv_Node16_data2.var_value(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node16_data2.nd_parent = UpdateNodeIDProperty(lv_Node16_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node16_data2.nd_fail = UpdateNodeIDProperty(lv_Node16_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node16_data2.nd_succ = UpdateNodeIDProperty(lv_Node16_data2.nd_succ, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node16_data2.var_value(0), DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 17
             CopyMemory lv_Node17_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node17_data2.nd_parent = UpdateNodeIDProperty(lv_Node17_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node17_data2.nd_fail = UpdateNodeIDProperty(lv_Node17_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node17_data2.nd_succ = UpdateNodeIDProperty(lv_Node17_data2.nd_succ, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node17_data2.vox_play, DEF_NODE_DATA2_LEN
        Case 18
             CopyMemory lv_Node18_data2.typeflags(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node18_data2.nd_parent = UpdateNodeIDProperty(lv_Node18_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node18_data2.nd_child = UpdateNodeIDProperty(lv_Node18_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node18_data2.typeflags(0), DEF_NODE_DATA2_LEN
        Case 19
             CopyMemory lv_Node19_data2.delaytime, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node19_data2.nd_parent = UpdateNodeIDProperty(lv_Node19_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node19_data2.delaytime, DEF_NODE_DATA2_LEN
        Case 20
             CopyMemory lv_Node20_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node20_data2.nd_parent = UpdateNodeIDProperty(lv_Node20_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node20_data2.vox_play, DEF_NODE_DATA2_LEN
        Case 21
             CopyMemory lv_Node21_data2.vox_pred, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node21_data2.nd_parent = UpdateNodeIDProperty(lv_Node21_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node21_data2.nd_child = UpdateNodeIDProperty(lv_Node21_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node21_data2.vox_pred, DEF_NODE_DATA2_LEN
        Case 22
             CopyMemory lv_Node22_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node22_data2.nd_parent = UpdateNodeIDProperty(lv_Node22_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node22_data2.nd_nodefail = UpdateNodeIDProperty(lv_Node22_data2.nd_nodefail, f_NewNodeIDList(), f_OldNodeIDList())
             For lv_SubLoop = 0 To UBound(lv_Node22_data2.nd_key)
                lv_Node22_data2.nd_key(lv_SubLoop) = UpdateNodeIDProperty(lv_Node22_data2.nd_key(lv_SubLoop), f_NewNodeIDList(), f_OldNodeIDList())
             Next
             CopyMemory lv_Gdata2(0), lv_Node22_data2.vox_play, DEF_NODE_DATA2_LEN
        Case 23
             CopyMemory lv_Node23_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node23_data2.nd_parent = UpdateNodeIDProperty(lv_Node23_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node23_data2.nd_goto = UpdateNodeIDProperty(lv_Node23_data2.nd_goto, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node23_data2.vox_play, DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2004-12-30
        Case 28
             CopyMemory lv_Node28_data2.vox_string, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node28_data2.nd_parent = UpdateNodeIDProperty(lv_Node28_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node28_data2.nd_fail = UpdateNodeIDProperty(lv_Node28_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node28_data2.nd_succ = UpdateNodeIDProperty(lv_Node28_data2.nd_succ, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node28_data2.vox_string, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 40
             CopyMemory lv_Node40_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node40_data2.nd_parent = UpdateNodeIDProperty(lv_Node40_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node40_data2.nd_child = UpdateNodeIDProperty(lv_Node40_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node40_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 41
             CopyMemory lv_Node41_data2.vox_play(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node41_data2.nd_parent = UpdateNodeIDProperty(lv_Node41_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node41_data2.nd_child = UpdateNodeIDProperty(lv_Node41_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node41_data2.vox_play(0), DEF_NODE_DATA2_LEN
        
        Case 50
             CopyMemory lv_Node50_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node50_data2.nd_parent = UpdateNodeIDProperty(lv_Node50_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node50_data2.nd_succ = UpdateNodeIDProperty(lv_Node50_data2.nd_succ, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node50_data2.nd_fail = UpdateNodeIDProperty(lv_Node50_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node50_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 51
             CopyMemory lv_Node51_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node51_data2.nd_parent = UpdateNodeIDProperty(lv_Node51_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node51_data2.nd_child = UpdateNodeIDProperty(lv_Node51_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node51_data2.vox_op, DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2006-12-31
        Case 55
             CopyMemory lv_Node55_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node55_data2.nd_parent = UpdateNodeIDProperty(lv_Node55_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node55_data2.nd_succ = UpdateNodeIDProperty(lv_Node55_data2.nd_succ, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node55_data2.nd_fail = UpdateNodeIDProperty(lv_Node55_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node55_data2.vox_op, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 60
             CopyMemory lv_Node60_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node60_data2.nd_parent = UpdateNodeIDProperty(lv_Node60_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node60_data2.nd_busy = UpdateNodeIDProperty(lv_Node60_data2.nd_busy, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node60_data2.nd_noanswer = UpdateNodeIDProperty(lv_Node60_data2.nd_noanswer, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node60_data2.nd_nobody = UpdateNodeIDProperty(lv_Node60_data2.nd_nobody, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node60_data2.nd_ok = UpdateNodeIDProperty(lv_Node60_data2.nd_ok, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node60_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 61
             CopyMemory lv_Node61_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node61_data2.nd_parent = UpdateNodeIDProperty(lv_Node61_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node61_data2.nd_busy = UpdateNodeIDProperty(lv_Node61_data2.nd_busy, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node61_data2.nd_noanswer = UpdateNodeIDProperty(lv_Node61_data2.nd_noanswer, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node61_data2.nd_nobody = UpdateNodeIDProperty(lv_Node61_data2.nd_nobody, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node61_data2.nd_ok = UpdateNodeIDProperty(lv_Node61_data2.nd_ok, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node61_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 62
             CopyMemory lv_Node62_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node62_data2.nd_parent = UpdateNodeIDProperty(lv_Node62_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node62_data2.nd_failed = UpdateNodeIDProperty(lv_Node62_data2.nd_failed, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node62_data2.nd_noans = UpdateNodeIDProperty(lv_Node62_data2.nd_noans, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node62_data2.nd_ok = UpdateNodeIDProperty(lv_Node62_data2.nd_ok, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node62_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 63
             CopyMemory lv_Node63_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node63_data2.nd_parent = UpdateNodeIDProperty(lv_Node63_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node63_data2.nd_busy = UpdateNodeIDProperty(lv_Node63_data2.nd_busy, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node63_data2.nd_nobody = UpdateNodeIDProperty(lv_Node63_data2.nd_nobody, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node63_data2.nd_ok = UpdateNodeIDProperty(lv_Node63_data2.nd_ok, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node63_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 69
             CopyMemory lv_Node69_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node69_data2.nd_parent = UpdateNodeIDProperty(lv_Node69_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node69_data2.nd_child = UpdateNodeIDProperty(lv_Node69_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node69_data2.vox_op, DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2005-06-27
        Case 70
             CopyMemory lv_Node70_data2.routepointid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node70_data2.nd_parent = UpdateNodeIDProperty(lv_Node70_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node70_data2.nd_fail = UpdateNodeIDProperty(lv_Node70_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node70_data2.nd_no = UpdateNodeIDProperty(lv_Node70_data2.nd_no, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node70_data2.nd_yes = UpdateNodeIDProperty(lv_Node70_data2.nd_yes, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node70_data2.routepointid, DEF_NODE_DATA2_LEN
        
        Case 71
             CopyMemory lv_Node71_data2.agentid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node71_data2.nd_parent = UpdateNodeIDProperty(lv_Node71_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node71_data2.nd_fail = UpdateNodeIDProperty(lv_Node71_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node71_data2.nd_no = UpdateNodeIDProperty(lv_Node71_data2.nd_no, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node71_data2.nd_yes = UpdateNodeIDProperty(lv_Node71_data2.nd_yes, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node71_data2.agentid, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 80
             CopyMemory lv_Node80_data2.vox_wait, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node80_data2.nd_parent = UpdateNodeIDProperty(lv_Node80_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node80_data2.nd_child = UpdateNodeIDProperty(lv_Node80_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node80_data2.vox_wait, DEF_NODE_DATA2_LEN
             
        Case 90
             CopyMemory lv_Node90_data2.predial(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node90_data2.nd_parent = UpdateNodeIDProperty(lv_Node90_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node90_data2.nd_fail = UpdateNodeIDProperty(lv_Node90_data2.nd_fail, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node90_data2.nd_succ = UpdateNodeIDProperty(lv_Node90_data2.nd_succ, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node90_data2.predial(0), DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2005-05-26
        Case 91
             CopyMemory lv_Node91_data2.vox_talklen, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node91_data2.nd_parent = UpdateNodeIDProperty(lv_Node91_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node91_data2.nd_child = UpdateNodeIDProperty(lv_Node91_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node91_data2.vox_talklen, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        '-------------------------------------------
        ' Sun added 2004-03-15
        Case 96
             CopyMemory lv_Node96_data2.command, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node96_data2.nd_parent = UpdateNodeIDProperty(lv_Node96_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node96_data2.nd_child = UpdateNodeIDProperty(lv_Node96_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node96_data2.nd_timeout = UpdateNodeIDProperty(lv_Node96_data2.nd_timeout, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node96_data2.command, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 100
             CopyMemory lv_Node100_data2.dll_fid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node100_data2.nd_parent = UpdateNodeIDProperty(lv_Node100_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node100_data2.nd_child = UpdateNodeIDProperty(lv_Node100_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node100_data2.dll_fid, DEF_NODE_DATA2_LEN
        Case 101
             CopyMemory lv_Node101_data2.reserved1(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node101_data2.nd_parent = UpdateNodeIDProperty(lv_Node101_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node101_data2.nd_child = UpdateNodeIDProperty(lv_Node101_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node101_data2.reserved1(0), DEF_NODE_DATA2_LEN
        
        Case 102
             CopyMemory lv_Node102_data2.reserved1(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node102_data2.nd_parent = UpdateNodeIDProperty(lv_Node102_data2.nd_parent, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node102_data2.nd_child = UpdateNodeIDProperty(lv_Node102_data2.nd_child, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node102_data2.reserved1(0), DEF_NODE_DATA2_LEN
        
        Case 255
             CopyMemory lv_Node255_data2.StartNode, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node255_data2.StartNode = UpdateNodeIDProperty(lv_Node255_data2.StartNode, f_NewNodeIDList(), f_OldNodeIDList())
             lv_Node255_data2.EndNode = UpdateNodeIDProperty(lv_Node255_data2.EndNode, f_NewNodeIDList(), f_OldNodeIDList())
             CopyMemory lv_Gdata2(0), lv_Node255_data2.StartNode, DEF_NODE_DATA2_LEN
        
        Case Else
        End Select
        
        '' Update Node data
        lv_Str = ""
        For lv_SubLoop = 0 To 63
            lv_Str = lv_Str & ChrB(lv_Gdata2(lv_SubLoop))
        Next
        m_ctlNode(f_nIndex).Data2 = lv_Str

    End If
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

End Sub

Private Function UpdateNodeIDProperty(ByVal f_nNodeID, f_NewNodeIDList() As Integer, f_OldNodeIDList() As Integer) As Integer
Dim lv_loop As Integer

If f_nNodeID > 255 Then
    For lv_loop = 0 To UBound(f_OldNodeIDList) - 1
        If f_nNodeID = f_OldNodeIDList(lv_loop) Then
            UpdateNodeIDProperty = f_NewNodeIDList(lv_loop)
            Exit Function
        End If
    Next
End If

UpdateNodeIDProperty = 0

End Function

' Change the parent node of specified node
' Note: all input parameters are node No. not select id
'
Public Sub SetNodeParentNote(ByVal f_Node As Integer, ByVal f_ParentNode As Integer)
On Error Resume Next

    Dim lv_Node6_data2 As SData2_006
    Dim lv_Node7_data2 As SData2_007
    Dim lv_Node8_data2 As SData2_008
    Dim lv_Node9_data2 As SData2_009
    Dim lv_Node10_data2 As SData2_010
    Dim lv_Node16_data2 As SData2_016
    Dim lv_Node17_data2 As SData2_017
    Dim lv_Node18_data2 As SData2_018
    Dim lv_Node19_data2 As SData2_019
    Dim lv_Node20_data2 As SData2_020
    Dim lv_Node21_data2 As SData2_021
    Dim lv_Node22_data2 As SData2_022
    Dim lv_Node23_data2 As SData2_023
    Dim lv_Node28_data2 As SData2_028
    Dim lv_Node40_data2 As SData2_040
    Dim lv_Node41_data2 As SData2_041
    Dim lv_Node50_data2 As SData2_050
    Dim lv_Node51_data2 As SData2_051
    Dim lv_Node55_data2 As SData2_055
    Dim lv_Node60_data2 As SData2_060
    Dim lv_Node61_data2 As SData2_061
    Dim lv_Node62_data2 As SData2_062
    Dim lv_Node63_data2 As SData2_063
    Dim lv_Node69_data2 As SData2_069
    Dim lv_Node70_data2 As SData2_070
    Dim lv_Node71_data2 As SData2_071
    Dim lv_Node80_data2 As SData2_080
    Dim lv_Node90_data2 As SData2_090
    Dim lv_Node91_data2 As SData2_091
    Dim lv_Node96_data2 As SData2_096
    Dim lv_Node100_data2 As SData2_100
    Dim lv_Node101_data2 As SData2_101
    Dim lv_Node102_data2 As SData2_102
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_ParentID As Integer, lv_Index As Integer
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    Dim lv_Str As String
    Dim lv_SubLoop As Integer
    
    ' Get out of Sub if Parent Node is 019
    lv_ParentID = SearchNodeIndexWithID(f_ParentNode)
    lv_Index = SearchNodeIndexWithID(f_Node)
    If lv_ParentID > 0 And lv_Index > 0 Then

        For lv_SubLoop = 0 To 63
            lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_Index).Data2, lv_SubLoop + 1, 1))
        Next
                    
        Select Case m_ctlNode(lv_Index).NodeNo
        Case 6
             CopyMemory lv_Node6_data2.nd_goto, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node6_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node6_data2.nd_goto, DEF_NODE_DATA2_LEN
             
        Case 7
             CopyMemory lv_Node7_data2.vox_userid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node7_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node7_data2.vox_userid, DEF_NODE_DATA2_LEN
        Case 8
             CopyMemory lv_Node8_data2.vox_password, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node8_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node8_data2.vox_password, DEF_NODE_DATA2_LEN
        Case 9
             CopyMemory lv_Node9_data2.workday, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node9_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node9_data2.workday, DEF_NODE_DATA2_LEN
        Case 10
             CopyMemory lv_Node10_data2.daytype(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node10_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node10_data2.daytype(0), DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2004-12-30
        Case 16
             CopyMemory lv_Node16_data2.var_value(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node16_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node16_data2.var_value(0), DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 17
             CopyMemory lv_Node17_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node17_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node17_data2.vox_play, DEF_NODE_DATA2_LEN
        Case 18
             CopyMemory lv_Node18_data2.typeflags(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node18_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node18_data2.typeflags(0), DEF_NODE_DATA2_LEN
        Case 19
             CopyMemory lv_Node19_data2.delaytime, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node19_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node19_data2.delaytime, DEF_NODE_DATA2_LEN
        Case 20
             CopyMemory lv_Node20_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node20_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node20_data2.vox_play, DEF_NODE_DATA2_LEN
        Case 21
             CopyMemory lv_Node21_data2.vox_pred, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node21_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node21_data2.vox_pred, DEF_NODE_DATA2_LEN
        Case 22
             CopyMemory lv_Node22_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node22_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node22_data2.vox_play, DEF_NODE_DATA2_LEN
        Case 23
             CopyMemory lv_Node23_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node23_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node23_data2.vox_play, DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2004-12-30
        Case 28
             CopyMemory lv_Node28_data2.vox_string, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node28_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node28_data2.vox_string, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 40
             CopyMemory lv_Node40_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node40_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node40_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 41
             CopyMemory lv_Node41_data2.vox_play(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node41_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node41_data2.vox_play(0), DEF_NODE_DATA2_LEN
        
        Case 50
             CopyMemory lv_Node50_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node50_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node50_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 51
             CopyMemory lv_Node51_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node51_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node51_data2.vox_op, DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2006-12-31
        Case 55
             CopyMemory lv_Node55_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node55_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node55_data2.vox_op, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 60
             CopyMemory lv_Node60_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node60_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node60_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 61
             CopyMemory lv_Node61_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node61_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node61_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 62
             CopyMemory lv_Node62_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node62_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node62_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 63
             CopyMemory lv_Node63_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node63_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node63_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 69
             CopyMemory lv_Node69_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node69_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node69_data2.vox_op, DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2005-06-27
        Case 70
             CopyMemory lv_Node70_data2.routepointid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node70_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node70_data2.routepointid, DEF_NODE_DATA2_LEN
        
        Case 71
             CopyMemory lv_Node71_data2.agentid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node71_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node71_data2.agentid, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 80
             CopyMemory lv_Node80_data2.vox_wait, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node80_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node80_data2.vox_wait, DEF_NODE_DATA2_LEN
        
        Case 90
             CopyMemory lv_Node90_data2.predial(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node90_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node90_data2.predial(0), DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2005-05-26
        Case 91
             CopyMemory lv_Node91_data2.vox_talklen, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node91_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node91_data2.vox_talklen, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        '-------------------------------------------
        ' Sun added 2004-03-15
        Case 96
             CopyMemory lv_Node96_data2.command, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node96_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node96_data2.command, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 100
             CopyMemory lv_Node100_data2.dll_fid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node100_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node100_data2.dll_fid, DEF_NODE_DATA2_LEN
        Case 101
             CopyMemory lv_Node101_data2.reserved1(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node101_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node101_data2.reserved1(0), DEF_NODE_DATA2_LEN
        
        Case 102
             CopyMemory lv_Node102_data2.reserved1(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node102_data2.nd_parent = f_ParentNode
             CopyMemory lv_Gdata2(0), lv_Node102_data2.reserved1(0), DEF_NODE_DATA2_LEN
        
        Case Else
        End Select
        
        '' Update Node data
        lv_Str = ""
        For lv_SubLoop = 0 To 63
            lv_Str = lv_Str & ChrB(lv_Gdata2(lv_SubLoop))
        Next
        m_ctlNode(lv_Index).Data2 = lv_Str

        UpdateAnotherIVRRecord lv_Index

    End If
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

End Sub

' 自动更新节点连线序号，在父节点转移属性改变时调用
'
Public Sub AutoChangeLineIndex(ByVal f_nNodeNo As Integer, ByVal f_StartNode As Integer, ByVal f_EndNode As Integer, ByVal f_nLineIndex As Byte)
On Error Resume Next
    
    Dim lv_Data2 As SData2_255
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_Index As Integer
    Dim lv_Str As String
    Dim lv_SubLoop As Integer
    
    If f_EndNode <= 255 Then
        Exit Sub
    End If
    
    '' 检索起止节点符合要求的连线
    For lv_Index = 1 To m_nNodeID
        If Not m_ctlNode(lv_Index) Is Nothing Then
            If Not m_ctlNode(lv_Index).NodeDestroyed Then
                
                '' 连线
                If m_ctlNode(lv_Index).NodeNo = 255 Then
                    
                    For lv_SubLoop = 0 To 63
                        lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_Index).Data2, lv_SubLoop + 1, 1))
                    Next
                    CopyMemory lv_Data2.StartNode, lv_Gdata2(0), DEF_NODE_DATA2_LEN
                    
                    '' 此节点之连线
                    If lv_Data2.StartNode = f_StartNode And lv_Data2.EndNode = f_EndNode Then
                        If lv_Data2.Index <> f_nLineIndex Then
                            lv_Data2.Index = f_nLineIndex
                            
                            CopyMemory lv_Gdata2(0), lv_Data2.StartNode, DEF_NODE_DATA2_LEN
                            
                            m_ctlNode(lv_Index).Description = F_GetLineCaption(f_nNodeNo, f_nLineIndex)
                            
                            lv_Str = ""
                            For lv_SubLoop = 0 To 63
                                lv_Str = lv_Str & ChrB(lv_Gdata2(lv_SubLoop))
                            Next
                            m_ctlNode(lv_Index).Data2 = lv_Str
                            UpdateAnotherIVRRecord lv_Index
                        End If
                    End If
                
                End If
            
            End If
        End If
    Next
    
End Sub

' 重新设置节点全部子节点，在修改连线序号时调用
' f_nNodeIndex - 连线节点下标
' f_Node - 起始节点
' f_ChildNode - 子节点
' f_nLineIndex - 连线节点序号
Public Function ResetNodeChildrenNodes(ByVal f_nNodeIndex As Integer, ByVal f_Node As Integer, ByVal f_ChildNode As Integer, ByVal f_nLineIndex As Byte) As String
On Error Resume Next

    Dim lv_Data2 As SData2_255
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_Index As Integer
    Dim lv_NodeIndexID As Integer
    Dim lv_SubLoop As Integer
    Dim lv_Str As String
    
    '' 需要重新排定
    lv_NodeIndexID = SearchNodeIndexWithID(f_Node)
    
    '' 更新当前连线
    ResetNodeChildrenNodes = SetNodeChildNode(f_Node, f_ChildNode, f_nLineIndex)
    If ResetNodeChildrenNodes <> "" Then
    
        '' 列举所有StartNode为f_Node的连线
        For lv_Index = 1 To m_nNodeID
            If Not m_ctlNode(lv_Index) Is Nothing Then
                If Not m_ctlNode(lv_Index).NodeDestroyed Then
                    
                    '' 连线
                    If m_ctlNode(lv_Index).NodeNo = 255 And lv_Index <> f_nNodeIndex Then
                            
                        For lv_SubLoop = 0 To 63
                            lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_Index).Data2, lv_SubLoop + 1, 1))
                        Next
                        CopyMemory lv_Data2.StartNode, lv_Gdata2(0), DEF_NODE_DATA2_LEN
                        
                        '' 此节点之连线
                        If lv_Data2.StartNode = f_Node And lv_Data2.EndNode > 255 Then
                            lv_Str = SetNodeChildNode(lv_Data2.StartNode, lv_Data2.EndNode, lv_Data2.Index)
                            If lv_Str <> "" Then
                                m_ctlNode(lv_Index).Description = lv_Str
                                UpdateAnotherIVRRecord lv_Index
                            End If
                        End If
                    End If
                End If
            End If
        Next
            
    End If
    
End Function

' Change the child node of specified node
' Note: all input parameters are node No. not select id
'
' f_Node - 起始节点
' f_ChildNode - 子节点
' f_nLineIndex - 连线节点序号
'
Public Function SetNodeChildNode(ByVal f_Node As Integer, ByVal f_ChildNode As Integer, ByVal f_nLineIndex As Byte) As String
On Error Resume Next

    Dim lv_Node6_data2 As SData2_006
    Dim lv_Node7_data2 As SData2_007
    Dim lv_Node8_data2 As SData2_008
    Dim lv_Node9_data2 As SData2_009
    Dim lv_Node10_data2 As SData2_010
    Dim lv_Node16_data2 As SData2_016
    Dim lv_Node17_data2 As SData2_017
    Dim lv_Node18_data2 As SData2_018
    Dim lv_Node19_data2 As SData2_019
    Dim lv_Node20_data2 As SData2_020
    Dim lv_Node21_data2 As SData2_021
    Dim lv_Node22_data2 As SData2_022
    Dim lv_Node23_data2 As SData2_023
    Dim lv_Node28_data2 As SData2_028
    Dim lv_Node40_data2 As SData2_040
    Dim lv_Node41_data2 As SData2_041
    Dim lv_Node50_data2 As SData2_050
    Dim lv_Node51_data2 As SData2_051
    Dim lv_Node55_data2 As SData2_055
    Dim lv_Node60_data2 As SData2_060
    Dim lv_Node61_data2 As SData2_061
    Dim lv_Node62_data2 As SData2_062
    Dim lv_Node63_data2 As SData2_063
    Dim lv_Node69_data2 As SData2_069
    Dim lv_Node70_data2 As SData2_070
    Dim lv_Node71_data2 As SData2_071
    Dim lv_Node80_data2 As SData2_080
    Dim lv_Node90_data2 As SData2_090
    Dim lv_Node91_data2 As SData2_091
    Dim lv_Node96_data2 As SData2_096
    Dim lv_Node100_data2 As SData2_100
    Dim lv_Node101_data2 As SData2_101
    Dim lv_Node102_data2 As SData2_102
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_ChildID As Integer, lv_Index As Integer
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    Dim lv_Str As String
    Dim lv_SubLoop As Integer
    
    lv_ChildID = SearchNodeIndexWithID(f_ChildNode)
    lv_Index = SearchNodeIndexWithID(f_Node)
    If lv_ChildID > 0 And lv_Index > 0 Then
    
        For lv_SubLoop = 0 To 63
            lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_Index).Data2, lv_SubLoop + 1, 1))
        Next
                    
        Select Case m_ctlNode(lv_Index).NodeNo
        Case 6
            CopyMemory lv_Node6_data2.nd_goto, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            lv_Node6_data2.nd_goto = f_ChildNode
            CopyMemory lv_Gdata2(0), lv_Node6_data2.nd_goto, DEF_NODE_DATA2_LEN
             
        Case 7
            CopyMemory lv_Node7_data2.vox_userid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node7_data2.nd_fail = f_ChildNode
            Else
                lv_Node7_data2.nd_succeed = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node7_data2.vox_userid, DEF_NODE_DATA2_LEN
            
        Case 8
            CopyMemory lv_Node8_data2.vox_password, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node8_data2.nd_fail = f_ChildNode
            Else
                lv_Node8_data2.nd_succeed = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node8_data2.vox_password, DEF_NODE_DATA2_LEN
            
        Case 9
            CopyMemory lv_Node9_data2.workday, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex <= 0 Then
                lv_Node9_data2.nd_sparetime = f_ChildNode
            Else
                lv_Node9_data2.nd_timesec((f_nLineIndex - 1) Mod 6) = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node9_data2.workday, DEF_NODE_DATA2_LEN
        
        Case 10
            CopyMemory lv_Node10_data2.daytype(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node10_data2.nd_daysec(0) = f_ChildNode
            ElseIf f_nLineIndex = 1 Then
                lv_Node10_data2.nd_daysec(1) = f_ChildNode
            Else
                lv_Node10_data2.nd_daysec(2) = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node10_data2.daytype(0), DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2004-12-30
        Case 16
            CopyMemory lv_Node16_data2.var_value(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node16_data2.nd_fail = f_ChildNode
            Else
                lv_Node16_data2.nd_succ = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node16_data2.var_value(0), DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 17
            CopyMemory lv_Node17_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node17_data2.nd_fail = f_ChildNode
            Else
                lv_Node17_data2.nd_succ = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node17_data2.vox_play, DEF_NODE_DATA2_LEN
        
        Case 18
             CopyMemory lv_Node18_data2.typeflags(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node18_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node18_data2.typeflags(0), DEF_NODE_DATA2_LEN
             
        'Case 19
        'Case 20
        
        Case 21
             CopyMemory lv_Node21_data2.vox_pred, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node21_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node21_data2.vox_pred, DEF_NODE_DATA2_LEN
        
        Case 22
            CopyMemory lv_Node22_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex >= 0 And f_nLineIndex <= 11 Then
                lv_Node22_data2.nd_key(f_nLineIndex) = f_ChildNode
            Else
                lv_Node22_data2.nd_nodefail = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node22_data2.vox_play, DEF_NODE_DATA2_LEN
        
        Case 23
             CopyMemory lv_Node23_data2.vox_play, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node23_data2.nd_goto = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node23_data2.vox_play, DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2004-12-30
        Case 28
            CopyMemory lv_Node28_data2.vox_string, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node28_data2.nd_fail = f_ChildNode
            Else
                lv_Node28_data2.nd_succ = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node28_data2.vox_string, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 40
             CopyMemory lv_Node40_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node40_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node40_data2.vox_op, DEF_NODE_DATA2_LEN
        Case 41
             CopyMemory lv_Node41_data2.vox_play(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node41_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node41_data2.vox_play(0), DEF_NODE_DATA2_LEN
        
        Case 50
             CopyMemory lv_Node50_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node50_data2.nd_fail = f_ChildNode
            Else
                lv_Node50_data2.nd_succ = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node50_data2.vox_op, DEF_NODE_DATA2_LEN
        
        Case 51
             CopyMemory lv_Node51_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node51_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node51_data2.vox_op, DEF_NODE_DATA2_LEN
             
        Case 55
            CopyMemory lv_Node55_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node55_data2.nd_fail = f_ChildNode
            Else
                lv_Node55_data2.nd_succ = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node55_data2.vox_op, DEF_NODE_DATA2_LEN
        
        'Case 60
        'Case 61
        'Case 62
        'Case 63
        Case 69
             CopyMemory lv_Node69_data2.vox_op, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node69_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node69_data2.vox_op, DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2005-06-27
        Case 70
            CopyMemory lv_Node70_data2.routepointid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node70_data2.nd_no = f_ChildNode
            ElseIf f_nLineIndex = 1 Then
                lv_Node70_data2.nd_yes = f_ChildNode
            Else
                lv_Node70_data2.nd_fail = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node70_data2.routepointid, DEF_NODE_DATA2_LEN
        
        Case 71
            CopyMemory lv_Node71_data2.agentid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node71_data2.nd_no = f_ChildNode
            ElseIf f_nLineIndex = 1 Then
                lv_Node71_data2.nd_yes = f_ChildNode
            Else
                lv_Node71_data2.nd_fail = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node71_data2.agentid, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 80
             CopyMemory lv_Node80_data2.vox_wait, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node80_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node80_data2.vox_wait, DEF_NODE_DATA2_LEN
        
        Case 90
            CopyMemory lv_Node90_data2.predial(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node90_data2.nd_fail = f_ChildNode
            Else
                lv_Node90_data2.nd_succ = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node90_data2.predial(0), DEF_NODE_DATA2_LEN
        
        '-------------------------------------------
        ' Sun added 2005-05-26
        Case 91
             CopyMemory lv_Node91_data2.vox_talklen, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node91_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node91_data2.vox_talklen, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        '-------------------------------------------
        ' Sun added 2005-03-15
        Case 96
            CopyMemory lv_Node96_data2.command, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            If f_nLineIndex = 0 Then
                lv_Node96_data2.nd_child = f_ChildNode
            Else
                lv_Node96_data2.nd_timeout = f_ChildNode
            End If
            CopyMemory lv_Gdata2(0), lv_Node96_data2.command, DEF_NODE_DATA2_LEN
        '-------------------------------------------
        
        Case 100
             CopyMemory lv_Node100_data2.dll_fid, lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node100_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node100_data2.dll_fid, DEF_NODE_DATA2_LEN
        Case 101
             CopyMemory lv_Node101_data2.reserved1(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node101_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node101_data2.reserved1(0), DEF_NODE_DATA2_LEN
        Case 102
             CopyMemory lv_Node102_data2.reserved1(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
             lv_Node102_data2.nd_child = f_ChildNode
             CopyMemory lv_Gdata2(0), lv_Node102_data2.reserved1(0), DEF_NODE_DATA2_LEN
        
        Case Else
        End Select
        
        SetNodeChildNode = F_GetLineCaption(m_ctlNode(lv_Index).NodeNo, f_nLineIndex)
        
        '' Update Node data
        lv_Str = ""
        For lv_SubLoop = 0 To 63
            lv_Str = lv_Str & ChrB(lv_Gdata2(lv_SubLoop))
        Next
        m_ctlNode(lv_Index).Data2 = lv_Str

        UpdateAnotherIVRRecord lv_Index
    
    End If

    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

End Function

Public Function GetNodeLogged(ByVal f_Index As Integer) As Integer
On Error Resume Next
    
    Dim lv_Node6_data1 As SData1_006
    Dim lv_Node7_data1 As SData1_007
    Dim lv_Node8_data1 As SData1_008
    Dim lv_Node9_data1 As SData1_009
    Dim lv_Node10_data1 As SData1_010
    Dim lv_Node16_data1 As SData1_016
    Dim lv_Node17_data1 As SData1_017
    Dim lv_Node18_data1 As SData1_018
    Dim lv_Node19_data1 As SData1_019
    Dim lv_Node20_data1 As SData1_020
    Dim lv_Node21_data1 As SData1_021
    Dim lv_Node22_data1 As SData1_022
    Dim lv_Node23_data1 As SData1_023
    Dim lv_Node28_data1 As SData1_028
    Dim lv_Node40_data1 As SData1_040
    Dim lv_Node41_data1 As SData1_041
    Dim lv_Node50_data1 As SData1_050
    Dim lv_Node51_data1 As SData1_051
    Dim lv_Node55_data1 As SData1_055
    Dim lv_Node60_data1 As SData1_060
    Dim lv_Node61_data1 As SData1_061
    Dim lv_Node62_data1 As SData1_062
    Dim lv_Node63_data1 As SData1_063
    Dim lv_Node69_data1 As SData1_069
    Dim lv_Node70_data1 As SData1_070
    Dim lv_Node71_data1 As SData1_071
    Dim lv_Node80_data1 As SData1_080
    Dim lv_Node90_data1 As SData1_090
    Dim lv_Node91_data1 As SData1_091
    Dim lv_Node96_data1 As SData1_096
    Dim lv_Node100_data1 As SData1_100
    Dim lv_Node101_data1 As SData1_101
    Dim lv_Node102_data1 As SData1_102
    Dim lv_Gdata1(DEF_NODE_DATA1_LEN) As Byte
    Dim lv_SubLoop As Integer
    
    GetNodeLogged = 0
    If Not m_ctlNode(f_Index) Is Nothing Then
    
        If Not m_ctlNode(f_Index).NodeDestroyed Then
        
            For lv_SubLoop = 0 To 12
                lv_Gdata1(lv_SubLoop) = AscB(MidB(m_ctlNode(f_Index).Data1, lv_SubLoop + 1, 1))
            Next
    
            Select Case m_ctlNode(f_Index).NodeNo
            Case 6
            Case 7
                 CopyMemory lv_Node7_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node7_data1.log
            Case 8
                 CopyMemory lv_Node8_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node8_data1.log
            Case 9
                 CopyMemory lv_Node9_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node9_data1.log
            Case 10
                 CopyMemory lv_Node10_data1.maincalendar, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node10_data1.log
            
            Case 16
            Case 17
            Case 18
            Case 19
            Case 20
                 CopyMemory lv_Node20_data1.reserved1, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node20_data1.log
            Case 21
                 CopyMemory lv_Node21_data1.reserved1, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node21_data1.log
            Case 22
                 CopyMemory lv_Node22_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node22_data1.log
            Case 23
                 CopyMemory lv_Node23_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node23_data1.log
            
            '-------------------------------------------------
            ' Sun added 2004-12-30
            Case 28
                 CopyMemory lv_Node28_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node28_data1.log
            '-------------------------------------------------
                 
            Case 40
                 CopyMemory lv_Node40_data1.rectime, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node40_data1.log
                                  
            Case 41
                 CopyMemory lv_Node41_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node41_data1.log
            
            Case 50
                 CopyMemory lv_Node50_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node50_data1.log
            Case 51
                 CopyMemory lv_Node51_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node51_data1.log
            
            '-------------------------------------------------
            ' Sun added 2006-12-31
            Case 55
                 CopyMemory lv_Node55_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node55_data1.log
            '-------------------------------------------------
            
            Case 60
                 CopyMemory lv_Node60_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node60_data1.log
            Case 61
                 CopyMemory lv_Node61_data1.maxwait, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node61_data1.log
            Case 62
                 CopyMemory lv_Node62_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node62_data1.log
            Case 63
                 CopyMemory lv_Node63_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node63_data1.log
            Case 69
                 CopyMemory lv_Node69_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node69_data1.log
            
            '-------------------------------------------------
            ' Sun added 2005-06-27
            Case 70
                 CopyMemory lv_Node70_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node70_data1.log
            
            Case 71
                 CopyMemory lv_Node71_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node71_data1.log
            '-------------------------------------------------
            
            Case 80
                 CopyMemory lv_Node80_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node80_data1.log
            
            Case 90
            
            '-------------------------------------------------
            ' Sun added 2005-05-26
            Case 91
                 CopyMemory lv_Node91_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node91_data1.log
            '-------------------------------------------------
            
            '-------------------------------------------------
            ' Sun added 2005-03-15
            Case 96
                 CopyMemory lv_Node96_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node96_data1.log
            '-------------------------------------------------
            
            Case 100
                 CopyMemory lv_Node100_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node100_data1.log
            Case 101
                 CopyMemory lv_Node101_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node101_data1.log
            Case 102
                 CopyMemory lv_Node102_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                 GetNodeLogged = lv_Node102_data1.log
            
            End Select
            
        End If
    
    End If
    
On Error GoTo 0
End Function

Public Function GetUserVarCount() As Integer
On Error Resume Next

    Dim lv_Node1_data1 As SData1_001
    Dim lv_Gdata1(DEF_NODE_DATA1_LEN) As Byte
    Dim lv_SubLoop As Integer
    
    GetUserVarCount = 0
    If Not m_ctlNode(2) Is Nothing Then
    
        If Not m_ctlNode(2).NodeDestroyed Then
        
            For lv_SubLoop = 0 To 12
                lv_Gdata1(lv_SubLoop) = AscB(MidB(m_ctlNode(2).Data1, lv_SubLoop + 1, 1))
            Next
            CopyMemory lv_Node1_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
            GetUserVarCount = lv_Node1_data1.uservars
            If GetUserVarCount < 0 Or GetUserVarCount > 255 Then GetUserVarCount = 0
        
        End If
        
    End If
    
On Error GoTo 0
End Function

Public Sub GetUserVarDefination(ByVal f_Index As Integer, f_Name As String, f_Type As Byte, f_Len As Byte)
On Error Resume Next

    Dim lv_Node2_data2 As SData2_002
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_SubLoop As Integer
    Dim lv_NodeID As Integer
    Dim lv_NodeIndex As Integer                     '' Sun added 2006-05-24
    Dim lv_VarIndex As Integer

    lv_NodeID = Int((f_Index - 1) / 4) + 3
    lv_VarIndex = (f_Index - 1) Mod 4
    
    '' Sun updated 2006-05-24
    lv_NodeIndex = SearchNodeIndexWithID(lv_NodeID)
    If lv_NodeIndex > 0 Then
        If Not m_ctlNode(lv_NodeIndex) Is Nothing Then
    
            If Not m_ctlNode(lv_NodeIndex).NodeDestroyed Then
            
                For lv_SubLoop = 0 To 63
                    lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_NodeIndex).Data2, lv_SubLoop + 1, 1))
                Next
                CopyMemory lv_Node2_data2.uservar(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
                
                f_Type = lv_Node2_data2.uservar(lv_VarIndex * 16 + 1)
                f_Len = lv_Node2_data2.uservar(lv_VarIndex * 16)
                For lv_SubLoop = 2 To 15
                    f_Name = f_Name & Chr(lv_Node2_data2.uservar(lv_SubLoop + lv_VarIndex * 16))
                Next
             
             End If
        
        End If
    End If
       
On Error GoTo 0
End Sub

' 设置用户变量
'
Public Sub SetUserVarDefination(ByVal f_Index As Integer, ByVal f_Name As String, ByVal f_Type As Byte, ByVal f_Len As Byte)
On Error Resume Next

    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_SubLoop As Integer
    Dim lv_NodeID As Integer
    Dim lv_NodeIndex As Integer
    Dim lv_VarIndex As Integer
    Dim lv_strName As String
    Dim lv_bytType As Byte
    Dim lv_bytLen As Byte
    Dim lv_Len As Byte

    lv_NodeID = Int((f_Index - 1) / 4) + 3
    lv_VarIndex = (f_Index - 1) Mod 4
    
    '' Sun updated 2006-05-24
    lv_NodeIndex = SearchNodeIndexWithID(lv_NodeID)
    If lv_NodeIndex > 0 Then
        If Not m_ctlNode(lv_NodeIndex) Is Nothing Then
    
            If Not m_ctlNode(lv_NodeIndex).NodeDestroyed Then
            
                For lv_SubLoop = 0 To 63
                    lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_NodeIndex).Data2, lv_SubLoop + 1, 1))
                Next
                CopyMemory Node2_Data2.uservar(0), lv_Gdata2(0), DEF_NODE_DATA2_LEN
                
                lv_bytType = Node2_Data2.uservar(lv_VarIndex * 16 + 1)
                lv_bytLen = Node2_Data2.uservar(lv_VarIndex * 16)
                lv_strName = ""
                For lv_SubLoop = 2 To 15
                    lv_strName = lv_strName & Chr(Node2_Data2.uservar(lv_SubLoop + lv_VarIndex * 16))
                Next
                
                If lv_bytType <> f_Type Or _
                        lv_bytLen <> f_Len Or _
                        lv_strName <> f_Name Then
                        
                    Node2_Data2.uservar(lv_VarIndex * 16 + 1) = f_Type
                    Node2_Data2.uservar(lv_VarIndex * 16) = f_Len
                    lv_Len = Len(f_Name)
                    If lv_Len > 14 Then lv_Len = 14
                        
                    For lv_SubLoop = 2 To 15
                        If lv_SubLoop <= lv_Len + 1 Then
                            Node2_Data2.uservar(lv_VarIndex * 16 + lv_SubLoop) = Asc(Mid(f_Name, lv_SubLoop - 1, 1))
                        Else
                            Node2_Data2.uservar(lv_VarIndex * 16 + lv_SubLoop) = 0
                        End If
                    Next
                    
                    F_NodeData lv_NodeIndex, 2
                    UpdateAnotherIVRRecord lv_NodeIndex
                    
                End If
             
             End If
        
        End If
    End If
       
On Error GoTo 0
End Sub

'' Sun modified 2002-06-12
Public Sub GetUserVarUsedStatus(ByVal f_Index As Integer, _
            f_WriteTimes As Integer, f_NodeCount As Integer, _
            f_RelateNode() As Integer, _
            f_ReadTimes As Integer, f_ReadCount As Integer, _
            f_ReadNode() As Integer)
On Error Resume Next

    Dim lv_Node6_data1 As SData1_006
    Dim lv_Node7_data1 As SData1_007
    Dim lv_Node8_data1 As SData1_008
    Dim lv_Node9_data1 As SData1_009
    Dim lv_Node10_data1 As SData1_010
    Dim lv_Node16_data1 As SData1_016
    Dim lv_Node17_data1 As SData1_017
    Dim lv_Node18_data1 As SData1_018
    Dim lv_Node19_data1 As SData1_019
    Dim lv_Node20_data1 As SData1_020
    Dim lv_Node21_data1 As SData1_021
    Dim lv_Node22_data1 As SData1_022
    Dim lv_Node23_data1 As SData1_023
    Dim lv_Node28_data1 As SData1_028
    Dim lv_Node40_data1 As SData1_040
    Dim lv_Node41_data1 As SData1_041
    Dim lv_Node50_data1 As SData1_050
    Dim lv_Node51_data1 As SData1_051
    Dim lv_Node55_data1 As SData1_055
    Dim lv_Node60_data1 As SData1_060
    Dim lv_Node61_data1 As SData1_061
    Dim lv_Node62_data1 As SData1_062
    Dim lv_Node63_data1 As SData1_063
    Dim lv_Node69_data1 As SData1_069
    Dim lv_Node70_data1 As SData1_070
    Dim lv_Node71_data1 As SData1_071
    Dim lv_Node80_data1 As SData1_080
    Dim lv_Node90_data1 As SData1_090
    Dim lv_Node91_data1 As SData1_091
    Dim lv_Node96_data1 As SData1_096
    Dim lv_Node100_data1 As SData1_100
    Dim lv_Node101_data1 As SData1_101
    Dim lv_Node102_data1 As SData1_102
    Dim lv_Gdata1(DEF_NODE_DATA1_LEN) As Byte
    Dim lv_Index As Integer
    Dim lv_SubLoop As Integer
    Dim lv_NodeTimes As Integer
    Dim lv_ReadTimes As Integer
    Dim lv_RefIndex As Integer
    Dim lv_ReadIndex As Integer
    
    f_WriteTimes = 0
    f_ReadTimes = 0
    lv_RefIndex = 0
    lv_ReadIndex = 0
    
    For lv_SubLoop = 0 To UBound(f_RelateNode)
        f_RelateNode(lv_SubLoop) = 0
    Next
    
    For lv_SubLoop = 0 To UBound(f_ReadNode)
        f_ReadNode(lv_SubLoop) = 0
    Next
        
    For lv_Index = 1 To m_nNodeID
    
        If Not m_ctlNode(lv_Index) Is Nothing Then
        
            If Not m_ctlNode(lv_Index).NodeDestroyed Then
            
                lv_NodeTimes = 0
                lv_ReadTimes = 0
                
                For lv_SubLoop = 0 To 12
                    lv_Gdata1(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_Index).Data1, lv_SubLoop + 1, 1))
                Next
        
                Select Case m_ctlNode(lv_Index).NodeNo
                Case 6
                Case 7
                     CopyMemory lv_Node7_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node7_data1.var_password = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                     If lv_Node7_data1.var_result = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                     If lv_Node7_data1.var_trytime = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                     If lv_Node7_data1.var_userid = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                     
                Case 8
                     CopyMemory lv_Node8_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node8_data1.var_password = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                     If lv_Node8_data1.var_result = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                     If lv_Node8_data1.var_trytime = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                     
                Case 9
                Case 10
                
                Case 16
                     CopyMemory lv_Node16_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node16_data1.var_id = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                
                Case 17
                     CopyMemory lv_Node17_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node17_data1.var_lang = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                Case 18
                Case 19
                Case 20
                Case 21
                     CopyMemory lv_Node21_data1.reserved1, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node21_data1.usevar = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                Case 22
                     CopyMemory lv_Node22_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node22_data1.var_key = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                
                Case 23
                     CopyMemory lv_Node23_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node23_data1.var_play = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                
                '--------------------------------------------------------------
                ' Sun added 2004-12-30
                Case 28
                     CopyMemory lv_Node28_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node28_data1.usevar = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                '--------------------------------------------------------------
                
                Case 40
                     CopyMemory lv_Node40_data1.rectime, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node40_data1.var_agent = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node40_data1.var_filename = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node40_data1.var_appfield(0) = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node40_data1.var_appfield(1) = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node40_data1.var_appfield(2) = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     
                Case 41
                     CopyMemory lv_Node41_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node41_data1.var_agent = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                
                '--------------------------------------------------------------
                ' Sun added 2006-12-31
                Case 50
                     CopyMemory lv_Node50_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node50_data1.var_faxfile = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node50_data1.var_fromno = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node50_data1.var_tono = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node50_data1.var_appfield(0) = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node50_data1.var_appfield(1) = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node50_data1.var_appfield(2) = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node50_data1.var_result = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                '--------------------------------------------------------------
                
                Case 51
                
                '--------------------------------------------------------------
                ' Sun added 2006-12-31
                Case 55
                     CopyMemory lv_Node55_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node55_data1.var_faxfile = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node55_data1.var_fromno = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node55_data1.var_tono = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node55_data1.var_extno = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node55_data1.var_appfield(0) = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node55_data1.var_appfield(1) = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node55_data1.var_appfield(2) = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node55_data1.var_result = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                '--------------------------------------------------------------
                
                Case 60
                     CopyMemory lv_Node60_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node60_data1.var_key = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                     
                Case 61
                Case 62
                     CopyMemory lv_Node62_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node62_data1.usevar = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                Case 63
                     CopyMemory lv_Node63_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node63_data1.usevar = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     
                Case 69
                     CopyMemory lv_Node69_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node69_data1.usevar = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                
                '--------------------------------------------------------------
                ' Sun added 2005-06-27
                Case 70
                     CopyMemory lv_Node70_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node70_data1.usevar = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     
                     '' Sun added 2007-04-11
                     If lv_Node70_data1.var_result = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                    
                Case 71
                     CopyMemory lv_Node71_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node71_data1.usevar = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                '--------------------------------------------------------------
                
                Case 80
                     CopyMemory lv_Node80_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node80_data1.usevar = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                
                Case 90
                     CopyMemory lv_Node90_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node90_data1.usevar = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node90_data1.resultvar = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                
                '-------------------------------------------------
                ' Sun added 2005-05-26
                Case 91
                     CopyMemory lv_Node91_data1.timeout, lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node91_data1.var_cardno = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node91_data1.var_telno = f_Index Then lv_ReadTimes = lv_ReadTimes + 1
                     If lv_Node91_data1.var_connectlength = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                '-------------------------------------------------
                
                '--------------------------------------------------------------
                ' Sun added 2005-03-15
                Case 96
                '--------------------------------------------------------------
                
                Case 100
                Case 101
                
                Case 102
                     CopyMemory lv_Node102_data1.reserved1(0), lv_Gdata1(0), DEF_NODE_DATA1_LEN
                     If lv_Node102_data1.var_chg = f_Index Then lv_NodeTimes = lv_NodeTimes + 1
                               
                End Select
                
                If lv_NodeTimes > 0 Then
                    f_WriteTimes = f_WriteTimes + lv_NodeTimes
                    If lv_RefIndex <= UBound(f_RelateNode) Then
                        f_RelateNode(lv_RefIndex) = lv_Index
                        lv_RefIndex = lv_RefIndex + 1
                    End If
                End If
                
                If lv_ReadTimes > 0 Then
                    f_ReadTimes = f_ReadTimes + lv_ReadTimes
                    If lv_ReadIndex <= UBound(f_ReadNode) Then
                        f_ReadNode(lv_ReadIndex) = lv_Index
                        lv_ReadIndex = lv_ReadIndex + 1
                    End If
                End If
                
            End If
        
        End If
    
    Next
    
    f_NodeCount = lv_RefIndex
    f_ReadCount = lv_ReadIndex

On Error GoTo 0
End Sub

Public Function FindNodeIDAtPoint(ByVal x As Single, ByVal y As Single) As Integer
On Error Resume Next

    Dim lv_Index As Integer

    For lv_Index = 1 To m_nNodeID
        If Not m_ctlNode(lv_Index) Is Nothing Then
            If Not m_ctlNode(lv_Index).NodeDestroyed Then
                If m_ctlNode(lv_Index).NodeNo <> 255 Then
                
                    '' Sun added 2007-12-01
                    If m_ctlNode(lv_Index).Visible = True Then
                
                        If (m_ctlNode(lv_Index).Left <= x And x <= m_ctlNode(lv_Index).Left + m_ctlNode(lv_Index).Width) _
                           And (m_ctlNode(lv_Index).Top <= y And y <= m_ctlNode(lv_Index).Top + m_ctlNode(lv_Index).Height) Then
                           FindNodeIDAtPoint = m_ctlNode(lv_Index).NodeID
                           Exit Function
                        End If
                        
                    End If
                End If
            End If
        End If
    Next
    
    FindNodeIDAtPoint = 0
    
End Function

Public Function FindLinesOnNode(f_Node As CNode, Optional f_Pos As Integer = 1) As Integer
On Error Resume Next

    Dim lv_Data2 As SData2_255
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_Index As Integer
    Dim lv_SubLoop As Integer
    
    If f_Pos <= 0 Then f_Pos = 1
    For lv_Index = f_Pos To m_nNodeID
        If Not m_ctlNode(lv_Index) Is Nothing Then
            If Not m_ctlNode(lv_Index).NodeDestroyed Then
                
                '' Sun updated 2007-12-01
                If m_ctlNode(lv_Index).NodeNo = 255 And Not m_ctlNode(lv_Index).IsSelected And _
                   m_ctlNode(lv_Index).Visible = True Then
                        
                    For lv_SubLoop = 0 To 63
                        lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_Index).Data2, lv_SubLoop + 1, 1))
                    Next
                    CopyMemory lv_Data2.StartNode, lv_Gdata2(0), DEF_NODE_DATA2_LEN
                    'Debug.Print lv_Gdata2(0), lv_Gdata2(1)
                    
                    If lv_Data2.StartNode = f_Node.NodeID Then
                        
                        FindLinesOnNode = m_ctlNode(lv_Index).NodeID
                        f_Pos = lv_Index
                        
                        m_ctlNode(lv_Index).Line_X1 = f_Node.Left + f_Node.Width / 2
                        m_ctlNode(lv_Index).Line_Y1 = f_Node.Top + f_Node.Height / 2
                        m_ctlNode(lv_Index).MoveImageWithLine
                    
                        ''' 保存节点
                        UpdateAnotherIVRRecord lv_Index
    
                        Exit Function
                        
                    ElseIf lv_Data2.EndNode = f_Node.NodeID Then
                        
                        FindLinesOnNode = m_ctlNode(lv_Index).NodeID
                        f_Pos = lv_Index
                        
                        m_ctlNode(lv_Index).Line_X2 = f_Node.Left + f_Node.Width / 2
                        m_ctlNode(lv_Index).Line_Y2 = f_Node.Top + f_Node.Height / 2
                        m_ctlNode(lv_Index).MoveImageWithLine
                    
                        ''' 保存节点
                        UpdateAnotherIVRRecord lv_Index
    
                        Exit Function
                        
                    End If
                    
                End If
            End If
        End If
    Next
    
    f_Pos = -1
    FindLinesOnNode = 0
    
End Function

Public Function GetFlowDataFileInfo(f_strFileName As String) As Byte
On Error GoTo BackDoor

    Dim FileNumber
    Dim lv_Header As Mdlnode.SCFFileHeader
    Dim lv_FileCaption As String
    Dim lv_bytPID As Byte
    Dim lv_sSQL As String
    
    ' 缺省返回值
    GetFlowDataFileInfo = 0
    
    '*****************************************************
    '* Open File for Read
    '*****************************************************
    If Not CheckExistFile(f_strFileName) Then
        Call MsgBox(LoadNationalResString(1669) & " ‘" & f_strFileName & "' " & LoadNationalResString(1125), vbOK + vbExclamation, App.Title)
        Exit Function
    End If
    'Debug.Print "Open File for Read: " & f_strFileName
    FileNumber = FreeFile   ' 取得未使用的文件号

    Open f_strFileName For Binary Access Read Lock Write As #FileNumber
    If Err Then
        Err.Clear
        MsgBox Err.Description
        MsgBox LoadNationalResString(1110) & f_strFileName & LoadNationalResString(1111), vbCritical + vbOKOnly, App.Title
        Exit Function
    Else
        '*****************************************************
        '* place loading flow information codes here         *
        '*****************************************************
        ''' Read File Header
        Get #FileNumber, , lv_Header
        If Err Then
            Err.Clear
            MsgBox LoadNationalResString(1126) & f_strFileName, vbCritical + vbOKOnly, App.Title
            GoTo BackDoor
        Else
            With lv_Header
                lv_FileCaption = LoadNationalResString(1669) & " ‘" & f_strFileName & "' " & LoadNationalResString(1128) & " :" & vbCrLf
                lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1114) & " :" & Str(.P_ID) & vbCrLf
                lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1116) & " :" & EraseInvalidCharacters(Trim(.P_Name)) & vbCrLf
                lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1117) & " :" & EraseInvalidCharacters(Trim(.P_Description)) & vbCrLf
                lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1105) & " :" & EraseInvalidCharacters(Trim(.P_Auther)) & vbCrLf
                lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1118) & " :" & EraseInvalidCharacters(Trim(.P_User)) & vbCrLf
                lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1119) & " :" & EraseInvalidCharacters(Trim(.P_Version)) & vbCrLf
                lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1106) & " :" & Str(.P_CreateTime) & vbCrLf
                lv_FileCaption = lv_FileCaption & Space(4) & LoadNationalResString(1120) & " :" & Str(.P_ModifyTime) & vbCrLf
                lv_FileCaption = lv_FileCaption & LoadNationalResString(1129)
            End With
            If MsgBox(lv_FileCaption, vbYesNo + vbQuestion, App.Title) = vbNo Then
                GoTo BackDoor
            End If
        End If
    End If

    '判断流程是否在数据库里存在
    lv_bytPID = lv_Header.P_ID
    Do While True
        
        If CallFlowID > 0 Then
            DestroyAllNodes
        End If
        
        '连接数据库
        OpenIvrRecordSet lv_bytPID
    
        ' 判断流程是否在数据库里存在
        If IsNewCallFlow Then
            
            '' 增加新流程
            If AddCallFlowIntoList(lv_bytPID, _
                    EraseInvalidCharacters(Trim(lv_Header.P_Name)), _
                    EraseInvalidCharacters(Trim(lv_Header.P_Description)), _
                    EraseInvalidCharacters(Trim(lv_Header.P_Auther)), _
                    EraseInvalidCharacters(Trim(lv_Header.P_User)), _
                    EraseInvalidCharacters(Trim(lv_Header.P_Version))) Then
                
                m_rsFS("P_modifytime") = Trim(lv_Header.P_ModifyTime)                   '修改时间
                m_rsFS("P_createtime") = Trim(lv_Header.P_CreateTime)                   '创建时间
                m_rsFS.Update
                
                Exit Do
            
            End If
           
        Else
            
            '' 覆盖或者另外编号
            frmImportConfirm.Show vbModal
            
            If CallFlowID = 0 Then
            '' 取消
                GoTo BackDoor
            ElseIf CallFlowID = lv_bytPID Then
            '' 覆盖
                m_rsFS.Delete
#If Programnew = 0 Then
                lv_sSQL = "DELETE FROM tbivrprogramNew WHERE P_ID= " & Str(lv_bytPID)
#Else
                lv_sSQL = "DELETE FROM tbivrprogram WHERE P_ID= " & Str(lv_bytPID)
#End If
                m_adoCN.Execute lv_sSQL, -1, adCmdText
'                If m_rsRS.RecordCount > 0 Then
'                    m_rsRS.MoveFirst
'                    Do While Not m_rsRS.EOF
'                        m_rsRS.Delete
'                        m_rsRS.MoveFirst
'                    Loop
'                End If
            Else
            '' 另外编号
                lv_bytPID = CallFlowID
            End If
       
        End If
    
    Loop

    GetFlowDataFileInfo = lv_bytPID

BackDoor:
    
    ' Close File
    Close #FileNumber

    If Err Then
        Debug.Print Err.Description
    End If
    
    'Mike added @ 2008-7-4
    Call WriteLogMessage(Err.Number, enu_Error, "Get callflow date file information failed when import a callflow!", Err.Description)
    
End Function

Public Function ImportFlowData(f_strFileName As String) As Boolean
'On Error Resume Next
On Error GoTo BackDoor

    Dim FileNumber
    Dim lv_Header As Mdlnode.SCFFileHeader
    Dim lv_Record As Mdlnode.SCFFileRecord
    Dim lv_loop As Long
    Dim lv_SubLoop As Integer
    Dim lv_StrD1, lv_StrD2
    
    ' 缺省返回值
    ImportFlowData = False
    
    '*****************************************************
    '* Open File for Read
    '*****************************************************
    If Not CheckExistFile(f_strFileName) Then
        Call MsgBox(LoadNationalResString(1669) & " ‘" & f_strFileName & "' " & LoadNationalResString(1125), vbOK + vbExclamation, App.Title)
        Exit Function
    End If
    'Debug.Print "Open File for Read: " & f_strFileName
    FileNumber = FreeFile   ' 取得未使用的文件号

    Open f_strFileName For Binary Access Read Lock Write As #FileNumber
    If Err Then
        Err.Clear
        MsgBox Err.Description
        MsgBox LoadNationalResString(1110) & f_strFileName & LoadNationalResString(1111), vbCritical + vbOKOnly, App.Title
        Exit Function
    Else
        '*****************************************************
        '* place loading flow information codes here         *
        '*****************************************************
        ''' Read File Header
        Get #FileNumber, , lv_Header
        If Err Then
            Err.Clear
            MsgBox LoadNationalResString(1126) & f_strFileName, vbCritical + vbOKOnly, App.Title
            GoTo BackDoor
        End If
    End If
        
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    '*****************************************************
    '* Add loading nodes database codes here...          *
    '*****************************************************
    If lv_Header.P_RCN > 0 Then
        
        For lv_loop = 1 To lv_Header.P_RCN
            
            ''' Write File Record
            Get #FileNumber, , lv_Record
            If Err Then
                Err.Clear
                MsgBox LoadNationalResString(1126) & f_strFileName & "’ " & LoadNationalResString(1123) & Trim(Str(lv_loop)), vbCritical + vbOKOnly, App.Title
                GoTo BackDoor
            End If
            
            m_rsRS.AddNew
            If Err Then
                Err.Clear
                MsgBox LoadNationalResString(1131) & " '" & f_strFileName & "' " & LoadNationalResString(1123) & Trim(Str(lv_loop)), vbCritical + vbOKOnly, App.Title
                GoTo BackDoor
            End If

            With lv_Record
                
                m_rsRS("p_id") = CallFlowID
                m_rsRS("n_id") = .N_ID
                m_rsRS("n_no") = .N_NO
                m_rsRS("n_page") = .N_Page
                m_rsRS("n_left") = .N_Left
                m_rsRS("n_top") = .N_Top
                m_rsRS("n_height") = .N_Height
                m_rsRS("n_width") = .N_Width
                m_rsRS("n_description") = EraseInvalidCharacters(Trim(.N_Description))
                
                '' Sun added 2008-01-18
                m_rsRS("n_tag") = EraseInvalidCharacters(Trim(.N_Tag))
                
                lv_StrD1 = ""
                For lv_SubLoop = 0 To 12
                    m_rsRS("N_D1_" & CStr(lv_SubLoop + 1)) = .N_Data1(lv_SubLoop)
                Next
                lv_StrD2 = ""
                For lv_SubLoop = 0 To 63
                    m_rsRS("N_D2_" & CStr(lv_SubLoop + 1)) = .N_Data2(lv_SubLoop)
                Next
            End With
            
            m_rsRS.Update
        
        Next
        
    End If
    
    'Mike added @ 2008-7-4
    Call WriteLogMessage(0, enu_Information, "ImportFlowData(" & f_strFileName & ") successful!")
    
    ImportFlowData = True

BackDoor:
    
    ' Close File
    Close #FileNumber

    If Err Then
        Debug.Print "ImportFlowData(" & f_strFileName & ") got error: " & Err.Description
    End If
    
    'Mike added @ 2008-7-4
    Call WriteLogMessage(Err.Number, enu_Error, "ImportFlowData(" & f_strFileName & ") got error,", Err.Description)
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

End Function

Public Function ExportFlowData(f_strFileName As String) As Boolean
On Error Resume Next

    Dim FileNumber
    Dim lv_loop As Long
    Dim lv_RCN As Long
    Dim lv_Sub As Integer
    Dim lv_StrD1 As String, lv_StrD2 As String
    
    ' 缺省返回值
    ExportFlowData = False
        
    '*****************************************************
    '* Open File for Write
    '*****************************************************
    If CheckExistFile(f_strFileName) Then
    
        If MsgBox(LoadNationalResString(1669) & " '" & f_strFileName & LoadNationalResString(1109), vbYesNo + vbQuestion, App.Title) = vbNo Then
            Exit Function
        End If
        '' Kill the File
        Call RemoveFile(f_strFileName)
    End If
    Debug.Print "Open File for Written: " & f_strFileName
    FileNumber = FreeFile   ' 取得未使用的文件号
    Open f_strFileName For Binary Access Write Lock Write As #FileNumber
    If Err Then
        Err.Clear
        MsgBox LoadNationalResString(1110) & f_strFileName & LoadNationalResString(1111), vbCritical + vbOKOnly, App.Title
        Exit Function
    End If
        
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
        
    '*****************************************************
    '* place saving flow information codes here          *
    '*****************************************************
    Dim lv_Header As Mdlnode.SCFFileHeader

    If IsNewCallFlow Then
        MsgBox LoadNationalResString(1112) & Trim(Str(CallFlowID)), vbCritical + vbOKOnly, App.Title
        GoTo BackDoor
    End If
    
    ''' Fill File Header
    m_rsRS.MoveLast
    lv_RCN = m_rsRS.RecordCount
    With lv_Header
    
        .P_ID_Name = LoadNationalResString(1114)
        .P_ID = m_rsFS("p_id") Mod 256
        
        .P_Type_Name = LoadNationalResString(1115)
        .P_Type = Asc("P")
        
        .P_Name_Name = LoadNationalResString(1116)
        If Not IsNull(m_rsFS("p_name")) Then
            .P_Name = Left(m_rsFS("p_name"), Len(.P_Name))
        End If
        
        .P_Description_Name = LoadNationalResString(1117)
        If Not IsNull(m_rsFS("p_description")) Then
            .P_Description = Left(m_rsFS("p_description"), Len(.P_Description))
        End If
        
        .P_Auther_Name = LoadNationalResString(1105)
        If Not IsNull(m_rsFS("p_auther")) Then
            .P_Auther = Left(m_rsFS("p_auther"), Len(.P_Auther))
        End If
        
        .P_User_Name = LoadNationalResString(1118)
        If Not IsNull(m_rsFS("p_user")) Then
            .P_User = Left(m_rsFS("p_user"), Len(.P_User))
        End If
        
        .P_Version_Name = LoadNationalResString(1119)
        If Not IsNull(m_rsFS("p_version")) Then
            .P_Version = Left(m_rsFS("p_version"), Len(.P_Version))
        End If
        
        .P_CreateTime_Name = LoadNationalResString(1106)
        .P_CreateTime = m_rsFS("P_createtime")
        
        .P_ModifyTime_Name = LoadNationalResString(1120)
        .P_ModifyTime = m_rsFS("P_modifytime")
        
        .P_RCN_Name = "RecordCnt"
        .P_RCN = lv_RCN
        
        .Reserved = ""
        
    End With
    
    ''' Write File Header
    Put #FileNumber, , lv_Header
    If Err Then
        Err.Clear
        MsgBox LoadNationalResString(1121) & "'" & f_strFileName & "' " & LoadNationalResString(1122), vbCritical + vbOKOnly, App.Title
        GoTo BackDoor
    End If
    
    '*****************************************************
    '* Add saving nodes database codes here...           *
    '*****************************************************
    Dim lv_Record As Mdlnode.SCFFileRecord
    If lv_RCN > 0 Then
        
        m_rsRS.MoveFirst
        Debug.Print m_rsRS.Source
        For lv_loop = 1 To lv_RCN
            
            With lv_Record
                
                .N_ID = m_rsRS("n_id")
                .N_NO = m_rsRS("n_no")
                .N_Page = m_rsRS("n_page")
                .N_Left = m_rsRS("n_left")
                .N_Top = m_rsRS("n_top")
                .N_Height = m_rsRS("n_height")
                .N_Width = m_rsRS("n_width")
                If Not IsNull(m_rsRS("n_description")) Then
                    .N_Description = Left(Trim(m_rsRS("n_description")), Len(.N_Description))
                Else
                    .N_Description = ""
                End If
                
                '' Sun added 2008-01-18
                If Not IsNull(m_rsRS("n_tag")) Then
                    .N_Tag = Left(Trim(m_rsRS("n_tag")), 7)
                Else
                    .N_Tag = ""
                End If
'
'                For lv_Sub = 0 To 12
'                    .N_Data1(lv_Sub) = CByte(0)
'                    Debug.Print "1.1" & .N_Data1(lv_Sub)
'                Next
'                For lv_Sub = 0 To 63
'                    .N_Data2(lv_Sub) = CByte(0)
'                    Debug.Print "1.2" & .N_Data1(lv_Sub)
'                Next
'Update Jeremy 2004-07-08
           '     If Not IsNull(m_rsRS.Fields("n_data1")) Then
               '     lv_StrD1 = LeftB(m_rsRS.Fields("n_data1"), 13)
                    For lv_Sub = 0 To 12
                        If Not IsNull(m_rsRS.Fields("N_D1_" & CStr(lv_Sub + 1))) Then
                            .N_Data1(lv_Sub) = m_rsRS.Fields("N_D1_" & CStr(lv_Sub + 1)) 'AscB(MidB(lv_StrD1, lv_Sub + 1, 1))
                        Else
                            .N_Data1(lv_Sub) = 0
                        End If
                    Next
            '    End If
             '   If Not IsNull(m_rsRS.Fields("n_data2")) Then
                '    lv_StrD2 = LeftB(m_rsRS.Fields("n_data2"), DEF_NODE_DATA2_LEN)
                    For lv_Sub = 0 To 63
                        If Not IsNull(m_rsRS.Fields("N_D2_" & CStr(lv_Sub + 1))) Then
                            .N_Data2(lv_Sub) = m_rsRS.Fields("N_D2_" & CStr(lv_Sub + 1)) 'AscB(MidB(lv_StrD2, lv_Sub + 1, 1))
                        Else
                            .N_Data2(lv_Sub) = 0
                        End If
                    Next
              '  End If
'Update end
                '.Reserved = ""
                
                Err.Clear
                
                ''Debug.Print "N_ID :" & Str(RS_RS!N_ID) & "  Data1:" & .N_Data1
                ''Debug.Print "   Data2:" & .N_Data2
                        
            End With
            
            ''' Write File Record
            Put #FileNumber, , lv_Record
            If Err Then
                Debug.Print Err.Description
                Err.Clear
                MsgBox LoadNationalResString(1121) & " '" & f_strFileName & "' " & LoadNationalResString(1123) & Trim(Str(lv_loop)), vbCritical + vbOKOnly, App.Title
                GoTo BackDoor
            End If
            
            m_rsRS.MoveNext
            
        Next
        
    End If
    
    'Mike Added @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Export Callflow Data to the file successful !", "File Path:" & f_strFileName)
    ExportFlowData = True
    
BackDoor:
    ' Close File
    Close #FileNumber
    
    'Mike Added @ 2008-7-7
    Call WriteLogMessage(Err.Number, enu_Error, "Failed to export Callflow Data to the file:" & f_strFileName & "!", Err.Description)
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True

End Function

'Michael Added @ 2008-6-27 for "Add New Resource Item"
Public Function AddResourceItem(ByRef lv_AbsPos As Integer, ByVal r_ID As Integer, ByVal r_LID As Integer, ByVal r_Type As Byte, _
                        ByVal r_Path As String, ByVal r_Desc As String, ByVal r_Note As String) As Boolean
On Error GoTo Catch

    If IsNewResource(r_ID, r_LID) Then
        m_rsResource.AddNew
        m_rsResource("P_ID") = m_nResourceID                     '项目ID
        m_rsResource("L_ID") = r_LID                             '资源语言
        m_rsResource("R_ID") = r_ID                              '资源编号
        m_rsResource("R_Type") = r_Type                          '资源类型
        m_rsResource("R_Description") = r_Desc                   '资源描述
        m_rsResource("R_Path") = r_Path                          '资源路径
        m_rsResource("R_Note") = r_Note                          '备注
        m_rsResource("ModifyTime") = Now()                       '修改时间
        m_rsResource("CreateTime") = Now()                       '创建时间
        m_rsResource.Update
        
        lv_AbsPos = m_rsResource.AbsolutePosition
        'm_blnSavedMark = False
        AddResourceItem = True
        'Mike Added @ 2008-7-7
        Call WriteLogMessage(0, enu_Information, "Add resource item:" & r_ID & " to Project:" & m_nResourceID)
    Else
        AddResourceItem = False
    End If
    
Exit Function
Catch:
    AddResourceItem = False
    Call WriteLogMessage(Err.Number, enu_Error, "Failed to add resource item:" & r_ID & " to Project:" & m_nResourceID, Err.Description)
End Function

'Michael Added @ 2008-6-27 for "Update Resource item"
Function UpdateResourceItem(ByVal Key As String, ByVal RID As Integer, ByVal LID As Integer, _
                            ByVal RType As Byte, ByVal path As String, ByVal Description As String, ByVal Note As String) As Boolean
On Error GoTo Catch
        
    m_rsResource.AbsolutePosition = Right(Key, Len(Key) - 1)
    'm_rsResource.Filter = ("R_ID = " & RID)
    
    If IsNull(m_rsResource("L_ID")) Or Trim(m_rsResource("L_ID")) <> LID Then
        m_rsResource("L_ID") = LID
        m_blnSavedMark = False
    End If
    
    If IsNull(m_rsResource("R_Type")) Or Trim(m_rsResource("R_Type")) <> RType Then
        m_rsResource("R_Type") = RType
        m_blnSavedMark = False
    End If
    
    If IsNull(m_rsResource("R_Path")) Or Trim(m_rsResource("R_Path")) <> path Then
        m_rsResource("R_Path") = path
        m_blnSavedMark = False
    End If
    
    If IsNull(m_rsResource("R_Description")) Or Trim(m_rsResource("R_Description")) <> Description Then
        m_rsResource("R_Description") = Description
        m_blnSavedMark = False
    End If
    
    If IsNull(m_rsResource("R_Note")) Or Trim(m_rsResource("R_Note")) <> Note Then
        m_rsResource("R_Note") = Note
        m_blnSavedMark = False
    End If
    
    If m_blnSavedMark = False Then
        m_rsResource("ModifyTime") = Now()
        m_rsResource.Update
        m_rsResource.Requery
    End If
    
    UpdateResourceItem = True
    'Mike @ 2008-7-7
    Call WriteLogMessage(0, enu_Information, "Update Resource item:" & RID & " in Project:" & m_rsResource("P_ID"))
    m_blnSavedMark = True
Exit Function

Catch:
    'Mike added @ 2008-7-7
    Call WriteLogMessage(Err.Number, enu_Error, "Failed Update Resource item:" & RID & " in Project:" & m_rsResource("P_ID"), Err.Description)
End Function

Public Sub ClearSet()
    Set m_rsResource = Nothing
End Sub

'Michael Added @ 2008-6-27
Public Function IsNewResource(iRID As Integer, iLID As Integer) As Boolean
On Error Resume Next
    IsNewResource = True
    Dim lv_count As Integer
    If m_rsResource.RecordCount > 0 Then
        m_rsResource.MoveFirst
        For lv_count = 1 To m_rsResource.RecordCount
            'michael modified @ 2008-10-7
            If (m_rsResource("R_ID") = iRID) And (m_rsResource("L_ID") = iLID) Then
                IsNewResource = False
                Exit For
            End If
            
            m_rsResource.MoveNext
        Next lv_count
    End If
End Function

'Mike added 2008-8-20
Public Sub EmptyVarNodeByIndex(iNodeIndex As Integer, iVarIndex As Integer)
On Error Resume Next
    Dim lv_NodeID As Integer
    Dim lv_NodeIndex As Integer
    Dim lv_SubLoop As Integer
    
    lv_NodeID = Int((iNodeIndex - 1) / 4) + 3
    lv_NodeIndex = SearchNodeIndexWithID(lv_NodeID)
    
    
    Node2_Data2.uservar(iVarIndex * 16) = 0
    Node2_Data2.uservar(iVarIndex * 16 + 1) = 2
    'empty name
    For lv_SubLoop = 2 To 15
        If lv_SubLoop = 2 Then
            Node2_Data2.uservar(iVarIndex * 16 + lv_SubLoop) = Asc("0")
        Else
            Node2_Data2.uservar(iVarIndex * 16 + lv_SubLoop) = " "
        End If
    Next
                    
    F_NodeData lv_NodeIndex, 2
    UpdateAnotherIVRRecord lv_NodeIndex
        
End Sub

'' Mike added 2008-8-28
'' 切换流程时更新相关系统节点->Node0_Data2
Public Sub UpdateSystemNode()
On Error Resume Next

    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_loop As Integer
    Dim lv_SubLoop As Integer
    ReDim Preserve m_ctlNode(m_nNodeID) As New CNode
    
    For lv_loop = 1 To m_nNodeID
        Select Case m_ctlNode(lv_loop).NodeNo
        Case 0:
            If m_ctlNode(lv_loop).Data2 <> "" Then
                For lv_SubLoop = 0 To 63
                    lv_Gdata2(lv_SubLoop) = AscB(MidB(m_ctlNode(lv_loop).Data2, lv_SubLoop + 1, 1))
                Next
                CopyMemory Node0_Data2.key_repeat, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            End If
            Exit For
        End Select
    Next

End Sub

