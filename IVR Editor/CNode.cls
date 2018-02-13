VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CNode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private m_objRootWnd As Object              '' 菜单窗口
Private m_objParent As Object               '' 母窗口
Private m_objContainer As Object            '' 容器

Private WithEvents ctlImage As VB.Image
Attribute ctlImage.VB_VarHelpID = -1
Private WithEvents ctlLine As VB.Line
Attribute ctlLine.VB_VarHelpID = -1
Private ctlCaption As VB.Label
'Mike Added @ 2008-1-29
Private ctlTag As VB.Label

Private WithEvents ctlSelected As VB.Image
Attribute ctlSelected.VB_VarHelpID = -1

Private m_FlowID As Byte
Private m_NodeIndex As Integer
Private m_Nodeid As Integer
Private m_NodeNo As Byte
Private m_Width As Integer
Private m_Height As Integer
Private m_Description As String
Private m_InPage As Byte
Private m_Data1 As String
Private m_Data2 As String
Private m_tooltiptext As String
Private m_blnRootNode As Boolean            '' Whether is Root Node
Private m_blnSelected As Boolean            '' Whether is Selected
Private m_bytModifyFlag As Byte             '' Node Modify Flag，参见“操作定义”
'Mike add 2008-1-25
'Private m_NodeTag  As String                '' Node统计标签(need not anymore)

'' Sun added 2002-06-12
Private m_blnDestroy As Boolean

Public Function Create(objRootWnd As Object, objParent As Object, objContainer As Object, nIndex As Integer) As Boolean
On Error Resume Next

    Dim xHandle As Single, yHandle As Single

    Set m_objRootWnd = objRootWnd
    Set m_objParent = objParent
    Set m_objContainer = objContainer

    If nIndex <= 0 Then
        Create = False
        Exit Function
    End If
        
    ' Create node
    Set ctlLine = Nothing
    Set ctlCaption = objParent.Controls.Add("VB.Label", "Node_Caption" & Format(nIndex), objContainer)
    If Err Then Debug.Print Err.Description & " on Node_Caption" & Format(nIndex): Err.Clear
    
    'Mike Added @ 2008-1-29 Node Tag
    Set ctlTag = objParent.Controls.Add("VB.Label", "Node_Tag" & Format(nIndex), objContainer)
    If Err Then Debug.Print Err.Description & " on Node_Tag" & Format(nIndex): Err.Clear
    
    Set ctlSelected = objParent.Controls.Add("VB.Image", "Node_Selected" & Format(nIndex), objContainer)
    If Err Then Debug.Print Err.Description & " on Node_Selected" & Format(nIndex): Err.Clear
    
    Set ctlImage = objParent.Controls.Add("VB.Image", "Node" & Format(nIndex), objContainer)
    If Err Then Debug.Print Err.Description & " on NodeImage" & Format(nIndex): Err.Clear
    If ctlImage Is Nothing Then
        MsgBox "Create node failed!", vbCritical
        Create = False
        Exit Function
    End If
        
    ctlSelected.ZOrder
    ctlSelected.Visible = False
    
    '' Sun replaced 2008-02-18
    ''ctlSelected.Picture = LoadPicture(App.path & "\Bitmaps\Selected.ico")
    ctlSelected.Picture = m_objParent.imgSelect.Picture
    If Err Then Debug.Print Err.Description & " on Selected.ico" & Format(nIndex): Err.Clear
    
    ctlSelected.MousePointer = 99
    ctlSelected.Appearance = 0
    
    '' Sun replaced 2008-02-18
    'ctlSelected.MouseIcon = LoadPicture(App.path & "\Bitmaps\HypeLink.cur")
    ctlSelected.MouseIcon = m_objParent.imgHyperLink.Picture
    If Err Then Debug.Print Err.Description & " on HypeLink.cur" & Format(nIndex): Err.Clear
    ctlSelected.Stretch = True
    
    '' Sun change "gif" to "bmp", 2008-02-18
    ' Set node properties
    ' DefaNode.bmp is default node image
    '' Sun replaced 2008-02-18
    'ctlImage.Picture = LoadPicture(App.path & "\Bitmaps\DefaNode.bmp")
    ctlImage.Picture = m_objParent.imgDefaNode.Picture
    If Err Then Debug.Print Err.Description & " on DefaNode.bmp" & Format(nIndex)
    ' default cursor pointer is hand shape
    ctlImage.MousePointer = 99
    ctlImage.Appearance = 0
    
    '' Sun replaced 2008-02-18
    'ctlImage.MouseIcon = LoadPicture(App.path & "\Bitmaps\HypeLink.cur")
    ctlImage.MouseIcon = m_objParent.imgHyperLink.Picture
    If Err Then Debug.Print Err.Description & " on HypeLink2.cur" & Format(nIndex): Err.Clear
    
    ctlImage.Stretch = True
    
    '' Sun added 2002-03-28 for Preoper Display
    ctlCaption.Visible = False
    ctlCaption.AutoSize = True
    ctlCaption.Appearance = 0       ' Flat
    ctlCaption.BorderStyle = 1
    ctlCaption.BackColor = &H80FFFF
    
    ' Mike Added @ 2008-1-29 for Tag Display
    ctlTag.Visible = False
    ctlTag.AutoSize = True
    ctlTag.Appearance = 0  'Flat
    ctlTag.BorderStyle = 1
    ctlTag.ForeColor = &HFFFFFF  'Write
    ctlTag.BackColor = &HFF0000  'Blue
    
    m_blnSelected = False
    m_blnRootNode = False
    m_blnDestroy = False
    m_NodeIndex = nIndex
    Create = True

End Function

Public Function AddLine() As Boolean
On Error Resume Next

    Dim lv_Data2 As SData2_255
    Dim lv_Gdata2(DEF_NODE_DATA2_LEN) As Byte
    Dim lv_loop As Integer
    Dim lv_Str As String
    
    Set ctlLine = m_objParent.Controls.Add("VB.Line", "LineNode" & Format(m_NodeIndex), m_objContainer)
    If ctlLine Is Nothing Then
        MsgBox "Add line failed!", vbCritical
        AddLine = False
        Exit Function
    End If
    ctlLine.x1 = Me.Left - Me.Width * (1 - Def_GOLDEN_DIV)
    ctlLine.x2 = Me.Left + Me.Width * Def_GOLDEN_DIV
    ctlLine.y1 = Me.Top - Me.Height * (1 - Def_GOLDEN_DIV)
    ctlLine.y2 = Me.Top + Me.Height * Def_GOLDEN_DIV
    
    If Not IsNull(Me.Data2) Then
    ''Debug.Print "Data2=<" & Me.Data2 & ">"
        If Trim(Me.Data2) <> "" Then
            For lv_loop = 0 To 63
                lv_Gdata2(lv_loop) = AscB(MidB(Me.Data2, lv_loop + 1, 1))
            Next
            CopyMemory lv_Data2.StartNode, lv_Gdata2(0), DEF_NODE_DATA2_LEN
            Me.Line_Color = lv_Data2.Color
            Me.Line_Style = lv_Data2.Style
            Me.Line_Width = lv_Data2.Width
        Else
            lv_Str = ""
            lv_Data2.Color = 0
            lv_Data2.Style = 1
            lv_Data2.Width = 1
            lv_Data2.StartNode = 0
            lv_Data2.EndNode = 0
            lv_Data2.Index = 0
            lv_Data2.reserved1(0) = 0
            CopyMemory lv_Gdata2(0), lv_Data2.StartNode, DEF_NODE_DATA2_LEN
            For lv_loop = 0 To 63
                lv_Str = lv_Str & ChrB(lv_Gdata2(lv_loop))
            Next
            Me.Data2 = lv_Str
        End If
    End If
    ctlLine.Visible = True
    
    ' Show an arrow
    ctlImage.Height = Screen.TwipsPerPixelY * 8
    ctlImage.Width = Screen.TwipsPerPixelX * 8
    
    AddLine = True
    
End Function

Public Sub Destroy()
On Error Resume Next

    If Not m_objParent Is Nothing Then
        
        m_objParent.Controls.Remove ctlImage
        If Err Then Debug.Print Err.Description
        
        m_objParent.Controls.Remove ctlCaption
        If Err Then Debug.Print Err.Description
        
        'Mike Add @ 2008-1-29
        m_objParent.Controls.Remove ctlTag
        If Err Then Debug.Print Err.Description
        
        m_objParent.Controls.Remove ctlSelected
        If Err Then Debug.Print Err.Description
        
        If Not ctlLine Is Nothing Then
            m_objParent.Controls.Remove ctlLine
            If Err Then Debug.Print Err.Description
        End If
        
    End If
    
    'Mike Added @ 2008-7-7
    'Call WriteLogMessage(0, enu_Information, "Delete Node, Node Number:" & NodeNo)
    
    NodeNo = 0
    InPage = 0
    
    '' Sun added 2002-06-12
    Set ctlImage = Nothing
    Set ctlCaption = Nothing
    Set ctlLine = Nothing
    Set m_objParent = Nothing
    Set m_objRootWnd = Nothing
    'Mike added @ 2008-1-29
    Set ctlTag = Nothing
    m_blnDestroy = True
        
End Sub

Private Sub Class_Initialize()
    Set m_objRootWnd = Nothing
    Set m_objParent = Nothing
    Set m_objContainer = Nothing
    m_bytModifyFlag = DEF_OPERATION_INIT
End Sub

Private Sub ctlSelected_Click()
    IsSelected = Not IsSelected
End Sub

Private Sub ctlSelected_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    If Button = vbLeftButton Then
        'm_objParent.DragBeginEx
    End If
    If Not m_objRootWnd Is Nothing Then
        If Button = vbRightButton Then
            m_objRootWnd.PopupMenu m_objRootWnd.mEdit
        End If
    End If

End Sub

Private Sub ctlImage_DblClick()
On Error Resume Next

    m_objParent.ShowNodeProp m_NodeNo
End Sub

Public Sub SelectThisNode()
    Call ctlImage_MouseDown(vbLeftButton, 0, 0, 0)
End Sub

Private Sub ctlImage_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
On Error Resume Next

    Dim num As Integer
    Dim CtrlDown
    
    CtrlDown = (Shift And vbCtrlMask) > 0
    
    ctlImage.ZOrder
    If Button = vbLeftButton Then
        
        m_objParent.MouseSelectNode m_NodeIndex
        If ctlLine Is Nothing Then
            m_objParent.DragBegin ctlImage
        Else
            m_objParent.DragBegin ctlLine
        End If
        
        If Not CtrlDown Then
            m_objParent.ChangeAllSelectedNodesStatus False
            IsSelected = True
        Else
            IsSelected = Not IsSelected
        End If
                
    End If
    
    If Not m_objRootWnd Is Nothing Then
        If Button = vbRightButton Then
            m_objRootWnd.PopupMenu m_objRootWnd.mEdit
        End If
    End If
    
End Sub

' Sun added 2002-12-04
Public Property Let IsRootNode(f_Root As Boolean)
On Error Resume Next

    If f_Root Then
        ctlImage.BorderStyle = 1
    Else
        ctlImage.BorderStyle = 0
    End If
    
    m_blnRootNode = f_Root
        
End Property

' Sun added 2002-12-04
Public Property Get IsRootNode() As Boolean
On Error Resume Next

    IsRootNode = m_blnRootNode
        
End Property

Public Property Let IsSelected(f_blnSelected As Boolean)
On Error Resume Next

    Dim lv_StartPT As Byte, lv_EndPT As Byte
    
    '' System Nodes don't allowed to select
    If f_blnSelected And m_Nodeid < 256 Then Exit Property
    
    If m_blnSelected <> f_blnSelected Then
        
        ctlSelected.Visible = f_blnSelected
        m_blnSelected = f_blnSelected
    
    End If
    
End Property

Public Property Get IsSelected() As Boolean
On Error Resume Next

    IsSelected = m_blnSelected
        
End Property

'
' Sun added 2002-06-12
'
Public Property Get NodeDestroyed() As Boolean
On Error Resume Next

    NodeDestroyed = m_blnDestroy
    
End Property

Public Property Get Visible() As Boolean
On Error Resume Next

    Visible = ctlImage.Visible
End Property

Public Property Let Visible(ByVal vNewValue As Boolean)
On Error Resume Next

    ctlImage.Visible = vNewValue
    ctlCaption.Visible = (ctlCaption.Visible And vNewValue)
    
    '' Sun added 2008-02-18
    ctlTag.Visible = (ctlTag.Visible And vNewValue)
    
    ctlSelected.Visible = (IsSelected And vNewValue)
    If Not ctlLine Is Nothing Then
        ctlLine.Visible = vNewValue
    End If
    
End Property

'Mike modified 2008-8-19
'Public Sub VerifyVisible(ByVal f_nCurrentPage As Integer)
Public Function VerifyVisible(ByVal f_nCurrentPage As Integer) As Boolean
On Error Resume Next
    VerifyVisible = False
    
    If Me.InPage = f_nCurrentPage Then
    
        '' Sun added 2007-03-25
        If gSystem.intShowSysNodes > 0 Or Me.NodeNo >= 5 Then
            Me.Visible = True
            VerifyVisible = True
        Else
            Me.Visible = False
        End If
    Else
        Me.Visible = False
    End If

End Function

Public Property Get Left() As Integer
On Error Resume Next

    Left = ctlImage.Left
End Property

Public Property Let Left(ByVal vNewValue As Integer)
On Error Resume Next

    ctlImage.Left = vNewValue
    MoveCaptionWithNode
    MoveTagWithNode     'Mike added @ 08-1-29
End Property

Public Property Get Top() As Integer
On Error Resume Next

    Top = ctlImage.Top
End Property

Public Property Let Top(ByVal vNewValue As Integer)
On Error Resume Next

    ctlImage.Top = vNewValue
    MoveCaptionWithNode
    MoveTagWithNode     'Mike added @ 08-1-29
End Property

Public Property Get Width() As Integer
On Error Resume Next

    Width = m_Width
End Property

Public Property Let Width(ByVal vNewValue As Integer)
On Error Resume Next

    m_Width = vNewValue
    If ctlLine Is Nothing Then
        ctlImage.Width = vNewValue
    End If
    MoveCaptionWithNode
    MoveTagWithNode     'Mike added @ 08-1-29
End Property

Public Property Get Height() As Integer
On Error Resume Next

    Height = m_Height
End Property

Public Property Let Height(ByVal vNewValue As Integer)
On Error Resume Next

    m_Height = vNewValue
    If ctlLine Is Nothing Then
        ctlImage.Height = vNewValue
    End If
    MoveCaptionWithNode
    MoveTagWithNode     'Mike added @ 08-1-29
End Property

Public Property Get NodeCaption() As String
On Error Resume Next

    NodeCaption = ctlCaption.Caption
End Property

Public Property Let NodeCaption(ByVal vNewValue As String)
On Error Resume Next

    ctlCaption.Caption = vNewValue
    
    ' Sun added 2003-04-24
    NodeCaptionVisible = gblnShowNodeCaption
    
End Property

'Mike add those 2 property for Node Tag @ 2008-1-25
Public Property Get NodeTag() As String
On Error Resume Next
    NodeTag = ctlTag.Caption
End Property

Public Property Let NodeTag(ByVal strNewNodeTag As String)
On Error Resume Next
    ctlTag.Caption = strNewNodeTag
    NodeTagVisible = gblnShowNodeTag
End Property
'*******************  add end  *********************

Public Property Get NodeCaptionLeft() As Integer
On Error Resume Next

    NodeCaptionLeft = ctlCaption.Left
End Property

Public Property Let NodeCaptionLeft(ByVal vNewValue As Integer)
On Error Resume Next

    ctlCaption.Left = vNewValue
End Property

Public Property Get NodeCaptionTop() As Integer
On Error Resume Next

    NodeCaptionTop = ctlCaption.Top
End Property

Public Property Let NodeCaptionTop(ByVal vNewValue As Integer)
On Error Resume Next

    ctlCaption.Top = vNewValue
End Property

Public Property Get NodeCaptionVisible() As Boolean
On Error Resume Next

    NodeCaptionVisible = (ctlCaption.Visible And Me.Visible)
End Property

Public Property Let NodeCaptionVisible(ByVal vNewValue As Boolean)
On Error Resume Next

    'If Me.NodeNo <> [255-节点连线] Then
        If Trim(Me.NodeCaption) = "" Then vNewValue = False
        ctlCaption.Visible = (vNewValue And Me.Visible)
    'End If
    
    'Debug.Print "Node Caption:(" & Str(ctlCaption.Left) & ":" & Str(ctlCaption.Top) & ")-(" & Str(ctlCaption.Left + ctlCaption.Width) & ":" & Str(ctlCaption.Top + ctlCaption.Height)
End Property

Public Property Get NodeTagLeft() As Integer
On Error Resume Next

    NodeTagLeft = ctlTag.Left
End Property

Public Property Let NodeTagLeft(ByVal vNewValue As Integer)
On Error Resume Next

    ctlTag.Left = vNewValue
End Property

Public Property Get NodeTagTop() As Integer
On Error Resume Next

    NodeTagTop = ctlTag.Top
End Property

Public Property Let NodeTagTop(ByVal vNewValue As Integer)
On Error Resume Next

    ctlTag.Top = vNewValue
End Property

'**********************************************
' Mike Added Those 2 property @ 2008-1-29
Public Property Get NodeTagVisible() As Boolean
On Error Resume Next
    NodeTagVisible = (ctlTag.Visible And Me.Visible)
End Property

Public Property Let NodeTagVisible(ByVal vNewValue As Boolean)
On Error Resume Next
    '[255-节点连线]
    If Me.NodeNo <> 255 Then
        If Trim(Me.NodeTag) = "" Then vNewValue = False
        ctlTag.Visible = (vNewValue And Me.Visible)
    End If
    
    'If NodeID = 283 Then Debug.Print "Node Tag visible:"; Str(ctlTag.Visible) & " @(" & Str(ctlTag.Left) & ":" & Str(ctlTag.Top) & ")-(" & Str(ctlTag.Left + ctlTag.Width) & ":" & Str(ctlTag.Top + ctlTag.Height)
End Property
'**********************************************

Public Property Get NodeID() As Integer
On Error Resume Next

    NodeID = m_Nodeid
End Property

Public Property Let NodeID(ByVal vNewValue As Integer)
On Error Resume Next

    m_Nodeid = vNewValue

    UpdateNodeCaption
End Property

Public Property Get FlowID() As Byte
On Error Resume Next

    FlowID = m_FlowID
End Property

Public Property Let FlowID(ByVal vNewValue As Byte)
On Error Resume Next

    m_FlowID = vNewValue
End Property

Public Property Get NodeIndex() As Integer
On Error Resume Next

    NodeIndex = m_NodeIndex
End Property

Public Property Get NodeNo() As Byte
On Error Resume Next

    NodeNo = m_NodeNo
End Property

Public Property Let NodeNo(ByVal vNewValue As Byte)
On Error Resume Next

    m_NodeNo = vNewValue
    
    '' Sun replaced 2008-02-18
    ''' From
    'ctlImage.Picture = LoadPicture(GetNodePictureFileName(m_NodeNo))
    ''' To
    ctlImage.Picture = m_objParent.imgListNodes.ListImages("node" & Right("000" & Trim(Str(m_NodeNo)), 3)).Picture
    If Err Then
        ctlImage.Picture = LoadPicture(GetNodePictureFileName(m_NodeNo))
        Err.Clear
    End If
    
End Property

Private Sub UpdateNodeCaption()
    
    Dim lv_Caption As String
    
    ' Sun added 2002-03-28 to change Node Caption on changing of description
    If m_Description = "N/A" Then
        lv_Caption = ""
    Else
        lv_Caption = m_Description
    End If
    
    If m_Nodeid < 255 Then
        NodeCaption = Trim(Str(m_Nodeid))
    ElseIf m_NodeNo = 255 Then
        NodeCaption = lv_Caption
    Else
        NodeCaption = Trim(Str(m_Nodeid)) & "：" & lv_Caption
    End If

End Sub
    
Public Property Get Description() As String
On Error Resume Next

    Description = m_Description
End Property

Public Property Let ToolTipText(ByVal vNewValue As String)
On Error Resume Next

    ctlImage.ToolTipText = vNewValue
End Property

Public Property Get ToolTipText() As String
On Error Resume Next

    ToolTipText = ctlImage.ToolTipText
End Property

Public Property Let Description(ByVal vNewValue As String)
On Error Resume Next

    m_Description = Trim(vNewValue)
    
    ' Sun added 2002-03-20 to change tooltiptext on changing of description
    ToolTipText = "节点ID:" & NodeID & "->描述:" & m_Description
    
    UpdateNodeCaption
    
End Property

Public Property Get InPage() As Byte
On Error Resume Next

    InPage = m_InPage
End Property

Public Property Let InPage(ByVal vNewValue As Byte)
On Error Resume Next

    m_InPage = vNewValue
End Property

Public Property Get Data1() As String
On Error Resume Next

    Data1 = m_Data1
End Property

Public Property Let Data1(ByVal vNewValue As String)
On Error Resume Next

    m_Data1 = vNewValue
End Property

Public Property Get Data2() As String
On Error Resume Next

    Data2 = m_Data2
End Property

Public Property Let Data2(ByVal vNewValue As String)
On Error Resume Next

    m_Data2 = vNewValue
End Property

Public Sub MoveImageWithLine()
On Error Resume Next

    Dim lv_X As Long
    Dim lv_Y As Long
    Dim lv_Angel As Double
    Dim lv_Index As Integer
    Dim lv_Sin As Double, lv_Cos As Double
    Dim lv_adX As Integer, lv_adY As Integer
    
    lv_X = ctlLine.x2 - ctlLine.x1
    lv_Y = ctlLine.y2 - ctlLine.y1
    
    ' Move image
    Me.Left = ctlLine.x1 + lv_X * (1 - Def_GOLDEN_DIV)
    Me.Top = ctlLine.y1 + lv_Y * (1 - Def_GOLDEN_DIV)
    
    ' Change direction of arrow image
    lv_Y = -lv_Y
    If lv_X > 0 Then
        lv_Angel = Atn(lv_Y / lv_X)
        If lv_Angel < 0 Then lv_Angel = 2 * Def_PI + lv_Angel
    ElseIf lv_X < 0 Then
        lv_Angel = Atn(lv_Y / -lv_X)
        lv_Angel = Def_PI - lv_Angel
    Else
        If lv_Y >= 0 Then
            lv_Angel = Def_PI / 2
        Else
            lv_Angel = Def_PI * 1.5
        End If
    End If
    lv_Sin = Abs(Sin(lv_Angel) * 4)
    lv_Cos = Abs(Cos(lv_Angel) * 4)
    lv_adX = lv_Sin * Screen.TwipsPerPixelX
    lv_adY = lv_Cos * Screen.TwipsPerPixelY
    lv_Angel = lv_Angel * 180 / Def_PI
    lv_Index = Int(lv_Angel / 22.5 + 0.5) Mod 16
    
    m_objParent.clpArrows.ClipX = 8 * Int(lv_Index Mod 4)
    m_objParent.clpArrows.ClipY = 8 * Int(lv_Index / 4)
    ctlImage.Picture = m_objParent.clpArrows.Clip
    
    ' Adjust Image
    Me.Left = Me.Left - lv_adX
    Me.Top = Me.Top - lv_adY
    
End Sub

Public Property Get Line_X1() As Single
On Error Resume Next
    
    If Not ctlLine Is Nothing Then
        Line_X1 = ctlLine.x1
    Else
        Line_X1 = 0
    End If
    
End Property

Public Property Let Line_X1(ByVal vNewValue As Single)
On Error Resume Next
    
    If Not ctlLine Is Nothing Then
        ctlLine.x1 = vNewValue
        m_Width = ctlLine.x2 - ctlLine.x1
    End If
    
End Property

Public Property Get Line_X2() As Single
On Error Resume Next
    
    If Not ctlLine Is Nothing Then
        Line_X2 = ctlLine.x2
    Else
        Line_X2 = 0
    End If
    
End Property

Public Property Let Line_X2(ByVal vNewValue As Single)
On Error Resume Next
    
    If Not ctlLine Is Nothing Then
        ctlLine.x2 = vNewValue
        m_Width = ctlLine.x2 - ctlLine.x1
    End If
    
End Property

Public Property Get Line_Y1() As Single
On Error Resume Next
    
    If Not ctlLine Is Nothing Then
        Line_Y1 = ctlLine.y1
    Else
        Line_Y1 = 0
    End If
    
End Property

Public Property Let Line_Y1(ByVal vNewValue As Single)
On Error Resume Next
    
    If Not ctlLine Is Nothing Then
        ctlLine.y1 = vNewValue
        m_Height = ctlLine.y2 - ctlLine.y1
    End If
    
End Property

Public Property Get Line_Y2() As Single
On Error Resume Next
    
    If Not ctlLine Is Nothing Then
        Line_Y2 = ctlLine.y2
    Else
        Line_Y2 = 0
    End If
    
End Property

Public Property Let Line_Y2(ByVal vNewValue As Single)
On Error Resume Next
    
    If Not ctlLine Is Nothing Then
        ctlLine.y2 = vNewValue
        m_Height = ctlLine.y2 - ctlLine.y1
    End If
    
End Property

Public Property Get Line_Color() As Long
On Error Resume Next

    Line_Color = ctlLine.BorderColor
End Property

Public Property Let Line_Color(ByVal vNewValue As Long)
On Error Resume Next

#If LOCALTEST = 1 Then
#Else
    ctlLine.BorderColor = vNewValue
#End If

End Property

Public Property Get Line_Width() As Integer
On Error Resume Next

    Line_Width = ctlLine.BorderWidth
End Property

Public Property Let Line_Width(ByVal vNewValue As Integer)
On Error Resume Next
    
#If LOCALTEST = 1 Then
#Else
    ctlLine.BorderWidth = vNewValue
#End If

End Property

Public Property Get Line_Style() As Integer
On Error Resume Next

    Line_Style = ctlLine.BorderStyle
End Property

Public Property Let Line_Style(ByVal vNewValue As Integer)
On Error Resume Next
    
    ctlLine.BorderStyle = vNewValue

End Property

Public Property Get Picture() As Picture
On Error Resume Next
    
    Set Picture = ctlImage.Picture
    
End Property

' Move Label object & Handles when Node is moved or Resized
'
Private Sub MoveCaptionWithNode()
    ctlCaption.Left = ctlImage.Left + ctlImage.Width * (1 - Def_GOLDEN_DIV)
    ctlCaption.Top = ctlImage.Top + ctlImage.Height + 20
    ctlSelected.Move ctlImage.Left - ctlSelected.Width, ctlImage.Top - ctlSelected.Height / 2
End Sub

'Mike added @ 2008-1-29
Private Sub MoveTagWithNode()
    ctlTag.Left = ctlImage.Left + ctlImage.Width * (1 - Def_GOLDEN_DIV)
    ctlTag.Top = ctlImage.Top - ctlTag.Height - 20
    'ctlSelected.Move ctlImage.Left - ctlSelected.Width, ctlImage.Top - ctlSelected.Height / 2
End Sub

Public Sub MoveLinesOnNode()
    m_objParent.MoveLinesOnNode m_NodeIndex
    
End Sub

Public Property Get ModifyFlag() As Byte
On Error Resume Next

    ModifyFlag = m_bytModifyFlag
End Property

Public Property Let ModifyFlag(ByVal vNewValue As Byte)
On Error Resume Next
    
    m_bytModifyFlag = vNewValue

End Property
