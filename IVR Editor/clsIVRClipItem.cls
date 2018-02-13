VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIVRClipItem"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'-----------------------------------------------
' IVR clipboard item class
'-----------------------------------------------
Option Explicit

'' Private variable
Private m_clpContents() As New clsIVRClipItemData               '' 节点内容数组

'' Properties
Private m_bytOperation As Byte                                  '' 粘贴板操作，参见“操作定义”
Private m_bytNodeCount As Byte                                  '' 内容数组包含的节点数量
Private m_bytNodePointer As Byte                                '' 节点当前指针

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
    
    m_bytOperation = 0
    
End Sub

Public Property Let Operation(f_bytOperation As Byte)
On Error Resume Next
  
    m_bytOperation = f_bytOperation
        
End Property

Public Property Get Operation() As Byte
On Error Resume Next

    Operation = m_bytOperation
        
End Property

Public Property Let NodeCount(f_bytNodeCount As Byte)
On Error Resume Next
  
    If m_bytNodeCount <> f_bytNodeCount Then
        ReDim Preserve m_clpContents(f_bytNodeCount) As New clsIVRClipItemData
        m_bytNodeCount = f_bytNodeCount
    End If
        
End Property

Public Property Get NodeCount() As Byte
On Error Resume Next

    NodeCount = m_bytNodeCount
        
End Property

Public Property Let NodePointer(f_bytNodePtr As Byte)
On Error Resume Next
  
    m_bytNodePointer = f_bytNodePtr
        
End Property

Public Property Get NodePointer() As Byte
On Error Resume Next

    NodePointer = m_bytNodePointer
        
End Property

Public Property Get Contents(ByVal f_Index As Byte) As clsIVRClipItemData
On Error Resume Next

    If f_Index < m_bytNodeCount Then
        Set Contents = m_clpContents(f_Index)
    End If
        
End Property

Public Sub GetItemRect(x1 As Long, y1 As Long, x2 As Long, y2 As Long)
On Error Resume Next

    Dim lv_Index As Byte
    
    If m_bytNodeCount > 0 Then
    
        x1 = m_clpContents(0).IntN_left
        y1 = m_clpContents(0).IntN_top
        x2 = m_clpContents(0).IntN_left + m_clpContents(0).IntN_width
        y2 = m_clpContents(0).IntN_top + m_clpContents(0).IntN_height
    
        For lv_Index = 1 To m_bytNodeCount - 1
            If x1 > m_clpContents(lv_Index).IntN_left Then
                x1 = m_clpContents(lv_Index).IntN_left
            End If
            If y1 > m_clpContents(lv_Index).IntN_top Then
                y1 = m_clpContents(lv_Index).IntN_top
            End If
            If x2 < m_clpContents(lv_Index).IntN_left + m_clpContents(lv_Index).IntN_width Then
                x2 = m_clpContents(lv_Index).IntN_left + m_clpContents(lv_Index).IntN_width
            End If
            If y2 < m_clpContents(lv_Index).IntN_top + m_clpContents(lv_Index).IntN_height Then
                y2 = m_clpContents(lv_Index).IntN_top + m_clpContents(lv_Index).IntN_height
            End If
        Next
        
    End If
    
End Sub
