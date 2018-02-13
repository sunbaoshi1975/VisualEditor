VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsIVRClipItemData"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'' Properties
Private m_IntP_id As Byte                                          ''流程编号
Private m_IntN_index As Integer                                    ''节点Index
Private m_IntN_id As Integer                                       ''节点ID
Private m_IntN_no As Byte                                          ''节点编号
Private m_IntN_page As Integer                                     ''节点所在页
Private m_IntN_left As Integer                                     ''左
Private m_IntN_top As Integer                                      ''上
Private m_IntN_height As Integer                                   ''高
Private m_IntN_width As Integer                                    ''宽
Private m_StrN_description As String                               ''描述
Private m_StrN_data1 As String                                     ''数据段1
Private m_StrN_data2 As String                                     ''数据段2
Private m_Line_X1 As Single
Private m_Line_X2 As Single
Private m_Line_Y1 As Single
Private m_Line_Y2 As Single
Private m_Line_Color As Long
Private m_Line_Width As Integer
Private m_Line_Style As Integer
Private m_bytModifyFlag As Byte                     '' Node Modify Flag，参见“操作定义”


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
    
    m_IntP_id = 0
    m_IntN_index = 0
    m_IntN_id = 0
    m_IntN_no = 0
    m_IntN_page = 0
    m_IntN_left = 0
    m_IntN_top = 0
    m_IntN_height = 0
    m_IntN_width = 0
    m_StrN_description = ""
    m_StrN_data1 = ""
    m_StrN_data2 = ""
    m_bytModifyFlag = DEF_OPERATION_INIT
    
End Sub

Public Property Let IntP_id(f_bytPID As Byte)
On Error Resume Next
  
    m_IntP_id = f_bytPID
        
End Property

Public Property Get IntP_id() As Byte
On Error Resume Next

    IntP_id = m_IntP_id
        
End Property

Public Property Let IntN_index(f_nIndex As Integer)
On Error Resume Next
  
    m_IntN_index = f_nIndex
        
End Property

Public Property Get IntN_index() As Integer
On Error Resume Next

    IntN_index = m_IntN_index
        
End Property

Public Property Let IntN_id(f_nID As Integer)
On Error Resume Next
  
    m_IntN_id = f_nID
        
End Property

Public Property Get IntN_id() As Integer
On Error Resume Next

    IntN_id = m_IntN_id
        
End Property

Public Property Let IntN_no(f_bytNO As Byte)
On Error Resume Next
  
    m_IntN_no = f_bytNO
        
End Property

Public Property Get IntN_no() As Byte
On Error Resume Next

    IntN_no = m_IntN_no
        
End Property

Public Property Let IntN_page(f_nPage As Integer)
On Error Resume Next
  
    m_IntN_page = f_nPage
        
End Property

Public Property Get IntN_page() As Integer
On Error Resume Next

    IntN_page = m_IntN_page
        
End Property

Public Property Let IntN_left(f_nLeft As Integer)
On Error Resume Next
  
    m_IntN_left = f_nLeft
        
End Property

Public Property Get IntN_left() As Integer
On Error Resume Next

    IntN_left = m_IntN_left
        
End Property

Public Property Let IntN_top(f_nTop As Integer)
On Error Resume Next
  
    m_IntN_top = f_nTop
        
End Property

Public Property Get IntN_top() As Integer
On Error Resume Next

    IntN_top = m_IntN_top
        
End Property

Public Property Let IntN_height(f_nHeight As Integer)
On Error Resume Next
  
    m_IntN_height = f_nHeight
        
End Property

Public Property Get IntN_height() As Integer
On Error Resume Next

    IntN_height = m_IntN_height
        
End Property

Public Property Let IntN_width(f_nWidth As Integer)
On Error Resume Next
  
    m_IntN_width = f_nWidth
        
End Property

Public Property Get IntN_width() As Integer
On Error Resume Next

    IntN_width = m_IntN_width
        
End Property

Public Property Let StrN_description(f_str As String)
On Error Resume Next
  
    m_StrN_description = f_str
        
End Property

Public Property Get StrN_description() As String
On Error Resume Next

    StrN_description = m_StrN_description
        
End Property

Public Property Let StrN_data1(f_str As String)
On Error Resume Next
  
    m_StrN_data1 = f_str
        
End Property

Public Property Get StrN_data1() As String
On Error Resume Next

    StrN_data1 = m_StrN_data1
        
End Property

Public Property Let StrN_data2(f_str As String)
On Error Resume Next
  
    m_StrN_data2 = f_str
        
End Property

Public Property Get StrN_data2() As String
On Error Resume Next

    StrN_data2 = m_StrN_data2
        
End Property

Public Property Get Line_X1() As Single
On Error Resume Next

    Line_X1 = m_Line_X1

End Property

Public Property Let Line_X1(ByVal vNewValue As Single)
On Error Resume Next

    m_Line_X1 = vNewValue

End Property

Public Property Get Line_X2() As Single
On Error Resume Next

    Line_X2 = m_Line_X2

End Property

Public Property Let Line_X2(ByVal vNewValue As Single)
On Error Resume Next

    m_Line_X2 = vNewValue

End Property

Public Property Get Line_Y1() As Single
On Error Resume Next

    Line_Y1 = m_Line_Y1

End Property

Public Property Let Line_Y1(ByVal vNewValue As Single)
On Error Resume Next

    m_Line_Y1 = vNewValue

End Property

Public Property Get Line_Y2() As Single
On Error Resume Next

    Line_Y2 = m_Line_Y2

End Property

Public Property Let Line_Y2(ByVal vNewValue As Single)
On Error Resume Next

    m_Line_Y2 = vNewValue

End Property

Public Property Get Line_Color() As Long
On Error Resume Next

    Line_Color = m_Line_Color
End Property

Public Property Let Line_Color(ByVal vNewValue As Long)
On Error Resume Next

    m_Line_Color = vNewValue

End Property

Public Property Get Line_Width() As Integer
On Error Resume Next

    Line_Width = m_Line_Width
End Property

Public Property Let Line_Width(ByVal vNewValue As Integer)
On Error Resume Next

    m_Line_Width = vNewValue

End Property

Public Property Get Line_Style() As Integer
On Error Resume Next

    Line_Style = m_Line_Style
End Property

Public Property Let Line_Style(ByVal vNewValue As Integer)
On Error Resume Next

    m_Line_Style = vNewValue

End Property

Public Property Get ModifyFlag() As Byte
On Error Resume Next

    ModifyFlag = m_bytModifyFlag
End Property

Public Property Let ModifyFlag(ByVal vNewValue As Byte)
On Error Resume Next
    
    m_bytModifyFlag = vNewValue

End Property

