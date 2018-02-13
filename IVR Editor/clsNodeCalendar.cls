VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsNodeCalendar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'=================================
' Calender Nodes properties
'=================================
' Calender Node exist flag
Private m_nNodeMain As Integer

' Count of Calendar Nodes
Private m_nNodeCount As Integer


Public Property Get MainNode() As Integer
On Error Resume Next
    MainNode = m_nNodeMain
End Property

Public Property Let MainNode(ByVal vNewValue As Integer)
On Error Resume Next
    m_nNodeMain = vNewValue
End Property
