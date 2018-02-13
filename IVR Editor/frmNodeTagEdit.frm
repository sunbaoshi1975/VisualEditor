VERSION 5.00
Begin VB.Form frmNodeTagEdit 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "节点标签"
   ClientHeight    =   1245
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   2145
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1245
   ScaleWidth      =   2145
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Tag             =   "1944"
   Begin VB.CommandButton cmdCancel 
      Caption         =   "取消"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Tag             =   "1144"
      Top             =   840
      Width           =   735
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "保存"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Tag             =   "1007"
      Top             =   840
      Width           =   735
   End
   Begin VB.TextBox txtNodeTag 
      BeginProperty Font 
         Name            =   "Arial Black"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1695
   End
   Begin VB.Label lblNodeTag 
      AutoSize        =   -1  'True
      Caption         =   "节点标签 ："
      Height          =   195
      Left            =   240
      TabIndex        =   1
      Tag             =   "1947"
      Top             =   120
      Width           =   945
   End
End
Attribute VB_Name = "frmNodeTagEdit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Public iNodeID As Integer
Public byNodeNo As Byte

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdSave_Click()
    gCallFlow.Node(gCallFlow.NodeSelectedID).NodeTag = Trim(txtNodeTag.Text)
    ' save the new value
    gCallFlow.UpdateIvrRecord iNodeID, byNodeNo
    
    cmdSave.Enabled = False
    Unload Me
End Sub

Private Sub Form_Load()
    txtNodeTag.Text = gCallFlow.Node(gCallFlow.NodeSelectedID).NodeTag
    
    cmdSave.Enabled = False
    
    LoadResStrings Me
End Sub

Private Sub txtNodeTag_Change()
    cmdSave.Enabled = True
End Sub

Private Sub txtNodeTag_GotFocus()
    txtNodeTag.SelStart = 0
    txtNodeTag.SelLength = Len(txtNodeTag.Text)
End Sub

Private Sub txtNodeTag_KeyPress(KeyAscii As Integer)
    If (Not (KeyAscii >= 48 And KeyAscii <= 57)) And _
        (Not (KeyAscii >= 97 And KeyAscii <= 122)) And _
        (Not (KeyAscii >= 65 And KeyAscii <= 90)) And _
        KeyAscii <> 8 And KeyAscii <> 7 Then KeyAscii = 0
End Sub
