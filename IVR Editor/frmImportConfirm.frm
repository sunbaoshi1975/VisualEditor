VERSION 5.00
Begin VB.Form frmImportConfirm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "导入提示"
   ClientHeight    =   2550
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4275
   Icon            =   "frmImportConfirm.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2550
   ScaleWidth      =   4275
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1660"
   Begin VB.Frame Frame1 
      Height          =   1905
      Left            =   60
      TabIndex        =   5
      Top             =   60
      Width           =   4155
      Begin VB.TextBox txtCallFlowID 
         Alignment       =   2  'Center
         Height          =   345
         Left            =   1920
         TabIndex        =   2
         Top             =   1320
         Width           =   1035
      End
      Begin VB.OptionButton optChoice 
         Caption         =   "新项目ID"
         Height          =   285
         Index           =   1
         Left            =   210
         TabIndex        =   1
         Tag             =   "1659"
         Top             =   1350
         Value           =   -1  'True
         Width           =   1605
      End
      Begin VB.OptionButton optChoice 
         Caption         =   "覆盖原有流程"
         Height          =   285
         Index           =   0
         Left            =   210
         TabIndex        =   0
         Tag             =   "1658"
         Top             =   930
         Width           =   2415
      End
      Begin VB.Label lblCaption 
         Caption         =   "流程：999 已经存在，请选择另外一个项目ID或者覆盖原有流程。"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FF0000&
         Height          =   750
         Left            =   150
         TabIndex        =   6
         Top             =   270
         Width           =   3855
      End
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定(&C)"
      Default         =   -1  'True
      Height          =   315
      Left            =   900
      TabIndex        =   3
      Tag             =   "1372"
      Top             =   2130
      Width           =   1065
   End
   Begin VB.CommandButton cmdExit 
      Caption         =   "退出(&E)"
      Height          =   315
      Left            =   2340
      TabIndex        =   4
      Tag             =   "1144"
      Top             =   2130
      Width           =   1065
   End
End
Attribute VB_Name = "frmImportConfirm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim m_bytCallFlowID As Byte

Private Sub cmdExit_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
    If optChoice(0).value Then
        m_bytCallFlowID = gCallFlow.CallFlowID
    Else
        m_bytCallFlowID = CByte(Val(txtCallFlowID) Mod 256)
    End If
    Unload Me
End Sub

Private Sub Form_Load()
On Error Resume Next

    Dim lv_Str As String
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True
   
    lblCaption = LoadNationalResString(1103) & Str(gCallFlow.CallFlowID) & LoadNationalResString(1130)
    
    txtCallFlowID = gCallFlow.CallFlowID + 1
    m_bytCallFlowID = 0
    
    '改变鼠标指针形状->箭头光标
    ChangeMousePointer vbDefault, True
    LoadResStrings Me

On Error GoTo 0
End Sub

Private Sub Form_Unload(Cancel As Integer)
    gCallFlow.CallFlowID = m_bytCallFlowID
End Sub

Private Sub optChoice_Click(Index As Integer)
    txtCallFlowID.Enabled = optChoice(1).value
End Sub

Private Sub txtCallFlowID_GotFocus()
    txtCallFlowID.SelStart = 0
    txtCallFlowID.SelLength = Len(txtCallFlowID)
End Sub
