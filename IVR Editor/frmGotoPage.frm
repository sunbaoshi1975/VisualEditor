VERSION 5.00
Begin VB.Form frmGotoPage 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ҳ����ת"
   ClientHeight    =   1485
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3690
   Icon            =   "frmGotoPage.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1485
   ScaleWidth      =   3690
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1731"
   Begin VB.CommandButton CommandCreate 
      Caption         =   "ȷ��(&C)"
      Default         =   -1  'True
      Height          =   315
      Left            =   540
      TabIndex        =   3
      Tag             =   "1372"
      Top             =   900
      Width           =   1065
   End
   Begin VB.CommandButton exit_Command 
      Caption         =   "�˳�(&E)"
      Height          =   315
      Left            =   1920
      TabIndex        =   2
      Tag             =   "1144"
      Top             =   900
      Width           =   1065
   End
   Begin VB.ComboBox cmbPage 
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
      ItemData        =   "frmGotoPage.frx":014A
      Left            =   2100
      List            =   "frmGotoPage.frx":014C
      Style           =   2  'Dropdown List
      TabIndex        =   1
      ToolTipText     =   "��ѡ��"
      Top             =   180
      Width           =   1365
   End
   Begin VB.Label lblHint 
      AutoSize        =   -1  'True
      Caption         =   "������Ŀ��ҳ��ҳ��"
      Height          =   195
      Left            =   120
      TabIndex        =   0
      Tag             =   "1108"
      Top             =   240
      Width           =   1620
   End
End
Attribute VB_Name = "frmGotoPage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub CommandCreate_Click()
    Dim lv_nPage As Integer
    
    If gCallFlow.CallFlowID > 0 Then
        If cmbPage.ListIndex >= 0 Then
            lv_nPage = cmbPage.ItemData(cmbPage.ListIndex)
            If gCallFlow.CurrentPage <> lv_nPage Then
                Mdlfunction.GotoAnotherPage lv_nPage
            End If
        End If
    End If
    
    Unload Me
End Sub

Private Sub exit_Command_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    Dim lv_nLoop As Integer
    
    With cmbPage
        .Clear
        If gCallFlow.CallFlowID > 0 Then
            For lv_nLoop = 1 To gCallFlow.PageCount
                .AddItem Str(lv_nLoop)
                .ItemData(.ListCount - 1) = lv_nLoop
            Next
            .ListIndex = SearchItemDataIndex(cmbPage, CLng(gCallFlow.CurrentPage), 0)
        End If
    End With
    
    LoadResStrings Me
End Sub
