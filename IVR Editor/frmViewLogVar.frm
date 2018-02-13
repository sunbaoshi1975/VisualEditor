VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F856EC8B-F03C-4515-BDC6-64CBD617566A}#7.0#0"; "FPSPR70.ocx"
Begin VB.Form frmViewLogVar 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "日志使用情况"
   ClientHeight    =   5145
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9615
   Icon            =   "frmViewLogVar.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5145
   ScaleWidth      =   9615
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Tag             =   "1455"
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      Height          =   300
      Index           =   2
      Left            =   5850
      Picture         =   "frmViewLogVar.frx":030A
      ScaleHeight     =   240
      ScaleWidth      =   240
      TabIndex        =   7
      Top             =   4950
      Visible         =   0   'False
      Width           =   300
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   1
      Left            =   5250
      Picture         =   "frmViewLogVar.frx":0454
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   6
      Top             =   4950
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.PictureBox picStatus 
      AutoSize        =   -1  'True
      Height          =   540
      Index           =   0
      Left            =   4890
      Picture         =   "frmViewLogVar.frx":0D1E
      ScaleHeight     =   480
      ScaleWidth      =   480
      TabIndex        =   5
      Top             =   4980
      Visible         =   0   'False
      Width           =   540
   End
   Begin VB.CommandButton CommandExit 
      Caption         =   "&X退出"
      Height          =   315
      Left            =   3510
      TabIndex        =   4
      Tag             =   "1144"
      Top             =   5010
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.CommandButton CommandSave 
      Caption         =   "&S保存"
      Default         =   -1  'True
      Height          =   315
      Left            =   2250
      TabIndex        =   3
      Tag             =   "1007"
      Top             =   5010
      Visible         =   0   'False
      Width           =   1035
   End
   Begin VB.PictureBox picLOG 
      BorderStyle     =   0  'None
      Height          =   4455
      Index           =   0
      Left            =   150
      ScaleHeight     =   4455
      ScaleWidth      =   9315
      TabIndex        =   1
      Top             =   480
      Width           =   9315
      Begin FPSpreadADO.fpSpread vasLOG 
         Height          =   4365
         Left            =   30
         TabIndex        =   2
         Top             =   60
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
         MaxCols         =   3
         MaxRows         =   16
         ScrollBarExtMode=   -1  'True
         SpreadDesigner  =   "frmViewLogVar.frx":1160
         UserResize      =   1
      End
   End
   Begin MSComctlLib.TabStrip tbsViewer 
      Height          =   4965
      Left            =   90
      TabIndex        =   0
      Top             =   90
      Width           =   9465
      _ExtentX        =   16695
      _ExtentY        =   8758
      _Version        =   393216
      BeginProperty Tabs {1EFB6598-857C-11D1-B16A-00C0F0283628} 
         NumTabs         =   1
         BeginProperty Tab1 {1EFB659A-857C-11D1-B16A-00C0F0283628} 
            Caption         =   "日志使用"
            Key             =   "日志使用"
            Object.Tag             =   "1456"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmViewLogVar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

' 数据是否被修改
Dim f_DataChanged As Boolean

Private Sub CommandExit_Click()
    Unload Me
End Sub

Private Sub CommandSave_Click()
On Error Resume Next

    If f_DataChanged Then

        '改变鼠标指针形状->沙漏光标
        mdlcommon.ChangeMousePointer vbHourglass, True
    
        '改变鼠标指针形状->箭头光标
        mdlcommon.ChangeMousePointer vbDefault, True
    
    End If

On Error GoTo 0
End Sub

Private Sub Form_KeyPress(KeyAscii As Integer)
    If KeyAscii = 27 Then
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    
    '改变鼠标指针形状->沙漏光标
    mdlcommon.ChangeMousePointer vbHourglass, True

    tbsViewer.Tabs(1).Selected = True
    Call RefreshLogSpread
    
    f_DataChanged = False
    
    '改变鼠标指针形状->箭头光标
    mdlcommon.ChangeMousePointer vbDefault, True
    LoadResStrings Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    If f_DataChanged Then
        If Message("Q005") = vbNo Then
            Cancel = True
        End If
    End If
End Sub

Private Sub tbsViewer_Click()
    Dim i As Integer
    
    'show and enable the selected tab's controls
    'and hide and disable all others
    For i = 0 To tbsViewer.Tabs.Count - 1
        If i = tbsViewer.SelectedItem.Index - 1 Then
            picLOG(i).Left = 150
            picLOG(i).Top = 480
            picLOG(i).Visible = True
            picLOG(i).Enabled = True
        Else
            picLOG(i).Left = -20000
            picLOG(i).Enabled = False
        End If
    Next
End Sub

Private Sub RefreshLogSpread()
On Error Resume Next

Dim lv_loop As Integer
Dim lv_LogUsed(16) As Integer
Dim lv_RelateNode(16, 20) As Integer
Dim lv_LogPosition As Integer

'' Sun added 2002-06-12
Dim lv_SubLoop As Integer
Dim lv_StrTemp As String

    For lv_loop = 1 To gCallFlow.NewNodeID
        lv_LogPosition = gCallFlow.GetNodeLogged(lv_loop)
        If lv_LogPosition > 0 And lv_LogPosition <= 16 Then
            If lv_LogUsed(lv_LogPosition) <= 20 Then
                lv_RelateNode(lv_LogPosition, lv_LogUsed(lv_LogPosition)) = lv_loop
            End If
            lv_LogUsed(lv_LogPosition) = lv_LogUsed(lv_LogPosition) + 1
        End If
    Next

    vasLOG.Enabled = False
    For lv_loop = 1 To vasLOG.MaxRows
        vasLOG.Row = lv_loop
        vasLOG.Col = 1
        vasLOG.Text = Str(lv_loop) & LoadNationalResString(1179)
        
        vasLOG.Col = 2
        If lv_LogUsed(lv_loop) <= 0 Then
            vasLOG.TypePictPicture = picStatus(0).Picture
            vasLOG.Col = 3
            vasLOG.Text = ""
        ElseIf lv_LogUsed(lv_loop) = 1 Then
            vasLOG.TypePictPicture = picStatus(1).Picture
            vasLOG.Col = 3
            vasLOG.Text = Trim(gCallFlow.Node(lv_RelateNode(lv_loop, 0)).NodeCaption)
        Else
            vasLOG.TypePictPicture = picStatus(2).Picture
            vasLOG.Col = 3
            lv_StrTemp = Trim(Str(lv_LogUsed(lv_loop))) & LoadNationalResString(1466) & Trim(gCallFlow.Node(lv_RelateNode(lv_loop, 0)).NodeCaption)
            
            '' Sun added 2002-06-12
            If lv_LogUsed(lv_loop) > 20 Then
                lv_LogPosition = 20
            Else
                lv_LogPosition = lv_LogUsed(lv_loop) - 1
            End If
            
            For lv_SubLoop = 1 To lv_LogPosition
                lv_StrTemp = lv_StrTemp & "； " & Trim(gCallFlow.Node(lv_RelateNode(lv_loop, lv_SubLoop)).NodeCaption)
            Next
            vasLOG.Text = lv_StrTemp
            
        End If
    Next
    
    vasLOG.Enabled = True
    
On Error GoTo 0
End Sub

Private Sub vasLOG_Change(ByVal Col As Long, ByVal Row As Long)
    If Col Then f_DataChanged = True
End Sub
